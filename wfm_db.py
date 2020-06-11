from pymongo import MongoClient
#from pymongo import MongoClient.errr
import logging
from datetime import datetime, timedelta
import wfm_xl
import wfm_helpers
#from openpyxl import load_workbook
import uuid

class Wfm_db:
    def __init__ (self, env:dict, cnf:dict, logger):
        try:
            self.__lgr      = logger
            self.__env      = env.copy()
            self.__dic      = cnf.copy()
            self.__dic_fc   = {}
            self.__client   = MongoClient()
            self.__cnt_iv   = 0
            self.__max_days = 0
            self.__max_rows = 0
            self.__max_cols = 0
            self.__strt_tm  = datetime.now()
            self.__xl_obj   = wfm_xl.wfm_xl (env, cnf, self.__lgr)

        except Exception as e:
            raise Exception (e)

    def __del__ (self):
        self.__client.close()
        del self.__client
        del self.__env
        del self.__lgr
        del self.__dic
        del self.__dic_fc
        del self.__strt_tm
        del self.__xl_obj

    def dbCheckConnectionStatus (self, db):
        """
        Checks the connection to a Mongo database.
        
        Parameters
        ----------
        db: object - MongoDB database connection object

        Returns
        ---------
        int:  1 for connection is live and database can accept commands
             -1 if database is not responsive
        """
        con_stat = db.command ( { "connectionStatus": 1 } )
        if con_stat ['ok'] != 1.0:
            return -1
        else:
            return 1

    def dbConnect (self, url, port):
        """
        Attempts to connect to a MongoDB server.
        
        Parameters
        ----------
        url: string connection string such as 'mongodb://10.10.10.10
        port: int the port to connect to, such as '27017'

        Returns
        ---------
        object: Mongo client
        """

        if self.__lgr.root.level == logging.DEBUG:
            self.__lgr.info (str.format(self.__dic['res_strings']['info']['0029'], url, port, self.__env['db']['connectTimeoutMS'], self.__env['db']['socketTimeoutMS']))
        else:
            self.__lgr.info (str.format(self.__dic['res_strings']['info']['0044'], url, port))

        self.__client = MongoClient(url, port, connectTimeoutMS=self.__env['db']['connectTimeoutMS'], socketTimeoutMS=self.__env['db']['socketTimeoutMS'])

        if self.__client is None:
            raise Exception (str.format (self.__dic['res_strings']['fatal']['0002'], self.__env['db']['url'], self.__env['db']['port']))

        self.__lgr.info (str.format(self.__dic['res_strings']['info']['0042'], url, port))
        return self.__client


    def dbImportXL (self):
        try:
            client = self.dbConnect (self.__env['db']['url'], self.__env['db']['port'])
            self.__lgr.info (str.format(self.__dic['res_strings']['info']['0043'], self.__env['db']['db-name']))

            db = client [str (self.__env['db']['db-name'])]

            if self.dbCheckConnectionStatus (db) != 1:
                raise Exception (str.format (self.__dic['res_strings']['fatal']['0002'], self.__env['db']['url'], self.__env['db']['port']))

            self.__lgr.info (str.format(self.__dic['res_strings']['info']['0030'], self.__env['db']['db-name']))
            
            # ================== Create wfm_obj ==================
            # Select the active sheet
            collection = db [str (self.__env['db']['col-in'])]
            # Process spreadsheet and update db
            # we are moving row by row
            cnt_days = 0
            cnt_omit = 0
            cnt_records = 0

            #TODO: Add mapping of Excel fields to database fields
            # For every day to be imported from spreadsheet do...
            self.__lgr.info (self.__dic['res_strings']['info']['0045'])
            for c in range (2, self.__xl_obj.max_days() + 2, 1):
                record = {
                    "_id":0, 
                    "day_ary": [{
                            "year":     "$int",       # The year
                            "month":    "$int",       # Month 1 - 12
                            "day":      "$int",       # Day 1 - 31
                            "dow":      "$string",    # Day of Week
                            "s_mnth":   "$string",    # Month string
                            "times":    [],           # Interval Times
                            "agents":   [],           # Number of agents
                            "tx_all":   [],           # Number of offered transactions
                            "tx_ans":   [],           # Number of handeled transactions
                            "tx_abn":   [],           # Number of abandoned tarnsactions
                            "r_abn":    [],           # Abandonment Rate
                            "sl":       [],           # Service Level
                            "a_sa":     [],           # Average Speed to Answer
                            "a_wait":   [],           # Average Wait Time
                            "q_rate":   [],           # Average Queue Rate
                            "q_time":   [],           # Average Incident Queue Time
                            "a_it":     [],           # Average Incident Time
                            "aiw":      [],           # After Incident Wrap Time
                            "a_ht":     [],           # Average Handle Time
                            "r_avail":  [],           # Agents Availability
                            "r_util":   []            # Agents Utilisation
                    }]}
                # Fill the record dictionary to be processed
                self.__xl_obj.xlFillSingleRecord (record, c)
                # Check if the last records year is different to this record - in that case need to create new _id
                qry = {"_id": record['_id']}
                # Check if a document with this year exists
                query_res = collection.find_one(qry)
                if query_res == None:
                    collection.insert_one(record)
                    cnt_days += 1
                    cnt_records += len (record['day_ary'][0]['tx_all'])
                else:
                    tmp_date = "{}/{}/{}".format(record['day_ary'][0]['year'], record['day_ary'][0]['month'], record['day_ary'][0]['day'])
                    self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0002'], tmp_date, collection.name))
                    cnt_omit += 1
                    continue
            self.__lgr.info (str.format(self.__dic['res_strings']['info']['0034'], cnt_days, cnt_records, cnt_omit))

#        logger_wfm.info(dic_cnf['res_strings']['info']['0034'])
                
            del record
            del collection
            client.close()
            del client
            del db
            return False

        except Exception as e:
            self.__lgr.fatal (e)    
            raise Exception (e)

# ================== Construct the forecast dictionary  -- NOT THREAD SAFE -- ==================
    def dbCreateDictionary (self):
        """
        Function creates the dictionary used for forecasting the agent numbers.
        
        Parameters
        ----------
        dic_fc : dictionary dict {}
            Generates the excel reports based on forecast data provided in fc.
            The report parameters are defined in file workforce.conf
        """
        try:
            # ================== Sanity check for data read - compare parts of framework file with actual data  ==================
            self.__xl_obj.xlCheckImportDataValidity ()
            # ================== Build the Dictionary template for the required calculations ==================
            self.__dic_fc = self.__xl_obj.xlCreateDictionary ()
            # ================== Sanity check of framework data supplied ==================
            wfm_helpers.helperCheckFrameworkData (self.__dic, self.__lgr)
            return
        except Exception as ex:
            raise (ex)

    def dbCreateForecast (self):
        # record = {"_id": '', "date": '', "sdate":"","smonth":"" , "dow":"", "times": [], "calls": [], "util": [], "sla": [], "asa": [], "att": [], "aiw": [], "abandon": [], "qpercent": [], "qtime": [], "qsize": []}
        pass


    def __checkSyntax (self, record):
# TODO: add syntax check rules to make sure the data imported makes sense
        return True

