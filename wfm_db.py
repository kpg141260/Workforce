from pymongo import MongoClient
from pymongo import errors as mdb_errors
#from pymongo import MongoClient.errr
import logging
from datetime import datetime, timedelta
import wfm_xl
import wfm_helpers
#from openpyxl import load_workbook
import uuid

class Wfm_db:
    def __init__ (self, env:dict, cnf:dict, logger):
        """
        Initialises the Wfm_db class.
        
        Parameters
        ----------
        env: dictionary - contains the configuration parameters for WFMS
        cnf: dictionary - contains filenames, dictionary for resource strings and forecast framework dictionary
        """
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
            self.__xl_obj   = ''
            if self.__env['data-io']['source-from']['excel'] and self.__env['data-io']['import-to']['db']:
                if self.__env['excel']['Simple']:
                    self.__xl_obj = wfm_xl.wfm_xl (self.__env, cnf, self.__lgr, xlfile=cnf['f_imp_simple'])
                else:
                    self.__xl_obj = wfm_xl.wfm_xl (self.__env, cnf, self.__lgr, xlfile=cnf['f_imp_complex'])
            else:
# TODO: add staright forecast here - excel to excel
                pass

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

    def dbCheckConnectionStatus (self, db, pingonly=False):
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
        if pingonly:
            return db.command ( { "ping": 1 } )
        else:
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

        if self.__lgr.level == logging.DEBUG:
            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0001'], url, port, self.__env['db']['connectTimeoutMS'], self.__env['db']['socketTimeoutMS']))
        else:
            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0008'], url, port))

        self.__client = MongoClient(url, port, connectTimeoutMS=self.__env['db']['connectTimeoutMS'], socketTimeoutMS=self.__env['db']['socketTimeoutMS'])

        if self.__client is None:
            raise Exception (str.format (self.__dic['res_strings']['db']['0003'], self.__env['db']['url'], self.__env['db']['port']))

        self.__lgr.info (str.format(self.__dic['res_strings']['db']['0006'], url, port))
        return self.__client


    def dbImportXLSimple (self):
        """
        Imports an Excel spreadsheet to the Mongo database.
        
        Parameters
        ----------
        None

        Returns
        ---------
        None
        """
        if not self.__env['excel']['Simple']:
            self.__lgr.info (self.__dic['res_strings']['db']['0017'])
            return
        try:
            client = self.dbConnect (self.__env['db']['url'], self.__env['db']['port'])
            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0007'], self.__env['db']['db-name']))

            db = client [str (self.__env['db']['db-name'])]

            if self.dbCheckConnectionStatus (db, pingonly=False) != 1:
                raise mdb_errors.ConnectionFailure (str.format (self.__dic['res_strings']['db']['0005'], self.__env['db']['url'], self.__env['db']['port']))

            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0002'], self.__env['db']['db-name']))
            
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
                            "times":    [],             # Interval Times
                            "agents":   [],             # Number of agents
                            "tx_all":   [],             # Number of offered transactions
                            "tx_ans":   [],             # Number of handeled transactions
                            "tx_abn":   [],             # Number of abandoned tarnsactions
                            "r_abn":    [],             # Abandonment Rate
                            "sl":       [],             # Service Level
                            "a_sa":     [],             # Average Speed to Answer
                            "a_wait":   [],             # Average Wait Time
                            "q_rate":   [],             # Average Queue Rate
                            "q_time":   [],             # Average Incident Queue Time
                            "itt":      [],             # Total Incident Talk Time
                            "aiw":      [],             # Total After Incident Wrap Time 
                            "iht":      [],             # Total incident Handle Time
                            "r_util":   [],             # Agents Utilisation
                            "txn_sum": {
                                "year":     "$int",     # The year
                                "month":    "$int",     # Month 1 - 12
                                "day":      "$int",     # Day 1 - 31
                                "dow":      "$string",  # Day of Week
                                "s_mnth":   "$string",  # Month string
                                "sum_iat":  "$int",     # Total Active Time for this day
                                "sum_aiw":  "$int",     # Total After Incident Wrap Time for this day
                                "sum_ht":   "$int",     # Total Handle Time (iat_sum + aiw_sum) for this day
                                "avg_ht":   "$int",     # Average Handle Time for this day
                                "avg_aiw":  "$int",     # Average After Call Wrap Time for this day
                                "avg_iat":  "$int",     # Average Incident Active Time for this day
                                "sum_txn":   "$int",    # Total Tarnsactions for this day
                                "sum_ans":  "$int",     # Total Transactions answered for this day
                                "sum_abn":  "$int",     # Total Abandoned Tarnsactions for this day
                                "r_abn":    "$float",   # Abandonment Rate for tis day
                                "r_avail":  "$float",   # Agents Availability
                                "r_util":   "$float"    # Utilisation for this day
                            }
                    }]}

                # Fill the record dictionary to be processed
                self.__xl_obj.xlImportTxnsOnly (record, c)
                # Simple means we only are importing transaction data, times and dates - nothing else
                # it means that data is organised:
                #       Date,   Date,   Date,   ...
                # Time  txns,   txns,   txns,   ...
                # Time  txns,   txns,   txns,   ...
                # Time  txns,   txns,   txns,   ...

                # Perform average calculations

                # Check if the last records year is different to this record - in that case need to create new _id
                qry = {"_id": record['_id']}

                # Check if a document with this year exists
                query_res = collection.find_one(qry)

                if query_res == None:
                    collection.insert_one(record)
                    cnt_days += 1
                    cnt_records += len (record['day_ary'][0]['tx_all'])
                else:
                    tmp_date = "{}/{}/{}".format(record['day_ary'][0]['txn_sum']['year'], record['day_ary'][0]['txn_sum']['month'], record['day_ary'][0]['txn_sum']['day'])
                    self.__lgr.warning (str.format(self.__dic['res_strings']['db']['0014'], tmp_date, collection.name))
                    cnt_omit += 1
                    continue
            if cnt_omit > 0:
                self.__lgr.info (str.format(self.__dic['res_strings']['db']['0015'], cnt_omit))
            if cnt_records > 0 or cnt_days > 0:
                self.__lgr.info (str.format(self.__dic['res_strings']['db']['0013'], cnt_days, cnt_records))
            else:
                self.__lgr.info (self.__dic['res_strings']['db']['0016'])

            del record
            del collection
            client.close()
            del client
            del db

        except Exception as e:
            self.__lgr.fatal (e)    
            raise Exception (e)
        
    def dbImportXLComplex (self):
        """
        Imports an Excel spreadsheet to the Mongo database.
        
        Parameters
        ----------
        None

        Returns
        ---------
        None
        """
        if self.__env['excel']['Simple']:
            self.__lgr.info (self.__dic['res_strings']['db']['0018'])
            return

        try:
            client = self.dbConnect (self.__env['db']['url'], self.__env['db']['port'])
            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0007'], self.__env['db']['db-name']))

            db = client [str (self.__env['db']['db-name'])]

            if self.dbCheckConnectionStatus (db, pingonly=False) != 1:
                raise mdb_errors.ConnectionFailure (str.format (self.__dic['res_strings']['db']['0005'], self.__env['db']['url'], self.__env['db']['port']))

            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0002'], self.__env['db']['db-name']))
            
            # ================== Create wfm_obj ==================
            # Select the active sheet
            collection = db [str (self.__env['db']['col-in'])]
            # Process spreadsheet and update db
            # we are moving row by row
            cnt_days = 0
            cnt_omit = 0
            cnt_update = 0
            cnt_replace = 0
            days = self.__xl_obj.getDateArray().copy()

            #TODO: Add mapping of Excel fields to database fields

            # For every day to be imported from spreadsheet do...
            self.__lgr.info (self.__dic['res_strings']['info']['0045'])

            for day in days:
                record = {
                    "_id":0, 
                    "day_ary": [{
                            "times":    [],             # Interval Times
                            "tx_all":   [],             # Number of offered transactions
                            "tx_ans":   [],             # Number of handeled transactions
                            "tx_abn":   [],             # Number of abandoned tarnsactions
                            "r_abn":    [],             # Abandonment Rate
                            "sl":       [],             # Service Level
                            "a_sa":     [],             # Average Speed to Answer
                            "a_wait":   [],             # Average Wait Time
                            "q_rate":   [],             # Average Queue Rate
                            "q_time":   [],             # Average Incident Queue Time
                            "iat":      [],             # Total Incident Active Time
                            "aiw":      [],             # Total After Incident Wrap Time 
                            "iht":      [],             # Total incident Handle Time
                            "agents":   [],             # Number of agents
                            "ag_lit":   [],             # Agent login time summed for this time segment
                            "ag_rt":    [],             # Agent rostered time
                            "r_util":   [],             # Agent's Utilisation
                            "r_avail":  [],             # Agent's Availability
                            "txn_sum": {
                                "year":     "$int",     # The year
                                "month":    "$int",     # Month 1 - 12
                                "day":      "$int",     # Day 1 - 31
                                "dow":      "$string",  # Day of Week
                                "s_mnth":   "$string",  # Month string
                                "sum_lit":  "$int",     # Time agents are logged in
                                "sum_rt":   "$int",     # Time agents are rostered
                                "sum_at":   "$int",     # Total Active Time for this day
                                "sum_aiw":  "$int",     # Total After Incident Wrap Time for this day
                                "sum_ht":   "$int",     # Total Handle Time (iat_sum + aiw_sum) for this day
                                "sum_txn":   "$int",    # Total Tarnsactions for this day
                                "sum_ans":  "$int",     # Total Transactions answered for this day
                                "sum_abn":  "$int",     # Total Abandoned Tarnsactions for this day
                                "avg_ht":   "$int",     # Average Handle Time for this day
                                "avg_aiw":  "$int",     # Average After Call Wrap Time for this day
                                "avg_at":   "$int",     # Average Incident Active Time for this day
                                "avg_ag":   "$float",   # Average Agent Count over this day
                                "r_abn":    "$float",   # Abandonment Rate for tis day
                                "r_avail":  "$float",   # Agents Availability
                                "r_util":   "$float"    # Utilisation for this day
                            }
                    }]}

                # Fill the record dictionary to be processed
                self.__xl_obj.xlImportTxnsKPIs (record, date=day)
                # Complex means we are importing all data specified in framework.conf.excel_map
                # it also means, that the data is organised:
                # date, time, metric 1, metric 2, ... 
                # date, time, metric 1, metric 2, ... 
                # date, time, metric 1, metric 2, ... 

                # Perform average calculations

                # Check if the last records year is different to this record - in that case need to create new _id
                qry = {"_id": record['_id']}

                # Check if a document with this year exists
                query_res = collection.find_one(qry)

                if query_res == None:
                    collection.insert_one(record)
                    cnt_days += 1
                else:
                    if not self.__env['db']['updateRecords'] and not self.__env['db']['replaceRecords']:
                        tmp_date = "{}/{}/{}".format(record['day_ary'][0]['txn_sum']['year'], record['day_ary'][0]['txn_sum']['month'], record['day_ary'][0]['txn_sum']['day'])
                        self.__lgr.warning (self.__dic['res_strings']['db']['0019'])
                        self.__lgr.warning (str.format(self.__dic['res_strings']['db']['0014'], tmp_date, collection.name))
                        cnt_omit += 1
                        continue
                    elif self.__env['db']['updateRecords']:
                        collection.replace_one (qry, record, upsert=True)
                        cnt_days += 1
                        cnt_update += 1
                        continue
                    else:
                        collection.replace_one (qry, record)
                        cnt_days += 1
                        cnt_replace += 1
                        continue
            if cnt_days > 0 or cnt_replace > 0 or cnt_update > 0:
                self.__lgr.info (str.format(self.__dic['res_strings']['db']['0013'], cnt_days, cnt_omit, cnt_update, cnt_replace))
            else:
                self.__lgr.info (self.__dic['res_strings']['db']['0016'])

            del record
            del collection
            client.close()
            del client
            del db

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
            wfm_helpers.helperCheckFrameworkData (self.__dic, self.__env, self.__lgr)
            return
        except Exception as ex:
            raise (ex)

    def dbCreateForecast (self):
        # record = {"_id": '', "date": '', "sdate":"","smonth":"" , "dow":"", "times": [], "calls": [], "util": [], "sla": [], "asa": [], "att": [], "aiw": [], "abandon": [], "qpercent": [], "qtime": [], "qsize": []}
        pass


    def __checkSyntax (self, record):
# TODO: add syntax check rules to make sure the data imported makes sense
        return True

# ================== Upload Forecast Framework to MongoDB -- THREAD SAFE -- ==================
    def dbUploadFramework (self):
        """
        Uploads the forecast Framework defined in Excel or json file to the Mongo database.
        If the framework already exists, it will be updated.
        """
        try:
            client = self.dbConnect (self.__env['db']['url'], self.__env['db']['port'])
            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0002'], self.__env['db']['db-name']))

            db = client [str (self.__env['db']['db-name'])]

            if self.dbCheckConnectionStatus (db, pingonly=True) != 1:
                raise mdb_errors.ConnectionFailure (str.format (self.__dic['res_strings']['db']['0005'], self.__env['db']['url'], self.__env['db']['port']))

            self.__lgr.info (str.format(self.__dic['res_strings']['db']['0002'], self.__env['db']['db-name']))



        except Exception as e:
            raise e
