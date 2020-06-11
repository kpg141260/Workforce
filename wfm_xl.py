# wfm_xl.py
# Copyright 2020 by Peter Gossler. All rights reserved.

import logging
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
#import threading

class wfm_xl:
#  ================== Initialise class ==================
    def __init__ (self, env:dict, cnf:dict, logger):
        try:
            self.__lgr      = logger
            self.__env      = env.copy()
            self.__dic      = cnf.copy()
            self.__dic_fc   = {}
            self.__wb       = Workbook()
            self.__ws       = Workbook.worksheets
            self.__cnt_iv   = 0
            self.__max_days = 0
            self.__max_rows = 0
            self.__max_row  = 0
            self.__max_cols = 0
            self.__max_col  = 0
            self.__max_sht  = 0
            self.__svc_time = 0
            self.__svc_intv = 0
            self.__strt_tm  = datetime.now()
            self.__strt_dt  = datetime.now()
            self.__isReady  = False
            self.__hasDict  = False
            self.__isSource = False
            # Pre-process the excel data source - get max rows, max columns, start time and Service Interval
            self.xlPreProcessSource ()
            # preprocess = threading.Thread(target=self.xlPreProcessSource, daemon=True)
        except Exception as e:
            self.__isReady = False
            raise Exception (e)
            
# ================== General Methods ==================
    def __del__ (self):
        if (self.__isReady):
            self.__wb.close()
            del self.__ws
            del self.__wb
            del self.__strt_tm
            del self.__strt_dt
            del self.__dic
            del self.__dic_fc
            del self.__env
            del self.__lgr
    
    def __str__ (self):
        self.__wb.worksheets.count
        return f"Class: wfm_xl, Current sheet: {self.__ws.sheet_name}, Total sheets: {self.__wb.worksheets.count}"

    def max_days (self):
        return self.__max_days

    def interval_count (self):
        return self.__cnt_iv

    def max_rows (self):
        return self.__max_rows

    def max_row (self):
        return self.__max_row

    def max_col (self):
        return self.__max_col

    def max_columns (self):
        return self.__max_cols

    def ServiceInterval (self):
        return self.__svc_intv
    
    def ServiceTime (self):
        return self.__svc_time
    
    def getWorksheet (self):
        return self.__ws

    def getWorkbook (self):
        return self.__wb

    def getFcDictionary (self):
        return self.__dic_fc

    def getStartDate (self):
        return self.__strt_dt
    
    def setStartDate (self, date:datetime):
        self.__strt_dt = date

# ================== Pre-process Excel Source Data -- NOT THREAD SAFE ==================
    def xlPreProcessSource (self):
        try:
            # load the source forecast data from xl
            self.__lgr.info (str.format (self.__dic['res_strings']['info']['0038'], self.__dic['f_xl_full']))
            self.__wb = load_workbook (self.__dic['f_xl_full'], data_only=True)
            self.__lgr.debug (f"Excel workbook {self.__dic['f_xl_full']} opened.")
            # Select the active sheet
            self.__ws = self.__wb[self.__env['excel']['SD_name']]
        except:
            # Something went wrong
            raise Exception (str.format(self.__dic['res_strings']['errors']['0017'], self.__dic['f_xl_full']))
        # Determine the row and column count
        self.__max_rows = self.__ws.max_row
        self.__max_cols = self.__ws.max_column
        # Determine the Service Interval from xl sheet
        try:
            # Check that all columns contain the valid format and are not blank
            for c in range (2, self.__max_cols, 1):
                # Seems to be more logical for log-file to check for empty cells first, so do not change the order of the two if statements
                if (self.__ws.cell(1, c).value is None):
                    self.__ws.cell(1, c).fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
                    self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0005'], 1, c, c - 2))
                    break
                if (not self.__ws.cell(1, c).is_date):
                    self.__ws.cell(1, c).fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
                    self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0006'], 1, c, type(self.__ws.cell(1, c).value), c - 2))
                    break
            # Get the start date of the Excel source data
            self.__strt_dt = self.__ws.cell(1, 2).value
            # assign correct max column value, if any
            self.__max_cols = c - 2 
            self.__max_col  = c
            # Check that the first row contains time data
            for c in range (2, self.__max_rows, 1):
                if (self.__ws.cell(1, c).value is None):
                    self.__ws.cell(1, c).fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
                    self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0007'], c, 1, c - 2))
                    break           
                if (not self.__ws.cell(1, c).is_date):
                    self.__ws.cell(1, c).fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
                    self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0008'], c, 1, type(self.__ws.cell(1, c).value), c - 2))
                    break
            # assign correct max row value, if any
            self.__max_rows = c - 2
            self.__max_row  = c

            # Iterate through all rows
            act_row = self.__ws.min_row + 1
            act_col = self.__ws.min_column
            diff_prev = timedelta(0)
            for x in range (0, self.__max_rows - 1, 1):
                t1 = self.__ws.cell(row=act_row, column=act_col).value
                t2 = self.__ws.cell(row=act_row + 1, column=act_col).value
                str_t1 = str.format('{}:{}', t1.hour, t1.minute)
                str_t2 = str.format('{}:{}', t2.hour, t2.minute)
                time1 = datetime.strptime(str_t1, self.__dic['sf_time'])
                time2 = datetime.strptime(str_t2, self.__dic['sf_time'])
                diff = time2 - time1
                # Compensate for roll over of time to next day
                if (diff.days == -1):
                    diff = diff + timedelta(days=1)
                if (diff_prev > timedelta(0) and diff_prev != diff):
                    self.__lgr.error (str.format(self.__dic['res_strings']['errors']['0014'], int(diff), act_row, act_col, int(diff_prev)))
                    raise ValueError (self.__dic['res_strings']['errors']['0005'].format(self.__dic['f_frame']))
                diff_prev = diff
                act_row += 1
            # Set start time
            self.__strt_tm = self.__ws.cell(row=self.__ws.min_row + 1, column=act_col).value
            # Set service interval
            self.__svc_intv = diff.total_seconds()/60
            self.__isReady  = True
            self.__isSource = True
            self.__lgr.debug ("{}: {} min, {}: {} hours.".format(self.__dic['res_strings']['prompts']['0018'], int(self.__svc_intv), self.__dic['res_strings']['prompts']['0019'], self.__strt_tm.strftime('%H:%M')))
            self.__lgr.debug (str.format(self.__dic['res_strings']['debug']['0003'], __name__))
            # Check the imported data for validity
            self.xlCheckImportDataValidity()
        except Exception as ex:
            raise (ex)

# ================== Construct the forecast dictionary  -- NOT THREAD SAFE -- ==================
    def xlCreateDictionary (self):
        """
        Function creates the dictionary used for forecasting the agent numbers.
        
        Parameters
        ----------
        dic_fc : dictionary dict {}
            Generates the excel reports based on forecast data provided in fc.
            The report parameters are defined in file workforce.conf
        """

        try:
            if (self.__isReady):
                # Create the array for the various times - the size depends on the forecast interval 15, 30, 45, 60 minutes and the operations hours
                # The times array is common to all other entries in the folowing arrays, so don't need to duplicate that for each Dayx
                self.__dic_fc['times'] = []
                # Use start time from before
                timeobj = self.__strt_tm
                # ================== Construct the forecast dictionary from Excel data ==================
                self.__lgr.debug (self.__dic['res_strings']['info']['0007'])
                # First creating the Interval times array

                for c in range (0, self.__cnt_iv, 1):
                    t2 = self.__ws.cell(row=self.__ws.min_row + c + 1, column=self.__ws.min_column).value
                    str_t2 = str.format('{}:{}', t2.hour, t2.minute)
                    timeobj = datetime.strptime(str_t2, self.__dic['sf_time'])
                    self.__dic_fc['times'].append(timeobj.strftime(self.__dic['sf_time']))
                self.__lgr.debug (str.format(self.__dic['res_strings']['info']['0002'], len(self.__dic_fc['times'])))
                # Create the rest of the arrays, all but Calls will be empty for now
                if (self.__env['export']['report-detail']):
                    self.__lgr.debug (self.__dic['res_strings']['debug']['0004'])
                else:
                    self.__lgr.debug (self.__dic['res_strings']['debug']['0005'])
                for i in range (0, self.__max_days, 1): # restrict to maximum forecast days as calculated before
                    # Read date from excel spread sheet and put into array
                    row_num = self.__ws.min_row
                    col_num = self.__ws.min_column + i + 1
                    dateobj = self.__ws.cell(row=row_num, column=col_num).value
                    self.__lgr.debug(str.format(self.__dic['res_strings']['info']['0001'], dateobj.strftime(self.__dic['sf_date'])))
                    # Format the dictionary key - Set0, Set1, Set2, ...
                    str_tmp = 'Day' + str(i) 
                    self.__dic_fc[str_tmp] = {}
                    self.__dic_fc[str_tmp]['count'] = self.__max_days
                    self.__dic_fc[str_tmp]['date'] = dateobj.strftime(self.__dic['sf_date'])
                    # Insert Call Numbers for this date
                    self.__dic_fc[str_tmp]['calls'] = []
                    for c in range(0, self.__cnt_iv):
                        r = c + 2
                        val = self.__ws.cell(row=r, column=col_num).value
                        self.__dic_fc[str_tmp]['calls'].append(int(val))
                    self.__dic_fc[str_tmp]['agents'] = []
                    if (self.__env['export']['report-detail']):
                        self.__dic_fc[str_tmp]['util']      = []
                        self.__dic_fc[str_tmp]['sla']       = []
                        self.__dic_fc[str_tmp]['asa']       = []
                        self.__dic_fc[str_tmp]['abandon']   = []
                        self.__dic_fc[str_tmp]['q-percent'] = []
                        self.__dic_fc[str_tmp]['q-time']    = []
                        self.__dic_fc[str_tmp]['q-count']   = []
                # Signal the the dictionary has been built and copied
                self.__hasDict = True
                self.__lgr.debug (self.__dic['res_strings']['info']['0008'])
                return self.__dic_fc
            else:
                raise RuntimeError (str.format(self.__dic['res_strings']['errors']['0016']))
        except Exception as ex:
            raise (ex)

# ================== Sanity check for data read - compare parts of framework file with actual data -- NOT THREAD SAFE -- ==================
    def xlCheckImportDataValidity (self):
        try:
            if (self.__isReady):
                # Check if Service Interval is 15, 30, 45 or 60 minutes
                if (self.__svc_intv) not in self.__dic['fw']['Intervals']:
                    raise ValueError (self.__dic['res_strings']['errors']['0001'].format(self.__svc_intv, self.__dic['f_xl_data']))
                # Check if Service Interval from framework file is the same as in excel file
                if (self.__svc_intv != self.__dic['fw']['ServiceInterval'] or self.__dic['fw']['ServiceInterval'] <= 0):
                    raise ValueError ((self.__dic['res_strings']['errors']['0002'].format(self.__dic['fw']['ServiceInterval'], self.__dic['f_frame'], int(self.__svc_intv), self.__dic['f_xl_data'])))
                if (self.__dic['fw']['OperationHours'] <= 0):
                    raise ValueError (str.format(self.__dic['res_strings']['errors']['0006'], self.__dic['fw']['OperationHours'], self.__dic['f_frame']))

                # Calculate maximum Number of Intervals per calculation loop
                self.__cnt_iv = int(self.__dic['fw']['OperationHours'] * (60 / self.__svc_intv))
                # Check if framework interval is different to what is in the excel forcast sheet
                if (self.__cnt_iv != self.__max_rows):
                    self.__lgr.warning (self.__dic['res_strings']['errors']['0004'])
                    self.__cnt_iv = int(self.__max_rows)
                # Check forecast framework forcast days against spreadsheet
                if ((self.__max_cols) != self.__dic['fw']['ForecastDays']):
                    # There are less forecast columns than forecast days specified in the framework file, using the columns available
                    if ((self.__max_cols) < self.__dic['fw']['ForecastDays']):
                        self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0001'], self.__max_cols, self.__dic['f_xl_data'], self.__dic['fw']['ForecastDays'], self.__dic['f_frame']))
                        self.__max_days = self.__max_cols
                    # There are more columns in the forecast file than specified in the framework file - using framework file value
                    if ((self.__max_cols) > self.__dic['fw']['ForecastDays']):
                        self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0001'], self.__max_cols, self.__dic['f_xl_data'], self.__dic['fw']['ForecastDays'], self.__dic['f_frame']))
                        self.__max_days = self.__dic['fw']['ForecastDays']
                else:
                    self.__max_days = self.__max_cols
                self.__lgr.info (str.format(self.__dic['res_strings']['info']['0004'], self.__cnt_iv * self.__max_days))
                self.__lgr.info (str.format(self.__dic['res_strings']['info']['0005'], self.__max_days))
                self.__lgr.info (str.format(self.__dic['res_strings']['info']['0006'], self.__cnt_iv))
                return
            else:
                raise RuntimeError (str.format(self.__dic['res_strings']['errors']['0016']))
        except Exception as ex:
            raise (ex)

# ================== Create the Excel report -- THREAD SAFE -- ==================
    def xlCreateReport (self, fc:dict): 
        """
        Function creates agent number forecast based on data provided.
        The forecast is exported to an Excel (.xlsx). Forecast outcome is based on
        parameters provided during class initialisation.

        Parameters
        ----------
        fc : custom dictionary
            Generates the excel reports based on forecast data provided in fc.
            The report parameters are defined in file workforce.conf
        """
        try:
            if (self.__isReady):
            # Load the forecast output excel file
                self.__wb = Workbook()
                self.__wb.save (self.__dic['f_result'])
                self.__wb = load_workbook (self.__dic['f_result'])
                self.__lgr.info (str.format(self.__dic['res_strings']['info']['0015'], self.__dic['f_result']))
                self.__isSource = False                
                # Create the summary worksheet
                # wb.create_sheet (title=env['excel']['summarysheet'])
                # Select the active sheet
                self.__ws = self.__wb.active
                self.__ws.title = self.__env['excel']['SS_name']
                self.__lgr.debug (str.format(self.__dic['res_strings']['info']['0022'], self.__ws.title))
                self.__wb.save (self.__dic['f_result'])
                self.__lgr.info (self.__dic['res_strings']['info']['0028'])

            # Start filling in the details - starting with the times
                cell = self.__ws.cell(1,1) # 'tis cell 'A1'
                cell.value = 'Times'
                times = fc['times']
                row = 2
                col = 1
                for time in times:
                    self.__ws.cell(row, col).value = time
                    row += 1
                self.__lgr.debug (self.__dic['res_strings']['info']['0023'])

            # Next fill the dates in the first row and at the same time create the various detail sheets if required
                s_day = ''
                i = 0
                row = 1
                col = 2
                for i in range(len(fc) - 1):
                    s_day = 'Day' + str(i)
                    self.__ws.cell(row, col).value = fc[s_day]['date']
                    # Only create detail excel sheets and report if configured
                    if (self.__env['export']['report-detail']):
                        dateobj = datetime.strptime(fc[s_day]['date'], self.__dic['sf_date'])
                        sheet_name = dateobj.strftime(self.__env['formats']['xl-detail'])
                        self.__wb.create_sheet (sheet_name)
                        self.__lgr.debug (str.format(self.__dic['res_strings']['debug']['0002'], sheet_name))
                    self.__ws = self.__wb[self.__env['excel']['SS_name']]
                    i += 1
                    col += 1
                self.__lgr.debug (self.__dic['res_strings']['info']['0024'])

            # Fill in the agents numbers into the summary sheetrequired 
                col = 2
                for i in range(len(fc) - 1):
                    row = 2
                    s_day = 'Day' + str(i)
                    agents = fc[s_day]['agents']
                    for agent in agents:
                        self.__ws.cell(row, col).value = agent
                        row += 1
                    i += 1
                    col += 1
                self.__lgr.info(self.__dic['res_strings']['info']['0025'])

            # Create detailed report - if required
                if (self.__env['export']['report-detail']):
                    self.__lgr.info(self.__dic['res_strings']['info']['0026'])
                    self.__max_sht = len(self.__wb.sheetnames) - 1
                    # Iterate through all sheets in workbook 
                    day = 0
                    for self.__ws in self.__wb:
                        row = 1
                        col = 1
                        # Skip the summary sheet
                        if (self.__ws.title == self.__env['excel']['SS_name']):
                            continue
                        # Fill the heading for each sheet
                        for head in self.__env['export']['headings']:
                            self.__ws.cell(1, col, value=head)
                            col += 1
                        # Fill spreadsheet with times and other details
                        times = fc['times']
                        x = 0
                        r = 2
                        s_day = 'Day' + str(day)
                        fc_day = self.__dic_fc[s_day]
                        for time in times:
                            self.__ws.cell(r, column=1).value = time
                            self.__ws.cell(r, column=2).value = fc_day['calls'][x]
                            self.__ws.cell(r, column=3).value = fc_day['agents'][x]
                            self.__ws.cell(r, column=4).value = fc_day['sla'][x]
                            self.__ws.cell(r, column=4).number_format = self.__env['formats']['percent']
                            self.__ws.cell(r, column=5).value = fc_day['asa'][x]
                            self.__ws.cell(r, column=6).value = fc_day['abandon'][x]
                            self.__ws.cell(r, column=6).number_format = self.__env['formats']['percent']
                            self.__ws.cell(r, column=7).value = fc_day['q-percent'][x]
                            self.__ws.cell(r, column=7).number_format = self.__env['formats']['percent']
                            self.__ws.cell(r, column=8).value = fc_day['q-time'][x]
                            self.__ws.cell(r, column=9).value = fc_day['q-count'][x]
                            x += 1
                            r += 1
                        day += 1
                else:
                    self.__lgr.info(self.__dic['res_strings']['info']['0027'])

            # Exit gracefully
                self.__wb.save (self.__dic['f_result'])
                self.__lgr.info(str.format(self.__dic['res_strings']['info']['0011'], self.__dic['f_result']))
                return
            else:
                raise RuntimeError (str.format(self.__dic['res_strings']['errors']['0016']))
        except Exception as ex:
            raise (ex)

# ================== Helper Function to create a single record for export to DB -- NOT THREAD SAFE -- ==================
    def xlFillSingleRecord (self, record:dict, col=0):
        try:
            if (not self.__isReady or not self.__isSource):
                self.xlPreProcessSource ()

            date_cell                       = self.__ws.cell(row=1, column=col).value
            record['_id']                   = int (str (datetime.strftime (date_cell, self.__env['formats']['dbq-date'])))
            record['day_ary'][0]['year']    = int (date_cell.year)
            record['day_ary'][0]['month']   = int (date_cell.month)
            record['day_ary'][0]['day']     = int (date_cell.day)
            record['day_ary'][0]['s_mnth']  = str (datetime.strftime (date_cell, '%B'))
            record['day_ary'][0]['s_date']  = str (datetime.strftime (date_cell, self.__env['formats']['dbq-date']))
            record['day_ary'][0]['dow']     = str (datetime.strftime (date_cell, '%A'))

            # TODO - create the 'xl_db_map dynamically from spreadsheet
            if not self.__env['excel']['Simple']:
                inc = -2 # account for the three keys that always need to be there, otherwise we have nothing to import
                # Calculate the increment based on how many columns of KPIs there are
                tmp_dict = {}
                tmp_dict = self.__env['excel']['xl_db_map']
                for k, v in tmp_dict.items():
                    if v[2]:
                        inc += 1
            for r in range (2, self.__max_row, 1):
                t1 = self.__ws.cell(row=r, column=1).value
                record['day_ary'][0]['times'].append (str.format ('{:02d}:{:02d}', t1.hour, t1.minute))
                record['day_ary'][0]['tx_all'].append (int (round (self.__ws.cell(row=r, column=col).value, 0)))
                # Process additional columns according to 'xl_db_map' from config file
                if not self.__env['excel']['Simple']:
                    if self.__env['excel']['xl_db_map'][' tx_ans'][2]:
                        record['day_ary'][0]['tx_abn'].append (record['day_ary'][0]['tx_all'][0] - int (round (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['tx_ans'][1]).value, 0)))

                    if self.__env['excel']['xl_db_map']['r_abn'][2]:
                        record['day_ary'][0]['r_abn'].append (round (record['day_ary'][0]['r_abn'][0] / record['day_ary'][0]['tx_all'][0], 2))

                    if self.__env['excel']['xl_db_map']['sl'][2]:
                        record['day_ary'][0]['sl'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['sl'][1]))

                    if self.__env['excel']['xl_db_map']['a_sa'][2]:
                        record['day_ary'][0]['a_sa'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['a_sa'][1]))

                    if self.__env['excel']['xl_db_map']['a_wait'][2]:
                        record['day_ary'][0]['a_wait'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['a_wait'][1]))

                    if self.__env['excel']['xl_db_map']['q-rate'][2]:
                        record['day_ary'][0]['q-rate'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['q-rate'][1]))

                    if self.__env['excel']['xl_db_map']['a_it'][2]:
                        record['day_ary'][0]['a_it'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['a_it'][1]))

                    if self.__env['excel']['xl_db_map']['aiw'][2]:
                        record['day_ary'][0]['aiw'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['aiw'][1]))

                    if self.__env['excel']['xl_db_map']['a_ht'][2]:
                        record['day_ary'][0]['a_ht'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['a_ht'][1]))

                    if self.__env['excel']['xl_db_map']['r_avail'][2]:
                        record['day_ary'][0]['r_avail'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['r_avail'][1]))

                    if self.__env['excel']['xl_db_map']['r_util'][2]:
                        record['day_ary'][0]['r_util'].append (self.__ws.cell(row=r, column=col + self.__env['excel']['xl_db_map']['r_util'][1]))

        except Exception as ex:
            raise (ex)