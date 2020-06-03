# wfm_xl.py
# Copyright 2020 by Peter Gossler. All rights reserved.

import logging
from openpyxl import load_workbook, Workbook
from datetime import datetime, timedelta
#import threading

class wfm_xl:
#  ================== Initialise class ==================
    def __init__ (self, env, dic, logger):
        try:
            self.__lgr      = logger
            self.__env      = env
            self.__dic      = dic
            self.__wb       = Workbook()
            self.__ws       = Workbook.worksheets
            self.__cnt_iv   = 0
            self.__max_days = 0
            self.__max_rows = 0
            self.__max_cols = 0
            self.__svc_time = 0
            self.__svc_intv = 0
            self.__strt_tm  = datetime.now()
            self.__isReady  = False
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
            del self.__dic
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

# ================== Pre-process Excel Source Data -- NOT THREAD SAFE ==================
    def xlPreProcessSource (self):
        try:
            # load the source forecast data from xl
            self.__wb = load_workbook (self.__dic['f_xl_data'], data_only=True)
            self.__lgr.debug (f"Excel workbook {self.__dic['f_xl_data']} opened.")
            # Select the active sheet
            self.__ws = self.__wb[self.__env['excel']['SD_name']]
        except:
            # Something went wrong
            raise Exception (str.format(self.__dic['res_strings']['errors']['0017'], self.__dic['f_xl_data']))
        # Determine the row and column count
        self.__max_rows = self.__ws.max_row
        self.__max_cols = self.__ws.max_column
        # Determine the Service Interval from xl sheet
        act_row = self.__ws.min_row + 1
        act_col = self.__ws.min_column
        x = 0
        diff_prev = timedelta(0)
        try:
            # Iterate through all rows
            while x < self.__ws.max_row - 2:
                if (self.__ws.cell(row=act_row, column=act_col).value is not None):
                    if (self.__ws.cell(row=act_row, column=act_col).is_date):
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
                            raise ValueError (self.__dic['res_strings']['errors']['0005'].format(self.__dic['f_frame']))
                        diff_prev = diff
                        act_row += 1
                        x += 1
                    else:
                        raise ValueError (str.format(self.__dic['res_strings']['errors']['0014'], act_row, act_col, self.__ws.cell(row=act_row, column=act_col).number_format))
                else:
                    raise ValueError (str.format(self.__dic['res_strings']['errors']['0015'], act_row, act_col))
            # Set start time and service interval
            self.__strt_tm = self.__ws.cell(row=self.__ws.min_row + 1, column=act_col).value
            # Set start time and service interval
            self.__SrvcTime = diff.total_seconds()/60
            self.__isReady = True
        except Exception as ex:
            raise (ex)

# ================== Construct the forecast dictionary  -- NOT THREAD SAFE -- ==================
    def xlCreateDictionary (self, dic_fc):
        try:
            if (self.__isReady):
                # Create the array for the various times - the size depends on the forecast interval 15, 30, 45, 60 minutes and the operations hours
                # The times array is common to all other entries in the folowing arrays, so don't need to duplicate that for each Dayx
                dic_fc['times'] = []
                # Use start time from before
                timeobj = self.__strt_tm
                # ================== Construct the forecast dictionary from Excel data ==================
                self.__lgr.debug (self.__dic['res_strings']['info']['0007'])
                # First creating the Interval times array
                for c in range (0, self.__cnt_iv, 1):
                    t2 = self.__ws.cell(row=self.__ws.min_row + c + 1, column=self.__ws.min_column).value
                    str_t2 = str.format('{}:{}', t2.hour, t2.minute)
                    timeobj = datetime.strptime(str_t2, self.__dic['sf_time'])
                    dic_fc['times'].append(timeobj.strftime(self.__dic['sf_time']))
                self.__lgr.debug (str.format(self.__dic['res_strings']['info']['0002'], len(dic_fc['times'])))
                # Create the rest of the arrays, all but Calls will be empty for now
                for i in range (0, self.__max_days, 1): # restrict to maximum forecast days as calculated before
                    # Read date from excel spread sheet and put into array
                    row_num = self.__ws.min_row
                    col_num = self.__ws.min_column + i + 1
                    dateobj = self.__ws.cell(row=row_num, column=col_num).value
                    self.__lgr.debug(str.format(self.__dic['res_strings']['info']['0001'], dateobj))
                    # Format the dictionary key - Set0, Set1, Set2, ...
                    str_tmp = 'Day' + str(i) 
                    dic_fc[str_tmp] = {}
                    dic_fc[str_tmp]['count'] = self.__max_days
                    dic_fc[str_tmp]['date'] = dateobj.strftime(self.__dic['sf_date'])
                    # Insert Call Numbers for this date
                    dic_fc[str_tmp]['calls'] = []
                    for c in range(0, self.__cnt_iv):
                        r = c + 2
                        val = self.__ws.cell(row=r, column=col_num).value
                        dic_fc[str_tmp]['calls'].append(int(val))
                    dic_fc[str_tmp]['agents'] = []
                    if (self.__env['excel']['report-detail']):
                        dic_fc[str_tmp]['util']      = []
                        dic_fc[str_tmp]['sla']       = []
                        dic_fc[str_tmp]['asa']       = []
                        dic_fc[str_tmp]['abandon']   = []
                        dic_fc[str_tmp]['q-percent'] = []
                        dic_fc[str_tmp]['q-time']    = []
                        dic_fc[str_tmp]['q-count']   = []

                self.__lgr.debug (self.__dic['res_strings']['info']['0008'])
                return
            else:
                raise RuntimeError (str.format(self.__dic['res_strings']['errors']['0016']))
        except Exception as ex:
            raise (ex)

# ================== Sanity check for data read - compare parts of framework file with actual data -- NOT THREAD SAFE -- ==================
    def xlCheckImportDataValidity (self):
        try:
            if (self.__isReady):
                # Check if Service Interval is 15, 30, 45 or 60 minutes
                if (self.__SrvcTime) not in self.__dic['fw']['Intervals']:
                    raise ValueError (self.__dic['res_strings']['errors']['0001'].format(self.__SrvcTime, self.__dic['f_xl_data']))
                # Check if Service Interval from framework file is the same as in excel file
                if (self.__SrvcTime != self.__dic['fw']['ServiceInterval'] or self.__dic['fw']['ServiceInterval'] <= 0):
                    raise ValueError ((self.__dic['res_strings']['errors']['0002'].format(self.__dic['fw']['ServiceInterval'], self.__dic['f_frame'], int(self.__SrvcTime), self.__dic['f_xl_data'])))
                if (self.__dic['fw']['OperationHours'] <= 0):
                    raise ValueError (str.format(self.__dic['res_strings']['errors']['0006'], self.__dic['fw']['OperationHours'], self.__dic['f_frame']))

                # Calculate maximum Number of Intervals per calculation loop
                self.__cnt_iv = int(self.__dic['fw']['OperationHours'] * (60 / self.__dic['fw']['ServiceInterval']))
                # Check if framework interval is different to what is in the excel forcast sheet
                if (self.__cnt_iv != self.__max_rows - 1):
                    self.__lgr.warning (self.__dic['res_strings']['errors']['0004'])
                    self.__cnt_iv = int(self.__max_rows - 1)
                # Check forecast framework forcast days against spreadsheet
                if ((self.__max_cols - 1) != self.__dic['fw']['ForecastDays']):
                    # There are less forecast columns than forecast days specified in the framework file, using the columns available
                    if ((self.__max_cols - 1) < self.__dic['fw']['ForecastDays']):
                        self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0001'], self.__max_cols - 1, self.__dic['f_xl_data'], self.__dic['fw']['ForecastDays'], self.__dic['f_frame'], self.__max_cols - 1))
                        self.__max_days = self.__max_cols - 1
                    # There are more columns in the forecast file than specified in the framework file - using framework file value
                    if ((self.__max_cols - 1) > self.__dic['fw']['ForecastDays']):
                        self.__lgr.warning (str.format(self.__dic['res_strings']['warnings']['0001'], self.__max_cols - 1, self.__dic['f_xl_data'], self.__dic['fw']['ForecastDays'], self.__dic['f_frame'], self.__dic['fw']['ForecastDays']))
                        self.__max_days = self.__dic['fw']['ForecastDays']
                else:
                    self.__max_days = self.__max_cols - 1
                self.__lgr.debug (str.format(self.__dic['res_strings']['info']['0005'], self.__max_days))
                self.__lgr.debug (str.format(self.__dic['res_strings']['info']['0006'], self.__cnt_iv))
                return
            else:
                raise RuntimeError (str.format(self.__dic['res_strings']['errors']['0016']))
        except Exception as ex:
            raise (ex)

# ================== Create the Excel Summary sheet -- THREAD SAFE -- ==================
    def xlCreateReport (self, fc): 
        """
        A function that creates an Excel spreadsheet based on parameters provided
        during class initialisation.

        Parameters
        ----------
        xlCreateReport(dictionary fc)
            Generates the excel report based on forecast data fc.
        """
        try:
            if (self.__isReady):
            # Load the forecast output excel file
                self.__wb = Workbook()
                self.__wb.save (self.__dic['f_result'])
                self.__wb = load_workbook (self.__dic['f_result'])
                self.__lgr.info (str.format(self.__dic['res_strings']['info']['0015'], self.__dic['f_result']))
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
                    if (self.__env['export']['report-detail']):
                        dateobj = datetime.strptime(fc[s_day]['date'], self.__dic['sf_date'])
                        sheet_name = dateobj.strftime(self.__env['formats']['xl-detail'])
                        self.__wb.create_sheet (sheet_name)
                    self.__ws = self.__wb[self.__env['excel']['SS_name']]
                    i += 1
                    col += 1
                self.__lgr.debug (self.__dic['res_strings']['info']['0024'])

            # Fill in the agents numbers required 
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
                    for self.__ws in self.__wb:
                        row = 1
                        col = 1
                        # If this is the summary sheet skip it
                        if (self.__ws.title == self.__env['excel']['SS_name']):
                            continue
                        # Fill the heading for each sheet
                        for head in self.__env['export']['headings']:
                            self.__ws.cell(1, col, value=head)
                            col += 1
                        # Fill the times for each sheet
                        times = fc['times']
                        row = 2
                        col = 1
                        for time in times:
                            self.__ws.cell(row, col).value = time
                            row += 1
                        row = 2
                        col = 1
                        for i in range(len(fc) - 1):
                            s_day = 'Day' + str(i)
                            col += 1
                        i += 1
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
