# workforce.py
# Copyright 2020 by Peter Gossler. All rights reserved.
# Version 0.2.0

import logging
import os
import json
from lib.erlang.erlang_base import Erlang_Base
from lib.erlang.erlang_c import Erlang
from datetime import datetime, timedelta
#import threading
#import lib.gui.workforce_gui

import wfm_helpers
import wfm_xl
import wfm_db
from __version import __version__

try:
# >>> Tinkering ===========
    #quickWindow()

# ================== Preparation to get everything going ==================
    # Load configuration file
    start_time = datetime.now()
    f_conf = os.path.join ('conf', 'wfms.conf')
    with open (f_conf) as json_file:
        env = json.load(json_file)
    del f_conf

    # Enable Debug level logging and set logging defaults
    logger_wfm = wfm_helpers.config_logger (env, 'wfm')
    # Create resource dictionary
    dic_cnf = {'f_summary':'', 'f_result':'', 'fs_nopath':'', 'fr_nopath':'', 'f_frame':'', 'f_xl_data':'', 'f_xl_full':'', 'sf_time':'', 'sf_date':'', 'f_resource':'', 'res_strings':{}, 'fw':{}}
    wfm_helpers.createFileNames(dic_cnf, env, logger_wfm)
    # Log the forecast framework
    if (env['log-conf']['log-summary']):
        wfm_helpers.logForecastFramework (dic_cnf, logger_wfm)

# TODO: #2 add some code to pop up a confirmation window here at a later state
    #if (len(wb.sheetnames) > 1):
    #    pass

# ================== Load and process source data - change here if database is source ==================
    # TODO: #3 Add code here to load data from database
    logger_wfm.info (dic_cnf['res_strings']['info']['0033'])
    wfm_helpers.logTimeStamp(start_time, dic_cnf ['res_strings']['info']['0041'], logger_wfm)
    if (env['data-io']['source-from']['db']): # Check if import should be from db
        logger_wfm.info(str.format(dic_cnf['res_strings']['info']['0035'], dic_cnf['res_strings']['prompts']['0021']))
# TODO: Remove after import from db is implemented
    if (env['data-io']['import-to']['db']):
        tgt = "{}:{}.{}.{}".format (env['db']['url'], env['db']['port'], env['db']['db-name'], env['db']['col-in'])
        logger_wfm.info(str.format (dic_cnf['res_strings']['info']['0039'], tgt ))
        db = wfm_db.Wfm_db (env, dic_cnf, logger_wfm)
        db.dbImportXL ()
        wfm_helpers.logTimeStamp(start_time, dic_cnf ['res_strings']['info']['0041'], logger_wfm)
    else:
        logger_wfm.info(dic_cnf['res_strings']['info']['0040'])
    del db

    # ================== Preprocess spreadsheet data - check if times in the sheet make sense ================== 
    xl_obj = wfm_xl.wfm_xl(env, dic_cnf, logger_wfm)
    if (env['data-io']['source-from']['excel']):
        # ================== Build the Dictionary template for the required calculations ==================
        dic_fc = xl_obj.xlCreateDictionary ()
        # ================== Sanity check of framework data supplied ==================
        wfm_helpers.helperCheckFrameworkData (dic_cnf, logger_wfm)
        wfm_helpers.logTimeStamp(start_time, dic_cnf ['res_strings']['info']['0041'], logger_wfm)

# >>>>>>>>>>>>>>>>>> The actual forecast module / function follows below <<<<<<<<<<<<<<<<<<<

# ================== Create Erlang object and start calculations ================== 
    # (SLA, TTA, ATT, ACW, ABNT, MAX_WAIT, NV, CCC, INTERVAL, OPS_HRS) <- this data comes form the forecast framework file
    logger_wfm.info(dic_cnf['res_strings']['info']['0012'])
    ec = Erlang(dic_cnf['fw']['SLA'], dic_cnf['fw']['AnswerTime'], dic_cnf['fw']['TalkTime'], dic_cnf['fw']['AfterCallWork'], dic_cnf['fw']['AbandonTime'], dic_cnf['fw']['MaxWait'], dic_cnf['fw']['NonVoice'], dic_cnf['fw']['Concurrency'], dic_cnf['fw']['ServiceInterval'], dic_cnf['fw']['OperationHours'], avail=dic_cnf['fw']['Availability'])
    max_calls = 0
    for c in range (0, xl_obj.max_days(), 1):
        str_tmp = 'Day' + str(c)
        i = 0
        for each_call in dic_fc[str_tmp]['calls']:
            # find maximum call volume for Trunk Calculation only if 'summary' flag in conf file is true
            if env['options']['summary'] == True:
                if (each_call > max_calls):
                    max_calls = each_call
                    max_date = dic_fc[str_tmp]['date']
                    max_time = dic_fc['times'][i]
            agents = ec.Agents (dic_cnf['fw']['ServiceTime'], each_call)
            dic_fc[str_tmp]['agents'].append(agents)
            dic_fc[str_tmp]['util'].append(round(ec.Utilisation (dic_cnf['fw']['ServiceTime'], each_call), 2))
            dic_fc[str_tmp]['sla'].append(ec.SLA(agents, each_call, dic_cnf['fw']['ServiceTime']))
            dic_fc[str_tmp]['asa'].append(ec.ASA(agents, each_call))
            dic_fc[str_tmp]['abandon'].append(round(ec.Abandon(agents, each_call), 2))
            dic_fc[str_tmp]['q-percent'].append(round(ec.Queued(agents, each_call), 2))
            dic_fc[str_tmp]['q-time'].append(ec.QueueTime(agents, each_call))
            dic_fc[str_tmp]['q-count'].append(ec.QueueSize(agents, each_call))
            i += 1
    wfm_helpers.logTimeStamp(start_time, dic_cnf ['res_strings']['info']['0041'], logger_wfm)
    del each_call
    del str_tmp

# ================== Export the calculated forecast data to JSON ====================
    # Save forcast to JSON file
    if (env['output-format']['json']):
        logger_wfm.info (str.format(dic_cnf['res_strings']['info']['0013'], ['f_result']))
        rec_set = json.dumps(dic_fc)
        with open(dic_cnf['f_result'], 'w') as f:
            json.dump(rec_set, f)
        logger_wfm.info (str.format(dic_cnf['res_strings']['info']['0019'], ['f_result']))
        del rec_set

# ================== Build summary report for max transaction volume =======================
    agents = ec.Agents (dic_cnf['fw']['ServiceTime'], max_calls)
    # Construct Summary Forecast file - only if 'summary' flag is set to 1
    if (env['output-format']['json']):
        logger_wfm.info (str.format(dic_cnf['res_strings']['info']['0013'], dic_cnf['f_summary']))
        wfm_helpers.createJSONSummary (agents, max_date, max_time, max_calls, ec, dic_cnf, logger_wfm)
    else:
        logger_wfm.info(str.format(dic_cnf['res_strings']['info']['0010'], dic_cnf['f_resource']))
    # Log summary data to log file, if summary-log flag is true    
    if (env['log-conf']['log-summary']):
        wfm_helpers.createLogSummary (agents, max_date, max_time, max_calls, ec, dic_cnf, logger_wfm)

# ================== Build Excel Report - Agent count by day by Interval ==================
# TODO: #1 Change export functions here...
    if (env['data-io']['export-to']['excel']):
        xl_obj.xlCreateReport(dic_fc)
        logger_wfm.info(dic_cnf['res_strings']['info']['0014'])
    else:
        logger_wfm.info(dic_cnf['res_strings']['info']['0036'])
    if (env['data-io']['export-to']['db']):
        logger_wfm.info(dic_cnf['res_strings']['info']['0035'])
    else:
        logger_wfm.info(dic_cnf['res_strings']['info']['0037'])
    
# ================== Final Cleanup before exit ==================
    del xl_obj
    del json_file
    del json
    del os
    del ec
    version = '.'.join(str(c) for c in __version__)
    logger_wfm.info(str.format(dic_cnf['res_strings']['info']['0009'], version, wfm_helpers.getExitString(dic_cnf)))
    wfm_helpers.logTimeStamp(start_time, dic_cnf ['res_strings']['info']['0041'], logger_wfm, force=True)
    del env
    del version
    del dic_fc
    del dic_cnf
    
# ================== Error Handling and Exit ===================================
except ValueError as v_err:
    logging.fatal (v_err, exc_info=True)
except FileNotFoundError as fnf_err:
    logging.fatal (fnf_err, exc_info=True)
except PermissionError as p_err:
    logging.error (p_err, exc_info=True)
except OSError as e:
    logging.error (e, exc_info=True)
except Exception as e:
    logging.fatal (e, exc_info=True)
finally:
    logging.shutdown()

