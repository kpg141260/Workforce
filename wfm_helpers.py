# wfm_helpers.py
# Copyright 2020 by Peter Gossler. All rights reserved.

# ================== Imports ==================
import os
import logging
from logging.handlers import RotatingFileHandler
from openpyxl import load_workbook, Workbook
from datetime import datetime
import json
from __version import __version__
from __version import __product__

# ================== Helper function to configure the python logger ==================
def config_logger (env, id):
    truncated = False
    f_log =  os.path.join (os.getcwd(), str(env['paths']['log-path']), str(env['log-conf']['log-file']))
    if not os.path.exists(env['paths']['log-path']):
        os.makedirs(env['paths']['log-path'], mode=0o666, exist_ok=True)
    if os.path.exists(f_log) == False:
        os.chdir (env['paths']['log-path'])
        fp = open (env['log-conf']['log-file'], "w+")
        fp.close()
        os.chdir ('..')
    else:
        # Check file size and truncate if required
        f_size = os.path.getsize(f_log)
        if (f_size > env['log-conf']['max_bytes'] and env['log-conf']['truncate'] and not env['log-conf']['rotate']):
            # truncate the file
            fp = open (f_log, 'w')
            truncated = True
            fp.close()

    # Setup formatter for logging system
    # Enable logging to file
    if (env['log-conf']['log-to-file']):
        logger = logging.getLogger(id)
        logger.setLevel(env['log-conf']['level'])
        if (env['log-conf']['rotate']):
            file_handler = RotatingFileHandler(f_log, maxBytes=env['log-conf']['max_bytes'], backupCount=env['log-conf']['backup_count'])
        else:
            file_handler = logging.FileHandler(f_log)
        logger.addHandler(file_handler)
        if (env['log-conf']['level'] == "DEBUG"):
            formatter = logging.Formatter (fmt=env['log-conf']['format-debug'], datefmt=env['log-conf']['dateformat'])
        else:
            formatter = logging.Formatter (fmt=env['log-conf']['format'], datefmt=env['log-conf']['dateformat'])
        file_handler.setFormatter (formatter)
    else:
        if (env['log-conf']['level'] == "DEBUG"):
            logging.basicConfig(level=env['log-conf']['level'], format=env['log-conf']['format-debug'], datefmt=env['log-conf']['dateformat'])
        else:
            logging.basicConfig(level=env['log-conf']['level'], format=env['log-conf']['format'], datefmt=env['log-conf']['dateformat'])
    # Enable logging to stdout - this works as if log-to-file is true, and log-to-stdout is true add another logger to stdout
    # if log-to-file is false, then enforce stdout logger regardless of log-to-stdout flag
    if (env['log-conf']['log-to-stdout'] or not env['log-conf']['log-to-file']):
        logger = logging.getLogger(id)
        consoleHandler = logging.StreamHandler()
        consoleHandler.setLevel(env['log-conf']['level'])
        logger.addHandler(consoleHandler)
        if (env['log-conf']['level'] == "DEBUG"):
            formatter = logging.Formatter(fmt=env['log-conf']['format-debug'], datefmt=env['log-conf']['dateformat'])
        else:
            formatter = logging.Formatter(fmt=env['log-conf']['format'], datefmt=env['log-conf']['dateformat'])
        consoleHandler.setFormatter(formatter)
    version = '.'.join(str(c) for c in __version__)
    logger.info (str.format("{} version {} loaded.", __product__, version))
    logger.debug ("Log system activated.")
    if (truncated):
        logger.warning (f'Log-file has been truncated! - Previous log-entries lost!')
    del f_log
    return logger

def set_log_color(msg):
    colors = {
        10: "\033[36m{}\033[0m",       # DEBUG
        20: "\033[32m{}\033[0m",       # INFO
        30: "\033[33m{}\033[0m",       # WARNING
        40: "\033[31m{}\033[0m",       # ERROR
        50: "\033[7;31;31m{}\033[0m"   # FATAL/CRITICAL/EXCEPTION
    }
    return colors[int(logging.root.level)].format(msg)

#  ================== Helper function for all the necessary file IO operations ==================
def createFileNames (dict, env, logger):
    # Check if append-date flag is set and construct file name accordingly:
    # forecast.json or
    # forecast-YYYYMMDD-HHMM.json
    if (env['options']['append-date'] == True):
        dict['f_summary'] = os.path.join (str(env['paths']['output-path']), str.format ("{}-{}.{}", env['files']['summary'], datetime.now().strftime(env['formats']['file-ext']), env['files']['ext-json']))
        dict['f_result'] = os.path.join (str(env['paths']['output-path']), str.format ("{}-{}.{}", env['files']['out-data'], datetime.now().strftime(env['formats']['file-ext']), env['files']['ext-xl']))
        dict['fs_nopath'] = os.path.join (str.format ("{}-{}.{}", env['files']['summary'], datetime.now().strftime(env['formats']['file-ext']), env['files']['ext-json']))
        dict['fr_nopath'] = os.path.join (str.format ("{}-{}.{}", env['files']['out-data'], datetime.now().strftime(env['formats']['file-ext']), env['files']['ext-xl']))
    else:
        dict['f_summary'] = os.path.join (str(env['paths']['output-path']), str.format("{}.{}", str(env['files']['summary']), str(env['files']['ext-json'])))
        dict['f_result'] = os.path.join (str(env['paths']['output-path']), str.format("{}.{}", str(env['files']['out-data']), str(env['files']['ext-xl'])))
        dict['fs_nopath'] = os.path.join (str.format("{}.{}", str(env['files']['summary']), str(env['files']['ext-json'])))
        dict['fr_nopath'] = os.path.join (str.format("{}.{}", str(env['files']['out-data']), str(env['files']['ext-xl'])))
    
    f_res = os.path.join (os.getcwd(), env['paths']['rsrc-path'], str.format("{}-{}.{}", env['files']['language-file'], env['encoding']['language'], env['files']['language-type']))

    dict['f_resource']      = f_res
    dict['f_imp_simple']    = os.path.join (env['paths']['import-path'], env['files']['import-simple'])
    dict['f_imp_complex']   = os.path.join (env['paths']['import-path'], env['files']['import-complex'])
    dict['f_template']      = os.path.join (env['paths']['import-path'], env['files']['import-template'])
    dict['f_frame']         = os.path.join (env['paths']['input-path'], env['files']['framework'])
    dict['f_xl_data']       = os.path.join (env['paths']['input-path'], env['files']['in-data'])
    dict['f_xl_full']       = os.path.join (os.getcwd(), env['paths']['input-path'], env['files']['in-data'])

# Check if forecast resource file exists and load it.
    if os.path.exists(f_res) == False:
        logger.error(f"Can't find input file {f_res}. Execution cannot continue!")
        raise FileNotFoundError (f_res)
    else:
        with open (f_res) as json_file:
            dict['res_strings'] = json.load(json_file)
            logger.info (str.format (dict['res_strings']['info']['0015'], f_res))

# Load the forecast framework for the forecast from json file
    with open (dict['f_frame']) as json_file:
        dict['fw'] = json.load(json_file)
        # Prepare format strings
        dict['sf_time'] = env['formats']['time']
        dict['sf_date'] = env['formats']['date']
        logger.info (str.format (dict['res_strings']['info']['0015'], dict['f_frame']))

# Check that all directories are present, if not create them
    for path in env['paths']:
        if not os.path.exists (env['paths'][path]):
            os.makedirs(env['paths'][path], mode=0o766, exist_ok=True)
            logger.info (str.format(dict['res_strings']['info']['0016'], env['paths'][path]))
        else:
            logger.debug (str.format(dict['res_strings']['info']['0017'], env['paths'][path]))

# Check if forecast input xlsx file exists - if not throw an error
    if os.path.exists(dict['f_xl_data']) == False:
        raise FileNotFoundError (dict['res_strings']['errors']['0003'].format(dict['f_xl_data']))
    else:
        logger.info (str.format(dict['res_strings']['info']['0018'], dict['f_xl_data']))

# Check if output files exist and if not create them
    if os.path.exists(dict['f_result']) == False:
        os.chdir (env['paths']['output-path'])
        fp = open (dict['fr_nopath'], "w+")
        fp.close()
        os.chdir('..')
        logger.info (str.format(dict['res_strings']['info']['0019'], dict['f_result']))
    else:
        logger.info (str.format(dict['res_strings']['info']['0018'], dict['f_result']))

    if os.path.exists(dict['f_summary']) == False:
        os.chdir (env['paths']['output-path'])
        fp = open (dict['fs_nopath'], "w")
        fp.close()
        os.chdir('..')
        logger.info (str.format(dict['res_strings']['info']['0019'], dict['f_summary']))
    else:
        logger.info (str.format(dict['res_strings']['info']['0018'], dict['f_summary']))

    return

# ================== Helper function to output forecast framework to log ==================
def logForecastFramework (dict, logger):
    for fw_item in dict['fw']:
        logger.info (str.format("{}: {}", fw_item, dict['fw'][fw_item]))
    return

# ================== Helper function to create the JSON summary file ==================
def createJSONSummary (agents, max_date, max_time, max_calls, ec, dict, logger):
    logger.info(dict['res_strings']['info']['0013'].format(dict['f_summary']))
    dict_summary = {}
    dict_summary ['Date'] = max_date
    dict_summary ['Time'] = max_time
    dict_summary ['Max Calls'] = max_calls
    dict_summary ['Max Agents'] = agents
    dict_summary ['Utilisation'] = round((ec.Utilisation(agents, max_calls) * 100), 2)
    dict_summary ['Max Trunks required'] = ec.Trunks(agents, max_calls)
    dict_summary ['SLA'] = round ((ec.SLA(agents, max_calls, dict['fw']['ServiceTime']) * 100), 2)
    dict_summary ['ASA'] = ec.ASA(agents, max_calls)
    dict_summary ['Abandoned'] = round ((ec.Abandon(agents, max_calls) * 100), 2)
    dict_summary ['Queued Percent'] = round ((ec.Queued(agents, max_calls) * 100), 2)
    dict_summary ['Queue Time'] = ec.QueueTime(agents, max_calls)
    dict_summary ['Queue Size'] = ec.QueueSize(agents, max_calls)

    summary = json.dumps (dict_summary)
    with open (dict['f_summary'], 'w') as f:
        json.dump (summary, f)
        logger.info(dict['res_strings']['info']['0011'].format(dict['f_summary']))
    # Delete object dict_summary
    del dict_summary
    return

# ================== Helper function to create log entries of the summary data ==================
def createLogSummary (agents, max_date, max_time, max_calls, ec, dict, logger):
    logger.debug (dict['res_strings']['prompts']['0001'])
    logger.debug (dict['res_strings']['prompts']['0002'].format(max_date, max_time, max_calls))
    logger.debug (dict['res_strings']['prompts']['0003'].format(max_calls))
    logger.debug (dict['res_strings']['prompts']['0004'].format(agents))
    logger.debug (dict['res_strings']['prompts']['0005'].format(ec.Utilisation(agents, max_calls) * 100))
    logger.debug (dict['res_strings']['prompts']['0006'].format(ec.Trunks(agents, max_calls)))
    logger.debug (dict['res_strings']['prompts']['0007'].format(ec.SLA(agents, max_calls, dict['fw']['ServiceTime']) * 100))
    logger.debug (dict['res_strings']['prompts']['0008'].format(ec.ASA(agents, max_calls)))
    logger.debug (dict['res_strings']['prompts']['0009'].format(ec.Abandon(agents, max_calls) * 100))
    logger.debug (dict['res_strings']['prompts']['0010'].format(ec.Queued(agents, max_calls) * 100))
    logger.debug (dict['res_strings']['prompts']['0011'].format(ec.QueueTime(agents, max_calls)))
    logger.debug (dict['res_strings']['prompts']['0012'].format(ec.QueueSize(agents, max_calls)))
    return

#  ================== Helper Function to check framework data for any errors ==================
def helperCheckFrameworkData (dic_cnf, env, logger):
    logger.info(dic_cnf['res_strings']['info']['0031'])
    # Check that we are using sound values
    if dic_cnf['fw']['SLA'] < env['minmax']['sla_min'] or dic_cnf['fw']['SLA'] > 1.0:
        err = str.format(dic_cnf['res_strings']['errors']['0015'], env['minmax']['sla_min'], 1.0, dic_cnf['fw']['SLA'], dic_cnf['f_frame'], dic_cnf['f_xl_data'])
        logger.error (err)
        raise ValueError (err)
    if (dic_cnf['fw']['AnswerTime'] <= 0):
        err = str.format(dic_cnf['res_strings']['errors']['0008'], dic_cnf['fw']['AnswerTime'], dic_cnf['f_frame'])
        logger.error (err)
        raise ValueError (err)
    if (dic_cnf['fw']['TalkTime'] <= 0):
        err = str.format(dic_cnf['res_strings']['errors']['0009'], dic_cnf['fw']['TalkTime'], dic_cnf['f_frame'])
        logger.error (err)
        raise ValueError (err)
    if (dic_cnf['fw']['ServiceTime'] <= 0):
        err = str.format(dic_cnf['res_strings']['errors']['0010'], dic_cnf['fw']['ServiceTime'], dic_cnf['f_frame'])
        logger.error (err)
        raise ValueError (err)
    if (dic_cnf['fw']['AfterCallWork'] <= 0):
        err = str.format(dic_cnf['res_strings']['errors']['0011'], dic_cnf['fw']['AfterCallWork'], dic_cnf['f_frame'])
        logger.error (err)
        raise ValueError (err)
    if (dic_cnf['fw']['AbandonTime'] <= 0):
        err = str.format(dic_cnf['res_strings']['errors']['0012'], dic_cnf['fw']['AbandonTime'], dic_cnf['f_frame'])
        logger.error (err)
        raise ValueError (err)
    if (dic_cnf['fw']['MaxWait'] <= 0):
        err = str.format(dic_cnf['res_strings']['errors']['0013'], dic_cnf['fw']['MaxWait'], dic_cnf['f_frame'])
        logger.error (err)
        raise ValueError (err)
    if dic_cnf['fw']['Availability'] < env['minmax']['avy_min'] or dic_cnf['fw']['Availability'] > 1.0:
        err = str.format(dic_cnf['res_strings']['errors']['0018'], env['minmax']['avy_min'], 1.0, dic_cnf['fw']['Availability'], dic_cnf['f_frame'], dic_cnf['f_xl_data'])
        logger.error (err)
        raise ValueError (err)
    if dic_cnf['fw']['UtilLimit'] < env['minmax']['utl_min'] or dic_cnf['fw']['UtilLimit'] > 1.0:
        err = str.format(dic_cnf['res_strings']['errors']['0019'], env['minmax']['utl_min'], 1.0, dic_cnf['fw']['UtilLimit'], dic_cnf['f_frame'], dic_cnf['f_xl_data'])
        logger.error (err)
        raise ValueError (err)

    logger.info(dic_cnf['res_strings']['info']['0032'])

# ================== Helper function to get exit string like good morning, good afternoon, etc. ==================
def getExitString (dic_cnf):
    try:
        str_tmp = ""
        tod = datetime.now().time()
        if (tod.hour >= 6 and tod.hour < 10):
            str_tmp = dic_cnf['res_strings']['prompts']['0013']
        if (tod.hour >= 10 and tod.hour < 12):
            str_tmp = dic_cnf['res_strings']['prompts']['0014']
        if (tod.hour >= 12 and tod.hour < 18):
            str_tmp = dic_cnf['res_strings']['prompts']['0015']
        if (tod.hour >= 18 and tod.hour < 20):
            str_tmp = dic_cnf['res_strings']['prompts']['0016']
        if (tod.hour >= 20 and tod.hour < 6):
            str_tmp = dic_cnf['res_strings']['prompts']['0017']
        return str_tmp
    except Exception as e:
        raise (e)

def logTimeStamp (start_time, msg, logger, force=False):
    dif_time = datetime.now() - start_time
    if force:
        logger.info (msg.format (round (dif_time.total_seconds(), 3)))
    else:
        logger.debug (msg.format (round (dif_time.total_seconds(), 3)))



