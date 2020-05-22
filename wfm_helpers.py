# wfm_helpers.py
# Copyright 2020 by Peter Gossler. All rights reserved.
# Version 0.1

import os
import logging
from openpyxl import load_workbook, Workbook
from datetime import datetime
import json

# Configure the logger
def config_logger (env, id):
    f_log =  os.path.join (str(env['paths']['log-path']), str(env['files']['log-file']))
    if not os.path.exists(env['paths']['log-path']):
        os.makedirs(env['paths']['log-path'], mode=0o666, exist_ok=True)
    if os.path.exists(f_log) == False:
        os.chdir (env['paths']['log-path'])
        fp = open (env['files']['log-file'], "w+")
        fp.close()
        os.chdir ('..')
    else:
        # Check file size and truncate if required
        f_size = os.path.getsize(f_log)
        if (f_size > env['log-conf']['max-log-size'] and env['log-conf']):
            # truncate the file
            fp = open (f_log, 'w')

    # Make sure that log directory and file exists
    if (env['log-conf']['log-to-file']):
        logging.basicConfig(filename=f_log, level=env['log-conf']['level'], format=env['log-conf']['format'], datefmt=env['log-conf']['dateformat'])
    else:
        logging.basicConfig(level=env['log-conf']['level'], format=env['log-conf']['format'], datefmt=env['log-conf']['dateformat'])
    logger = logging.getLogger(id)
    consoleHandler = logging.StreamHandler()
    consoleHandler.setLevel(env['log-conf']['level'])
    logger.addHandler(consoleHandler)
    formatter = logging.Formatter(fmt=env['log-conf']['format'], datefmt=env['log-conf']['dateformat'])
    consoleHandler.setFormatter(formatter)
    logger.debug ("Log system activated.")
    del f_log
    return logger

# Does all the necessary file IO operations
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
    
    f_res             = os.path.join (env['paths']['config-path'], str.format("{}-{}.{}", env['files']['language-file'], env['encoding']['language'], env['files']['language-type']))
    dict['f_frame']   = os.path.join (env['paths']['config-path'], env['files']['framework'])
    dict['f_xl_data'] = os.path.join (env['paths']['input-path'], env['files']['in-data'])

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

    # Create output directory if it does not already exist
    if not os.path.exists(env['paths']['output-path']):
        os.makedirs(env['paths']['output-path'], mode=0o777, exist_ok=True)
        logger.info (str.format(dict['res_strings']['info']['0016'], env['paths']['output-path']))
    else:
        logger.info (str.format(dict['res_strings']['info']['0017'], env['paths']['output-path']))

        # Check if forecast input xlsx file exists - if not throw an error
    if os.path.exists(dict['f_xl_data']) == False:
        raise FileNotFoundError (dict['res_strings']['errors']['0003'].format(dict['f_xl_data']))
    else:
        logger.info (str.format(dict['res_strings']['info']['0018'], dict['f_xl_data']))

    # Check if output path exist and if not create it.
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

def logForecastFramework (dict, logger):
    for fw_item in dict['fw']:
        logger.info (str.format("{}: {}", fw_item, dict['fw'][fw_item]))
    return

# Create the JSON summary file
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

# Create log entries of the summary data
def createLogSummary (agents, max_date, max_time, max_calls, ec, dict, logger):
    logger.info (dict['res_strings']['prompts']['0001'])
    logger.info (dict['res_strings']['prompts']['0002'].format(max_date, max_time, max_calls))
    logger.info (dict['res_strings']['prompts']['0003'].format(max_calls))
    logger.info (dict['res_strings']['prompts']['0004'].format(agents))
    logger.info (dict['res_strings']['prompts']['0005'].format(ec.Utilisation(agents, max_calls) * 100))
    logger.info (dict['res_strings']['prompts']['0006'].format(ec.Trunks(agents, max_calls)))
    logger.info (dict['res_strings']['prompts']['0007'].format(ec.SLA(agents, max_calls, dict['fw']['ServiceTime']) * 100))
    logger.info (dict['res_strings']['prompts']['0008'].format(ec.ASA(agents, max_calls)))
    logger.info (dict['res_strings']['prompts']['0009'].format(ec.Abandon(agents, max_calls) * 100))
    logger.info (dict['res_strings']['prompts']['0010'].format(ec.Queued(agents, max_calls) * 100))
    logger.info (dict['res_strings']['prompts']['0011'].format(ec.QueueTime(agents, max_calls)))
    logger.info (dict['res_strings']['prompts']['0012'].format(ec.QueueSize(agents, max_calls)))
    return

# Create the Excel Summary sheet
def xlCreateReport (env, fc, dict, logger): 
# Load the forecast output excel file
    wb = Workbook()
    wb.save (dict['f_result'])
    wb = load_workbook (dict['f_result'])
    logger.info (str.format(dict['res_strings']['info']['0015'], dict['f_result']))
    # Select the active sheet
    ws = wb.active
    ws.title = env['excel']['SS_name']
    logger.info (str.format(dict['res_strings']['info']['0022'], ws.title))
    wb.save (dict['f_result'])

# Start filling in the details - starting with the times
    cell = ws.cell(1,1) # 'tis cell 'A1'
    cell.value = 'Times'
    times = fc['times']
    row = 2
    col = 1
    for time in times:
        ws.cell(row, col).value = time
        row += 1
    logger.info(dict['res_strings']['info']['0023'])

# Next fill the dates in the first row and at the same time create the various detail sheets
    s_day = ''
    i = 0
    row = 1
    col = 2
    for i in range(len(fc) - 1):
        s_day = 'Day' + str(i)
        ws.cell(row, col).value = fc[s_day]['date']
        if (env['excel']['report-detail']):
            dateobj = datetime.strptime(fc[s_day]['date'], dict['sf_date'])
            sheet_name = dateobj.strftime(env['formats']['xl-detail'])
            wb.create_sheet (sheet_name)
        ws = wb[env['excel']['SS_name']]
        i += 1
        col += 1
    logger.info(dict['res_strings']['info']['0024'])

# Fill in the agents numbers required 
    col = 2
    for i in range(len(fc) - 1):
        row = 2
        s_day = 'Day' + str(i)
        agents = fc[s_day]['agents']
        for agent in agents:
            ws.cell(row, col).value = agent
            row += 1
        i += 1
        col += 1
    logger.info(dict['res_strings']['info']['0025'])

# Create detailed repport - if required
    if (env['excel']['report-detail']):
        logger.info(dict['res_strings']['info']['0026'])
        i = 0
        for ws in wb:
            row = 1
            col = 1
            s_day = 'Day' + str(i)
            # If this is the summary sheet skip it
            if (ws.title == env['excel']['SS_name']):
                continue
            # Fill the heading for each sheet
            for head in env['excel']['headings']:
                ws.cell(1, col, value=head)
                col += 1
            # Fill the times for each sheet
            times = fc['times']
            row = 2
            col = 1
            for time in times:
                ws.cell(row, col).value = time
                row += 1
            # Fill transactions column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['calls'])):
                ws.cell(row, col).value = fc[s_day]['calls'][x]
                row += 1
            # Fill agents column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['agents'])):
                ws.cell(row, col).value = fc[s_day]['agents'][x]
                row += 1
            # Fill SLA column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['sla'])):
                ws.cell(row, col).value = fc[s_day]['sla'][x]
                ws.cell(row, col).number_format = '0.00%'
                row += 1
            # Fill ASA column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['asa'])):
                ws.cell(row, col).value = fc[s_day]['asa'][x]
                row += 1
            # Fill ASA column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['abandon'])):
                ws.cell(row, col).value = fc[s_day]['abandon'][x]
                ws.cell(row, col).number_format = '0%'
                row += 1
            # Fill Q-percent column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['q-percent'])):
                ws.cell(row, col).value = fc[s_day]['q-percent'][x]
                ws.cell(row, col).number_format = '0%'
                row += 1
            # Fill Q-percent column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['q-time'])):
                ws.cell(row, col).value = fc[s_day]['q-time'][x]
                row += 1
            # Fill Q-percent column
            row = 2
            col += 1
            for x in range (len(fc[s_day]['q-count'])):
                ws.cell(row, col).value = fc[s_day]['q-count'][x]
                row += 1
            i += 1
    else:
        logger.info(dict['res_strings']['info']['0027'])



# Exit gracefully
    wb.save (dict['f_result'])
    logger.info(str.format(dict['res_strings']['info']['0011'], dict['f_result']))
    return
