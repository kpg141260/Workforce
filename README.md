# Workforce Management System *WFMS*

Workforce Management System (*WFMS*) for Contact Centers supporting agent number and Key Performance Indicator (KPIs) forecasting based on either historical data stored in a database or provided in spreadsheet.
The current version (0.3.0) only supports the import of historical call data - no other parameters - the support for import of other parameters will be available in a near future version (I'm working on it).
Other KPIs such as ATT, ACW, ASA, Max_Wait, Agent Availability, etc. need to be provided through a json configuration file (for the time being - I am working on the option to provide these parameters in the same spreadsheet as the transaction volumes).

## General

*WFMS* can be used stand-alone (Excel to Excel) or with an associated database (Excel to DB, DB to DB, DB to Excel).
The general use case for now is such:
1. User creates a transaction volume forecast for a period of time and saves that forecast in an Excel spreadsheet organised as detailed [here](#_h_data_source).
2. User stores that report in the location defined [here](#h_files_folders).
3. User edits [framework.json](#h_fw_file) file to provide forecast parameters.
4. User edits [configuration file](#h_sys_config) to control what input and output they desire.
5. User executes wfms.

### Summary of Major Functions

*WFMS* supports the following major functionality:

1. Read an excel spreadsheet that contains forecasted transaction volumes and generate a spreadsheet that projects the required agents numbers and associated KPIs. The spreadsheet containing the forecasted transaction volume has to meet the requirements detailed [here](#h_data_source).
> **Status: implemented**
2. Import into the database an Excel Spreadsheet containing historical transaction volumes. The spreadsheet containing the forecasted transaction volume has to meet the requirements detailed [here](#h_data_source).
> **Status: implemented**
3. Retrieve future transaction volumes from database and generate a forecast exported to Excel that projects the required agent numbers and the projected associated KPIs.
> **Status: Not yet implemented**
4. Retrieve future transaction volumes from database and generate a forecast stored in the database that projects the required agent numbers and the projected associated KPIs.
> **Status: Not yet implemented**
5. Retrieve historical transaction volumes from database, analyse trends, generate a transaction volume forecast based on the analysed trends, and generate a spreadsheet that projects the required agent numbers and the projected associated KPIs.
> **Status: implementing**
6. Retrieve historical transaction volumes from database, analyse trends, generate a transaction volume forecast which is based on the analysed trends, predict the required agent numbers and associated KPIs, and store that forecast in the database.
> **Status: Not yet implemented**


## Dependencies

**Python version >= 3.8**

- **json**
- **datetime**
- **openpyxl**
- **logging**
- **os**
- **pymongo**
- **erlang-c**

---

## <a name="h_dir_struct"></a>Folder Structure

```
wfm-folder
    ├── conf      - configuration files
    ├── data      - import data and forecast parameters - such as ATT, ACW, ASA, etc.
    ├── forecast  - contains forecast data (xlsx or json or both)
    ├── lib       - library files
    └── logs      - log files
```
### Folder *'conf'*

The ```conf``` folder contains configuration files that control the input, output and general program flows, as well as names, locations of files.
The main configuration file is ```wfms.conf```. <span style="color:red">The name and location of this file is hard-coded into the *WFMS* code and therefore this file or folder cannot be renamed nor moved!</span>

### Folder *'data'*

Data to be imported should be placed in this folder. Per default the WFM system expects a file called ```call_data.xlsx``` containing the forecasted call volume for a period of days (minimum is 1 max is limited by Excel limitation) in this folder. Additionally it expects a file called ```framework.json``` in this folder.

There are two ways to provide the forecast framework to the WFM system, one is to provide the ```framework.json``` file, the second is to provide the parameters within the ```call_data.xlsx``` file in a separate tab called 'framework'. The WFM system will try to load the framework file first, if it cannot find teh framework file or the framework file contains errors, it will try to load the forecast parameters from the 'framework' tab in the workbook. If that fails th WFM system will signal a Fatal error and quit. Otherwise it will process with the import of the forecasting parameters.

### Folder *'forecast'*

This folder will contain the Excel forecast file if Excel output is configured in the [configuration file](#h_sys_config).

### Folder *'lib'*

This folder contains library files used by the *WFMS*.

### Folder *'logs'*

This folder contains the *WFMS* [log files](#h_logs).


### Folder *'res'*

This folder contains any resource files needed by *WFMS*.
---

## <a name="h_sys_config"></a>System Configuration

The software is written entirely in Python (version 3.8) and has been designed to work on an application server without user interface. Thus its workflow and function depends on a configuration file called ```conf/wfms.conf```.

### <a name="h_logs"></a>*WFMS* Log System

The log system logs certain events that occur when running the *WFMS*. There are 5 log levels that can be configured in the [configuration file](#h_sys_config):
- *DEBUG* - <a name="h_log_sys"></a>this is a mode that logs a lot of events and should not be used during normal operation
- *INFO* - less event messages than DEBUG but still quite verbose
- *WARNING* - does only display messages that are warnings, errors or Fatal
- *ERROR* - only logs error and Fatal events
- *FATAL* - only logs events that are Fatal

The default setting is *WARNING*. If you experience any issues with the software, set the log level to *INFO*.

#### Configuring the Log System

Configuration of the log system is done in the file ```conf/wfms.conf```. The file is a JSON file and thus can easily be edited with any text editor.
To customise the Log System the section called ```log-config``` can be changed as shown below.
```json
"log-conf": {
        "log-to-file": true,
        "log-to-stdout": true,
        "rotate": true,
        "max_bytes": 50000,
        "backup_count": 10,
        "truncate": false,
        "level": "INFO",
        "log-summary": false,
        "log-file": "wfms.log"
    }
```
#### Log Configuration Parameters
- **log-to-file:** ```boolean -- true/false``` -- specifies whether the log function is recording log events to the log-file.
> **true** = log to file</br>
> **false** = do not log to file
- **log-to-stdout:**  ```boolean -- true/false``` -- specifies whether the log function displays events on the console.
> **true** = display events on console</br>
> **false** = do not display events on console
- **rotate:** ```boolean -- true/false``` -- specifies whether the log system will automatically over-write older events.
> **true** = over-write older events</br>
> **false** = do not over-write older events
- **truncate:** ```boolean -- true/false``` -- specifies whether the log function will truncate once it reaches *max_bytes* (erase the content of the log file) regardless of the setting of rotate. This means that *truncate* takes precedence over rotate.
> **true** = truncate file once reaching max_bytes and disregard setting of *rotate*</br>
> **false** = do not truncate file once reaching *max_bytes* and respect the setting of *rotate*
- **log-summary:** ```boolean -- true/false``` -- specifies whether the log function shall write the forecasted summary data to the log-file.
> **true** = log summary</br>
> **false** = do not log summary
- **max_bytes:** ```int``` -- specifies the size of the log file in bytes before it gets 'rotated' or truncated. This value is set to 50000 bytes by default
- **backup_count:** ```int``` -- specifies that the rotation of the log file shall create backup copies (value > 0) of the log file before rotating it, and how many backup copies shall be kept. This values is set to 10 by default. The backup copies are named with a sequential number at the end of the file name > *workforce.log.1*, *workforce.log.2* and so on, until the number of specified copies are reached. then the system will over-write the oldest backup.
- **level:** ```string``` -- defines the log level (detail) that the log system will record. See [above](#h_log_sys) for a more detailed explanation.
- **log-file:** ```string``` -- specifies the name of the log-file. This can be changed. The default is ```wfms.log```.

<span style="color:red">**Any other settings in this section should not be edited!**</span>

#### Sample Event Log

Below you can see the beginning of the log-file after starting the application with instructions to import data from a spreadsheet into the database using a log level of 'INFO'.
```
'2020-06-09:14:05:29,660  INFO     wfm  Workforce Forecast System version 0.3.0 loaded.
'2020-06-09:14:05:29,661  INFO     wfm  File '/Users/username/Workforce/res/workforce-enGB.json' successfully loaded.
'2020-06-09:14:05:29,661  INFO     wfm  File '/Users/username/Workforce/data/framework.json' successfully loaded.
'2020-06-09:14:05:29,661  INFO     wfm  Path 'forecast' exists.
'2020-06-09:14:05:29,661  INFO     wfm  File 'data/call_data.xlsx' exists.
'2020-06-09:14:05:29,662  INFO     wfm  File 'forecast/forecast.xlsx' exists.
'2020-06-09:14:05:29,662  INFO     wfm  File 'forecast/forecast-summary.json' exists.
'2020-06-09:14:05:29,662  INFO     wfm  Workforce Forecast Module started.
'2020-06-09:14:05:29,662  INFO     wfm  Function to produce forecast based on database data not yet implemented.
'2020-06-09:14:05:33,208  INFO     wfm  Starting data import from spreadsheet. Target: 10.10.0.8:27017.wfm_main_db.history.
```

### <a name="h_files_folders"></a>Configuring Folders and Filenames

*WFMS* provisions the configuration of filenames folders for data import and export. The configuration is done in the file ```conf/wfms.conf```. To customise Folders and Filenames the sections called ```files and paths``` can be changed as shown below.
```json
"files": {
    "framework": "framework.json",
    "in-data": "call_data-1.xlsx",
    "out-data": "forecast",
    "summary": "forecast-summary",
    "language-file": "workforce",
    "language-type": "json",
    "ext-json": "json",
    "ext-xl": "xlsx",
    "in-type": "xlsx"
},
"paths": {
    "rsrc-path" : "res",
    "config-path" : "conf",
    "output-path" : "forecast",
    "input-path": "data",
    "log-path": "logs"
},
```

#### File Configuration Parameters

- **framework** ```string``` -- the name of the framework file, default is ```framework.json```; the location of this file is specified under ```"paths": {"input-path"}```.
- **in-data** ```string``` -- the name of the file that contains the data to be processed, either to create a report using forecast data or to import historical data - the default filename is ```tx_data.xlsx```. The location of this file is specified under ```"paths": {"input-path"}```.
- **out-data** ```string``` -- the name of the output file, the forecast. The default name is ```forecast```. The location of the output file is specified under ```"paths": {"out`put-path"}```.
> There is no extension specified in this entry. The file extension is automatically appended based on the output options chosen in settings described further below.</br>
These options are specified in the section ```output-format```
- **summary** ```string``` -- specifies the filename of the forecast summary file refer to [Forecast Summary](#h_fc_summary)" for more information. The default name for the forecast summary file is ```forecast-summary```.
> There is no extension specified in this entry. The file extension is automatically appended based on the output options chosen in settings described further below.</br>
These options are specified in the section ```"options": "summary": true```. If set to ```true``` the summary file will be created; if set to ```false``` the summary file will not be created.
- **language-file** ```string``` -- defines the name of the language file. <span style="color:red">Do not change this entry!</span>
- **language-type** ```string``` -- defines the file type of the language resource file. <span style="color:red">Do not change this entry!</span>
- **ext-json** ```string``` -- defines the extension of json files. <span style="color:red">Do not change this entry!</span>
- **ext-xl** ```string``` -- defines the extension of spreadsheet files. <span style="color:red">Do not change this entry!</span>
- **in-type** ```string``` -- defines the extension of the data input file. For the time being *WFMS* only supports .xlsx files. <span style="color:red">Do not change this entry!</span>

#### Folder Configuration Parameters

- **rsrc-path** ```string``` -- This entry specifies the folder that contains any resource files needed by *WFMS*. <span style="color:red">Do not change this entry!</span>
- **config-path** ```string``` -- This entry specifies the folder that contains the configuration files needed by *WFMS*. <span style="color:red">Do not change this entry!</span>
- **output-path** ```string``` -- This entry specifies the path to the output (forecast) files *WFMS* generates. <span style="color:red">Do not change this entry!</span>
- **input-path** ```string``` -- This entry specifies the path to the input files *WFMS* needs if input is through Excel spreadsheets. <span style="color:red">Do not change this entry!</span>
- **log-path** ```string``` -- This entry specifies the folder that will hold the log files generated by *WFMS*. <span style="color:red">Do not change this entry!</span>

### Mongo DB Connection

To use database functionality a Mongo database has to be available, either locally or on the network.
The mongo connection parameters are configured in the file ```/conf/wfms.conf```, as shown below.

```json
"db": {
    "url": "10.10.0.8",
    "port": 27017,
    "connectTimeoutMS": 10000,
    "socketTimeoutMS": 5000,
    "db-name": "wfm_main_db",
    "col-in": "history",
    "col-out": "forecast",
    "col-frame": "framework",
    "col-log": "log",
    "col-resources": "resources"
}
```
| Key | Type | Description | Notes |
| :--- | :--- | :--- | :--- |
| **url** | ```string``` | This entry specifies the url of the Mongo database server.</br>Either a IP address (```i.e. 10.10.1.4```) or a url ```mongodb://my.mongos.com``` | <span style="color:green">Edit this entry to point to your server.</span> |
| **port** | ```string``` | This entry specifies the port of the Mongo database server. | <span style="color:green">Edit this entry to point to the Mongo DB port.</span> |
| **connectTimeoutMS** | ```string``` | Connection timeout for the Mongo DB server connection.| Default is set to 10000 or 10 seconds. |
| **socketTimeoutMS** | ```string``` | Connection timeout for the Mongo DB socket connection.| Default is set to 5000 or 5 seconds. |
| **db-name** | ```string``` | The name of the *WFMS* database on the Mongo server. | <span style="color:red">Do not change this entry!</span> |
| **col-in** | ```string``` | The name of the Collection in the Mongo DB that holds historical transaction data. | <span style="color:red">Do not change this entry!</span> |
| **col-out** | ```string``` | The name of the Collection in the Mongo DB that holds forecast transaction data. | <span style="color:red">Do not change this entry!</span> |
| **col-frame** | ```string``` | The name of the Collection in the Mongo DB that holds forecast framework data.</br><span style="color:blue">Not yet implemented.</span>  | <span style="color:red">Do not change this entry!</span> |
| **col-log** | ```string``` | The name of the Collection in the Mongo DB that holds the log data.</br><span style="color:blue">Not yet implemented.</span>  | <span style="color:red">Do not change this entry!</span> |
| **col-resources** | ```string``` | The name of the Collection in the Mongo DB that holds resource strings.</br><span style="color:blue">Not yet implemented.</span> | <span style="color:red">Do not change this entry!</span> |

## <a name="h_data_source"></a>Using Data Provided through an Excel Workbook

*WFMS* supports 2 types of import files:
1. a file only containing dates, times and transaction data in a single spreadsheet [Transaction Volume Only](#h_io_simple);
2. a files containing dates, times, transaction data, and KPIs distributed over several sheets, each sheet representing one week of data [Transaction Volume Plus KPIs](#h_io_complex).

### <a name="h_io_simple"></a>Transaction Volume Only Spreadsheet

#### Spreadsheet Layout for Transaction Volume Only

| | A | B | C | ... |
| :---: | ---: | :---: | :---: | :---: |
| **1** | **Times** | **2019/06/31** | **2019/06/30** | ... |
| **2** | **08:00** | 120 | 112 | ... |
| **3** | **09:00** | 150 | 160 | ... |
| **4** | **10:00** | 220 | 221 | ... |
| **...** | **...** | ... | ... | ... |
> Note that the Times difference (the interval) can be 15 minutes, 30 minutes, 45 minutes, and 60 minutes.</br>
For example a 15 min interval would be 06:00, 06:15, 06:30, 06:45, and so on;</br>
a 30 min interval would be 06:00, 06:30, 07:00, 07:30, and so on;</br>
When loading any data from a spreadsheet the processing of the data will continue until a bracket is found that does not fall into the defined intervals (15, 30, 45, 60).</br>
Data up to that point will be imported. So, if you are missing import data or your forecast report is incomplete, check the log file, as the system will tell you which cell is in error.

#### Expected Cell Formatting for Transaction Volume Only

| | A | B | C | ... |
| :---: | :--- | :--- | :--- | :--- |
| **1** | Not Used | **Datetime** | **Datetime** | ... |
| **2** | **Datetime** | Transactions | Transactions | ... |
| **3** | **Datetime** | Transactions | Transactions | ... |
| **4** | **Datetime** | Transactions | Transactions | ... |
| **...** | **...** | ... | ... | ... |

> The *Transaction Volume Only* workbook may contain more than one sheet, but the actual data to be processed needs to be located on a **`single sheet** with the structure as shown above.</br>
The name of the sheet that contains the data can be configured in the **configuration file** ```wfms.conf```, **Section** ```excel```, **Key** ```SD_name``` The default is: ```Sheet1```.

### <a name="h_io_complex"></a>Transaction Volume Plus KPIs Spreadsheet

#### <a name="h_fw_file></a>Framework File

## WFMS File Output

### <a name="h_fc_summary"></a>Forecast Summary

#### Forecast Summary JSON

#### Forecast Summary XLSX

## <a name="h_db_1"></a>WFMS Database

### WFMS Database Output

## Communicating with WFMS from a WEB Application
