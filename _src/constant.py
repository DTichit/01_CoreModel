class Constant:


    NAME_MAIN_SHEET     = "main"
    ROW_TABLE_HEADER    = 6
    OUTPUT_CONTROLER    ="output.txt"

    # COLUMNS
    COL_RUN         = 1
    COL_NAME        = 2
    COL_FUNCTION    = 3
    COL_PYTHON_PARAM = 4
    COL_BOOLEAN_RUN = 5
    COL_DEPENDENCY  = 6
    COL_RUN_STATUS  = 7
    COL_START_TIME  = 8
    COL_END_TIME    = 9
    COL_RUN_OUTPUT  = 10

    # ID STATUS FOR RUN
    ID_RUN_STATUS_NOT_STARTED   = 0
    ID_RUN_STATUS_STARTED       = 1
    ID_RUN_STATUS_END_OK        = 2
    ID_RUN_STATUS_INSTANCE_KO   = 5
    ID_RUN_STATUS_FUNCTION_KO   = 6

    ROW_POSITION = 0

    # RUN STATUS
    RUN_STATUS = {
        '0' :   "Run not launched yet",
        '1' :   "Run started",
        '2' :   "Run completed",
        '5' :   "Function defined in Excel is not known",
        '6' :   "Error while launch function"
    }


    # Messages
    RUN_MESSAGES = {
        'BeginRuns'     : "{TimeStamp} - Begining of the runs",
        'Begin1Run'     : "{TimeStamp} - Begining of the run number {RunNumber} [{Function}]: {Name}",
        'End1Run'       : "{TimeStamp} - End of the run number {RunNumber}",
        'EndRuns'       : "{TimeStamp} - End of the runs",
        'ExportMessage' : "{TimeStamp} - Exporting output controler",
        'PathControler' : "{TimeStamp} - Output controler exported: {Path}",
        'ErrorOccured'  : "{TimeStamp} - An error occured while launching function {Function} for run number {RunNumber}. See message below: \n {Error}",
        'ErrorInstance' : "{TimeStamp} - An error occured while creating instance for function {Function} for run number {RunNumber}"

    }