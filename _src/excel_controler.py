import xlwings

from constant import Constant


class MainSheetControler:

    def __init__(self, path):
        # Open workbook
        t_wb = self.OpenWorkbook(path)

        # Open main sheet
        self.main_sheet = self.OpenMainSheet(t_wb)

    
    # Open Workbook
    def OpenWorkbook(self, path):
        return(xlwings.Book(path))


    # Open main sheet, with the table of runs
    def OpenMainSheet(self, workbook):
        return(workbook.sheets[Constant.NAME_MAIN_SHEET])
    

    # Set start time and date for a given run
    def SetStartTime(self, run_id, time):
        self.main_sheet.range(self.GetRowPosition(run_id), Constant.COL_START_TIME).value = time


    # Set end time and date for a given run
    def SetEndTime(self, run_id, time):
        self.main_sheet.range(self.GetRowPosition(run_id), Constant.COL_END_TIME).value = time


    # Get row position for a given run
    def GetRowPosition(self, run_id):
        return(Constant.ROW_TABLE_HEADER + run_id)


    # Get BooleanStatus
    def GetBooleanRun(self, run_id):
        return(
            self.main_sheet.range(self.GetRowPosition(run_id), Constant.COL_BOOLEAN_RUN).value
        )


    # Get Name
    def GetNameRun(self, run_id):
        return(
            self.main_sheet.range(self.GetRowPosition(run_id), Constant.COL_NAME).value
        )


    # Get Function to call ofr a given run
    def GetFunctionRun(self, run_id) -> str:
        return(
            self.main_sheet.range(self.GetRowPosition(run_id), Constant.COL_FUNCTION).value
        )


    # Set run status
    def SetRunStatus(self, run_id, status_id):
        self.main_sheet.range(self.GetRowPosition(run_id), Constant.COL_RUN_STATUS).value = self.GetRunDescription(status_id)


    # Get run description
    def GetRunDescription(self, id_status):
        return(Constant.RUN_STATUS[str(id_status)])


    # Get number of run to be launched : check on the colums ofrun status
    def get_no_runs(self):

            # Initialize count
            t_count = 0

            # Ignore headers
            row = Constant.ROW_TABLE_HEADER + 1
            
            # Loop on the rows
            while self.main_sheet.range(Constant.COL_BOOLEAN_RUN, row).value is not None:
                t_count += 1
                row += 1

            
            # Output
            return(t_count)


    