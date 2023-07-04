from datetime import datetime
import glob as glob
import importlib.util
import sys
import xlwings
import os

from constant import Constant
from excel_controler import MainSheetControler
from outputs_controler import OutputControler




class Automate:

    def __init__(self, path_workbook, path_functions):

        # Main sheet controler
        self.main_sheet = MainSheetControler(path=path_workbook)

        # Initialize object output
        self.output = OutputControler()

        # Load functions
        self.functions = self.load_functions(path=path_functions)

        # Status run
        # 03.07 - DT : we might include all information from the runs in the list to avoid reading several times in XL
        self.runs_status = {}


    # Change status for a run
    def set_run_status(self, run_id, status_id):
        self.runs_status[str(run_id)] = status_id
    
    
    # Get run status
    def get_run_status(self, run_id):
        return(self.runs_status[str(run_id)])
    

    # Load all the functions
    def load_functions(self, path):

        # Initialize list
        t_functions = {}

        # Sys append
        sys.path.append(path)

        # Loop on all the files
        for file_name in os.listdir(path):

            # Load only python files
            if file_name.endswith(".py"):
                module_name = file_name[:-3] 
                module_path = os.path.join(path, file_name)

                # Importer le module
                module = importlib.import_module(module_name, module_path)
        
            # Go through all the functions
            for attr_name in dir(module):
                attr = getattr(module, attr_name)
                if isinstance(attr, type):
                    # Créer une instance de la classe dynamiquement
                    instance = attr()
                    # Stocker l'instance dans le dictionnaire
                    t_functions[attr_name] = instance
        
        # Outputs
        return(t_functions)
            



    # Fonction pour créer une instance de classe dynamiquement
    def create_instance(self, class_name):
        if class_name in self.functions:
            return self.functions[class_name]
        else:
            return None



    # Main function which allows to loop on all the runs
    def LoopRuns(self):

        # Handle begining of the runs
        self.handle_begin_runs()
        
        # Loop on all the runs
        for run_id in range(3):

            # Define row position for the given run
            ROW_POSITION = self.main_sheet.GetRowPosition(run_id)

            # Launch 1 run
            if self.main_sheet.GetBooleanRun(run_id=run_id):
                self.LaunchRun(run_id)
            else:
                False
        
        # Handle end of runs 
        self.handle_end_runs()


    # Launch a given run : launch function and handle various other things (time, status, outputs ...)
    def LaunchRun(self, run_id):

        # Handle start run
        self.handle_start_run(run_id)

        # Launch function
        self.LaunchFunction(run_id)

        # Set End Time
        self.handle_end_run(run_id)

        

    # For a given run, launch the function
    def LaunchFunction(self, run_id):

        # Get function to be launched
        function_name = self.main_sheet.GetFunctionRun(run_id=run_id)

        # Create instance
        function = self.CreateInstanceFunction(function_name, run_id)
        
        # Check whether everything went well so far
        if (self.get_run_status(run_id=run_id) == Constant.ID_RUN_STATUS_STARTED):
            # Launch function and cathc the error (if any)
            try:
                function.run()
            except Exception as exception:
                self.handle_exception(run_id, exception)


    # Handle exception if occurs
    def handle_exception(self, run_id, exception):

        # Set run status
        self.set_run_status(run_id, Constant.ID_RUN_STATUS_FUNCTION_KO)

        # Error message
        self.handle_messages(run_id=run_id, message="ErrorOccured", exception=exception)



    # Handle messages 
    def handle_messages(self, message, run_id=None, time=datetime.now(), exception=None, path=None):

        # List with all parameters for the comment
        t_list = {
            'TimeStamp' : time,
            'RunNumber' : run_id,
            'Error'     : exception,
            'Name'      : self.main_sheet.GetNameRun(run_id) if run_id is not None else None,
            'Function'  : self.main_sheet.GetFunctionRun(run_id) if run_id is not None else None,
            'Path'      : path
        }

        # Message
        t_message = Constant.RUN_MESSAGES[message].format(**t_list)

        # Update output control
        self.output.add_text(t_message)


    # For a given run, create instance for the function to be launched 
    def CreateInstanceFunction(self, function_name, run_id):
        if function_name in self.functions:
            return self.functions[function_name]
        else:
            self.handle_error_create_instance(run_id)
            return None      


    # Handle when error during creation of instance
    def handle_error_create_instance(self, run_id):

        # Set run status
        self.set_run_status(run_id, Constant.ID_RUN_STATUS_INSTANCE_KO)

        # Message output
        self.handle_messages(run_id=run_id, message="ErrorInstance")



    def handle_begin_runs(self):

        # Start time of the runs
        self.handle_messages(message="BeginRuns")

        # Other, 1 day ... maybe ...



    # Handle start of 1 run
    def handle_start_run(self, run_id):

        # Set status to start
        self.set_run_status(run_id, Constant.ID_RUN_STATUS_STARTED)

        # Get time
        t_time = self.get_time()

        # Update main sheet controler
        self.main_sheet.SetStartTime(run_id, t_time)

        # Handle message 
        self.handle_messages(run_id=run_id, time=t_time, message="Begin1Run")



    # Handle the end of 1 run
    def handle_end_run(self, run_id):

        # Set status to end : only if everything went well
        if (self.get_run_status(run_id=run_id) == Constant.ID_RUN_STATUS_STARTED):
            self.set_run_status(run_id, Constant.ID_RUN_STATUS_END_OK)

        # Get time
        t_time = self.get_time()

        # Update main sheet controler
        self.handle_sheet_controler_end_1run(run_id, t_time)

        # Handle message
        self.handle_messages(run_id=run_id, message="End1Run", time=t_time)



    def handle_sheet_controler_end_1run(self, run_id, t_time):

        # Update main sheet controler
        self.main_sheet.SetEndTime(run_id, t_time)

        # Update run status
        self.main_sheet.SetRunStatus(run_id, self.get_run_status(run_id))



    # Hnadles end of run
    def handle_end_runs(self):

        # Message pour end of the runs
        self.handle_messages(message="EndRuns")

        # Export output controller
        self.handle_export_controler()

    

    # Handle export 
    def handle_export_controler(self):

        # Message export
        self.handle_messages(message="ExportMessage")

        # Export controler
        r_path = self.output.export_control()

        # Path of controler
        self.handle_messages(message="PathControler", path=r_path)



    def get_time(self):
        return(datetime.now())


    