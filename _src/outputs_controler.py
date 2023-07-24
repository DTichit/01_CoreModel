from datetime import datetime
import os
from constant import Constant


class OutputControler:

    def __init__(self):
        # String with all the outputs
        self.output = ""

        # Print outputs in consol
        self.consol = True 
    

    def print_text(self, text):
        print(text)

    # Add a time stamp to a text
    def add_timestamp_to_text(self, text):
        return(
            str(datetime.now()) + " - " + text
        )

    # Add text to the 
    def add_text(self, text):

        # Concat string
        self.output += "\n" + text

        # Print message
        if self.consol:
            self.print_text(text)
    
    
    # Export output control
    def export_control(self):
        
        # File 
        # DT - 03.07 : can be changed, if we want it to be dynamic
        t_file = Constant.OUTPUT_CONTROLER
        t_folder_path = os.path.dirname(os.path.abspath(__file__))
        t_path = os.path.join(t_folder_path, t_file)

        # Open the file in write mode
        with open(t_path, "w") as file:
            file.write(self.output)

        # Return   
        return(t_path)



