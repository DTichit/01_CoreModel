
import xlwings as xw

class LaunchMacroXL:

    # Attributes
    path_wb = "C://Users//DTichit//Downloads//Book1.xlsm"
    macro_name = "HelloWorld"


    # Class
    def __init__(self) -> None:
        self.path_wb = "C://Users//DTichit//Downloads//Book1.xlsm"
        self.macro_name = "HelloWorld"


    # Run function
    def run(self):
        
        # Call main function
        self.launch_macro()
        


    def launch_macro(self):

        # Connect to XL application
        app = xw.App(visible=False)

        # Open XL file
        wb = app.books.open(self.path_wb)

        # Lanch macro
        app.macro(self.macro_name)()

        # Close XL file & application
        wb.close()
        app.quit()
        

