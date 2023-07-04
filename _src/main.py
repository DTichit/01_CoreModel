from datetime import datetime
import glob as glob
import importlib.util
import sys
import xlwings
import os


from automate import Automate





if __name__ == '__main__':

    # Few parameters
    path_workbook = "MainFile.xlsx"
    path_function = "C://Users//DTichit//OneDrive - Deloitte (O365D)//Documents//TRAVAUX INTERNES//202307_Automatisation//01_CoreModel//_src//_functions"

    # Object
    auto = Automate(path_workbook, path_function)

    # Loop on all the runs
    auto.LoopRuns()
