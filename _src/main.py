from automate import Automate


if __name__ == '__main__':

    # Few parameters
    path_workbook = "MainFile.xlsx"
    path_function = "_src//_functions"

    # Object
    auto = Automate(path_workbook, path_function)

    # Loop on all the runs
    auto.LoopRuns()
