from Route import Route



def open_excel_from_macro():
    """
    Gets file directory from DATA column, in example:
    Wt 02-01-2021
    sheet_name = Wt 02-01-2021
    folder = 2021
    file_name = 01-2021 + 'xlsm'
    base_directory = ex. poziom0/home/logist/specyfikacja/
    file_directory = base_directory + 2021 + '/' + filename
    gets sheet_name, file_directory from static file location
    returns Route.create_dataframe_from_excel(file_directory, sheet_name)
    """
    pass


def open_test_excel():
    """
    hardcoded file directory to test_object, later to be replaced with open_excel_from_macro()
    returns
        path to a test excel
    """
    file_directory = '01-2021.xlsm'
    sheet_name = 'WT 05-01-2021'
    return Route.create_dataframe_from_excel(file_directory, sheet_name)

