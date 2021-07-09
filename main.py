#   Marcin Knap 2021
#    for MPT TABO
# knapmarcin@icloud.com
from Route import Route
import utils

#insert your APIKEY from google with directions, distance matrix, geocode
api_key = ''


# RUN IN TABO \/


# excel = win32com.client.Dispatch("Excel.Application")
# wb1 = pd.ExcelFile('\\\\tabo-srv1\\logist\\dokumenty.xls')
# sheet = pd.read_excel(wb1, sheet_name="Arkusz1", usecols=["DATA"])
# df = sheet
# nazwa_arkusza = df['DATA'].iloc[0]
#
# folder_name = nazwa_arkusza[-4:]
# file_name = nazwa_arkusza[-7:]+'.xlsm'
# file_directory = 'T:\\poziom 0\\optima_analizy_tabo upg\\logist\\SPECYFIKACJA\\'+ folder_name +'\\'+ file_name
# df = Route.create_dataframe_from_excel(file_directory, nazwa_arkusza)
# df = Route.complete_km(Route.directions_dataframe(Route.geocode_dataframe(df, api_key), api_key), api_key)
# Route.connect_excel_with_df(df, file_directory, nazwa_arkusza)


# RUN IN TABO /\



# RUN IN HOME \/

df = utils.open_test_excel()
df = Route.complete_km(Route.directions_dataframe(Route.geocode_dataframe(df, api_key), api_key), api_key)
Route.connect_excel_with_df(df, "01-2021.xlsm", "WT 05-01-2021")
Route.create_map()

# RUN IN TABO /\
