import pandas as pd 
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

wb = Workbook()
ws = wb.active


pd.set_option("display.max_columns",500)
raw_df=pd.read_csv("all_delays.csv")
#print(raw_df.head())

station_analysis=raw_df.groupby(["Station"])['Day'].count().reset_index()
print(raw_df.groupby(["Station","Code"])['Day'].count().reset_index())

for r in dataframe_to_rows(station_analysis, index=False, header=True):
    ws.append(r)

wb.save("pandas_openpyxl.xlsx")