import glob
import logging
import os
import pathlib
import time
from datetime import datetime, date

import numpy as np
import pandas as pd
import win32com.client

Start_time = datetime.now()
path = os.getcwd()
print(path)
# ##############################################
# # open the VBA macro file and run the formatting macro
# xl = win32com.client.Dispatch('Excel.Application')
# xl.Workbooks.Open(Filename=path + r'\Formatting Macro.xlsm', ReadOnly=1)
# xl.Application.Run('Format_1')
# xl.Application.Quit()
# del xl
# print('Formatting Done')
# ###############################################

# exporting final file with time stamp
file = pathlib.Path(os.getcwd() + "\\OutputFolder")
if file.exists() is False:
    os.mkdir(os.getcwd() + "\\OutputFolder")
    print('OutputFolder Created')

# load data of excel into pandas dataframe
# df_master_vlookup = pd.concat([pd.read_csv(f, engine='python') for f in glob.glob(path + "\\input_files\\GM_Master_merged.csv")])
df_master_vlookup = pd.concat([pd.read_csv(f, engine='python') for f in glob.glob(path + "\\input_files\\GM_Master_merged.csv")])
df_master_vlookup["POD"] = pd.to_datetime(df_master_vlookup['POD'], errors='coerce')
df_master_vlookup["APPT"] = pd.to_datetime(df_master_vlookup['APPT'], errors='coerce')
df_master_vlookup["For ocean exception Event Date"] = pd.to_datetime(df_master_vlookup['For ocean exception Event Date'], errors='coerce')
df_master_vlookup["ETA Destination Rail"] = pd.to_datetime(df_master_vlookup['ETA Destination Rail'], errors='coerce')
df_master_vlookup["Date Arrived Destination Rail"] = pd.to_datetime(df_master_vlookup['Date Arrived Destination Rail'], errors='coerce')
df_master_vlookup["Actual Outgate from Port"] = pd.to_datetime(df_master_vlookup['Actual Outgate from Port'], errors='coerce')
df_master_vlookup["Outgate but not departed"] = pd.to_datetime(df_master_vlookup['Outgate but not departed'], errors='coerce')
df_master_vlookup["For ocean exception Event Date"] = pd.to_datetime(df_master_vlookup['For ocean exception Event Date'], errors='coerce')
df_master_vlookup["ATD Rail"] = pd.to_datetime(df_master_vlookup['ATD Rail'], errors='coerce')
df_master_vlookup["Actual Departure from Dest Rail"] = pd.to_datetime(df_master_vlookup['Actual Departure from Dest Rail'], errors='coerce')
df_master_vlookup["ATD Rail"] = pd.to_datetime(df_master_vlookup['ATD Rail'], errors='coerce')
df_master_vlookup["Estimated Arrival Date To Origin Port"] = pd.to_datetime(df_master_vlookup['Estimated Arrival Date To Origin Port'], errors='coerce')
df_master_vlookup["DELAY EDA"] = pd.to_datetime(df_master_vlookup['DELAY EDA'], errors='coerce')
df_master_vlookup["ORIGINAL EDA"] = pd.to_datetime(df_master_vlookup['ORIGINAL EDA'], errors='coerce')
df_master_vlookup["DELAY ETA"] = pd.to_datetime(df_master_vlookup['DELAY ETA'], errors='coerce')
df_master_vlookup["ETA"] = pd.to_datetime(df_master_vlookup['ETA'], errors='coerce')

df_direct_truck_comment = pd.concat([pd.read_excel(f) for f in glob.glob(path + "\\input_files\\DIRECT TRUCK LOG*.xlsx")])
df_direct_truck_comment.dropna(subset=["Container"], inplace=True)
df_direct_truck_comment.drop_duplicates(subset="Container", inplace=True)
df_direct_truck_comment['COMMENTS'] = df_direct_truck_comment['COMMENTS'].astype(str).str.split(";", 1)
df_direct_truck_comment['COMMENTS'] = df_direct_truck_comment['COMMENTS'].str[0] + ";"
df_direct_truck_comment.rename(columns={"Container": "Container_direct_truck_comment", "COMMENTS": "COMMENTS_direct_truck"}, inplace=True)
df_master_vlookup = df_master_vlookup.merge(df_direct_truck_comment[["Container_direct_truck_comment", "COMMENTS_direct_truck"]], left_on="ContainerID", right_on="Container_direct_truck_comment", how='left')
df_master_vlookup["Last Staus on Rail"] = np.where((df_master_vlookup['COMMENTS_direct_truck'].notnull()), (df_master_vlookup['COMMENTS_direct_truck']), df_master_vlookup["Last Staus on Rail"])

##############################################################
df_master_vlookup.to_csv(path + "\\OutputFolder\\Final_master.csv", index=False)
##############################################################
End_time = datetime.now()
print("Process Completed!")
print("Time of Processing :", End_time - Start_time)
time.sleep(5)
