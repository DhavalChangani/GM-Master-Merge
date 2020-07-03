import glob
import logging
import os
import pathlib
import time
from datetime import datetime, date
import re

import numpy as np
import pandas as pd
import win32com.client
import win32com.client

logging.basicConfig(filename='Master_merge_error_file.log', level=logging.DEBUG)
try:
    Start_time = datetime.now()
    path = os.getcwd()
    print(path)
    ###############################################
    # open the VBA macro file and run the formatting macro
    xl = win32com.client.Dispatch('Excel.Application')
    xl.Workbooks.Open(Filename=path + r'\Formatting Macro.xlsm', ReadOnly=1)
    xl.Application.Run('Format_1')
    xl.Application.Quit()
    del xl
    print('Formatting Done')
    ###############################################
except:
    logging.exception("\n \n \n Error logged: Excel Formatting")

# exporting final file with time stamp
file = pathlib.Path(os.getcwd() + "\\OutputFolder")
if file.exists() is False:
    os.mkdir(os.getcwd() + "\\OutputFolder")
    print('OutputFolder Created')

try:
    # load data of excel into pandas dataframe
    # df_master_vlookup = pd.concat([pd.read_csv(f, engine='python') for f in glob.glob(path + "\\input_files\\GM_Master_merged.csv")])
    df_master_vlookup = pd.concat([pd.read_csv(f, engine='python', dtype={'MBL#': str}) for f in glob.glob(path + "\\input_files\\GM_Master_merged.csv")])
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
    df_master_vlookup["Estimated Arrival Date To Origin Port"] = pd.to_datetime(df_master_vlookup['Estimated Arrival Date To Origin Port'], errors='coerce')
    df_master_vlookup["DELAY EDA"] = pd.to_datetime(df_master_vlookup['DELAY EDA'], errors='coerce')
    df_master_vlookup["ORIGINAL EDA"] = pd.to_datetime(df_master_vlookup['ORIGINAL EDA'], errors='coerce')
    df_master_vlookup["DELAY ETA"] = pd.to_datetime(df_master_vlookup['DELAY ETA'], errors='coerce')
    df_master_vlookup["ETA"] = pd.to_datetime(df_master_vlookup['ETA'], errors='coerce')
    df_master_vlookup["ATA to POL"] = pd.to_datetime(df_master_vlookup['ATA to POL'], errors='coerce')
except:
    logging.exception("\n \n \n Error logged: Master Read \n")

#################################################################################################################################
# Non Transload POD
try:
    df_non_transload = pd.concat([pd.read_csv(f, engine='python') for f in glob.glob(path + "\\input_files\\GM_Non_Transload.csv")])
    df_non_transload.dropna(subset=["ContainerID"], inplace=True)
    df_non_transload.drop_duplicates(subset="ContainerID", inplace=True)
    df_non_transload.rename(columns={"ContainerID": "ContainerID_non_transload",
                                     "Rail/DT": "Rail/DT_non_transload",
                                     "ATD Rail": "ATD Rail_non_transload",
                                     "Date Arrived Destination Rail": "Date Arrived Destination Rail_non_transload",
                                     "City / Rail Yard Name": "City / Rail Yard Name_non_transload",
                                     "Actual Departure from Dest Rail": "Actual Departure from Dest Rail_non_transload",
                                     "POD": "POD_non_transload",
                                     "POD Source": "POD Source_non_transload"}, inplace=True)

    df_master_vlookup = df_master_vlookup.merge(df_non_transload[["ContainerID_non_transload", "Rail/DT_non_transload",
                                                                  "ATD Rail_non_transload", "Date Arrived Destination Rail_non_transload",
                                                                  "City / Rail Yard Name_non_transload", "Actual Departure from Dest Rail_non_transload",
                                                                  "POD_non_transload", "POD Source_non_transload"]],
                                                left_on="ContainerID", right_on="ContainerID_non_transload", how='left')

    df_master_vlookup["POD_non_transload"] = pd.to_datetime(df_master_vlookup['POD_non_transload'], errors='coerce')
    df_master_vlookup["ATD Rail_non_transload"] = pd.to_datetime(df_master_vlookup['ATD Rail_non_transload'], errors='coerce')
    df_master_vlookup["POD_non_transload"] = pd.to_datetime(df_master_vlookup['POD_non_transload'], errors='coerce')
    df_master_vlookup["Actual Departure from Dest Rail_non_transload"] = pd.to_datetime(df_master_vlookup['Actual Departure from Dest Rail_non_transload'], errors='coerce')

    df_master_vlookup["Rail/DT"] = np.where((df_master_vlookup["Rail/DT"].isnull()) & (df_master_vlookup['Rail/DT_non_transload'].notnull()), df_master_vlookup['Rail/DT_non_transload'], df_master_vlookup["Rail/DT"])
    df_master_vlookup["ATD Rail"] = np.where((df_master_vlookup['ATD Rail'].isnull()) & (df_master_vlookup['ATD Rail_non_transload'].notnull()), df_master_vlookup['ATD Rail_non_transload'], df_master_vlookup['ATD Rail'])
    df_master_vlookup["Date Arrived Destination Rail"] = np.where((df_master_vlookup['Date Arrived Destination Rail'].isnull()) & (df_master_vlookup['Date Arrived Destination Rail_non_transload'].notnull()), df_master_vlookup['Date Arrived Destination Rail_non_transload'], df_master_vlookup['Date Arrived Destination Rail'])
    df_master_vlookup["City / Rail Yard Name"] = np.where((df_master_vlookup['City / Rail Yard Name'].isnull()) & (df_master_vlookup['City / Rail Yard Name_non_transload'].notnull()), df_master_vlookup['City / Rail Yard Name_non_transload'], df_master_vlookup['City / Rail Yard Name'])
    df_master_vlookup["Actual Departure from Dest Rail"] = np.where((df_master_vlookup['Actual Departure from Dest Rail'].isnull()) & (df_master_vlookup['Actual Departure from Dest Rail_non_transload'].notnull()), df_master_vlookup['Actual Departure from Dest Rail_non_transload'], df_master_vlookup['Actual Departure from Dest Rail'])
    df_master_vlookup["POD"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['POD_non_transload'].notnull()), df_master_vlookup['POD_non_transload'], df_master_vlookup['POD'])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup['POD Source'].isnull()) & (df_master_vlookup['POD Source_non_transload'].notnull()), df_master_vlookup['POD Source_non_transload'], df_master_vlookup['POD Source'])
except:
    logging.exception("\n \n \n Error logged: Non Transload POD \n")
#################################################################################################################################
# ODC GIT
try:
    df_ODC_GIT = pd.concat([pd.read_excel(f) for f in glob.glob(path + "\\input_files\\odc*git*.xlsx")])
    df_ODC_GIT.dropna(subset=["Container ID"], inplace=True)
    df_ODC_GIT.drop_duplicates(subset="Container ID", inplace=True)
    df_ODC_GIT.rename(columns={"Container ID": "ContainerID_odc_git"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_ODC_GIT[["ContainerID_odc_git"]], left_on="ContainerID", right_on="ContainerID_odc_git", how='left')
    df_master_vlookup["ODC"] = np.where(df_master_vlookup['ContainerID_odc_git'].notnull(), "YES", "NO")
    df_master_vlookup["Place of Delivery"] = np.where(df_master_vlookup['ODC'].astype(str).str.upper().str.contains("YES"), "DRDC ILG Detroit Regional Distribution", df_master_vlookup["Place of Delivery"])
    df_master_vlookup["ODC"] = np.where((df_master_vlookup['Destination Name per pre-alert'].notnull()) & (df_master_vlookup['Destination Name per pre-alert'].str.lower().str.contains("drdc")), "YES", df_master_vlookup["ODC"])
    df_master_vlookup["ODC"] = np.where((df_master_vlookup['Place of Delivery'].notnull()) & (df_master_vlookup['Place of Delivery'].str.lower().str.contains("drdc")), "YES", df_master_vlookup["ODC"])
except:
    logging.exception("\n \n \n Error logged: ODC GIT \n")
#################################################################################################################################
# Part Level - Expected Outgate Date From Destination Port
try:
    df_part_level = pd.concat([pd.read_csv(f, engine='python') for f in glob.glob(path + "\\input_files\\Part_level_report.csv")])
    df_part_level.dropna(subset=["Departed Origin Port MTI"], inplace=True)
    df_part_level.drop_duplicates(subset="Departed Origin Port MTI", inplace=True)
    df_part_level.rename(columns={"Departed Origin Port MTI": "Departed Origin Port MTI_part_level",
                                  "Outgated Expected Date/Time": "Outgated Expected Date/Time_part_level"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_part_level[["Departed Origin Port MTI_part_level", "Outgated Expected Date/Time_part_level"]], left_on="DepartedonShipMTI", right_on="Departed Origin Port MTI_part_level", how='left')
    df_master_vlookup["Expected Outgate Date From Destination Port"] = pd.to_datetime(df_master_vlookup['Expected Outgate Date From Destination Port']).dt.normalize()
    df_master_vlookup["Outgated Expected Date/Time_part_level"] = pd.to_datetime(df_master_vlookup['Outgated Expected Date/Time_part_level']).dt.normalize()
    df_master_vlookup["Expected Outgate Date From Destination Port"] = np.where(df_master_vlookup['Outgated Expected Date/Time_part_level'].notnull(), df_master_vlookup['Outgated Expected Date/Time_part_level'], df_master_vlookup["Expected Outgate Date From Destination Port"])
except:
    logging.exception("\n \n \n Error logged: Part Level \n")
#################################################################################################################################
# shipment report - Expected Outgate Date From Destination Port
try:
    df_shipment_report = pd.concat([pd.read_excel(f) for f in glob.glob(path + "\\input_files\\Shipment_Report*.xlsx")])
    df_shipment_report.dropna(subset=["MTI"], inplace=True)
    df_shipment_report.drop_duplicates(subset="MTI", inplace=True)
    df_shipment_report.rename(columns={"MTI": "MTI_shipment_report",
                                       "ESTIMATED_DEPARTURE_FROM_PORT_OF_DISCHARGE": "ESTIMATED_DEPARTURE_FROM_PORT_OF_DISCHARGE"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_shipment_report[["MTI_shipment_report", "ESTIMATED_DEPARTURE_FROM_PORT_OF_DISCHARGE"]], left_on="DepartedonShipMTI", right_on="MTI_shipment_report", how='left')
    df_master_vlookup["Expected Outgate Date From Destination Port"] = pd.to_datetime(df_master_vlookup['Expected Outgate Date From Destination Port']).dt.normalize()
    df_master_vlookup["ESTIMATED_DEPARTURE_FROM_PORT_OF_DISCHARGE"] = pd.to_datetime(df_master_vlookup['ESTIMATED_DEPARTURE_FROM_PORT_OF_DISCHARGE']).dt.normalize()
    df_master_vlookup["Expected Outgate Date From Destination Port"] = np.where((df_master_vlookup["Expected Outgate Date From Destination Port"].isnull()) & (df_master_vlookup['ESTIMATED_DEPARTURE_FROM_PORT_OF_DISCHARGE'].notnull()), df_master_vlookup['ESTIMATED_DEPARTURE_FROM_PORT_OF_DISCHARGE'], df_master_vlookup["Expected Outgate Date From Destination Port"])
except:
    logging.exception("\n \n \n Error logged: shipment report \n")
#################################################################################################################################
# Transload (Y/N) on water file
try:
    df_on_water = pd.read_excel(path + "\\input_files\\2020 INTERNAL - CNWW Compiled Vessel File.xlsx")
    df_on_water.dropna(subset=["MTI"], inplace=True)
    df_on_water.drop_duplicates(subset="MTI", inplace=True)
    df_on_water.rename(columns={"MTI": "MTI_on_water", "Transload (Yes/No)": "Transload (Yes/No)_on_water"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_on_water[["MTI_on_water", "Transload (Yes/No)_on_water"]], left_on="DepartedonShipMTI", right_on="MTI_on_water", how='left')
    df_master_vlookup["Transload"] = np.where(df_master_vlookup['Transload (Yes/No)_on_water'].notnull(), df_master_vlookup['Transload (Yes/No)_on_water'].astype(str).str.upper(), "NO")
except:
    logging.exception("\n \n \n Error logged: Transload (Y/N) on water file \n")
#################################################################################################################################
# Demurrage report
try:
    df_demurrage = pd.concat([pd.read_csv(f, engine="python") for f in glob.glob(path + "\\input_files\\demurrage_file*.csv")])
    df_demurrage.dropna(subset=["ContainerID"], inplace=True)
    df_demurrage.drop_duplicates(subset="ContainerID", inplace=True)

    df_demurrage.rename(columns={"APPT Date": "APPT Date_demmurage", "POD": "POD_demmurage",
                                 "Comment": "Comment_demmurage",
                                 "Door move/Trucker move": "Door move/Trucker move_demmurage",
                                 "Switched to Universal": "Switched to Universal_demmurage",
                                 "Plant receiving status": "Plant receiving status_demmurage",
                                 "ContainerID": "ContainerID_demmurage"}, inplace=True)

    df_master_vlookup = df_master_vlookup.merge(df_demurrage[["ContainerID_demmurage", "APPT Date_demmurage", "POD_demmurage", "Comment_demmurage",
                                                              "Door move/Trucker move_demmurage", "Switched to Universal_demmurage",
                                                              "Plant receiving status_demmurage"]], left_on="ContainerID", right_on="ContainerID_demmurage", how='left')

    df_master_vlookup["APPT Date_demmurage"] = pd.to_datetime(df_master_vlookup['APPT Date_demmurage'], errors='coerce')
    df_master_vlookup["POD_demmurage"] = pd.to_datetime(df_master_vlookup['POD_demmurage'], errors='coerce')
    df_master_vlookup["APPT Source"] = np.where((df_master_vlookup['APPT Date_demmurage'].notnull()) & (df_master_vlookup["ContainerID_demmurage"] == df_master_vlookup['ContainerID']), "Demurrage", df_master_vlookup["APPT Source"])
    df_master_vlookup["APPT"] = np.where((df_master_vlookup['APPT Date_demmurage'].notnull()) & (df_master_vlookup["ContainerID_demmurage"] == df_master_vlookup['ContainerID']), df_master_vlookup['APPT Date_demmurage'], df_master_vlookup["APPT"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['POD_demmurage'].notnull()), df_master_vlookup['POD_demmurage'], df_master_vlookup["POD"])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup['POD Source'].isnull()) & (df_master_vlookup['POD_demmurage'].notnull()), "Demurrage", df_master_vlookup["POD Source"])
    df_master_vlookup["Demurrage Comments"] = np.where((df_master_vlookup['Comment_demmurage'].notnull()), df_master_vlookup['Comment_demmurage'], df_master_vlookup["Demurrage Comments"])
    df_master_vlookup["Remarks"] = np.where((df_master_vlookup['Plant receiving status_demmurage'].notnull()), df_master_vlookup['Plant receiving status_demmurage'], df_master_vlookup["Remarks"])
    df_master_vlookup["Type of Move for Ceva shipments"] = np.where((df_master_vlookup['Door move/Trucker move_demmurage'].notnull()), df_master_vlookup['Door move/Trucker move_demmurage'], df_master_vlookup["Type of Move for Ceva shipments"])
    df_master_vlookup["Type of Move for Ceva shipments"] = np.where((df_master_vlookup['Switched to Universal_demmurage'].notnull()), df_master_vlookup['Switched to Universal_demmurage'], df_master_vlookup["Type of Move for Ceva shipments"])
except:
    logging.exception("\n \n \n Error logged: Demurrage report \n")
#################################################################################################################################
# Plant APPT
try:
    df_master_vlookup["APPT"] = pd.to_datetime(df_master_vlookup['APPT'], errors='coerce')
    df_master_vlookup['Master_plant_pod_combination'] = df_master_vlookup['PLT'].astype(str) + df_master_vlookup['InvoiceNumber'].astype(str)
    df_plant_pod_appt = pd.concat([pd.read_excel(f, sheet_name="Appt") for f in glob.glob(path + "\\input_files\\Plant POD*.xlsx")])
    df_plant_pod_appt["Combination_plant_appt"] = df_plant_pod_appt['PLANT CODE'].astype(str) + df_plant_pod_appt["SHPR ID"].astype(str)
    df_plant_pod_appt.dropna(subset=["Combination_plant_appt"], inplace=True)
    df_plant_pod_appt.drop_duplicates(subset="Combination_plant_appt", inplace=True)
    df_plant_pod_appt.rename(columns={"APPT": "APPT_plant_pod_appt"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_plant_pod_appt[["Combination_plant_appt", "APPT_plant_pod_appt"]], left_on="Master_plant_pod_combination", right_on="Combination_plant_appt", how='left')
    df_master_vlookup["APPT_plant_pod_appt"] = pd.to_datetime(df_master_vlookup['APPT_plant_pod_appt'], errors='coerce')
    df_master_vlookup["APPT Source"] = np.where((df_master_vlookup['APPT_plant_pod_appt'].notnull()) & (df_master_vlookup["APPT"] != df_master_vlookup['APPT_plant_pod_appt']), df_master_vlookup['PLT'], df_master_vlookup["APPT Source"])
    df_master_vlookup["APPT"] = np.where((df_master_vlookup['APPT_plant_pod_appt'].notnull()) & (df_master_vlookup["APPT"] != df_master_vlookup['APPT_plant_pod_appt']), df_master_vlookup['APPT_plant_pod_appt'], df_master_vlookup["APPT"])
except:
    logging.exception("\n \n \n Error logged: Plant APPT \n")
#################################################################################################################################
# Plant POD
try:
    df_plant_pod_pod = pd.concat([pd.read_excel(f, sheet_name="Delivered") for f in glob.glob(path + "\\input_files\\Plant POD*.xlsx")])
    df_plant_pod_pod["Combination_plant_pod"] = df_plant_pod_pod['PLANT CODE'].astype(str) + df_plant_pod_pod["SHPR ID"].astype(str)
    df_plant_pod_pod.dropna(subset=["Combination_plant_pod"], inplace=True)
    df_plant_pod_pod.drop_duplicates(subset="Combination_plant_pod", inplace=True)
    df_plant_pod_pod.rename(columns={"POD": "POD_plant_pod_pod"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_plant_pod_pod[["Combination_plant_pod", "POD_plant_pod_pod"]], left_on="Master_plant_pod_combination", right_on="Combination_plant_pod", how='left')
    df_master_vlookup["POD_plant_pod_pod"] = pd.to_datetime(df_master_vlookup['POD_plant_pod_pod'], errors='coerce')
    df_master_vlookup["POD"] = np.where((df_master_vlookup["POD"].isnull()) & (df_master_vlookup['POD_plant_pod_pod'].notnull()), df_master_vlookup['POD_plant_pod_pod'], df_master_vlookup["POD"])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup["POD Source"].isnull()) & (df_master_vlookup['POD_plant_pod_pod'].notnull()), df_master_vlookup['PLT'], df_master_vlookup["POD Source"])
except:
    logging.exception("\n \n \n Error logged: Plant POD \n")
    #################################################################################################################################
    # Pyramid Lines open tab
df_master_vlookup["master_pyramid_key"] = df_master_vlookup["ContainerID"].astype(str) + df_master_vlookup["HBL#"].astype(str)
try:
    df_pyrd_lines = pd.concat([pd.read_excel(f, sheet_name="Open") for f in glob.glob(path + "\\input_files\\PYRD Moves FCL and LCL*.xlsx")])
    df_pyrd_lines["PYRD_pyramid_key"] = df_pyrd_lines["CONTAINER"].astype(str) + df_pyrd_lines["BL#"].astype(str)
    df_pyrd_lines.dropna(subset=["PYRD_pyramid_key"], inplace=True)
    df_pyrd_lines.drop_duplicates(subset="PYRD_pyramid_key", inplace=True)
    df_pyrd_lines.rename(columns={"SSL Door Or Ramp Move": "SSL Door Or Ramp Move_pyrd",
                                  "DELIVERED DATE": "DELIVERED DATE_pyrd"}, inplace=True)

    df_master_vlookup = df_master_vlookup.merge(df_pyrd_lines[["PYRD_pyramid_key", "SSL Door Or Ramp Move_pyrd", "DELIVERED DATE_pyrd"]], left_on="master_pyramid_key", right_on="PYRD_pyramid_key", how='left')
    df_master_vlookup["DELIVERED DATE_pyrd"] = pd.to_datetime(df_master_vlookup['DELIVERED DATE_pyrd'], errors='coerce')
    df_master_vlookup["Type of Move for Ceva shipments"] = np.where(df_master_vlookup['SSL Door Or Ramp Move_pyrd'].notnull(), df_master_vlookup['SSL Door Or Ramp Move_pyrd'].astype(str).str.upper(), df_master_vlookup["Type of Move for Ceva shipments"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup['DELIVERED DATE_pyrd'].notnull()), df_master_vlookup['DELIVERED DATE_pyrd'], df_master_vlookup['POD'])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup['DELIVERED DATE_pyrd'].notnull()), "PYRAMID", df_master_vlookup['POD Source'])
except:
    logging.exception("\n \n \n Error logged: Pyramid Lines open tab \n")

# Pyramid Lines delivered tab
try:
    df_pyrd_DLVD_lines = pd.concat([pd.read_excel(f, sheet_name="Dlvd 2020") for f in glob.glob(path + "\\input_files\\PYRD Moves FCL and LCL*.xlsx")])
    df_pyrd_DLVD_lines["PYRD_pyramid_key_DLVD"] = df_pyrd_DLVD_lines["CONTAINER"].astype(str) + df_pyrd_DLVD_lines["BL#"].astype(str)
    df_pyrd_DLVD_lines.dropna(subset=["PYRD_pyramid_key_DLVD"], inplace=True)
    df_pyrd_DLVD_lines.drop_duplicates(subset="PYRD_pyramid_key_DLVD", inplace=True)
    df_pyrd_DLVD_lines.rename(columns={"SSL Door Or Ramp move": "SSL Door Or Ramp Move_pyrd_DLVD",
                                       "DELIVERED DATE": "DELIVERED DATE_pyrd_DLVD"}, inplace=True)

    df_master_vlookup = df_master_vlookup.merge(df_pyrd_DLVD_lines[["PYRD_pyramid_key_DLVD", "SSL Door Or Ramp Move_pyrd_DLVD", "DELIVERED DATE_pyrd_DLVD"]], left_on="master_pyramid_key", right_on="PYRD_pyramid_key_DLVD", how='left')
    df_master_vlookup["DELIVERED DATE_pyrd_DLVD"] = pd.to_datetime(df_master_vlookup['DELIVERED DATE_pyrd_DLVD'], errors='coerce')
    df_master_vlookup["Type of Move for Ceva shipments"] = np.where(df_master_vlookup['SSL Door Or Ramp Move_pyrd_DLVD'].notnull(), df_master_vlookup['SSL Door Or Ramp Move_pyrd_DLVD'].astype(str).str.upper(), df_master_vlookup["Type of Move for Ceva shipments"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup['DELIVERED DATE_pyrd_DLVD'].notnull()), df_master_vlookup['DELIVERED DATE_pyrd_DLVD'], df_master_vlookup['POD'])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup['DELIVERED DATE_pyrd_DLVD'].notnull()), "PYRAMID", df_master_vlookup['POD Source'])
except:
    logging.exception("\n \n \n Error logged: Pyramid Lines delivered tab \n")
#################################################################################################################################
# Invoice Log - Open Tab
try:
    df_invoice_log = pd.concat([pd.read_csv(f, engine="python") for f in glob.glob(path + "\\input_files\\GM_Master_invoice_log_open.csv")])
    df_invoice_log.dropna(subset=["File#"], inplace=True)
    df_invoice_log.drop_duplicates(subset="File#", inplace=True)
    df_invoice_log.rename(columns={"File#": "File#_invoice_log", "Path": "Path_invoice_log"}, inplace=True)
    df_invoice_log = df_invoice_log[df_invoice_log["Path_invoice_log"] != "FCL"]
    df_invoice_log = df_invoice_log[df_invoice_log["Path_invoice_log"] != "PLCL"]
    df_master_vlookup = df_master_vlookup.merge(df_invoice_log[["File#_invoice_log", "Path_invoice_log"]], left_on="File#", right_on="File#_invoice_log", how='left')
    df_master_vlookup["Path/Consol Center"] = np.where(df_master_vlookup['Path_invoice_log'].notnull(), df_master_vlookup["Path_invoice_log"], df_master_vlookup["Path/Consol Center"])
except:
    logging.exception("\n \n \n Error logged: Invoice Log - Open Tab \n")

# Invoice Log - Delivered
try:
    df_invoice_log_dlvd = pd.concat([pd.read_csv(f, engine="python") for f in glob.glob(path + "\\input_files\\GM_Master_invoice_log_delivered.csv")])
    df_invoice_log_dlvd.dropna(subset=["FILENUMBER"], inplace=True)
    df_invoice_log_dlvd.drop_duplicates(subset="FILENUMBER", inplace=True)
    df_invoice_log_dlvd.rename(columns={"FILENUMBER": "File#_invoice_log_dlvd", "Air/Ocean": "Path_invoice_log_dlvd"}, inplace=True)
    df_invoice_log_dlvd = df_invoice_log_dlvd[df_invoice_log_dlvd["Path_invoice_log_dlvd"] != "FCL"]
    df_invoice_log_dlvd = df_invoice_log_dlvd[df_invoice_log_dlvd["Path_invoice_log_dlvd"] != "PLCL"]
    df_master_vlookup = df_master_vlookup.merge(df_invoice_log_dlvd[["File#_invoice_log_dlvd", "Path_invoice_log_dlvd"]], left_on="File#", right_on="File#_invoice_log_dlvd", how='left')
    df_master_vlookup["Path/Consol Center"] = np.where(df_master_vlookup['Path_invoice_log_dlvd'].notnull(), df_master_vlookup["Path_invoice_log_dlvd"], df_master_vlookup["Path/Consol Center"])
except:
    logging.exception("\n \n \n Error logged: Invoice Log - Delivered \n")
#################################################################################################################################
# Oneview log
try:
    df_oneview_log = pd.concat([pd.read_excel(f) for f in glob.glob(path + "\\input_files\\oneview.xls")])
    df_oneview_log.dropna(subset=["Container Number"], inplace=True)
    df_oneview_log.drop_duplicates(subset="Container Number", inplace=True)
    df_oneview_log.rename(columns={"Container Number": "Container Number_oneview_log", "Load Type": "Load Type_oneview_log"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_oneview_log[["Container Number_oneview_log", "Load Type_oneview_log", "Master BOL"]], left_on="ContainerID", right_on="Container Number_oneview_log", how='left')
    df_master_vlookup["Path Type"] = np.where((df_master_vlookup['Path Type'].isnull()) & (df_master_vlookup['Load Type_oneview_log'].notnull()), df_master_vlookup["Load Type_oneview_log"], df_master_vlookup["Path Type"])
    df_master_vlookup["CNWW BOL #"] = np.where((df_master_vlookup['CNWW BOL #'].isnull()) & (df_master_vlookup['Master BOL'].notnull()) & (df_master_vlookup['Master BOL'].astype(str).str.contains("CNWW")), df_master_vlookup["Master BOL"], df_master_vlookup["CNWW BOL #"])
    df_master_vlookup["HBL#"] = np.where((df_master_vlookup['HBL#'].isnull()) & (df_master_vlookup['Master BOL'].notnull()) & (df_master_vlookup['Master BOL'].astype(str).str.contains("CNWW")), df_master_vlookup["Master BOL"], df_master_vlookup["HBL#"])
    df_master_vlookup["MBL#"] = np.where((df_master_vlookup['MBL#'].isnull()) & (df_master_vlookup['Master BOL'].notnull()) & (~df_master_vlookup['Master BOL'].astype(str).str.contains("CNWW")), df_master_vlookup["Master BOL"], df_master_vlookup["MBL#"])
except:
    logging.exception("\n \n \n Error logged: Oneview log \n")
#################################################################################################################################
# Direct_truck
try:
    df_direct_truck = pd.concat([pd.read_excel(f) for f in glob.glob(path + "\\input_files\\DIRECT TRUCK LOG*.xlsx")])
    df_direct_truck.dropna(subset=["Container"], inplace=True)
    df_direct_truck.drop_duplicates(subset="Container", inplace=True)
    df_direct_truck.rename(columns={"Container": "Container_direct_truck", "STATUS": "STATUS_direct_truck"}, inplace=True)
    status_list = ["at delivery", "in transit", "storage", "delivered"]
    df_direct_truck = df_direct_truck[~df_direct_truck["STATUS_direct_truck"].str.lower().str.contains('|'.join(status_list))]
    df_master_vlookup = df_master_vlookup.merge(df_direct_truck[["Container_direct_truck", "STATUS_direct_truck"]], left_on="ContainerID", right_on="Container_direct_truck", how='left')
    df_master_vlookup["Transload"] = np.where((df_master_vlookup['Container_direct_truck'].notnull()), "NO", df_master_vlookup["Transload"])
    df_master_vlookup["Rail/DT"] = np.where((df_master_vlookup['Container_direct_truck'].notnull()), "DIRECT TRUCK", df_master_vlookup['Rail/DT'])
    df_master_vlookup["Outgate but not departed"] = np.where((df_master_vlookup['Container_direct_truck'].notnull()) & (df_master_vlookup['Actual Outgate from Port'].notnull()), df_master_vlookup["Actual Outgate from Port"], df_master_vlookup["Outgate but not departed"])
    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['Container_direct_truck'].notnull()) & (df_master_vlookup['Actual Outgate from Port'].notnull()), np.datetime64('NaT'), df_master_vlookup["Actual Outgate from Port"])
    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['Container_direct_truck'].notnull()) & (df_master_vlookup['Actual Outgate from Port'].notnull()), np.datetime64('NaT'), df_master_vlookup["Actual Outgate from Port"])

    df_direct_truck_comment = pd.concat([pd.read_excel(f) for f in glob.glob(path + "\\input_files\\DIRECT TRUCK LOG*.xlsx")])
    df_direct_truck_comment.dropna(subset=["Container"], inplace=True)
    df_direct_truck_comment.drop_duplicates(subset="Container", inplace=True)
    df_direct_truck_comment['COMMENTS'] = df_direct_truck_comment['COMMENTS'].astype(str).str.split(";", 1)
    df_direct_truck_comment['COMMENTS'] = df_direct_truck_comment['COMMENTS'].str[0] + ";"
    df_direct_truck_comment.rename(columns={"Container": "Container_direct_truck_comment", "COMMENTS": "COMMENTS_direct_truck"}, inplace=True)
    df_master_vlookup = df_master_vlookup.merge(df_direct_truck_comment[["Container_direct_truck_comment", "COMMENTS_direct_truck"]], left_on="ContainerID", right_on="Container_direct_truck_comment", how='left')
    df_master_vlookup["Last Staus on Rail"] = np.where((df_master_vlookup['COMMENTS_direct_truck'].notnull()), (df_master_vlookup['COMMENTS_direct_truck']), df_master_vlookup["Last Staus on Rail"])

except:
    logging.exception("\n \n \n Error logged: Direct_truck \n")

#################################################################################################################################
#################################################################################################################################
# rail exception
df_master_vlookup["Event status per railinc"] = ""
df_master_vlookup["For ocean exception Event Date"] = ""
do_not_take_city = ["vancouver", "halifax", "prince rupert", "montreal", "surrey", "robert"]

#################################################################################################################################
# cn_direct_intact
try:
    df_master_vlookup["For ocean exception Event Date"] = pd.to_datetime(df_master_vlookup['For ocean exception Event Date'], errors='coerce')
    df_master_vlookup["ETA Destination Rail"] = pd.to_datetime(df_master_vlookup['ETA Destination Rail'], errors='coerce')
    df_master_vlookup["Date Arrived Destination Rail"] = pd.to_datetime(df_master_vlookup['Date Arrived Destination Rail'], errors='coerce')

    df_cn_direct_intact = pd.concat([pd.read_excel(f, sheet_name="CN-Direct Intact") for f in glob.glob(path + "\\input_files\\Master rail*.xlsx")])
    df_cn_direct_intact.dropna(subset=["Container"], inplace=True)
    df_cn_direct_intact.drop_duplicates(subset="Container", inplace=True)
    df_cn_direct_intact.rename(columns={"Container": "Container_cn_direct_intact", "Destination City": "Destination City_cn_direct_intact",
                                        "Comment": "Comment_cn_direct_intact", "ETA Date": "ETA Date_cn_direct_intact", "Date": "Date_cn_direct_intact",
                                        "Event": "Event_cn_direct_intact", "City": "City_cn_direct_intact", "Load": "Load__cn_direct_intact"}, inplace=True)

    df_master_vlookup = df_master_vlookup.merge(df_cn_direct_intact[["Container_cn_direct_intact", "Destination City_cn_direct_intact", "Comment_cn_direct_intact", "ETA Date_cn_direct_intact",
                                                                     "Date_cn_direct_intact", "Event_cn_direct_intact", "City_cn_direct_intact", "Load__cn_direct_intact"]], left_on="ContainerID", right_on="Container_cn_direct_intact", how='left')

    df_master_vlookup["ETA Date_cn_direct_intact"] = pd.to_datetime(df_master_vlookup['ETA Date_cn_direct_intact'], errors="coerce").dt.normalize()
    df_master_vlookup["Date_cn_direct_intact"] = pd.to_datetime(df_master_vlookup['Date_cn_direct_intact'], errors="coerce").dt.normalize()

    df_master_vlookup["Direct Intact"] = np.where((df_master_vlookup['Container_cn_direct_intact'].notnull()), "YES", df_master_vlookup['Direct Intact'])
    df_master_vlookup["CNRU#"] = np.where((df_master_vlookup['Container_cn_direct_intact'].notnull()), "Direct Intact", df_master_vlookup['CNRU#'])
    df_master_vlookup["Rail/DT"] = np.where((df_master_vlookup['Container_cn_direct_intact'].notnull()), "CN", df_master_vlookup['Rail/DT'])

    df_master_vlookup["Destination City_cn_direct_intact"] = np.where((df_master_vlookup['Destination City_cn_direct_intact'].notnull()) & (df_master_vlookup['Destination City_cn_direct_intact'].str.lower().str.contains("detroit")), "detroit", df_master_vlookup["Destination City_cn_direct_intact"])
    df_master_vlookup["City_cn_direct_intact"] = np.where((df_master_vlookup['City_cn_direct_intact'].notnull()) & (df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("detroit")), "detroit", df_master_vlookup["City_cn_direct_intact"])

    df_master_vlookup["City / Rail Yard Name"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Destination City_cn_direct_intact'].notnull()) & (~df_master_vlookup['Destination City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)), df_master_vlookup["Destination City_cn_direct_intact"], df_master_vlookup["City / Rail Yard Name"])
    df_master_vlookup["Last Staus on Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Comment_cn_direct_intact'].notnull()), df_master_vlookup["Comment_cn_direct_intact"], df_master_vlookup["Last Staus on Rail"])
    df_master_vlookup["ETA Destination Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Date Arrived Destination Rail"].isnull()) & (df_master_vlookup['ETA Date_cn_direct_intact'].notnull()), df_master_vlookup["ETA Date_cn_direct_intact"], df_master_vlookup["ETA Destination Rail"])
    df_master_vlookup["Event status per railinc"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Event_cn_direct_intact'].notnull()), df_master_vlookup["Event_cn_direct_intact"], df_master_vlookup["Event status per railinc"])
    df_master_vlookup["For ocean exception Event Date"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Date_cn_direct_intact'].notnull()), df_master_vlookup["Date_cn_direct_intact"], df_master_vlookup["For ocean exception Event Date"])
    df_master_vlookup["Date Arrived Destination Rail"] = np.where(
        (df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Load__cn_direct_intact'] == "Load") & (df_master_vlookup['Event_cn_direct_intact'].str.lower().str.contains("at destination|deramped|pad placement|out-gated|notified|arrived")) & (~df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup['Date_cn_direct_intact'].notnull()) & (df_master_vlookup["City_cn_direct_intact"] == df_master_vlookup["Destination "
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              "City_cn_direct_intact"]),
        df_master_vlookup["Date_cn_direct_intact"], df_master_vlookup["Date Arrived Destination Rail"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() != "PLCL") & (df_master_vlookup['Load__cn_direct_intact'] == "Empty") & (df_master_vlookup['Date_cn_direct_intact'].notnull()) & (~df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup["City_cn_direct_intact"] == df_master_vlookup["Destination City_cn_direct_intact"]),
                                        df_master_vlookup["Date_cn_direct_intact"], df_master_vlookup["POD"])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() != "PLCL") & (df_master_vlookup['Load__cn_direct_intact'] == "Empty") & (df_master_vlookup['Date_cn_direct_intact'].notnull()) & (~df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup["City_cn_direct_intact"] == df_master_vlookup["Destination City_cn_direct_intact"]), "CN Rail",
                                               df_master_vlookup["POD Source"])

    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Actual Outgate from Port"].isnull()) & (df_master_vlookup['Load__cn_direct_intact'] == "Load") & (df_master_vlookup['Event_cn_direct_intact'] == "Arrived") & (df_master_vlookup['Date_cn_direct_intact'].notnull()) & (~df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)),
                                                             df_master_vlookup["Date_cn_direct_intact"], df_master_vlookup["Actual Outgate from Port"])
    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Actual Outgate from Port"].isnull()) & (df_master_vlookup['Load__cn_direct_intact'] == "Load") & (df_master_vlookup['Event_cn_direct_intact'] == "Departed") & (df_master_vlookup['Date_cn_direct_intact'].notnull()), df_master_vlookup["Date_cn_direct_intact"], df_master_vlookup["Actual Outgate from Port"])

    df_master_vlookup["Outgate but not departed"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Outgate but not departed"].isnull()) & (df_master_vlookup['Load__cn_direct_intact'] == "Load") & (df_master_vlookup['Event_cn_direct_intact'].str.lower().str.contains("arrive|ramped|flatcar", na=False)) & (df_master_vlookup['Date_cn_direct_intact'].notnull()) & (df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)),
                                                             df_master_vlookup["Date_cn_direct_intact"],
                                                             df_master_vlookup["Outgate but not departed"])
    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup['Load__cn_direct_intact'] == "Load") & (df_master_vlookup['Event_cn_direct_intact'].str.lower().str.contains("arrive|ramped|flatcar", na=False)) & (df_master_vlookup['Date_cn_direct_intact'].notnull()) & (df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)),
                                                             np.datetime64("NaT"),
                                                             df_master_vlookup["Actual Outgate from Port"])

    df_master_vlookup["Actual Departure from Dest Rail"] = np.where((df_master_vlookup["Actual Departure from Dest Rail"].isnull()) & (df_master_vlookup['Load__cn_direct_intact'] == "Load") & (df_master_vlookup['Event_cn_direct_intact'] == "Out-Gate") & (df_master_vlookup["City_cn_direct_intact"] == df_master_vlookup["Destination City_cn_direct_intact"]), df_master_vlookup["Date_cn_direct_intact"], df_master_vlookup["Actual Departure from Dest Rail"])
except:
    logging.exception("\n \n \n Error logged: cn_direct_intact \n")
# #################################################################################################################################
# CN Rail
try:
    df_master_vlookup["For ocean exception Event Date"] = pd.to_datetime(df_master_vlookup['For ocean exception Event Date'], errors='coerce')
    df_master_vlookup["ETA Destination Rail"] = pd.to_datetime(df_master_vlookup['ETA Destination Rail'], errors='coerce')
    df_master_vlookup["Date Arrived Destination Rail"] = pd.to_datetime(df_master_vlookup['Date Arrived Destination Rail'], errors='coerce')

    df_CN = pd.concat([pd.read_excel(f, sheet_name="CNRU") for f in glob.glob(path + "\\input_files\\Master rail*.xlsx")])
    df_CN.dropna(subset=["CNRU"], inplace=True)
    df_CN.drop_duplicates(subset="CNRU", inplace=True)
    df_CN.rename(columns={"CNRU": "CNRU_CN", "Destination City": "Destination City_CN", "Comment": "Comment_CN",
                          "ETA Date": "ETA Date_CN", "Date": "Date_CN", "Event": "Event_CN", "City": "City_CN", "Load": "Load_CN"}, inplace=True)

    df_master_vlookup = df_master_vlookup.merge(df_CN[["CNRU_CN", "Destination City_CN", "Comment_CN", "ETA Date_CN",
                                                       "Date_CN", "Event_CN", "City_CN", "Load_CN"]],
                                                left_on="CNRU#", right_on="CNRU_CN", how='left')

    df_master_vlookup["ETA Date_CN"] = pd.to_datetime(df_master_vlookup['ETA Date_CN'], errors="coerce").dt.normalize()
    df_master_vlookup["Date_CN"] = pd.to_datetime(df_master_vlookup['Date_CN'], errors="coerce").dt.normalize()

    df_master_vlookup["Destination City_CN"] = np.where((df_master_vlookup['Destination City_CN'].notnull()) & (df_master_vlookup['Destination City_CN'].str.lower().str.contains("detroit")), "detroit", df_master_vlookup["Destination City_CN"])
    df_master_vlookup["City_CN"] = np.where((df_master_vlookup['City_CN'].notnull()) & (df_master_vlookup['City_CN'].str.lower().str.contains("detroit")), "detroit", df_master_vlookup["City_CN"])

    df_master_vlookup["City / Rail Yard Name"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Destination City_CN'].notnull()) & (~df_master_vlookup['Destination City_CN'].str.lower().str.contains("|".join(do_not_take_city), na=False)), df_master_vlookup["Destination City_CN"], df_master_vlookup["City / Rail Yard Name"])
    df_master_vlookup["Last Staus on Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Comment_CN'].notnull()), df_master_vlookup["Comment_CN"], df_master_vlookup["Last Staus on Rail"])
    df_master_vlookup["ETA Destination Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Date Arrived Destination Rail"].isnull()) & (df_master_vlookup['ETA Date_CN'].notnull()), df_master_vlookup["ETA Date_CN"], df_master_vlookup["ETA Destination Rail"])
    df_master_vlookup["Event status per railinc"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Event_CN'].notnull()), df_master_vlookup["Event_CN"], df_master_vlookup["Event status per railinc"])
    df_master_vlookup["For ocean exception Event Date"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Date_CN'].notnull()), df_master_vlookup["Date_CN"], df_master_vlookup["For ocean exception Event Date"])
    df_master_vlookup["Date Arrived Destination Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Load_CN'] == "Load") & (df_master_vlookup['Event_CN'].str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup['Date_CN'].notnull()) & (df_master_vlookup['Event_CN'].str.lower().str.contains("at destination|deramped|pad placement|out-gated|notified|arrived")) & (df_master_vlookup["City_CN"] == df_master_vlookup["Destination City_CN"]),
                                                                  df_master_vlookup["Date_CN"],
                                                                  df_master_vlookup["Date Arrived Destination Rail"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() != "PLCL") & (df_master_vlookup['Load_CN'] == "Empty") & (df_master_vlookup['Date_CN'].notnull()) & (~df_master_vlookup['City_CN'].str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup["City_CN"] == df_master_vlookup["Destination City_CN"]), df_master_vlookup["Date_CN"], df_master_vlookup["POD"])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() != "PLCL") & (df_master_vlookup['Load_CN'] == "Empty") & (df_master_vlookup['Date_CN'].notnull()) & (~df_master_vlookup['City_CN'].str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup["City_CN"] == df_master_vlookup["Destination City_CN"]), "CN Rail", df_master_vlookup["POD Source"])

    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Actual Outgate from Port"].isnull()) & (df_master_vlookup['Load_CN'] == "Load") & (df_master_vlookup['Event_CN'] == "Arrived") & (df_master_vlookup['Date_CN'].notnull()) & (~df_master_vlookup['City_CN'].str.lower().str.contains("|".join(do_not_take_city), na=False)), df_master_vlookup["Date_CN"] - pd.Timedelta(1, unit='d'), df_master_vlookup["Actual Outgate from Port"])
    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Actual Outgate from Port"].isnull()) & (df_master_vlookup['Load_CN'] == "Load") & (df_master_vlookup['Event_CN'] == "Departed") & (df_master_vlookup['Date_CN'].notnull()), df_master_vlookup["Date_CN"], df_master_vlookup["Actual Outgate from Port"])
    df_master_vlookup["Outgate but not departed"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Outgate but not departed"].isnull()) & (df_master_vlookup['Load_CN'] == "Load") & (df_master_vlookup['Event_CN'].str.lower().str.contains("arrive|ramped|flatcar", na=False)) & (df_master_vlookup['Date_CN'].notnull()) & (df_master_vlookup['City_CN'].str.lower().str.contains("|".join(do_not_take_city), na=False)), df_master_vlookup["Date_CN"],
                                                             df_master_vlookup["Outgate but not departed"])
    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup['Load_CN'] == "Load") & (df_master_vlookup['Event_CN'].str.lower().str.contains("arrive|ramped|flatcar", na=False)) & (df_master_vlookup['Date_cn_direct_intact'].notnull()) & (df_master_vlookup['City_cn_direct_intact'].str.lower().str.contains("|".join(do_not_take_city), na=False)), np.datetime64("NaT"),
                                                             df_master_vlookup["Actual Outgate from Port"])

    df_master_vlookup["Actual Departure from Dest Rail"] = np.where((df_master_vlookup["Actual Departure from Dest Rail"].isnull()) & (df_master_vlookup['Load_CN'] == "Load") & (df_master_vlookup['Event_CN'] == "Out-Gate") & (df_master_vlookup["City_CN"] == df_master_vlookup["Destination City_CN"]), df_master_vlookup["Date_CN"], df_master_vlookup["Actual Departure from Dest Rail"])
except:
    logging.exception("\n \n \n Error logged: CN Rail \n")
# #################################################################################################################################
# cp_rail
try:
    df_cp_rail = pd.concat([pd.read_excel(f, sheet_name="CP") for f in glob.glob(path + "\\input_files\\Master rail*.xlsx")])
    df_cp_rail.dropna(subset=["Equipment"], inplace=True)
    df_cp_rail.drop_duplicates(subset="Equipment", inplace=True)
    df_cp_rail.rename(columns={"Equipment.1": "Equipment_cp_rail", "Current Position": "Current Position_cp_rail",
                               "Cons Point": "Cons Point_cp_rail", "Comment": "Comment_cp_rail",
                               "Act Arrival <i>(or Exp)</i>": "Act Arrival <i>(or Exp)</i>_cp_rail"}, inplace=True)

    df_cp_rail["Act Arrival <i>(or Exp)</i>_cp_rail"] = df_cp_rail["Act Arrival <i>(or Exp)</i>_cp_rail"].astype(str)
    df_cp_rail["ETA_cp_rail"] = ""
    df_cp_rail["ATA_cp_rail"] = ""
    df_cp_rail["ETA_cp_rail"] = np.where(df_cp_rail["Act Arrival <i>(or Exp)</i>_cp_rail"].str.contains("\(", na=False), df_cp_rail["Act Arrival <i>(or Exp)</i>_cp_rail"], np.nan)
    df_cp_rail["ATA_cp_rail"] = np.where(~df_cp_rail["Act Arrival <i>(or Exp)</i>_cp_rail"].str.contains("\(", na=False), df_cp_rail["Act Arrival <i>(or Exp)</i>_cp_rail"], np.nan)
    df_cp_rail["ETA_cp_rail"] = pd.to_datetime(df_cp_rail["ETA_cp_rail"].astype(str).str[1:-7])
    df_cp_rail["ATA_cp_rail"] = pd.to_datetime(df_cp_rail["ATA_cp_rail"].astype(str).str[:-8])

    df_master_vlookup = df_master_vlookup.merge(df_cp_rail[["Equipment_cp_rail", "Current Position_cp_rail", "Cons Point_cp_rail", "Comment_cp_rail",
                                                            "ETA_cp_rail", "ATA_cp_rail"]], left_on="ContainerID", right_on="Equipment_cp_rail", how='left')
    df_master_vlookup["Rail/DT"] = np.where((df_master_vlookup['Equipment_cp_rail'].notnull()), "CPRS", df_master_vlookup['Rail/DT'])
    df_master_vlookup["ETA_cp_rail"] = pd.to_datetime(df_master_vlookup['ETA_cp_rail'], errors='coerce')
    df_master_vlookup["ATA_cp_rail"] = pd.to_datetime(df_master_vlookup['ATA_cp_rail'], errors='coerce')
    df_master_vlookup["Last Staus on Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Comment_cp_rail'].notnull()), df_master_vlookup["Comment_cp_rail"], df_master_vlookup["Last Staus on Rail"])
    df_master_vlookup["City / Rail Yard Name"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Cons Point_cp_rail'].notnull()) & (~df_master_vlookup['Cons Point_cp_rail'].str.lower().str.contains("|".join(do_not_take_city), na=False)), df_master_vlookup["Cons Point_cp_rail"], df_master_vlookup["City / Rail Yard Name"])
    df_master_vlookup["ETA Destination Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Date Arrived Destination Rail"].isnull()) & (df_master_vlookup['ETA_cp_rail'].notnull()), df_master_vlookup["ETA_cp_rail"], df_master_vlookup["ETA Destination Rail"])
    df_master_vlookup["Date Arrived Destination Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['ATA_cp_rail'].notnull()) & (~df_master_vlookup['Current Position_cp_rail'].astype(str).str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup["Cons Point_cp_rail"] == df_master_vlookup["Current Position_cp_rail"]), df_master_vlookup["ATA_cp_rail"], df_master_vlookup["Date Arrived Destination Rail"])
except:
    logging.exception("\n \n \n Error logged: cp_rail \n")
################################################################################################################################
# Railinc
try:
    df_railinc = pd.concat([pd.read_excel(f, sheet_name="Raillinc") for f in glob.glob(path + "\\input_files\\Master rail*.xlsx")])

    df_railinc.dropna(subset=["Equipment ID"], inplace=True)
    df_railinc.drop_duplicates(subset="Equipment ID", inplace=True)
    df_railinc.rename(columns={"Equipment ID": "Equipment ID_railinc", "Destination City": "Destination City_railinc", "Comment": "Comment_railinc", "ETA time": "ETA time_railinc",
                               "Event Description": "Event Description_railinc", "Event Time": "Event Time_railinc", "Posting Mark": "Posting Mark_railinc", "Last Event Location": "Last Event Location_railinc"}, inplace=True)

    df_master_vlookup = df_master_vlookup.merge(df_railinc[["Equipment ID_railinc", "Destination City_railinc", "Comment_railinc", "ETA time_railinc", "Event Description_railinc",
                                                            "Event Time_railinc", "Posting Mark_railinc", "Last Event Location_railinc"]], left_on="ContainerID", right_on="Equipment ID_railinc", how='left')

    df_master_vlookup["ETA time_railinc"] = pd.to_datetime(df_master_vlookup['ETA time_railinc'], errors="coerce").dt.normalize()
    df_master_vlookup["Event Time_railinc"] = pd.to_datetime(df_master_vlookup['Event Time_railinc'], errors="coerce").dt.normalize()
    df_master_vlookup["City / Rail Yard Name"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Destination City_railinc'].notnull()) & (~df_master_vlookup['Destination City_railinc'].astype(str).str.lower().str.contains("|".join(do_not_take_city), na=False)), df_master_vlookup["Destination City_railinc"], df_master_vlookup["City / Rail Yard Name"])
    df_master_vlookup["Last Staus on Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Comment_railinc'].notnull()), df_master_vlookup["Comment_railinc"], df_master_vlookup["Last Staus on Rail"])
    df_master_vlookup["ETA Destination Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup["Date Arrived Destination Rail"].isnull()) & (df_master_vlookup['ETA time_railinc'].notnull()), df_master_vlookup["ETA time_railinc"], df_master_vlookup["ETA Destination Rail"])
    df_master_vlookup["Event status per railinc"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Event Description_railinc'].notnull()), df_master_vlookup["Event Description_railinc"], df_master_vlookup["Event status per railinc"])
    df_master_vlookup["For ocean exception Event Date"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Event Time_railinc'].notnull()), df_master_vlookup["Event Time_railinc"], df_master_vlookup["For ocean exception Event Date"])
    df_master_vlookup["Rail/DT"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Posting Mark_railinc'].notnull()), df_master_vlookup["Posting Mark_railinc"], df_master_vlookup["Rail/DT"])
    df_master_vlookup["Date Arrived Destination Rail"] = np.where((df_master_vlookup['POD'].isnull()) & (df_master_vlookup['Event Time_railinc'].notnull()) & (~df_master_vlookup['Last Event Location_railinc'].astype(str).str.lower().str.contains("|".join(do_not_take_city), na=False)) & (df_master_vlookup["Last Event Location_railinc"] == df_master_vlookup["Destination City_railinc"]), df_master_vlookup["Event Time_railinc"], df_master_vlookup["Date Arrived Destination Rail"])
except:
    logging.exception("\n \n \n Error logged: Railinc \n")
#################################################################################################################################

try:
    # if outgate is less then outgate from port
    df_master_vlookup["Actual Outgate from Port"] = np.where((df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup["Outgate but not departed"].notnull()) & (df_master_vlookup["Actual Outgate from Port"] < df_master_vlookup["Outgate but not departed"]), df_master_vlookup["Outgate but not departed"], df_master_vlookup["Actual Outgate from Port"])
    # if ATD Rail is empty update outgate in rail atd
    df_master_vlookup["ATD Rail"] = np.where((df_master_vlookup["ATD Rail"].isnull()) & (df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup["Rail/DT"].notnull()), df_master_vlookup["Actual Outgate from Port"], df_master_vlookup["ATD Rail"])
    # actual outgate from the destination rail
    df_master_vlookup["Actual Departure from Dest Rail"] = np.where((df_master_vlookup["Actual Departure from Dest Rail"].isnull()) & (df_master_vlookup['Load_CN'] == "Empty") & (df_master_vlookup['Event_CN'] == "Out-Gate") & (df_master_vlookup["City_CN"] == df_master_vlookup["Destination City_CN"]), df_master_vlookup["Date_CN"], df_master_vlookup["Actual Departure from Dest Rail"])

    # if POD is available and then these dates should be filled
    df_master_vlookup["Date Arrived Destination Rail"] = np.where((df_master_vlookup['Date Arrived Destination Rail'].isnull()) & (df_master_vlookup["Rail/DT"].notnull()) & (~df_master_vlookup["Rail/DT"].astype(str).str.lower().str.contains("direct truck", na=False)) & (df_master_vlookup["POD"].notnull()), (df_master_vlookup["POD"] - pd.Timedelta(1, unit='d')), df_master_vlookup["Date Arrived Destination Rail"])
    df_master_vlookup["Actual Departure from Dest Rail"] = np.where((df_master_vlookup['Actual Departure from Dest Rail'].isnull()) & (df_master_vlookup["Rail/DT"].notnull()) & (~df_master_vlookup["Rail/DT"].astype(str).str.lower().str.contains("direct truck", na=False)) & (df_master_vlookup["POD"].notnull()), (df_master_vlookup["POD"] - pd.Timedelta(1, unit='d')), df_master_vlookup["Actual Departure from Dest Rail"])
except:
    logging.exception("\n \n \n Error logged: Rail Exceptions \n")
#################################################################################################################################
# Final Check
# Updating bill of lading
# noinspection PyBroadException
try:
    # if Date Arrived Destination Rail is present then ETA rail should not be blnk
    df_master_vlookup["ETA Destination Rail"] = np.where((df_master_vlookup["ETA Destination Rail"].isnull()) & (df_master_vlookup["Date Arrived Destination Rail"].notnull()), df_master_vlookup["Date Arrived Destination Rail"], df_master_vlookup["ETA Destination Rail"])

    df_master_vlookup["Bill of Lading"] = np.where(df_master_vlookup["MBL#"].notnull(), df_master_vlookup["MBL#"], df_master_vlookup["Bill of Lading"])
    df_master_vlookup["Bill of Lading"] = np.where(df_master_vlookup["HBL#"].notnull(), df_master_vlookup["HBL#"], df_master_vlookup["Bill of Lading"])

    # Original carrier
    df_master_vlookup["Carrier"] = np.where((df_master_vlookup["HBL#"].notnull()) & (df_master_vlookup["HBL#"].astype(str).str.upper().str.contains("CNWW")), "CNWW", df_master_vlookup["Carrier"])
    df_master_vlookup["Carrier"] = np.where((df_master_vlookup["HBL#"].notnull()) & (~df_master_vlookup["HBL#"].astype(str).str.upper().str.contains("CNWW", na=False)), "CEVV", df_master_vlookup["Carrier"])
    df_master_vlookup["Carrier"] = np.where((df_master_vlookup["HBL#"].isnull()) & (df_master_vlookup["Carrier"].str.contains("CNWW|CEVV", regex=True)), np.nan, df_master_vlookup["Carrier"])
    df_master_vlookup["Carrier"] = np.where((df_master_vlookup["Original carrier"].isnull()) & (~df_master_vlookup["Carrier"].str.contains("CNWW|CEVV", regex=True, na=False)), np.nan, df_master_vlookup["Carrier"])
    df_master_vlookup["Carrier"] = np.where((df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"] == "PLCL"), "CEVV", df_master_vlookup["Carrier"])
    df_master_vlookup["Carrier"] = np.where((df_master_vlookup["Original carrier"].notnull()) & (~df_master_vlookup["Carrier"].str.contains("CNWW|CEVV", regex=True, na=False)), df_master_vlookup["Original carrier"], df_master_vlookup["Carrier"])

    # Contracted Carrier SCAC
    df_master_vlookup["Contracted Carrier SCAC"] = np.where((df_master_vlookup["Original carrier"].notnull()) & (df_master_vlookup["Carrier"].str.contains("CNWW|CEVV", regex=True)), df_master_vlookup["Original carrier"], df_master_vlookup["Contracted Carrier SCAC"])
    df_master_vlookup["Contracted Carrier BL#"] = np.where((df_master_vlookup["MBL#"].notnull()) & (df_master_vlookup["Carrier"].str.contains("CNWW|CEVV", regex=True)), df_master_vlookup["MBL#"], df_master_vlookup["Contracted Carrier BL#"])
    df_master_vlookup["CNWW BOL #"] = np.where((df_master_vlookup["Bill of Lading"].notnull()) & (df_master_vlookup["Bill of Lading"].astype(str).str.upper().str.contains("CNWW")), df_master_vlookup["Bill of Lading"], df_master_vlookup["CNWW BOL #"])

    df_port_list = pd.read_excel(path + "\\Input_Files\\Pol and Pod Alias.xlsx", sheet_name="ports")
    port_list = df_port_list["Full Port Name"].tolist()
    df_master_vlookup["Place of Delivery"] = np.where((df_master_vlookup["Place of Delivery"].notnull()) & (df_master_vlookup["Place of Delivery"].astype(str).str.upper().str.contains("|".join(port_list), na=False)), np.nan, df_master_vlookup["Place of Delivery"])
    df_master_vlookup["Place of Delivery"] = np.where((df_master_vlookup["Destination Name per pre-alert"].notnull()) & (df_master_vlookup["Destination Name per pre-alert"].astype(str).str.upper().str.contains("MICHIGAN CROSS")), "GM CCA C/O MICHIGAN CROSSDOCK", df_master_vlookup["Place of Delivery"])
    df_master_vlookup["Place of Delivery"] = np.where((df_master_vlookup["Destination Name per pre-alert"].notnull()) & (df_master_vlookup["Destination Name per pre-alert"].astype(str).str.upper().str.contains("DRDC")), "DRDC ILG Detroit Regional Distribution", df_master_vlookup["Place of Delivery"])

    df_ultimate_dest_name = pd.read_excel(path + "\\Input_Files\\Pol and Pod Alias.xlsx", sheet_name="BTC")
    df_ultimate_dest_name["BTC"] = df_ultimate_dest_name["BTC"].astype(str)
    df_master_vlookup["BillToCisco"] = df_master_vlookup["BillToCisco"].astype(str)
    df_master_vlookup = df_master_vlookup.merge(df_ultimate_dest_name[["BTC", "Plant name"]], left_on="BillToCisco", right_on="BTC", how="left")

    df_master_vlookup["Ultimate Dest. Name"] = np.where((df_master_vlookup["Ultimate Dest. Name"].isnull()) & (df_master_vlookup["Plant name"].notnull()), df_master_vlookup["Plant name"], df_master_vlookup["Ultimate Dest. Name"])
    df_master_vlookup["Place of Delivery"] = np.where((df_master_vlookup["Place of Delivery"].isnull()), df_master_vlookup["Destination Name per pre-alert"], df_master_vlookup["Place of Delivery"])
    df_master_vlookup["Place of Delivery"] = np.where((df_master_vlookup["Place of Delivery"].isnull()), df_master_vlookup["Ultimate Dest. Name"], df_master_vlookup["Place of Delivery"])
    df_master_vlookup["CRITERIA"] = np.where((df_master_vlookup["Place of Delivery"].notnull()) & (df_master_vlookup["Place of Delivery"].astype(str).str.upper().str.contains("MICHIGAN CROSS")) & (df_master_vlookup["Path Type"].astype(str).str.upper().str.contains("LCL")), "LCL through MCD: +20 days", df_master_vlookup["CRITERIA"])
    df_master_vlookup["Estimated Arrival Date To Origin Port"] = np.where((df_master_vlookup["Estimated Arrival Date To Origin Port"].isnull()), df_master_vlookup["ATA to POL"], df_master_vlookup["Estimated Arrival Date To Origin Port"])

    # CRITERIA
    df_master_vlookup["CRITERIA"] = np.where((df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "FCL"), "FCL + 14 days", df_master_vlookup["CRITERIA"])
    df_master_vlookup["CRITERIA"] = np.where((df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "LCL"), "LCL + 14 Days direct to consignee", df_master_vlookup["CRITERIA"])
    df_master_vlookup["CRITERIA"] = np.where((df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "FCL\\LCL"), "FCL\\LCL + 20 Days", df_master_vlookup["CRITERIA"])
    df_master_vlookup["CRITERIA"] = np.where((df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "PLCL"), "PLCL + 21 Days", df_master_vlookup["CRITERIA"])
    df_master_vlookup["CRITERIA"] = np.where((df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "LCL") & (df_master_vlookup["Place of Delivery"].astype(str).str.upper().str.contains("MICHIGAN CROSS")), "LCL through MCD: + 20 days", df_master_vlookup["CRITERIA"])
    df_master_vlookup["CRITERIA"] = np.where((df_master_vlookup["Transload"].notnull()) & (df_master_vlookup["Transload"].astype(str).str.upper().str.contains("YES")), "Transload + 20 days", df_master_vlookup["CRITERIA"])
    # ORIGINAL EDA
    df_master_vlookup["ETA"] = pd.to_datetime(df_master_vlookup['ETA'], errors='coerce')
    df_master_vlookup["ORIGINAL EDA"] = np.where((df_master_vlookup["ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "FCL"), (df_master_vlookup["ETA"] + pd.Timedelta(14, unit='d')), df_master_vlookup["ORIGINAL EDA"])
    df_master_vlookup["ORIGINAL EDA"] = np.where((df_master_vlookup["ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "LCL"), (df_master_vlookup["ETA"] + pd.Timedelta(14, unit='d')), df_master_vlookup["ORIGINAL EDA"])
    df_master_vlookup["ORIGINAL EDA"] = np.where((df_master_vlookup["ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "FCL\\LCL"), (df_master_vlookup["ETA"] + pd.Timedelta(20, unit='d')), df_master_vlookup["ORIGINAL EDA"])
    df_master_vlookup["ORIGINAL EDA"] = np.where((df_master_vlookup["ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "PLCL"), (df_master_vlookup["ETA"] + pd.Timedelta(21, unit='d')), df_master_vlookup["ORIGINAL EDA"])
    df_master_vlookup["ORIGINAL EDA"] = np.where((df_master_vlookup["ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "LCL") & (df_master_vlookup["Place of Delivery"].astype(str).str.upper().str.contains("MICHIGAN CROSS")), (df_master_vlookup["ETA"] + pd.Timedelta(20, unit='d')), df_master_vlookup["ORIGINAL EDA"])
    df_master_vlookup["ORIGINAL EDA"] = np.where((df_master_vlookup["ETA"].notnull()) & (df_master_vlookup["Transload"].notnull()) & (df_master_vlookup["Transload"].astype(str).str.upper().str.contains("YES")), (df_master_vlookup["ETA"] + pd.Timedelta(20, unit='d')), df_master_vlookup["ORIGINAL EDA"])
    # DELAY EDA
    df_master_vlookup["Delay ETA"] = pd.to_datetime(df_master_vlookup['Delay ETA'], errors='coerce')
    df_master_vlookup["DELAY EDA"] = np.where((df_master_vlookup["Delay ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "FCL"), (df_master_vlookup["Delay ETA"] + pd.Timedelta(14, unit='d')), df_master_vlookup["DELAY EDA"])
    df_master_vlookup["DELAY EDA"] = np.where((df_master_vlookup["Delay ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "LCL"), (df_master_vlookup["Delay ETA"] + pd.Timedelta(14, unit='d')), df_master_vlookup["DELAY EDA"])
    df_master_vlookup["DELAY EDA"] = np.where((df_master_vlookup["Delay ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "FCL\\LCL"), (df_master_vlookup["Delay ETA"] + pd.Timedelta(20, unit='d')), df_master_vlookup["DELAY EDA"])
    df_master_vlookup["DELAY EDA"] = np.where((df_master_vlookup["Delay ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "PLCL"), (df_master_vlookup["Delay ETA"] + pd.Timedelta(21, unit='d')), df_master_vlookup["DELAY EDA"])
    df_master_vlookup["DELAY EDA"] = np.where((df_master_vlookup["Delay ETA"].notnull()) & (df_master_vlookup["Path Type"].notnull()) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "LCL") & (df_master_vlookup["Place of Delivery"].astype(str).str.upper().str.contains("MICHIGAN CROSS")), (df_master_vlookup["Delay ETA"] + pd.Timedelta(20, unit='d')), df_master_vlookup["DELAY EDA"])
    df_master_vlookup["DELAY EDA"] = np.where((df_master_vlookup["Delay ETA"].notnull()) & (df_master_vlookup["Transload"].notnull()) & (df_master_vlookup["Transload"].astype(str).str.upper().str.contains("YES")), (df_master_vlookup["Delay ETA"] + pd.Timedelta(20, unit='d')), df_master_vlookup["DELAY EDA"])

    # mx pod
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup["POD"].isnull()) & (df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup["Actual Outgate from Port"] < (np.datetime64("today", "D") - 1)) & (df_master_vlookup["Ultimate Dest. Country"].astype(str).str.upper().str.contains("MEXICO")) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "PLCL"), "Outgate + 7 Days (Mexico)", df_master_vlookup["POD Source"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup["POD"].isnull()) & (df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup["Actual Outgate from Port"] < (np.datetime64("today", "D") - 1)) & (df_master_vlookup["Ultimate Dest. Country"].astype(str).str.upper().str.contains("MEXICO")) & (df_master_vlookup["Path Type"].astype(str).str.upper() == "PLCL"), (df_master_vlookup["Actual Outgate from Port"] + pd.Timedelta(7, unit='d')), df_master_vlookup["POD"])

    df_master_vlookup["POD Source"] = np.where((df_master_vlookup["POD"].isnull()) & (df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup["Actual Outgate from Port"] < (np.datetime64("today", "D") - 1)) & (df_master_vlookup["Ultimate Dest. Country"].astype(str).str.upper().str.contains("MEXICO")) & (df_master_vlookup["Path Type"].astype(str).str.upper().str.contains("FCL|LCL", regex=True)), "Outgate + 1 Day (Mexico)", df_master_vlookup["POD Source"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup["POD"].isnull()) & (df_master_vlookup["Actual Outgate from Port"].notnull()) & (df_master_vlookup["Actual Outgate from Port"] < (np.datetime64("today", "D") - 1)) & (df_master_vlookup["Ultimate Dest. Country"].astype(str).str.upper().str.contains("MEXICO")) & (df_master_vlookup["Path Type"].astype(str).str.upper().str.contains("FCL|LCL", regex=True)), (df_master_vlookup["Actual Outgate from Port"] + pd.Timedelta(1, unit='d')),
                                        df_master_vlookup["POD"])

    # if POD is more then today
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup["POD"].notnull()) & (df_master_vlookup["POD"] > np.datetime64('today', 'D')), np.datetime64("NaT"), df_master_vlookup["POD Source"])
    df_master_vlookup["POD"] = np.where((df_master_vlookup["POD"].notnull()) & (df_master_vlookup["POD"] > np.datetime64('today', 'D')), np.datetime64("NaT"), df_master_vlookup["POD"])

    # remove POD for shut down plants
    df_place_of_delivery = pd.read_excel(path + "\\Input_Files\\Pol and Pod Alias.xlsx", sheet_name="Place_of_delivery")
    Port_of_delivery_list = df_place_of_delivery["Place of delivery"].tolist()
    BTC_list = df_place_of_delivery["BTC"].tolist()

    # Port_of_delivery_list
    df_master_vlookup["POD"] = np.where((df_master_vlookup["POD"].notnull()) & (df_master_vlookup["TODAY DATE"] == np.datetime64('today', 'D')) & (df_master_vlookup["POD Source"].str.contains("cn|ssl", na=False, flags=re.IGNORECASE)) & (df_master_vlookup["Place of Delivery"].astype(str).str.lower().str.contains("|".join(Port_of_delivery_list), na=False)), np.datetime64("NaT"), df_master_vlookup["POD"])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup["POD Source"].notnull()) & (df_master_vlookup["TODAY DATE"] == np.datetime64('today', 'D')) & (df_master_vlookup["POD Source"].str.contains("cn|ssl", na=False, flags=re.IGNORECASE)) & (df_master_vlookup["Place of Delivery"].astype(str).str.lower().str.contains("|".join(Port_of_delivery_list), na=False)), np.nan, df_master_vlookup["POD Source"])

    # BTC list
    df_master_vlookup["POD"] = np.where((df_master_vlookup["POD"].notnull()) & (df_master_vlookup["TODAY DATE"] == np.datetime64('today', 'D')) & (df_master_vlookup["POD Source"].str.contains("cn|ssl", na=False, flags=re.IGNORECASE)) & (df_master_vlookup["BillToCisco"].str.contains("|".join(str(x) for x in BTC_list), na=False)), np.datetime64("NaT"), df_master_vlookup["POD"])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup["POD Source"].notnull()) & (df_master_vlookup["TODAY DATE"] == np.datetime64('today', 'D')) & (df_master_vlookup["POD Source"].str.contains("cn|ssl", na=False, flags=re.IGNORECASE)) & (df_master_vlookup["BillToCisco"].str.contains("|".join(str(x) for x in BTC_list), na=False)), np.nan, df_master_vlookup["POD Source"])

    # if ramp move with SSL pod remove POD
    df_master_vlookup["POD"] = np.where((df_master_vlookup["POD"].notnull()) & (df_master_vlookup["Type of Move for Ceva shipments"].str.contains("RAMP", na=False, flags=re.IGNORECASE)) & (df_master_vlookup["POD Source"].str.contains("SSL", na=False, flags=re.IGNORECASE)), np.datetime64("NaT"), df_master_vlookup["POD"])
    df_master_vlookup["POD Source"] = np.where((df_master_vlookup["POD Source"].notnull()) & (df_master_vlookup["Type of Move for Ceva shipments"].str.contains("RAMP", na=False, flags=re.IGNORECASE)) & (df_master_vlookup["POD Source"].str.contains("SSL", na=False, flags=re.IGNORECASE)), np.nan, df_master_vlookup["POD Source"])

    # TODAY DATE
    df_master_vlookup["TODAY DATE"] = np.where((df_master_vlookup["TODAY DATE"].isnull()) & (df_master_vlookup["POD"].notnull()), date.today().strftime("%m/%d/%Y"), df_master_vlookup["TODAY DATE"])

    # Ultimate Dest country from POD alias
    df_Ultimate_Dest_Country = pd.read_excel(path + "\\Input_Files\\Pol and Pod Alias.xlsx", sheet_name="ports")
    df_master_vlookup = df_master_vlookup.merge(df_Ultimate_Dest_Country[["Port Alias", "Port Country"]], left_on="POD Alias", right_on="Port Alias", how="left")
    df_master_vlookup["Ultimate Dest. Country"] = np.where((df_master_vlookup["Ultimate Dest. Country"].isnull()) & (df_master_vlookup["Port Country"].notnull()), df_master_vlookup["Port Country"], df_master_vlookup["Ultimate Dest. Country"])

    # Atd to Rail is Blank - City Rail Yard name is there and Outgate date is also there then We will Update
    df_master_vlookup["ATD Rail"] = np.where((df_master_vlookup['ATD Rail'].isnull()) & (df_master_vlookup['City / Rail Yard Name'].notnull()) & (df_master_vlookup['Actual Outgate from Port'].notnull()), df_master_vlookup['Actual Outgate from Port'], df_master_vlookup['ATD Rail'])

    # If Outgate date is Bigger Than ATD to rail then We will change - Outgate Date = ATD to Rail
    df_master_vlookup["ATD Rail"] = np.where((df_master_vlookup['Actual Outgate from Port'].notnull()) & (df_master_vlookup['ATD Rail'].notnull()) & (df_master_vlookup['ATD Rail'] < df_master_vlookup['Actual Outgate from Port']), df_master_vlookup["Actual Outgate from Port"], df_master_vlookup["ATD Rail"])

    # Current Status of Shipment
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["Shipment Date/Time"].notnull(), "Depart supplier", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["ATA to POL"].notnull(), "Arrived at port of Departure", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["Origin Port Departure"].notnull(), "Departed on Ship", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["ATD"].notnull(), "Departed on Ship", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["ATA"].notnull(), "Arrived at port of entry", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["Outgate but not departed"].notnull(), "Outgate But Pending departed on rail", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["Actual Outgate from Port"].notnull(), "Outgated", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where(df_master_vlookup["POD"].notnull(), "Delivered", df_master_vlookup["Current Status of Shipment"])
    df_master_vlookup["Current Status of Shipment"] = np.where((df_master_vlookup["POD"].isnull()) & (df_master_vlookup["Actual Departure from Dest Rail"].notnull()) & ((df_master_vlookup["APPT"].isnull()) | (df_master_vlookup["APPT"] < np.datetime64('today', 'D'))) & (~df_master_vlookup["Remarks"].astype(str).astype(str).str.upper().str.contains("PLANT SHUT DOWN", na=False)), "Pending Delivery confirmation", df_master_vlookup["Current Status of Shipment"])

except:
    logging.exception("\n \n \n Error logged: Final Check Points \n")

##############################################################
df_master_vlookup.to_csv(path + "\\OutputFolder\\Final_master.csv", index=False)
##############################################################
End_time = datetime.now()
print("Process Completed!")
print("Time of Processing :", End_time - Start_time)
time.sleep(5)
