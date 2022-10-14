import pandas as pd
import validate
import counts

import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

### LIMITS ###
# 1. Workbooks must be single sheet; 
# 2. Can't have serials or MACs split between multiple columns, 
#    All Serials must be in single column and MACs must be in a single column

 

alphabet = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
            'u', 'v', 'w', 'x', 'y', 'z']



STAGING_HOSTNAME_COL = alphabet.index(input('In STAGING Sheet, What column is hostname?').lower())
STAGING_SERIAL_COL = alphabet.index(input('In STAGING Sheet, What column is Serial Numbers?').lower())
STAGING_MAC_COL = alphabet.index(input('In STAGING Sheet, What column is MAC Addresss?').lower())


EXPORT_LIST_HOSTNAME_COL = 0
EXPORT_LIST_SERIAL_COL = 6
EXPORT_LIST_MAC_COL = 7

print('Upload EXPORT list')
export_filepath = filedialog.askopenfilename()

print('Upload STAGING Sheet')
staging_filepath = filedialog.askopenfilename()



df_export = pd.read_csv(export_filepath)
df_staging = pd.read_excel(staging_filepath, engine='openpyxl')



####   Iterates through the exported list from Aruba and grabs the MAC on each row and saves it to a variable #####

for row in range(len(df_export)):
    currentRow = df_export.iloc[row]
    mac = currentRow[EXPORT_LIST_MAC_COL].upper()
    serial = currentRow[EXPORT_LIST_SERIAL_COL]
    
    if (validate.isMacUnique(mac, STAGING_MAC_COL, df_staging) 
        & validate.isSerialUnique(serial, STAGING_SERIAL_COL, df_staging)
        & validate.isMacAndSerialOnSameRow(mac, serial, STAGING_MAC_COL, STAGING_SERIAL_COL, df_staging)
        & validate.isNotAlreadyRenamed(currentRow, mac, EXPORT_LIST_HOSTNAME_COL)):        

        filt = df_staging.iloc[:, STAGING_MAC_COL] == mac
        rowFound = df_staging[filt]
        hostname = rowFound.iloc[0][STAGING_HOSTNAME_COL]
        df_export.iloc[row, EXPORT_LIST_HOSTNAME_COL] = hostname
        counts.rename_successful += 1




writer = pd.ExcelWriter('Output/RenamedAPs.xlsx', engine='xlsxwriter')
df_export.to_excel(writer, sheet_name='sheet1', index=False)
writer.save()

print(f'{counts.rename_successful} have been renamed succesfully')
print(f'{counts.already_renamed} device(s) already renamed')
print(f'{counts.mac_not_found} MACs from export sheet were not found in staging sheet')
print(f'{counts.mac_duplicated} MACs are duplicated in staging sheet')
print(f'{counts.serial_not_found} Serial Numbers from export sheet were not found in staging sheet')
print(f'{counts.serial_duplicated} Serial Numbers are duplicated in staging sheet')
print(f'{counts.mac_serial_mismatch} MAC and Serial mismatches on the staging sheet')


endMessage = print('Script Complete.')