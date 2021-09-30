# openpyxl to read xlsx files
import openpyxl
import csv
import os
from openpyxl import load_workbook

#loop through xlsx files and add to array
directory = r'C:\PythonParser\mondayAccounts'
client_sheets = []
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):
        client_sheets.append(filename)
    else:
        continue

#loop through client sheets array and concat filename
for i in client_sheets:
    #open workbook, get client name and contact info
    workbook = load_workbook(filename = 'mondayAccounts\\' + i)
    sheet_ranges = workbook.worksheets[0]
    client_name = [sheet_ranges['A1'].value]
    client_name.append(sheet_ranges['A2'].value)

    #write to csv, write mode for first iteration, append for every one after that.
    if i == 0:
        with open('clientwb.csv', 'w', encoding='UTF8') as f:
            writer = csv.writer(f)
            writer.writerow(client_name)
    else:
        with open('clientwb.csv', 'a', encoding='UTF8') as f:
            writer = csv.writer(f)
            writer.writerow(client_name)


