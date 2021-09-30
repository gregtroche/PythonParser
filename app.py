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
    #open workbook get client name and contact info
    workbook = load_workbook(filename = 'mondayAccounts\\' + i)
    worksheet = workbook.worksheets[0]
    client_list = [worksheet['A1'].value]
    client_list.append(worksheet['A2'].value)
    
    # find credential area and add
    amount_of_rows = worksheet.max_row
    for j in range(amount_of_rows)[1:]:
        if worksheet.cell(j, 1).value == 'CREDENTIALS' or worksheet.cell(j, 1).value == 'LOGIN CREDENTIALS':
            credential_name = []
            credential_list = []
            for j in range(amount_of_rows)[j+1:]:
                credential_name.append(worksheet.cell(j, 1).value)
                credential_list.append(worksheet.cell(j, 4).value)
            break
    

    # write to csv, write mode for first iteration, append for every one after that.
    if i == 0:
        with open('clientwb.csv', 'w', encoding='UTF8') as f:
            writer = csv.writer(f)
            writer.writerow(client_list)
            writer.writerow(credential_name)
            writer.writerow(credential_list)
    else:
        with open('clientwb.csv', 'a', encoding='UTF8') as f:
            writer = csv.writer(f)
            writer.writerow(client_list)
            writer.writerow(credential_name)
            writer.writerow(credential_list)


