# openpyxl to read xlsx files
import openpyxl
import csv
import os
import time
from openpyxl import load_workbook

#loop through xlsx files and add to array
#directory = r'C:\PythonParser\mondayAccounts'
directory = r'C:\\Users\\steve\\Desktop\\py-parser'
client_sheets = []
clientinfo =[]
client_info_split = []
client_info_container = []
client_info_title = []
client_info_data = []
split_two = []
a =0 
start = time.time()
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):
        client_sheets.append(filename)
    else:
        continue

#loop through client sheets array and concat filename
for i in client_sheets:
    #open workbook get client name and contact info
    #workbook = load_workbook(filename = 'mondayAccounts\\' + i)
    workbook = load_workbook(filename = 'C:\\Users\\steve\\Desktop\\py-parser\\' + i)
    worksheet = workbook.worksheets[0]
    client_list = [worksheet['A1'].value]
    #client_list.append(worksheet['A2'].value)
    clientinfo.append(worksheet['A2'].value.split('\n'))
    client_info_split = clientinfo[0]
    
    # find credential area and add
    amount_of_rows = worksheet.max_row
    for j in range(amount_of_rows)[1:]:
        if worksheet.cell(j, 1).value == 'CREDENTIALS' or worksheet.cell(j, 1).value == 'LOGIN CREDENTIALS':
            credential_name = []
            credential_list = []
            for j in range(amount_of_rows)[j+1:]:
                credential_name.append(worksheet.cell(j, 1).value)
                credential_list.append(worksheet.cell(j, 4).value.split('\n'))
            break

    
    
    # print(sec_half)
    #print(client_info_container)
    # write to csv, write mode for first iteration, append for every one after that.
    #if i == 0:
     
    if a ==0:
        with open('C:\\Users\\steve\\Desktop\\py-parser\\clientwb.csv', 'w',  encoding='UTF8') as f:
            for b in range(len(client_info_split)):
                client_info_split[b] = client_info_split[b].split(":",1)
                client_info_container = client_info_split[b]
                client_info_title.append(client_info_container[0])
                client_info_data.append(client_info_container[1])
                # print(client_info_container)
                # length = len(client_info_container)
                # mid_index = length//2
                # first_half = client_info_container[:mid_index]
                # sec_half = client_info_container[mid_index:]
                # client_info_title.append(first_half)
                # client_info_data.append(sec_half)
            # for c in range(len(client_info_container) - 1):
            #     client_info_title.append(client_info_container[c][0])
            #     client_info_data.append(client_info_container[c][1])

            writer = csv.writer(f)
            writer.writerow(client_list)
            
                # writer.writerow(first_half)
                # writer.writerow(sec_half)    
            # print(client_info_container)   
            writer.writerow(client_info_title)
            writer.writerow(client_info_data)
            writer.writerow(credential_name)
           # for c in credential_list:
                #writer.writerow(c)
            writer.writerow(credential_list)
           
    else:
        with open('C:\\Users\\steve\\Desktop\\py-parser\\clientwb.csv', 'a', encoding='UTF8') as f:
            writer = csv.writer(f)
            writer.writerow(client_list)
            
            writer.writerow(credential_name)
            writer.writerow(credential_list)
    a+=1

print('file updated')
print(time.time()-start)