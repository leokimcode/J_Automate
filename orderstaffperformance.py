import csv
import pandas as pd
from openpyxl import load_workbook

## Copying Staff Performance Overview to template sheet

writer = pd.ExcelWriter('MasterTest.xlsx', engine='xlsxwriter')

year = "2022" 
month = '02' #write in format (always include 2 digits eg: Jan 2nd, 2022 = 2022, 01, 02)
date = '03' #write in format (always include 2 digits eg: Jan 2nd, 2022 = 2022, 01, 02)

filename = year + '-' + month  + '-' + date + ' Staff Performance Overview.csv'
rows = []

with open(filename, 'r') as file:
    csvreader = csv.reader(file)
    header = next(csvreader)
    for row in csvreader:
        rows.append(row)

## print(rows)

##writing array into excel file

df2 = pd.DataFrame(rows)
data2 = df2.drop([0, 1, 2])

print(data2)

data2 = data2.sort_index()

print(data2)

data2.to_excel(writer, sheet_name = "Staff Performance Overview", index = False, header = False)


#------------ Save writer --------------#
writer.save()