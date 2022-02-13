import csv
import pandas as pd
from openpyxl import load_workbook
import openpyxl as xl

#date variables
year = "2022"
month = '02' #keep these to two digits eg: feb = 02
date = '03' #keep these to two digits eg: 3rd day = 03

#defining writer
writer = pd.ExcelWriter('MasterTest.xlsx', engine='xlsxwriter')

## Copying Financial Totals to template sheet

filename = year + '-' + month + '-' + date + ' Financial Totals.csv'
rows = []

with open(filename, 'r') as file:
    csvreader = csv.reader(file)
    header = next(csvreader)
    for row in csvreader:
        rows.append(row)

## print(rows)

##writing array into excel file

df = pd.DataFrame(rows)
data = df.drop([0, 1, 2])

#print(data)

data.to_excel(writer, sheet_name = "Financial Totals", index = False, header = False)

## Copying Staff Performance Overview to template sheet 

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

data2.to_excel(writer, sheet_name = "Staff Performance Overview", index = False, header = False)

#------------ Save writer --------------#
writer.save()

#---------------- Copying "Entry" into new excel file ---------------#

path1 = "EntryTemplate.xlsx"
path2 = "MasterTest.xlsx"

wb1 = xl.load_workbook(filename = path1)
ws1 = wb1.worksheets[0]

wb2 = xl.load_workbook(filename = path2)
ws2 = wb2.create_sheet(ws1.title)

for row in ws1:
    for cell in row:
        ws2[cell.coordinate].value = cell.value

wb2.save(path2)