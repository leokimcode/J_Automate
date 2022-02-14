import csv
import pandas as pd
from openpyxl import load_workbook
import openpyxl as xl
from shutil import copyfile

#date variables
year = "2022"
month = '02' #keep these to two digits eg: feb = 02
date = '03' #keep these to two digits eg: 3rd day = 03

#template

template_file = 'DailySalesTest.xlsx' 
book = load_workbook(template_file)

savefilename = year + '-' + month + '-' + date + "Daily Sales.xlsx"

writer = pd.ExcelWriter('DailySalesTest.xlsx' , engine='openpyxl', mode = 'a', if_sheet_exists="replace")
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

## Copying Financial Totals to template sheet

filename = year + '-' + month + '-' + date + ' Financial Totals.csv'
rows = []

with open(filename, 'r') as file:
    csvreader = csv.reader(file)
    header = next(csvreader)
    for row in csvreader:
        rows.append(row)

df = pd.DataFrame(rows)
data = df.drop([0, 1, 2])

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

df2 = pd.DataFrame(rows)

data2 = df2.drop([0, 1, 2])

ashrow = data2[data2[0].str.contains('Ashlee')]
britrow = data2[data2[0].str.contains('Brittany')]
katierow = data2[data2[0].str.contains('Katie')]
julrow = data2[data2[0].str.contains('Julia')]
aliciarow =  data2[data2[0].str.contains('Alicia')]
emprow = data2[data2[0].str.contains('Emplyee')]

empData = pd.DataFrame()
empData = empData.append(ashrow)
empData = empData.append(britrow)
empData = empData.append(katierow)
empData = empData.append(julrow)
empData = empData.append(aliciarow)
empData = empData.append(emprow)

print(empData)

##data.to_excel(writer, sheet_name = "Financial Totals", header = False, index = False)
##empData.to_excel(writer, sheet_name = "Staff Performance Overview", index = False, header = False)

data.to_excel(writer, sheet_name = "Financial Totals", index = False, header = False)
empData.to_excel(writer, sheet_name = "Staff Performance Overview", startrow = 2, index = False, header = False)

writer.save(filename = savefilename)
writer.close()
