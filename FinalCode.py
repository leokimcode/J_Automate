## By Leo Kim Feb 2022

import csv
import pandas as pd
from openpyxl import load_workbook
import openpyxl as xl
from shutil import copyfile
import shutil

#-----------------------------Create copies of needed files-------------------------------#
year = "2022"
month = '02' #keep these to two digits eg: feb = 02
daysinmonth = 2

#---------------------------Read and write on the files--------------%

## Copying Financial Totals to template sheet

for i in range(1, daysinmonth + 1):
    
    date = str("0" + str(i))

    savefilemame = str(year + '-' + month + '-' + date + "Daily Sales.xlsx")
    shutil.copy("DailySalesTest.xlsx", savefilemame)

    book = load_workbook(savefilemame)

    writer = pd.ExcelWriter(savefilemame, engine='openpyxl', mode = 'a', if_sheet_exists="replace")
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    filename = str(year + '-' + month + '-' + date + ' Financial Totals.csv') ##2022-02-01 Financial Totals
    rows = []

    with open(filename, 'r') as file:
        csvreader = csv.reader(file)
        header = next(csvreader)
        for row in csvreader:
            rows.append(row)

    df = pd.DataFrame(rows)
    data = df.drop([0, 1, 2])

    ## Copying Staff Performance Overview to template sheet 

    filename = str(year + '-' + month  + '-' + date + ' Staff Performance Overview.csv')
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

    writer.save()
    writer.close()

        

