import csv
import pandas as pd
from openpyxl import load_workbook
import openpyxl as xl
from shutil import copyfile
import shutil

year = "2022"
month = '02' #keep these to two digits eg: feb = 02

for i in range(1, 3):
    date = str('0' + str(i))
    savefilemame = year + '-' + month + '-' + date + "Daily Sales.xlsx"
    shutil.copy("DailySalesTest.xlsx", savefilemame)

    filename = str(year + '-' + month + '-' + date + ' Financial Totals.csv') ##2022-02-01 Financial Totals
    rows = []

    with open(filename, 'r') as file:
        csvreader = csv.reader(file)
        header = next(csvreader)
        for row in csvreader:
            rows.append(row)

    df = pd.DataFrame(rows)
    data = df.drop([0, 1, 2])