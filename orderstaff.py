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
template_file = 'TestCopyPaste.xlsx' 
path = template_file
book = load_workbook(path)

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
