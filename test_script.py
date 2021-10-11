import os
import settings
import openpyxl

directory = settings.PATH

files = os.listdir(directory) 

wb = openpyxl.reader.excel.load_workbook(filename=directory + '/' + files[0])

wb.active = 0

sheet = wb.active





print(sheet['A1'].value)
print(files[0])
