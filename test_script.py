import os
import settings
import openpyxl

# Обращаемся к файлу настроек
directory = settings.PATH

# Берём директорию директорию
files = os.listdir(directory) 

# Сохраняем в переменную сконкатенированные директорию и название файла 
wb = openpyxl.reader.excel.load_workbook(filename=directory + '/' + files[0])

# Назначаем активным первый Лист
wb.active = 0

# Сохраняем в переменную активный лист
sheet = wb.active

sheet['D1'].value = "Объект"

wb.save(settings.SAVE_PATH + '/' + files[0])


