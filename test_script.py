import os
import settings
import openpyxl

# Берём директорию из настроек
directory = settings.PATH

# Сохраняем имена всех файлов в директории
files = os.listdir(directory) 

new_folder = settings.PATH + "/Готовые файлы"

if "Готовые файлы" in files:
    inp = input("Папка 'Готовые файлы' уже существует. Перезаписать?\ny/n\n")
    if inp == "y":
        pass
    elif inp == "n":
        exit()

if "Готовые файлы" not in files:
    # Создаём новую папку
    os.mkdir(new_folder)   

for file_name in files:
    if file_name.endswith(".xlsx"):
        # Открываем файл
        wb = openpyxl.reader.excel.load_workbook(filename=directory + '/' + file_name)

        # Назначаем активным первый Лист
        wb.active = 0

        # Сохраняем в переменную активный лист
        sheet = wb.active

        sheet['D1'].value = "Объект"

        for str_count in range(2, 16):
            sheet[f'D{str_count}'].value = file_name[:-5]
        
        # Сохраняем файл
        wb.save(str(new_folder) + '/' + file_name)
        
        print(f"Файл '{file_name[:-5]}' перезаписан!")

print("Всё сработало, расходимся...")
