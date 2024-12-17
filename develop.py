import pandas as pd
from openpyxl import load_workbook
from functions import update_ids

# Пути к файлам
file_paths = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']

# Вызов функции
update_ids(file_paths)

print("Обновление ID завершено во всех файлах.")