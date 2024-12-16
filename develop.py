import pandas as pd
from openpyxl import load_workbook

def update_files(file_paths):
    last_id = None
    print("Начинаем обновление файлов...")

# Пути к файлам
file_paths = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']

# Вызов функции
update_files(file_paths)