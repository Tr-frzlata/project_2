import pandas as pd
from openpyxl import load_workbook

def update_files(file_paths):
    last_id = None
    print("Начинаем обновление файлов...")

    for i, file_path in enumerate(file_paths):
        # Загрузка файла
        df = pd.read_excel(file_path)
        print(f"Файл {file_path} загружен.")

# Пути к файлам
file_paths = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']

# Вызов функции
update_files(file_paths)