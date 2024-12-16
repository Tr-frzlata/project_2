import pandas as pd
from openpyxl import load_workbook


def find_id_column(df):
    # Ищем столбец, содержащий 'id' (без учета регистра)
    id_columns = df.columns[df.columns.str.lower() == 'id']
    if len(id_columns) > 0:
        return id_columns[0]
    return None


def update_ids(file_paths):
    last_id = None

    for i, file_path in enumerate(file_paths):
        # Загрузка файла
        df = pd.read_excel(file_path)

        # Поиск столбца с 'id'
        id_column = find_id_column(df)

        if id_column is None:
            print(f"Столбец 'id' не найден в файле {file_path}. Пропускаем этот файл.")
            continue

        if i == 0:
            # Для первого файла берем значение ID из второй строки
            last_id = df[id_column].iloc[1]

        # Если last_id все еще None (например, если первый файл пустой), начинаем с 1
        if last_id is None:
            last_id = 0

        # Создание новых ID
        new_ids = range(last_id + 1, last_id + len(df) + 1)

        # Замена ID в DataFrame
        df[id_column] = new_ids

        # Обновление last_id для следующего файла
        last_id = last_id + len(df)

        # Сохранение изменений в файл
        book = load_workbook(file_path)
        sheet_name = book.sheetnames[0]

        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            # Копирование существующих листов, кроме того, который мы обновляем
            for sheet in book.sheetnames:
                if sheet != sheet_name:
                    book[sheet].copy(writer.book, sheet)

            # Запись обновленных данных
            df.to_excel(writer, index=False, sheet_name=sheet_name)


# Пути к файлам
file_paths = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']

# Вызов функции
update_ids(file_paths)

print("Обновление ID завершено во всех файлах.")