import openpyxl
import csv
from datetime import datetime

# Шаг 1: Загрузка данных из Excel
excel_file = 'herm_data.xlsx'
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Шаг 2: Функция обработки дат
def process_dates(date):
    """
    Обрабатывает значения в столбце 'Дата'.
    Возвращает два значения: date_from и date_to.
    """
    if date is None or str(date).strip() == "":
        return None, None

    date = str(date).strip()

    if '-' in date and date.count('-') == 2:  # Полная дата (например, "28-03-1776")
        try:
            exact_date = datetime.strptime(date, "%d-%m-%Y").date()
            return exact_date.year, date
        except ValueError:
            pass

    if '-' in date:  # Диапазон дат, например "1785-1790"
        parts = date.split('-')
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            return int(parts[0]), int(parts[1])

    if 'XVIII' in date:
        return 1700, 1800
    if 'XVII' in date:
        return 1600, 1700

    if date.isdigit():  # Если это одиночный год
        year = int(date)
        if 1000 <= year <= 9999:
            return year, year

    try:  # Если дата представлена как число с точкой
        year = int(float(date))
        if 1000 <= year <= 9999:
            return year, year
    except ValueError:
        pass

    return None, None

# Шаг 3: Функция извлечения английского названия
def extract_eng_name(text):
    """
    Извлекает только корректное английское название из строки.
    Учитывает, что название может быть в кавычках.
    """
    if text is None:
        return None

    parts = text.split('"')  # Разбиваем текст на части по кавычкам
    for part in parts:
        part = part.strip()
        if part and not part.startswith(("Plate", "Fig", "I", "II", "III", "VIII")):
            return part
    return None

# Шаг 4: Обработка данных
processed_data = []
for row in sheet.iter_rows(min_row=2, values_only=True):  # Пропускаем заголовок
    acc_num = row[0]
    rus_name = row[1]
    eng_name = extract_eng_name(rus_name)
    date_from, date_to = process_dates(row[2])

    material = ""
    technique = ""
    if row[3] and ',' in row[3]:
        parts = row[3].split(',')
        material = parts[0].strip()
        if len(parts) > 1:
            technique = parts[1].strip()

    size = row[4]
    description = rus_name

    processed_data.append([
        acc_num,
        eng_name,
        rus_name,
        description,
        date_from,
        date_to,
        material,
        technique,
        size
    ])

# Шаг 5: Сохранение данных в CSV
output_file = 'processed_data_simple.csv'
with open(output_file, mode='w', encoding='utf-8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow([
        'acc_num', 'eng_name', 'rus_name', 'description', 'date_from', 'date_to', 'material', 'technique', 'size'
    ])
    writer.writerows(processed_data)

print(f"Обработка завершена! Результат сохранен в '{output_file}'")
