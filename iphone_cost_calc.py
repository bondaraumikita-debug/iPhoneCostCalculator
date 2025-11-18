from openpyxl import Workbook
from openpyxl.styles import Font, numbers
import os
import re

# Базовый путь
base_path = r"D:\Python projects\P4Git"
base_filename = "iPhone_17_Pro_Max_Calc.xlsx"
file_path = os.path.join(base_path, base_filename)

# Функция проверки файла доступа
def get_available_filename(path):
    if not os.path.exists(path):
        return path

    folder, filename = os.path.split(path)
    name, ext = os.path.splitext(filename)

    pattern = re.compile(rf"{re.escape(name)}_v(\d+){re.escape(ext)}")
    version = 1

    for f in os.listdir(folder):
        match = pattern.match(f)
        if match:
            version = max(version, int(match.group(1)) + 1)

    return os.path.join(folder, f"{name}_v{version}{ext}")

file_path = get_available_filename(file_path)

# Создание Excel
wb = Workbook()
ws = wb.active
ws.title = "Экономика"

# Параметры
params = [
    ("Цена за шт, $", 1400),
    ("Количество, шт", 50),
    ("Доставка, $", 300),
    ("Серт 024 (связь), $", 500),
    ("Серт 020, $", 300),
    ("Серт 037, $", 200),
    ("Утильсбор, %", 3),
    ("Таможня, %", 10),
    ("НДС, %", 20),
    ("Маржа, %", 3),
]

ws["A1"] = "ПАРАМЕТРЫ"
ws["A1"].font = Font(bold=True)

row = 2
for name, value in params:
    ws[f"A{row}"] = name
    ws[f"B{row}"] = value
    row += 1

# Расчёт
start_row = row + 1
ws[f"A{start_row}"] = "РАСЧЁТ"
ws[f"A{start_row}"].font = Font(bold=True)

headers = [
    "Закупка, $", "Утильсбор, $", "Доставка, $",
    "Серт. расходы, $", "Таможня, $",
    "СС без НДС, $", "НДС, $",
    "Маржа, $", "Итог цена парт, $",
    "Цена 1шт, $ без НДС", "Цена 1шт, $ с НДС"
]

col_row = start_row + 1
for col, h in enumerate(headers, start=1):
    cell = ws.cell(row=col_row, column=col, value=h)
    cell.font = Font(bold=True)

data_row = col_row + 1

ws[f"A{data_row}"] = "=B2*B3"
ws[f"B{data_row}"] = f"=A{data_row}*(B8/100)"
ws[f"C{data_row}"] = "=B4"
ws[f"D{data_row}"] = "=B5 + B6 + B7"
ws[f"E{data_row}"] = f"=A{data_row}*(B9/100)"
ws[f"F{data_row}"] = f"=A{data_row}+B{data_row}+C{data_row}+D{data_row}+E{data_row}"
ws[f"G{data_row}"] = f"=F{data_row}*(B10/100)"
ws[f"H{data_row}"] = f"=F{data_row}*(B11/100)"
ws[f"I{data_row}"] = f"=F{data_row}+G{data_row}+H{data_row}"
ws[f"J{data_row}"] = f"=I{data_row}/B3"
ws[f"K{data_row}"] = f"=I{data_row}/B3"

for col in "ABCDEFGHIJK":
    ws[f"{col}{data_row}"].number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.column_dimensions[col].width = 22

wb.save(file_path)

try:
    os.startfile(file_path)
except:
    pass

print(f"✔ Файл создан: {file_path}")
