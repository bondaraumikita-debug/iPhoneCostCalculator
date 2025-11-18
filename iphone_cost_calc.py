from openpyxl import Workbook
from openpyxl.styles import Font, numbers
import os

# Путь сохранения файла
file_path = r"D:\Python projects\P4Git\iPhone_17_Pro_Max_Calc.xlsx"
os.makedirs(os.path.dirname(file_path), exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = "Экономика"

# --- Блок параметров ---
ws["A1"] = "ПАРАМЕТРЫ"
ws["A1"].font = Font(bold=True)

params = [
    ("Цена за шт, $", 1400),
    ("Количество, шт", 50),
    ("Доставка, $", 300),
    ("Серт Связь, $", 700),
    ("Утильсбор, %", 3),
    ("Таможня, %", 10),
    ("НДС, %", 20),
    ("Маржа, %", 3),
]

row = 2
for name, value in params:
    ws[f"A{row}"] = name
    ws[f"B{row}"] = value
    row += 1

# --- Заголовок блока расчётов ---
start_row = row + 1
ws[f"A{start_row}"] = "РАСЧЁТ"
ws[f"A{start_row}"].font = Font(bold=True)

headers = [
    "Закупка, $", "Утильсбор, $", "Доставка, $", "Серт Связь, $",
    "Таможня, $", "СС без НДС, $", "НДС, $",
    "Маржа, $", "Итог цена парт, $", "Цена 1шт, $ без НДС"
]

col_row = start_row + 1
for col, h in enumerate(headers, start=1):
    cell = ws.cell(row=col_row, column=col, value=h)
    cell.font = Font(bold=True)

# --- Формулы расчёта ---
data_row = col_row + 1

ws[f"A{data_row}"] = "=B2*B3"                                  # Закупка
ws[f"B{data_row}"] = f"=A{data_row}*(B6/100)"                  # Утильсбор
ws[f"C{data_row}"] = "=B4"                                     # Доставка
ws[f"D{data_row}"] = "=B5"                                     # Сертификат
ws[f"E{data_row}"] = f"=A{data_row}*(B7/100)"                  # Таможня
ws[f"F{data_row}"] = f"=A{data_row}+B{data_row}+C{data_row}+D{data_row}+E{data_row}"  # СС без НДС
ws[f"G{data_row}"] = f"=F{data_row}*(B8/100)"                  # НДС
ws[f"H{data_row}"] = f"=F{data_row}*(B9/100)"                  # Маржа
ws[f"I{data_row}"] = f"=F{data_row}+G{data_row}+H{data_row}"   # Итог цена парт
ws[f"J{data_row}"] = f"=I{data_row}/B3"                        # Цена 1шт без НДС

# --- Формат валюты и ширина столбцов ---
for col in "ABCDEFGHIJ":
    ws[f"{col}{data_row}"].number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.column_dimensions[col].width = 20

# --- Сохраняем и открываем Excel ---
wb.save(file_path)

try:
    os.startfile(file_path)
except Exception:
    pass

print("✔ Расчёт выполнен")
print("📁 Файл создан:", file_path)
