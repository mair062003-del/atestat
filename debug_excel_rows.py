import openpyxl
import sys
sys.stdout.reconfigure(encoding='utf-8')

file_path = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

print(f"Opening {file_path}...")
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb.worksheets[0]

print("--- Inspecting Header Rows 9-12 ---")
for r in range(9, 13):
    row_cells = sheet[r]
    values = [c.value for c in row_cells[:10]] # First 10 columns
    print(f"Row {r}: {values}")
