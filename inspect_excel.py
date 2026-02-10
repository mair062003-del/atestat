import pandas as pd
import os

file_path = 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'

if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
    # try looking in current dir just in case
    file_path = 'ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
    if not os.path.exists(file_path):
         print(f"File not found in current dir either: {file_path}")
         exit(1)

try:
    xl = pd.ExcelFile(file_path)
    print("Sheet names:", xl.sheet_names)
except Exception as e:
    print("Error reading Excel file:", e)
