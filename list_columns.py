import pandas as pd
import os

import sys

# Force UTF-8 encoding for stdout
sys.stdout.reconfigure(encoding='utf-8')

file_path = 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'

if os.path.exists(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='Лист2', header=10)
        print("Columns found in Excel:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
    except Exception as e:
        print(f"Error reading Excel: {e}")
else:
    print(f"File not found: {file_path}")
