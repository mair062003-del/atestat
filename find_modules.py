import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

excel_file = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

# Read row 9 (module headers)
df_modules = pd.read_excel(excel_file, sheet_name=0, header=None, skiprows=8, nrows=1)

# Read row 10 (subject names)
df_subjects = pd.read_excel(excel_file, sheet_name=0, header=None, skiprows=9, nrows=1)

print("Модули и предметы из Excel:\n")

for col_idx in range(len(df_modules.columns)):
    module_name = df_modules.iloc[0, col_idx]
    subject_name = df_subjects.iloc[0, col_idx]
    
    if pd.notna(module_name) and str(module_name) != 'nan' and col_idx >= 3:
        print(f"\n{'='*60}")
        print(f"Col {col_idx}: МОДУЛЬ: {module_name}")
        print(f"{'='*60}")
    
    if pd.notna(subject_name) and str(subject_name) != 'nan' and col_idx >= 3:
        print(f"  Col {col_idx}: {subject_name}")
