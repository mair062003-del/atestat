import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

excel_file = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

# Read row 9 (subject names) and row 10 (hours)
df_subjects = pd.read_excel(excel_file, sheet_name='экви', header=None, skiprows=8, nrows=3)

print("Row 9 (subject names):")
for i in range(min(50, len(df_subjects.columns))):
    val = df_subjects.iloc[0, i]
    if pd.notna(val) and str(val) != 'nan':
        print(f"Col {i:2d}: {val}")

print("\n\nRow 10 (hours):")
for i in range(min(50, len(df_subjects.columns))):
    val = df_subjects.iloc[1, i]
    if pd.notna(val) and str(val) != 'nan':
        print(f"Col {i:2d}: {val}")

print("\n\nRow 11 (first student grades):")
for i in range(min(50, len(df_subjects.columns))):
    val = df_subjects.iloc[2, i]
    if pd.notna(val) and str(val) != 'nan':
        print(f"Col {i:2d}: {val}")
