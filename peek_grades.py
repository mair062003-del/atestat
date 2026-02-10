import pandas as pd
import sys

# Force UTF-8
sys.stdout.reconfigure(encoding='utf-8')

file_path = 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'

try:
    df = pd.read_excel(file_path, sheet_name='Лист2', header=10)
    # Pick a few subject columns (e.g., indices 2, 10, 20) and show unique values
    cols = df.columns[[2, 10, 20]]
    
    print("Sample Grade Data:")
    for col in cols:
        print(f"\nColumn: {col}")
        print(df[col].unique()[:10]) # Show first 10 unique values
        
except Exception as e:
    print(f"Error: {e}")
