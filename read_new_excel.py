import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

excel_file = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

# Read first sheet
xl = pd.ExcelFile(excel_file)
sheet_name = xl.sheet_names[0]

print(f"Reading sheet: {sheet_name}\n")

# Try different header rows
for header_row in [0, 1, 2, 10]:
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row)
        print(f"\n=== Header row {header_row} ===")
        print(f"Columns ({len(df.columns)}):")
        for i, col in enumerate(df.columns[:20], 1):
            print(f"{i:2d}. {col}")
        
        print(f"\nFirst row data:")
        if len(df) > 0:
            print(df.iloc[0].head(10))
        break
    except Exception as e:
        print(f"Header row {header_row} failed: {e}")
