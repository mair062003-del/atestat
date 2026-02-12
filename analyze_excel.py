import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

excel_file = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

xl = pd.ExcelFile(excel_file)
sheet_name = xl.sheet_names[0]

print(f"Sheet name: {sheet_name}\n")

# Try reading with different header rows
for header_row in range(0, 15):
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row, nrows=5)
        
        # Check if this looks like the data row
        # Look for student names in column 1 or 2
        if len(df) > 0:
            first_col_values = df.iloc[:, 1].dropna().astype(str).tolist()
            
            # Check if we have subject names in columns
            col_names = [str(col) for col in df.columns if 'Unnamed' not in str(col)]
            
            if len(col_names) > 10:  # Likely found the header row
                print(f"\n{'='*60}")
                print(f"HEADER ROW {header_row} looks promising!")
                print(f"{'='*60}")
                print(f"\nColumns ({len(df.columns)}):")
                
                for i, col in enumerate(df.columns[:30], 1):
                    print(f"{i:2d}. {col}")
                
                print(f"\n\nFirst 3 rows of data:")
                print(df.iloc[:3, :5])
                
                break
    except Exception as e:
        pass

# Also try to find where student names start
print(f"\n\n{'='*60}")
print("Looking for student names...")
print(f"{'='*60}\n")

for row in range(0, 20):
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, skiprows=row, nrows=3)
        if len(df) > 0 and len(df.columns) > 1:
            # Check second column for potential student names
            val = str(df.iloc[0, 1])
            if len(val) > 10 and not val.startswith('Unnamed'):
                print(f"Row {row}: {val[:50]}")
    except:
        pass
