import pandas as pd
import sys

# Load the excel file
file_path = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

try:
    # Load with header=None to access by row index easily
    df = pd.read_excel(file_path, sheet_name=0, header=None)
    
    print("--- Checking Module Headers (Row 9 / Index 8) ---")
    row_modules = df.iloc[8] # Row 9 is index 8
    row_subjects = df.iloc[9] # Row 10 is index 9 (Subject names)
    
    current_module = None
    
    # Iterate through columns starting from where data usually begins (e.g. col 4)
    # Based on previous analysis, subjects start around column 4 (E)
    
    modules_found = []
    
    for col in range(len(row_modules)):
        val = row_modules[col]
        subj = row_subjects[col]
        
        if pd.notna(val):
            # Clean up module name
            module_name = str(val).strip().replace('\n', ' ')
            if module_name != current_module:
                current_module = module_name
                modules_found.append({
                    "start_col": col,
                    "name": current_module,
                    "first_subject": str(subj).strip() if pd.notna(subj) else "Unknown"
                })
    
    print(f"Found {len(modules_found)} modules:")
    for m in modules_found:
        print(f"Column {m['start_col']}: {m['name']} (Starts with: {m['first_subject']})")

except Exception as e:
    print(f"Error: {e}")
