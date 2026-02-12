import pandas as pd
import json
import sys

# Force UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

file_path = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

# Load the file
df = pd.read_excel(file_path, sheet_name=0, header=None)

# Extract module headers (Row 9, index 8) and subject names (Row 10, index 9)
modules_row = df.iloc[8]
subjects_row = df.iloc[9]

structure = []
current_module = None
current_subjects = []

# Start iterating from column 4 (index 4) where subjects typically begin
start_col = 4
end_col = len(df.columns)

for col in range(start_col, end_col):
    module_val = modules_row[col]
    subject_val = subjects_row[col]
    
    # Check if subject exists
    if pd.isna(subject_val):
        continue
        
    subject_name = str(subject_val).strip()
    
    # Check for new module
    if pd.notna(module_val):
        module_name = str(module_val).strip().replace('\n', ' ')
        
        # If we had a previous module, save it
        if current_module:
            structure.append({
                "module_name": current_module,
                "subjects": current_subjects
            })
        
        # Start new module
        current_module = module_name
        current_subjects = []
    
    # Add subject to current module
    current_subjects.append(subject_name)

# Add the last module if exists
if current_module and current_subjects:
    structure.append({
        "module_name": current_module,
        "subjects": current_subjects
    })

# Save structure to JSON
with open('module_structure.json', 'w', encoding='utf-8') as f:
    json.dump(structure, f, ensure_ascii=False, indent=2)

print(f"✅ Extracted {len(structure)} modules.")
print("Saved to module_structure.json")

# Print short summary
for m in structure:
    print(f"Module: {m['module_name'][:50]}... ({len(m['subjects'])} subjects)")
