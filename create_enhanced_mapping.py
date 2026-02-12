import sys
import io
import json
import pandas as pd

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

excel_file = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

# Read modules (row 9) and subjects (row 10)
df_modules = pd.read_excel(excel_file, sheet_name=0, header=None, skiprows=8, nrows=1)
df_subjects = pd.read_excel(excel_file, sheet_name=0, header=None, skiprows=9, nrows=1)

# Build module structure
modules = []
current_module = None

for col_idx in range(3, len(df_modules.columns)):
    module_name = df_modules.iloc[0, col_idx]
    subject_name = df_subjects.iloc[0, col_idx]
    
    # Check if this is a new module
    if pd.notna(module_name) and str(module_name) != 'nan':
        if current_module:
            modules.append(current_module)
        current_module = {
            'name': str(module_name).strip(),
            'subjects': []
        }
    
    # Add subject to current module
    if pd.notna(subject_name) and str(subject_name) != 'nan' and current_module:
        current_module['subjects'].append(str(subject_name).strip())

# Add last module
if current_module:
    modules.append(current_module)

print(f"Found {len(modules)} modules:\n")
for i, module in enumerate(modules, 1):
    print(f"{i}. {module['name']}")
    print(f"   Subjects: {len(module['subjects'])}")
    for subj in module['subjects'][:3]:
        print(f"     - {subj}")
    if len(module['subjects']) > 3:
        print(f"     ... and {len(module['subjects']) - 3} more")
    print()

# Create enhanced mapping with modules
mapping = {
    "page_1_left": {
        "start_y": 269.45,
        "line_height": 17.64,
        "module_line_height": 13.0,  # Extra space before module header
        "x_number": 24.36,
        "x_subject": 43.56,
        "x_hours": 184.10,
        "x_score": 258.98,
        "x_letter": 281.45,
        "x_point": 319.49,
        "items": []  # Will contain both modules and subjects
    },
    "page_1_right": {
        "start_y": 548.52,
        "line_height": 14.16,
        "module_line_height": 11.0,
        "x_number": 448.75,
        "x_subject": 467.83,
        "x_hours": 609.07,
        "x_score": 685.66,
        "x_letter": 708.10,
        "x_point": 746.14,
        "items": []
    },
    "page_2_left": {
        "start_y": 549.96,
        "line_height": 13.38,
        "module_line_height": 10.0,
        "x_number": 21.60,
        "x_subject": 39.72,
        "x_hours": 179.18,
        "x_score": 252.98,
        "x_letter": 275.57,
        "x_point": 313.49,
        "items": []
    }
}

# Distribute modules and subjects across pages
# Based on PDF template analysis:
# Page 1 Left: БМ 01 (subjects 1-2)
# Page 1 Right: БМ 02-04, КМ 01-03 (subjects 3-35)
# Page 2 Left: КМ 04-05, Ф 00 (subjects 36-48)

subject_counter = 1

# Page 1 Left: First 10 subjects (ЖБП 00 partial + БМ 01)
page1_left_count = 0
for module in modules[:2]:  # ЖБП 00 and БМ 01
    mapping["page_1_left"]["items"].append({
        "type": "module",
        "name": module['name']
    })
    for subj in module['subjects']:
        if page1_left_count >= 10:
            break
        mapping["page_1_left"]["items"].append({
            "type": "subject",
            "row_number": subject_counter,
            "excel_name": subj
        })
        subject_counter += 1
        page1_left_count += 1
    if page1_left_count >= 10:
        break

# Page 1 Right: Next 25 subjects
page1_right_count = 0
start_module_idx = 2 if page1_left_count >= 10 else 0

for module in modules[start_module_idx:]:
    if page1_right_count >= 25:
        break
    mapping["page_1_right"]["items"].append({
        "type": "module",
        "name": module['name']
    })
    for subj in module['subjects']:
        if page1_right_count >= 25:
            break
        mapping["page_1_right"]["items"].append({
            "type": "subject",
            "row_number": subject_counter,
            "excel_name": subj
        })
        subject_counter += 1
        page1_right_count += 1

# Page 2 Left: Remaining subjects
for module in modules:
    # Find subjects not yet added
    for subj in module['subjects']:
        # Check if already added
        already_added = False
        for section in [mapping["page_1_left"], mapping["page_1_right"]]:
            for item in section["items"]:
                if item.get("type") == "subject" and item.get("excel_name") == subj:
                    already_added = True
                    break
        
        if not already_added:
            # Add module header if not already in page 2
            module_in_page2 = any(item.get("name") == module['name'] for item in mapping["page_2_left"]["items"] if item.get("type") == "module")
            if not module_in_page2:
                mapping["page_2_left"]["items"].append({
                    "type": "module",
                    "name": module['name']
                })
            
            mapping["page_2_left"]["items"].append({
                "type": "subject",
                "row_number": subject_counter,
                "excel_name": subj
            })
            subject_counter += 1

# Save mapping
with open('attestat_generator/pdf_subject_mapping_v2.json', 'w', encoding='utf-8') as f:
    json.dump(mapping, f, ensure_ascii=False, indent=2)

print(f"\n✅ Enhanced mapping saved to: attestat_generator/pdf_subject_mapping_v2.json")
print(f"\nPage 1 Left: {sum(1 for item in mapping['page_1_left']['items'] if item['type'] == 'subject')} subjects, {sum(1 for item in mapping['page_1_left']['items'] if item['type'] == 'module')} modules")
print(f"Page 1 Right: {sum(1 for item in mapping['page_1_right']['items'] if item['type'] == 'subject')} subjects, {sum(1 for item in mapping['page_1_right']['items'] if item['type'] == 'module')} modules")
print(f"Page 2 Left: {sum(1 for item in mapping['page_2_left']['items'] if item['type'] == 'subject')} subjects, {sum(1 for item in mapping['page_2_left']['items'] if item['type'] == 'module')} modules")
