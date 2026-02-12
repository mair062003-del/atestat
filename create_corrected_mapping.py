import json

# Based on visual analysis of шаблон каз.jpg and шаблон каз 2.jpg
# Page 1 has TWO tables: left (10 rows) and right (25 rows)
# Page 2 has TWO tables: left and right

# Create corrected mapping with accurate coordinates
mapping = {
    "page_1_left": {
        "start_y": 580,  # Top of first row in left table
        "line_height": 19.5,  # Space between rows
        "x_number": 35,  # Row number column
        "x_subject": 55,  # Subject name column
        "x_hours": 270,  # Hours column
        "x_score": 330,  # Score column
        "x_letter": 365,  # Letter grade column
        "x_point": 410,  # GPA point column
        "subjects": []
    },
    "page_1_right": {
        "start_y": 580,  # Top of first row in right table  
        "line_height": 13.5,  # Smaller spacing for more rows
        "x_number": 455,  # Row number column (right side)
        "x_subject": 475,  # Subject name column (right side)
        "x_hours": 690,  # Hours column (right side)
        "x_score": 750,  # Score column (right side)
        "x_letter": 785,  # Letter grade column (right side)
        "x_point": 830,  # GPA point column (right side)
        "subjects": []
    },
    "page_2_left": {
        "start_y": 750,  # Top of first row on page 2 left
        "line_height": 16,  # Space between rows
        "x_number": 35,  # Row number column
        "x_subject": 55,  # Subject name column  
        "x_hours": 270,  # Hours column
        "x_score": 330,  # Score column
        "x_letter": 365,  # Letter grade column
        "x_point": 410,  # GPA point column
        "subjects": []
    },
    "page_2_right": {
        "start_y": 750,  # Top of first row on page 2 right
        "line_height": 16,  # Space between rows
        "x_number": 455,  # Row number column (right side)
        "x_subject": 475,  # Subject name column (right side)
        "x_hours": 690,  # Hours column (right side)
        "x_score": 750,  # Score column (right side)
        "x_letter": 785,  # Letter grade column (right side)
        "x_point": 830,  # GPA point column (right side)
        "subjects": []
    }
}

# Load subjects from our data
import sys
sys.path.insert(0, 'attestat_generator')
from new_data_loader import NewDataLoader

loader = NewDataLoader('data/ПОЛОТНО - 4аКШО-тексерілді.xlsx')
students = loader.load_data()

if students:
    subjects = students[0]['subjects_list']
    
    # Page 1 Left: First 10 subjects
    for i in range(min(10, len(subjects))):
        mapping["page_1_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Page 1 Right: Next 25 subjects (11-35)
    for i in range(10, min(35, len(subjects))):
        mapping["page_1_right"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Page 2 Left: Subjects 36-48 (first 13 on page 2)
    for i in range(35, min(48, len(subjects))):
        mapping["page_2_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Page 2 Right: Remaining subjects if any (49+)
    for i in range(48, min(len(subjects), 65)):  # Up to row 65 if needed
        mapping["page_2_right"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })

# Save corrected mapping
with open('attestat_generator/pdf_subject_mapping_simple.json', 'w', encoding='utf-8') as f:
    json.dump(mapping, f, ensure_ascii=False, indent=2)

print("OK Corrected mapping saved!")
print(f"\nPage 1 Left: {len(mapping['page_1_left']['subjects'])} subjects")
print(f"Page 1 Right: {len(mapping['page_1_right']['subjects'])} subjects")
print(f"Page 2 Left: {len(mapping['page_2_left']['subjects'])} subjects")
print(f"Page 2 Right: {len(mapping['page_2_right']['subjects'])} subjects")
