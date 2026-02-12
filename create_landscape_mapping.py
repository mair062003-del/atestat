import json
import sys
sys.path.insert(0, 'attestat_generator')
from new_data_loader import NewDataLoader

# LANDSCAPE A4: 841.92 x 595.32 points
# Y coordinates start from BOTTOM (0) to TOP (595)
# Header area is approximately y=500-595, so tables start below that

# Create mapping with corrected coordinates for landscape orientation
mapping = {
    "page_1_left": {
        "start_y": 450,  # Start below headers (approximately y=450)
        "line_height": 18.5,  # Space between rows (10 rows total)
        "x_number": 30,  # Row number column
        "x_subject": 55,  # Subject name column
        "x_hours": 260,  # Hours column
        "x_score": 320,  # Score column
        "x_letter": 355,  # Letter grade column
        "x_point": 395,  # GPA point column
        "subjects": []
    },
    "page_1_right": {
        "start_y": 540,  # Start higher on right side (more rows)
        "line_height": 13.8,  # Smaller spacing for 25 rows
        "x_number": 445,  # Row number column (right table)
        "x_subject": 470,  # Subject name column (right table)
        "x_hours": 675,  # Hours column (right table)
        "x_score": 735,  # Score column (right table)
        "x_letter": 770,  # Letter grade column (right table)
        "x_point": 810,  # GPA point column (right table)
        "subjects": []
    },
    "page_2_left": {
        "start_y": 540,  # Start near top of page 2
        "line_height": 17,  # Space between rows
        "x_number": 30,  # Row number column
        "x_subject": 55,  # Subject name column
        "x_hours": 260,  # Hours column
        "x_score": 320,  # Score column
        "x_letter": 355,  # Letter grade column
        "x_point": 395,  # GPA point column
        "subjects": []
    },
    "page_2_right": {
        "start_y": 540,  # Start near top of page 2 right
        "line_height": 17,  # Space between rows
        "x_number": 445,  # Row number column (right table)
        "x_subject": 470,  # Subject name column (right table)
        "x_hours": 675,  # Hours column (right table)
        "x_score": 735,  # Score column (right table)
        "x_letter": 770,  # Letter grade column (right table)
        "x_point": 810,  # GPA point column (right table)
        "subjects": []
    }
}

# Load subjects
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
    
    # Page 2 Left: Subjects 36-48
    for i in range(35, min(48, len(subjects))):
        mapping["page_2_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Page 2 Right: Remaining subjects if any
    for i in range(48, min(len(subjects), 65)):
        mapping["page_2_right"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })

# Save mapping
with open('attestat_generator/pdf_subject_mapping_simple.json', 'w', encoding='utf-8') as f:
    json.dump(mapping, f, ensure_ascii=False, indent=2)

print("Mapping updated for LANDSCAPE A4!")
print(f"\nPage 1 Left: {len(mapping['page_1_left']['subjects'])} subjects (rows 1-10)")
print(f"Page 1 Right: {len(mapping['page_1_right']['subjects'])} subjects (rows 11-35)")
print(f"Page 2 Left: {len(mapping['page_2_left']['subjects'])} subjects (rows 36-48)")
print(f"Page 2 Right: {len(mapping['page_2_right']['subjects'])} subjects (rows 49+)")
print("\nCoordinates adjusted for landscape orientation (842 x 595 points)")
