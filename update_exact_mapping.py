import json

# EXACT COORDINATES FROM шаблон каз.pdf ANALYSIS
mapping = {
    "page_1_left": {
        "start_y": 269.45,
        "line_height": 17.64,
        "x_number": 24.36,
        "x_subject": 43.56,
        "x_hours": 184.10,
        "x_score": 258.98,
        "x_letter": 281.45,
        "x_point": 319.49,
        "subjects": []
    },
    "page_1_right": {
        "start_y": 548.52,
        "line_height": 14.16,
        "x_number": 448.75,
        "x_subject": 467.83,
        "x_hours": 609.07,
        "x_score": 685.66,
        "x_letter": 708.10,
        "x_point": 746.14,
        "subjects": []
    },
    "page_2_left": {
        "start_y": 549.96,
        "line_height": 13.38,
        "x_number": 21.60,
        "x_subject": 39.72,
        "x_hours": 179.18,
        "x_score": 252.98,
        "x_letter": 275.57,
        "x_point": 313.49,
        "subjects": []
    },
    "page_2_right": {
        "start_y": 549.96,
        "line_height": 13.38,
        "x_number": 448.75,
        "x_subject": 467.83,
        "x_hours": 609.07,
        "x_score": 685.66,
        "x_letter": 708.10,
        "x_point": 746.14,
        "subjects": []
    }
}

# Distribute subjects (assuming 48 subjects as in sample)
import sys
import os
sys.path.insert(0, os.path.join(os.getcwd(), 'attestat_generator'))
from new_data_loader import NewDataLoader

loader = NewDataLoader('data/ПОЛОТНО - 4аКШО-тексерілді.xlsx')
students = loader.load_data()

if students:
    subjects = students[0]['subjects_list']
    
    # Page 1 Left: 1-10
    for i in range(min(10, len(subjects))):
        mapping["page_1_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Page 1 Right: 11-35 (25 subjects)
    for i in range(10, min(35, len(subjects))):
        mapping["page_1_right"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Page 2 Left: 36-48 (13 subjects)
    for i in range(35, min(48, len(subjects))):
        mapping["page_2_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Page 2 Right: 49-65
    for i in range(48, len(subjects)):
        mapping["page_2_right"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })

with open('attestat_generator/pdf_subject_mapping_simple.json', 'w', encoding='utf-8') as f:
    json.dump(mapping, f, ensure_ascii=False, indent=2)

print("Mapping file updated with EXACT template coordinates.")
