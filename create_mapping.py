import sys
import io
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from new_data_loader import NewDataLoader

# Load subjects from Excel
loader = NewDataLoader('data/ПОЛОТНО - 4аКШО-тексерілді.xlsx')
students = loader.load_data()

if students:
    student = students[0]
    subjects = student['subjects_list']
    
    print(f"Total subjects in Excel: {len(subjects)}\n")
    print("Subjects from Excel (in order):\n")
    
    for i, subj in enumerate(subjects, 1):
        print(f"{i:2d}. {subj['name_kz']}")
    
    # Based on pdf_structure_analysis.txt, create mapping
    # Page 1 has subjects 1-10 (left column) and 11-35 (right column)
    # Page 2 has subjects 36-48 (left column)
    
    print("\n\n" + "="*60)
    print("Creating mapping based on PDF template structure...")
    print("="*60)
    
    # From pdf_structure_analysis.txt:
    # Page 1 Left: subjects 1-10
    # Page 1 Right: subjects 11-35
    # Page 2 Left: subjects 36-48
    
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
            "line_height": 14.16,  # Approximate
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
            "line_height": 13.38,  # Approximate
            "x_number": 21.60,
            "x_subject": 39.72,
            "x_hours": 179.18,
            "x_score": 252.98,
            "x_letter": 275.57,
            "x_point": 313.49,
            "subjects": []
        }
    }
    
    # Assign subjects to positions
    # First 10 subjects go to page 1 left
    for i in range(min(10, len(subjects))):
        mapping["page_1_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Next 25 subjects go to page 1 right (11-35)
    for i in range(10, min(35, len(subjects))):
        mapping["page_1_right"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Remaining subjects go to page 2 left (36-48)
    for i in range(35, min(48, len(subjects))):
        mapping["page_2_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Save mapping
    with open('attestat_generator/pdf_subject_mapping.json', 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    
    print("\n✅ Mapping saved to: attestat_generator/pdf_subject_mapping.json")
    print(f"\nPage 1 Left: {len(mapping['page_1_left']['subjects'])} subjects")
    print(f"Page 1 Right: {len(mapping['page_1_right']['subjects'])} subjects")
    print(f"Page 2 Left: {len(mapping['page_2_left']['subjects'])} subjects")
