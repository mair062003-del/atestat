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
    
    # Create simple mapping - just first 48 subjects
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
        }
    }
    
    # First 10 subjects -> page 1 left
    for i in range(min(10, len(subjects))):
        mapping["page_1_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Next 25 subjects (11-35) -> page 1 right
    for i in range(10, min(35, len(subjects))):
        mapping["page_1_right"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Remaining subjects (36-48) -> page 2 left
    for i in range(35, min(48, len(subjects))):
        mapping["page_2_left"]["subjects"].append({
            "row_number": i + 1,
            "excel_name": subjects[i]['name_kz']
        })
    
    # Save simple mapping
    with open('attestat_generator/pdf_subject_mapping_simple.json', 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    
    print("✅ Simple mapping saved to: attestat_generator/pdf_subject_mapping_simple.json")
    print(f"\nPage 1 Left: {len(mapping['page_1_left']['subjects'])} subjects (rows 1-10)")
    print(f"Page 1 Right: {len(mapping['page_1_right']['subjects'])} subjects (rows 11-35)")
    print(f"Page 2 Left: {len(mapping['page_2_left']['subjects'])} subjects (rows 36-48)")
    
    print("\nSubjects distribution:")
    print("\nPage 1 Left:")
    for s in mapping["page_1_left"]["subjects"]:
        print(f"  {s['row_number']:2d}. {s['excel_name']}")
    
    print("\nPage 1 Right (first 5):")
    for s in mapping["page_1_right"]["subjects"][:5]:
        print(f"  {s['row_number']:2d}. {s['excel_name']}")
    print(f"  ... and {len(mapping['page_1_right']['subjects'])-5} more")
    
    print("\nPage 2 Left (first 5):")
    for s in mapping["page_2_left"]["subjects"][:5]:
        print(f"  {s['row_number']:2d}. {s['excel_name']}")
    print(f"  ... and {len(mapping['page_2_left']['subjects'])-5} more")
