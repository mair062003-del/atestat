import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from data_loader import DataLoader
from docxtpl import DocxTemplate
import datetime

# Load data
loader = DataLoader('data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx', 'attestat_generator/subjects_mapping.json')
students = loader.load_data()

print(f"Loaded {len(students)} students")

# Check if template exists
template_path = input("Enter path to Word template (or press Enter to skip): ").strip()

if not template_path or not os.path.exists(template_path):
    print("No template provided or file not found.")
    print("\nTo create a Word template:")
    print("1. Create a .docx file")
    print("2. Use these tags:")
    print("   {{ student_name_kz }} - student name")
    print("   {{ document_number }} - document number")
    print("   {{ date }} - current date")
    print("   {% for subj in subjects %}")
    print("     {{ subj.name_kz }}: {{ subj.score }} ({{ subj.letter }})")
    print("   {% endfor %}")
else:
    # Test rendering
    student = students[0]
    
    context = {
        'student_name_kz': student['name_kz'],
        'student_name_ru': student['name_ru'],
        'document_number': 2129475,
        'date': datetime.datetime.now().strftime("%d.%m.%Y"),
        'subjects': student['subjects_list'],
        's': student['subjects']
    }
    
    print(f"\nRendering for: {context['student_name_kz']}")
    print(f"Subjects count: {len(context['subjects'])}")
    
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        
        output_path = "test_output.docx"
        doc.save(output_path)
        
        print(f"\n[OK] Saved to {output_path}")
        print(f"Check if the document contains:")
        print(f"  - Student name: {context['student_name_kz']}")
        print(f"  - Document number: {context['document_number']}")
        print(f"  - {len(context['subjects'])} subjects with grades")
        
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
