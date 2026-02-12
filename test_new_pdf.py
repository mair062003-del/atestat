import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from new_data_loader import NewDataLoader
from pdf_generator import PDFGenerator

# Load data
loader = NewDataLoader('data/ПОЛОТНО - 4аКШО-тексерілді.xlsx')
students = loader.load_data()

print(f"Loaded {len(students)} students")

# Get first student
student = students[0]
print(f"\nTesting with: {student['name_kz']}")
print(f"Subjects: {len(student['subjects_list'])}")

# Generate PDF
bg1_path = "data/шаблон каз.jpg"
bg2_path = "data/шаблон каз2.jpg"

if not os.path.exists(bg1_path):
    print(f"ERROR: Background image not found: {bg1_path}")
else:
    print(f"\nUsing background: {bg1_path}")
    
    pdf_gen = PDFGenerator(bg1_path)
    pdf_gen.bg2_path = bg2_path
    
    output_path = "test_new_attestat.pdf"
    pdf_gen.generate(student, output_path)
    
    print(f"\n✅ PDF generated: {output_path}")
    print("Open this file to check the output!")
    
    # Show first 10 subjects that will be in PDF
    print("\nFirst 10 subjects in PDF:")
    for i, subj in enumerate(student['subjects_list'][:10], 1):
        print(f"{i:2d}. {subj['name_kz']:50s} | {subj['score']:3s} {subj['letter']:3s} {subj['point']}")
