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
    
    output_path = "test_mapped_attestat.pdf"
    pdf_gen.generate(student, output_path)
    
    print(f"\n✅ PDF generated: {output_path}")
    print("\n📋 Subjects that should appear in PDF:")
    print("\nPage 1 Left (1-10):")
    for i in range(min(10, len(student['subjects_list']))):
        subj = student['subjects_list'][i]
        print(f"  {i+1:2d}. [{subj.get('module', '')[:10]}] {subj['name_kz']:45s} {str(subj['score']):3s} {subj['letter']:3s} {subj['point']}")
    
    print("\nPage 1 Right (11-35):")
    for i in range(10, min(35, len(student['subjects_list']))):
        subj = student['subjects_list'][i]
        mod = subj.get('module', '')[:10]
        print(f"  {i+1:2d}. [{mod}] {subj['name_kz']:45s} {str(subj['score']):3s} {subj['letter']:3s} {subj['point']}")
    
    print("\nPage 2 Left (36-48):")
    for i in range(35, min(48, len(student['subjects_list']))):
        subj = student['subjects_list'][i]
        mod = subj.get('module', '')[:10]
        print(f"  {i+1:2d}. [{mod}] {subj['name_kz']:45s} {str(subj['score']):3s} {subj['letter']:3s} {subj['point']}")
        
    print("\nPage 2 Right (49+):")
    for i in range(48, len(student['subjects_list'])):
        subj = student['subjects_list'][i]
        mod = subj.get('module', '')[:10]
        print(f"  {i+1:2d}. [{mod}] {subj['name_kz']:45s} {str(subj['score']):3s} {subj['letter']:3s} {subj['point']}")
