import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from data_loader import DataLoader
from pdf_generator import PDFGenerator

# Load data
loader = DataLoader('data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx', 'attestat_generator/subjects_mapping.json')
students = loader.load_data()

print(f"Loaded {len(students)} students")

# Get first student
student = students[0]
print(f"Testing with: {student['name_kz']}")
print(f"Subjects: {len(student['subjects_list'])}")

# Generate PDF
bg1_path = "data/шаблон каз.jpg"
bg2_path = "data/шаблон каз2.jpg"

if not os.path.exists(bg1_path):
    print(f"ERROR: Background image not found: {bg1_path}")
else:
    print(f"Using background: {bg1_path}")
    
    pdf_gen = PDFGenerator(bg1_path)
    pdf_gen.bg2_path = bg2_path
    
    output_path = "test_attestat.pdf"
    pdf_gen.generate(student, output_path)
    
    print(f"\n✅ PDF generated: {output_path}")
    print("Open this file to check if subjects are aligned correctly!")
