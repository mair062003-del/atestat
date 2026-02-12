import sys
import os
import io

# Force UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from data_loader import DataLoader
from jinja2 import Template

# Load data
loader = DataLoader('data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx', 'attestat_generator/subjects_mapping.json')
students = loader.load_data()

print(f"Loaded {len(students)} students")

# Load HTML template
with open('attestat_generator/templates/template_kz.htm', 'r', encoding='cp1251') as f:
    template_content = f.read()

print(f"Template loaded, length: {len(template_content)}")

# Check for placeholders
if '{{ student_name_kz }}' in template_content:
    print("[OK] Found {{ student_name_kz }} placeholder")
else:
    print("[WARNING] {{ student_name_kz }} placeholder NOT found!")

if '{% for s in subjects %}' in template_content:
    print("[OK] Found subjects loop")
else:
    print("[WARNING] subjects loop NOT found!")

# Try rendering
jinja_template = Template(template_content)
student = students[0]

context = {
    'student_name_kz': student['name_kz'],
    'student_name_ru': student['name_ru'],
    'document_number': 2129475,
    'subjects': student['subjects_list'],
    's': student['subjects']
}

print(f"\nRendering template with:")
print(f"  student_name_kz: {context['student_name_kz']}")
print(f"  subjects count: {len(context['subjects'])}")

rendered = jinja_template.render(context)

# Check if name appears in output
if student['name_kz'] in rendered:
    print(f"\n[OK] Student name FOUND in rendered output")
else:
    print(f"\n[ERROR] Student name NOT found in rendered output!")
    
# Save sample
with open('test_output.htm', 'w', encoding='cp1251', errors='xmlcharrefreplace') as f:
    f.write(rendered)
    
print(f"\nSaved test output to test_output.htm")
