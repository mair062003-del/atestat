import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from data_loader import DataLoader
from jinja2 import Template

# Load data
loader = DataLoader('data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx', 'attestat_generator/subjects_mapping.json')
students = loader.load_data()

print(f"Loaded {len(students)} students")

# Get first student
student = students[0]

# Prepare context exactly as app.py does
context = {
    'student_name_kz': student['name_kz'],
    'student_name_ru': student['name_ru'],
    'document_number': 2129475,
    'date': '10.02.2026',
    'subjects': student['subjects_list'],
    's': student['subjects']
}

print(f"\nContext prepared:")
print(f"  student_name_kz: {context['student_name_kz']}")
print(f"  document_number: {context['document_number']}")
print(f"  subjects type: {type(context['subjects'])}")
print(f"  subjects length: {len(context['subjects'])}")

if context['subjects']:
    print(f"\n  First subject:")
    first = context['subjects'][0]
    print(f"    Type: {type(first)}")
    print(f"    Keys: {first.keys() if isinstance(first, dict) else 'N/A'}")
    print(f"    Data: {first}")

# Load template
with open('data/test_template.htm', 'r', encoding='utf-8') as f:
    template_content = f.read()

print(f"\nTemplate loaded, length: {len(template_content)}")

# Render
jinja_template = Template(template_content)
rendered = jinja_template.render(context)

print(f"\nRendered HTML length: {len(rendered)}")

# Check if subjects appear
if 'Қазақ тілі' in rendered or 'Математика' in rendered:
    print("[OK] Subjects found in rendered HTML!")
else:
    print("[WARNING] Subjects NOT found in rendered HTML!")

# Check table rows
table_rows = rendered.count('<tr>')
print(f"Table rows in output: {table_rows}")
print(f"Expected rows: {len(context['subjects']) + 1} (header + subjects)")

# Save output
with open('test_html_output.htm', 'w', encoding='cp1251', errors='xmlcharrefreplace') as f:
    f.write(rendered)

print(f"\nSaved to test_html_output.htm")
print("Open this file in browser to check the result.")
