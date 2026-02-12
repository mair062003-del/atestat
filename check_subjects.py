import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from data_loader import DataLoader

# Load data
loader = DataLoader('data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx', 'attestat_generator/subjects_mapping.json')
students = loader.load_data()

student = students[0]

print(f"Student: {student['name_kz']}")
print(f"\nFirst 10 subjects:")
for i, subj in enumerate(student['subjects_list'][:10]):
    print(f"{i+1}. {subj['name_kz']} ({subj['name_ru']}) - {subj['score']} ({subj['letter']})")

print(f"\nTotal subjects: {len(student['subjects_list'])}")
