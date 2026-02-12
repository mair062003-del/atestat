import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from data_loader import DataLoader

loader = DataLoader('data/ПОЛОТНО - 4аКШО-тексерілді.xlsx', 'attestat_generator/subjects_mapping.json')
students = loader.load_data()

student = students[0]

print("Все предметы из Excel (в порядке появления):\n")
for i, subj in enumerate(student['subjects_list'], 1):
    print(f"{i:2d}. {subj['name_kz']:50s} | {subj['name_ru']}")

print(f"\nВсего предметов: {len(student['subjects_list'])}")
