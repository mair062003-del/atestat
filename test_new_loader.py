import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'attestat_generator')

from new_data_loader import NewDataLoader

loader = NewDataLoader('data/ПОЛОТНО - 4аКШО-тексерілді.xlsx')
students = loader.load_data()

print(f"\n✅ Loaded {len(students)} students\n")

if students:
    student = students[0]
    print(f"First student: {student['name_kz']}")
    print(f"Subjects: {len(student['subjects_list'])}\n")
    
    print("First 10 subjects:")
    for i, subj in enumerate(student['subjects_list'][:10], 1):
        print(f"{i:2d}. {subj['name_kz']:50s} | {subj['score']:3s} | {subj['letter']:3s} | {subj['point']}")
