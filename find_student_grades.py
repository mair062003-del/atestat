import sys
import os
import io

# Force stdout to utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Add package to path
sys.path.append(os.path.join(os.getcwd(), 'attestat_generator'))

from attestat_generator.data_loader import DataLoader
import json

excel_path = 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
# Mapping is inside the package folder relative to CWD?
# No, it's in attestat_generator/subjects_mapping.json
mapping_path = 'attestat_generator/subjects_mapping.json'

# Initialize loader
loader = DataLoader(excel_path, mapping_path)
students = loader.load_data()

print(f"Loaded {len(students)} students.")
# Print all names to help find the right one
print("All loaded names:")
for s in students:
    print(f" - {s['name_kz']}")

target_fragment = "Амангелдина" # Searching for this fragment
print(f"\nSearching for '{target_fragment}' in names...")

found = False
for s in students:
    # Case insensitive search
    if target_fragment.lower() in s['name_kz'].lower():
        print(f"Found Student: {s['name_kz']}")
        print("-" * 30)
        # Sort by subject name to make it readable
        sorted_subjects = sorted(s['subjects_list'], key=lambda x: x['name_ru'])
        
        for subj in sorted_subjects:
            print(f"Subject (RU): {subj['name_ru']}")
            print(f"  Score: {subj['score']}")
            print(f"  Letter: {subj['letter']}")
            print(f"  Point: {subj['point']}")
        
        # Also print raw dictionary for copy-pasting into replacement script
        print("-" * 30)
        print("RAW DATA FOR SCRIPT:")
        print(json.dumps(s['subjects'], ensure_ascii=False, indent=2))
        found = True
        break

if not found:
    print("Student not found.")
