import sys
import os
import io

# Force UTF-8 output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Add attestat_generator to path
sys.path.insert(0, 'attestat_generator')

from data_loader import DataLoader

# Test with the actual file
excel_file = 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
mapping_file = 'attestat_generator/subjects_mapping.json'

print(f"Testing data loading...")
print(f"Excel file exists: {os.path.exists(excel_file)}")
print(f"Mapping file exists: {os.path.exists(mapping_file)}")

try:
    loader = DataLoader(excel_file, mapping_file)
    students = loader.load_data()
    
    print(f"\n[OK] Loaded {len(students)} students")
    
    if students:
        print(f"\n[INFO] First student sample:")
        s = students[0]
        print(f"  Name (KZ): {s['name_kz']}")
        print(f"  Name (RU): {s['name_ru']}")
        print(f"  Subjects count: {len(s['subjects_list'])}")
        
        if s['subjects_list']:
            print(f"\n  First 3 subjects:")
            for subj in s['subjects_list'][:3]:
                print(f"    - {subj['name_kz']}: {subj['score']} ({subj['letter']}, {subj['point']})")
        else:
            print("  [WARNING] NO SUBJECTS FOUND!")
            
        print(f"\n  Subjects dict keys: {list(s['subjects'].keys())[:5]}")
    else:
        print("[ERROR] No students loaded!")
        
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()
