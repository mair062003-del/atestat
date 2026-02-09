import os
import sys
from docxtpl import DocxTemplate
from data_loader import DataLoader
import datetime

# Force UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Constants
EXCEL_PATH = '../data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
MAPPING_PATH = 'subjects_mapping.json'
TEMPLATE_PATH = 'templates/template.docx'
OUTPUT_DIR = 'output'

def main():
    print("=== Attestat Generator ===")
    
    # Check paths
    if not os.path.exists(TEMPLATE_PATH):
        print(f"Error: Template not found at {TEMPLATE_PATH}")
        print("Please place your 'template.docx' in the 'templates' folder.")
        # For demonstration, we might have generated a dummy one, so this check is valid
        return

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    # Load Data
    try:
        loader = DataLoader(EXCEL_PATH, MAPPING_PATH)
        students = loader.load_data()
        print(f"Successfully loaded {len(students)} students.")
    except Exception as e:
        print(f"Error loading data: {e}")
        return

    # Process Students
    print("\nStarting generation...")
    print("Press Enter to generate for a student, 's' to skip, 'q' to quit.")
    
    for i, student in enumerate(students):
        print(f"\n[{i+1}/{len(students)}] Student: {student['name_kz']}")
        
        user_input = input("Generate? [Y/n/q]: ").strip().lower()
        if user_input == 'q':
            break
        if user_input == 'n':
            print("Skipping...")
            continue
            
        try:
            generate_attestat(student)
            print("Done.")
        except Exception as e:
            print(f"Error generating attestat: {e}")

def generate_attestat(student):
    doc = DocxTemplate(TEMPLATE_PATH)
    
    context = {
        'student_name_kz': student['name_kz'],
        'student_name_ru': student['name_ru'],
        'date': datetime.datetime.now().strftime("%d.%m.%Y"),
        'subjects': student['subjects_list'], # kept for backward compat with loops
        's': student['subjects'] # Shortcut for direct access: {{ s['Math'].score }}
    }
    
    doc.render(context)
    
    safe_name = "".join([c for c in student['name_kz'] if c.isalpha() or c.isdigit() or c==' ']).strip()
    output_path = os.path.join(OUTPUT_DIR, f"{safe_name}.docx")
    
    doc.save(output_path)
    print(f"Saved to: {output_path}")

if __name__ == "__main__":
    # Ensure running from correct dir
    if not os.path.exists('subjects_mapping.json'):
        print("Please run this script from the 'attestat_generator' directory.")
    else:
        main()
