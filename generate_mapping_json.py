import pandas as pd
import json
import os
import sys

# Force UTF-8 for any console output
sys.stdout.reconfigure(encoding='utf-8')

INPUT_FILE = '../data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
OUTPUT_FILE = 'subjects_mapping.json'

def generate_mapping():
    if not os.path.exists(INPUT_FILE):
        # Fallback for dev environment or if running from root
        potential_path = 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
        if os.path.exists(potential_path):
           file_path = potential_path
        else:
            # Try absolute path from previous context if needed, but relative is safer
            file_path = r'c:/Users/УГТК/Desktop/Atestat 2026/data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
    else:
        file_path = INPUT_FILE
        
    try:
        print(f"Reading {file_path}...")
        df = pd.read_excel(file_path, sheet_name='Лист2', header=10)
        
        mapping = {}
        
        # Taking columns 2 to 60 (as per user's earlier script logic implies)
        # Verify bounds based on list_columns output
        # Columns 2 to 59 seem to be subjects. 60 is Unnamed. 61, 62 are exams.
        
        columns_of_interest = df.columns[2:63] 
        
        print(f"Found {len(columns_of_interest)} potential subjects.")
        
        for col in columns_of_interest:
            col_name = str(col).strip()
            if "Unnamed" in col_name:
                continue
                
            # Heuristic for Russian translation (very basic, user must verify)
            ru_name = "ПЕРЕВЕСТИ" 
            if "Орыс тілі" in col_name: ru_name = "Русский язык"
            elif "Математика" in col_name: ru_name = "Математика"
            elif "Физика" in col_name: ru_name = "Физика"
            elif "Химия" in col_name: ru_name = "Химия"
            # ... and so on. Better to let user fill it.
            
            mapping[col_name] = {
                "kz": col_name,
                "ru": ru_name
            }
            
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=4)
            
        print(f"Mapping generated at {os.path.abspath(OUTPUT_FILE)}")
        print("Please edit this file to provide correct Russian translations.")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    generate_mapping()
