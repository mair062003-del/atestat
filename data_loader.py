import pandas as pd
import json
import os
import sys

class DataLoader:
    def __init__(self, excel_file, mapping_path):
        """
        excel_file: str path or file-like object
        mapping_path: str path to json mapping
        """
        self.excel_file = excel_file
        self.mapping_path = mapping_path
        self.mapping = self._load_mapping()
        
    def _load_mapping(self):
        if not os.path.exists(self.mapping_path):
            raise FileNotFoundError(f"Mapping file not found: {self.mapping_path}")
        with open(self.mapping_path, 'r', encoding='utf-8') as f:
            return json.load(f)
            
    def load_data(self):
        # excel_file can be a path string or a file-like object (buffer)
        print(f"Loading data...")
        try:
            # Pandora read_excel accepts both paths and buffers
            df = pd.read_excel(self.excel_file, sheet_name='Лист2', header=10)
        except Exception as e:
            raise Exception(f"Failed to read Excel file: {e}")
            
        # Filter empty students (Column 1 is name)
        name_col = df.columns[1]
        df = df.dropna(subset=[name_col])
        
        students = []
        for index, row in df.iterrows():
            student = self._process_row(row, df.columns)
            if student:
                students.append(student)
                
        return students
        
    def _process_row(self, row, columns):
        student_name = str(row.iloc[1]).strip()
        
        student_data = {
            "name_kz": student_name,
            "name_ru": student_name,
            "subjects_list": [], # Ordered list for simple loops
            "subjects": {}       # Dictionary for direct access by RU name
        }
        
        # Iterate over mapped subjects
        for col_name, translations in self.mapping.items():
            if col_name in row:
                raw_score = row[col_name]
                if pd.notna(raw_score):
                    grade_info = self._convert_grade(raw_score)
                    
                    subject_entry = {
                        "name_kz": translations.get("kz", col_name),
                        "name_ru": translations.get("ru", col_name),
                        "score": grade_info['score'],
                        "letter": grade_info['letter'],
                        "point": grade_info['point'],
                        "raw": str(raw_score)
                    }
                    
                    student_data["subjects_list"].append(subject_entry)
                    # Use Russian name as key for dictionary access (removing spaces for safer access if needed, but dict keys can have spaces)
                    # User will use {{ subjects['Математика'].score }}
                    key = translations.get("ru", col_name).strip()
                    student_data["subjects"][key] = subject_entry
        
        return student_data

    def _convert_grade(self, raw_score):
        try:
            score = float(raw_score)
        except:
            return {"score": str(raw_score), "letter": "", "point": ""}

        # Handle 5-point scale mapping to approximate ECTS
        if score <= 5.0:
            if score >= 4.5: return {"score": "95", "letter": "A", "point": "4.0"} # 5
            if score >= 3.5: return {"score": "85", "letter": "B", "point": "3.0"} # 4
            if score >= 2.5: return {"score": "75", "letter": "C", "point": "2.0"} # 3
            return {"score": "50", "letter": "D", "point": "1.0"} # 2

        # Handle 100-point scale
        # Standard KZ scale approx: 
        # 95-100 A 4.0, 90-94 A- 3.67, 85-89 B+ 3.33, 80-84 B 3.0
        # 75-79 B- 2.67, 70-74 C+ 2.33, 65-69 C 2.0, 60-64 C- 1.67
        # 55-59 D+ 1.33, 50-54 D 1.0, 0-49 F 0
        s = int(score)
        if s >= 95: return {"score": str(s), "letter": "A", "point": "4.0"}
        if s >= 90: return {"score": str(s), "letter": "A-", "point": "3.67"}
        if s >= 85: return {"score": str(s), "letter": "B+", "point": "3.33"}
        if s >= 80: return {"score": str(s), "letter": "B", "point": "3.0"}
        if s >= 75: return {"score": str(s), "letter": "B-", "point": "2.67"}
        if s >= 70: return {"score": str(s), "letter": "C+", "point": "2.33"}
        if s >= 65: return {"score": str(s), "letter": "C", "point": "2.0"}
        if s >= 60: return {"score": str(s), "letter": "C-", "point": "1.67"}
        if s >= 55: return {"score": str(s), "letter": "D+", "point": "1.33"}
        if s >= 50: return {"score": str(s), "letter": "D", "point": "1.0"}
        return {"score": str(s), "letter": "F", "point": "0"}

if __name__ == "__main__":
    # Test run
    loader = DataLoader('../data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx', 'subjects_mapping.json')
    try:
        data = loader.load_data()
        print(f"Loaded {len(data)} students.")
        if data:
            print("First student sample:", json.dumps(data[0], ensure_ascii=False, indent=2))
    except Exception as e:
        print(e)
