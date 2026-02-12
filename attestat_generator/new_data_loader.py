import openpyxl
import sys
import os

class NewDataLoader:
    """Data loader for the new Excel structure (ПОЛОТНО - 4аКШО-тексерілді.xlsx)"""
    
    def __init__(self, excel_file):
        self.excel_file = excel_file
        
    def load_data(self):
        """Load student data from the new Excel format"""
        print("Loading data from new Excel format...")
        
        try:
            # Load workbook
            wb = openpyxl.load_workbook(self.excel_file, data_only=True)
            # Use index 0 instead of name 'экви' to be safe
            sheet = wb.worksheets[0]
            
            students = []
            
            # Row indices (openpyxl is 1-based)
            # Row 9 -> Module Names
            # Row 10 -> Subject Names
            # Row 11 -> Hours
            # Row 12+ -> Student Data
            
            # Extract headers
            # Get all values from the row to ensure we handle sparse/merged cells correctly
            # openpyxl rows are generators
            
            row_modules = [cell.value for cell in sheet[9]]
            row_subjects = [cell.value for cell in sheet[10]]
            row_hours = [cell.value for cell in sheet[11]]
            
            # Pre-process subjects metadata
            subjects_meta = []
            current_module = "ЖБП 00. Жалпы білім беретін пәндер" # Default initial module
            
            # Start from column D (index 3 in 0-based list)
            # Ensure we don't go out of bounds
            max_col = len(row_subjects)
            
            for col_idx in range(3, max_col):
                subj_name = row_subjects[col_idx]
                
                # Stop if no subject name (end of data)
                if not subj_name:
                    continue
                
                # Check for module in this column
                mod_val = row_modules[col_idx]
                if mod_val:
                    current_module = str(mod_val).strip().replace('\n', ' ')
                
                subjects_meta.append({
                    'col_idx': col_idx,
                    'name_kz': str(subj_name).strip(),
                    'module': current_module,
                    'hours': row_hours[col_idx]
                })

            print(f"Found {len(subjects_meta)} subjects.")

            # Iterate through students (Row 12 to 150)
            for row_idx in range(12, 150):
                # row_idx is 1-based, sheet[row_idx] returns tuple of cells
                row_cells = sheet[row_idx]
                
                if not row_cells:
                    continue
                
                # Check if row has data (check ID or Name columns)
                # Cell indices in tuple are 0-based
                # Column A (index 0), Column B (index 1) - Name
                if len(row_cells) < 2:
                     continue

                student_name = row_cells[1].value
                if not student_name:
                    continue
                    
                student = {
                    'id': row_cells[0].value,
                    'full_name': student_name, 
                    'name_kz': student_name,
                    'subjects_list': []
                }
                
                # Collect grades for this student
                for meta in subjects_meta:
                    if meta['col_idx'] < len(row_cells):
                        cell_value = row_cells[meta['col_idx']].value
                    else:
                        cell_value = None
                    
                    # Convert score to grade point
                    score_num = 0
                    try:
                        if cell_value and str(cell_value).strip():
                             score_num = int(float(cell_value))
                    except:
                        pass
                        
                    letter, point = self._get_grade_info(score_num)
                    
                    # Handle hours
                    h_val = meta['hours']
                    h_str = str(h_val) if h_val is not None else ""
                    
                    student['subjects_list'].append({
                        'module': meta['module'], # Attach module name to subject
                        'name_kz': meta['name_kz'],
                        'hours': h_str,
                        'score': score_num,
                        'letter': letter,
                        'point': point
                    })
                
                students.append(student)
            
            print(f"Loaded {len(students)} students.")
            return students
            
        except Exception as e:
            print(f"Error loading new Excel format: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _get_grade_info(self, raw_score):
        """Convert score to letter and point using standard scale"""
        try:
            score = float(raw_score)
            if score >= 95: return 'A', '4.0'
            if score >= 90: return 'A', '4.0' # Some variations exist, usually >=95 is A, 90-94 A- but user data suggests 90 is A
            if score >= 85: return 'A-', '3.67'
            if score >= 80: return 'B+', '3.33'
            if score >= 75: return 'B', '3.0'
            if score >= 70: return 'B-', '2.67'
            if score >= 65: return 'C+', '2.33'
            if score >= 60: return 'C', '2.0'
            if score >= 55: return 'C-', '1.67'
            if score >= 50: return 'D+', '1.33'
            if score >= 25: return 'F', '0.0' # Fail
            if score > 0: return 'F', '0.0'
            return 'сын', '0.0' # сыналды / passed (if score is 0 or non-numeric)
        except:
             return 'сын', '0.0'
