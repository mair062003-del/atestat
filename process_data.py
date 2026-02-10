import pandas as pd
import os

def process_data():
    try:
        # Define file path - assuming structure is: 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
        # Adjust if needed based on actual file location
        file_path = 'data/ПОЛОТНО - 4аКШИ - 2021-2022.xlsx'
        
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return

        print(f"Reading {file_path}...")
        
        # Read Excel sheet 'Лист2', skipping first 10 rows (header is row index 10 / 11th row)
        # Note: header=10 in read_excel is equivalent to skiprows=10 + implicit header
        df = pd.read_excel(file_path, sheet_name='Лист2', header=10)
        
        # Remove empty rows based on the Student Name column (2nd column, index 1)
        # The column names are dynamically changing so we access the name of the column by position
        if len(df.columns) < 2:
            print("Error: DataFrame has fewer than 2 columns.")
            return

        name_col = df.columns[1]
        print(f"Filtering based on column: {name_col}")
        
        # Drop rows where name_col is NaN
        original_len = len(df)
        df = df.dropna(subset=[name_col])
        print(f"Dropped {original_len - len(df)} rows with missing names.")

        # Extract relevant data based on user request
        # Subjects start from 3rd column (index 2) up to 60th column (index 59 -> slice 2:60)
        if len(df.columns) >= 60:
            subjects = df.columns[2:60].tolist()
            print(f"Extracted {len(subjects)} subjects.") 
        else:
            print(f"Warning: Expected at least 60 columns, found {len(df.columns)}")
            subjects = df.columns[2:].tolist()

        students = df.iloc[:, 1].tolist()
        print(f"Extracted {len(students)} students.")

        # Ensure output directory exists
        output_dir = 'reports'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Generate reports
        print("Generating reports...")
        for index, row in df.iterrows():
            report = create_student_report(row, subjects)
            
            # Create a safe filename
            student_name = str(row[df.columns[1]]).strip()
            safe_filename = "".join([c for c in student_name if c.isalpha() or c.isdigit() or c==' ']).rstrip()
            report_path = os.path.join(output_dir, f"{safe_filename}.txt")
            
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(report)
        
        print(f"Reports generated in '{output_dir}' directory.")

    except Exception as e:
        print(f"An error occurred: {e}")

def create_student_report(row, subjects):
    name = row.iloc[1] # ФИО студента (using iloc for position-based access on row)
    
    # Шапка шаблона на основе данных из файла
    report = f"ВЫПИСКА ИЗ ВЕДОМОСТИ\n"
    report += f"Студент: {name}\n"
    report += f"Группа: 4аКШИ\n" 
    report += f"Специальность: 0111083 Шетел тілі мұғалімі\n" 
    report += f"Период обучения: 2018-2022 гг.\n" 
    report += "="*40 + "\n"
    report += f"{'Предмет':<50} | {'Оценка':<5}\n"
    report += "-"*60 + "\n"
    
    # Заполняем оценки по каждому предмету 
    for subject in subjects:
        if subject in row:
            grade = row[subject]
            if pd.notna(grade): # Если оценка выставлена
                report += f"{str(subject)[:48]:<50} | {str(grade):<5}\n"
        
    return report

if __name__ == "__main__":
    process_data()
