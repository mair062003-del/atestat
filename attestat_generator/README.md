# Attestat Generator

Tool for automating the generation of student attestats from Excel data into Word templates.

## Features
- **Web Interface**: Easy-to-use browser interface using Streamlit.
- **Bilingual Support**: Handles Kazakh and Russian subject names.
- **Complex Layouts**: Supports arbitrary Word template layouts.
- **Grade Conversion**: Automatically converts numeric grades to Letter grades and GPA points.

## Installation

1. Install Python 3.8+
2. Install dependencies:
   ```bash
   pip install pandas openpyxl docxtpl streamlit
   ```

## Usage

1. Run the app:
   ```bash
   streamlit run app.py
   ```
2. Upload your Excel file (Canvas) and Word Template.
3. Download the generated attestats.

## Project Structure
- `app.py`: Main web application.
- `data_loader.py`: logic for reading Excel and processing grades.
- `subjects_mapping.json`: Configuration for subject translations.
- `templates/`: Directory for default templates.
