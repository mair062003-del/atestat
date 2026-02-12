import sys
import os

# Get current dir for absolute paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Add to sys.path to ensure local imports work
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)

try:
    from new_data_loader import NewDataLoader
    from pdf_generator import PDFGenerator
except ImportError:
    # Fallback if run from parent
    from attestat_generator.new_data_loader import NewDataLoader
    from attestat_generator.pdf_generator import PDFGenerator

import streamlit as st
from jinja2 import Template
from docxtpl import DocxTemplate
import zipfile
import io
from datetime import datetime

# Set page config
st.set_page_config(page_title="Attestat Generator", layout="wide")

st.title("📜 Генератор Аттестатов")

# Sidebar
st.sidebar.header("1. Загрузка данных")

# Use new Excel file by default
default_excel_path = os.path.join(os.path.dirname(BASE_DIR), "data", "ПОЛОТНО - 4аКШО-тексерілді.xlsx")

uploaded_excel = st.sidebar.file_uploader("Excel файл (необязательно)", type=["xlsx"])

if uploaded_excel:
    with open("temp_excel.xlsx", "wb") as f:
        f.write(uploaded_excel.getbuffer())
    excel_path = "temp_excel.xlsx"
elif os.path.exists(default_excel_path):
    excel_path = default_excel_path
    st.sidebar.success(f"Используется: {os.path.basename(default_excel_path)}")
else:
    st.sidebar.error("Excel файл не найден!")
    excel_path = None

st.sidebar.header("2. Настройки генерации")
output_format = st.sidebar.selectbox("Формат вывода", ["PDF", "HTML", "Word (.docx)"])

template_content = None
template_path = None

# Specific Template Options
if output_format == "HTML":
    st.sidebar.info("HTML генерация с шаблоном")
    uploaded_html = st.sidebar.file_uploader("HTML Шаблон", type=["htm", "html"])
    
    # Default template
    default_html_path = os.path.join(os.path.dirname(BASE_DIR), "data", "test_template.htm")
    
    if uploaded_html:
        try:
            template_content = uploaded_html.getvalue().decode('cp1251')
        except:
            try:
                template_content = uploaded_html.getvalue().decode('utf-8', errors='ignore')
            except:
                template_content = uploaded_html.getvalue().decode('windows-1251', errors='ignore')
    elif os.path.exists(default_html_path):
        with open(default_html_path, 'r', encoding='utf-8') as f:
            template_content = f.read()
        st.sidebar.success("Используется: test_template.htm")

elif output_format == "PDF":
    st.sidebar.info("PDF генерируется на основе фона.")
    bg1 = st.sidebar.file_uploader("Фон Стр. 1 (JPG)", type=["jpg", "jpeg"])
    bg2 = st.sidebar.file_uploader("Фон Стр. 2 (JPG)", type=["jpg", "jpeg"])
    
    default_data_dir = os.path.join(os.path.dirname(BASE_DIR), "data")
    if not os.path.exists(default_data_dir):
        default_data_dir = os.path.join(BASE_DIR, "data")

    bg1_path = os.path.join(default_data_dir, "шаблон каз.jpg")
    bg2_path = os.path.join(default_data_dir, "шаблон каз2.jpg")
    
    # Handle uploads
    if bg1:
        with open("temp_bg1.jpg", "wb") as f:
            f.write(bg1.getbuffer())
        bg1_path = "temp_bg1.jpg"
    if bg2:
        with open("temp_bg2.jpg", "wb") as f:
            f.write(bg2.getbuffer())
        bg2_path = "temp_bg2.jpg"

elif output_format == "Word (.docx)":
    uploaded_docx = st.sidebar.file_uploader("Word Шаблон", type=["docx"])
    if uploaded_docx:
        with open("temp_template.docx", "wb") as f:
            f.write(uploaded_docx.getbuffer())
        template_path = "temp_template.docx"
    else:
        st.sidebar.warning("Загрузите .docx шаблон (с {{tags}})")

# Main Logic
if st.button("🚀 Сгенерировать"):
    if not excel_path:
        st.error("Пожалуйста, загрузите Excel файл")
    else:
        with st.spinner("Загрузка данных..."):
            try:
                loader = NewDataLoader(excel_path)
                students = loader.load_data()
                st.success(f"Загружено {len(students)} студентов")
            except Exception as e:
                st.error(f"Ошибка загрузки данных: {e}")
                st.stop()
        
        # Validation
        valid = True
        if output_format == "HTML" and not template_content:
            st.error("Ошибка: HTML шаблон не загружен.")
            valid = False
        if output_format == "Word (.docx)" and not template_path:
            st.error("Ошибка: Шаблон Word не загружен.")
            valid = False
        
        if valid:
            zip_buffer = io.BytesIO()
            progress_bar = st.progress(0)
            
            # Pre-compile Jinja if HTML
            jinja_template = None
            if output_format == "HTML":
                jinja_template = Template(template_content)
            
            # Pre-init PDF Gen if PDF
            pdf_gen = None
            if output_format == "PDF":
                pdf_gen = PDFGenerator(bg1_path)
                pdf_gen.bg2_path = bg2_path
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for i, student in enumerate(students):
                    progress_bar.progress((i + 1) / len(students))
                    
                    safe_name = student['name_kz'].replace(' ', '_').replace('/', '_')
                    
                    if output_format == "HTML":
                        html_out = jinja_template.render(student=student)
                        zf.writestr(f"{safe_name}.html", html_out.encode('utf-8'))
                    
                    elif output_format == "Word (.docx)":
                        doc = DocxTemplate(template_path)
                        doc.render(student)
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        zf.writestr(f"{safe_name}.docx", doc_io.getvalue())
                    
                    elif output_format == "PDF":
                        temp_pdf = f"temp_{i}.pdf"
                        pdf_gen.generate(student, temp_pdf)
                        with open(temp_pdf, "rb") as f:
                            zf.writestr(f"{safe_name}.pdf", f.read())
                        if os.path.exists(temp_pdf):
                            os.remove(temp_pdf)
            
            st.success("Готово! Скачайте архив ниже.")
            st.download_button("📥 Скачать ZIP", zip_buffer.getvalue(), "attestats.zip", "application/zip")
else:
    if excel_path:
        st.info("Нажмите кнопку '🚀 Сгенерировать' для начала")
    else:
        st.info("👈 Загрузите Excel файл слева для начала.")
