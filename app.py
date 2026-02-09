import streamlit as st
import pandas as pd
import os
import json
import zipfile
import io
import datetime
from docxtpl import DocxTemplate
from data_loader import DataLoader

# Set page config
st.set_page_config(page_title="Attestat Generator", layout="wide")

st.title("🎓 Генератор Аттестатов")
st.markdown("Загрузите Excel файл и шаблон, чтобы создать аттестаты.")

# Sidebar for Setup
st.sidebar.header("Настройки")

# 1. Subject Mapping
mapping_path = 'subjects_mapping.json'
if os.path.exists(mapping_path):
    st.sidebar.success(f"✅ Файл переводов найден ({mapping_path})")
else:
    st.sidebar.error(f"❌ Файл переводов не найден ({mapping_path})")
    st.sidebar.info("Пожалуйста, убедитесь, что subjects_mapping.json находится в папке приложения.")

# 2. Upload Excel
uploaded_excel = st.file_uploader("1. Загрузите Excel файл (Полотно)", type=['xlsx'])

# 3. Upload Template (Optional, default provided)
uploaded_template = st.file_uploader("2. Загрузите шаблон Word (.docx)", type=['docx'])

# Logic
if uploaded_excel and os.path.exists(mapping_path):
    try:
        # Load Data
        loader = DataLoader(uploaded_excel, mapping_path)
        students = loader.load_data()
        
        st.success(f"Загружено {len(students)} студентов")
        
        # Preview Data
        if st.checkbox("Показать найденных студентов"):
            preview_data = []
            for s in students:
                preview_data.append({
                    "ФИО (KZ)": s['name_kz'],
                    "ФИО (RU)": s['name_ru'],
                    "Предметов": len(s['subjects_list'])
                })
            st.dataframe(pd.DataFrame(preview_data))

        # Selection
        st.subheader("Генерация")
        
        # Template Handling
        template_file = None
        if uploaded_template:
            template_file = uploaded_template
        elif os.path.exists('templates/template.docx'):
            template_file = 'templates/template.docx'
            st.info("Используется стандартный шаблон из папки templates/")
        
        if not template_file:
            st.error("Шаблон не найден! Загрузите файл .docx или поместите его в папку templates/")
        else:
            if st.button("Сгенерировать аттестаты"):
                progress_bar = st.progress(0)
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for i, student in enumerate(students):
                        # Update Progress
                        progress_bar.progress((i + 1) / len(students))
                        
                        # Generate Doc
                        doc = DocxTemplate(template_file)
                        
                        # Fix for stream reuse if using uploaded file
                        if hasattr(template_file, 'seek'):
                            template_file.seek(0)
                            
                        context = {
                            'student_name_kz': student['name_kz'],
                            'student_name_ru': student['name_ru'],
                            'date': datetime.datetime.now().strftime("%d.%m.%Y"),
                            'subjects': student['subjects_list'],
                            's': student['subjects']
                        }
                        
                        doc.render(context)
                        
                        # Save to memory buffer
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # Add to Zip
                        safe_name = "".join([c for c in student['name_kz'] if c.isalpha() or c.isdigit() or c==' ']).strip()
                        zf.writestr(f"{safe_name}.docx", doc_io.getvalue())
                
                progress_bar.progress(100)
                
                # Download Button
                st.success("Готово!")
                st.download_button(
                    label="⬇️ Скачать все аттестаты (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="attestats.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
        st.exception(e)

else:
    st.info("Ожидание загрузки файлов...")
