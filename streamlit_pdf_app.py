import streamlit as st
import os
import zipfile
import io
import sys

# Ensure local modules are found
sys.path.append(os.path.join(os.path.dirname(__file__), 'attestat_generator'))
from attestat_generator.new_data_loader import NewDataLoader
from attestat_generator.pdf_generator import PDFGenerator

st.set_page_config(page_title="PDF Attestat Generator", layout="wide")

st.title("🎓 Генератор Аттестатов (PDF)")
st.markdown("""
Этот инструмент автоматически создает PDF аттестаты на основе загруженного Excel файла.
Структура предметов и модулей определяется **автоматически**.
""")

# Sidebar
st.sidebar.header("Настройки")
uploaded_file = st.sidebar.file_uploader("Загрузите Excel файл (Полотно)", type=['xlsx'])

# Check for Template Images
bg_path = os.path.join("data", "шаблон каз.jpg")
if not os.path.exists(bg_path):
    st.error(f"❌ Шаблон фона не найден: {bg_path}")
    st.info("Пожалуйста, убедитесь, что файл 'шаблон каз.jpg' находится в папке 'data/'.")

if uploaded_file and os.path.exists(bg_path):
    try:
        # Load Data
        with st.spinner("Чтение Excel файла..."):
            # Save uploaded file temporarily because openpyxl needs a path or file-like object
            # NewDataLoader expects a path currently, let's fix that or save temp
            temp_path = "temp_uploaded.xlsx"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
                
            loader = NewDataLoader(temp_path)
            students = loader.load_data()
            
        st.success(f"✅ Загружено {len(students)} студентов.")
        
        # Preview
        with st.expander("Просмотр списка студентов"):
            for s in students:
                st.write(f"**{s['full_name']}**: {len(s['subjects_list'])} предметов")
                
        # Generate Button
        if st.button("🚀 Сгенерировать PDF для всех"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Zip Buffer
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                generator = PDFGenerator(background_image_path=bg_path)
                
                for i, student in enumerate(students):
                    status_text.text(f"Обработка: {student['full_name']}")
                    
                    # Generate PDF in memory
                    pdf_buffer = io.BytesIO()
                    # We need to adapt PDFGenerator to accept a file-like object or save to temp
                    # Current PDFGenerator takes output_path string.
                    # We will save to a temp file and read it back.
                    
                    safe_name = "".join([c for c in student['name_kz'] if c.isalpha() or c.isdigit() or c==' ']).strip()
                    temp_pdf_name = f"temp_{i}.pdf"
                    
                    generator.generate(student, temp_pdf_name)
                    
                    # Read and add to zip
                    with open(temp_pdf_name, "rb") as f:
                        zf.writestr(f"{safe_name}.pdf", f.read())
                    
                    # Cleanup
                    os.remove(temp_pdf_name)
                    
                    progress_bar.progress((i + 1) / len(students))
            
            st.success("🎉 Генерация завершена!")
            
            # Download
            st.download_button(
                label="⬇️ Скачать PDF Аттестаты (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="attestats_pdf.zip",
                mime="application/zip"
            )
            
            # Cleanup excel
            if os.path.exists(temp_path):
                os.remove(temp_path)

    except Exception as e:
        st.error(f"Ошибка: {e}")
        import traceback
        st.text(traceback.format_exc())

else:
    st.info("👈 Пожалуйста, загрузите Excel файл в меню слева.")

