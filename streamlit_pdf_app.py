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

if st.sidebar.button("🔄 Сбросить загруженные данные"):
    for key in ['students_data', 'layout_overrides']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

# Check for Template Images
bg_path = os.path.join("data", "шаблон каз.jpg")
if not os.path.exists(bg_path):
    st.error(f"❌ Шаблон фона не найден: {bg_path}")
    st.info("Пожалуйста, убедитесь, что файл 'шаблон каз.jpg' находится в папке 'data/'.")

if uploaded_file and os.path.exists(bg_path):
    try:
        # Load Data if not already loaded or if file changed
        # For simplicity in this app, we reload if students_data is missing
        if 'students_data' not in st.session_state:
            with st.spinner("Чтение Excel файла..."):
                # Save uploaded file temporarily because openpyxl needs a path or file-like object
                temp_path = "temp_uploaded.xlsx"
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                    
                loader = NewDataLoader(temp_path)
                students = loader.load_data()
                
                # Cleanup excel
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            
            st.session_state.students_data = students
            st.success(f"✅ Загружено {len(students)} студентов.")
        
        # --- TABBED INTERFACE ---
        tab1, tab2, tab3 = st.tabs(["📝 Редактор данных", "🎨 Настройка макета", "🚀 Генерация"])
        
        # --- TAB 1: DATA EDITOR ---
        with tab1:
            st.info("Здесь можно изменить ФИО и данные предметов перед генерацией.")
            
            # Student Selector for Data Editor
            selected_student_idx = st.selectbox(
                "Выберите студента для редактирования:", 
                range(len(st.session_state.students_data)), 
                format_func=lambda i: st.session_state.students_data[i]['full_name'],
                key="data_editor_student_select"
            )
            
            student = st.session_state.students_data[selected_student_idx]
            
            # Edit Metadata
            col1, col2 = st.columns(2)
            new_name_kz = col1.text_input("ФИО (KZ)", student['name_kz'])
            new_doc_num = col2.text_input("№ Документа", str(student.get('document_number', '')))
            
            if new_name_kz != student['name_kz']:
                student['name_kz'] = new_name_kz
                st.success("ФИО обновлено!")
            
            if new_doc_num != str(student.get('document_number', '')):
                student['document_number'] = new_doc_num
                st.success("Номер документа обновлен!")
                
                # Auto-increment for subsequent students
                # Only if input is numeric to avoid errors with complex strings
                if new_doc_num.isdigit():
                    try:
                        start_num = int(new_doc_num)
                        count_updated = 0
                        
                        # Iterate from next student to end
                        for i in range(selected_student_idx + 1, len(st.session_state.students_data)):
                            offset = i - selected_student_idx
                            next_num = start_num + offset
                            
                            # Preserve zero-padding (e.g. "005" -> "006")
                            next_str = str(next_num)
                            if len(new_doc_num) > len(next_str):
                                 next_str = next_str.zfill(len(new_doc_num))
                            
                            st.session_state.students_data[i]['document_number'] = next_str
                            count_updated += 1
                        
                        if count_updated > 0:
                            st.info(f"🔢 Автоматически заполнены номера для {count_updated} следующих студентов.")
                    except:
                        pass # Ignore errors in auto-increment logic

            # Edit Extended Metadata (Institution, Specialty, etc.)
            with st.expander("🏢 Данные организации и специальности"):
                col_inst, col_spec = st.columns(2)
                
                # Defaults
                def_inst = "“Орал гуманитарлық-техникалық колледжі” мекемесінде"
                def_spec = "01140100 “Бастауыш білім беру педагогикасы мен әдістемесі” мамандығында"
                def_qual = "4S01140101 “Бастауыш білім беру мұғалімі”"
                def_qual2 = "біліктілігі бойынша"
                
                # Get current values
                curr_inst = student.get('institution', def_inst)
                curr_spec = student.get('specialty', def_spec)
                curr_qual = student.get('qualification', def_qual)
                curr_qual2 = student.get('qualification_2', def_qual2)
                
                new_inst = col_inst.text_area("Организация:", curr_inst, height=68)
                new_spec = col_spec.text_area("Специальность:", curr_spec, height=68)
                new_qual = col_inst.text_input("Квалификация:", curr_qual)
                new_qual2 = col_spec.text_input("Квалификация (стр 2):", curr_qual2)
                
                if new_inst != curr_inst: 
                    student['institution'] = new_inst
                if new_spec != curr_spec: 
                    student['specialty'] = new_spec
                if new_qual != curr_qual: 
                    student['qualification'] = new_qual
                if new_qual2 != curr_qual2: 
                    student['qualification_2'] = new_qual2

            # Edit Subjects (DataFrame)
            st.subheader("Предметы")
            
            # Convert subjects to simple flat list for editor
            subjects_flat = []
            for s in student['subjects_list']:
                subjects_flat.append({
                    "Модуль": s.get('module', ''),
                    "Предмет": s['name_kz'],
                    "Часы": s.get('hours', ''),
                    "Оценка": s.get('score', ''),
                    "Тамға": s.get('letter', ''),
                    "Балл": s.get('point', '')
                })
            
            edited_df = st.data_editor(subjects_flat, num_rows="dynamic", use_container_width=True)
            
            # Save back to session state
            if st.button("💾 Сохранить изменения предметов"):
                new_subjects_list = []
                for row in edited_df:
                    new_subjects_list.append({
                        "module": row["Модуль"],
                        "name_kz": row["Предмет"],
                        "hours": row["Часы"],
                        "score": row["Оценка"],
                        "letter": row["Тамға"],
                        "point": row["Балл"]
                    })
                st.session_state.students_data[selected_student_idx]['subjects_list'] = new_subjects_list
                st.success("Список предметов сохранен!")

        # --- TAB 2: VISUAL EDITOR (LAYOUT) ---
        with tab2:
            st.info("⬅️ Слева настройки смещения. ➡️ Справа результат (прокрутите вниз если не видно).")
            
            if 'layout_overrides' not in st.session_state:
                st.session_state.layout_overrides = {}

            # Student Selector
            layout_student_idx = st.selectbox(
                "Выберите студента для настройки макета:", 
                range(len(st.session_state.students_data)), 
                format_func=lambda i: st.session_state.students_data[i]['full_name'],
                key="layout_student_select"
            )
            
            student_id = str(layout_student_idx)
            student_layout = st.session_state.students_data[layout_student_idx]
            
            # Ensure overrides exist
            if student_id not in st.session_state.layout_overrides:
                st.session_state.layout_overrides[student_id] = {'subjects': {}}
            
            current_overrides = st.session_state.layout_overrides[student_id]

            # --- SPLIT LAYOUT ---
            col_left, col_right = st.columns([1, 1.5]) # Left for controls, Right for PDF
            
            with col_left:
                st.subheader("🛠️ Настройка предмета")
                
                # Navigation Controls (Prev/Next)
                # Store selected subject index in session state to allow buttons to change it
                if f"sel_subj_{student_id}" not in st.session_state:
                    st.session_state[f"sel_subj_{student_id}"] = 0
                
                current_subj_idx = st.session_state[f"sel_subj_{student_id}"]
                subjs_count = len(student_layout['subjects_list'])
                
                c_prev, c_sel, c_next = st.columns([1, 2, 1])
                with c_prev:
                    if st.button("⬅️ Пред", key="btn_prev", use_container_width=True):
                         if current_subj_idx > 0:
                             st.session_state[f"sel_subj_{student_id}"] -= 1
                             st.rerun()
                with c_next:
                    if st.button("След ➡️", key="btn_next", use_container_width=True):
                         if current_subj_idx < subjs_count - 1:
                             st.session_state[f"sel_subj_{student_id}"] += 1
                             st.rerun()
                
                # Dropdown sync
                selected_subj_idx = c_sel.selectbox(
                    "Выбрать:", 
                    range(subjs_count),
                    index=st.session_state[f"sel_subj_{student_id}"],
                    format_func=lambda i: f"{i+1}. {student_layout['subjects_list'][i]['name_kz'][:20]}...",
                    key=f"dropdown_{student_id}"
                )
                
                # Update state if dropdown changed
                if selected_subj_idx != st.session_state[f"sel_subj_{student_id}"]:
                     st.session_state[f"sel_subj_{student_id}"] = selected_subj_idx
                     st.rerun()
                
                # --- ACTIVE SUBJECT CONTROLS ---
                subj_data = student_layout['subjects_list'][selected_subj_idx]
                current_ov = current_overrides['subjects'].get(selected_subj_idx, {})
                
                st.info(f"**Предмет:** {subj_data['name_kz']}\n\n**Модуль:** {subj_data.get('module', '-')}")
                
                # 1. Y-Offset Slider (Real-time)
                # Range +/- 50pt
                current_y = float(current_ov.get('y_offset', 0.0))
                new_y = st.slider(
                    "↕️ Сдвиг по вертикали (Y)", 
                    min_value=-50.0, 
                    max_value=50.0, 
                    value=current_y,
                    step=1.0,
                    help="Вниз (-) или Вверх (+). Результат обновляется сразу."
                )
                
                # 2. Force Break Toggle
                is_break = bool(current_ov.get('force_break', False))
                new_break = st.checkbox("⤵️ Новая колонка (начать с начала)", value=is_break)
                
                # 3. Custom Text Input
                current_text = str(current_ov.get('custom_text', ''))
                # Show placeholder as original name
                st.caption("Используйте символ `|` для переноса строки (например: `Алғашқы әскери|және технологи...`)")
                new_text = st.text_input("✏️ Свой текст (для переносов)", value=current_text, placeholder=subj_data['name_kz'])
                
                # -- UPDATE STATE LOGIC --
                # Check for changes
                changes_made = False
                if new_y != current_y:
                    current_ov['y_offset'] = new_y
                    changes_made = True
                
                if new_break != is_break:
                    current_ov['force_break'] = new_break
                    changes_made = True
                    
                if new_text != current_text:
                    current_ov['custom_text'] = new_text
                    changes_made = True
                
                if changes_made:
                    # Update session state override
                    if 'subjects' not in st.session_state.layout_overrides[student_id]:
                         st.session_state.layout_overrides[student_id]['subjects'] = {}
                    
                    st.session_state.layout_overrides[student_id]['subjects'][selected_subj_idx] = current_ov
                    # RERUN to update PDF immediately
                    st.rerun()

                st.write("---")
                st.caption("Изменения применяются автоматически.")


            with col_right:
                st.subheader("📄 Предпросмотр (PDF)")
                
                # Generate PDF in memory
                generator = PDFGenerator(background_image_path=bg_path)
                preview_buffer = io.BytesIO()
                
                # Use a temp file because reportlab canvas needs a filename or file-like object
                # But my PDFGenerator implementation (checked earlier) accepts output_path as string for canvas constructor
                # Wait, canvas.Canvas(filename) accepts string OR file-like object. 
                # My implementation: c = canvas.Canvas(output_path, ...)
                # So passing io.BytesIO() should work!
                
                # However, earlier I used a temp string path. Let's stick to temp file to be practically safe with image loading etc.
                preview_temp_path = f"preview_{student_id}.pdf"
                
                try:
                    generator.generate(student_layout, preview_temp_path, overrides=st.session_state.layout_overrides[student_id])
                    
                    # Read back
                    with open(preview_temp_path, "rb") as f:
                        pdf_bytes = f.read()
                    
                    # Embed PDF
                    import base64
                    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                    pdf_display = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf">'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                    
                except Exception as e:
                    st.error(f"Ошибка генерации превью: {e}")
                finally:
                    if os.path.exists(preview_temp_path):
                        os.remove(preview_temp_path)

        # --- TAB 3: GENERATION ---
        with tab3:
            st.write("Генерация всех аттестатов с учетом ваших изменений.")
            
            if st.button("🚀 Сгенерировать ВСЕ PDF"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    generator = PDFGenerator(background_image_path=bg_path)
                    
                    students_source = st.session_state.students_data
                    
                    for i, student in enumerate(students_source):
                        status_text.text(f"Обработка: {student['full_name']}")
                        
                        pdf_buffer = io.BytesIO()
                        
                        safe_name = "".join([c for c in student['name_kz'] if c.isalpha() or c.isdigit() or c==' ']).strip()
                        temp_pdf_name = f"temp_{i}.pdf"
                        
                        # Apply overrides
                        s_id = str(i)
                        overrides = st.session_state.layout_overrides.get(s_id, {})
                        
                        generator.generate(student, temp_pdf_name, overrides=overrides)
                        
                        with open(temp_pdf_name, "rb") as f:
                            zf.writestr(f"{safe_name}.pdf", f.read())
                        
                        os.remove(temp_pdf_name)
                        progress_bar.progress((i + 1) / len(students_source))
                
                st.success("🎉 Генерация завершена!")
                st.download_button(
                    label="⬇️ Скачать Архив (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="attestats_final.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Ошибка: {e}")
        import traceback
        st.text(traceback.format_exc())

else:
    st.info("👈 Пожалуйста, загрузите Excel файл в меню слева.")

