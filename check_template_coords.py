import sys
import io
from pypdf import PdfReader

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Read the template PDF
pdf_path = 'data/шаблон каз.pdf'

try:
    reader = PdfReader(pdf_path)
    print(f"PDF найден: {len(reader.pages)} страниц")
    print()
    
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        print(f"=== Страница {page_num + 1} ===")
        
        # Get page dimensions
        mediabox = page.mediabox
        width = float(mediabox.width)
        height = float(mediabox.height)
        print(f"Размер страницы: {width} x {height} points")
        print()
        
        # Extract text with positions
        text = page.extract_text()
        
        # Try to find specific patterns
        if '№' in text or 'Пәннің атауы' in text:
            print("Найдены заголовки таблицы!")
            
        print("Первые 500 символов текста:")
        print(text[:500])
        print()
        print("-" * 80)
        print()

except FileNotFoundError:
    print(f"Файл не найден: {pdf_path}")
    print("Проверьте путь к файлу.")
except Exception as e:
    print(f"Ошибка при чтении PDF: {e}")
