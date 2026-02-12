import sys
import io
from pypdf import PdfReader

# Force UTF-8 for output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def extract_coords(pdf_path):
    reader = PdfReader(pdf_path)
    for i, page in enumerate(reader.pages):
        print(f"--- Page {i+1} ---")
        
        def visitor_body(text, cm, tm, font_dict, font_size):
            cleaned_text = text.strip()
            if cleaned_text:
                # tm[4] is x, tm[5] is y
                print(f"[{tm[4]:2.2f}, {tm[5]:2.2f}] {cleaned_text}")

        page.extract_text(visitor_text=visitor_body)

if __name__ == "__main__":
    pdf_path = 'data/шаблон каз.pdf' # Use the template provided by user
    extract_coords(pdf_path)
