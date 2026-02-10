import sys
import io
from pypdf import PdfReader

# Force stdout to utf-8 just in case we print to console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def make_visitor(f):
    def visitor(text, cm, tm, fontDict, fontSize):
        y = tm[5]
        x = tm[4]
        if text.strip():
            f.write(f"TEXT: '{text}' | X: {x:.2f} | Y: {y:.2f}\n")
    return visitor

target = 'data/шаблон каз.pdf'

try:
    with open('pdf_layout.txt', 'w', encoding='utf-8') as f:
        reader = PdfReader(target)
        visitor = make_visitor(f)
        
        for i, page in enumerate(reader.pages):
             f.write(f"--- Page {i+1} ---\n")
             page.extract_text(visitor_text=visitor)
    
    print("Dumped layout to pdf_layout.txt")

except Exception as e:
    print(f"Error: {e}")

except Exception as e:
    print(f"Error: {e}")
