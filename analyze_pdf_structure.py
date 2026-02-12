import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from pypdf import PdfReader

target = 'data/шаблон каз.pdf'

reader = PdfReader(target)
print(f"Total pages: {len(reader.pages)}\n")

for page_num, page in enumerate(reader.pages, 1):
    print(f"=== PAGE {page_num} ===")
    
    text_elements = []
    
    def visitor_body(text, cm, tm, fontDict, fontSize):
        y = tm[5]
        x = tm[4]
        if text.strip():
            text_elements.append((y, x, text.strip()))
    
    page.extract_text(visitor_text=visitor_body)
    
    # Sort by Y (descending) then X
    text_elements.sort(key=lambda item: (-item[0], item[1]))
    
    # Group by approximate Y coordinate (within 2 pixels)
    lines = []
    current_line = []
    last_y = None
    
    for y, x, text in text_elements:
        if last_y is None or abs(y - last_y) < 2:
            current_line.append((x, text))
            last_y = y
        else:
            if current_line:
                lines.append((last_y, current_line))
            current_line = [(x, text)]
            last_y = y
    
    if current_line:
        lines.append((last_y, current_line))
    
    # Print lines
    for y, line_items in lines:
        line_items.sort(key=lambda item: item[0])  # Sort by X
        line_text = ' '.join([text for x, text in line_items])
        print(f"Y={y:6.2f}: {line_text}")
    
    print()
