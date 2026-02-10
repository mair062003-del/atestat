file_path = 'attestat_generator/templates/template_kz.htm'

with open(file_path, 'r', encoding='cp1251') as f:
    content = f.read()

start_marker = ">1</span>"
start_pos = content.find(start_marker)

if start_pos != -1:
    tr_start = content.rfind("<tr", 0, start_pos)
    tr_end = content.find("</tr>", start_pos) + 5
    
    row_html = content[tr_start:tr_end]
    print(f"Row HTML ({len(row_html)} chars):")
    print(row_html)
else:
    print("Row 1 not found")
