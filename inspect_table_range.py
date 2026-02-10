file_path = 'attestat_generator/templates/template_kz.htm'

with open(file_path, 'r', encoding='cp1251') as f:
    content = f.read()

start_marker = ">1</span>"
start_pos = content.find(start_marker)

if start_pos != -1:
    # Find start of TR
    # Scan backwards for <tr
    tr_start = content.rfind("<tr", 0, start_pos)
    print(f"Row 1 starts at {tr_start}")
    
    # Find end of Table
    # Scan forwards for </table
    table_end = content.find("</table", start_pos)
    print(f"Table ends at {table_end}")
    
    if table_end != -1:
        chunk = content[tr_start:table_end]
        print(f"Chunk size: {len(chunk)} chars")
        print("Last 500 chars of chunk:")
        print(chunk[-500:])
        
        # Check for footer text
        if "Директор" in chunk or "Director" in chunk:
            print("WARNING: 'Director' found in chunk, might have included footer!")
    else:
        print("Table end not found")
else:
    print("Row 1 not found")
