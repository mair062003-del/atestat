import os

file_path = 'attestat_generator/templates/template_kz.htm'

try:
    # Use cp1251 as seen in the meta tag
    with open(file_path, 'r', encoding='cp1251', errors='replace') as f:
        content = f.read()
        
    print(f"File size: {len(content)} bytes")
    
    keywords = ["Аманжолова", "2129475"]
    print("\n--- Searching Keywords (cp1251) ---")
    for kw in keywords:
        pos = content.find(kw)
        if pos != -1:
            print(f"Found '{kw}' at position {pos}")
            start = max(0, pos - 50)
            end = min(len(content), pos + 100)
            print(f"Context repr: {repr(content[start:end])}")
        else:
            print(f"'{kw}' not found.")
            
except Exception as e:
    print(f"Error: {e}")
