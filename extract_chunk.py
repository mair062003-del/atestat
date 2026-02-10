import os

file_path = 'attestat_generator/templates/template_kz.htm'

keyword = "Математика".encode('cp1251')

with open(file_path, 'rb') as f:
    content = f.read()

pos = content.find(keyword)
if pos != -1:
    print(f"Found keyword at {pos}")
    # Extract 5000 bytes before and after to see full table structure
    start = max(0, pos - 5000)
    end = min(len(content), pos + 5000)
    chunk = content[start:end]
    
    # Save to file for viewing
    with open('chunk.txt', 'wb') as f:
        f.write(chunk)
    print("Saved chunk.txt")
else:
    print("Keyword not found")
