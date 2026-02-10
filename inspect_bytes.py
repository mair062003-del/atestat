import os

file_path = 'attestat_generator/templates/template_kz.htm'

with open(file_path, 'rb') as f:
    content = f.read()

print(f"File size: {len(content)} bytes")

# Search for "Математика" in cp1251
target = "Математика".encode('cp1251')
pos = content.find(target)
if pos != -1:
    print(f"Found 'Математика' at {pos}")
    # Show extensive context to see the grade cells following it
    print(f"Context: {content[pos:pos+400]}") # 400 bytes forward
else:
    print("'Математика' NOT found")

# Search for "Қазақ тілі" (Kazakh)
# Note: 'қ' is \xba in cp1251 usually, but might be HTML entity
# Let's try matching the known bytes or just 'аза' if unsure
target_kz = "аза".encode('cp1251') # Fragment of Qazaq
pos = content.find(target_kz)
if pos != -1:
    print(f"Found 'аза' at {pos}")
    print(f"Context: {content[pos-20:pos+400]}")
