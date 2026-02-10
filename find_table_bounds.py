file_path = 'attestat_generator/templates/template_kz.htm'

with open(file_path, 'r', encoding='cp1251') as f:
    content = f.read()

# Search for Index 1
start_marker = ">1</span>"
start_pos = content.find(start_marker)

# Search for End Marker (e.g., "Additional info" or just a high index)
# Let's try to find "Қосымша"
end_marker = "Қосымша"
end_pos = content.find(end_marker)

if start_pos != -1:
    print(f"Start Marker found at {start_pos}")
    # Show context around start
    print(f"Start Context: {content[start_pos-200:start_pos+200]}")
else:
    print("Start Marker NOT found")

if end_pos != -1:
    print(f"End Marker found at {end_pos}")
    # Show context around end
    print(f"End Context: {content[end_pos-200:end_pos+200]}")
else:
    print("End Marker NOT found")
