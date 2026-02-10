import os

# Target files
files = [
    'attestat_generator/templates/template_kz.htm',
    'attestat_generator/templates/template_ru.htm'
]

# Replacements (Static String -> Jinja2 Tag)
# Note: encoding is cp1251
replacements = [
    # Exact match from byte inspection
    ("Аманжолова Акмаржан Жанарбек&#1179;ызы", "{{ student_name_kz }}"), 
    ("2129475", "{{ document_number }}"),                          # Document ID
]

def fix_file(path):
    if not os.path.exists(path):
        print(f"File not found: {path}")
        return

    try:
        with open(path, 'r', encoding='cp1251') as f:
            content = f.read()
            
        new_content = content
        for old, new in replacements:
            if old in new_content:
                print(f"Replacing '{old}' with '{new}' in {path}")
                new_content = new_content.replace(old, new)
            else:
                print(f"String '{old}' not found in {path}")
        
        if new_content != content:
            with open(path, 'w', encoding='cp1251') as f:
                f.write(new_content)
            print(f"Saved updates to {path}")
        else:
            print(f"No changes needed for {path}")
            
    except Exception as e:
        print(f"Error processing {path}: {e}")

if __name__ == "__main__":
    for f in files:
        fix_file(f)
