try:
    with open('chunk.txt', 'rb') as f:
        content = f.read()
    
    text = content.decode('cp1251', errors='replace')
    print(text)
except Exception as e:
    print(f"Error: {e}")
