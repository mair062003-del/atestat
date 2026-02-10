import re
import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# We need to map Subject Name to (Page, Y)
# We also need X coordinates for columns (Hours, Score, Letter, Point)
# Based on analysis:
# Page 1 Left Cols: Hours=184, Score=259, Letter=281, Point=319
# Page 1 Right Cols: Hours=609, Score=685, Letter=708, Point=746
# Page 2 Left Cols: Hours=179, Score=253, Letter=275, Point=313
# Page 2 Right Cols (Empty in log? No, Page 2 seems to be single column in log?)

# Wait, Page 2 in layout:
# X=21.60 for Index.
# X=39.72 for Name.
# X=179 for Hours.
# X=252 for Score.
# It seems Page 2 in the log corresponds to ONE SIDE of the visual page 2.
# Let's verify Page 2 width.
# Page Size: 841 x 595.
# Page 2 text X goes up to ~350.
# So Page 2 in PDF is HALF?
# "--- Page 2 ---"
# Text X max is ~348.
# This implies Page 2 in the PDF might be just the Left side?
# Or maybe the Right side is empty?
# Let's assume standard layout.

mapping = {
    "page_1_left": {"hours": 184, "score": 259, "letter": 281, "point": 319},
    "page_1_right": {"hours": 609, "score": 685, "letter": 708, "point": 746},
    "page_2_left": {"hours": 179, "score": 253, "letter": 275, "point": 313},
    "subjects": {}
}

current_page = 1

with open('pdf_layout.txt', 'r', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        if "--- Page 2 ---" in line:
            current_page = 2
            continue
        
        # TEXT: 'Subject Name' | X: 123 | Y: 456
        match = re.search(r"TEXT: '(.+?)' \| X: ([\d\.]+) \| Y: ([\d\.]+)", line)
        if match:
            text = match.group(1).strip()
            x = float(match.group(2))
            y = float(match.group(3))
            
            # Heuristics to identify Subjects
            # Usually starts with capital letter, X is near 43 (Page 1 Left) or 467 (Page 1 Right) or 39 (Page 2)
            # and is NOT a number.
            
            is_subject = False
            region = ""
            
            # Page 1 Left
            if current_page == 1 and 40 < x < 50:
                is_subject = True
                region = "page_1_left"
            # Page 1 Right
            elif current_page == 1 and 460 < x < 475:
                is_subject = True
                region = "page_1_right"
            # Page 2 Left (assuming it's left)
            elif current_page == 2 and 35 < x < 45:
                is_subject = True
                region = "page_2_left"
            
            if is_subject and not text[0].isdigit():
                # Filter out garbage
                if len(text) > 3:
                     mapping["subjects"][text] = {
                         "page": current_page,
                         "y": y,
                         "region": region
                     }

print(json.dumps(mapping, indent=2, ensure_ascii=False))
