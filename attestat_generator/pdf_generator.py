from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import json
import os
import sys

# Constants
# A4 Landscape: 841.89 x 595.28 points
PAGE_WIDTH, PAGE_HEIGHT = landscape(A4)

# Load Positions
# Get the parent directory (project root)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
POSITIONS_PATH = os.path.join(os.path.dirname(SCRIPT_DIR), 'pdf_positions.json')

try:
    with open(POSITIONS_PATH, 'r', encoding='utf-8') as f:
        POSITIONS = json.load(f)
except Exception as e:
    print(f"Error loading positions from {POSITIONS_PATH}: {e}")
    POSITIONS = {}

# Register Fonts
try:
    pdfmetrics.registerFont(TTFont('Arial', 'C:\\Windows\\Fonts\\arial.ttf'))
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'C:\\Windows\\Fonts\\arialbd.ttf'))
    GLOBAL_FONT_NAME = 'Arial'
    GLOBAL_FONT_BOLD_NAME = 'Arial-Bold'
except:
    print("Warning: Arial font not found. Using Helvetica.")
    GLOBAL_FONT_NAME = 'Helvetica'
    GLOBAL_FONT_BOLD_NAME = 'Helvetica-Bold'

class PDFGenerator:
    def __init__(self, background_image_path=None):
        self.bg_path = background_image_path
        self.FONT_NAME = GLOBAL_FONT_NAME
        self.FONT_BOLD_NAME = GLOBAL_FONT_BOLD_NAME

    def generate(self, student, output_path, overrides=None):
        """
        Generate PDF attestat for a single student
        Uses LANDSCAPE A4 to match template
        """
        c = canvas.Canvas(output_path, pagesize=landscape(A4))
        
        # --- Page 1 ---
        # Draw Background if available
        if self.bg_path and os.path.exists(self.bg_path):
            # Draw image to fit A4 Landscape
            c.drawImage(self.bg_path, 0, 0, width=PAGE_WIDTH, height=PAGE_HEIGHT)
        
        c.setFont(self.FONT_NAME, 7)
        
        # Draw metadata (Page 1)
        # Apply overrides for metadata if any
        name_kz = student.get('name_kz', '')
        if overrides and 'metadata' in overrides and 'name_kz' in overrides['metadata']:
             name_kz = overrides['metadata']['name_kz']

        if name_kz:
             # User requested Bold and Larger for Name
             c.setFont(self.FONT_BOLD_NAME, 11)
             c.drawString(103.58, 517.18, name_kz)
        
        # Document Number (Bold 11 as requested)
        c.setFont(self.FONT_BOLD_NAME, 11)
        doc_num = student.get('document_number', '')
        c.drawString(129.26, 538.56, str(doc_num))
        
        # User requested same font as Name (Bold 11) for Years and Info
        c.setFont(self.FONT_BOLD_NAME, 11) 

        # Years (2023 - 2026)
        c.drawString(83.30, 502.02, "2023") 
        c.drawString(283.61, 502.50, "2026")
        
        # Institution and Qualification Info (Centered ~ X=200)
        center_x = 210
        
        # Defaults
        default_inst = "“Орал гуманитарлық-техникалық колледжі” мекемесінде"
        default_spec = "01140100 “Бастауыш білім беру педагогикасы мен әдістемесі” мамандығында"
        default_qual = "4S01140101 “Бастауыш білім беру мұғалімі”"
        default_qual2 = "біліктілігі бойынша"

        # Get from Student Data (First priority)
        txt_inst = student.get('institution', default_inst)
        txt_spec = student.get('specialty', default_spec)
        txt_qual = student.get('qualification', default_qual)
        txt_qual2 = student.get('qualification_2', default_qual2)
        
        # Check Overrides (Second priority - e.g. from a config file?? Not used by UI but good to keep)
        if overrides and 'metadata' in overrides:
             md = overrides['metadata']
             txt_inst = md.get('institution', txt_inst)
             txt_spec = md.get('specialty', txt_spec)
             txt_qual = md.get('qualification', txt_qual)
             txt_qual2 = md.get('qualification_2', txt_qual2)
        
        # 1. Institution
        c.drawCentredString(center_x, 480, txt_inst)
        
        # Change font to 9 for Specialty and Qualification per user request
        c.setFont(self.FONT_BOLD_NAME, 9)

        # 2. Specialty
        c.drawCentredString(center_x, 460, txt_spec) # +5
        
        # 3. Qualification
        c.drawCentredString(center_x, 440, txt_qual) # +5
        c.drawCentredString(center_x, 425, txt_qual2) # +5
        
        # Draw grades using simple mapping
        self._draw_grades(c, student, page_num=1, overrides=overrides)
        
        c.showPage()
        
        # --- Page 2 ---
        # Draw background for page 2
        bg2 = getattr(self, 'bg2_path', None)
        if not bg2 and self.bg_path:
             bg2 = self.bg_path.replace('.jpg', ' 2.jpg')
             if not os.path.exists(bg2):
                 bg2 = self.bg_path.replace('шаблон каз.jpg', 'шаблон каз2.jpg')

        if bg2 and os.path.exists(bg2):
            c.drawImage(bg2, 0, 0, width=PAGE_WIDTH, height=PAGE_HEIGHT)
        
        c.setFont(self.FONT_NAME, 7)
        self._draw_grades(c, student, page_num=2, overrides=overrides)
        
        c.showPage()
        c.save()
        print(f"PDF saved to {output_path}")

    def _draw_grades(self, c, student, page_num, overrides=None):
        """
        Draw all grades sequentially, flowing from column to column.
        Supports overrides for layout customization.
        overrides format: { 'subjects': { index: { 'y_offset': float, 'force_break': bool, 'custom_text': str, ... } } }
        """
        # Define areas for content flow based on template analysis
        # Format: (x_start, y_start, y_limit, page_num)
        flow_areas = [
            {"page": 1, "x": 24, "y_start": 280, "y_limit": 20, "width": 400},   # Page 1 Left (Reverted to +10)
            {"page": 1, "x": 448, "y_start": 563, "y_limit": 20, "width": 400},  # Page 1 Right (Kept at +15)
            {"page": 2, "x": 21, "y_start": 550, "y_limit": 20, "width": 400},   # Page 2 Left
            {"page": 2, "x": 448, "y_start": 550, "y_limit": 20, "width": 400}   # Page 2 Right
        ]
        
        # We need to process ALL subjects from start to finish to track layout state correctly
        # Layout state: which area index we are in, current Y coordinate, current row number
        current_area_idx = 0
        current_y = flow_areas[0]["y_start"]
        row_num = 1
        previous_module = None
        
        subjects = student.get('subjects_list', [])
        subj_overrides = overrides.get('subjects', {}) if overrides else {}
        
        for idx, subj in enumerate(subjects):
            # Check for Module Change
            current_module = subj.get('module', '').strip()
            
            # Application of Layout Overrides: Forced Break
            force_break = False
            y_offset = 0.0
            
            if idx in subj_overrides:
                 force_break = subj_overrides[idx].get('force_break', False)
                 y_offset = float(subj_overrides[idx].get('y_offset', 0.0))
            
            # Handling Force Break
            if force_break:
                 current_area_idx += 1
                 if current_area_idx < len(flow_areas):
                      current_y = flow_areas[current_area_idx]["y_start"]

            # If module changed (and it's not empty), draw header
            if current_module and current_module != previous_module:
                # Apply the Y-offset of the FIRST subject to the Header as well
                # This "glues" the header to the subject
                current_y += y_offset
                
                # Wrap header text (limit 65 chars for wider column)
                header_lines = self._wrap_text(current_module, max_chars=65)
                header_height = len(header_lines) * 9 + 2 # line height + padding
                
                # Calculate space needed for header + space for at least one subject (approx 15pt)
                # We do a simple lookahead for the NEXT subject height
                
                # Default safety margin: Header Height + 1 line subject (15)
                # Reduced buffer to allow tighter packing (User requested)
                safety_margin = header_height + 10 
                
                if current_y - safety_margin < flow_areas[current_area_idx]["y_limit"]:
                    # Move to next area
                    current_area_idx += 1
                    if current_area_idx >= len(flow_areas): break # Out of space
                    current_y = flow_areas[current_area_idx]["y_start"]
                    
                # Draw Header if on correct page
                if flow_areas[current_area_idx]["page"] == page_num:
                    c.setFont(self.FONT_BOLD_NAME, 7)
                    
                    text_y = current_y
                    for line in header_lines:
                        c.drawString(flow_areas[current_area_idx]["x"], text_y, line)
                        text_y -= 9
                
                current_y -= header_height # Space after header
                previous_module = current_module
                
                # Consume y_offset so it's not applied again to the subject
                y_offset = 0 

            # Prepare Subject Text (Wrapping)
            # Check for text override
            subj_name = subj.get('name_kz', '')
            if idx in subj_overrides and 'custom_text' in subj_overrides[idx]:
                 subj_name = subj_overrides[idx]['custom_text']
            
            c.setFont(self.FONT_NAME, 7) # Ensure font is set for width calc (approx)
            lines = self._wrap_text(subj_name, 28) # Limit 28 chars
            
            # Calculate total height for this item (lines + padding)
            # Add Y Offset to current Y (negative moves down, positive moves up, but here Y grows up so + is up)
            # WAIT: In PDF coords, Y=0 is bottom. So + is Up, - is Down.
            # But we are subtracting from current_y.
            # Let's say we want to "move down" -> layout creates a gap.
            # We apply Offset to the starting Y of this item.
            
            current_y += y_offset 

            item_height = len(lines) * 9 + 2 # 9pt per line + reduced padding (2pt)
            
            # Check if items fit in current area
            if current_y - item_height < flow_areas[current_area_idx]["y_limit"]:
                # Move to next area
                current_area_idx += 1
                if current_area_idx >= len(flow_areas): break # Out of space
                current_y = flow_areas[current_area_idx]["y_start"]
                # Re-apply offset for new column? Usually offsets are relative to previous item.
                # If we moved column, offset might mean "start lower in this column".
                current_y += y_offset

            # Draw Item if on correct page
            if flow_areas[current_area_idx]["page"] == page_num:
                x = flow_areas[current_area_idx]["x"]
                y = current_y
                
                # Row Number
                c.drawString(x, y, str(row_num))
                
                # Subject Lines
                text_y = y
                for line in lines:
                    c.drawString(x + 20, text_y, line)
                    text_y -= 9
                
                # Data (aligned with top line)
                # Adjust X coordinates based on template columns
                hours_val = str(subj.get('hours', ''))
                score_val = str(subj.get('score', ''))
                letter_val = str(subj.get('letter', ''))
                point_val = str(subj.get('point', ''))
                
                c.drawString(x + 160, y, hours_val)
                
                # Check for "сын" (Exam) trigger in Letter column
                if 'сын' in letter_val.lower():
                    # Move 'point' (the numeric grade) to the "Numeric 5-point" column (Last column)
                    # User requested text "сынақ" instead of number
                    # Shift -5 horizontal as requested
                    c.drawString(x + 345, y, "сынақ")
                else:
                    # Standard columns
                    c.drawString(x + 235, y, score_val)
                    c.drawString(x + 258, y, letter_val)
                    c.drawString(x + 295, y, point_val)
            
            # Update State
            current_y -= item_height
            row_num += 1

    def _wrap_text(self, text, max_chars=28):
        """
        Wrap text based on character count.
        Default max_chars = 28 (User requested limit).
        """
        if not text: return []
        
        # Support manual line breaks
        text = text.replace('|', '\n')
        paragraphs = text.split('\n')
        
        final_lines = []
        
        for paragraph in paragraphs:
            words = paragraph.split()
            if not words:
                continue
            
            current_line = []
            current_len = 0
            
            for word in words:
                # Calculate length with new word
                # If current_len > 0, we add 1 for space
                line_len_with_word = current_len + len(word) + (1 if current_len > 0 else 0)
                
                if line_len_with_word > max_chars and current_line:
                    # Line full
                    final_lines.append(" ".join(current_line))
                    current_line = [word]
                    current_len = len(word)
                else:
                    current_line.append(word)
                    current_len = line_len_with_word
            
            if current_line:
                final_lines.append(" ".join(current_line))
                
        return final_lines
