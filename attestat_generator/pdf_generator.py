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

    def generate(self, student, output_path):
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
        
        c.setFont(self.FONT_NAME, 9)
        
        # Draw metadata (Page 1)
        # Coordinates below are placeholders or should match your landscape analysis
        if 'name_kz' in student:
             c.drawString(103.58, 517.18, student['name_kz'])
        
        doc_num = student.get('document_number', '')
        c.drawString(129.26, 538.56, str(doc_num))
        
        # Years (2023 - 2026)
        c.drawString(83.30, 497.02, "2023")
        c.drawString(283.61, 497.50, "2026")
        
        # Draw grades using simple mapping
        self._draw_grades(c, student, page_num=1)
        
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
        
        c.setFont(self.FONT_NAME, 9)
        self._draw_grades(c, student, page_num=2)
        
        c.showPage()
        c.save()
        print(f"PDF saved to {output_path}")

    def _draw_grades(self, c, student, page_num):
        """
        Draw all grades sequentially, flowing from column to column.
        """
        # Define areas for content flow based on template analysis
        # Format: (x_start, y_start, y_limit, page_num)
        flow_areas = [
            {"page": 1, "x": 24, "y_start": 270, "y_limit": 50, "width": 400},   # Page 1 Left
            {"page": 1, "x": 448, "y_start": 548, "y_limit": 50, "width": 400},  # Page 1 Right
            {"page": 2, "x": 21, "y_start": 550, "y_limit": 50, "width": 400},   # Page 2 Left
            {"page": 2, "x": 448, "y_start": 550, "y_limit": 50, "width": 400}   # Page 2 Right
        ]
        
        # We need to process ALL subjects from start to finish to track layout state correctly
        # Layout state: which area index we are in, current Y coordinate, current row number
        current_area_idx = 0
        current_y = flow_areas[0]["y_start"]
        row_num = 1
        previous_module = None
        
        subjects = student.get('subjects_list', [])
        
        for subj in subjects:
            # Check for Module Change
            current_module = subj.get('module', '').strip()
            
            # If module changed (and it's not empty), draw header
            if current_module and current_module != previous_module:
                # Calculate space needed for header (approx 14pt)
                if current_y - 14 < flow_areas[current_area_idx]["y_limit"]:
                    # Move to next area
                    current_area_idx += 1
                    if current_area_idx >= len(flow_areas): break # Out of space
                    current_y = flow_areas[current_area_idx]["y_start"]
                    
                # Draw Header if on correct page
                if flow_areas[current_area_idx]["page"] == page_num:
                    c.setFont(self.FONT_BOLD_NAME, 9)
                    c.drawString(flow_areas[current_area_idx]["x"], current_y, current_module)
                
                current_y -= 14 # Space after header
                previous_module = current_module

            # Prepare Subject Text (Wrapping)
            subj_name = subj.get('name_kz', '')
            c.setFont(self.FONT_NAME, 9) # Ensure font is set for width calc (approx)
            lines = self._wrap_text(subj_name, 45) # ~45 chars max width
            
            # Calculate total height for this item (lines + padding)
            item_height = len(lines) * 11 + 4 # 11pt per line + padding
            
            # Check if items fit in current area
            if current_y - item_height < flow_areas[current_area_idx]["y_limit"]:
                # Move to next area
                current_area_idx += 1
                if current_area_idx >= len(flow_areas): break # Out of space
                current_y = flow_areas[current_area_idx]["y_start"]
                
                # Re-draw module header on new page/column for context? 
                # User asked for sequential flow, usually you don't repeat header unless requested.
                # We will just continue content.

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
                    text_y -= 11
                
                # Data (aligned with top line)
                # Adjust X coordinates based on template columns
                c.drawString(x + 160, y, str(subj.get('hours', '')))
                c.drawString(x + 235, y, str(subj.get('score', '')))
                c.drawString(x + 258, y, str(subj.get('letter', '')))
                c.drawString(x + 295, y, str(subj.get('point', '')))
            
            # Update State
            current_y -= item_height
            row_num += 1

    def _wrap_text(self, text, max_chars):
        if not text: return []
        words = text.split()
        lines = []
        current_line = []
        current_len = 0
        for word in words:
            if current_len + len(word) + 1 > max_chars:
                lines.append(" ".join(current_line))
                current_line = [word]
                current_len = len(word)
            else:
                current_line.append(word)
                current_len += len(word) + 1
        if current_line:
            lines.append(" ".join(current_line))
        return lines
