import re

files = [
    ('attestat_generator/templates/template_kz.htm', 'name_kz'),
    ('attestat_generator/templates/template_ru.htm', 'name_ru')
]

# improved regex to capture cell content within spans
# Row structure:
# 1. Index
# 2. Name
# 3. Hours (288) -> Clear
# 4. Empty
# 5. Score (87)
# 6. Letter
# 7. Point (3,33)
# 8. Empty

def apply_fix(path, name_key):
    print(f"Processing {path}...")
    with open(path, 'r', encoding='cp1251') as f:
        content = f.read()

    # 1. Find Table Start (Row 1)
    # Search for ">1</span>" and backtrack to "<tr>"
    start_anchor = ">1</span>"
    start_idx = content.find(start_anchor)
    if start_idx == -1:
        print("  Start anchor '>1</span>' not found.")
        return

    tr_start = content.rfind("<tr", 0, start_idx)
    if tr_start == -1:
        print("  Could not find start <tr>.")
        return

    # 2. Find Table End
    # Search for </table> after the start
    table_end = content.find("</table", tr_start)
    if table_end == -1:
        print("  Could not find end </table>.")
        return
    
    # 3. Extract the First Row to use as Template
    # Find first </tr>
    tr_end = content.find("</tr>", tr_start) + 5
    first_row = content[tr_start:tr_end]
    
    # 4. Create Jinja2 Row
    # We use simple string replacement on the first row to inject tags
    # We replace the *content* of the spans.
    # Regex to find >Content< inside spans
    
    # Helper to replace Nth occurrence of >Content<
    # But HTML is messy.
    # Let's split by "<td" and process each cell?
    cells = first_row.split("<td")
    # cell[0] is start of tr
    # cell[1] is Index td
    # cell[2] is Name td
    # ...
    
    if len(cells) < 8:
        print("  Row structure unexpected (columns < 8).")
        return

    # Cell 1: Index (>1<)
    cells[1] = re.sub(r'>\s*1\s*<', '>{{ loop.index }}<', cells[1])
    
    # Cell 2: Name
    # Replace anything between > and < inside the span?
    # Usually the content is deep inside: <p><span>TEXT</span></p>
    # Regex: replace content of the last span?
    # Look for the last ">" before "</span>"
    cells[2] = re.sub(r'>[^<>]+</span>', f'>{{{{ s.{name_key} }}}}</span>', cells[2])
    
    # Cell 3: Hours (288) -> Clear
    cells[3] = re.sub(r'>[^<>]+</span>', '>&nbsp;</span>', cells[3])
    
    # Cell 5: Score (87)
    cells[5] = re.sub(r'>[^<>]+</span>', '>{{ s.score }}</span>', cells[5], count=1)
    
    # Cell 6: Letter (Garbled/Unknown)
    # Just match whatever is in there
    cells[6] = re.sub(r'>[^<>]+</span>', '>{{ s.letter }}</span>', cells[6], count=1)
    
    # Cell 7: Point (3,33)
    cells[7] = re.sub(r'>[^<>]+</span>', '>{{ s.point }}</span>', cells[7], count=1)
    
    # Reassemble row
    jinja_row = "<td".join(cells)
    
    # 5. Create Loop Block
    loop_block = (
        "{% for s in subjects %}\n" +
        jinja_row + "\n" +
        "{% endfor %}\n"
    )
    
    # 6. Replace the Whole Table Body (from tr_start to table_end)
    # Note: table_end is index of </table>. We replace up to that.
    # BEWARE: This removes all rows from 1 to End.
    # We should ensure we don't delete headers (headers are before tr_start).
    # We should ensure we don't delete footer (footer might be after table_end? No, footer is usually inside table or separate).
    # Based on `inspect_table_range`, the chunk ended cleanly.
    
    new_content = content[:tr_start] + loop_block + content[table_end:]
    
    with open(path, 'w', encoding='cp1251') as f:
        f.write(new_content)
    print("  Updated successfully.")

if __name__ == "__main__":
    for p, k in files:
        apply_fix(p, k)
