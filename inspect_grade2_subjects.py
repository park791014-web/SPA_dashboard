import os
import zipfile
import xml.etree.ElementTree as ET
import re

folder_path = r"c:\Users\부산동성고등학교\Documents\AntiGravity\내신등급 산출\성적파일"

def inspect_file(filename):
    p = os.path.join(folder_path, filename)
    with zipfile.ZipFile(p, 'r') as z:
        shared_strings = []
        try:
            with z.open('xl/sharedStrings.xml') as f:
                root = ET.parse(f).getroot()
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for t in root.findall('.//ns:t', ns):
                    shared_strings.append(t.text if t.text else "")
        except KeyError:
            pass

        with z.open('xl/worksheets/sheet1.xml') as f:
            root = ET.parse(f).getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            grid = {}
            for row in root.findall('.//ns:row', ns):
                r_idx = int(row.attrib.get('r'))
                for cell in row.findall('ns:c', ns):
                    c_ref = cell.attrib.get('r')
                    col_letter = re.match(r'([A-Z]+)', c_ref).group(1)
                    col_idx = 0
                    for char in col_letter:
                        col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
                    col_idx -= 1
                    
                    cell_type = cell.attrib.get('t')
                    val_node = cell.find('ns:v', ns)
                    val = val_node.text if val_node is not None else ""
                    
                    if cell_type == 's' and val:
                        idx = int(val)
                        if idx < len(shared_strings):
                            val = shared_strings[idx]
                    
                    if r_idx not in grid:
                        grid[r_idx] = {}
                    grid[r_idx][col_idx] = val
            
            # Find "단위수" cell
            unit_row_idx = -1
            unit_col_idx = -1
            for r in sorted(grid.keys()):
                for c in grid[r]:
                    if str(grid[r][c]).strip() == "단위수":
                        unit_row_idx = r
                        unit_col_idx = c
                        break
                if unit_row_idx != -1:
                    break
            
            print(f"\nFile: {filename}")
            if unit_row_idx == -1:
                print("Could not find '단위수' in this file!")
                return
                
            subject_row_idx = unit_row_idx + 1
            unit_row = grid[unit_row_idx]
            subject_row = grid.get(subject_row_idx, {})
            
            print(f"Row containing '단위수': {unit_row_idx}")
            print(f"Row containing subjects: {subject_row_idx}")
            
            # Print subjects and unit values
            subjects_found = []
            for c in sorted(subject_row.keys()):
                if c > unit_col_idx:
                    sub_name = str(subject_row[c]).strip()
                    raw_unit = unit_row.get(c, "")
                    subjects_found.append(f"{chr(65+c)}:{sub_name}(단위수:{raw_unit})")
            print("Detected Subjects:", ", ".join(subjects_found[:12]))
            
            # Let's inspect some student rows to see what values are in columns for Japanese/Chinese
            # Let's print row 15
            if 15 in grid:
                row_15 = grid[15]
                vals = [f"{chr(65+c)}:{row_15[c]}" for c in sorted(row_15.keys()) if c > unit_col_idx]
                print("Row 15 Grades:", ", ".join(vals[:12]))

inspect_file("2026_1학기_학기말_201.xlsx")
inspect_file("2026_1학기_학기말_205.xlsx")
inspect_file("2026_1학기_학기말_206.xlsx")
