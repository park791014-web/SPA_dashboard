import os
import zipfile
import xml.etree.ElementTree as ET
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"c:\Users\부산동성고등학교\Documents\AntiGravity\내신등급 산출\성적파일\2026_1학기_학기말_201.xlsx"

with zipfile.ZipFile(file_path, 'r') as z:
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

        print("--- Rows in 201.xlsx ---")
        for r in sorted(grid.keys())[:15]:
            row = grid[r]
            line = [f"{chr(65+c)}: '{row[c]}'" for c in sorted(row.keys())]
            print(f"Row {r}: " + " | ".join(line))
