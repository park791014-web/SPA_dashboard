import os
import zipfile
import xml.etree.ElementTree as ET
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')
p = r"c:\Users\부산동성고등학교\Documents\AntiGravity\내신등급 산출\성적파일\교과학습발달상황(2-2).xlsx"

with zipfile.ZipFile(p, 'r') as z:
    shared_strings = []
    try:
        with z.open('xl/sharedStrings.xml') as f:
            root = ET.parse(f).getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for t in root.findall('.//ns:t', ns):
                shared_strings.append(t.text if t.text else "")
    except KeyError: pass

    sheet_files = [x for x in z.namelist() if "worksheets/sheet" in x]
    sheet_name = sheet_files[0]
    with z.open(sheet_name) as f: root = ET.parse(f).getroot()
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    grid = {}
    for row in root.findall('.//ns:row', ns):
        r_idx = int(row.attrib.get('r'))
        for cell in row.findall('ns:c', ns):
            c_ref = cell.attrib.get('r')
            col_letter = re.match(r'([A-Z]+)', c_ref).group(1)
            col_idx = 0
            for char in col_letter: col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
            col_idx -= 1
            cell_type = cell.attrib.get('t')
            val_node = cell.find('ns:v', ns)
            val = val_node.text if val_node is not None else ""
            if cell_type == 's' and val:
                idx = int(val)
                if idx < len(shared_strings): val = shared_strings[idx]
            if r_idx not in grid: grid[r_idx] = {}
            grid[r_idx][col_idx] = val
    rows = []
    max_r = max(grid.keys()) if grid else 0
    for r in range(1, max_r + 1):
        if r in grid:
            row_data = []
            max_c = max(grid[r].keys()) if grid[r] else 0
            for c in range(max_c + 1): row_data.append(grid[r].get(c, ""))
            rows.append(row_data)
        else: rows.append([])

    # Header Row is Row 4 (index 3)
    headers = [str(x).strip().replace(" ", "") for x in rows[3]]
    print("Headers:", headers)
    
    def findColumnIndex(headers, keys):
        for i in range(len(headers)):
            h = headers[i]
            for key in keys:
                if key in h:
                    return i
        return -1
        
    idxRoll = findColumnIndex(headers, ["번호", "번 호"])
    idxName = findColumnIndex(headers, ["성명", "이름"])
    idxGrade = findColumnIndex(headers, ["학년"])
    idxSemester = findColumnIndex(headers, ["학기"])
    idxSubject = findColumnIndex(headers, ["과목"])
    idxUnits = findColumnIndex(headers, ["학점", "단위수", "단위"])
    idxAchievement = findColumnIndex(headers, ["성취도"])
    idxRank = findColumnIndex(headers, ["석차등급", "등급"])
    
    print("idxRoll:", idxRoll)
    print("idxName:", idxName)
    print("idxGrade:", idxGrade)
    print("idxSemester:", idxSemester)
    print("idxSubject:", idxSubject)
    print("idxUnits:", idxUnits)
    print("idxAchievement:", idxAchievement)
    print("idxRank:", idxRank)
    
    # Let's print rows[4] values mapped
    print("\nRow 5 mapping:")
    row = rows[4]
    for idx, val in enumerate(row):
        print(f"Col {idx} ({rows[3][idx]}): {val}")
