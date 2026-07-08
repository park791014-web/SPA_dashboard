import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
p = r"c:\Users\부산동성고등학교\Documents\AntiGravity\내신등급 산출\app.js"
with open(p, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for idx, line in enumerate(lines):
    if ".length" in line:
        print(f"Line {idx+1}: {line.strip()}")
