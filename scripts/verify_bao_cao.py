# -*- coding: utf-8 -*-
"""
Kịch bản kiểm tra toàn diện chất lượng file Word Bao_Cao_Khoa_Luan.docx
"""
import os
import sys
import docx
from docx.oxml import parse_xml

sys.stdout.reconfigure(encoding="utf-8")

DOCX_PATH = r"d:\Aziel\KLTN_2026\Bao_Cao_Khoa_Luan.docx"

doc = docx.Document(DOCX_PATH)

print("=" * 60)
print(f"File: {DOCX_PATH}")
print(f"File size: {os.path.getsize(DOCX_PATH):,} bytes")
print(f"Total Sections: {len(doc.sections)}")
print(f"Total Paragraphs: {len(doc.paragraphs)}")
print(f"Total Tables: {len(doc.tables)}")

# 1. Kiểm tra Font chữ Times New Roman
non_tnr_runs = []
for pi, p in enumerate(doc.paragraphs):
    for ri, r in enumerate(p.runs):
        font_name = r.font.name
        # Bỏ qua symbol font như Wingdings ở trang bìa
        if font_name and font_name not in ["Times New Roman", "Wingdings"]:
            non_tnr_runs.append((pi, ri, font_name, r.text[:30]))

print(f"\nNon-Times New Roman runs count: {len(non_tnr_runs)}")
if non_tnr_runs:
    for item in non_tnr_runs[:5]:
        print("  Non-TNR:", item)
else:
    print("  -> 100% Times New Roman verified across all standard runs!")

# 2. Kiểm tra Table Styling (Nền trắng, viền đen)
non_white_cells = 0
total_cells = 0
for ti, tbl in enumerate(doc.tables):
    for row in tbl.rows:
        for cell in row.cells:
            total_cells += 1
            tcPr = cell._tc.tcPr
            if tcPr is not None:
                for child in tcPr:
                    if child.tag.endswith("shd"):
                        fill = child.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
                        if fill and fill.upper() not in ["FFFFFF", "AUTO", "NONE"]:
                            non_white_cells += 1

print(f"\nTotal Table Cells: {total_cells}")
print(f"Non-white / Shaded cells: {non_white_cells}")
if non_white_cells == 0:
    print("  -> 100% Pure White cells verified across all tables (No gray/zebra shading)!")

# 3. Kiểm tra Section Page Numbering
print("\nSection Page Numbering:")
for si, s in enumerate(doc.sections):
    sectPr = s._sectPr
    pgNumType = None
    for child in sectPr:
        if child.tag.endswith("pgNumType"):
            fmt = child.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fmt")
            start = child.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}start")
            pgNumType = f"fmt={fmt}, start={start}"
    has_footer = bool(s.footer.paragraphs[0].text or len(s.footer.paragraphs[0]._p))
    print(f"  Section {si+1}: {pgNumType or 'Default'}, Footer present: {has_footer}")

# 4. Kiểm tra sự hiện diện của OMML equations
omml_count = 0
for p in doc.paragraphs:
    if "oMath" in p._p.xml:
        omml_count += 1

print(f"\nTotal OMML Equations found: {omml_count}")

# 5. Kiểm tra các Heading
headings = []
for p in doc.paragraphs:
    pPr = p._p.pPr
    if pPr is not None:
        for child in pPr:
            if child.tag.endswith("outlineLvl"):
                lvl = child.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                headings.append((lvl, p.text[:60]))

print(f"\nTotal Outlined Headings: {len(headings)}")
for lvl, text in headings[:15]:
    print(f"  [Lvl {lvl}] {text}")
print("  ...")
for lvl, text in headings[-10:]:
    print(f"  [Lvl {lvl}] {text}")

print("=" * 60)
