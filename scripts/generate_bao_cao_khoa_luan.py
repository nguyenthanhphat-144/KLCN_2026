# -*- coding: utf-8 -*-
"""
Kịch bản sinh toàn bộ Báo cáo Khóa luận Tốt nghiệp Đại học: Bao_Cao_Khoa_Luan.docx
Tuân thủ 100% Thể thức HUIT 2025 và Yêu cầu Định dạng từ Ảnh 1, Ảnh 2, Ảnh 3.
Phần 1: Thiết lập thư viện, hàm định dạng, và helper tạo công thức toán OMML Word chuẩn.
"""

import os
import sys
import json
import subprocess
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = r"d:\Aziel\KLTN_2026"
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
IMAGES_DIR = os.path.join(BASE_DIR, "images")
BENCH_JSON = os.path.join(BASE_DIR, "docs", "Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json")
OUTPUT_DOCX = os.path.join(BASE_DIR, "Bao_Cao_Khoa_Luan.docx")

# Đọc số liệu thực nghiệm gốc ("Sự thật việc thật")
with open(BENCH_JSON, "r", encoding="utf-8") as f:
    bench_data = json.load(f)

pho_m = bench_data["phobert_metrics"]
svm_m = bench_data["svm_metrics"]
sel_m = bench_data["selected_model"]
man_m = bench_data["manual_vs_automated"]
ds_m = bench_data["dataset_stats"]
per_f = bench_data["per_field_comparison"]

chapter_fig_count = {}
chapter_tbl_count = {}

# ==================== ĐỊNH DẠNG CƠ BẢN ====================

def set_page_setup(section):
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3.5)
    section.right_margin = Cm(2.5)
    section.header_distance = Cm(1.27)
    section.footer_distance = Cm(1.27)

def set_section_page_numbering(section, fmt="decimal", start=1):
    sectPr = section._sectPr
    for child in list(sectPr):
        if child.tag.endswith("pgNumType"):
            sectPr.remove(child)
    if fmt == "lowerRoman":
        xml = f'<w:pgNumType {nsdecls("w")} w:fmt="lowerRoman" w:start="{start}"/>'
    else:
        xml = f'<w:pgNumType {nsdecls("w")} w:fmt="decimal" w:start="{start}"/>'
    sectPr.append(parse_xml(xml))

def add_centered_page_number_to_footer(section):
    section.footer.is_linked_to_previous = False
    p = section.footer.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    fld = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
    p._p.append(fld)

def configure_doc_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(13)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 1.3
    normal.paragraph_format.space_before = Pt(6)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def add_chapter_title(doc, title_text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(18)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.page_break_before = True
    
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
    
    r = p.add_run(title_text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_h1(doc, title_text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.first_line_indent = Cm(0)
    
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="1"/>'))
    
    r = p.add_run(title_text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_h2(doc, title_text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.first_line_indent = Cm(0)
    
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="2"/>'))
    
    r = p.add_run(title_text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_h3(doc, title_text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.first_line_indent = Cm(0)
    
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="3"/>'))
    
    r = p.add_run(title_text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_body(doc, text, bold_prefix=None, italic_prefix=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.first_line_indent = Cm(1.0)
    
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Times New Roman"
        r_pre.font.size = Pt(13)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(0, 0, 0)
        
    if italic_prefix:
        r_it = p.add_run(italic_prefix)
        r_it.font.name = "Times New Roman"
        r_it.font.size = Pt(13)
        r_it.font.italic = True
        r_it.font.color.rgb = RGBColor(0, 0, 0)

    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_bullet_1(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.5)

    r_dash = p.add_run("–\t")
    r_dash.font.name = "Times New Roman"
    r_dash.font.size = Pt(13)
    r_dash.font.color.rgb = RGBColor(0, 0, 0)

    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = "Times New Roman"
        r_b.font.size = Pt(13)
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(0, 0, 0)

    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_bullet_2(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(1.3)
    p.paragraph_format.first_line_indent = Cm(-0.5)

    r_plus = p.add_run("+\t")
    r_plus.font.name = "Times New Roman"
    r_plus.font.size = Pt(13)
    r_plus.font.color.rgb = RGBColor(0, 0, 0)

    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = "Times New Roman"
        r_b.font.size = Pt(13)
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(0, 0, 0)

    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_omml_equation(doc, math_inner_xml):
    xml = f"""<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <w:pPr>
        <w:spacing w:before="120" w:after="120" w:line="312" w:lineRule="auto"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <m:oMathPara>
        <m:oMath>
          {math_inner_xml}
        </m:oMath>
      </m:oMathPara>
    </w:p>"""
    doc._body._element.append(parse_xml(xml))

def add_figure_custom(doc, img_rel_path, chapter_idx, caption_text, width_cm=14.5):
    full_path = os.path.join(BASE_DIR, img_rel_path) if not os.path.isabs(img_rel_path) else img_rel_path
    if not os.path.exists(full_path):
        print(f"WARNING: Image not found: {full_path}")
        return

    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(12)
    p_img.paragraph_format.space_after = Pt(4)
    p_img.paragraph_format.left_indent = Cm(0)
    p_img.paragraph_format.first_line_indent = Cm(0)
    run_img = p_img.add_run()
    run_img.add_picture(full_path, width=Cm(width_cm))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(4)
    p_cap.paragraph_format.space_after = Pt(12)
    p_cap.paragraph_format.line_spacing = 1.3
    p_cap.paragraph_format.left_indent = Cm(0)
    p_cap.paragraph_format.first_line_indent = Cm(0)

    is_first = chapter_idx not in chapter_fig_count
    chapter_fig_count[chapter_idx] = chapter_fig_count.get(chapter_idx, 0) + 1

    r1 = p_cap.add_run(f"Hình {chapter_idx}.")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0, 0, 0)

    if is_first:
        fld_xml = r'<w:fldSimple %s w:instr="SEQ Hình \r 1 \* ARABIC"/>' % nsdecls('w')
    else:
        fld_xml = r'<w:fldSimple %s w:instr="SEQ Hình \* ARABIC"/>' % nsdecls('w')
    p_cap._p.append(parse_xml(fld_xml))

    r2 = p_cap.add_run(f": {caption_text}")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(12)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0, 0, 0)

def add_table_custom(doc, chapter_idx, caption_text, headers, data, col_widths=None, has_caption=True):
    if has_caption:
        is_first = chapter_idx not in chapter_tbl_count
        chapter_tbl_count[chapter_idx] = chapter_tbl_count.get(chapter_idx, 0) + 1

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(12)
        p_cap.paragraph_format.space_after = Pt(4)
        p_cap.paragraph_format.line_spacing = 1.3
        p_cap.paragraph_format.left_indent = Cm(0)
        p_cap.paragraph_format.first_line_indent = Cm(0)

        r1 = p_cap.add_run(f"Bảng {chapter_idx}.")
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(0, 0, 0)

        if is_first:
            fld_xml = r'<w:fldSimple %s w:instr="SEQ Bảng \r 1 \* ARABIC"/>' % nsdecls('w')
        else:
            fld_xml = r'<w:fldSimple %s w:instr="SEQ Bảng \* ARABIC"/>' % nsdecls('w')
        p_cap._p.append(parse_xml(fld_xml))

        r2 = p_cap.add_run(f": {caption_text}")
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(12)
        r2.font.bold = True
        r2.font.color.rgb = RGBColor(0, 0, 0)

    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
    )
    table._tbl.tblPr.append(tblBorders)

    hdr_cells = table.rows[0].cells
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    for ci, h in enumerate(headers):
        hdr_cells[ci].text = h
        p = hdr_cells[ci].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.3
        p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.first_line_indent = Cm(0)
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF" w:val="clear"/>')
        hdr_cells[ci]._tc.get_or_add_tcPr().append(shd)

    for ri, row_data in enumerate(data):
        row = table.rows[ri + 1]
        r_trPr = row._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for ci, val in enumerate(row_data):
            row.cells[ci].text = str(val)
            p = row.cells[ci].paragraphs[0]
            val_str = str(val).strip()
            if ci == 0 or (len(val_str) <= 12 and not any(k in headers[ci].lower() for k in ["tên", "nội dung", "mô tả", "phạm vi", "giải pháp"])):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.3
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(12)
                r.font.color.rgb = RGBColor(0, 0, 0)
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF" w:val="clear"/>')
            row.cells[ci]._tc.get_or_add_tcPr().append(shd)

    if col_widths:
        for row in table.rows:
            for ci, w in enumerate(col_widths):
                if ci < len(row.cells):
                    row.cells[ci].width = Cm(w)

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(6)
    p_after.paragraph_format.space_after = Pt(0)
    p_after.paragraph_format.line_spacing = 1.3

# ==================== CÁC CÔNG THỨC TOÁN HỌC WORD OMML CHUẨN ====================

def add_omml_tfidf(doc):
    add_omml_equation(doc, """
    <m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>TF-IDF</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>t, d, D</m:t></m:r></m:e></m:d>
    <m:r><m:t> = </m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>1 + ln</m:t></m:r><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>tf</m:t></m:r><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>t, d</m:t></m:r></m:e></m:d></m:e></m:d></m:e></m:d>
    <m:r><m:t> × </m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>
      <m:r><m:t>ln</m:t></m:r>
      <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
        <m:f>
          <m:num><m:r><m:t>1 + |D|</m:t></m:r></m:num>
          <m:den><m:r><m:t>1 + df(t, D)</m:t></m:r></m:den>
        </m:f>
      </m:e></m:d>
      <m:r><m:t> + 1</m:t></m:r>
    </m:e></m:d>
    """)

def add_omml_linearsvc(doc):
    add_omml_equation(doc, """
    <m:r><m:t>min </m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>
      <m:f>
        <m:num><m:r><m:t>1</m:t></m:r></m:num>
        <m:den><m:r><m:t>2</m:t></m:r></m:den>
      </m:f>
      <m:r><m:t> ‖w‖² + C </m:t></m:r>
      <m:nary>
        <m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/><m:subHide m:val="0"/><m:supHide m:val="0"/></m:naryPr>
        <m:sub><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>i=1</m:t></m:r></m:sub>
        <m:sup><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>N</m:t></m:r></m:sup>
        <m:e>
          <m:r><m:t>max</m:t></m:r>
          <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
            <m:r><m:t>0, 1 - </m:t></m:r>
            <m:sSub><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>y</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>i</m:t></m:r></m:sub></m:sSub>
            <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
              <m:r><m:t>w^T </m:t></m:r>
              <m:sSub><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>i</m:t></m:r></m:sub></m:sSub>
              <m:r><m:t> + b</m:t></m:r>
            </m:e></m:d>
          </m:e></m:d>
        </m:e>
      </m:nary>
    </m:e></m:d>
    """)

def add_omml_platt_scaling(doc):
    add_omml_equation(doc, """
    <m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>P</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
      <m:r><m:t>y = 1 | f(x)</m:t></m:r>
    </m:e></m:d>
    <m:r><m:t> = </m:t></m:r>
    <m:f>
      <m:num><m:r><m:t>1</m:t></m:r></m:num>
      <m:den><m:r><m:t>1 + exp(A × f(x) + B)</m:t></m:r></m:den>
    </m:f>
    """)

def add_omml_scaled_attention(doc):
    add_omml_equation(doc, """
    <m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>Attention</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>Q, K, V</m:t></m:r></m:e></m:d>
    <m:r><m:t> = softmax</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
      <m:f>
        <m:num><m:r><m:t>Q K^T</m:t></m:r></m:num>
        <m:den><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:e><m:sSub><m:e><m:r><m:t>d</m:t></m:r></m:e><m:sub><m:r><m:t>k</m:t></m:r></m:sub></m:sSub></m:e></m:rad></m:den>
      </m:f>
    </m:e></m:d>
    <m:r><m:t> V</m:t></m:r>
    """)

def add_omml_multihead_attention(doc):
    add_omml_equation(doc, """
    <m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>MultiHead</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>Q, K, V</m:t></m:r></m:e></m:d>
    <m:r><m:t> = Concat</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
      <m:sSub><m:e><m:r><m:t>head</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>
      <m:r><m:t>, ..., </m:t></m:r>
      <m:sSub><m:e><m:r><m:t>head</m:t></m:r></m:e><m:sub><m:r><m:t>12</m:t></m:r></m:sub></m:sSub>
    </m:e></m:d>
    <m:sSup><m:e><m:r><m:t> W</m:t></m:r></m:e><m:sup><m:r><m:t>O</m:t></m:r></m:sup></m:sSup>
    """)

def add_omml_joint_loss(doc):
    add_omml_equation(doc, """
    <m:sSub><m:e><m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>L</m:t></m:r></m:e><m:sub><m:r><m:t>total</m:t></m:r></m:sub></m:sSub>
    <m:r><m:t> = </m:t></m:r>
    <m:sSub><m:e><m:r><m:t>L</m:t></m:r></m:e><m:sub><m:r><m:t>CE</m:t></m:r></m:sub></m:sSub>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
      <m:sSub><m:e><m:r><m:t>z</m:t></m:r></m:e><m:sub><m:r><m:t>field</m:t></m:r></m:sub></m:sSub>
      <m:r><m:t>, </m:t></m:r>
      <m:sSub><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sub><m:r><m:t>field</m:t></m:r></m:sub></m:sSub>
    </m:e></m:d>
    <m:r><m:t> + 0.5 × </m:t></m:r>
    <m:sSub><m:e><m:r><m:t>L</m:t></m:r></m:e><m:sub><m:r><m:t>CE</m:t></m:r></m:sub></m:sSub>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
      <m:sSub><m:e><m:r><m:t>z</m:t></m:r></m:e><m:sub><m:r><m:t>type</m:t></m:r></m:sub></m:sSub>
      <m:r><m:t>, </m:t></m:r>
      <m:sSub><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sub><m:r><m:t>type</m:t></m:r></m:sub></m:sSub>
    </m:e></m:d>
    """)

def add_omml_degree_centrality(doc):
    add_omml_equation(doc, """
    <m:sSub><m:e><m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>C</m:t></m:r></m:e><m:sub><m:r><m:t>D</m:t></m:r></m:sub></m:sSub>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>v</m:t></m:r></m:e></m:d>
    <m:r><m:t> = </m:t></m:r>
    <m:f>
      <m:num><m:r><m:t>deg(v)</m:t></m:r></m:num>
      <m:den><m:r><m:t>|V| - 1</m:t></m:r></m:den>
    </m:f>
    """)

def add_omml_isolation_forest(doc):
    add_omml_equation(doc, """
    <m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>s</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>x, n</m:t></m:r></m:e></m:d>
    <m:r><m:t> = </m:t></m:r>
    <m:sSup>
      <m:e><m:r><m:t>2</m:t></m:r></m:e>
      <m:sup>
        <m:r><m:t>- </m:t></m:r>
        <m:f>
          <m:num><m:r><m:t>E(h(x))</m:t></m:r></m:num>
          <m:den><m:r><m:t>c(n)</m:t></m:r></m:den>
        </m:f>
      </m:sup>
    </m:sSup>
    <m:r><m:t>   với   c(n) = 2 </m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e><m:r><m:t>ln(n - 1) + 0.5772156649</m:t></m:r></m:e></m:d>
    <m:r><m:t> - </m:t></m:r>
    <m:f>
      <m:num><m:r><m:t>2(n - 1)</m:t></m:r></m:num>
      <m:den><m:r><m:t>n</m:t></m:r></m:den>
    </m:f>
    """)

def add_omml_pytorch_autoencoder(doc):
    add_omml_equation(doc, """
    <m:sSub><m:e><m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>MSE</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>i</m:t></m:r></m:sub></m:sSub>
    <m:r><m:t> = </m:t></m:r>
    <m:f>
      <m:num><m:r><m:t>1</m:t></m:r></m:num>
      <m:den><m:r><m:t>6</m:t></m:r></m:den>
    </m:f>
    <m:nary>
      <m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/><m:subHide m:val="0"/><m:supHide m:val="0"/></m:naryPr>
      <m:sub><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>k=1</m:t></m:r></m:sub>
      <m:sup><m:r><m:t>6</m:t></m:r></m:sup>
      <m:e>
        <m:sSup>
          <m:e>
            <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
              <m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>ik</m:t></m:r></m:sub></m:sSub>
              <m:r><m:t> - </m:t></m:r>
              <m:acc>
                <m:accPr><m:chr m:val="̂"/></m:accPr>
                <m:e><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>ik</m:t></m:r></m:sub></m:sSub></m:e>
              </m:acc>
            </m:e></m:d>
          </m:e>
          <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
      </m:e>
    </m:nary>
    """)

def add_omml_cosine_similarity(doc):
    add_omml_equation(doc, """
    <m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>Sim</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
      <m:sSub><m:e><m:r><m:t>d</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
      <m:r><m:t>, </m:t></m:r>
      <m:sSub><m:e><m:r><m:t>d</m:t></m:r></m:e><m:sub><m:r><m:t>j</m:t></m:r></m:sub></m:sSub>
    </m:e></m:d>
    <m:r><m:t> = </m:t></m:r>
    <m:f>
      <m:num>
        <m:sSub><m:e><m:r><m:t>v</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t> · </m:t></m:r>
        <m:sSub><m:e><m:r><m:t>v</m:t></m:r></m:e><m:sub><m:r><m:t>j</m:t></m:r></m:sub></m:sSub>
      </m:num>
      <m:den>
        <m:r><m:t>‖</m:t></m:r>
        <m:sSub><m:e><m:r><m:t>v</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>‖₂ × ‖</m:t></m:r>
        <m:sSub><m:e><m:r><m:t>v</m:t></m:r></m:e><m:sub><m:r><m:t>j</m:t></m:r></m:sub></m:sSub>
        <m:r><m:t>‖₂</m:t></m:r>
      </m:den>
    </m:f>
    """)

def add_omml_risk_score(doc):
    add_omml_equation(doc, """
    <m:sSub><m:e><m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>R</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>i</m:t></m:r></m:sub></m:sSub>
    <m:r><m:t> = min</m:t></m:r>
    <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
      <m:r><m:t>0.99,  0.35 × Norm</m:t></m:r>
      <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
        <m:sSub><m:e><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>s</m:t></m:r></m:e><m:sub><m:r><m:t>IF, i</m:t></m:r></m:sub></m:sSub>
      </m:e></m:d>
      <m:r><m:t> + 0.35 × Norm</m:t></m:r>
      <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>
        <m:sSub><m:e><m:r><m:t>MSE</m:t></m:r></m:e><m:sub><m:r><m:t>AE, i</m:t></m:r></m:sub></m:sSub>
      </m:e></m:d>
      <m:r><m:t> + 0.30 × </m:t></m:r>
      <m:sSub><m:e><m:r><m:t>RuleSeverity</m:t></m:r></m:e><m:sub><m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>i</m:t></m:r></m:sub></m:sSub>
    </m:e></m:d>
    """)

def add_omml_classification_metrics(doc):
    add_omml_equation(doc, """
    <m:r><m:t>Precision = </m:t></m:r>
    <m:f><m:num><m:r><m:t>TP</m:t></m:r></m:num><m:den><m:r><m:t>TP + FP</m:t></m:r></m:den></m:f>
    <m:r><m:t>,   Recall = </m:t></m:r>
    <m:f><m:num><m:r><m:t>TP</m:t></m:r></m:num><m:den><m:r><m:t>TP + FN</m:t></m:r></m:den></m:f>
    <m:r><m:t>,   F1-score = 2 × </m:t></m:r>
    <m:f><m:num><m:r><m:t>Precision × Recall</m:t></m:r></m:num><m:den><m:r><m:t>Precision + Recall</m:t></m:r></m:den></m:f>
    """)

def add_cover_pages(doc):
    # ==================== SECTION 1: TRANG BÌA (KHÔNG ĐÁNH SỐ TRANG) ====================
    s1 = doc.sections[0]
    set_page_setup(s1)
    s1.footer.is_linked_to_previous = False
    s1.footer.paragraphs[0].text = ""

    # Thiết lập khung viền hoa văn góc tam giác (cornerTriangles sz="15") CHỈ HIỂN THỊ Ở TRANG BÌA ĐẦU TIÊN
    sectPr = s1._sectPr
    for child in list(sectPr):
        if child.tag.endswith("pgBorders"):
            sectPr.remove(child)
    borders_xml = (
        f'<w:pgBorders {nsdecls("w")} w:display="firstPage">'
        '<w:top w:val="cornerTriangles" w:sz="15" w:space="1" w:color="auto"/>'
        '<w:left w:val="cornerTriangles" w:sz="15" w:space="4" w:color="auto"/>'
        '<w:bottom w:val="cornerTriangles" w:sz="15" w:space="1" w:color="auto"/>'
        '<w:right w:val="cornerTriangles" w:sz="15" w:space="4" w:color="auto"/>'
        '</w:pgBorders>'
    )
    sectPr.append(parse_xml(borders_xml))

    # ---- TRANG BÌA CHÍNH (In trên giấy bìa màu) ----
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BỘ CÔNG THƯƠNG\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = False
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("TRƯỜNG ĐẠI HỌC CÔNG THƯƠNG TP. HCM\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("KHOA CÔNG NGHỆ THÔNG TIN")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    # Dòng biểu tượng Wingdings 🕮
    p_sym = doc.add_paragraph()
    p_sym.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sym.paragraph_format.space_before = Pt(2)
    p_sym.paragraph_format.space_after = Pt(6)
    sym_xml = (
        f'<w:pPr {nsdecls("w")}><w:jc w:val="center"/></w:pPr>'
        f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="32"/></w:rPr><w:t>---</w:t></w:r>'
        f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/><w:sz w:val="32"/></w:rPr><w:sym w:font="Wingdings" w:char="F098"/><w:sym w:font="Wingdings" w:char="F026"/><w:sym w:font="Wingdings" w:char="F099"/></w:r>'
        f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="32"/></w:rPr><w:t>---</w:t></w:r>'
    )
    p_sym._p.append(parse_xml(f'<w:span {nsdecls("w")}>{sym_xml}</w:span>'))

    # Logo HUIT tròn màu chính thức
    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(2)
    p_logo.paragraph_format.space_after = Pt(12)
    logo_path = os.path.join(IMAGES_DIR, "logo_huit.png")
    if os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Cm(3.5))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("KHÓA LUẬN TỐT NGHIỆP CỬ NHÂN CNTT")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC CÔNG NGHỆ\nBẰNG PHƯƠNG PHÁP PHÂN LOẠI TỰ ĐỘNG\nVÀ PHÁT HIỆN BẤT THƯỜNG")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(16)
    r = p.add_run("(Mã đề tài: CNTT-KLCN001 | Định hướng: Tích hợp Trí tuệ Nhân tạo)")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(24)
    r = p.add_run("Chuyên ngành: CÔNG NGHỆ THÔNG TIN")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.left_indent = Cm(4.5)
    p_info.paragraph_format.space_before = Pt(6)
    p_info.paragraph_format.space_after = Pt(4)
    p_info.paragraph_format.line_spacing = 1.3
    r = p_info.add_run("SINH VIÊN THỰC HIỆN:\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    students = [
        ("1. Nguyễn Thành Phát", "MSSV: 201230641", "Lớp: 14DHTH11"),
        ("2. Lê Thanh Phong", "MSSV: 2001230669", "Lớp: 14DHTH10"),
        ("3. Nguyễn Thành Tài", "MSSV: 2001230767", "Lớp: 14DHTH13"),
    ]
    for name, mssv, lop in students:
        r = p_info.add_run(f"   {name} – {mssv} – {lop}\n")
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(0, 0, 0)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.paragraph_format.space_before = Pt(36)
    p_date.paragraph_format.space_after = Pt(10)
    r = p_date.add_run("TP. HỒ CHÍ MINH, THÁNG 10 NĂM 2026")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    doc.add_page_break()

    # ---- TRANG BÌA PHỤ (In giấy trắng thường, KHÔNG CÓ KHUNG VIỀN) ----
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BỘ CÔNG THƯƠNG\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = False
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("TRƯỜNG ĐẠI HỌC CÔNG THƯƠNG TP. HCM\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("KHOA CÔNG NGHỆ THÔNG TIN")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_sym2 = doc.add_paragraph()
    p_sym2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sym2.paragraph_format.space_before = Pt(2)
    p_sym2.paragraph_format.space_after = Pt(6)
    p_sym2._p.append(parse_xml(f'<w:span {nsdecls("w")}>{sym_xml}</w:span>'))

    p_logo2 = doc.add_paragraph()
    p_logo2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo2.paragraph_format.space_before = Pt(2)
    p_logo2.paragraph_format.space_after = Pt(12)
    if os.path.exists(logo_path):
        p_logo2.add_run().add_picture(logo_path, width=Cm(3.5))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("KHÓA LUẬN TỐT NGHIỆP CỬ NHÂN CNTT")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC CÔNG NGHỆ\nBẰNG PHƯƠNG PHÁP PHÂN LOẠI TỰ ĐỘNG\nVÀ PHÁT HIỆN BẤT THƯỜNG")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run("Chuyên ngành: CÔNG NGHỆ THÔNG TIN")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_gv = doc.add_paragraph()
    p_gv.paragraph_format.left_indent = Cm(4.5)
    p_gv.paragraph_format.space_before = Pt(4)
    p_gv.paragraph_format.space_after = Pt(14)
    p_gv.paragraph_format.line_spacing = 1.3
    r = p_gv.add_run("GIẢNG VIÊN HƯỚNG DẪN:\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p_gv.add_run("   ThS. BÙI CÔNG DANH (Khoa Công nghệ Thông tin – HUIT)\n   Email: danhbc@huit.edu.vn | Điện thoại: 0936598931")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_info2 = doc.add_paragraph()
    p_info2.paragraph_format.left_indent = Cm(4.5)
    p_info2.paragraph_format.space_before = Pt(4)
    p_info2.paragraph_format.space_after = Pt(4)
    p_info2.paragraph_format.line_spacing = 1.3
    r = p_info2.add_run("SINH VIÊN THỰC HIỆN:\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    for name, mssv, lop in students:
        r = p_info2.add_run(f"   {name} – {mssv} – {lop}\n")
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(0, 0, 0)

    p_date2 = doc.add_paragraph()
    p_date2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date2.paragraph_format.space_before = Pt(30)
    p_date2.paragraph_format.space_after = Pt(10)
    r = p_date2.add_run("TP. HỒ CHÍ MINH, THÁNG 10 NĂM 2026")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

def add_front_matter_sections(doc):
    # ==================== SECTION 2: FRONT MATTER (i, ii, iii...) ====================
    s2 = doc.add_section()
    set_page_setup(s2)
    set_section_page_numbering(s2, fmt="lowerRoman", start=1)
    add_centered_page_number_to_footer(s2)

    # 1. LỜI CAM ĐOAN
    add_chapter_title(doc, "LỜI CAM ĐOAN")
    add_body(doc, "Chúng tôi xin cam đoan khóa luận tốt nghiệp với đề tài \"Hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường\" (Mã đề tài: CNTT-KLCN001) là công trình nghiên cứu khoa học do chính nhóm sinh viên chúng tôi thực hiện dưới sự chỉ dẫn khoa học tận tình của Thầy hướng dẫn ThS. Bùi Công Danh.")
    add_body(doc, "Các số liệu khảo sát, kết quả thực nghiệm mô hình trí tuệ nhân tạo (PhoBERT, LinearSVC, Isolation Forest, Deep Autoencoder, NetworkX) và toàn bộ mã nguồn chương trình ứng dụng web được trình bày trong bản báo cáo này hoàn toàn trung thực, khách quan và được nhóm tự tay thiết kế, cài đặt, kiểm thử nghiêm ngặt tại Trường Đại học Công Thương TP. Hồ Chí Minh.")
    add_body(doc, "Tất cả các tài liệu, ấn phẩm khoa học, thuật toán và thư viện mã nguồn mở kế thừa đều được trích dẫn nguồn gốc xuất xứ rõ ràng theo đúng chuẩn mực liêm chính học thuật quốc tế. Chúng tôi hoàn toàn chịu mọi trách nhiệm trước Hội đồng chấm khóa luận tốt nghiệp, Ban Chủ nhiệm Khoa Công nghệ Thông tin và Nhà trường nếu có bất kỳ hành vi sao chép hay gian lận học thuật nào xảy ra.")

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.left_indent = Cm(8.5)
    p_sig.paragraph_format.space_before = Pt(20)
    p_sig.paragraph_format.space_after = Pt(4)
    r = p_sig.add_run("TP. Hồ Chí Minh, ngày 19 tháng 10 năm 2026\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p_sig.add_run("Đại diện nhóm sinh viên thực hiện\n\n\n\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p_sig.add_run("Nguyễn Thành Phát – Lê Thanh Phong – Nguyễn Thành Tài")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    # 2. LỜI CẢM ƠN
    add_chapter_title(doc, "LỜI CẢM ƠN")
    add_body(doc, "Để hoàn thành khóa luận tốt nghiệp đại học này một cách trọn vẹn và đạt chất lượng cao, nhóm chúng tôi đã nhận được sự quan tâm, động viên và hỗ trợ vô cùng to lớn từ quý Thầy Cô, gia đình và bạn bè.")
    add_body(doc, "Trước hết, chúng tôi xin bày tỏ lòng biết ơn sâu sắc nhất tới Ban Giám hiệu Trường Đại học Công Thương TP. Hồ Chí Minh (HUIT) và quý Thầy Cô Khoa Công nghệ Thông tin đã tận tình giảng dạy, truyền đạt những nền tảng tri thức vững chắc về khoa học máy tính, kỹ thuật phần mềm, cơ sở dữ liệu và trí tuệ nhân tạo trong suốt những năm tháng học tập tại giảng đường đại học.")
    add_body(doc, "Đặc biệt, nhóm chúng tôi xin gửi lời cảm ơn chân thành và sâu sắc nhất tới Thầy hướng dẫn ThS. Bùi Công Danh. Trong suốt quá trình thực hiện đề tài từ Tuần 1 đến Tuần 12, Thầy đã luôn định hướng khoa học chuẩn xác, truyền cảm hứng nghiên cứu, tận tình góp ý và kiểm duyệt từng giải thuật, từng chức năng hệ thống để nhóm hoàn thiện sản phẩm một cách chỉn chu và thiết thực nhất.")
    add_body(doc, "Nhóm cũng xin chân thành cảm ơn các Thầy Cô cán bộ Phòng Quản lý Khoa học và các Giảng viên tại HUIT đã nhiệt tình tham gia trả lời phiếu khảo sát hiện trạng, cung cấp các biểu mẫu nghiệp vụ thực tế và dữ liệu mẫu quý báu để nhóm thử nghiệm mô hình phân loại và phát hiện bất thường.")
    add_body(doc, "Sau cùng, chúng tôi xin gửi lời tri ân vô hạn đến Cha Mẹ và Gia đình – điểm tựa tinh thần vững chắc luôn là nguồn động lực to lớn giúp chúng tôi vượt qua mọi khó khăn trong suốt chặng đường học tập.")

    # 3. NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN
    add_chapter_title(doc, "NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN")
    for _ in range(18):
        p_dot = doc.add_paragraph()
        p_dot.paragraph_format.space_before = Pt(6)
        p_dot.paragraph_format.space_after = Pt(6)
        p_dot.paragraph_format.line_spacing = 1.3
        pPr = p_dot._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="8504"/></w:tabs>'))
        r = p_dot.add_run()
        r.add_tab()

    p_gv_sig = doc.add_paragraph()
    p_gv_sig.paragraph_format.left_indent = Cm(8.5)
    p_gv_sig.paragraph_format.space_before = Pt(20)
    p_gv_sig.paragraph_format.space_after = Pt(4)
    r = p_gv_sig.add_run("TP. Hồ Chí Minh, ngày ... tháng 10 năm 2026\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p_gv_sig.add_run("GIẢNG VIÊN HƯỚNG DẪN\n(Ký và ghi rõ họ tên)\n\n\n\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p_gv_sig.add_run("ThS. BÙI CÔNG DANH")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    # 4. TÓM TẮT ĐỀ TÀI & ABSTRACT
    add_chapter_title(doc, "TÓM TẮT ĐỀ TÀI")
    add_body(doc, "Trong bối cảnh đẩy mạnh chuyển đổi số giáo dục đại học, công tác quản lý hồ sơ Khoa học và Công nghệ (KH&CN) đóng vai trò then chốt trong việc đánh giá năng lực nghiên cứu, quy đổi giờ chuẩn giảng dạy và kiểm định chất lượng cơ sở giáo dục. Tuy nhiên, quy trình rà soát thủ công hiện nay tại các trường đại học đang bộc lộ nhiều hạn chế: thời gian đối chiếu kéo dài, áp lực lớn đối với cán bộ quản lý khoa học, và nguy cơ bỏ sót các hồ sơ khai sai lĩnh vực chuyên ngành, đăng trên tạp chí cảnh báo/ngừng chỉ mục hoặc trùng lặp nội dung nghiên cứu.")
    add_body(doc, f"Khóa luận tốt nghiệp này tập trung nghiên cứu, thiết kế và xây dựng hoàn chỉnh 'Hệ thống hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường' (Mã đề tài: CNTT-KLCN001). Hệ thống triển khai đầy đủ 06 nhóm chức năng cốt lõi theo đề cương: (1) Quản lý hồ sơ KH&CN hỗ trợ cả 03 nhóm công trình (Bài báo khoa học, Đề tài NCKH và Sáng chế / Giải pháp hữu ích) tích hợp gợi ý trực tuyến Live AI Preview; (2) Quản lý đầy đủ thao tác Thêm – Sửa – Ẩn/Hiện cho 04 danh mục chuẩn hóa (Lĩnh vực nghiên cứu, Loại công trình, Tạp chí/Hội thảo và Danh mục Tác giả); (3) Quy trình 5 bước xét duyệt tính hợp lệ với trạng thái 'Yêu cầu bổ sung' và lưu vết lịch sử bất biến; (4) Tra cứu, lọc đa tiêu chí (theo từ khóa, tác giả, đơn vị, lĩnh vực, loại hình, năm, cảnh báo) và xuất báo cáo Excel/PDF; (5) Tích hợp mô hình học sâu PyTorch VinAI PhoBERT (vinai/phobert-base) phân loại tự động đồng thời Lĩnh vực chuyên ngành (đạt Macro F1-score {pho_m['f1_score']}%, vượt +{sel_m['delta_f1']}% so với TF-IDF + LinearSVC và vượt +{man_m['phobert_auto']['delta_f1_vs_manual']}% so với phân loại thủ công trên tập kiểm thử độc lập) và Nhóm loại công trình (F1 đạt {pho_m['work_type_f1']}%) kèm vòng lặp phản hồi (Feedback Loop); (6) Phát hiện bất thường đa tầng kết hợp Đồ thị tri thức NetworkX, Rừng cô lập (Isolation Forest) và Mạng nơ-ron Tự mã hóa PyTorch Autoencoder (torch.nn.Module) cùng Hàng đợi xét duyệt ưu tiên theo điểm rủi ro liên tục.")
    add_body(doc, "Từ khóa: Quản lý hồ sơ KH&CN; Phân loại văn bản đa nhiệm; PyTorch PhoBERT; TF-IDF + LinearSVC; Isolation Forest; PyTorch Autoencoder; Đồ thị tri thức NetworkX; Hàng đợi ưu tiên.", bold_prefix="Từ khóa: ")

    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abs.paragraph_format.space_before = Pt(16)
    p_abs.paragraph_format.space_after = Pt(10)
    r_abs = p_abs.add_run("ABSTRACT")
    r_abs.font.name = 'Times New Roman'
    r_abs.font.size = Pt(16)
    r_abs.font.bold = True
    r_abs.font.color.rgb = RGBColor(0, 0, 0)

    add_body(doc, "In the context of digital transformation in higher education, managing scientific and technological (S&T) records plays an essential role in evaluating faculty research productivity, converting academic hours, and supporting institutional accreditation. Nevertheless, traditional manual review workflows suffer from prolonged processing times, heavy administrative workloads, and high risks of overlooking misclassified research domains, predatory or discontinued journals, and duplicate submissions.")
    add_body(doc, f"This graduation thesis researches, designs, and develops an 'S&T Record Management Support System Using Automated Classification and Anomaly Detection' (Topic Code: CNTT-KLCN001). The web system delivers 6 complete functional modules: (1) Dynamic S&T profile management covering scientific articles, research projects, and patents/utility solutions with real-time Live AI Preview; (2) Full CRUD catalog management across 4 standardized catalogs (Research Fields, Work Types, Valid Journals, and Faculty Authors); (3) A 5-step validity appraisal workflow with 'Revision Requested' status and immutable audit trails; (4) Multi-criteria filtering (by keyword, author, department, field, work type, year, and risk flag) with Excel/PDF report generation; (5) Real PyTorch fine-tuning of VinAI PhoBERT (vinai/phobert-base) for multi-task classification of both research fields (achieving {pho_m['f1_score']}% Macro F1-score, outperforming TF-IDF + LinearSVC by +{sel_m['delta_f1']}% and manual classification by +{man_m['phobert_auto']['delta_f1_vs_manual']}% on an independent 60-record test set) and work-type groups ({pho_m['work_type_f1']}% F1-score) with an interactive human-in-the-loop feedback mechanism; (6) Unsupervised anomaly detection combining a NetworkX knowledge graph, Isolation Forest, and a PyTorch Autoencoder (torch.nn.Module) to rank pending submissions in a continuous risk-prioritized review queue.")
    add_body(doc, "Keywords: S&T Record Management; Multi-task Text Classification; PyTorch PhoBERT; TF-IDF + LinearSVC; Isolation Forest; PyTorch Autoencoder; NetworkX Knowledge Graph; Priority Review Queue.", bold_prefix="Keywords: ")

    # 5. MỤC LỤC TỰ ĐỘNG
    add_chapter_title(doc, "MỤC LỤC")
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_before = Pt(6)
    p_toc.paragraph_format.space_after = Pt(6)
    p_toc.paragraph_format.line_spacing = 1.3
    p_toc._p.append(parse_xml(r'<w:fldSimple %s w:instr="TOC \o &quot;1-3&quot; \h \z \u"/>' % nsdecls('w')))

    # 6. DANH MỤC CÁC KÝ HIỆU VÀ CHỮ VIẾT TẮT
    add_chapter_title(doc, "DANH MỤC CÁC KÝ HIỆU VÀ CHỮ VIẾT TẮT")
    abbr_headers = ["Ký hiệu viết tắt", "Tên tiếng Anh đầy đủ", "Ý nghĩa / Tên tiếng Việt"]
    abbr_data = [
        ["AI", "Artificial Intelligence", "Trí tuệ nhân tạo"],
        ["NLP", "Natural Language Processing", "Xử lý ngôn ngữ tự nhiên"],
        ["PhoBERT", "Pre-trained RoBERTa for Vietnamese", "Mô hình ngôn ngữ tiền huấn luyện cho tiếng Việt (VinAI)"],
        ["SVM / LinearSVC", "Support Vector Machine / Linear Classifier", "Máy vectơ hỗ trợ phân loại tuyến tính"],
        ["TF-IDF", "Term Frequency - Inverse Document Frequency", "Tần suất từ - Nghịch đảo tần suất văn bản"],
        ["IF / iTree", "Isolation Forest / Isolation Tree", "Thuật toán Rừng cây cô lập phát hiện ngoại lai"],
        ["AE / Autoencoder", "PyTorch Neural Network Autoencoder", "Mạng nơ-ron tự mã hóa nén tái tạo đặc trưng (PyTorch)"],
        ["MSE", "Mean Squared Error", "Sai số bình phương trung bình"],
        ["BPE", "Byte-Pair Encoding", "Thuật toán mã hóa cặp byte tách từ phụ (subword)"],
        ["WoS", "Web of Science (Clarivate Analytics)", "Hệ thống cơ sở dữ liệu ấn phẩm khoa học uy tín"],
        ["Scopus", "Scopus Database (Elsevier)", "Hệ thống cơ sở dữ liệu thư mục học thuật quốc tế"],
        ["HĐGSNN", "State Council for Professorship", "Hội đồng Giáo sư Nhà nước"],
        ["SHTT / GPHI", "Intellectual Property / Utility Solution", "Sở hữu trí tuệ / Giải pháp hữu ích"],
        ["ISSN", "International Standard Serial Number", "Mã số chuẩn quốc tế cho ấn phẩm định kỳ / tạp chí"],
        ["RBAC", "Role-Based Access Control", "Kiểm soát truy cập dựa trên vai trò người dùng"],
        ["CRUD", "Create, Read, Update, Delete", "Nhóm thao tác Thêm mới, Tra cứu, Cập nhật, Xóa/Ẩn dữ liệu"],
        ["DBMS", "Database Management System", "Hệ quản trị cơ sở dữ liệu quan hệ (SQLite 3)"],
        ["BCE", "Boundary - Control - Entity", "Mẫu kiến trúc lớp phân tích tương tác đối tượng"],
        ["FIFO", "First In, First Out", "Nguyên lý hàng đợi vào trước ra trước"],
        ["KH&CN", "Science and Technology", "Khoa học và Công nghệ"],
        ["HUIT", "Ho Chi Minh City University of Industry and Trade", "Trường Đại học Công Thương TP. Hồ Chí Minh"],
    ]
    add_table_custom(doc, 0, "Danh mục các ký hiệu và chữ viết tắt sử dụng trong khóa luận",
                     abbr_headers, abbr_data, col_widths=[3.0, 5.5, 6.5], has_caption=False)

    # 7. DANH MỤC HÌNH TỰ ĐỘNG
    add_chapter_title(doc, "DANH MỤC HÌNH")
    p_fig_toc = doc.add_paragraph()
    p_fig_toc.paragraph_format.space_before = Pt(6)
    p_fig_toc.paragraph_format.space_after = Pt(6)
    p_fig_toc.paragraph_format.line_spacing = 1.3
    p_fig_toc._p.append(parse_xml(r'<w:fldSimple %s w:instr="TOC \h \z \c &quot;Hình&quot;"/>' % nsdecls('w')))

    # 8. DANH MỤC BẢNG TỰ ĐỘNG
    add_chapter_title(doc, "DANH MỤC BẢNG")
    p_tbl_toc = doc.add_paragraph()
    p_tbl_toc.paragraph_format.space_before = Pt(6)
    p_tbl_toc.paragraph_format.space_after = Pt(6)
    p_tbl_toc.paragraph_format.line_spacing = 1.3
    p_tbl_toc._p.append(parse_xml(r'<w:fldSimple %s w:instr="TOC \h \z \c &quot;Bảng&quot;"/>' % nsdecls('w')))


def add_main_content(doc, helpers):
    # ==================== SECTION 3: NỘI DUNG CHÍNH (ĐÁNH SỐ TRANG 1, 2, 3...) ====================
    s3 = doc.add_section()
    set_page_setup(s3)
    set_section_page_numbering(s3, fmt="decimal", start=1)
    add_centered_page_number_to_footer(s3)

    from scripts.report_chapters import build_all_chapters
    from scripts.report_chuong1 import add_chuong_1
    from scripts.report_chuong2 import add_chuong_2
    from scripts.report_chuong3 import add_chuong_3
    from scripts.report_chuong4 import add_chuong_4
    from scripts.report_chuong5 import add_chuong_5
    from scripts.report_closing import add_closing_sections

    print("Building MỞ ĐẦU...")
    build_all_chapters(doc, helpers)

    print("Building CHƯƠNG 1...")
    add_chuong_1(doc, helpers)

    print("Building CHƯƠNG 2...")
    add_chuong_2(doc, helpers)

    print("Building CHƯƠNG 3...")
    add_chuong_3(doc, helpers)

    print("Building CHƯƠNG 4...")
    add_chuong_4(doc, helpers)

    print("Building CHƯƠNG 5...")
    add_chuong_5(doc, helpers)

    print("Building KẾT LUẬN & TÀI LIỆU THAM KHẢO...")
    add_closing_sections(doc, helpers)


def update_fields_via_word_com(docx_path):
    ps_cmd = f"""
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone
try {{
    $doc = $word.Documents.Open('{docx_path}')
    $doc.Fields.Update()
    foreach ($toc in $doc.TablesOfContents) {{ $toc.Update() }}
    foreach ($tof in $doc.TablesOfFigures) {{ $tof.Update() }}
    $doc.Save()
    $doc.Close()
    Write-Output "WORD_COM_SUCCESS"
}} catch {{
    Write-Output "WORD_COM_ERROR: $($_.Exception.Message)"
}} finally {{
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}}
"""
    ps_file = os.path.join(BASE_DIR, "scripts", "run_word_update.ps1")
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_cmd)
    
    res = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_file],
                         capture_output=True, text=True, encoding="utf-8")
    print(res.stdout)
    if os.path.exists(ps_file):
        os.remove(ps_file)


def main():
    print("Generating complete Bao_Cao_Khoa_Luan.docx...")
    doc = Document()
    configure_doc_styles(doc)

    helpers = {
        "add_chapter_title": add_chapter_title,
        "add_h1": add_h1,
        "add_h2": add_h2,
        "add_h3": add_h3,
        "add_body": add_body,
        "add_bullet_1": add_bullet_1,
        "add_bullet_2": add_bullet_2,
        "add_figure_custom": add_figure_custom,
        "add_table_custom": add_table_custom,
        "add_omml_tfidf": add_omml_tfidf,
        "add_omml_linearsvc": add_omml_linearsvc,
        "add_omml_platt_scaling": add_omml_platt_scaling,
        "add_omml_scaled_attention": add_omml_scaled_attention,
        "add_omml_multihead_attention": add_omml_multihead_attention,
        "add_omml_joint_loss": add_omml_joint_loss,
        "add_omml_degree_centrality": add_omml_degree_centrality,
        "add_omml_isolation_forest": add_omml_isolation_forest,
        "add_omml_pytorch_autoencoder": add_omml_pytorch_autoencoder,
        "add_omml_cosine_similarity": add_omml_cosine_similarity,
        "add_omml_risk_score": add_omml_risk_score,
        "add_omml_classification_metrics": add_omml_classification_metrics,
        "bench_data": bench_data
    }

    print("Adding Section 1: Cover Pages...")
    add_cover_pages(doc)

    print("Adding Section 2: Front Matter (Roman numerals i, ii, iii...)...")
    add_front_matter_sections(doc)

    print("Adding Section 3: Main Content (Arabic numerals 1, 2, 3...)...")
    add_main_content(doc, helpers)

    print("Configuring document settings to update fields on open...")
    doc.settings.element.append(parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>'))

    print(f"Saving to {OUTPUT_DOCX}...")
    doc.save(OUTPUT_DOCX)

    print("Updating fields and TOCs via Word COM...")
    update_fields_via_word_com(OUTPUT_DOCX)
    print("DONE! File successfully generated.")


if __name__ == "__main__":
    main()

