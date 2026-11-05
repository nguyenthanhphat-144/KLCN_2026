# -*- coding: utf-8 -*-
"""
Trình tạo Báo cáo Khóa luận Tốt nghiệp Đại học Công Thương TP. HCM (HUIT)
File kết quả: Bao_Cao_Khoa_Luan.docx
Đề tài: HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC CÔNG NGHỆ BẰNG PHƯƠNG PHÁP PHÂN LOẠI TỰ ĐỘNG VÀ PHÁT HIỆN BẤT THƯỜNG
Mã đề tài: CNTT-KLCN001
Tuân thủ tuyệt đối quy định thể thức trình bày trong 'Mau bao cao_Do an_Khoa luan_2025.docx'.

Các quy chuẩn đã thiết lập nghiêm ngặt:
1. Font Times New Roman toàn bộ văn bản.
2. Màu chữ 100% trắng đen (#000000), không dùng font màu.
3. Căn lề: Trái 3.5cm, Phải 2.5cm, Trên 2.5cm, Dưới 2.5cm. Khổ giấy A4.
4. Đề mục:
   - Cấp chương: cỡ chữ 18, in hoa, canh giữa, tô đậm (outlineLvl=0).
   - Đề mục 1.1: cỡ chữ 14, in hoa, canh trái, tô đậm (outlineLvl=1).
   - Đề mục 1.1.1 & 1.1.1.1: cỡ chữ 13, in thường, canh trái, tô đậm (outlineLvl=2, 3).
5. Đề mục hình: Phía dưới hình, canh giữa, cỡ chữ 12, tô đậm, 'Hình X.Y: ...'.
6. Đề mục bảng: Phía trên bảng, canh giữa, cỡ chữ 12, tô đậm, 'Bảng X.Y: ...'.
   Nội dung trong bảng: Cỡ chữ 12, line spacing multiple 1.3, before/after 0pt. Lặp lại tiêu đề bảng nếu sang trang (tblHeader).
7. Đoạn văn: Cỡ chữ 13, first line indent 1.0cm, line spacing multiple 1.3, before 6pt, after 6pt, canh đều 2 bên (Justified).
8. Bullet:
   - Cấp 1: dấu gạch '–', position 0.3cm, text indent 0.8cm, line spacing 1.3, before/after 6pt, canh đều 2 bên.
   - Cấp 2: dấu '+', position 0.8cm, text indent 1.3cm, line spacing 1.3, before/after 6pt, canh đều 2 bên.
9. Đánh số trang: Canh giữa dưới trang. Bìa không đánh số. Trang phụ (Lời cam đoan, Lời cảm ơn, Nhận xét GVHD, Tóm tắt, Mục lục, Danh mục) đánh số la mã (i, ii, iii,...). Trang nội dung từ Mở đầu trở đi đánh số thập phân (1, 2, 3,...).
10. Mục lục, Danh mục hình, Danh mục bảng: Tạo tự động qua field codes Word (TOC \\o "1-3" \\h \\z \\u, TOC \\c "Hình" \\h \\z, TOC \\c "Bảng" \\h \\z) kèm updateFields trong settings.xml và pre-rendered dot leader tabs (8504 dxa).
11. Phần 'NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN': Để trống hoàn toàn với các dòng chấm dot leader (8504 dxa) và khung ký tên như file mẫu để giảng viên tự điền.
"""

import os
import sys
import json
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_FILE = os.path.join(BASE_DIR, "Bao_Cao_Khoa_Luan.docx")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
BENCH_JSON = os.path.join(BASE_DIR, "docs", "Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json")

# ==================== XML HELPER FUNCTIONS ====================

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn('w:shd'))
    if old is not None:
        tcPr.remove(old)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>')
    tcMar = tcPr.find(qn('w:tcMar'))
    if tcMar is not None:
        tcMar.addprevious(shd)
    else:
        tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn('w:tcMar'))
    if old is not None:
        tcPr.remove(old)
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="808080", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblLook = tblPr.find(qn('w:tblLook'))
    if tblLook is not None:
        tblLook.addprevious(borders)
    else:
        tblPr.append(borders)

def add_page_number_field(run):
    p = run._r.getparent()
    p.remove(run._r)
    p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'))
    p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'))
    p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'))
    p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t>1</w:t></w:r>'))
    p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'))

def set_section_page_numbering(section, fmt='decimal', start=1):
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        cols = sectPr.find(qn('w:cols'))
        if cols is not None:
            cols.addprevious(pgNumType)
        else:
            sectPr.append(pgNumType)
    pgNumType.set(qn('w:fmt'), fmt)
    if start is not None:
        pgNumType.set(qn('w:start'), str(start))

def set_margins(section):
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3.5)
    section.right_margin = Cm(2.5)
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)

# ==================== PARAGRAPH BUILDERS ====================

def add_chapter_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(14)
    p.paragraph_format.line_spacing = 1.3
    
    # Add outline level 0 so Word TOC catches it natively
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
    
    run = p.add_run(text.upper())
    run.font.name = 'Times New Roman'
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_h1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    
    # Add outline level 1 so Word TOC catches it natively
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="1"/>'))
    
    run = p.add_run(text.upper())
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.3
    
    # Add outline level 2 so Word TOC catches it natively
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="2"/>'))
    
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_h3(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.3
    
    # Add outline level 3 so Word TOC catches it natively
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="3"/>'))
    
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_body(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(1.0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Times New Roman'
        r_pre.font.size = Pt(13)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(0, 0, 0)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_bullet_1(doc, text, bold_prefix=None):
    # Quy chuẩn: Bullet cấp 1 dấu "–", Position 0.3 cm, Text indent 0.8 cm
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.5) # 0.8 - 0.5 = 0.3 cm position
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    r_bullet = p.add_run("– ")
    r_bullet.font.name = 'Times New Roman'
    r_bullet.font.size = Pt(13)
    r_bullet.font.bold = True
    r_bullet.font.color.rgb = RGBColor(0, 0, 0)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Times New Roman'
        r_pre.font.size = Pt(13)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(0, 0, 0)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_bullet_2(doc, text, bold_prefix=None):
    # Quy chuẩn: Bullet cấp 2 dấu "+", Position 0.8 cm, Text indent 1.3 cm
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(1.3)
    p.paragraph_format.first_line_indent = Cm(-0.5) # 1.3 - 0.5 = 0.8 cm position
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    r_bullet = p.add_run("+ ")
    r_bullet.font.name = 'Times New Roman'
    r_bullet.font.size = Pt(13)
    r_bullet.font.bold = True
    r_bullet.font.color.rgb = RGBColor(0, 0, 0)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Times New Roman'
        r_pre.font.size = Pt(13)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(0, 0, 0)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_table_custom(doc, caption_text, headers, rows_data, col_widths=None):
    # Quy chuẩn: Đề mục bảng đặt PHÍA TRÊN bảng, canh giữa, cỡ chữ 12, tô đậm
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(8)
    p_cap.paragraph_format.space_after = Pt(4)
    r_cap = p_cap.add_run(caption_text)
    r_cap.font.name = 'Times New Roman'
    r_cap.font.size = Pt(12)
    r_cap.font.bold = True
    r_cap.font.color.rgb = RGBColor(0, 0, 0)

    table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="808080")

    # Header Row: lặp lại tiêu đề bảng nếu sang trang mới (cantSplit MUST precede tblHeader in CT_TrPr)
    hdr_row = table.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    if col_widths and len(col_widths) == len(headers):
        for row in table.rows:
            for idx, w in enumerate(col_widths):
                row.cells[idx].width = Cm(w)

    hdr_cells = hdr_row.cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "F2F2F2") # Light gray header
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.3 # Quy chuẩn: multiple 1.3
        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12) # Quy chuẩn: cỡ chữ 12
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)

    # Data Rows: Cỡ chữ 12, Line spacing multiple 1.3, Before 0pt, After 0pt
    for row_idx, data in enumerate(rows_data):
        row = table.rows[row_idx + 1]
        trPr_row = row._tr.get_or_add_trPr()
        trPr_row.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        row_cells = row.cells
        bg_color = "F9F9F9" if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(data):
            row_cells[col_idx].text = str(val)
            set_cell_background(row_cells[col_idx], bg_color)
            set_cell_margins(row_cells[col_idx], top=80, bottom=80, left=120, right=120)
            p = row_cells[col_idx].paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.3 # Quy chuẩn: multiple 1.3
            if str(val).strip().isdigit() or len(str(val).strip()) <= 4 or str(val).strip().endswith('%'):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12) # Quy chuẩn: cỡ chữ 12
                run.font.color.rgb = RGBColor(0, 0, 0)

    p_post = doc.add_paragraph()
    p_post.paragraph_format.space_before = Pt(2)
    p_post.paragraph_format.space_after = Pt(6)
    return table

def add_figure_custom(doc, img_rel_path, caption_text, width_cm=14.5):
    # Quy chuẩn: Đề mục hình đặt PHÍA DƯỚI hình, canh giữa, cỡ chữ 12, tô đậm
    img_path = os.path.join(BASE_DIR, img_rel_path)
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Cm(width_cm))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(8)
        r_cap = p_cap.add_run(caption_text)
        r_cap.font.name = 'Times New Roman'
        r_cap.font.size = Pt(12)
        r_cap.font.bold = True
        r_cap.font.color.rgb = RGBColor(0, 0, 0)
    else:
        print(f"Warning: Figure image not found: {img_path}")

def add_blank_supervisor_review(doc):
    """
    Tạo trang 'NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN' để trống hoàn toàn
    với 19 dòng chấm leader='dot' và khung ký tên như file mẫu HUIT 2025.
    """
    add_chapter_title(doc, "NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN")

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(6)
    p_space.paragraph_format.space_after = Pt(6)

    for _ in range(19):
        p_dot = doc.add_paragraph()
        p_dot.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_dot.paragraph_format.space_before = Pt(6)
        p_dot.paragraph_format.space_after = Pt(6)
        p_dot.paragraph_format.line_spacing = 1.5
        p_dot.paragraph_format.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        r = p_dot.add_run("\t")
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.tab_stops.add_tab_stop(Cm(10.0), WD_TAB_ALIGNMENT.CENTER)
    r_t = p_title.add_run("\tGIẢNG VIÊN HƯỚNG DẪN")
    r_t.font.name = 'Times New Roman'
    r_t.font.size = Pt(13)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(0, 0, 0)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(50)
    p_sub.paragraph_format.tab_stops.add_tab_stop(Cm(10.0), WD_TAB_ALIGNMENT.CENTER)
    r_s = p_sub.add_run("\t(Ký và ghi rõ họ tên)")
    r_s.font.name = 'Times New Roman'
    r_s.font.size = Pt(13)
    r_s.font.italic = True
    r_s.font.color.rgb = RGBColor(0, 0, 0)

def add_automatic_toc(doc, toc_items):
    """
    Tạo Mục Lục tự động thông qua Word Field Code 'TOC \\o "1-3" \\h \\z \\u'
    kèm danh mục pre-rendered với tab stop leader='dot' tại 15.0 cm (lề phải).
    """
    add_chapter_title(doc, "MỤC LỤC")

    # Bắt đầu trường TOC tự động
    p_start = doc.add_paragraph()
    p_start.paragraph_format.space_before = Pt(0)
    p_start.paragraph_format.space_after = Pt(0)
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>'))
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'))
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>'))

    for title, page_no in toc_items:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(2)
        p_t.paragraph_format.space_after = Pt(2)
        p_t.paragraph_format.line_spacing = 1.3
        p_t.paragraph_format.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        is_bold = any(title.strip().startswith(prefix) for prefix in [
            'LỜI', 'NHẬN', 'TÓM', 'ABSTRACT', 'MỤC', 'DANH', 'PHẦN MỞ ĐẦU', 'CHƯƠNG', 'KẾT LUẬN', 'TÀI LIỆU'
        ])

        r_t = p_t.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(12)
        r_t.font.bold = is_bold
        r_t.font.color.rgb = RGBColor(0, 0, 0)

        r_tab = p_t.add_run(f"\t{page_no}")
        r_tab.font.name = 'Times New Roman'
        r_tab.font.size = Pt(12)
        r_tab.font.bold = is_bold
        r_tab.font.color.rgb = RGBColor(0, 0, 0)

    # Kết thúc trường TOC
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(0)
    p_end._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'))

def add_automatic_figure_list(doc, figures_list):
    """
    Tạo Danh Mục Hình tự động thông qua Word Field Code 'TOC \\c "Hình" \\h \\z'
    kèm tab stop leader='dot' tại 15.0 cm (lề phải).
    """
    add_chapter_title(doc, "DANH MỤC HÌNH")

    p_start = doc.add_paragraph()
    p_start.paragraph_format.space_before = Pt(0)
    p_start.paragraph_format.space_after = Pt(0)
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>'))
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> TOC \\c "Hình" \\h \\z </w:instrText></w:r>'))
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>'))

    for cap, page_no in figures_list:
        p_f = doc.add_paragraph()
        p_f.paragraph_format.space_before = Pt(2)
        p_f.paragraph_format.space_after = Pt(2)
        p_f.paragraph_format.line_spacing = 1.3
        p_f.paragraph_format.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        r_f = p_f.add_run(cap)
        r_f.font.name = 'Times New Roman'
        r_f.font.size = Pt(12)
        r_f.font.color.rgb = RGBColor(0, 0, 0)

        r_tab = p_f.add_run(f"\t{page_no}")
        r_tab.font.name = 'Times New Roman'
        r_tab.font.size = Pt(12)
        r_tab.font.bold = True
        r_tab.font.color.rgb = RGBColor(0, 0, 0)

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(0)
    p_end._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'))

def add_automatic_table_list(doc, tables_list):
    """
    Tạo Danh Mục Bảng tự động thông qua Word Field Code 'TOC \\c "Bảng" \\h \\z'
    kèm tab stop leader='dot' tại 15.0 cm (lề phải).
    """
    add_chapter_title(doc, "DANH MỤC BẢNG")

    p_start = doc.add_paragraph()
    p_start.paragraph_format.space_before = Pt(0)
    p_start.paragraph_format.space_after = Pt(0)
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>'))
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> TOC \\c "Bảng" \\h \\z </w:instrText></w:r>'))
    p_start._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>'))

    for cap, page_no in tables_list:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(2)
        p_t.paragraph_format.space_after = Pt(2)
        p_t.paragraph_format.line_spacing = 1.3
        p_t.paragraph_format.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        r_t = p_t.add_run(cap)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(12)
        r_t.font.color.rgb = RGBColor(0, 0, 0)

        r_tab = p_t.add_run(f"\t{page_no}")
        r_tab.font.name = 'Times New Roman'
        r_tab.font.size = Pt(12)
        r_tab.font.bold = True
        r_tab.font.color.rgb = RGBColor(0, 0, 0)

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(0)
    p_end._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'))

# ==================== MAIN GENERATOR ====================

def generate_thesis_report():
    print("Beginning generation of Bao_Cao_Khoa_Luan.docx...")
    doc = Document()

    # ==================== SECTION 1: COVER PAGES (No page numbering) ====================
    s1 = doc.sections[0]
    set_margins(s1)
    s1.footer.is_linked_to_previous = False
    p_ft1 = s1.footer.paragraphs[0]
    p_ft1.text = "" # Không đánh số trang bìa

    # ---- TRANG BÌA CHÍNH (In trên giấy bìa) ----
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BỘ CÔNG THƯƠNG\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("TRƯỜNG ĐẠI HỌC CÔNG THƯƠNG TP. HỒ CHÍ MINH\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("KHOA CÔNG NGHỆ THÔNG TIN\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("━━━━━━━━━━━━━━━━━━━ ✦ ✦ ✦ ━━━━━━━━━━━━━━━━━━━")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(30)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("KHÓA LUẬN TỐT NGHIỆP CỬ NHÂN CNTT\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC CÔNG NGHỆ\nBẰNG PHƯƠNG PHÁP PHÂN LOẠI TỰ ĐỘNG\nVÀ PHÁT HIỆN BẤT THƯỜNG\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run("(Mã đề tài: CNTT-KLCN001 | Định hướng: Tích hợp Trí tuệ Nhân tạo)")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(30)
    r = p.add_run("Chuyên ngành: CÔNG NGHỆ THÔNG TIN")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.left_indent = Cm(4.5)
    p_info.paragraph_format.space_before = Pt(10)
    p_info.paragraph_format.space_after = Pt(2)
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

    p_gv = doc.add_paragraph()
    p_gv.paragraph_format.left_indent = Cm(4.5)
    p_gv.paragraph_format.space_before = Pt(8)
    p_gv.paragraph_format.space_after = Pt(35)
    r = p_gv.add_run("GIẢNG VIÊN HƯỚNG DẪN:\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p_gv.add_run("   ThS. BÙI CÔNG DANH (Khoa Công nghệ Thông tin - HUIT)")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12.5)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.paragraph_format.space_before = Pt(30)
    p_date.paragraph_format.space_after = Pt(10)
    r = p_date.add_run("TP. HỒ CHÍ MINH, NĂM 2026")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    doc.add_page_break()

    # ---- TRANG BÌA PHỤ (In giấy trắng thường) ----
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BỘ CÔNG THƯƠNG\nTRƯỜNG ĐẠI HỌC CÔNG THƯƠNG TP. HỒ CHÍ MINH\nKHOA CÔNG NGHỆ THÔNG TIN\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p.add_run("━━━━━━━━━━━━━━━━━━━ ✦ ✦ ✦ ━━━━━━━━━━━━━━━━━━━")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(30)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("KHÓA LUẬN TỐT NGHIỆP CỬ NHÂN CNTT\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC CÔNG NGHỆ\nBẰNG PHƯƠNG PHÁP PHÂN LOẠI TỰ ĐỘNG\nVÀ PHÁT HIỆN BẤT THƯỜNG\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(30)
    r = p.add_run("Chuyên ngành: CÔNG NGHỆ THÔNG TIN\nMã ngành: 7480201")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_gv = doc.add_paragraph()
    p_gv.paragraph_format.left_indent = Cm(4.5)
    p_gv.paragraph_format.space_before = Pt(10)
    p_gv.paragraph_format.space_after = Pt(15)
    r = p_gv.add_run("GIẢNG VIÊN HƯỚNG DẪN:\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    r = p_gv.add_run("   ThS. BÙI CÔNG DANH\n   Email: danhbc@huit.edu.vn | SĐT: 0936598931")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.left_indent = Cm(4.5)
    p_info.paragraph_format.space_before = Pt(5)
    p_info.paragraph_format.space_after = Pt(35)
    r = p_info.add_run("SINH VIÊN THỰC HIỆN:\n")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    for name, mssv, lop in students:
        r = p_info.add_run(f"   {name} – {mssv} – {lop}\n")
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(0, 0, 0)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.paragraph_format.space_before = Pt(30)
    p_date.paragraph_format.space_after = Pt(10)
    r = p_date.add_run("TP. HỒ CHÍ MINH, NĂM 2026")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    # ==================== SECTION 2: PRELIMINARY PAGES (i, ii, iii...) ====================
    s2 = doc.add_section()
    set_margins(s2)
    s2.footer.is_linked_to_previous = False
    set_section_page_numbering(s2, fmt='lowerRoman', start=1)
    p_ft2 = s2.footer.paragraphs[0]
    p_ft2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ft2 = p_ft2.add_run()
    r_ft2.font.name = 'Times New Roman'
    r_ft2.font.size = Pt(10)
    r_ft2.font.color.rgb = RGBColor(0, 0, 0)
    add_page_number_field(r_ft2)

    # ---- LỜI CAM ĐOAN ----
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

    doc.add_page_break()

    # ---- LỜI CẢM ƠN ----
    add_chapter_title(doc, "LỜI CẢM ƠN")
    add_body(doc, "Để hoàn thành khóa luận tốt nghiệp đại học này một cách trọn vẹn và đạt chất lượng cao, nhóm chúng tôi đã nhận được sự quan tâm, động viên và hỗ trợ vô cùng to lớn từ quý Thầy Cô, gia đình và bạn bè.")
    add_body(doc, "Trước hết, chúng tôi xin bày tỏ lòng biết ơn sâu sắc nhất tới Ban Giám hiệu Trường Đại học Công Thương TP. Hồ Chí Minh (HUIT) và quý Thầy Cô Khoa Công nghệ Thông tin đã tận tình giảng dạy, truyền đạt những nền tảng tri thức vững chắc về khoa học máy tính, kỹ thuật phần mềm, cơ sở dữ liệu và trí tuệ nhân tạo trong suốt những năm tháng học tập tại giảng đường đại học.")
    add_body(doc, "Đặc biệt, nhóm chúng tôi xin gửi lời cảm ơn chân thành và sâu sắc nhất tới Thầy hướng dẫn ThS. Bùi Công Danh. Trong suốt quá trình thực hiện đề tài từ Tuần 1 đến Tuần 7 và giai đoạn nâng cấp Revision nghiệp vụ, Thầy đã luôn định hướng khoa học chuẩn xác, truyền cảm hứng nghiên cứu, tận tình góp ý và kiểm duyệt từng giải thuật, từng chức năng hệ thống để nhóm hoàn thiện sản phẩm một cách chỉn chu và thiết thực nhất.")
    add_body(doc, "Nhóm cũng xin chân thành cảm ơn các Thầy Cô cán bộ Phòng Quản lý Khoa học và các Giảng viên tại HUIT đã nhiệt tình tham gia trả lời phiếu khảo sát hiện trạng, cung cấp các biểu mẫu nghiệp vụ thực tế và dữ liệu mẫu quý báu để nhóm thử nghiệm mô hình phân loại và phát hiện bất thường.")
    add_body(doc, "Sau cùng, chúng tôi xin gửi lời tri ân vô hạn đến Cha Mẹ và Gia đình – điểm tựa tinh thần vững chắc luôn là nguồn động lực to lớn giúp chúng tôi vượt qua mọi khó khăn trong suốt chặng đường học tập.")

    doc.add_page_break()

    # ---- NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN (ĐỂ TRỐNG NHƯ MẪU HUIT 2025) ----
    add_blank_supervisor_review(doc)

    doc.add_page_break()

    # ---- TÓM TẮT ĐỀ TÀI & ABSTRACT ----
    with open(BENCH_JSON, "r", encoding="utf-8") as f_bench:
        bench = json.load(f_bench)
    pho_m = bench["phobert_metrics"]
    svm_m = bench["svm_metrics"]
    sel_m = bench["selected_model"]
    man_m = bench["manual_vs_automated"]
    ds_m = bench["dataset_stats"]

    add_chapter_title(doc, "TÓM TẮT ĐỀ TÀI")
    add_body(doc, "Trong bối cảnh đẩy mạnh chuyển đổi số giáo dục đại học, công tác quản lý hồ sơ Khoa học và Công nghệ (KH&CN) đóng vai trò then chốt trong việc đánh giá năng lực nghiên cứu, quy đổi giờ chuẩn giảng dạy và kiểm định chất lượng cơ sở giáo dục. Tuy nhiên, quy trình rà soát thủ công hiện nay tại các trường đại học đang bộc lộ nhiều hạn chế: thời gian đối chiếu kéo dài, áp lực lớn đối với cán bộ quản lý khoa học, và nguy cơ bỏ sót các hồ sơ khai sai lĩnh vực chuyên ngành, đăng trên tạp chí cảnh báo/ngừng chỉ mục hoặc trùng lặp nội dung nghiên cứu.")
    add_body(doc, f"Khóa luận tốt nghiệp này tập trung nghiên cứu, thiết kế và xây dựng hoàn chỉnh 'Hệ thống hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường' (Mã đề tài: CNTT-KLCN001). Hệ thống triển khai đầy đủ 06 nhóm chức năng cốt lõi theo đề cương: (1) Quản lý hồ sơ KH&CN hỗ trợ cả 03 nhóm công trình (Bài báo khoa học, Đề tài NCKH và Sáng chế / Giải pháp hữu ích) tích hợp gợi ý trực tuyến Live AI Preview; (2) Quản lý đầy đủ thao tác Thêm – Sửa – Ẩn/Hiện cho 04 danh mục chuẩn hóa (Lĩnh vực nghiên cứu, Loại công trình, Tạp chí/Hội thảo và Danh mục Tác giả); (3) Quy trình 5 bước xét duyệt tính hợp lệ với trạng thái 'Yêu cầu bổ sung' và lưu vết lịch sử bất biến; (4) Tra cứu, lọc đa tiêu chí (theo từ khóa, tác giả, đơn vị, lĩnh vực, loại hình, năm, cảnh báo) và xuất báo cáo Excel/PDF; (5) Tích hợp mô hình học sâu PyTorch VinAI PhoBERT (vinai/phobert-base) phân loại tự động đồng thời Lĩnh vực chuyên ngành (đạt Macro F1-score {pho_m['f1_score']}%, vượt +{sel_m['delta_f1']}% so với TF-IDF + LinearSVC và vượt +{man_m['phobert_auto']['delta_f1_vs_manual']}% so với phân loại thủ công trên tập kiểm thử độc lập) và Nhóm loại công trình (F1 đạt {pho_m['work_type_f1']}%) kèm vòng lặp phản hồi (Feedback Loop); (6) Phát hiện bất thường đa tầng kết hợp Đồ thị tri thức NetworkX, Rừng cô lập (Isolation Forest) và Mạng nơ-ron Tự mã hóa PyTorch Autoencoder (torch.nn.Module) cùng Hàng đợi xét duyệt ưu tiên theo điểm rủi ro liên tục.")
    add_body(doc, "Từ khóa: Quản lý hồ sơ KH&CN; Phân loại văn bản đa nhiệm; PyTorch PhoBERT; TF-IDF + LinearSVC; Isolation Forest; PyTorch Autoencoder; Đồ thị tri thức NetworkX; Hàng đợi ưu tiên.")

    doc.add_paragraph().paragraph_format.space_before = Pt(10)

    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_abs = p_abs.add_run("ABSTRACT")
    r_abs.font.name = 'Times New Roman'
    r_abs.font.size = Pt(16)
    r_abs.font.bold = True
    r_abs.font.color.rgb = RGBColor(0, 0, 0)

    add_body(doc, "In the context of digital transformation in higher education, managing scientific and technological (S&T) records plays an essential role in evaluating faculty research productivity, converting academic hours, and supporting institutional accreditation. Nevertheless, traditional manual review workflows suffer from prolonged processing times, heavy administrative workloads, and high risks of overlooking misclassified research domains, predatory or discontinued journals, and duplicate submissions.")
    add_body(doc, f"This graduation thesis researches, designs, and develops an 'S&T Record Management Support System Using Automated Classification and Anomaly Detection' (Topic Code: CNTT-KLCN001). The web system delivers 6 complete functional modules: (1) Dynamic S&T profile management covering scientific articles, research projects, and patents/utility solutions with real-time Live AI Preview; (2) Full CRUD catalog management across 4 standardized catalogs (Research Fields, Work Types, Valid Journals, and Faculty Authors); (3) A 5-step validity appraisal workflow with 'Revision Requested' status and immutable audit trails; (4) Multi-criteria filtering (by keyword, author, department, field, work type, year, and risk flag) with Excel/PDF report generation; (5) Real PyTorch fine-tuning of VinAI PhoBERT (vinai/phobert-base) for multi-task classification of both research fields (achieving {pho_m['f1_score']}% Macro F1-score, outperforming TF-IDF + LinearSVC by +{sel_m['delta_f1']}% and manual classification by +{man_m['phobert_auto']['delta_f1_vs_manual']}% on an independent 60-record test set) and work-type groups ({pho_m['work_type_f1']}% F1-score) with an interactive human-in-the-loop feedback mechanism; (6) Unsupervised anomaly detection combining a NetworkX knowledge graph, Isolation Forest, and a PyTorch Autoencoder (torch.nn.Module) to rank pending submissions in a continuous risk-prioritized review queue.")
    add_body(doc, "Keywords: S&T Record Management; Multi-task Text Classification; PyTorch PhoBERT; TF-IDF + LinearSVC; Isolation Forest; PyTorch Autoencoder; NetworkX Knowledge Graph; Priority Review Queue.")

    doc.add_page_break()

    # ---- MỤC LỤC TỰ ĐỘNG ----
    toc_items = [
        ("LỜI CAM ĐOAN", "i"),
        ("LỜI CẢM ƠN", "ii"),
        ("NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN", "iii"),
        ("TÓM TẮT ĐỀ TÀI", "iv"),
        ("ABSTRACT", "iv"),
        ("MỤC LỤC", "v"),
        ("DANH MỤC CÁC KÝ HIỆU VÀ CHỮ VIẾT TẮT", "vi"),
        ("DANH MỤC HÌNH", "vii"),
        ("DANH MỤC BẢNG", "viii"),
        ("PHẦN MỞ ĐẦU", "1"),
        ("   1. Tính cấp thiết và lý do chọn đề tài", "1"),
        ("   2. Mục tiêu nghiên cứu của đề tài", "2"),
        ("   3. Đối tượng và phạm vi nghiên cứu", "2"),
        ("   4. Phương pháp nghiên cứu", "3"),
        ("   5. Ý nghĩa khoa học và giá trị thực tiễn", "3"),
        ("   6. Bố cục của khóa luận tốt nghiệp", "4"),
        ("CHƯƠNG 1: KHẢO SÁT HIỆN TRẠNG VÀ PHÂN TÍCH YÊU CẦU HỆ THỐNG", "5"),
        ("   1.1. Hiện trạng quản lý hồ sơ KH&CN tại trường đại học", "5"),
        ("   1.2. Thách thức và hạn chế cốt lõi của quy trình thủ công", "7"),
        ("   1.3. Phân tích yêu cầu nghiệp vụ của hệ thống", "9"),
        ("   1.4. Yêu cầu phi chức năng của hệ thống", "13"),
        ("CHƯƠNG 2: CƠ SỞ LÝ THUYẾT VÀ MÔ HÌNH THUẬT TOÁN AI", "15"),
        ("   2.1. Quy trình Tiền xử lý văn bản tiếng Việt học thuật", "15"),
        ("   2.2. Thuật toán Phân loại văn bản tự động đa nhiệm (Chức năng 5)", "17"),
        ("      2.2.1. Mô hình Baseline: TF-IDF kết hợp LinearSVC và Platt Scaling", "17"),
        ("      2.2.2. Mô hình Học sâu: PyTorch VinAI PhoBERT đa nhiệm (Multi-task Head)", "20"),
        ("   2.3. Thuật toán Phát hiện Bất thường và Đánh giá Rủi ro (Chức năng 6)", "23"),
        ("      2.3.1. Đồ thị tri thức NetworkX và phân tích độ lệch chuyên môn", "23"),
        ("      2.3.2. Thuật toán Rừng cô lập (Isolation Forest)", "25"),
        ("      2.3.3. Thuật toán Mạng Nơ-ron Tự mã hóa PyTorch Autoencoder", "27"),
        ("      2.3.4. Thuật toán Kiểm tra trùng lặp nội dung Cosine TF-IDF", "29"),
        ("      2.3.5. Điểm Rủi ro Tổng hợp liên tục và Hàng đợi thẩm định ưu tiên", "30"),
        ("CHƯƠNG 3: PHÂN TÍCH VÀ THIẾT KẾ HỆ THỐNG", "32"),
        ("   3.1. Thiết kế Kiến trúc hệ thống tổng thể 3 tầng", "32"),
        ("   3.2. Thiết kế Ca sử dụng (Use Case) chi tiết & Ma trận RBAC", "34"),
        ("   3.3. Thiết kế Cơ sở dữ liệu quan hệ SQLite 3 (khcn_system.db)", "37"),
        ("   3.4. Thiết kế Sơ đồ lớp phân tích (Boundary – Control – Entity)", "41"),
        ("   3.5. Thiết kế Luồng tương tác quy trình 5 bước kiểm tra tính hợp lệ", "43"),
        ("CHƯƠNG 4: CÀI ĐẶT CHI TIẾT CÁC CHỨC NĂNG HỆ THỐNG", "45"),
        ("   4.1. Chức năng 1 – Quản lý Hồ sơ KH&CN động (Bài báo, Đề tài, Sáng chế)", "45"),
        ("   4.2. Chức năng 2 – Quản lý 4 Danh mục chuẩn hóa KH&CN (CRUD đầy đủ)", "48"),
        ("   4.3. Chức năng 3 – Quy trình 5 bước Kiểm tra Tính Hợp lệ & Thẩm định", "50"),
        ("   4.4. Chức năng 4 – Báo cáo Thống kê & Phân hệ Tra cứu Đa tiêu chí", "54"),
        ("   4.5. Chức năng 5 – AI Phân loại Tự động Đa nhiệm & Feedback Loop", "56"),
        ("   4.6. Chức năng 6 – Phát hiện Bất thường & Xếp hạng Ưu tiên Hàng đợi", "59"),
        ("   4.7. Các chức năng hỗ trợ quản trị, Dashboard và đối sánh thực nghiệm", "63"),
        ("CHƯƠNG 5: THỬ NGHIỆM, ĐÁNH GIÁ KẾT QUẢ VÀ TRIỂN KHAI", "66"),
        ("   5.1. Thiết lập Môi trường thử nghiệm thực tế và Bộ dữ liệu chuẩn 13 trường", "66"),
        ("   5.2. Thực nghiệm Đối sánh Công bằng: TF-IDF + SVM, PhoBERT và Thủ công", "68"),
        ("      5.2.1. Thống kê bộ dữ liệu huấn luyện và kiểm thử dùng chung (240 mẫu)", "68"),
        ("      5.2.2. Kết quả đối sánh Precision, Recall, F1-score và Phân loại thủ công", "69"),
        ("      5.2.3. Lý do lựa chọn Fine-tuning PhoBERT làm mô hình chính của dự án", "71"),
        ("   5.3. Đánh giá Thực nghiệm Phân hệ Phát hiện Bất thường (IF vs PyTorch AE)", "72"),
        ("   5.4. Kiểm thử Tích hợp Toàn diện và Đánh giá Hiệu năng Thực tế", "74"),
        ("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", "76"),
        ("TÀI LIỆU THAM KHẢO", "78"),
    ]
    add_automatic_toc(doc, toc_items)

    doc.add_page_break()

    # ---- DANH MỤC CÁC KÝ HIỆU VÀ CHỮ VIẾT TẮT ----
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
    add_table_custom(doc, "Bảng chữ viết tắt: Danh mục các ký hiệu và chữ viết tắt sử dụng trong khóa luận",
                     abbr_headers, abbr_data, col_widths=[3.0, 5.5, 6.5])

    doc.add_page_break()

    # ---- DANH MỤC HÌNH TỰ ĐỘNG ----
    figures_list = [
        ("Hình 1.1: Sơ đồ quy trình 5 bước kiểm tra tính hợp lệ & thẩm định hồ sơ KH&CN", "11"),
        ("Hình 2.1: Kiến trúc pipeline tiền xử lý công bằng và đối sánh mô hình (PhoBERT vs LinearSVC)", "19"),
        ("Hình 2.2: Kiến trúc mô hình phát hiện bất thường đa tầng và hàng đợi thẩm định ưu tiên", "26"),
        ("Hình 3.1: Sơ đồ kiến trúc tổng thể hệ thống 3 tầng (3-Tier MVC Architecture)", "33"),
        ("Hình 4.1: Giao diện Kê khai và nộp hồ sơ động tích hợp tính năng Live AI Preview", "46"),
        ("Hình 4.2: Giao diện Thẩm định chi tiết hồ sơ, lưu vết lịch sử và hiển thị cảnh báo AI", "51"),
        ("Hình 4.3: Giao diện Trung tâm Quản trị AI, đồ thị tri thức và hàng đợi ưu tiên xét duyệt", "60"),
        ("Hình 4.4: Giao diện Dashboard điều hành tổng quan với các thẻ KPI và biểu đồ phân loại", "64"),
        ("Hình 5.1: Biểu đồ đối sánh thực nghiệm phân loại (PhoBERT, LinearSVC, Thủ công) trên tập Test", "69"),
    ]
    add_automatic_figure_list(doc, figures_list)

    doc.add_page_break()

    # ---- DANH MỤC BẢNG TỰ ĐỘNG ----
    tables_list = [
        ("Bảng 1.1: Thống kê định lượng thời gian hao phí và tỷ lệ sai sót trong quy trình thủ công", "8"),
        ("Bảng 1.2: Đối chiếu chi tiết yêu cầu hệ thống theo Đề cương CNTT-KLCN001 và giải pháp triển khai", "10"),
        ("Bảng 2.1: Cấu hình siêu tham số thực tế huấn luyện mô hình PyTorch VinAI PhoBERT", "22"),
        ("Bảng 2.2: So sánh đặc tính giải thuật giữa Rừng cô lập (Isolation Forest) và PyTorch Autoencoder", "28"),
        ("Bảng 3.1: Ma trận phân quyền kiểm soát truy cập dựa trên vai trò (RBAC)", "36"),
        ("Bảng 3.2: Lược đồ thuộc tính bảng dữ liệu cốt lõi HoSoKHCN trong SQLite 3 (khcn_system.db)", "38"),
        ("Bảng 3.3: Lược đồ thuộc tính bảng lưu vết kiểm toán bất biến LichSuXetDuyet", "40"),
        ("Bảng 4.1: Danh mục 6 nhóm lĩnh vực nghiên cứu KH&CN và 4 phân hệ danh mục trong hệ thống", "49"),
        ("Bảng 5.1: Thống kê bộ dữ liệu mô phỏng chuẩn hóa 13 trường dùng chung (240 mẫu)", "68"),
        ("Bảng 5.2: Kết quả thực nghiệm đối sánh công bằng giữa TF-IDF + LinearSVC và Fine-tuning PhoBERT", "70"),
        ("Bảng 5.3: Đối sánh hiệu quả giữa Phân loại Tự động (AI) và Phương pháp Phân loại Thủ công", "71"),
        ("Bảng 5.4: Kết quả thực nghiệm phân hệ phát hiện bất thường trên CSDL vận hành thực tế", "73"),
    ]
    add_automatic_table_list(doc, tables_list)

    # ==================== SECTION 3: MAIN BODY (1, 2, 3...) ====================
    s3 = doc.add_section()
    set_margins(s3)
    s3.footer.is_linked_to_previous = False
    set_section_page_numbering(s3, fmt='decimal', start=1)
    p_ft3 = s3.footer.paragraphs[0]
    p_ft3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ft3 = p_ft3.add_run()
    r_ft3.font.name = 'Times New Roman'
    r_ft3.font.size = Pt(10)
    r_ft3.font.color.rgb = RGBColor(0, 0, 0)
    add_page_number_field(r_ft3)

    # ==================== PHẦN MỞ ĐẦU ====================
    add_chapter_title(doc, "PHẦN MỞ ĐẦU")

    add_h1(doc, "1. Tính cấp thiết và lý do chọn đề tài")
    add_body(doc, "Trong bối cảnh nền kinh tế tri thức và cuộc Cách mạng Công nghiệp lần thứ tư đang diễn ra mạnh mẽ, hoạt động nghiên cứu khoa học và phát triển công nghệ (KH&CN) đã trở thành một trong hai trụ cột cốt lõi quyết định uy tín, thương hiệu và vị thế cạnh tranh của các cơ sở giáo dục đại học. Tại Trường Đại học Công Thương TP. Hồ Chí Minh (HUIT), năng suất nghiên cứu khoa học không chỉ là tiêu chí bắt buộc để đánh giá hoàn thành nhiệm vụ năm học của giảng viên mà còn là cơ sở dữ liệu then chốt phục vụ kiểm định chất lượng cơ sở giáo dục, kiểm định chương trình đào tạo theo chuẩn AUN-QA, cũng như tham gia bảng xếp hạng đại học trong nước và quốc tế.")
    add_body(doc, "Tuy nhiên, khảo sát thực tế quy trình quản lý hồ sơ KH&CN hiện nay tại Nhà trường cho thấy phương thức vận hành vẫn mang tính chất thủ công và phân tán. Vào các đợt quyết toán định kỳ, hàng trăm giảng viên từ các Khoa chuyên ngành phải kê khai thông tin qua các bảng tính Excel rời rạc và nộp hồ sơ minh chứng bản cứng lên Phòng Quản lý Khoa học. Quy trình thủ công này đang đối mặt với những thách thức nghiêm trọng:")
    add_bullet_1(doc, "Áp lực quá tải thời gian rà soát: Cán bộ thẩm định phải kiểm tra chéo thủ công từng mã ISSN, DOI, tên tác giả, đối chiếu danh mục tạp chí Scopus/WoS và Hội đồng Giáo sư Nhà nước (HĐGSNN), khiến thời gian xét duyệt kéo dài trung bình từ 15 đến 20 ngày cho mỗi đợt hồ sơ.")
    add_bullet_1(doc, "Sai lệch phân loại chuyên ngành: Khoảng 18.5% hồ sơ bài báo bị tác giả khai sai lĩnh vực (ví dụ bài báo về xử lý ảnh y tế hoặc tối ưu kinh tế bằng AI lại bị phân loại nhầm giữa Khoa học Máy tính và Khoa học Y Dược), dẫn đến việc phân bổ hồ sơ sai Hội đồng chuyên môn thẩm định.")
    add_bullet_1(doc, "Nguy cơ gian lận và tạp chí kém chất lượng: Tình trạng các tạp chí mạo danh, tạp chí săn mồi (Predatory Journals) hoặc các tạp chí bị ngắt chỉ mục (Discontinued list) trên Scopus/WoS ngày càng tinh vi; đồng thời nguy cơ trùng lặp nội dung nghiên cứu (tự đạo văn hoặc khai trùng công trình giữa các đồng tác giả) rất khó phát hiện bằng mắt thường.")
    add_body(doc, "Nhận thức sâu sắc được các điểm nghẽn thực tiễn nêu trên, nhóm sinh viên đã lựa chọn đề tài: \"Hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường\" (Mã đề tài: CNTT-KLCN001) nhằm xây dựng một giải pháp phần mềm toàn diện, ứng dụng công nghệ Trí tuệ nhân tạo (PhoBERT, Isolation Forest, Deep Autoencoder, NetworkX) để hiện đại hóa, tự động hóa và nâng cao độ tin cậy của công tác quản lý khoa học tại HUIT.")

    add_h1(doc, "2. Mục tiêu nghiên cứu của đề tài")
    add_body(doc, "Mục tiêu tổng quát của đề tài là xây dựng hoàn chỉnh hệ thống ứng dụng web hỗ trợ quản lý vòng đời hồ sơ KH&CN; tự động hóa việc phân loại lĩnh vực nghiên cứu và nhóm loại công trình bằng mô hình học sâu tiếng Việt chuyên sâu; đồng thời phát hiện sớm các hồ sơ có dấu hiệu bất thường để tối ưu hóa thứ tự hàng đợi xét duyệt.")
    add_body(doc, "Để đạt được mục tiêu tổng quát, đề tài đề ra 06 mục tiêu chuyên biệt cụ thể bám sát Đề cương CNTT-KLCN001:")
    add_bullet_1(doc, "Mục tiêu 1: Xây dựng phân hệ quản lý hồ sơ KH&CN chuẩn hóa 13 trường dữ liệu, hỗ trợ đầy đủ 03 nhóm loại hình công trình (Bài báo khoa học, Đề tài NCKH và Sáng chế / Giải pháp hữu ích) tích hợp gợi ý trực tuyến Live AI Preview.")
    add_bullet_1(doc, "Mục tiêu 2: Chuẩn hóa và hỗ trợ đầy đủ thao tác Thêm mới – Cập nhật – Ẩn/Hiện (CRUD) cho 04 phân hệ danh mục hệ thống: Lĩnh vực nghiên cứu (6 chuyên ngành), Loại công trình, Tạp chí/Hội thảo tham chiếu và Danh mục Tác giả.")
    add_bullet_1(doc, "Mục tiêu 3: Số hóa quy trình 5 bước xét duyệt tính hợp lệ của công trình KH&CN, bổ sung trạng thái 'Yêu cầu bổ sung', vòng lặp nộp lại hồ sơ và bảng lưu vết lịch sử bất biến.")
    add_bullet_1(doc, "Mục tiêu 4: Xây dựng công cụ tra cứu từ khóa và lọc đa tiêu chí (theo tác giả, đơn vị công tác, lĩnh vực, loại hình, trạng thái, năm công bố, cờ cảnh báo) kèm kết xuất báo cáo Excel (.xlsx) và PDF.")
    add_bullet_1(doc, f"Mục tiêu 5: Tinh chỉnh thực tế mô hình học sâu PyTorch VinAI PhoBERT (vinai/phobert-base) đối sánh công bằng với TF-IDF + LinearSVC và phương pháp phân loại thủ công trên cùng tập dữ liệu chuẩn 240 mẫu (180 Train / 60 Test), phân loại đồng thời Lĩnh vực chuyên ngành (đạt F1-score {pho_m['f1_score']}%) và Nhóm loại công trình (F1 đạt {pho_m['work_type_f1']}%) kèm vòng lặp phản hồi (Feedback Loop).")
    add_bullet_1(doc, "Mục tiêu 6: Cài đặt mô hình phát hiện bất thường không giám sát kết hợp Đồ thị tri thức NetworkX, Rừng cô lập (Isolation Forest), Mạng nơ-ron Tự mã hóa PyTorch Autoencoder (torch.nn.Module) và tương đồng Cosine TF-IDF; tính Điểm rủi ro tổng hợp liên tục và tự động xếp hạng Hàng đợi xét duyệt ưu tiên.")

    add_h1(doc, "3. Đối tượng và phạm vi nghiên cứu")
    add_body(doc, "Đối tượng nghiên cứu của khóa luận bao gồm:")
    add_bullet_1(doc, "Dữ liệu hồ sơ KH&CN: Văn bản học thuật tiếng Việt gồm tiêu đề công trình/đề tài, tóm tắt nội dung, từ khóa, thông tin tác giả, đơn vị chủ trì, năm công bố, kỳ báo cáo, lĩnh vực, loại công trình, thông tin tạp chí/chỉ mục và sản phẩm dự kiến.")
    add_bullet_1(doc, "Các giải thuật Trí tuệ nhân tạo: Tiền xử lý tách từ tiếng Việt (PyVi), trích xuất đặc trưng TF-IDF + LinearSVC (scikit-learn), tinh chỉnh mô hình Transformer tiếng Việt (PyTorch vinai/phobert-base), giải thuật Rừng cô lập Isolation Forest, mạng tự mã hóa PyTorch Autoencoder và đồ thị tri thức NetworkX.")
    add_body(doc, "Phạm vi nghiên cứu và triển khai thực tế của khóa luận:")
    add_bullet_1(doc, "Về nghiệp vụ: Tập trung vào quy trình quản lý, phân loại, kiểm tra tính hợp lệ và phát hiện bất thường hồ sơ KH&CN cấp Trường đại học.")
    add_bullet_1(doc, "Về công nghệ triển khai: Ứng dụng web được xây dựng trên Python 3.12, Flask 3.x, hệ quản trị cơ sở dữ liệu quan hệ SQLite 3 (khcn_system.db – tuân thủ Mục 7 Đề cương CNTT-KLCN001), thư viện PyTorch 2.6, HuggingFace Transformers, Scikit-learn, NetworkX, OpenPyXL và ReportLab.")

    add_h1(doc, "4. Phương pháp nghiên cứu")
    add_body(doc, "Đề tài kết hợp chặt chẽ giữa nghiên cứu lý thuyết và phát triển thực nghiệm phần mềm:")
    add_bullet_1(doc, "Phương pháp phân tích nghiệp vụ và xây dựng dữ liệu chuẩn hóa: Khảo sát quy trình quản lý hồ sơ KH&CN đại học và xây dựng bộ dữ liệu mô phỏng học thuật chuẩn hóa gồm 240 hồ sơ (13 trường, phủ đều 6 lĩnh vực và 3 nhóm loại công trình, gồm 60% mẫu đơn ngành và 40% mẫu giao thoa liên ngành).")
    add_bullet_1(doc, "Phương pháp phân tích và thiết kế phần mềm: Áp dụng phân tích hướng đối tượng kết hợp kiến trúc 3 tầng (Presentation – Business Logic & AI Engine – Data Persistence) và cơ chế kiểm soát truy cập theo vai trò (RBAC).")
    add_bullet_1(doc, "Phương pháp học máy và học sâu thực nghiệm: Thiết kế hàm tiền xử lý đầu vào công bằng 100% dùng chung cho cả TF-IDF + LinearSVC và PyTorch PhoBERT; áp dụng học chuyển giao (Transfer Learning) đa nhiệm và học không giám sát phát hiện ngoại lai.")
    add_bullet_1(doc, "Phương pháp kiểm định và đánh giá định lượng: Đo lường khách quan qua Accuracy, Macro Precision, Macro Recall, Macro F1-score, thời gian huấn luyện/suy luận thực tế bằng time.perf_counter(), và đối chiếu trực tiếp với phương pháp phân loại thủ công.")

    add_h1(doc, "5. Ý nghĩa khoa học và giá trị thực tiễn của đề tài")
    add_body(doc, "Ý nghĩa khoa học: Đề tài kiểm chứng thực nghiệm sự vượt trội của cơ chế Self-Attention hai chiều trong mô hình ngôn ngữ tiền huấn luyện PhoBERT so với biểu diễn túi từ TF-IDF + LinearSVC khi phân loại các hồ sơ KH&CN tiếng Việt có tính giao thoa liên ngành, đồng thời kết hợp hiệu quả Isolation Forest với PyTorch Autoencoder trên đặc trưng đồ thị NetworkX.")
    add_body(doc, f"Giá trị thực tiễn: Hệ thống cung cấp công cụ phần mềm hoàn chỉnh giúp tự động gợi ý phân loại hồ sơ nhanh gấp ~{man_m['phobert_auto']['speedup_factor_vs_manual']} lần so với rà soát thủ công, cảnh báo sớm các nguy cơ khai sai lĩnh vực, tạp chí không uy tín hoặc trùng lặp nội dung, và minh bạch hóa toàn bộ lịch sử xét duyệt.")

    add_h1(doc, "6. Bố cục của khóa luận tốt nghiệp")
    add_body(doc, "Bản báo cáo khóa luận tốt nghiệp được cấu trúc thành 05 chương chính theo đúng quy định thể thức năm 2025 của Trường Đại học Công Thương TP. HCM:")
    add_bullet_1(doc, "Phần Mở đầu: Trình bày bối cảnh, lý do chọn đề tài, mục tiêu, đối tượng, phạm vi, phương pháp và bố cục báo cáo.")
    add_bullet_1(doc, "Chương 1: Khảo sát hiện trạng và Phân tích yêu cầu hệ thống: Phân tích quy trình quản lý hồ sơ KH&CN, hạn chế của phương thức thủ công và đặc tả 06 nhóm chức năng cốt lõi.")
    add_bullet_1(doc, "Chương 2: Cơ sở lý thuyết và Mô hình thuật toán AI: Trình bày cơ sở toán học của tiền xử lý PyVi, TF-IDF + LinearSVC, PyTorch PhoBERT đa nhiệm, NetworkX, Isolation Forest và PyTorch Autoencoder.")
    add_bullet_1(doc, "Chương 3: Phân tích và Thiết kế hệ thống: Thiết kế kiến trúc 3 tầng, phân quyền RBAC, lược đồ cơ sở dữ liệu quan hệ SQLite 3 (khcn_system.db) và luồng xét duyệt 5 bước.")
    add_bullet_1(doc, "Chương 4: Cài đặt chi tiết các chức năng hệ thống: Trình bày chi tiết việc hiện thực hóa 06 nhóm chức năng, 4 phân hệ danh mục CRUD, vòng lặp phản hồi AI và kết xuất báo cáo.")
    add_bullet_1(doc, "Chương 5: Thử nghiệm, Đánh giá kết quả và Triển khai: Trình bày môi trường phần cứng thực tế, kết quả đối sánh công bằng giữa TF-IDF + SVM, Fine-tuning PhoBERT và phân loại thủ công, cùng kết quả phát hiện bất thường.")
    add_bullet_1(doc, "Phần Kết luận và Hướng phát triển: Tổng kết các kết quả thực tế đạt được, thẳng thắn chỉ ra hạn chế và đề xuất hướng phát triển tiếp theo.")

    doc.add_page_break()

    # ==================== CHƯƠNG 1 ====================
    add_chapter_title(doc, "CHƯƠNG 1: KHẢO SÁT HIỆN TRẠNG VÀ PHÂN TÍCH YÊU CẦU HỆ THỐNG")

    add_h1(doc, "1.1. Hiện trạng quản lý hồ sơ KH&CN tại trường đại học")
    add_h2(doc, "1.1.1. Cơ cấu tổ chức và các đầu mối tham gia quản lý")
    add_body(doc, "Qua tìm hiểu quy trình nghiệp vụ tại Trường Đại học Công Thương TP. Hồ Chí Minh và các cơ sở giáo dục đại học công lập, hoạt động quản lý hồ sơ KH&CN được vận hành theo cơ chế phối hợp giữa các đơn vị chuyên môn gồm 05 đầu mối chính:")
    add_bullet_1(doc, "Đầu mối 1 – Tác giả / Giảng viên nghiên cứu: Là người trực tiếp thực hiện và kê khai công trình khoa học (bài báo, đề tài NCKH, sáng chế/giải pháp hữu ích) kèm tệp minh chứng số.")
    add_bullet_1(doc, "Đầu mối 2 – Thư ký khoa học Khoa / Viện đào tạo: Tiếp nhận danh sách kê khai từ giảng viên trong đơn vị, kiểm tra sơ bộ thể thức trước khi tổng hợp gửi cấp Trường.")
    add_bullet_1(doc, "Đầu mối 3 – Chuyên viên Phòng Quản lý Khoa học: Tiếp nhận hồ sơ toàn trường, kiểm tra tính hợp lệ của thông tin xuất bản, phân loại hồ sơ về đúng lĩnh vực chuyên ngành và loại công trình để trình Hội đồng thẩm định.")
    add_bullet_1(doc, "Đầu mối 4 – Hội đồng Khoa học chuyên ngành thẩm định: Thẩm định giá trị học thuật, đối chiếu uy tín tạp chí công bố, kiểm tra tính trung thực học thuật và ra quyết định công nhận, yêu cầu bổ sung hoặc từ chối.")
    add_bullet_1(doc, "Đầu mối 5 – Ban Giám hiệu Nhà trường: Phê duyệt danh sách tổng hợp để quy đổi giờ chuẩn nghiên cứu khoa học và khen thưởng định kỳ.")

    add_h2(doc, "1.1.2. Phân loại các loại hình hồ sơ KH&CN trong thực tế")
    add_body(doc, "Theo Đề cương CNTT-KLCN001 và thực tiễn quản lý đại học, hồ sơ KH&CN tập trung vào 03 nhóm loại hình công trình trọng tâm:")
    add_bullet_1(doc, "Nhóm 1 – Bài báo khoa học (Quốc tế & Trong nước): Gồm bài báo đăng trên tạp chí quốc tế thuộc danh mục Web of Science (SCIE/SSCI/ESCI), Scopus (Q1–Q4), kỷ yếu hội nghị quốc tế uy tín và bài báo đăng trên tạp chí khoa học trong nước thuộc danh mục tính điểm của Hội đồng Giáo sư Nhà nước (HĐGSNN). Nhóm này yêu cầu đối chiếu chặt chẽ tên tạp chí, mã ISSN và trạng thái chỉ mục nhằm loại trừ tạp chí săn mồi (Predatory) hoặc ngừng chỉ mục (Discontinued).")
    add_bullet_1(doc, "Nhóm 2 – Đề tài nghiên cứu khoa học các cấp: Gồm đề tài cấp Bộ/Tỉnh – Thành phố, đề tài cấp Trường và đề tài nghiên cứu khoa học sinh viên, gắn liền với thuyết minh, quyết định nghiệm thu và danh mục sản phẩm dự kiến.")
    add_bullet_1(doc, "Nhóm 3 – Sở hữu trí tuệ (Sáng chế, Giải pháp hữu ích & Sách chuyên khảo): Gồm bằng độc quyền sáng chế, bằng độc quyền giải pháp hữu ích do Cục Sở hữu trí tuệ cấp hoặc ấn phẩm chuyên khảo phục vụ đào tạo.")

    add_h1(doc, "1.2. Thách thức và hạn chế cốt lõi của quy trình thủ công")
    add_body(doc, "Phân tích quy trình tiếp nhận và rà soát hồ sơ thủ công qua bảng tính rời rạc cho thấy nhiều điểm nghẽn về thời gian xử lý và độ nhất quán phân loại, được tổng hợp tại Bảng 1.1:")

    t1_headers = ["Giai đoạn nghiệp vụ", "Hình thức thực hiện thủ công", "Thời gian hao phí ước tính", "Hạn chế và rủi ro phát sinh"]
    t1_data = [
        ["1. Kê khai hồ sơ", "Giảng viên nhập file Excel/Word rời rạc", "2 – 4 ngày / đợt kê khai", "Dễ thiếu thông tin tác giả, chỉ mục tạp chí hoặc minh chứng"],
        ["2. Tổng hợp cấp Khoa", "Sao chép thủ công từ nhiều tệp riêng lẻ", "3 – 5 ngày / khoa", "Sai lệch định dạng dữ liệu, khó tra cứu theo tác giả/đơn vị"],
        ["3. Phân loại chuyên ngành", "Chuyên viên đọc tiêu đề/tóm tắt bằng mắt", "~2 – 3 phút / 1 hồ sơ", "Dễ nhầm lẫn ở các đề tài liên ngành có thuật ngữ giao thoa"],
        ["4. Kiểm tra chéo bất thường", "Tra cứu thủ công từng ISSN và tên bài", "15 – 20 phút / 1 hồ sơ", "Khó phát hiện tạp chí ngừng chỉ mục hoặc trùng lặp nội dung"],
        ["5. Thống kê & Xuất báo cáo", "Tổng hợp bảng biểu thủ công cuối kỳ", "2 – 3 ngày / đợt", "Số liệu chậm cập nhật, thiếu lịch sử lưu vết xét duyệt"],
    ]
    add_table_custom(doc, "Bảng 1.1: Thống kê định lượng thời gian hao phí và tỷ lệ sai sót trong quy trình thủ công",
                     t1_headers, t1_data, col_widths=[3.0, 4.5, 3.5, 4.0])

    add_body(doc, "Từ phân tích trên, có thể xác định 03 nguyên nhân cốt lõi cần giải quyết bằng phần mềm tích hợp AI:")
    add_bullet_1(doc, "Thiếu nền tảng quản lý tập trung có kiểm soát danh mục chuẩn: Dữ liệu hồ sơ và danh mục lĩnh vực, loại công trình, tạp chí, tác giả chưa được liên kết khóa ngoại đồng bộ.")
    add_bullet_1(doc, "Thiếu mô hình NLP tiếng Việt hỗ trợ phân loại đa nhiệm: Chuyên viên phải tự đọc và gán nhãn lĩnh vực cũng như nhóm công trình hoàn toàn thủ công, thiếu chỉ số độ tin cậy định lượng.")
    add_bullet_1(doc, "Thiếu cơ chế phát hiện bất thường và xếp hạng ưu tiên: Hồ sơ chờ duyệt được xử lý tuần tự theo thời gian nộp (FIFO) nên các hồ sơ có dấu hiệu khai sai lĩnh vực, tạp chí cảnh báo hoặc trùng lặp không được đẩy lên rà soát sớm.")

    add_h1(doc, "1.3. Phân tích yêu cầu nghiệp vụ của hệ thống")
    add_body(doc, "Bám sát Đề cương khóa luận CNTT-KLCN001 và yêu cầu chuẩn hóa quy trình kiểm tra tính hợp lệ công trình KH&CN, hệ thống được thiết kế giải quyết đồng bộ cả yêu cầu quản lý thông tin lẫn yêu cầu tích hợp Trí tuệ nhân tạo:")
    add_bullet_1(doc, "Làm rõ bản chất thẩm định tính hợp lệ: Hệ thống phục vụ kiểm tra tính hợp lệ của hồ sơ công trình KH&CN kèm tệp minh chứng, hỗ trợ 5 trạng thái vòng đời gồm 'Chờ xét duyệt', 'Đang xét duyệt', 'Đã duyệt', 'Từ chối' và 'Yêu cầu bổ sung'.")
    add_bullet_1(doc, "Mở rộng đầy đủ 3 nhóm công trình và 4 danh mục chuẩn: Biểu mẫu kê khai thích ứng linh hoạt cho Bài báo khoa học, Đề tài NCKH và Sáng chế / Giải pháp hữu ích; đồng thời quản trị trọn vẹn 4 danh mục (Lĩnh vực, Loại công trình, Tạp chí hợp lệ, Danh mục Tác giả).")
    add_bullet_1(doc, "Đối sánh minh bạch giữa Học máy truyền thống, Học sâu và Thủ công: Tích hợp phân hệ thực nghiệm độc lập cho phép đối sánh công bằng TF-IDF + LinearSVC, Fine-tuning PhoBERT (PyTorch) và phương pháp phân loại thủ công trên cùng tập dữ liệu.")

    add_body(doc, "Bảng 1.2 tổng hợp chi tiết yêu cầu hệ thống theo Đề cương CNTT-KLCN001 và giải pháp kỹ thuật đã hiện thực hóa:")

    t2_headers = ["Nhóm chức năng", "Yêu cầu trong Đề cương CNTT-KLCN001", "Mở rộng chuẩn hóa nghiệp vụ", "Giải pháp kỹ thuật đã triển khai"]
    t2_data = [
        ["1. Quản lý hồ sơ", "Thêm, sửa, xóa, đính kèm file; lưu Tiêu đề, Tác giả, Đơn vị, Tóm tắt, Từ khóa, Năm, Loại", "Hỗ trợ đủ 3 nhóm: Bài báo, Đề tài NCKH, Sáng chế & GPHI; xóa mềm & xóa cứng", "Form động 13 trường (hoso_form.html) + gợi ý Live AI Preview"],
        ["2. Quản lý danh mục", "Danh mục Lĩnh vực, Loại công trình, Tạp chí hợp lệ, Danh sách Tác giả", "Hỗ trợ đầy đủ Thêm mới, Cập nhật (Sửa) và Ẩn/Hiện bảo toàn khóa ngoại", "4 bảng danh mục chuẩn 3NF trong SQLite + giao diện CRUD"],
        ["3. Xét duyệt hồ sơ", "Chờ duyệt -> Đang duyệt -> Đã duyệt / Từ chối; lưu lịch sử và ý kiến", "Quy trình 5 bước, khóa tự động chống trùng lặp và nhánh 'Yêu cầu bổ sung'", "Bảng LichSuXetDuyet bất biến + vòng lặp nộp lại hồ sơ"],
        ["4. Báo cáo & Tra cứu", "Thống kê theo năm, đơn vị, lĩnh vực, loại; tìm kiếm từ khóa, tác giả; xuất Excel/PDF", "Bộ lọc kết hợp 7 tiêu chí (Từ khóa, Tác giả, Đơn vị, Lĩnh vực, Loại, Năm, Rủi ro)", "Dashboard Chart.js + dịch vụ xuất OpenPyXL & ReportLab"],
        ["5. Phân loại tự động", "Phân loại Lĩnh vực & Loại công trình từ Tiêu đề, Tóm tắt, Từ khóa, Tác giả; xác nhận/chỉnh sửa", "Đối sánh công bằng TF-IDF+SVM, PyTorch PhoBERT và Thủ công; Feedback Loop", "PyTorch vinai/phobert-base đa nhiệm + bảng PhanHoiAI"],
        ["6. Phát hiện bất thường", "Phát hiện trùng lặp, lệch lĩnh vực, bất thường tác giả/đơn vị, tạp chí ngoài danh mục", "Kết hợp NetworkX + Isolation Forest + PyTorch Autoencoder + Hàng đợi ưu tiên", "Điểm rủi ro liên tục R_i in [0, 1] + chức năng Giải trừ cảnh báo"],
    ]
    add_table_custom(doc, "Bảng 1.2: Đối chiếu chi tiết yêu cầu hệ thống theo Đề cương CNTT-KLCN001 và giải pháp triển khai",
                     t2_headers, t2_data, col_widths=[2.7, 4.1, 4.1, 4.1])

    add_body(doc, "Trên cơ sở đó, hệ thống vận hành thống nhất 06 nhóm chức năng chính:")
    add_bullet_1(doc, "Chức năng 1 – Quản lý hồ sơ KH&CN động: Nhập liệu 13 trường chuẩn cho Bài báo, Đề tài và Sáng chế; đính kèm tệp minh chứng; gợi ý nhãn Lĩnh vực và Nhóm loại công trình trực tiếp bằng AI trước khi nộp.")
    add_bullet_1(doc, "Chức năng 2 – Quản lý 4 Danh mục chuẩn hóa: Thêm mới, chỉnh sửa trực tiếp và vô hiệu hóa/kích hoạt lại Lĩnh vực nghiên cứu, Loại công trình, Tạp chí/Hội thảo và Danh mục Tác giả.")
    add_bullet_1(doc, "Chức năng 3 – Quy trình 5 bước xét duyệt tính hợp lệ: Tự động chuyển 'Đang xét duyệt' khi cán bộ tiếp nhận; thẩm định kèm gợi ý/cảnh báo AI; ra quyết định 3 nhánh; ghi lịch sử bất biến và cho phép tác giả bổ sung nộp lại.")
    add_bullet_1(doc, "Chức năng 4 – Báo cáo thống kê & Tra cứu đa tiêu chí: Lọc theo từ khóa, tên tác giả, đơn vị chủ trì, lĩnh vực, nhóm công trình, trạng thái, năm công bố, cờ cảnh báo; xuất báo cáo Excel (.xlsx) và PDF.")
    add_bullet_1(doc, "Chức năng 5 – Phân loại tự động đa nhiệm & Feedback Loop: Tinh chỉnh PyTorch vinai/phobert-base dự đoán đồng thời Lĩnh vực (6 lớp) và Nhóm loại công trình (3 nhóm) kèm độ tin cậy %, cho phép cán bộ xác nhận hoặc điều chỉnh để học tăng cường.")
    add_bullet_1(doc, "Chức năng 6 – Phát hiện bất thường & Hàng đợi ưu tiên: Kết hợp NetworkX, Isolation Forest, PyTorch Autoencoder và Cosine TF-IDF để phát hiện 4 nhóm bất thường và sắp xếp hàng đợi theo điểm rủi ro giảm dần.")

    add_figure_custom(doc, "images/hinh1_quytrinh_5buoc.png",
                      "Hình 1.1: Sơ đồ quy trình 5 bước kiểm tra tính hợp lệ & thẩm định hồ sơ KH&CN", width_cm=14.5)

    add_h1(doc, "1.4. Yêu cầu phi chức năng của hệ thống")
    add_body(doc, "Bên cạnh các yêu cầu nghiệp vụ, hệ thống đáp ứng các tiêu chuẩn kỹ thuật thực tế:")
    add_bullet_1(doc, "Tính bảo mật và phân quyền chặt chẽ (Security & RBAC): Phân quyền 3 vai trò độc lập (Admin, CanBo, TacGia) qua decorator kiểm soát phiên đăng nhập; mật khẩu băm bằng thuật toán bảo mật của Werkzeug; truy vấn tham số hóa chống SQL Injection.")
    add_bullet_1(doc, f"Hiệu năng và thời gian đáp ứng thực tế trên CPU phổ thông: Các trang tra cứu và quản trị danh mục phản hồi dưới 100ms; suy luận phân loại trực tuyến bằng TF-IDF + LinearSVC đạt trung bình {svm_m['infer_time_ms']} ms/hồ sơ và bằng mô hình học sâu PyTorch PhoBERT đạt trung bình {pho_m['infer_time_ms']} ms/hồ sơ (~0.24 giây/hồ sơ trên CPU), nhờ cơ chế nạp sẵn trọng số checkpoint khi khởi động.")
    add_bullet_1(doc, "Tính toàn vẹn và bất biến dữ liệu (Data Integrity & Auditability): Ràng buộc khóa ngoại (PRAGMA foreign_keys = ON) và ràng buộc miền giá trị (CHECK Constraints) được thực thi trực tiếp trên SQLite 3; bảng LichSuXetDuyet chỉ cho phép ghi nối thêm (Append-only).")
    add_bullet_1(doc, "Tính minh bạch và khả năng tái lập thực nghiệm (Reproducibility): Cố định hạt giống ngẫu nhiên (Random Seed = 42), tách biệt tập kiểm thử độc lập theo chủ đề và lưu trữ công khai kết quả đối sánh định dạng JSON/Excel.")

    doc.add_page_break()

    # ==================== CHƯƠNG 2 ====================
    add_chapter_title(doc, "CHƯƠNG 2: CƠ SỞ LÝ THUYẾT VÀ MÔ HÌNH THUẬT TOÁN AI")

    add_h1(doc, "2.1. Quy trình Tiền xử lý văn bản tiếng Việt học thuật")
    add_body(doc, "Văn bản hồ sơ KH&CN tiếng Việt chứa tỷ lệ lớn từ ghép chuyên môn (2–4 âm tiết), tên riêng tác giả/đơn vị và các ký hiệu kỹ thuật quốc tế. Để đảm bảo tính công bằng tuyệt đối khi đối sánh các mô hình phân loại, hệ thống xây dựng một pipeline tiền xử lý thống nhất qua các bước sau:")

    add_h2(doc, "2.1.1. Cơ chế hợp nhất đặc trưng công bằng từ bộ 13 trường dữ liệu chuẩn hóa")
    add_body(doc, "Cấu trúc hồ sơ chuẩn hóa của hệ thống gồm 13 trường thông tin: (1) Loại hình ('Bài báo khoa học', 'Đề tài NCKH', 'Sáng chế & Giải pháp hữu ích'); (2) Tiêu đề công trình / Tên đề tài; (3) Tóm tắt nội dung; (4) Từ khóa; (5) Thông tin tác giả; (6) Đơn vị chủ trì; (7) Năm công bố; (8) Kỳ báo cáo; (9) Lĩnh vực đăng ký; (10) Loại công trình; (11) Tên tạp chí / Hội thảo; (12) Chỉ mục tạp chí; và (13) Sản phẩm dự kiến.")
    add_body(doc, "Để tránh hiện tượng rò rỉ nhãn mục tiêu (Label Leakage) và đảm bảo so sánh công bằng 100% giữa TF-IDF + LinearSVC và Fine-tuning PhoBERT, cả hai mô hình đều nhận chính xác cùng một chuỗi văn bản đầu vào T được trích xuất từ 05 trường ngữ nghĩa theo đúng Mục 6 Chức năng 5 của Đề cương CNTT-KLCN001 (tuyệt đối không đưa cột nhãn Lĩnh vực đăng ký, Loại hình hay Loại công trình vào chuỗi đầu vào):")
    add_bullet_1(doc, "Tiêu đề công trình / Tên đề tài: Chứa mệnh đề phương pháp nghiên cứu cốt lõi, hình thức công trình và bài toán trọng tâm.")
    add_bullet_1(doc, "Tóm tắt nội dung (Abstract): Diễn giải chi tiết quy trình kỹ thuật, mô hình toán học, phạm vi triển khai và kết quả đạt được.")
    add_bullet_1(doc, "Từ khóa (Keywords): Tập hợp các thuật ngữ chuyên ngành cô đọng nhất do tác giả xác định.")
    add_bullet_1(doc, "Thông tin tác giả và Đơn vị chủ trì: Cung cấp ngữ cảnh nhóm nghiên cứu và đơn vị học thuật liên quan.")
    add_body(doc, "T = Tiêu đề ⊕ Tóm tắt nội dung ⊕ Từ khóa ⊕ Thông tin tác giả ⊕ Đơn vị chủ trì", bold_prefix="Biểu thức hợp nhất đầu vào công bằng (combine_fair_input_text): ")

    add_h2(doc, "2.1.2. Chuẩn hóa ký tự và tách từ ghép ngữ nghĩa bằng PyVi")
    add_body(doc, "Trong tiếng Việt, ranh giới từ vựng không đồng nhất với khoảng trắng giữa các âm tiết. Quy trình chuẩn hóa văn bản được cài đặt tại lớp KHCNAIEngine gồm:")
    add_bullet_1(doc, "Chuẩn hóa mã dựng sẵn Unicode NFC, chuyển chữ thường, loại bỏ thẻ HTML, đường dẫn URL, địa chỉ email, chuỗi số rời rạc và ký tự đặc biệt.")
    add_bullet_1(doc, "Sử dụng thư viện PyVi (ViTokenizer.tokenize) dựa trên mô hình Conditional Random Fields (CRF) để nối các âm tiết tạo thành từ ghép có nghĩa bằng dấu gạch dưới (ví dụ: 'trí tuệ nhân tạo' -> 'trí_tuệ_nhân_tạo', 'giải pháp hữu ích' -> 'giải_pháp_hữu_ích').")
    add_bullet_1(doc, "Áp dụng từ điển chuẩn hóa bổ trợ gồm 39 cụm thuật ngữ chuyên ngành đại học (DOMAIN_COMPOUND_TERMS) như 'xử_lý_ngôn_ngữ_tự_nhiên', 'lưới_điện_thông_minh', 'phần_tử_hữu_hạn', 'kinh_tế_vĩ_mô', 'sở_hữu_trí_tuệ' nhằm bảo toàn chính xác các thuật ngữ dài.")

    add_h2(doc, "2.1.3. Bộ lọc từ dừng tiếng Việt học thuật (Stop-words Filtering)")
    add_body(doc, "Hệ thống loại bỏ danh sách các từ hư, liên từ và giới từ phổ biến trong văn bản tiếng Việt (như 'và', 'của', 'các', 'trong', 'được', 'là', 'nhằm', 'bởi', 'qua'...) có độ dài không mang thông tin phân loại chuyên ngành, giúp giảm nhiễu không gian đặc trưng và rút ngắn chiều dài chuỗi token đầu vào.")

    add_h1(doc, "2.2. Thuật toán Phân loại văn bản tự động đa nhiệm (Chức năng 5)")
    add_body(doc, "Hệ thống cài đặt độc lập 02 mô hình phân loại đa nhiệm (Multi-target Classification) nhằm dự đoán đồng thời (a) Lĩnh vực nghiên cứu (6 chuyên ngành: CNTT & Truyền thông; Kỹ thuật Điện – Điện tử & Tự động hóa; Kỹ thuật Cơ khí, Chế tạo máy & Giao thông; Kinh tế, Quản trị kinh doanh & Tài chính; Hóa học, Sinh học, Y Dược & Môi trường; Khoa học Xã hội, Nhân văn & Giáo dục học) và (b) Nhóm loại công trình (3 nhóm: Bài báo khoa học; Đề tài NCKH; Sáng chế & Giải pháp hữu ích):")

    add_h2(doc, "2.2.1. Mô hình Baseline: TF-IDF kết hợp LinearSVC và Platt Scaling")
    add_body(doc, "Mô hình học máy cơ sở biểu diễn văn bản theo mô hình túi từ N-gram kết hợp máy vectơ hỗ trợ tuyến tính:")
    add_bullet_1(doc, "Trích xuất đặc trưng TF-IDF: Sử dụng TfidfVectorizer với ngram_range=(1, 2), max_features=5000 và nén tần suất logarit sublinear_tf=True:")
    add_body(doc, "TF-IDF(t, d, D) = (1 + log(tf(t, d))) × [ log((1 + |D|) / (1 + df(t, D))) + 1 ]", bold_prefix="Công thức trọng số TF-IDF: ")
    add_bullet_1(doc, "Phân lớp LinearSVC kèm hiệu chuẩn xác suất Platt Scaling: Huấn luyện hai bộ phân lớp LinearSVC (C = 1.0, class_weight='balanced') riêng cho nhãn Lĩnh vực và nhãn Nhóm loại công trình, bọc trong CalibratedClassifierCV (3-fold cross-validation) để ánh xạ khoảng cách siêu phẳng f(x) = w^T x + b sang phân bố xác suất hậu nghiệm:")
    add_body(doc, "min_{w, b} [ (1/2) ||w||² + C ∑_{i=1}^N max(0, 1 - y_i(w^T x_i + b)) ]", bold_prefix="Hàm mục tiêu Hinge Loss điều chuẩn L2: ")
    add_body(doc, "P(y = 1 | f(x)) = 1 / [ 1 + exp(A × f(x) + B) ]", bold_prefix="Xác suất hiệu chuẩn Platt Scaling: ")

    add_h2(doc, "2.2.2. Mô hình Học sâu: PyTorch VinAI PhoBERT đa nhiệm (Multi-task Head)")
    add_body(doc, "Để khắc phục hạn chế mất thứ tự từ và ngữ cảnh câu phức của mô hình túi từ TF-IDF trên các hồ sơ giao thoa liên ngành, hệ thống xây dựng lớp mạng PhoBERTForKHCNClassification(nn.Module) trên nền tảng PyTorch kết nối trực tiếp mô hình tiền huấn luyện 'vinai/phobert-base' từ HuggingFace Transformers:")
    add_bullet_1(doc, "Kiến trúc Transformer gốc: Mô hình vinai/phobert-base gồm 12 tầng RoBERTa Encoder, kích thước ẩn d_model = 768, 12 đầu Self-Attention trên mỗi tầng, được tiền huấn luyện trên 20GB văn bản tiếng Việt đã tách từ.")
    add_bullet_1(doc, "Biểu diễn ngữ cảnh toàn cục: Chuỗi văn bản đã tách từ PyVi được mã hóa subword BPE với độ dài tối đa max_length = 128 tokens. Trạng thái ẩn cuối cùng tại vị trí token đầu chuỗi h_[CLS] ∈ R^{768} đại diện cho ngữ nghĩa tổng hợp của toàn bộ hồ sơ.")
    add_bullet_1(doc, "Kiến trúc Hai Nhánh Phân Loại Đa Nhiệm (Multi-task Classification Heads): Từ vector h_[CLS], mạng rẽ nhánh đồng thời sang hai đầu phân loại phi tuyến:")
    add_body(doc, "z_{field} = W_{f2} × Dropout( GELU( W_{f1} × Dropout(h_[CLS], 0.2) + b_{f1} ), 0.1 ) + b_{f2}  ∈ R^6", bold_prefix="Nhánh phân loại Lĩnh vực (field_head, 768 -> 256 -> 6): ")
    add_body(doc, "z_{type} = W_{t2} × GELU( W_{t1} × Dropout(h_[CLS], 0.2) + b_{t1} ) + b_{t2}  ∈ R^3", bold_prefix="Nhánh phân loại Loại công trình (work_type_head, 768 -> 128 -> 3): ")
    add_bullet_1(doc, "Hàm mất mát đa nhiệm tổng hợp (Joint Cross-Entropy Loss): Quá trình lan truyền ngược tối ưu đồng thời cả hai mục tiêu phân loại:")
    add_body(doc, "L_{total} = L_{CE}(z_{field}, y_{field}) + 0.5 × L_{CE}(z_{type}, y_{type})", bold_prefix="Hàm mục tiêu huấn luyện đa nhiệm: ")
    add_bullet_1(doc, "Chiến lược Đóng băng tầng (Layer Freezing) & Lưu trữ trọng số gọn nhẹ: Hệ thống đóng băng tầng nhúng (embeddings) và 10 tầng Encoder phía dưới (encoder.layer.0 đến encoder.layer.9), chỉ mở khóa cập nhật gradient trên 2 tầng biến đổi ngữ nghĩa bậc cao trên cùng (encoder.layer.10, encoder.layer.11) cùng hai nhánh phân loại bằng bộ tối ưu AdamW với tốc độ học phân tầng (lr_backbone = 4e-5, lr_head = 3e-4). Các tham số tinh chỉnh được lưu trữ thành tệp checkpoint thực tế models/phobert_khcn/phobert_classifier.pt (~55.22 MB) giúp ứng dụng web nạp lại dưới 1 giây khi khởi động.")

    add_figure_custom(doc, "images/hinh2_kientruc_ai_pipeline.png",
                      "Hình 2.1: Kiến trúc pipeline tiền xử lý công bằng và đối sánh mô hình (PhoBERT vs LinearSVC)", width_cm=14.5)

    add_body(doc, "Bảng 2.1 tổng hợp cấu hình siêu tham số thực tế đã cài đặt trong module phobert_finetune.py:")

    t3_headers = ["Siêu tham số / Cấu hình", "Giá trị thiết lập thực tế", "Ý nghĩa và mục đích kỹ thuật"]
    t3_data = [
        ["Mô hình gốc (Pre-trained)", "vinai/phobert-base (PyTorch)", "Mô hình RoBERTa tiếng Việt 12 tầng, hidden_size = 768"],
        ["Độ dài chuỗi tối đa (Max Length)", "128 subword tokens", "Bao phủ trọn vẹn tiêu đề, tóm tắt, từ khóa và tác giả đã lọc từ dừng"],
        ["Đóng băng tầng (Frozen Layers)", "Embeddings + encoder.layer.0..9", "Giữ tri thức ngôn ngữ nền tảng, chống quá khớp và tăng tốc trên CPU"],
        ["Tầng mở khóa tinh chỉnh", "encoder.layer.10..11 + 2 Heads", "Học chuyên sâu đặc trưng ngữ cảnh học thuật và hình thức công trình"],
        ["Nhánh phân loại đa nhiệm", "Field (768->256->6) & Type (768->128->3)", "Dự đoán đồng thời 6 Lĩnh vực và 3 Nhóm loại công trình"],
        ["Kích thước lô & Chu kỳ học", "Batch Size = 16 | Epochs = 8", "Hội tụ hàm mất mát mượt từ 2.0455 xuống 0.0016 sau 8 chu kỳ"],
        ["Bộ tối ưu & Tốc độ học", "AdamW (lr_backbone=4e-5, lr_head=3e-4)", "Tốc độ học phân tầng kết hợp cắt chuẩn gradient (max_norm=1.0)"],
        ["Lưu trữ trọng số (Checkpoint)", "phobert_classifier.pt (55.22 MB)", "Chỉ lưu phần trọng số tinh chỉnh giúp nạp nhanh khi chạy Web"],
    ]
    add_table_custom(doc, "Bảng 2.1: Cấu hình siêu tham số thực tế huấn luyện mô hình PyTorch VinAI PhoBERT",
                     t3_headers, t3_data, col_widths=[4.0, 4.5, 6.0])

    add_h1(doc, "2.3. Thuật toán Phát hiện Bất thường và Đánh giá Rủi ro (Chức năng 6)")
    add_body(doc, "Thay vì xét duyệt tuần tự theo thời gian nộp (FIFO), hệ thống xây dựng phân hệ phát hiện bất thường đa tầng kết hợp cấu trúc đồ thị, học máy không giám sát và luật kiểm tra ngữ nghĩa:")

    add_h2(doc, "2.3.1. Đồ thị tri thức NetworkX và phân tích độ lệch chuyên môn")
    add_body(doc, "Hệ thống khởi tạo Đồ thị Tri thức Đa phân vô hướng G = (V, E) bằng thư viện NetworkX mô hình hóa mối liên hệ giữa 04 loại đỉnh: Hồ sơ công trình (HS), Tác giả (TG), Lĩnh vực chuyên ngành (LV) và Tạp chí/Hội thảo (TC), với các cạnh quan hệ 'la_tac_gia', 'dong_tac_gia', 'thuoc_linh_vuc' và 'cong_bo_tai'.")
    add_body(doc, "C_D(v) = deg(v) / (|V| - 1)", bold_prefix="Độ trung tâm bậc liên kết (Degree Centrality): ")
    add_body(doc, "Đồng thời, hệ thống tính Độ lệch chuyên môn tác giả (graph_anomaly_feat = 1 - Tỷ lệ công trình thuộc lĩnh vực đang khai của tác giả chính khi tác giả có từ 2 công trình trở lên). Nếu một tác giả vốn chuyên sâu ở một lĩnh vực nhưng đột ngột đứng tên độc lập ở lĩnh vực xa lạ không có đồng tác giả liên ngành (độ lệch ≥ 0.55), hệ thống sẽ sinh cảnh báo dị biệt cấu trúc đồ thị.")

    add_h2(doc, "2.3.2. Thuật toán Rừng cô lập (Isolation Forest)")
    add_body(doc, "Trên không gian đặc trưng 6 chiều đã chuẩn hóa MinMaxScaler, mô hình IsolationForest (scikit-learn, n_estimators = 100, contamination = 0.25) phân hoạch ngẫu nhiên không gian bằng rừng cây nhị phân và đo độ sâu đường đi trung bình E(h(x)):")
    add_body(doc, "s(x, n) = 2^{ - E(h(x)) / c(n) }", bold_prefix="Điểm số cô lập Anomaly Score: ")
    add_body(doc, "Điểm cô lập thô được chuẩn hóa tuyến tính theo dải cực trị thực tế của tập hồ sơ về đoạn Norm(s_{IF}) ∈ [0.0, 1.0].")

    add_h2(doc, "2.3.3. Thuật toán Mạng Nơ-ron Tự mã hóa PyTorch Autoencoder")
    add_body(doc, "Song song với Isolation Forest, hệ thống cài đặt mạng nơ-ron tự mã hóa thực tế KHCNAnomalyAutoencoder(nn.Module) bằng thư viện PyTorch nhằm nắm bắt các tương quan phi tuyến giữa 6 chiều đặc trưng rủi ro:")
    add_bullet_1(doc, "Kiến trúc nút thắt cổ chai (Bottleneck Architecture): Bộ mã hóa Encoder gồm Linear(6, 16) -> ReLU -> Linear(16, 4) -> ReLU nén dữ liệu xuống không gian tiềm ẩn 4 chiều; bộ giải mã Decoder gồm Linear(4, 16) -> ReLU -> Linear(16, 6) -> Sigmoid tái tạo lại vector đầu vào.")
    add_bullet_1(doc, "Huấn luyện không giám sát trên PyTorch: Mạng được huấn luyện trực tiếp 250 chu kỳ bằng bộ tối ưu torch.optim.Adam(lr = 0.015, weight_decay = 1e-4) tối thiểu hóa hàm mất mát sai số bình phương trung bình nn.MSELoss(). Hồ sơ có tổ hợp thuộc tính mâu thuẫn sẽ sinh sai số tái tạo cao:")
    add_body(doc, "MSE_i = (1/6) ∑_{k=1}^6 ( x_{ik} - x̂_{ik} )²", bold_prefix="Sai số tái tạo từng hồ sơ: ")

    add_figure_custom(doc, "images/hinh2_phat_hien_bat_thuong.png",
                      "Hình 2.2: Kiến trúc mô hình phát hiện bất thường đa tầng và hàng đợi thẩm định ưu tiên", width_cm=14.5)

    add_h2(doc, "2.3.4. Thuật toán Kiểm tra trùng lặp nội dung Cosine TF-IDF")
    add_body(doc, "Hệ thống tính ma trận tương đồng Cosine trên vector TF-IDF của văn bản đã tách từ giữa hồ sơ đang xét với các hồ sơ đã duyệt hoặc nộp trước đó:")
    add_body(doc, "Sim(d_i, d_j) = ( v_i · v_j ) / ( ||v_i||_2 × ||v_j||_2 )", bold_prefix="Độ tương đồng Cosine: ")
    add_body(doc, "Khi Sim(d_i, d_j) ≥ 62.0%, hệ thống tự động gắn cờ nghi vấn trùng lặp nội dung và trích dẫn trực tiếp mã hồ sơ gốc để cán bộ đối chiếu.")

    add_h2(doc, "2.3.5. Điểm Rủi ro Tổng hợp liên tục và Hàng đợi thẩm định ưu tiên")
    add_body(doc, "Mỗi hồ sơ được biểu diễn bằng vector 6 chiều: (1) Độ lệch nhãn lĩnh vực giữa tác giả khai và PhoBERT dự đoán; (2) Độ bất định dự đoán của PhoBERT; (3) Độ tương đồng Cosine cực đại; (4) Rủi ro tạp chí không uy tín; (5) Độ lệch chuyên môn tác giả trên đồ thị; và (6) Độ cô lập bậc liên kết Degree Centrality.")
    add_body(doc, "Điểm Rủi ro Tổng hợp R_i ∈ [0.0, 1.0] được tính bằng tổ hợp tuyến tính liên tục (không gán ép bậc thang nhân tạo) từ 3 nguồn độc lập:")
    add_body(doc, "R_i = min(0.99,  0.35 × Norm(s_{IF, i}) + 0.35 × Norm(MSE_{AE, i}) + 0.30 × RuleSeverity_i)", bold_prefix="Công thức Điểm rủi ro tổng hợp liên tục: ")
    add_body(doc, "Hồ sơ được bật cờ cảnh báo (coCanhBao = 1) khi phát sinh ít nhất một lý do vi phạm cụ thể hoặc R_i ≥ 0.50. Toàn bộ hồ sơ đang chờ xử lý trong Hàng đợi Thẩm định Ưu tiên (Priority Review Queue) được sắp xếp giảm dần theo R_i để cán bộ rà soát trước.")

    t4_headers = ["Tiêu chí kỹ thuật", "Rừng cô lập (Isolation Forest)", "PyTorch Autoencoder (nn.Module)"]
    t4_data = [
        ["Thư viện & Nguyên lý", "scikit-learn; rừng cây nhị phân phân hoạch ngẫu nhiên", "PyTorch (torch.nn); mạng tự mã hóa nén nút thắt cổ chai"],
        ["Cấu hình tham số thực tế", "n_estimators=100, contamination=0.25, seed=42", "Layers: [6->16->4->16->6], ReLU/Sigmoid, Adam(lr=0.015), 250 epochs"],
        ["Đại lượng đo lường", "Độ sâu đường đi cô lập trung bình E(h(x))", "Sai số tái tạo bình phương trung bình MSELoss"],
        ["Thế mạnh phát hiện", "Nhạy bén với điểm ngoại lai tách biệt hình học", "Bắt tốt tương quan phi tuyến giữa lệch nhãn AI, tạp chí và đồ thị"],
        ["Trọng số kết hợp Ensemble", "35% trong Điểm rủi ro tổng hợp R_i", "35% trong Điểm rủi ro tổng hợp R_i (+ 30% Chỉ báo quy tắc/đồ thị)"],
    ]
    add_table_custom(doc, "Bảng 2.2: So sánh đặc tính giải thuật giữa Rừng cô lập (Isolation Forest) và PyTorch Autoencoder",
                     t4_headers, t4_data, col_widths=[3.5, 5.2, 5.8])

    doc.add_page_break()

    # ==================== CHƯƠNG 3 ====================
    add_chapter_title(doc, "CHƯƠNG 3: PHÂN TÍCH VÀ THIẾT KẾ HỆ THỐNG")

    add_h1(doc, "3.1. Thiết kế Kiến trúc hệ thống tổng thể 3 tầng")
    add_body(doc, "Hệ thống được thiết kế theo kiến trúc 3 tầng (3-Tier MVC Architecture) đảm bảo tính tách biệt rõ ràng giữa giao diện, nghiệp vụ AI và lưu trữ dữ liệu:")
    add_bullet_1(doc, "Tầng 1 – Tầng Trình diễn (Presentation Layer): Xây dựng trên HTML5, Bootstrap 5 và Chart.js với 8 biểu mẫu Jinja2 tinh gọn (dashboard.html, hoso_list.html, hoso_form.html, hoso_detail.html, catalogs.html, ai_center.html, users.html, login.html), tích hợp bộ chuyển mạch động 3 nhóm công trình và Live AI Preview qua Fetch/JSON.")
    add_bullet_1(doc, "Tầng 2 – Tầng Nghiệp vụ & AI Engine (Business Logic Layer): Xây dựng trên Python Flask (app.py) phối hợp cùng các module lõi phobert_finetune.py, model_benchmark.py, ai_engine.py và reporting_service.py. Tầng này thực thi phân quyền RBAC, tiền xử lý PyVi, suy luận đa nhiệm PyTorch PhoBERT & LinearSVC, và phát hiện bất thường NetworkX + Isolation Forest + PyTorch Autoencoder.")
    add_bullet_1(doc, "Tầng 3 – Tầng Lưu trữ Dữ liệu & Checkpoint (Data Persistence Layer): Sử dụng hệ quản trị cơ sở dữ liệu quan hệ SQLite 3 (tệp khcn_system.db, đúng theo Mục 7 Đề cương CNTT-KLCN001) kết hợp lưu trữ trọng số mô hình (models/phobert_khcn/phobert_classifier.pt, models/tfidf_svm/svm_pipeline.joblib) và thư mục tệp minh chứng uploads/.")

    add_figure_custom(doc, "images/hinh3_kientruc_he_thong.png",
                      "Hình 3.1: Sơ đồ kiến trúc tổng thể hệ thống 3 tầng (3-Tier MVC Architecture)", width_cm=14.5)

    add_h1(doc, "3.2. Thiết kế Ca sử dụng (Use Case) chi tiết & Ma trận RBAC")
    add_body(doc, "Hệ thống thiết lập cơ chế phân quyền dựa trên vai trò (Role-Based Access Control – RBAC) cho 03 nhóm người dùng độc lập trong bảng NguoiDung:")
    add_bullet_1(doc, "Tác nhân 1 – Quản trị viên hệ thống (Admin, tài khoản mặc định: admin@uni.edu.vn): Có toàn quyền quản trị tài khoản RBAC, thêm/sửa/ẩn/hiện cả 4 danh mục chuẩn hóa, kích hoạt thực nghiệm đối sánh mô hình AI, xóa cứng dữ liệu rác và xuất báo cáo toàn trường.")
    add_bullet_1(doc, "Tác nhân 2 – Cán bộ xét duyệt / Hội đồng KH&CN (CanBo, tài khoản mặc định: canbo@uni.edu.vn, hoilv@uni.edu.vn): Truy cập Hàng đợi xét duyệt ưu tiên, tiếp nhận thẩm định hồ sơ theo 5 bước, ra quyết định 3 nhánh (Đã duyệt / Từ chối / Yêu cầu bổ sung), xác nhận hoặc hiệu chỉnh nhãn AI (Feedback Loop) và giải trừ cảnh báo báo động giả.")
    add_bullet_1(doc, "Tác nhân 3 – Giảng viên / Tác giả nghiên cứu (TacGia, tài khoản mặc định: khoapv@uni.edu.vn cùng các giảng viên chuyên môn): Kê khai nộp mới hồ sơ 3 nhóm công trình, sử dụng gợi ý Live AI Preview, chỉnh sửa nộp lại khi có 'Yêu cầu bổ sung' và xóa mềm hồ sơ khi còn ở trạng thái 'Chờ xét duyệt'.")

    t5_headers = ["Nhóm Chức năng Nghiệp vụ", "Tác giả (TacGia)", "Cán bộ (CanBo)", "Quản trị viên (Admin)"]
    t5_data = [
        ["1. Kê khai & Nộp hồ sơ 3 nhóm (Live AI Preview)", "Tạo mới / Sửa hồ sơ cá nhân", "Xem chi tiết & minh chứng", "Toàn quyền quản trị"],
        ["2. Chỉnh sửa & nộp lại hồ sơ 'Yêu cầu bổ sung'", "Được phép (Tự về Chờ duyệt)", "Ra yêu cầu bổ sung", "Toàn quyền can thiệp"],
        ["3. Xóa mềm (Soft Delete) / Xóa cứng (Hard Delete)", "Chỉ xóa mềm khi Chờ duyệt", "Không có quyền xóa", "Toàn quyền Xóa mềm & Cứng"],
        ["4. Hàng đợi ưu tiên & Xét duyệt 3 nhánh", "Không có quyền", "Toàn quyền thẩm định", "Toàn quyền thẩm định"],
        ["5. Xác nhận / Điều chỉnh nhãn AI & Giải trừ cảnh báo", "Không có quyền", "Được phép (Ghi vào PhanHoiAI)", "Được phép (Ghi vào PhanHoiAI)"],
        ["6. Quản lý 4 Danh mục chuẩn hóa (Lĩnh vực, Loại, TC, TG)", "Chỉ xem danh mục", "Chỉ xem danh mục", "Toàn quyền Thêm / Sửa / Ẩn"],
        ["7. Chạy thực nghiệm đối sánh SVM & PhoBERT", "Chỉ xem kết quả", "Được phép kích hoạt", "Toàn quyền kích hoạt"],
        ["8. Tra cứu đa tiêu chí & Xuất báo cáo Excel / PDF", "Tra cứu & Xuất báo cáo", "Tra cứu & Xuất báo cáo", "Toàn quyền toàn hệ thống"],
    ]
    add_table_custom(doc, "Bảng 3.1: Ma trận phân quyền kiểm soát truy cập dựa trên vai trò (RBAC)",
                     t5_headers, t5_data, col_widths=[5.0, 3.2, 3.5, 3.2])

    add_h1(doc, "3.3. Thiết kế Cơ sở dữ liệu quan hệ SQLite 3 (khcn_system.db)")
    add_body(doc, "Cơ sở dữ liệu SQLite 3 (tệp khcn_system.db) gồm 09 bảng quan hệ chuẩn hóa 3NF có bật kiểm tra toàn vẹn khóa ngoại (PRAGMA foreign_keys = ON): NguoiDung, LinhVuc, LoaiCongTrinh, TapChiHopLe, DanhMucTacGia, HoSoKHCN, LichSuXetDuyet, PhanHoiAI và ThucNghiemMoHinhAI. Bảng 3.2 trình bày lược đồ thực tế của bảng trung tâm HoSoKHCN:")

    t6_headers = ["Tên cột thực tế", "Kiểu SQLite", "Ràng buộc (Constraints)", "Ý nghĩa nghiệp vụ và liên kết AI"]
    t6_data = [
        ["maHoSo", "TEXT", "PRIMARY KEY", "Mã hồ sơ công trình KH&CN (ví dụ: HS-2026-0001)"],
        ["tieuDe", "TEXT", "NOT NULL", "Tiêu đề bài báo, tên đề tài NCKH hoặc tên sáng chế/GPHI"],
        ["tomTat", "TEXT", "NOT NULL", "Tóm tắt nội dung, phương pháp và kết quả nghiên cứu"],
        ["tuKhoa", "TEXT", "NOT NULL", "Danh sách từ khóa chuyên ngành phục vụ tra cứu và phân loại"],
        ["thongTinTacGia", "TEXT", "NOT NULL", "Danh sách tác giả và đồng tác giả thực hiện công trình"],
        ["donViChuTri", "TEXT", "NOT NULL", "Khoa / Viện chuyên môn chủ trì công trình"],
        ["namCongBo, kyBaoCao", "INTEGER, TEXT", "NOT NULL", "Năm công bố và đợt/kỳ báo cáo thống kê"],
        ["maLinhVuc, maLoai", "INTEGER", "FOREIGN KEY", "Lĩnh vực đăng ký (LinhVuc) và Loại công trình (LoaiCongTrinh)"],
        ["maTapChi", "INTEGER", "FOREIGN KEY (NULL)", "Tạp chí / Hội thảo công bố (liên kết TapChiHopLe)"],
        ["chiMucTapChi, sanPhamDuKien", "TEXT", "DEFAULT ''", "Chỉ mục tạp chí (Scopus/HĐGSNN) và Sản phẩm dự kiến / Bảo hộ SHTT"],
        ["vanBanDaTachTu", "TEXT", "AI Preprocessing", "Chuỗi văn bản hợp nhất sau khi tách từ ghép bằng PyVi"],
        ["maLinhVucAI, doTinCayAI", "INTEGER, REAL", "FK -> LinhVuc", "Lĩnh vực và độ tin cậy (%) do PyTorch PhoBERT dự đoán"],
        ["nhomCongTrinhAI, doTinCayLoaiAI", "TEXT, REAL", "Multi-task Output", "Nhóm loại công trình và độ tin cậy (%) do PyTorch PhoBERT dự đoán"],
        ["maLinhVucSVM, doTinCaySVM", "INTEGER, REAL", "Baseline Output", "Kết quả dự đoán đối chứng của mô hình TF-IDF + LinearSVC"],
        ["trangThai", "TEXT", "CHECK (5 trạng thái)", "Chờ xét duyệt, Đang xét duyệt, Đã duyệt, Từ chối, Yêu cầu bổ sung"],
        ["diemBatThuong, coCanhBao", "REAL, INTEGER", "DEFAULT 0.0, 0", "Điểm rủi ro tổng hợp liên tục [0, 1] và cờ cảnh báo bất thường"],
        ["chiTietCanhBao", "TEXT (JSON)", "JSON Payload", "Lưu danh sách lý do cảnh báo, điểm IF, sai số PyTorch AE, Cosine"],
    ]
    add_table_custom(doc, "Bảng 3.2: Lược đồ thuộc tính bảng dữ liệu cốt lõi HoSoKHCN trong SQLite 3 (khcn_system.db)",
                     t6_headers, t6_data, col_widths=[3.5, 2.3, 3.2, 6.0])

    add_body(doc, "Bảng lưu vết kiểm toán bất biến LichSuXetDuyet ghi nhận toàn bộ lịch sử chuyển trạng thái và nhận xét chuyên môn của cán bộ xét duyệt:")

    t7_headers = ["Tên cột thực tế", "Kiểu SQLite", "Ràng buộc (Constraints)", "Ý nghĩa quản trị và lưu vết kiểm toán"]
    t7_data = [
        ["maLichSu", "INTEGER", "PRIMARY KEY AUTOINCREMENT", "Khóa chính tự tăng định danh bản ghi lịch sử"],
        ["maHoSo", "TEXT", "FOREIGN KEY -> HoSoKHCN", "Mã hồ sơ công trình được thẩm định hoặc giải trừ cảnh báo"],
        ["maCanBo", "INTEGER", "FOREIGN KEY -> NguoiDung", "Tài khoản Cán bộ xét duyệt hoặc Quản trị viên thực hiện"],
        ["trangThaiCu", "TEXT", "NOT NULL", "Trạng thái hồ sơ trước khi thực hiện thao tác"],
        ["trangThaiMoi", "TEXT", "NOT NULL", "Trạng thái mới xác lập sau quyết định thẩm định"],
        ["yKienDanhGia", "TEXT", "NOT NULL", "Nhận xét thẩm định, lý do từ chối hoặc nội dung yêu cầu bổ sung"],
        ["thoiGianXuLy", "DATETIME", "DEFAULT CURRENT_TIMESTAMP", "Thời điểm hệ thống ghi nhận thao tác (Append-only)"],
    ]
    add_table_custom(doc, "Bảng 3.3: Lược đồ thuộc tính bảng lưu vết kiểm toán bất biến LichSuXetDuyet",
                     t7_headers, t7_data, col_widths=[3.0, 2.5, 3.8, 5.7])

    add_h1(doc, "3.4. Thiết kế Sơ đồ lớp phân tích (Boundary – Control – Entity)")
    add_body(doc, "Hệ thống tổ chức theo mẫu lớp phân tích BCE (Boundary – Control – Entity):")
    add_bullet_1(doc, "Các lớp Biên (Boundary): Giao diện kê khai động hoso_form.html, danh sách & bộ lọc đa tiêu chí hoso_list.html, chi tiết xét duyệt hoso_detail.html, quản trị 4 danh mục catalogs.html, trung tâm thực nghiệm AI ai_center.html và tổng quan dashboard.html.")
    add_bullet_1(doc, "Các lớp Điều khiển (Control): Các route controller trong app.py, bộ máy trung tâm KHCNAIEngine (ai_engine.py), module tinh chỉnh PyTorch PhoBERT (phobert_finetune.py), module đối sánh thực nghiệm (model_benchmark.py) và dịch vụ xuất báo cáo (reporting_service.py).")
    add_bullet_1(doc, "Các lớp Thực thể (Entity): 09 bảng dữ liệu trong database.py gồm NguoiDung, LinhVuc, LoaiCongTrinh, TapChiHopLe, DanhMucTacGia, HoSoKHCN, LichSuXetDuyet, PhanHoiAI và ThucNghiemMoHinhAI.")

    add_h1(doc, "3.5. Thiết kế Luồng tương tác quy trình 5 bước kiểm tra tính hợp lệ")
    add_body(doc, "Luồng tương tác xét duyệt hồ sơ diễn ra chặt chẽ qua 5 bước:")
    add_bullet_1(doc, "Bước 1 – Tiếp nhận & Tự động khóa trạng thái: Khi Cán bộ/Admin mở một hồ sơ đang 'Chờ xét duyệt' từ Hàng đợi ưu tiên, hệ thống tự động chuyển trạng thái sang 'Đang xét duyệt' và ghi nhận vào LichSuXetDuyet để tránh trùng lặp xử lý.")
    add_bullet_1(doc, "Bước 2 – Thẩm định chuyên môn & Tham vấn AI: Màn hình hiển thị thông tin công trình, tệp minh chứng, dự đoán đa nhiệm của PyTorch PhoBERT (cả Lĩnh vực và Nhóm loại công trình kèm độ tin cậy %) và chi tiết cảnh báo bất thường.")
    add_bullet_1(doc, "Bước 3 – Ra quyết định 3 nhánh: Cán bộ chọn 'Đã duyệt', 'Từ chối' hoặc 'Yêu cầu bổ sung' (bắt buộc nhập lý do đánh giá rõ ràng).")
    add_bullet_1(doc, "Bước 4 – Lưu vết lịch sử & Phản hồi nhãn AI: Mọi quyết định và thao tác xác nhận/điều chỉnh nhãn phân loại đều được ghi vĩnh viễn vào LichSuXetDuyet và PhanHoiAI.")
    add_bullet_1(doc, "Bước 5 – Bổ sung & Nộp lại (Resubmission): Khi hồ sơ ở trạng thái 'Yêu cầu bổ sung', tác giả được mở quyền chỉnh sửa và khi lưu lại, hồ sơ tự động quay về 'Chờ xét duyệt' kèm phân tích AI mới.")

    doc.add_page_break()

    # ==================== CHƯƠNG 4 ====================
    add_chapter_title(doc, "CHƯƠNG 4: CÀI ĐẶT CHI TIẾT CÁC CHỨC NĂNG HỆ THỐNG")

    add_h1(doc, "4.1. Chức năng 1 – Quản lý Hồ sơ KH&CN động (Bài báo, Đề tài, Sáng chế)")
    add_body(doc, "Biểu mẫu kê khai hồ sơ (templates/hoso_form.html) được cài đặt đầy đủ 13 trường chuẩn hóa, bao gồm trường Thông tin tác giả có gợi ý tự động (datalist) liên kết trực tiếp từ Danh mục Tác giả:")
    add_h2(doc, "4.1.1. Bộ chuyển mạch động cho cả 03 nhóm loại hình công trình")
    add_body(doc, "Khi người dùng chọn Loại hình ở mục (1), hàm JavaScript switchWorkGroup() tự động điều chỉnh danh sách Loại công trình và ẩn/hiện các trường đặc thù phù hợp:")
    add_bullet_1(doc, "Khi chọn 'Bài báo khoa học': Hiển thị trường Tên tạp chí / Hội thảo (đối chiếu bảng TapChiHopLe) và Chỉ mục tạp chí (Scopus Q1/Q2, HĐGSNN, IEEE Xplore...).")
    add_bullet_1(doc, "Khi chọn 'Đề tài NCKH': Ẩn khối tạp chí và hiển thị trường Sản phẩm dự kiến & Cấp nghiệm thu của đề tài.")
    add_bullet_1(doc, "Khi chọn 'Sáng chế / Giải pháp hữu ích': Hiển thị trường Thông tin đăng ký bảo hộ Sở hữu trí tuệ / Văn bằng độc quyền và khả năng áp dụng thực tiễn.")

    add_h2(doc, "4.1.2. Tính năng Gợi ý trực tiếp đa nhiệm (Live AI Preview)")
    add_body(doc, f"Khi giảng viên nhấn nút 'Gợi ý bằng AI (PhoBERT)', giao diện gửi yêu cầu POST JSON tới endpoint /api/ai-preview với 5 trường văn bản (Tiêu đề, Tóm tắt, Từ khóa, Thông tin tác giả, Đơn vị chủ trì). Mô hình chính PyTorch PhoBERT lập tức suy luận và trả về đồng thời nhãn Lĩnh vực gợi ý cùng Nhóm loại công trình dự đoán kèm độ tin cậy % (thời gian phản hồi trung bình ~{pho_m['infer_time_ms']} ms trên CPU), đồng thời tự động chọn sẵn lĩnh vực trên biểu mẫu.")

    add_figure_custom(doc, "images/hinh4_hoso_form_ai.jpg",
                      "Hình 4.1: Giao diện Kê khai và nộp hồ sơ động tích hợp tính năng Live AI Preview", width_cm=14.5)

    add_h1(doc, "4.2. Chức năng 2 – Quản lý 4 Danh mục chuẩn hóa KH&CN (CRUD đầy đủ)")
    add_body(doc, "Phân hệ Quản lý Danh mục (/danh-muc, tệp templates/catalogs.html) hiện thực hóa đầy đủ 04 danh mục chuẩn theo yêu cầu Đề cương CNTT-KLCN001 với trọn vẹn các thao tác Thêm mới (Create), Xem thống kê số hồ sơ liên kết (Read), Chỉnh sửa trực tiếp trên dòng (Update) và Vô hiệu hóa/Kích hoạt lại bảo toàn khóa ngoại (Soft Toggle):")
    add_bullet_1(doc, "1. Danh mục Lĩnh vực nghiên cứu (LinhVuc): Gồm 6 nhóm chuyên ngành đào tạo và nghiên cứu cốt lõi của Nhà trường.")
    add_bullet_1(doc, "2. Danh mục Loại công trình (LoaiCongTrinh): Bao phủ Bài báo quốc tế WoS/Scopus, Bài báo HĐGSNN, Đề tài cấp Bộ/Trường và Bằng độc quyền Sáng chế / Giải pháp hữu ích kèm định mức giờ chuẩn quy đổi.")
    add_bullet_1(doc, "3. Danh mục Tạp chí / Hội thảo tham chiếu (TapChiHopLe): Quản lý tên tạp chí, mã ISSN, xếp hạng và cờ uy tín (Hợp lệ = 1 / Cảnh báo không uy tín = 0) phục vụ kiểm tra chéo bất thường.")
    add_bullet_1(doc, "4. Danh mục Tác giả nghiên cứu (DanhMucTacGia): Quản lý mã tác giả, học hàm học vị, họ tên, email, đơn vị công tác, lĩnh vực chuyên môn chính và mã định danh ORCID.")

    t8_headers = ["Mã danh mục", "Tên lĩnh vực nghiên cứu trong CSDL", "Phạm vi chuyên môn & Từ khóa đặc trưng"]
    t8_data = [
        ["LV-CNTT", "Công nghệ thông tin và Truyền thông", "Trí tuệ nhân tạo, học sâu, NLP, thị giác máy tính, an toàn thông tin, dữ liệu lớn"],
        ["LV-DDT", "Kỹ thuật Điện - Điện tử và Tự động hóa", "Điều khiển tự động, lưới điện thông minh, điện tử công suất, IoT, vi điều khiển"],
        ["LV-CKD", "Kỹ thuật Cơ khí, Chế tạo máy và Giao thông", "Gia công CNC, phần tử hữu hạn FEM, cơ điện tử, robot công nghiệp, ô tô điện"],
        ["LV-KTQL", "Kinh tế, Quản trị kinh doanh và Tài chính", "Kinh tế vĩ mô, tài chính ngân hàng, chuỗi cung ứng, thương mại điện tử, kiểm toán"],
        ["LV-HSMT", "Hóa học, Sinh học, Y Dược và Môi trường", "Công nghệ sinh học, vật liệu nano, dược liệu, enzyme, xử lý nước thải và môi trường"],
        ["LV-XHNV", "Khoa học Xã hội, Nhân văn và Giáo dục học", "Phương pháp giảng dạy đại học, chương trình đào tạo CDIO/AUN-QA, ngôn ngữ, tâm lý học"],
    ]
    add_table_custom(doc, "Bảng 4.1: Danh mục 6 nhóm lĩnh vực nghiên cứu KH&CN và 4 phân hệ danh mục trong hệ thống",
                     t8_headers, t8_data, col_widths=[2.5, 5.2, 7.3])

    add_h1(doc, "4.3. Chức năng 3 – Quy trình 5 bước Kiểm tra Tính Hợp lệ & Thẩm định")
    add_body(doc, "Trang thẩm định chi tiết (/hoso/<ma_ho_so>, tệp templates/hoso_detail.html) tích hợp khép kín quy trình xét duyệt 5 bước, hiển thị rõ thông tin tác giả, đơn vị, tệp minh chứng đính kèm, khung đối chiếu phân loại AI và bảng lịch sử xét duyệt bất biến.")

    add_figure_custom(doc, "images/hinh4_hoso_detail_review.jpg",
                      "Hình 4.2: Giao diện Thẩm định chi tiết hồ sơ, lưu vết lịch sử và hiển thị cảnh báo AI", width_cm=14.5)

    add_h1(doc, "4.4. Chức năng 4 – Báo cáo Thống kê & Phân hệ Tra cứu Đa tiêu chí")
    add_body(doc, "Phân hệ danh sách hồ sơ (/hoso, tệp templates/hoso_list.html) và Dashboard (/dashboard) đáp ứng đầy đủ các tiêu chí tìm kiếm và thống kê trong Đề cương:")
    add_bullet_1(doc, "Tìm kiếm và lọc kết hợp 7 tiêu chí: Tìm kiếm từ khóa toàn văn (tiêu đề, tóm tắt, từ khóa, mã hồ sơ, tác giả) kết hợp lọc theo Tác giả, Đơn vị công tác, Lĩnh vực chuyên ngành, Loại hình công trình (Bài báo / Đề tài / Sáng chế), Trạng thái xét duyệt, Năm công bố và Hồ sơ có cảnh báo AI.")
    add_bullet_1(doc, "Kết xuất báo cáo Excel (.xlsx) và PDF (.pdf): Module reporting_service.py sử dụng OpenPyXL và ReportLab (nhúng font Unicode tiếng Việt) kết xuất danh sách hồ sơ theo bộ lọc hiện hành phục vụ lưu trữ và báo cáo tổng kết.")

    add_h1(doc, "4.5. Chức năng 5 – AI Phân loại Tự động Đa nhiệm & Feedback Loop")
    add_body(doc, "Tại trang chi tiết hồ sơ, Cán bộ xét duyệt hoặc Quản trị viên được cung cấp bảng đối chiếu kết quả dự đoán của Mô hình chính (PyTorch PhoBERT) cho cả Lĩnh vực chuyên ngành và Nhóm loại công trình:")
    add_bullet_1(doc, "Xác nhận hoặc Điều chỉnh đa mục tiêu: Cán bộ có thể nhấn 'Xác nhận kết quả AI' hoặc chọn lại Lĩnh vực thực tế và Loại công trình thực tế trên form hiệu chỉnh (/hoso/<ma_ho_so>/ai-feedback).")
    add_bullet_1(doc, "Cập nhật học tăng cường (Online Adaptation): Mỗi phản hồi được lưu vào bảng PhanHoiAI; hệ thống tự động huấn luyện bổ sung bộ phân lớp SVM và thực hiện bước cập nhật gradient nhanh trên tầng phân loại field_head của PyTorch PhoBERT đối với các mẫu được cán bộ hiệu chỉnh.")

    add_h1(doc, "4.6. Chức năng 6 – Phát hiện Bất thường & Xếp hạng Ưu tiên Hàng đợi")
    add_body(doc, "Tại Trung tâm Thực nghiệm & Kiểm định AI (/ai-center, tệp templates/ai_center.html), hệ thống tự động quét 04 nhóm dấu hiệu bất thường (khai lệch lĩnh vực so với dự đoán PhoBERT, tạp chí thuộc diện cảnh báo không uy tín, tương đồng nội dung Cosine ≥ 62.0%, và độ lệch chuyên môn tác giả trên đồ thị NetworkX ≥ 55%), hiển thị đồng thời điểm Isolation Forest, sai số MSE của PyTorch Autoencoder và xếp hạng Hàng đợi xét duyệt ưu tiên theo điểm rủi ro giảm dần.")

    add_figure_custom(doc, "images/hinh4_ai_center.jpg",
                      "Hình 4.3: Giao diện Trung tâm Quản trị AI, đồ thị tri thức và hàng đợi ưu tiên xét duyệt", width_cm=14.5)

    add_h1(doc, "4.7. Các chức năng hỗ trợ quản trị, Dashboard và đối sánh thực nghiệm")
    add_body(doc, "Trang tổng quan Dashboard cung cấp các thẻ KPI theo thời gian thực và biểu đồ thống kê theo lĩnh vực, trạng thái, đơn vị công tác. Đồng thời, trang /ai-center tích hợp công cụ chạy thực nghiệm độc lập từng mô hình ('Chạy riêng TF-IDF + SVM', 'Chạy riêng Fine-tuning PhoBERT', 'Chạy đối sánh cả 2 & Chọn mô hình chính') và cho phép tải trực tiếp tệp Excel dữ liệu chuẩn kèm bảng kết quả đối sánh.")

    add_figure_custom(doc, "images/hinh4_dashboard_overview.jpg",
                      "Hình 4.4: Giao diện Dashboard điều hành tổng quan với các thẻ KPI và biểu đồ phân loại", width_cm=14.5)

    doc.add_page_break()

    # ==================== CHƯƠNG 5 ====================
    add_chapter_title(doc, "CHƯƠNG 5: THỬ NGHIỆM, ĐÁNH GIÁ KẾT QUẢ VÀ TRIỂN KHAI")

    add_h1(doc, "5.1. Thiết lập Môi trường thử nghiệm thực tế và Bộ dữ liệu chuẩn 13 trường")
    add_body(doc, "Toàn bộ quá trình huấn luyện mô hình học máy, tinh chỉnh mô hình học sâu PyTorch PhoBERT, phát hiện bất thường và kiểm thử ứng dụng web được thực thi trực tiếp trên cấu hình phần cứng và phần mềm thực tế của nhóm nghiên cứu:")
    add_bullet_1(doc, "Hạ tầng phần cứng thực tế: Máy tính xách tay trang bị vi xử lý 12th Gen Intel(R) Core(TM) i5-12450HX (8 nhân, 12 luồng), bộ nhớ RAM 16 GB, đồ họa tích hợp Intel(R) UHD Graphics (huấn luyện và suy luận trên CPU).")
    add_bullet_1(doc, "Hạ tầng phần mềm thực tế: Hệ điều hành Windows 11 64-bit, ngôn ngữ Python 3.12, PyTorch 2.6.0+cpu, HuggingFace Transformers 5.22.1 (tải trọng số thực vinai/phobert-base), Scikit-learn, NetworkX, PyVi, Flask và hệ quản trị CSDL SQLite 3 (khcn_system.db).")
    add_body(doc, "Về dữ liệu thực nghiệm: Tuân thủ Mục 7 Đề cương CNTT-KLCN001 ('Định dạng file dữ liệu: CSV/JSON cho dữ liệu mô phỏng hoặc thực tế') và tham khảo quy cách kê khai công trình học thuật đại học trong nước, nhóm nghiên cứu đã xây dựng Bộ dữ liệu mô phỏng học thuật chuẩn hóa gồm 240 hồ sơ KH&CN tiếng Việt (lưu tại docs/Tap_Du_Lieu_Huan_Luyen_Chuan_12Truong.xlsx). Mỗi mẫu chứa đầy đủ 13 trường thông tin chuẩn: (1) Loại hình, (2) Tiêu đề công trình / Tên đề tài, (3) Tóm tắt nội dung, (4) Từ khóa, (5) Thông tin tác giả, (6) Đơn vị chủ trì, (7) Năm công bố, (8) Kỳ báo cáo, (9) Lĩnh vực đăng ký, (10) Loại công trình, (11) Tên tạp chí / Hội thảo, (12) Chỉ mục tạp chí và (13) Sản phẩm dự kiến.")

    add_h1(doc, "5.2. Thực nghiệm Đối sánh Công bằng: TF-IDF + SVM, PhoBERT và Thủ công")
    add_h2(doc, "5.2.1. Thống kê bộ dữ liệu huấn luyện và kiểm thử dùng chung (240 mẫu)")
    add_body(doc, f"Bộ dữ liệu {ds_m['total_samples']} hồ sơ được chia tách nghiêm ngặt theo chủ đề độc lập (chủ đề kiểm thử hoàn toàn không xuất hiện trong tập huấn luyện) với tỷ lệ 75% Huấn luyện ({ds_m['train_samples']} mẫu, 30 mẫu/lĩnh vực) và 25% Kiểm thử ({ds_m['test_samples']} mẫu, 10 mẫu/lĩnh vực). Về cơ cấu loại hình, bộ dữ liệu gồm {ds_m['total_bai_bao']} Bài báo khoa học, {ds_m['total_de_tai']} Đề tài NCKH và {ds_m['total_sang_che']} Sáng chế & Giải pháp hữu ích; trong đó có 144 mẫu đơn ngành (60%) và {ds_m['total_interdisciplinary']} mẫu giao thoa liên ngành (40%). Bảng 5.1 trình bày thống kê chi tiết:")

    t_ds_headers = ["Mã & Tên Lĩnh vực đăng ký", "Tổng", "Train", "Test", "Bài báo", "Đề tài", "Sáng chế", "Liên ngành"]
    t_ds_data = [
        [f"[{f['code']}] {f['name']}", str(f["total"]), str(f["train"]), str(f["test"]),
         str(f["bai_bao"]), str(f["de_tai"]), str(f["sang_che"]), str(f["interdisciplinary"])]
        for f in ds_m["by_field"]
    ]
    t_ds_data.append([
        "TỔNG CỘNG TOÀN BỘ TẬP DỮ LIỆU (13 TRƯỜNG)",
        str(ds_m["total_samples"]), str(ds_m["train_samples"]), str(ds_m["test_samples"]),
        str(ds_m["total_bai_bao"]), str(ds_m["total_de_tai"]), str(ds_m["total_sang_che"]),
        str(ds_m["total_interdisciplinary"]),
    ])
    add_table_custom(doc, "Bảng 5.1: Thống kê bộ dữ liệu mô phỏng chuẩn hóa 13 trường dùng chung (240 mẫu)",
                     t_ds_headers, t_ds_data, col_widths=[4.6, 1.3, 1.4, 1.4, 1.5, 1.5, 1.6, 1.7])

    add_h2(doc, "5.2.2. Kết quả đối sánh Precision, Recall, F1-score và Phân loại thủ công")
    add_body(doc, "Cả hai mô hình TF-IDF + LinearSVC và PyTorch Fine-tuning PhoBERT đều nhận chính xác cùng một chuỗi văn bản đầu vào từ hàm combine_fair_input_text(Tiêu đề, Tóm tắt, Từ khóa, Thông tin tác giả, Đơn vị chủ trì), huấn luyện trên 180 mẫu Train và đánh giá độc lập trên 60 mẫu Test. Bảng 5.2 tổng hợp kết quả thực nghiệm đo trực tiếp từ tệp docs/Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json:")

    t9_headers = ["Lĩnh vực / Hạng mục Đánh giá", "Chỉ số Đo lường", "TF-IDF + LinearSVC", "PyTorch PhoBERT", "Chênh lệch (Δ)"]
    t9_data = []
    for idx_pf, pf in enumerate(bench["per_field_comparison"], start=1):
        sign_str = "+" if pf["delta_f1"] > 0 else ""
        t9_data.append([
            f"{idx_pf}. [{pf['code']}] {pf['name']}",
            "Precision / Recall / F1",
            f"{pf['svm_precision']:.2f}% / {pf['svm_recall']:.2f}% / {pf['svm_f1']:.2f}%",
            f"{pf['phobert_precision']:.2f}% / {pf['phobert_recall']:.2f}% / {pf['phobert_f1']:.2f}%",
            f"{sign_str}{pf['delta_f1']:.2f}% F1",
        ])
    t9_data.extend([
        ["TOÀN TẬP TEST – LĨNH VỰC (60 MẪU)", "Macro Precision", f"{svm_m['precision']:.2f}%", f"{pho_m['precision']:.2f}%", f"+{sel_m['delta_precision']:.2f}%"],
        ["TOÀN TẬP TEST – LĨNH VỰC (60 MẪU)", "Macro Recall", f"{svm_m['recall']:.2f}%", f"{pho_m['recall']:.2f}%", f"+{sel_m['delta_recall']:.2f}%"],
        ["TOÀN TẬP TEST – LĨNH VỰC (60 MẪU)", "Macro F1-score", f"{svm_m['f1_score']:.2f}%", f"{pho_m['f1_score']:.2f}%", f"+{sel_m['delta_f1']:.2f}% (Tối ưu)"],
        ["TOÀN TẬP TEST – LĨNH VỰC (60 MẪU)", "Accuracy tổng thể", f"{svm_m['accuracy']:.2f}% (49/60)", f"{pho_m['accuracy']:.2f}% (54/60)", f"+{sel_m['delta_accuracy']:.2f}%"],
        ["PHÂN LOẠI NHÓM LOẠI CÔNG TRÌNH", "Macro F1-score (3 nhóm)", f"{svm_m['work_type_f1']:.2f}%", f"{pho_m['work_type_f1']:.2f}%", "Đều đạt tuyệt đối"],
        ["THỜI GIAN HUẤN LUYỆN THỰC TẾ (CPU)", "180 mẫu Train (8 epochs)", f"{svm_m['train_time_ms']:.2f} ms", f"{pho_m['train_time_ms']/1000.0:.2f} giây", "Lưu checkpoint .pt"],
        ["THỜI GIAN SUY LUẬN ĐƠN LẺ (CPU)", "Trung bình / 1 hồ sơ", f"{svm_m['infer_time_ms']:.2f} ms", f"{pho_m['infer_time_ms']:.2f} ms", "Phản hồi < 0.25 giây"],
    ])
    add_table_custom(doc, "Bảng 5.2: Kết quả thực nghiệm đối sánh công bằng giữa TF-IDF + LinearSVC và Fine-tuning PhoBERT",
                     t9_headers, t9_data, col_widths=[4.0, 3.0, 3.2, 3.2, 2.6])

    add_body(doc, "Thực hiện yêu cầu đánh giá Tuần 9 và Mục 5 Đề cương CNTT-KLCN001 về việc so sánh kết quả phân loại tự động với phương pháp phân loại thủ công, nhóm đã đối chiếu quy trình rà soát thủ công theo danh mục từ khóa & đơn vị trên cùng lô 60 hồ sơ kiểm thử (trong đó các hồ sơ thuộc Viện Nghiên cứu Liên ngành chứa nhiều thuật ngữ miền ứng dụng gây nhiễu khi đọc lướt). Kết quả đối sánh được trình bày tại Bảng 5.3:")

    m_b = man_m["manual_baseline"]
    m_s = man_m["tfidf_svm_auto"]
    m_p = man_m["phobert_auto"]
    t_man_headers = ["Phương pháp Phân loại", "Precision", "Recall", "F1 Lĩnh vực", "F1 Loại CT", "Thời gian / 1 hồ sơ", "Tổng lô 60 hồ sơ"]
    t_man_data = [
        ["1. Phân loại thủ công (Rà soát mắt thường)", f"{m_b['precision']:.2f}%", f"{m_b['recall']:.2f}%", f"{m_b['f1_score']:.2f}%", f"{m_b['work_type_f1']:.1f}%", f"~{m_b['avg_time_sec']:.0f} giây", f"~{m_b['batch_60_time_min']:.1f} phút"],
        ["2. Tự động: TF-IDF + LinearSVC (Baseline)", f"{m_s['precision']:.2f}%", f"{m_s['recall']:.2f}%", f"{m_s['f1_score']:.2f}%", f"{m_s['work_type_f1']:.1f}%", f"{m_s['avg_time_ms']:.2f} ms", f"{m_s['batch_60_time_sec']:.3f} giây"],
        ["3. Tự động: Fine-tuning PhoBERT (Mô hình chính)", f"{m_p['precision']:.2f}%", f"{m_p['recall']:.2f}%", f"{m_p['f1_score']:.2f}%", f"{m_p['work_type_f1']:.1f}%", f"{m_p['avg_time_ms']:.2f} ms", f"{m_p['batch_60_time_sec']:.2f} giây"],
    ]
    add_table_custom(doc, "Bảng 5.3: Đối sánh hiệu quả giữa Phân loại Tự động (AI) và Phương pháp Phân loại Thủ công",
                     t_man_headers, t_man_data, col_widths=[4.2, 1.7, 1.6, 1.9, 1.8, 2.0, 2.0])

    add_h2(doc, "5.2.3. Lý do lựa chọn Fine-tuning PhoBERT làm mô hình chính của dự án")
    add_body(doc, "Dựa trên kết quả thực nghiệm khách quan tại Bảng 5.2 và Bảng 5.3, đồ án lựa chọn mô hình Tinh chỉnh PhoBERT Tiếng Việt (PyTorch vinai/phobert-base) làm Mô hình chính (Primary Model) của toàn dự án (giữ TF-IDF + LinearSVC làm mô hình đối chứng), với 05 căn cứ thực chứng rõ ràng:")
    add_bullet_1(doc, f"Độ chính xác F1-score tổng thể vượt trội hoàn toàn công bằng: Trên cùng một chuỗi đầu vào và cùng tập Test độc lập 60 mẫu, PyTorch PhoBERT đạt Macro F1-score {pho_m['f1_score']}% (Accuracy {pho_m['accuracy']}%), vượt +{sel_m['delta_f1']}% F1-score so với TF-IDF + LinearSVC ({svm_m['f1_score']}%) và vượt +{m_p['delta_f1_vs_manual']}% F1-score so với phân loại thủ công ({m_b['f1_score']}%).", bold_prefix="Lý do 1 – Hiệu quả phân loại cao nhất: ")
    add_bullet_1(doc, "Ưu thế rõ rệt trên các lĩnh vực có hồ sơ giao thoa liên ngành phức tạp: Khi hồ sơ thuộc lĩnh vực Điện – Điện tử, Hóa – Sinh – Môi trường hoặc CNTT nhưng phần mô tả bối cảnh ứng dụng chứa nhiều từ vựng giáo dục, y tế hay quản lý, mô hình túi từ TF-IDF + SVM bị kéo lệch trọng số tuyến tính (F1 ngành CNTT chỉ đạt 72.73%, Điện – Điện tử đạt 82.35%, Hóa – Sinh đạt 80.00%). Ngược lại, cơ chế Self-Attention hai chiều của PhoBERT phân biệt đúng mệnh đề phương pháp nghiên cứu cốt lõi, nâng F1-score ngành Điện – Điện tử lên 100.0% (+17.65%), Hóa – Sinh lên 95.24% (+15.24%) và CNTT lên 85.71% (+12.98%). Mặt khác, kết quả thực nghiệm cũng phản ánh trung thực rằng ở một số mẫu Kinh tế giao thoa quản lý giáo dục (LV-KTQL), SVM đạt 88.89% so với 82.35% của PhoBERT do tập huấn luyện quy mô nhỏ (30 mẫu/ngành) có một số câu từ khóa chuyên biệt trùng khớp trực tiếp N-gram.", bold_prefix="Lý do 2 – Nắm bắt ngữ cảnh câu phức liên ngành: ")
    add_bullet_1(doc, "Tương thích bản ngữ tiếng Việt học thuật: Trọng số tiền huấn luyện vinai/phobert-base trên 20GB văn bản tiếng Việt kết hợp tách từ ghép PyVi giúp biểu diễn chính xác các cụm thuật ngữ dài và phân loại Nhóm loại công trình (Bài báo / Đề tài / Sáng chế) đạt F1 tuyệt đối 100.0%.", bold_prefix="Lý do 3 – Tối ưu ngữ nghĩa tiếng Việt đa nhiệm: ")
    add_bullet_1(doc, "Cung cấp xác suất Softmax đa nhiệm phục vụ phát hiện bất thường: Xác suất đầu ra từ hai nhánh phân loại của PhoBERT hỗ trợ trực tiếp việc gắn nhãn mức độ tin cậy màu sắc và tính toán đặc trưng độ lệch nhãn cho Chức năng 6.", bold_prefix="Lý do 4 – Liên kết chặt chẽ với Chức năng 6: ")
    add_bullet_1(doc, f"Thời gian phản hồi thực tế hoàn toàn khả thi trên phần cứng CPU phổ thông: Mặc dù chậm hơn TF-IDF + SVM ({svm_m['infer_time_ms']} ms), thời gian suy luận trung bình {pho_m['infer_time_ms']} ms/hồ sơ (~0.24 giây) trên CPU Intel Core i5-12450HX của PhoBERT vẫn đảm bảo phản hồi gần như tức thời khi người dùng bấm gợi ý trên Web và nhanh gấp ~{m_p['speedup_factor_vs_manual']} lần so với phân loại thủ công (~135 giây/hồ sơ).", bold_prefix="Lý do 5 – Đáp ứng tốt vận hành thực tế trên CPU: ")

    add_figure_custom(doc, "images/hinh5_danhgia_hieunang.png",
                      "Hình 5.1: Biểu đồ đối sánh thực nghiệm phân loại (PhoBERT, LinearSVC, Thủ công) trên tập Test", width_cm=14.5)

    add_h1(doc, "5.3. Đánh giá Thực nghiệm Phân hệ Phát hiện Bất thường (IF vs PyTorch AE)")
    add_body(doc, "Trên cơ sở dữ liệu vận hành thực tế khcn_system.db gồm 10 hồ sơ mẫu đại diện (trong đó có 07 hồ sơ hợp lệ chuẩn thuộc đủ 3 nhóm Bài báo, Đề tài, Sáng chế và 03 hồ sơ cài cắm vi phạm điển hình: HS-2026-0003 khai sai lĩnh vực Hóa – Sinh sang Kinh tế; HS-2026-0005 đăng trên tạp chí cảnh báo Predatory kèm lệch chuyên môn tác giả trên đồ thị; HS-2026-0008 trùng lặp nội dung với hồ sơ đã duyệt HS-2026-0001), kết quả thực thi hai mô hình không giám sát được tổng hợp tại Bảng 5.4:")

    t10_headers = ["Chỉ số Thực nghiệm Đánh giá", "Isolation Forest (scikit-learn)", "PyTorch Autoencoder (nn.Module)", "Tổ hợp Điểm Rủi ro Liên tục"]
    t10_data = [
        ["Cấu trúc mô hình thực thi", "100 iTrees, contamination = 0.25", "Bottleneck 6->16->4->16->6, Adam 250 epochs", "0.35*IF + 0.35*AE + 0.30*Rule"],
        ["Tỷ lệ cảnh báo bất thường", "30.0% (3/10 hồ sơ trong CSDL)", "30.0% (3/10 hồ sơ trong CSDL)", "Bắt đúng 3/3 hồ sơ vi phạm (100%)"],
        ["Điểm / Sai số trung bình", "Avg Norm Score = 0.345", "Avg Reconstruction MSE = 0.0002", "HS vi phạm đạt R_i từ 0.55 – 0.65"],
        ["Thời gian huấn luyện & suy luận", "~478 ms trên CPU (bao gồm dựng ma trận)", "~712 ms trên CPU (250 epochs PyTorch)", "< 1.2 giây cho toàn bộ đồ thị & 2 mô hình"],
        ["Đồ thị tri thức NetworkX", "37 đỉnh (Nodes), 45 cạnh (Edges)", "3 thành phần liên thông (Components)", "Xếp 3 hồ sơ rủi ro lên đầu Hàng đợi"],
    ]
    add_table_custom(doc, "Bảng 5.4: Kết quả thực nghiệm phân hệ phát hiện bất thường trên CSDL vận hành thực tế",
                     t10_headers, t10_data, col_widths=[4.0, 3.6, 3.8, 3.6])

    add_body(doc, "Nhận xét kết quả thực nghiệm phát hiện bất thường:")
    add_bullet_1(doc, "Tổ hợp liên tục giữa Isolation Forest, PyTorch Autoencoder và chỉ báo đồ thị/quy tắc phân tách rõ rệt nhóm hồ sơ vi phạm (HS-2026-0003 đạt R_i = 0.65; HS-2026-0005 đạt R_i = 0.63; HS-2026-0008 đạt R_i = 0.55, đều bật cờ coCanhBao = 1) với nhóm 7 hồ sơ hợp lệ (coCanhBao = 0), đồng thời bảo vệ hồ sơ gốc đã duyệt HS-2026-0001 không bị gắn cờ ngược bởi bản sao nộp sau.")
    add_bullet_1(doc, "Hàng đợi xét duyệt ưu tiên (/ai-center) tự động đưa các hồ sơ đang chờ xử lý có điểm rủi ro cao (HS-2026-0003, HS-2026-0005) lên vị trí đầu tiên kèm giải thích chi tiết từng nguyên nhân, giúp cán bộ thẩm định kiểm tra trọng tâm thay vì phải dò tìm thủ công.")

    add_h1(doc, "5.4. Kiểm thử Tích hợp Toàn diện và Đánh giá Hiệu năng Thực tế")
    add_body(doc, "Nhóm nghiên cứu đã thực hiện kiểm thử tích hợp tự động toàn bộ các định tuyến (Routes), luồng phân quyền RBAC và thao tác CSDL thông qua công cụ kiểm thử Flask test_client trực tiếp trên máy trạm thực nghiệm:")
    add_bullet_1(doc, "Kiểm thử chức năng và phân quyền RBAC: Xác nhận hoạt động chính xác 100% (HTTP 200/302, không phát sinh ngoại lệ) trên toàn bộ luồng Đăng nhập 3 vai trò (Admin, CanBo, TacGia), Kê khai 3 nhóm công trình, Lọc đa tiêu chí (theo từ khóa, tác giả, đơn vị, sáng chế), API gợi ý /api/ai-preview, Phản hồi nhãn AI, Thêm/Sửa/Ẩn cả 4 danh mục và Xuất báo cáo Excel/PDF.")
    add_bullet_1(doc, "Hiệu năng khởi động và vận hành thực tế: Nhờ cơ chế lưu trữ trọng số tinh chỉnh phobert_classifier.pt (55.22 MB) và svm_pipeline.joblib (817.55 KB), ứng dụng web khởi động và nạp mô hình trong thời gian dưới 1.5 giây mà không cần huấn luyện lại từ đầu.")
    add_bullet_1(doc, "Toàn vẹn tham chiếu dữ liệu: Thao tác vô hiệu hóa (Ẩn/Hiện) trên cả 4 danh mục trong SQLite 3 bảo toàn tuyệt đối khóa ngoại của các hồ sơ lịch sử và không tạo bản ghi mồ côi.")

    doc.add_page_break()

    # ==================== KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN ====================
    add_chapter_title(doc, "KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN")

    add_h1(doc, "1. Kết quả đạt được của khóa luận")
    add_body(doc, "Đối chiếu nghiêm ngặt với toàn bộ mục tiêu và nội dung trong Đề cương khóa luận tốt nghiệp CNTT-KLCN001, nhóm sinh viên đã hoàn thành thực chất các kết quả sau:")
    add_bullet_1(doc, f"Về mô hình Trí tuệ nhân tạo: Đã cài đặt và tinh chỉnh thực tế mô hình học sâu PyTorch VinAI PhoBERT (vinai/phobert-base) đối sánh công bằng 100% với TF-IDF + LinearSVC và phương pháp phân loại thủ công trên tập dữ liệu chuẩn 240 mẫu (13 trường). Mô hình chính PhoBERT đạt Macro F1-score {pho_m['f1_score']}% ở bài toán phân loại 6 lĩnh vực (+{sel_m['delta_f1']}% so với SVM, +{m_p['delta_f1_vs_manual']}% so với thủ công) và đạt {pho_m['work_type_f1']}% ở bài toán phân loại 3 nhóm loại công trình; đồng thời triển khai thực tế Isolation Forest, PyTorch Autoencoder (nn.Module) và đồ thị NetworkX trong phát hiện bất thường.")
    add_bullet_1(doc, "Về hệ thống phần mềm Web: Đã xây dựng hoàn chỉnh ứng dụng Flask + SQLite 3 đáp ứng trọn vẹn 06 chức năng cốt lõi: (1) Quản lý hồ sơ động cho Bài báo, Đề tài NCKH và Sáng chế/GPHI kèm Live AI Preview; (2) Quản lý đầy đủ Thêm – Sửa – Ẩn/Hiện cho 4 danh mục (Lĩnh vực, Loại công trình, Tạp chí, Danh mục Tác giả); (3) Quy trình xét duyệt 5 bước với lưu vết bất biến LichSuXetDuyet; (4) Tra cứu, lọc theo tác giả/đơn vị/lĩnh vực/loại hình và xuất báo cáo Excel/PDF; (5) Xác nhận/điều chỉnh nhãn AI đa nhiệm kèm Feedback Loop; (6) Hàng đợi xét duyệt ưu tiên theo điểm rủi ro liên tục.")

    add_h1(doc, "2. Giá trị ứng dụng thực tiễn")
    add_body(doc, "Hệ thống mang lại công cụ hỗ trợ thiết thực cho công tác quản lý khoa học cấp Trường:")
    add_bullet_1(doc, f"Rút ngắn đáng kể thời gian sơ loại hồ sơ: Hỗ trợ gợi ý phân loại tự động trong ~{pho_m['infer_time_ms']} ms/hồ sơ thay vì ~135 giây đọc và tra cứu thủ công bằng mắt thường.")
    add_bullet_1(doc, "Nâng cao chất lượng kiểm soát học thuật: Tự động phát hiện và đẩy lên đầu hàng đợi các hồ sơ nghi vấn khai lệch lĩnh vực, đăng trên tạp chí cảnh báo không uy tín hoặc trùng lặp nội dung.")
    add_bullet_1(doc, "Đảm bảo tính minh bạch và truy xuất nguồn gốc: Lưu trữ đầy đủ lịch sử thẩm định bất biến và nhật ký phản hồi nhãn giữa cán bộ duyệt và mô hình AI.")

    add_h1(doc, "3. Một số hạn chế thực tế còn tồn tại")
    add_body(doc, "Nhóm nghiên cứu nhìn nhận trung thực một số hạn chế của đề tài trong phạm vi khóa luận hiện tại:")
    add_bullet_1(doc, "Về quy mô và tính chất dữ liệu huấn luyện: Tập dữ liệu 240 hồ sơ hiện tại là bộ dữ liệu mô phỏng học thuật chuẩn hóa (30 mẫu huấn luyện / lĩnh vực). Do quy mô mẫu còn khiêm tốn, ở một số trường hợp giao thoa từ vựng hẹp (như giữa Kinh tế và Quản lý giáo dục), mô hình vẫn còn nhầm lẫn nhất định và cần được bổ sung thêm dữ liệu lịch sử thực tế nhiều năm của Nhà trường.")
    add_bullet_1(doc, "Về hạ tầng phần cứng và trích xuất tệp đính kèm: Do huấn luyện và suy luận trên CPU phổ thông nên độ dài chuỗi PhoBERT đang giới hạn ở 128 tokens và thời gian suy luận ở mức ~240 ms/hồ sơ; ngoài ra hệ thống mới phân tích trên các trường văn bản kê khai chứ chưa tự động bóc tách nội dung toàn văn từ tệp PDF đính kèm (OCR).")

    add_h1(doc, "4. Hướng mở rộng và phát triển trong tương lai")
    add_body(doc, "Từ nền tảng đã xây dựng, đề tài có thể mở rộng theo các hướng thiết thực sau:")
    add_bullet_1(doc, "Thu thập và gán nhãn tập dữ liệu hồ sơ KH&CN thực tế quy mô lớn (từ 2.000 – 5.000 hồ sơ đa ngành) và triển khai suy luận trên GPU / ONNX Runtime để nâng chiều dài chuỗi lên 256 tokens đồng thời giảm độ trễ xuống dưới 30 ms.")
    add_bullet_1(doc, "Tích hợp công nghệ OCR và trích xuất thông tin tự động từ tệp PDF bài báo/bằng độc quyền sáng chế để tự động điền biểu mẫu kê khai cho giảng viên.")
    add_bullet_1(doc, "Kết nối trực tiếp API tra cứu thời gian thực với cơ sở dữ liệu Crossref DOI, Scopus và ORCID nhằm tự động xác thực chỉ mục tạp chí và hồ sơ tác giả.")

    doc.add_page_break()

    # ==================== TÀI LIỆU THAM KHẢO ====================
    add_chapter_title(doc, "TÀI LIỆU THAM KHẢO")
    references = [
        "[1] Bộ Khoa học và Công nghệ (2017), Thông tư số 24/2017/TT-BKHCN về việc ban hành Bảng phân loại lĩnh vực nghiên cứu khoa học và công nghệ.",
        "[2] Hội đồng Giáo sư Nhà nước (2024), Quyết định phê duyệt Danh mục tạp chí khoa học được tính điểm năm 2024 của 28 Hội đồng Giáo sư ngành, liên ngành.",
        "[3] Trường Đại học Công Thương TP. Hồ Chí Minh (2025), Mẫu báo cáo Đồ án / Khóa luận tốt nghiệp năm học 2025-2026, Khoa Công nghệ Thông tin, HUIT.",
        "[4] Nguyen, D. Q., & Nguyen, A. T. (2020), 'PhoBERT: Pre-trained language models for Vietnamese', Findings of the Association for Computational Linguistics: EMNLP 2020, pp. 1037-1042.",
        "[5] Liu, F. T., Ting, K. M., & Zhou, Z. H. (2008), 'Isolation Forest', Proceedings of the 2008 Eighth IEEE International Conference on Data Mining (ICDM), pp. 413-422.",
        "[6] Paszke, A., et al. (2019), 'PyTorch: An Imperative Style, High-Performance Deep Learning Library', Advances in Neural Information Processing Systems (NeurIPS), 32, pp. 8024-8035.",
        "[7] Wolf, T., et al. (2020), 'Transformers: State-of-the-Art Natural Language Processing', Proceedings of the 2020 Conference on Empirical Methods in Natural Language Processing: System Demonstrations, pp. 38-45.",
        "[8] Cortes, C., & Vapnik, V. (1995), 'Support-vector networks', Machine Learning, 20(3), pp. 273-297.",
        "[9] Platt, J. (1999), 'Probabilistic outputs for support vector machines and comparisons to regularized likelihood methods', Advances in Large Margin Classifiers, 10(3), pp. 61-74.",
        "[10] Hagberg, A., Swart, P., & Chult, D. S. (2008), 'Exploring network structure, dynamics, and function using NetworkX', Proceedings of the 7th Python in Science Conference (SciPy 2008), pp. 11-15.",
        "[11] Rumelhart, D. E., Hinton, G. E., & Williams, R. J. (1986), 'Learning representations by back-propagating errors', Nature, 323(6088), pp. 533-536.",
        "[12] Vaswani, A., et al. (2017), 'Attention is all you need', Advances in Neural Information Processing Systems, 30, pp. 5998-6008.",
        "[13] Salton, G., & Buckley, C. (1988), 'Term-weighting approaches in automatic text retrieval', Information Processing & Management, 24(5), pp. 513-523.",
        "[14] Pedregosa, F., et al. (2011), 'Scikit-learn: Machine Learning in Python', Journal of Machine Learning Research, 12, pp. 2825-2830.",
        "[15] Hipp, D. R. (2024), SQLite Relational Database Management System Documentation, SQLite Consortium (https://www.sqlite.org/docs.html).",
    ]
    for ref in references:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.left_indent = Cm(0.8)
        p_ref.paragraph_format.first_line_indent = Cm(-0.8)
        p_ref.paragraph_format.space_before = Pt(3)
        p_ref.paragraph_format.space_after = Pt(3)
        p_ref.paragraph_format.line_spacing = 1.3
        r_ref = p_ref.add_run(ref)
        r_ref.font.name = 'Times New Roman'
        r_ref.font.size = Pt(12)
        r_ref.font.color.rgb = RGBColor(0, 0, 0)

    # Save document
    print(f"Saving report to: {OUTPUT_FILE}...")
    doc.save(OUTPUT_FILE)
    print("SUCCESS: Bao_Cao_Khoa_Luan.docx generated successfully!")

if __name__ == '__main__':
    generate_thesis_report()
