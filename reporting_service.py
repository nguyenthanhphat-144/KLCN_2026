# -*- coding: utf-8 -*-
"""
Module Kết xuất Báo cáo Thống kê ra định dạng Excel (.xlsx) và PDF (.pdf) (Tuần 5 - Mục 3b)
Đề tài: Hệ thống Hỗ trợ Quản lý Hồ sơ KH&CN Tích hợp Phân loại Tự động (CNTT-KLCN001)
"""

import io
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def _register_vietnamese_font():
    """Đăng ký phông chữ hỗ trợ Unicode tiếng Việt từ hệ thống Windows cho ReportLab."""
    font_candidates = [
        ("VNFont", "C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("VNFont", "C:/Windows/Fonts/times.ttf", "C:/Windows/Fonts/timesbd.ttf"),
        ("VNFont", "C:/Windows/Fonts/tahoma.ttf", "C:/Windows/Fonts/tahomabd.ttf"),
    ]
    for family_name, reg_path, bold_path in font_candidates:
        if os.path.exists(reg_path):
            try:
                pdfmetrics.registerFont(TTFont(family_name, reg_path))
                if os.path.exists(bold_path):
                    pdfmetrics.registerFont(TTFont(f"{family_name}-Bold", bold_path))
                else:
                    pdfmetrics.registerFont(TTFont(f"{family_name}-Bold", reg_path))
                return family_name, f"{family_name}-Bold"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"


def generate_excel_report(records: list, filter_desc: str = "Tất cả hồ sơ KH&CN") -> io.BytesIO:
    """
    Xuất danh sách chi tiết hồ sơ KH&CN kèm kết quả phân loại AI ra tệp Excel (.xlsx)
    sử dụng thư viện openpyxl.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "DanhSach_HoSo_KHCN"

    # Styles
    title_font = Font(name="Times New Roman", size=14, bold=True, color="1E3A8A")
    sub_font = Font(name="Times New Roman", size=11, italic=True, color="334155")
    hdr_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    data_font = Font(name="Times New Roman", size=11)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    # Header Title
    ws.merge_cells("A1:K1")
    ws["A1"] = "BÁO CÁO THỐNG KÊ DANH SÁCH HỒ SƠ KHOA HỌC VÀ CÔNG NGHỆ (CNTT-KLCN001)"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:K2")
    ws["A2"] = f"Điều kiện lọc: {filter_desc} | Thời gian kết xuất: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    ws["A2"].font = sub_font
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        "STT", "Mã Hồ Sơ", "Tiêu Đề Công Trình KH&CN", "Tác Giả", "Đơn Vị Chủ Trì",
        "Kỳ Báo Cáo", "Lĩnh Vực Khai Báo", "Loại Công Trình", "Nhãn AI Gợi Ý (PhoBERT)",
        "Độ Tin Cậy AI (%)", "Trạng Thái Xét Duyệt"
    ]

    ws.append([])
    ws.append(headers)
    header_row_idx = 4
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for idx, r in enumerate(records, start=1):
        row_vals = [
            idx,
            r.get("maHoSo", ""),
            r.get("tieuDe", ""),
            r.get("thongTinTacGia", ""),
            r.get("donViChuTri", ""),
            r.get("kyBaoCao", ""),
            r.get("tenLinhVuc", ""),
            r.get("tenLoai", ""),
            r.get("tenLinhVucAI", ""),
            f"{float(r.get('doTinCayAI') or 0.0):.1f}%",
            r.get("trangThai", ""),
        ]
        ws.append(row_vals)
        curr_row = header_row_idx + idx
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=curr_row, column=col_idx)
            c.font = data_font
            c.border = thin_border
            c.alignment = Alignment(
                horizontal="center" if col_idx in (1, 2, 6, 10, 11) else "left",
                vertical="center",
                wrap_text=True,
            )

    col_widths = [6, 15, 42, 26, 26, 22, 26, 28, 26, 16, 18]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def generate_pdf_report(records: list, filter_desc: str = "Tất cả hồ sơ KH&CN") -> io.BytesIO:
    """
    Xuất báo cáo hành chính chuẩn định dạng PDF (.pdf) hỗ trợ phông chữ tiếng Việt Unicode.
    """
    reg_font, bold_font = _register_vietnamese_font()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=24,
        rightMargin=24,
        topMargin=24,
        bottomMargin=24,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "VNTitle",
        parent=styles["Heading1"],
        fontName=bold_font,
        fontSize=15,
        leading=20,
        alignment=1,
        textColor=colors.HexColor("#1E3A8A"),
    )
    sub_style = ParagraphStyle(
        "VNSub",
        parent=styles["Normal"],
        fontName=reg_font,
        fontSize=10,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#475569"),
    )
    cell_style = ParagraphStyle(
        "VNCell",
        parent=styles["Normal"],
        fontName=reg_font,
        fontSize=8.5,
        leading=11,
    )
    hdr_cell_style = ParagraphStyle(
        "VNHdrCell",
        parent=styles["Normal"],
        fontName=bold_font,
        fontSize=9,
        leading=12,
        textColor=colors.white,
        alignment=1,
    )

    elements = [
        Paragraph("BÁO CÁO TỔNG HỢP HỒ SƠ KHOA HỌC VÀ CÔNG NGHỆ (CNTT-KLCN001)", title_style),
        Spacer(1, 4),
        Paragraph(
            f"Phạm vi: {filter_desc} | Ngày xuất báo cáo: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            sub_style,
        ),
        Spacer(1, 14),
    ]

    table_data = [[
        Paragraph("STT", hdr_cell_style),
        Paragraph("Mã HS", hdr_cell_style),
        Paragraph("Tiêu đề công trình KH&CN", hdr_cell_style),
        Paragraph("Tác giả & Đơn vị", hdr_cell_style),
        Paragraph("Lĩnh vực khai báo", hdr_cell_style),
        Paragraph("Nhãn AI (Tin cậy)", hdr_cell_style),
        Paragraph("Trạng thái", hdr_cell_style),
    ]]

    for idx, r in enumerate(records, start=1):
        table_data.append([
            Paragraph(str(idx), cell_style),
            Paragraph(str(r.get("maHoSo", "")), cell_style),
            Paragraph(str(r.get("tieuDe", "")), cell_style),
            Paragraph(f"{r.get('thongTinTacGia', '')}<br/><i>({r.get('donViChuTri', '')})</i>", cell_style),
            Paragraph(str(r.get("tenLinhVuc", "")), cell_style),
            Paragraph(f"{r.get('tenLinhVucAI', '')} ({float(r.get('doTinCayAI') or 0.0):.1f}%)", cell_style),
            Paragraph(str(r.get("trangThai", "")), cell_style),
        ])

    col_widths = [30, 75, 250, 145, 115, 115, 75]
    pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    pdf_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    elements.append(pdf_table)
    doc.build(elements)
    buffer.seek(0)
    return buffer
