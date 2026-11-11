# -*- coding: utf-8 -*-
"""Báo cáo KH&CN: giao diện, bộ lọc, Excel và PDF trong một tệp."""

import io
import json
import math
import os
from collections import Counter
from datetime import date, datetime, timedelta
from functools import wraps
from urllib.parse import urlencode
from xml.sax.saxutils import escape

from flask import (
    Blueprint,
    abort,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
)

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.graphics.shapes import Drawing, Rect, String


BLUE = "17365D"

STATES = [
    "Chờ xét duyệt",
    "Đang xét duyệt",
    "Đã duyệt",
    "Từ chối",
    "Yêu cầu bổ sung",
]

FILTER_KEYS = [
    "nhom",
    "nam",
    "tuNgay",
    "denNgay",
    "donVi",
    "maLinhVuc",
    "maLoai",
    "trangThai",
    "kyBaoCao",
    "chiMuc",
    "sanPhamDuKien",
]


def text(value):
    return "" if value is None else str(value)


def group_name(record):
    name = (
        text(record.get("nhomCongTrinh"))
        + " "
        + text(record.get("tenLoai"))
    ).casefold()

    if "bài báo" in name:
        return "Bài báo"
    if "đề tài" in name:
        return "Đề tài"
    return "Khác"


def distribution(records, key):
    counts = Counter(
        text(record.get(key)).strip() or "Chưa kê khai"
        for record in records
    )
    return sorted(
        counts.items(),
        key=lambda item: (-item[1], item[0]),
    )


def grouped_statistics(records, key):
    result = []

    for label, total in distribution(records, key):
        approved = sum(
            record.get("trangThai") == "Đã duyệt"
            for record in records
            if (
                text(record.get(key)).strip()
                or "Chưa kê khai"
            ) == label
        )

        result.append([
            label,
            total,
            approved,
            approved / total if total else 0,
        ])

    return result


def report_statistics(records):
    records = [dict(record) for record in records]

    articles = [
        record for record in records
        if group_name(record) == "Bài báo"
    ]

    projects = [
        record for record in records
        if group_name(record) == "Đề tài"
    ]

    states = Counter(
        record.get("trangThai")
        for record in records
    )

    total = len(records)
    approved = states["Đã duyệt"]

    funding = 0.0
    missing_funding = 0
    zero_funding = 0

    for record in projects:
        value = record.get("kinhPhi")

        if value is None or value == "":
            missing_funding += 1
            continue

        try:
            amount = float(value)
        except (TypeError, ValueError):
            missing_funding += 1
            continue

        funding += amount
        if amount == 0:
            zero_funding += 1

    summary = [
        ["Tổng hồ sơ", total],
        ["Bài báo", len(articles)],
        ["Đề tài", len(projects)],
        ["Công trình khác", total - len(articles) - len(projects)],
        *[[state, states[state]] for state in STATES],
        ["Tỷ lệ đã duyệt", approved / total if total else 0],
    ]

    sections = [
        (
            "Bài báo theo năm công bố",
            distribution(articles, "namCongBo"),
        ),
        (
            "Bài báo theo loại công trình",
            distribution(articles, "tenLoai"),
        ),
        (
            "Bài báo theo chỉ mục tạp chí",
            distribution(articles, "chiMucTapChi"),
        ),
        (
            "Bài báo theo đơn vị chủ trì",
            distribution(articles, "donViChuTri"),
        ),
        (
            "Đề tài theo loại công trình",
            distribution(projects, "tenLoai"),
        ),
        (
            "Đề tài theo đơn vị chủ trì",
            distribution(projects, "donViChuTri"),
        ),
        (
            "Đề tài theo sản phẩm dự kiến",
            distribution(projects, "sanPhamDuKien"),
        ),
    ]

    return {
        "total": total,
        "articles": len(articles),
        "projects": len(projects),
        "approved": approved,
        "approval_rate": approved / total if total else 0,
        "summary": summary,
        "states": [
            [state, states[state]]
            for state in STATES
        ],
        "units": grouped_statistics(records, "donViChuTri"),
        "fields": grouped_statistics(records, "tenLinhVuc"),
        "sections": sections,
        "funding": funding,
        "missing_funding": missing_funding,
        "zero_funding": zero_funding,
    }


# =========================================================
# EXCEL
# =========================================================

def excel_value(value):
    if isinstance(value, str) and value.startswith(
        ("=", "+", "-", "@")
    ):
        return "'" + value
    return value


def _excel_status_colors(sheet, headers):
    """Rules remain live when users change a status in Excel."""
    if "Trạng thái xét duyệt" not in headers or sheet.max_row < 6:
        return
    column = get_column_letter(headers.index("Trạng thái xét duyệt") + 1)
    palette = {
        "Chờ xét duyệt": ("FFF0CD", "936006"),
        "Đang xét duyệt": ("DFEEFF", "1957A5"),
        "Đã duyệt": ("DFF2E8", "146C43"),
        "Từ chối": ("FCE3E6", "A1263B"),
        "Yêu cầu bổ sung": ("FCE9D7", "934D13"),
    }
    for state, (fill, color) in palette.items():
        sheet.conditional_formatting.add(
            f"{column}6:{column}{sheet.max_row}",
            FormulaRule(formula=[f'{column}6="{state}"'],
                        fill=PatternFill("solid", fgColor=fill),
                        font=Font(bold=True, color=color)),
        )


def _excel_summary(sheet, listing, stats, description):
    """Dashboard generated from the current filtered records, never fixed demo data."""
    navy, teal, pale, muted = "193653", "087F8C", "F1F5F9", "61758A"
    sheet.sheet_properties.tabColor = navy
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A4"
    for col in range(1, 13):
        sheet.column_dimensions[get_column_letter(col)].width = 12
    for row in range(1, 51):
        sheet.row_dimensions[row].height = 23
        for col in range(1, 13):
            cell = sheet.cell(row, col)
            cell.font = Font(name="Arial", size=11, color=navy)
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    def merged(address, value, fill=None, size=11, bold=False, color=navy, align="left"):
        sheet.merge_cells(address)
        cell = sheet[address.split(":")[0]]
        cell.value = excel_value(value)
        cell.font = Font(name="Arial", size=size, bold=bold, color=color)
        cell.alignment = Alignment(vertical="center", horizontal=align, wrap_text=True)
        if fill:
            for row in sheet[address]:
                for target in row:
                    target.fill = PatternFill("solid", fgColor=fill)
        return cell

    merged("A1:L2", "HUIT | BÁO CÁO HỒ SƠ KHOA HỌC & CÔNG NGHỆ", navy, 19, True, "FFFFFF")
    merged("A3:L3", "Phạm vi: " + description, pale, 10, color=muted)
    sheet.row_dimensions[3].height = 32
    merged("A4:L4", "Ngày xuất: " + datetime.now().strftime("%d/%m/%Y %H:%M"), size=10, color=muted)
    last = max(6, listing.max_row)
    statuses = f"DanhSach!$J$6:$J${last}"
    types = f"DanhSach!$C$6:$C${last}"
    cards = [
        ("A", "C", "TỔNG HỒ SƠ", f"=COUNTA(DanhSach!B6:B{last})", navy, "#,##0"),
        ("D", "F", "CHỜ XÉT DUYỆT", f'=COUNTIF({statuses},"Chờ xét duyệt")', "A56B08", "#,##0"),
        ("G", "I", "ĐÃ DUYỆT", f'=COUNTIF({statuses},"Đã duyệt")', "13845B", "#,##0"),
        ("J", "L", "TỶ LỆ ĐÃ DUYỆT", '=IF(A6=0,0,G6/A6)', teal, "0.00%"),
    ]
    for left, right, title, formula, color, fmt in cards:
        merged(f"{left}5:{right}5", title, color, 11, True, "FFFFFF")
        cell = merged(f"{left}6:{right}8", "", pale, 30, True, color, "center")
        cell.value, cell.number_format = formula, fmt
    merged("A10:F10", "01  TIẾN ĐỘ XÉT DUYỆT", navy, bold=True, color="FFFFFF")
    merged("H10:L10", "02  CƠ CẤU CÔNG TRÌNH", navy, bold=True, color="FFFFFF")
    merged("A12:D12", "Trạng thái", pale, bold=True, color=muted)
    merged("H12:J12", "Loại hình", pale, bold=True, color=muted)
    for address, label in [("E12", "Số lượng"), ("F12", "Tỷ trọng"), ("K12", "Số lượng"), ("L12", "Tỷ trọng")]:
        sheet[address] = label
        sheet[address].fill = PatternFill("solid", fgColor=pale)
        sheet[address].font = Font(name="Arial", size=10, bold=True, color=muted)
    sheet.row_dimensions[12].height = 30
    for row, state in enumerate(STATES, 13):
        merged(f"A{row}:D{row}", state)
        sheet[f"E{row}"] = f'=COUNTIF({statuses},A{row})'
        sheet[f"F{row}"] = f'=IF($A$6=0,0,E{row}/$A$6)'
        sheet[f"F{row}"].number_format = "0.0%"
        sheet.row_dimensions[row].height = 29
    for row, name in enumerate(["Bài báo", "Đề tài", "Công trình khác"], 13):
        merged(f"H{row}:J{row}", name)
        sheet[f"K{row}"] = f'=COUNTIF({types},H{row})' if row < 15 else '=$A$6-SUM(K13:K14)'
        sheet[f"L{row}"] = f'=IF($A$6=0,0,K{row}/$A$6)'
        sheet[f"L{row}"].number_format = "0.0%"

    for title, cat_col, val_col, end, anchor, width, color in [
        ("Số hồ sơ theo trạng thái", 1, 5, 17, "A20", 17.5, teal),
        ("Số hồ sơ theo loại hình", 8, 11, 15, "H20", 12.5, navy),
    ]:
        chart = BarChart()
        chart.title, chart.legend = title, None
        chart.add_data(Reference(sheet, min_col=val_col, min_row=12, max_row=end), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=cat_col, min_row=13, max_row=end))
        chart.width, chart.height = width, 10
        chart.y_axis.numFmt = "0"
        chart.series[0].graphicalProperties.solidFill = color
        chart.series[0].graphicalProperties.line.solidFill = color
        sheet.add_chart(chart, anchor)

    merged("A36:L36", "03  KINH PHÍ ĐỀ TÀI & CÁCH ĐỌC BÁO CÁO", navy, bold=True, color="FFFFFF")
    for row, (label, value) in enumerate([
        ("Tổng kinh phí đã kê khai (đồng)", stats["funding"]),
        ("Đề tài chưa có kinh phí", stats["missing_funding"]),
        ("Đề tài lưu kinh phí bằng 0", stats["zero_funding"]),
    ], 38):
        fill = pale if row % 2 == 0 else "FFFFFF"
        merged(f"A{row}:H{row}", label, fill)
        merged(f"I{row}:L{row}", value, fill, align="right").number_format = "#,##0"
        sheet.row_dimensions[row].height = 28
    for row, note in [
        (43, "Tỷ lệ đã duyệt = số hồ sơ đã duyệt / tổng hồ sơ trong phạm vi báo cáo."),
        (45, "Phê duyệt tính hợp lệ của hồ sơ không đồng nghĩa với nghiệm thu đề tài."),
        (47, "Chỉ số số lượng liên kết với DanhSach. Kinh phí và ThongKe là số liệu tại thời điểm xuất."),
        (49, "Đơn vị chủ trì giữ nguyên theo dữ liệu nguồn; các tên gần giống chưa được tự động gộp."),
    ]:
        merged(f"A{row}:L{row+1}", note, size=10, color=muted)
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = sheet.page_setup.fitToHeight = 1
    sheet.print_area = "A1:L50"
    sheet.print_title_rows = "1:3"
    sheet.oddFooter.center.text = "Trang &P / &N"


def excel_table(
    workbook,
    name,
    title,
    description,
    headers,
    rows,
    widths,
    percentage_columns=(),
):
    sheet = workbook.create_sheet(name)
    last_column = get_column_letter(len(headers))

    for row, value in [
        [1, title],
        [2, "Phạm vi: " + description],
        [
            3,
            "Ngày xuất: "
            + datetime.now().strftime("%d/%m/%Y %H:%M"),
        ],
    ]:
        sheet.merge_cells(
            f"A{row}:{last_column}{row}"
        )

        cell = sheet.cell(
            row, 1, excel_value(value)
        )

        cell.font = Font(
            name="Arial",
            size=16 if row == 1 else 10,
            bold=row == 1,
            color="FFFFFF" if row == 1 else "61758A",
        )
        cell.fill = PatternFill("solid", fgColor="193653" if row == 1 else "F1F5F9")

        cell.alignment = Alignment(
            wrap_text=True,
            vertical="center",
        )

        sheet.row_dimensions[row].height = (
            38 if row < 3 else 24
        )

    for column, label in enumerate(headers, 1):
        cell = sheet.cell(5, column, label)
        cell.font = Font(
            name="Arial",
            bold=True,
            color="FFFFFF",
        )
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(
            wrap_text=True,
            vertical="center",
        )

    sheet.row_dimensions[5].height = 36

    for row_number, values in enumerate(rows, 6):
        estimated_lines = 1

        for column, value in enumerate(values, 1):
            cell = sheet.cell(
                row_number,
                column,
                excel_value(value),
            )

            cell.font = Font(
                name="Arial",
                size=11,
                color="243447",
            )

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

            if row_number % 2 == 0:
                cell.fill = PatternFill(
                    "solid",
                    fgColor="F3F6FA",
                )

            if column in percentage_columns:
                cell.number_format = "0.0%"
            elif isinstance(value, (int, float)):
                cell.number_format = "0" if headers[column - 1] in ("Năm công bố", "STT") else "#,##0.##"

            capacity = max(
                6, int(widths[column - 1] * 0.8)
            )

            lines = sum(
                max(
                    1,
                    (len(part) + capacity - 1) // capacity,
                )
                for part in text(value).split("\n")
            )

            estimated_lines = max(
                estimated_lines, lines
            )

        sheet.row_dimensions[row_number].height = min(
            409,
            max(28, estimated_lines * 16 + 8),
        )

    for column, width in enumerate(widths, 1):
        sheet.column_dimensions[
            get_column_letter(column)
        ].width = width

    sheet.sheet_properties.tabColor = "193653"
    _excel_status_colors(sheet, headers)
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A6"

    sheet.auto_filter.ref = (
        f"A5:{last_column}{max(sheet.max_row, 5)}"
    )

    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"

    sheet.page_setup.paperSize = (
        sheet.PAPERSIZE_A3
        if len(headers) > 7
        else sheet.PAPERSIZE_A4
    )

    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0

    sheet.print_title_rows = "1:5"
    sheet.print_area = (
        f"A1:{last_column}{max(sheet.max_row, 5)}"
    )

    sheet.oddFooter.center.text = "Trang &P / &N"

    return sheet


def detail_columns(group):
    if group == "Bài báo":
        return [
            ("maHoSo", "Mã hồ sơ"),
            ("tieuDe", "Tiêu đề công trình"),
            ("donViChuTri", "Đơn vị chủ trì"),
            ("namCongBo", "Năm công bố"),
            ("kyBaoCao", "Kỳ báo cáo"),
            ("tenLinhVuc", "Lĩnh vực đăng ký"),
            ("tenLoai", "Loại công trình"),
            ("tenTapChi", "Tên tạp chí / Hội thảo"),
            ("chiMucTapChi", "Chỉ mục tạp chí"),
            ("trangThai", "Trạng thái xét duyệt"),
        ]

    return [
        ("maHoSo", "Mã hồ sơ"),
        ("tieuDe", "Tên đề tài"),
        ("donViChuTri", "Đơn vị chủ trì"),
        ("namCongBo", "Năm công bố"),
        ("kyBaoCao", "Kỳ báo cáo"),
        ("tenLinhVuc", "Lĩnh vực đăng ký"),
        ("tenLoai", "Loại công trình"),
        ("sanPhamDuKien", "Sản phẩm dự kiến"),
        ("trangThai", "Trạng thái xét duyệt"),
    ]


def review_support_statistics(records):
    """Summarize stored analysis and active alerts in the export scope."""
    reason_labels = {
        "nghi vấn trùng lặp nội dung": "Nghi vấn trùng nội dung",
        "lệch lĩnh vực đăng ký": "Lệch lĩnh vực đăng ký",
        "tạp chí cần lưu ý": "Tạp chí cần lưu ý",
        "khác biệt hướng nghiên cứu": "Khác biệt hướng nghiên cứu",
    }
    categories = dict.fromkeys([*reason_labels.values(), "Lý do khác", "Chưa ghi lý do"], 0)
    analyzed = warned = 0
    priority = []

    for record in records:
        try:
            payload = json.loads(record.get("chiTietCanhBao") or "{}")
        except (ValueError, TypeError):
            payload = {}
        raw_reasons = payload.get("reasons", []) if isinstance(payload, dict) else payload
        reasons = (
            [item.strip() for item in raw_reasons if isinstance(item, str) and item.strip()]
            if isinstance(raw_reasons, list) else []
        )
        active_alert = text(record.get("coCanhBao")).strip().casefold() in {"1", "true"}
        has_score = isinstance(payload, dict) and payload.get("composite_risk") is not None
        if record.get("maLinhVucAI") or record.get("maLinhVucSVM") or has_score or reasons or active_alert:
            analyzed += 1
        if not active_alert:
            continue

        warned += 1
        groups = {
            reason_labels.get(reason.partition(":")[0].casefold().strip(), "Lý do khác")
            for reason in reasons
        }
        for group in groups or {"Chưa ghi lý do"}:
            categories[group] += 1

        state = record.get("trangThai")
        if state not in ("Chờ xét duyệt", "Đang xét duyệt"):
            continue
        try:
            score = float(record.get("diemBatThuong"))
            if not math.isfinite(score):
                score = None
        except (TypeError, ValueError):
            score = None
        priority.append([
            record.get("maHoSo"), record.get("tieuDe"), state,
            score, "\n".join(dict.fromkeys(reasons)) or "Chưa ghi lý do",
        ])

    priority.sort(key=lambda row: (row[3] is None, -(row[3] or 0), text(row[0])))
    return {
        "metrics": [
            ("Tổng hồ sơ trong phạm vi báo cáo", len(records)),
            ("Hồ sơ có kết quả phân tích đã lưu", analyzed),
            ("Hồ sơ có cảnh báo đang bật", warned),
            ("Hồ sơ cần ưu tiên xem xét", len(priority)),
        ],
        "categories": list(categories.items()),
        "priority": priority,
    }


REVIEW_NOTE = (
    "Mỗi hồ sơ được đếm một lần trong tổng cảnh báo và một lần trong mỗi nhóm lý do. "
    "Danh sách ưu tiên gồm hồ sơ chờ hoặc đang xét duyệt có cảnh báo đang bật, "
    "sắp theo điểm bất thường giảm dần."
)
REVIEW_HEADERS = ["Mã hồ sơ", "Tên công trình", "Trạng thái xét duyệt", "Điểm bất thường", "Lý do cảnh báo"]


def _excel_review_support(workbook, review, filter_desc):
    rows = [
        *[list(row) for row in review["metrics"]],
        ["Lý do cảnh báo", "Số hồ sơ"],
        *[list(row) for row in review["categories"]],
        ["HỒ SƠ CẦN ƯU TIÊN XEM XÉT"],
        REVIEW_HEADERS,
        *(review["priority"] or [["Không có hồ sơ cần ưu tiên"]]),
    ]
    sheet = excel_table(
        workbook, "HoTroXetDuyet", "KẾT QUẢ HỖ TRỢ XÉT DUYỆT",
        filter_desc, ["Chỉ tiêu", "Số hồ sơ", "", "", ""], rows, [32, 49, 23, 17, 58],
    )
    workbook.move_sheet(sheet, offset=1 - workbook.index(sheet))
    heading_row = 6 + len(review["metrics"])
    priority_title = heading_row + 1 + len(review["categories"])
    priority_header = priority_title + 1
    sheet.merge_cells(start_row=priority_title, start_column=1, end_row=priority_title, end_column=5)
    for row in (heading_row, priority_title, priority_header):
        for cell in sheet[row]:
            cell.fill = PatternFill("solid", fgColor="193653")
            cell.font = Font(name="Arial", bold=True, color="FFFFFF")
    sheet.auto_filter.ref = (
        f"A{priority_header}:E{sheet.max_row}" if review["priority"] else None
    )
    for row in range(priority_header + 1, sheet.max_row + 1):
        sheet.cell(row, 4).number_format = "0.00"
    note_row = sheet.max_row + 2
    sheet.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=5)
    note = sheet.cell(note_row, 1, REVIEW_NOTE)
    note.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[note_row].height = 32
    sheet.print_area = f"A1:E{note_row}"
    sheet.print_title_rows = "1:3"


def generate_excel_report(
    records: list,
    filter_desc: str = "Tất cả hồ sơ KH&CN",
) -> io.BytesIO:
    records = [dict(record) for record in records]
    stats = report_statistics(records)

    workbook = Workbook()
    workbook.remove(workbook.active)

    summary = workbook.create_sheet("TongHop")

    statistical_rows = []

    for title, values in [
        ["Theo đơn vị", stats["units"]],
        ["Theo lĩnh vực đang lưu", stats["fields"]],
    ]:
        statistical_rows.append([
            title, None, None, None
        ])

        statistical_rows.extend(values)

        statistical_rows.append([
            "", None, None, None
        ])

    for title, values in stats["sections"]:
        statistical_rows.append([
            title, None, None, None
        ])

        for label, count in values:
            statistical_rows.append([
                label, count, None, None
            ])

        if not values:
            statistical_rows.append([
                "Không có dữ liệu", None, None, None
            ])

        statistical_rows.append([
            "", None, None, None
        ])

    analysis = excel_table(
        workbook,
        "ThongKe",
        "THỐNG KÊ THEO NHÓM",
        filter_desc,
        [
            "Phân nhóm",
            "Số hồ sơ",
            "Đã duyệt",
            "Tỷ lệ duyệt trong nhóm",
        ],
        statistical_rows,
        [58, 18, 18, 26],
        percentage_columns=(4,),
    )

    analysis.auto_filter.ref = None
    analysis.sheet_properties.tabColor = "087F8C"
    for row in range(6, analysis.max_row + 1):
        if analysis.cell(row, 1).value and analysis.cell(row, 2).value is None:
            analysis.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            cell = analysis.cell(row, 1)
            cell.fill = PatternFill("solid", fgColor="087F8C")
            cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            analysis.row_dimensions[row].height = 30

    list_rows = []

    for index, record in enumerate(records, 1):
        list_rows.append([
            index,
            record.get("maHoSo", ""),
            group_name(record),
            record.get("tieuDe", ""),
            record.get("donViChuTri", ""),
            record.get("namCongBo", ""),
            record.get("kyBaoCao", ""),
            record.get("tenLinhVuc", ""),
            record.get("tenLoai", ""),
            record.get("trangThai", ""),
        ])

    listing = excel_table(
        workbook,
        "DanhSach",
        "DANH SÁCH HỒ SƠ KH&CN",
        filter_desc,
        [
            "STT",
            "Mã hồ sơ",
            "Loại hình",
            "Tiêu đề công trình / Tên đề tài",
            "Đơn vị chủ trì",
            "Năm công bố",
            "Kỳ báo cáo",
            "Lĩnh vực đăng ký",
            "Loại công trình",
            "Trạng thái xét duyệt",
        ],
        list_rows,
        [6, 18, 16, 56, 32, 16, 26, 30, 32, 22],
    )

    listing.freeze_panes = "C6"

    for group, sheet_name in [
        ["Bài báo", "BaiBao"],
        ["Đề tài", "DeTai"],
    ]:
        subset = [
            record for record in records
            if group_name(record) == group
        ]

        columns = detail_columns(group)
        detail_rows = []

        for record in subset:
            values = []

            for key, _ in columns:
                value = record.get(key)

                if value is None or value == "":
                    value = "Chưa kê khai"

                values.append(value)

            detail_rows.append(values)

        widths = [
            60 if key in ("tieuDe", "sanPhamDuKien")
            else 35 if key in (
                "thongTinTacGia",
                "thanhVienThamGia",
                "donViChuTri",
                "tenTapChi",
            )
            else 23
            for key, _ in columns
        ]

        sheet = excel_table(
            workbook,
            sheet_name,
            "CHI TIẾT " + group.upper(),
            filter_desc,
            [label for _, label in columns],
            detail_rows,
            widths,
        )

        sheet.freeze_panes = "C6"

    _excel_summary(summary, listing, stats, filter_desc)
    _excel_review_support(workbook, review_support_statistics(records), filter_desc)
    workbook.active = 0

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output


# =========================================================
# PDF
# =========================================================

def register_fonts():
    base = os.path.dirname(os.path.abspath(__file__))

    candidates = [
        (
            os.path.join(base, "fonts", "DejaVuSans.ttf"),
            os.path.join(base, "fonts", "DejaVuSans-Bold.ttf"),
        ),
        (
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
        ),
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ),
    ]

    for regular, bold in candidates:
        if os.path.isfile(regular):
            pdfmetrics.registerFont(
                TTFont("KHCNRegular", regular)
            )

            pdfmetrics.registerFont(
                TTFont(
                    "KHCNBold",
                    bold if os.path.isfile(bold) else regular,
                )
            )

            return "KHCNRegular", "KHCNBold"

    raise RuntimeError(
        "Không tìm thấy font tiếng Việt. "
        "Thêm DejaVuSans.ttf và DejaVuSans-Bold.ttf "
        "vào thư mục fonts."
    )


def generate_pdf_report(
    records: list,
    filter_desc: str = "Tất cả hồ sơ KH&CN",
) -> io.BytesIO:
    """Báo cáo PDF với tổng hợp một trang, thống kê và danh sách chi tiết."""
    from reportlab.platypus import PageBreak, KeepTogether
    from reportlab.pdfgen.canvas import Canvas

    records = [dict(record) for record in records]
    stats = report_statistics(records)
    regular, bold = register_fonts()
    navy = colors.HexColor("#193653")
    teal = colors.HexColor("#087F8C")
    pale = colors.HexColor("#F1F5F9")
    muted = colors.HexColor("#61758A")
    white = colors.white
    palette = {
        "Chờ xét duyệt": ("#FFF0CD", "#936006"),
        "Đang xét duyệt": ("#DFEEFF", "#1957A5"),
        "Đã duyệt": ("#DFF2E8", "#146C43"),
        "Từ chối": ("#FCE3E6", "#A1263B"),
        "Yêu cầu bổ sung": ("#FCE9D7", "#934D13"),
    }
    stamp = datetime.now().strftime("%d/%m/%Y %H:%M")
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=landscape(A4),
        leftMargin=32, rightMargin=32, topMargin=46, bottomMargin=36,
        title="Báo cáo thống kê hồ sơ KH&CN", author="HUIT Research",
    )
    body = ParagraphStyle("ReportBody", fontName=regular, fontSize=8.5,
                          leading=12, textColor=navy, splitLongWords=True)
    small = ParagraphStyle("ReportSmall", parent=body, fontSize=8, leading=11, textColor=muted)
    title_style = ParagraphStyle("ReportTitle", parent=body, fontName=bold,
                                fontSize=19, leading=25, spaceAfter=8)
    heading = ParagraphStyle("ReportHeading", parent=body, fontName=bold,
                             fontSize=12, leading=17, spaceBefore=10,
                             spaceAfter=8, keepWithNext=True)
    header = ParagraphStyle("ReportHeader", parent=body, fontName=bold,
                            fontSize=8.5, textColor=white)

    def p(value, style=body):
        return Paragraph(escape(text(value)).replace("\n", "<br/>") or "-", style)

    def table(headers, rows, ratios, status_column=None, caption=None):
        rows = list(rows)
        if not rows:
            rows = [["Không có dữ liệu"] + [""] * (len(headers) - 1)]
        data = [[p(h, header) for h in headers]]
        for row in rows:
            cells = []
            for i, value in enumerate(row):
                style = body
                if i == status_column and value in palette:
                    style = ParagraphStyle("ReportState", parent=body, fontName=bold,
                                           textColor=colors.HexColor(palette[value][1]))
                cells.append(p(value, style))
            data.append(cells)
        if caption:
            data.insert(0, [p(caption, header)] + [""] * (len(headers) - 1))
        first_data = 2 if caption else 1
        # Keep ordinary records together; split within a row only for exceptionally long content.
        oversized_row = any(
            isinstance(cell, Paragraph)
            and cell.wrap(doc.width * ratios[col] - 16, doc.height)[1] > doc.height - 100
            for row in data[first_data:] for col, cell in enumerate(row)
        )
        result = Table(data, colWidths=[doc.width * r for r in ratios],
                       repeatRows=first_data, hAlign="LEFT", splitByRow=1,
                       splitInRow=1 if oversized_row else 0)
        commands = [
            ("BACKGROUND", (0, 0), (-1, first_data - 1), navy),
            ("ROWBACKGROUNDS", (0, first_data), (-1, -1), [white, pale]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("LINEBELOW", (0, 0), (-1, 0), 1, teal),
            ("LINEBELOW", (0, 1), (-1, -1), .35, colors.HexColor("#DBE4ED")),
        ]
        if caption:
            commands.extend([("SPAN", (0, 0), (-1, 0)),
                             ("BACKGROUND", (0, 0), (-1, 0), teal)])
        if status_column is not None:
            for i, row in enumerate(rows, first_data):
                value = row[status_column]
                if value in palette:
                    commands.append(("BACKGROUND", (status_column, i), (status_column, i),
                                     colors.HexColor(palette[value][0])))
        result.setStyle(TableStyle(commands))
        return result

    elements = [p("BÁO CÁO HỒ SƠ KHOA HỌC & CÔNG NGHỆ", title_style),
                p("Phạm vi: " + filter_desc, small),
                p("Thời điểm xuất: " + stamp, small), Spacer(1, 14)]
    labels = ["TỔNG HỒ SƠ", "CHỜ XÉT DUYỆT", "ĐÃ DUYỆT", "TỶ LỆ ĐÃ DUYỆT"]
    values = [str(stats["total"]), str(dict(stats["states"])["Chờ xét duyệt"]),
              str(stats["approved"]), f'{stats["approval_rate"]:.2%}']
    card_colors = [navy, colors.HexColor("#A56B08"), colors.HexColor("#13845B"), teal]
    card_cells = []
    for label, value, color in zip(labels, values, card_colors):
        label_style = ParagraphStyle("CardLabel", parent=small, fontName=bold, textColor=color)
        value_style = ParagraphStyle("CardValue", parent=body, fontName=bold,
                                    fontSize=25, leading=32, textColor=color)
        card_cells.append([p(label, label_style), Spacer(1, 5), p(value, value_style)])
    cards = Table([card_cells], colWidths=[doc.width / 4] * 4)
    cards.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), pale),
                              ("BOX", (0, 0), (-1, -1), .4, colors.HexColor("#DBE4ED")),
                              ("VALIGN", (0, 0), (-1, -1), "TOP"),
                              ("TOPPADDING", (0, 0), (-1, -1), 12),
                              ("BOTTOMPADDING", (0, 0), (-1, -1), 12)]))
    elements.extend([cards, Spacer(1, 14)])

    # Vector charts remain crisp in print. Counts and percentages are explicit.
    drawing = Drawing(doc.width, 178)
    half = doc.width / 2
    for x, title in [(0, "01  TIẾN ĐỘ XÉT DUYỆT"), (half + 12, "02  CƠ CẤU CÔNG TRÌNH")]:
        drawing.add(String(x, 164, title, fontName=bold, fontSize=10, fillColor=navy))
    for offset, pairs in [(0, stats["states"]), (half + 12, [
        ["Bài báo", stats["articles"]], ["Đề tài", stats["projects"]],
        ["Công trình khác", stats["total"] - stats["articles"] - stats["projects"]],
    ])]:
        maximum = max([count for _, count in pairs] + [1])
        for i, (label, count) in enumerate(pairs):
            y = 137 - i * 27
            color = colors.HexColor(palette[label][1]) if label in palette else teal
            drawing.add(String(offset, y, label, fontName=regular, fontSize=8.5, fillColor=navy))
            bar_x = offset + 108
            width = (half - 190) * count / maximum
            drawing.add(Rect(bar_x, y - 3, half - 190, 11, fillColor=pale, strokeColor=None))
            if count:
                drawing.add(Rect(bar_x, y - 3, width, 11, fillColor=color, strokeColor=None))
            percent = count / stats["total"] if stats["total"] else 0
            drawing.add(String(offset + half - 73, y, f"{count} | {percent:.1%}",
                               fontName=bold, fontSize=8, fillColor=color))
    elements.extend([drawing, p("03  KINH PHÍ ĐỀ TÀI", heading),
                     p(f"Tổng kinh phí đã kê khai: {stats['funding']:,.0f} đồng. "
                       f"Chưa kê khai: {stats['missing_funding']} đề tài. "
                       f"Lưu bằng 0: {stats['zero_funding']} đề tài."),
                     Spacer(1, 8),
                     p("Tỷ lệ đã duyệt = số đã duyệt / tổng hồ sơ trong phạm vi lọc. "
                       "Phê duyệt tính hợp lệ không đồng nghĩa với nghiệm thu đề tài.", small)])

    review = review_support_statistics(records)
    elements.extend([
        PageBreak(),
        p("KẾT QUẢ HỖ TRỢ XÉT DUYỆT", title_style),
        table(["Chỉ tiêu", "Số hồ sơ"], review["metrics"], [.78, .22]),
        Spacer(1, 12),
        table(["Nhóm lý do", "Số hồ sơ"], review["categories"], [.78, .22],
              caption="Lý do của các cảnh báo đang bật"),
        Spacer(1, 9),
        p(REVIEW_NOTE, small),
        Spacer(1, 12),
    ])
    priority_rows = [
        [code, title, status, f"{score:.2f}" if score is not None else "", reasons]
        for code, title, status, score, reasons in review["priority"]
    ]
    elements.append(table(
        REVIEW_HEADERS, priority_rows, [.13, .27, .15, .10, .35],
        status_column=2, caption="Hồ sơ cần ưu tiên xem xét",
    ))

    elements.extend([PageBreak(), p("THỐNG KÊ THEO NHÓM", title_style)])
    for label, rows in [("Theo đơn vị chủ trì", stats["units"]),
                        ("Theo lĩnh vực đang lưu", stats["fields"])]:
        elements.append(table(["Phân nhóm", "Số hồ sơ", "Đã duyệt", "Tỷ lệ duyệt"],
                              [[name, total, approved, f"{rate:.1%}"]
                               for name, total, approved, rate in rows], [.58, .14, .14, .14], caption=label))
        elements.append(Spacer(1, 12))
    for label, rows in stats["sections"]:
        if not rows:
            continue
        elements.extend([table(["Phân nhóm", "Số hồ sơ"], rows, [.85, .15], caption=label), Spacer(1, 12)])

    for group in ["Bài báo", "Đề tài", "Khác"]:
        subset = [r for r in records if group_name(r) == group]
        if not subset:
            continue
        elements.extend([PageBreak(), p("DANH SÁCH " + group.upper(), title_style),
                         p(f"{len(subset)} hồ sơ trong phạm vi báo cáo", small), Spacer(1, 10)])
        rows = []
        for r in subset:
            lines = [text(r.get("tieuDe")) or "Chưa kê khai tên công trình"]
            authors = r.get("thongTinTacGia")
            if authors:
                lines.append("Tác giả: " + text(authors))
            if group != "Khác":
                missing = []
                for key, label in detail_columns(group):
                    if key in ("maHoSo", "tieuDe", "donViChuTri", "trangThai", "kyBaoCao"):
                        continue
                    value = r.get(key)
                    if value is None or value == "":
                        missing.append(label)
                    else:
                        lines.append(label + ": " + text(value))
                if missing:
                    lines.append("Chưa kê khai: " + ", ".join(missing))
            rows.append([r.get("maHoSo", ""), "\n".join(lines),
                         (text(r.get("donViChuTri")) or "Chưa kê khai đơn vị")
                         + "\n" + text(r.get("kyBaoCao")), r.get("trangThai", "")])
        elements.append(table(["Mã hồ sơ", "Công trình và thông tin chi tiết",
                               "Đơn vị chủ trì / Kỳ báo cáo", "Trạng thái xét duyệt"],
                              rows, [.13, .49, .23, .15], status_column=3))
    if not records:
        elements.append(p("Không có hồ sơ phù hợp với phạm vi lọc.", heading))

    def page_header(canvas, document):
        canvas.saveState()
        width, height = landscape(A4)
        canvas.setFillColor(navy)
        canvas.rect(0, height - 30, width, 30, fill=1, stroke=0)
        canvas.setFillColor(white)
        canvas.setFont(bold, 9)
        canvas.drawString(32, height - 19, "HUIT RESEARCH  |  QUẢN LÝ KHOA HỌC & CÔNG NGHỆ")
        canvas.setFont(regular, 8)
        canvas.drawRightString(width - 32, height - 19, stamp)
        canvas.restoreState()

    class NumberedCanvas(Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._report_pages = []

        def showPage(self):
            self._report_pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._report_pages)
            for state in self._report_pages:
                self.__dict__.update(state)
                self.saveState()
                self.setStrokeColor(colors.HexColor("#DBE4ED"))
                self.line(32, 28, landscape(A4)[0] - 32, 28)
                self.setFillColor(muted)
                self.setFont(regular, 8)
                self.drawString(32, 16, "Báo cáo hồ sơ KH&CN | Theo phạm vi lọc tại thời điểm xuất")
                self.drawRightString(landscape(A4)[0] - 32, 16, f"Trang {self._pageNumber} / {total}")
                self.restoreState()
                super().showPage()
            super().save()

    doc.build(elements, onFirstPage=page_header, onLaterPages=page_header, canvasmaker=NumberedCanvas)
    output.seek(0)
    return output


# =========================================================
# BỘ LỌC DỮ LIỆU
# =========================================================

def positive_integer(value, label):
    try:
        result = int(value)

        if result < 1:
            raise ValueError()

        return result

    except (TypeError, ValueError):
        abort(
            400,
            description=label + " không hợp lệ.",
        )


def filtered_records():
    # Import tại đây để tránh phụ thuộc vòng khi nạp app.py.
    from database import get_db_connection

    filters = {
        key: request.args.get(key, "").strip()
        for key in FILTER_KEYS
    }

    conditions = ["h.isDeleted = 0"]
    parameters = []
    descriptions = []

    group = filters["nhom"]

    if group not in ("", "baibao", "detai", "khac"):
        abort(400, description="Nhóm công trình không hợp lệ.")

    article_condition = (
        "(lct.nhomCongTrinh LIKE '%Bài báo%' "
        "OR lct.tenLoai LIKE '%Bài báo%')"
    )

    project_condition = (
        "(lct.nhomCongTrinh LIKE '%Đề tài%' "
        "OR lct.tenLoai LIKE '%Đề tài%')"
    )

    if group:
        if group == "baibao":
            conditions.append(article_condition)
        elif group == "detai":
            conditions.append(project_condition)
        else:
            conditions.append(
                f"(NOT {article_condition} "
                f"AND NOT {project_condition})"
            )

        descriptions.append(
            "Nhóm: "
            + {
                "baibao": "Bài báo",
                "detai": "Đề tài",
                "khac": "Khác",
            }[group]
        )

    for key, column, label in [
        ["nam", "h.namCongBo", "Năm công bố/kê khai"],
        ["maLinhVuc", "h.maLinhVuc", "Mã lĩnh vực"],
        ["maLoai", "h.maLoai", "Mã loại công trình"],
    ]:
        if filters[key]:
            number = positive_integer(
                filters[key], label
            )

            conditions.append(column + " = ?")
            parameters.append(number)
            descriptions.append(f"{label}: {number}")

    for key, column, label in [
        ["donVi", "h.donViChuTri", "Đơn vị chủ trì"],
        ["trangThai", "h.trangThai", "Trạng thái"],
        ["kyBaoCao", "h.kyBaoCao", "Kỳ báo cáo"],
    ]:
        if filters[key]:
            conditions.append(column + " = ?")
            parameters.append(filters[key])
            descriptions.append(
                label + ": " + filters[key]
            )

    if filters["sanPhamDuKien"]:
        conditions.append("h.sanPhamDuKien LIKE ?")
        parameters.append(
            "%" + filters["sanPhamDuKien"] + "%"
        )
        descriptions.append(
            "Sản phẩm dự kiến: " + filters["sanPhamDuKien"]
        )

    if filters["chiMuc"]:
        conditions.append("h.chiMucTapChi LIKE ?")

        parameters.append(
            "%" + filters["chiMuc"] + "%"
        )

        descriptions.append(
            "Chỉ mục tạp chí: " + filters["chiMuc"]
        )

    parsed_dates = {}

    for key in ["tuNgay", "denNgay"]:
        if filters[key]:
            try:
                parsed_dates[key] = date.fromisoformat(
                    filters[key]
                )
            except ValueError:
                abort(
                    400,
                    description="Ngày không hợp lệ.",
                )

    if (
        "tuNgay" in parsed_dates
        and "denNgay" in parsed_dates
        and parsed_dates["tuNgay"] > parsed_dates["denNgay"]
    ):
        abort(
            400,
            description="Ngày bắt đầu phải trước ngày kết thúc.",
        )

    if "tuNgay" in parsed_dates:
        conditions.append("h.ngayNop >= ?")
        parameters.append(
            parsed_dates["tuNgay"].isoformat()
        )
        descriptions.append(
            "Ngày nộp từ: " + filters["tuNgay"]
        )

    if "denNgay" in parsed_dates:
        try:
            next_day = (
                parsed_dates["denNgay"]
                + timedelta(days=1)
            )
        except OverflowError:
            abort(400, description="Ngày kết thúc quá lớn.")

        conditions.append("h.ngayNop < ?")
        parameters.append(next_day.isoformat())
        descriptions.append(
            "Ngày nộp đến: " + filters["denNgay"]
        )

    sql = """
        SELECT
            h.*,
            lv.tenLinhVuc,
            lct.tenLoai,
            lct.nhomCongTrinh,
            tc.tenTapChi,
            tc.chiSoISSN,
            tc.xepHang,
            ai.tenLinhVuc AS tenLinhVucAI
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv
            ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LoaiCongTrinh lct
            ON h.maLoai = lct.maLoai
        LEFT JOIN TapChiHopLe tc
            ON h.maTapChi = tc.maTapChi
        LEFT JOIN LinhVuc ai
            ON h.maLinhVucAI = ai.maLinhVuc
        WHERE
    """

    sql += " AND ".join(conditions)
    sql += " ORDER BY h.maHoSo"

    connection = get_db_connection()

    try:
        rows = connection.execute(
            sql, parameters
        ).fetchall()

        return (
            [dict(row) for row in rows],
            "; ".join(descriptions)
            or "Toàn bộ hồ sơ KH&CN",
            filters,
        )

    finally:
        connection.close()


# =========================================================
# GIAO DIỆN BÁO CÁO TRONG CHÍNH TỆP PYTHON
# =========================================================

REPORT_TEMPLATE = """
<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Báo cáo & Thống kê</title>
<style>
*{box-sizing:border-box}
body{margin:0;background:#f2f5f9;color:#243447;
font-family:Arial,sans-serif;font-size:14px}
main{max-width:1400px;margin:auto;padding:24px}
h1,h2,h3{color:#17365d}
.panel{background:white;padding:22px;border:1px solid #dce3eb;
border-radius:10px;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));
gap:14px}
label{display:block;margin-bottom:6px;font-weight:bold}
input,select{width:100%;padding:10px;border:1px solid #cbd5e1;
border-radius:6px;background:white}
.btn{display:inline-block;padding:10px 16px;background:#17365d;
color:white;text-decoration:none;border:0;border-radius:6px;
cursor:pointer;margin:4px 4px 4px 0}
.secondary{background:#64748b}.green{background:#187447}
.red{background:#a52d38}
.note{color:#64748b;font-size:13px;line-height:1.6}
.kpi{font-size:30px;font-weight:bold;color:#17365d;margin-top:8px}
.scroll{overflow:auto}
table{border-collapse:collapse;width:100%;font-size:14px}
td,th{padding:11px;border-bottom:1px solid #e2e8f0;
text-align:left;vertical-align:top}
th{background:#edf2f7}
tr:nth-child(even){background:#f8fafc}
.bar{height:12px;background:#17365d;border-radius:3px}
.bar-label{margin:12px 0 5px}
.actions{margin-top:16px}
a{color:#17365d}
</style>
</head>
<body>
<main>
<a href="{{ url_for('dashboard') }}">← Về Dashboard</a>
<h1>Báo cáo & Thống kê</h1>
<p class="note">Các bảng, biểu đồ và file xuất dùng cùng bộ lọc.
Hồ sơ đã xóa mềm không được tính.</p>

<form class="panel" method="get">
<div class="grid">
<div>
<label>Nhóm công trình</label>
<select name="nhom">
{% for value,label in [
('','Tất cả'),('baibao','Bài báo'),
('detai','Đề tài'),('khac','Khác')] %}
<option value="{{ value }}"
{% if filters.nhom==value %}selected{% endif %}>
{{ label }}</option>
{% endfor %}
</select>
</div>

<div>
<label>Năm công bố/kê khai</label>
<input type="number" min="1" max="9999"
name="nam" value="{{ filters.nam }}">
</div>

<div><label>Ngày nộp từ</label>
<input type="date" name="tuNgay" value="{{ filters.tuNgay }}">
</div>

<div><label>Ngày nộp đến</label>
<input type="date" name="denNgay" value="{{ filters.denNgay }}">
</div>

<div>
<label>Đơn vị</label>
<select name="donVi">
<option value="">Tất cả</option>
{% for value in units %}
<option {% if filters.donVi==value %}selected{% endif %}>
{{ value }}</option>
{% endfor %}
</select>
</div>

<div>
<label>Lĩnh vực</label>
<select name="maLinhVuc">
<option value="">Tất cả</option>
{% for value in fields %}
<option value="{{ value.maLinhVuc }}"
{% if filters.maLinhVuc==value.maLinhVuc|string %}selected{% endif %}>
{{ value.tenLinhVuc }}</option>
{% endfor %}
</select>
</div>

<div>
<label>Loại công trình</label>
<select name="maLoai">
<option value="">Tất cả</option>
{% for value in types %}
<option value="{{ value.maLoai }}"
{% if filters.maLoai==value.maLoai|string %}selected{% endif %}>
{{ value.tenLoai }}</option>
{% endfor %}
</select>
</div>

<div>
<label>Trạng thái hợp lệ</label>
<select name="trangThai">
<option value="">Tất cả</option>
{% for value in states %}
<option {% if filters.trangThai==value %}selected{% endif %}>
{{ value }}</option>
{% endfor %}
</select>
</div>

{% for key,label in [
('kyBaoCao','Kỳ báo cáo'),
('chiMuc','Chỉ mục tạp chí'),
('sanPhamDuKien','Sản phẩm dự kiến')
] %}
<div>
<label>{{ label }}</label>
<input name="{{ key }}" value="{{ filters[key] }}">
</div>
{% endfor %}
</div>

<p class="note">
Năm lọc dùng trường năm công bố/kê khai hiện có,
không phải năm thực hiện đề tài.
Kỳ, cấp và nghiệm thu lọc theo giá trị chính xác đã lưu.
Tác giả và chỉ mục tìm theo một phần nội dung.
</p>

<button class="btn" type="submit">Áp dụng</button>
<a class="btn secondary" href="{{ url_for('reports.index') }}">
Xóa bộ lọc</a>
</form>

<div class="panel">
<strong>Phạm vi:</strong> {{ description }}
<div class="actions">
<a class="btn green"
href="{{ url_for('reports.download',kind='excel') }}?{{ query }}">
Tải Excel</a>
<a class="btn red"
href="{{ url_for('reports.download',kind='pdf') }}?{{ query }}">
Tải PDF</a>
</div>
</div>

<div class="grid">
{% for label,value in [
('Tổng hồ sơ',stats.total),
('Bài báo',stats.articles),
('Đề tài',stats.projects),
('Đã duyệt',stats.approved)] %}
<div class="panel">
{{ label }}
<div class="kpi">{{ value }}</div>
</div>
{% endfor %}
</div>

<div class="panel">
<h2>Kết quả kiểm tra tính hợp lệ</h2>
<p>Tỷ lệ đã duyệt:
<strong>{{ '%.1f'|format(stats.approval_rate*100) }}%</strong>
trên tổng hồ sơ trong phạm vi lọc.</p>

{% for label,count in stats.states %}
<div class="bar-label">{{ label }}: {{ count }}</div>
<div class="bar"
style="width:{{ count/stats.total*100 if stats.total else 0 }}%">
</div>
{% endfor %}
</div>

{% for title,rows in [
('Theo đơn vị',stats.units),
('Theo lĩnh vực đang lưu',stats.fields)
] %}
<div class="panel">
<h2>{{ title }}</h2>
<div class="scroll">
<table>
<thead><tr>
<th>Phân nhóm</th><th>Tổng</th>
<th>Đã duyệt</th><th>Tỷ lệ duyệt trong nhóm</th>
</tr></thead>
<tbody>
{% for label,total,approved,rate in rows %}
<tr>
<td>{{ label }}</td><td>{{ total }}</td>
<td>{{ approved }}</td>
<td>{{ '%.1f'|format(rate*100) }}%</td>
</tr>
{% else %}
<tr><td colspan="4">Không có dữ liệu.</td></tr>
{% endfor %}
</tbody>
</table>
</div>
</div>
{% endfor %}

<div class="grid">
{% for title,rows in stats.sections %}
<div class="panel">
<h3>{{ title }}</h3>
<table>
{% for label,count in rows %}
<tr><td>{{ label }}</td><td>{{ count }}</td></tr>
{% else %}
<tr><td>Không có dữ liệu.</td></tr>
{% endfor %}
</table>
</div>
{% endfor %}
</div>

<div class="panel">
<h2>Kinh phí đề tài</h2>
<p>Tổng kinh phí đã kê khai:
<strong>{{ '{:,.0f}'.format(stats.funding) }} đồng</strong>.</p>
<p class="note">
{{ stats.missing_funding }} đề tài chưa có kinh phí;
{{ stats.zero_funding }} đề tài lưu bằng 0.
Giá trị 0 được giữ nguyên, không tự coi là thiếu.
</p>
<p class="note">
Trạng thái duyệt tính hợp lệ không đồng nghĩa nghiệm thu.
Thống kê tác giả dùng nguyên trường tác giả đã lưu.
Cơ sở dữ liệu hiện chưa có tiến độ và năm thực hiện riêng.
</p>
</div>

<div class="panel">
<h2>Danh sách hồ sơ</h2>
<div class="scroll">
<table>
<thead>
<tr><th>Mã</th><th>Tên công trình</th>
<th>Tác giả</th><th>Đơn vị</th><th>Trạng thái</th></tr>
</thead>
<tbody>
{% for record in records %}
<tr>
<td>{{ record.maHoSo }}</td>
<td>{{ record.tieuDe }}</td>
<td>{{ record.thongTinTacGia }}</td>
<td>{{ record.donViChuTri }}</td>
<td>{{ record.trangThai }}</td>
</tr>
{% else %}
<tr><td colspan="5">Không có hồ sơ phù hợp.</td></tr>
{% endfor %}
</tbody>
</table>
</div>

<div class="actions">
Trang {{ page }}/{{ pages }}
{% if page>1 %}
<a class="btn secondary"
href="?{{ query }}&page={{ page-1 }}">Trước</a>
{% endif %}
{% if page<pages %}
<a class="btn secondary"
href="?{{ query }}&page={{ page+1 }}">Sau</a>
{% endif %}
</div>
</div>
</main>
</body>
</html>
"""


# =========================================================
# ĐĂNG KÝ ROUTE
# =========================================================

def staff_only(function):
    @wraps(function)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))

        if session.get("vai_tro") not in ("Admin", "CanBo"):
            abort(403)

        return function(*args, **kwargs)

    return wrapped


def register_reports(app):
    # Tránh đăng ký trùng khi người dùng gọi lại.
    if "reports" in app.blueprints:
        return

    reports = Blueprint("reports", __name__)

    @reports.route("/bao-cao")
    @staff_only
    def index():
        from database import get_db_connection

        records, description, filters = filtered_records()

        page = positive_integer(
            request.args.get("page", "1"),
            "Trang",
        )

        pages = max(1, (len(records) + 19) // 20)
        page = min(page, pages)

        connection = get_db_connection()

        try:
            fields = [
                dict(row)
                for row in connection.execute("""
                    SELECT maLinhVuc, tenLinhVuc
                    FROM LinhVuc
                    ORDER BY tenLinhVuc
                """).fetchall()
            ]

            types = [
                dict(row)
                for row in connection.execute("""
                    SELECT maLoai, tenLoai
                    FROM LoaiCongTrinh
                    ORDER BY tenLoai
                """).fetchall()
            ]

            units = [
                row[0]
                for row in connection.execute("""
                    SELECT DISTINCT donViChuTri
                    FROM HoSoKHCN
                    WHERE isDeleted = 0
                    ORDER BY donViChuTri
                """).fetchall()
                if row[0]
            ]

        finally:
            connection.close()

        query = urlencode({
            key: value
            for key, value in filters.items()
            if value
        })

        return render_template_string(
            REPORT_TEMPLATE,
            records=records[(page - 1) * 20:page * 20],
            stats=report_statistics(records),
            description=description,
            filters=filters,
            fields=fields,
            types=types,
            units=units,
            states=STATES,
            page=page,
            pages=pages,
            query=query,
        )

    @reports.route("/bao-cao/tai/<kind>")
    @staff_only
    def download(kind):
        if kind not in ("excel", "pdf"):
            abort(404)

        records, description, _ = filtered_records()

        if kind == "excel":
            stream = generate_excel_report(
                records, description
            )

            extension = "xlsx"

            mimetype = (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )

        else:
            stream = generate_pdf_report(
                records, description
            )

            extension = "pdf"
            mimetype = "application/pdf"

        filename = (
            "BaoCao_KHCN_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + "."
            + extension
        )

        return send_file(
            stream,
            as_attachment=True,
            download_name=filename,
            mimetype=mimetype,
        )

    app.register_blueprint(reports)

    # Giữ các nút xuất cũ của ứng dụng.
    if "export_excel" in app.view_functions:
        app.view_functions["export_excel"] = (
            lambda: download("excel")
        )

    if "export_pdf" in app.view_functions:
        app.view_functions["export_pdf"] = (
            lambda: download("pdf")
        )
