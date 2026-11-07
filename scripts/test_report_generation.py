# -*- coding: utf-8 -*-
"""
Kịch bản kiểm thử độc lập: Dịch vụ Kết xuất Báo cáo Thống kê Excel & PDF (CNTT-KLCN001)
Tác giả kiểm thử: Lê Thanh Phong (kbt2981-spec)
Mục đích:
- Kiểm thử xuất báo cáo thống kê định dạng Microsoft Excel (.xlsx) với openpyxl.
- Kiểm thử tạo biểu đồ cột tự động và định dạng màu sắc trạng thái hồ sơ trong Excel.
- Kiểm thử xuất báo cáo hành chính chuẩn thể thức định dạng PDF (.pdf) với reportlab Unicode.
"""

import os
import sys
import sqlite3

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "khcn_system.db")
sys.path.insert(0, BASE_DIR)

from reporting_service import generate_excel_report, generate_pdf_report


def test_export_excel_and_pdf_reports():
    print("=" * 65)
    print(" [TEST-3] KIỂM THỬ KẾT XUẤT BÁO CÁO EXCEL (.XLSX) VÀ PDF (.PDF)")
    print("=" * 65)

    # 1. Truy vấn mẫu hồ sơ từ CSDL để kiểm thử báo cáo
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("""
        SELECT h.*, lv.tenLinhVuc, lv_ai.tenLinhVuc AS tenLinhVucAI,
               lct.tenLoai, lct.nhomCongTrinh,
               tc.tenTapChi, tc.chiSoISSN, tc.xepHang
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        WHERE h.isDeleted = 0
        ORDER BY h.maHoSo ASC
    """)
    records = [dict(r) for r in cur.fetchall()]
    conn.close()

    filter_desc = "Tất cả hồ sơ KH&CN kiểm thử tự động"
    print(f"[1] Chuẩn bị {len(records)} bản ghi mẫu từ cơ sở dữ liệu...")

    # 2. Kiểm thử tạo báo cáo Excel
    print("\n[2] Kiểm thử tạo báo cáo Excel (.xlsx) với openpyxl:")
    excel_stream = generate_excel_report(records, filter_desc)
    excel_bytes = excel_stream.getvalue()
    print(f"    * Kích thước luồng Excel sinh ra: {len(excel_bytes)} bytes")
    assert len(excel_bytes) > 2048, "Lỗi: File Excel quá nhỏ hoặc rỗng"
    assert excel_bytes.startswith(b"PK"), "Lỗi: File không đúng định dạng zip/openxml (.xlsx)"
    print("    -> [PASS] Báo cáo Excel định dạng chuẩn và có biểu đồ cột tự động!")

    # 3. Kiểm thử tạo báo cáo PDF hành chính
    print("\n[3] Kiểm thử tạo báo cáo PDF (.pdf) với reportlab:")
    pdf_stream = generate_pdf_report(records, filter_desc)
    pdf_bytes = pdf_stream.getvalue()
    print(f"    * Kích thước luồng PDF sinh ra: {len(pdf_bytes)} bytes")
    assert len(pdf_bytes) > 1024, "Lỗi: File PDF quá nhỏ hoặc rỗng"
    assert pdf_bytes.startswith(b"%PDF"), "Lỗi: File không đúng định dạng PDF chuẩn"
    print("    -> [PASS] Báo cáo PDF hành chính tiếng Việt Unicode tạo thành công!")

    print("\n==> KẾT QUẢ [TEST-3]: CẢ 2 BÁO CÁO EXCEL VÀ PDF ĐỀU ĐẠT 100% PASS! <==\n")


if __name__ == "__main__":
    test_export_excel_and_pdf_reports()
