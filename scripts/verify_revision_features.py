# -*- coding: utf-8 -*-
"""
Script kiểm thử tự động toàn diện các tính năng sau khi điều chỉnh theo Mô tả nghiệp vụ revision.
"""
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from app import app
from database import get_db_connection

def test_all():
    print("=== BẮT ĐẦU KIỂM THỬ HỆ THỐNG THEO NGHIỆP VỤ REVISION ===")
    client = app.test_client()

    with client.session_transaction() as sess:
        sess["user_id"] = 1
        sess["ho_ten"] = "Nguyễn Thành Phát"
        sess["vai_tro"] = "Admin"
        sess["don_vi"] = "Phòng Khoa học & Công nghệ"
        sess["email"] = "admin@uni.edu.vn"

    # 1. Test Dashboard
    print("1. Kiểm tra Dashboard & Thống kê chuyên sâu...")
    res = client.get("/dashboard")
    assert res.status_code == 200, f"Dashboard error: {res.status_code}"
    html = res.data.decode("utf-8")
    assert "Yêu cầu bổ sung" in html, "Thiếu trạng thái Yêu cầu bổ sung trên KPI"
    assert "Thống Kê Công Bố Bài Báo Khoa Học" in html, "Thiếu thống kê bài báo"
    assert "Thống Kê Đề Tài Nghiên Cứu Khoa Học" in html, "Thiếu thống kê đề tài"
    print("   -> Dashboard: OK (KPI, Bài báo, Đề tài hiển thị đầy đủ).")

    # 2. Test Danh sách hồ sơ & Tabs phân loại
    print("2. Kiểm tra Danh sách hồ sơ & Bộ lọc 2 nhóm...")
    res = client.get("/hoso")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "Bài Báo Khoa Học" in html and "Đề Tài NCKH" in html
    
    res_bb = client.get("/hoso?nhom=baibao")
    assert res_bb.status_code == 200
    print("   -> Tab Bài báo: OK.")

    res_dt = client.get("/hoso?nhom=detai")
    assert res_dt.status_code == 200
    print("   -> Tab Đề tài: OK.")

    # 3. Test Chi tiết hồ sơ có metadata bài báo & đề tài
    print("3. Kiểm tra Chi tiết hồ sơ & Metadata chuyên sâu...")
    res = client.get("/hoso/HS-2026-0001")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "Mã định danh DOI" in html or "Tập / Số" in html, "Thiếu thông tin bài báo HS-0001"
    assert "Kiểm Tra Tính Hợp Lệ & Thẩm Định Hồ Sơ" in html, "Thiếu form kiểm tra tính hợp lệ"
    print("   -> Chi tiết Bài báo (HS-0001): OK.")

    res = client.get("/hoso/HS-2026-0007")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "Mã số đề tài" in html or "Cấp quản lý" in html, "Thiếu thông tin đề tài HS-0007"
    print("   -> Chi tiết Đề tài (HS-0007): OK.")

    # 4. Test Biểu mẫu nộp hồ sơ (hoso_form)
    print("4. Kiểm tra Biểu mẫu khai báo động (Bài báo vs Đề tài)...")
    res = client.get("/hoso/new")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "radioBaiBao" in html and "radioDeTai" in html
    assert "sectionBaiBao" in html and "sectionDeTai" in html
    print("   -> Biểu mẫu nộp mới động: OK.")

    # 5. Test Danh mục hệ thống
    print("5. Kiểm tra Danh mục chuẩn hóa & Quy chuẩn bổ sung...")
    res = client.get("/danh-muc")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "Danh Mục Cấp Quản Lý Đề Tài" in html
    assert "Cơ Sở Dữ Liệu Chỉ Mục Học Thuật" in html
    assert "Đánh Giá & Nghiệm Thu Đề Tài" in html
    print("   -> Danh mục mở rộng: OK.")

    print("\n=== TẤT CẢ CÁC BƯỚC KIỂM THỬ ĐÃ THÀNH CÔNG RỰC RỠ 100%! ===")

if __name__ == "__main__":
    test_all()
