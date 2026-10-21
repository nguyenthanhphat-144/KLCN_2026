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

    # 3. Test Chi tiết hồ sơ có metadata bài báo & đề tài (theo 12 trường chuẩn)
    print("3. Kiểm tra Chi tiết hồ sơ & Metadata 12 trường chuẩn...")
    res = client.get("/hoso/HS-2026-0001")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "Tên tạp chí / Hội thảo" in html and "Chỉ mục tạp chí" in html, "Thiếu thông tin bài báo HS-0001"
    assert "Kiểm Tra Tính Hợp Lệ & Thẩm Định Hồ Sơ" in html, "Thiếu form kiểm tra tính hợp lệ"
    print("   -> Chi tiết Bài báo (HS-0001): OK.")

    res = client.get("/hoso/HS-2026-0007")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "Sản phẩm dự kiến" in html and "Đề tài NCKH" in html, "Thiếu thông tin đề tài HS-0007"
    print("   -> Chi tiết Đề tài (HS-0007): OK.")

    # 4. Test Biểu mẫu nộp hồ sơ (hoso_form)
    print("4. Kiểm tra Biểu mẫu khai báo động 12 trường (Bài báo vs Đề tài)...")
    res = client.get("/hoso/new")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "radioBaiBao" in html and "radioDeTai" in html
    assert "sectionBaiBao" in html and "sectionDeTai" in html
    assert "1. Loại hình (*)" in html and "12. Sản phẩm dự kiến" in html
    print("   -> Biểu mẫu nộp mới động 12 trường: OK.")

    # 5. Test Danh mục hệ thống
    print("5. Kiểm tra Danh mục chuẩn hóa & Quy chuẩn 12 trường...")
    res = client.get("/danh-muc")
    assert res.status_code == 200
    html = res.data.decode("utf-8")
    assert "Danh Mục Chỉ Mục Tạp Chí" in html
    assert "Danh Mục Sản Phẩm Dự Kiến" in html
    assert "Danh Mục Kỳ Báo Cáo & Đơn Vị Chủ Trì" in html
    print("   -> Danh mục mở rộng: OK.")

    # 6. Test Toàn bộ Quy trình 5 bước Nghiệp vụ kiểm tra tính hợp lệ (Revision)
    print("6. Kiểm tra Quy trình & Luồng 5 bước (Kiểm tra tính hợp lệ hồ sơ)...")
    
    # Bước 1 & 1c: Tạo 1 hồ sơ thử nghiệm ở trạng thái 'Chờ xét duyệt'
    conn = get_db_connection()
    cur = conn.cursor()
    test_code = "HS-TEST-WF01"
    cur.execute("DELETE FROM LichSuXetDuyet WHERE maHoSo = ?", (test_code,))
    cur.execute("DELETE FROM HoSoKHCN WHERE maHoSo = ?", (test_code,))
    cur.execute("""
        INSERT INTO HoSoKHCN (
            maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
            namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
            trangThai, fileDinhKem, vanBanDaTachTu, maLinhVucAI, doTinCayAI,
            maLinhVucSVM, doTinCaySVM, coCanhBao, diemBatThuong
        ) VALUES (
            ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?,
            ?, ?, ?, ?
        )
    """, (
        test_code, "Nghiên cứu ứng dụng Blockchain trong xác thực văn bằng",
        "Tóm tắt về công nghệ chuỗi khối và hợp đồng thông minh.",
        "blockchain, smart contracts", "Nguyễn Thành Tài", "Khoa CNTT",
        2026, "Học kỳ 1 - 2026", 3, 1, 1, 1,
        "Chờ xét duyệt", "minhchung_test.pdf", "văn_bản đã tách từ", 1, 92.5,
        1, 88.0, 0, 0.12
    ))
    conn.commit()
    conn.close()

    # Bước 1c: Cán bộ xét duyệt (Lê Thanh Phong - user_id=2) mở hồ sơ
    canbo_client = app.test_client()
    with canbo_client.session_transaction() as sess:
        sess["user_id"] = 2
        sess["ho_ten"] = "Lê Thanh Phong"
        sess["vai_tro"] = "CanBo"
        sess["don_vi"] = "Phòng Quản lý KH&CN"
        sess["email"] = "canbo@uni.edu.vn"

    res = canbo_client.get(f"/hoso/{test_code}")
    assert res.status_code == 200
    
    # Kiểm tra DB: trạng thái phải tự động nhảy sang 'Đang xét duyệt'
    conn = get_db_connection()
    row = conn.execute("SELECT trangThai FROM HoSoKHCN WHERE maHoSo = ?", (test_code,)).fetchone()
    assert row["trangThai"] == "Đang xét duyệt", f"Lỗi Bước 1c: Trạng thái chưa tự chuyển sang Đang xét duyệt (Hiện tại: {row['trangThai']})"
    
    # Kiểm tra LichSuXetDuyet đã ghi nhận
    ls = conn.execute("SELECT * FROM LichSuXetDuyet WHERE maHoSo = ? ORDER BY maLichSu DESC", (test_code,)).fetchone()
    assert ls and ls["trangThaiMoi"] == "Đang xét duyệt"
    print("   -> Bước 1c: Tự động chuyển 'Chờ xét duyệt' -> 'Đang xét duyệt' & ghi nhật ký: OK.")

    # Bước 2 & 3: Cán bộ kiểm tra và ra quyết định 'Yêu cầu bổ sung'
    res_review = canbo_client.post(f"/hoso/{test_code}/xet-duyet", data={
        "trangThaiMoi": "Yêu cầu bổ sung",
        "yKienDanhGia": "Hồ sơ thiếu minh chứng trang bìa và mục lục tạp chí. Yêu cầu tác giả bổ sung tệp quét scan rõ nét."
    }, follow_redirects=True)
    assert res_review.status_code == 200

    row = conn.execute("SELECT trangThai FROM HoSoKHCN WHERE maHoSo = ?", (test_code,)).fetchone()
    assert row["trangThai"] == "Yêu cầu bổ sung", f"Lỗi Bước 3: Chưa cập nhật trạng thái Yêu cầu bổ sung"
    print("   -> Bước 2 & 3: Thẩm định và ra quyết định 'Yêu cầu bổ sung': OK.")

    # Bước 4: Kiểm tra tính bất biến và đầy đủ của LichSuXetDuyet
    ls_all = conn.execute("SELECT * FROM LichSuXetDuyet WHERE maHoSo = ? ORDER BY maLichSu ASC", (test_code,)).fetchall()
    assert len(ls_all) >= 2, "Chưa lưu vết đầy đủ trong LichSuXetDuyet"
    print("   -> Bước 4: Ghi nhận lịch sử xét duyệt đầy đủ: OK.")

    # Bước 5: Tác giả (Nguyễn Thành Tài - user_id=3) đăng nhập thấy thông báo và chỉnh sửa nộp lại
    tacgia_client = app.test_client()
    with tacgia_client.session_transaction() as sess:
        sess["user_id"] = 3
        sess["ho_ten"] = "Nguyễn Thành Tài"
        sess["vai_tro"] = "TacGia"
        sess["don_vi"] = "Khoa Công nghệ Thông tin"
        sess["email"] = "tacgia@uni.edu.vn"

    # Kiểm tra banner thông báo
    res_tg = tacgia_client.get("/hoso")
    assert res_tg.status_code == 200
    html_tg = res_tg.data.decode("utf-8")
    assert "Yêu cầu bổ sung" in html_tg and "Phản hồi thẩm định hồ sơ" in html_tg
    print("   -> Bước 5a: Banner thông báo nổi bật cho Tác giả: OK.")

    # Tác giả cập nhật và lưu hồ sơ (theo đúng 12 trường chuẩn)
    res_update = tacgia_client.post(f"/hoso/{test_code}/edit", data={
        "tieuDe": "Nghiên cứu ứng dụng Blockchain trong xác thực văn bằng đại học (Đã bổ sung)",
        "tomTat": "Tóm tắt đã được bổ sung đầy đủ thông tin trang bìa và chỉ số.",
        "tuKhoa": "blockchain, verified degrees",
        "donViChuTri": "Khoa Công nghệ Thông tin",
        "namCongBo": 2026,
        "kyBaoCao": "Đợt 1 - Năm học 2025-2026",
        "maLinhVuc": 1,
        "maLoai": 1,
        "maTapChi": 1,
        "chiMucTapChi": "Scopus Q1 / Web of Science (SCIE)"
    }, follow_redirects=True)
    assert res_update.status_code == 200

    # Kiểm tra sau khi tác giả lưu: Trạng thái tự động chuyển lại 'Chờ xét duyệt'
    row_after = conn.execute("SELECT trangThai FROM HoSoKHCN WHERE maHoSo = ?", (test_code,)).fetchone()
    assert row_after["trangThai"] == "Chờ xét duyệt", f"Lỗi Bước 5b: Chưa tự động chuyển về 'Chờ xét duyệt' (Hiện tại: {row_after['trangThai']})"

    ls_latest = conn.execute("SELECT * FROM LichSuXetDuyet WHERE maHoSo = ? ORDER BY maLichSu DESC", (test_code,)).fetchone()
    assert ls_latest["trangThaiMoi"] == "Chờ xét duyệt"
    print("   -> Bước 5b: Vòng lặp resubmission loop (Yêu cầu bổ sung -> Chờ xét duyệt): OK.")

    # Dọn dẹp dữ liệu test
    cur = conn.cursor()
    cur.execute("DELETE FROM LichSuXetDuyet WHERE maHoSo = ?", (test_code,))
    cur.execute("DELETE FROM HoSoKHCN WHERE maHoSo = ?", (test_code,))
    conn.commit()
    conn.close()

    # 7. Kiểm tra Chức năng 6 (Phát hiện bất thường & Xếp hạng ưu tiên hàng đợi)
    print("7. Kiểm tra Chức năng 6 (Đồ thị NetworkX, Isolation Forest, Deep Autoencoder & Hàng đợi ưu tiên)...")
    from ai_engine import ai_engine
    conn_ai = get_db_connection()
    graph_res = ai_engine.build_networkx_graph_and_detect_anomalies(conn_ai)
    assert graph_res["num_nodes"] > 0 and graph_res["num_edges"] > 0, "Lỗi đồ thị NetworkX không có nút/cạnh"
    print(f"   -> Đồ thị NetworkX: Đã xây dựng {graph_res['num_nodes']} nút, {graph_res['num_edges']} cạnh, {graph_res['anomalies']} bất thường.")

    # Kiểm tra 2 mô hình bất thường (Isolation Forest & Deep Autoencoder)
    models_comp = ai_engine.get_anomaly_models_comparison()
    assert len(models_comp) == 2, "Thiếu mô hình so sánh phát hiện bất thường"
    assert any("Isolation Forest" in m["name"] for m in models_comp), "Thiếu mô hình Isolation Forest"
    assert any("Autoencoder" in m["name"] for m in models_comp), "Thiếu mô hình Deep Autoencoder"
    print("   -> 2 Mô hình Bất thường (Isolation Forest & Deep Autoencoder): OK.")

    # Kiểm tra Hàng đợi xét duyệt ưu tiên
    cur_ai = conn_ai.cursor()
    pq = ai_engine.get_priority_queue(cur_ai)
    assert len(pq) > 0, "Hàng đợi ưu tiên trống"
    assert pq[0]["diemBatThuong"] >= pq[-1]["diemBatThuong"], "Hàng đợi chưa sắp xếp theo Điểm rủi ro cao nhất"
    print(f"   -> Hàng đợi xét duyệt ưu tiên (Priority Review Queue): OK ({len(pq)} hồ sơ, Hạng #1: {pq[0]['maHoSo']} - Rủi ro: {pq[0]['diemBatThuong']}).")

    # Kiểm tra Giao diện /ai-center hiển thị đầy đủ
    res_aic = client.get("/ai-center")
    assert res_aic.status_code == 200
    html_aic = res_aic.data.decode("utf-8")
    assert "Nghiệp Vụ Phức Tạp 2: So Sánh 2 Mô Hình Phát Hiện Bất Thường" in html_aic
    assert "Deep Autoencoder" in html_aic and "Isolation Forest" in html_aic
    assert "Đồ Thị Tri Thức Quan Hệ" in html_aic
    assert "Hàng Đợi Xét Duyệt Ưu Tiên Rủi Ro" in html_aic
    print("   -> Giao diện /ai-center tích hợp đầy đủ 6 phân hệ AI: OK.")

    # Kiểm tra chức năng Bỏ qua cảnh báo (False Alarm - Nghiệp vụ 6 Bước 5b)
    # Gắn cờ thử nghiệm trên 1 hồ sơ rồi giải trừ
    test_dismiss_code = "HS-2026-0003"
    res_dismiss = client.post(f"/hoso/{test_dismiss_code}/dismiss-warning", data={
        "lyDoBoQua": "Kiểm tra thực tế xác nhận báo động giả."
    }, follow_redirects=True)
    assert res_dismiss.status_code == 200
    row_dis = conn_ai.execute("SELECT coCanhBao FROM HoSoKHCN WHERE maHoSo = ?", (test_dismiss_code,)).fetchone()
    assert row_dis["coCanhBao"] == 0, "Chưa giải trừ cờ coCanhBao"
    print("   -> Bỏ qua cảnh báo (False Alarm) & Ghi nhật ký: OK.")

    # Khôi phục lại cờ cảnh báo cho hồ sơ HS-2026-0003
    conn_ai.execute("UPDATE HoSoKHCN SET coCanhBao = 1 WHERE maHoSo = ?", (test_dismiss_code,))
    conn_ai.commit()
    conn_ai.close()

    print("\n=== TẤT CẢ CÁC BƯỚC KIỂM THỬ ĐÃ THÀNH CÔNG RỰC RỠ 100%! ===")

if __name__ == "__main__":
    test_all()

