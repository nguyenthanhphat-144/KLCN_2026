# -*- coding: utf-8 -*-
"""
Script nạp bổ sung các bộ dữ liệu mẫu (Sample Inputs) chuẩn vào hệ thống KHCN 2026
Bao gồm: Bài báo khoa học, Đề tài NCKH, Hồ sơ chuẩn an toàn và các Kịch bản kiểm thử bất thường.
"""
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from database import get_db_connection
from ai_engine import ai_engine

SAMPLE_RECORDS = [
    # 1. Bài báo khoa học Quốc tế Chuẩn (CNTT - AI)
    {
        "maHoSo": "HS-SAMPLE-01",
        "tieuDe": "Tối ưu hóa phân cụm dữ liệu lớn sử dụng thuật toán bầy đàn PSO kết hợp học sâu trên cụm máy chủ Spark",
        "tomTat": "Bài báo nghiên cứu cải tiến thuật toán tối ưu hóa bầy đàn hạt kết hợp kiến trúc mạng nơ ron sâu để xử lý và phân loại các tập dữ liệu lớn phân tán trên nền tảng đám mây Apache Spark, giúp tăng tốc độ hội tụ và giảm độ trễ tính toán.",
        "tuKhoa": "dữ liệu lớn, học sâu, điện toán đám mây, trí tuệ nhân tạo, tối ưu hóa",
        "thongTinTacGia": "Nguyễn Thành Tài, Lê Minh Tuấn",
        "donViChuTri": "Khoa Công nghệ Thông tin",
        "namCongBo": 2026,
        "kyBaoCao": "Học kỳ 1 - 2026",
        "maNguoiNop": 4, # Nguyễn Thành Tài
        "maLinhVuc": 1,  # CNTT
        "maLoai": 1,     # Bài báo quốc tế
        "maTapChi": 1,   # Engineering Applications of Artificial Intelligence
        "doi": "10.1016/j.engappai.2026.109821",
        "tapSo": "Tập 142, Số 1",
        "trangBaiBao": "215-230",
        "chiMucTapChi": "Scopus Q1 / Web of Science (SCIE)",
        "trangThai": "Chờ xét duyệt",
    },

    # 2. Đề tài NCKH Cấp Bộ Chuẩn (Kỹ thuật Điện - Tự động hóa)
    {
        "maHoSo": "HS-SAMPLE-02",
        "tieuDe": "Nghiên cứu làm chủ công nghệ thiết kế và chế tạo trạm sạc nhanh thông minh công suất cao cho xe ô tô điện",
        "tomTat": "Đề tài tập trung nghiên cứu cấu trúc biến đổi nguồn DC-DC đa bậc hai chiều, thuật toán điều khiển tự động bám dòng sạc tối ưu bảo vệ pin Lithium-ion và tích hợp giao thức truyền thông tiêu chuẩn ISO 15118 cho lưới điện thông minh.",
        "tuKhoa": "điều khiển tự động, điện tử công suất, ô tô điện, lưới điện thông minh, trạm sạc",
        "thongTinTacGia": "TS. Vũ Đình Trọng, ThS. Hoàng Văn Nam",
        "donViChuTri": "Khoa Điện - Điện tử",
        "namCongBo": 2026,
        "kyBaoCao": "Học kỳ 1 - 2026",
        "maNguoiNop": 6,
        "maLinhVuc": 2,  # Kỹ thuật Điện
        "maLoai": 4,     # Đề tài cấp Bộ
        "maDeTai": "DT-2026-BO-05",
        "capDeTai": "Cấp Bộ / Nhà nước",
        "chuNhiemDeTai": "TS. Vũ Đình Trọng",
        "thanhVienThamGia": "ThS. Hoàng Văn Nam, KS. Phạm Quốc Hưng",
        "thoiGianThucHien": "01/2025 - 12/2026",
        "kinhPhi": 450000000.0,
        "sanPhamDuKien": "01 Bộ trạm sạc nhanh thử nghiệm 50kW; 02 bài báo Scopus; Báo cáo nghiệm thu.",
        "ketQuaNghiemThu": "Xuất sắc",
        "trangThai": "Chờ xét duyệt",
    },

    # 3. Kịch bản Bất thường 1: Khai sai lĩnh vực chuyên ngành (Khai vào Kinh tế nhưng nội dung Kỹ thuật Cơ khí)
    {
        "maHoSo": "HS-SAMPLE-03",
        "tieuDe": "Nghiên cứu ứng dụng hợp kim nhớ hình Nitinol và cảm biến sợi quang FBG trong giám sát kết cấu cánh quạt tuabin gió",
        "tomTat": "Công trình phân tích ứng xử cơ học mỏi biến dạng của vật liệu cơ khí và mô phỏng dao động khí động học cánh tuabin gió khi chịu tải trọng gió giật bằng phần mềm ANSYS và thực nghiệm cơ học biến dạng dẻo.",
        "tuKhoa": "chế tạo máy, phần tử hữu hạn, cơ điện tử, vật liệu cơ khí, khí động học",
        "thongTinTacGia": "TS. Hoàng Gia Huy",
        "donViChuTri": "Khoa Cơ khí Động lực",
        "namCongBo": 2026,
        "kyBaoCao": "Học kỳ 1 - 2026",
        "maNguoiNop": 8,
        "maLinhVuc": 4,  # Cố tình khai: Kinh tế & Quản trị kinh doanh (Sai lệch với PhoBERT: Cơ khí maLinhVuc=3)
        "maLoai": 1,
        "maTapChi": 4,   # International Journal of Mechanical Sciences
        "doi": "10.1016/j.ijmecsci.2026.04412",
        "tapSo": "Tập 210",
        "trangBaiBao": "88-99",
        "chiMucTapChi": "Scopus Q1 / Web of Science (SCIE)",
        "trangThai": "Chờ xét duyệt",
    },

    # 4. Kịch bản Bất thường 2: Tạp chí không hợp lệ / Ngoài danh mục uy tín (Predatory)
    {
        "maHoSo": "HS-SAMPLE-04",
        "tieuDe": "Tổng hợp nhanh các hợp chất dẫn xuất dị vòng và khảo sát hoạt tính kháng oxy hóa trên mô hình in vitro",
        "tomTat": "Nghiên cứu tổng hợp hóa học hữu cơ các dẫn xuất mới, phân tích phổ cộng hưởng từ hạt nhân NMR và kiểm tra khả năng ức chế gốc tự do phục vụ định hướng phát triển dược liệu y sinh học.",
        "tuKhoa": "tổng hợp hữu cơ, dược liệu, vật liệu nano, sinh học phân tử",
        "thongTinTacGia": "PGS.TS. Ngô Thanh Hà",
        "donViChuTri": "Khoa Kỹ thuật Hóa học & Môi trường",
        "namCongBo": 2026,
        "kyBaoCao": "Học kỳ 1 - 2026",
        "maNguoiNop": 7,
        "maLinhVuc": 5,  # Hóa học Sinh học
        "maLoai": 1,
        "maTapChi": 9,   # Global Predatory Journal of Multidisciplinary Research (uyTin = 0)
        "doi": "10.9999/predatory.2026.0123",
        "tapSo": "Tập 5, Số 1",
        "trangBaiBao": "12-18",
        "chiMucTapChi": "Tạp chí Khoa học Quốc gia",
        "trangThai": "Chờ xét duyệt",
    },

    # 5. Đề tài NCKH Cấp Cơ sở Chuẩn (Kinh tế & Quản trị kinh doanh)
    {
        "maHoSo": "HS-SAMPLE-05",
        "tieuDe": "Đánh giá các yếu tố thúc đẩy ý định sử dụng ví điện tử và dịch vụ ngân hàng số của người tiêu dùng thế hệ Z",
        "tomTat": "Nghiên cứu ứng dụng mô hình chấp nhận công nghệ mở rộng UTAUT2 và phân tích hồi quy PLS-SEM trên mẫu khảo sát 600 khách hàng trẻ nhằm đề xuất giải pháp phát triển thương mại điện tử và tài chính ngân hàng số bền vững.",
        "tuKhoa": "hành vi người tiêu dùng, tài chính ngân hàng, thương mại điện tử, kinh tế vĩ mô",
        "thongTinTacGia": "ThS. Đặng Thị Mai Hương, ThS. Nguyễn Bích Ngọc",
        "donViChuTri": "Khoa Kinh tế & Quản trị Kinh doanh",
        "namCongBo": 2026,
        "kyBaoCao": "Học kỳ 1 - 2026",
        "maNguoiNop": 5,
        "maLinhVuc": 4,  # Kinh tế
        "maLoai": 5,     # Đề tài cấp trường
        "maDeTai": "DT-2026-CS-12",
        "capDeTai": "Cấp Cơ sở / Trường",
        "chuNhiemDeTai": "ThS. Đặng Thị Mai Hương",
        "thanhVienThamGia": "ThS. Nguyễn Bích Ngọc, CN. Lê Hoàng Yến",
        "thoiGianThucHien": "03/2026 - 11/2026",
        "kinhPhi": 85000000.0,
        "sanPhamDuKien": "01 Bài báo hội thảo quốc gia; 01 báo cáo tổng kết đề tài; 01 bộ số liệu khảo sát.",
        "ketQuaNghiemThu": "Đạt",
        "trangThai": "Chờ xét duyệt",
    }
]

def seed_sample_data():
    conn = get_db_connection()
    cur = conn.cursor()
    print("=== BẮT ĐẦU NẠP DỮ LIỆU INPUT MẪU VÀO HỆ THỐNG ===")
    
    count_added = 0
    for r in SAMPLE_RECORDS:
        # Kiểm tra xem mã hồ sơ đã tồn tại chưa
        check = cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE maHoSo = ?", (r["maHoSo"],)).fetchone()[0]
        if check > 0:
            print(f"- Hồ sơ {r['maHoSo']} đã tồn tại trong CSDL, bỏ qua.")
            continue

        pred = ai_engine.predict_record_field(r["tieuDe"], r["tomTat"], r["tuKhoa"], r["thongTinTacGia"])
        cur.execute("""
            INSERT INTO HoSoKHCN (
                maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
                namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
                doi, tapSo, trangBaiBao, chiMucTapChi,
                maDeTai, capDeTai, chuNhiemDeTai, thanhVienThamGia, thoiGianThucHien,
                kinhPhi, sanPhamDuKien, ketQuaNghiemThu,
                trangThai, fileDinhKem, vanBanDaTachTu,
                maLinhVucAI, doTinCayAI, maLinhVucSVM, doTinCaySVM
            ) VALUES (
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?,
                ?, ?, ?,
                ?, ?, ?, ?
            )
        """, (
            r["maHoSo"], r["tieuDe"], r["tomTat"], r["tuKhoa"], r["thongTinTacGia"], r["donViChuTri"],
            r["namCongBo"], r["kyBaoCao"], r["maNguoiNop"], r["maLinhVuc"], r["maLoai"], r.get("maTapChi"),
            r.get("doi"), r.get("tapSo"), r.get("trangBaiBao"), r.get("chiMucTapChi"),
            r.get("maDeTai"), r.get("capDeTai"), r.get("chuNhiemDeTai"), r.get("thanhVienThamGia"), r.get("thoiGianThucHien"),
            r.get("kinhPhi", 0.0), r.get("sanPhamDuKien"), r.get("ketQuaNghiemThu"),
            r["trangThai"], f"MinhChung_{r['maHoSo']}.pdf", pred["segmented_text"],
            pred["phobert_label_id"], pred["phobert_confidence"], pred["svm_label_id"], pred["svm_confidence"]
        ))
        
        # Thêm nhật ký khởi tạo
        cur.execute("""
            INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
            VALUES (?, ?, ?, ?, ?)
        """, (r["maHoSo"], r["maNguoiNop"], "Khởi tạo mới", "Chờ xét duyệt", "Tác giả khởi tạo và nộp hồ sơ dữ liệu mẫu lên hệ thống."))
        count_added += 1
        print(f"✔ Đã thêm thành công hồ sơ: {r['maHoSo']} - {r['tieuDe'][:45]}...")

    conn.commit()
    # Chạy lại mạng đồ thị tri thức NetworkX, Isolation Forest & Deep Autoencoder
    print("\nĐang cập nhật Đồ thị NetworkX, Isolation Forest và Deep Autoencoder...")
    ai_engine.build_networkx_graph_and_detect_anomalies(conn)
    conn.close()
    print(f"\n=== HOÀN TẤT: ĐÃ THÊM MỚI {count_added} HỒ SƠ VÀ CẬP NHẬT TOÀN DIỆN MÔ HÌNH AI ===")

if __name__ == "__main__":
    seed_sample_data()
