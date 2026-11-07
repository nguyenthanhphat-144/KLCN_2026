# -*- coding: utf-8 -*-
"""
Kịch bản trích lọc 25 hồ sơ nghiên cứu thật tiêu biểu từ Mau_Nhap_Lieu_HoSo_KHCN.xlsx
Kết hợp cùng 10 hồ sơ gốc tạo thành bộ dữ liệu chuẩn 35 hồ sơ.
"""
import os
import sys
import shutil
import sqlite3
import json
import time
import re
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = r"d:\Aziel\KLTN_2026"
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
DB_PATH = os.path.join(BASE_DIR, "khcn_system.db")
BACKUP_PATH = os.path.join(BASE_DIR, "khcn_system_backup_before_import.db")
EXCEL_PATH = os.path.join(BASE_DIR, "docs", "Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
OUT_JSON_PATH = os.path.join(BASE_DIR, "docs", "Ket_Qua_Tap_Du_Lieu_Chon_Loc.json")

print("=== BẮT ĐẦU QUY TRÌNH CHỌN LỌC DỮ LIỆU THẬT & CẬP NHẬT HỆ THỐNG ===")

# Bước 1: Khôi phục database từ bản backup sạch ban đầu (10 hồ sơ gốc)
if not os.path.exists(BACKUP_PATH):
    raise FileNotFoundError(f"Không tìm thấy file backup {BACKUP_PATH}")

shutil.copyfile(BACKUP_PATH, DB_PATH)
print("-> [1/5] Đã khôi phục khcn_system.db về trạng thái chuẩn ban đầu (10 hồ sơ gốc).")

# Bước 2: Đọc file Excel
df_excel = pd.read_excel(EXCEL_PATH)
print(f"-> [2/5] Đã nạp file Excel: {len(df_excel)} dòng dữ liệu gốc.")

# 25 dòng được chọn lọc có tính toán kỹ lưỡng (0-based index)
CURATED_INDICES = [
    # Nhóm 1: Ca bất thường đặc biệt & Cặp trùng lặp thực tế
    98,  # Dòng 100: Làm chậm vận tốc nhóm ánh sáng (MSE Outlier = 0.169, Ri = 0.74, Vật lý/Quang học)
    49,  # Dòng 51: Quantum steering in a PT-symmetric system (Lệch lĩnh vực Cơ khí vs Vật lý/IT, trùng 77.4%)
    15,  # Dòng 16: On modules and rings having large absolute direct summands (Cặp đối chứng trùng lặp)
    59,  # Dòng 61: Exploring the Influence of Gender and L1 Conceptual Transfer (Trùng lặp 83.4%)
    52,  # Dòng 53: Jamming-and-Relaying Strategy for Non-linear Energy Scavenging Networks (Cặp đối chứng)
    20,  # Dòng 21: Effects of Reflective Thinking On Deep Learning (Giao thoa ngôn ngữ học & Deep Learning)
    2,   # Dòng 3: Giải pháp mô phỏng Simatic IOT 2050 chiết rót nước ngọt (Liên ngành Tự động hóa & Thực phẩm)
    48,  # Dòng 49: All optical switching in a three-level V-type atomic medium (Laser/Quang học - Ts. Lê Văn Đoài)

    # Nhóm 2: Minh chứng Tự động phân loại điền khuyết thiếu lĩnh vực (Khuyết Lĩnh vực trong Excel)
    62,  # Dòng 63: Nghiên cứu thử nghiệm sản phẩm thạch jelly atiso đỏ (Khoa CN Thực phẩm)
    63,  # Dòng 64: Thu nhận các chất có hoạt tính sinh học từ lá Xạ đen (Khoa CN Thực phẩm & Sinh học)
    65,  # Dòng 66: Nghiên cứu hoạt tính kháng nấm của nano bạc kết hợp màng bao sinh học (Vật liệu & Môi trường)
    1,   # Dòng 2: Nghiên cứu công nghệ xử lý vi nhựa và kim loại nặng trong nước thải đô thị (Môi trường)
    66,  # Dòng 67: Thiết kế đầm đa năng với nhiều kiểu dáng thay đổi tùy thuộc mục đích (Khoa May & Thời trang)
    68,  # Dòng 69: Nghiên cứu kỹ thuật nhuộm vải Shibori từ màu tự nhiên của hoa đậu biếc (Dệt nhuộm & Thời trang)
    82,  # Dòng 83: Quyền nuôi con và nghĩa vụ cấp dưỡng của cha, mẹ sau ly hôn (Khoa Luật - KHXH)
    74,  # Dòng 75: Đánh giá tác động của ChatGPT và các hệ thống AI đối với sinh viên (Giáo dục & Sư phạm)

    # Nhóm 3: Bài báo khoa học quốc tế uy tín & Tạp chí HUIT
    0,   # Dòng 1: Ứng dụng thị giác máy tính YOLOv8 trong nhận dạng biển số phương tiện (Scopus Q1/Q2, CNTT)
    50,  # Dòng 51: Numerical study of rectified electroosmotic flow in nanofluidics (Scopus Q3)
    5,   # Dòng 6: The effects of extraction conditions on antioxidant activity (Tạp chí Khoa học & CN HUIT)
    4,   # Dòng 5: Các chiến lược ứng dụng quyền chọn trên thị trường chứng khoán (Khoa Tài chính - Kế toán)

    # Nhóm 4: Đề tài Giảng viên, Sinh viên, Đặt hàng & Sáng kiến
    64,  # Dòng 65: Ứng dụng tinh bột chuối kết hợp chitosan lên quá trình bảo quản dâu tây (Đề tài SV)
    83,  # Dòng 84: Khảo sát hoạt tính chống oxy hóa và ức chế enzyme alpha-glucosidase (Đề tài GV)
    105, # Dòng 106: Thiết kế và thử nghiệm mô đun điều khiển hệ thống điện-khí nén bằng PLC (Đề tài đặt hàng)
    126, # Dòng 127: Xây dựng mô hình Điều khiển và giám sát qua webserver với PLC S7 1200 (Đề tài đặt hàng)
    71   # Dòng 72: Cẩm nang hướng dẫn sử dụng thiết bị tại phòng thí nghiệm (Sáng kiến cấp Trường)
]

print(f"-> Đã chọn {len(CURATED_INDICES)} hồ sơ đại diện tiêu biểu.")

# Import helper functions from import_mau_nhap_lieu
from scripts.import_mau_nhap_lieu import (
    clean_text, parse_int, parse_float, map_raw_field, map_work_type
)
from ai_engine import KHCNAIEngine

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

# Khởi tạo mô hình AI
print("-> [3/5] Khởi tạo KHCNAIEngine...")
ai_engine = KHCNAIEngine()
ai_engine.train_and_evaluate(db_conn=conn, mode="load_or_init")

cur.execute("SELECT maLinhVuc, maCode, tenLinhVuc FROM LinhVuc")
linh_vuc_dict = {r["maLinhVuc"]: r["tenLinhVuc"] for r in cur.fetchall()}

cur.execute("SELECT maTapChi, tenTapChi FROM TapChiHopLe")
tap_chi_dict = {r["tenTapChi"].lower().strip(): r["maTapChi"] for r in cur.fetchall()}

cur.execute("SELECT maTacGia, hoTen FROM DanhMucTacGia")
tac_gia_dict = {r["hoTen"].lower().strip(): r["maTacGia"] for r in cur.fetchall()}

curated_records_data = []
curated_details = []

for i, row_idx in enumerate(CURATED_INDICES, start=11):
    row = df_excel.iloc[row_idx]
    ma_ho_so = f"HS-2026-{i:04d}"

    raw_loai_hinh = clean_text(row.get("Loại hình (*)", "Bài báo khoa học"))
    tieu_de = clean_text(row.get("Tiêu đề công trình / Tên đề tài (*)", ""))
    tom_tat = clean_text(row.get("Tóm tắt nội dung (*)", ""))
    tu_khoa = clean_text(row.get("Từ khóa (*)", ""))
    raw_tac_gia = clean_text(row.get("Thông tin nhóm tác giả (*)", ""))
    raw_chu_nhiem = clean_text(row.get("Chủ nhiệm đề tài", ""))
    don_vi = clean_text(row.get("Đơn vị chủ trì (*)", "Khoa Công nghệ Thông tin")) or "Khoa Công nghệ Thông tin"
    nam_cong_bo = parse_int(row.get("Năm công bố (*)", 2024), 2024)
    ky_bao_cao = clean_text(row.get("Kỳ báo cáo", "Học kỳ 1 - 2024")) or "Học kỳ 1 - 2024"
    raw_linh_vuc = clean_text(row.get("Lĩnh vực đăng ký (*)", ""))
    raw_loai_ct = clean_text(row.get("Loại công trình (*)", ""))
    ten_tap_chi = clean_text(row.get("Tên tạp chí / hội thảo", ""))
    doi = clean_text(row.get("Mã DOI", ""))
    tap_so = clean_text(row.get("Tập / Số", ""))
    trang_bai_bao = clean_text(row.get("Trang bài báo", ""))
    chi_muc = clean_text(row.get("Chỉ mục tạp chí", ""))
    ma_de_tai = clean_text(row.get("Mã số đề tài", ""))
    cap_de_tai = clean_text(row.get("Cấp đề tài", ""))
    thanh_vien = clean_text(row.get("Thành viên tham gia", ""))
    thoi_gian_th = clean_text(row.get("Thời gian thực hiện", ""))
    kinh_phi = parse_float(row.get("Kinh phí (VNĐ)", 0.0))
    san_pham_dk = clean_text(row.get("Sản phẩm dự kiến", ""))
    ket_qua_nt = clean_text(row.get("Kết quả nghiệm thu", ""))

    thong_tin_tg = raw_tac_gia if raw_tac_gia else (raw_chu_nhiem if raw_chu_nhiem else f"Cán bộ {don_vi}")

    # Tác giả
    main_author_name = thong_tin_tg.split(",")[0].split(";")[0].strip()
    if main_author_name and main_author_name.lower() not in tac_gia_dict:
        cur.execute("SELECT count(*) FROM DanhMucTacGia")
        tg_count = cur.fetchone()[0] + 1
        code_tg = f"TG-{tg_count:03d}"
        cur.execute("""
            INSERT INTO DanhMucTacGia (maCodeTG, hoTen, hocHamHocVi, email, donViCongTac, maLinhVucChinh)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (code_tg, main_author_name, "Giảng viên", "", don_vi, 1))
        new_id = cur.lastrowid
        tac_gia_dict[main_author_name.lower()] = new_id

    # Tạp chí
    ma_tap_chi = None
    if ten_tap_chi:
        tc_key = ten_tap_chi.lower()
        if tc_key in tap_chi_dict:
            ma_tap_chi = tap_chi_dict[tc_key]
        else:
            cur.execute("""
                INSERT INTO TapChiHopLe (tenTapChi, chiSoISSN, xepHang, maLinhVuc, uyTin, trangThai)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (ten_tap_chi, "", "Trong nước/Hội thảo", 1, 1, 1))
            ma_tap_chi = cur.lastrowid
            tap_chi_dict[tc_key] = ma_tap_chi

    ma_loai = map_work_type(raw_loai_hinh, raw_loai_ct, cap_de_tai)

    # Chạy AI
    ai_res = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, thong_tin_tg, don_vi)
    phobert_field_id = ai_res["phobert_label_id"]
    phobert_conf = ai_res["phobert_confidence"]
    svm_field_id = ai_res["svm_label_id"]
    svm_conf = ai_res["svm_confidence"]
    phobert_work_type = ai_res["work_type_phobert"]
    phobert_work_type_conf = ai_res["work_type_phobert_conf"]
    svm_work_type = ai_res["work_type_svm"]
    svm_work_type_conf = ai_res["work_type_svm_conf"]
    seg_text = ai_res["segmented_text"]

    original_mapped_field = map_raw_field(raw_linh_vuc)
    if original_mapped_field is not None:
        ma_linh_vuc = original_mapped_field
        da_xac_nhan_ai = 1
        was_missing = False
    else:
        ma_linh_vuc = phobert_field_id
        da_xac_nhan_ai = 0
        was_missing = True

    if ket_qua_nt in ["Đã nghiệm thu", "Xuất sắc"] or "bài báo" in raw_loai_hinh.lower():
        trang_thai = "Đã duyệt"
    else:
        trang_thai = "Chờ duyệt"

    record_tuple = (
        ma_ho_so, tieu_de, tom_tat, tu_khoa, thong_tin_tg, don_vi,
        nam_cong_bo, ky_bao_cao, 4, ma_linh_vuc, ma_loai, ma_tap_chi,
        trang_thai, "", 0, doi, tap_so, trang_bai_bao, chi_muc,
        ma_de_tai, cap_de_tai, raw_chu_nhiem, thanh_vien, thoi_gian_th,
        kinh_phi, san_pham_dk, ket_qua_nt, seg_text,
        phobert_field_id, phobert_conf, svm_field_id, svm_conf,
        phobert_work_type, phobert_work_type_conf,
        svm_work_type, svm_work_type_conf,
        "phobert", da_xac_nhan_ai, 0.0, 0, "", 0.0, ""
    )
    curated_records_data.append(record_tuple)
    
    curated_details.append({
        "stt": i - 10,
        "maHoSo": ma_ho_so,
        "excel_row": row_idx + 2,
        "tieuDe": tieu_de,
        "loaiHinh": raw_loai_hinh,
        "donVi": don_vi,
        "linhVucGoc": raw_linh_vuc or "(Trống - Tự động điền)",
        "linhVucApDung": linh_vuc_dict.get(ma_linh_vuc, "Khác"),
        "was_missing": was_missing,
        "phoBertDuDoan": linh_vuc_dict.get(phobert_field_id, ""),
        "phoBertConfidence": phobert_conf,
        "svmDuDoan": linh_vuc_dict.get(svm_field_id, ""),
        "svmConfidence": svm_conf
    })

# Chèn 25 bản ghi vào DB
cur.executemany("""
    INSERT INTO HoSoKHCN (
        maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
        namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
        trangThai, fileDinhKem, isDeleted, doi, tapSo, trangBaiBao, chiMucTapChi,
        maDeTai, capDeTai, chuNhiemDeTai, thanhVienThamGia, thoiGianThucHien,
        kinhPhi, sanPhamDuKien, ketQuaNghiemThu, vanBanDaTachTu,
        maLinhVucAI, doTinCayAI, maLinhVucSVM, doTinCaySVM,
        nhomCongTrinhAI, doTinCayLoaiAI, nhomCongTrinhSVM, doTinCayLoaiSVM,
        moHinhChinh, daXacNhanAI, diemBatThuong, coCanhBao, chiTietCanhBao,
        doTuongDongMax, hoSoTuongDongNhat, ngayNop
    ) VALUES (
        ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?,
        ?, ?, ?, ?,
        ?, ?, ?, ?,
        ?, ?, ?, ?,
        ?, ?, ?, ?, ?,
        ?, ?, datetime('now')
    )
""", curated_records_data)

conn.commit()
print(f"-> [4/5] Đã chèn thành công {len(curated_records_data)} hồ sơ chọn lọc vào CSDL.")

# Bước 5: Chạy Đồ thị tri thức NetworkX, Mô hình phát hiện bất thường (Isolation Forest + Autoencoder + Rule-based)
print("-> [5/5] Đang xây dựng Đồ thị tri thức NetworkX và tính toán phát hiện bất thường trên toàn bộ 35 hồ sơ...")
anomaly_results = ai_engine.build_networkx_graph_and_detect_anomalies(conn)

# Lấy danh sách hồ sơ có cảnh báo trong bộ 35 hồ sơ
cur.execute("""
    SELECT maHoSo, tieuDe, diemBatThuong, coCanhBao, chiTietCanhBao, doTuongDongMax, hoSoTuongDongNhat
    FROM HoSoKHCN
    WHERE coCanhBao = 1
    ORDER BY diemBatThuong DESC
""")
flagged_rows = cur.fetchall()
print(f"\n=== DANH SÁCH {len(flagged_rows)} HỒ SƠ CÓ CẢNH BÁO TRONG BỘ 35 HỒ SƠ ===")
for r in flagged_rows:
    print(f"[{r['maHoSo']}] Điểm R_i = {r['diemBatThuong']} | MaxSim = {r['doTuongDongMax']}% ({r['hoSoTuongDongNhat']}) | {r['tieuDe'][:65]}")
    try:
        details = json.loads(r['chiTietCanhBao'])
        print(f"    -> Lý do: {', '.join(details.get('reasons', []))}")
        print(f"    -> Điểm thành phần: IF={details.get('if_score')}, AE_MSE={details.get('ae_mse')}, AE_Score={details.get('ae_score')}, RuleSev={details.get('rule_severity')}")
    except Exception:
        pass

# Thống kê tổng hợp
cur.execute("SELECT count(*) FROM HoSoKHCN")
total_now = cur.fetchone()[0]

summary_data = {
    "ten_tap_du_lieu": "Bộ dữ liệu 35 hồ sơ KHCN chọn lọc HUIT (10 hồ sơ hạt nhân + 25 hồ sơ thực tế tiêu biểu)",
    "tong_so_ho_so": total_now,
    "so_ho_so_goc": 10,
    "so_ho_so_that_chon_loc": len(curated_records_data),
    "so_ho_so_co_canh_bao": len(flagged_rows),
    "danh_sach_canh_bao": [
        {
            "maHoSo": r["maHoSo"],
            "tieuDe": r["tieuDe"],
            "diemBatThuong": r["diemBatThuong"],
            "doTuongDongMax": r["doTuongDongMax"],
            "hoSoTuongDongNhat": r["hoSoTuongDongNhat"],
            "chiTiet": json.loads(r["chiTietCanhBao"]) if r["chiTietCanhBao"] else {}
        }
        for r in flagged_rows
    ],
    "chi_tiet_25_ho_so_that": curated_details
}

with open(OUT_JSON_PATH, "w", encoding="utf-8") as f:
    json.dump(summary_data, f, ensure_ascii=False, indent=2)

print(f"\n-> Đã lưu báo cáo phân tích bộ dữ liệu chọn lọc tại: {OUT_JSON_PATH}")

conn.close()
print("=== HOÀN TẤT QUY TRÌNH CHỌN LỌC DỮ LIỆU THẬT THÀNH CÔNG! ===")
