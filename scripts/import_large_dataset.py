# -*- coding: utf-8 -*-
"""
Công cụ Nạp Dữ Liệu Lớn Tự Động (Bulk Data Importer) cho Hệ thống KH&CN 2026.
Hỗ trợ đọc file Excel (.xlsx) hoặc CSV, tự động ánh xạ thông minh (Smart Mapping),
tự động xử lý khóa ngoại, làm sạch dữ liệu và kích hoạt AI phân tích hàng loạt.

Cách sử dụng:
    python scripts/import_large_dataset.py <duong_dan_file_excel_hoac_csv>
"""

import os
import re
import sys
import unicodedata
import pandas as pd
from werkzeug.security import generate_password_hash

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from database import get_db_connection
from ai_engine import ai_engine


def clean_str(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


def parse_float(val, default=0.0):
    if pd.isna(val) or val is None:
        return default
    # Loại bỏ ký tự tiền tệ, dấu chấm phân cách ngàn nếu có
    s = str(val).replace(",", "").replace(".", "").replace("VNĐ", "").replace("vnd", "").replace("VND", "").strip()
    try:
        return float(s)
    except ValueError:
        return default


def parse_int(val, default=2026):
    if pd.isna(val) or val is None:
        return default
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return default


def import_dataset(file_path: str):
    if not os.path.exists(file_path):
        print(f"[LỖI] Không tìm thấy file dữ liệu tại đường dẫn: {file_path}")
        return False

    print(f"=== BẮT ĐẦU NẠP DỮ LIỆU LỚN TỪ FILE: {os.path.basename(file_path)} ===")
    
    # 1. Đọc dữ liệu từ Excel hoặc CSV
    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path)
    elif ext == ".csv":
        df = pd.read_csv(file_path, encoding="utf-8-sig")
    else:
        print("[LỖI] Định dạng file không được hỗ trợ. Vui lòng sử dụng file .xlsx hoặc .csv.")
        return False

    print(f"- Đã nạp thành công file: Phát hiện {len(df)} dòng dữ liệu.")
    if len(df) == 0:
        print("[CẢNH BÁO] File trống, không có dữ liệu để nạp.")
        return False

    # 2. Kết nối CSDL và tải sẵn danh mục để đối chiếu
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT maLinhVuc, maCode, tenLinhVuc FROM LinhVuc WHERE trangThai = 1")
    linh_vuc_list = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT maLoai, maCodeLoai, tenLoai, nhomCongTrinh FROM LoaiCongTrinh WHERE trangThai = 1")
    loai_ct_list = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT maTapChi, tenTapChi, chiSoISSN FROM TapChiHopLe WHERE trangThai = 1")
    tap_chi_list = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT maNguoiDung, email, hoTen FROM NguoiDung")
    users_cache = {r["email"].lower(): r["maNguoiDung"] for r in cur.fetchall()}

    # Lấy số thứ tự hồ sơ lớn nhất hiện tại để sinh mã HS-YYYY-XXXX tuần tự
    cur.execute("SELECT maHoSo FROM HoSoKHCN WHERE maHoSo LIKE 'HS-2026-%'")
    existing_codes = [r[0] for r in cur.fetchall()]
    max_idx = 0
    for code in existing_codes:
        m = re.search(r"HS-\d{4}-(\d+)", code)
        if m:
            max_idx = max(max_idx, int(m.group(1)))

    # 3. Chuẩn hóa tên 12 cột đầu vào chuẩn (hỗ trợ cả tiếng Việt có dấu và không dấu)
    col_mapping = {}
    for col in df.columns:
        c_norm = unicodedata.normalize("NFC", str(col)).lower()
        if "loại hình" in c_norm or "loai_hinh" in c_norm:
            col_mapping["loai_hinh"] = col
        elif "tiêu đề" in c_norm or "tieu_de" in c_norm or "tên đề tài" in c_norm:
            col_mapping["tieu_de"] = col
        elif "tóm tắt" in c_norm or "tom_tat" in c_norm or "abstract" in c_norm:
            col_mapping["tom_tat"] = col
        elif "từ khóa" in c_norm or "tu_khoa" in c_norm or "keyword" in c_norm:
            col_mapping["tu_khoa"] = col
        elif "đơn vị" in c_norm or "don_vi" in c_norm:
            col_mapping["don_vi"] = col
        elif "năm" in c_norm or "nam" in c_norm:
            col_mapping["nam"] = col
        elif "kỳ" in c_norm or "ky" in c_norm:
            col_mapping["ky"] = col
        elif "lĩnh vực" in c_norm or "linh_vuc" in c_norm:
            col_mapping["linh_vuc"] = col
        elif "loại công trình" in c_norm or "loai_ct" in c_norm:
            col_mapping["loai_ct"] = col
        elif "tên tạp chí" in c_norm or "hội thảo" in c_norm or ("tạp chí" in c_norm and "chỉ mục" not in c_norm) or "tap_chi" in c_norm:
            col_mapping["tap_chi"] = col
        elif "chỉ mục" in c_norm or "chi_muc" in c_norm:
            col_mapping["chi_muc"] = col
        elif "sản phẩm" in c_norm or "san_pham" in c_norm:
            col_mapping["san_pham"] = col

    # 4. Lặp qua từng dòng dữ liệu và xử lý nạp theo đúng 12 trường chuẩn
    success_count = 0
    error_count = 0

    for idx_row, row in df.iterrows():
        loai_hinh = clean_str(row.get(col_mapping.get("loai_hinh", ""), "Bài báo khoa học"))
        tieu_de = clean_str(row.get(col_mapping.get("tieu_de", ""), ""))
        tom_tat = clean_str(row.get(col_mapping.get("tom_tat", ""), ""))
        tu_khoa = clean_str(row.get(col_mapping.get("tu_khoa", ""), ""))
        don_vi = clean_str(row.get(col_mapping.get("don_vi", ""), "Khoa Công nghệ Thông tin")) or "Khoa Công nghệ Thông tin"
        nam_cong_bo = parse_int(row.get(col_mapping.get("nam", ""), 2026), 2026)
        ky_bao_cao = clean_str(row.get(col_mapping.get("ky", ""), "Đợt 1 - Năm học 2025-2026")) or "Đợt 1 - Năm học 2025-2026"

        if not tieu_de or not tom_tat:
            print(f"[Bỏ qua Dòng {idx_row + 2}] Thiếu Tiêu đề hoặc Tóm tắt bắt buộc.")
            error_count += 1
            continue

        ma_nguoi_nop = 4  # Gán người nộp mặc định
        tac_gia_auto = don_vi  # Gán tự động để đảm bảo ràng buộc NOT NULL trong CSDL

        # a) Ánh xạ thông minh Lĩnh vực đăng ký (LinhVuc)
        raw_lv = clean_str(row.get(col_mapping.get("linh_vuc", ""), "")).lower()
        ma_linh_vuc = 1
        for lv in linh_vuc_list:
            if lv["maCode"].lower() in raw_lv or lv["tenLinhVuc"].lower() in raw_lv:
                ma_linh_vuc = lv["maLinhVuc"]
                break

        # b) Ánh xạ thông minh Loại công trình (LoaiCongTrinh)
        raw_lct = clean_str(row.get(col_mapping.get("loai_ct", ""), "")).lower()
        raw_nhom = loai_hinh.lower()
        ma_loai = 1 if "đề tài" not in raw_nhom and "đề tài" not in raw_lct else 4
        for lct in loai_ct_list:
            if lct["maCodeLoai"].lower() in raw_lct or lct["tenLoai"].lower() in raw_lct:
                ma_loai = lct["maLoai"]
                break

        # c) Ánh xạ Tên tạp chí / Hội thảo (TapChiHopLe)
        raw_tc = clean_str(row.get(col_mapping.get("tap_chi", ""), "")).lower()
        ma_tap_chi = None
        if raw_tc:
            for tc in tap_chi_list:
                if tc["tenTapChi"].lower() in raw_tc or tc["chiSoISSN"].lower() in raw_tc:
                    ma_tap_chi = tc["maTapChi"]
                    break

        # d) Chỉ mục tạp chí & Sản phẩm dự kiến
        chi_muc = clean_str(row.get(col_mapping.get("chi_muc", ""), "")) or None
        san_pham = clean_str(row.get(col_mapping.get("san_pham", ""), "")) or None

        # e) Sinh mã hồ sơ chuẩn tuần tự
        max_idx += 1
        ma_ho_so = f"HS-2026-{max_idx:04d}"

        # f) Chạy AI PhoBERT & SVM dự đoán nhãn lĩnh vực
        pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, don_vi)

        # g) Lưu vào CSDL
        cur.execute("""
            INSERT INTO HoSoKHCN (
                maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
                namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
                chiMucTapChi, sanPhamDuKien,
                trangThai, fileDinhKem, vanBanDaTachTu,
                maLinhVucAI, doTinCayAI, maLinhVucSVM, doTinCaySVM
            ) VALUES (
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?,
                ?, ?, ?,
                ?, ?, ?, ?
            )
        """, (
            ma_ho_so, tieu_de, tom_tat, tu_khoa, tac_gia_auto, don_vi,
            nam_cong_bo, ky_bao_cao, ma_nguoi_nop, ma_linh_vuc, ma_loai, ma_tap_chi,
            chi_muc, san_pham,
            "Chờ xét duyệt", f"MinhChung_{ma_ho_so}.pdf", pred["segmented_text"],
            pred["phobert_label_id"], pred["phobert_confidence"],
            pred["svm_label_id"], pred["svm_confidence"]
        ))

        # Ghi nhật ký vào LichSuXetDuyet
        cur.execute("""
            INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
            VALUES (?, ?, ?, ?, ?)
        """, (ma_ho_so, ma_nguoi_nop, "Khởi tạo mới", "Chờ xét duyệt", "Hồ sơ được nạp tự động qua công cụ Bulk Import vào hệ thống."))

        success_count += 1
        if success_count % 10 == 0 or success_count == len(df):
            print(f"  -> Đã nạp thành công {success_count}/{len(df)} hồ sơ...")

    conn.commit()

    # 5. Kích hoạt cập nhật Đồ thị NetworkX, Isolation Forest & Deep Autoencoder MỘT LẦN DUY NHẤT sau khi nạp xong
    print("\n--- ĐANG KÍCH HOẠT MÔ HÌNH AI PHÂN TÍCH TOÀN BỘ DỮ LIỆU ĐÃ NẠP ---")
    ai_engine.build_networkx_graph_and_detect_anomalies(conn)
    conn.close()

    print(f"\n=== HOÀN TẤT NHẬP DỮ LIỆU LỚN ===")
    print(f"✔ Thành công: {success_count} hồ sơ.")
    print(f"✖ Bỏ qua / Lỗi: {error_count} hồ sơ.")
    print(f"Đồ thị tri thức NetworkX, Isolation Forest & Deep Autoencoder đã được cập nhật toàn diện.")
    return True


if __name__ == "__main__":
    if len(sys.argv) > 1:
        target_file = sys.argv[1]
    else:
        target_file = os.path.join(os.path.dirname(__file__), "..", "docs", "Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
    import_dataset(target_file)
