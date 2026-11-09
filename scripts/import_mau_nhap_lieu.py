# -*- coding: utf-8 -*-
"""
Kịch bản nạp toàn bộ tập dữ liệu thực tế: Mau_Nhap_Lieu_HoSo_KHCN.xlsx (149 hồ sơ thực tế từ HUIT)
Thực hiện "Sự thật việc thật":
1. Chuẩn hóa & làm sạch 149 hồ sơ.
2. Đồng bộ danh mục Tác giả (DanhMucTacGia) & Tạp chí (TapChiHopLe).
3. Chạy AI phân loại tự động (PhoBERT & LinearSVC) trên từng hồ sơ:
   - Tự động điền Lĩnh vực nghiên cứu cho 87 đề tài bị khuyết nhãn.
   - Lưu trữ kết quả dự đoán của cả 2 mô hình (nhãn, % tin cậy, nhóm loại hình, thời gian suy luận ms).
4. Nạp vào CSDL khcn_system.db.
5. Chạy phân hệ Phát hiện Bất thường đa tầng:
   - Đồ thị tri thức NetworkX
   - Đo độ tương đồng Cosine TF-IDF
   - Rừng cô lập Isolation Forest
   - Mạng nơ-ron PyTorch Deep Autoencoder
   - Tính Điểm rủi ro tổng hợp (Composite Risk Score) & Hàng đợi ưu tiên.
6. Xuất báo cáo tổng kết chi tiết ra docs/Ket_Qua_Nhap_Lieu_Mau_149_HoSo.json.
"""

import os
import re
import sys
import time
import json
import shutil
import sqlite3
import unicodedata
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding="utf-8")
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from ai_engine import ai_engine
from database import get_db_connection

EXCEL_PATH = (
    os.path.abspath(sys.argv[1])
    if len(sys.argv) > 1
    else os.path.join(BASE_DIR, "docs", "Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
)
OUTPUT_JSON_PATH = os.path.join(BASE_DIR, "docs", "Ket_Qua_Nhap_Lieu_Mau_149_HoSo.json")
DB_PATH = os.path.join(BASE_DIR, "khcn_system.db")
BACKUP_DB_PATH = os.path.join(BASE_DIR, "khcn_system_backup_before_import.db")

FIELD_KEYWORD_MAP = {
    'công nghệ thông tin': 1,
    'truyền thông': 1,
    'toán học': 1,
    'điện': 2,
    'điện tử': 2,
    'tự động': 2,
    'cơ khí': 3,
    'chế tạo máy': 3,
    'động lực': 3,
    'giao thông': 3,
    'vật lý': 3,
    'kinh tế': 4,
    'quản trị': 4,
    'tài chính': 4,
    'kế toán': 4,
    'thương mại': 4,
    'hóa': 5,
    'sinh': 5,
    'dược': 5,
    'môi trường': 5,
    'thực phẩm': 5,
    'thủy sản': 5,
    'tài nguyên': 5,
    'ngôn ngữ': 6,
    'triết': 6,
    'xã hội': 6,
    'nhân văn': 6,
    'chính trị': 6,
    'giáo dục': 6,
    'luật': 6,
}

def clean_text(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).replace("_x000D_\n", "\n").replace("_x000D_", "\n").strip()
    return unicodedata.normalize("NFC", s)

def parse_float(val, default=0.0):
    if pd.isna(val) or val is None:
        return default
    s = str(val).replace(",", "").replace(".", "").replace("VNĐ", "").replace("VND", "").replace("vnd", "").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return default

def parse_int(val, default=2024):
    if pd.isna(val) or val is None:
        return default
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return default

def map_raw_field(raw_field_name: str):
    if not raw_field_name:
        return None
    raw_lower = raw_field_name.lower()
    for kw, fid in FIELD_KEYWORD_MAP.items():
        if kw in raw_lower:
            return fid
    return None

def map_work_type(loai_hinh: str, loai_ct: str, cap_dt: str):
    lh = loai_hinh.lower()
    lct = loai_ct.lower()
    cdt = cap_dt.lower()
    
    if "sáng kiến" in lh or "sáng chế" in lh or "giải pháp hữu ích" in lh:
        return 6  # LCT-SANGCHE
    if "đề tài" in lh:
        if "bộ" in cdt or "tỉnh" in cdt or "nhà nước" in cdt:
            return 4  # LCT-DT-BO
        else:
            return 5  # LCT-DT-TRG
    if "bài báo" in lh:
        if "q1" in lct or "q2" in lct or "isi" in lct or "web of science" in lct:
            return 1  # LCT-ISI-Q1
        elif "scopus" in lct or "q3" in lct or "q4" in lct or "quốc tế" in lct:
            return 2  # LCT-SCOPUS
        else:
            return 3  # LCT-HDGSNN
    return 3

def main():
    print("=" * 70)
    print("BẮT ĐẦU QUY TRÌNH NẠP THỰC TẾ 149 HỒ SƠ TỪ Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
    print("=" * 70)

    if not os.path.exists(EXCEL_PATH):
        print(f"[LỖI] Không tìm thấy file: {EXCEL_PATH}")
        sys.exit(1)

    # 1. Sao lưu CSDL hiện tại
    if os.path.exists(DB_PATH):
        shutil.copy2(DB_PATH, BACKUP_DB_PATH)
        print(f"[1/6] Đã tạo bản sao lưu an toàn CSDL tại: {os.path.basename(BACKUP_DB_PATH)}")

    # 2. Đọc file Excel
    df = pd.read_excel(EXCEL_PATH, sheet_name=0)
    total_rows = len(df)
    print(f"[2/6] Đã đọc thành công tệp Excel: {total_rows} dòng dữ liệu thực tế.")

    # 3. Kết nối CSDL và chuẩn bị dữ liệu danh mục
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Lấy mã hồ sơ hiện tại lớn nhất để đánh số tiếp theo: HS-2026-XXXX
    cur.execute("SELECT maHoSo FROM HoSoKHCN WHERE maHoSo LIKE 'HS-%'")
    existing_codes = [r[0] for r in cur.fetchall()]
    max_seq = 0
    for code in existing_codes:
        m = re.search(r"HS-\d{4}-(\d+)", code)
        if m:
            max_seq = max(max_seq, int(m.group(1)))
    print(f"  -> Mã hồ sơ hiện tại đã có đến số thứ tự: {max_seq}. Các hồ sơ mới sẽ bắt đầu từ HS-2026-{(max_seq+1):04d}.")

    # Đọc danh mục hiện có
    cur.execute("SELECT maLinhVuc, maCode, tenLinhVuc FROM LinhVuc")
    linh_vuc_dict = {r["maLinhVuc"]: r["tenLinhVuc"] for r in cur.fetchall()}

    cur.execute("SELECT maTapChi, tenTapChi FROM TapChiHopLe")
    tap_chi_dict = {r["tenTapChi"].lower().strip(): r["maTapChi"] for r in cur.fetchall()}

    cur.execute("SELECT maTacGia, hoTen FROM DanhMucTacGia")
    tac_gia_dict = {r["hoTen"].lower().strip(): r["maTacGia"] for r in cur.fetchall()}

    # Khởi tạo mô hình AI
    print("[3/6] Đang nạp Động cơ AI (VinAI PhoBERT & TF-IDF + LinearSVC)...")
    ai_engine.train_and_evaluate(db_conn=conn, mode="load_or_init")
    print(f"  -> Đã nạp thành công mô hình. Trọng số PhoBERT & SVM sẵn sàng trên RAM.")

    # 4. Xử lý từng dòng dữ liệu và suy luận AI
    print("[4/6] Đang phân tích văn bản, chạy AI phân loại tự động và chuẩn bị bản ghi...")
    t_start_all = time.perf_counter()

    imported_records = []
    stats = {
        "total_records": total_rows,
        "already_had_field": 0,
        "missing_field_auto_filled": 0,
        "phobert_agreed_with_original": 0,
        "phobert_svm_agreed": 0,
        "by_predicted_field": {},
        "by_work_type": {},
        "total_inference_time_ms": 0.0,
        "records_detail": []
    }

    new_journals_added = 0
    new_authors_added = 0

    for idx, row in df.iterrows():
        seq_num = max_seq + idx + 1
        ma_ho_so = f"HS-2026-{seq_num:04d}"

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

        # Xác định tác giả chính & đại diện
        thong_tin_tg = raw_tac_gia if raw_tac_gia else (raw_chu_nhiem if raw_chu_nhiem else f"Cán bộ {don_vi}")

        # Thêm tác giả vào DanhMucTacGia nếu chưa có
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
            new_authors_added += 1

        # Thêm tạp chí vào TapChiHopLe nếu là bài báo có tên tạp chí mới
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
                new_journals_added += 1

        # Xác định loại công trình
        ma_loai = map_work_type(raw_loai_hinh, raw_loai_ct, cap_de_tai)

        # Chạy AI dự đoán đa nhiệm (PhoBERT và LinearSVC)
        t0_inf = time.perf_counter()
        ai_res = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, thong_tin_tg, don_vi)
        infer_time_ms = round((time.perf_counter() - t0_inf) * 1000.0, 2)
        stats["total_inference_time_ms"] += infer_time_ms

        phobert_field_id = ai_res["phobert_label_id"]
        phobert_conf = ai_res["phobert_confidence"]
        svm_field_id = ai_res["svm_label_id"]
        svm_conf = ai_res["svm_confidence"]
        phobert_work_type = ai_res["work_type_phobert"]
        phobert_work_type_conf = ai_res["work_type_phobert_conf"]
        svm_work_type = ai_res["work_type_svm"]
        svm_work_type_conf = ai_res["work_type_svm_conf"]
        seg_text = ai_res["segmented_text"]

        if phobert_field_id == svm_field_id:
            stats["phobert_svm_agreed"] += 1

        # Xác định Lĩnh vực chính thức:
        # Nếu dòng đã có 'Lĩnh vực đăng ký (*)', ánh xạ theo quy chuẩn
        # Nếu dòng bị khuyết (NaN) -> Áp dụng tự động nhãn do PhoBERT phân loại!
        original_mapped_field = map_raw_field(raw_linh_vuc)
        if original_mapped_field is not None:
            ma_linh_vuc = original_mapped_field
            da_xac_nhan_ai = 1
            stats["already_had_field"] += 1
            if phobert_field_id == original_mapped_field:
                stats["phobert_agreed_with_original"] += 1
        else:
            ma_linh_vuc = phobert_field_id  # Tự động điền bằng PhoBERT!
            da_xac_nhan_ai = 0
            stats["missing_field_auto_filled"] += 1

        field_name = linh_vuc_dict.get(ma_linh_vuc, "Khác")
        stats["by_predicted_field"][field_name] = stats["by_predicted_field"].get(field_name, 0) + 1
        stats["by_work_type"][raw_loai_hinh] = stats["by_work_type"].get(raw_loai_hinh, 0) + 1

        # Xác định trạng thái
        if ket_qua_nt in ["Đã nghiệm thu", "Xuất sắc"]:
            trang_thai = "Đã duyệt"
        elif "bài báo" in raw_loai_hinh.lower():
            trang_thai = "Đã duyệt"
        else:
            trang_thai = "Chờ duyệt"

        # Chuẩn bị bản ghi chèn CSDL
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
        imported_records.append(record_tuple)

        stats["records_detail"].append({
            "stt": idx + 1,
            "maHoSo": ma_ho_so,
            "tieuDe": tieu_de,
            "loaiHinh": raw_loai_hinh,
            "donVi": don_vi,
            "linhVucGoc": raw_linh_vuc or "(Trống - Tự động điền)",
            "linhVucApDung": field_name,
            "phoBertDuDoan": linh_vuc_dict.get(phobert_field_id, ""),
            "phoBertConfidence": phobert_conf,
            "svmDuDoan": linh_vuc_dict.get(svm_field_id, ""),
            "svmConfidence": svm_conf,
            "inferMs": infer_time_ms
        })

        if (idx + 1) % 25 == 0 or (idx + 1) == total_rows:
            print(f"  -> Đã suy luận & chuẩn bị {idx + 1}/{total_rows} hồ sơ ({(idx + 1)*100//total_rows}%)...")

    # 5. Ghi vào CSDL
    print("[5/6] Đang chèn 149 hồ sơ vào bảng HoSoKHCN trong CSDL...")
    insert_sql = """
        INSERT INTO HoSoKHCN (
            maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
            namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
            trangThai, fileDinhKem, isDeleted, doi, tapSo, trangBaiBao, chiMucTapChi,
            maDeTai, capDeTai, chuNhiemDeTai, thanhVienThamGia, thoiGianThucHien,
            kinhPhi, sanPhamDuKien, ketQuaNghiemThu, vanBanDaTachTu,
            maLinhVucAI, doTinCayAI, maLinhVucSVM, doTinCaySVM,
            nhomCongTrinhAI, doTinCayLoaiAI, nhomCongTrinhSVM, doTinCayLoaiSVM,
            moHinhChinh, daXacNhanAI, diemBatThuong, coCanhBao, chiTietCanhBao,
            doTuongDongMax, hoSoTuongDongNhat
        ) VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?
        )
    """
    cur.executemany(insert_sql, imported_records)
    conn.commit()
    print(f"  -> Đã nạp thành công 149 hồ sơ vào HoSoKHCN!")
    print(f"  -> Đã thêm mới {new_journals_added} tạp chí vào TapChiHopLe.")
    print(f"  -> Đã thêm mới {new_authors_added} tác giả vào DanhMucTacGia.")

    # 6. Chạy Phát hiện Bất thường đa tầng trên toàn bộ CSDL (bao gồm cả hồ sơ cũ + mới)
    print("[6/6] Đang chạy chuỗi giải thuật Phát hiện Bất thường đa tầng (NetworkX + Cosine + Isolation Forest + PyTorch Autoencoder)...")
    anomaly_res = ai_engine.build_networkx_graph_and_detect_anomalies(conn)
    conn.commit()

    # Thống kê sau khi quét bất thường
    cur.execute("SELECT count(*) FROM HoSoKHCN WHERE coCanhBao = 1 AND isDeleted = 0")
    total_anomalies = cur.fetchone()[0]

    cur.execute("SELECT count(*) FROM HoSoKHCN WHERE isDeleted = 0")
    total_db_records = cur.fetchone()[0]

    cur.execute("SELECT maHoSo, tieuDe, diemBatThuong, chiTietCanhBao FROM HoSoKHCN WHERE coCanhBao = 1 ORDER BY diemBatThuong DESC LIMIT 5")
    top_anomalies = [dict(r) for r in cur.fetchall()]

    t_total_sec = round(time.perf_counter() - t_start_all, 2)
    avg_infer_ms = round(stats["total_inference_time_ms"] / total_rows, 2)

    stats["summary"] = {
        "total_records_imported": total_rows,
        "total_records_in_db_now": total_db_records,
        "new_journals_added": new_journals_added,
        "new_authors_added": new_authors_added,
        "missing_field_auto_filled_by_phobert": stats["missing_field_auto_filled"],
        "already_had_field_count": stats["already_had_field"],
        "phobert_original_match_rate_pct": round((stats["phobert_agreed_with_original"] / max(stats["already_had_field"], 1)) * 100.0, 2),
        "phobert_svm_agreement_rate_pct": round((stats["phobert_svm_agreed"] / total_rows) * 100.0, 2),
        "total_anomalies_detected_in_db": total_anomalies,
        "total_execution_time_sec": t_total_sec,
        "avg_inference_time_per_record_ms": avg_infer_ms,
        "networkx_nodes": anomaly_res.get("num_nodes", 0),
        "networkx_edges": anomaly_res.get("num_edges", 0),
        "top_prioritized_anomalies": top_anomalies
    }

    # Lưu báo cáo JSON
    with open(OUTPUT_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    conn.close()

    print("\n" + "=" * 70)
    print("HOÀN THÀNH XUẤT SẮC TOÀN BỘ QUY TRÌNH NẠP THỰC TẾ!")
    print("=" * 70)
    print(f"- Tổng số hồ sơ đã nạp thành công: {total_rows} hồ sơ.")
    print(f"- Tổng số hồ sơ trong CSDL hiện tại: {total_db_records} hồ sơ.")
    print(f"- Số đề tài bị khuyết Lĩnh vực được PhoBERT tự động phân loại: {stats['missing_field_auto_filled']} đề tài.")
    print(f"- Độ khớp giữa PhoBERT và nhãn người dùng khai sẵn (62 mẫu): {stats['summary']['phobert_original_match_rate_pct']}%.")
    print(f"- Độ đồng thuận giữa PhoBERT và LinearSVC trên 149 hồ sơ: {stats['summary']['phobert_svm_agreement_rate_pct']}%.")
    print(f"- Thời gian suy luận trung bình: {avg_infer_ms} ms / hồ sơ.")
    print(f"- Tổng số cảnh báo bất thường được phát hiện & xếp hàng ưu tiên: {total_anomalies} hồ sơ.")
    print(f"- Báo cáo chi tiết đã được ghi vào: {os.path.basename(OUTPUT_JSON_PATH)}")
    print("=" * 70)

if __name__ == "__main__":
    main()
