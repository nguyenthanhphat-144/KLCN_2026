# -*- coding: utf-8 -*-
import json
import sys

sys.stdout.reconfigure(encoding="utf-8")

with open("docs/Ket_Qua_Tap_Du_Lieu_Chon_Loc.json", "r", encoding="utf-8") as f:
    data = json.load(f)

print(f"Tên tập dữ liệu: {data['ten_tap_du_lieu']}")
print(f"Tổng số hồ sơ trong CSDL: {data['tong_so_ho_so']}")
print(f"Số hồ sơ gốc (Baseline Seed): {data['so_ho_so_goc']}")
print(f"Số hồ sơ thật chọn lọc: {data['so_ho_so_that_chon_loc']}")
print(f"Số hồ sơ phát hiện có cảnh báo: {data['so_ho_so_co_canh_bao']}")

print("\n--- DANH SÁCH 13 HỒ SƠ CẢNH BÁO TIÊU BIỂU ---")
for idx, r in enumerate(data["danh_sach_canh_bao"], 1):
    reasons = r["chiTiet"].get("reasons", [])
    print(f"{idx}. [{r['maHoSo']}] R_i = {r['diemBatThuong']} | {r['tieuDe'][:60]}")
    print(f"   Lý do: {'; '.join(reasons)}")
    print(f"   Max Sim: {r['doTuongDongMax']}% ({r['hoSoTuongDongNhat']}) | AE_MSE: {r['chiTiet'].get('ae_mse')} | IF: {r['chiTiet'].get('if_score')}\n")

print("\n--- 8 CA TỰ ĐỘNG ĐIỀN KHUYẾT THIẾU LĨNH VỰC BẰNG PHOBERT ---")
missing_cases = [r for r in data["chi_tiet_25_ho_so_that"] if r.get("was_missing")]
for idx, r in enumerate(missing_cases, 1):
    print(f"{idx}. [{r['maHoSo']}] {r['tieuDe'][:60]}")
    print(f"   Đơn vị: {r['donVi']} | Lĩnh vực dự đoán: {r['phoBertDuDoan']} ({r['phoBertConfidence']}%) | SVM: {r['svmDuDoan']} ({r['svmConfidence']}%)\n")
