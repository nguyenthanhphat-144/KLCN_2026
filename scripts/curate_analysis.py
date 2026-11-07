# -*- coding: utf-8 -*-
import pandas as pd
import json
import sys

sys.stdout.reconfigure(encoding="utf-8")

df = pd.read_excel("docs/Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
with open("docs/Ket_Qua_Nhap_Lieu_Mau_149_HoSo.json", "r", encoding="utf-8") as f:
    json_data = json.load(f)

# Connect to sqlite to get anomaly info for each record HS-2026-0011 to HS-2026-0159
import sqlite3
conn = sqlite3.connect("khcn_system.db")
c = conn.cursor()
c.execute("SELECT maHoSo, tieuDe, diemBatThuong, coCanhBao, chiTietCanhBao, doTuongDongMax, hoSoTuongDongNhat, maLinhVucAI, doTinCayAI, maLinhVucSVM, doTinCaySVM FROM HoSoKHCN")
db_map = {row[0]: row for row in c.fetchall()}
conn.close()

print(f"Total in Excel: {len(df)}")

candidates = []

for idx, row in df.iterrows():
    maHoSo = f"HS-2026-{idx+11:04d}"
    db_rec = db_map.get(maHoSo)
    if not db_rec:
        continue
    
    title = str(row["Tiêu đề công trình / Tên đề tài (*)"])
    loaiHinh = str(row["Loại hình (*)"])
    donVi = str(row["Đơn vị chủ trì (*)"])
    linhVucGoc = str(row["Lĩnh vực đăng ký (*)"])
    loaiCongTrinh = str(row["Loại công trình (*)"])
    tapChi = str(row["Tên tạp chí / hội thảo"])
    
    diemBatThuong = db_rec[2]
    coCanhBao = db_rec[3]
    chiTiet = db_rec[4]
    maxSim = db_rec[5]
    mostSim = db_rec[6]
    aiField = db_rec[7]
    aiConf = db_rec[8]
    svmField = db_rec[9]
    svmConf = db_rec[10]
    
    candidates.append({
        "row_idx": idx,
        "excel_row": idx + 2, # in Excel 1-based header
        "maHoSo": maHoSo,
        "title": title,
        "loaiHinh": loaiHinh,
        "donVi": donVi,
        "linhVucGoc": linhVucGoc,
        "loaiCongTrinh": loaiCongTrinh,
        "tapChi": tapChi,
        "diemBatThuong": diemBatThuong,
        "coCanhBao": coCanhBao,
        "chiTiet": chiTiet,
        "maxSim": maxSim,
        "mostSim": mostSim,
        "aiField": aiField,
        "aiConf": aiConf,
        "svmField": svmField,
        "svmConf": svmConf
    })

print(f"Mapped candidates: {len(candidates)}")

# 1. Anomalies with warning
anom_list = [c for c in candidates if c["coCanhBao"] == 1]
print(f"\n--- Anomaly candidates (total {len(anom_list)}):")
for c in sorted(anom_list, key=lambda x: -x["diemBatThuong"])[:10]:
    print(f"Row {c['excel_row']} ({c['maHoSo']}) R={c['diemBatThuong']}: {c['title'][:70]}")

# 2. Missing fields auto-filled
missing_list = [c for c in candidates if c["linhVucGoc"] == "nan" or not c["linhVucGoc"].strip()]
print(f"\n--- Missing fields candidates (total {len(missing_list)}):")
for c in missing_list[:8]:
    print(f"Row {c['excel_row']} ({c['maHoSo']}) AI_Conf={c['aiConf']}% Field={c['aiField']}: {c['title'][:70]}")

# 3. Work types check
for lh in df["Loại hình (*)"].unique():
    sub = [c for c in candidates if c["loaiHinh"] == lh]
    print(f"Type '{lh}': {len(sub)} candidates")
