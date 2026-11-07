# -*- coding: utf-8 -*-
import json
import sys
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

with open("docs/Ket_Qua_Nhap_Lieu_Mau_149_HoSo.json", "r", encoding="utf-8") as f:
    data = json.load(f)

records = data.get("records_detail", [])
print(f"Total records in benchmark json: {len(records)}")

# Group by anomaly score
anomalies = [r for r in records if r.get("diemBatThuong", 0) >= 0.35]
print(f"Total anomalies with score >= 0.35: {len(anomalies)}")

print("\n--- TOP ANOMALIES ---")
for r in sorted(records, key=lambda x: -x.get("diemBatThuong", 0))[:12]:
    print(f"[{r['maHoSo']}] R={r.get('diemBatThuong')} | Sim={r.get('doTuongDongMax')}% | AI: {r.get('predicted_field_phobert')[:30]} | Reg: {str(r.get('original_field'))[:30]}")
    print(f"   Title: {r.get('tieuDe')[:80]}")
    if r.get("chiTietCanhBao"):
        print(f"   Alerts: {r.get('chiTietCanhBao')}")

# Group by missing field auto-filled
missing_fields = [r for r in records if r.get("was_missing_field")]
print(f"\nTotal records with missing registered field: {len(missing_fields)}")

# Group by work type
df_excel = pd.read_excel("docs/Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
print("\nWork types in Excel:")
print(df_excel["Loại hình (*)"].value_counts())
print("\nTypes of projects/articles:")
print(df_excel["Loại công trình (*)"].value_counts())
