# -*- coding: utf-8 -*-
import pandas as pd
import json
import sys

sys.stdout.reconfigure(encoding="utf-8")

df = pd.read_excel("docs/Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
print(f"Excel rows: {len(df)}")

# Let's see some key rows by searching titles
keywords = [
    "Quantum steering",
    "vận tốc nhóm ánh sáng",
    "Influence of Gender",
    "Simatic IOT",
    "nano ZnO",
    "Reflective Thinking",
    "thị giác máy tính YOLO",
    "Sáng kiến",
    "Giải pháp hữu ích",
    "sinh viên",
    "đặt hàng"
]

for kw in keywords:
    matches = df[df["Tiêu đề công trình / Tên đề tài (*)"].str.contains(kw, case=False, na=False)]
    print(f"\nKeyword '{kw}': {len(matches)} matches")
    for idx, row in matches.head(2).iterrows():
        print(f"  Row {idx+1}: {row['Loại hình (*)']} | {row['Loại công trình (*)']} | {row['Đơn vị chủ trì (*)']}")
        print(f"    Title: {row['Tiêu đề công trình / Tên đề tài (*)'][:75]}")
        print(f"    Lĩnh vực: {row['Lĩnh vực đăng ký (*)']}")
