# -*- coding: utf-8 -*-
import pandas as pd
import json
import sys

sys.stdout.reconfigure(encoding="utf-8")

df = pd.read_excel("docs/Mau_Nhap_Lieu_HoSo_KHCN.xlsx")

# Define chosen row indices (0-based in df)
chosen_indices = [
    # Top anomalies & duplicate pairs
    98,  # HS-2026-0109: Làm chậm vận tốc nhóm ánh sáng (MSE outlier)
    49,  # HS-2026-0060: Quantum steering in a PT-symmetric system (misdeclared field + duplicate)
    15,  # HS-2026-0026: Squeezing and entanglement (paired with Quantum steering)
    59,  # HS-2026-0070: Exploring the Influence of Gender (duplicate 83.4%)
    52,  # HS-2026-0063: Prepositional usage / conceptual transfer (paired duplicate)
    20,  # HS-2026-0031: Effects of Reflective Thinking On Deep Learning (linguistics vs ML)
    2,   # HS-2026-0013: Giải pháp mô phỏng Simatic IOT 2050 chiết rót nước ngọt (interdisciplinary)
    
    # Missing fields auto-classified by PhoBERT
    62,  # HS-2026-0073: Thạch jelly atiso đỏ (Food Tech)
    63,  # HS-2026-0074: Hoạt tính sinh học lá Xạ đen (Food & Bio-Chem)
    65,  # HS-2026-0076: Hoạt tính kháng nấm nano bạc màng bao (Material / Chem)
    1,   # HS-2026-0012: Vi nhựa và kim loại nặng nước thải đô thị (Environment)
    105, # HS-2026-0116: Thiết kế và thử nghiệm mô đun điện-khí nén PLC (Electrical)
    73,  # HS-2026-0084: Quyết định sử dụng dịch vụ ngân hàng số (Economics)
    74,  # HS-2026-0085: Chat GPT và AI viết nội dung trong học tập (Social Sci / Edu)
    66,  # HS-2026-0077: Thiết kế đầm đa năng biến kiểu (Textile / Mech)
    
    # Representative Publications (Journals & Conferences)
    0,   # HS-2026-0011: Thị giác máy tính YOLOv8 (Scopus Q1/Q2, IT)
    51,  # HS-2026-0061: Numerical study electroosmotic flow (Scopus Q3)
    5,   # HS-2026-0016: Lưu biến bột nhào bánh mì bã đậu nành (Tạp chí Khoa học & CN HUIT)
    3,   # HS-2026-0014: Bảo mật phân quyền hệ thống quản lý văn bản (Conference)
    
    # Projects & Sáng kiến
    64,  # HS-2026-0075: Tinh bột chuối kết hợp chitosan bảo quản dâu tây (Đề tài SV)
    68,  # HS-2026-0079: Nhuộm vải Shibori hoa đậu biếc (Đề tài SV)
    83,  # HS-2026-0094: Hoạt tính chống oxy hóa dịch chiết dược liệu (Đề tài GV)
    106, # HS-2026-0117: Thiết kế thử nghiệm mô đun điều khiển động cơ (Đề tài đặt hàng)
    126, # HS-2026-0137: Điều khiển giám sát PLC S7 1200 qua Webserver (Đề tài đặt hàng)
    71   # HS-2026-0082: Cẩm nang hướng dẫn sử dụng thiết bị PTN (Sáng kiến)
]

print(f"Total curated selected rows: {len(chosen_indices)}")
assert len(chosen_indices) == 25, "Must be exactly 25 curated records!"

for i, idx in enumerate(chosen_indices, start=11):
    row = df.iloc[idx]
    title = str(row["Tiêu đề công trình / Tên đề tài (*)"])
    lh = str(row["Loại hình (*)"])
    dv = str(row["Đơn vị chủ trì (*)"])
    lv = str(row["Lĩnh vực đăng ký (*)"])
    print(f"HS-2026-{i:04d} (Excel row {idx+2}): [{lh}] {title[:60]} | {dv[:25]} | Reg: {lv[:20]}")
