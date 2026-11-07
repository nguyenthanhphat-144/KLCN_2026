# -*- coding: utf-8 -*-
import sqlite3
import sys

sys.stdout.reconfigure(encoding="utf-8")

conn = sqlite3.connect("khcn_system.db")
c = conn.cursor()
c.execute("SELECT maHoSo, tieuDe, doTuongDongMax, hoSoTuongDongNhat, chiTietCanhBao FROM HoSoKHCN WHERE maHoSo IN ('HS-2026-0060', 'HS-2026-0026', 'HS-2026-0070', 'HS-2026-0063', 'HS-2026-0109')")
for row in c.fetchall():
    print(f"{row[0]}: {row[1]}")
    print(f"   MaxSim: {row[2]}% with {row[3]}")
    print(f"   Alerts: {row[4]}\n")
conn.close()
