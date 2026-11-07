# -*- coding: utf-8 -*-
import sqlite3
import sys

sys.stdout.reconfigure(encoding="utf-8")

conn = sqlite3.connect("khcn_system.db")
c = conn.cursor()

c.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0")
print("Total records:", c.fetchone()[0])

c.execute("SELECT donViChuTri, COUNT(*), SUM(CASE WHEN trangThai = 'Đã duyệt' THEN 1 ELSE 0 END) FROM HoSoKHCN WHERE isDeleted = 0 GROUP BY donViChuTri ORDER BY COUNT(*) DESC")
print("\nStats don vi (total faculties):", len(c.fetchall()))

c.execute("""
    SELECT namCongBo, COUNT(*)
    FROM HoSoKHCN h
    JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
    WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Bài báo%' OR lct.tenLoai LIKE '%Bài báo%')
    GROUP BY namCongBo
    ORDER BY namCongBo DESC
""")
print("\nStats bai bao nam:", c.fetchall())

c.execute("""
    SELECT COALESCE(tc.xepHang, h.chiMucTapChi, 'Chưa phân hạng') AS chiMuc, COUNT(*)
    FROM HoSoKHCN h
    JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
    LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
    WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Bài báo%' OR lct.tenLoai LIKE '%Bài báo%')
    GROUP BY 1 ORDER BY 2 DESC
""")
print("\nStats bai bao chi muc:", c.fetchall())

c.execute("""
    SELECT lct.tenLoai, COUNT(*)
    FROM HoSoKHCN h
    JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
    WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')
    GROUP BY lct.maLoai, lct.tenLoai ORDER BY 2 DESC
""")
print("\nStats de tai loai:", c.fetchall())

c.execute("""
    SELECT h.donViChuTri, COUNT(*)
    FROM HoSoKHCN h
    JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
    WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')
    GROUP BY h.donViChuTri ORDER BY 2 DESC
""")
print("\nStats de tai don vi:", c.fetchall())

conn.close()
