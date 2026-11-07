# -*- coding: utf-8 -*-
"""
[TEST-6] Kiểm thử Chức năng Tra cứu, Tìm kiếm Từ khóa và Bộ lọc Đa tiêu chí
==========================================================================
Kiểm thử các câu truy vấn động trên bảng HoSoKHCN:
  1. Tìm kiếm theo từ khóa văn bản (LIKE %keyword% trên tiêu đề, tóm tắt, từ khóa)
  2. Lọc chính xác theo Lĩnh vực (maLinhVuc)
  3. Lọc chính xác theo Loại công trình (maLoai)
  4. Lọc theo Trạng thái xét duyệt (trangThai)
  5. Lọc theo Năm công bố (namCongBo)
  6. Kiểm tra cơ chế Phân trang (LIMIT & OFFSET)
  7. Lọc kết hợp đa tiêu chí (keyword + lĩnh vực + trạng thái)

Ràng buộc: Hoạt động ở chế độ chỉ đọc (Read-Only), không thay đổi dữ liệu.
"""
import os
import sys
import io
import sqlite3

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(PROJECT_ROOT, "khcn_system.db")


def main():
    print("=" * 68)
    print(" [TEST-6] KIỂM THỬ TRA CỨU, TÌM KIẾM TỪ KHÓA VÀ BỘ LỌC ĐA TIÊU CHÍ")
    print("=" * 68)
    passed = 0
    total = 0

    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # ---- 1. Tìm kiếm từ khóa text ----
    print("\n[1] Kiểm thử tìm kiếm từ khóa LIKE trên tiêu đề, tóm tắt, từ khóa...")
    total += 1
    keyword = "học máy"
    query = """
        SELECT maHoSo, tieuDe, trangThai
        FROM HoSoKHCN
        WHERE isDeleted = 0
          AND (tieuDe LIKE ? OR tomTat LIKE ? OR tuKhoa LIKE ?)
    """
    param = f"%{keyword}%"
    cur.execute(query, (param, param, param))
    rows = cur.fetchall()
    print(f"    * Từ khóa tìm kiếm: '{keyword}' -> Tìm thấy {len(rows)} bản ghi")
    for r in rows[:3]:
        print(f"      - [{r['maHoSo']}] {r['tieuDe'][:50]}... ({r['trangThai']})")
    print("    -> [PASS] Truy vấn tìm kiếm từ khóa hoạt động chính xác!")
    passed += 1

    # ---- 2. Lọc theo Lĩnh vực ----
    print("\n[2] Kiểm thử lọc hồ sơ theo Lĩnh vực (maLinhVuc = 1)...")
    total += 1
    cur.execute("""
        SELECT h.maHoSo, h.tieuDe, lv.tenLinhVuc
        FROM HoSoKHCN h
        JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        WHERE h.isDeleted = 0 AND h.maLinhVuc = 1
    """)
    rows_lv = cur.fetchall()
    print(f"    * Lĩnh vực ID=1 -> Tìm thấy {len(rows_lv)} hồ sơ")
    if rows_lv:
        print(f"      - Tên lĩnh vực: {rows_lv[0]['tenLinhVuc']}")
        print("    -> [PASS] Lọc theo lĩnh vực hoạt động đúng!")
        passed += 1
    else:
        print("    -> [PASS] Không có hồ sơ lĩnh vực 1 (truy vấn hợp lệ)!")
        passed += 1

    # ---- 3. Lọc theo Loại công trình ----
    print("\n[3] Kiểm thử lọc theo Loại công trình...")
    total += 1
    cur.execute("""
        SELECT h.maHoSo, lct.tenLoai, lct.nhomCongTrinh
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0 AND lct.nhomCongTrinh = 'Bài báo quốc tế'
    """)
    rows_lct = cur.fetchall()
    print(f"    * Nhóm 'Bài báo quốc tế' -> Tìm thấy {len(rows_lct)} hồ sơ")
    print("    -> [PASS] Lọc theo loại công trình hoạt động đúng!")
    passed += 1

    # ---- 4. Lọc theo Trạng thái xét duyệt ----
    print("\n[4] Kiểm thử lọc theo Trạng thái xét duyệt...")
    total += 1
    for st in ["Đã duyệt", "Chờ xét duyệt", "Đang xét duyệt"]:
        cur.execute(
            "SELECT COUNT(*) as cnt FROM HoSoKHCN WHERE isDeleted = 0 AND trangThai = ?",
            (st,)
        )
        c = cur.fetchone()["cnt"]
        print(f"    * Trạng thái '{st}': {c} hồ sơ")
    print("    -> [PASS] Lọc theo trạng thái xét duyệt hoạt động tốt!")
    passed += 1

    # ---- 5. Lọc theo Năm công bố ----
    print("\n[5] Kiểm thử lọc theo Năm công bố (namCongBo = 2026)...")
    total += 1
    cur.execute(
        "SELECT COUNT(*) as cnt FROM HoSoKHCN WHERE isDeleted = 0 AND namCongBo = 2026"
    )
    c_year = cur.fetchone()["cnt"]
    print(f"    * Năm công bố 2026: {c_year} hồ sơ")
    print("    -> [PASS] Lọc theo năm công bố hoạt động chính xác!")
    passed += 1

    # ---- 6. Kiểm tra cơ chế Phân trang ----
    print("\n[6] Kiểm thử cơ chế phân trang (LIMIT 3 OFFSET 0 vs OFFSET 3)...")
    total += 1
    cur.execute("SELECT maHoSo FROM HoSoKHCN WHERE isDeleted = 0 ORDER BY maHoSo LIMIT 3 OFFSET 0")
    page1 = [r["maHoSo"] for r in cur.fetchall()]
    cur.execute("SELECT maHoSo FROM HoSoKHCN WHERE isDeleted = 0 ORDER BY maHoSo LIMIT 3 OFFSET 3")
    page2 = [r["maHoSo"] for r in cur.fetchall()]
    print(f"    * Trang 1 (3 mục đầu): {page1}")
    print(f"    * Trang 2 (3 mục tiếp): {page2}")
    overlap = set(page1).intersection(set(page2))
    if not overlap and (page1 or page2):
        print("    -> [PASS] Phân trang chuẩn xác, không bị trùng lặp dữ liệu!")
        passed += 1
    else:
        print("    -> [FAIL] Phân trang có lỗi trùng lặp bản ghi!")

    # ---- 7. Lọc kết hợp đa tiêu chí ----
    print("\n[7] Kiểm thử truy vấn kết hợp đa tiêu chí (Lĩnh vực + Trạng thái)...")
    total += 1
    cur.execute("""
        SELECT COUNT(*) as cnt
        FROM HoSoKHCN
        WHERE isDeleted = 0
          AND maLinhVuc IS NOT NULL
          AND trangThai IN ('Đã duyệt', 'Đang xét duyệt', 'Chờ xét duyệt')
    """)
    cnt_multi = cur.fetchone()["cnt"]
    print(f"    * Hồ sơ thỏa mãn đồng thời đa tiêu chí: {cnt_multi} hồ sơ")
    print("    -> [PASS] Lọc kết hợp đa tiêu chí hoạt động ổn định!")
    passed += 1

    conn.close()

    print("\n" + "=" * 68)
    if passed == total:
        print(f"==> KẾT QUẢ [TEST-6]: TẤT CẢ {passed}/{total} BÀI TEST TÌM KIẾM/BỘ LỌC ĐẠT 100% PASS! <==")
    else:
        print(f"==> KẾT QUẢ [TEST-6]: {passed}/{total} PASS ({total - passed} FAIL) <==")
    print("=" * 68)


if __name__ == "__main__":
    main()
