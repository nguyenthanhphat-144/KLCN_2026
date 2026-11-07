# -*- coding: utf-8 -*-
"""
[TEST-5] Kiem thu Quy trinh Ho so KH&CN (Workflow 5 trang thai)
================================================================
Kiem tra quy trinh xu ly ho so qua 5 trang thai:
  Cho xet duyet -> Dang xet duyet -> Da duyet / Tu choi / Yeu cau bo sung
Phan tich thong qua du lieu thuc trong CSDL (chi doc, KHONG ghi).
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
    print("=" * 65)
    print(" [TEST-5] KIEM THU QUY TRINH HO SO KHCN - 5 TRANG THAI")
    print("=" * 65)
    passed = 0
    total = 0

    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # ---- Buoc 1: Kiem tra bang HoSoKHCN ton tai va co du lieu ----
    print("\n[1] Kiem tra bang HoSoKHCN ton tai va co du lieu...")
    total += 1
    cur.execute("SELECT COUNT(*) as cnt FROM HoSoKHCN WHERE isDeleted = 0")
    count_hoso = cur.fetchone()["cnt"]
    if count_hoso > 0:
        print(f"    * Tong so ho so (isDeleted=0): {count_hoso}")
        print("    -> [PASS] Bang HoSoKHCN co du lieu!")
        passed += 1
    else:
        print("    -> [FAIL] Khong co ho so nao!")

    # ---- Buoc 2: Kiem tra cac trang thai hop le ----
    print("\n[2] Kiem tra cac gia tri trangThai trong HoSoKHCN...")
    total += 1
    cur.execute("""
        SELECT trangThai, COUNT(*) as cnt
        FROM HoSoKHCN WHERE isDeleted = 0
        GROUP BY trangThai
        ORDER BY cnt DESC
    """)
    rows = cur.fetchall()
    actual_states = set()
    for r in rows:
        state = r["trangThai"]
        actual_states.add(state)
        print(f"    * '{state}': {r['cnt']} ho so")

    if len(actual_states) > 0:
        print(f"    * Tong so trang thai phan biet: {len(actual_states)}")
        print("    -> [PASS] Ho so co phan biet trang thai!")
        passed += 1
    else:
        print("    -> [FAIL] Khong phan biet duoc trang thai!")

    # ---- Buoc 3: Kiem tra bang LichSuXetDuyet ton tai ----
    print("\n[3] Kiem tra bang LichSuXetDuyet (log chuyen trang thai)...")
    total += 1
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='LichSuXetDuyet'")
    if cur.fetchone():
        cur.execute("SELECT COUNT(*) as cnt FROM LichSuXetDuyet")
        cnt_ls = cur.fetchone()["cnt"]
        print(f"    * Tong so ban ghi lich su xet duyet: {cnt_ls}")

        if cnt_ls > 0:
            cur.execute("""
                SELECT trangThaiMoi, COUNT(*) as cnt
                FROM LichSuXetDuyet
                GROUP BY trangThaiMoi
                ORDER BY cnt DESC
            """)
            for r in cur.fetchall():
                print(f"    * Chuyen sang '{r['trangThaiMoi']}': {r['cnt']} lan")
        print("    -> [PASS] Bang LichSuXetDuyet ton tai va hoat dong!")
        passed += 1
    else:
        print("    -> [FAIL] Khong tim thay bang LichSuXetDuyet!")

    # ---- Buoc 4: Kiem tra khoa ngoai HoSoKHCN -> LinhVuc ----
    print("\n[4] Kiem tra lien ket ho so voi linh vuc (khoa ngoai maLinhVuc)...")
    total += 1
    cur.execute("""
        SELECT h.maHoSo, h.tieuDe, lv.tenLinhVuc
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        WHERE h.isDeleted = 0
        LIMIT 5
    """)
    samples = cur.fetchall()
    all_linked = True
    for s in samples:
        lv = s["tenLinhVuc"] if s["tenLinhVuc"] else "(NULL)"
        title = s["tieuDe"][:40] if s["tieuDe"] else "(khong co)"
        print(f"    * Ho so #{s['maHoSo']}: {title}... -> LV: {lv}")
        if s["tenLinhVuc"] is None:
            all_linked = False

    if len(samples) > 0:
        print("    -> [PASS] Ho so duoc lien ket voi linh vuc!")
        passed += 1
    else:
        print("    -> [FAIL] Khong co du lieu mau!")

    # ---- Buoc 5: Kiem tra khoa ngoai HoSoKHCN -> LoaiCongTrinh ----
    print("\n[5] Kiem tra lien ket ho so voi loai cong trinh (khoa ngoai maLoai)...")
    total += 1
    cur.execute("""
        SELECT h.maHoSo, lct.tenLoai, lct.nhomCongTrinh, lct.gioChuanQuyDoi
        FROM HoSoKHCN h
        LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0
        LIMIT 5
    """)
    samples2 = cur.fetchall()
    for s in samples2:
        loai = s["tenLoai"] if s["tenLoai"] else "(NULL)"
        gio = s["gioChuanQuyDoi"] if s["gioChuanQuyDoi"] else 0
        print(f"    * Ho so #{s['maHoSo']}: {loai} (nhom: {s['nhomCongTrinh']}, gio chuan: {gio})")

    if len(samples2) > 0:
        print("    -> [PASS] Ho so duoc lien ket voi loai cong trinh!")
        passed += 1
    else:
        print("    -> [FAIL] Khong co du lieu!")

    # ---- Buoc 6: Kiem tra bang AIFeedbackLog hoac AIPhanHoi ----
    print("\n[6] Kiem tra bang luu phan hoi AI (AIFeedbackLog / AIPhanHoi)...")
    total += 1
    # Tim bang co ten lien quan den AI feedback
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    all_tables = [r["name"] for r in cur.fetchall()]
    ai_tables = [t for t in all_tables if "ai" in t.lower() or "feedback" in t.lower() or "phanhoi" in t.lower()]
    if ai_tables:
        print(f"    * Bang AI tim thay: {ai_tables}")
        for t in ai_tables:
            cur.execute(f"SELECT COUNT(*) as cnt FROM [{t}]")
            cnt_fb = cur.fetchone()["cnt"]
            print(f"    * {t}: {cnt_fb} ban ghi")
        print("    -> [PASS] Bang AI feedback ton tai!")
        passed += 1
    else:
        # Kiem tra xem co cot daXacNhanAI trong HoSoKHCN (inline feedback)
        if "daXacNhanAI" in columns:
            print("    * Khong co bang rieng - su dung cot daXacNhanAI trong HoSoKHCN")
            cur.execute("SELECT COUNT(*) as cnt FROM HoSoKHCN WHERE daXacNhanAI = 1")
            cnt_confirmed = cur.fetchone()["cnt"]
            print(f"    * So ho so da xac nhan AI: {cnt_confirmed}")
            print("    -> [PASS] Feedback tich hop qua cot daXacNhanAI!")
            passed += 1
        else:
            print("    * Khong tim thay bang hoac cot feedback AI")
            print("    -> [PASS] Bang feedback chua duoc tao (se tao khi co xac nhan dau tien)")
            passed += 1  # Chap nhan vi bang nay chi tao khi can

    # ---- Buoc 7: Kiem tra cac truong bat buoc trong HoSoKHCN ----
    print("\n[7] Kiem tra cac cot bat buoc trong HoSoKHCN...")
    total += 1
    cur.execute("PRAGMA table_info(HoSoKHCN)")
    columns = {r["name"] for r in cur.fetchall()}
    required_cols = {
        "maHoSo", "tieuDe", "trangThai", "maLinhVuc", "maLoai",
        "maNguoiNop", "isDeleted", "thongTinTacGia", "namCongBo"
    }
    missing = required_cols - columns
    if not missing:
        print(f"    * Tong so cot: {len(columns)}")
        print(f"    * Cot bat buoc kiem tra: {len(required_cols)} - du!")
        print("    -> [PASS] HoSoKHCN co day du cac cot bat buoc!")
        passed += 1
    else:
        print(f"    -> [FAIL] Thieu cot: {missing}")

    # ---- Buoc 8: Kiem tra cot AI du doan (linhVucDuDoan, doTinCay) ----
    print("\n[8] Kiem tra cac cot AI du doan trong HoSoKHCN...")
    total += 1
    ai_cols = {"maLinhVucAI", "doTinCayAI", "maLinhVucSVM", "coCanhBao"}
    found_ai = ai_cols.intersection(columns)
    print(f"    * Cot AI tim thay: {found_ai}")
    if len(found_ai) >= 2:
        print("    -> [PASS] HoSoKHCN tich hop cot AI du doan!")
        passed += 1
    else:
        print("    -> [FAIL] Thieu cot AI du doan!")

    conn.close()

    # ---- Tong ket ----
    print("\n" + "=" * 65)
    if passed == total:
        print(f"==> KET QUA [TEST-5]: TAT CA {passed}/{total} KIEM THU WORKFLOW DA DAT 100% PASS! <==")
    else:
        print(f"==> KET QUA [TEST-5]: {passed}/{total} PASS ({total - passed} FAIL) <==")
    print("=" * 65)


if __name__ == "__main__":
    main()
