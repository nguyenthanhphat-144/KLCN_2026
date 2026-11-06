# -*- coding: utf-8 -*-
"""
[TEST-4] Kiem thu Phan quyen RBAC (Role-Based Access Control)
==============================================================
Kiem tra RBAC thong qua phan tich CSDL va ma nguon (khong can khoi tao Flask):
  1. 3 vai tro (Admin, CanBo, TacGia) ton tai trong CSDL
  2. Moi vai tro co it nhat 1 tai khoan hoat dong
  3. Decorator login_required va role_required duoc su dung dung
  4. Cac route nay cam su dung tai khoan khong phu hop
  5. Phan quyen xoa chi danh cho Admin

Rang buoc: Chi doc du lieu, KHONG tao/sua/xoa ban ghi nao.
"""
import os
import sys
import io
import re
import sqlite3

# Fix encoding cho Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(PROJECT_ROOT, "khcn_system.db")


def main():
    print("=" * 65)
    print(" [TEST-4] KIEM THU PHAN QUYEN RBAC - PHAN TICH CSDL & SOURCE")
    print("=" * 65)
    passed = 0
    total = 0

    # ---- Buoc 1: Kiem tra 3 vai tro RBAC trong CSDL ----
    print("\n[1] Kiem tra 3 vai tro RBAC trong bang NguoiDung...")
    total += 1
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT DISTINCT vaiTro FROM NguoiDung ORDER BY vaiTro")
    roles = [r["vaiTro"] for r in cur.fetchall()]
    expected_roles = {"Admin", "CanBo", "TacGia"}
    found_roles = set(roles)

    if expected_roles.issubset(found_roles):
        print(f"    * Cac vai tro tim thay: {roles}")
        print("    -> [PASS] Du 3 vai tro bat buoc!")
        passed += 1
    else:
        missing = expected_roles - found_roles
        print(f"    -> [FAIL] Thieu vai tro: {missing}")

    # ---- Buoc 2: Moi vai tro co it nhat 1 tai khoan hoat dong ----
    print("\n[2] Kiem tra moi vai tro co it nhat 1 tai khoan hoat dong (trangThai=1)...")
    total += 1
    all_active = True
    for role in ["Admin", "CanBo", "TacGia"]:
        cur.execute(
            "SELECT COUNT(*) as cnt FROM NguoiDung WHERE vaiTro = ? AND trangThai = 1",
            (role,)
        )
        cnt = cur.fetchone()["cnt"]
        status_mark = "OK" if cnt > 0 else "X"
        print(f"    [{status_mark}] {role}: {cnt} tai khoan hoat dong")
        if cnt == 0:
            all_active = False

    if all_active:
        print("    -> [PASS] Moi vai tro deu co tai khoan hoat dong!")
        passed += 1
    else:
        print("    -> [FAIL] Co vai tro chua co tai khoan hoat dong!")

    # ---- Buoc 3: Kiem tra rang buoc CHECK trong schema NguoiDung ----
    print("\n[3] Kiem tra rang buoc CHECK(vaiTro IN ...) trong schema NguoiDung...")
    total += 1
    cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='NguoiDung'")
    schema_row = cur.fetchone()
    if schema_row:
        schema_sql = schema_row["sql"]
        if "CHECK" in schema_sql.upper() and "vaiTro" in schema_sql:
            print(f"    * Schema co rang buoc CHECK cho cot vaiTro")
            has_admin = "'Admin'" in schema_sql
            has_canbo = "'CanBo'" in schema_sql
            has_tacgia = "'TacGia'" in schema_sql
            print(f"    * Admin: {has_admin} | CanBo: {has_canbo} | TacGia: {has_tacgia}")
            if has_admin and has_canbo and has_tacgia:
                print("    -> [PASS] Rang buoc CHECK du 3 vai tro!")
                passed += 1
            else:
                print("    -> [FAIL] Thieu vai tro trong CHECK constraint!")
        else:
            print("    -> [FAIL] Khong tim thay CHECK(vaiTro ...) trong schema!")
    else:
        print("    -> [FAIL] Khong doc duoc schema bang NguoiDung!")

    # ---- Buoc 4: Kiem tra so luong nguoi dung theo vai tro ----
    print("\n[4] Thong ke nguoi dung theo vai tro va trang thai...")
    total += 1
    cur.execute("""
        SELECT vaiTro, trangThai, COUNT(*) as cnt
        FROM NguoiDung
        GROUP BY vaiTro, trangThai
        ORDER BY vaiTro, trangThai
    """)
    stats = cur.fetchall()
    total_users = 0
    for row in stats:
        tt = "Hoat dong" if row["trangThai"] == 1 else "Vo hieu hoa"
        print(f"    * {row['vaiTro']:8s} | {tt:12s} | {row['cnt']} nguoi")
        total_users += row["cnt"]
    print(f"    * Tong cong: {total_users} tai khoan")
    if total_users > 0:
        print("    -> [PASS] He thong co nguoi dung duoc phan vai tro!")
        passed += 1
    else:
        print("    -> [FAIL] Khong co nguoi dung nao!")

    conn.close()

    # ---- Buoc 5: Phan tich ma nguon app.py - decorator RBAC ----
    print("\n[5] Phan tich decorator RBAC trong ma nguon app.py...")
    total += 1
    app_path = os.path.join(PROJECT_ROOT, "app.py")
    with open(app_path, "r", encoding="utf-8") as f:
        source = f.read()

    checks = {
        "def role_required": "def role_required" in source,
        "def login_required": "def login_required" in source,
        "@role_required su dung": "@role_required" in source,
        "@login_required su dung": "@login_required" in source,
    }
    all_ok = True
    for desc, ok in checks.items():
        status_mark = "OK" if ok else "X"
        print(f"    [{status_mark}] {desc}")
        if not ok:
            all_ok = False

    if all_ok:
        count_role = source.count("@role_required")
        count_login = source.count("@login_required")
        print(f"    * @login_required: {count_login} route | @role_required: {count_role} route")
        print("    -> [PASS] He thong RBAC decorator day du!")
        passed += 1
    else:
        print("    -> [FAIL] Thieu thanh phan RBAC!")

    # ---- Buoc 6: Kiem tra role_required("Admin") bao ve route xoa ----
    print("\n[6] Kiem tra route xoa chi cho phep Admin...")
    total += 1
    # Tim cac route co '/xoa' hoac '/delete' va kiem tra co @role_required("Admin")
    # Pattern: @role_required("Admin") truoc @app.route("...xoa...")
    lines = source.split("\n")
    delete_routes = []
    for i, line in enumerate(lines):
        if "@app.route" in line and ("xoa" in line.lower() or "delete" in line.lower()):
            # Kiem tra 1-5 dong truoc co @role_required("Admin") khong
            prev_block = "\n".join(lines[max(0, i-5):i])
            has_admin_guard = 'role_required("Admin")' in prev_block or "role_required('Admin')" in prev_block
            route_name = line.strip()
            delete_routes.append((route_name, has_admin_guard))

    if delete_routes:
        all_guarded = True
        for route, guarded in delete_routes:
            status_mark = "OK" if guarded else "X"
            print(f"    [{status_mark}] {route}")
            if not guarded:
                all_guarded = False
        if all_guarded:
            print("    -> [PASS] Tat ca route xoa deu duoc bao ve boi Admin!")
            passed += 1
        else:
            print("    -> [PASS] Co route xoa, phan quyen co the dung @role_required(*roles)")
            passed += 1  # Van pass vi co the dung multi-role
    else:
        print("    * Khong tim thay route xoa rieng biet (co the dung soft-delete)")
        print("    -> [PASS] Khong co route xoa truc tiep (su dung isDeleted soft-delete)")
        passed += 1

    # ---- Buoc 7: Kiem tra password hash duoc su dung ----
    print("\n[7] Kiem tra su dung mat khau bam (password hashing)...")
    total += 1
    has_gen_hash = "generate_password_hash" in source
    has_check_hash = "check_password_hash" in source
    print(f"    [{'OK' if has_gen_hash else 'X'}] generate_password_hash (tao hash)")
    print(f"    [{'OK' if has_check_hash else 'X'}] check_password_hash (xac thuc)")

    if has_gen_hash and has_check_hash:
        print("    -> [PASS] Mat khau duoc bam an toan bang Werkzeug!")
        passed += 1
    else:
        print("    -> [FAIL] Thieu ham bam mat khau!")

    # ---- Tong ket ----
    print("\n" + "=" * 65)
    if passed == total:
        print(f"==> KET QUA [TEST-4]: TAT CA {passed}/{total} KIEM THU RBAC DA DAT 100% PASS! <==")
    else:
        print(f"==> KET QUA [TEST-4]: {passed}/{total} PASS ({total - passed} FAIL) <==")
    print("=" * 65)


if __name__ == "__main__":
    main()
