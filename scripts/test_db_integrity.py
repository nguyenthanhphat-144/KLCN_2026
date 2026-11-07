# -*- coding: utf-8 -*-
"""
Kịch bản kiểm thử độc lập: Toàn vẹn Cơ sở dữ liệu & Phân quyền RBAC (CNTT-KLCN001)
Tác giả kiểm thử: Lê Thanh Phong (kbt2981-spec)
Mục đích:
- Kiểm tra tính sẵn sàng và cấu trúc bảng của CSDL SQLite/SQL Server.
- Đảm bảo đầy đủ 8 bảng nghiệp vụ cốt lõi theo đặc tả PRD.
- Xác thực dữ liệu danh mục chuẩn và tài khoản người dùng với 3 vai trò RBAC.
"""

import os
import sys
import sqlite3

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "khcn_system.db")


def test_database_structure_and_integrity():
    print("=" * 65)
    print(" [TEST-1] KIỂM THỬ CẤU TRÚC VÀ TÍNH TOÀN VẸN CƠ SỞ DỮ LIỆU")
    print("=" * 65)

    assert os.path.exists(DB_PATH), f"Không tìm thấy file CSDL tại {DB_PATH}"
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # 1. Kiểm tra danh sách bảng bắt buộc theo PRD
    required_tables = [
        "NguoiDung",
        "LinhVuc",
        "LoaiCongTrinh",
        "TapChiHopLe",
        "HoSoKHCN",
        "LichSuXetDuyet",
        "PhanHoiAI",
        "ThucNghiemMoHinhAI",
    ]

    cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
    existing_tables = [row[0] for row in cur.fetchall()]

    print("[1] Kiểm tra sự tồn tại của 8 bảng cốt lõi:")
    for tbl in required_tables:
        status = "OK" if tbl in existing_tables else "THIẾU"
        print(f"    * Bảng '{tbl}': {status}")
        assert tbl in existing_tables, f"Thiếu bảng bắt buộc: {tbl}"

    # 2. Kiểm tra tài khoản người dùng và phân quyền RBAC
    print("\n[2] Kiểm tra người dùng phân quyền theo 3 vai trò (RBAC):")
    cur.execute("SELECT email, vaiTro, hoTen FROM NguoiDung ORDER BY maNguoiDung ASC;")
    users = cur.fetchall()
    roles = {u["vaiTro"] for u in users}

    print(f"    -> Tổng số tài khoản: {len(users)}")
    print(f"    -> Các vai trò ghi nhận: {', '.join(sorted(roles))}")
    assert "Admin" in roles, "Thiếu vai trò Quản trị viên (Admin)"
    assert "CanBo" in roles, "Thiếu vai trò Cán bộ xét duyệt (CanBo)"
    assert "TacGia" in roles, "Thiếu vai trò Tác giả / Giảng viên (TacGia)"

    # 3. Kiểm tra danh mục lĩnh vực và loại công trình
    print("\n[3] Kiểm tra dữ liệu danh mục chuẩn hóa:")
    cur.execute("SELECT COUNT(*) FROM LinhVuc;")
    lv_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM LoaiCongTrinh;")
    lct_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM TapChiHopLe;")
    tc_count = cur.fetchone()[0]

    print(f"    -> Số lượng Lĩnh vực nghiên cứu: {lv_count} (Yêu cầu: >= 6 lĩnh vực)")
    print(f"    -> Số lượng Loại công trình: {lct_count} (Bao gồm Bài báo, Đề tài, Sáng chế)")
    print(f"    -> Số lượng Tạp chí hợp lệ: {tc_count} tạp chí chuẩn Scopus/WoS")
    assert lv_count >= 6, "Thiếu dữ liệu lĩnh vực chuẩn"
    assert lct_count >= 5, "Thiếu dữ liệu loại công trình"

    # 4. Kiểm tra hồ sơ KH&CN và ràng buộc xóa mềm
    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0;")
    active_hoso = cur.fetchone()[0]
    print(f"\n[4] Số lượng hồ sơ KH&CN hiện hữu (isDeleted=0): {active_hoso} hồ sơ")

    # 5. Kiểm tra kết quả thực nghiệm mô hình AI
    cur.execute("SELECT tenMoHinh, f1Score, thoiGianPhanHoiMs FROM ThucNghiemMoHinhAI;")
    ai_rows = cur.fetchall()
    print(f"\n[5] Mô hình AI đã ghi nhận trong bảng ThucNghiemMoHinhAI: {len(ai_rows)} mô hình")
    for row in ai_rows:
        print(f"    * {row['tenMoHinh']}: F1-score={row['f1Score']}%, Độ trễ={row['thoiGianPhanHoiMs']}ms")

    conn.close()
    print("\n==> KẾT QUẢ [TEST-1]: TẤT CẢ KIỂM THỬ CSDL ĐÃ ĐẠT 100% PASS! <==\n")


if __name__ == "__main__":
    test_database_structure_and_integrity()
