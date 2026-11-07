# -*- coding: utf-8 -*-
"""
Đề tài: Hệ thống Hỗ trợ Quản lý Hồ sơ KH&CN Tích hợp Phân loại Tự động (CNTT-KLCN001)

Module database.py – Tầng truy cập dữ liệu (Data Access Layer)
----------------------------------------------------------------
Quản lý kết nối và khởi tạo lược đồ CSDL SQLite (khcn_system.db).
Hỗ trợ tùy chọn chuyển sang Microsoft SQL Server (SSMS) qua db_config.py
nhờ các lớp bọc SQLServerCursorWrapper / SQLServerConnectionWrapper tự động
dịch cú pháp LIMIT ↔ TOP / OFFSET FETCH.

Cấu trúc 8 bảng chính:
  1. NguoiDung         – Tài khoản người dùng (Admin / CanBo / TacGia)
  2. LinhVuc           – Danh mục 6 lĩnh vực KH&CN
  3. LoaiCongTrinh     – Danh mục 6 loại công trình (Bài báo, Sách, Đề tài, ...)
  4. TapChiHopLe       – Danh mục 10 tạp chí uy tín (ISSN, xếp hạng Q1–Q4)
  5. DanhMucTacGia     – Danh mục tác giả nghiên cứu
  6. HoSoKHCN          – Hồ sơ khoa học công nghệ (bảng trung tâm)
  7. LichSuXetDuyet    – Log chuyển trạng thái xét duyệt hồ sơ
  8. AIFeedbackLog     – Log phản hồi người dùng với dự đoán AI (Feedback Loop)

Quan hệ khóa ngoại:
  HoSoKHCN.maLinhVuc      → LinhVuc.maLinhVuc
  HoSoKHCN.maLoai         → LoaiCongTrinh.maLoai
  HoSoKHCN.maTapChi       → TapChiHopLe.maTapChi
  HoSoKHCN.maNguoiNop     → NguoiDung.maNguoiDung
  LichSuXetDuyet.maHoSo   → HoSoKHCN.maHoSo
  AIFeedbackLog.maHoSo     → HoSoKHCN.maHoSo
"""

import os
import re
import sqlite3
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "khcn_system.db")


class SQLServerRow(dict):
    """Lớp bọc dòng dữ liệu trả về từ pyodbc (SQL Server) để tương thích hoàn toàn với sqlite3.Row."""
    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self.values())[key]
        return super().__getitem__(key)


class SQLServerCursorWrapper:
    """Bộ chuyển đổi cú pháp truy vấn tự động giữa SQLite và T-SQL (SQL Server SSMS)."""
    def __init__(self, raw_cursor):
        self._cur = raw_cursor

    @staticmethod
    def _translate_sql(sql: str, params=()):
        translated = sql.strip()
        params_list = list(params) if params else []

        # Chuyển SELECT ... LIMIT <number> ở cuối câu lệnh sang SELECT TOP <number> ...
        m_limit_const = re.search(r"\bLIMIT\s+(\d+)\s*;?\s*$", translated, flags=re.IGNORECASE)
        if m_limit_const and "OFFSET" not in translated.upper():
            top_n = m_limit_const.group(1)
            translated = re.sub(r"\bLIMIT\s+\d+\s*;?\s*$", "", translated, flags=re.IGNORECASE)
            translated = re.sub(r"^SELECT\s+", f"SELECT TOP {top_n} ", translated, count=1, flags=re.IGNORECASE)

        # Chuyển SELECT ... LIMIT ? ở cuối câu lệnh sang SELECT TOP <val> ...
        m_limit_param = re.search(r"\bLIMIT\s+\?\s*;?\s*$", translated, flags=re.IGNORECASE)
        if m_limit_param and "OFFSET" not in translated.upper() and params_list:
            top_n = params_list.pop()
            translated = re.sub(r"\bLIMIT\s+\?\s*;?\s*$", "", translated, flags=re.IGNORECASE)
            translated = re.sub(r"^SELECT\s+", f"SELECT TOP {int(top_n)} ", translated, count=1, flags=re.IGNORECASE)

        # Chuyển LIMIT ? OFFSET ? (SQLite) sang OFFSET ? ROWS FETCH NEXT ? ROWS ONLY (SQL Server)
        m_limit_offset = re.search(r"\bLIMIT\s+\?\s+OFFSET\s+\?\s*;?\s*$", translated, flags=re.IGNORECASE)
        if m_limit_offset and len(params_list) >= 2:
            limit_val = params_list[-2]
            offset_val = params_list[-1]
            params_list[-2] = offset_val
            params_list[-1] = limit_val
            translated = re.sub(
                r"\bLIMIT\s+\?\s+OFFSET\s+\?\s*;?\s*$",
                "OFFSET ? ROWS FETCH NEXT ? ROWS ONLY",
                translated,
                flags=re.IGNORECASE,
            )

        return translated, tuple(params_list)

    def execute(self, sql: str, params=()):
        tsql, tparams = self._translate_sql(sql, params)
        if tparams:
            self._cur.execute(tsql, tparams)
        else:
            self._cur.execute(tsql)
        return self

    def fetchone(self):
        row = self._cur.fetchone()
        if row is None:
            return None
        cols = [col[0] for col in self._cur.description]
        return SQLServerRow(zip(cols, row))

    def fetchall(self):
        rows = self._cur.fetchall()
        if not rows or not self._cur.description:
            return []
        cols = [col[0] for col in self._cur.description]
        return [SQLServerRow(zip(cols, r)) for r in rows]

    def __iter__(self):
        return iter(self.fetchall())


class SQLServerConnectionWrapper:
    """Lớp bọc kết nối pyodbc cho Microsoft SQL Server (SSMS)."""
    def __init__(self, raw_conn):
        self._conn = raw_conn

    def cursor(self):
        return SQLServerCursorWrapper(self._conn.cursor())

    def execute(self, sql: str, params=()):
        cur = self.cursor()
        return cur.execute(sql, params)

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


_MIGRATED_ONCE = False


def get_db_connection():
    """
    Kết nối tới SQLite (khcn_system.db) theo đề cương CNTT-KLCN001,
    hoặc Microsoft SQL Server (SSMS) nếu bật USE_SQL_SERVER = True trong db_config.py.
    """
    global _MIGRATED_ONCE
    try:
        import db_config
        if getattr(db_config, "USE_SQL_SERVER", False):
            import pyodbc
            conn_str = db_config.get_sql_server_connection_string()
            raw_conn = pyodbc.connect(conn_str, timeout=3)
            return SQLServerConnectionWrapper(raw_conn)
    except Exception as exc:
        print(f"[DB Notice] Không kết nối được SQL Server ({exc}), tự động chuyển sang SQLite.")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    if not _MIGRATED_ONCE and os.path.exists(DB_PATH):
        try:
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='HoSoKHCN'")
            if cur.fetchone():
                ensure_schema_migrations(conn)
                _MIGRATED_ONCE = True
        except Exception:
            pass
    return conn


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    cur = conn.cursor()

    # 1. Bảng NguoiDung (Tuần 3 - 1a & Tuần 4 - 1a, 1b, 1c)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS NguoiDung (
        maNguoiDung INTEGER PRIMARY KEY AUTOINCREMENT,
        hoTen TEXT NOT NULL,
        email TEXT NOT NULL UNIQUE,
        matKhau TEXT NOT NULL,
        vaiTro TEXT NOT NULL CHECK(vaiTro IN ('Admin', 'CanBo', 'TacGia')),
        donViCongTac TEXT NOT NULL,
        hocHamHocVi TEXT DEFAULT 'ThS.',
        trangThai INTEGER DEFAULT 1,
        ngayTao DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 2. Bảng danh mục LinhVuc (Tuần 3 - 1a & Tuần 5 - 1a)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS LinhVuc (
        maLinhVuc INTEGER PRIMARY KEY AUTOINCREMENT,
        maCode TEXT NOT NULL UNIQUE,
        tenLinhVuc TEXT NOT NULL,
        moTa TEXT,
        tuKhoaDacTrung TEXT,
        trangThai INTEGER DEFAULT 1
    );
    """)

    # 3. Bảng danh mục LoaiCongTrinh (Tuần 3 - 1a & Tuần 5 - 1a)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS LoaiCongTrinh (
        maLoai INTEGER PRIMARY KEY AUTOINCREMENT,
        maCodeLoai TEXT NOT NULL UNIQUE,
        tenLoai TEXT NOT NULL,
        nhomCongTrinh TEXT NOT NULL,
        gioChuanQuyDoi REAL DEFAULT 0.0,
        trangThai INTEGER DEFAULT 1
    );
    """)

    # 4. Bảng danh mục TapChiHopLe (Tuần 5 - 1b)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS TapChiHopLe (
        maTapChi INTEGER PRIMARY KEY AUTOINCREMENT,
        tenTapChi TEXT NOT NULL,
        chiSoISSN TEXT NOT NULL,
        xepHang TEXT NOT NULL,
        maLinhVuc INTEGER,
        uyTin INTEGER DEFAULT 1,
        trangThai INTEGER DEFAULT 1,
        FOREIGN KEY (maLinhVuc) REFERENCES LinhVuc(maLinhVuc)
    );
    """)

    # 4b. Bảng danh mục Tác giả - DanhMucTacGia (Chức năng 2 theo Đề cương CNTT-KLCN001)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS DanhMucTacGia (
        maTacGia INTEGER PRIMARY KEY AUTOINCREMENT,
        maCodeTG TEXT NOT NULL UNIQUE,
        hoTen TEXT NOT NULL,
        hocHamHocVi TEXT DEFAULT 'TS.',
        email TEXT,
        donViCongTac TEXT NOT NULL,
        maLinhVucChinh INTEGER,
        orcid TEXT DEFAULT '',
        trangThai INTEGER DEFAULT 1,
        FOREIGN KEY (maLinhVucChinh) REFERENCES LinhVuc(maLinhVuc)
    );
    """)

    # 5. Bảng HoSoKHCN (Tuần 3 - 1a, 1b & Tuần 4 - 2a, 2b, 2c & Tuần 6 - 2d)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS HoSoKHCN (
        maHoSo TEXT PRIMARY KEY,
        tieuDe TEXT NOT NULL,
        tomTat TEXT NOT NULL,
        tuKhoa TEXT NOT NULL,
        thongTinTacGia TEXT NOT NULL,
        donViChuTri TEXT NOT NULL,
        namCongBo INTEGER DEFAULT 2026,
        kyBaoCao TEXT DEFAULT 'Đợt 1 - Năm học 2025-2026',
        maNguoiNop INTEGER NOT NULL,
        maLinhVuc INTEGER NOT NULL,
        maLoai INTEGER NOT NULL,
        maTapChi INTEGER,
        trangThai TEXT DEFAULT 'Chờ xét duyệt' CHECK(trangThai IN ('Chờ xét duyệt', 'Đang xét duyệt', 'Đã duyệt', 'Từ chối', 'Yêu cầu bổ sung')),
        fileDinhKem TEXT,
        isDeleted INTEGER DEFAULT 0,
        -- Các trường nghiệp vụ đặc thù cho Bài báo khoa học (Revision 2026)
        doi TEXT,
        tapSo TEXT,
        trangBaiBao TEXT,
        chiMucTapChi TEXT,
        -- Các trường nghiệp vụ đặc thù cho Đề tài NCKH (Revision 2026)
        maDeTai TEXT,
        capDeTai TEXT,
        chuNhiemDeTai TEXT,
        thanhVienThamGia TEXT,
        thoiGianThucHien TEXT,
        kinhPhi REAL DEFAULT 0.0,
        sanPhamDuKien TEXT,
        ketQuaNghiemThu TEXT DEFAULT 'Đang thực hiện',
        -- Các trường tích hợp AI Phân loại (Chức năng 5: Lĩnh vực & Loại công trình) & Phát hiện bất thường (Chức năng 6)
        vanBanDaTachTu TEXT,
        maLinhVucAI INTEGER,
        doTinCayAI REAL DEFAULT 0.0,
        maLinhVucSVM INTEGER,
        doTinCaySVM REAL DEFAULT 0.0,
        nhomCongTrinhAI TEXT DEFAULT '',
        doTinCayLoaiAI REAL DEFAULT 0.0,
        nhomCongTrinhSVM TEXT DEFAULT '',
        doTinCayLoaiSVM REAL DEFAULT 0.0,
        moHinhChinh TEXT DEFAULT 'PhoBERT (vinai/phobert-base)',
        daXacNhanAI INTEGER DEFAULT 0,
        diemBatThuong REAL DEFAULT 0.0,
        coCanhBao INTEGER DEFAULT 0,
        chiTietCanhBao TEXT DEFAULT '[]',
        doTuongDongMax REAL DEFAULT 0.0,
        hoSoTuongDongNhat TEXT DEFAULT '',
        ngayNop DATETIME DEFAULT CURRENT_TIMESTAMP,
        ngayCapNhat DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (maNguoiNop) REFERENCES NguoiDung(maNguoiDung),
        FOREIGN KEY (maLinhVuc) REFERENCES LinhVuc(maLinhVuc),
        FOREIGN KEY (maLoai) REFERENCES LoaiCongTrinh(maLoai),
        FOREIGN KEY (maTapChi) REFERENCES TapChiHopLe(maTapChi),
        FOREIGN KEY (maLinhVucAI) REFERENCES LinhVuc(maLinhVuc),
        FOREIGN KEY (maLinhVucSVM) REFERENCES LinhVuc(maLinhVuc)
    );
    """)

    # 6. Bảng LichSuXetDuyet (Tuần 3 - 1a & Tuần 5 - 2a, 2b)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS LichSuXetDuyet (
        maLichSu INTEGER PRIMARY KEY AUTOINCREMENT,
        maHoSo TEXT NOT NULL,
        maCanBo INTEGER NOT NULL,
        trangThaiCu TEXT,
        trangThaiMoi TEXT NOT NULL,
        yKienDanhGia TEXT NOT NULL,
        thoiGian DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (maHoSo) REFERENCES HoSoKHCN(maHoSo) ON DELETE CASCADE,
        FOREIGN KEY (maCanBo) REFERENCES NguoiDung(maNguoiDung)
    );
    """)

    # 7. Bảng PhanHoiAI - Feedback Loop (Tuần 6 - 2d)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS PhanHoiAI (
        maPhanHoi INTEGER PRIMARY KEY AUTOINCREMENT,
        maHoSo TEXT NOT NULL,
        maCanBo INTEGER NOT NULL,
        maLinhVucAI INTEGER NOT NULL,
        maLinhVucThucTe INTEGER NOT NULL,
        doTinCayLucDo REAL NOT NULL,
        ketQuaKhop INTEGER NOT NULL,
        vanBanHuanLuyen TEXT NOT NULL,
        ghiChuPhanHoi TEXT,
        thoiGian DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (maHoSo) REFERENCES HoSoKHCN(maHoSo) ON DELETE CASCADE,
        FOREIGN KEY (maCanBo) REFERENCES NguoiDung(maNguoiDung),
        FOREIGN KEY (maLinhVucAI) REFERENCES LinhVuc(maLinhVuc),
        FOREIGN KEY (maLinhVucThucTe) REFERENCES LinhVuc(maLinhVuc)
    );
    """)

    # 8. Bảng ThucNghiemMoHinhAI (Tuần 6 - 2c)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS ThucNghiemMoHinhAI (
        maThucNghiem INTEGER PRIMARY KEY AUTOINCREMENT,
        tenMoHinh TEXT NOT NULL,
        precisionScore REAL NOT NULL,
        recallScore REAL NOT NULL,
        f1Score REAL NOT NULL,
        accuracyScore REAL NOT NULL,
        thoiGianPhanHoiMs REAL NOT NULL,
        soMauHuanLuyen INTEGER NOT NULL,
        soMauKiemThu INTEGER NOT NULL,
        moTaChiTiet TEXT,
        ngayDanhGia DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    """)

    conn.commit()
    ensure_schema_columns_and_tables(conn)
    seed_initial_data(conn)
    seed_author_catalog_if_empty(conn)
    conn.close()


def ensure_schema_columns_and_tables(conn):
    """Đảm bảo CSDL hiện hữu tự động được bổ sung đầy đủ cột & bảng DanhMucTacGia mà không mất dữ liệu."""
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS DanhMucTacGia (
        maTacGia INTEGER PRIMARY KEY AUTOINCREMENT,
        maCodeTG TEXT NOT NULL UNIQUE,
        hoTen TEXT NOT NULL,
        hocHamHocVi TEXT DEFAULT 'TS.',
        email TEXT,
        donViCongTac TEXT NOT NULL,
        maLinhVucChinh INTEGER,
        orcid TEXT DEFAULT '',
        trangThai INTEGER DEFAULT 1,
        FOREIGN KEY (maLinhVucChinh) REFERENCES LinhVuc(maLinhVuc)
    );
    """)
    cur.execute("PRAGMA table_info(HoSoKHCN)")
    existing_cols = {r[1] for r in cur.fetchall()}
    alter_columns = [
        ("nhomCongTrinhAI", "TEXT DEFAULT ''"),
        ("doTinCayLoaiAI", "REAL DEFAULT 0.0"),
        ("nhomCongTrinhSVM", "TEXT DEFAULT ''"),
        ("doTinCayLoaiSVM", "REAL DEFAULT 0.0"),
    ]
    for col_name, col_def in alter_columns:
        if col_name not in existing_cols:
            cur.execute(f"ALTER TABLE HoSoKHCN ADD COLUMN {col_name} {col_def}")
    conn.commit()


def seed_author_catalog_if_empty(conn):
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM LinhVuc")
    if cur.fetchone()[0] == 0:
        return
    cur.execute("SELECT COUNT(*) FROM DanhMucTacGia")
    if cur.fetchone()[0] == 0:
        tac_gias = [
            ("TG-001", "Phạm Văn Khoa", "TS.", "khoapv@uni.edu.vn", "Khoa Công nghệ Thông tin", 1, "0000-0002-1452-8801"),
            ("TG-002", "Trần Quốc Bảo", "PGS.TS.", "baotq@uni.edu.vn", "Khoa Công nghệ Thông tin", 1, "0000-0001-9834-2104"),
            ("TG-003", "Vũ Đình Trọng", "TS.", "trongvd@uni.edu.vn", "Khoa Điện - Điện tử", 2, "0000-0003-4120-5567"),
            ("TG-004", "Lê Hoàng Việt", "PGS.TS.", "vietlh@uni.edu.vn", "Khoa Điện - Điện tử", 2, "0000-0002-7719-3390"),
            ("TG-005", "Hoàng Gia Huy", "TS.", "huyhg@uni.edu.vn", "Khoa Cơ khí Động lực", 3, "0000-0001-6650-7812"),
            ("TG-006", "Nguyễn Đức Mạnh", "TS.", "manhnd@uni.edu.vn", "Khoa Cơ khí Động lực", 3, "0000-0002-8941-1029"),
            ("TG-007", "Đặng Thị Mai Hương", "ThS.", "huongdtm@uni.edu.vn", "Khoa Kinh tế & Quản trị Kinh doanh", 4, "0000-0003-1198-4502"),
            ("TG-008", "Lê Thu Trang", "TS.", "tranglt@uni.edu.vn", "Khoa Kinh tế & Quản trị Kinh doanh", 4, "0000-0002-5401-6789"),
            ("TG-009", "Ngô Thanh Hà", "PGS.TS.", "hant@uni.edu.vn", "Khoa Kỹ thuật Hóa học & Môi trường", 5, "0000-0001-3349-9012"),
            ("TG-010", "Đỗ Văn Kiên", "TS.", "kiendv@uni.edu.vn", "Khoa Khoa học Xã hội & Nhân văn", 6, "0000-0003-8821-3456"),
        ]
        cur.executemany("""
            INSERT INTO DanhMucTacGia (maCodeTG, hoTen, hocHamHocVi, email, donViCongTac, maLinhVucChinh, orcid)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, tac_gias)
        conn.commit()


def seed_initial_data(conn):
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM NguoiDung")
    if cur.fetchone()[0] > 0:
        seed_sample_records_with_ai(conn)
        return

    # 1. Seed NguoiDung (Mật khẩu mặc định được hash bảo mật bằng Werkzeug)
    users = [
        ("PGS.TS. Trần Quốc Bảo (Quản trị viên)", "admin@uni.edu.vn", generate_password_hash("Admin@123"), "Admin", "Phòng Quản lý Khoa học & HTQT", "PGS.TS."),
        ("TS. Nguyễn Minh Tâm (Cán bộ duyệt)", "canbo@uni.edu.vn", generate_password_hash("CanBo@123"), "CanBo", "Hội đồng KH&CN Trường", "TS."),
        ("PGS.TS. Lê Hoàng Việt (Cán bộ duyệt)", "vietlh@uni.edu.vn", generate_password_hash("CanBo@123"), "CanBo", "Phòng Quản lý Khoa học & HTQT", "PGS.TS."),
        ("TS. Phạm Văn Khoa (Giảng viên CNTT)", "tacgia@uni.edu.vn", generate_password_hash("TacGia@123"), "TacGia", "Khoa Công nghệ Thông tin", "TS."),
        ("ThS. Đặng Thị Mai Hương (GV Kinh tế)", "huongdtm@uni.edu.vn", generate_password_hash("TacGia@123"), "TacGia", "Khoa Kinh tế & Quản trị Kinh doanh", "ThS."),
        ("TS. Vũ Đình Trọng (GV Điện - ĐT)", "trongvd@uni.edu.vn", generate_password_hash("TacGia@123"), "TacGia", "Khoa Điện - Điện tử", "TS."),
        ("PGS.TS. Ngô Thanh Hà (GV Hóa - Sinh)", "hant@uni.edu.vn", generate_password_hash("TacGia@123"), "TacGia", "Khoa Kỹ thuật Hóa học & Môi trường", "PGS.TS."),
        ("TS. Hoàng Gia Huy (GV Cơ khí)", "huyhg@uni.edu.vn", generate_password_hash("TacGia@123"), "TacGia", "Khoa Cơ khí Động lực", "TS."),
    ]
    cur.executemany("""
        INSERT INTO NguoiDung (hoTen, email, matKhau, vaiTro, donViCongTac, hocHamHocVi)
        VALUES (?, ?, ?, ?, ?, ?)
    """, users)

    # 2. Seed Danh mục LinhVuc (6 lĩnh vực KH&CN chuẩn)
    linh_vucs = [
        ("LV-CNTT", "Công nghệ thông tin và Truyền thông",
         "Trí tuệ nhân tạo, học máy, học sâu, xử lý ngôn ngữ tự nhiên, thị giác máy tính, an toàn thông tin, mạng máy tính, kỹ thuật phần mềm, dữ liệu lớn.",
         "trí_tuệ_nhân_tạo học_máy học_sâu phobert xử_lý_ngôn_ngữ_tự_nhiên thị_giác_máy_tính mạng_nơ_ron dữ_liệu_lớn an_toàn_thông_tin chuỗi_khối phần_mềm đám_mây phân_loại_văn_bản"),
        ("LV-DDT", "Kỹ thuật Điện - Điện tử và Tự động hóa",
         "Hệ thống nhúng, vi mạch bán dẫn, điều khiển tự động, mạng cảm biến không dây IoT, truyền thông vô tuyến 5G/6G, lưới điện thông minh, điện tử công suất.",
         "hệ_thống_nhúng vi_mạch bán_dẫn điều_khiển_tự_động cảm_biến_không_dây lưới_điện_thông_minh biến_tần_công_suất ăng_ten vi_điều_khiển plc_scada năng_lượng_mặt_trời"),
        ("LV-CKD", "Kỹ thuật Cơ khí, Chế tạo máy và Giao thông",
         "Thiết kế chế tạo máy, cơ điện tử, động cơ đốt trong, ô tô điện, gia công CNC, mô phỏng phần tử hữu hạn FEM, vật liệu cơ khí, khí động học.",
         "chế_tạo_máy cơ_điện_tử động_cơ_đốt_trong gia_công_cnc phần_tử_hữu_hạn kết_cấu_thép truyền_động_thủy_lực ma_sát_mài_mòn cánh_tay_robot_cơ_khí ô_tô_điện"),
        ("LV-KTQL", "Kinh tế, Quản trị kinh doanh và Tài chính",
         "Kinh tế vĩ mô, tài chính ngân hàng, kế toán kiểm toán, quản trị chuỗi cung ứng, thương mại điện tử, marketing số, hành vi người tiêu dùng, kinh tế xanh.",
         "kinh_tế_vĩ_mô tài_chính_ngân_hàng chuỗi_cung_ứng chứng_khoán quản_trị_nhân_lực hành_vi_người_tiêu_dùng thương_mại_điện_tử tăng_trưởng_kinh_tế lạm_phát kiểm_toán"),
        ("LV-HSMT", "Hóa học, Sinh học, Y Dược và Môi trường",
         "Công nghệ sinh học, vật liệu nano, tổng hợp hữu cơ, xử lý nước thải, biến đổi khí hậu, dược liệu học, phân tích độc chất, sinh học phân tử.",
         "công_nghệ_sinh_học vật_liệu_nano tổng_hợp_hữu_cơ xử_lý_nước_thải kim_loại_nặng dược_liệu kháng_khuẩn xúc_tác_quang_hóa enzyme_sinh_học ô_nhiễm_môi_trường"),
        ("LV-XHNV", "Khoa học Xã hội, Nhân văn và Giáo dục học",
         "Phương pháp giảng dạy đại học, tâm lý học giáo dục, ngôn ngữ học đối chiếu, văn hóa học, xã hội học, chính sách công, pháp luật đại cương.",
         "khoa_học_giáo_dục phương_pháp_giảng_dạy tâm_lý_học ngôn_ngữ_học văn_hóa_dân_gian xã_hội_học năng_lực_người_học chương_trình_đào_tạo giáo_dục_đại_học")
    ]
    cur.executemany("""
        INSERT INTO LinhVuc (maCode, tenLinhVuc, moTa, tuKhoaDacTrung)
        VALUES (?, ?, ?, ?)
    """, linh_vucs)

    # 3. Seed Danh mục LoaiCongTrinh
    loai_cts = [
        ("LCT-ISI-Q1", "Bài báo tạp chí quốc tế Web of Science / Scopus (Q1/Q2)", "Bài báo quốc tế", 180.0),
        ("LCT-SCOPUS", "Bài báo tạp chí quốc tế Scopus (Q3/Q4) & Hội nghị quốc tế uy tín", "Bài báo quốc tế", 120.0),
        ("LCT-HDGSNN", "Bài báo tạp chí khoa học trong nước thuộc danh mục HĐGSNN", "Bài báo trong nước", 60.0),
        ("LCT-DT-BO", "Đề tài nghiên cứu khoa học cấp Bộ / cấp Tỉnh - Thành phố", "Đề tài NCKH", 240.0),
        ("LCT-DT-TRG", "Đề tài nghiên cứu khoa học cấp Trường / Đề tài Sinh viên", "Đề tài NCKH", 90.0),
        ("LCT-SANGCHE", "Bằng độc quyền Sáng chế / Giải pháp hữu ích / Sách chuyên khảo", "Sở hữu trí tuệ & Sách", 200.0),
    ]
    cur.executemany("""
        INSERT INTO LoaiCongTrinh (maCodeLoai, tenLoai, nhomCongTrinh, gioChuanQuyDoi)
        VALUES (?, ?, ?, ?)
    """, loai_cts)

    # 4. Seed Danh mục TapChiHopLe (Bao gồm cả tạp chí uy tín và tạp chí cảnh báo/không uy tín)
    tap_chis = [
        ("Engineering Applications of Artificial Intelligence", "0952-1976", "ISI-Q1 / Scopus Q1", 1, 1),
        ("Journal of Computer Science and Cybernetics (VAST)", "1813-9663", "HĐGSNN (1.0 điểm)", 1, 1),
        ("IEEE Transactions on Industrial Electronics", "0278-0046", "ISI-Q1 / Scopus Q1", 2, 1),
        ("International Journal of Mechanical Sciences", "0020-7403", "ISI-Q1 / Scopus Q1", 3, 1),
        ("Journal of Asian Economics & Management", "1049-0078", "Scopus Q2 / HĐGSNN", 4, 1),
        ("Journal of Cleaner Production & Environmental Tech", "0959-6526", "ISI-Q1 / Scopus Q1", 5, 1),
        ("Tạp chí Khoa học Giáo dục Việt Nam", "2615-8957", "HĐGSNN (0.75 điểm)", 6, 1),
        ("Hội đồng Nghiệm thu Đề tài KH&CN Cấp Trường / Cấp Bộ", "HDNT-VN", "Quyết định nghiệm thu", 1, 1),
        ("Tạp chí Đa ngành Quốc tế Chưa Kiểm Định (Cảnh báo)", "9999-001X", "Ngoài danh mục uy tín", 1, 0),
        ("Tạp chí Công bố Nhanh Đã Ngừng Chỉ Mục", "8888-404X", "Ngừng chỉ mục / Không uy tín", 4, 0),
    ]
    cur.executemany("""
        INSERT INTO TapChiHopLe (tenTapChi, chiSoISSN, xepHang, maLinhVuc, uyTin)
        VALUES (?, ?, ?, ?, ?)
    """, tap_chis)

    conn.commit()
    seed_sample_records_with_ai(conn)


def seed_sample_records_with_ai(conn):
    from ai_engine import ai_engine

    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM HoSoKHCN")
    if cur.fetchone()[0] > 0:
        ai_engine.train_and_evaluate(conn)
        cur.execute("SELECT maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri FROM HoSoKHCN WHERE isDeleted = 0")
        for r in cur.fetchall():
            pred = ai_engine.predict_record_field(
                r["tieuDe"], r["tomTat"], r["tuKhoa"], r["thongTinTacGia"], r["donViChuTri"]
            )
            cur.execute("""
                UPDATE HoSoKHCN
                SET vanBanDaTachTu = ?,
                    maLinhVucAI = ?,
                    doTinCayAI = ?,
                    maLinhVucSVM = ?,
                    doTinCaySVM = ?,
                    nhomCongTrinhAI = ?,
                    doTinCayLoaiAI = ?,
                    nhomCongTrinhSVM = ?,
                    doTinCayLoaiSVM = ?
                WHERE maHoSo = ?
            """, (
                pred["segmented_text"],
                pred["phobert_label_id"],
                pred["phobert_confidence"],
                pred["svm_label_id"],
                pred["svm_confidence"],
                pred["work_type_phobert"],
                pred["work_type_phobert_conf"],
                pred["work_type_svm"],
                pred["work_type_svm_conf"],
                r["maHoSo"],
            ))
        conn.commit()
        ai_engine.build_networkx_graph_and_detect_anomalies(conn)
        return

    ai_engine.train_and_evaluate(conn)

    sample_hoso = [
        # 1. Hồ sơ CNTT chuẩn (Độ tin cậy > 80% - Xanh lá, Đã duyệt)
        ("HS-2026-0001",
         "Nghiên cứu tinh chỉnh mô hình ngôn ngữ PhoBERT kết hợp học máy SVM trong phân loại tự động hồ sơ khoa học tiếng Việt",
         "Bài báo đề xuất kiến trúc tiền xử lý tách từ tiếng Việt bằng thư viện PyVi kết hợp fine-tuning mô hình học sâu vinai/phobert-base và thuật toán Support Vector Machine để phân loại chính xác lĩnh vực chuyên ngành cho các công trình nghiên cứu khoa học tại trường đại học.",
         "trí tuệ nhân tạo, học sâu, xử lý ngôn ngữ tự nhiên, phobert, phân loại văn bản, học máy",
         "TS. Phạm Văn Khoa, PGS.TS. Trần Quốc Bảo",
         "Khoa Công nghệ Thông tin", 2026, "Đợt 1 - Năm học 2025-2026",
         4, 1, 1, 1, "Đã duyệt", "MinhChung_HS20260001_PhoBERT_Q1.pdf", 1),

        # 2. Hồ sơ Điện - Điện tử chuẩn (Chờ xét duyệt)
        ("HS-2026-0002",
         "Thiết kế bộ điều khiển thích nghi bền vững cho hệ thống nghịch lưu ba pha nối lưới điện thông minh tích hợp điện mặt trời",
         "Nghiên cứu trình bày thuật toán điều khiển tự động trên vi điều khiển DSP TMS320F28335 nhằm ổn định điện áp, giảm méo hài tổng THD và tối ưu công suất phát của biến tần năng lượng mặt trời trong lưới điện thông minh.",
         "lưới điện thông minh, năng lượng mặt trời, điều khiển tự động, điện tử công suất, vi điều khiển",
         "TS. Vũ Đình Trọng, PGS.TS. Lê Hoàng Việt",
         "Khoa Điện - Điện tử", 2026, "Đợt 1 - Năm học 2025-2026",
         6, 2, 1, 3, "Chờ xét duyệt", "MinhChung_HS20260002_SmartGrid_IEEE.pdf", 0),

        # 3. Hồ sơ BẤT THƯỜNG 1: Khai sai lĩnh vực chuyên ngành (Nội dung Hóa-Sinh Nano nhưng khai báo nhầm vào Kinh tế Quản lý!)
        ("HS-2026-0003",
         "Tổng hợp vật liệu nano ZnO pha tạp bạc ứng dụng xúc tác quang hóa phân hủy chất ô nhiễm hữu cơ và kháng khuẩn trong xử lý nước thải",
         "Công trình thực hiện tổng hợp hạt vật liệu nano bằng phương pháp thủy nhiệt, phân tích cấu trúc tinh thể XRD, SEM và đánh giá hiệu suất xử lý nước thải dệt nhuộm nhiễm kim loại nặng cùng hoạt tính kháng khuẩn E.coli.",
         "vật liệu nano, xúc tác quang hóa, xử lý nước thải, kháng khuẩn, ô nhiễm môi trường",
         "PGS.TS. Ngô Thanh Hà, TS. Bùi Thị Lan",
         "Khoa Kỹ thuật Hóa học & Môi trường", 2026, "Đợt 1 - Năm học 2025-2026",
         7, 4, 1, 6, "Chờ xét duyệt", "MinhChung_HS20260003_Nano_Water.pdf", 0),

        # 4. Hồ sơ Kinh tế chuẩn (Đã duyệt)
        ("HS-2026-0004",
         "Tác động của quản trị chuỗi cung ứng xanh và chuyển đổi số đến hiệu quả tài chính của các doanh nghiệp niêm yết trên thị trường chứng khoán",
         "Sử dụng mô hình cấu trúc tuyến tính SEM và dữ liệu bảng từ báo cáo tài chính kiểm toán của 250 doanh nghiệp nhằm đánh giá mối quan hệ giữa chuỗi cung ứng bền vững, quản trị rủi ro tài chính ngân hàng và tăng trưởng kinh tế.",
         "chuỗi cung ứng, tài chính ngân hàng, chứng khoán, kinh tế vĩ mô, kế toán kiểm toán",
         "ThS. Đặng Thị Mai Hương, TS. Lê Thu Trang",
         "Khoa Kinh tế & Quản trị Kinh doanh", 2026, "Đợt 1 - Năm học 2025-2026",
         5, 4, 2, 5, "Đã duyệt", "MinhChung_HS20260004_GreenSupplyChain.pdf", 1),

        # 5. Hồ sơ BẤT THƯỜNG 2: Đăng trên Tạp chí Săn mồi / Không uy tín (Predatory Journal) + Lệch đồ thị Tác giả
        ("HS-2026-0005",
         "Ứng dụng mạng nơ ron học sâu và thị giác máy tính trong nhận dạng khuôn mặt điểm danh lớp học thông minh",
         "Hệ thống tích hợp camera nhận diện khuôn mặt thời gian thực bằng mạng tích chập sâu CNN và cơ sở dữ liệu đám mây hỗ trợ quản lý điểm danh sinh viên tự động.",
         "thị giác máy tính, học sâu, mạng nơ ron, trí tuệ nhân tạo, công nghệ thông tin",
         "ThS. Đặng Thị Mai Hương",
         "Khoa Kinh tế & Quản trị Kinh doanh", 2026, "Đợt 1 - Năm học 2025-2026",
         5, 1, 2, 9, "Đang xét duyệt", "MinhChung_HS20260005_FaceRec_Warning.pdf", 0),

        # 6. Hồ sơ Cơ khí chuẩn (Đang xét duyệt)
        ("HS-2026-0006",
         "Mô phỏng phần tử hữu hạn FEM và tối ưu hóa chế độ cắt khi gia công phay CNC hợp kim titan dùng trong chế tạo máy",
         "Nghiên cứu phân tích trường nhiệt độ, lực cắt cơ học, biến dạng dẻo và mài mòn dụng cụ cắt khi gia công tốc độ cao hợp kim Ti-6Al-4V trên trung tâm gia công CNC 5 trục.",
         "gia công cnc, chế tạo máy, phần tử hữu hạn, ma sát mài mòn, cơ khí động lực",
         "TS. Hoàng Gia Huy, TS. Nguyễn Đức Mạnh",
         "Khoa Cơ khí Động lực", 2026, "Đợt 1 - Năm học 2025-2026",
         8, 3, 1, 4, "Đang xét duyệt", "MinhChung_HS20260006_CNC_FEM.pdf", 1),

        # 7. Hồ sơ Đề tài NCKH Liên ngành (Chờ xét duyệt)
        ("HS-2026-0007",
         "Đề tài nghiên cứu xây dựng và triển khai hệ thống cảm biến đeo tay IoT phân tích tín hiệu sinh học theo dõi nhịp tim cho người cao tuổi",
         "Đề tài nghiên cứu khoa học cấp Trường thực hiện khảo sát thực tiễn, xây dựng thuyết minh, thiết kế mạch thu nhận tín hiệu điện tâm đồ ECG trên vi điều khiển công suất thấp và tổ chức nghiệm thu sản phẩm thiết bị cảnh báo sớm rối loạn nhịp tim.",
         "mạng cảm biến không dây, hệ thống nhúng, y sinh học, xử lý tín hiệu, đề tài nghiên cứu",
         "TS. Vũ Đình Trọng, TS. Phạm Văn Khoa",
         "Khoa Điện - Điện tử", 2026, "Đợt 1 - Năm học 2025-2026",
         6, 2, 4, 8, "Chờ xét duyệt", "MinhChung_HS20260007_IoT_Bio.pdf", 0),

        # 8. Hồ sơ BẤT THƯỜNG 3: Nghi vấn trùng lặp nội dung cao với HS-2026-0001
        ("HS-2026-0008",
         "Nghiên cứu ứng dụng mô hình ngôn ngữ tiền huấn luyện PhoBERT và SVM phân loại tự động văn bản hồ sơ khoa học tiếng Việt",
         "Đề tài trình bày quy trình tiền xử lý tách từ tiếng Việt bằng PyVi kết hợp fine-tuning mô hình học sâu vinai/phobert-base và thuật toán học máy Support Vector Machine nhằm tự động phân loại lĩnh vực chuyên ngành hồ sơ nghiên cứu khoa học.",
         "trí tuệ nhân tạo, học sâu, xử lý ngôn ngữ tự nhiên, phobert, phân loại văn bản",
         "TS. Phạm Văn Khoa",
         "Khoa Công nghệ Thông tin", 2025, "Đợt 2 - Năm học 2024-2025",
         4, 1, 5, 2, "Từ chối", "MinhChung_HS20260008_Duplicate.pdf", 1),

        # 9. Hồ sơ Giáo dục & XHNV chuẩn (Đã duyệt)
        ("HS-2026-0009",
         "Đổi mới phương pháp giảng dạy đại học theo tiếp cận CDIO và chuẩn đầu ra AUN-QA nhằm phát triển năng lực tự học cho sinh viên",
         "Nghiên cứu khảo sát thực trạng và đề xuất quy trình thiết kế chương trình đào tạo kết hợp học tập qua dự án giúp nâng cao kỹ năng tư duy phản biện và năng lực nghề nghiệp.",
         "phương pháp giảng dạy, giáo dục đại học, chương trình đào tạo, khoa học giáo dục, năng lực người học",
         "TS. Đỗ Văn Kiên, TS. Nguyễn Thị Hồng",
         "Khoa Sư phạm & Khoa học Xã hội Nhân văn", 2026, "Đợt 1 - Năm học 2025-2026",
         4, 6, 3, 7, "Đã duyệt", "MinhChung_HS20260009_GiaoDuc_CDIO.pdf", 1),

        # 10. Hồ sơ Sáng chế & Giải pháp hữu ích Hóa - Môi trường (Chờ xét duyệt)
        ("HS-2026-0010",
         "Giải pháp hữu ích đề xuất quy trình độc quyền sản xuất chế phẩm sinh học giàu enzyme cellulase và than sinh học biến tính xử lý kim loại nặng",
         "Giải pháp hữu ích mô tả chi tiết các bước quy trình kỹ thuật lên men chủng xạ khuẩn bản địa kết hợp vật liệu hấp phụ sinh học loại bỏ chì Pb(II) và cadimi Cd(II); hồ sơ đăng ký bảo hộ sở hữu trí tuệ xác lập tính mới, tính sáng tạo và khả năng áp dụng công nghiệp.",
         "giải pháp hữu ích, sở hữu trí tuệ, công nghệ sinh học, enzyme sinh học, xử lý nước thải",
         "PGS.TS. Ngô Thanh Hà",
         "Khoa Kỹ thuật Hóa học & Môi trường", 2026, "Đợt 1 - Năm học 2025-2026",
         7, 5, 6, 6, "Chờ xét duyệt", "MinhChung_HS20260010_SangChe_Bio.pdf", 0),
    ]

    for item in sample_hoso:
        (ma_hs, tieu_de, tom_tat, tu_khoa, tac_gia, don_vi, nam, ky,
         ma_nguoi_nop, ma_lv, ma_loai, ma_tc, trang_thai, file_dk, da_xac_nhan) = item

        pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, tac_gia, don_vi)

        cur.execute("""
            INSERT INTO HoSoKHCN (
                maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
                namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
                trangThai, fileDinhKem, vanBanDaTachTu, maLinhVucAI, doTinCayAI,
                maLinhVucSVM, doTinCaySVM, nhomCongTrinhAI, doTinCayLoaiAI,
                nhomCongTrinhSVM, doTinCayLoaiSVM, daXacNhanAI
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ma_hs, tieu_de, tom_tat, tu_khoa, tac_gia, don_vi,
            nam, ky, ma_nguoi_nop, ma_lv, ma_loai, ma_tc,
            trang_thai, file_dk, pred["segmented_text"],
            pred["phobert_label_id"], pred["phobert_confidence"],
            pred["svm_label_id"], pred["svm_confidence"],
            pred["work_type_phobert"], pred["work_type_phobert_conf"],
            pred["work_type_svm"], pred["work_type_svm_conf"], da_xac_nhan
        ))

    lich_su = [
        ("HS-2026-0001", 4, "Khởi tạo", "Chờ xét duyệt", "Tác giả nộp hồ sơ mới lên hệ thống kèm tệp minh chứng toàn văn."),
        ("HS-2026-0001", 2, "Chờ xét duyệt", "Đang xét duyệt", "Cán bộ tiếp nhận thẩm định; nhãn gợi ý tự động khớp 100% với chuyên ngành CNTT."),
        ("HS-2026-0001", 2, "Đang xét duyệt", "Đã duyệt", "Hồ sơ hợp lệ, công bố trên tạp chí quốc tế uy tín, đủ điều kiện công nhận giờ chuẩn NCKH."),
        ("HS-2026-0004", 5, "Khởi tạo", "Chờ xét duyệt", "Tác giả nộp hồ sơ bài báo khoa học quốc tế."),
        ("HS-2026-0004", 3, "Chờ xét duyệt", "Đã duyệt", "Đã đối chiếu mã chuẩn ấn phẩm và xác nhận lĩnh vực Kinh tế, Quản trị kinh doanh và Tài chính."),
        ("HS-2026-0005", 5, "Khởi tạo", "Chờ xét duyệt", "Nộp hồ sơ bài báo nghiên cứu nhận diện khuôn mặt."),
        ("HS-2026-0005", 2, "Chờ xét duyệt", "Đang xét duyệt", "Hệ thống cảnh báo tạp chí thuộc diện ngừng chỉ mục (9999-001X) và khác biệt hướng nghiên cứu trước đây của tác giả. Đang yêu cầu giải trình."),
        ("HS-2026-0006", 8, "Khởi tạo", "Chờ xét duyệt", "Khai báo bài báo quốc tế chuyên ngành Cơ khí chế tạo máy."),
        ("HS-2026-0006", 3, "Chờ xét duyệt", "Đang xét duyệt", "Đã xác nhận nhãn lĩnh vực Kỹ thuật Cơ khí. Đang chờ Hội đồng khoa xác minh minh chứng."),
        ("HS-2026-0008", 4, "Khởi tạo", "Chờ xét duyệt", "Khai báo đề tài cấp Trường."),
        ("HS-2026-0008", 2, "Chờ xét duyệt", "Từ chối", "Từ chối phê duyệt: Phát hiện nội dung tóm tắt trùng lặp mức cao với hồ sơ HS-2026-0001."),
        ("HS-2026-0009", 4, "Khởi tạo", "Chờ xét duyệt", "Khai báo bài báo Tạp chí Khoa học Giáo dục Việt Nam."),
        ("HS-2026-0009", 3, "Chờ xét duyệt", "Đã duyệt", "Hồ sơ đầy đủ minh chứng, thuộc danh mục HĐGSNN 0.75 điểm, phê duyệt."),
    ]
    cur.executemany("""
        INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
        VALUES (?, ?, ?, ?, ?)
    """, lich_su)

    cur.execute("SELECT maHoSo, maLinhVucAI, maLinhVuc, doTinCayAI, vanBanDaTachTu FROM HoSoKHCN WHERE maHoSo IN ('HS-2026-0001', 'HS-2026-0004', 'HS-2026-0006', 'HS-2026-0009')")
    for r in cur.fetchall():
        cur.execute("""
            INSERT INTO PhanHoiAI (maHoSo, maCanBo, maLinhVucAI, maLinhVucThucTe, doTinCayLucDo, ketQuaKhop, vanBanHuanLuyen, ghiChuPhanHoi)
            VALUES (?, 2, ?, ?, ?, 1, ?, 'Cán bộ xác nhận nhãn lĩnh vực dự đoán tự động là chính xác.')
        """, (r["maHoSo"], r["maLinhVucAI"], r["maLinhVuc"], r["doTinCayAI"], r["vanBanDaTachTu"]))

    conn.commit()
    ai_engine.build_networkx_graph_and_detect_anomalies(conn)
