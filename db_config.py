# -*- coding: utf-8 -*-
"""
Cấu hình kết nối Cơ sở dữ liệu Microsoft SQL Server (SSMS)
Đề tài: Hệ thống Hỗ trợ Quản lý Hồ sơ KH&CN Tích hợp Phân loại Tự động (CNTT-KLCN001)
"""

import os
import socket
import pyodbc

# Chế độ sử dụng 100% Microsoft SQL Server (SSMS)
USE_SQL_SERVER = True

# Tên máy chủ (Hostname) hiện tại
HOSTNAME = socket.gethostname()

# Bạn có thể gán cố định tên Server trong SSMS tại đây (hoặc để None để tự động nhận diện)
SQL_SERVER_NAME = os.environ.get("KHCN_SQL_SERVER", None)

# Tên Database hiển thị trong SSMS Object Explorer
SQL_DATABASE_NAME = os.environ.get("KHCN_SQL_DB", "KHCN_AI_System")

# Chế độ xác thực SSMS:
# True  -> Windows Authentication (Trusted_Connection=yes)
# False -> SQL Server Authentication (UID / PWD)
USE_WINDOWS_AUTH = True
SQL_USER = os.environ.get("KHCN_SQL_USER", "sa")
SQL_PASSWORD = os.environ.get("KHCN_SQL_PASSWORD", "123456")

# Danh sách các tên Server phổ biến trên máy Windows/SSMS để tự động kết nối
CANDIDATE_SERVERS = [
    HOSTNAME,
    rf"{HOSTNAME}\SQLEXPRESS",
    r"localhost",
    r".\SQLEXPRESS",
    r"localhost\SQLEXPRESS",
    r"(local)",
    r"(localdb)\MSSQLLocalDB",
]

_CACHED_SERVER_AND_DRIVER = None


def get_preferred_odbc_driver() -> str:
    """Lấy Driver ODBC tốt nhất đã cài trên máy cho SQL Server."""
    drivers = pyodbc.drivers()
    for preferred in [
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 18 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]:
        if preferred in drivers:
            return preferred
    return "ODBC Driver 17 for SQL Server"


def build_conn_str(server: str, driver: str, database: str = "master") -> str:
    """Tạo chuỗi kết nối ODBC chuẩn cho SQL Server (SSMS)."""
    trust_cert = "TrustServerCertificate=yes;" if "18" in driver or "17" in driver else ""
    if USE_WINDOWS_AUTH:
        return (
            f"DRIVER={{{driver}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Trusted_Connection=yes;"
            f"{trust_cert}"
        )
    return (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={SQL_USER};"
        f"PWD={SQL_PASSWORD};"
        f"{trust_cert}"
    )


def detect_active_sql_server():
    """Tự động dò tìm SQL Server Instance đang chạy trên máy tính."""
    global _CACHED_SERVER_AND_DRIVER
    if _CACHED_SERVER_AND_DRIVER is not None:
        return _CACHED_SERVER_AND_DRIVER

    driver = get_preferred_odbc_driver()
    candidates = [SQL_SERVER_NAME] if SQL_SERVER_NAME else CANDIDATE_SERVERS

    last_err = None
    for srv in candidates:
        if not srv:
            continue
        try:
            conn_str = build_conn_str(srv, driver, database="master")
            conn = pyodbc.connect(conn_str, timeout=2)
            conn.close()
            _CACHED_SERVER_AND_DRIVER = (srv, driver)
            return _CACHED_SERVER_AND_DRIVER
        except Exception as exc:
            last_err = exc

    raise RuntimeError(
        f"Không thể kết nối tới dịch vụ SQL Server (SSMS). Chi tiết lỗi: {last_err}"
    )


def get_sql_server_connection_string(database: str = None) -> str:
    """Trả về chuỗi kết nối tới Database KHCN_AI_System trên SQL Server (SSMS)."""
    srv, driver = detect_active_sql_server()
    target_db = database or SQL_DATABASE_NAME
    return build_conn_str(srv, driver, database=target_db)
