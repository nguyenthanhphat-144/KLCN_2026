# -*- coding: utf-8 -*-
"""
Script nâng cấp CSDL theo Mô tả nghiệp vụ revision:
- Bổ sung trạng thái 'Yêu cầu bổ sung' vào bảng HoSoKHCN.
- Bổ sung các trường nghiệp vụ riêng cho Bài báo khoa học: doi, tapSo, trangBaiBao, chiMucTapChi.
- Bổ sung các trường nghiệp vụ riêng cho Đề tài NCKH: maDeTai, capDeTai, chuNhiemDeTai, thanhVienThamGia, thoiGianThucHien, kinhPhi, sanPhamDuKien, ketQuaNghiemThu.
- Cập nhật dữ liệu mẫu thực tế cho cả 2 nhóm.
"""
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pyodbc
import db_config

def migrate_sql_server():
    print("Connecting to SQL Server...")
    conn = pyodbc.connect(db_config.get_sql_server_connection_string())
    cur = conn.cursor()

    # 1. Update CHECK constraint
    cur.execute("""
        SELECT cc.name 
        FROM sys.check_constraints cc 
        JOIN sys.tables t ON cc.parent_object_id = t.object_id 
        WHERE t.name = 'HoSoKHCN' AND cc.definition LIKE '%trangThai%'
    """)
    ck_rows = cur.fetchall()
    for ck in ck_rows:
        ck_name = ck[0]
        print(f"Dropping constraint: {ck_name}")
        cur.execute(f"ALTER TABLE dbo.HoSoKHCN DROP CONSTRAINT [{ck_name}]")

    cur.execute("""
        ALTER TABLE dbo.HoSoKHCN ADD CONSTRAINT CK_HoSoKHCN_trangThai 
        CHECK (trangThai IN (N'Chờ xét duyệt', N'Đang xét duyệt', N'Đã duyệt', N'Từ chối', N'Yêu cầu bổ sung'))
    """)
    print("Added updated CK_HoSoKHCN_trangThai successfully.")

    # 2. Add new columns
    new_columns = [
        ("doi", "VARCHAR(100) NULL"),
        ("tapSo", "NVARCHAR(50) NULL"),
        ("trangBaiBao", "NVARCHAR(50) NULL"),
        ("chiMucTapChi", "NVARCHAR(100) NULL"),
        ("maDeTai", "VARCHAR(50) NULL"),
        ("capDeTai", "NVARCHAR(100) NULL"),
        ("chuNhiemDeTai", "NVARCHAR(150) NULL"),
        ("thanhVienThamGia", "NVARCHAR(500) NULL"),
        ("thoiGianThucHien", "NVARCHAR(100) NULL"),
        ("kinhPhi", "FLOAT DEFAULT 0.0"),
        ("sanPhamDuKien", "NVARCHAR(MAX) NULL"),
        ("ketQuaNghiemThu", "NVARCHAR(50) DEFAULT N'Đang thực hiện'")
    ]

    cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'HoSoKHCN'")
    existing_cols = {r[0].lower() for r in cur.fetchall()}

    for col_name, col_def in new_columns:
        if col_name.lower() not in existing_cols:
            print(f"Adding column: {col_name} {col_def}")
            cur.execute(f"ALTER TABLE dbo.HoSoKHCN ADD [{col_name}] {col_def}")
        else:
            print(f"Column {col_name} already exists.")

    conn.commit()

    # 3. Cập nhật dữ liệu mẫu thực tế cho các hồ sơ bài báo và đề tài
    print("Updating sample profiles with realistic article and project metadata...")

    # Cập nhật metadata cho nhóm Bài báo khoa học
    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET doi = '10.1016/j.engappai.2025.106821',
            tapSo = N'Tập 138, Số 2',
            trangBaiBao = '104-118',
            chiMucTapChi = N'Scopus Q1 / Web of Science (SCIE)'
        WHERE maHoSo = 'HS-2026-0001'
    """)

    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET doi = '10.15625/1813-9663/39/3/18201',
            tapSo = N'Tập 39, Số 3',
            trangBaiBao = '245-256',
            chiMucTapChi = N'Hội đồng GSNN (1.0 điểm)'
        WHERE maHoSo = 'HS-2026-0002'
    """)

    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET doi = '10.1109/TIE.2025.3241908',
            tapSo = N'Tập 72, Số 8',
            trangBaiBao = '8910-8921',
            chiMucTapChi = N'Scopus Q1 / Web of Science (SCIE)'
        WHERE maHoSo = 'HS-2026-0003'
    """)

    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET doi = '10.1016/j.ijmecsci.2025.107412',
            tapSo = N'Tập 245',
            trangBaiBao = '55-68',
            chiMucTapChi = N'Scopus Q1 / Web of Science'
        WHERE maHoSo = 'HS-2026-0004'
    """)

    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET doi = '10.1016/j.jclepro.2025.138902',
            tapSo = N'Tập 412',
            trangBaiBao = '210-224',
            chiMucTapChi = N'Scopus Q1 / Web of Science (SCIE)'
        WHERE maHoSo = 'HS-2026-0006'
    """)

    # Cập nhật metadata cho nhóm Đề tài NCKH
    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET maDeTai = 'DT-2025-DH01',
            capDeTai = N'Cấp Đại học / Bộ GD&ĐT',
            chuNhiemDeTai = N'Nguyễn Thành Tài',
            thanhVienThamGia = N'ThS. Lê Hoàng Long, KS. Đỗ Minh Triết',
            thoiGianThucHien = N'01/2025 - 12/2026',
            kinhPhi = 150000000,
            sanPhamDuKien = N'02 bài báo Scopus/WoS, 01 chương trình phần mềm ứng dụng AI',
            ketQuaNghiemThu = N'Xuất sắc'
        WHERE maHoSo = 'HS-2026-0007'
    """)

    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET maDeTai = 'DT-2025-CS08',
            capDeTai = N'Cấp Trường / Cơ sở',
            chuNhiemDeTai = N'Nguyễn Thành Phát',
            thanhVienThamGia = N'ThS. Phạm Hải Đăng, CN. Vũ Bích Ngọc',
            thoiGianThucHien = N'03/2025 - 03/2026',
            kinhPhi = 45000000,
            sanPhamDuKien = N'01 bài báo tạp chí HĐGSNN, 01 báo cáo tổng kết đề tài',
            ketQuaNghiemThu = N'Tốt'
        WHERE maHoSo = 'HS-2026-0008'
    """)

    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET maDeTai = 'DT-2026-TINH03',
            capDeTai = N'Cấp Tỉnh / Thành phố',
            chuNhiemDeTai = N'Lê Thanh Phong',
            thanhVienThamGia = N'TS. Nguyễn Thành Tài, ThS. Hoàng Văn Bình',
            thoiGianThucHien = N'06/2025 - 06/2027',
            kinhPhi = 350000000,
            sanPhamDuKien = N'01 mô hình thử nghiệm tại địa phương, 02 bài báo quốc tế',
            ketQuaNghiemThu = N'Đang thực hiện'
        WHERE maHoSo = 'HS-2026-0009'
    """)

    # Gán 1 hồ sơ vào trạng thái 'Yêu cầu bổ sung' để kiểm thử tính năng mới
    cur.execute("""
        UPDATE dbo.HoSoKHCN
        SET trangThai = N'Yêu cầu bổ sung'
        WHERE maHoSo = 'HS-2026-0005'
    """)

    conn.commit()
    print("Database updated with realistic article & project metadata successfully!")
    conn.close()

if __name__ == "__main__":
    migrate_sql_server()
