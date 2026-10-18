-- =============================================================================

PRAGMA foreign_keys = ON;

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

CREATE TABLE IF NOT EXISTS LinhVuc (
    maLinhVuc INTEGER PRIMARY KEY AUTOINCREMENT,
    maCode TEXT NOT NULL UNIQUE,
    tenLinhVuc TEXT NOT NULL,
    moTa TEXT,
    tuKhoaDacTrung TEXT,
    trangThai INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS LoaiCongTrinh (
    maLoai INTEGER PRIMARY KEY AUTOINCREMENT,
    maCodeLoai TEXT NOT NULL UNIQUE,
    tenLoai TEXT NOT NULL,
    nhomCongTrinh TEXT NOT NULL,
    gioChuanQuyDoi REAL DEFAULT 0.0,
    trangThai INTEGER DEFAULT 1
);

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
    trangThai TEXT DEFAULT 'Chờ xét duyệt' CHECK(trangThai IN ('Chờ xét duyệt', 'Đang xét duyệt', 'Đã duyệt', 'Từ chối')),
    fileDinhKem TEXT,
    isDeleted INTEGER DEFAULT 0,
    vanBanDaTachTu TEXT,
    maLinhVucAI INTEGER,
    doTinCayAI REAL DEFAULT 0.0,
    maLinhVucSVM INTEGER,
    doTinCaySVM REAL DEFAULT 0.0,
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
