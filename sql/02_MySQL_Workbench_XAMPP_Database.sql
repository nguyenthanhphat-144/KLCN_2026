-- =============================================================================

CREATE DATABASE IF NOT EXISTS khcn_ai_system
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE khcn_ai_system;

DROP TABLE IF EXISTS PhanHoiAI;
DROP TABLE IF EXISTS LichSuXetDuyet;
DROP TABLE IF EXISTS HoSoKHCN;
DROP TABLE IF EXISTS TapChiHopLe;
DROP TABLE IF EXISTS LoaiCongTrinh;
DROP TABLE IF EXISTS LinhVuc;
DROP TABLE IF EXISTS ThucNghiemMoHinhAI;
DROP TABLE IF EXISTS NguoiDung;

-- 1. Bảng NguoiDung
CREATE TABLE NguoiDung (
    maNguoiDung INT AUTO_INCREMENT PRIMARY KEY,
    hoTen VARCHAR(150) NOT NULL,
    email VARCHAR(150) NOT NULL UNIQUE,
    matKhau VARCHAR(255) NOT NULL,
    vaiTro ENUM('Admin', 'CanBo', 'TacGia') NOT NULL DEFAULT 'TacGia',
    donViCongTac VARCHAR(200) NOT NULL,
    hocHamHocVi VARCHAR(50) DEFAULT 'ThS.',
    trangThai TINYINT DEFAULT 1,
    ngayTao DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 2. Bảng LinhVuc
CREATE TABLE LinhVuc (
    maLinhVuc INT AUTO_INCREMENT PRIMARY KEY,
    maCode VARCHAR(30) NOT NULL UNIQUE,
    tenLinhVuc VARCHAR(200) NOT NULL,
    moTa TEXT,
    tuKhoaDacTrung TEXT,
    trangThai TINYINT DEFAULT 1
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 3. Bảng LoaiCongTrinh
CREATE TABLE LoaiCongTrinh (
    maLoai INT AUTO_INCREMENT PRIMARY KEY,
    maCodeLoai VARCHAR(30) NOT NULL UNIQUE,
    tenLoai VARCHAR(250) NOT NULL,
    nhomCongTrinh VARCHAR(100) NOT NULL,
    gioChuanQuyDoi DOUBLE DEFAULT 0.0,
    trangThai TINYINT DEFAULT 1
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 4. Bảng TapChiHopLe
CREATE TABLE TapChiHopLe (
    maTapChi INT AUTO_INCREMENT PRIMARY KEY,
    tenTapChi VARCHAR(300) NOT NULL,
    chiSoISSN VARCHAR(30) NOT NULL,
    xepHang VARCHAR(100) NOT NULL,
    maLinhVuc INT NULL,
    uyTin TINYINT DEFAULT 1,
    trangThai TINYINT DEFAULT 1,
    CONSTRAINT fk_tapchi_linhvuc FOREIGN KEY (maLinhVuc) REFERENCES LinhVuc(maLinhVuc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 5. Bảng HoSoKHCN
CREATE TABLE HoSoKHCN (
    maHoSo VARCHAR(30) PRIMARY KEY,
    tieuDe VARCHAR(500) NOT NULL,
    tomTat TEXT NOT NULL,
    tuKhoa VARCHAR(300) NOT NULL,
    thongTinTacGia VARCHAR(300) NOT NULL,
    donViChuTri VARCHAR(200) NOT NULL,
    namCongBo INT DEFAULT 2026,
    kyBaoCao VARCHAR(100) DEFAULT 'Đợt 1 - Năm học 2025-2026',
    maNguoiNop INT NOT NULL,
    maLinhVuc INT NOT NULL,
    maLoai INT NOT NULL,
    maTapChi INT NULL,
    trangThai VARCHAR(50) DEFAULT 'Chờ xét duyệt',
    fileDinhKem VARCHAR(255) NULL,
    isDeleted TINYINT DEFAULT 0,
    vanBanDaTachTu TEXT NULL,
    maLinhVucAI INT NULL,
    doTinCayAI DOUBLE DEFAULT 0.0,
    maLinhVucSVM INT NULL,
    doTinCaySVM DOUBLE DEFAULT 0.0,
    moHinhChinh VARCHAR(100) DEFAULT 'PhoBERT (vinai/phobert-base)',
    daXacNhanAI TINYINT DEFAULT 0,
    diemBatThuong DOUBLE DEFAULT 0.0,
    coCanhBao TINYINT DEFAULT 0,
    chiTietCanhBao TEXT,
    doTuongDongMax DOUBLE DEFAULT 0.0,
    hoSoTuongDongNhat VARCHAR(30) DEFAULT '',
    ngayNop DATETIME DEFAULT CURRENT_TIMESTAMP,
    ngayCapNhat DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    CONSTRAINT fk_hoso_nguoidung FOREIGN KEY (maNguoiNop) REFERENCES NguoiDung(maNguoiDung),
    CONSTRAINT fk_hoso_linhvuc FOREIGN KEY (maLinhVuc) REFERENCES LinhVuc(maLinhVuc),
    CONSTRAINT fk_hoso_loai FOREIGN KEY (maLoai) REFERENCES LoaiCongTrinh(maLoai),
    CONSTRAINT fk_hoso_tapchi FOREIGN KEY (maTapChi) REFERENCES TapChiHopLe(maTapChi),
    CONSTRAINT fk_hoso_linhvuc_ai FOREIGN KEY (maLinhVucAI) REFERENCES LinhVuc(maLinhVuc),
    CONSTRAINT fk_hoso_linhvuc_svm FOREIGN KEY (maLinhVucSVM) REFERENCES LinhVuc(maLinhVuc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 6. Bảng LichSuXetDuyet
CREATE TABLE LichSuXetDuyet (
    maLichSu INT AUTO_INCREMENT PRIMARY KEY,
    maHoSo VARCHAR(30) NOT NULL,
    maCanBo INT NOT NULL,
    trangThaiCu VARCHAR(50),
    trangThaiMoi VARCHAR(50) NOT NULL,
    yKienDanhGia TEXT NOT NULL,
    thoiGian DATETIME DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_lichsu_hoso FOREIGN KEY (maHoSo) REFERENCES HoSoKHCN(maHoSo) ON DELETE CASCADE,
    CONSTRAINT fk_lichsu_canbo FOREIGN KEY (maCanBo) REFERENCES NguoiDung(maNguoiDung)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 7. Bảng PhanHoiAI
CREATE TABLE PhanHoiAI (
    maPhanHoi INT AUTO_INCREMENT PRIMARY KEY,
    maHoSo VARCHAR(30) NOT NULL,
    maCanBo INT NOT NULL,
    maLinhVucAI INT NOT NULL,
    maLinhVucThucTe INT NOT NULL,
    doTinCayLucDo DOUBLE NOT NULL,
    ketQuaKhop TINYINT NOT NULL,
    vanBanHuanLuyen TEXT NOT NULL,
    ghiChuPhanHoi VARCHAR(500),
    thoiGian DATETIME DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_phanhoi_hoso FOREIGN KEY (maHoSo) REFERENCES HoSoKHCN(maHoSo) ON DELETE CASCADE,
    CONSTRAINT fk_phanhoi_canbo FOREIGN KEY (maCanBo) REFERENCES NguoiDung(maNguoiDung),
    CONSTRAINT fk_phanhoi_lv_ai FOREIGN KEY (maLinhVucAI) REFERENCES LinhVuc(maLinhVuc),
    CONSTRAINT fk_phanhoi_lv_thucte FOREIGN KEY (maLinhVucThucTe) REFERENCES LinhVuc(maLinhVuc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 8. Bảng ThucNghiemMoHinhAI
CREATE TABLE ThucNghiemMoHinhAI (
    maThucNghiem INT AUTO_INCREMENT PRIMARY KEY,
    tenMoHinh VARCHAR(200) NOT NULL,
    precisionScore DOUBLE NOT NULL,
    recallScore DOUBLE NOT NULL,
    f1Score DOUBLE NOT NULL,
    accuracyScore DOUBLE NOT NULL,
    thoiGianPhanHoiMs DOUBLE NOT NULL,
    soMauHuanLuyen INT NOT NULL,
    soMauKiemThu INT NOT NULL,
    moTaChiTiet TEXT,
    ngayDanhGia DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
