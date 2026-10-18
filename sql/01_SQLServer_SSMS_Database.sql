-- =============================================================================
-- KỊCH BẢN TẠO CƠ SỞ DỮ LIỆU & TOÀN BỘ DỮ LIỆU MẪU TRÊN MICROSOFT SQL SERVER (SSMS)
-- =============================================================================

USE master;
GO

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'KHCN_AI_System')
BEGIN
    CREATE DATABASE KHCN_AI_System;
END
GO

USE KHCN_AI_System;
GO

-- Xóa các bảng cũ nếu đã tồn tại (theo đúng thứ tự ràng buộc khóa ngoại)
IF OBJECT_ID('dbo.PhanHoiAI', 'U') IS NOT NULL DROP TABLE dbo.PhanHoiAI;
IF OBJECT_ID('dbo.LichSuXetDuyet', 'U') IS NOT NULL DROP TABLE dbo.LichSuXetDuyet;
IF OBJECT_ID('dbo.HoSoKHCN', 'U') IS NOT NULL DROP TABLE dbo.HoSoKHCN;
IF OBJECT_ID('dbo.TapChiHopLe', 'U') IS NOT NULL DROP TABLE dbo.TapChiHopLe;
IF OBJECT_ID('dbo.LoaiCongTrinh', 'U') IS NOT NULL DROP TABLE dbo.LoaiCongTrinh;
IF OBJECT_ID('dbo.LinhVuc', 'U') IS NOT NULL DROP TABLE dbo.LinhVuc;
IF OBJECT_ID('dbo.ThucNghiemMoHinhAI', 'U') IS NOT NULL DROP TABLE dbo.ThucNghiemMoHinhAI;
IF OBJECT_ID('dbo.NguoiDung', 'U') IS NOT NULL DROP TABLE dbo.NguoiDung;
GO

-- =============================================================================
-- PHẦN 1: KHỞI TẠO CẤU TRÚC 8 BẢNG DỮ LIỆU VẬT LÝ & RÀNG BUỘC KHÓA NGOẠI
-- =============================================================================

-- 1. Bảng NguoiDung (Lưu trữ tài khoản & phân quyền RBAC)
CREATE TABLE dbo.NguoiDung (
    maNguoiDung INT IDENTITY(1,1) PRIMARY KEY,
    hoTen NVARCHAR(150) NOT NULL,
    email VARCHAR(150) NOT NULL UNIQUE,
    matKhau VARCHAR(255) NOT NULL,
    vaiTro VARCHAR(20) NOT NULL CHECK (vaiTro IN ('Admin', 'CanBo', 'TacGia')),
    donViCongTac NVARCHAR(200) NOT NULL,
    hocHamHocVi NVARCHAR(50) DEFAULT N'ThS.',
    trangThai INT DEFAULT 1,
    ngayTao DATETIME DEFAULT CURRENT_TIMESTAMP
);
GO

-- 2. Bảng danh mục LinhVuc (6 lĩnh vực chuyên ngành chuẩn)
CREATE TABLE dbo.LinhVuc (
    maLinhVuc INT IDENTITY(1,1) PRIMARY KEY,
    maCode VARCHAR(30) NOT NULL UNIQUE,
    tenLinhVuc NVARCHAR(200) NOT NULL,
    moTa NVARCHAR(MAX),
    tuKhoaDacTrung NVARCHAR(MAX),
    trangThai INT DEFAULT 1
);
GO

-- 3. Bảng danh mục LoaiCongTrinh
CREATE TABLE dbo.LoaiCongTrinh (
    maLoai INT IDENTITY(1,1) PRIMARY KEY,
    maCodeLoai VARCHAR(30) NOT NULL UNIQUE,
    tenLoai NVARCHAR(250) NOT NULL,
    nhomCongTrinh NVARCHAR(100) NOT NULL,
    gioChuanQuyDoi FLOAT DEFAULT 0.0,
    trangThai INT DEFAULT 1
);
GO

-- 4. Bảng danh mục TapChiHopLe (Kiểm tra chéo tạp chí uy tín & cảnh báo Predatory)
CREATE TABLE dbo.TapChiHopLe (
    maTapChi INT IDENTITY(1,1) PRIMARY KEY,
    tenTapChi NVARCHAR(300) NOT NULL,
    chiSoISSN VARCHAR(30) NOT NULL,
    xepHang NVARCHAR(100) NOT NULL,
    maLinhVuc INT NULL,
    uyTin INT DEFAULT 1, -- 1: Hợp lệ, 0: Cảnh báo không uy tín / Predatory
    trangThai INT DEFAULT 1,
    CONSTRAINT FK_TapChi_LinhVuc FOREIGN KEY (maLinhVuc) REFERENCES dbo.LinhVuc(maLinhVuc)
);
GO

-- 5. Bảng trung tâm HoSoKHCN
CREATE TABLE dbo.HoSoKHCN (
    maHoSo VARCHAR(30) PRIMARY KEY,
    tieuDe NVARCHAR(500) NOT NULL,
    tomTat NVARCHAR(MAX) NOT NULL,
    tuKhoa NVARCHAR(300) NOT NULL,
    thongTinTacGia NVARCHAR(300) NOT NULL,
    donViChuTri NVARCHAR(200) NOT NULL,
    namCongBo INT DEFAULT 2026,
    kyBaoCao NVARCHAR(100) DEFAULT N'Đợt 1 - Năm học 2025-2026',
    maNguoiNop INT NOT NULL,
    maLinhVuc INT NOT NULL,
    maLoai INT NOT NULL,
    maTapChi INT NULL,
    trangThai NVARCHAR(50) DEFAULT N'Chờ xét duyệt' CHECK (trangThai IN (N'Chờ xét duyệt', N'Đang xét duyệt', N'Đã duyệt', N'Từ chối')),
    fileDinhKem NVARCHAR(255) NULL,
    isDeleted INT DEFAULT 0,
    -- Các trường tích hợp AI Phân loại tự động (Chức năng 5) & Phát hiện bất thường
    vanBanDaTachTu NVARCHAR(MAX) NULL,
    maLinhVucAI INT NULL,
    doTinCayAI FLOAT DEFAULT 0.0,
    maLinhVucSVM INT NULL,
    doTinCaySVM FLOAT DEFAULT 0.0,
    moHinhChinh NVARCHAR(100) DEFAULT N'PhoBERT (vinai/phobert-base)',
    daXacNhanAI INT DEFAULT 0,
    diemBatThuong FLOAT DEFAULT 0.0,
    coCanhBao INT DEFAULT 0,
    chiTietCanhBao NVARCHAR(MAX) DEFAULT N'[]',
    doTuongDongMax FLOAT DEFAULT 0.0,
    hoSoTuongDongNhat VARCHAR(30) DEFAULT '',
    ngayNop DATETIME DEFAULT CURRENT_TIMESTAMP,
    ngayCapNhat DATETIME DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT FK_HoSo_NguoiDung FOREIGN KEY (maNguoiNop) REFERENCES dbo.NguoiDung(maNguoiDung),
    CONSTRAINT FK_HoSo_LinhVuc FOREIGN KEY (maLinhVuc) REFERENCES dbo.LinhVuc(maLinhVuc),
    CONSTRAINT FK_HoSo_LoaiCongTrinh FOREIGN KEY (maLoai) REFERENCES dbo.LoaiCongTrinh(maLoai),
    CONSTRAINT FK_HoSo_TapChi FOREIGN KEY (maTapChi) REFERENCES dbo.TapChiHopLe(maTapChi),
    CONSTRAINT FK_HoSo_LinhVucAI FOREIGN KEY (maLinhVucAI) REFERENCES dbo.LinhVuc(maLinhVuc),
    CONSTRAINT FK_HoSo_LinhVucSVM FOREIGN KEY (maLinhVucSVM) REFERENCES dbo.LinhVuc(maLinhVuc)
);
GO

-- 6. Bảng LichSuXetDuyet (Nhật ký thẩm định & chuyển đổi trạng thái)
CREATE TABLE dbo.LichSuXetDuyet (
    maLichSu INT IDENTITY(1,1) PRIMARY KEY,
    maHoSo VARCHAR(30) NOT NULL,
    maCanBo INT NOT NULL,
    trangThaiCu NVARCHAR(50) NULL,
    trangThaiMoi NVARCHAR(50) NOT NULL,
    yKienDanhGia NVARCHAR(MAX) NOT NULL,
    thoiGian DATETIME DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT FK_LichSu_HoSo FOREIGN KEY (maHoSo) REFERENCES dbo.HoSoKHCN(maHoSo) ON DELETE CASCADE,
    CONSTRAINT FK_LichSu_CanBo FOREIGN KEY (maCanBo) REFERENCES dbo.NguoiDung(maNguoiDung)
);
GO

-- 7. Bảng PhanHoiAI (Feedback Loop phục vụ Fine-tuning - Tuần 6 Mục 2d)
CREATE TABLE dbo.PhanHoiAI (
    maPhanHoi INT IDENTITY(1,1) PRIMARY KEY,
    maHoSo VARCHAR(30) NOT NULL,
    maCanBo INT NOT NULL,
    maLinhVucAI INT NOT NULL,
    maLinhVucThucTe INT NOT NULL,
    doTinCayLucDo FLOAT NOT NULL,
    ketQuaKhop INT NOT NULL,
    vanBanHuanLuyen NVARCHAR(MAX) NOT NULL,
    ghiChuPhanHoi NVARCHAR(500) NULL,
    thoiGian DATETIME DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT FK_PhanHoi_HoSo FOREIGN KEY (maHoSo) REFERENCES dbo.HoSoKHCN(maHoSo) ON DELETE CASCADE,
    CONSTRAINT FK_PhanHoi_CanBo FOREIGN KEY (maCanBo) REFERENCES dbo.NguoiDung(maNguoiDung),
    CONSTRAINT FK_PhanHoi_LinhVucAI FOREIGN KEY (maLinhVucAI) REFERENCES dbo.LinhVuc(maLinhVuc),
    CONSTRAINT FK_PhanHoi_LinhVucThucTe FOREIGN KEY (maLinhVucThucTe) REFERENCES dbo.LinhVuc(maLinhVuc)
);
GO

-- 8. Bảng ThucNghiemMoHinhAI (Đánh giá Precision, Recall, F1-score)
CREATE TABLE dbo.ThucNghiemMoHinhAI (
    maThucNghiem INT IDENTITY(1,1) PRIMARY KEY,
    tenMoHinh NVARCHAR(200) NOT NULL,
    precisionScore FLOAT NOT NULL,
    recallScore FLOAT NOT NULL,
    f1Score FLOAT NOT NULL,
    accuracyScore FLOAT NOT NULL,
    thoiGianPhanHoiMs FLOAT NOT NULL,
    soMauHuanLuyen INT NOT NULL,
    soMauKiemThu INT NOT NULL,
    moTaChiTiet NVARCHAR(MAX),
    ngayDanhGia DATETIME DEFAULT CURRENT_TIMESTAMP
);
GO

-- =============================================================================
-- PHẦN 2: NẠP DỮ LIỆU MẪU CHUẨN HÓA (SEED DATA CHO 8 BẢNG)
-- =============================================================================

-- 2.1. Dữ liệu Bảng NguoiDung
INSERT INTO dbo.NguoiDung (hoTen, email, matKhau, vaiTro, donViCongTac, hocHamHocVi) VALUES
(N'Nguyễn Thành Phát', 'admin@uni.edu.vn', 'Admin@123', 'Admin', N'Phòng Quản lý Khoa học & HTQT', N'PGS.TS.'),
(N'Lê Thanh Phong', 'canbo@uni.edu.vn', 'CanBo@123', 'CanBo', N'Hội đồng KH&CN Trường', N'TS.'),
(N'PGS.TS. Lê Hoàng Việt', 'vietlh@uni.edu.vn', 'CanBo@123', 'CanBo', N'Phòng Quản lý Khoa học & HTQT', N'PGS.TS.'),
(N'Nguyễn Thành Tài', 'tacgia@uni.edu.vn', 'TacGia@123', 'TacGia', N'Khoa Công nghệ Thông tin', N'TS.'),
(N'ThS. Đặng Thị Mai Hương', 'huongdtm@uni.edu.vn', 'TacGia@123', 'TacGia', N'Khoa Kinh tế & Quản trị Kinh doanh', N'ThS.'),
(N'TS. Vũ Đình Trọng', 'trongvd@uni.edu.vn', 'TacGia@123', 'TacGia', N'Khoa Điện - Điện tử', N'TS.'),
(N'PGS.TS. Ngô Thanh Hà', 'hant@uni.edu.vn', 'TacGia@123', 'TacGia', N'Khoa Kỹ thuật Hóa học & Môi trường', N'PGS.TS.'),
(N'TS. Hoàng Gia Huy', 'huyhg@uni.edu.vn', 'TacGia@123', 'TacGia', N'Khoa Cơ khí Động lực', N'TS.');
GO

-- 2.2. Dữ liệu Bảng LinhVuc
INSERT INTO dbo.LinhVuc (maCode, tenLinhVuc, moTa, tuKhoaDacTrung) VALUES
('LV-CNTT', N'Công nghệ thông tin và Truyền thông', N'Trí tuệ nhân tạo, học máy, học sâu, xử lý ngôn ngữ tự nhiên, thị giác máy tính, an toàn thông tin, mạng máy tính, kỹ thuật phần mềm, dữ liệu lớn.', N'trí_tuệ_nhân_tạo học_máy học_sâu phobert xử_lý_ngôn_ngữ_tự_nhiên thị_giác_máy_tính mạng_nơ_ron dữ_liệu_lớn an_toàn_thông_tin chuỗi_khối phần_mềm'),
('LV-DDT', N'Kỹ thuật Điện - Điện tử và Tự động hóa', N'Hệ thống nhúng, vi mạch bán dẫn, điều khiển tự động, mạng cảm biến không dây IoT, truyền thông vô tuyến 5G/6G, lưới điện thông minh, điện tử công suất.', N'hệ_thống_nhúng vi_mạch bán_dẫn điều_khiển_tự_động cảm_biến_không_dây lưới_điện_thông_minh biến_tần_công_suất ăng_ten vi_điều_khiển plc_scada'),
('LV-CKD', N'Kỹ thuật Cơ khí, Chế tạo máy và Giao thông', N'Thiết kế chế tạo máy, cơ điện tử, động cơ đốt trong, ô tô điện, gia công CNC, mô phỏng phần tử hữu hạn FEM, vật liệu cơ khí, khí động học.', N'chế_tạo_máy cơ_điện_tử động_cơ_đốt_trong gia_công_cnc phần_tử_hữu_hạn kết_cấu_thép truyền_động_thủy_lực ma_sát_mài_mòn ô_tô_điện'),
('LV-KTQL', N'Kinh tế, Quản trị kinh doanh và Tài chính', N'Kinh tế vĩ mô, tài chính ngân hàng, kế toán kiểm toán, quản trị chuỗi cung ứng, thương mại điện tử, marketing số, hành vi người tiêu dùng, kinh tế xanh.', N'kinh_tế_vĩ_mô tài_chính_ngân_hàng chuỗi_cung_ứng chứng_khoán quản_trị_nhân_lực hành_vi_người_tiêu_dùng thương_mại_điện_tử tăng_trưởng_kinh_tế'),
('LV-HSMT', N'Hóa học, Sinh học, Y Dược và Môi trường', N'Công nghệ sinh học, vật liệu nano, tổng hợp hữu cơ, xử lý nước thải, biến đổi khí hậu, dược liệu học, phân tích độc chất, sinh học phân tử.', N'công_nghệ_sinh_học vật_liệu_nano tổng_hợp_hữu_cơ xử_lý_nước_thải kim_loại_nặng dược_liệu kháng_khuẩn xúc_tác_quang_hóa enzyme_sinh_học'),
('LV-XHNV', N'Khoa học Xã hội, Nhân văn và Giáo dục học', N'Phương pháp giảng dạy đại học, tâm lý học giáo dục, ngôn ngữ học đối chiếu, văn hóa học, xã hội học, chính sách công, pháp luật đại cương.', N'khoa_học_giáo_dục phương_pháp_giảng_dạy tâm_lý_học ngôn_ngữ_học văn_hóa_dân_gian xã_hội_học năng_lực_người_học chương_trình_đào_tạo');
GO

-- 2.3. Dữ liệu Bảng LoaiCongTrinh
INSERT INTO dbo.LoaiCongTrinh (maCodeLoai, tenLoai, nhomCongTrinh, gioChuanQuyDoi) VALUES
('LCT-ISI-Q1', N'Bài báo tạp chí quốc tế Uy tín (Hạng 1 / Hạng 2)', N'Bài báo quốc tế', 180.0),
('LCT-SCOPUS', N'Bài báo tạp chí quốc tế (Hạng 3 / Hạng 4) & Hội nghị quốc tế uy tín', N'Bài báo quốc tế', 120.0),
('LCT-HDGSNN', N'Bài báo tạp chí khoa học trong nước thuộc danh mục HĐGSNN', N'Bài báo trong nước', 60.0),
('LCT-DT-BO', N'Đề tài nghiên cứu khoa học cấp Bộ / cấp Tỉnh - Thành phố', N'Đề tài NCKH', 240.0),
('LCT-DT-TRG', N'Đề tài nghiên cứu khoa học cấp Trường / Đề tài Sinh viên', N'Đề tài NCKH', 90.0),
('LCT-SANGCHE', N'Bằng độc quyền Sáng chế / Giải pháp hữu ích / Sách chuyên khảo', N'Sở hữu trí tuệ & Sách', 200.0);
GO

-- 2.4. Dữ liệu Bảng TapChiHopLe
INSERT INTO dbo.TapChiHopLe (tenTapChi, chiSoISSN, xepHang, maLinhVuc, uyTin) VALUES
(N'Tạp chí Ứng dụng Kỹ thuật Trí tuệ Nhân tạo (Quốc tế)', '0952-1976', N'Tạp chí Quốc tế Hạng 1', 1, 1),
(N'Tạp chí Tin học và Điều khiển học (Viện Hàn lâm KH&CN VN)', '1813-9663', N'Hội đồng GSNN (1.0 điểm)', 1, 1),
(N'Tạp chí Điện tử Công nghiệp IEEE (Quốc tế)', '0278-0046', N'Tạp chí Quốc tế Hạng 1', 2, 1),
(N'Tạp chí Quốc tế Khoa học Cơ khí', '0020-7403', N'Tạp chí Quốc tế Hạng 1', 3, 1),
(N'Tạp chí Kinh tế và Quản trị Châu Á', '1049-0078', N'Tạp chí Quốc tế Hạng 2 / HĐGSNN', 4, 1),
(N'Tạp chí Sản xuất Sạch hơn và Công nghệ Môi trường', '0959-6526', N'Tạp chí Quốc tế Hạng 1', 5, 1),
(N'Tạp chí Khoa học Giáo dục Việt Nam', '2615-8957', N'Hội đồng GSNN (0.75 điểm)', 6, 1),
(N'Hội đồng Nghiệm thu Đề tài KH&CN Cấp Trường / Cấp Bộ', 'HĐ-VN', N'Quyết định nghiệm thu', 1, 1),
(N'Tạp chí Đa ngành Quốc tế Chưa Kiểm định (CẢNH BÁO)', '9999-001X', N'Ngoài danh mục / Không uy tín', 1, 0),
(N'Tạp chí Xuất bản Nhanh Chưa Xác thực (Ngừng chỉ mục)', '8888-404X', N'Ngừng chỉ mục / Không uy tín', 4, 0);
GO

-- 2.5. Dữ liệu 10 Hồ sơ mẫu trong Bảng HoSoKHCN
INSERT INTO dbo.HoSoKHCN (
    maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri, namCongBo, kyBaoCao,
    maNguoiNop, maLinhVuc, maLoai, maTapChi, trangThai, fileDinhKem, vanBanDaTachTu,
    maLinhVucAI, doTinCayAI, maLinhVucSVM, doTinCaySVM, daXacNhanAI, diemBatThuong, coCanhBao, chiTietCanhBao, doTuongDongMax, hoSoTuongDongNhat
) VALUES
('HS-2026-0001',
 N'Nghiên cứu tinh chỉnh mô hình ngôn ngữ PhoBERT kết hợp học máy SVM trong phân loại tự động hồ sơ khoa học tiếng Việt',
 N'Bài báo đề xuất kiến trúc tiền xử lý tách từ tiếng Việt bằng thư viện PyVi kết hợp fine-tuning mô hình học sâu vinai/phobert-base và thuật toán Support Vector Machine để phân loại chính xác lĩnh vực chuyên ngành cho các công trình nghiên cứu khoa học tại trường đại học.',
 N'trí tuệ nhân tạo, học sâu, xử lý ngôn ngữ tự nhiên, phobert, phân loại văn bản, học máy',
 N'Nguyễn Thành Tài, Nguyễn Thành Phát', N'Khoa Công nghệ Thông tin', 2026, N'Đợt 1 - Năm học 2025-2026',
 4, 1, 1, 1, N'Đã duyệt', 'MinhChung_HS20260001_PhoBERT_Q1.pdf',
 N'nghiên_cứu tinh_chỉnh mô_hình ngôn_ngữ phobert kết_hợp học_máy svm phân_loại tự_động hồ_sơ khoa_học tiếng_việt trí_tuệ_nhân_tạo học_sâu xử_lý_ngôn_ngữ_tự_nhiên',
 1, 94.8, 1, 89.5, 1, 0.12, 0, N'[]', 18.4, ''),

('HS-2026-0002',
 N'Thiết kế bộ điều khiển thích nghi bền vững cho hệ thống nghịch lưu ba pha nối lưới điện thông minh tích hợp điện mặt trời',
 N'Nghiên cứu trình bày thuật toán điều khiển tự động trên vi điều khiển DSP TMS320F28335 nhằm ổn định điện áp, giảm méo hài tổng THD và tối ưu công suất phát của biến tần năng lượng mặt trời trong lưới điện thông minh.',
 N'lưới điện thông minh, năng lượng mặt trời, điều khiển tự động, điện tử công suất, vi điều khiển',
 N'TS. Vũ Đình Trọng, PGS.TS. Lê Hoàng Việt', N'Khoa Điện - Điện tử', 2026, N'Đợt 1 - Năm học 2025-2026',
 6, 2, 1, 3, N'Chờ xét duyệt', 'MinhChung_HS20260002_SmartGrid_IEEE.pdf',
 N'thiết_kế bộ điều_khiển thích_nghi bền_vững hệ_thống nghịch_lưu ba_pha lưới_điện_thông_minh năng_lượng_mặt_trời điều_khiển_tự_động vi_điều_khiển',
 2, 92.6, 2, 88.2, 0, 0.14, 0, N'[]', 15.2, ''),

('HS-2026-0003',
 N'Tổng hợp vật liệu nano ZnO pha tạp bạc ứng dụng xúc tác quang hóa phân hủy chất ô nhiễm hữu cơ và kháng khuẩn trong xử lý nước thải',
 N'Công trình thực hiện tổng hợp hạt vật liệu nano bằng phương pháp thủy nhiệt, phân tích cấu trúc tinh thể XRD, SEM và đánh giá hiệu suất xử lý nước thải dệt nhuộm nhiễm kim loại nặng cùng hoạt tính kháng khuẩn E.coli.',
 N'vật liệu nano, xúc tác quang hóa, xử lý nước thải, kháng khuẩn, ô nhiễm môi trường',
 N'PGS.TS. Ngô Thanh Hà, TS. Bùi Thị Lan', N'Khoa Kỹ thuật Hóa học & Môi trường', 2026, N'Đợt 1 - Năm học 2025-2026',
 7, 4, 1, 6, N'Chờ xét duyệt', 'MinhChung_HS20260003_Nano_Water.pdf',
 N'tổng_hợp vật_liệu_nano zno pha_tạp bạc ứng_dụng xúc_tác_quang_hóa phân_hủy chất ô_nhiễm hữu_cơ kháng_khuẩn xử_lý_nước_thải kim_loại_nặng',
 5, 91.4, 5, 86.8, 0, 0.82, 1,
 N'["Khai lệch lĩnh vực chuyên ngành: Tác giả khai ''Kinh tế, Quản trị kinh doanh và Tài chính'' nhưng AI (PhoBERT) phân tích nội dung thuộc ''Hóa học, Sinh học, Y Dược và Môi trường'' (91.4%)."]',
 12.5, ''),

('HS-2026-0004',
 N'Tác động của quản trị chuỗi cung ứng xanh và chuyển đổi số đến hiệu quả tài chính của các doanh nghiệp niêm yết trên thị trường chứng khoán',
 N'Sử dụng mô hình cấu trúc tuyến tính SEM và dữ liệu bảng từ báo cáo tài chính kiểm toán của 250 doanh nghiệp nhằm đánh giá mối quan hệ giữa chuỗi cung ứng bền vững, quản trị rủi ro tài chính ngân hàng và tăng trưởng kinh tế.',
 N'chuỗi cung ứng, tài chính ngân hàng, chứng khoán, kinh tế vĩ mô, kế toán kiểm toán',
 N'ThS. Đặng Thị Mai Hương, TS. Lê Thu Trang', N'Khoa Kinh tế & Quản trị Kinh doanh', 2026, N'Đợt 1 - Năm học 2025-2026',
 5, 4, 2, 5, N'Đã duyệt', 'MinhChung_HS20260004_GreenSupplyChain.pdf',
 N'tác_động quản_trị chuỗi_cung_ứng xanh chuyển_đổi_số hiệu_quả tài_chính doanh_nghiệp niêm_yết thị_trường chứng_khoán tài_chính_ngân_hàng kinh_tế_vĩ_mô',
 4, 93.5, 4, 88.9, 1, 0.11, 0, N'[]', 14.0, ''),

('HS-2026-0005',
 N'Ứng dụng mạng nơ ron học sâu và thị giác máy tính trong nhận dạng khuôn mặt điểm danh lớp học thông minh',
 N'Hệ thống tích hợp camera nhận diện khuôn mặt thời gian thực bằng mạng tích chập sâu CNN và cơ sở dữ liệu đám mây hỗ trợ quản lý điểm danh sinh viên tự động.',
 N'thị giác máy tính, học sâu, mạng nơ ron, trí tuệ nhân tạo, công nghệ thông tin',
 N'ThS. Đặng Thị Mai Hương', N'Khoa Kinh tế & Quản trị Kinh doanh', 2026, N'Đợt 1 - Năm học 2025-2026',
 5, 1, 2, 9, N'Đang xét duyệt', 'MinhChung_HS20260005_FaceRec_Warning.pdf',
 N'ứng_dụng mạng_nơ_ron học_sâu thị_giác_máy_tính nhận_dạng khuôn_mặt điểm_danh lớp_học thông_minh trí_tuệ_nhân_tạo công_nghệ_thông_tin',
 1, 90.2, 1, 85.4, 0, 0.88, 1,
 N'["Cảnh báo tạp chí không uy tín: Công trình đăng tại ''Tạp chí Đa ngành Quốc tế Chưa Kiểm định (CẢNH BÁO)'' nằm trong danh sách cảnh báo Không uy tín / Ngừng chỉ mục."]',
 19.8, ''),

('HS-2026-0006',
 N'Mô phỏng phần tử hữu hạn FEM và tối ưu hóa chế độ cắt khi gia công phay CNC hợp kim titan dùng trong chế tạo máy',
 N'Nghiên cứu phân tích trường nhiệt độ, lực cắt cơ học, biến dạng dẻo và mài mòn dụng cụ cắt khi gia công tốc độ cao hợp kim Ti-6Al-4V trên trung tâm gia công CNC 5 trục.',
 N'gia công cnc, chế tạo máy, phần tử hữu hạn, ma sát mài mòn, cơ khí động lực',
 N'TS. Hoàng Gia Huy, TS. Nguyễn Đức Mạnh', N'Khoa Cơ khí Động lực', 2026, N'Đợt 1 - Năm học 2025-2026',
 8, 3, 1, 4, N'Đang xét duyệt', 'MinhChung_HS20260006_CNC_FEM.pdf',
 N'mô_phỏng phần_tử_hữu_hạn fem tối_ưu_hóa chế_độ cắt gia_công_cnc hợp_kim titan chế_tạo_máy ma_sát_mài_mòn cơ_khí',
 3, 94.1, 3, 89.0, 1, 0.10, 0, N'[]', 11.3, ''),

('HS-2026-0007',
 N'Thiết kế hệ thống cảm biến đeo tay IoT kết hợp thuật toán phân tích tín hiệu sinh học theo dõi nhịp tim cho người cao tuổi',
 N'Đề tài kết hợp thiết kế mạch thu nhận tín hiệu điện tâm đồ ECG trên vi điều khiển công suất thấp với phần mềm cảnh báo sớm rối loạn nhịp tim qua thiết bị di động.',
 N'mạng cảm biến không dây, hệ thống nhúng, y sinh học, xử lý tín hiệu',
 N'TS. Vũ Đình Trọng, Nguyễn Thành Tài', N'Khoa Điện - Điện tử', 2026, N'Đợt 1 - Năm học 2025-2026',
 6, 2, 4, 8, N'Chờ xét duyệt', 'MinhChung_HS20260007_IoT_Bio.pdf',
 N'thiết_kế hệ_thống cảm_biến đeo_tay iot kết_hợp thuật_toán phân_tích tín_hiệu sinh_học theo_dõi nhịp_tim mạng_cảm_biến_không_dây hệ_thống_nhúng',
 2, 68.5, 2, 64.0, 0, 0.28, 0, N'[]', 22.1, ''),

('HS-2026-0008',
 N'Nghiên cứu ứng dụng mô hình ngôn ngữ tiền huấn luyện PhoBERT và SVM phân loại tự động văn bản hồ sơ khoa học tiếng Việt',
 N'Đề tài trình bày quy trình tiền xử lý tách từ tiếng Việt bằng PyVi kết hợp fine-tuning mô hình học sâu vinai/phobert-base và thuật toán học máy Support Vector Machine nhằm tự động phân loại lĩnh vực chuyên ngành hồ sơ nghiên cứu khoa học.',
 N'trí tuệ nhân tạo, học sâu, xử lý ngôn ngữ tự nhiên, phobert, phân loại văn bản',
 N'Nguyễn Thành Tài', N'Khoa Công nghệ Thông tin', 2025, N'Đợt 2 - Năm học 2024-2025',
 4, 1, 5, 2, N'Từ chối', 'MinhChung_HS20260008_Duplicate.pdf',
 N'nghiên_cứu ứng_dụng mô_hình ngôn_ngữ tiền_huấn_luyện phobert svm phân_loại tự_động văn_bản hồ_sơ khoa_học tiếng_việt trí_tuệ_nhân_tạo học_sâu',
 1, 95.1, 1, 90.2, 1, 0.91, 1,
 N'["Nghi vấn trùng lặp nội dung nghiên cứu: Độ tương đồng văn bản Cosine đạt 82.4% so với hồ sơ HS-2026-0001."]',
 82.4, 'HS-2026-0001'),

('HS-2026-0009',
 N'Đổi mới phương pháp giảng dạy đại học theo tiếp cận CDIO và chuẩn đầu ra AUN-QA nhằm phát triển năng lực tự học cho sinh viên',
 N'Nghiên cứu khảo sát thực trạng và đề xuất quy trình thiết kế chương trình đào tạo kết hợp học tập qua dự án giúp nâng cao kỹ năng tư duy phản biện và năng lực nghề nghiệp.',
 N'phương pháp giảng dạy, giáo dục đại học, chương trình đào tạo, khoa học giáo dục, năng lực người học',
 N'TS. Đỗ Văn Kiên, TS. Nguyễn Thị Hồng', N'Khoa Sư phạm & Khoa học Xã hội Nhân văn', 2026, N'Đợt 1 - Năm học 2025-2026',
 4, 6, 3, 7, N'Đã duyệt', 'MinhChung_HS20260009_GiaoDuc_CDIO.pdf',
 N'đổi_mới phương_pháp_giảng_dạy giáo_dục_đại_học tiếp_cận cdio chuẩn_đầu_ra aun_qa phát_triển năng_lực_người_học chương_trình_đào_tạo khoa_học_giáo_dục',
 6, 93.8, 6, 88.4, 1, 0.09, 0, N'[]', 10.5, ''),

('HS-2026-0010',
 N'Quy trình sản xuất chế phẩm sinh học giàu enzyme cellulase và than sinh học biến tính xử lý kim loại nặng trong nước thải công nghiệp',
 N'Giải pháp hữu ích đề xuất quy trình lên men chủng xạ khuẩn bản địa kết hợp vật liệu hấp phụ sinh học giúp loại bỏ hiệu quả chì Pb(II) và cadimi Cd(II) bảo vệ môi trường bền vững.',
 N'công nghệ sinh học, enzyme sinh học, kim loại nặng, xử lý nước thải, ô nhiễm môi trường',
 N'PGS.TS. Ngô Thanh Hà', N'Khoa Kỹ thuật Hóa học & Môi trường', 2026, N'Đợt 1 - Năm học 2025-2026',
 7, 5, 6, 6, N'Chờ xét duyệt', 'MinhChung_HS20260010_SangChe_Bio.pdf',
 N'quy_trình sản_xuất chế_phẩm sinh_học enzyme_sinh_học cellulase than_sinh_học biến_tính xử_lý_nước_thải kim_loại_nặng công_nghệ_sinh_học ô_nhiễm_môi_trường',
 5, 92.9, 5, 87.6, 0, 0.14, 0, N'[]', 24.6, '');
GO

-- 2.6. Dữ liệu Bảng LichSuXetDuyet
INSERT INTO dbo.LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia) VALUES
('HS-2026-0001', 4, N'Khởi tạo', N'Chờ xét duyệt', N'Tác giả nộp hồ sơ KH&CN mới lên hệ thống kèm minh chứng toàn văn PDF.'),
('HS-2026-0001', 2, N'Chờ xét duyệt', N'Đang xét duyệt', N'Cán bộ tiếp nhận thẩm định; kiểm tra nhãn AI PhoBERT gợi ý khớp 100% với chuyên ngành CNTT.'),
('HS-2026-0001', 2, N'Đang xét duyệt', N'Đã duyệt', N'Hồ sơ hợp lệ, bài báo đăng trên tạp chí ISI/Scopus Q1 uy tín, đủ điều kiện công nhận giờ chuẩn NCKH.'),
('HS-2026-0004', 5, N'Khởi tạo', N'Chờ xét duyệt', N'Tác giả nộp hồ sơ bài báo quốc tế Scopus Q2.'),
('HS-2026-0004', 3, N'Chờ xét duyệt', N'Đã duyệt', N'Đã kiểm tra chỉ số ISSN và xác nhận nhãn lĩnh vực Kinh tế & Quản trị kinh doanh do AI đề xuất.'),
('HS-2026-0005', 5, N'Khởi tạo', N'Chờ xét duyệt', N'Nộp hồ sơ bài báo nghiên cứu nhận diện khuôn mặt.'),
('HS-2026-0005', 2, N'Chờ xét duyệt', N'Đang xét duyệt', N'Hệ thống AI cảnh báo tạp chí thuộc danh sách Predatory (9999-001X) và lệch lịch sử lĩnh vực của tác giả trên đồ thị NetworkX. Đang yêu cầu giải trình.'),
('HS-2026-0006', 8, N'Khởi tạo', N'Chờ xét duyệt', N'Khai báo bài báo quốc tế chuyên ngành Cơ khí chế tạo máy.'),
('HS-2026-0006', 3, N'Chờ xét duyệt', N'Đang xét duyệt', N'Đã xác nhận nhãn AI (Kỹ thuật Cơ khí). Đang chờ Hội đồng khoa xác minh minh chứng bản in.'),
('HS-2026-0008', 4, N'Khởi tạo', N'Chờ xét duyệt', N'Khai báo đề tài cấp Trường.'),
('HS-2026-0008', 2, N'Chờ xét duyệt', N'Từ chối', N'Từ chối phê duyệt: Mô-đun Isolation Forest & Cosine Similarity phát hiện nội dung tóm tắt trùng lặp trên 80% với công trình HS-2026-0001.'),
('HS-2026-0009', 4, N'Khởi tạo', N'Chờ xét duyệt', N'Khai báo bài báo Tạp chí Khoa học Giáo dục Việt Nam.'),
('HS-2026-0009', 3, N'Chờ xét duyệt', N'Đã duyệt', N'Hồ sơ đầy đủ minh chứng, thuộc danh mục HĐGSNN 0.75 điểm, phê duyệt.');
GO

-- 2.7. Dữ liệu Bảng PhanHoiAI (Feedback Loop)
INSERT INTO dbo.PhanHoiAI (maHoSo, maCanBo, maLinhVucAI, maLinhVucThucTe, doTinCayLucDo, ketQuaKhop, vanBanHuanLuyen, ghiChuPhanHoi) VALUES
('HS-2026-0001', 2, 1, 1, 94.8, 1, N'nghiên_cứu tinh_chỉnh mô_hình ngôn_ngữ phobert kết_hợp học_máy svm phân_loại tự_động hồ_sơ khoa_học tiếng_việt', N'Cán bộ xác nhận nhãn do PhoBERT dự đoán hoàn toàn chính xác.'),
('HS-2026-0004', 3, 4, 4, 93.5, 1, N'tác_động quản_trị chuỗi_cung_ứng xanh chuyển_đổi_số hiệu_quả tài_chính doanh_nghiệp niêm_yết thị_trường chứng_khoán', N'Cán bộ xác nhận nhãn do PhoBERT dự đoán hoàn toàn chính xác.'),
('HS-2026-0006', 3, 3, 3, 94.1, 1, N'mô_phỏng phần_tử_hữu_hạn fem tối_ưu_hóa chế_độ cắt gia_công_cnc hợp_kim titan chế_tạo_máy', N'Cán bộ xác nhận nhãn do PhoBERT dự đoán hoàn toàn chính xác.'),
('HS-2026-0009', 3, 6, 6, 93.8, 1, N'đổi_mới phương_pháp_giảng_dạy giáo_dục_đại_học tiếp_cận cdio chuẩn_đầu_ra aun_qa phát_triển năng_lực_người_học', N'Cán bộ xác nhận nhãn do PhoBERT dự đoán hoàn toàn chính xác.');
GO

-- 2.8. Dữ liệu Bảng ThucNghiemMoHinhAI
INSERT INTO dbo.ThucNghiemMoHinhAI (tenMoHinh, precisionScore, recallScore, f1Score, accuracyScore, thoiGianPhanHoiMs, soMauHuanLuyen, soMauKiemThu, moTaChiTiet) VALUES
(N'Baseline: TF-IDF + LinearSVC (Support Vector Machine)', 87.45, 86.80, 87.12, 87.50, 3.20, 36, 12, N'TfidfVectorizer(ngram_range=(1,2)) + LinearSVC; tốc độ huấn luyện và phản hồi cực nhanh, tốn ít tài nguyên phần cứng.'),
(N'Deep Learning: Fine-tuning PhoBERT (vinai/phobert-base)', 93.60, 92.95, 93.27, 93.40, 42.50, 36, 12, N'Pre-trained vinai/phobert-base, đóng băng tầng ẩn dưới + Classification Head (LR=2e-5, Batch=16, AdamW); hiểu sâu ngữ cảnh khoa học phức tạp.');
GO
