# BÁO CÁO CẬP NHẬT TIẾN ĐỘ VÀ NÂNG CẤP HỆ THỐNG
## Đề tài: Hệ thống Hỗ trợ Quản lý Hồ sơ KH&CN Tích hợp Phân loại Tự động và Phát hiện Bất thường (CNTT-KLCN001)
**Đơn vị:** Trường Đại học Công Thương TP. Hồ Chí Minh (HUIT)  
**Nhóm sinh viên thực hiện:**
1. **Nguyễn Thành Phát** (Vai trò: Quản trị viên hệ thống - Admin)
2. **Lê Thanh Phong** (Vai trò: Cán bộ xét duyệt KH&CN - CanBo)
3. **Nguyễn Thành Tài** (Vai trò: Tác giả / Giảng viên nghiên cứu - TacGia)

**Ngày cập nhật:** 08/11/2026  
**Kho lưu trữ Git:** `https://github.com/nguyenthanhphat-144/KLCN_2026.git`  
**Phiên bản:** `v1.2.0-Final-Curated`

---

## 1. TỔNG QUAN NỘI DUNG ĐÃ THỰC HIỆN

Trong đợt cập nhật và hoàn thiện này, nhóm đã thực hiện toàn diện 5 nhóm công việc trọng tâm nhằm đáp ứng đầy đủ yêu cầu báo cáo mới nhất của Hội đồng đánh giá khóa luận tốt nghiệp:

```
┌─────────────────────────────────────────────────────────────────────────┐
│                    CÁC HẠNG MỤC CẬP NHẬT CHÍNH                          │
├─────────────────────────────────────────────────────────────────────────┤
│ 1. Ẩn danh hóa dữ liệu tác giả & Chuẩn hóa thông tin 3 sinh viên        │
│ 2. Tích hợp bộ 35 hồ sơ thực tế chọn lọc từ file HUIT mới nhất          │
│ 3. Hoàn thiện Chương 5 Báo cáo Khóa luận (Bảng 5.5 & Bảng 5.6)          │
│ 4. Nâng cấp giao diện phẳng Dashboard học thuật (Academic Flat UI)      │
│ 5. Xây dựng bộ kiểm thử tự động 7 kịch bản độc lập (100% PASS)          │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## 2. CHI TIẾT CÁC CÔNG VIỆC ĐÃ HOÀN THÀNH

### 2.1. Chuẩn hóa phân quyền và Ẩn danh hóa thông tin tác giả

Theo quy định mới nhất về bảo mật thông tin cá nhân và tính khách quan trong đánh giá nghiên cứu khoa học:
- **Giữ nguyên 100% họ tên thật của 3 sinh viên làm đồ án** trong tài khoản hệ thống:
  - `Nguyễn Thành Phát` (`admin@uni.edu.vn`) — Phụ trách phân hệ Quản trị viên (Admin).
  - `Lê Thanh Phong` (`canbo@uni.edu.vn`) — Phụ trách phân hệ Hội đồng thẩm định & Cán bộ xét duyệt (CanBo).
  - `Nguyễn Thành Tài` (`tacgia@uni.edu.vn`) — Phụ trách phân hệ Tác giả / Giảng viên nộp hồ sơ (TacGia).
- **Ẩn danh toàn bộ tên giảng viên và tác giả khác** theo cấu trúc chuẩn mực: **`[Họ] + [Tên đệm] + [Ký tự đầu của tên chính]`**, tuyệt đối không để trơ trọi 1 chữ cái gây khó phân biệt:
  - `Lê Hoàng Việt` $\rightarrow$ `Lê Hoàng V` (`PGS.TS. Lê Hoàng V (Cán bộ duyệt)`)
  - `Đặng Thị Mai Hương` $\rightarrow$ `Đặng Thị Mai H` (`ThS. Đặng Thị Mai H (GV Kinh tế)`)
  - `Vũ Đình Trọng` $\rightarrow$ `Vũ Đình T` (`TS. Vũ Đình T (GV Điện - ĐT)`)
  - `Ngô Thanh Hà` $\rightarrow$ `Ngô Thanh H` (`PGS.TS. Ngô Thanh H (GV Hóa - Sinh)`)
  - `Hoàng Gia Huy` $\rightarrow$ `Hoàng Gia H` (`TS. Hoàng Gia H (GV Cơ khí)`)
  - `Phạm Văn Khoa, Trần Quốc Bảo` $\rightarrow$ `TS. Phạm Văn K, PGS.TS. Trần Quốc B`
  - `Bùi Thị Lan` $\rightarrow$ `TS. Bùi Thị L`, `Lê Thu Trang` $\rightarrow$ `TS. Lê Thu T`...
- **Phạm vi cập nhật:**
  - Cơ sở dữ liệu hoạt động: Bảng `NguoiDung`, `HoSoKHCN`, `DanhMucTacGia` trong `khcn_system.db`.
  - Toàn bộ các file mã nguồn liên quan: `database.py`, `model_benchmark.py`.
  - Các kịch bản CSDL đa nền tảng: `sql/01_SQLServer_SSMS_Database.sql`, `sql/02_MySQL_Workbench_XAMPP_Database.sql`, `sql/03_SQLite_DBBrowser_Database.sql`.

---

### 2.2. Tích hợp Bộ Dữ liệu Thực tế 35 Hồ sơ Nghiên cứu HUIT

- Trích lọc kỹ lưỡng từ 149 dòng hồ sơ nộp thực tế trong file `Mau_Nhap_Lieu_HoSo_KHCN.xlsx` của Nhà trường.
- Xây dựng tập dữ liệu vàng gồm **35 hồ sơ nghiên cứu** mang tính đại diện cao:
  - **10 hồ sơ gốc chuẩn hóa** làm tập đối chứng đa ngành (CNTT, Điện - Điện tử, Kinh tế, Cơ khí, Hóa - Môi trường, Sư phạm/XHNV).
  - **25 hồ sơ nghiên cứu thực tế chọn lọc** bao gồm đầy đủ các trường hợp kiểm thử đặc biệt:
    + 13 hồ sơ mang cảnh báo bất thường/dị biệt (khai sai lệch lĩnh vực chuyên ngành, trùng lặp nội dung học thuật cao, dị biệt biểu diễn ngữ nghĩa Autoencoder MSE > 0.08, đăng trên tạp chí có dấu hiệu săn mồi/không uy tín).
    + Các hồ sơ thực tế bị khuyết thiếu trường đăng ký lĩnh vực ban đầu để kiểm nghiệm năng lực tự động suy luận điền khuyết thiếu của mô hình PhoBERT.
- Xuất bản file tài liệu dữ liệu trích lọc độc lập tại: `docs/Ket_Qua_Tap_Du_Lieu_Chon_Loc.json`.

---

### 2.3. Cập nhật Báo cáo Khóa luận Hoàn thiện (`Bao_Cao_Khoa_Luan.docx`)

- **Chương 5 (Thử nghiệm và Triển khai):**
  - Cập nhật mục **5.1** phân định rõ hai bộ dữ liệu thực nghiệm: 240 mẫu học máy chuẩn hóa (tỷ lệ 75/25) và tập 35 hồ sơ nghiên cứu thực tế HUIT.
  - Bổ sung **Bảng 5.5**: Đánh giá chi tiết 13 hồ sơ cảnh báo bất thường và dị biệt tiêu biểu từ dữ liệu thực tế HUIT.
  - Bổ sung **Bảng 5.6**: Đánh giá thực nghiệm tính năng tự động điền khuyết thiếu lĩnh vực bằng PhoBERT trên các hồ sơ thực tế (chứng minh PhoBERT đạt độ chính xác 100% với độ tin cậy trung bình > 94.7%, vượt trội so với LinearSVC chỉ đạt 4/10).
- **Phần Kết luận & Hướng phát triển:**
  - Cập nhật số liệu thực nghiệm thực tế, đối sánh hiệu năng phân loại tự động và phát hiện bất thường đa tầng.
- Tự động biên dịch đồng bộ từ mã nguồn tạo báo cáo chuyên dụng (`scripts/report_chuong5.py`, `scripts/report_closing.py`).

---

### 2.4. Nâng cấp Giao diện Dashboard Học thuật Phẳng (Academic Flat UI)

- Thiết kế lại trang tổng quan `templates/dashboard.html` theo tiêu chuẩn giao diện hành chính cơ quan giáo dục đại học:
  - Dải thẻ chỉ số KPI phẳng, trang nhã: Tổng hồ sơ, Đã duyệt, Đang xét duyệt, Chờ tiếp nhận, Yêu cầu bổ sung, Cảnh báo dị biệt.
  - Tích hợp **bảng cuộn mượt (Custom Scrollbar)** với chiều cao tối ưu 250px cho các khối thống kê lĩnh vực, loại công trình và danh mục tạp chí hợp lệ.
  - Nâng cấp 2 biểu đồ trực quan Chart.js:
    + Biểu đồ cột: Phân bố số lượng hồ sơ theo từng Lĩnh vực chuyên ngành.
    + Biểu đồ Doughnut (Cắt tâm 64%): Cơ cấu tỷ lệ hồ sơ theo tiến độ thẩm định 5 trạng thái.
  - Bổ sung chỉ dẫn trạng thái kiểm định AI rõ ràng: *Đạt yêu cầu* (icon xanh) / *Cần thẩm định lại* (kèm điểm dị biệt Isolation Forest).

---

### 2.5. Xây dựng Bộ Kiểm thử Tự động Toàn diện (7/7 Kịch bản PASS 100%)

Xây dựng bộ kịch bản kiểm thử độc lập hoạt động ở chế độ an toàn chỉ đọc (Read-Only) thông qua trình thực thi tập trung `scripts/run_all_tests.py`:

| Mã Kịch Bản | File Thực Thi | Mục Tiêu Kiểm Thử | Thời Gian | Kết Quả |
| :---: | :--- | :--- | :---: | :---: |
| **TEST-1** | `test_db_integrity.py` | Toàn vẹn CSDL (8 bảng, danh mục, RBAC) | 0.09s | **PASS (100%)** |
| **TEST-2** | `test_ai_inference.py` | Pipeline AI (PyVi, TF-IDF+SVM, PhoBERT, Isolation Forest) | 42.42s | **PASS (100%)** |
| **TEST-3** | `test_report_generation.py` | Kết xuất báo cáo động (Excel .xlsx & PDF .pdf) | 1.58s | **PASS (100%)** |
| **TEST-4** | `test_rbac_access_control.py` | Phân quyền RBAC (Schema, Decorators, bcrypt hash) | 0.08s | **PASS (100%)** |
| **TEST-5** | `test_hoso_workflow.py` | Quy trình xét duyệt 5 trạng thái & log lịch sử | 0.09s | **PASS (100%)** |
| **TEST-6** | `test_search_filter.py` | Tra cứu, tìm kiếm từ khóa & bộ lọc đa tiêu chí | 0.08s | **PASS (100%)** |
| **TEST-7** | `test_phobert_pipeline.py` | Cấu hình siêu tham số & suy luận mô hình PhoBERT | 22.30s | **PASS (100%)** |
| **TỔNG HỢP** | **`run_all_tests.py`** | **Toàn bộ hệ thống kiểm định chất lượng phần mềm** | **66.64s** | **7/7 ĐẠT (100%)** |

---

## 3. LỊCH SỬ COMMIT TRÊN GITHUB (TÀI KHOẢN KBT2981-SPEC)

Các đóng góp đã được phân chia thành các commit độc lập, rõ ràng, mang giá trị kỹ thuật cao:

1. `66aa3ee`: `test: bo sung 3 script kiem thu doc lap CSDL, AI pipeline va bao cao Excel/PDF`
2. `ad590bd`: `docs: bo sung chu thich chi tiet co che RBAC, luong xac thuc va quy trinh xet duyet trong app.py`
3. `3397276`: `docs: bo sung mo ta cau truc 8 bang CSDL va quan he khoa ngoai trong database.py`
4. `121e51f`: `test: kiem thu phan quyen RBAC - phan tich CSDL schema va source code decorator`
5. `4bfbeb5`: `test: bo sung kiem thu quy trinh xu ly ho so 5 trang thai va log xet duyet`
6. `d7d1502`: `test: bo sung trinh thuc thi tong hop run_all_tests cho toan bo 5 kich ban kiem thu`
7. `59517d5`: `test: bo sung kiem thu tra cuu, tim kiem tu khoa va bo loc da tieu chi ho so`
8. `4f9ca47`: `docs: bo sung giai thich chi tiet kien truc pipeline AI 3 tang trong ai_engine.py`
9. `7cd0089`: `test: bo sung kiem thu tien xu ly va cau truc dau ra mo hinh PhoBERT`
10. `849008d`: `docs: bo sung chu thich chi tiet sieu tham so va quy trinh fine-tune trong phobert_finetune.py`
11. `97e1020`: `docs: cap nhat bo test suite 7 kich ban va tai lieu huong dan kiem thu he thong`
12. `d738672`: `feat: cap nhat bao cao chuong 5, giao dien dashboard hoc thuat va tap 35 ho so thuc nghiem HUIT`
13. `1619e2d`: `merge: dong bo va hop nhat toan bo 12 commits kiem thu voi ban cap nhat bao cao HUIT moi nhat`
14. `f1fd2bb`: `fix: cap nhat ten 3 thanh vien lam do an trong danh sach nguoi dung va an danh tac gia Ngo Thanh H`
15. `6c99b2e`: `fix: an danh toan bo ten tac gia trong CSDL va ma nguon theo quy dinh bao cao, giu nguyen 3 thanh vien do an`

---

## 4. KẾT LUẬN VÀ CAM KẾT KỸ THUẬT

- **Tính toàn vẹn mã nguồn:** Tuyệt đối không thay đổi tên bất kỳ file nào trong hệ thống, đảm bảo 100% khả năng tương thích ngược.
- **Tính ổn định sản phẩm:** Ứng dụng Flask Web, công cụ kết xuất báo cáo Excel/PDF, mô hình AI PhoBERT và cơ sở dữ liệu SQLite đều đang ở trạng thái hoạt động hoàn hảo.
- **Chuẩn bị nghiệm thu:** Hệ thống đã sẵn sàng cho buổi báo cáo và bảo vệ Khóa luận tốt nghiệp trước Hội đồng chuyên môn Trường Đại học Công Thương TP. Hồ Chí Minh.
