# HỆ THỐNG HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC CÔNG NGHỆ BẰNG PHƯƠNG PHÁP PHÂN LOẠI TỰ ĐỘNG VÀ PHÁT HIỆN BẤT THƯỜNG (CNTT-KLCN001)

Dự án Khóa luận Tốt nghiệp triển khai đầy đủ theo đề cương chính thức **`CNTT-KLCN001_BuiCongDanh.pdf`** (lộ trình **Tuần 1 đến Tuần 10**), tích hợp phân loại tự động đa mục tiêu (**Lĩnh vực** & **Loại công trình**) bằng `PyTorch` + `vinai/phobert-base` và phát hiện bất thường bằng đồ thị `NetworkX` + `Isolation Forest` + `PyTorch Autoencoder`.

---

## 📁 Cấu trúc Thư mục Dự án

1. **Báo cáo Khóa luận Tốt nghiệp & Bộ dữ liệu chuẩn 13 trường**:
   - [`Bao_Cao_Khoa_Luan.docx`](Bao_Cao_Khoa_Luan.docx): Báo cáo Khóa luận Tốt nghiệp hoàn chỉnh được định dạng chuẩn theo mẫu `Mau bao cao_Do an_Khoa luan_2025.docx`.
   - [`docs/Tap_Du_Lieu_Huan_Luyen_Chuan_12Truong.xlsx`](docs/Tap_Du_Lieu_Huan_Luyen_Chuan_12Truong.xlsx): Bộ dữ liệu thực nghiệm chuẩn hóa dùng chung gồm 240 mẫu (180 Train / 60 Test độc lập theo chủ đề) với đầy đủ 13 trường dữ liệu đầu vào (gồm 3 sheet: Tập dữ liệu 240 mẫu, Thống kê phân bố dữ liệu, Bảng đối sánh thực nghiệm 2 mô hình & so sánh với phân loại thủ công).
   - [`docs/Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json`](docs/Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json): Kết quả thống kê thực nghiệm đo thực tế giữa `TF-IDF + LinearSVC` (`F1 = 82.26%`), `Fine-tuning PhoBERT` (`F1 = 89.92%`) và phân loại thủ công (`F1 = 73.94%`).

2. **Hệ thống Sơ đồ Vẽ Trực Tiếp cho Extension Draw.io trên VS Code — `diagrams/`**:
   *(Mở trực tiếp các tệp `.drawio` trong VS Code bằng extension **Draw.io Integration** - `hediet.vscode-drawio`)*
   - [`diagrams/01_UseCase_NghiepVu_TongQuan.drawio`](diagrams/01_UseCase_NghiepVu_TongQuan.drawio): Sơ đồ Use-Case nghiệp vụ tổng quan.
   - [`diagrams/02_UseCase_HeThong_ChiTiet.drawio`](diagrams/02_UseCase_HeThong_ChiTiet.drawio): Sơ đồ Use-Case hệ thống phân rã chi tiết.
   - [`diagrams/03_Analysis_Class_Diagram_BCE.drawio`](diagrams/03_Analysis_Class_Diagram_BCE.drawio): Sơ đồ lớp phân tích Boundary - Control - Entity.
   - [`diagrams/04_ERD_CoSoDuLieu_VatLy.drawio`](diagrams/04_ERD_CoSoDuLieu_VatLy.drawio): Sơ đồ ERD cơ sở dữ liệu vật lý.
   - [`diagrams/05_Wireframe_GiaoDien_HeThong.drawio`](diagrams/05_Wireframe_GiaoDien_HeThong.drawio): Thiết kế Wireframe các màn hình cốt lõi của ứng dụng Web.
   - [`diagrams/06_KienTruc_AI_PhanLoai_Va_BatThuong.drawio`](diagrams/06_KienTruc_AI_PhanLoai_Va_BatThuong.drawio): Sơ đồ kiến trúc Pipeline AI (Tiền xử lý tiếng Việt PyVi, TF-IDF+SVM, PyTorch PhoBERT Fine-tuning, NetworkX + Isolation Forest & PyTorch Autoencoder).

3. **Mã nguồn Ứng dụng Web & Mô-đun Trí tuệ Nhân tạo**:
   - [`app.py`](app.py): Ứng dụng Flask chính, định tuyến API, xác thực Session (`check_password_hash`), phân quyền RBAC (`@role_required`), CRUD hồ sơ KH&CN 13 trường (lọc theo Từ khóa, Nhóm công trình, Lĩnh vực, Loại, Tác giả, Đơn vị, Trạng thái, Cảnh báo), quản lý 4 danh mục (Thêm/Sửa/Ẩn-Hiện), quy trình xét duyệt và Feedback Loop.
   - [`database.py`](database.py): Thiết kế CSDL `SQLite 3` (`khcn_system.db`) gồm 9 bảng chuẩn 3NF: `NguoiDung`, `LinhVuc`, `LoaiCongTrinh`, `TapChiHopLe`, `DanhMucTacGia`, `HoSoKHCN`, `LichSuXetDuyet`, `PhanHoiAI`, `ThucNghiemMoHinhAI`.
   - [`model_benchmark.py`](model_benchmark.py): Phân hệ chạy độc lập và đối sánh công bằng 100% giữa `TF-IDF + LinearSVC` và `Fine-tuning PhoBERT` trên bộ dữ liệu 240 mẫu, kèm đối sánh với phân loại thủ công.
   - [`ai_engine.py`](ai_engine.py): Pipeline tiền xử lý tiếng Việt (`PyVi`), suy luận phân loại đa mục tiêu (Lĩnh vực & Loại công trình) bằng `PyTorch vinai/phobert-base`, và phát hiện bất thường bằng `NetworkX + Isolation Forest + PyTorch Autoencoder (nn.Module)`.
   - [`phobert_finetune.py`](phobert_finetune.py): Kiến trúc `PyTorch` Fine-tuning thực tế `vinai/phobert-base` (mở khóa 4 tầng Transformer trên cùng `layer 8..11` kết hợp 2 đầu phân loại `field_head` & `work_type_head`, tối ưu bằng `AdamW`, `epochs=8`, `batch_size=8`, lưu trọng số tại `models/phobert_khcn/phobert_classifier.pt`).
   - [`reporting_service.py`](reporting_service.py): Xuất báo cáo thống kê ra định dạng Excel (`.xlsx` với `openpyxl`) và PDF hành chính (`.pdf` hỗ trợ Unicode tiếng Việt với `reportlab`).

---

## 🚀 Hướng dẫn Chạy Ứng dụng Web

```powershell
# 1. Cài đặt các thư viện cần thiết
pip install -r requirements.txt

# 2. Khởi chạy máy chủ Web Flask
python app.py
```

Sau đó truy cập tại trình duyệt: **`http://127.0.0.1:5000`**

### 🔑 Tài khoản Demo Phân quyền RBAC:
- **Quản trị viên (Admin):** `admin@uni.edu.vn` / `Admin@123`
- **Cán bộ xét duyệt (CanBo):** `canbo@uni.edu.vn` / `CanBo@123`
- **Tác giả / Giảng viên (TacGia):** `tacgia@uni.edu.vn` / `TacGia@123`
