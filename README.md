# HỆ THỐNG HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC VÀ CÔNG NGHỆ TÍCH HỢP PHÂN LOẠI TỰ ĐỘNG (CNTT-KLCN001)

Dự án Khóa luận Tốt nghiệp triển khai đầy đủ theo **Tài liệu Yêu cầu Sản phẩm (PRD)** từ **Tuần 1 đến Tuần 6** (kèm mô-đun Phát hiện bất thường đồ thị `NetworkX` + `Isolation Forest`).

---

## 📁 Cấu trúc Thư mục Dự án

1. **Báo cáo Khóa luận Tốt nghiệp & Bộ dữ liệu chuẩn 12 trường**:
   - [`Bao_Cao_Khoa_Luan.docx`](Bao_Cao_Khoa_Luan.docx): Báo cáo Khóa luận Tốt nghiệp hoàn chỉnh được định dạng chuẩn theo mẫu `Mau bao cao_Do an_Khoa luan_2025.docx`.
   - [`docs/Tap_Du_Lieu_Huan_Luyen_Chuan_12Truong.xlsx`](docs/Tap_Du_Lieu_Huan_Luyen_Chuan_12Truong.xlsx): Bộ dữ liệu huấn luyện & kiểm thử dùng chung gồm 240 mẫu (180 Train / 60 Test) với đầy đủ 12 trường dữ liệu đầu vào chuẩn.
   - [`docs/Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json`](docs/Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json): Kết quả thống kê thực nghiệm đối sánh độc lập giữa `TF-IDF + SVM` và `Fine-tuning PhoBERT`.

2. **Hệ thống Sơ đồ Vẽ Trực Tiếp cho Extension Draw.io trên VS Code — `diagrams/`**:
   *(Mở trực tiếp các tệp `.drawio` trong VS Code bằng extension **Draw.io Integration** - `hediet.vscode-drawio`)*
   - [`diagrams/01_UseCase_NghiepVu_TongQuan.drawio`](diagrams/01_UseCase_NghiepVu_TongQuan.drawio): Sơ đồ Use-Case nghiệp vụ tổng quan.
   - [`diagrams/02_UseCase_HeThong_ChiTiet.drawio`](diagrams/02_UseCase_HeThong_ChiTiet.drawio): Sơ đồ Use-Case hệ thống phân rã chi tiết 5 phân hệ.
   - [`diagrams/03_Analysis_Class_Diagram_BCE.drawio`](diagrams/03_Analysis_Class_Diagram_BCE.drawio): Sơ đồ lớp phân tích Boundary - Control - Entity.
   - [`diagrams/04_ERD_CoSoDuLieu_VatLy.drawio`](diagrams/04_ERD_CoSoDuLieu_VatLy.drawio): Sơ đồ ERD cơ sở dữ liệu vật lý.
   - [`diagrams/05_Wireframe_GiaoDien_HeThong.drawio`](diagrams/05_Wireframe_GiaoDien_HeThong.drawio): Thiết kế Wireframe các màn hình cốt lõi của ứng dụng Web.
   - [`diagrams/06_KienTruc_AI_PhanLoai_Va_BatThuong.drawio`](diagrams/06_KienTruc_AI_PhanLoai_Va_BatThuong.drawio): Sơ đồ kiến trúc Pipeline AI (Tiền xử lý tiếng Việt PyVi, TF-IDF+SVM, PhoBERT Fine-tuning, NetworkX + Isolation Forest & Autoencoder).

3. **Mã nguồn Ứng dụng Web & Mô-đun Trí tuệ Nhân tạo**:
   - [`app.py`](app.py): Ứng dụng Flask chính, định tuyến API, phân hệ xác thực Session (`check_password_hash`), Decorator phân quyền RBAC (`@role_required`), CRUD hồ sơ KH&CN 12 trường, quản lý danh mục, quy trình xét duyệt 5 bước và thực nghiệm mô hình.
   - [`database.py`](database.py): Thiết kế CSDL (`khcn_system.db` / SQL Server) gồm các bảng `NguoiDung`, `LinhVuc`, `LoaiCongTrinh`, `TapChiHopLe`, `HoSoKHCN`, `LichSuXetDuyet`, `PhanHoiAI`, `ThucNghiemMoHinhAI`.
   - [`model_benchmark.py`](model_benchmark.py): Phân hệ chạy độc lập và đối sánh triệt để `TF-IDF + SVM (LinearSVC)` và `Fine-tuning PhoBERT` trên bộ dữ liệu 240 mẫu 12 trường, tự động chọn `Fine-tuning PhoBERT` làm mô hình chính.
   - [`ai_engine.py`](ai_engine.py): Pipeline tiền xử lý tiếng Việt (`PyVi`), suy luận phân loại chuyên ngành với mô hình chính `Fine-tuning PhoBERT`, và phát hiện bất thường bằng `NetworkX + Isolation Forest + Deep Autoencoder`.
   - [`phobert_finetune.py`](phobert_finetune.py): Kiến trúc PyTorch Fine-tuning `vinai/phobert-base` (đóng băng các tầng ẩn bên dưới, bổ sung `Classification Head`, tối ưu bằng `AdamW`, `lr=2e-5`, `batch_size=16`).
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
