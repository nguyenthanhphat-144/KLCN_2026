# HỆ THỐNG HỖ TRỢ QUẢN LÝ HỒ SƠ KHOA HỌC VÀ CÔNG NGHỆ TÍCH HỢP PHÂN LOẠI TỰ ĐỘNG (CNTT-KLCN001)

Dự án Khóa luận Tốt nghiệp triển khai đầy đủ theo **Tài liệu Yêu cầu Sản phẩm (PRD)** từ **Tuần 1 đến Tuần 6** (kèm mô-đun Phát hiện bất thường đồ thị `NetworkX` + `Isolation Forest`).

---

## 📁 Cấu trúc Thư mục Dự án

1. **Tài liệu Khảo sát & Báo cáo Word (`.docx`) — `docs/`**:
   - [`docs/01_BaoCao_KhaoSat_HienTrang_MucTieu_Va_KeHoach_Tuan1.docx`](docs/01_BaoCao_KhaoSat_HienTrang_MucTieu_Va_KeHoach_Tuan1.docx): Báo cáo khảo sát hiện trạng thực tế cực kỳ chi tiết (Cơ cấu tổ chức 5 cấp, phân tích 5 nhóm hồ sơ KH&CN, quy trình thủ công 5 bước, mẫu phiếu khảo sát thực địa, 6 hạn chế cốt lõi, mục tiêu SMART, ma trận phân quyền RBAC 15 chức năng, kế hoạch phân công nhóm Backend/Frontend/AI và ma trận RACI).
   - [`docs/02_TaiLieu_PRD_PhanTich_ThietKe_ToanTap_CNTT_KLCN001.docx`](docs/02_TaiLieu_PRD_PhanTich_ThietKe_ToanTap_CNTT_KLCN001.docx): Tài liệu tổng hợp toàn bộ dự án từ Tuần 1 đến Tuần 7 (Khảo sát, Đặc tả Use-Case, Sơ đồ lớp BCE, Lý thuyết NLP TF-IDF+SVM & PhoBERT, Thiết kế CSDL SQLite, Lý thuyết NetworkX + Isolation Forest, và Bảng đối sánh thực nghiệm).

2. **Hệ thống Sơ đồ Vẽ Trực Tiếp cho Extension Draw.io trên VS Code — `diagrams/`**:
   *(Mở trực tiếp các tệp `.drawio` trong VS Code bằng extension **Draw.io Integration** - `hediet.vscode-drawio`)*
   - [`diagrams/01_UseCase_NghiepVu_TongQuan.drawio`](diagrams/01_UseCase_NghiepVu_TongQuan.drawio): Sơ đồ Use-Case nghiệp vụ tổng quan (Tuần 2 - Mục 1a).
   - [`diagrams/02_UseCase_HeThong_ChiTiet.drawio`](diagrams/02_UseCase_HeThong_ChiTiet.drawio): Sơ đồ Use-Case hệ thống phân rã chi tiết 5 phân hệ (Tuần 2 - Mục 1b).
   - [`diagrams/03_Analysis_Class_Diagram_BCE.drawio`](diagrams/03_Analysis_Class_Diagram_BCE.drawio): Sơ đồ lớp phân tích Boundary - Control - Entity (Tuần 2 - Mục 2a, 2b, 2c).
   - [`diagrams/04_ERD_CoSoDuLieu_VatLy.drawio`](diagrams/04_ERD_CoSoDuLieu_VatLy.drawio): Sơ đồ ERD cơ sở dữ liệu vật lý SQLite (Tuần 3 - Mục 1a, 1b).
   - [`diagrams/05_Wireframe_GiaoDien_HeThong.drawio`](diagrams/05_Wireframe_GiaoDien_HeThong.drawio): Thiết kế Wireframe các màn hình cốt lõi của ứng dụng Web (Tuần 3 - Mục 2a & Tuần 6 - Mục 2d).
   - [`diagrams/06_KienTruc_AI_PhanLoai_Va_BatThuong.drawio`](diagrams/06_KienTruc_AI_PhanLoai_Va_BatThuong.drawio): Sơ đồ kiến trúc Pipeline AI (Tiền xử lý tiếng Việt PyVi, TF-IDF+SVM, PhoBERT Fine-tuning, NetworkX + Isolation Forest & Feedback Loop).

3. **Mã nguồn Ứng dụng Web & Mô-đun Trí tuệ Nhân tạo**:
   - [`app.py`](app.py): Ứng dụng Flask chính, định tuyến API, phân hệ xác thực Session (`check_password_hash`), Decorator phân quyền RBAC (`@role_required`), CRUD hồ sơ KH&CN, quản lý danh mục, quy trình xét duyệt và Feedback Loop.
   - [`database.py`](database.py): Thiết kế CSDL vật lý SQLite (`khcn_system.db`) gồm các bảng `NguoiDung`, `LinhVuc`, `LoaiCongTrinh`, `TapChiHopLe`, `HoSoKHCN`, `LichSuXetDuyet`, `PhanHoiAI`, `ThucNghiemMoHinhAI`.
   - [`ai_engine.py`](ai_engine.py): Pipeline tiền xử lý tiếng Việt (hợp nhất 4 trường, làm sạch, lọc từ dừng, tách từ bằng `PyVi`), huấn luyện mô hình Baseline (`TF-IDF + LinearSVC`) và Deep Learning (`PhoBERT`), đánh giá `Precision`, `Recall`, `F1-score`, và phát hiện bất thường bằng `NetworkX + Isolation Forest`.
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
