# HƯỚNG DẪN KIỂM THỬ HỆ THỐNG QUẢN LÝ HỒ SƠ KH&CN (KLTN 2026)

Tài liệu hướng dẫn thực thi bộ kiểm thử độc lập cho Hệ thống Quản lý Hồ sơ KH&CN tích hợp AI (Đề tài CNTT-KLCN001).

---

## 1. Danh sách 7 Kịch bản Kiểm thử

| Mã | Tên Script | Mục Tiêu Kiểm Thử | Chế Độ |
| :--- | :--- | :--- | :---: |
| **TEST-1** | `scripts/test_db_integrity.py` | Kiểm tra tính toàn vẹn 8 bảng CSDL, quan hệ khóa ngoại, 3 vai trò RBAC | Read-Only |
| **TEST-2** | `scripts/test_ai_inference.py` | Kiểm thử tiền xử lý PyVi, mô hình TF-IDF+LinearSVC, đồ thị NetworkX, Isolation Forest | Read-Only |
| **TEST-3** | `scripts/test_report_generation.py` | Kiểm thử sinh luồng báo cáo Excel (.xlsx kèm biểu đồ) và PDF (.pdf Unicode tiếng Việt) | Memory Stream |
| **TEST-4** | `scripts/test_rbac_access_control.py` | Kiểm thử phân quyền RBAC (Admin, CanBo, TacGia), mã hóa mật khẩu bcrypt | Read-Only |
| **TEST-5** | `scripts/test_hoso_workflow.py` | Kiểm thử quy trình 5 trạng thái hồ sơ và bảng log lịch sử xét duyệt | Read-Only |
| **TEST-6** | `scripts/test_search_filter.py` | Kiểm thử tra cứu từ khóa full-text, bộ lọc đa tiêu chí (lĩnh vực, loại, năm, phân trang) | Read-Only |
| **TEST-7** | `scripts/test_phobert_pipeline.py` | Kiểm thử cấu hình siêu tham số PhoBERT, ánh xạ 6 nhãn, cờ fallback an toàn | Read-Only |

---

## 2. Cách Thực Thi

### Chạy Toàn Bộ Bộ Kiểm Thử (Khuyên dùng)
```powershell
python -X utf8 scripts/run_all_tests.py
```

### Chạy Từng Kịch Bản Độc Lập
```powershell
python -X utf8 scripts/test_db_integrity.py
python -X utf8 scripts/test_ai_inference.py
python -X utf8 scripts/test_report_generation.py
python -X utf8 scripts/test_rbac_access_control.py
python -X utf8 scripts/test_hoso_workflow.py
python -X utf8 scripts/test_search_filter.py
python -X utf8 scripts/test_phobert_pipeline.py
```

---

## 3. Đặc Điểm An Toàn Kỹ Thuật
- Mọi kiểm thử kết nối SQLite đều kích hoạt cờ `file:...db?mode=ro` (chỉ đọc), cam kết **không làm biến đổi hay hư hại dữ liệu sản phẩm đang chạy**.
- Tích hợp cờ môi trường `HF_HUB_OFFLINE=1` và `TRANSFORMERS_OFFLINE=1` đảm bảo kiểm thử chạy ngoại tuyến 100%, không bị ảnh hưởng bởi kết nối Internet.
