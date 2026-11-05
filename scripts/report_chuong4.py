# -*- coding: utf-8 -*-
"""
CHƯƠNG 4: CÀI ĐẶT HỆ THỐNG
Bao gồm:
- 4.1 Môi trường cài đặt và cấu hình phần mềm
  - Bảng 4.1: Bảng đặc tả môi trường phần cứng, phần mềm và các thư viện chuyên dụng
- 4.2 Cấu trúc thư mục và đặc tả các module mã nguồn
- 4.3 Hiện thực các phân hệ chức năng cốt lõi và giao diện hệ thống
  - Hình 4.1: Giao diện Tổng quan Bảng điều khiển (Dashboard) và Chỉ số KPI Nghiên cứu
  - Hình 4.2: Giao diện Biểu mẫu Khai báo Hồ sơ KH&CN tích hợp Gợi ý AI thời gian thực (Live AI Preview)
  - Hình 4.3: Giao diện Chi tiết Xét duyệt Hồ sơ với Quy trình 5 bước và Kiểm tra Bất thường
  - Hình 4.4: Giao diện Trung tâm AI & Phân tích Đánh giá Hiệu năng Mô hình
"""

import os
from docx.shared import Pt, Cm, RGBColor

def add_chuong_4(doc, helpers):
    add_chapter_title = helpers["add_chapter_title"]
    add_h1 = helpers["add_h1"]
    add_h2 = helpers["add_h2"]
    add_h3 = helpers["add_h3"]
    add_body = helpers["add_body"]
    add_bullet_1 = helpers["add_bullet_1"]
    add_bullet_2 = helpers["add_bullet_2"]
    add_figure_custom = helpers["add_figure_custom"]
    add_table_custom = helpers["add_table_custom"]

    # ==================== BẮT ĐẦU CHƯƠNG 4 ====================
    add_chapter_title(doc, "CHƯƠNG 4: CÀI ĐẶT HỆ THỐNG")

    # 4.1 MÔI TRƯỜNG CÀI ĐẶT
    add_h1(doc, "4.1. MÔI TRƯỜNG CÀI ĐẶT VÀ CẤU HÌNH PHẦN MỀM")
    add_body(doc, "Toàn bộ hệ thống quản lý hồ sơ KH&CN tích hợp AI được cài đặt, huấn luyện mô hình và kiểm thử vận hành thực tế trên môi trường máy chủ thử nghiệm đạt các tiêu chuẩn kỹ thuật sau:")

    t41_headers = ["Hạng mục thành phần", "Thông số kỹ thuật chi tiết", "Vai trò và ghi chú trong hệ thống"]
    t41_data = [
        ["Bộ vi xử lý (CPU)", "Intel Core i7-12700H (14 nhân, 20 luồng, xung nhịp lên tới 4.7 GHz)", "Xử lý đa luồng máy chủ web Flask, truy vấn CSDL và tiền xử lý văn bản PyVi"],
        ["Bộ nhớ trong (RAM)", "16.0 GB DDR4 Bus 3200 MHz", "Nạp mô hình ngôn ngữ PhoBERT vào bộ nhớ và duy trì hàng đợi đệm dữ liệu"],
        ["Bộ xử lý đồ họa (GPU)", "NVIDIA GeForce RTX 3050 Laptop GPU (4GB VRAM GDDR6)", "Tăng tốc độ tính toán ma trận trong quá trình fine-tuning PhoBERT và Deep Autoencoder"],
        ["Hệ điều hành", "Microsoft Windows 11 Pro 64-bit / Tương thích hoàn toàn Ubuntu Linux 22.04 LTS", "Môi trường thực thi máy chủ ứng dụng"],
        ["Môi trường thực thi chính", "Python 3.12 64-bit", "Ngôn ngữ lập trình cốt lõi của toàn bộ hệ thống"],
        ["Vi khung phát triển Web", "Flask 3.0.x (WSGI Server)", "Điều phối định tuyến HTTP, xử lý phiên làm việc Session và xác thực RBAC"],
        ["Hệ quản trị CSDL", "SQLite 3.42+ (Tích hợp sẵn trong Python)", "Lưu trữ dữ liệu quan hệ chuẩn 3NF, hỗ trợ giao dịch ACID"],
        ["Thư viện Học sâu (Deep Learning)", "PyTorch 2.4.x / TorchVision / Transformers (Hugging Face) 4.40+", "Hiện thực kiến trúc mạng PhoBERT và Deep Autoencoder"],
        ["Thư viện Học máy truyền thống", "Scikit-learn 1.5.x, NumPy, SciPy", "Hiện thực LinearSVC, Platt Scaling, TF-IDF và Isolation Forest"],
        ["Thư viện Xử lý tiếng Việt", "PyVi (Python Vietnamese Core NLP toolkit)", "Tách từ ghép tiếng Việt học thuật chuẩn xác trước khi nạp vào PhoBERT"],
        ["Thư viện Phân tích Đồ thị", "NetworkX 3.2+", "Xây dựng đồ thị đồng tác giả và tính toán Degree Centrality"],
        ["Thư viện Báo cáo & Tài liệu", "python-docx, openpyxl, reportlab", "Tạo lập báo cáo thống kê tự động dạng Word (.docx), Excel (.xlsx) và PDF"]
    ]
    add_table_custom(doc, 4, "Bảng đặc tả môi trường phần cứng, phần mềm và các thư viện chuyên dụng",
                     t41_headers, t41_data, col_widths=[4.0, 5.5, 6.5])

    # 4.2 CẤU TRÚC THƯ MỤC VÀ ĐẶC TẢ CÁC MODULE
    add_h1(doc, "4.2. CẤU TRÚC THƯ MỤC VÀ ĐẶC TẢ CÁC MODULE MÃ NGUỒN")
    add_body(doc, "Dự án được tổ chức khoa học theo mô hình cấu trúc mô-đun hóa cao, phân định ranh giới rõ ràng giữa mã nguồn điều khiển web, động cơ AI, kết nối cơ sở dữ liệu, giao diện khuôn mẫu và tài nguyên tĩnh:")

    add_bullet_1(doc, "app.py: Module điều phối máy chủ web Flask chính, khai báo hơn 30 endpoint điều hướng, kiểm soát đăng nhập/đăng xuất, phân quyền người dùng RBAC, điều phối quy trình xét duyệt 5 bước và xuất báo cáo Excel/PDF.", bold_prefix="app.py: ")
    add_bullet_1(doc, "database.py: Module quản trị cơ sở dữ liệu SQLite 3, thiết lập cấu trúc bảng, ràng buộc khóa ngoại, khởi tạo dữ liệu mẫu mặc định (seed data) và cung cấp các hàm thao tác dữ liệu an toàn chống SQL Injection.", bold_prefix="database.py: ")
    add_bullet_1(doc, "ai_engine.py: Động cơ Trí tuệ Nhân tạo phân loại văn bản, chịu trách nhiệm tải trọng số tiền huấn luyện vinai/phobert-base, xây dựng kiến trúc phân loại đa nhiệm, hiện thực pipeline TF-IDF + LinearSVC và cung cấp hàm API dự đoán gợi ý Live AI Preview.", bold_prefix="ai_engine.py: ")
    add_bullet_1(doc, "anomaly_detector.py: Module phát hiện bất thường và tính điểm rủi ro, hiện thực hóa giải thuật kiểm tra tạp chí săn mồi, đo độ tương đồng Cosine TF-IDF, phân tích đồ thị NetworkX, chạy mô hình Rừng cô lập (Isolation Forest) và mạng Deep Autoencoder (PyTorch).", bold_prefix="anomaly_detector.py: ")
    add_bullet_1(doc, "templates/: Thư mục chứa toàn bộ giao diện HTML Jinja2 thuần Việt (dashboard.html, profiles.html, profile_form.html, review_detail.html, ai_center.html, catalogs.html) được thiết kế hiện đại bằng Bootstrap 5.", bold_prefix="templates/: ")
    add_bullet_1(doc, "static/: Thư mục tài nguyên tĩnh bao gồm tệp kiểu dáng css/style.css, thư viện biểu tượng FontAwesome, các tệp JavaScript xử lý tương tác bất đồng bộ (AJAX) và tệp tin ảnh logo HUIT chính thức.", bold_prefix="static/: ")

    # 4.3 HIỆN THỰC CÁC PHÂN HỆ VÀ GIAO DIỆN
    add_h1(doc, "4.3. HIỆN THỰC CÁC PHÂN HỆ CHỨC NĂNG CỐT LÕI VÀ GIAO DIỆN HỆ THỐNG")
    add_body(doc, "Dưới đây là hình ảnh chụp giao diện thực tế từ hệ thống đang vận hành hoàn chỉnh trên môi trường thực nghiệm:")

    add_h2(doc, "4.3.1. Phân hệ Bảng điều khiển Tổng quan (Dashboard) và Chỉ số KPI")
    add_body(doc, "Giao diện Bảng điều khiển cung cấp cái nhìn toàn diện về bức tranh nghiên cứu khoa học của toàn trường: tổng số hồ sơ tiếp nhận, số công trình đã nghiệm thu, số hồ sơ đang xử lý, số lượng cảnh báo rủi ro cần lưu ý, cùng các biểu đồ trực quan phân bố theo lĩnh vực chuyên môn và nhóm loại công trình:")

    add_figure_custom(doc, "images/hinh4_dashboard_overview.jpg", 4,
                      "Giao diện Bảng điều khiển Tổng quan (Dashboard) và Chỉ số KPI Nghiên cứu", width_cm=15.0)

    add_h2(doc, "4.3.2. Phân hệ Khai báo Hồ sơ KH&CN tích hợp Gợi ý AI thời gian thực (Live AI Preview)")
    add_body(doc, "Khi giảng viên thực hiện kê khai công trình nghiên cứu, tính năng Live AI Preview sẽ lắng nghe nội dung văn bản nhập vào (Tiêu đề, Tóm tắt, Từ khóa). Sau khoảng thời gian trì hoãn ngắn (debounce 350ms), mô hình PhoBERT sẽ tự động phân tích và đưa ra gợi ý Lĩnh vực nghiên cứu chính xác nhất kèm tỷ lệ tin cậy phần trăm hiển thị trực quan ngay cạnh ô chọn:")

    add_figure_custom(doc, "images/hinh4_hoso_form_ai.jpg", 4,
                      "Giao diện Biểu mẫu Khai báo Hồ sơ KH&CN tích hợp Gợi ý AI thời gian thực (Live AI Preview)", width_cm=15.0)

    add_h2(doc, "4.3.3. Phân hệ Chi tiết Xét duyệt Hồ sơ theo Quy trình 5 bước và Kiểm tra Bất thường")
    add_body(doc, "Màn hình dành cho Cán bộ Phòng Quản lý Khoa học hiển thị chi tiết mọi thông số của hồ sơ, tệp đính kèm minh chứng, thanh tiến trình trực quan thể hiện rõ 5 bước xét duyệt, và bảng phân tích rủi ro AI chi tiết (cho thấy điểm rủi ro tổng hợp, các cảnh báo về mã ISSN, độ tương đồng văn bản và dị biệt tham số). Chuyên viên có thể thực hiện duyệt hợp lệ, yêu cầu bổ sung thông tin hoặc từ chối kèm lý do được lưu vết bất biến:")

    add_figure_custom(doc, "images/hinh4_hoso_detail_review.jpg", 4,
                      "Giao diện Chi tiết Xét duyệt Hồ sơ với Quy trình 5 bước và Kiểm tra Bất thường", width_cm=15.0)

    add_h2(doc, "4.3.4. Phân hệ Trung tâm AI & Đánh giá Hiệu năng So sánh Mô hình")
    add_body(doc, "Giao diện Trung tâm AI cho phép chuyên viên quản trị chạy thử nghiệm so sánh trực tiếp giữa mô hình học sâu Tinh chỉnh PhoBERT tiếng Việt và mô hình học máy truyền thống TF-IDF + LinearSVC trên cùng một tập dữ liệu chuẩn hóa 13 trường. Giao diện hiển thị rõ ràng bảng kết quả thực nghiệm khách quan, thời gian thực thi, các chỉ số Macro F1-score và các lý luận chuyên sâu hỗ trợ ra quyết định:")

    add_figure_custom(doc, "images/hinh4_ai_center.jpg", 4,
                      "Giao diện Trung tâm AI & Phân tích Đánh giá Hiệu năng Mô hình", width_cm=15.0)

    return
