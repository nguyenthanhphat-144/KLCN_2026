# -*- coding: utf-8 -*-
"""
CHƯƠNG 3: THIẾT KẾ HỆ THỐNG
Bao gồm:
- 3.1 Thiết kế kiến trúc tổng thể hệ thống (3-Tier MVC)
  - Hình 3.1: Kiến trúc phân tầng (3-Tier Architecture) của hệ thống
- 3.2 Thiết kế cơ sở dữ liệu quan hệ (SQLite 3, chuẩn 3NF)
  - Hình 3.2: Sơ đồ quan hệ thực thể (ERD) cơ sở dữ liệu hệ thống
  - Bảng 3.1: Đặc tả cấu trúc bảng hồ sơ khoa học công nghệ (scientific_profiles)
  - Bảng 3.2: Đặc tả cấu trúc bảng cảnh báo bất thường (anomaly_audit_logs)
- 3.3 Thiết kế kiến trúc phân hệ Trí tuệ Nhân tạo (AI Engine)
  - Hình 3.3: Pipeline xử lý và phân loại tự động hồ sơ bằng mô hình PhoBERT và SVM
  - Hình 3.4: Quy trình đa tầng phát hiện bất thường và tính điểm rủi ro tổng hợp
- 3.4 Thiết kế giao diện người dùng và tương tác thời gian thực
"""

import os
from docx.shared import Pt, Cm, RGBColor

def add_chuong_3(doc, helpers):
    add_chapter_title = helpers["add_chapter_title"]
    add_h1 = helpers["add_h1"]
    add_h2 = helpers["add_h2"]
    add_h3 = helpers["add_h3"]
    add_body = helpers["add_body"]
    add_bullet_1 = helpers["add_bullet_1"]
    add_bullet_2 = helpers["add_bullet_2"]
    add_figure_custom = helpers["add_figure_custom"]
    add_table_custom = helpers["add_table_custom"]

    # ==================== BẮT ĐẦU CHƯƠNG 3 ====================
    add_chapter_title(doc, "CHƯƠNG 3: THIẾT KẾ HỆ THỐNG")

    # 3.1 THIẾT KẾ KIẾN TRÚC TỔNG THỂ HỆ THỐNG
    add_h1(doc, "3.1. THIẾT KẾ KIẾN TRÚC TỔNG THỂ HỆ THỐNG")
    add_body(doc, "Nhằm đáp ứng các yêu cầu cao về tính sẵn sàng, khả năng mở rộng, khả năng bảo trì và tách biệt rõ ràng các mối quan tâm (Separation of Concerns), hệ thống được thiết kế theo mô hình Kiến trúc 3 tầng (3-Tier Architecture) kết hợp mẫu hình MVC (Model - View - Controller). Kiến trúc tổng thể được biểu diễn trong sơ đồ phân tầng sau:")

    add_figure_custom(doc, "images/hinh3_kientruc_he_thong.png", 3,
                      "Kiến trúc phân tầng (3-Tier Architecture) của hệ thống quản lý hồ sơ KH&CN", width_cm=14.5)

    add_body(doc, "Các tầng kiến trúc trong hệ thống đảm nhiệm các vai trò cụ thể như sau:")
    add_bullet_1(doc, "Tầng Trình diễn (Presentation Layer / View): Xây dựng trên nền tảng HTML5, CSS3, JavaScript và thư viện giao diện Bootstrap 5 thuần Việt. Tầng này cung cấp các giao diện tương tác: Bảng điều khiển KPI (Dashboard), Biểu mẫu khai báo hồ sơ động tích hợp gợi ý Live AI Preview, Màn hình xét duyệt hồ sơ chi tiết theo quy trình 5 bước, và Trung tâm phân tích AI so sánh mô hình.", bold_prefix="Tầng 1: ")
    add_bullet_1(doc, "Tầng Nghiệp vụ và Động cơ Trí tuệ Nhân tạo (Application & AI Engine Layer / Controller): Được viết bằng ngôn ngữ Python sử dụng vi khung Flask. Tầng này xử lý toàn bộ logic nghiệp vụ (xác thực RBAC, điều chuyển trạng thái hồ sơ, xuất báo cáo Excel/PDF) và tích hợp các động cơ AI chuyên sâu: module tiền xử lý văn bản PyVi, bộ trích xuất đặc trưng TF-IDF, bộ phân loại LinearSVC, mô hình học sâu VinAI PhoBERT chạy trên PyTorch, và module phát hiện bất thường kết hợp NetworkX, Isolation Forest và Deep Autoencoder.", bold_prefix="Tầng 2: ")
    add_bullet_1(doc, "Tầng Dữ liệu và Lưu trữ (Data Persistence Layer / Model): Sử dụng hệ quản trị cơ sở dữ liệu quan hệ SQLite 3 được thiết kế chuẩn hóa ở mức Dạng chuẩn 3 (3NF). Tầng này chịu trách nhiệm lưu trữ bền vững thông tin hồ sơ KH&CN, các bảng danh mục chuẩn hóa, nhật ký kiểm toán (Audit Logs), báo cáo bất thường và các tệp tin minh chứng số.", bold_prefix="Tầng 3: ")

    # 3.2 THIẾT KẾ CƠ SỞ DỮ LIỆU QUAN HỆ
    add_h1(doc, "3.2. THIẾT KẾ CƠ SỞ DỮ LIỆU QUAN HỆ (SQLite 3)")
    add_body(doc, "Cơ sở dữ liệu được thiết kế tuân thủ nghiêm ngặt các quy tắc chuẩn hóa nhằm loại bỏ hoàn toàn hiện tượng dư thừa dữ liệu (Data Redundancy) và các dị thường khi Thêm, Sửa, Xóa (Anomalies). Dưới đây là Sơ đồ quan hệ thực thể (Entity-Relationship Diagram - ERD) thể hiện cấu trúc các bảng và mối liên kết khóa ngoại giữa chúng:")

    add_figure_custom(doc, "images/hinh_erd.png", 3,
                      "Sơ đồ quan hệ thực thể (ERD) cơ sở dữ liệu hệ thống", width_cm=14.5)

    add_body(doc, "Dưới đây là bảng đặc tả chi tiết cấu trúc trường dữ liệu của 02 bảng trọng yếu nhất trong hệ thống:")

    t31_headers = ["Tên trường (Field Name)", "Kiểu dữ liệu", "Ràng buộc", "Mô tả ý nghĩa nghiệp vụ"]
    t31_data = [
        ["id", "INTEGER", "PRIMARY KEY AUTOINCREMENT", "Định danh duy nhất của hồ sơ công trình"],
        ["title", "TEXT", "NOT NULL", "Tiêu đề bài báo / Tên đề tài NCKH / Tên sáng chế"],
        ["abstract", "TEXT", "NULL", "Nội dung tóm tắt công trình (dùng cho trích xuất đặc trưng AI)"],
        ["keywords", "TEXT", "NULL", "Danh sách từ khóa chuyên ngành, phân tách bằng dấu chấm phẩy"],
        ["work_type", "TEXT", "NOT NULL", "Nhóm loại hình: Bài báo khoa học / Đề tài NCKH / Sáng chế"],
        ["field_id", "INTEGER", "FOREIGN KEY -> fields(id)", "Khóa ngoại liên kết lĩnh vực chuyên môn được duyệt"],
        ["ai_suggested_field_id", "INTEGER", "FOREIGN KEY -> fields(id)", "Lĩnh vực do AI phân loại tự động gợi ý"],
        ["ai_confidence", "REAL", "DEFAULT 0.0", "Độ tin cậy phần trăm của mô hình AI (0.00 đến 1.00)"],
        ["journal_name", "TEXT", "NULL", "Tên tạp chí xuất bản hoặc Hội thảo khoa học"],
        ["issn", "TEXT", "NULL", "Mã chuẩn quốc tế ISSN / eISSN của tạp chí"],
        ["status", "TEXT", "DEFAULT 'pending'", "Trạng thái quy trình: pending, processing, revision, approved, rejected, archived"],
        ["risk_score", "REAL", "DEFAULT 0.0", "Điểm rủi ro bất thường tổng hợp (Composite Risk Score)"],
        ["evidence_file", "TEXT", "NULL", "Đường dẫn lưu trữ tệp tài liệu minh chứng điện tử"],
        ["created_at", "DATETIME", "DEFAULT CURRENT_TIMESTAMP", "Thời điểm nộp hồ sơ vào hệ thống"]
    ]
    add_table_custom(doc, 3, "Đặc tả cấu trúc bảng hồ sơ khoa học công nghệ (scientific_profiles)",
                     t31_headers, t31_data, col_widths=[3.8, 2.5, 3.8, 5.9])

    t32_headers = ["Tên trường", "Kiểu dữ liệu", "Ràng buộc", "Mô tả ý nghĩa nghiệp vụ"]
    t32_data = [
        ["id", "INTEGER", "PRIMARY KEY AUTOINCREMENT", "Định danh bản ghi cảnh báo"],
        ["profile_id", "INTEGER", "FOREIGN KEY -> scientific_profiles(id)", "Liên kết đến hồ sơ KH&CN bị cảnh báo"],
        ["anomaly_type", "TEXT", "NOT NULL", "Loại rủi ro: predatory_journal, duplicate_title, missing_evidence, outlier_score"],
        ["severity_level", "TEXT", "NOT NULL", "Mức độ nghiêm trọng: info, warning, high, critical"],
        ["risk_value", "REAL", "NOT NULL", "Giá trị điểm rủi ro thành phần tính toán được"],
        ["description", "TEXT", "NOT NULL", "Mô tả chi tiết nguyên nhân cảnh báo để chuyên viên kiểm tra"],
        ["detected_at", "DATETIME", "DEFAULT CURRENT_TIMESTAMP", "Thời điểm thuật toán AI phát hiện cảnh báo"]
    ]
    add_table_custom(doc, 3, "Đặc tả cấu trúc bảng cảnh báo bất thường (anomaly_audit_logs)",
                     t32_headers, t32_data, col_widths=[3.5, 2.5, 4.0, 6.0])

    # 3.3 THIẾT KẾ PHÂN HỆ AI
    add_h1(doc, "3.3. THIẾT KẾ KIẾN TRÚC PHÂN HỆ TRÍ TUỆ NHÂN TẠO (AI ENGINE)")
    add_body(doc, "Phân hệ Trí tuệ Nhân tạo là trái tim thông minh của hệ thống, bao gồm hai module độc lập nhưng phối hợp chặt chẽ: (1) Module Phân loại tự động đa nhiệm và (2) Module Phát hiện bất thường đa tầng.")

    add_h2(doc, "3.3.1. Thiết kế Pipeline xử lý và phân loại văn bản tự động")
    add_body(doc, "Pipeline phân loại văn bản được thiết kế với chuỗi bước chuẩn hóa khép kín: Dữ liệu văn bản thô từ 13 trường chuẩn hóa -> Tiền xử lý và làm sạch ký tự lạ -> Tách từ ghép tiếng Việt PyVi -> Trích xuất biểu diễn ngữ cảnh -> Đưa vào bộ phân loại PhoBERT hoặc TF-IDF + LinearSVC -> Trả về nhãn dự đoán và xác suất tin cậy:")

    add_figure_custom(doc, "images/hinh2_kientruc_ai_pipeline.png", 3,
                      "Pipeline xử lý và phân loại tự động hồ sơ bằng mô hình PhoBERT và SVM", width_cm=14.5)

    add_body(doc, "Cả hai mô hình đều dùng chung một hàm ghép nối văn bản đầu vào thống nhất (get_model_input_text) để bảo đảm sự công bằng tuyệt đối trong đánh giá so sánh hiệu năng thực nghiệm.")

    add_h2(doc, "3.3.2. Thiết kế Kiến trúc phát hiện bất thường đa tầng")
    add_body(doc, "Kiến trúc phát hiện bất thường hoạt động theo cơ chế kiểm tra đa tầng nối tiếp (Multi-layered Anomaly Detection Pipeline), phối hợp giữa các luật kiểm tra chuyên gia (Heuristic Rules) và các mô hình học máy không giám sát:")

    add_figure_custom(doc, "images/hinh2_phat_hien_bat_thuong.png", 3,
                      "Quy trình đa tầng phát hiện bất thường và tính điểm rủi ro tổng hợp", width_cm=14.5)

    add_body(doc, "Quy trình bao gồm 4 tầng rà soát liên hoàn:")
    add_bullet_1(doc, "Tầng 1 - Kiểm tra quy chuẩn pháp quy & Tạp chí săn mồi: So khớp mã ISSN với danh mục tạp chí uy tín của Nhà trường và danh mục các tạp chí có dấu hiệu săn mồi (Beall's List / Predatory Journals); kiểm tra sự tồn tại của tệp minh chứng PDF.", bold_prefix="Tầng 1: ")
    add_bullet_1(doc, "Tầng 2 - Phát hiện trùng lặp văn bản: Trích xuất vector TF-IDF và tính độ tương đồng Cosine giữa hồ sơ hiện tại với toàn bộ kho dữ liệu lịch sử để phát hiện sao chép công trình.", bold_prefix="Tầng 2: ")
    add_bullet_1(doc, "Tầng 3 - Khai phá đồ thị đồng tác giả (NetworkX): Phân tích mạng lưới liên kết học thuật, đo độ trung tâm Degree Centrality để tìm các liên kết ảo hoặc tác giả đứng tên bất thường.", bold_prefix="Tầng 3: ")
    add_bullet_1(doc, "Tầng 4 - Mô hình học máy phát hiện dị biệt (Isolation Forest & Deep Autoencoder): Tính điểm dị biệt cấu trúc s_IF từ rừng cây cô lập và sai số tái tạo đặc trưng định lượng MSE_AE từ mạng nơ-ron sâu PyTorch.", bold_prefix="Tầng 4: ")

    # 3.4 THIẾT KẾ GIAO DIỆN NGƯỜI DÙNG VÀ TƯƠNG TÁC THỜI GIAN THỰC
    add_h1(doc, "3.4. THIẾT KẾ GIAO DIỆN NGƯỜI DÙNG VÀ TƯƠNG TÁC THỜI GIAN THỰC")
    add_body(doc, "Giao diện hệ thống được thiết kế theo nguyên lý Công thái học phần mềm (Software Ergonomics) và tư duy lấy người dùng làm trung tâm (User-Centered Design):")
    add_bullet_1(doc, "Cơ chế Live AI Preview với kỹ thuật Trì hoãn sự kiện (Debounce): Để tránh gửi quá nhiều yêu cầu suy luận AI gây quá tải máy chủ khi người dùng đang gõ phím, hệ thống cài đặt bộ đệm trễ 350ms. Khi người dùng dừng gõ sau 350ms, hàm JavaScript mới tự động kích hoạt cuộc gọi API nền (Fetch API) để nhận gợi ý phân loại từ PhoBERT.", bold_prefix="Debounce 350ms: ")
    add_bullet_1(doc, "Hệ thống màu sắc cảnh báo trực quan: Sử dụng chuẩn mã màu quốc tế dễ nhận biết (Xanh lá - An toàn, Vàng cam - Lưu ý, Đỏ tươi - Nguy cơ vi phạm cao) kết hợp các biểu tượng đồ họa FontAwesome, giúp chuyên viên nắm bắt ngay tình trạng hồ sơ chỉ trong vài giây quan sát.", bold_prefix="Trực quan hóa màu sắc: ")
    add_bullet_1(doc, "Tính tương thích cao trên mọi thiết bị: Bố cục lưới Bootstrap Grid linh hoạt, tự động co giãn từ màn hình máy tính bàn làm việc (1920x1080) đến máy tính bảng và điện thoại thông minh của giảng viên.", bold_prefix="Responsive: ")

    return
