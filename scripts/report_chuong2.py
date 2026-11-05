# -*- coding: utf-8 -*-
"""
CHƯƠNG 2: PHÂN TÍCH HỆ THỐNG
Bao gồm:
- 2.1 Xác định các tác nhân và mô hình phân quyền (RBAC)
  - Bảng 2.1: Danh sách các tác nhân người dùng tham gia tương tác trong hệ thống
  - Bảng 2.2: Ma trận phân quyền chi tiết (RBAC) trên các phân hệ chức năng
- 2.2 Đặc tả yêu cầu chức năng hệ thống (Đầy đủ 6 phân hệ)
- 2.3 Đặc tả yêu cầu phi chức năng
- 2.4 Biểu đồ Use Case và đặc tả Use Case chính
  - Hình 2.1: Biểu đồ Use Case tổng thể hệ thống quản lý hồ sơ KH&CN
  - Đặc tả Use Case UC01, UC02, UC03
- 2.5 Quy trình nghiệp vụ xét duyệt hồ sơ 5 bước và phân tích lớp BCE
  - Hình 2.2: Quy trình nghiệp vụ 5 bước xét duyệt tính hợp lệ của hồ sơ KH&CN
  - Hình 2.3: Sơ đồ phân tích lớp Boundary - Control - Entity (BCE)
"""

import os
from docx.shared import Pt, Cm, RGBColor

def add_chuong_2(doc, helpers):
    add_chapter_title = helpers["add_chapter_title"]
    add_h1 = helpers["add_h1"]
    add_h2 = helpers["add_h2"]
    add_h3 = helpers["add_h3"]
    add_body = helpers["add_body"]
    add_bullet_1 = helpers["add_bullet_1"]
    add_bullet_2 = helpers["add_bullet_2"]
    add_figure_custom = helpers["add_figure_custom"]
    add_table_custom = helpers["add_table_custom"]

    # ==================== BẮT ĐẦU CHƯƠNG 2 ====================
    add_chapter_title(doc, "CHƯƠNG 2: PHÂN TÍCH HỆ THỐNG")

    # 2.1 XÁC ĐỊNH TÁC NHÂN VÀ MÔ HÌNH PHÂN QUYỀN RBAC
    add_h1(doc, "2.1. XÁC ĐỊNH CÁC TÁC NHÂN VÀ MÔ HÌNH PHÂN QUYỀN (RBAC)")
    add_body(doc, "Dựa trên kết quả khảo sát thực tế cơ cấu tổ chức và quy trình phân định trách nhiệm trong công tác quản lý nghiên cứu khoa học tại Trường Đại học Công Thương TP. Hồ Chí Minh, hệ thống xác định 04 nhóm tác nhân người dùng (Actors) chính tham gia tương tác trực tiếp:")

    t21_headers = ["Tác nhân (Actor)", "Mô tả vai trò trong trường đại học", "Mục đích sử dụng chính"]
    t21_data = [
        ["Giảng viên / Nhà nghiên cứu (Lecturer)", "Cán bộ giảng dạy, nghiên cứu viên thuộc các Khoa, Viện, Trung tâm trực thuộc trường", "Khai báo hồ sơ công trình KH&CN, nhận gợi ý AI tức thì, tra cứu hồ sơ cá nhân và theo dõi trạng thái thẩm định"],
        ["Cán bộ Phòng QLKH (Science Officer)", "Chuyên viên chuyên trách tiếp nhận và thẩm định hồ sơ thuộc Phòng Quản lý Khoa học và Công nghệ", "Rà soát tính hợp lệ của hồ sơ, đối chiếu minh chứng, chuyển bước quy trình, phản hồi yêu cầu bổ sung và xem cảnh báo rủi ro AI"],
        ["Trưởng bộ môn / Ban Giám hiệu (Dean / Board)", "Lãnh đạo đơn vị chuyên môn và Ban Giám hiệu phụ trách công tác nghiên cứu khoa học", "Xem bảng điều khiển chỉ số KPI nghiên cứu toàn trường, phê duyệt công nhận nghiệm thu cuối cùng và xuất báo cáo thống kê"],
        ["Quản trị viên hệ thống (System Admin)", "Chuyên viên quản trị hạ tầng máy chủ và hệ thống phần mềm thuộc Trung tâm CNTT", "Quản lý tài khoản, phân quyền người dùng, quản trị 04 danh mục chuẩn hóa, cấu hình tham số AI và bảo trì cơ sở dữ liệu"]
    ]
    add_table_custom(doc, 2, "Danh sách các tác nhân người dùng tham gia tương tác trong hệ thống",
                     t21_headers, t21_data, col_widths=[4.0, 5.5, 6.5])

    add_body(doc, "Nhằm bảo đảm an toàn thông tin, bảo mật dữ liệu và ngăn chặn các truy cập trái phép vượt thẩm quyền, hệ thống áp dụng mô hình Kiểm soát truy cập dựa trên vai trò (Role-Based Access Control - RBAC). Ma trận phân quyền chi tiết được quy định như sau:")

    t22_headers = ["Phân hệ chức năng", "Giảng viên", "Cán bộ QLKH", "Lãnh đạo / BGH", "Quản trị viên"]
    t22_data = [
        ["Xem Tổng quan Dashboard & KPI nghiên cứu", "Xem số liệu chung", "Xem chi tiết", "Xem toàn quyền", "Xem toàn quyền"],
        ["Khai báo hồ sơ KH&CN mới (Live AI Preview)", "Toàn quyền thực hiện", "Toàn quyền thực hiện", "Không áp dụng", "Toàn quyền thực hiện"],
        ["Chỉnh sửa / Cập nhật hồ sơ của chính mình", "Chỉ khi 'Chờ duyệt' hoặc 'Cần bổ sung'", "Toàn quyền sửa đổi", "Chỉ xem", "Toàn quyền sửa đổi"],
        ["Thực hiện Quy trình xét duyệt 5 bước", "Không có quyền", "Toàn quyền thực hiện", "Phê duyệt bước 4 & 5", "Toàn quyền thực hiện"],
        ["Xem Cảnh báo bất thường & Điểm rủi ro AI", "Không xem cảnh báo nội bộ", "Xem chi tiết cảnh báo", "Xem chi tiết cảnh báo", "Xem toàn quyền"],
        ["Quản lý 04 Danh mục chuẩn hóa (Thêm/Sửa/Ẩn)", "Chỉ được tra cứu", "Thêm/Sửa danh mục", "Chỉ xem", "Toàn quyền CRUD"],
        ["Tra cứu đa tiêu chí & Xuất báo cáo Excel/PDF", "Tra cứu hồ sơ cá nhân", "Xuất báo cáo toàn trường", "Xuất báo cáo chiến lược", "Toàn quyền"],
        ["Cấu hình mô hình AI & Quản trị hệ thống", "Không có quyền", "Xem trạng thái AI", "Xem hiệu năng AI", "Toàn quyền quản trị"]
    ]
    add_table_custom(doc, 2, "Ma trận phân quyền chi tiết (RBAC) trên các phân hệ chức năng",
                     t22_headers, t22_data, col_widths=[4.5, 2.8, 2.8, 3.0, 2.9])

    # 2.2 ĐẶC TẢ YÊU CẦU CHỨC NĂNG
    add_h1(doc, "2.2. ĐẶC TẢ YÊU CẦU CHỨC NĂNG HỆ THỐNG")
    add_body(doc, "Căn cứ vào mục tiêu và yêu cầu nghiệp vụ của đề tài, hệ thống hiện thực hóa đầy đủ 06 phân hệ chức năng cốt lõi:")

    add_h2(doc, "2.2.1. Phân hệ 1: Quản lý hồ sơ KH&CN đa loại hình tích hợp Live AI Preview")
    add_bullet_1(doc, "Hỗ trợ khai báo toàn diện cả 03 nhóm loại công trình nghiên cứu: (1) Bài báo khoa học (tạp chí quốc tế Scopus/WoS, tạp chí trong nước HĐGSNN, kỷ yếu hội thảo); (2) Đề tài NCKH (cấp Cơ sở, cấp Bộ, cấp Tỉnh, cấp Nhà nước); (3) Sáng chế và Giải pháp hữu ích (đã được cấp bằng hoặc có đơn chấp nhận hợp lệ của Cục Sở hữu Trí tuệ).", bold_prefix="Hỗ trợ đa loại hình: ")
    add_bullet_1(doc, "Biểu mẫu nhập liệu động, tự động hiển thị hoặc ẩn các trường thông tin đặc thù tương ứng với loại công trình được chọn (ví dụ: chỉ số ISSN, mã DOI đối với bài báo; mã số hợp đồng, kinh phí cấp đối với đề tài; số văn bằng độc quyền, ngày nộp đơn đối với sáng chế).", bold_prefix="Biểu mẫu linh hoạt: ")
    add_bullet_1(doc, "Tích hợp tính năng Gợi ý trực tuyến thời gian thực (Live AI Preview): khi giảng viên gõ tiêu đề công trình hoặc nội dung tóm tắt, hệ thống tự động kích hoạt tiến trình nền phân tích và phản hồi ngay đề xuất Lĩnh vực nghiên cứu và Loại công trình có xác suất cao nhất kèm độ tin cậy phần trăm (%) chỉ sau 350ms, giúp giảng viên kê khai nhanh chóng và chính xác.", bold_prefix="Live AI Preview: ")
    add_bullet_1(doc, "Hỗ trợ đính kèm tệp tin tài liệu minh chứng (dạng PDF, Word, hoặc ảnh scan) với cơ chế kiểm tra định dạng và giới hạn dung lượng tải lên an toàn (tối đa 25MB/tệp).", bold_prefix="Minh chứng điện tử: ")

    add_h2(doc, "2.2.2. Phân hệ 2: Quản lý 04 danh mục dữ liệu chuẩn hóa")
    add_bullet_1(doc, "Danh mục Lĩnh vực nghiên cứu: Quản lý mã lĩnh vực, tên tiếng Việt, mô tả chuyên môn, trạng thái kích hoạt (Hoạt động / Ẩn).", bold_prefix="Danh mục 1: ")
    add_bullet_1(doc, "Danh mục Loại công trình: Quản lý tên nhóm loại hình, định mức giờ NCKH quy chuẩn, hệ số điểm HĐGSNN tương ứng.", bold_prefix="Danh mục 2: ")
    add_bullet_1(doc, "Danh mục Tạp chí / Hội thảo hợp lệ: Quản lý tên tạp chí, mã chuẩn quốc tế ISSN/eISSN, tổ chức xuất bản, xếp hạng Scopus (Q1-Q4), chỉ số WoS (IF), điểm số HĐGSNN và cờ cảnh báo tạp chí không uy tín / ngừng chỉ mục.", bold_prefix="Danh mục 3: ")
    add_bullet_1(doc, "Danh mục Tác giả / Giảng viên: Quản lý thông tin định danh nhà khoa học (Họ tên, Mã số cán bộ, Email nội bộ @huit.edu.vn, Học hàm, Học vị, Đơn vị Khoa/Viện trực thuộc).", bold_prefix="Danh mục 4: ")
    add_bullet_1(doc, "Toàn bộ 04 danh mục đều hỗ trợ đầy đủ các thao tác Thêm mới, Chỉnh sửa, và Ẩn/Hiện (Soft Delete - xóa mềm) nhằm bảo đảm toàn vẹn dữ liệu tham chiếu lịch sử.", bold_prefix="Quy chuẩn quản trị: ")

    add_h2(doc, "2.2.3. Phân hệ 3: Quy trình 5 bước xét duyệt tính hợp lệ và lưu vết lịch sử")
    add_body(doc, "Quy trình xét duyệt được thiết kế chuẩn hóa và chặt chẽ theo 5 bước tuần tự:")
    add_bullet_1(doc, "Bước 1: Tiếp nhận hồ sơ (Trạng thái: 'Chờ duyệt' - Pending) - Hồ sơ do giảng viên nộp được đưa vào hàng đợi tiếp nhận chung.", bold_prefix="Bước 1: ")
    add_bullet_1(doc, "Bước 2: Rà soát điều kiện hợp lệ sơ bộ (Trạng thái: 'Đang xử lý' - Processing) - Cán bộ QLKH đối chiếu hồ sơ với các quy định hành chính.", bold_prefix="Bước 2: ")
    add_bullet_1(doc, "Bước 3: Phản hồi và Bổ sung thông tin (Trạng thái: 'Cần bổ sung' - Revision Requested) - Nếu hồ sơ thiếu minh chứng, cán bộ gửi yêu cầu kèm ghi chú lý do rõ ràng; giảng viên cập nhật lại hồ sơ mà không bị hủy bỏ lượt nộp.", bold_prefix="Bước 3: ")
    add_bullet_1(doc, "Bước 4: Đánh giá hợp lệ / Không hợp lệ (Trạng thái: 'Đã duyệt' - Approved hoặc 'Từ chối' - Rejected) - Hội đồng chuyên môn xác nhận kết quả thẩm định.", bold_prefix="Bước 4: ")
    add_bullet_1(doc, "Bước 5: Hoàn tất nghiệm thu và Lưu trữ (Trạng thái: 'Đã nghiệm thu' - Archived/Finalized) - Hồ sơ chính thức được ghi nhận giờ chuẩn NCKH và đưa vào kho dữ liệu thành tích học thuật của Nhà trường.", bold_prefix="Bước 5: ")
    add_bullet_1(doc, "Hệ thống tự động ghi nhật ký lưu vết (Audit Log) bất biến cho từng hành động thay đổi trạng thái, ghi nhận chính xác thời gian (timestamp), người thực hiện (actor) và nội dung ghi chú.", bold_prefix="Lưu vết kiểm toán: ")

    add_h2(doc, "2.2.4. Phân hệ 4: Tra cứu, lọc đa tiêu chí và kết xuất báo cáo thống kê")
    add_bullet_1(doc, "Bộ lọc tìm kiếm mạnh mẽ: cho phép tra cứu kết hợp đồng thời theo Từ khóa (trong tên hoặc tóm tắt), Tác giả, Đơn vị công tác, Lĩnh vực chuyên ngành, Nhóm loại công trình, Năm công bố, Trạng thái thẩm định, và Mức độ cảnh báo rủi ro AI.", bold_prefix="Lọc đa tiêu chí: ")
    add_bullet_1(doc, "Kết xuất dữ liệu linh hoạt: hỗ trợ xuất danh sách hồ sơ lọc được ra định dạng bảng tính Microsoft Excel (.xlsx) chuẩn cấu trúc báo cáo của Bộ GD&ĐT và định dạng tài liệu PDF có định dạng chữ ký số thẩm định.", bold_prefix="Kết xuất báo cáo: ")

    add_h2(doc, "2.2.5. Phân hệ 5: Tích hợp mô hình AI phân loại tự động và Vòng lặp phản hồi")
    add_bullet_1(doc, "Phân loại tự động Lĩnh vực chuyên môn và Nhóm loại công trình bằng mô hình học sâu VinAI PhoBERT đạt F1-score 89.92% kèm mô hình đối chứng TF-IDF + LinearSVC (F1-score 82.26%).", bold_prefix="Động cơ AI phân loại: ")
    add_bullet_1(doc, "Giao diện Trung tâm AI cho phép chuyên viên chạy thử nghiệm so sánh trực tiếp hai thuật toán trên cùng một tập dữ liệu hoặc từng hồ sơ đơn lẻ, hiển thị trực quan biểu đồ độ tin cậy và ma trận nhầm lẫn.", bold_prefix="Trung tâm AI so sánh: ")
    add_bullet_1(doc, "Cơ chế Vòng lặp phản hồi tích cực (Feedback Loop): khi chuyên viên thẩm định điều chỉnh nhãn phân loại do AI gợi ý, hệ thống tự động ghi nhận mẫu dữ liệu hiệu chỉnh vào kho dữ liệu Active Learning để tái huấn luyện tinh chỉnh mô hình định kỳ.", bold_prefix="Feedback Loop: ")

    add_h2(doc, "2.2.6. Phân hệ 6: Phát hiện bất thường đa tầng và Hàng đợi xét duyệt ưu tiên")
    add_bullet_1(doc, "Tự động kích hoạt chuỗi kiểm tra đa tầng ngay khi hồ sơ được nộp: kiểm tra trùng lặp nội dung (Cosine Similarity), kiểm tra tạp chí săn mồi qua CSDL ISSN, phân tích mạng lưới đồng tác giả (NetworkX), và phát hiện dị biệt tham số qua Rừng cô lập (Isolation Forest) và PyTorch Deep Autoencoder.", bold_prefix="Kiểm tra đa tầng: ")
    add_bullet_1(doc, "Tính toán Điểm rủi ro tổng hợp (Composite Risk Score từ 0.00 đến 0.99) và xếp hạng hồ sơ trong Hàng đợi xét duyệt ưu tiên (Priority Review Queue), giúp chuyên viên nhận diện ngay các hồ sơ rủi ro cao để ưu tiên thẩm định trước.", bold_prefix="Hàng đợi ưu tiên: ")

    # 2.3 ĐẶC TẢ YÊU CẦU PHI CHỨC NĂNG
    add_h1(doc, "2.3. ĐẶC TẢ YÊU CẦU PHI CHỨC NĂNG")
    add_bullet_1(doc, "Thời gian tải trang ban đầu dưới 1.5 giây; thời gian phản hồi của các thao tác lọc, tìm kiếm và duyệt hồ sơ dưới 150ms; thời gian suy luận (Inference Time) của mô hình PhoBERT dưới 120ms/hồ sơ trên phần cứng tiêu chuẩn.", bold_prefix="Yêu cầu hiệu năng (Performance): ")
    add_bullet_1(doc, "Giao diện hoàn toàn bằng tiếng Việt, thiết kế thích ứng (Responsive) chuẩn Bootstrap 5, tối ưu hóa công thái học trải nghiệm người dùng, tương thích hoàn hảo trên các trình duyệt hiện đại (Chrome, Edge, Firefox, Safari).", bold_prefix="Tính khả dụng (Usability): ")
    add_bullet_1(doc, "Độ sẵn sàng hệ thống đạt tối thiểu 99.5%, cơ chế sao lưu cơ sở dữ liệu định kỳ tự động, tính toàn vẹn khóa ngoại được thực thi nghiêm ngặt trong SQLite 3 (PRAGMA foreign_keys = ON).", bold_prefix="Độ tin cậy (Reliability): ")
    add_bullet_1(doc, "Mã hóa mật khẩu bằng thuật toán băm an toàn Werkzeug PBKDF2-HMAC-SHA256; bảo vệ chống tấn công CSRF (Cross-Site Request Forgery) trên mọi biểu mẫu web; kiểm tra và lọc sạch dữ liệu đầu vào chống SQL Injection và XSS.", bold_prefix="Bảo mật và An toàn (Security): ")
    add_bullet_1(doc, "Kiến trúc mô-đun hóa cao, dễ dàng mở rộng quy mô dữ liệu lên hàng chục nghìn hồ sơ hoặc chuyển đổi hệ quản trị CSDL sang PostgreSQL/MySQL khi cần nâng cấp cấp cụm trường.", bold_prefix="Khả năng mở rộng (Scalability): ")

    # 2.4 BIỂU ĐỒ USE CASE VÀ ĐẶC TẢ
    add_h1(doc, "2.4. BIỂU ĐỒ USE CASE VÀ ĐẶC TẢ USE CASE CHÍNH")
    add_body(doc, "Dưới đây là Biểu đồ Use Case tổng thể mô tả mối quan hệ tương tác giữa 04 nhóm tác nhân người dùng với các phân hệ chức năng cốt lõi của hệ thống:")

    add_figure_custom(doc, "images/hinh_usecase.png", 2,
                      "Biểu đồ Use Case tổng thể hệ thống quản lý hồ sơ KH&CN", width_cm=14.5)

    add_h2(doc, "2.4.1. Đặc tả Use Case UC01: Khai báo hồ sơ KH&CN trực tuyến")
    add_bullet_1(doc, "Giảng viên / Nhà nghiên cứu (đã đăng nhập tài khoản hợp lệ).", bold_prefix="Tác nhân: ")
    add_bullet_1(doc, "Giảng viên chọn 'Khai báo hồ sơ mới' -> Chọn nhóm loại công trình -> Nhập tên đề tài, tóm tắt, từ khóa, tác giả, năm công bố -> Hệ thống kích hoạt Live AI Preview hiển thị gợi ý Lĩnh vực nghiên cứu kèm % tin cậy -> Giảng viên xác nhận hoặc điều chỉnh -> Đính kèm tệp PDF minh chứng -> Nhấn 'Nộp hồ sơ'.", bold_prefix="Luồng sự kiện chính: ")
    add_bullet_1(doc, "Nếu tệp đính kèm vượt quá 25MB hoặc sai định dạng, hệ thống hiển thị thông báo lỗi và yêu cầu chọn lại tệp tin hợp lệ.", bold_prefix="Luồng sự kiện ngoại lệ: ")
    add_bullet_1(doc, "Hồ sơ được tạo mới trong CSDL ở trạng thái 'Chờ duyệt', Điểm rủi ro ban đầu được tự động tính toán, tạo bản ghi Audit Log đầu tiên.", bold_prefix="Hậu điều kiện: ")

    add_h2(doc, "2.4.2. Đặc tả Use Case UC02: Xét duyệt hồ sơ theo quy trình 5 bước")
    add_bullet_1(doc, "Cán bộ Phòng QLKH, Lãnh đạo Đơn vị.", bold_prefix="Tác nhân: ")
    add_bullet_1(doc, "Cán bộ truy cập Danh sách hồ sơ cần duyệt -> Chọn hồ sơ ở đầu Hàng đợi ưu tiên -> Xem chi tiết thông tin, tệp minh chứng và bảng phân tích rủi ro AI -> Chọn hành động chuyển bước: 'Tiếp nhận' (Bước 2), 'Yêu cầu bổ sung' (Bước 3), hoặc 'Phê duyệt hợp lệ' (Bước 4) -> Nhập nội dung biên bản/ghi chú -> Nhấn 'Xác nhận'.", bold_prefix="Luồng sự kiện chính: ")
    add_bullet_1(doc, "Nếu chọn 'Yêu cầu bổ sung', hệ thống yêu cầu bắt buộc phải nhập lý do cụ thể gửi thông báo cho giảng viên.", bold_prefix="Luồng sự kiện nhánh: ")
    add_bullet_1(doc, "Trạng thái hồ sơ được cập nhật tức thì, hiển thị trên giao diện theo dõi tiến độ của giảng viên.", bold_prefix="Hậu điều kiện: ")

    add_h2(doc, "2.4.3. Đặc tả Use Case UC03: Phân tích AI và phát hiện cảnh báo rủi ro")
    add_bullet_1(doc, "Cán bộ Phòng QLKH, Quản trị viên hệ thống.", bold_prefix="Tác nhân: ")
    add_bullet_1(doc, "Hệ thống tự động kích hoạt tiến trình trích xuất đặc trưng văn bản, gọi đồng thời mô hình PhoBERT và pipeline phát hiện bất thường -> So khớp mã ISSN với danh mục tạp chí uy tín và tạp chí săn mồi -> Tính toán Cosine Similarity phát hiện trùng lặp -> Chạy Isolation Forest và Deep Autoencoder -> Tổng hợp Điểm rủi ro R_i -> Gắn nhãn màu cảnh báo và xuất chi tiết vi phạm lên màn hình thẩm định.", bold_prefix="Luồng sự kiện chính: ")

    # 2.5 QUY TRÌNH NGHIỆP VỤ 5 BƯỚC VÀ PHÂN TÍCH LỚP BCE
    add_h1(doc, "2.5. QUY TRÌNH NGHIỆP VỤ XÉT DUYỆT HỒ SƠ 5 BƯỚC VÀ PHÂN TÍCH LỚP BCE")
    add_body(doc, "Để bảo đảm tính minh bạch, khách quan và loại bỏ hoàn toàn các lỗ hổng tắc trách trong xét duyệt, hệ thống chuẩn hóa luồng công việc thành Quy trình 5 bước chặt chẽ, được mô hình hóa bằng sơ đồ hoạt động (Activity Diagram) dưới đây:")

    add_figure_custom(doc, "images/hinh1_quytrinh_5buoc.png", 2,
                      "Quy trình nghiệp vụ 5 bước xét duyệt tính hợp lệ của hồ sơ KH&CN", width_cm=14.5)

    add_body(doc, "Đặc điểm ưu việt nổi bật của quy trình là nhánh xử lý 'Yêu cầu bổ sung' tại Bước 3: khi phát hiện thiếu minh chứng hoặc thông tin kê khai chưa rõ ràng, chuyên viên không cần phải từ chối hủy hồ sơ (khiến giảng viên phải nhập lại từ đầu), mà có thể yêu cầu bổ sung kèm lời nhắn chi tiết, giúp tiết kiệm thời gian cho cả hai bên.")
    add_body(doc, "Nhằm phân rã rõ ràng ranh giới giữa giao diện người dùng, logic xử lý nghiệp vụ và thực thể dữ liệu, hệ thống áp dụng mẫu kiến trúc phân tích lớp Biên - Điều khiển - Thực thể (Boundary - Control - Entity - BCE). Sơ đồ BCE cho chức năng Xét duyệt hồ sơ và Phân tích AI được thể hiện trong hình sau:")

    add_figure_custom(doc, "images/hinh_bce.png", 2,
                      "Sơ đồ phân tích lớp Biên - Điều khiển - Thực thể (BCE) cho chức năng Xét duyệt và Phân tích AI", width_cm=14.5)

    add_body(doc, "Trong mô hình BCE:")
    add_bullet_1(doc, "Lớp Giao diện (Boundary): Bao gồm các màn hình tương tác người dùng như ProfileSubmissionForm, ReviewDetailView, AICenterPanel, tiếp nhận yêu cầu từ người dùng và trình bày dữ liệu phản hồi trực quan.", bold_prefix="Boundary: ")
    add_bullet_1(doc, "Lớp Điều khiển (Control): Gồm ProfileWorkflowController, AIInferenceEngine, AnomalyDetectionController, chịu trách nhiệm điều phối luồng nghiệp vụ, kiểm tra ràng buộc phân quyền, gọi thuật toán AI và tính toán điểm rủi ro.", bold_prefix="Control: ")
    add_bullet_1(doc, "Lớp Thực thể (Entity): Gồm các bảng thực thể cơ sở dữ liệu như ScientificProfile, CatalogItem, AuditLog, AnomalyReport, ánh xạ trực tiếp và lưu trữ trạng thái dữ liệu một cách nhất quán.", bold_prefix="Entity: ")

    return
