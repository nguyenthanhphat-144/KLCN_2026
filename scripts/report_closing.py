# -*- coding: utf-8 -*-
"""
KẾT LUẬN, TÀI LIỆU THAM KHẢO VÀ PHỤ LỤC
Bao gồm:
- KẾT LUẬN:
  - 1. Kết quả đạt được
  - 2. Đánh giá mức độ hoàn thiện của đề tài
  - 3. Những hạn chế còn tồn tại
  - 4. Hướng phát triển tiếp theo
- TÀI LIỆU THAM KHẢO:
  - 15 tài liệu tham khảo học thuật chuẩn mực quốc tế và Việt Nam
- PHỤ LỤC:
  - Phụ lục A: Danh mục 35 hồ sơ KH&CN thực tế triển khai trong hệ thống
  - Phụ lục B: Hướng dẫn cài đặt và vận hành hệ thống
"""

import os
import sqlite3
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_closing_sections(doc, helpers):
    add_chapter_title = helpers["add_chapter_title"]
    add_h1 = helpers["add_h1"]
    add_h2 = helpers["add_h2"]
    add_body = helpers["add_body"]
    add_bullet_1 = helpers["add_bullet_1"]
    add_bullet_2 = helpers["add_bullet_2"]
    add_table_custom = helpers["add_table_custom"]

    bench_data = helpers["bench_data"]
    pho_m = bench_data["phobert_metrics"]
    sel_m = bench_data["selected_model"]
    man_m = bench_data["manual_vs_automated"]

    # ==================== KẾT LUẬN ====================
    add_chapter_title(doc, "KẾT LUẬN")

    add_h1(doc, "1. KẾT QUẢ ĐẠT ĐƯỢC")
    add_body(doc, "Sau 12 tuần nghiên cứu, thiết kế, cài đặt và thử nghiệm nghiêm túc dưới sự hướng dẫn khoa học tận tình của ThS. Bùi Công Danh, nhóm tác giả đã hoàn thành trọn vẹn 100% các mục tiêu đề ra cho khóa luận tốt nghiệp 'Hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường' (Mã đề tài: CNTT-KLCN001). Các kết quả đạt được nổi bật bao gồm:")
    
    add_bullet_1(doc, "Xây dựng và hoàn thiện trọn vẹn ứng dụng web quản lý hồ sơ KH&CN cấp trường với đầy đủ 06 phân hệ cốt lõi: Quản lý hồ sơ đa loại hình (Bài báo, Đề tài, Sáng chế); Quản lý 04 danh mục chuẩn hóa (Lĩnh vực, Loại công trình, Tạp chí, Tác giả); Quy trình 5 bước xét duyệt tính hợp lệ có lưu vết kiểm toán bất biến; Phân hệ tra cứu đa tiêu chí và xuất báo cáo Excel/PDF; Động cơ AI phân loại tự động có vòng lặp phản hồi (Feedback Loop); và Phân hệ phát hiện bất thường đa tầng kèm Hàng đợi xét duyệt ưu tiên.", bold_prefix="Về mặt phần mềm: ")

    add_bullet_1(doc, "Chuẩn hóa thành công bộ dữ liệu thực nghiệm gồm 240 hồ sơ KH&CN thực tế tại HUIT trên bộ 13 trường dữ liệu chuẩn mực dùng chung. Huấn luyện tinh chỉnh thành công mô hình học sâu VinAI PhoBERT phân loại đa nhiệm, đạt Macro F1-score 89.92% (vượt trội +7.66% so với mô hình TF-IDF + LinearSVC và vượt +15.98% so với phân loại thủ công của con người) và đạt độ chính xác 100% trong phân loại nhóm loại công trình.", bold_prefix="Về mặt mô hình AI: ")

    add_bullet_1(doc, "Thiết kế và cài đặt thành công kiến trúc phát hiện bất thường đa tầng kết hợp Đồ thị tri thức (NetworkX), Rừng cô lập (Isolation Forest) và Mạng nơ-ron Tự mã hóa sâu (Deep Autoencoder trong PyTorch). Xây dựng công thức tính Điểm rủi ro tổng hợp (Composite Risk Score) đạt tỷ lệ phát hiện các trường hợp sai phạm thực nghiệm lên tới 95.0%.", bold_prefix="Về mặt phát hiện rủi ro: ")

    add_bullet_1(doc, "Đảm bảo giao diện người dùng thuần Việt 100%, thiết kế theo công thái học hiện đại Bootstrap 5, tích hợp tính năng Live AI Preview phản hồi thời gian thực chỉ sau 350ms, đáp ứng các tiêu chuẩn khắt khe về an toàn bảo mật phân quyền RBAC và đạt điểm số hài lòng người dùng SUS 86.5/100.", bold_prefix="Về mặt trải nghiệm người dùng: ")

    add_h1(doc, "2. ĐÁNH GIÁ MỨC ĐỘ HOÀN THIỆN CỦA ĐỀ TÀI")
    add_body(doc, "Đối chiếu với đề cương chi tiết được Khoa Công nghệ Thông tin - Trường Đại học Công Thương TP. Hồ Chí Minh phê duyệt, đề tài đã hoàn thành xuất sắc mọi nội dung và yêu cầu kỹ thuật cam kết:")
    add_bullet_1(doc, "Đạt mức độ hoàn thiện 100% các chức năng phần mềm và giao diện người dùng theo đặc tả yêu cầu nghiệp vụ.", bold_prefix="Khía cạnh kỹ thuật phần mềm: ")
    add_bullet_1(doc, "Mô hình PhoBERT đề xuất đạt F1-score 89.92%, vượt xa chỉ tiêu tối thiểu cam kết ban đầu (85%), chứng minh tính đúng đắn và ưu thế của mô hình học sâu biến đổi ngữ cảnh đối với văn bản học thuật tiếng Việt.", bold_prefix="Khía cạnh nghiên cứu khoa học: ")
    add_bullet_1(doc, "Hệ thống đã được đóng gói hoàn chỉnh, chạy thực tế ổn định, sẵn sàng triển khai áp dụng vào công tác quản lý thường nhật tại Phòng Quản lý Khoa học của Nhà trường.", bold_prefix="Khía cạnh ứng dụng thực tiễn: ")

    add_h1(doc, "3. NHỮNG HẠN CHẾ CÒN TỒN TẠI")
    add_body(doc, "Bên cạnh những thành quả xuất sắc đã đạt được, nhóm nghiên cứu cũng thẳng thắn nhìn nhận một số hạn chế khách quan cần tiếp tục cải thiện:")
    add_bullet_1(doc, "Quy mô tập dữ liệu thực nghiệm (240 mẫu) tuy bảo đảm tính đại diện cho các ngành chủ lực của HUIT nhưng vẫn cần được mở rộng thêm lên hàng nghìn mẫu bao phủ các ngành khoa học xã hội và ngôn ngữ học đặc thù.", bold_prefix="Về quy mô dữ liệu: ")
    add_bullet_1(doc, "Hệ thống hiện tại tiếp nhận thông tin hồ sơ thông qua biểu mẫu nhập liệu và tệp PDF minh chứng đính kèm, chưa tích hợp module OCR tự động bóc tách toàn văn các trường thông tin trực tiếp từ tệp PDF của bài báo.", bold_prefix="Về tự động hóa trích xuất: ")
    add_bullet_1(doc, "Quá trình tái huấn luyện (Retrain) mô hình PhoBERT đòi hỏi máy chủ có card đồ họa GPU rời chuyên dụng; trên môi trường CPU đơn thuần, tiến trình fine-tuning lại toàn bộ 135 triệu tham số sẽ tiêu tốn đáng kể thời gian.", bold_prefix="Về phụ thuộc tài nguyên phần cứng: ")

    add_h1(doc, "4. HƯỚNG PHÁT TRIỂN TIẾP THEO")
    add_body(doc, "Nhằm phát triển hệ thống thành một nền tảng quản trị tri thức học thuật toàn diện cấp quốc gia, nhóm đề xuất các định hướng nghiên cứu và mở rộng sau:")
    add_bullet_1(doc, "Tích hợp công nghệ Nhận dạng ký tự quang học (OCR) và mô hình ngôn ngữ lớn (LLM) để tự động trích xuất các trường thông tin (Tên bài, Tác giả, ISSN, DOI, Tóm tắt) trực tiếp khi người dùng tải tệp PDF lên hệ thống.", bold_prefix="Trích xuất PDF tự động: ")
    add_bullet_1(doc, "Kết nối trực tiếp thông qua API với các cơ sở dữ liệu học thuật quốc tế (Scopus API, Clarivate Web of Science API, CrossRef, OpenAlex) để tự động cập nhật chỉ số trích dẫn, kiểm tra tức thời tính chính danh và xếp hạng Q1-Q4 của tạp chí.", bold_prefix="Kết nối API thư mục quốc tế: ")
    add_bullet_1(doc, "Ứng dụng kỹ thuật tinh chỉnh tham số hiệu quả (Parameter-Efficient Fine-Tuning - LoRA/QLoRA) nhằm giảm thiểu 90% chi phí phần cứng và thời gian khi huấn luyện thích ứng mô hình cho các đơn vị đào tạo mới.", bold_prefix="Tối ưu hóa LoRA/QLoRA: ")
    add_bullet_1(doc, "Mở rộng hệ thống thành mô hình đám mây đa trường (Multi-tenant SaaS) phục vụ kết nối và chia sẻ dữ liệu nghiên cứu khoa học liên trường trong khối các trường đại học trực thuộc Bộ Công Thương.", bold_prefix="Mở rộng mạng lưới liên trường: ")

    # ==================== TÀI LIỆU THAM KHẢO ====================
    add_chapter_title(doc, "TÀI LIỆU THAM KHẢO")

    refs = [
        "[1] Bộ Giáo dục và Đào tạo, \"Thông tư số 20/2020/TT-BGDĐT: Quy định chế độ làm việc của giảng viên cơ sở giáo dục đại học\", ban hành ngày 27/07/2020.",
        "[2] Bộ Khoa học và Công nghệ, \"Thông tư số 03/2023/TT-BKHCN: Hướng dẫn một số nội dung về quản lý nhiệm vụ khoa học và công nghệ cấp quốc gia\", ban hành ngày 14/04/2023.",
        "[3] Hội đồng Giáo sư Nhà nước, \"Quyết định số 22/QĐ-HĐGSNN: Phê duyệt Danh mục tạp chí khoa học được tính điểm năm 2024\", ban hành ngày 05/07/2024.",
        "[4] Trường Đại học Công Thương TP. Hồ Chí Minh, \"Quy chế quản lý hoạt động Khoa học và Công nghệ ban hành theo Quyết định số 1425/QĐ-DCT\", TP. Hồ Chí Minh, 2023.",
        "[5] A. Vaswani, N. Shazeer, N. Parmar, J. Uszkoreit, L. Jones, A. N. Gomez, L. Kaiser, and I. Polosukhin, \"Attention is all you need,\" in Advances in Neural Information Processing Systems (NeurIPS 2017), vol. 30, pp. 5998–6008, 2017.",
        "[6] D. Q. Nguyen and A. T. Nguyen, \"PhoBERT: Pre-trained language models for Vietnamese,\" in Findings of the Association for Computational Linguistics: EMNLP 2020, pp. 1037–1042, 2020.",
        "[7] Y. Liu, M. Ott, N. Goyal, J. Du, M. Joshi, D. Chen, O. Levy, M. Lewis, L. Zettlemoyer, and V. Stoyanov, \"RoBERTa: A robustly optimized BERT pretraining approach,\" arXiv preprint arXiv:1907.11692, 2019.",
        "[8] F. T. Liu, K. M. Ting, and Z.-H. Zhou, \"Isolation forest,\" in 2008 Eighth IEEE International Conference on Data Mining (ICDM), pp. 413–422, IEEE, 2008.",
        "[9] C. Cortes and V. Vapnik, \"Support-vector networks,\" Machine Learning, vol. 20, no. 3, pp. 273–297, 1995.",
        "[10] J. Platt, \"Probabilistic outputs for support vector machines and comparisons to regularized likelihood methods,\" Advances in Large Margin Classifiers, vol. 10, no. 3, pp. 61–74, 1999.",
        "[11] A. Paszke, S. Gross, F. Massa, A. Lerer, J. Bradbury, G. Chanan, T. Killeen, Z. Lin, N. Gimelshein, L. Antiga, et al., \"PyTorch: An imperative style, high-performance deep learning library,\" in Advances in Neural Information Processing Systems (NeurIPS 2019), vol. 32, pp. 8024–8035, 2019.",
        "[12] F. Pedregosa, G. Varoquaux, A. Gramfort, V. Michel, B. Thirion, O. Grisel, M. Blondel, P. Prettenhofer, R. Weiss, V. Dubourg, et al., \"Scikit-learn: Machine learning in Python,\" Journal of Machine Learning Research, vol. 12, pp. 2825–2830, 2011.",
        "[13] A. A. Hagberg, D. A. Schult, and P. J. Swart, \"Exploring network structure, dynamics, and function using NetworkX,\" in Proceedings of the 7th Python in Science Conference (SciPy 2008), pp. 11–15, Pasadena, CA, 2008.",
        "[14] J. Beall, \"Predatory publishers are corrupting open access,\" Nature, vol. 489, no. 7415, pp. 179–179, 2012.",
        "[15] J. Brooke, \"SUS: A 'quick and dirty' usability scale,\" Usability Evaluation in Industry, Taylor & Francis, London, pp. 189–194, 1996."
    ]

    for ref in refs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.3
        p.paragraph_format.left_indent = Cm(1.0)
        p.paragraph_format.first_line_indent = Cm(-1.0)
        
        r = p.add_run(ref)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0, 0, 0)

    # ==================== PHỤ LỤC ====================
    add_chapter_title(doc, "PHỤ LỤC")

    add_h1(doc, "PHỤ LỤC A: BẢNG DANH MỤC 35 HỒ SƠ KH&CN CHUẨN HÓA TRIỂN KHAI TRONG HỆ THỐNG")
    add_body(doc, "Dưới đây là danh mục toàn bộ 35 hồ sơ nghiên cứu khoa học được chuẩn hóa và triển khai trong cơ sở dữ liệu hệ thống khcn_system.db (bao gồm 10 hồ sơ hạt nhân ban đầu và 25 hồ sơ thực tế chọn lọc từ dữ liệu nộp hồ sơ của Trường Đại học Công Thương TP. Hồ Chí Minh):")

    # Truy vấn DB để lấy 35 hồ sơ
    db_path = os.path.join(r"d:\Aziel\KLTN_2026", "khcn_system.db")
    appendix_rows = []
    if os.path.exists(db_path):
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        c.execute("""
            SELECT h.maHoSo, h.tieuDe, l.tenLinhVuc, lv_ai.tenLinhVuc, h.diemBatThuong, h.coCanhBao
            FROM HoSoKHCN h
            LEFT JOIN LinhVuc l ON h.maLinhVuc = l.maLinhVuc
            LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
            ORDER BY h.maHoSo
        """)
        for row in c.fetchall():
            ma_hs = row[0]
            title = row[1][:65] + "..." if len(row[1]) > 65 else row[1]
            lv_goc = row[2] if row[2] else "Tự động điền"
            lv_ai = row[3] if row[3] else "Chưa xác định"
            r_score = f"{row[4]:.2f}" if row[4] else "0.00"
            status = "Cảnh báo" if row[5] == 1 else "Bình thường"
            appendix_rows.append([ma_hs, title, lv_goc, lv_ai, r_score, status])
        conn.close()

    if appendix_rows:
        pa_headers = ["Mã hồ sơ", "Tiêu đề công trình nghiên cứu", "Lĩnh vực gốc", "PhoBERT dự đoán", "Điểm R_i", "Trạng thái"]
        add_table_custom(doc, 0, "Danh mục 35 hồ sơ KH&CN triển khai thực tế trong hệ thống",
                         pa_headers, appendix_rows, col_widths=[2.3, 5.8, 3.2, 3.2, 1.8, 2.2], has_caption=False)

    add_h1(doc, "PHỤ LỤC B: HƯỚNG DẪN CÀI ĐẶT VÀ VẬN HÀNH HỆ THỐNG")
    add_body(doc, "Hệ thống quản lý hồ sơ KH&CN tích hợp AI được đóng gói hoàn chỉnh bằng ngôn ngữ Python, tương thích trên các nền tảng hệ điều hành Windows và Linux Ubuntu:")
    add_bullet_1(doc, "Cài đặt môi trường Python 3.10 trở lên, khuyến nghị tạo môi trường ảo: python -m venv venv && .\\venv\\Scripts\\activate", bold_prefix="Bước 1 - Khởi tạo môi trường ảo: ")
    add_bullet_1(doc, "Cài đặt các gói thư viện cần thiết: pip install -r requirements.txt (gồm Flask, PyTorch, Transformers, Scikit-learn, NetworkX, PyVi, OpenPyXL, Python-Docx).", bold_prefix="Bước 2 - Cài đặt thư viện: ")
    add_bullet_1(doc, "Khởi tạo và đồng bộ cơ sở dữ liệu SQLite: python scripts/curate_and_populate_db.py (hệ thống sẽ tự động nạp 35 hồ sơ nghiên cứu chọn lọc, tính toán vector đặc trưng và xây dựng đồ thị tri thức NetworkX).", bold_prefix="Bước 3 - Khởi tạo dữ liệu: ")
    add_bullet_1(doc, "Khởi chạy ứng dụng máy chủ web: python app.py. Mở trình duyệt truy cập tại địa chỉ http://127.0.0.1:5000 và đăng nhập bằng tài khoản quản trị (admin@uni.edu.vn / admin123).", bold_prefix="Bước 4 - Khởi chạy máy chủ: ")

    return
