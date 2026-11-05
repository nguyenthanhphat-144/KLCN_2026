# -*- coding: utf-8 -*-
"""
Các chương nội dung chính cho Báo cáo Khóa luận Tốt nghiệp HUIT: Bao_Cao_Khoa_Luan.docx
Bao gồm:
- MỞ ĐẦU
- CHƯƠNG 1: KHẢO SÁT HỆ THỐNG VÀ CƠ SỞ LÝ THUYẾT
- CHƯƠNG 2: PHÂN TÍCH HỆ THỐNG
- CHƯƠNG 3: THIẾT KẾ HỆ THỐNG
- CHƯƠNG 4: CÀI ĐẶT HỆ THỐNG
- CHƯƠNG 5: THỬ NGHIỆM VÀ TRIỂN KHAI
- KẾT LUẬN
- TÀI LIỆU THAM KHẢO
"""

import os
import sys
import json
from docx.shared import Pt, Cm, RGBColor

def build_all_chapters(doc, helpers):
    # Trích xuất các helper functions từ dict helpers
    add_chapter_title = helpers["add_chapter_title"]
    add_h1 = helpers["add_h1"]
    add_h2 = helpers["add_h2"]
    add_h3 = helpers["add_h3"]
    add_body = helpers["add_body"]
    add_bullet_1 = helpers["add_bullet_1"]
    add_bullet_2 = helpers["add_bullet_2"]
    add_figure_custom = helpers["add_figure_custom"]
    add_table_custom = helpers["add_table_custom"]
    
    # Các hàm công thức OMML
    add_omml_tfidf = helpers["add_omml_tfidf"]
    add_omml_linearsvc = helpers["add_omml_linearsvc"]
    add_omml_platt_scaling = helpers["add_omml_platt_scaling"]
    add_omml_scaled_attention = helpers["add_omml_scaled_attention"]
    add_omml_multihead_attention = helpers["add_omml_multihead_attention"]
    add_omml_joint_loss = helpers["add_omml_joint_loss"]
    add_omml_degree_centrality = helpers["add_omml_degree_centrality"]
    add_omml_isolation_forest = helpers["add_omml_isolation_forest"]
    add_omml_pytorch_autoencoder = helpers["add_omml_pytorch_autoencoder"]
    add_omml_cosine_similarity = helpers["add_omml_cosine_similarity"]
    add_omml_risk_score = helpers["add_omml_risk_score"]
    add_omml_classification_metrics = helpers["add_omml_classification_metrics"]

    # Dữ liệu thực nghiệm
    bench_data = helpers["bench_data"]
    pho_m = bench_data["phobert_metrics"]
    svm_m = bench_data["svm_metrics"]
    sel_m = bench_data["selected_model"]
    man_m = bench_data["manual_vs_automated"]
    ds_m = bench_data["dataset_stats"]
    per_f = bench_data["per_field_comparison"]

    # =========================================================================
    # MỞ ĐẦU
    # =========================================================================
    add_chapter_title(doc, "MỞ ĐẦU")

    add_h1(doc, "1. GIỚI THIỆU")
    add_body(doc, "Trong kỷ nguyên của cuộc Cách mạng Công nghiệp lần thứ tư và làn sóng chuyển đổi số sâu rộng trong giáo dục đại học, hoạt động Khoa học và Công nghệ (KH&CN) đã và đang trở thành một trong những trụ cột chiến lược cốt lõi, quyết định uy tín học thuật, chất lượng đào tạo và vị thế cạnh tranh của mỗi cơ sở giáo dục đại học. Theo các văn bản chỉ đạo của Bộ Giáo dục và Đào tạo, đặc biệt là Thông tư số 20/2020/TT-BGDĐT quy định chế độ làm việc của giảng viên cơ sở giáo dục đại học, hoạt động nghiên cứu khoa học là nghĩa vụ bắt buộc, được lượng hóa cụ thể thành định mức giờ chuẩn nghiên cứu khoa học hàng năm cho từng chức danh giảng viên. Đồng thời, khối lượng công trình công bố quốc tế (thuộc danh mục Scopus, Web of Science), bài báo trên các tạp chí khoa học chuyên ngành trong nước được Hội đồng Giáo sư Nhà nước tính điểm, cùng các đề tài NCKH các cấp và bằng độc quyền sáng chế / giải pháp hữu ích là những tiêu chí then chốt phục vụ công tác kiểm định chất lượng cơ sở giáo dục đại học theo tiêu chuẩn quốc gia và quốc tế (AUN-QA, QS, THE).")
    add_body(doc, "Tại Trường Đại học Công Thương TP. Hồ Chí Minh (HUIT) cũng như nhiều cơ sở đào tạo đại học lớn trong cả nước, số lượng hồ sơ kê khai hoạt động KH&CN của cán bộ, giảng viên có sự gia tăng nhanh chóng về cả quy mô và tính đa dạng qua từng năm học. Mỗi đợt quyết toán kinh phí hoặc bình xét thi đua cuối học kỳ, Phòng Quản lý Khoa học phải tiếp nhận, phân loại và thẩm định hàng trăm đến hàng ngàn hồ sơ thuộc nhiều lĩnh vực chuyên môn rất khác nhau. Tuy nhiên, quy trình xử lý hiện nay tại đa số các trường vẫn mang tính thủ công hoặc bán tự động dựa trên các biểu mẫu bảng tính phân tán (Microsoft Excel, Google Sheets) và hồ sơ minh chứng rời rạc qua email. Thực trạng này dẫn đến những nút thắt và rủi ro nghiêm trọng:")
    add_bullet_1(doc, "Thời gian rà soát, thẩm định kéo dài: Cán bộ chuyên trách phải đọc từng hồ sơ, tra cứu thủ công danh mục tạp chí, rà soát từng quyết định giao đề tài và đối chiếu số lượng tác giả, gây quá tải công việc vào các thời điểm cao điểm quyết toán.", bold_prefix="Nút thắt về thời gian: ")
    add_bullet_1(doc, "Nguy cơ phân loại sai lệch lĩnh vực chuyên ngành: Các nghiên cứu có tính liên ngành cao (Interdisciplinary) thường chứa các từ khóa đan xen giữa lĩnh vực ứng dụng và phương pháp cốt lõi, khiến việc phân loại thủ công dễ bị nhầm lẫn, dẫn đến sai lệch trong việc phân công hội đồng thẩm định chuyên môn phù hợp.", bold_prefix="Nút thắt về phân loại: ")
    add_bullet_1(doc, "Khó khăn trong kiểm soát liêm chính học thuật và phát hiện bất thường: Việc rà soát bằng mắt thường gần như bất khả thi trong việc phát hiện các hành vi kê khai trùng lặp công trình giữa các tác giả/đơn vị, công trình đăng tải trên các tạp chí 'săn mồi' (predatory journals) hoặc các tạp chí đã bị hệ thống Scopus/WoS ngừng chỉ mục (discontinued), cũng như các trường hợp kê khai vượt quá định mức giờ nghiên cứu thực tế.", bold_prefix="Nút thắt về kiểm soát rủi ro: ")
    add_body(doc, "Xuất phát từ những đòi hỏi cấp thiết của thực tiễn quản lý và tiềm năng ứng dụng vượt bậc của Trí tuệ Nhân tạo (AI), đề tài \"Hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường\" (Mã đề tài: CNTT-KLCN001) được hình thành và thực hiện. Khóa luận hướng đến việc xây dựng một hệ sinh thái phần mềm hoàn chỉnh, ứng dụng mô hình học sâu biến đổi ngữ cảnh tiếng Việt VinAI PhoBERT kết hợp các thuật toán học máy phân tích dị biệt (Isolation Forest, Deep Autoencoder, NetworkX) nhằm chuẩn hóa quy trình, tự động hóa phân loại và thiết lập cơ chế giám sát rủi ro thông minh, minh bạch cho các trường đại học.")

    add_h1(doc, "2. MỤC TIÊU ĐỀ TÀI")
    add_body(doc, "Mục tiêu tổng quát của khóa luận là nghiên cứu, thiết kế và phát triển hoàn chỉnh một hệ thống ứng dụng web chuyên nghiệp hỗ trợ quản lý hồ sơ KH&CN cấp trường, trong đó tích hợp sâu các mô hình Trí tuệ Nhân tạo hiện đại nhằm tự động hóa tối đa quy trình phân loại chuyên ngành và nâng cao hiệu quả phát hiện các sai lệch, rủi ro học thuật.")
    add_body(doc, "Để đạt được mục tiêu tổng quát nêu trên, đề tài xác định các mục tiêu cụ thể sau:")
    add_bullet_1(doc, "Hiện thực hóa toàn diện 06 phân hệ chức năng nghiệp vụ cốt lõi theo đúng đề cương: (1) Quản lý hồ sơ KH&CN đa loại hình; (2) Quản lý 04 danh mục chuẩn hóa; (3) Quy trình 5 bước xét duyệt tính hợp lệ; (4) Tra cứu, lọc đa tiêu chí và xuất báo cáo Excel/PDF; (5) Tích hợp mô hình AI phân loại tự động; (6) Phân hệ phát hiện bất thường và thiết lập Hàng đợi xét duyệt ưu tiên.", bold_prefix="Mục tiêu phần mềm: ")
    add_bullet_1(doc, "Nghiên cứu và xây dựng pipeline tiền xử lý ngôn ngữ tiếng Việt học thuật, chuẩn hóa thống nhất bộ 13 trường thông tin đầu vào dùng chung cho các mô hình AI.", bold_prefix="Mục tiêu xử lý dữ liệu: ")
    add_bullet_1(doc, "Huấn luyện tinh chỉnh (Fine-tuning) mô hình học sâu Transformer tiếng Việt VinAI PhoBERT (vinai/phobert-base) phân loại đồng thời Lĩnh vực chuyên ngành và Nhóm loại công trình; xây dựng mô hình đối chứng truyền thống TF-IDF + LinearSVC; đánh giá khách quan trên cùng tập kiểm thử độc lập 60 mẫu nhằm đạt F1-score trên 88%.", bold_prefix="Mục tiêu phân loại AI: ")
    add_bullet_1(doc, "Thiết kế và hiện thực giải thuật phát hiện bất thường đa tầng kết hợp Đồ thị tri thức đồng tác giả NetworkX, Rừng cây cô lập (Isolation Forest) và Mạng nơ-ron Tự mã hóa PyTorch Autoencoder (torch.nn.Module); thiết lập công thức tính Điểm rủi ro tổng hợp (Composite Risk Score) để cảnh báo tức thời các hồ sơ có nguy cơ vi phạm.", bold_prefix="Mục tiêu phát hiện bất thường: ")
    add_bullet_1(doc, "Đảm bảo giao diện người dùng thuần Việt 100%, thao tác trực quan, tốc độ phản hồi nhanh (< 150ms), hỗ trợ cơ chế gợi ý trực tuyến Live AI Preview và đảm bảo an toàn phân quyền người dùng (RBAC).", bold_prefix="Mục tiêu trải nghiệm người dùng: ")

    add_h1(doc, "3. ĐỐI TƯỢNG VÀ PHẠM VI ĐỀ TÀI")
    add_body(doc, "Đối tượng nghiên cứu của đề tài bao gồm:", bold_prefix="Đối tượng nghiên cứu: ")
    add_bullet_1(doc, "Nghiệp vụ quản lý, tiếp nhận, xét duyệt, phân loại và lưu vết thẩm định hồ sơ KH&CN trong môi trường giáo dục đại học theo quy định của Bộ GD&ĐT và Trường Đại học Công Thương TP. HCM.")
    add_bullet_1(doc, "Các thuật toán Xử lý ngôn ngữ tự nhiên (NLP) cho văn bản tiếng Việt: trích xuất đặc trưng TF-IDF, thuật toán phân loại tuyến tính Support Vector Machine (LinearSVC), kỹ thuật hiệu chuẩn xác suất Platt Scaling, và mô hình mạng nơ-ron Transformer học sâu VinAI PhoBERT.")
    add_bullet_1(doc, "Các giải thuật phát hiện dị biệt và khai phá mạng lưới xã hội: Phân tích đồ thị đồng tác giả (NetworkX Degree Centrality), Rừng cây cô lập (Isolation Forest), Mạng nơ-ron Tự mã hóa nén tái tạo đặc trưng (PyTorch Autoencoder), và phép đo độ tương đồng góc Cosine.")
    add_body(doc, "Phạm vi nghiên cứu của đề tài được giới hạn như sau:", bold_prefix="Phạm vi nghiên cứu: ")
    add_bullet_1(doc, "Phạm vi dữ liệu: Tập dữ liệu thực nghiệm gồm 240 hồ sơ KH&CN được thu thập, chuẩn hóa từ các công trình nghiên cứu thực tế tại Trường Đại học Công Thương TP. HCM (180 mẫu huấn luyện, 60 mẫu kiểm thử độc lập). Dữ liệu bao quát cả 03 nhóm công trình (Bài báo khoa học, Đề tài NCKH, Bằng sáng chế/GPHI) và 05 lĩnh vực chuyên môn trọng điểm: Công nghệ Thông tin, Công nghệ Thực phẩm, Kỹ thuật Hóa học, Kinh tế & Quản trị Kinh doanh, và Khoa học Môi trường.")
    add_bullet_1(doc, "Phạm vi công nghệ: Hệ thống web được phát triển bằng ngôn ngữ Python, vi khung Flask, cơ sở dữ liệu quan hệ chuẩn hóa SQLite 3, giao diện Bootstrap 5 thuần Việt. Module AI được lập trình trên PyTorch 2.x, HuggingFace Transformers, PyVi, Scikit-learn và NetworkX.")
    add_bullet_1(doc, "Phạm vi triển khai: Ứng dụng đóng vai trò là hệ thống hỗ trợ ra quyết định (Decision Support System) nội bộ dành cho Ban Quản lý Khoa học, Giảng viên và Hội đồng xét duyệt tại các trường đại học.")

    add_h1(doc, "4. PHƯƠNG PHÁP NGHIÊN CỨU")
    add_body(doc, "Khóa luận kết hợp đồng bộ các phương pháp nghiên cứu khoa học sau:")
    add_bullet_1(doc, "Nghiên cứu các quy chế, biểu mẫu quản lý KH&CN hiện hành; tổng hợp các công trình nghiên cứu trong và ngoài nước về bài toán phân loại văn bản học thuật và phát hiện gian lận nghiên cứu.", bold_prefix="Phương pháp nghiên cứu lý thuyết: ")
    add_bullet_1(doc, "Khảo sát thực tế quy trình tiếp nhận, luân chuyển và thẩm định hồ sơ tại Phòng Quản lý Khoa học HUIT; áp dụng phương pháp phân tích hướng đối tượng (OOAD) với chuẩn UML (Use Case, BCE, Activity Diagram) để mô hình hóa toàn diện các yêu cầu nghiệp vụ.", bold_prefix="Phương pháp khảo sát và phân tích hệ thống: ")
    add_bullet_1(doc, "Tiến hành tiền xử lý, làm sạch và gán nhãn tập dữ liệu 240 hồ sơ; thực hiện huấn luyện và kiểm thử chéo hai mô hình trên cùng tập dữ liệu test độc lập 60 mẫu với cùng 13 trường dữ liệu; đo lường chính xác các chỉ số Precision, Recall, Macro F1-score và thời gian đáp ứng để bảo đảm tính khách quan, trung thực.", bold_prefix="Phương pháp thực nghiệm và đánh giá: ")
    add_bullet_1(doc, "Áp dụng mô hình Agile/Scrum theo chu kỳ 12 tuần, chia nhỏ các giai đoạn phát triển, tích hợp liên tục và kiểm thử phần mềm ở từng cấp độ (Unit Test, Integration Test, UI/UX Test).", bold_prefix="Phương pháp xây dựng phần mềm: ")

    add_h1(doc, "5. Ý NGHĨA KHOA HỌC VÀ THỰC TIỄN")
    add_body(doc, "Về mặt khoa học, đề tài chứng minh tính khả thi và ưu thế vượt trội của việc áp dụng mô hình Transformer tiền huấn luyện bản ngữ tiếng Việt (VinAI PhoBERT) trong việc giải quyết bài toán phân loại văn bản học thuật đa ngành. Kết quả thực nghiệm khẳng định cơ chế Self-Attention hai chiều có khả năng nắm bắt chính xác quan hệ cú pháp giữa phương pháp nghiên cứu cốt lõi và lĩnh vực ứng dụng, giải quyết triệt để nhược điểm của mô hình túi từ truyền thống (TF-IDF). Đồng thời, đề tài đề xuất mô hình lai ghép đa tầng hiệu quả giữa Đồ thị tri thức, Isolation Forest và Deep Autoencoder để chấm điểm rủi ro liên tục cho hồ sơ nghiên cứu.", bold_prefix="Ý nghĩa khoa học: ")
    add_body(doc, "Về mặt thực tiễn, đề tài cung cấp một sản phẩm phần mềm hoàn chỉnh, có tính ứng dụng cao và sẵn sàng vận hành tại Trường Đại học Công Thương TP. HCM. Hệ thống giúp giảm trên 80% thời gian rà soát sơ bộ, tự động cảnh báo các rủi ro tạp chí săn mồi, phát hiện trùng lặp đề tài, tối ưu hóa công tác phân công hội đồng chấm và hỗ trợ báo cáo thống kê chính xác, góp phần đẩy nhanh tiến trình chuyển đổi số toàn diện của Nhà trường.", bold_prefix="Ý nghĩa thực tiễn: ")

    add_h1(doc, "6. BỐ CỤC KHÓA LUẬN")
    add_body(doc, "Báo cáo khóa luận tốt nghiệp được trình bày trong 05 chương chính, phần Mở đầu, Kết luận và Tài liệu tham khảo theo quy chuẩn HUIT:")
    add_bullet_1(doc, "Trình bày tính cấp thiết, mục tiêu nghiên cứu, đối tượng, phạm vi, phương pháp và ý nghĩa khoa học - thực tiễn của đề tài.", bold_prefix="Phần Mở đầu: ")
    add_bullet_1(doc, "Khảo sát thực trạng quy trình xét duyệt hồ sơ tại các trường đại học; trình bày chi tiết cơ sở toán học và lý thuyết của các thuật toán: TF-IDF, LinearSVC, Platt Scaling, Transformer, PhoBERT, NetworkX, Isolation Forest, PyTorch Autoencoder và Điểm rủi ro tổng hợp.", bold_prefix="Chương 1 – Khảo sát hệ thống và cơ sở lý thuyết: ")
    add_bullet_1(doc, "Xác định các tác nhân người dùng, ma trận phân quyền RBAC, đặc tả các yêu cầu chức năng và phi chức năng, xây dựng biểu đồ Use Case tổng thể, quy trình nghiệp vụ xét duyệt 5 bước và phân tích tương tác lớp BCE.", bold_prefix="Chương 2 – Phân tích hệ thống: ")
    add_bullet_1(doc, "Thiết kế kiến trúc phân tầng 3-Tier MVC, thiết kế cơ sở dữ liệu quan hệ SQLite 3 (chuẩn hóa 3NF), thiết kế chi tiết kiến trúc AI Pipeline và module phát hiện bất thường, cùng thiết kế giao diện tương tác người dùng thời gian thực.", bold_prefix="Chương 3 – Thiết kế hệ thống: ")
    add_bullet_1(doc, "Đặc tả chi tiết môi trường phần cứng, phần mềm, cấu trúc thư mục mã nguồn và minh chứng việc hiện thực hoàn chỉnh 06 phân hệ chức năng trên giao diện web thực tế.", bold_prefix="Chương 4 – Cài đặt hệ thống: ")
    add_bullet_1(doc, "Công bố chi tiết tập dữ liệu thực nghiệm 240 mẫu; đánh giá so sánh công bằng hiệu năng giữa PhoBERT và SVM; phân tích sâu sắc 05 lý do lựa chọn PhoBERT làm mô hình chính; đánh giá phân hệ phát hiện bất thường và thử nghiệm triển khai hệ thống.", bold_prefix="Chương 5 – Thử nghiệm và triển khai: ")
    add_bullet_1(doc, "Tổng kết toàn bộ các kết quả đạt được, đánh giá mức độ hoàn thiện so với mục tiêu ban đầu, thẳng thắn chỉ ra các hạn chế và đề xuất phương hướng phát triển tiếp theo của đề tài.", bold_prefix="Phần Kết luận: ")
    add_bullet_1(doc, "Liệt kê đầy đủ 15 tài liệu học thuật và văn bản pháp quy được trích dẫn theo đúng thể thức chuẩn mực.", bold_prefix="Tài liệu tham khảo: ")

    return
