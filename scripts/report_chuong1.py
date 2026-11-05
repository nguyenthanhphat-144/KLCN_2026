# -*- coding: utf-8 -*-
"""
CHƯƠNG 1: KHẢO SÁT HỆ THỐNG VÀ CƠ SỞ LÝ THUYẾT
Bao gồm:
- 1.1 Khảo sát hiện trạng quản lý hồ sơ KH&CN tại các trường đại học
  - Bảng 1.1: Bảng đối chiếu hiện trạng quy trình thủ công và giải pháp ứng dụng AI
  - Hình 1.1: Quy trình rà soát thủ công truyền thống và các nút thắt hiện nay
- 1.2 Cơ sở lý thuyết về mô hình phân loại văn bản học máy truyền thống
  - Công thức TF-IDF (OMML)
  - Công thức LinearSVC (OMML)
  - Công thức Platt Scaling (OMML)
- 1.3 Cơ sở lý thuyết về mô hình học sâu biến đổi ngữ cảnh PhoBERT
  - Công thức Scaled Attention (OMML)
  - Công thức Multi-Head Attention (OMML)
  - Bảng 1.2: Bảng thông số kỹ thuật cốt lõi của mô hình VinAI PhoBERT Base
  - Công thức Joint Loss (OMML)
- 1.4 Cơ sở lý thuyết về phát hiện bất thường và đồ thị tri thức
  - Công thức Degree Centrality (OMML)
  - Công thức Isolation Forest (OMML)
  - Công thức PyTorch Autoencoder MSE (OMML)
  - Công thức Cosine Similarity (OMML)
  - Công thức Composite Risk Score (OMML)
  - Bảng 1.3: Bảng trọng số thành phần cấu thành Điểm rủi ro tổng hợp
"""

import os
from docx.shared import Pt, Cm, RGBColor

def add_chuong_1(doc, helpers):
    add_chapter_title = helpers["add_chapter_title"]
    add_h1 = helpers["add_h1"]
    add_h2 = helpers["add_h2"]
    add_h3 = helpers["add_h3"]
    add_body = helpers["add_body"]
    add_bullet_1 = helpers["add_bullet_1"]
    add_bullet_2 = helpers["add_bullet_2"]
    add_figure_custom = helpers["add_figure_custom"]
    add_table_custom = helpers["add_table_custom"]

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

    # ==================== BẮT ĐẦU CHƯƠNG 1 ====================
    add_chapter_title(doc, "CHƯƠNG 1: KHẢO SÁT HỆ THỐNG VÀ CƠ SỞ LÝ THUYẾT")

    # 1.1 KHẢO SÁT HIỆN TRẠNG
    add_h1(doc, "1.1. KHẢO SÁT HIỆN TRẠNG QUẢN LÝ HỒ SƠ KH&CN TẠI CÁC TRƯỜNG ĐẠI HỌC")
    
    add_h2(doc, "1.1.1. Thực trạng quy trình tiếp nhận và xét duyệt hồ sơ thủ công")
    add_body(doc, "Hoạt động nghiên cứu khoa học và chuyển giao công nghệ tại các trường đại học đóng vai trò thước đo then chốt phản ánh uy tín học thuật của đội ngũ giảng viên và chất lượng đào tạo của nhà trường. Hàng năm, căn cứ vào Quy chế chi tiêu nội bộ và Quy định về định mức giờ nghiên cứu khoa học, cán bộ giảng viên phải thực hiện kê khai toàn bộ các sản phẩm khoa học công nghệ phát sinh trong năm học, bao gồm: bài báo xuất bản trên các tạp chí quốc tế (thuộc chỉ mục Scopus, Web of Science), kỷ yếu hội thảo quốc gia/quốc tế, bài báo trên các tạp chí trong nước có điểm của Hội đồng Giáo sư Nhà nước (HĐGSNN), các đề tài NCKH các cấp (cấp Cơ sở, cấp Bộ, cấp Tỉnh/Thành phố, cấp Nhà nước), và các đơn đăng ký hoặc bằng độc quyền Sáng chế / Giải pháp hữu ích.")
    add_body(doc, "Qua khảo sát thực tế tại Trường Đại học Công Thương TP. Hồ Chí Minh và một số cơ sở giáo dục đại học đối tác, quy trình tiếp nhận, xét duyệt và tổng hợp hồ sơ KH&CN hiện nay chủ yếu vận hành theo phương thức bán thủ công với nhiều hạn chế cố hữu:")
    add_bullet_1(doc, "Quy trình nộp phân tán: Giảng viên tự tổng hợp thông tin, điền vào các mẫu biểu Excel hoặc biểu mẫu Google Forms, sau đó gửi kèm các tập tin scan minh chứng (dạng PDF, hình ảnh) qua email hoặc nộp trực tiếp bản in tại Văn phòng Khoa/Bộ môn.", bold_prefix="Khâu tiếp nhận: ")
    add_bullet_1(doc, "Rà soát tra cứu hoàn toàn bằng sức người: Chuyên viên Phòng Quản lý Khoa học phải đối chiếu từng dòng thông tin: tra cứu chỉ số ISSN của tạp chí trên cơ sở dữ liệu Scopus/WoS, kiểm tra danh mục điểm số HĐGSNN tương ứng với năm công bố, xác thực vai trò tác giả (tác giả chính, tác giả liên hệ hay thành viên), và kiểm tra tính hợp lệ của biên lai nộp lệ phí hoặc quyết định nghiệm thu đề tài.", bold_prefix="Khâu thẩm định: ")
    add_bullet_1(doc, "Tổng hợp dữ liệu dễ phát sinh sai sót: Sau khi các Hội đồng Khoa học họp thẩm định, dữ liệu kết quả được sao chép thủ công qua lại giữa các file Excel tổng hợp cấp Khoa và cấp Trường. Việc này dẫn đến nguy cơ sai lệch số liệu, nhầm lẫn tác giả, và khó khăn khi cần thống kê tra cứu nhanh.", bold_prefix="Khâu tổng hợp: ")

    t1_headers = ["Tiêu chí so sánh", "Phương thức quản lý thủ công hiện tại", "Hệ thống tự động hóa tích hợp AI"]
    t1_data = [
        ["Thời gian thẩm định trung bình", "Từ 15 đến 30 ngày cho mỗi đợt quyết toán", "Chỉ từ 1 đến 3 phút (gợi ý AI và cảnh báo tức thời)"],
        ["Phân loại lĩnh vực nghiên cứu", "Phân loại bằng cảm tính chuyên viên, dễ nhầm liên ngành", "Mô hình PhoBERT phân loại ngữ cảnh tự động (F1 = 89.92%)"],
        ["Xác thực tạp chí uy tín", "Chuyên viên tra cứu tay trên cổng Scopus/WoS/HĐGSNN", "Tự động đối chiếu CSDL chuẩn hóa ISSN, cảnh báo săn mồi"],
        ["Kiểm soát trùng lặp hồ sơ", "Gần như bất khả thi khi số lượng hồ sơ lớn", "Đo độ tương đồng Cosine TF-IDF phát hiện trùng lập tức"],
        ["Phát hiện gian lận tác giả", "Không phát hiện được nhóm liên kết chéo ảo", "Phân tích đồ thị đồng tác giả (NetworkX Co-authorship)"],
        ["Lưu vết thẩm định (Audit Trail)", "Ghi chép biên bản giấy, dễ thất lạc hoặc sửa đổi", "Lưu vết bất biến từng bước xét duyệt, minh bạch 100%"],
        ["Trải nghiệm giảng viên", "Mất nhiều công sức điền mẫu, không biết tiến độ", "Giao diện web trực quan, Live AI Preview, xem tiến độ 24/7"]
    ]
    add_table_custom(doc, 1, "Bảng đối chiếu hiện trạng quy trình thủ công và giải pháp ứng dụng AI",
                     t1_headers, t1_data, col_widths=[4.0, 5.5, 6.0])

    add_body(doc, "Sự quá tải trong quy trình rà soát thủ công được minh họa chi tiết qua sơ đồ các nút thắt nghiệp vụ dưới đây:")
    add_figure_custom(doc, "images/hinh1_hien_trang_thu_cong.png", 1,
                      "Quy trình rà soát thủ công truyền thống và các nút thắt hiện nay", width_cm=14.5)

    add_h2(doc, "1.1.2. Khảo sát các hệ thống quản lý nghiên cứu khoa học hiện hữu")
    add_body(doc, "Trên thế giới, các hệ thống Quản lý Thông tin Nghiên cứu (Current Research Information Systems - CRIS) tiêu biểu như Pure (Elsevier), Converis (Clarivate), Symplectic Elements, hay Vivo (Duraspace) đã được triển khai rộng rãi tại các viện - trường đại học hàng đầu. Các hệ thống này sở hữu ưu điểm vượt trội về khả năng kết nối tự động với các kho dữ liệu thư mục toàn cầu thông qua API chuyên dụng. Tuy nhiên, rào cản lớn nhất đối với các trường đại học tại Việt Nam là chi phí bản quyền và duy trì hàng năm cực kỳ đắt đỏ (lên đến hàng chục nghìn USD mỗi năm), đồng thời các hệ thống ngoại nhập này không tương thích với quy định quản lý đề tài KH&CN đặc thù của Việt Nam (như cách tính điểm công trình theo Hội đồng Giáo sư Nhà nước, định mức giờ chuẩn theo quy chế Bộ GD&ĐT).")
    add_body(doc, "Tại Việt Nam, một số đơn vị phần mềm trong nước đã phát triển các phân hệ quản lý đề tài NCKH tích hợp trong hệ thống ERP trường học. Mặc dù vậy, các giải pháp này mới chỉ dừng lại ở mức độ lưu trữ văn bản số hóa dạng CRUD (Thêm, Xóa, Sửa đơn thuần), chưa có khả năng hiểu ngữ nghĩa văn bản tiếng Việt để tự động phân loại chuyên ngành, và hoàn toàn thiếu vắng các cơ chế Trí tuệ Nhân tạo thông minh để phát hiện các dị biệt hay dấu hiệu gian lận học thuật.")

    add_h2(doc, "1.1.3. Nhu cầu cấp thiết về ứng dụng trí tuệ nhân tạo và phát hiện bất thường")
    add_body(doc, "Từ những bất cập thực tiễn nêu trên, việc nghiên cứu xây dựng một hệ thống hỗ trợ quản lý hồ sơ KH&CN hoàn chỉnh, được thiết kế chuyên biệt cho môi trường giáo dục đại học Việt Nam là nhu cầu vô cùng cấp thiết. Hệ thống không chỉ giải quyết bài toán tự động hóa nghiệp vụ hành chính mà còn phải tích hợp sâu các thuật toán AI hiện đại:")
    add_bullet_1(doc, "Mô hình ngôn ngữ học sâu (PhoBERT): Tự động đọc hiểu tóm tắt, tên công trình và từ khóa tiếng Việt để dự đoán chính xác lĩnh vực nghiên cứu và nhóm loại công trình, hỗ trợ giảng viên kê khai nhanh và giúp ban thẩm định phân bổ đúng hội đồng chuyên môn.")
    add_bullet_1(doc, "Mô hình phát hiện bất thường đa tầng (NetworkX + Isolation Forest + PyTorch Autoencoder): Tự động phát hiện các hồ sơ có nguy cơ vi phạm cao (như trùng lặp tên công trình, tạp chí có dấu hiệu săn mồi, thiếu minh chứng hoặc bất thường trong cơ cấu phân bổ giờ), từ đó chấm Điểm rủi ro tổng hợp và sắp xếp vào Hàng đợi xét duyệt ưu tiên, giúp cán bộ tập trung thời gian kiểm tra các ca nghi vấn.")

    # 1.2 CƠ SỞ LÝ THUYẾT MÔ HÌNH HỌC MÁY TRUYỀN THỐNG
    add_h1(doc, "1.2. CƠ SỞ LÝ THUYẾT VỀ MÔ HÌNH PHÂN LOẠI VĂN BẢN HỌC MÁY TRUYỀN THỐNG")

    add_h2(doc, "1.2.1. Mô hình không gian vector TF-IDF (Term Frequency - Inverse Document Frequency)")
    add_body(doc, "Trong bài toán Xử lý ngôn ngữ tự nhiên (NLP) cổ điển, văn bản được biểu diễn thành các vector số học thông qua mô hình không gian vector (Vector Space Model). Phương pháp trích xuất đặc trưng phổ biến và hiệu quả nhất là trọng số TF-IDF, phản ánh tầm quan trọng tương đối của một từ (hoặc cụm n-gram) đối với một tài liệu cụ thể trong toàn bộ tập ngữ liệu (corpus).")
    add_body(doc, "Trọng số TF-IDF của từ t trong tài liệu d thuộc tập ngữ liệu D được tính theo công thức chuẩn của thư viện Scikit-learn (sử dụng thang đo logarit mịn và chuẩn hóa l2):")

    add_omml_tfidf(doc)

    add_body(doc, "Trong đó: tf(t, d) là tần suất xuất hiện của từ t trong tài liệu d; |D| là tổng số tài liệu có trong tập ngữ liệu; df(t, D) là tần suất văn bản, biểu thị số lượng tài liệu trong tập D có chứa từ t. Thành phần nghịch đảo tần suất văn bản IDF đóng vai trò gán trọng số thấp cho các hư từ xuất hiện phổ biến trong mọi văn bản (như 'và', 'của', 'là', 'trong') và gán trọng số cao cho các từ khóa chuyên ngành mang tính phân biệt cao (như 'deep learning', 'lên men', 'chất xúc tác').")

    add_h2(doc, "1.2.2. Thuật toán Support Vector Machine (LinearSVC) và hiệu chỉnh xác suất Platt Scaling")
    add_body(doc, "Máy vectơ hỗ trợ (Support Vector Machine - SVM) là một trong những thuật toán học máy có giám sát mạnh mẽ nhất đối với bài toán phân loại văn bản dạng vector thưa chiều cao. Đối với không gian đặc trưng TF-IDF với hàng nghìn chiều, Linear Support Vector Classifier (LinearSVC) giải bài toán tối ưu lồi nhằm tìm một siêu phẳng (hyperplane) phân cách tối ưu có biên phân cách (margin) lớn nhất giữa các lớp dữ liệu:")

    add_omml_linearsvc(doc)

    add_body(doc, "Trong đó: w là vector trọng số pháp tuyến của siêu phẳng; b là hệ số tự do; x_i là vector đặc trưng TF-IDF của mẫu văn bản thứ i; y_i thuộc {-1, +1} là nhãn lớp; C > 0 là tham số điều hòa (regularization parameter) kiểm soát sự cân bằng giữa độ rộng biên phân cách và việc phạt các điểm dữ liệu vi phạm biên.")
    add_body(doc, "Do LinearSVC xuất ra giá trị khoảng cách hình học có dấu f(x) = w^T x + b thay vì xác suất hậu nghiệm thực sự, đề tài áp dụng kỹ thuật hiệu chuẩn xác suất Platt Scaling (thông qua CalibratedClassifierCV với phương pháp sigmoid) để chuyển đổi giá trị quyết định f(x) thành xác suất tin cậy P(y = 1 | x) nằm trong khoảng [0, 1]:")

    add_omml_platt_scaling(doc)

    add_body(doc, "Trong đó A và B là các tham số vô hướng được ước lượng thông qua phương pháp tối đa hóa hợp lý (Maximum Likelihood Estimation) trên tập dữ liệu kiểm thực chéo nội bộ. Nhờ đó, hệ thống có thể cung cấp độ tin cậy phần trăm (%) cho người dùng khi sử dụng mô hình SVM.")

    # 1.3 CƠ SỞ LÝ THUYẾT MÔ HÌNH HỌC SÂU BIẾN ĐỔI NGỮ CẢNH PHOBERT
    add_h1(doc, "1.3. CƠ SỞ LÝ THUYẾT VỀ MÔ HÌNH HỌC SÂU BIẾN ĐỔI NGỮ CẢNH PHOBERT")

    add_h2(doc, "1.3.1. Kiến trúc mạng Transformer và cơ chế Multi-Head Self-Attention")
    add_body(doc, "Khác biệt căn bản với các kiến trúc chuỗi truyền thống như RNN hay LSTM bị giới hạn bởi tính tuần tự và vấn đề tiêu biến đạo hàm khi xử lý chuỗi dài, kiến trúc Transformer (Vaswani et al., 2017) dựa hoàn toàn trên cơ chế Tự chú ý (Self-Attention), cho phép mô hình nắm bắt mối quan hệ phụ thuộc ngữ nghĩa hai chiều giữa tất cả các từ trong câu mà không phụ thuộc vào khoảng cách từ vựng.")
    add_body(doc, "Cơ chế Chú ý tích vô hướng tỷ lệ (Scaled Dot-Product Attention) ánh xạ một tập hợp gồm ma trận Truy vấn (Query - Q), Khóa (Key - K) và Giá trị (Value - V) thành một ma trận đầu ra:")

    add_omml_scaled_attention(doc)

    add_body(doc, "Trong đó d_k là số chiều của vector khóa, phép chia cho căn bậc hai của d_k giúp ổn định gradient khi số chiều lớn. Nhằm cho phép mô hình đồng thời chú ý đến thông tin từ các không gian biểu diễn khác nhau tại các vị trí khác nhau, cơ chế Multi-Head Attention thực hiện chiếu tuyến tính Q, K, V qua h đầu chú ý độc lập rồi ghép nối lại:")

    add_omml_multihead_attention(doc)

    add_h2(doc, "1.3.2. Mô hình ngôn ngữ tiếng Việt tiền huấn luyện VinAI PhoBERT (vinai/phobert-base)")
    add_body(doc, "VinAI PhoBERT (Nguyễn & Nguyễn, 2020) là mô hình ngôn ngữ ngữ cảnh đơn ngữ hiện đại đầu tiên được huấn luyện chuyên sâu cho tiếng Việt dựa trên kiến trúc RoBERTa (Robustly Optimized BERT Approach). Mô hình giải quyết bài toán đặc thù của tiếng Việt - ngôn ngữ đơn lập, ranh giới từ không trùng với khoảng trắng - bằng cách kết hợp bộ tách từ ghép tiếng Việt PyVi và thuật toán mã hóa cặp byte (Byte-Pair Encoding - BPE) với vốn từ vựng 64.000 subwords.")

    t2_headers = ["Thông số kỹ thuật", "Giá trị cấu hình mô hình vinai/phobert-base"]
    t2_data = [
        ["Kiến trúc nền tảng", "RoBERTa (Chỉ khối Encoder của Transformer)"],
        ["Tập dữ liệu tiền huấn luyện", "20GB văn bản tiếng Việt (~3 tỷ từ, gồm Wikipedia tiếng Việt và Báo chí)"],
        ["Số lớp Transformer (Layers)", "12 lớp ẩn (Hidden Layers)"],
        ["Số chiều vector ẩn (Hidden Size)", "768 chiều"],
        ["Số đầu chú ý (Attention Heads)", "12 đầu chú ý độc lập"],
        ["Quy mô từ vựng (Vocabulary Size)", "64.000 token (BPE subwords)"],
        ["Tổng số tham số mạng", "~135 triệu tham số (135M parameters)"],
        ["Phương pháp tiền xử lý từ ngữ", "Tách từ ghép PyVi (nối từ ghép bằng dấu gạch dưới '_')"]
    ]
    add_table_custom(doc, 1, "Bảng thông số kỹ thuật cốt lõi của mô hình VinAI PhoBERT Base áp dụng trong đề tài",
                     t2_headers, t2_data, col_widths=[5.5, 9.5])

    add_h2(doc, "1.3.3. Phương pháp Fine-tuning đa nhiệm (Multi-task Classification) và hàm mất mát liên hợp")
    add_body(doc, "Thay vì xây dựng hai mô hình mạng nơ-ron riêng biệt gây lãng phí bộ nhớ và tăng thời gian xử lý, nhóm nghiên cứu thiết kế một kiến trúc mạng phân loại đa nhiệm (Multi-task Classification Head) kế thừa chung thân mạng nơ-ron PhoBERT (PhoBERT Backbone). Vector biểu diễn ngữ cảnh tổng thể của toàn bộ văn bản hồ sơ được trích xuất từ token đặc biệt <s> (hoặc [CLS]) tại lớp ẩn thứ 12: h_CLS thuộc R^768.")
    add_body(doc, "Vector h_CLS được truyền qua lớp Dropout (tỷ lệ 0.2) và đưa đồng thời vào hai đầu phân loại tuyến tính độc lập:")
    add_bullet_1(doc, "Đầu phân loại Lĩnh vực nghiên cứu: z_field = W_field * h_CLS + b_field (với 5 lớp phân loại).", bold_prefix="Head 1: ")
    add_bullet_1(doc, "Đầu phân loại Nhóm loại công trình: z_type = W_type * h_CLS + b_type (với 3 nhóm lớp: Bài báo, Đề tài, Sáng chế).", bold_prefix="Head 2: ")
    add_body(doc, "Toàn bộ mạng được huấn luyện tinh chỉnh (Fine-tuning) đồng thời (End-to-End) dựa trên hàm mất mát liên hợp tổng (Joint Cross-Entropy Loss):")

    add_omml_joint_loss(doc)

    add_body(doc, "Hệ số phạt 0.5 cho bài toán phân loại loại công trình giúp cân bằng tỷ trọng lan truyền ngược, bảo đảm mô hình tối ưu đồng thời cả hai mục tiêu học máy mà không bị hiện tượng lấn át gradient.")

    # 1.4 CƠ SỞ LÝ THUYẾT PHÁT HIỆN BẤT THƯỜNG
    add_h1(doc, "1.4. CƠ SỞ LÝ THUYẾT VỀ PHÁT HIỆN BẤT THƯỜNG VÀ ĐỒ THỊ TRI THỨC")

    add_h2(doc, "1.4.1. Đồ thị tri thức đồng tác giả và Độ đo mức độ trung tâm Degree Centrality (NetworkX)")
    add_body(doc, "Để phát hiện các nhóm tác giả liên kết chéo bất thường hoặc các hồ sơ kê khai đứng tên chung thiếu căn cứ thực tế, hệ thống xây dựng một Đồ thị tri thức đồng tác giả G = (V, E) bằng thư viện NetworkX. Trong đó V là tập hợp các nhà khoa học, và E là tập hợp các cạnh vô hướng có trọng số, đại diện cho số lượng công trình mà hai tác giả đã từng cùng công bố.")
    add_body(doc, "Mức độ trung tâm bậc (Degree Centrality) của một tác giả v trong đồ thị thể hiện mức độ kết nối học thuật của tác giả đó:")

    add_omml_degree_centrality(doc)

    add_body(doc, "Trong đó deg(v) là bậc của đỉnh v (tổng số đồng tác giả trực tiếp của v) và |V| là tổng số tác giả trong toàn hệ thống. Nếu một hồ sơ xuất hiện một tác giả mới hoàn toàn, không có mối liên hệ chuyên môn trong mạng lưới hoặc có tần suất tham gia đột biến bất thường, độ đo đồ thị sẽ kích hoạt chỉ số cảnh báo rủi ro hợp tác học thuật.")

    add_h2(doc, "1.4.2. Thuật toán Rừng cây cô lập (Isolation Forest)")
    add_body(doc, "Rừng cây cô lập (Isolation Forest - Liu et al., 2008) là giải thuật học máy không giám sát tiên tiến chuyên dụng cho bài toán phát hiện ngoại lai. Khác với các thuật toán dựa trên khoảng cách (như k-NN, DBSCAN) vốn đòi hỏi chi phí tính toán O(n^2), Isolation Forest dựa trên nguyên lý: các điểm dữ liệu dị biệt (anomalies) thường có số lượng ít và nằm cách biệt trong không gian đặc trưng, do đó chúng sẽ bị cô lập ở những lát cắt đầu tiên (độ sâu ngắn nhất) trong cây nhị phân ngẫu nhiên (iTree).")
    add_body(doc, "Điểm bất thường s(x, n) của một mẫu hồ sơ x trên một rừng gồm 100 cây cô lập được định nghĩa theo công thức:")

    add_omml_isolation_forest(doc)

    add_body(doc, "Trong đó E(h(x)) là độ sâu trung bình để cô lập mẫu x qua toàn bộ các cây; c(n) là độ sâu trung bình của cây nhị phân tìm kiếm trên tập n mẫu; hằng số Euler-Mascheroni xấp xỉ 0.5772156649. Khi s(x, n) tiến sát về 1.0, mẫu x có xác suất dị biệt cấu trúc rất cao; khi s(x, n) < 0.5, mẫu được xem là thuộc phân phối bình thường.")

    add_h2(doc, "1.4.3. Mạng nơ-ron Tự mã hóa sâu (Deep Autoencoder trong PyTorch)")
    add_body(doc, "Nhằm phát hiện các dạng vi phạm phi tuyến tính phức tạp trong các trường định lượng (như số tác giả, số trang, năm công bố, điểm tạp chí, giờ quy đổi), hệ thống xây dựng một Mạng nơ-ron Tự mã hóa sâu (Deep Autoencoder) kế thừa lớp nn.Module của PyTorch. Mạng gồm hai phần:")
    add_bullet_1(doc, "Bộ mã hóa (Encoder): Nén vector đặc trưng 6 chiều x thuộc R^6 qua các lớp Fully-Connected [6 -> 16 -> 8] kết hợp hàm kích hoạt ReLU và Batch Normalization, tạo ra biểu diễn tiềm ẩn z thuộc R^8 tại nút thắt (bottleneck).", bold_prefix="Encoder: ")
    add_bullet_1(doc, "Bộ giải mã (Decoder): Tái tạo lại dữ liệu ban đầu từ vector tiềm ẩn qua các lớp [8 -> 16 -> 6] tạo thành vector đầu ra x_hat.", bold_prefix="Decoder: ")
    add_body(doc, "Mạng được huấn luyện để giảm thiểu sai số tái tạo bình phương trung bình (Mean Squared Error - MSE) trên tập các hồ sơ hợp lệ chuẩn hóa:")

    add_omml_pytorch_autoencoder(doc)

    add_body(doc, "Khi một hồ sơ có các trường thông tin bất thường (ví dụ: bài báo hội thảo nhưng khai số tác giả quá đông hoặc điểm quy đổi cao vượt khung quy chế), bộ giải mã sẽ không thể tái tạo chính xác các giá trị này, dẫn đến giá trị sai số tái tạo MSE tăng vọt, báo hiệu một dị biệt định lượng đáng ngờ.")

    add_h2(doc, "1.4.4. Độ tương đồng Cosine TF-IDF và Điểm rủi ro tổng hợp Composite Risk Score")
    add_body(doc, "Nhằm ngăn chặn hành vi kê khai trùng lặp công trình nghiên cứu (cùng một bài báo nhưng nộp nhiều lần hoặc đổi thứ tự tác giả), hệ thống tính toán ma trận độ tương đồng góc Cosine trên vector đặc trưng TF-IDF giữa hồ sơ mới nộp d_i và toàn bộ các hồ sơ đã có trong cơ sở dữ liệu d_j:")

    add_omml_cosine_similarity(doc)

    add_body(doc, "Khi Sim(d_i, d_j) vượt quá ngưỡng an toàn (0.85 đối với cùng một tác giả, hoặc 0.92 đối với các tác giả khác nhau), hệ thống lập tức gắn cờ cảnh báo nghi vấn trùng lặp nội dung.")
    add_body(doc, "Để cung cấp cho cán bộ thẩm định một cái nhìn toàn diện, hệ thống tích hợp tất cả các nguồn tín hiệu cảnh báo thành một Điểm rủi ro tổng hợp (Composite Risk Score - R_i) chuẩn hóa trong đoạn [0.0, 0.99]:")

    add_omml_risk_score(doc)

    t3_headers = ["Thành phần rủi ro", "Mô hình / Nguồn dữ liệu", "Trọng số (%)", "Ý nghĩa trong việc phát hiện vi phạm"]
    t3_data = [
        ["Điểm bất thường rừng cô lập (s_IF)", "Scikit-learn Isolation Forest (100 cây)", "35%", "Phát hiện dị biệt trong cấu trúc phân bố các trường dữ liệu"],
        ["Sai số tái tạo đặc trưng (MSE_AE)", "PyTorch Deep Autoencoder (nn.Module)", "35%", "Phát hiện sự bất thường trong tương quan giữa số tác giả, giờ NCKH và loại bài"],
        ["Mức độ vi phạm luật cứng (RuleSeverity)", "Cơ sở tri thức luật chuyên gia (Rule-based)", "30%", "Cảnh báo tạp chí săn mồi, thiếu minh chứng PDF, trùng lặp tên công trình"]
    ]
    add_table_custom(doc, 1, "Bảng trọng số thành phần cấu thành Điểm rủi ro tổng hợp (Composite Risk Score)",
                     t3_headers, t3_data, col_widths=[4.5, 4.5, 2.5, 4.5])

    add_body(doc, "Dựa trên Điểm rủi ro tổng hợp R_i, hệ thống tự động phân loại hồ sơ vào 3 mức độ cảnh báo: An toàn (R_i < 0.35 - màu xanh lá), Cần lưu ý (0.35 <= R_i < 0.65 - màu vàng cam), và Rủi ro cao (R_i >= 0.65 - màu đỏ đậm). Các hồ sơ rủi ro cao sẽ tự động được đẩy lên đầu Hàng đợi xét duyệt ưu tiên (Priority Review Queue) để Hội đồng thẩm định tập trung kiểm tra chuyên sâu.")

    return
