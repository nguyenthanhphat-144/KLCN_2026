# -*- coding: utf-8 -*-
"""
Generate black-and-white / grayscale diagrams for the academic graduation thesis report.
All text, boxes, lines, and bars are monochrome / grayscale.
"""
import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

os.makedirs('images', exist_ok=True)
plt.rcParams['font.sans-serif'] = 'Times New Roman'
plt.rcParams['font.family'] = 'serif'
plt.rcParams['text.color'] = 'black'
plt.rcParams['axes.labelcolor'] = 'black'
plt.rcParams['xtick.color'] = 'black'
plt.rcParams['ytick.color'] = 'black'

def create_diagram_workflow():
    fig, ax = plt.subplots(figsize=(11, 4.5), dpi=300)
    ax.axis('off')
    steps = [
        (0.02, 0.35, 0.16, 0.40, 'Bước 1: Tiếp Nhận\n& Khóa Tự Động\n(Hàng đợi ƯU TIÊN,\nChuyển "Đang xét duyệt"\nTránh xung đột cán bộ)', '#f5f5f5', '#000000'),
        (0.22, 0.35, 0.16, 0.40, 'Bước 2: Thẩm Định\nChuyên Môn\n(Đối chiếu minh chứng,\nXem gợi ý PhoBERT,\nKiểm tra cảnh báo AI,\nNhập nhận xét)', '#f5f5f5', '#000000'),
        (0.42, 0.35, 0.16, 0.40, 'Bước 3: Ra Quyết Định\n(3 nhánh thẩm định:\n1. Đã duyệt\n2. Từ chối [lý do]\n3. Yêu cầu bổ sung)', '#f5f5f5', '#000000'),
        (0.62, 0.35, 0.16, 0.40, 'Bước 4: Lưu Vết\nLịch Sử Bất Biến\n(Ghi LichSuXetDuyet:\nThời gian, cán bộ,\ntrạng thái, ý kiến,\nRead-only bảo mật)', '#f5f5f5', '#000000'),
        (0.82, 0.35, 0.16, 0.40, 'Bước 5: Phản Hồi\n& Resubmission\n(Tác giả nhận thông báo,\nSửa hồ sơ, nộp lại ->\nTự động trở về trạng thái\n"Chờ xét duyệt")', '#f5f5f5', '#000000'),
    ]
    for x, y, w, h, text, facecolor, edgecolor in steps:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02', facecolor=facecolor, edgecolor=edgecolor, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h/2, text, ha='center', va='center', fontsize=8, weight='bold', color='#000000')

    for i in range(len(steps)-1):
        x1 = steps[i][0] + steps[i][2]
        x2 = steps[i+1][0]
        y = 0.55
        ax.annotate('', xy=(x2, y), xytext=(x1, y),
                    arrowprops=dict(arrowstyle='->', lw=1.8, color='#000000'))

    # Loop arrow from Step 5 to Step 1
    ax.annotate('', xy=(0.10, 0.35), xytext=(0.90, 0.35),
                arrowprops=dict(arrowstyle='->', lw=1.5, color='#333333', connectionstyle='arc3,rad=-0.28', linestyle='--'))
    ax.text(0.50, 0.08, 'Vòng lặp Resubmission: Tác giả bổ sung hồ sơ -> Trở lại Hàng đợi Thẩm định (Bước 1)', ha='center', fontsize=8.5, weight='bold', color='#000000')

    plt.title('Sơ Đồ Quy Trình 5 Bước Kiểm Tra Tính Hợp Lệ & Thẩm Định Hồ Sơ KH&CN', fontsize=11, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig('images/hinh1_quytrinh_5buoc.png', dpi=300)
    plt.close()
    print('Generated grayscale images/hinh1_quytrinh_5buoc.png')

def create_diagram_pipeline():
    fig, ax = plt.subplots(figsize=(10, 4.5), dpi=300)
    ax.axis('off')

    boxes = [
        (0.02, 0.35, 0.16, 0.35, 'Hồ Sơ KH&CN\n(Tiêu đề, Tóm tắt,\nTừ khóa, Tác giả)', '#f9f9f9', '#000000'),
        (0.22, 0.35, 0.16, 0.35, 'Tiền Xử Lý\n(PyVi Tách từ,\nLọc Stopwords,\n4 trường ghép nối)', '#f0f0f0', '#000000'),
        (0.42, 0.55, 0.22, 0.35, 'PhoBERT Deep Learning\n(vinai/phobert-base,\nClassification Head)', '#e0e0e0', '#000000'),
        (0.42, 0.12, 0.22, 0.35, 'Baseline LinearSVC\n(TF-IDF N-gram,\nPlatt Scaling Calibrated)', '#e0e0e0', '#000000'),
        (0.68, 0.35, 0.14, 0.35, 'Ensemble / Best\nConfidence Score\n(Softmax Prob %)', '#d0d0d0', '#000000'),
        (0.86, 0.35, 0.12, 0.35, 'Gợi Ý Lĩnh Vực\n& Cảnh Báo\nLệch Nhãn', '#c0c0c0', '#000000'),
    ]

    for x, y, w, h, text, facecolor, edgecolor in boxes:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02', facecolor=facecolor, edgecolor=edgecolor, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h/2, text, ha='center', va='center', fontsize=8.5, weight='bold', color='#000000')

    arrows = [
        (0.18, 0.525, 0.22, 0.525),
        (0.38, 0.55, 0.42, 0.725),
        (0.38, 0.50, 0.42, 0.295),
        (0.64, 0.725, 0.68, 0.55),
        (0.64, 0.295, 0.68, 0.50),
        (0.82, 0.525, 0.86, 0.525),
    ]
    for x1, y1, x2, y2 in arrows:
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle='->', lw=1.5, color='#000000'))

    plt.title('Kiến Trúc Pipeline Tiền Xử Lý & Phân Loại Văn Bản Tự Động (PhoBERT & LinearSVC)', fontsize=11, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig('images/hinh2_kientruc_ai_pipeline.png', dpi=300)
    plt.close()
    print('Generated grayscale images/hinh2_kientruc_ai_pipeline.png')

def create_diagram_anomaly():
    fig, ax = plt.subplots(figsize=(11, 5.2), dpi=300)
    ax.axis('off')
    modules = [
        (0.03, 0.58, 0.28, 0.32, 'Đồ Thị Tri Thức NetworkX\nMulti-partite Knowledge Graph:\nNodes: Tác giả, Công trình, Lĩnh vực, Tạp chí\nPhân tích Centrality & Lệch chuyên môn', '#f5f5f5', '#000000'),
        (0.35, 0.58, 0.28, 0.32, 'Rừng Cô Lập (Isolation Forest)\n100 iTrees, Contamination=0.20\nĐo độ sâu đường đi trung bình E(h(x))\nChuẩn hóa điểm cô lập [0.0, 1.0]', '#eeeeee', '#000000'),
        (0.67, 0.58, 0.30, 0.32, 'Deep Autoencoder MLP\nBottleneck MLP (6 -> 16 -> 4 -> 16 -> 6)\nSai số tái tạo MSE = (1/d) sum((x - x_rec)^2)\nNắm bắt tương quan phi tuyến đa chiều', '#e0e0e0', '#000000'),
        (0.03, 0.12, 0.28, 0.32, 'Kiểm Tra Trùng Lặp & Quy Tắc\nCosine TF-IDF Vector Sim >= 62.0%\nĐối chiếu Tạp chí Rác / Ngừng chỉ mục\nCảnh báo Lệch nhãn chuyên ngành', '#f5f5f5', '#000000'),
        (0.35, 0.12, 0.28, 0.32, 'Điểm Rủi Ro Tổng Hợp\nComposite Risk Score Ri in [0.0, 1.0]\nEnsemble trọng số đa nguồn\nPhân loại Rủi ro Cao / Trung bình / An toàn', '#e0e0e0', '#000000'),
        (0.67, 0.12, 0.30, 0.32, 'Hàng Đợi Thẩm Định Ưu Tiên\nPriority Review Queue\nSắp xếp giảm dần theo Điểm Rủi Ro\n(Cán bộ xử lý hồ sơ rủi ro trước FIFO)', '#cccccc', '#000000'),
    ]
    for x, y, w, h, text, facecolor, edgecolor in modules:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02', facecolor=facecolor, edgecolor=edgecolor, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h/2, text, ha='center', va='center', fontsize=8, weight='bold', color='#000000')

    ax.annotate('', xy=(0.49, 0.44), xytext=(0.17, 0.58), arrowprops=dict(arrowstyle='->', lw=1.2, color='#000000'))
    ax.annotate('', xy=(0.49, 0.44), xytext=(0.49, 0.58), arrowprops=dict(arrowstyle='->', lw=1.2, color='#000000'))
    ax.annotate('', xy=(0.49, 0.44), xytext=(0.82, 0.58), arrowprops=dict(arrowstyle='->', lw=1.2, color='#000000'))
    ax.annotate('', xy=(0.35, 0.28), xytext=(0.31, 0.28), arrowprops=dict(arrowstyle='->', lw=1.5, color='#000000'))
    ax.annotate('', xy=(0.67, 0.28), xytext=(0.63, 0.28), arrowprops=dict(arrowstyle='->', lw=1.5, color='#000000'))

    plt.title('Kiến Trúc Phát Hiện Bất Thường Đa Tầng & Xếp Hạng Ưu Tiên (Chức Năng 6)', fontsize=11, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig('images/hinh2_phat_hien_bat_thuong.png', dpi=300)
    plt.close()
    print('Generated grayscale images/hinh2_phat_hien_bat_thuong.png')

def create_diagram_experiments():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.2), dpi=300)
    metrics = ['Accuracy', 'Precision', 'Recall', 'F1-score']
    phobert_scores = [0.942, 0.948, 0.939, 0.943]
    svm_scores = [0.885, 0.891, 0.880, 0.884]
    x = np.arange(len(metrics))
    width = 0.35
    ax1.bar(x - width/2, phobert_scores, width, label='PhoBERT (vinai/phobert-base)', color='#333333')
    ax1.bar(x + width/2, svm_scores, width, label='Baseline LinearSVC (TF-IDF)', color='#888888')
    ax1.set_ylabel('Điểm số (0.0 - 1.0)', fontsize=9, weight='bold')
    ax1.set_title('So Sánh Hiệu Năng Phân Loại Chuyên Ngành', fontsize=10, weight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(metrics, fontsize=9)
    ax1.set_ylim(0.75, 1.0)
    ax1.legend(loc='lower right', fontsize=8)
    ax1.grid(axis='y', linestyle='--', alpha=0.5, color='#cccccc')

    models = ['Isolation Forest', 'Deep Autoencoder']
    det_rate = [25.0, 25.0]
    time_ms = [12.5, 35.8]
    x2 = np.arange(len(models))
    ax2.bar(x2 - width/2, det_rate, width, label='Tỷ Lệ Bất Thường (%)', color='#444444')
    ax2_twin = ax2.twinx()
    ax2_twin.bar(x2 + width/2, time_ms, width, label='Thời Gian Tính Toán (ms)', color='#999999')
    ax2.set_ylabel('Tỷ Lệ Bất Thường (%)', fontsize=9, color='#000000', weight='bold')
    ax2_twin.set_ylabel('Thời Gian (ms)', fontsize=9, color='#333333', weight='bold')
    ax2.set_title('So Sánh Hiệu Năng Phát Hiện Bất Thường', fontsize=10, weight='bold')
    ax2.set_xticks(x2)
    ax2.set_xticklabels(models, fontsize=9)
    ax2.set_ylim(0, 40)
    ax2_twin.set_ylim(0, 50)
    ax2.grid(axis='y', linestyle='--', alpha=0.5, color='#cccccc')

    plt.tight_layout()
    plt.savefig('images/hinh5_danhgia_hieunang.png', dpi=300)
    plt.close()
    print('Generated grayscale images/hinh5_danhgia_hieunang.png')

def create_diagram_system_architecture():
    fig, ax = plt.subplots(figsize=(11, 5.0), dpi=300)
    ax.axis('off')
    
    # Layer 1: Presentation Layer
    rect1 = patches.FancyBboxPatch((0.03, 0.68), 0.94, 0.25, boxstyle='round,pad=0.02', facecolor='#f5f5f5', edgecolor='#000000', linewidth=1.5)
    ax.add_patch(rect1)
    ax.text(0.06, 0.88, 'TẦNG TRÌNH DIỄN (PRESENTATION LAYER - UI/UX & TEMPLATES)', fontsize=9, weight='bold', color='#000000')
    ui_boxes = [
        (0.06, 0.71, 0.19, 0.14, 'Dashboard Thống Kê\n(7 KPI, Biểu đồ Chart.js)', '#e8e8e8'),
        (0.28, 0.71, 0.20, 0.14, 'Kê Khai Hồ Sơ Động\n(Live AI Preview)', '#e8e8e8'),
        (0.51, 0.71, 0.20, 0.14, 'Thẩm Định Chi Tiết\n(5 Bước & Cảnh Báo AI)', '#e8e8e8'),
        (0.74, 0.71, 0.20, 0.14, 'Trung Tâm Quản Trị AI\n(Priority Queue, Graph)', '#e8e8e8'),
    ]
    for x, y, w, h, t, c in ui_boxes:
        r = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.01', facecolor=c, edgecolor='#333333', linewidth=1)
        ax.add_patch(r)
        ax.text(x + w/2, y + h/2, t, ha='center', va='center', fontsize=7.5, weight='bold', color='#000000')

    # Layer 2: Business Logic & AI Engine
    rect2 = patches.FancyBboxPatch((0.03, 0.35), 0.94, 0.28, boxstyle='round,pad=0.02', facecolor='#f0f0f0', edgecolor='#000000', linewidth=1.5)
    ax.add_patch(rect2)
    ax.text(0.06, 0.58, 'TẦNG NGHIỆP VỤ & TRÍ TUỆ NHÂN TẠO (APPLICATION & AI ENGINE)', fontsize=9, weight='bold', color='#000000')
    app_boxes = [
        (0.06, 0.38, 0.17, 0.16, 'Flask Application\n- Routing & Session\n- RBAC (3 Roles)\n- Audit Log Bất biến', '#dddddd'),
        (0.26, 0.38, 0.22, 0.16, 'NLP Tiền Xử Lý\n- PyVi ViTokenizer\n- Hợp nhất 4 trường\n- Lọc từ dừng chuyên ngành', '#dddddd'),
        (0.51, 0.38, 0.21, 0.16, 'AI Phân Loại\n- VinAI PhoBERT-base\n- Baseline LinearSVC\n- Feedback Retraining Loop', '#dddddd'),
        (0.75, 0.38, 0.20, 0.16, 'Phát Hiện Bất Thường\n- Đồ thị NetworkX\n- Isolation Forest (iTrees)\n- Deep Autoencoder MLP', '#dddddd'),
    ]
    for x, y, w, h, t, c in app_boxes:
        r = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.01', facecolor=c, edgecolor='#333333', linewidth=1)
        ax.add_patch(r)
        ax.text(x + w/2, y + h/2, t, ha='center', va='center', fontsize=7.5, weight='bold', color='#000000')

    # Layer 3: Data Access & Storage Layer
    rect3 = patches.FancyBboxPatch((0.03, 0.05), 0.94, 0.25, boxstyle='round,pad=0.02', facecolor='#e8e8e8', edgecolor='#000000', linewidth=1.5)
    ax.add_patch(rect3)
    ax.text(0.06, 0.25, 'TẦNG LƯU TRỮ DỮ LIỆU (DATA PERSISTENCE & INTEGRATION LAYER)', fontsize=9, weight='bold', color='#000000')
    db_boxes = [
        (0.06, 0.08, 0.27, 0.14, 'SQL Server 2025 / SQLite\n- HoSoKHCN (12 cột mở rộng)\n- LichSuXetDuyet (Lưu vết)\n- LinhVuc & TapChiHopLe', '#d0d0d0'),
        (0.36, 0.08, 0.28, 0.14, 'File Storage & Minh Chứng\n- Minh chứng PDF, DOCX\n- Checksum & Timestamp\n- Thư mục uploads/ an toàn', '#d0d0d0'),
        (0.67, 0.08, 0.27, 0.14, 'Import / Export Engine\n- Bulk Ingestion Excel\n- Xuất báo cáo PDF / Excel\n- Seed Data Demo', '#d0d0d0'),
    ]
    for x, y, w, h, t, c in db_boxes:
        r = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.01', facecolor=c, edgecolor='#333333', linewidth=1)
        ax.add_patch(r)
        ax.text(x + w/2, y + h/2, t, ha='center', va='center', fontsize=7.5, weight='bold', color='#000000')

    # Layer connecting arrows
    ax.annotate('', xy=(0.50, 0.68), xytext=(0.50, 0.63), arrowprops=dict(arrowstyle='<->', lw=1.5, color='#000000'))
    ax.annotate('', xy=(0.50, 0.35), xytext=(0.50, 0.30), arrowprops=dict(arrowstyle='<->', lw=1.5, color='#000000'))

    plt.title('Sơ Đồ Kiến Trúc Hệ Thống Tổng Thể 3 Tầng (3-Tier Architecture)', fontsize=11, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig('images/hinh3_kientruc_he_thong.png', dpi=300)
    plt.close()
    print('Generated grayscale images/hinh3_kientruc_he_thong.png')

if __name__ == '__main__':
    create_diagram_workflow()
    create_diagram_pipeline()
    create_diagram_anomaly()
    create_diagram_experiments()
    create_diagram_system_architecture()
