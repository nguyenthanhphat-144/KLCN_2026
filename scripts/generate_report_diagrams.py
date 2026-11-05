# -*- coding: utf-8 -*-
"""
Generate 100% monochrome / grayscale diagrams for the graduation thesis report (Bao_Cao_Khoa_Luan.docx).
All experimental numbers are loaded directly from docs/Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json.
"""
import os
import json
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMAGES_DIR = os.path.join(BASE_DIR, 'images')
BENCH_JSON = os.path.join(BASE_DIR, 'docs', 'Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json')

os.makedirs(IMAGES_DIR, exist_ok=True)
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
        ax.text(x + w / 2, y + h / 2, text, ha='center', va='center', fontsize=8, weight='bold', color='#000000')

    for i in range(len(steps) - 1):
        x1 = steps[i][0] + steps[i][2]
        x2 = steps[i + 1][0]
        y = 0.55
        ax.annotate('', xy=(x2, y), xytext=(x1, y),
                    arrowprops=dict(arrowstyle='->', lw=1.8, color='#000000'))

    ax.annotate('', xy=(0.10, 0.35), xytext=(0.90, 0.35),
                arrowprops=dict(arrowstyle='->', lw=1.5, color='#333333', connectionstyle='arc3,rad=-0.28', linestyle='--'))
    ax.text(0.50, 0.08, 'Vòng lặp Resubmission: Tác giả bổ sung hồ sơ -> Trở lại Hàng đợi Thẩm định (Bước 1)',
            ha='center', fontsize=8.5, weight='bold', color='#000000')

    plt.title('Sơ Đồ Quy Trình 5 Bước Kiểm Tra Tính Hợp Lệ & Thẩm Định Hồ Sơ KH&CN', fontsize=11, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, 'hinh1_quytrinh_5buoc.png'), dpi=300)
    plt.close()
    print('Generated grayscale images/hinh1_quytrinh_5buoc.png')


def create_diagram_pipeline(bench):
    pho_f1 = bench['phobert_metrics']['f1_score']
    svm_f1 = bench['svm_metrics']['f1_score']
    delta_f1 = bench['selected_model']['delta_f1']

    fig, ax = plt.subplots(figsize=(10.8, 4.5), dpi=300)
    ax.axis('off')

    boxes = [
        (0.02, 0.35, 0.16, 0.36, 'Hồ Sơ KH&CN\n(13 trường chuẩn:\nTiêu đề, Tóm tắt,\nTừ khóa, Tác giả,\nĐơn vị chủ trì...)', '#f9f9f9', '#000000'),
        (0.22, 0.35, 0.16, 0.36, 'Tiền Xử Lý Chung\n(PyVi Tách từ,\nLọc Stopwords,\nĐầu vào công bằng\n100% cho 2 mô hình)', '#f0f0f0', '#000000'),
        (0.42, 0.55, 0.22, 0.35, f'Fine-tuning PhoBERT\n(PyTorch vinai/phobert-base,\nF1 Lĩnh vực: {pho_f1}%\nF1 Loại CT: 100.0%)', '#e0e0e0', '#000000'),
        (0.42, 0.12, 0.22, 0.35, f'Baseline TF-IDF + SVM\n(LinearSVC Calibrated,\nF1 Lĩnh vực: {svm_f1}%\nF1 Loại CT: 100.0%)', '#e8e8e8', '#000000'),
        (0.68, 0.35, 0.14, 0.36, f'Chọn Mô Hình\nChính: PhoBERT\n(+{delta_f1}% F1-score\nso với SVM)', '#d0d0d0', '#000000'),
        (0.86, 0.35, 0.12, 0.36, 'Dự Đoán Đa Nhiệm\n(Lĩnh vực & Loại CT)\n& Cảnh Báo Lệch', '#c0c0c0', '#000000'),
    ]

    for x, y, w, h, text, facecolor, edgecolor in boxes:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02', facecolor=facecolor, edgecolor=edgecolor, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + w / 2, y + h / 2, text, ha='center', va='center', fontsize=8.2, weight='bold', color='#000000')

    arrows = [
        (0.18, 0.53, 0.22, 0.53),
        (0.38, 0.56, 0.42, 0.725),
        (0.38, 0.50, 0.42, 0.295),
        (0.64, 0.725, 0.68, 0.56),
        (0.64, 0.295, 0.68, 0.50),
        (0.82, 0.53, 0.86, 0.53),
    ]
    for x1, y1, x2, y2 in arrows:
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle='->', lw=1.5, color='#000000'))

    plt.title('Kiến Trúc Pipeline Tiền Xử Lý Công Bằng, Đối Sánh & Chọn Mô Hình Chính (PyTorch PhoBERT vs TF-IDF + SVM)', fontsize=10, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, 'hinh2_kientruc_ai_pipeline.png'), dpi=300)
    plt.close()
    print('Generated grayscale images/hinh2_kientruc_ai_pipeline.png')


def create_diagram_anomaly():
    fig, ax = plt.subplots(figsize=(11, 5.2), dpi=300)
    ax.axis('off')
    modules = [
        (0.03, 0.58, 0.28, 0.32, 'Đồ Thị Tri Thức NetworkX\nMulti-partite Knowledge Graph:\nNodes: Tác giả, Công trình, Lĩnh vực, Tạp chí\nPhân tích Centrality & Lệch chuyên môn', '#f5f5f5', '#000000'),
        (0.35, 0.58, 0.28, 0.32, 'Rừng Cô Lập (Isolation Forest)\n100 iTrees, Contamination=0.25\nĐo độ sâu đường đi trung bình E(h(x))\nChuẩn hóa điểm cô lập [0.0, 1.0]', '#eeeeee', '#000000'),
        (0.67, 0.58, 0.30, 0.32, 'PyTorch Autoencoder (nn.Module)\nBottleneck (6 -> 16 -> 4 -> 16 -> 6)\nMSELoss + Adam (250 epochs)\nNắm bắt tương quan phi tuyến đa chiều', '#e0e0e0', '#000000'),
        (0.03, 0.12, 0.28, 0.32, 'Kiểm Tra Trùng Lặp & Quy Tắc\nCosine TF-IDF Vector Sim >= 62.0%\nĐối chiếu Tạp chí Cảnh báo / Ngừng chỉ mục\nCảnh báo Lệch nhãn chuyên ngành', '#f5f5f5', '#000000'),
        (0.35, 0.12, 0.28, 0.32, 'Điểm Rủi Ro Tổng Hợp Liên Tục\nR_i = 0.35*IF + 0.35*AE + 0.30*Rule\nĐiểm liên tục thực tế [0.0, 1.0]\nNgưỡng cảnh báo vi phạm >= 0.50', '#e0e0e0', '#000000'),
        (0.67, 0.12, 0.30, 0.32, 'Hàng Đợi Thẩm Định Ưu Tiên\nPriority Review Queue\nSắp xếp giảm dần theo Điểm Rủi Ro\n(Cán bộ xử lý hồ sơ rủi ro trước FIFO)', '#cccccc', '#000000'),
    ]
    for x, y, w, h, text, facecolor, edgecolor in modules:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02', facecolor=facecolor, edgecolor=edgecolor, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + w / 2, y + h / 2, text, ha='center', va='center', fontsize=8, weight='bold', color='#000000')

    ax.annotate('', xy=(0.49, 0.44), xytext=(0.17, 0.58), arrowprops=dict(arrowstyle='->', lw=1.2, color='#000000'))
    ax.annotate('', xy=(0.49, 0.44), xytext=(0.49, 0.58), arrowprops=dict(arrowstyle='->', lw=1.2, color='#000000'))
    ax.annotate('', xy=(0.49, 0.44), xytext=(0.82, 0.58), arrowprops=dict(arrowstyle='->', lw=1.2, color='#000000'))
    ax.annotate('', xy=(0.35, 0.28), xytext=(0.31, 0.28), arrowprops=dict(arrowstyle='->', lw=1.5, color='#000000'))
    ax.annotate('', xy=(0.67, 0.28), xytext=(0.63, 0.28), arrowprops=dict(arrowstyle='->', lw=1.5, color='#000000'))

    plt.title('Kiến Trúc Phát Hiện Bất Thường Đa Tầng (Isolation Forest + PyTorch Autoencoder + NetworkX)', fontsize=10.5, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, 'hinh2_phat_hien_bat_thuong.png'), dpi=300)
    plt.close()
    print('Generated grayscale images/hinh2_phat_hien_bat_thuong.png')


def create_diagram_experiments(bench):
    pho = bench['phobert_metrics']
    svm = bench['svm_metrics']
    man = bench['manual_vs_automated']['manual_baseline']
    per_f = bench['per_field_comparison']

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.5, 4.4), dpi=300)

    metrics = ['Accuracy', 'Precision', 'Recall', 'F1-score']
    phobert_scores = [pho['accuracy'], pho['precision'], pho['recall'], pho['f1_score']]
    svm_scores = [svm['accuracy'], svm['precision'], svm['recall'], svm['f1_score']]
    manual_scores = [man['accuracy'], man['precision'], man['recall'], man['f1_score']]

    x = np.arange(len(metrics))
    width = 0.26
    b1 = ax1.bar(x - width, phobert_scores, width, label='Fine-tuning PhoBERT (Chính)', color='#222222', edgecolor='#000000')
    b2 = ax1.bar(x, svm_scores, width, label='TF-IDF + LinearSVC (Baseline)', color='#777777', edgecolor='#000000')
    b3 = ax1.bar(x + width, manual_scores, width, label='Phân loại Thủ công', color='#cccccc', edgecolor='#000000')

    for bars in (b1, b2, b3):
        for bar in bars:
            h = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width() / 2, h + 0.6, f'{h:.1f}', ha='center', va='bottom', fontsize=6.5)

    ax1.set_ylabel('Tỷ lệ (%)', fontsize=9, weight='bold')
    ax1.set_title('Đối Sánh Tổng Thể Trên Tập Test Độc Lập (60 Mẫu)', fontsize=9.5, weight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(metrics, fontsize=8.5)
    ax1.set_ylim(60, 102)
    ax1.legend(loc='lower right', fontsize=7.2)
    ax1.grid(axis='y', linestyle='--', alpha=0.5, color='#bbbbbb')

    codes = [r['code'].replace('LV-', '') for r in per_f]
    pho_f1_cls = [r['phobert_f1'] for r in per_f]
    svm_f1_cls = [r['svm_f1'] for r in per_f]
    x2 = np.arange(len(codes))
    w2 = 0.35
    ax2.bar(x2 - w2 / 2, pho_f1_cls, w2, label='Fine-tuning PhoBERT (PyTorch)', color='#333333', edgecolor='#000000')
    ax2.bar(x2 + w2 / 2, svm_f1_cls, w2, label='TF-IDF + LinearSVC', color='#999999', edgecolor='#000000')
    ax2.set_ylabel('F1-score (%)', fontsize=9, weight='bold')
    ax2.set_title('F1-score Theo Từng Lĩnh Vực Chuyên Ngành (Test)', fontsize=9.5, weight='bold')
    ax2.set_xticks(x2)
    ax2.set_xticklabels(codes, fontsize=8.5)
    ax2.set_ylim(55, 108)
    ax2.legend(loc='lower right', fontsize=7.5)
    ax2.grid(axis='y', linestyle='--', alpha=0.5, color='#bbbbbb')

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, 'hinh5_danhgia_hieunang.png'), dpi=300)
    plt.close()
    print('Generated grayscale images/hinh5_danhgia_hieunang.png')


def create_diagram_system_architecture():
    fig, ax = plt.subplots(figsize=(11, 5.0), dpi=300)
    ax.axis('off')

    # Layer 1: Presentation Layer
    rect1 = patches.FancyBboxPatch((0.03, 0.68), 0.94, 0.25, boxstyle='round,pad=0.02', facecolor='#f5f5f5', edgecolor='#000000', linewidth=1.5)
    ax.add_patch(rect1)
    ax.text(0.06, 0.88, 'TẦNG TRÌNH DIỄN (PRESENTATION LAYER - BOOTSTRAP 5 & CHART.JS)', fontsize=9, weight='bold', color='#000000')
    ui_boxes = [
        (0.06, 0.71, 0.19, 0.14, 'Dashboard Thống Kê\n(KPI & Biểu đồ Chart.js)', '#e8e8e8'),
        (0.28, 0.71, 0.20, 0.14, 'Kê Khai 3 Nhóm Hồ Sơ\n(Bài báo, Đề tài, Sáng chế)', '#e8e8e8'),
        (0.51, 0.71, 0.20, 0.14, 'Thẩm Định & 4 Danh Mục\n(5 Bước & CRUD Danh mục)', '#e8e8e8'),
        (0.74, 0.71, 0.20, 0.14, 'Trung Tâm Thực Nghiệm\n(SVM, PhoBERT, Thủ công)', '#e8e8e8'),
    ]
    for x, y, w, h, t, c in ui_boxes:
        r = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.01', facecolor=c, edgecolor='#333333', linewidth=1)
        ax.add_patch(r)
        ax.text(x + w / 2, y + h / 2, t, ha='center', va='center', fontsize=7.5, weight='bold', color='#000000')

    # Layer 2: Business Logic & AI Engine
    rect2 = patches.FancyBboxPatch((0.03, 0.35), 0.94, 0.28, boxstyle='round,pad=0.02', facecolor='#f0f0f0', edgecolor='#000000', linewidth=1.5)
    ax.add_patch(rect2)
    ax.text(0.06, 0.58, 'TẦNG NGHIỆP VỤ & TRÍ TUỆ NHÂN TẠO (PYTHON FLASK & PYTORCH AI ENGINE)', fontsize=9, weight='bold', color='#000000')
    app_boxes = [
        (0.06, 0.38, 0.17, 0.16, 'Flask Application\n- Routing & Session\n- RBAC (3 Roles)\n- Audit Log Bất biến', '#dddddd'),
        (0.26, 0.38, 0.22, 0.16, 'NLP Tiền Xử Lý\n- PyVi ViTokenizer\n- Hợp nhất 5 trường văn bản\n- Lọc từ dừng chuyên ngành', '#dddddd'),
        (0.51, 0.38, 0.21, 0.16, 'AI Phân Loại Đa Nhiệm\n- Chính: PyTorch PhoBERT\n- Đối chứng: TF-IDF + SVM\n- Lĩnh vực & Loại công trình', '#dddddd'),
        (0.75, 0.38, 0.20, 0.16, 'Phát Hiện Bất Thường\n- Đồ thị NetworkX\n- Isolation Forest (iTrees)\n- PyTorch Autoencoder', '#dddddd'),
    ]
    for x, y, w, h, t, c in app_boxes:
        r = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.01', facecolor=c, edgecolor='#333333', linewidth=1)
        ax.add_patch(r)
        ax.text(x + w / 2, y + h / 2, t, ha='center', va='center', fontsize=7.5, weight='bold', color='#000000')

    # Layer 3: Data Access & Storage Layer
    rect3 = patches.FancyBboxPatch((0.03, 0.05), 0.94, 0.25, boxstyle='round,pad=0.02', facecolor='#e8e8e8', edgecolor='#000000', linewidth=1.5)
    ax.add_patch(rect3)
    ax.text(0.06, 0.25, 'TẦNG LƯU TRỮ DỮ LIỆU & TRỌNG SỐ MÔ HÌNH (SQLITE 3 & PYTORCH CHECKPOINTS)', fontsize=9, weight='bold', color='#000000')
    db_boxes = [
        (0.06, 0.08, 0.28, 0.14, 'SQLite 3 (khcn_system.db)\n- HoSoKHCN, LichSuXetDuyet\n- LinhVuc, LoaiCongTrinh\n- TapChiHopLe, DanhMucTacGia', '#d0d0d0'),
        (0.37, 0.08, 0.27, 0.14, 'Model Checkpoints & File\n- phobert_classifier.pt\n- svm_pipeline.joblib\n- Minh chứng PDF/DOCX', '#d0d0d0'),
        (0.67, 0.08, 0.27, 0.14, 'Báo Cáo & Thực Nghiệm\n- Bộ chuẩn 240 mẫu (13 cột)\n- Xuất báo cáo PDF / Excel\n- Lọc Tác giả, Đơn vị, Rủi ro', '#d0d0d0'),
    ]
    for x, y, w, h, t, c in db_boxes:
        r = patches.FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.01', facecolor=c, edgecolor='#333333', linewidth=1)
        ax.add_patch(r)
        ax.text(x + w / 2, y + h / 2, t, ha='center', va='center', fontsize=7.5, weight='bold', color='#000000')

    ax.annotate('', xy=(0.50, 0.68), xytext=(0.50, 0.63), arrowprops=dict(arrowstyle='<->', lw=1.5, color='#000000'))
    ax.annotate('', xy=(0.50, 0.35), xytext=(0.50, 0.30), arrowprops=dict(arrowstyle='<->', lw=1.5, color='#000000'))

    plt.title('Sơ Đồ Kiến Trúc Hệ Thống Tổng Thể 3 Tầng (3-Tier MVC Architecture)', fontsize=11, weight='bold', pad=15)
    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, 'hinh3_kientruc_he_thong.png'), dpi=300)
    plt.close()
    print('Generated grayscale images/hinh3_kientruc_he_thong.png')


if __name__ == '__main__':
    with open(BENCH_JSON, 'r', encoding='utf-8') as f:
        bench_data = json.load(f)
    create_diagram_workflow()
    create_diagram_pipeline(bench_data)
    create_diagram_anomaly()
    create_diagram_experiments(bench_data)
    create_diagram_system_architecture()
