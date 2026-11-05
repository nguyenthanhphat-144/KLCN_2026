# -*- coding: utf-8 -*-
"""
Sinh bộ sơ đồ và biểu đồ học thuật đơn sắc (100% font Times New Roman, 100% tiếng Việt)
phục vụ Báo cáo Khóa luận Tốt nghiệp (Bao_Cao_Khoa_Luan.docx).
Toàn bộ số liệu thực nghiệm được đọc trực tiếp từ docs/Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json.
"""
import os
import json
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMAGES_DIR = os.path.join(BASE_DIR, "images")
BENCH_JSON = os.path.join(BASE_DIR, "docs", "Ket_Qua_Thuc_Nghiem_So_Sanh_2_Mo_Hinh.json")

os.makedirs(IMAGES_DIR, exist_ok=True)

# Bắt buộc 100% chữ trong tất cả sơ đồ/biểu đồ sử dụng font Times New Roman
plt.rcParams["font.family"] = "Times New Roman"
plt.rcParams["font.serif"] = ["Times New Roman"]
plt.rcParams["font.sans-serif"] = ["Times New Roman"]
plt.rcParams["mathtext.fontset"] = "stix"
plt.rcParams["text.color"] = "#000000"
plt.rcParams["axes.labelcolor"] = "#000000"
plt.rcParams["xtick.color"] = "#000000"
plt.rcParams["ytick.color"] = "#000000"


def _draw_box(ax, x, y, w, h, text, fc="#ffffff", ec="#000000", lw=1.2, fs=8.5, bold=False, align="center"):
    rect = patches.FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.015",
        facecolor=fc, edgecolor=ec, linewidth=lw
    )
    ax.add_patch(rect)
    weight = "bold" if bold else "normal"
    tx = x + w / 2 if align == "center" else x + 0.015
    ha = "center" if align == "center" else "left"
    ax.text(tx, y + h / 2, text, ha=ha, va="center", fontsize=fs, weight=weight, color="#000000", family="Times New Roman")


def create_diagram_usecase():
    fig, ax = plt.subplots(figsize=(11.2, 5.4), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    # Khung hệ thống
    sys_box = patches.Rectangle((0.23, 0.06), 0.54, 0.86, facecolor="#fafafa", edgecolor="#000000", linewidth=1.4)
    ax.add_patch(sys_box)
    ax.text(0.50, 0.89, "HỆ THỐNG HỖ TRỢ QUẢN LÝ HỒ SƠ KH&CN TÍCH HỢP TRÍ TUỆ NHÂN TẠO",
            ha="center", va="center", fontsize=9.5, weight="bold", family="Times New Roman")

    # Tác nhân bên trái & bên phải
    actors = [
        (0.03, 0.66, 0.15, 0.14, "Giảng viên /\nTác giả\n(TacGia)"),
        (0.03, 0.22, 0.15, 0.14, "Cán bộ xét duyệt\nKH&CN\n(CanBo)"),
        (0.82, 0.44, 0.15, 0.14, "Quản trị viên\nhệ thống\n(Admin)"),
    ]
    for x, y, w, h, label in actors:
        _draw_box(ax, x, y, w, h, label, fc="#eeeeee", lw=1.3, fs=8.8, bold=True)

    # Các ca sử dụng chính (elip)
    usecases = [
        (0.50, 0.79, 0.42, 0.09, "UC01: Đăng nhập, đổi mật khẩu & phân quyền RBAC"),
        (0.50, 0.67, 0.42, 0.09, "UC02: Kê khai hồ sơ động 3 nhóm & gợi ý PhoBERT trực tiếp"),
        (0.50, 0.55, 0.42, 0.09, "UC03: Tra cứu, lọc đa tiêu chí & xuất báo cáo Excel / PDF"),
        (0.50, 0.43, 0.42, 0.09, "UC04: Xét duyệt tính hợp lệ 5 bước & lưu vết lịch sử"),
        (0.50, 0.31, 0.42, 0.09, "UC05: Xác nhận / hiệu chỉnh nhãn AI & giải trừ cảnh báo"),
        (0.50, 0.19, 0.42, 0.09, "UC06: Phát hiện bất thường đa tầng & hàng đợi ưu tiên"),
        (0.50, 0.09, 0.42, 0.075, "UC07: Quản trị 4 danh mục chuẩn & kiểm định 2 mô hình AI"),
    ]
    for cx, cy, ew, eh, label in usecases:
        el = patches.Ellipse((cx, cy), ew, eh, facecolor="#ffffff", edgecolor="#000000", linewidth=1.1)
        ax.add_patch(el)
        ax.text(cx, cy, label, ha="center", va="center", fontsize=8.0, family="Times New Roman")

    # Đường nối Tác giả
    for uy in [0.79, 0.67, 0.55]:
        ax.plot([0.18, 0.29], [0.73, uy], color="#000000", lw=1.0)
    # Đường nối Cán bộ
    for uy in [0.79, 0.55, 0.43, 0.31, 0.19]:
        ax.plot([0.18, 0.29], [0.29, uy], color="#000000", lw=1.0)
    # Đường nối Admin
    for uy in [0.79, 0.67, 0.55, 0.43, 0.31, 0.19, 0.09]:
        ax.plot([0.82, 0.71], [0.51, uy], color="#000000", lw=1.0)

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh_usecase.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh_usecase.png")


def create_diagram_workflow():
    fig, ax = plt.subplots(figsize=(11.2, 4.2), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    steps = [
        (0.02, 0.34, 0.165, 0.48, "Bước 1: Tiếp nhận\n& Khóa trạng thái\n\nLấy từ Hàng đợi ưu tiên,\ntự chuyển sang\n“Đang xét duyệt”"),
        (0.22, 0.34, 0.165, 0.48, "Bước 2: Thẩm định\n& Tham vấn AI\n\nKiểm tra minh chứng,\nxem dự đoán PhoBERT\nvà cảnh báo bất thường"),
        (0.42, 0.34, 0.165, 0.48, "Bước 3: Quyết định\nthẩm định 3 nhánh\n\n1. Đã duyệt\n2. Từ chối (nêu lý do)\n3. Yêu cầu bổ sung"),
        (0.62, 0.34, 0.165, 0.48, "Bước 4: Lưu vết\nlịch sử bất biến\n\nGhi LichSuXetDuyet\nvà nhật ký phản hồi\nnhãn PhanHoiAI"),
        (0.82, 0.34, 0.165, 0.48, "Bước 5: Bổ sung\n& Nộp lại hồ sơ\n\nTác giả chỉnh sửa,\nhệ thống phân tích lại\nvà về “Chờ xét duyệt”"),
    ]
    for x, y, w, h, text in steps:
        _draw_box(ax, x, y, w, h, text, fc="#f7f7f7", lw=1.3, fs=8.2, bold=False)

    for i in range(len(steps) - 1):
        x1 = steps[i][0] + steps[i][2]
        x2 = steps[i + 1][0]
        ax.annotate("", xy=(x2, 0.58), xytext=(x1, 0.58),
                    arrowprops=dict(arrowstyle="->", lw=1.5, color="#000000"))

    ax.annotate("", xy=(0.10, 0.34), xytext=(0.90, 0.34),
                arrowprops=dict(arrowstyle="->", lw=1.3, color="#222222", connectionstyle="arc3,rad=-0.22", linestyle="--"))
    ax.text(0.50, 0.07, "Vòng lặp nộp lại: Tác giả cập nhật hồ sơ theo yêu cầu bổ sung và quay lại Bước 1",
            ha="center", va="center", fontsize=8.8, weight="bold", family="Times New Roman")

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh1_quytrinh_5buoc.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh1_quytrinh_5buoc.png")


def create_diagram_bce():
    fig, ax = plt.subplots(figsize=(11.2, 4.8), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    cols = [
        (0.03, 0.10, 0.28, 0.80, "NHÓM LỚP BIÊN\n(BOUNDARY CLASSES)", [
            "LoginView (login.html)",
            "DashboardView (dashboard.html)",
            "HoSoFormView (hoso_form.html)",
            "HoSoListView (hoso_list.html)",
            "HoSoDetailView (hoso_detail.html)",
            "CatalogManageView (catalogs.html)",
            "AICenterView (ai_center.html)",
        ]),
        (0.36, 0.10, 0.28, 0.80, "NHÓM LỚP ĐIỀU KHIỂN\n(CONTROL CLASSES)", [
            "AuthRBACController (app.py)",
            "HoSoWorkflowController (app.py)",
            "CatalogCRUDController (app.py)",
            "KHCNAIEngine (ai_engine.py)",
            "PhoBERTClassifier (phobert_finetune.py)",
            "ModelBenchmark (model_benchmark.py)",
            "ReportingService (reporting_service.py)",
        ]),
        (0.69, 0.10, 0.28, 0.80, "NHÓM LỚP THỰC THỂ\n(ENTITY CLASSES)", [
            "NguoiDung (Tài khoản RBAC)",
            "HoSoKHCN (Hồ sơ công trình)",
            "LinhVuc & LoaiCongTrinh",
            "TapChiHopLe & DanhMucTacGia",
            "LichSuXetDuyet (Lưu vết duyệt)",
            "PhanHoiAI (Nhật ký nhãn AI)",
            "ThucNghiemMoHinhAI (Kiểm định)",
        ]),
    ]

    for x, y, w, h, header, items in cols:
        outer = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.015", facecolor="#f8f8f8", edgecolor="#000000", linewidth=1.3)
        ax.add_patch(outer)
        ax.text(x + w / 2, y + h - 0.06, header, ha="center", va="center", fontsize=8.8, weight="bold", family="Times New Roman")
        for idx, item in enumerate(items):
            iy = y + h - 0.16 - idx * 0.095
            _draw_box(ax, x + 0.02, iy, w - 0.04, 0.068, item, fc="#ffffff", lw=0.9, fs=7.9)

    ax.annotate("", xy=(0.36, 0.50), xytext=(0.31, 0.50), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#000000"))
    ax.annotate("", xy=(0.69, 0.50), xytext=(0.64, 0.50), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#000000"))

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh_bce.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh_bce.png")


def create_diagram_system_architecture():
    fig, ax = plt.subplots(figsize=(11.2, 5.0), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    # Tầng 1
    rect1 = patches.FancyBboxPatch((0.03, 0.69), 0.94, 0.25, boxstyle="round,pad=0.015", facecolor="#f8f8f8", edgecolor="#000000", linewidth=1.4)
    ax.add_patch(rect1)
    ax.text(0.05, 0.90, "TẦNG TRÌNH DIỄN (GIAO DIỆN WEB TIẾNG VIỆT – HTML5, BOOTSTRAP 5, CHART.JS)", fontsize=9.0, weight="bold")
    ui_boxes = [
        (0.05, 0.72, 0.20, 0.14, "Tổng quan & Báo cáo\n(Thẻ chỉ số, biểu đồ,\nxuất Excel / PDF)"),
        (0.28, 0.72, 0.21, 0.14, "Kê khai 3 nhóm hồ sơ\n(Bài báo, Đề tài, Sáng chế\nvà gợi ý PhoBERT)"),
        (0.52, 0.72, 0.20, 0.14, "Xét duyệt & 4 Danh mục\n(Quy trình 5 bước,\nthêm/sửa/ẩn danh mục)"),
        (0.75, 0.72, 0.20, 0.14, "Phân tích & Kiểm định\n(Chạy riêng / đối sánh\nSVM, PhoBERT, Rủi ro)"),
    ]
    for x, y, w, h, t in ui_boxes:
        _draw_box(ax, x, y, w, h, t, fc="#ffffff", lw=1.0, fs=8.0)

    # Tầng 2
    rect2 = patches.FancyBboxPatch((0.03, 0.36), 0.94, 0.27, boxstyle="round,pad=0.015", facecolor="#f2f2f2", edgecolor="#000000", linewidth=1.4)
    ax.add_patch(rect2)
    ax.text(0.05, 0.59, "TẦNG XỬ LÝ NGHIỆP VỤ & TRÍ TUỆ NHÂN TẠO (PYTHON FLASK & PYTORCH ENGINE)", fontsize=9.0, weight="bold")
    app_boxes = [
        (0.05, 0.39, 0.18, 0.16, "Điều phối nghiệp vụ\n• Phân quyền 3 vai trò\n• Quản lý xét duyệt\n• Lưu vết kiểm toán"),
        (0.26, 0.39, 0.22, 0.16, "Tiền xử lý tiếng Việt\n• Hợp nhất 5 trường văn bản\n• Tách từ ghép PyVi\n• Lọc từ dừng học thuật"),
        (0.51, 0.39, 0.21, 0.16, "Phân loại đa nhiệm\n• Mô hình chính: PhoBERT\n• Đối chứng: TF-IDF + SVM\n• Học tăng cường phản hồi"),
        (0.75, 0.39, 0.20, 0.16, "Phát hiện bất thường\n• Đồ thị NetworkX\n• Rừng cô lập (iTrees)\n• PyTorch Autoencoder"),
    ]
    for x, y, w, h, t in app_boxes:
        _draw_box(ax, x, y, w, h, t, fc="#ffffff", lw=1.0, fs=8.0)

    # Tầng 3
    rect3 = patches.FancyBboxPatch((0.03, 0.05), 0.94, 0.25, boxstyle="round,pad=0.015", facecolor="#ececec", edgecolor="#000000", linewidth=1.4)
    ax.add_patch(rect3)
    ax.text(0.05, 0.26, "TẦNG LƯU TRỮ DỮ LIỆU & TRỌNG SỐ MÔ HÌNH (SQLITE 3 & PYTORCH CHECKPOINTS)", fontsize=9.0, weight="bold")
    db_boxes = [
        (0.05, 0.08, 0.29, 0.14, "CSDL SQLite 3 (khcn_system.db)\n9 bảng chuẩn hóa 3NF:\nHoSoKHCN, LichSuXetDuyet, 4 Danh mục..."),
        (0.37, 0.08, 0.28, 0.14, "Trọng số & Tệp minh chứng\nphobert_classifier.pt (55.22 MB),\nsvm_pipeline.joblib, thư mục uploads/"),
        (0.68, 0.08, 0.27, 0.14, "Dữ liệu chuẩn & Xuất báo cáo\nBộ 240 mẫu chuẩn hóa 13 trường,\nkết quả kiểm định JSON / Excel"),
    ]
    for x, y, w, h, t in db_boxes:
        _draw_box(ax, x, y, w, h, t, fc="#ffffff", lw=1.0, fs=8.0)

    ax.annotate("", xy=(0.50, 0.69), xytext=(0.50, 0.63), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#000000"))
    ax.annotate("", xy=(0.50, 0.36), xytext=(0.50, 0.30), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#000000"))

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh3_kientruc_he_thong.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh3_kientruc_he_thong.png")


def create_diagram_pipeline(bench):
    pho_f1 = bench["phobert_metrics"]["f1_score"]
    svm_f1 = bench["svm_metrics"]["f1_score"]
    delta_f1 = bench["selected_model"]["delta_f1"]

    fig, ax = plt.subplots(figsize=(11.2, 4.4), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    boxes = [
        (0.02, 0.32, 0.16, 0.38, "Hồ sơ KH&CN\n(13 trường chuẩn:\nTiêu đề, Tóm tắt,\nTừ khóa, Tác giả,\nĐơn vị chủ trì...)", "#ffffff", True),
        (0.21, 0.32, 0.17, 0.38, "Tiền xử lý dùng chung\n(Hợp nhất 5 trường,\ntách từ ghép PyVi,\nlọc từ dừng tiếng Việt)", "#f5f5f5", False),
        (0.42, 0.55, 0.23, 0.35, f"Tinh chỉnh PhoBERT\n(PyTorch vinai/phobert-base)\nF1 Lĩnh vực: {pho_f1:.2f}%\nF1 Loại công trình: 100.0%", "#eaeaea", True),
        (0.42, 0.12, 0.23, 0.35, f"Mô hình đối chứng\nTF-IDF + LinearSVC\nF1 Lĩnh vực: {svm_f1:.2f}%\nF1 Loại công trình: 100.0%", "#f7f7f7", False),
        (0.69, 0.32, 0.14, 0.38, f"Lựa chọn mô hình\nchính: PhoBERT\n(+{delta_f1:.2f}% F1\nso với SVM)", "#e2e2e2", True),
        (0.86, 0.32, 0.12, 0.38, "Đầu ra đa nhiệm\n(6 Lĩnh vực &\n3 Nhóm loại hình)\nkèm độ tin cậy", "#f0f0f0", False),
    ]
    for x, y, w, h, text, fc, is_b in boxes:
        _draw_box(ax, x, y, w, h, text, fc=fc, lw=1.3, fs=8.2, bold=is_b)

    arrows = [
        (0.18, 0.51, 0.21, 0.51),
        (0.38, 0.55, 0.42, 0.72),
        (0.38, 0.47, 0.42, 0.29),
        (0.65, 0.72, 0.69, 0.55),
        (0.65, 0.29, 0.69, 0.47),
        (0.83, 0.51, 0.86, 0.51),
    ]
    for x1, y1, x2, y2 in arrows:
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle="->", lw=1.4, color="#000000"))

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh2_kientruc_ai_pipeline.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh2_kientruc_ai_pipeline.png")


def create_diagram_anomaly():
    fig, ax = plt.subplots(figsize=(11.2, 4.9), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    modules = [
        (0.03, 0.57, 0.28, 0.34, "Đồ thị tri thức NetworkX\nĐỉnh: Hồ sơ, Tác giả, Lĩnh vực, Tạp chí\nĐo độ trung tâm bậc liên kết và\nđộ lệch chuyên môn lịch sử tác giả", "#f8f8f8", False),
        (0.36, 0.57, 0.28, 0.34, "Rừng cô lập (Isolation Forest)\n100 cây cô lập, tỷ lệ nhiễu 0.25\nĐo độ sâu đường đi trung bình\nvà chuẩn hóa điểm về đoạn [0, 1]", "#f2f2f2", False),
        (0.69, 0.57, 0.28, 0.34, "Mạng tự mã hóa PyTorch Autoencoder\nCấu trúc nén: 6 → 16 → 4 → 16 → 6\nTối ưu hàm mất mát MSE (250 chu kỳ)\nphát hiện tương quan phi tuyến bất thường", "#eaeaea", False),
        (0.03, 0.11, 0.28, 0.34, "Kiểm tra trùng lặp & Quy tắc\nĐộ tương đồng Cosine TF-IDF ≥ 62%\nĐối chiếu tạp chí cảnh báo không uy tín\nvà phát hiện khai lệch lĩnh vực AI", "#f8f8f8", False),
        (0.36, 0.11, 0.28, 0.34, "Hợp nhất điểm rủi ro liên tục\nTổ hợp tuyến tính trọng số:\n35% IF + 35% AE + 30% Quy tắc\nNgưỡng bật cờ cảnh báo ≥ 0.50", "#e6e6e6", True),
        (0.69, 0.11, 0.28, 0.34, "Hàng đợi xét duyệt ưu tiên\nSắp xếp hồ sơ chờ xử lý giảm dần\ntheo điểm rủi ro tổng hợp để cán bộ\nthẩm định trước các hồ sơ nguy cơ cao", "#dcdcdc", True),
    ]
    for x, y, w, h, text, fc, b in modules:
        _draw_box(ax, x, y, w, h, text, fc=fc, lw=1.3, fs=8.1, bold=b)

    ax.annotate("", xy=(0.50, 0.45), xytext=(0.17, 0.57), arrowprops=dict(arrowstyle="->", lw=1.2, color="#000000"))
    ax.annotate("", xy=(0.50, 0.45), xytext=(0.50, 0.57), arrowprops=dict(arrowstyle="->", lw=1.2, color="#000000"))
    ax.annotate("", xy=(0.50, 0.45), xytext=(0.83, 0.57), arrowprops=dict(arrowstyle="->", lw=1.2, color="#000000"))
    ax.annotate("", xy=(0.36, 0.28), xytext=(0.31, 0.28), arrowprops=dict(arrowstyle="->", lw=1.4, color="#000000"))
    ax.annotate("", xy=(0.69, 0.28), xytext=(0.64, 0.28), arrowprops=dict(arrowstyle="->", lw=1.4, color="#000000"))

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh2_phat_hien_bat_thuong.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh2_phat_hien_bat_thuong.png")


def create_diagram_erd():
    fig, ax = plt.subplots(figsize=(11.4, 5.4), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    tables = [
        (0.02, 0.68, 0.25, 0.26, "LinhVuc\n─────────────────\nPK maLinhVuc\n   maCode, tenLinhVuc\n   trangThaiHoatDong"),
        (0.02, 0.36, 0.25, 0.26, "LoaiCongTrinh\n─────────────────\nPK maLoai\n   nhomCongTrinh, tenLoai\n   gioChuanQuyDoi"),
        (0.02, 0.04, 0.25, 0.26, "TapChiHopLe\n─────────────────\nPK maTapChi\n   tenTapChi, issn\n   xepHang, trangThaiUyTin"),
        (0.36, 0.26, 0.28, 0.50, "HoSoKHCN (Bảng trung tâm)\n─────────────────────\nPK maHoSo\n   tieuDe, tomTat, tuKhoa\n   thongTinTacGia, donViChuTri\n   namCongBo, kyBaoCao\nFK maLinhVuc, maLoai, maTapChi\nFK nguoiNop\n   maLinhVucAI, doTinCayAI\n   nhomCongTrinhAI, maLinhVucSVM\n   trangThai, diemBatThuong, coCanhBao"),
        (0.73, 0.72, 0.25, 0.23, "NguoiDung\n─────────────────\nPK maNguoiDung\n   hoTen, email, vaiTro\n   donViCongTac"),
        (0.73, 0.45, 0.25, 0.22, "LichSuXetDuyet\n─────────────────\nPK maLichSu\nFK maHoSo, maCanBo\n   trangThaiMoi, yKienDanhGia"),
        (0.73, 0.20, 0.25, 0.20, "PhanHoiAI\n─────────────────\nPK maPhanHoi\nFK maHoSo, maCanBo\n   nhanAIDeXuat, nhanCanBoChon"),
        (0.36, 0.81, 0.28, 0.14, "DanhMucTacGia\n─────────────────────\nPK maTacGia | hoTen, email, orcid\nFK maLinhVucChinh"),
        (0.36, 0.04, 0.28, 0.15, "ThucNghiemMoHinhAI\n─────────────────────\nPK maThucNghiem | tenMoHinh\n   precision, recall, f1Score, laMoHinhChinh"),
    ]
    for x, y, w, h, text in tables:
        is_center = "HoSoKHCN" in text
        _draw_box(ax, x, y, w, h, text, fc="#efefef" if is_center else "#ffffff", lw=1.4 if is_center else 1.1, fs=7.7, bold=is_center)

    # Các đường quan hệ khóa ngoại (1 - n)
    links = [
        ((0.27, 0.80), (0.36, 0.62)),
        ((0.27, 0.49), (0.36, 0.50)),
        ((0.27, 0.17), (0.36, 0.38)),
        ((0.73, 0.82), (0.64, 0.65)),
        ((0.64, 0.54), (0.73, 0.56)),
        ((0.64, 0.40), (0.73, 0.30)),
        ((0.27, 0.86), (0.36, 0.88)),
    ]
    for (x1, y1), (x2, y2) in links:
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle="->", lw=1.1, color="#000000"))

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh_erd.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh_erd.png")


def create_diagram_experiments(bench):
    pho = bench["phobert_metrics"]
    svm = bench["svm_metrics"]
    man = bench["manual_vs_automated"]["manual_baseline"]
    per_f = bench["per_field_comparison"]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.4, 4.4), dpi=300)

    metrics = ["Độ chính xác\n(Accuracy)", "Độ chuẩn xác\n(Precision)", "Độ phủ\n(Recall)", "Điểm F1\n(Macro F1)"]
    phobert_scores = [pho["accuracy"], pho["precision"], pho["recall"], pho["f1_score"]]
    svm_scores = [svm["accuracy"], svm["precision"], svm["recall"], svm["f1_score"]]
    manual_scores = [man["accuracy"], man["precision"], man["recall"], man["f1_score"]]

    x = np.arange(len(metrics))
    width = 0.25
    b1 = ax1.bar(x - width, phobert_scores, width, label="Tinh chỉnh PhoBERT (Mô hình chính)", color="#222222", edgecolor="#000000")
    b2 = ax1.bar(x, svm_scores, width, label="TF-IDF + LinearSVC (Đối chứng)", color="#888888", edgecolor="#000000")
    b3 = ax1.bar(x + width, manual_scores, width, label="Phân loại thủ công", color="#dddddd", edgecolor="#000000")

    for bars in (b1, b2, b3):
        for bar in bars:
            h = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width() / 2, h + 0.6, f"{h:.2f}", ha="center", va="bottom", fontsize=7.0, family="Times New Roman")

    ax1.set_ylabel("Tỷ lệ (%)", fontsize=9.5, weight="bold", family="Times New Roman")
    ax1.set_title("Đối sánh tổng thể trên tập kiểm thử độc lập (60 mẫu)", fontsize=10.0, weight="bold", family="Times New Roman")
    ax1.set_xticks(x)
    ax1.set_xticklabels(metrics, fontsize=8.5, family="Times New Roman")
    ax1.set_ylim(60, 103)
    ax1.legend(loc="lower right", fontsize=7.8)
    ax1.grid(axis="y", linestyle="--", alpha=0.5, color="#bbbbbb")

    codes = [r["code"].replace("LV-", "") for r in per_f]
    pho_f1_cls = [r["phobert_f1"] for r in per_f]
    svm_f1_cls = [r["svm_f1"] for r in per_f]
    x2 = np.arange(len(codes))
    w2 = 0.35
    b4 = ax2.bar(x2 - w2 / 2, pho_f1_cls, w2, label="Tinh chỉnh PhoBERT", color="#333333", edgecolor="#000000")
    b5 = ax2.bar(x2 + w2 / 2, svm_f1_cls, w2, label="TF-IDF + LinearSVC", color="#aaaaaa", edgecolor="#000000")

    for bars in (b4, b5):
        for bar in bars:
            h = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width() / 2, h + 0.6, f"{h:.1f}", ha="center", va="bottom", fontsize=6.8, family="Times New Roman")

    ax2.set_ylabel("Điểm F1 (%)", fontsize=9.5, weight="bold", family="Times New Roman")
    ax2.set_title("Điểm F1 theo từng lĩnh vực chuyên ngành (Tập kiểm thử)", fontsize=10.0, weight="bold", family="Times New Roman")
    ax2.set_xticks(x2)
    ax2.set_xticklabels(codes, fontsize=8.5, family="Times New Roman")
    ax2.set_ylim(55, 108)
    ax2.legend(loc="lower right", fontsize=8.0)
    ax2.grid(axis="y", linestyle="--", alpha=0.5, color="#bbbbbb")

    plt.tight_layout()
    plt.savefig(os.path.join(IMAGES_DIR, "hinh5_danhgia_hieunang.png"), dpi=300, bbox_inches="tight")
    plt.close()
    print("Generated images/hinh5_danhgia_hieunang.png")


if __name__ == "__main__":
    with open(BENCH_JSON, "r", encoding="utf-8") as f:
        bench_data = json.load(f)
    create_diagram_usecase()
    create_diagram_workflow()
    create_diagram_bce()
    create_diagram_system_architecture()
    create_diagram_pipeline(bench_data)
    create_diagram_anomaly()
    create_diagram_erd()
    create_diagram_experiments(bench_data)
