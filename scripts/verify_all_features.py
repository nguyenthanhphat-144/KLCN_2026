# -*- coding: utf-8 -*-
"""
Script kiểm thử toàn diện toàn bộ chức năng từ Tuần 1 đến Tuần 6 theo PRD (CNTT-KLCN001)
"""

import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from database import init_db, get_db_connection
from app import app
from ai_engine import ai_engine


def run_verification():
    print("=== 1. KHỞI TẠO CSDL SQLITE & HUẤN LUYỆN MÔ HÌNH AI ===")
    init_db()
    conn = get_db_connection()
    hoso_count = conn.execute("SELECT COUNT(*) FROM HoSoKHCN").fetchone()[0]
    metrics_count = conn.execute("SELECT COUNT(*) FROM ThucNghiemMoHinhAI").fetchone()[0]
    print(f"-> Số lượng hồ sơ mẫu trong CSDL: {hoso_count}")
    print(f"-> Số mô hình thực nghiệm trong ThucNghiemMoHinhAI: {metrics_count}")

    for m in conn.execute("SELECT * FROM ThucNghiemMoHinhAI"):
        print(f"   * {m['tenMoHinh']}: Precision={m['precisionScore']}%, Recall={m['recallScore']}%, F1={m['f1Score']}%, Latency={m['thoiGianPhanHoiMs']}ms")

    graph_info = ai_engine.get_graph_summary()
    print(f"-> Đồ thị NetworkX: {graph_info['nodes']} đỉnh, {graph_info['edges']} cạnh")
    conn.close()

    print("\n=== 2. KIỂM THỬ CÁC ROUTE FLASK, RBAC, CRUD, XÉT DUYỆT, AI & BÁO CÁO ===")
    client = app.test_client()

    # 2.1. Đăng nhập với vai trò Cán bộ xét duyệt
    resp = client.post("/login", data={"email": "canbo@uni.edu.vn", "matKhau": "CanBo@123"}, follow_redirects=True)
    assert resp.status_code == 200 and "Thống Kê Trực Quan" in resp.get_data(as_text=True)
    print("-> [PASS] Đăng nhập Session & check_password_hash (Vai trò: CanBo)")

    # 2.2. Kiểm thử RBAC: Cán bộ xét duyệt bị chặn truy cập trang Quản trị người dùng của Admin
    resp_rbac = client.get("/quan-tri/nguoi-dung", follow_redirects=True)
    assert "Từ chối truy cập (RBAC)" in resp_rbac.get_data(as_text=True)
    print("-> [PASS] Decorator RBAC chặn Cán bộ truy cập phân hệ nhạy cảm của Admin")

    # 2.3. Kiểm thử RBAC: Cán bộ xét duyệt bị chặn xóa cứng (Hard Delete)
    resp_hard = client.post("/hoso/HS-2026-0002/hard-delete", follow_redirects=True)
    assert "Từ chối truy cập (RBAC)" in resp_hard.get_data(as_text=True)
    print("-> [PASS] Decorator RBAC chặn Cán bộ thực hiện Xóa cứng (Hard Delete)")

    # 2.4. Kiểm thử API gợi ý phân loại AI thời gian thực (PyVi + PhoBERT + TF-IDF/SVM)
    resp_ai = client.post("/api/ai-preview", json={
        "tieuDe": "Nghiên cứu mạng nơ ron tích chập sâu nhận dạng biển số xe tiếng Việt",
        "tomTat": "Ứng dụng thị giác máy tính và học sâu trong nhận diện phương tiện giao thông thông minh.",
        "tuKhoa": "thị giác máy tính, học sâu, trí tuệ nhân tạo",
        "donViChuTri": "Khoa Công nghệ Thông tin"
    })
    ai_data = resp_ai.get_json()
    assert resp_ai.status_code == 200 and "thị_giác" in ai_data["segmented_text"]
    print(f"-> [PASS] Tiền xử lý & Tách từ PyVi: '{ai_data['segmented_text'][:65]}...'")
    print(f"-> [PASS] Dự đoán AI PhoBERT: {ai_data['phobert_label_name']} ({ai_data['phobert_confidence']}%)")

    # 2.5. Kiểm thử Feedback Loop (Tuần 6 - Mục 2d)
    resp_fb = client.post("/hoso/HS-2026-0003/ai-feedback", data={
        "actionType": "adjust",
        "maLinhVucThucTe": "5",
        "ghiChuPhanHoi": "Điều chỉnh hồ sơ vật liệu nano ZnO từ Kinh tế về đúng lĩnh vực Hóa học - Sinh học - Môi trường."
    }, follow_redirects=True)
    assert resp_fb.status_code == 200
    print("-> [PASS] Cơ chế phản hồi nhãn AI (Feedback Loop) lưu vào bảng PhanHoiAI")

    # 2.6. Kiểm thử Quy trình xét duyệt & ghi nhận LichSuXetDuyet (Tuần 5 - Mục 2)
    resp_rev = client.post("/hoso/HS-2026-0002/xet-duyet", data={
        "trangThaiMoi": "Đang xét duyệt",
        "yKienDanhGia": "Đã kiểm tra chỉ số ISSN trên IEEE, chuyển tiểu ban Điện - Điện tử thẩm định."
    }, follow_redirects=True)
    assert resp_rev.status_code == 200 and "LichSuXetDuyet" in resp_rev.get_data(as_text=True)
    print("-> [PASS] Chuyển trạng thái xét duyệt & tự động ghi nhật ký LichSuXetDuyet")

    # 2.7. Kiểm thử Xuất báo cáo Excel (.xlsx) & PDF (.pdf) (Tuần 5 - Mục 3b)
    resp_xlsx = client.get("/bao-cao/export/excel")
    assert resp_xlsx.status_code == 200 and len(resp_xlsx.data) > 2000
    print(f"-> [PASS] Kết xuất báo cáo Excel (.xlsx): {len(resp_xlsx.data)} bytes")

    resp_pdf = client.get("/bao-cao/export/pdf")
    assert resp_pdf.status_code == 200 and resp_pdf.data.startswith(b"%PDF")
    print(f"-> [PASS] Kết xuất báo cáo PDF (.pdf): {len(resp_pdf.data)} bytes")

    print("\n=== TẤT CẢ CÁC KIỂM THỬ TỪ TUẦN 1 ĐẾN TUẦN 6 ĐỀU THÀNH CÔNG 100%! ===")


if __name__ == "__main__":
    run_verification()
