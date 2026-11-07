# -*- coding: utf-8 -*-
"""
Kịch bản kiểm thử độc lập: Pipeline Trí tuệ Nhân tạo & Phát hiện Bất thường (CNTT-KLCN001)
Tác giả kiểm thử: Lê Thanh Phong (kbt2981-spec)
Mục đích:
- Kiểm thử tiền xử lý tiếng Việt (chuẩn hóa Unicode, tách từ ghép chuyên ngành với PyVi).
- Kiểm thử suy luận phân loại tự động hồ sơ KH&CN (TF-IDF + SVM).
- Kiểm thử cấu trúc đồ thị quan hệ học thuật NetworkX và thuật toán Isolation Forest.
"""

import os
import sys

os.environ["HF_HUB_OFFLINE"] = "1"
os.environ["TRANSFORMERS_OFFLINE"] = "1"

import time
import sqlite3
import numpy as np
import networkx as nx
from sklearn.ensemble import IsolationForest

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "khcn_system.db")
sys.path.insert(0, BASE_DIR)

from ai_engine import ai_engine


def test_ai_pipeline_and_anomaly_detection():
    print("=" * 65)
    print(" [TEST-2] KIỂM THỬ PIPELINE AI & MÔ HÌNH PHÁT HIỆN BẤT THƯỜNG")
    print("=" * 65)

    # 1. Kiểm thử tiền xử lý văn bản tiếng Việt
    print("[1] Kiểm thử Tiền xử lý & Tách từ ghép tiếng Việt (PyVi):")
    sample_title = "Nghiên cứu ứng dụng trí tuệ nhân tạo và học sâu trong an toàn thông tin mạng máy tính"
    sample_abstract = "Đề tài chế tạo mô hình mạng nơ ron tích chập sâu nhằm phân tích và nhận dạng xâm nhập bất thường."
    sample_keywords = "trí tuệ nhân tạo, học sâu, an toàn thông tin, mạng nơ ron"

    segmented = ai_engine.preprocess_record(sample_title, sample_abstract, sample_keywords, tac_gia="Lê Thanh Phong")

    print(f"    * Văn bản sau tiền xử lý & tách từ PyVi:\n      '{segmented[:90]}...'")
    assert "trí_tuệ_nhân_tạo" in segmented or "học_sâu" in segmented or "an_toàn_thông_tin" in segmented, \
        "Lỗi: Không nhận diện được từ ghép chuyên ngành tiếng Việt"
    print("    -> [PASS] Tách từ ghép chuyên ngành tiếng Việt thành công!")

    # 2. Kiểm thử tải mô hình và suy luận phân loại
    print("\n[2] Kiểm thử Tải mô hình & Suy luận Phân loại Tự động:")
    ai_engine.train_and_evaluate(mode="load_or_init")

    t0 = time.perf_counter()
    pred = ai_engine.predict_record_field(sample_title, sample_abstract, sample_keywords, tac_gia="Lê Thanh Phong")
    latency_ms = round((time.perf_counter() - t0) * 1000.0, 2)

    print(f"    * Lĩnh vực dự đoán (SVM ID): {pred['svm_label_id']} (Độ tin cậy: {pred['svm_confidence']}%)")
    print(f"    * Nhóm công trình dự đoán: {pred.get('work_type_svm', 'N/A')} ({pred.get('work_type_svm_conf', 0)}%)")
    print(f"    * Thời gian phản hồi thực tế: {latency_ms} ms")
    assert pred["svm_label_id"] in range(1, 7), "Lỗi: ID lĩnh vực nằm ngoài phạm vi 1..6"
    assert latency_ms < 500, f"Độ trễ quá cao: {latency_ms} ms"
    print("    -> [PASS] Phân loại hồ sơ hợp lệ với độ trễ tối ưu!")

    # 3. Kiểm thử Đồ thị tri thức NetworkX
    print("\n[3] Kiểm thử Xây dựng Đồ thị quan hệ học thuật NetworkX:")
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT maHoSo, tieuDe, maLinhVuc, thongTinTacGia, maTapChi FROM HoSoKHCN WHERE isDeleted = 0;")
    rows = cur.fetchall()
    conn.close()

    G = nx.Graph()
    for r in rows:
        work_node = f"HS:{r['maHoSo']}"
        field_node = f"LV:{r['maLinhVuc']}"
        G.add_node(work_node, type="Hồ sơ", title=r["tieuDe"])
        G.add_node(field_node, type="Lĩnh vực")
        G.add_edge(work_node, field_node, relation="thuoc_linh_vuc")

        for auth in str(r["thongTinTacGia"] or "").split(","):
            auth_clean = auth.strip()
            if auth_clean:
                auth_node = f"TG:{auth_clean}"
                G.add_node(auth_node, type="Tác giả")
                G.add_edge(work_node, auth_node, relation="la_tac_gia")

    print(f"    * Tổng số đỉnh trên đồ thị NetworkX: {G.number_of_nodes()} đỉnh")
    print(f"    * Tổng số cạnh quan hệ liên kết: {G.number_of_edges()} cạnh")
    assert G.number_of_nodes() > 0, "Đồ thị tri thức không có đỉnh nào"
    assert G.number_of_edges() > 0, "Đồ thị tri thức không có cạnh liên kết"
    print("    -> [PASS] Cấu trúc đồ thị quan hệ Tác giả - Công trình - Lĩnh vực toàn vẹn!")

    # 4. Kiểm thử Thuật toán Phát hiện Bất thường Isolation Forest
    print("\n[4] Kiểm thử Thuật toán Cô lập Bất thường Isolation Forest:")
    dummy_features = np.array([
        [0.1, 0.2, 0.15, 0.05, 0.1, 0.2],
        [0.12, 0.18, 0.14, 0.06, 0.11, 0.19],
        [0.95, 0.88, 0.92, 0.85, 0.99, 0.91],  # Điểm bất thường (outlier)
        [0.11, 0.21, 0.16, 0.04, 0.12, 0.18],
    ], dtype=np.float32)

    iso = IsolationForest(n_estimators=50, contamination=0.25, random_state=42)
    iso.fit(dummy_features)
    preds = iso.predict(dummy_features)  # 1: bình thường, -1: bất thường
    print(f"    * Nhãn dự đoán (1: Hợp lệ, -1: Bất thường): {preds.tolist()}")
    assert preds[2] == -1, "Isolation Forest không nhận diện được điểm dị biệt rõ rệt"
    print("    -> [PASS] Thuật toán Isolation Forest cô lập thành công hồ sơ bất thường!")

    print("\n==> KẾT QUẢ [TEST-2]: TOÀN BỘ PIPELINE AI ĐÃ ĐẠT 100% PASS! <==\n")


if __name__ == "__main__":
    test_ai_pipeline_and_anomaly_detection()
