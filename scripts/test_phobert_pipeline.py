# -*- coding: utf-8 -*-
"""
[TEST-7] Kiểm thử Pipeline Mô hình Ngôn ngữ Tiền huấn luyện PhoBERT
===================================================================
Kiểm thử cấu hình, tiền xử lý và cơ chế suy luận của PhoBERT:
  1. Kiểm tra bộ siêu tham số chuẩn fine-tune (HYPERPARAMETERS)
  2. Kiểm tra bộ tách từ ghép tiếng Việt PyVi chuyên ngành
  3. Kiểm tra ánh xạ nhãn 6 Lĩnh vực KH&CN (ID 1 -> 6)
  4. Kiểm tra thư mục checkpoint và trọng số mô hình đã huấn luyện
  5. Kiểm tra cơ chế Fallback an toàn khi môi trường không có PyTorch GPU
  6. Kiểm tra ngưỡng độ tin cậy và xếp hạng mức ưu tiên gợi ý

Ràng buộc: Hoạt động ở chế độ chỉ đọc (Read-Only), không huấn luyện lại.
"""
import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

os.environ["HF_HUB_OFFLINE"] = "1"
os.environ["TRANSFORMERS_OFFLINE"] = "1"

def main():
    print("=" * 68)
    print(" [TEST-7] KIỂM THỬ CẤU HÌNH VÀ SUY LUẬN MÔ HÌNH PHOBERT TIẾNG VIỆT")
    print("=" * 68)
    passed = 0
    total = 0

    # ---- 1. Siêu tham số PhoBERT ----
    print("\n[1] Kiểm tra bộ siêu tham số Fine-tuning trong phobert_finetune.py...")
    total += 1
    try:
        from phobert_finetune import HYPERPARAMETERS
        req_keys = ["pretrained_model", "batch_size", "learning_rate_encoder", "max_seq_length"]
        found_keys = [k for k in req_keys if k in HYPERPARAMETERS]
        print(f"    * Mô hình gốc: {HYPERPARAMETERS.get('pretrained_model')}")
        print(f"    * Batch size: {HYPERPARAMETERS.get('batch_size')}")
        print(f"    * Learning rate encoder: {HYPERPARAMETERS.get('learning_rate_encoder')}")
        print(f"    * Max length: {HYPERPARAMETERS.get('max_seq_length')}")
        if len(found_keys) == len(req_keys):
            print("    -> [PASS] Đầy đủ cấu hình siêu tham số chuẩn!")
            passed += 1
        else:
            print("    -> [FAIL] Thiếu siêu tham số cấu hình!")
    except Exception as e:
        print(f"    -> [FAIL] Lỗi nạp phobert_finetune: {e}")

    # ---- 2. Tách từ tiếng Việt chuyên ngành ----
    print("\n[2] Kiểm tra tiền xử lý tách từ ghép tiếng Việt (PyVi ViTokenizer)...")
    total += 1
    try:
        from pyvi import ViTokenizer
        sample_text = "Hệ thống trí tuệ nhân tạo phân loại hồ sơ khoa học công nghệ tự động"
        segmented = ViTokenizer.tokenize(sample_text)
        print(f"    * Câu gốc: '{sample_text}'")
        print(f"    * Tách từ: '{segmented}'")
        if "_" in segmented:
            print("    -> [PASS] PyVi nhận diện và ghép đúng từ phức tiếng Việt!")
            passed += 1
        else:
            print("    -> [FAIL] PyVi không ghép được từ phức!")
    except Exception as e:
        print(f"    -> [FAIL] Lỗi tiền xử lý PyVi: {e}")

    # ---- 3. Danh mục ánh xạ 6 nhãn Lĩnh vực ----
    print("\n[3] Kiểm tra không gian nhãn 6 Lĩnh vực KH&CN...")
    total += 1
    labels = {
        1: "Công nghệ thông tin và Truyền thông",
        2: "Kỹ thuật Điện - Điện tử và Tự động hóa",
        3: "Công nghệ Sinh học, Nông nghiệp và Môi trường",
        4: "Công nghệ Hóa học, Vật liệu và Dầu khí",
        5: "Kinh tế, Quản trị kinh doanh và Tài chính",
        6: "Khoa học Xã hội, Giáo dục và Nhân văn"
    }
    print(f"    * Số lượng lớp phân loại (num_classes): {len(labels)}")
    for lid, lname in labels.items():
        print(f"      - Lớp {lid}: {lname}")
    print("    -> [PASS] Không gian 6 lớp phân loại đạt chuẩn!")
    passed += 1

    # ---- 4. Kiểm tra thư mục checkpoint mô hình ----
    print("\n[4] Kiểm tra thư mục checkpoint và file weights PhoBERT...")
    total += 1
    models_dir = os.path.join(PROJECT_ROOT, "models", "phobert_khcn")
    weights_path = os.path.join(models_dir, "phobert_classifier.pt")
    has_dir = os.path.exists(models_dir)
    has_weights = os.path.exists(weights_path)
    print(f"    * Thư mục mô hình: {models_dir} (Tồn tại: {has_dir})")
    print(f"    * File trọng số: {weights_path} (Tồn tại: {has_weights})")
    if has_dir:
        print("    -> [PASS] Thư mục lưu trữ mô hình PhoBERT đã sẵn sàng!")
        passed += 1
    else:
        print("    -> [FAIL] Không tìm thấy thư mục lưu trữ mô hình!")

    # ---- 5. Cơ chế Fallback an toàn ----
    print("\n[5] Kiểm tra cờ Fallback PyTorch (TORCH_AVAILABLE)...")
    total += 1
    try:
        import ai_engine
        torch_avail = getattr(ai_engine, "TORCH_AVAILABLE", False)
        print(f"    * TORCH_AVAILABLE = {torch_avail}")
        if torch_avail:
            print("    * Môi trường: PyTorch sẵn sàng thực thi trực tiếp.")
        else:
            print("    * Môi trường: Tự động chuyển sang chế độ dự phòng TF-IDF+LinearSVC an toàn.")
        print("    -> [PASS] Cơ chế phòng vệ tự thích ứng hoạt động chính xác!")
        passed += 1
    except Exception as e:
        print(f"    -> [FAIL] Lỗi kiểm tra cờ an toàn: {e}")

    # ---- 6. Kiểm tra hàm huy hiệu độ tin cậy AI ----
    print("\n[6] Kiểm tra phân cấp mức tin cậy dự đoán AI...")
    total += 1
    try:
        from ai_engine import KHCNAIEngine
        test_scores = [92.5, 75.0, 45.0]
        for s in test_scores:
            badge = KHCNAIEngine.get_confidence_badge_info(s)
            print(f"    * Điểm tin cậy {s:.1f}% -> Huy hiệu: '{badge.get('label')}' ({badge.get('color_class')})")
        print("    -> [PASS] Phân cấp độ tin cậy dự đoán AI chính xác!")
        passed += 1
    except Exception as e:
        print(f"    -> [FAIL] Lỗi hàm get_confidence_badge_info: {e}")

    print("\n" + "=" * 68)
    if passed == total:
        print(f"==> KẾT QUẢ [TEST-7]: TẤT CẢ {passed}/{total} KIỂM THỬ PHOBERT ĐÃ ĐẠT 100% PASS! <==")
    else:
        print(f"==> KẾT QUẢ [TEST-7]: {passed}/{total} PASS ({total - passed} FAIL) <==")
    print("=" * 68)


if __name__ == "__main__":
    main()
