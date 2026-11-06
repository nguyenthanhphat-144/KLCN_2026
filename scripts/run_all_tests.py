# -*- coding: utf-8 -*-
"""
[TEST SUITE] Trình chạy tự động toàn bộ 5 kịch bản kiểm thử độc lập
==================================================================
1. test_db_integrity.py        - Kiểm tra tính toàn vẹn 8 bảng CSDL, quan hệ khóa ngoại
2. test_ai_inference.py        - Kiểm tra pipeline AI (PyVi, TF-IDF+SVM, PhoBERT, NetworkX, Isolation Forest)
3. test_report_generation.py   - Kiểm tra kết xuất báo cáo động Excel (.xlsx) & PDF (.pdf)
4. test_rbac_access_control.py - Kiểm tra cơ chế phân quyền RBAC (Admin, CanBo, TacGia)
5. test_hoso_workflow.py       - Kiểm tra quy trình 5 trạng thái hồ sơ KH&CN & lịch sử xét duyệt

Tất cả các bài kiểm thử đều hoạt động ở chế độ đọc an toàn (Read-Only),
không làm biến đổi bất kỳ dữ liệu đang chạy nào của hệ thống.
"""
import os
import sys
import io
import time
import subprocess

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_EXE = sys.executable

TEST_SCRIPTS = [
    ("TEST-1", "test_db_integrity.py", "Kiểm thử toàn vẹn CSDL (8 bảng, danh mục, RBAC)"),
    ("TEST-2", "test_ai_inference.py", "Kiểm thử Pipeline AI (PyVi, TF-IDF+SVM, PhoBERT, Isolation Forest)"),
    ("TEST-3", "test_report_generation.py", "Kiểm thử Kết xuất Báo cáo Động (Excel & PDF)"),
    ("TEST-4", "test_rbac_access_control.py", "Kiểm thử Phân quyền RBAC (Schema & Decorators)"),
    ("TEST-5", "test_hoso_workflow.py", "Kiểm thử Quy trình Xét duyệt Hồ sơ (5 trạng thái & Log)"),
]

def main():
    print("=" * 72)
    print(" BỘ KIỂM THỬ TỰ ĐỘNG TOÀN DIỆN HỆ THỐNG QUẢN LÝ HỒ SƠ KH&CN (KLTN 2026)")
    print("=" * 72)
    start_total_time = time.time()
    results = []

    for idx, (test_id, filename, desc) in enumerate(TEST_SCRIPTS, 1):
        script_path = os.path.join(SCRIPTS_DIR, filename)
        print(f"\n[{idx}/5] Đang thực thi {test_id}: {desc}...")
        print("-" * 72)
        t0 = time.time()
        
        try:
            p = subprocess.run(
                [PYTHON_EXE, "-X", "utf8", script_path],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=60
            )
            elapsed = time.time() - t0
            passed = (p.returncode == 0)
            results.append((test_id, filename, desc, passed, elapsed, p.stdout.strip()))
            
            # In dòng kết quả chính
            lines = p.stdout.strip().splitlines()
            summary_line = [l for l in lines if "==>" in l or "PASS" in l]
            if summary_line:
                print(f"    Kết quả: {summary_line[-1]}")
            print(f"    Thời gian chạy: {elapsed:.2f}s | Trạng thái: {'[OK - ĐẠT]' if passed else '[FAIL - LỖI]'}")
        except subprocess.TimeoutExpired:
            results.append((test_id, filename, desc, False, 60.0, "Timeout > 60s"))
            print(f"    -> [TIMEOUT] Quá thời gian thực thi 60s!")
        except Exception as e:
            results.append((test_id, filename, desc, False, 0.0, str(e)))
            print(f"    -> [ERROR] Lỗi thực thi: {e}")

    total_elapsed = time.time() - start_total_time
    all_passed = all(r[3] for r in results)

    print("\n" + "=" * 72)
    print(" BẢNG TỔNG HỢP KẾT QUẢ KIỂM THỬ HỆ THỐNG")
    print("=" * 72)
    print(f"{'Mã':<8} | {'Tên Kịch Bản':<28} | {'Thời Gian':<10} | {'Kết Quả'}")
    print("-" * 72)
    for test_id, filename, desc, passed, elapsed, _ in results:
        status_str = "PASS (100%)" if passed else "FAIL"
        print(f"{test_id:<8} | {filename:<28} | {elapsed:>8.2f}s | {status_str}")
    print("-" * 72)
    print(f"Tổng thời gian kiểm thử: {total_elapsed:.2f} giây")
    print(f"Tổng số kịch bản: {len(results)} | Đạt: {sum(1 for r in results if r[3])} | Không đạt: {sum(1 for r in results if not r[3])}")
    print("=" * 72)

    if all_passed:
        print("==> CHÚC MỪNG: TOÀN BỘ 5/5 KỊCH BẢN KIỂM THỬ ĐÃ ĐẠT TIÊU CHUẨN 100%! <==")
    else:
        print("==> CẢNH BÁO: MỘT SỐ KỊCH BẢN KIỂM THỬ CHƯA ĐẠT! VUI LÒNG KIỂM TRA LẠI <==")
    print("=" * 72)

    sys.exit(0 if all_passed else 1)

if __name__ == "__main__":
    main()
