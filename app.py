# -*- coding: utf-8 -*-
"""
Ứng dụng Web Chính (Flask Backend & Controllers)
"""

import os
import json
import math
from functools import wraps
from datetime import datetime
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, jsonify, send_file, abort
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

from database import get_db_connection, init_db
from ai_engine import ai_engine
from phobert_finetune import HYPERPARAMETERS as PHOBERT_HYPERPARAMS
from reporting_service import generate_excel_report, generate_pdf_report

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
ALLOWED_EXTENSIONS = {"pdf", "doc", "docx", "zip", "png", "jpg"}

app = Flask(__name__)
app.secret_key = "kltn-2026-cntt-klcn001-secret-key-secure"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# =============================================================================
# TUẦN 4 - MỤC 1: XÁC THỰC (SESSION) VÀ PHÂN QUYỀN DỰA TRÊN VAI TRÒ (RBAC)
# =============================================================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Vui lòng đăng nhập để truy cập hệ thống.", "warning")
            return redirect(url_for("login", next=request.url))
        return f(*args, **kwargs)
    return decorated_function


def role_required(*allowed_roles):
    """
    Decorator kiểm tra vai trò người dùng (RBAC - Tuần 4 Mục 1b):
    Ngăn chặn Cán bộ xét duyệt hoặc Tác giả truy cập các phân hệ quản trị nhạy cảm
    hoặc thao tác xóa vĩnh viễn dữ liệu.
    """
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if "user_id" not in session:
                flash("Vui lòng đăng nhập để tiếp tục.", "warning")
                return redirect(url_for("login"))
            user_role = session.get("vai_tro")
            if user_role not in allowed_roles:
                flash(
                    f"Từ chối truy cập (RBAC): Vai trò '{user_role}' không có quyền thực hiện thao tác này. "
                    f"Yêu cầu quyền: {', '.join(allowed_roles)}.",
                    "danger"
                )
                return redirect(url_for("dashboard"))
            return f(*args, **kwargs)
        return decorated_function
    return decorator


@app.context_processor
def inject_globals():
    pending_corrections = 0
    pending_reviews = 0
    user_id = session.get("user_id")
    vai_tro = session.get("vai_tro")

    if user_id:
        try:
            conn = get_db_connection()
            cur = conn.cursor()
            if vai_tro == "TacGia":
                cur.execute(
                    "SELECT COUNT(*) FROM HoSoKHCN WHERE maNguoiNop = ? AND trangThai = ? AND isDeleted = 0",
                    (user_id, "Yêu cầu bổ sung")
                )
                pending_corrections = cur.fetchone()[0]
            elif vai_tro in ["CanBo", "Admin"]:
                cur.execute(
                    "SELECT COUNT(*) FROM HoSoKHCN WHERE trangThai = ? AND isDeleted = 0",
                    ("Chờ xét duyệt",)
                )
                pending_reviews = cur.fetchone()[0]
            conn.close()
        except Exception:
            pass

    return {
        "current_user": {
            "id": session.get("user_id"),
            "ho_ten": session.get("ho_ten"),
            "email": session.get("email"),
            "vai_tro": session.get("vai_tro"),
            "don_vi": session.get("don_vi"),
            "hoc_vi": session.get("hoc_vi"),
        },
        "get_confidence_badge": ai_engine.get_confidence_badge_info,
        "now_year": 2026,
        "pending_corrections": pending_corrections,
        "pending_reviews": pending_reviews,
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    """Tuần 4 - Mục 1a: Đăng nhập, kiểm tra check_password_hash và khởi tạo Session."""
    if "user_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        email = request.form.get("email", "").strip()
        mat_khau = request.form.get("matKhau", "")

        conn = get_db_connection()
        user = conn.execute(
            "SELECT * FROM NguoiDung WHERE email = ?", (email,)
        ).fetchone()
        conn.close()

        if user is None:
            flash("Email không tồn tại trong hệ thống.", "danger")
        elif user["trangThai"] == 0:
            flash("Tài khoản của bạn đã bị vô hiệu hóa. Vui lòng liên hệ Quản trị viên.", "danger")
        elif not check_password_hash(user["matKhau"], mat_khau):
            flash("Mật khẩu không chính xác. Vui lòng kiểm tra lại.", "danger")
        else:
            session.clear()
            session["user_id"] = user["maNguoiDung"]
            session["ho_ten"] = user["hoTen"]
            session["email"] = user["email"]
            session["vai_tro"] = user["vaiTro"]
            session["don_vi"] = user["donViCongTac"]
            session["hoc_vi"] = user["hocHamHocVi"]
            flash(f"Chào mừng {user['hoTen']} ({user['vaiTro']}) đã đăng nhập thành công!", "success")
            return redirect(url_for("dashboard"))

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Đã đăng xuất khỏi phiên làm việc an toàn.", "info")
    return redirect(url_for("login"))


@app.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    """Tuần 4 - Mục 1c: Đổi mật khẩu an toàn sau khi xác thực mật khẩu hiện tại."""
    if request.method == "POST":
        mk_hien_tai = request.form.get("matKhauHienTai", "")
        mk_moi = request.form.get("matKhauMoi", "")
        mk_xac_nhan = request.form.get("xacNhanMatKhauMoi", "")

        conn = get_db_connection()
        user = conn.execute(
            "SELECT * FROM NguoiDung WHERE maNguoiDung = ?", (session["user_id"],)
        ).fetchone()

        if not check_password_hash(user["matKhau"], mk_hien_tai):
            conn.close()
            flash("Mật khẩu hiện tại không chính xác!", "danger")
            return redirect(url_for("change_password"))

        if len(mk_moi) < 6:
            conn.close()
            flash("Mật khẩu mới phải có tối thiểu 6 ký tự.", "warning")
            return redirect(url_for("change_password"))

        if mk_moi != mk_xac_nhan:
            conn.close()
            flash("Mật khẩu mới và xác nhận mật khẩu không khớp.", "warning")
            return redirect(url_for("change_password"))

        new_hash = generate_password_hash(mk_moi)
        conn.execute(
            "UPDATE NguoiDung SET matKhau = ? WHERE maNguoiDung = ?",
            (new_hash, session["user_id"])
        )
        conn.commit()
        conn.close()
        flash("Cập nhật mật khẩu mới thành công!", "success")
        return redirect(url_for("dashboard"))

    return render_template("change_password.html")


# =============================================================================
# TUẦN 5 - MỤC 3A: DASHBOARD THỐNG KÊ TRỰC QUAN
# =============================================================================
@app.route("/")
@app.route("/dashboard")
@login_required
def dashboard():
    conn = get_db_connection()
    cur = conn.cursor()

    # Tổng quan KPI
    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0")
    total_hoso = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0 AND trangThai = ?", ("Chờ xét duyệt",))
    cho_duyet = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0 AND trangThai = ?", ("Đang xét duyệt",))
    dang_duyet = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0 AND trangThai = ?", ("Đã duyệt",))
    da_duyet = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0 AND trangThai = ?", ("Từ chối",))
    tu_choi = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0 AND trangThai = ?", ("Yêu cầu bổ sung",))
    bo_sung = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM HoSoKHCN WHERE isDeleted = 0 AND coCanhBao = 1")
    canh_bao_count = cur.fetchone()[0]

    # Thống kê theo Lĩnh vực
    cur.execute("""
        SELECT lv.tenLinhVuc, lv.maCode, COUNT(h.maHoSo) AS soLuong
        FROM LinhVuc lv
        LEFT JOIN HoSoKHCN h ON lv.maLinhVuc = h.maLinhVuc AND h.isDeleted = 0
        GROUP BY lv.maLinhVuc, lv.tenLinhVuc, lv.maCode
        ORDER BY soLuong DESC
    """)
    stats_linhvuc = [dict(r) for r in cur.fetchall()]

    # Thống kê theo Loại công trình
    cur.execute("""
        SELECT lct.tenLoai, lct.nhomCongTrinh, COUNT(h.maHoSo) AS soLuong,
               SUM(CASE WHEN h.trangThai = 'Đã duyệt' THEN lct.gioChuanQuyDoi ELSE 0 END) AS tongGioChuan
        FROM LoaiCongTrinh lct
        LEFT JOIN HoSoKHCN h ON lct.maLoai = h.maLoai AND h.isDeleted = 0
        GROUP BY lct.maLoai, lct.tenLoai, lct.nhomCongTrinh
        ORDER BY soLuong DESC
    """)
    stats_loai = [dict(r) for r in cur.fetchall()]

    # Thống kê chuyên sâu nhóm BÀI BÁO KHOA HỌC (Revision)
    cur.execute("""
        SELECT namCongBo, COUNT(*) AS soLuong
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Bài báo%' OR lct.tenLoai LIKE '%Bài báo%')
        GROUP BY namCongBo
        ORDER BY namCongBo DESC
    """)
    stats_baibao_nam = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT COALESCE(tc.xepHang, h.chiMucTapChi, N'Chưa phân hạng') AS chiMuc, COUNT(*) AS soLuong
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Bài báo%' OR lct.tenLoai LIKE '%Bài báo%')
        GROUP BY COALESCE(tc.xepHang, h.chiMucTapChi, N'Chưa phân hạng')
        ORDER BY soLuong DESC
    """)
    stats_baibao_chimuc = [dict(r) for r in cur.fetchall()]

    # Thống kê chuyên sâu nhóm ĐỀ TÀI NCKH (Revision)
    cur.execute("""
        SELECT COALESCE(capDeTai, N'Cấp Cơ sở') AS capDeTai, COUNT(*) AS soLuong,
               SUM(COALESCE(kinhPhi, 0)) AS tongKinhPhi
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')
        GROUP BY COALESCE(capDeTai, N'Cấp Cơ sở')
        ORDER BY soLuong DESC
    """)
    stats_detai_cap = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT COALESCE(ketQuaNghiemThu, N'Đang thực hiện') AS ketQua, COUNT(*) AS soLuong
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')
        GROUP BY COALESCE(ketQuaNghiemThu, N'Đang thực hiện')
        ORDER BY soLuong DESC
    """)
    stats_detai_nghiemthu = [dict(r) for r in cur.fetchall()]

    # Thống kê theo Đơn vị công tác / chủ trì
    cur.execute("""
        SELECT donViChuTri, COUNT(*) AS soLuong,
               SUM(CASE WHEN trangThai = 'Đã duyệt' THEN 1 ELSE 0 END) AS soDaDuyet
        FROM HoSoKHCN
        WHERE isDeleted = 0
        GROUP BY donViChuTri
        ORDER BY soLuong DESC
    """)
    stats_donvi = [dict(r) for r in cur.fetchall()]

    # Thống kê theo Kỳ báo cáo
    cur.execute("""
        SELECT kyBaoCao, COUNT(*) AS soLuong
        FROM HoSoKHCN
        WHERE isDeleted = 0
        GROUP BY kyBaoCao
        ORDER BY kyBaoCao DESC
    """)
    stats_ky = [dict(r) for r in cur.fetchall()]

    # Danh sách hồ sơ ưu tiên cần xét duyệt / có cảnh báo AI
    cur.execute("""
        SELECT h.*, lv.tenLinhVuc, lv_ai.tenLinhVuc AS tenLinhVucAI, lct.tenLoai, lct.nhomCongTrinh
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0
        ORDER BY h.coCanhBao DESC, h.ngayNop DESC
        LIMIT 6
    """)
    recent_hoso = [dict(r) for r in cur.fetchall()]

    # Lấy chỉ số thực nghiệm AI
    cur.execute("SELECT * FROM ThucNghiemMoHinhAI ORDER BY maThucNghiem ASC")
    ai_metrics = [dict(r) for r in cur.fetchall()]

    conn.close()

    return render_template(
        "dashboard.html",
        kpi={
            "total": total_hoso,
            "cho_duyet": cho_duyet,
            "dang_duyet": dang_duyet,
            "da_duyet": da_duyet,
            "tu_choi": tu_choi,
            "bo_sung": bo_sung,
            "canh_bao": canh_bao_count,
        },
        stats_linhvuc=stats_linhvuc,
        stats_loai=stats_loai,
        stats_baibao_nam=stats_baibao_nam,
        stats_baibao_chimuc=stats_baibao_chimuc,
        stats_detai_cap=stats_detai_cap,
        stats_detai_nghiemthu=stats_detai_nghiemthu,
        stats_donvi=stats_donvi,
        stats_ky=stats_ky,
        recent_hoso=recent_hoso,
        ai_metrics=ai_metrics,
    )


# =============================================================================
# TUẦN 4 - MỤC 2: PHÂN HỆ QUẢN LÝ HỒ SƠ KH&CN (CRUD, TÌM KIẾM, LỌC, PHÂN TRANG)
# =============================================================================
@app.route("/hoso")
@login_required
def hoso_list():
    """Tuần 4 - Mục 2b: Tra cứu, tìm kiếm từ khóa và lọc hồ sơ theo tiêu chí có phân trang."""
    q = request.args.get("q", "").strip()
    nhom = request.args.get("nhom", "").strip()
    ma_linh_vuc = request.args.get("maLinhVuc", "").strip()
    ma_loai = request.args.get("maLoai", "").strip()
    trang_thai = request.args.get("trangThai", "").strip()
    canh_bao = request.args.get("coCanhBao", "").strip()
    page = max(1, int(request.args.get("page", 1)))
    per_page = 6

    conn = get_db_connection()
    cur = conn.cursor()

    where_clauses = ["h.isDeleted = 0"]
    params = []

    if q:
        where_clauses.append("(h.tieuDe LIKE ? OR h.tomTat LIKE ? OR h.tuKhoa LIKE ? OR h.maHoSo LIKE ? OR h.thongTinTacGia LIKE ? OR h.maDeTai LIKE ?)")
        like_q = f"%{q}%"
        params.extend([like_q, like_q, like_q, like_q, like_q, like_q])

    if nhom == "baibao":
        where_clauses.append("(lct.nhomCongTrinh LIKE '%Bài báo%' OR lct.tenLoai LIKE '%Bài báo%')")
    elif nhom == "detai":
        where_clauses.append("(lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')")

    if ma_linh_vuc:
        where_clauses.append("h.maLinhVuc = ?")
        params.append(int(ma_linh_vuc))

    if ma_loai:
        where_clauses.append("h.maLoai = ?")
        params.append(int(ma_loai))

    if trang_thai:
        where_clauses.append("h.trangThai = ?")
        params.append(trang_thai)

    if canh_bao in ("0", "1"):
        where_clauses.append("h.coCanhBao = ?")
        params.append(int(canh_bao))

    where_sql = " AND ".join(where_clauses)

    cur.execute(f"SELECT COUNT(*) FROM HoSoKHCN h LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai WHERE {where_sql}", params)
    total_records = cur.fetchone()[0]
    total_pages = max(1, math.ceil(total_records / per_page))
    page = min(page, total_pages)
    offset = (page - 1) * per_page

    cur.execute(f"""
        SELECT h.*, lv.tenLinhVuc, lv.maCode AS maCodeLV,
               lv_ai.tenLinhVuc AS tenLinhVucAI,
               lct.tenLoai, lct.nhomCongTrinh, tc.tenTapChi, tc.uyTin AS tapChiUyTin,
               nd.hoTen AS tenNguoiNop
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        LEFT JOIN NguoiDung nd ON h.maNguoiNop = nd.maNguoiDung
        WHERE {where_sql}
        ORDER BY h.coCanhBao DESC, h.diemBatThuong DESC, h.maHoSo DESC
        LIMIT ? OFFSET ?
    """, params + [per_page, offset])
    records = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM LinhVuc WHERE trangThai = 1 ORDER BY maLinhVuc")
    linh_vucs = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM LoaiCongTrinh WHERE trangThai = 1 ORDER BY maLoai")
    loai_cts = [dict(r) for r in cur.fetchall()]

    conn.close()

    return render_template(
        "hoso_list.html",
        records=records,
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        filters={
            "q": q,
            "nhom": nhom,
            "maLinhVuc": ma_linh_vuc,
            "maLoai": ma_loai,
            "trangThai": trang_thai,
            "coCanhBao": canh_bao,
        },
        pagination={
            "page": page,
            "per_page": per_page,
            "total_records": total_records,
            "total_pages": total_pages,
        }
    )


@app.route("/hoso/new", methods=["GET", "POST"])
@login_required
def hoso_create():
    """
    Tuần 4 - Mục 2a: Thêm mới hồ sơ KH&CN.
    Nhập đầy đủ tiêu đề, tóm tắt, từ khóa, lĩnh vực, loại công trình, tệp đính kèm,
    tự động gán trạng thái ban đầu là 'Chờ xét duyệt' và kích hoạt mô-đun AI phân loại.
    """
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == "POST":
        tieu_de = request.form.get("tieuDe", "").strip()
        tom_tat = request.form.get("tomTat", "").strip()
        tu_khoa = request.form.get("tuKhoa", "").strip()
        thong_tin_tac_gia = request.form.get("thongTinTacGia", "").strip() or session.get("ho_ten")
        don_vi_chu_tri = request.form.get("donViChuTri", "").strip() or session.get("don_vi")
        nam_cong_bo = int(request.form.get("namCongBo", 2026))
        ky_bao_cao = request.form.get("kyBaoCao", "Đợt 1 - Năm học 2025-2026").strip()
        ma_linh_vuc = int(request.form.get("maLinhVuc", 1))
        ma_loai = int(request.form.get("maLoai", 1))
        ma_tap_chi = request.form.get("maTapChi")
        ma_tap_chi = int(ma_tap_chi) if ma_tap_chi else None

        if not tieu_de or not tom_tat or not tu_khoa:
            conn.close()
            flash("Vui lòng nhập đầy đủ các trường bắt buộc: Tiêu đề, Tóm tắt và Từ khóa.", "danger")
            return redirect(url_for("hoso_create"))

        # Sinh mã hồ sơ tự động
        cur.execute("SELECT COUNT(*) FROM HoSoKHCN")
        next_num = cur.fetchone()[0] + 1
        ma_ho_so = f"HS-2026-{next_num:04d}"
        while cur.execute("SELECT 1 FROM HoSoKHCN WHERE maHoSo = ?", (ma_ho_so,)).fetchone():
            next_num += 1
            ma_ho_so = f"HS-2026-{next_num:04d}"

        # Xử lý tệp đính kèm
        file_name_saved = "MinhChung_MacDinh.pdf"
        uploaded_file = request.files.get("fileDinhKem")
        if uploaded_file and uploaded_file.filename:
            if allowed_file(uploaded_file.filename):
                safe_name = secure_filename(uploaded_file.filename)
                file_name_saved = f"{ma_ho_so}_{safe_name}"
                uploaded_file.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name_saved))

        # Các trường đặc thù của Bài báo khoa học (Revision)
        doi = request.form.get("doi", "").strip() or None
        tap_so = request.form.get("tapSo", "").strip() or None
        trang_bai_bao = request.form.get("trangBaiBao", "").strip() or None
        chi_muc_tap_chi = request.form.get("chiMucTapChi", "").strip() or None

        # Các trường đặc thù của Đề tài NCKH (Revision)
        ma_de_tai = request.form.get("maDeTai", "").strip() or None
        cap_de_tai = request.form.get("capDeTai", "").strip() or None
        chu_nhiem_de_tai = request.form.get("chuNhiemDeTai", "").strip() or None
        thanh_vien_tham_gia = request.form.get("thanhVienThamGia", "").strip() or None
        thoi_gian_thuc_hien = request.form.get("thoiGianThucHien", "").strip() or None
        try:
            kinh_phi = float(request.form.get("kinhPhi", 0) or 0)
        except ValueError:
            kinh_phi = 0.0
        san_pham_du_kien = request.form.get("sanPhamDuKien", "").strip() or None
        ket_qua_nghiem_thu = request.form.get("ketQuaNghiemThu", "Đang thực hiện").strip() or "Đang thực hiện"

        if chu_nhiem_de_tai and not thong_tin_tac_gia:
            thong_tin_tac_gia = chu_nhiem_de_tai

        # Chạy mô-đun AI Tiền xử lý tiếng Việt + Phân loại tự động (Chức năng 5)
        pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia)

        # Trạng thái ban đầu mặc định là 'Chờ xét duyệt' (Tuần 4 - Mục 2a)
        trang_thai_ban_dau = "Chờ xét duyệt"

        cur.execute("""
            INSERT INTO HoSoKHCN (
                maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
                namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
                trangThai, fileDinhKem, vanBanDaTachTu, maLinhVucAI, doTinCayAI,
                maLinhVucSVM, doTinCaySVM,
                doi, tapSo, trangBaiBao, chiMucTapChi,
                maDeTai, capDeTai, chuNhiemDeTai, thanhVienThamGia, thoiGianThucHien,
                kinhPhi, sanPhamDuKien, ketQuaNghiemThu
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ma_ho_so, tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia, don_vi_chu_tri,
            nam_cong_bo, ky_bao_cao, session["user_id"], ma_linh_vuc, ma_loai, ma_tap_chi,
            trang_thai_ban_dau, file_name_saved, pred["segmented_text"],
            pred["phobert_label_id"], pred["phobert_confidence"],
            pred["svm_label_id"], pred["svm_confidence"],
            doi, tap_so, trang_bai_bao, chi_muc_tap_chi,
            ma_de_tai, cap_de_tai, chu_nhiem_de_tai, thanh_vien_tham_gia, thoi_gian_thuc_hien,
            kinh_phi, san_pham_du_kien, ket_qua_nghiem_thu
        ))

        # Ghi nhật ký khởi tạo vào bảng LichSuXetDuyet
        cur.execute("""
            INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
            VALUES (?, ?, ?, ?, ?)
        """, (ma_ho_so, session["user_id"], "Khởi tạo mới", "Chờ xét duyệt", "Tác giả khai báo và nộp mới hồ sơ KH&CN lên hệ thống."))

        conn.commit()
        # Cập nhật đồ thị NetworkX & Phát hiện bất thường Isolation Forest
        ai_engine.build_networkx_graph_and_detect_anomalies(conn)
        conn.close()

        flash(
            f"Đã thêm mới hồ sơ {ma_ho_so} (Trạng thái: Chờ xét duyệt). "
            f"Mô-đun AI PhoBERT gợi ý phân loại với độ tin cậy {pred['phobert_confidence']}%.",
            "success"
        )
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    cur.execute("SELECT * FROM LinhVuc WHERE trangThai = 1 ORDER BY maLinhVuc")
    linh_vucs = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM LoaiCongTrinh WHERE trangThai = 1 ORDER BY maLoai")
    loai_cts = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM TapChiHopLe WHERE trangThai = 1 ORDER BY maTapChi")
    tap_chis = [dict(r) for r in cur.fetchall()]
    conn.close()

    return render_template(
        "hoso_form.html",
        mode="create",
        hoso=None,
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        tap_chis=tap_chis,
    )


@app.route("/hoso/<ma_ho_so>/edit", methods=["GET", "POST"])
@login_required
def hoso_edit(ma_ho_so):
    """
    Tuần 4 - Mục 2c: Chỉnh sửa hồ sơ theo phân quyền.
    Cho phép cập nhật thông tin khi hồ sơ ở trạng thái 'Chờ xét duyệt' (hoặc bởi Admin).
    """
    conn = get_db_connection()
    cur = conn.cursor()
    hoso = cur.execute(
        "SELECT * FROM HoSoKHCN WHERE maHoSo = ? AND isDeleted = 0", (ma_ho_so,)
    ).fetchone()

    if not hoso:
        conn.close()
        flash("Không tìm thấy hồ sơ cần chỉnh sửa.", "danger")
        return redirect(url_for("hoso_list"))

    # Kiểm tra điều kiện trạng thái và phân quyền (Revision: Cho phép sửa khi 'Chờ xét duyệt' hoặc 'Yêu cầu bổ sung')
    if hoso["trangThai"] not in ("Chờ xét duyệt", "Yêu cầu bổ sung") and session.get("vai_tro") != "Admin":
        conn.close()
        flash(
            f"Chỉ được phép chỉnh sửa khi hồ sơ ở trạng thái 'Chờ xét duyệt' hoặc 'Yêu cầu bổ sung' (Trạng thái hiện tại: {hoso['trangThai']}).",
            "warning"
        )
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    if session.get("vai_tro") == "TacGia" and hoso["maNguoiNop"] != session.get("user_id"):
        conn.close()
        flash("Bạn chỉ được phép chỉnh sửa hồ sơ do chính mình nộp.", "danger")
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    if request.method == "POST":
        tieu_de = request.form.get("tieuDe", "").strip()
        tom_tat = request.form.get("tomTat", "").strip()
        tu_khoa = request.form.get("tuKhoa", "").strip()
        thong_tin_tac_gia = request.form.get("thongTinTacGia", "").strip()
        don_vi_chu_tri = request.form.get("donViChuTri", "").strip()
        nam_cong_bo = int(request.form.get("namCongBo", hoso["namCongBo"]))
        ky_bao_cao = request.form.get("kyBaoCao", hoso["kyBaoCao"]).strip()
        ma_linh_vuc = int(request.form.get("maLinhVuc", hoso["maLinhVuc"]))
        ma_loai = int(request.form.get("maLoai", hoso["maLoai"]))
        ma_tap_chi = request.form.get("maTapChi")
        ma_tap_chi = int(ma_tap_chi) if ma_tap_chi else None

        # Các trường đặc thù của Bài báo khoa học (Revision)
        doi = request.form.get("doi", "").strip() or None
        tap_so = request.form.get("tapSo", "").strip() or None
        trang_bai_bao = request.form.get("trangBaiBao", "").strip() or None
        chi_muc_tap_chi = request.form.get("chiMucTapChi", "").strip() or None

        # Các trường đặc thù của Đề tài NCKH (Revision)
        ma_de_tai = request.form.get("maDeTai", "").strip() or None
        cap_de_tai = request.form.get("capDeTai", "").strip() or None
        chu_nhiem_de_tai = request.form.get("chuNhiemDeTai", "").strip() or None
        thanh_vien_tham_gia = request.form.get("thanhVienThamGia", "").strip() or None
        thoi_gian_thuc_hien = request.form.get("thoiGianThucHien", "").strip() or None
        try:
            kinh_phi = float(request.form.get("kinhPhi", hoso["kinhPhi"] or 0) or 0)
        except (ValueError, TypeError):
            kinh_phi = 0.0
        san_pham_du_kien = request.form.get("sanPhamDuKien", "").strip() or None
        ket_qua_nghiem_thu = request.form.get("ketQuaNghiemThu", hoso["ketQuaNghiemThu"] or "Đang thực hiện").strip() or "Đang thực hiện"

        if chu_nhiem_de_tai and not thong_tin_tac_gia:
            thong_tin_tac_gia = chu_nhiem_de_tai

        file_name_saved = hoso["fileDinhKem"]
        uploaded_file = request.files.get("fileDinhKem")
        if uploaded_file and uploaded_file.filename and allowed_file(uploaded_file.filename):
            safe_name = secure_filename(uploaded_file.filename)
            file_name_saved = f"{ma_ho_so}_{safe_name}"
            uploaded_file.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name_saved))

        # Chạy lại AI phân loại & phát hiện bất thường sau khi cập nhật nội dung
        pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia)

        # Nghiệp vụ 3 - Bước 5 & Nghiệp vụ 1: Vòng lặp phản hồi và chỉnh sửa (Resubmission loop)
        # Khi tác giả chỉnh sửa hồ sơ đang ở trạng thái 'Yêu cầu bổ sung', tự động chuyển lại 'Chờ xét duyệt'
        trang_thai_moi = hoso["trangThai"]
        if hoso["trangThai"] == "Yêu cầu bổ sung":
            trang_thai_moi = "Chờ xét duyệt"
            cur.execute("""
                INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
                VALUES (?, ?, ?, ?, ?)
            """, (ma_ho_so, session["user_id"], "Yêu cầu bổ sung", "Chờ xét duyệt", "Tác giả đã cập nhật nội dung/minh chứng theo yêu cầu và gửi lại để thẩm định."))

        cur.execute("""
            UPDATE HoSoKHCN
            SET tieuDe = ?, tomTat = ?, tuKhoa = ?, thongTinTacGia = ?, donViChuTri = ?,
                namCongBo = ?, kyBaoCao = ?, maLinhVuc = ?, maLoai = ?, maTapChi = ?,
                trangThai = ?,
                fileDinhKem = ?, vanBanDaTachTu = ?, maLinhVucAI = ?, doTinCayAI = ?,
                maLinhVucSVM = ?, doTinCaySVM = ?,
                doi = ?, tapSo = ?, trangBaiBao = ?, chiMucTapChi = ?,
                maDeTai = ?, capDeTai = ?, chuNhiemDeTai = ?, thanhVienThamGia = ?,
                thoiGianThucHien = ?, kinhPhi = ?, sanPhamDuKien = ?, ketQuaNghiemThu = ?,
                ngayCapNhat = CURRENT_TIMESTAMP
            WHERE maHoSo = ?
        """, (
            tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia, don_vi_chu_tri,
            nam_cong_bo, ky_bao_cao, ma_linh_vuc, ma_loai, ma_tap_chi,
            trang_thai_moi,
            file_name_saved, pred["segmented_text"], pred["phobert_label_id"],
            pred["phobert_confidence"], pred["svm_label_id"], pred["svm_confidence"],
            doi, tap_so, trang_bai_bao, chi_muc_tap_chi,
            ma_de_tai, cap_de_tai, chu_nhiem_de_tai, thanh_vien_tham_gia,
            thoi_gian_thuc_hien, kinh_phi, san_pham_du_kien, ket_qua_nghiem_thu,
            ma_ho_so
        ))

        conn.commit()
        ai_engine.build_networkx_graph_and_detect_anomalies(conn)
        conn.close()
        if hoso["trangThai"] == "Yêu cầu bổ sung":
            flash(f"Đã cập nhật hồ sơ {ma_ho_so} và chuyển trạng thái về 'Chờ xét duyệt' để Cán bộ thẩm định lại.", "success")
        else:
            flash(f"Đã cập nhật thông tin hồ sơ {ma_ho_so} và phân tích lại bằng AI.", "success")
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    cur.execute("SELECT * FROM LinhVuc WHERE trangThai = 1 ORDER BY maLinhVuc")
    linh_vucs = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM LoaiCongTrinh WHERE trangThai = 1 ORDER BY maLoai")
    loai_cts = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM TapChiHopLe WHERE trangThai = 1 ORDER BY maTapChi")
    tap_chis = [dict(r) for r in cur.fetchall()]
    conn.close()

    return render_template(
        "hoso_form.html",
        mode="edit",
        hoso=dict(hoso),
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        tap_chis=tap_chis,
    )


@app.route("/hoso/<ma_ho_so>/soft-delete", methods=["POST"])
@login_required
def hoso_soft_delete(ma_ho_so):
    """Tuần 4 - Mục 2c: Xóa mềm (Soft Delete) hồ sơ khi ở trạng thái chờ duyệt."""
    conn = get_db_connection()
    hoso = conn.execute("SELECT * FROM HoSoKHCN WHERE maHoSo = ?", (ma_ho_so,)).fetchone()
    if not hoso:
        conn.close()
        flash("Hồ sơ không tồn tại.", "danger")
        return redirect(url_for("hoso_list"))

    if session.get("vai_tro") != "Admin":
        if hoso["trangThai"] != "Chờ xét duyệt" or hoso["maNguoiNop"] != session.get("user_id"):
            conn.close()
            flash("Bạn chỉ được phép xóa mềm hồ sơ của mình khi đang ở trạng thái 'Chờ xét duyệt'.", "danger")
            return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    conn.execute("UPDATE HoSoKHCN SET isDeleted = 1 WHERE maHoSo = ?", (ma_ho_so,))
    conn.commit()
    conn.close()
    flash(f"Đã chuyển hồ sơ {ma_ho_so} vào thùng rác (Xóa mềm).", "info")
    return redirect(url_for("hoso_list"))


@app.route("/hoso/<ma_ho_so>/hard-delete", methods=["POST"])
@role_required("Admin")
def hoso_hard_delete(ma_ho_so):
    """
    Tuần 4 - Mục 1b & 2c: Xóa cứng (Hard Delete) vĩnh viễn dữ liệu.
    Được bảo vệ bởi Decorator @role_required('Admin') ngăn chặn Cán bộ xét duyệt và Tác giả.
    """
    conn = get_db_connection()
    conn.execute("DELETE FROM HoSoKHCN WHERE maHoSo = ?", (ma_ho_so,))
    conn.commit()
    conn.close()
    flash(f"Quản trị viên đã xóa vĩnh viễn (Hard Delete) hồ sơ {ma_ho_so} khỏi CSDL.", "warning")
    return redirect(url_for("hoso_list"))


# =============================================================================
# TUẦN 5 - MỤC 2 & TUẦN 6 - MỤC 2D: CHI TIẾT HỒ SƠ, XÉT DUYỆT & AI FEEDBACK LOOP
# =============================================================================
@app.route("/hoso/<ma_ho_so>")
@login_required
def hoso_detail(ma_ho_so):
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT h.*,
               lv.tenLinhVuc, lv.maCode AS maCodeLV,
               lv_ai.tenLinhVuc AS tenLinhVucAI, lv_ai.maCode AS maCodeLVAI,
               lv_svm.tenLinhVuc AS tenLinhVucSVM,
               lct.tenLoai, lct.nhomCongTrinh, lct.gioChuanQuyDoi,
               tc.tenTapChi, tc.chiSoISSN, tc.xepHang, tc.uyTin AS tapChiUyTin,
               nd.hoTen AS tenNguoiNop, nd.email AS emailNguoiNop
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN LinhVuc lv_svm ON h.maLinhVucSVM = lv_svm.maLinhVuc
        LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        LEFT JOIN NguoiDung nd ON h.maNguoiNop = nd.maNguoiDung
        WHERE h.maHoSo = ? AND h.isDeleted = 0
    """, (ma_ho_so,))
    row = cur.fetchone()
    if not row:
        conn.close()
        flash("Hồ sơ không tồn tại hoặc đã bị xóa.", "danger")
        return redirect(url_for("hoso_list"))

    hoso = dict(row)
    try:
        hoso["canhBaoList"] = json.loads(hoso.get("chiTietCanhBao") or "[]")
    except Exception:
        hoso["canhBaoList"] = []

    # Nghiệp vụ 3 - Bước 1c: Tiếp nhận và Chuyển trạng thái xử lý
    # Khi Cán bộ xét duyệt / Quản trị viên mở hồ sơ ở trạng thái 'Chờ xét duyệt',
    # hệ thống tự động cập nhật trạng thái sang 'Đang xét duyệt' và ghi nhận cán bộ phụ trách vào LichSuXetDuyet để tránh trùng lặp xử lý.
    if session.get("vai_tro") in ["CanBo", "Admin"] and hoso.get("trangThai") == "Chờ xét duyệt":
        cur.execute("""
            UPDATE HoSoKHCN
            SET trangThai = ?, ngayCapNhat = CURRENT_TIMESTAMP
            WHERE maHoSo = ?
        """, ("Đang xét duyệt", ma_ho_so))
        cur.execute("""
            INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
            VALUES (?, ?, ?, ?, ?)
        """, (ma_ho_so, session["user_id"], "Chờ xét duyệt", "Đang xét duyệt", "Cán bộ xét duyệt mở hồ sơ tiếp nhận thẩm định tính hợp lệ (hệ thống tự động chuyển trạng thái)."))
        conn.commit()
        hoso["trangThai"] = "Đang xét duyệt"

    # Lịch sử xét duyệt (Bảng LichSuXetDuyet - Tuần 5 Mục 2b)
    cur.execute("""
        SELECT ls.*, nd.hoTen AS tenCanBo, nd.vaiTro AS vaiTroCanBo
        FROM LichSuXetDuyet ls
        LEFT JOIN NguoiDung nd ON ls.maCanBo = nd.maNguoiDung
        WHERE ls.maHoSo = ?
        ORDER BY ls.maLichSu DESC
    """, (ma_ho_so,))
    lich_su = [dict(r) for r in cur.fetchall()]

    # Lịch sử phản hồi AI (Feedback Loop - Tuần 6 Mục 2d)
    cur.execute("""
        SELECT pf.*, nd.hoTen AS tenCanBo,
               lv1.tenLinhVuc AS tenNhanAI, lv2.tenLinhVuc AS tenNhanThucTe
        FROM PhanHoiAI pf
        LEFT JOIN NguoiDung nd ON pf.maCanBo = nd.maNguoiDung
        LEFT JOIN LinhVuc lv1 ON pf.maLinhVucAI = lv1.maLinhVuc
        LEFT JOIN LinhVuc lv2 ON pf.maLinhVucThucTe = lv2.maLinhVuc
        WHERE pf.maHoSo = ?
        ORDER BY pf.maPhanHoi DESC
    """, (ma_ho_so,))
    phan_hoi_ai = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM LinhVuc WHERE trangThai = 1 ORDER BY maLinhVuc")
    linh_vucs = [dict(r) for r in cur.fetchall()]

    conn.close()

    badge_info = ai_engine.get_confidence_badge_info(float(hoso.get("doTinCayAI") or 0.0))

    return render_template(
        "hoso_detail.html",
        hoso=hoso,
        lich_su=lich_su,
        phan_hoi_ai=phan_hoi_ai,
        linh_vucs=linh_vucs,
        badge_info=badge_info,
    )


@app.route("/hoso/<ma_ho_so>/xet-duyet", methods=["POST"])
@role_required("Admin", "CanBo")
def hoso_review_action(ma_ho_so):
    """
    Tuần 5 - Mục 2a & 2b: Quản lý quy trình xét duyệt hồ sơ (Chức năng 3).
    Chuyển đổi trạng thái qua các mốc: Chờ xét duyệt -> Đang xét duyệt -> Đã duyệt / Từ chối.
    Tự động ghi nhận bản ghi vào bảng LichSuXetDuyet lưu thời gian, cán bộ thực hiện và ý kiến đánh giá.
    """
    trang_thai_moi = request.form.get("trangThaiMoi", "").strip()
    y_kien = request.form.get("yKienDanhGia", "").strip()

    valid_states = {"Chờ xét duyệt", "Đang xét duyệt", "Đã duyệt", "Từ chối", "Yêu cầu bổ sung"}
    if trang_thai_moi not in valid_states:
        flash("Trạng thái xét duyệt không hợp lệ.", "danger")
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    if not y_kien:
        flash("Vui lòng nhập nội dung ý kiến đánh giá thẩm định để lưu vào Lịch sử xét duyệt.", "warning")
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    if trang_thai_moi in ("Từ chối", "Yêu cầu bổ sung") and len(y_kien) < 5:
        flash("Bắt buộc phải nhập chi tiết lý do từ chối hoặc nội dung yêu cầu bổ sung (tối thiểu 5 ký tự) để phản hồi cho tác giả.", "warning")
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    conn = get_db_connection()
    cur = conn.cursor()
    hoso = cur.execute("SELECT trangThai FROM HoSoKHCN WHERE maHoSo = ?", (ma_ho_so,)).fetchone()
    if not hoso:
        conn.close()
        flash("Hồ sơ không tồn tại.", "danger")
        return redirect(url_for("hoso_list"))

    trang_thai_cu = hoso["trangThai"]
    cur.execute("""
        UPDATE HoSoKHCN
        SET trangThai = ?, ngayCapNhat = CURRENT_TIMESTAMP
        WHERE maHoSo = ?
    """, (trang_thai_moi, ma_ho_so))

    cur.execute("""
        INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
        VALUES (?, ?, ?, ?, ?)
    """, (ma_ho_so, session["user_id"], trang_thai_cu, trang_thai_moi, y_kien))

    conn.commit()
    conn.close()

    flash(
        f"Đã cập nhật trạng thái hồ sơ {ma_ho_so} từ '{trang_thai_cu}' sang '{trang_thai_moi}' và ghi nhận vào LichSuXetDuyet.",
        "success"
    )
    return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))


@app.route("/hoso/<ma_ho_so>/ai-feedback", methods=["POST"])
@role_required("Admin", "CanBo")
def hoso_ai_feedback(ma_ho_so):
    """
    Tuần 6 - Mục 2d: Cơ chế phản hồi nhãn phân loại AI (Feedback Loop).
    - Cán bộ nhấn nút 'Xác nhận nhãn AI' nếu dự đoán chính xác
    - Hoặc chọn 'Điều chỉnh' và chọn lại nhãn đúng từ danh sách danh mục chuẩn của hệ thống
    - Lưu cặp đối sánh (nhãn AI đề xuất vs nhãn thực tế cán bộ chỉnh sửa) vào bảng PhanHoiAI
      để tiếp tục cập nhật và fine-tuning lại mô hình.
    """
    action_type = request.form.get("actionType", "confirm")
    ghi_chu = request.form.get("ghiChuPhanHoi", "").strip()

    conn = get_db_connection()
    cur = conn.cursor()
    hoso = cur.execute("SELECT * FROM HoSoKHCN WHERE maHoSo = ?", (ma_ho_so,)).fetchone()
    if not hoso:
        conn.close()
        flash("Hồ sơ không tồn tại.", "danger")
        return redirect(url_for("hoso_list"))

    ma_lv_ai = int(hoso["maLinhVucAI"] or hoso["maLinhVuc"])
    do_tin_cay = float(hoso["doTinCayAI"] or 80.0)
    van_ban_tach_tu = hoso["vanBanDaTachTu"] or ai_engine.preprocess_record(
        hoso["tieuDe"], hoso["tomTat"], hoso["tuKhoa"], hoso["thongTinTacGia"]
    )

    if action_type == "confirm":
        ma_lv_thuc_te = ma_lv_ai
        ket_qua_khop = 1
        if not ghi_chu:
            ghi_chu = "Cán bộ xét duyệt xác nhận nhãn lĩnh vực do AI (PhoBERT) đề xuất là chính xác."
    else:
        ma_lv_thuc_te = int(request.form.get("maLinhVucThucTe", ma_lv_ai))
        ket_qua_khop = 1 if ma_lv_thuc_te == ma_lv_ai else 0
        if not ghi_chu:
            ghi_chu = "Cán bộ xét duyệt điều chỉnh lại nhãn lĩnh vực chuẩn cho hồ sơ phục vụ tái huấn luyện (Fine-tuning)."

    # Lưu vào bảng PhanHoiAI
    cur.execute("""
        INSERT INTO PhanHoiAI
        (maHoSo, maCanBo, maLinhVucAI, maLinhVucThucTe, doTinCayLucDo, ketQuaKhop, vanBanHuanLuyen, ghiChuPhanHoi)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        ma_ho_so, session["user_id"], ma_lv_ai, ma_lv_thuc_te,
        do_tin_cay, ket_qua_khop, van_ban_tach_tu, ghi_chu
    ))

    # Cập nhật lĩnh vực chuẩn của hồ sơ và đánh dấu đã xác nhận AI
    cur.execute("""
        UPDATE HoSoKHCN
        SET maLinhVuc = ?, daXacNhanAI = 1, ngayCapNhat = CURRENT_TIMESTAMP
        WHERE maHoSo = ?
    """, (ma_lv_thuc_te, ma_ho_so))

    conn.commit()
    # Cập nhật lại mô hình và đồ thị bất thường với dữ liệu phản hồi mới
    ai_engine.train_and_evaluate(conn)
    ai_engine.build_networkx_graph_and_detect_anomalies(conn)
    conn.close()

    if ket_qua_khop == 1:
        flash("Đã xác nhận nhãn gợi ý từ AI! Dữ liệu đã được ghi nhận vào kho Feedback Loop.", "success")
    else:
        flash("Đã điều chỉnh nhãn lĩnh vực thực tế và cập nhật vào kho dữ liệu Fine-tuning của mô hình AI!", "info")

    return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))


# =============================================================================
# API GỢI Ý NHANH AI TRỰC TIẾP TRÊN FORM KHAI BÁO
# =============================================================================
@app.route("/api/ai-preview", methods=["POST"])
@login_required
def api_ai_preview():
    data = request.get_json(silent=True) or {}
    tieu_de = data.get("tieuDe", "")
    tom_tat = data.get("tomTat", "")
    tu_khoa = data.get("tuKhoa", "")
    tac_gia = data.get("thongTinTacGia", "")

    pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, tac_gia)

    conn = get_db_connection()
    lv_pho = conn.execute("SELECT tenLinhVuc FROM LinhVuc WHERE maLinhVuc = ?", (pred["phobert_label_id"],)).fetchone()
    lv_svm = conn.execute("SELECT tenLinhVuc FROM LinhVuc WHERE maLinhVuc = ?", (pred["svm_label_id"],)).fetchone()
    conn.close()

    return jsonify({
        "segmented_text": pred["segmented_text"],
        "phobert_label_id": pred["phobert_label_id"],
        "phobert_label_name": lv_pho["tenLinhVuc"] if lv_pho else "",
        "phobert_confidence": pred["phobert_confidence"],
        "svm_label_id": pred["svm_label_id"],
        "svm_label_name": lv_svm["tenLinhVuc"] if lv_svm else "",
        "svm_confidence": pred["svm_confidence"],
        "badge": pred["badge"],
    })


# =============================================================================
# TUẦN 5 - MỤC 1: QUẢN LÝ DANH MỤC HỆ THỐNG (CHỨC NĂNG 2)
# =============================================================================
@app.route("/danh-muc")
@login_required
def catalog_management():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT lv.maLinhVuc, lv.maCode, lv.tenLinhVuc, lv.moTa, lv.trangThai, COUNT(h.maHoSo) AS soHoSoLienKet
        FROM LinhVuc lv
        LEFT JOIN HoSoKHCN h ON lv.maLinhVuc = h.maLinhVuc AND h.isDeleted = 0
        GROUP BY lv.maLinhVuc, lv.maCode, lv.tenLinhVuc, lv.moTa, lv.trangThai
        ORDER BY lv.maLinhVuc
    """)
    linh_vucs = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT lct.maLoai, lct.maCodeLoai, lct.tenLoai, lct.nhomCongTrinh, lct.trangThai, COUNT(h.maHoSo) AS soHoSoLienKet
        FROM LoaiCongTrinh lct
        LEFT JOIN HoSoKHCN h ON lct.maLoai = h.maLoai AND h.isDeleted = 0
        GROUP BY lct.maLoai, lct.maCodeLoai, lct.tenLoai, lct.nhomCongTrinh, lct.trangThai
        ORDER BY lct.maLoai
    """)
    loai_cts = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT tc.maTapChi, tc.tenTapChi, tc.chiSoISSN, tc.xepHang, tc.maLinhVuc, tc.uyTin, tc.trangThai, lv.tenLinhVuc, COUNT(h.maHoSo) AS soHoSoLienKet
        FROM TapChiHopLe tc
        LEFT JOIN LinhVuc lv ON tc.maLinhVuc = lv.maLinhVuc
        LEFT JOIN HoSoKHCN h ON tc.maTapChi = h.maTapChi AND h.isDeleted = 0
        GROUP BY tc.maTapChi, tc.tenTapChi, tc.chiSoISSN, tc.xepHang, tc.maLinhVuc, tc.uyTin, tc.trangThai, lv.tenLinhVuc
        ORDER BY tc.maTapChi
    """)
    tap_chis = [dict(r) for r in cur.fetchall()]

    conn.close()
    return render_template(
        "catalogs.html",
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        tap_chis=tap_chis,
    )


@app.route("/danh-muc/linh-vuc/add", methods=["POST"])
@role_required("Admin")
def catalog_add_linhvuc():
    ma_code = request.form.get("maCode", "").strip().upper()
    ten_lv = request.form.get("tenLinhVuc", "").strip()
    mo_ta = request.form.get("moTa", "").strip()
    tu_khoa = request.form.get("tuKhoaDacTrung", "").strip()

    if not ma_code or not ten_lv:
        flash("Mã lĩnh vực và Tên lĩnh vực không được để trống.", "danger")
        return redirect(url_for("catalog_management"))

    conn = get_db_connection()
    try:
        conn.execute(
            "INSERT INTO LinhVuc (maCode, tenLinhVuc, moTa, tuKhoaDacTrung, trangThai) VALUES (?, ?, ?, ?, 1)",
            (ma_code, ten_lv, mo_ta, tu_khoa)
        )
        conn.commit()
        flash(f"Đã thêm danh mục lĩnh vực '{ten_lv}'.", "success")
    except Exception as e:
        flash(f"Không thể thêm lĩnh vực (Mã có thể bị trùng): {e}", "danger")
    finally:
        conn.close()
    return redirect(url_for("catalog_management"))


@app.route("/danh-muc/loai-cong-trinh/add", methods=["POST"])
@role_required("Admin")
def catalog_add_loai():
    ma_code = request.form.get("maCodeLoai", "").strip().upper()
    ten_loai = request.form.get("tenLoai", "").strip()
    nhom = request.form.get("nhomCongTrinh", "Bài báo khoa học").strip()
    gio_chuan = float(request.form.get("gioChuanQuyDoi", 60.0))

    conn = get_db_connection()
    try:
        conn.execute(
            "INSERT INTO LoaiCongTrinh (maCodeLoai, tenLoai, nhomCongTrinh, gioChuanQuyDoi, trangThai) VALUES (?, ?, ?, ?, 1)",
            (ma_code, ten_loai, nhom, gio_chuan)
        )
        conn.commit()
        flash(f"Đã thêm loại công trình '{ten_loai}'.", "success")
    except Exception as e:
        flash(f"Lỗi khi thêm loại công trình: {e}", "danger")
    finally:
        conn.close()
    return redirect(url_for("catalog_management"))


@app.route("/danh-muc/tap-chi/add", methods=["POST"])
@role_required("Admin")
def catalog_add_tapchi():
    ten_tc = request.form.get("tenTapChi", "").strip()
    issn = request.form.get("chiSoISSN", "").strip()
    xep_hang = request.form.get("xepHang", "HĐGSNN").strip()
    ma_lv = int(request.form.get("maLinhVuc", 1))
    uy_tin = int(request.form.get("uyTin", 1))

    conn = get_db_connection()
    conn.execute(
        "INSERT INTO TapChiHopLe (tenTapChi, chiSoISSN, xepHang, maLinhVuc, uyTin, trangThai) VALUES (?, ?, ?, ?, ?, 1)",
        (ten_tc, issn, xep_hang, ma_lv, uy_tin)
    )
    conn.commit()
    conn.close()
    flash(f"Đã bổ sung tạp chí '{ten_tc}' vào danh mục kiểm tra chéo.", "success")
    return redirect(url_for("catalog_management"))


@app.route("/danh-muc/toggle/<cat_type>/<int:item_id>", methods=["POST"])
@role_required("Admin")
def catalog_toggle_status(cat_type, item_id):
    """
    Tuần 5 - Mục 1b: Cho phép Quản trị viên vô hiệu hóa hoặc kích hoạt lại danh mục
    mà không làm ảnh hưởng đến tính toàn vẹn liên kết khóa ngoại của các hồ sơ cũ.
    """
    table_map = {
        "linhvuc": ("LinhVuc", "maLinhVuc"),
        "loai": ("LoaiCongTrinh", "maLoai"),
        "tapchi": ("TapChiHopLe", "maTapChi"),
    }
    if cat_type not in table_map:
        abort(400)
    tbl, pk = table_map[cat_type]
    conn = get_db_connection()
    conn.execute(f"UPDATE {tbl} SET trangThai = CASE WHEN trangThai = 1 THEN 0 ELSE 1 END WHERE {pk} = ?", (item_id,))
    conn.commit()
    conn.close()
    flash("Đã cập nhật trạng thái hiển thị của danh mục (Bảo toàn toàn vẹn tham chiếu hồ sơ cũ).", "info")
    return redirect(url_for("catalog_management"))


# =============================================================================
# TUẦN 5 - MỤC 3B: KẾT XUẤT DỮ LIỆU BÁO CÁO RA EXCEL (.XLSX) & PDF (.PDF)
# =============================================================================
def _fetch_filtered_records_for_export():
    ma_linh_vuc = request.args.get("maLinhVuc", "").strip()
    trang_thai = request.args.get("trangThai", "").strip()
    ky_bao_cao = request.args.get("kyBaoCao", "").strip()

    where_clauses = ["h.isDeleted = 0"]
    params = []
    desc_parts = []

    if ma_linh_vuc:
        where_clauses.append("h.maLinhVuc = ?")
        params.append(int(ma_linh_vuc))
        desc_parts.append(f"Mã lĩnh vực={ma_linh_vuc}")
    if trang_thai:
        where_clauses.append("h.trangThai = ?")
        params.append(trang_thai)
        desc_parts.append(f"Trạng thái={trang_thai}")
    if ky_bao_cao:
        where_clauses.append("h.kyBaoCao = ?")
        params.append(ky_bao_cao)
        desc_parts.append(f"Kỳ={ky_bao_cao}")

    where_sql = " AND ".join(where_clauses)
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT h.*, lv.tenLinhVuc, lv_ai.tenLinhVuc AS tenLinhVucAI, lct.tenLoai
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE {where_sql}
        ORDER BY h.maHoSo ASC
    """, params).fetchall()
    conn.close()
    filter_desc = ", ".join(desc_parts) if desc_parts else "Toàn bộ hồ sơ KH&CN"
    return [dict(r) for r in rows], filter_desc


@app.route("/bao-cao/export/excel")
@login_required
def export_excel():
    records, filter_desc = _fetch_filtered_records_for_export()
    excel_stream = generate_excel_report(records, filter_desc)
    filename = f"BaoCao_HoSo_KHCN_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        excel_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/bao-cao/export/pdf")
@login_required
def export_pdf():
    records, filter_desc = _fetch_filtered_records_for_export()
    pdf_stream = generate_pdf_report(records, filter_desc)
    filename = f"BaoCao_HoSo_KHCN_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    return send_file(
        pdf_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf",
    )


# =============================================================================
# TUẦN 6: TRUNG TÂM THỰC NGHIỆM AI (NLP PHOBERT vs SVM & ĐỒ THỊ NETWORKX)
# =============================================================================
@app.route("/ai-center")
@login_required
def ai_center():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM ThucNghiemMoHinhAI ORDER BY maThucNghiem ASC")
    metrics = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT pf.*, nd.hoTen AS tenCanBo,
               lv1.tenLinhVuc AS tenNhanAI, lv2.tenLinhVuc AS tenNhanThucTe
        FROM PhanHoiAI pf
        LEFT JOIN NguoiDung nd ON pf.maCanBo = nd.maNguoiDung
        LEFT JOIN LinhVuc lv1 ON pf.maLinhVucAI = lv1.maLinhVuc
        LEFT JOIN LinhVuc lv2 ON pf.maLinhVucThucTe = lv2.maLinhVuc
        ORDER BY pf.maPhanHoi DESC
    """)
    feedbacks = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT h.*, lv.tenLinhVuc, lv_ai.tenLinhVuc AS tenLinhVucAI, tc.tenTapChi
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        WHERE h.isDeleted = 0 AND h.coCanhBao = 1
        ORDER BY h.diemBatThuong DESC
    """)
    anomalies = []
    for r in cur.fetchall():
        d = dict(r)
        try:
            d["reasons"] = json.loads(d.get("chiTietCanhBao") or "[]")
        except Exception:
            d["reasons"] = []
        anomalies.append(d)

    conn.close()
    graph_summary = ai_engine.get_graph_summary()

    return render_template(
        "ai_center.html",
        metrics=metrics,
        feedbacks=feedbacks,
        anomalies=anomalies,
        graph_summary=graph_summary,
        phobert_params=PHOBERT_HYPERPARAMS,
    )


@app.route("/ai-center/retrain", methods=["POST"])
@role_required("Admin", "CanBo")
def ai_retrain():
    conn = get_db_connection()
    ai_engine.train_and_evaluate(conn)
    ai_engine.build_networkx_graph_and_detect_anomalies(conn)
    conn.close()
    flash(
        "Đã huấn luyện lại (Retrain / Fine-tune) mô hình TF-IDF+SVM, PhoBERT và cập nhật đồ thị NetworkX + Isolation Forest từ kho dữ liệu Feedback Loop!",
        "success"
    )
    return redirect(url_for("ai_center"))


# =============================================================================
# QUẢN TRỊ NGƯỜI DÙNG & PHÂN QUYỀN RBAC (CHỈ ADMIN)
# =============================================================================
@app.route("/quan-tri/nguoi-dung", methods=["GET", "POST"])
@role_required("Admin")
def user_management():
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == "POST":
        action = request.form.get("action")
        if action == "create":
            ho_ten = request.form.get("hoTen", "").strip()
            email = request.form.get("email", "").strip()
            mat_khau = request.form.get("matKhau", "123456")
            vai_tro = request.form.get("vaiTro", "TacGia")
            don_vi = request.form.get("donViCongTac", "").strip()
            hoc_vi = request.form.get("hocHamHocVi", "ThS.").strip()
            try:
                cur.execute("""
                    INSERT INTO NguoiDung (hoTen, email, matKhau, vaiTro, donViCongTac, hocHamHocVi)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (ho_ten, email, generate_password_hash(mat_khau), vai_tro, don_vi, hoc_vi))
                conn.commit()
                flash(f"Đã tạo tài khoản '{ho_ten}' với vai trò '{vai_tro}'.", "success")
            except Exception as e:
                flash(f"Lỗi tạo tài khoản (Email có thể đã tồn tại): {e}", "danger")
        elif action == "update_role":
            ma_nd = int(request.form.get("maNguoiDung"))
            vai_tro_moi = request.form.get("vaiTroMoi")
            cur.execute("UPDATE NguoiDung SET vaiTro = ? WHERE maNguoiDung = ?", (vai_tro_moi, ma_nd))
            conn.commit()
            flash("Đã cập nhật phân quyền RBAC cho tài khoản.", "info")

    cur.execute("SELECT * FROM NguoiDung ORDER BY maNguoiDung ASC")
    users = [dict(r) for r in cur.fetchall()]
    conn.close()
    return render_template("users.html", users=users)


if __name__ == "__main__":
    init_db()
    app.run(host="127.0.0.1", port=5000, debug=False)
