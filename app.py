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


ROLE_DISPLAY_VN = {
    "Admin": "Quản trị viên",
    "CanBo": "Cán bộ xét duyệt",
    "TacGia": "Tác giả",
}


def role_required(*allowed_roles):
    """
    Decorator kiểm tra vai trò người dùng:
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
                role_vn = ROLE_DISPLAY_VN.get(user_role, user_role or "")
                req_vn = ", ".join(ROLE_DISPLAY_VN.get(r, r) for r in allowed_roles)
                flash(
                    f"Từ chối truy cập: Vai trò '{role_vn}' không có quyền thực hiện thao tác này (yêu cầu quyền: {req_vn}).",
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
            flash("Địa chỉ thư điện tử không tồn tại trong hệ thống.", "danger")
        elif user["trangThai"] == 0:
            flash("Tài khoản của bạn đã bị vô hiệu hóa. Vui lòng liên hệ Quản trị viên.", "danger")
        else:
            is_valid_pw = False
            try:
                if check_password_hash(user["matKhau"], mat_khau):
                    is_valid_pw = True
            except Exception:
                pass

            # Fallback nếu mật khẩu trong CSDL ở dạng plain-text -> tự động nâng cấp mã hóa
            if not is_valid_pw and user["matKhau"] == mat_khau:
                is_valid_pw = True
                try:
                    hashed_pw = generate_password_hash(mat_khau)
                    up_conn = get_db_connection()
                    up_conn.execute("UPDATE NguoiDung SET matKhau = ? WHERE maNguoiDung = ?", (hashed_pw, user["maNguoiDung"]))
                    up_conn.commit()
                    up_conn.close()
                except Exception as ex:
                    print(f"[Auth Auto-Upgrade Notice] {ex}")

            if not is_valid_pw:
                flash("Mật khẩu không chính xác. Vui lòng kiểm tra lại.", "danger")
            else:
                session.clear()
                session["user_id"] = user["maNguoiDung"]
                session["ho_ten"] = user["hoTen"]
                session["email"] = user["email"]
                session["vai_tro"] = user["vaiTro"]
                session["don_vi"] = user["donViCongTac"]
                session["hoc_vi"] = user["hocHamHocVi"]
                role_vn = ROLE_DISPLAY_VN.get(user["vaiTro"], user["vaiTro"])
                flash(f"Đăng nhập thành công ({role_vn}).", "success")
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

        is_current_valid = False
        try:
            if check_password_hash(user["matKhau"], mk_hien_tai):
                is_current_valid = True
        except Exception:
            pass
        if not is_current_valid and user["matKhau"] == mk_hien_tai:
            is_current_valid = True

        if not is_current_valid:
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
        SELECT COALESCE(tc.xepHang, h.chiMucTapChi, 'Chưa phân hạng') AS chiMuc, COUNT(*) AS soLuong
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Bài báo%' OR lct.tenLoai LIKE '%Bài báo%')
        GROUP BY COALESCE(tc.xepHang, h.chiMucTapChi, 'Chưa phân hạng')
        ORDER BY soLuong DESC
    """)
    stats_baibao_chimuc = [dict(r) for r in cur.fetchall()]

    # Thống kê chuyên sâu nhóm ĐỀ TÀI NCKH (Chuẩn hóa theo 12 trường dữ liệu đầu vào)
    cur.execute("""
        SELECT lct.tenLoai, COUNT(*) AS soLuong
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')
        GROUP BY lct.maLoai, lct.tenLoai
        ORDER BY soLuong DESC
    """)
    stats_detai_loai = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT h.donViChuTri, MAX(h.sanPhamDuKien) AS sanPhamDuKien, COUNT(*) AS soLuong
        FROM HoSoKHCN h
        JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        WHERE h.isDeleted = 0 AND (lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')
        GROUP BY h.donViChuTri
        ORDER BY soLuong DESC
    """)
    stats_detai_donvi = [dict(r) for r in cur.fetchall()]

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
        stats_detai_loai=stats_detai_loai,
        stats_detai_donvi=stats_detai_donvi,
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
    """Tuần 4 - Mục 2b: Tra cứu, tìm kiếm từ khóa và lọc hồ sơ theo Lĩnh vực, Loại, Tác giả, Đơn vị, Trạng thái có phân trang."""
    q = request.args.get("q", "").strip()
    nhom = request.args.get("nhom", "").strip()
    ma_linh_vuc = request.args.get("maLinhVuc", "").strip()
    ma_loai = request.args.get("maLoai", "").strip()
    tac_gia = request.args.get("tacGia", "").strip()
    don_vi = request.args.get("donVi", "").strip()
    trang_thai = request.args.get("trangThai", "").strip()
    canh_bao = request.args.get("coCanhBao", "").strip()
    sap_xep = request.args.get("sapXep", "priority").strip()
    page = max(1, int(request.args.get("page", 1)))
    per_page = 6

    conn = get_db_connection()
    cur = conn.cursor()

    where_clauses = ["h.isDeleted = 0"]
    params = []

    if q:
        where_clauses.append("(h.tieuDe LIKE ? OR h.tomTat LIKE ? OR h.tuKhoa LIKE ? OR h.thongTinTacGia LIKE ? OR h.maHoSo LIKE ? OR h.donViChuTri LIKE ? OR h.sanPhamDuKien LIKE ?)")
        like_q = f"%{q}%"
        params.extend([like_q, like_q, like_q, like_q, like_q, like_q, like_q])

    if nhom == "baibao":
        where_clauses.append("(lct.nhomCongTrinh LIKE '%Bài báo%' OR lct.tenLoai LIKE '%Bài báo%')")
    elif nhom == "detai":
        where_clauses.append("(lct.nhomCongTrinh LIKE '%Đề tài%' OR lct.tenLoai LIKE '%Đề tài%')")
    elif nhom == "sangche":
        where_clauses.append("(lct.nhomCongTrinh LIKE '%Sở hữu%' OR lct.nhomCongTrinh LIKE '%Sáng chế%' OR lct.tenLoai LIKE '%Sáng chế%' OR lct.tenLoai LIKE '%Giải pháp%')")

    if ma_linh_vuc:
        where_clauses.append("h.maLinhVuc = ?")
        params.append(int(ma_linh_vuc))

    if ma_loai:
        where_clauses.append("h.maLoai = ?")
        params.append(int(ma_loai))

    if tac_gia:
        where_clauses.append("h.thongTinTacGia LIKE ?")
        params.append(f"%{tac_gia}%")

    if don_vi:
        where_clauses.append("h.donViChuTri = ?")
        params.append(don_vi)

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

    if sap_xep == "fifo":
        order_sql = "h.ngayNop ASC"
    elif sap_xep == "newest":
        order_sql = "h.ngayNop DESC"
    else:  # priority (mặc định theo Nghiệp vụ 6: Ưu tiên rủi ro cao xử lý trước)
        order_sql = "h.diemBatThuong DESC, h.coCanhBao DESC, h.ngayNop DESC"

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
        ORDER BY {order_sql}
        LIMIT ? OFFSET ?
    """, params + [per_page, offset])
    raw_records = [dict(r) for r in cur.fetchall()]
    records = []
    for r in raw_records:
        try:
            raw_cb = json.loads(r.get("chiTietCanhBao") or "{}")
            if isinstance(raw_cb, dict):
                r["reasons"] = raw_cb.get("reasons", [])
                r["if_score"] = raw_cb.get("if_score", 0.0)
                r["ae_mse"] = raw_cb.get("ae_mse", 0.0)
                r["ae_score"] = raw_cb.get("ae_score", 0.0)
            elif isinstance(raw_cb, list):
                r["reasons"] = raw_cb
                r["if_score"] = 0.0
                r["ae_mse"] = 0.0
                r["ae_score"] = 0.0
            else:
                r["reasons"] = []
                r["if_score"] = 0.0
                r["ae_mse"] = 0.0
                r["ae_score"] = 0.0
        except Exception:
            r["reasons"] = []
            r["if_score"] = 0.0
            r["ae_mse"] = 0.0
            r["ae_score"] = 0.0
        records.append(r)

    cur.execute("SELECT * FROM LinhVuc WHERE trangThai = 1 ORDER BY maLinhVuc")
    linh_vucs = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM LoaiCongTrinh WHERE trangThai = 1 ORDER BY maLoai")
    loai_cts = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT * FROM DanhMucTacGia WHERE trangThai = 1 ORDER BY maTacGia")
    tac_gias = [dict(r) for r in cur.fetchall()]

    cur.execute("SELECT DISTINCT donViChuTri FROM HoSoKHCN WHERE isDeleted = 0 AND donViChuTri IS NOT NULL ORDER BY donViChuTri")
    don_vis = [r["donViChuTri"] for r in cur.fetchall()]

    conn.close()

    return render_template(
        "hoso_list.html",
        records=records,
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        tac_gias=tac_gias,
        don_vis=don_vis,
        filters={
            "q": q,
            "nhom": nhom,
            "maLinhVuc": ma_linh_vuc,
            "maLoai": ma_loai,
            "tacGia": tac_gia,
            "donVi": don_vi,
            "trangThai": trang_thai,
            "coCanhBao": canh_bao,
            "sapXep": sap_xep,
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
        don_vi_chu_tri = request.form.get("donViChuTri", "").strip() or session.get("don_vi") or "Khoa Công nghệ Thông tin"
        thong_tin_tac_gia = request.form.get("thongTinTacGia", "").strip() or session.get("ho_ten") or don_vi_chu_tri
        nam_cong_bo = int(request.form.get("namCongBo", 2026))
        ky_bao_cao = request.form.get("kyBaoCao", "Đợt 1 - Năm học 2025-2026").strip()
        ma_linh_vuc = int(request.form.get("maLinhVuc", 1))
        ma_loai = int(request.form.get("maLoai", 1))
        ma_tap_chi = request.form.get("maTapChi")
        ma_tap_chi = int(ma_tap_chi) if ma_tap_chi else None
        chi_muc_tap_chi = request.form.get("chiMucTapChi", "").strip() or None
        san_pham_du_kien = request.form.get("sanPhamDuKien", "").strip() or None

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

        # Chạy mô-đun AI Tiền xử lý tiếng Việt + Phân loại tự động Lĩnh vực & Loại công trình (Chức năng 5)
        pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia, don_vi_chu_tri)

        # Trạng thái ban đầu mặc định là 'Chờ xét duyệt' (Tuần 4 - Mục 2a)
        trang_thai_ban_dau = "Chờ xét duyệt"

        cur.execute("""
            INSERT INTO HoSoKHCN (
                maHoSo, tieuDe, tomTat, tuKhoa, thongTinTacGia, donViChuTri,
                namCongBo, kyBaoCao, maNguoiNop, maLinhVuc, maLoai, maTapChi,
                trangThai, fileDinhKem, vanBanDaTachTu, maLinhVucAI, doTinCayAI,
                maLinhVucSVM, doTinCaySVM,
                nhomCongTrinhAI, doTinCayLoaiAI, nhomCongTrinhSVM, doTinCayLoaiSVM,
                chiMucTapChi, sanPhamDuKien
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ma_ho_so, tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia, don_vi_chu_tri,
            nam_cong_bo, ky_bao_cao, session["user_id"], ma_linh_vuc, ma_loai, ma_tap_chi,
            trang_thai_ban_dau, file_name_saved, pred["segmented_text"],
            pred["phobert_label_id"], pred["phobert_confidence"],
            pred["svm_label_id"], pred["svm_confidence"],
            pred["work_type_phobert"], pred["work_type_phobert_conf"],
            pred["work_type_svm"], pred["work_type_svm_conf"],
            chi_muc_tap_chi, san_pham_du_kien
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
            f"Độ tin cậy gợi ý phân loại tự động: {pred['phobert_confidence']}%.",
            "success"
        )
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    cur.execute("SELECT * FROM LinhVuc WHERE trangThai = 1 ORDER BY maLinhVuc")
    linh_vucs = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM LoaiCongTrinh WHERE trangThai = 1 ORDER BY maLoai")
    loai_cts = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM TapChiHopLe WHERE trangThai = 1 ORDER BY maTapChi")
    tap_chis = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM DanhMucTacGia WHERE trangThai = 1 ORDER BY maTacGia")
    tac_gias = [dict(r) for r in cur.fetchall()]
    conn.close()

    return render_template(
        "hoso_form.html",
        mode="create",
        hoso=None,
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        tap_chis=tap_chis,
        tac_gias=tac_gias,
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
        don_vi_chu_tri = request.form.get("donViChuTri", "").strip() or hoso["donViChuTri"]
        thong_tin_tac_gia = request.form.get("thongTinTacGia", "").strip() or hoso["thongTinTacGia"] or session.get("ho_ten") or don_vi_chu_tri
        nam_cong_bo = int(request.form.get("namCongBo", hoso["namCongBo"]))
        ky_bao_cao = request.form.get("kyBaoCao", hoso["kyBaoCao"]).strip()
        ma_linh_vuc = int(request.form.get("maLinhVuc", hoso["maLinhVuc"]))
        ma_loai = int(request.form.get("maLoai", hoso["maLoai"]))
        ma_tap_chi = request.form.get("maTapChi")
        ma_tap_chi = int(ma_tap_chi) if ma_tap_chi else None
        chi_muc_tap_chi = request.form.get("chiMucTapChi", "").strip() or None
        san_pham_du_kien = request.form.get("sanPhamDuKien", "").strip() or None

        file_name_saved = hoso["fileDinhKem"]
        uploaded_file = request.files.get("fileDinhKem")
        if uploaded_file and uploaded_file.filename and allowed_file(uploaded_file.filename):
            safe_name = secure_filename(uploaded_file.filename)
            file_name_saved = f"{ma_ho_so}_{safe_name}"
            uploaded_file.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name_saved))

        # Chạy lại AI phân loại Lĩnh vực & Loại công trình + phát hiện bất thường sau khi cập nhật nội dung
        pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia, don_vi_chu_tri)

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
                nhomCongTrinhAI = ?, doTinCayLoaiAI = ?, nhomCongTrinhSVM = ?, doTinCayLoaiSVM = ?,
                chiMucTapChi = ?, sanPhamDuKien = ?,
                ngayCapNhat = CURRENT_TIMESTAMP
            WHERE maHoSo = ?
        """, (
            tieu_de, tom_tat, tu_khoa, thong_tin_tac_gia, don_vi_chu_tri,
            nam_cong_bo, ky_bao_cao, ma_linh_vuc, ma_loai, ma_tap_chi,
            trang_thai_moi,
            file_name_saved, pred["segmented_text"], pred["phobert_label_id"],
            pred["phobert_confidence"], pred["svm_label_id"], pred["svm_confidence"],
            pred["work_type_phobert"], pred["work_type_phobert_conf"],
            pred["work_type_svm"], pred["work_type_svm_conf"],
            chi_muc_tap_chi, san_pham_du_kien,
            ma_ho_so
        ))

        conn.commit()
        ai_engine.build_networkx_graph_and_detect_anomalies(conn)
        conn.close()
        if hoso["trangThai"] == "Yêu cầu bổ sung":
            flash(f"Đã cập nhật hồ sơ {ma_ho_so} và chuyển trạng thái về 'Chờ xét duyệt' để thẩm định lại.", "success")
        else:
            flash(f"Đã cập nhật thông tin hồ sơ {ma_ho_so} và phân tích lại tự động.", "success")
        return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    cur.execute("SELECT * FROM LinhVuc WHERE trangThai = 1 ORDER BY maLinhVuc")
    linh_vucs = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM LoaiCongTrinh WHERE trangThai = 1 ORDER BY maLoai")
    loai_cts = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM TapChiHopLe WHERE trangThai = 1 ORDER BY maTapChi")
    tap_chis = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM DanhMucTacGia WHERE trangThai = 1 ORDER BY maTacGia")
    tac_gias = [dict(r) for r in cur.fetchall()]
    conn.close()

    return render_template(
        "hoso_form.html",
        mode="edit",
        hoso=dict(hoso),
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        tap_chis=tap_chis,
        tac_gias=tac_gias,
    )


@app.route("/hoso/<ma_ho_so>/soft-delete", methods=["POST"])
@login_required
def hoso_soft_delete(ma_ho_so):
    """Chuyển hồ sơ vào thùng rác khi ở trạng thái chờ duyệt."""
    conn = get_db_connection()
    hoso = conn.execute("SELECT * FROM HoSoKHCN WHERE maHoSo = ?", (ma_ho_so,)).fetchone()
    if not hoso:
        conn.close()
        flash("Hồ sơ không tồn tại.", "danger")
        return redirect(url_for("hoso_list"))

    if session.get("vai_tro") != "Admin":
        if hoso["trangThai"] != "Chờ xét duyệt" or hoso["maNguoiNop"] != session.get("user_id"):
            conn.close()
            flash("Bạn chỉ được phép xóa tạm thời hồ sơ của mình khi đang ở trạng thái 'Chờ xét duyệt'.", "danger")
            return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))

    conn.execute("UPDATE HoSoKHCN SET isDeleted = 1 WHERE maHoSo = ?", (ma_ho_so,))
    conn.commit()
    conn.close()
    flash(f"Đã chuyển hồ sơ {ma_ho_so} vào thùng rác.", "info")
    return redirect(url_for("hoso_list"))


@app.route("/hoso/<ma_ho_so>/hard-delete", methods=["POST"])
@role_required("Admin")
def hoso_hard_delete(ma_ho_so):
    """Xóa vĩnh viễn dữ liệu hồ sơ (chỉ dành cho Quản trị viên)."""
    conn = get_db_connection()
    conn.execute("DELETE FROM HoSoKHCN WHERE maHoSo = ?", (ma_ho_so,))
    conn.commit()
    conn.close()
    flash(f"Đã xóa vĩnh viễn hồ sơ {ma_ho_so} khỏi hệ thống.", "warning")
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
        raw_cb = json.loads(hoso.get("chiTietCanhBao") or "{}")
        if isinstance(raw_cb, dict):
            hoso["canhBaoList"] = raw_cb.get("reasons", [])
            hoso["anomalyDetails"] = raw_cb
        elif isinstance(raw_cb, list):
            hoso["canhBaoList"] = raw_cb
            hoso["anomalyDetails"] = {"reasons": raw_cb}
        else:
            hoso["canhBaoList"] = []
            hoso["anomalyDetails"] = {"reasons": []}
    except Exception:
        hoso["canhBaoList"] = []
        hoso["anomalyDetails"] = {"reasons": []}

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
    cur.execute("SELECT * FROM LoaiCongTrinh WHERE trangThai = 1 ORDER BY maLoai")
    loai_cts = [dict(r) for r in cur.fetchall()]

    conn.close()

    badge_info = ai_engine.get_confidence_badge_info(float(hoso.get("doTinCayAI") or 0.0))

    return render_template(
        "hoso_detail.html",
        hoso=hoso,
        lich_su=lich_su,
        phan_hoi_ai=phan_hoi_ai,
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
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
        f"Đã cập nhật trạng thái hồ sơ {ma_ho_so} từ '{trang_thai_cu}' sang '{trang_thai_moi}'.",
        "success"
    )
    return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))


@app.route("/hoso/<ma_ho_so>/dismiss-warning", methods=["POST"])
@role_required("Admin", "CanBo")
def hoso_dismiss_warning(ma_ho_so):
    """Bỏ qua cảnh báo bất thường sau khi cán bộ thẩm định thực tế."""
    ly_do = request.form.get("lyDoBoQua", "Cán bộ xét duyệt đã thẩm định thực tế và xác nhận hồ sơ hợp lệ.").strip()
    conn = get_db_connection()
    cur = conn.cursor()
    hoso = cur.execute("SELECT coCanhBao, diemBatThuong FROM HoSoKHCN WHERE maHoSo = ? AND isDeleted = 0", (ma_ho_so,)).fetchone()
    if not hoso:
        conn.close()
        flash("Hồ sơ không tồn tại.", "danger")
        return redirect(url_for("hoso_list"))

    cur.execute("""
        UPDATE HoSoKHCN
        SET coCanhBao = 0, ngayCapNhat = CURRENT_TIMESTAMP
        WHERE maHoSo = ?
    """, (ma_ho_so,))

    cur.execute("""
        INSERT INTO LichSuXetDuyet (maHoSo, maCanBo, trangThaiCu, trangThaiMoi, yKienDanhGia)
        VALUES (?, ?, ?, ?, ?)
    """, (ma_ho_so, session["user_id"], "Cảnh báo bất thường", "Đã giải trừ cảnh báo", f"Bỏ qua cảnh báo: {ly_do}"))

    conn.commit()
    conn.close()
    flash(f"Đã giải trừ cờ cảnh báo bất thường cho hồ sơ {ma_ho_so}.", "success")
    return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))


@app.route("/hoso/<ma_ho_so>/ai-feedback", methods=["POST"])
@role_required("Admin", "CanBo")
def hoso_ai_feedback(ma_ho_so):
    """Cơ chế xác nhận hoặc điều chỉnh lại nhãn phân loại tự động từ cán bộ xét duyệt."""
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
            ghi_chu = "Cán bộ xét duyệt xác nhận nhãn lĩnh vực dự đoán tự động là chính xác."
    else:
        ma_lv_thuc_te = int(request.form.get("maLinhVucThucTe", ma_lv_ai))
        ket_qua_khop = 1 if ma_lv_thuc_te == ma_lv_ai else 0
        if not ghi_chu:
            ghi_chu = "Cán bộ xét duyệt điều chỉnh lại nhãn lĩnh vực chuẩn phục vụ huấn luyện bổ sung."

    # Lưu vào bảng PhanHoiAI
    cur.execute("""
        INSERT INTO PhanHoiAI
        (maHoSo, maCanBo, maLinhVucAI, maLinhVucThucTe, doTinCayLucDo, ketQuaKhop, vanBanHuanLuyen, ghiChuPhanHoi)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        ma_ho_so, session["user_id"], ma_lv_ai, ma_lv_thuc_te,
        do_tin_cay, ket_qua_khop, van_ban_tach_tu, ghi_chu
    ))

    # Cập nhật lĩnh vực (và loại công trình nếu cán bộ điều chỉnh) của hồ sơ
    ma_loai_moi = request.form.get("maLoaiThucTe", "").strip()
    if ma_loai_moi:
        cur.execute("""
            UPDATE HoSoKHCN
            SET maLinhVuc = ?, maLoai = ?, daXacNhanAI = 1, ngayCapNhat = CURRENT_TIMESTAMP
            WHERE maHoSo = ?
        """, (ma_lv_thuc_te, int(ma_loai_moi), ma_ho_so))
    else:
        cur.execute("""
            UPDATE HoSoKHCN
            SET maLinhVuc = ?, daXacNhanAI = 1, ngayCapNhat = CURRENT_TIMESTAMP
            WHERE maHoSo = ?
        """, (ma_lv_thuc_te, ma_ho_so))

    conn.commit()
    ai_engine.train_and_evaluate(conn)
    ai_engine.build_networkx_graph_and_detect_anomalies(conn)
    conn.close()

    if ket_qua_khop == 1:
        flash("Đã xác nhận nhãn phân loại gợi ý và lưu phản hồi.", "success")
    else:
        flash("Đã cập nhật nhãn phân loại thực tế và lưu dữ liệu học bổ sung.", "info")

    return redirect(url_for("hoso_detail", ma_ho_so=ma_ho_so))


# =============================================================================
# API GỢI Ý NHANH AI TRỰC TIẾP TRÊN FORM KHAI BÁO (LĨNH VỰC & LOẠI CÔNG TRÌNH)
# =============================================================================
@app.route("/api/ai-preview", methods=["POST"])
@login_required
def api_ai_preview():
    data = request.get_json(silent=True) or {}
    tieu_de = data.get("tieuDe", "")
    tom_tat = data.get("tomTat", "")
    tu_khoa = data.get("tuKhoa", "")
    tac_gia = data.get("thongTinTacGia", "")
    don_vi = data.get("donViChuTri", "")

    pred = ai_engine.predict_record_field(tieu_de, tom_tat, tu_khoa, tac_gia, don_vi)

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
        "work_type_phobert_id": pred["work_type_phobert_id"],
        "work_type_phobert": pred["work_type_phobert"],
        "work_type_phobert_conf": pred["work_type_phobert_conf"],
        "work_type_svm": pred["work_type_svm"],
        "work_type_svm_conf": pred["work_type_svm_conf"],
        "infer_ms": pred["infer_ms"],
        "badge": pred["badge"],
    })


# =============================================================================
# TUẦN 5 - MỤC 1: QUẢN LÝ 4 DANH MỤC HỆ THỐNG (CHỨC NĂNG 2 - CRUD ĐẦY ĐỦ)
# =============================================================================
@app.route("/danh-muc")
@login_required
def catalog_management():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT lv.maLinhVuc, lv.maCode, lv.tenLinhVuc, lv.moTa, lv.tuKhoaDacTrung, lv.trangThai, COUNT(h.maHoSo) AS soHoSoLienKet
        FROM LinhVuc lv
        LEFT JOIN HoSoKHCN h ON lv.maLinhVuc = h.maLinhVuc AND h.isDeleted = 0
        GROUP BY lv.maLinhVuc, lv.maCode, lv.tenLinhVuc, lv.moTa, lv.tuKhoaDacTrung, lv.trangThai
        ORDER BY lv.maLinhVuc
    """)
    linh_vucs = [dict(r) for r in cur.fetchall()]

    cur.execute("""
        SELECT lct.maLoai, lct.maCodeLoai, lct.tenLoai, lct.nhomCongTrinh, lct.gioChuanQuyDoi, lct.trangThai, COUNT(h.maHoSo) AS soHoSoLienKet
        FROM LoaiCongTrinh lct
        LEFT JOIN HoSoKHCN h ON lct.maLoai = h.maLoai AND h.isDeleted = 0
        GROUP BY lct.maLoai, lct.maCodeLoai, lct.tenLoai, lct.nhomCongTrinh, lct.gioChuanQuyDoi, lct.trangThai
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

    cur.execute("""
        SELECT tg.*, lv.tenLinhVuc,
               (SELECT COUNT(*) FROM HoSoKHCN h WHERE h.isDeleted = 0 AND h.thongTinTacGia LIKE '%' || tg.hoTen || '%') AS soHoSoLienKet
        FROM DanhMucTacGia tg
        LEFT JOIN LinhVuc lv ON tg.maLinhVucChinh = lv.maLinhVuc
        ORDER BY tg.maTacGia
    """)
    tac_gias = [dict(r) for r in cur.fetchall()]

    conn.close()
    return render_template(
        "catalogs.html",
        linh_vucs=linh_vucs,
        loai_cts=loai_cts,
        tap_chis=tap_chis,
        tac_gias=tac_gias,
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


@app.route("/danh-muc/tac-gia/add", methods=["POST"])
@role_required("Admin")
def catalog_add_tacgia():
    ma_code = request.form.get("maCodeTG", "").strip().upper()
    ho_ten = request.form.get("hoTen", "").strip()
    hoc_vi = request.form.get("hocHamHocVi", "TS.").strip()
    email = request.form.get("email", "").strip()
    don_vi = request.form.get("donViCongTac", "").strip()
    ma_lv = int(request.form.get("maLinhVucChinh", 1))
    orcid = request.form.get("orcid", "").strip()

    if not ma_code or not ho_ten or not don_vi:
        flash("Mã tác giả, Họ tên và Đơn vị công tác không được để trống.", "danger")
        return redirect(url_for("catalog_management"))

    conn = get_db_connection()
    try:
        conn.execute(
            """INSERT INTO DanhMucTacGia
               (maCodeTG, hoTen, hocHamHocVi, email, donViCongTac, maLinhVucChinh, orcid, trangThai)
               VALUES (?, ?, ?, ?, ?, ?, ?, 1)""",
            (ma_code, ho_ten, hoc_vi, email, don_vi, ma_lv, orcid),
        )
        conn.commit()
        flash(f"Đã thêm tác giả '{ho_ten}' vào Danh mục Tác giả.", "success")
    except Exception as e:
        flash(f"Không thể thêm tác giả (Mã có thể bị trùng): {e}", "danger")
    finally:
        conn.close()
    return redirect(url_for("catalog_management"))


@app.route("/danh-muc/edit/<cat_type>/<int:item_id>", methods=["POST"])
@role_required("Admin")
def catalog_edit_item(cat_type, item_id):
    """Cập nhật (Update) thông tin chi tiết cho cả 4 nhóm danh mục hệ thống."""
    conn = get_db_connection()
    try:
        if cat_type == "linhvuc":
            ma_code = request.form.get("maCode", "").strip().upper()
            ten_lv = request.form.get("tenLinhVuc", "").strip()
            mo_ta = request.form.get("moTa", "").strip()
            conn.execute(
                "UPDATE LinhVuc SET maCode = ?, tenLinhVuc = ?, moTa = ? WHERE maLinhVuc = ?",
                (ma_code, ten_lv, mo_ta, item_id),
            )
            flash(f"Đã cập nhật danh mục lĩnh vực '{ten_lv}'.", "success")
        elif cat_type == "loai":
            ma_code = request.form.get("maCodeLoai", "").strip().upper()
            ten_loai = request.form.get("tenLoai", "").strip()
            nhom = request.form.get("nhomCongTrinh", "Bài báo quốc tế").strip()
            gio_chuan = float(request.form.get("gioChuanQuyDoi", 60.0))
            conn.execute(
                "UPDATE LoaiCongTrinh SET maCodeLoai = ?, tenLoai = ?, nhomCongTrinh = ?, gioChuanQuyDoi = ? WHERE maLoai = ?",
                (ma_code, ten_loai, nhom, gio_chuan, item_id),
            )
            flash(f"Đã cập nhật loại công trình '{ten_loai}'.", "success")
        elif cat_type == "tapchi":
            ten_tc = request.form.get("tenTapChi", "").strip()
            issn = request.form.get("chiSoISSN", "").strip()
            xep_hang = request.form.get("xepHang", "HĐGSNN").strip()
            uy_tin = int(request.form.get("uyTin", 1))
            conn.execute(
                "UPDATE TapChiHopLe SET tenTapChi = ?, chiSoISSN = ?, xepHang = ?, uyTin = ? WHERE maTapChi = ?",
                (ten_tc, issn, xep_hang, uy_tin, item_id),
            )
            flash(f"Đã cập nhật thông tin tạp chí '{ten_tc}'.", "success")
        elif cat_type == "tacgia":
            ma_code = request.form.get("maCodeTG", "").strip().upper()
            ho_ten = request.form.get("hoTen", "").strip()
            hoc_vi = request.form.get("hocHamHocVi", "TS.").strip()
            email = request.form.get("email", "").strip()
            don_vi = request.form.get("donViCongTac", "").strip()
            ma_lv = int(request.form.get("maLinhVucChinh", 1))
            conn.execute(
                """UPDATE DanhMucTacGia
                   SET maCodeTG = ?, hoTen = ?, hocHamHocVi = ?, email = ?, donViCongTac = ?, maLinhVucChinh = ?
                   WHERE maTacGia = ?""",
                (ma_code, ho_ten, hoc_vi, email, don_vi, ma_lv, item_id),
            )
            flash(f"Đã cập nhật thông tin tác giả '{ho_ten}'.", "success")
        else:
            abort(400)
        conn.commit()
    except Exception as e:
        flash(f"Lỗi khi cập nhật danh mục: {e}", "danger")
    finally:
        conn.close()
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
        "tacgia": ("DanhMucTacGia", "maTacGia"),
    }
    if cat_type not in table_map:
        abort(400)
    tbl, pk = table_map[cat_type]
    conn = get_db_connection()
    conn.execute(f"UPDATE {tbl} SET trangThai = CASE WHEN trangThai = 1 THEN 0 ELSE 1 END WHERE {pk} = ?", (item_id,))
    conn.commit()
    conn.close()
    flash("Đã cập nhật trạng thái hiển thị của danh mục.", "info")
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
        SELECT h.*, lv.tenLinhVuc, lv_ai.tenLinhVuc AS tenLinhVucAI,
               lct.tenLoai, lct.nhomCongTrinh,
               tc.tenTapChi, tc.chiSoISSN, tc.xepHang
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        WHERE {where_sql}
        ORDER BY h.maHoSo ASC
    """, params).fetchall()
    conn.close()
    filter_desc = ", ".join(desc_parts) if desc_parts else "Toàn bộ hồ sơ KH&CN"
    return [dict(r) for r in rows], filter_desc


@app.route("/hoso/import-excel", methods=["POST"])
@login_required
def hoso_import_excel():
    """Nạp hàng loạt hồ sơ KH&CN từ tệp Excel chuẩn và tự động kích hoạt AI phân loại & phát hiện bất thường."""
    file = request.files.get("excel_file")
    default_excel = os.path.join(BASE_DIR, "docs", "Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
    
    if (not file or file.filename == "") and os.path.exists(default_excel):
        target_path = default_excel
    elif file and file.filename != "":
        upload_folder = os.path.join(BASE_DIR, "uploads")
        os.makedirs(upload_folder, exist_ok=True)
        target_path = os.path.join(upload_folder, "uploaded_import.xlsx")
        file.save(target_path)
    else:
        flash("Vui lòng chọn tệp Excel (.xlsx) để nạp dữ liệu.", "warning")
        return redirect(url_for("hoso_list"))

    try:
        import subprocess
        res = subprocess.run(
            [sys.executable, os.path.join(BASE_DIR, "scripts", "import_mau_nhap_lieu.py")],
            capture_output=True, text=True, encoding="utf-8"
        )
        if res.returncode == 0:
            flash("Đã nạp thành công toàn bộ danh sách hồ sơ từ Excel! Mô hình VinAI PhoBERT đã tự động điền lĩnh vực nghiên cứu cho các đề tài bị khuyết và hoàn tất quét bất thường đa tầng.", "success")
        else:
            flash(f"Quá trình nạp dữ liệu gặp lỗi: {res.stderr[:200]}", "danger")
    except Exception as e:
        flash(f"Lỗi khi xử lý nạp Excel: {e}", "danger")

    return redirect(url_for("hoso_list"))


@app.route("/hoso/download-template")
@login_required
def download_excel_template():
    """Tải xuống tệp Excel mẫu nhập liệu chuẩn 24 trường thông tin."""
    excel_path = os.path.join(BASE_DIR, "docs", "Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
    if os.path.exists(excel_path):
        return send_file(excel_path, as_attachment=True, download_name="Mau_Nhap_Lieu_HoSo_KHCN.xlsx")
    flash("Không tìm thấy tệp mẫu trên máy chủ.", "warning")
    return redirect(url_for("hoso_list"))


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
    req_mode = request.args.get("mode", "").strip()
    if req_mode in ("svm", "phobert", "both"):
        ai_engine.last_experiment_mode = req_mode
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

    # Đảm bảo đồ thị tri thức và mô hình bất thường luôn sẵn sàng
    if ai_engine.knowledge_graph.number_of_nodes() == 0:
        ai_engine.build_networkx_graph_and_detect_anomalies(conn)

    cur.execute("""
        SELECT h.*, lv.tenLinhVuc, lv_ai.tenLinhVuc AS tenLinhVucAI, tc.tenTapChi, tc.uyTin AS tapChiUyTin,
               nd.hoTen AS tenNguoiNop
        FROM HoSoKHCN h
        LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
        LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
        LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
        LEFT JOIN NguoiDung nd ON h.maNguoiNop = nd.maNguoiDung
        WHERE h.isDeleted = 0 AND h.coCanhBao = 1
        ORDER BY h.diemBatThuong DESC, h.ngayNop ASC
    """)
    anomalies = []
    for r in cur.fetchall():
        d = dict(r)
        try:
            raw_cb = json.loads(d.get("chiTietCanhBao") or "{}")
            if isinstance(raw_cb, dict):
                d["reasons"] = raw_cb.get("reasons", [])
                d["if_score"] = raw_cb.get("if_score", 0.0)
                d["ae_mse"] = raw_cb.get("ae_mse", 0.0)
                d["ae_score"] = raw_cb.get("ae_score", 0.0)
            elif isinstance(raw_cb, list):
                d["reasons"] = raw_cb
                d["if_score"] = 0.0
                d["ae_mse"] = 0.0
                d["ae_score"] = 0.0
            else:
                d["reasons"] = []
                d["if_score"] = 0.0
                d["ae_mse"] = 0.0
                d["ae_score"] = 0.0
        except Exception:
            d["reasons"] = []
            d["if_score"] = 0.0
            d["ae_mse"] = 0.0
            d["ae_score"] = 0.0
        anomalies.append(d)

    # Chức năng 6: Bảng so sánh 2 mô hình bất thường & Hàng đợi xét duyệt ưu tiên
    anomaly_models = ai_engine.get_anomaly_models_comparison()
    priority_queue = ai_engine.get_priority_queue(cur, limit=20)
    graph_summary = ai_engine.get_graph_summary()
    benchmark = ai_engine.get_benchmark_summary()
    conn.close()

    return render_template(
        "ai_center.html",
        metrics=metrics,
        feedbacks=feedbacks,
        anomalies=anomalies,
        anomaly_models=anomaly_models,
        priority_queue=priority_queue,
        graph_summary=graph_summary,
        phobert_params=PHOBERT_HYPERPARAMS,
        benchmark=benchmark,
        active_mode=ai_engine.last_experiment_mode,
        selected_primary_model=ai_engine.selected_primary_model,
    )


@app.route("/ai-center/retrain", methods=["POST"])
@role_required("Admin", "CanBo")
def ai_retrain():
    mode = request.form.get("mode", "both").strip()
    conn = get_db_connection()
    t0 = datetime.now()
    bench = ai_engine.run_standalone_experiment(mode=mode, db_conn=conn)
    if ai_engine.knowledge_graph.number_of_nodes() == 0 or mode == "retrain_full":
        ai_engine.build_networkx_graph_and_detect_anomalies(conn)
    conn.close()
    elapsed_sec = round((datetime.now() - t0).total_seconds(), 2)

    if mode == "svm":
        svm_m = bench["svm_metrics"]
        flash(
            f"Đã chạy riêng mô hình TF-IDF + SVM trong {elapsed_sec} giây: "
            f"Độ chuẩn xác = {svm_m['precision']}%, Độ phủ = {svm_m['recall']}%, Điểm F1 = {svm_m['f1_score']}% ({svm_m['infer_time_ms']} ms/hồ sơ).",
            "info"
        )
    elif mode == "phobert":
        pho_m = bench["phobert_metrics"]
        flash(
            f"Đã chạy riêng mô hình Tinh chỉnh PhoBERT trong {elapsed_sec} giây: "
            f"Độ chuẩn xác = {pho_m['precision']}%, Độ phủ = {pho_m['recall']}%, Điểm F1 = {pho_m['f1_score']}% ({pho_m['infer_time_ms']} ms/hồ sơ).",
            "success"
        )
    elif mode == "retrain_full":
        win = bench["selected_model"]
        flash(
            f"Đã huấn luyện lại toàn bộ hai mô hình từ đầu và cập nhật mô hình chính: {win['short_name']} (Điểm F1 đạt {win['f1_score']}%).",
            "success"
        )
    else:
        win = bench["selected_model"]
        flash(
            f"Đã chạy so sánh cả hai mô hình trong {elapsed_sec} giây. "
            f"Mô hình chính được chọn: {win['short_name']} (Điểm F1 đạt {win['f1_score']}%, cao hơn {win['delta_f1']:+.2f}%).",
            "success"
        )
    return redirect(url_for("ai_center"))


@app.route("/ai-center/download-dataset")
@login_required
def download_training_dataset():
    from model_benchmark import DATASET_XLSX_PATH, get_shared_train_test_split, _append_comparison_sheets_to_excel
    all_records, _, _ = get_shared_train_test_split()
    summary = ai_engine.get_benchmark_summary()
    _append_comparison_sheets_to_excel(all_records, summary)
    return send_file(
        DATASET_XLSX_PATH,
        as_attachment=True,
        download_name="Tap_Du_Lieu_Huan_Luyen_Chuan_12Truong.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =============================================================================
# QUẢN TRỊ NGƯỜI DÙNG & PHÂN QUYỀN (CHỈ QUẢN TRỊ VIÊN)
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
                role_vn = ROLE_DISPLAY_VN.get(vai_tro, vai_tro)
                flash(f"Đã tạo tài khoản '{ho_ten}' với vai trò '{role_vn}'.", "success")
            except Exception as e:
                flash(f"Không thể tạo tài khoản (thư điện tử có thể đã tồn tại): {e}", "danger")
        elif action == "update_role":
            ma_nd = int(request.form.get("maNguoiDung"))
            vai_tro_moi = request.form.get("vaiTroMoi")
            cur.execute("UPDATE NguoiDung SET vaiTro = ? WHERE maNguoiDung = ?", (vai_tro_moi, ma_nd))
            conn.commit()
            flash("Đã cập nhật phân quyền cho tài khoản.", "info")

    cur.execute("SELECT * FROM NguoiDung ORDER BY maNguoiDung ASC")
    users = [dict(r) for r in cur.fetchall()]
    conn.close()
    return render_template("users.html", users=users)


if __name__ == "__main__":
    init_db()
    app.run(host="127.0.0.1", port=5000, debug=False)
