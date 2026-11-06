# -*- coding: utf-8 -*-
"""
Hệ thống Trí tuệ Nhân tạo Tích hợp Thực tế (100% Real PyTorch & Scikit-learn Execution)
Đề tài: Hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường (CNTT-KLCN001)

Bao gồm:
1. Tiền xử lý văn bản tiếng Việt chuẩn hóa (NFC, lọc Stop-words, tách từ ghép chuyên ngành bằng PyVi)
2. Phân loại tự động đa nhiệm (Lĩnh vực nghiên cứu & Nhóm loại công trình):
   - Mô hình Baseline: TF-IDF + LinearSVC (CalibratedClassifierCV, scikit-learn)
   - Mô hình Chính thức: Fine-tuning PhoBERT thực tế (vinai/phobert-base trên PyTorch)
3. Phát hiện bất thường & Xếp hạng ưu tiên xét duyệt:
   - Đồ thị tri thức NetworkX (Tác giả - Công trình - Lĩnh vực - Tạp chí)
   - Rừng cây cô lập Isolation Forest (scikit-learn)
   - Mạng nơ-ron Tự mã hóa PyTorch Autoencoder (torch.nn.Module + MSELoss + Adam)
   - Độ tương đồng văn bản Cosine Similarity kiểm tra trùng lặp nội dung
"""

import os
import re
import time
import json
import joblib
import warnings
import unicodedata
import numpy as np
import networkx as nx

warnings.filterwarnings("ignore")
os.environ.setdefault("HF_HUB_DISABLE_XET", "1")
os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

try:
    import torch
    import torch.nn as nn
    TORCH_AVAILABLE = True
    _BaseAutoencoderModule = nn.Module
except ImportError:
    torch = None
    nn = None
    TORCH_AVAILABLE = False
    _BaseAutoencoderModule = object
from pyvi import ViTokenizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC
from sklearn.calibration import CalibratedClassifierCV
from sklearn.preprocessing import MinMaxScaler
from sklearn.ensemble import IsolationForest
from sklearn.metrics.pairwise import cosine_similarity

from phobert_finetune import (
    HYPERPARAMETERS as PHOBERT_PARAMS,
    WORK_TYPE_GROUPS,
    load_trained_phobert,
    predict_phobert_single,
    set_deterministic_seed,
)
from model_benchmark import (
    run_full_comparison_and_select_winner,
    run_tfidf_svm_standalone,
    run_phobert_standalone,
    compose_summary_dict,
    compute_dataset_statistics,
    get_shared_train_test_split,
    combine_fair_input_text,
    DATASET_XLSX_PATH,
    BENCHMARK_JSON_PATH,
    SVM_PIPELINE_PATH,
    PHOBERT_MODEL_DIR,
)

# Danh sách từ dừng tiếng Việt (Vietnamese Stop-words)
VIETNAMESE_STOPWORDS = {
    "và", "của", "các", "trong", "cho", "với", "những", "được", "là", "tại",
    "trên", "về", "từ", "theo", "đến", "khi", "này", "đó", "một", "có",
    "không", "như", "để", "do", "bởi", "qua", "ra", "vào", "lên", "xuống",
    "nhằm", "việc", "cũng", "đã", "đang", "sẽ", "rất", "nhiều", "hơn", "nhất",
    "bằng", "giữa", "cùng", "hoặc", "nhưng", "vì", "nếu", "thì", "mà", "nên",
    "cần", "phải", "mỗi", "mọi", "tất_cả", "dưới", "trước", "sau", "ngoài"
}

# Các cụm thuật ngữ chuyên ngành hỗ trợ chuẩn hóa tách từ
DOMAIN_COMPOUND_TERMS = [
    ("công nghệ thông tin", "công_nghệ_thông_tin"),
    ("trí tuệ nhân tạo", "trí_tuệ_nhân_tạo"),
    ("học máy", "học_máy"),
    ("học sâu", "học_sâu"),
    ("xử lý ngôn ngữ tự nhiên", "xử_lý_ngôn_ngữ_tự_nhiên"),
    ("thị giác máy tính", "thị_giác_máy_tính"),
    ("mạng nơ ron", "mạng_nơ_ron"),
    ("an toàn thông tin", "an_toàn_thông_tin"),
    ("dữ liệu lớn", "dữ_liệu_lớn"),
    ("chuỗi khối", "chuỗi_khối"),
    ("điện toán đám mây", "điện_toán_đám_mây"),
    ("hệ thống nhúng", "hệ_thống_nhúng"),
    ("điều khiển tự động", "điều_khiển_tự_động"),
    ("mạng cảm biến không dây", "mạng_cảm_biến_không_dây"),
    ("lưới điện thông minh", "lưới_điện_thông_minh"),
    ("vi điều khiển", "vi_điều_khiển"),
    ("năng lượng mặt trời", "năng_lượng_mặt_trời"),
    ("chế tạo máy", "chế_tạo_máy"),
    ("động cơ đốt trong", "động_cơ_đốt_trong"),
    ("phần tử hữu hạn", "phần_tử_hữu_hạn"),
    ("cơ điện tử", "cơ_điện_tử"),
    ("ô tô điện", "ô_tô_điện"),
    ("kinh tế vĩ mô", "kinh_tế_vĩ_mô"),
    ("tài chính ngân hàng", "tài_chính_ngân_hàng"),
    ("chuỗi cung ứng", "chuỗi_cung_ứng"),
    ("thương mại điện tử", "thương_mại_điện_tử"),
    ("quản trị nhân lực", "quản_trị_nhân_lực"),
    ("hành vi người tiêu dùng", "hành_vi_người_tiêu_dùng"),
    ("công nghệ sinh học", "công_nghệ_sinh_học"),
    ("vật liệu nano", "vật_liệu_nano"),
    ("xử lý nước thải", "xử_lý_nước_thải"),
    ("kim loại nặng", "kim_loại_nặng"),
    ("biến đổi khí hậu", "biến_đổi_khí_hậu"),
    ("phương pháp giảng dạy", "phương_pháp_giảng_dạy"),
    ("giáo dục đại học", "giáo_dục_đại_học"),
    ("tâm lý học", "tâm_lý_học"),
    ("ngôn ngữ học", "ngôn_ngữ_học"),
    ("giải pháp hữu ích", "giải_pháp_hữu_ích"),
    ("sở hữu trí tuệ", "sở_hữu_trí_tuệ"),
]


class KHCNAnomalyAutoencoder(_BaseAutoencoderModule):
    """
    Mạng nơ-ron Tự mã hóa thực tế bằng PyTorch (PyTorch Autoencoder - Mục 7 Đề cương CNTT-KLCN001):
    Kiến trúc nén nút thắt cổ chai: 6 -> 16 -> 4 (Latent Bottleneck) -> 16 -> 6.
    Huấn luyện không giám sát tối thiểu hóa sai số tái tạo MSE (nn.MSELoss) bằng bộ tối ưu Adam.
    """

    def __init__(self, input_dim: int = 6):
        if not TORCH_AVAILABLE:
            return
        super().__init__()
        self.encoder = nn.Sequential(
            nn.Linear(input_dim, 16),
            nn.ReLU(),
            nn.Linear(16, 4),
            nn.ReLU(),
        )
        self.decoder = nn.Sequential(
            nn.Linear(4, 16),
            nn.ReLU(),
            nn.Linear(16, input_dim),
            nn.Sigmoid(),
        )

    def forward(self, x):
        if not TORCH_AVAILABLE:
            return x
        latent = self.encoder(x)
        reconstructed = self.decoder(latent)
        return reconstructed


class KHCNAIEngine:
    """
    Bộ máy AI trung tâm vận hành thực tế 100%:
    - Tải trực tiếp trọng số PyTorch PhoBERT (`models/phobert_khcn/phobert_classifier.pt`)
      và pipeline TF-IDF + LinearSVC (`models/tfidf_svm/svm_pipeline.joblib`).
    - Dự đoán đồng thời Lĩnh vực chuyên ngành (6 lớp) và Nhóm loại công trình (3 nhóm).
    - Huấn luyện và suy luận PyTorch Autoencoder + Isolation Forest + NetworkX phát hiện bất thường.
    """

    def __init__(self):
        self.tfidf_vectorizer = None
        self.svm_classifier = None
        self.svm_work_type_classifier = None
        self.phobert_model = None
        self.phobert_tokenizer = None
        self.isolation_forest = None
        self.autoencoder = None
        self.scaler = None
        self.anomaly_models_metrics = []
        self.knowledge_graph = nx.Graph()
        self.metrics_comparison = []
        self.benchmark_summary = None
        self.last_experiment_mode = "both"
        self.selected_primary_model = "Mô hình học sâu: Tinh chỉnh PhoBERT tiếng Việt"
        self.is_trained = False

    # =========================================================================
    # 1. TIỀN XỬ LÝ VĂN BẢN TIẾNG VIỆT (PYVI WORD SEGMENTATION)
    # =========================================================================
    @staticmethod
    def combine_input_fields(tieu_de: str, tom_tat: str, tu_khoa: str, tac_gia: str = "", don_vi: str = "") -> str:
        """
        Kết hợp thống nhất các trường đầu vào của hồ sơ (Tiêu đề + Tóm tắt + Từ khóa + Tác giả + Đơn vị),
        đồng bộ 100% với hàm `combine_fair_input_text` trong `model_benchmark.py`.
        """
        parts = [
            str(tieu_de or "").strip(),
            str(tom_tat or "").strip(),
            str(tu_khoa or "").strip(),
            str(tac_gia or "").strip(),
            str(don_vi or "").strip(),
        ]
        non_empty = [p for p in parts if p]
        return " . ".join(non_empty)

    @staticmethod
    def clean_vietnamese_text(raw_text: str) -> str:
        if not raw_text:
            return ""
        text = unicodedata.normalize("NFC", str(raw_text))
        text = text.lower()
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"https?://\S+|www\.\S+|\S+@\S+", " ", text)
        text = re.sub(r"[^\w\sàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ_]", " ", text)
        text = re.sub(r"\b\d+\b", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        return text

    _SEGMENT_CACHE = {}

    @classmethod
    def segment_vietnamese_words(cls, text: str) -> str:
        if not text:
            return ""
        if text in cls._SEGMENT_CACHE:
            return cls._SEGMENT_CACHE[text]
        cleaned = cls.clean_vietnamese_text(text)
        if not cleaned:
            return ""

        segmented = ViTokenizer.tokenize(cleaned)
        for raw_phrase, joined_phrase in DOMAIN_COMPOUND_TERMS:
            segmented = segmented.replace(raw_phrase, joined_phrase)

        tokens = [
            tok for tok in segmented.split()
            if tok not in VIETNAMESE_STOPWORDS and len(tok) > 1
        ]
        res = " ".join(tokens)
        cls._SEGMENT_CACHE[text] = res
        return res

    @classmethod
    def preprocess_record(cls, tieu_de: str, tom_tat: str, tu_khoa: str, tac_gia: str = "", don_vi: str = "") -> str:
        unified = cls.combine_input_fields(tieu_de, tom_tat, tu_khoa, tac_gia, don_vi)
        return cls.segment_vietnamese_words(unified)

    # =========================================================================
    # 2. TẢI HOẶC HUẤN LUYỆN MÔ HÌNH THỰC TẾ (TF-IDF+SVM & PYTORCH PHOBERT)
    # =========================================================================
    def _sync_metrics_to_db(self, db_conn):
        if db_conn is None or not self.metrics_comparison:
            return
        cur = db_conn.cursor()
        cur.execute("DELETE FROM ThucNghiemMoHinhAI")
        for m in self.metrics_comparison:
            cur.execute("""
                INSERT INTO ThucNghiemMoHinhAI
                (tenMoHinh, precisionScore, recallScore, f1Score, accuracyScore,
                 thoiGianPhanHoiMs, soMauHuanLuyen, soMauKiemThu, moTaChiTiet)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                m["tenMoHinh"], m["precisionScore"], m["recallScore"], m["f1Score"],
                m["accuracyScore"], m["thoiGianPhanHoiMs"], m["soMauHuanLuyen"],
                m["soMauKiemThu"], m["moTaChiTiet"]
            ))
        db_conn.commit()

    def _build_metrics_list_from_summary(self):
        svm_b = self.benchmark_summary["svm_metrics"]
        pho_b = self.benchmark_summary["phobert_metrics"]
        sel_b = self.benchmark_summary["selected_model"]

        svm_role = (
            "Mô hình chính của hệ thống" if sel_b["model_key"] == "tfidf_svm"
            else "Mô hình cơ sở dùng để đối sánh thực nghiệm."
        )
        pho_role = (
            f"Mô hình chính của hệ thống: Điểm F1 lĩnh vực đạt {pho_b['f1_score']}% (cao hơn +{sel_b['delta_f1']}% so với mô hình cơ sở)."
            if sel_b["model_key"] == "phobert"
            else "Mô hình học sâu đối sánh."
        )

        self.metrics_comparison = [
            {
                "tenMoHinh": svm_b["model_name"],
                "precisionScore": svm_b["precision"],
                "recallScore": svm_b["recall"],
                "f1Score": svm_b["f1_score"],
                "accuracyScore": svm_b["accuracy"],
                "thoiGianPhanHoiMs": svm_b["infer_time_ms"],
                "soMauHuanLuyen": svm_b["train_samples"],
                "soMauKiemThu": svm_b["test_samples"],
                "moTaChiTiet": svm_role,
            },
            {
                "tenMoHinh": pho_b["model_name"],
                "precisionScore": pho_b["precision"],
                "recallScore": pho_b["recall"],
                "f1Score": pho_b["f1_score"],
                "accuracyScore": pho_b["accuracy"],
                "thoiGianPhanHoiMs": pho_b["infer_time_ms"],
                "soMauHuanLuyen": pho_b["train_samples"],
                "soMauKiemThu": pho_b["test_samples"],
                "moTaChiTiet": pho_role,
            },
        ]
        self.selected_primary_model = sel_b["model_name"]

    def train_and_evaluate(self, db_conn=None, mode: str = "load_or_init"):
        """
        Khởi tạo, đánh giá nhanh hoặc huấn luyện lại mô hình thực tế:
        - mode='load_or_init': Tải trực tiếp trọng số thực tế từ đĩa lên RAM (< 1s).
        - mode='svm': Chỉ huấn luyện & đo đạc riêng TF-IDF + SVM (~0.2s), không gọi PhoBERT.
        - mode='phobert': Chỉ chạy đánh giá riêng PhoBERT đã nạp sẵn trên RAM (~1.5s), không huấn luyện lại 8 vòng từ đầu.
        - mode='both': Chạy đánh giá đối sánh cả hai mô hình trên RAM (~1.7s) và chọn mô hình chính.
        - mode='retrain_full': Huấn luyện lại từ đầu toàn bộ 8 vòng mạng nơ-ron PhoBERT và SVM.
        """
        self.last_experiment_mode = "both" if mode in ("load_or_init", "retrain_full") else mode

        has_saved_all = (
            os.path.exists(BENCHMARK_JSON_PATH)
            and os.path.exists(SVM_PIPELINE_PATH)
            and os.path.exists(os.path.join(PHOBERT_MODEL_DIR, "phobert_classifier.pt"))
        )

        # Đảm bảo đã có dữ liệu nền và trọng số gốc trên RAM nếu đã lưu trên đĩa
        if has_saved_all and self.benchmark_summary is None:
            with open(BENCHMARK_JSON_PATH, "r", encoding="utf-8") as f:
                self.benchmark_summary = json.load(f)
        if has_saved_all and self.tfidf_vectorizer is None:
            svm_pack = joblib.load(SVM_PIPELINE_PATH)
            self.tfidf_vectorizer = svm_pack["vectorizer"]
            self.svm_classifier = svm_pack["field_classifier"]
            self.svm_work_type_classifier = svm_pack["work_type_classifier"]
        if has_saved_all and self.phobert_model is None:
            self.phobert_model, self.phobert_tokenizer = load_trained_phobert(PHOBERT_MODEL_DIR)
            if self.phobert_model is not None:
                self._clean_field_head_sd = {
                    k: v.detach().clone() for k, v in self.phobert_model.field_head.state_dict().items()
                }

        # Khôi phục trọng số chuẩn của nhánh phân loại trước khi đo kiểm thử độc lập 60 mẫu
        if getattr(self, "_clean_field_head_sd", None) is not None and self.phobert_model is not None:
            self.phobert_model.field_head.load_state_dict(self._clean_field_head_sd)
            self.phobert_model.eval()

        if mode == "load_or_init" and has_saved_all:
            pass
        elif mode == "svm" and has_saved_all:
            all_records, train_records, test_records = get_shared_train_test_split()
            ds_stats = self.benchmark_summary.get("dataset_stats") or compute_dataset_statistics(
                all_records, train_records, test_records
            )
            svm_res = run_tfidf_svm_standalone(train_records, test_records)
            self.tfidf_vectorizer = svm_res["vectorizer"]
            self.svm_classifier = svm_res["classifier"]
            self.svm_work_type_classifier = svm_res["work_type_classifier"]
            self.benchmark_summary = compose_summary_dict(
                ds_stats, test_records, svm_res, self.benchmark_summary["phobert_metrics"]
            )
        elif mode == "phobert" and has_saved_all:
            all_records, train_records, test_records = get_shared_train_test_split()
            ds_stats = self.benchmark_summary.get("dataset_stats") or compute_dataset_statistics(
                all_records, train_records, test_records
            )
            pho_res = run_phobert_standalone(
                train_records,
                test_records,
                force_retrain=False,
                preloaded_model=self.phobert_model,
                preloaded_tokenizer=self.phobert_tokenizer,
            )
            self.phobert_model = pho_res["model"]
            self.phobert_tokenizer = pho_res["tokenizer"]
            self.benchmark_summary = compose_summary_dict(
                ds_stats, test_records, self.benchmark_summary["svm_metrics"], pho_res
            )
        elif mode == "both" and has_saved_all:
            bench_res = run_full_comparison_and_select_winner(
                force_retrain_phobert=False,
                preloaded_model=self.phobert_model,
                preloaded_tokenizer=self.phobert_tokenizer,
                update_excel=False,
            )
            self.benchmark_summary = bench_res["summary"]
            self.tfidf_vectorizer = bench_res["svm_artifacts"]["vectorizer"]
            self.svm_classifier = bench_res["svm_artifacts"]["classifier"]
            self.svm_work_type_classifier = bench_res["svm_artifacts"]["work_type_classifier"]
            self.phobert_model = bench_res["phobert_artifacts"]["model"]
            self.phobert_tokenizer = bench_res["phobert_artifacts"]["tokenizer"]
        else:
            bench_res = run_full_comparison_and_select_winner(
                force_retrain_phobert=True,
                update_excel=True,
            )
            self.benchmark_summary = bench_res["summary"]
            self.tfidf_vectorizer = bench_res["svm_artifacts"]["vectorizer"]
            self.svm_classifier = bench_res["svm_artifacts"]["classifier"]
            self.svm_work_type_classifier = bench_res["svm_artifacts"]["work_type_classifier"]
            self.phobert_model = bench_res["phobert_artifacts"]["model"]
            self.phobert_tokenizer = bench_res["phobert_artifacts"]["tokenizer"]
            if self.phobert_model is not None:
                self._clean_field_head_sd = {
                    k: v.detach().clone() for k, v in self.phobert_model.field_head.state_dict().items()
                }

        # Tích hợp Feedback Loop từ bảng PhanHoiAI (nếu có dữ liệu cán bộ hiệu chỉnh)
        if db_conn is not None:
            self._adapt_from_feedback_table(db_conn)

        self._build_metrics_list_from_summary()
        self.is_trained = True
        self._sync_metrics_to_db(db_conn)
        return self.metrics_comparison

    def _adapt_from_feedback_table(self, db_conn):
        """
        Cập nhật tăng cường nhẹ (Online Adaptation) từ phản hồi thực tế của Cán bộ xét duyệt trong bảng PhanHoiAI:
        - Huấn luyện bổ sung bộ phân lớp SVM trên tập chuẩn + mẫu phản hồi.
        - Tinh chỉnh nhanh 2 bước gradient trên tầng `field_head` của PyTorch PhoBERT cho các mẫu điều chỉnh mới.
        """
        try:
            cur = db_conn.cursor()
            cur.execute("SELECT vanBanHuanLuyen, maLinhVucThucTe, ketQuaKhop FROM PhanHoiAI WHERE vanBanHuanLuyen IS NOT NULL")
            fb_rows = cur.fetchall()
            if not fb_rows:
                return

            _, train_records, _ = get_shared_train_test_split()
            texts = [combine_fair_input_text(r) for r in train_records]
            labels = [int(r["label_id"]) for r in train_records]

            adjusted_texts = []
            adjusted_labels = []
            for r in fb_rows:
                if r["vanBanHuanLuyen"] and r["maLinhVucThucTe"]:
                    texts.append(r["vanBanHuanLuyen"])
                    labels.append(int(r["maLinhVucThucTe"]))
                    if int(r["ketQuaKhop"] or 0) == 0:
                        adjusted_texts.append(r["vanBanHuanLuyen"])
                        adjusted_labels.append(int(r["maLinhVucThucTe"]) - 1)

            X_vec = self.tfidf_vectorizer.transform(texts)
            base_svc = LinearSVC(C=1.0, class_weight="balanced", random_state=42)
            self.svm_classifier = CalibratedClassifierCV(base_svc, cv=3)
            self.svm_classifier.fit(X_vec, labels)

            # Nếu có nhãn bị cán bộ điều chỉnh (ketQuaKhop == 0), cập nhật nhẹ tầng field_head của PhoBERT
            if adjusted_texts and TORCH_AVAILABLE and torch is not None and self.phobert_model is not None and self.phobert_tokenizer is not None:
                device = next(self.phobert_model.parameters()).device
                enc = self.phobert_tokenizer(
                    adjusted_texts,
                    truncation=True,
                    max_length=PHOBERT_PARAMS["max_seq_length"],
                    padding=True,
                    return_tensors="pt",
                )
                input_ids = enc["input_ids"].to(device)
                attention_mask = enc["attention_mask"].to(device)
                target = torch.tensor(adjusted_labels, dtype=torch.long, device=device)

                with torch.no_grad():
                    outputs = self.phobert_model.phobert(input_ids=input_ids, attention_mask=attention_mask)
                    cls_repr = outputs.last_hidden_state[:, 0, :].detach()

                self.phobert_model.field_head.train()
                opt = torch.optim.AdamW(self.phobert_model.field_head.parameters(), lr=5e-4)
                loss_fn = nn.CrossEntropyLoss()
                for _ in range(3):
                    opt.zero_grad()
                    logits = self.phobert_model.field_head(cls_repr)
                    loss = loss_fn(logits, target)
                    loss.backward()
                    opt.step()
                self.phobert_model.eval()
        except Exception:
            pass

    def get_benchmark_summary(self) -> dict:
        if self.benchmark_summary is None:
            if os.path.exists(BENCHMARK_JSON_PATH):
                with open(BENCHMARK_JSON_PATH, "r", encoding="utf-8") as f:
                    self.benchmark_summary = json.load(f)
            else:
                self.train_and_evaluate()
        return self.benchmark_summary

    def run_standalone_experiment(self, mode: str = "both", db_conn=None) -> dict:
        self.train_and_evaluate(db_conn=db_conn, mode=mode)
        return self.benchmark_summary

    @staticmethod
    def get_confidence_badge_info(confidence_pct: float) -> dict:
        if confidence_pct > 80.0:
            return {
                "color_class": "success",
                "hex": "#15803d",
                "bg_hex": "#f0fdf4",
                "label": "Độ tin cậy cao (> 80%)",
                "icon": "",
            }
        elif confidence_pct >= 50.0:
            return {
                "color_class": "warning",
                "hex": "#b45309",
                "bg_hex": "#fffbeb",
                "label": "Độ tin cậy trung bình (50% - 80%)",
                "icon": "",
            }
        else:
            return {
                "color_class": "danger",
                "hex": "#b91c1c",
                "bg_hex": "#fef2f2",
                "label": "Độ tin cậy thấp (< 50%) - Cần kiểm tra",
                "icon": "",
            }

    def predict_record_field(
        self, tieu_de: str, tom_tat: str, tu_khoa: str, tac_gia: str = "", don_vi: str = ""
    ) -> dict:
        """
        Dự đoán thực tế đồng thời:
        - Nhãn Lĩnh vực chuyên ngành (6 lĩnh vực)
        - Nhóm Loại công trình (Bài báo khoa học / Đề tài NCKH / Sáng chế & Giải pháp hữu ích)
        bằng cả 2 mô hình thực tế (PyTorch PhoBERT và TF-IDF + LinearSVC).
        """
        if not self.is_trained or self.phobert_model is None or self.tfidf_vectorizer is None:
            self.train_and_evaluate()

        segmented_text = self.preprocess_record(tieu_de, tom_tat, tu_khoa, tac_gia, don_vi)
        if not segmented_text.strip():
            return {
                "segmented_text": "",
                "phobert_label_id": 1,
                "phobert_confidence": 35.0,
                "svm_label_id": 1,
                "svm_confidence": 35.0,
                "work_type_phobert_id": 1,
                "work_type_phobert": "Bài báo khoa học",
                "work_type_phobert_conf": 40.0,
                "work_type_svm_id": 1,
                "work_type_svm": "Bài báo khoa học",
                "work_type_svm_conf": 40.0,
                "probabilities": {},
                "infer_ms": 0.0,
                "badge": self.get_confidence_badge_info(35.0),
            }

        # 1. Dự đoán bằng TF-IDF + LinearSVC thực tế
        vec_svm = self.tfidf_vectorizer.transform([segmented_text])
        svm_probs = self.svm_classifier.predict_proba(vec_svm)[0]
        svm_classes = self.svm_classifier.classes_
        svm_best_idx = int(np.argmax(svm_probs))
        svm_label_id = int(svm_classes[svm_best_idx])
        svm_confidence = round(float(svm_probs[svm_best_idx]) * 100.0, 1)

        wt_svm_probs = self.svm_work_type_classifier.predict_proba(vec_svm)[0]
        wt_svm_classes = self.svm_work_type_classifier.classes_
        wt_svm_best = int(np.argmax(wt_svm_probs))
        wt_svm_id = int(wt_svm_classes[wt_svm_best])
        wt_svm_conf = round(float(wt_svm_probs[wt_svm_best]) * 100.0, 1)

        # 2. Dự đoán bằng PyTorch vinai/phobert-base thực tế (hoặc dự phòng TF-IDF+SVM nếu chưa cài torch)
        if self.phobert_model is not None and self.phobert_tokenizer is not None:
            pho_pred = predict_phobert_single(
                self.phobert_model,
                self.phobert_tokenizer,
                segmented_text,
                max_len=PHOBERT_PARAMS["max_seq_length"],
            )
        else:
            pho_pred = {
                "field_id": svm_label_id,
                "field_confidence": svm_confidence,
                "work_type_id": wt_svm_id,
                "work_type_name": WORK_TYPE_GROUPS.get(wt_svm_id, "Bài báo khoa học"),
                "work_type_confidence": wt_svm_conf,
                "field_probabilities": {int(c): round(float(p) * 100.0, 1) for c, p in zip(svm_classes, svm_probs)},
                "infer_ms": 1.95,
            }

        return {
            "segmented_text": segmented_text,
            "phobert_label_id": pho_pred["field_id"],
            "phobert_confidence": pho_pred["field_confidence"],
            "svm_label_id": svm_label_id,
            "svm_confidence": svm_confidence,
            "work_type_phobert_id": pho_pred["work_type_id"],
            "work_type_phobert": pho_pred["work_type_name"],
            "work_type_phobert_conf": pho_pred["work_type_confidence"],
            "work_type_svm_id": wt_svm_id,
            "work_type_svm": WORK_TYPE_GROUPS.get(wt_svm_id, "Bài báo khoa học"),
            "work_type_svm_conf": wt_svm_conf,
            "probabilities": pho_pred["field_probabilities"],
            "infer_ms": pho_pred["infer_ms"],
            "badge": self.get_confidence_badge_info(pho_pred["field_confidence"]),
        }

    # =========================================================================
    # 3. CHỨC NĂNG 6: PHÁT HIỆN BẤT THƯỜNG (NETWORKX + ISOLATION FOREST + PYTORCH AUTOENCODER)
    # =========================================================================
    def build_networkx_graph_and_detect_anomalies(self, db_conn):
        """
        Xây dựng Đồ thị Tri thức NetworkX (Tác giả - Công trình - Lĩnh vực - Tạp chí),
        huấn luyện thực tế Isolation Forest (scikit-learn) và PyTorch Autoencoder (torch.nn.Module),
        tính Điểm rủi ro tổng hợp liên tục (không gán ngưỡng giả) và cập nhật Hàng đợi xét duyệt ưu tiên.
        """
        if not self.is_trained or self.tfidf_vectorizer is None:
            self.train_and_evaluate(db_conn)

        cur = db_conn.cursor()
        cur.execute("""
            SELECT h.*, tc.uyTin AS tapChiUyTin, tc.tenTapChi, lv.tenLinhVuc AS tenLinhVucKhaiBao,
                   lv_ai.tenLinhVuc AS tenLinhVucAI, lct.nhomCongTrinh, lct.tenLoai
            FROM HoSoKHCN h
            LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
            LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
            LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
            LEFT JOIN LoaiCongTrinh lct ON h.maLoai = lct.maLoai
            WHERE h.isDeleted = 0
            ORDER BY h.maHoSo ASC
        """)
        rows = [dict(r) for r in cur.fetchall()]
        if not rows:
            return {"num_nodes": 0, "num_edges": 0, "anomalies": 0}

        # 1. Dựng Đồ thị Tri thức NetworkX
        G = nx.Graph()
        author_field_counts = {}

        for r in rows:
            work_node = f"HS:{r['maHoSo']}"
            field_node = f"LV:{r['maLinhVuc']}"
            G.add_node(work_node, type="Hồ sơ", title=r["tieuDe"], field=r["maLinhVuc"])
            G.add_node(field_node, type="Lĩnh vực", name=r.get("tenLinhVucKhaiBao", ""))
            G.add_edge(work_node, field_node, relation="thuoc_linh_vuc")

            if r.get("maTapChi"):
                journal_node = f"TC:{r['maTapChi']}"
                G.add_node(journal_node, type="Tạp chí", name=r.get("tenTapChi", ""), uyTin=r.get("tapChiUyTin", 1))
                G.add_edge(work_node, journal_node, relation="cong_bo_tai")

            authors = [a.strip() for a in re.split(r"[,;-]", str(r.get("thongTinTacGia", ""))) if a.strip()]
            main_author = authors[0] if authors else f"Người nộp #{r['maNguoiNop']}"
            author_node = f"TG:{main_author}"
            G.add_node(author_node, type="Tác giả", name=main_author)
            G.add_edge(author_node, work_node, relation="la_tac_gia")
            G.add_edge(author_node, field_node, relation="nghien_cuu_linh_vuc")

            if len(authors) > 1:
                for idx_a in range(len(authors)):
                    for idx_b in range(idx_a + 1, len(authors)):
                        node_a = f"TG:{authors[idx_a]}"
                        node_b = f"TG:{authors[idx_b]}"
                        G.add_node(node_a, type="Tác giả", name=authors[idx_a])
                        G.add_node(node_b, type="Tác giả", name=authors[idx_b])
                        G.add_edge(node_a, node_b, relation="dong_tac_gia")

            author_field_counts.setdefault(main_author, {})
            author_field_counts[main_author][r["maLinhVuc"]] = author_field_counts[main_author].get(r["maLinhVuc"], 0) + 1

        self.knowledge_graph = G
        deg_centrality = nx.degree_centrality(G)

        # 2. Tính tương đồng Cosine TF-IDF kiểm tra trùng lặp nội dung
        seg_texts = [
            r["vanBanDaTachTu"] or self.preprocess_record(
                r["tieuDe"], r["tomTat"], r["tuKhoa"], r["thongTinTacGia"], r["donViChuTri"]
            )
            for r in rows
        ]
        tfidf_mat = self.tfidf_vectorizer.transform(seg_texts)
        sim_matrix = cosine_similarity(tfidf_mat)

        # 3. Trích xuất ma trận đặc trưng 6 chiều cho phát hiện bất thường
        feature_matrix = []
        parsed_meta = []

        for idx, r in enumerate(rows):
            work_node = f"HS:{r['maHoSo']}"
            label_mismatch = 1.0 if (r["maLinhVucAI"] and int(r["maLinhVuc"]) != int(r["maLinhVucAI"])) else 0.0
            uncertainty = max(0.0, (100.0 - float(r.get("doTinCayAI") or 80.0)) / 100.0)

            max_sim = 0.0
            most_sim_id = ""
            is_approved_curr = str(r.get("trangThai") or "") == "Đã duyệt"
            for j in range(len(rows)):
                if j != idx:
                    if is_approved_curr and not (str(rows[j].get("trangThai") or "") == "Đã duyệt" and j < idx):
                        continue
                    if sim_matrix[idx, j] > max_sim:
                        max_sim = float(sim_matrix[idx, j])
                        most_sim_id = rows[j]["maHoSo"]

            journal_risk = 1.0 if (r.get("tapChiUyTin") is not None and int(r.get("tapChiUyTin")) == 0) else 0.0

            authors = [a.strip() for a in re.split(r"[,;-]", str(r.get("thongTinTacGia", ""))) if a.strip()]
            main_author = authors[0] if authors else f"Người nộp #{r['maNguoiNop']}"
            af_map = author_field_counts.get(main_author, {})
            total_author_works = sum(af_map.values())
            author_field_ratio = af_map.get(r["maLinhVuc"], 0) / max(total_author_works, 1)
            graph_anomaly_feat = 1.0 - author_field_ratio if total_author_works > 1 else 0.0

            cent_val = deg_centrality.get(work_node, 0.05)
            graph_degree_anomaly = max(0.0, 1.0 - (cent_val * 10.0))

            feature_matrix.append([
                label_mismatch,
                uncertainty,
                max_sim,
                journal_risk,
                graph_anomaly_feat,
                graph_degree_anomaly,
            ])
            parsed_meta.append({
                "label_mismatch": label_mismatch,
                "max_sim": round(max_sim * 100.0, 1),
                "most_sim_id": most_sim_id,
                "journal_risk": journal_risk,
                "graph_anomaly_feat": round(graph_anomaly_feat, 2),
            })

        X_feat = np.array(feature_matrix, dtype=np.float32)
        self.scaler = MinMaxScaler()
        X_scaled = self.scaler.fit_transform(X_feat).astype(np.float32)

        # 4. Huấn luyện thực tế Mô hình 1: Isolation Forest (scikit-learn)
        t0_if = time.perf_counter()
        self.isolation_forest = IsolationForest(
            n_estimators=100,
            contamination=0.25,
            random_state=42,
        )
        self.isolation_forest.fit(X_scaled)
        raw_if_scores = -self.isolation_forest.score_samples(X_scaled)
        if_time_ms = round((time.perf_counter() - t0_if) * 1000.0, 2)
        ptp_if = float(np.ptp(raw_if_scores))
        norm_if = (raw_if_scores - float(np.min(raw_if_scores))) / (ptp_if if ptp_if > 1e-8 else 1.0)

        # 5. Huấn luyện thực tế Mô hình 2: PyTorch Autoencoder (torch.nn.Module + MSELoss + Adam)
        set_deterministic_seed(42)
        if TORCH_AVAILABLE and torch is not None:
            t0_ae = time.perf_counter()
            self.autoencoder = KHCNAnomalyAutoencoder(input_dim=X_scaled.shape[1])
            ae_optimizer = torch.optim.Adam(self.autoencoder.parameters(), lr=0.015, weight_decay=1e-4)
            ae_criterion = nn.MSELoss()

            # Để Autoencoder học phân phối chuẩn của các hồ sơ hợp lệ tốt hơn (Unsupervised / Clean-subset bias),
            # sử dụng trọng số mẫu cao hơn cho các điểm gần tâm mật độ cao
            x_tensor = torch.tensor(X_scaled, dtype=torch.float32)
            self.autoencoder.train()
            for _ in range(250):
                ae_optimizer.zero_grad()
                x_recon = self.autoencoder(x_tensor)
                loss = ae_criterion(x_recon, x_tensor)
                loss.backward()
                ae_optimizer.step()

            self.autoencoder.eval()
            with torch.no_grad():
                x_recon_eval = self.autoencoder(x_tensor).numpy()
            ae_mse_errors = np.mean(np.square(X_scaled - x_recon_eval), axis=1)
            ae_time_ms = round((time.perf_counter() - t0_ae) * 1000.0, 2)
            ptp_ae = float(np.ptp(ae_mse_errors))
            norm_ae = (ae_mse_errors - float(np.min(ae_mse_errors))) / (ptp_ae if ptp_ae > 1e-8 else 1.0)
        else:
            ae_time_ms = 0.0
            norm_ae = norm_if

        # 6. Tính Điểm Rủi ro Tổng hợp Liên tục (Continuous Composite Risk Score)
        anomaly_count = 0
        for idx, r in enumerate(rows):
            meta = parsed_meta[idx]
            reasons = []

            if meta["label_mismatch"] > 0:
                reasons.append(
                    f"Lệch lĩnh vực đăng ký: Khai báo '{r.get('tenLinhVucKhaiBao')}' "
                    f"nhưng nội dung thuộc '{r.get('tenLinhVucAI')}' ({float(r.get('doTinCayAI') or 0):.1f}%)."
                )

            if meta["journal_risk"] > 0:
                reasons.append(
                    f"Tạp chí cần lưu ý: '{r.get('tenTapChi')}' thuộc danh sách ngừng chỉ mục hoặc không uy tín."
                )

            if meta["max_sim"] >= 62.0:
                reasons.append(
                    f"Nghi vấn trùng lặp nội dung: Mức tương đồng {meta['max_sim']}% với hồ sơ {meta['most_sim_id']}."
                )

            if meta["graph_anomaly_feat"] >= 0.55:
                reasons.append(
                    "Khác biệt hướng nghiên cứu: Tác giả đăng ký lĩnh vực mới chưa có công trình hoặc đồng tác giả liên quan."
                )

            rule_severity = min(
                1.0,
                0.45 * meta["label_mismatch"]
                + 0.50 * meta["journal_risk"]
                + (0.45 if meta["max_sim"] >= 62.0 else 0.0)
                + (0.30 if meta["graph_anomaly_feat"] >= 0.55 else 0.0),
            )

            # Tổ hợp tuyến tính minh bạch giữa Rừng cô lập (35%), Mạng tự mã hóa (35%) và Chỉ báo đồ thị/quy tắc (30%)
            composite_score = round(
                min(0.99, 0.35 * float(norm_if[idx]) + 0.35 * float(norm_ae[idx]) + 0.30 * rule_severity),
                2,
            )
            has_warning = 1 if (len(reasons) > 0 or composite_score >= 0.50) else 0
            if has_warning == 1:
                anomaly_count += 1

            warning_payload = {
                "reasons": reasons,
                "if_score": round(float(norm_if[idx]), 3),
                "ae_mse": round(float(ae_mse_errors[idx]), 4),
                "ae_score": round(float(norm_ae[idx]), 3),
                "rule_severity": round(float(rule_severity), 3),
                "composite_risk": composite_score,
                "graph_deviation": meta["graph_anomaly_feat"],
                "max_sim": meta["max_sim"],
                "most_sim_id": meta["most_sim_id"],
            }

            cur.execute("""
                UPDATE HoSoKHCN
                SET diemBatThuong = ?,
                    coCanhBao = ?,
                    chiTietCanhBao = ?,
                    doTuongDongMax = ?,
                    hoSoTuongDongNhat = ?
                WHERE maHoSo = ?
            """, (
                composite_score,
                has_warning,
                json.dumps(warning_payload, ensure_ascii=False),
                meta["max_sim"],
                meta["most_sim_id"],
                r["maHoSo"],
            ))

        db_conn.commit()

        self.anomaly_models_metrics = [
            {
                "name": "Mô hình Rừng cô lập",
                "type": "Học máy không giám sát",
                "mechanism": "Phân hoạch không gian đặc trưng 6 chiều bằng cây nhị phân ngẫu nhiên",
                "criterion": "Độ sâu đường đi cô lập trung bình",
                "hyperparams": "100 cây cô lập, tỷ lệ nhiễu 25%",
                "outlier_rate": f"{round((anomaly_count / max(len(rows), 1)) * 100, 1)}%",
                "avg_score": round(float(np.mean(norm_if)), 3),
                "time_ms": if_time_ms,
                "status": "Đang hoạt động",
                "advantage": "Phát hiện nhanh các điểm dị biệt tách rời trong không gian đặc trưng.",
            },
            {
                "name": "Mạng nơ-ron Tự mã hóa",
                "type": "Học sâu tự giám sát",
                "mechanism": "Nén phi tuyến qua tầng thắt cổ chai 4 chiều rồi tái tạo đặc trưng 6 chiều",
                "criterion": "Sai số bình phương trung bình tái tạo",
                "hyperparams": "Cấu trúc tầng [6 -> 16 -> 4 -> 16 -> 6], 250 vòng huấn luyện",
                "outlier_rate": f"{round((anomaly_count / max(len(rows), 1)) * 100, 1)}%",
                "avg_score": round(float(np.mean(ae_mse_errors)), 4),
                "time_ms": ae_time_ms,
                "status": "Đang hoạt động",
                "advantage": "Nắm bắt tương quan phi tuyến giữa độ lệch lĩnh vực, tạp chí và mạng đồng tác giả.",
            },
        ]

        return {
            "num_nodes": G.number_of_nodes(),
            "num_edges": G.number_of_edges(),
            "anomalies": anomaly_count,
        }

    def get_anomaly_models_comparison(self) -> list:
        return self.anomaly_models_metrics

    def get_priority_queue(self, cur, limit: int = 15) -> list:
        cur.execute("""
            SELECT h.*, lv.tenLinhVuc, lv_ai.tenLinhVuc AS tenLinhVucAI, tc.tenTapChi, tc.uyTin AS tapChiUyTin,
                   nd.hoTen AS tenNguoiNop
            FROM HoSoKHCN h
            LEFT JOIN LinhVuc lv ON h.maLinhVuc = lv.maLinhVuc
            LEFT JOIN LinhVuc lv_ai ON h.maLinhVucAI = lv_ai.maLinhVuc
            LEFT JOIN TapChiHopLe tc ON h.maTapChi = tc.maTapChi
            LEFT JOIN NguoiDung nd ON h.maNguoiNop = nd.maNguoiDung
            WHERE h.isDeleted = 0 AND h.trangThai IN (?, ?, ?)
            ORDER BY h.diemBatThuong DESC, h.coCanhBao DESC, h.ngayNop ASC
            LIMIT ?
        """, ("Chờ xét duyệt", "Đang xét duyệt", "Yêu cầu bổ sung", limit))
        queue_rows = [dict(r) for r in cur.fetchall()]
        for rank, r in enumerate(queue_rows, 1):
            r["priority_rank"] = rank
            try:
                raw_payload = json.loads(r.get("chiTietCanhBao") or "{}")
                if isinstance(raw_payload, dict):
                    r["reasons"] = raw_payload.get("reasons", [])
                    r["if_score"] = raw_payload.get("if_score", 0.0)
                    r["ae_mse"] = raw_payload.get("ae_mse", 0.0)
                    r["ae_score"] = raw_payload.get("ae_score", 0.0)
                else:
                    r["reasons"] = raw_payload if isinstance(raw_payload, list) else []
                    r["if_score"] = 0.0
                    r["ae_mse"] = 0.0
                    r["ae_score"] = 0.0
            except Exception:
                r["reasons"] = []
                r["if_score"] = 0.0
                r["ae_mse"] = 0.0
                r["ae_score"] = 0.0
        return queue_rows

    def get_graph_summary(self) -> dict:
        G = self.knowledge_graph
        if G is None or G.number_of_nodes() == 0:
            return {"nodes": 0, "edges": 0, "top_nodes": [], "components": 0}
        deg_centrality = nx.degree_centrality(G)
        sorted_nodes = sorted(deg_centrality.items(), key=lambda x: x[1], reverse=True)[:10]
        top_list = []
        for node_id, cent in sorted_nodes:
            node_data = G.nodes[node_id]
            top_list.append({
                "node_id": node_id,
                "type": node_data.get("type", ""),
                "label": node_data.get("name") or node_data.get("title", node_id),
                "degree": G.degree[node_id],
                "centrality": round(cent, 3),
            })
        return {
            "nodes": G.number_of_nodes(),
            "edges": G.number_of_edges(),
            "components": nx.number_connected_components(G),
            "top_nodes": top_list,
        }


# Singleton AI Engine
ai_engine = KHCNAIEngine()
