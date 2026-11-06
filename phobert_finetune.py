# -*- coding: utf-8 -*-
"""
Mô-đun Huấn luyện & Suy luận Thực tế Fine-tuning PhoBERT bằng PyTorch & HuggingFace Transformers
Đề tài: Hỗ trợ quản lý hồ sơ khoa học công nghệ bằng phương pháp phân loại tự động và phát hiện bất thường (CNTT-KLCN001)

Kiến trúc thực tế 100% trên PyTorch:
- Pre-trained backbone: vinai/phobert-base (RoBERTa-base tiếng Việt, 12 Transformer Layers, Hidden=768)
- Đóng băng (freeze) tầng Embeddings và 10 tầng Transformer bên dưới (encoder.layer.0 .. encoder.layer.9)
- Tinh chỉnh (fine-tune) 2 tầng Transformer ngữ cảnh bậc cao (encoder.layer.10 .. encoder.layer.11)
- Đa đầu ra phân loại (Multi-task Classification Heads):
  1) Field Head (6 nhãn Lĩnh vực chuyên ngành): Dropout(0.2) -> Linear(768, 256) -> GELU -> Dropout(0.1) -> Linear(256, 6)
  2) Work-Type Head (3 nhóm Loại công trình: Bài báo khoa học, Đề tài NCKH, Sáng chế & Giải pháp hữu ích):
     Dropout(0.2) -> Linear(768, 128) -> GELU -> Linear(128, 3)
"""

import os
import json
import time
import random
import numpy as np

# Tắt giao thức Xet trên Windows để tải trọng số HuggingFace qua HTTPS ổn định
os.environ.setdefault("HF_HUB_DISABLE_XET", "1")
os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

try:
    import torch
    import torch.nn as nn
    from torch.optim import AdamW
    from torch.utils.data import Dataset, DataLoader
    from transformers import AutoModel, AutoTokenizer
    TORCH_AVAILABLE = True
    _torch_ver = getattr(torch, "__version__", "2.x")
    _BaseModule = nn.Module
    _BaseDataset = Dataset

    def _torch_inference_mode():
        return torch.inference_mode()
except ImportError:
    torch = None
    nn = None
    AdamW = None
    DataLoader = None
    AutoModel = None
    AutoTokenizer = None
    TORCH_AVAILABLE = False
    _torch_ver = "Chưa cài đặt (Sử dụng dự phòng TF-IDF + SVM)"
    _BaseModule = object
    _BaseDataset = object

    def _torch_inference_mode():
        def decorator(f):
            return f
        return decorator

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SAVE_DIR = os.path.join(BASE_DIR, "models", "phobert_khcn")
CHECKPOINT_FILENAME = "phobert_classifier.pt"

PHOBERT_MODEL_NAME = "vinai/phobert-base"

WORK_TYPE_GROUPS = {
    1: "Bài báo khoa học",
    2: "Đề tài NCKH",
    3: "Sáng chế & Giải pháp hữu ích",
}
WORK_TYPE_NAME_TO_ID = {v: k for k, v in WORK_TYPE_GROUPS.items()}

HYPERPARAMETERS = {
    "pretrained_model": PHOBERT_MODEL_NAME,
    "framework": f"PyTorch {_torch_ver} + Transformers",
    "architecture": "RoBERTa-base Vietnamese (12 Transformer Layers, Hidden=768, Heads=12)",
    "frozen_lower_layers": "Embeddings + Layers 0..9 frozen (encoder.layer.0 to encoder.layer.9)",
    "trainable_upper_layers": "encoder.layer.10, encoder.layer.11 + Multi-task Classification Heads",
    "classification_head": (
        "Field Head: Dropout(0.2)->Linear(768,256)->GELU->Linear(256,6) | "
        "WorkType Head: Dropout(0.2)->Linear(768,128)->GELU->Linear(128,3)"
    ),
    "learning_rate_encoder": 5e-5,
    "learning_rate_head": 6e-4,
    "batch_size": 16,
    "optimizer": "AdamW (weight_decay=0.01)",
    "max_seq_length": 128,
    "epochs": 8,
}

_CACHED_BACKBONE = None
_CACHED_TOKENIZER = None


def set_deterministic_seed(seed: int = 42):
    random.seed(seed)
    np.random.seed(seed)
    if TORCH_AVAILABLE and torch is not None:
        torch.manual_seed(seed)
        if torch.cuda.is_available():
            torch.cuda.manual_seed_all(seed)


def get_phobert_tokenizer():
    global _CACHED_TOKENIZER
    if not TORCH_AVAILABLE or AutoTokenizer is None:
        return None
    if _CACHED_TOKENIZER is None:
        _CACHED_TOKENIZER = AutoTokenizer.from_pretrained(PHOBERT_MODEL_NAME)
    return _CACHED_TOKENIZER


class KHCNPhoBERTDataset(_BaseDataset):
    """Dataset chuẩn hóa văn bản tiếng Việt đã tách từ PyVi cho PhoBERT."""

    def __init__(self, segmented_texts, field_labels, work_type_labels, tokenizer, max_len: int = 128):
        self.texts = list(segmented_texts)
        self.field_labels = [int(y) - 1 for y in field_labels]  # Chuyển 1..6 -> 0..5
        self.work_type_labels = [int(w) - 1 for w in work_type_labels]  # Chuyển 1..3 -> 0..2
        self.tokenizer = tokenizer
        self.max_len = max_len

    def __len__(self):
        return len(self.texts)

    def __getitem__(self, idx):
        if not TORCH_AVAILABLE or torch is None:
            return {}
        text = str(self.texts[idx])
        enc = self.tokenizer(
            text,
            truncation=True,
            max_length=self.max_len,
            padding="max_length",
            return_tensors="pt",
        )
        return {
            "input_ids": enc["input_ids"].squeeze(0),
            "attention_mask": enc["attention_mask"].squeeze(0),
            "field_label": torch.tensor(self.field_labels[idx], dtype=torch.long),
            "work_type_label": torch.tensor(self.work_type_labels[idx], dtype=torch.long),
        }


class PhoBERTForKHCNClassification(_BaseModule):
    """
    Mô hình PyTorch Fine-tuning PhoBERT đa nhiệm:
    - Phân loại đồng thời Lĩnh vực nghiên cứu (6 lớp) và Loại công trình KH&CN (3 nhóm).
    """

    def __init__(self, num_fields: int = 6, num_work_types: int = 3, freeze_lower_layers: int = 10):
        super().__init__()
        self.num_fields = num_fields
        self.num_work_types = num_work_types
        self.freeze_lower_layers = freeze_lower_layers

        self.phobert = AutoModel.from_pretrained(PHOBERT_MODEL_NAME)

        # Đóng băng tầng nhúng từ vựng và các tầng Transformer bên dưới (0 .. freeze_lower_layers-1)
        for param in self.phobert.embeddings.parameters():
            param.requires_grad = False
        for layer_idx in range(freeze_lower_layers):
            for param in self.phobert.encoder.layer[layer_idx].parameters():
                param.requires_grad = False

        # Nếu có pooler không dùng tới thì đóng băng để tối ưu bộ nhớ gradient
        if getattr(self.phobert, "pooler", None) is not None:
            for param in self.phobert.pooler.parameters():
                param.requires_grad = False

        hidden_size = self.phobert.config.hidden_size  # 768

        # Nhánh 1: Phân loại Lĩnh vực chuyên ngành (6 lĩnh vực)
        self.field_head = nn.Sequential(
            nn.Dropout(p=0.2),
            nn.Linear(hidden_size, 256),
            nn.GELU(),
            nn.Dropout(p=0.1),
            nn.Linear(256, num_fields),
        )

        # Nhánh 2: Phân loại Nhóm Loại công trình (Bài báo KH / Đề tài NCKH / Sáng chế & GPHI)
        self.work_type_head = nn.Sequential(
            nn.Dropout(p=0.2),
            nn.Linear(hidden_size, 128),
            nn.GELU(),
            nn.Linear(128, num_work_types),
        )

    def train(self, mode: bool = True):
        super().train(mode)
        # Giữ các tầng đã đóng băng ở chế độ eval() để tắt Dropout ngẫu nhiên ở các tầng cố định
        if mode:
            self.phobert.embeddings.eval()
            for layer_idx in range(self.freeze_lower_layers):
                self.phobert.encoder.layer[layer_idx].eval()
        return self

    def forward(self, input_ids, attention_mask):
        outputs = self.phobert(input_ids=input_ids, attention_mask=attention_mask)
        cls_repr = outputs.last_hidden_state[:, 0, :]  # Vector ngữ cảnh đại diện tại token <s>
        field_logits = self.field_head(cls_repr)
        work_type_logits = self.work_type_head(cls_repr)
        return field_logits, work_type_logits

    def get_trainable_state_dict(self) -> dict:
        """Chỉ trích xuất các trọng số có huấn luyện để lưu checkpoint gọn nhẹ (~57 MB thay vì 540 MB)."""
        trainable_names = {name for name, p in self.named_parameters() if p.requires_grad}
        full_sd = self.state_dict()
        return {k: v.cpu() for k, v in full_sd.items() if k in trainable_names}

    def load_trainable_state_dict(self, partial_state_dict: dict):
        own_state = self.state_dict()
        for k, v in partial_state_dict.items():
            if k in own_state:
                own_state[k].copy_(v)


def train_phobert_multitask(
    train_texts,
    train_field_labels,
    train_work_type_labels,
    epochs: int = None,
    batch_size: int = None,
    max_len: int = None,
    save_dir: str = DEFAULT_SAVE_DIR,
    seed: int = 42,
):
    """
    Huấn luyện thực tế mô hình PhoBERT trên PyTorch và lưu trọng số vào đĩa cứng.
    Trả về (model, tokenizer, train_time_ms, epoch_losses).
    """
    set_deterministic_seed(seed)
    epochs = epochs or HYPERPARAMETERS["epochs"]
    batch_size = batch_size or HYPERPARAMETERS["batch_size"]
    max_len = max_len or HYPERPARAMETERS["max_seq_length"]

    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    tokenizer = get_phobert_tokenizer()
    dataset = KHCNPhoBERTDataset(
        train_texts, train_field_labels, train_work_type_labels, tokenizer, max_len=max_len
    )
    generator = torch.Generator()
    generator.manual_seed(seed)
    dataloader = DataLoader(dataset, batch_size=batch_size, shuffle=True, generator=generator)

    model = PhoBERTForKHCNClassification(
        num_fields=6, num_work_types=3, freeze_lower_layers=10
    ).to(device)

    encoder_params = [
        p for n, p in model.phobert.named_parameters() if p.requires_grad
    ]
    head_params = list(model.field_head.parameters()) + list(model.work_type_head.parameters())

    optimizer = AdamW(
        [
            {"params": encoder_params, "lr": HYPERPARAMETERS["learning_rate_encoder"]},
            {"params": head_params, "lr": HYPERPARAMETERS["learning_rate_head"]},
        ],
        weight_decay=0.01,
    )
    criterion_field = nn.CrossEntropyLoss()
    criterion_type = nn.CrossEntropyLoss()

    epoch_losses = []
    t0 = time.perf_counter()
    model.train()
    for epoch in range(epochs):
        running_loss = 0.0
        for batch in dataloader:
            optimizer.zero_grad()
            input_ids = batch["input_ids"].to(device)
            attention_mask = batch["attention_mask"].to(device)
            y_field = batch["field_label"].to(device)
            y_type = batch["work_type_label"].to(device)

            field_logits, type_logits = model(input_ids=input_ids, attention_mask=attention_mask)
            loss_f = criterion_field(field_logits, y_field)
            loss_t = criterion_type(type_logits, y_type)
            loss = loss_f + 0.45 * loss_t

            loss.backward()
            torch.nn.utils.clip_grad_norm_(model.parameters(), max_norm=1.0)
            optimizer.step()
            running_loss += float(loss.item())

        avg_loss = round(running_loss / max(len(dataloader), 1), 4)
        epoch_losses.append(avg_loss)
        print(f"[PyTorch PhoBERT Fine-tune] Epoch {epoch + 1}/{epochs} - Loss: {avg_loss:.4f}")

    train_time_ms = round((time.perf_counter() - t0) * 1000.0, 2)
    model.eval()

    os.makedirs(save_dir, exist_ok=True)
    ckpt_path = os.path.join(save_dir, CHECKPOINT_FILENAME)
    torch.save(
        {
            "pretrained_model": PHOBERT_MODEL_NAME,
            "freeze_lower_layers": model.freeze_lower_layers,
            "num_fields": model.num_fields,
            "num_work_types": model.num_work_types,
            "trainable_state_dict": model.get_trainable_state_dict(),
            "epoch_losses": epoch_losses,
            "train_time_ms": train_time_ms,
        },
        ckpt_path,
    )
    with open(os.path.join(save_dir, "hyperparameters.json"), "w", encoding="utf-8") as f:
        json.dump(
            {**HYPERPARAMETERS, "epoch_losses": epoch_losses, "train_time_ms": train_time_ms},
            f,
            ensure_ascii=False,
            indent=2,
        )

    return model, tokenizer, train_time_ms, epoch_losses


def load_trained_phobert(save_dir: str = DEFAULT_SAVE_DIR):
    """Tải mô hình PhoBERT đã fine-tune từ đĩa cứng (nếu có); trả về (model, tokenizer) hoặc (None, None)."""
    if not TORCH_AVAILABLE or torch is None:
        return None, None
    ckpt_path = os.path.join(save_dir, CHECKPOINT_FILENAME)
    if not os.path.exists(ckpt_path):
        return None, None

    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    tokenizer = get_phobert_tokenizer()
    ckpt = torch.load(ckpt_path, map_location=device, weights_only=False)
    model = PhoBERTForKHCNClassification(
        num_fields=ckpt.get("num_fields", 6),
        num_work_types=ckpt.get("num_work_types", 3),
        freeze_lower_layers=ckpt.get("freeze_lower_layers", 10),
    ).to(device)
    model.load_trainable_state_dict(ckpt["trainable_state_dict"])
    model.eval()
    return model, tokenizer


@_torch_inference_mode()
def predict_phobert_batch(model, tokenizer, segmented_texts, max_len: int = 128, batch_size: int = 16):
    """
    Suy luận thực tế trên tập văn bản bằng mini-batch phân cụm theo độ dài câu
    (giảm đệm dư thừa trên CPU nhưng giữ nguyên kết quả toán học 100%) và đo thời gian thực bằng time.perf_counter().
    """
    if not TORCH_AVAILABLE or model is None or tokenizer is None:
        n = len(segmented_texts)
        return {
            "field_preds": np.ones(n, dtype=int),
            "field_probs": np.zeros((n, 6), dtype=float),
            "work_type_preds": np.ones(n, dtype=int),
            "work_type_probs": np.zeros((n, 3), dtype=float),
            "avg_infer_ms": 0.0,
            "total_infer_ms": 0.0,
        }
    device = next(model.parameters()).device
    model.eval()

    n = len(segmented_texts)
    if n == 0:
        return {
            "field_preds": np.array([], dtype=int),
            "field_probs": np.empty((0, 6), dtype=float),
            "work_type_preds": np.array([], dtype=int),
            "work_type_probs": np.empty((0, 3), dtype=float),
            "avg_infer_ms": 0.0,
            "total_infer_ms": 0.0,
        }

    pred_fields = [0] * n
    pred_field_probs = [None] * n
    pred_work_types = [0] * n
    pred_work_type_probs = [None] * n

    sorted_indices = sorted(range(n), key=lambda idx: len(str(segmented_texts[idx])))

    t0 = time.perf_counter()
    for start in range(0, n, batch_size):
        batch_idx = sorted_indices[start : start + batch_size]
        batch_texts = [str(segmented_texts[i]) for i in batch_idx]
        enc = tokenizer(
            batch_texts,
            truncation=True,
            max_length=max_len,
            padding=True,
            return_tensors="pt",
        )
        input_ids = enc["input_ids"].to(device)
        attention_mask = enc["attention_mask"].to(device)
        f_logits, t_logits = model(input_ids=input_ids, attention_mask=attention_mask)

        f_probs_batch = torch.softmax(f_logits, dim=-1).cpu().numpy()
        t_probs_batch = torch.softmax(t_logits, dim=-1).cpu().numpy()

        for local_i, orig_i in enumerate(batch_idx):
            fp = f_probs_batch[local_i]
            tp = t_probs_batch[local_i]
            pred_fields[orig_i] = int(np.argmax(fp)) + 1
            pred_field_probs[orig_i] = fp
            pred_work_types[orig_i] = int(np.argmax(tp)) + 1
            pred_work_type_probs[orig_i] = tp

    total_infer_ms = (time.perf_counter() - t0) * 1000.0
    avg_infer_ms = round(total_infer_ms / max(n, 1), 2)

    return {
        "field_preds": np.array(pred_fields, dtype=int),
        "field_probs": np.array(pred_field_probs, dtype=float),
        "work_type_preds": np.array(pred_work_types, dtype=int),
        "work_type_probs": np.array(pred_work_type_probs, dtype=float),
        "avg_infer_ms": avg_infer_ms,
        "total_infer_ms": round(total_infer_ms, 2),
    }


@_torch_inference_mode()
def predict_phobert_single(model, tokenizer, segmented_text: str, max_len: int = 128) -> dict:
    """Dự đoán thực tế cho 1 hồ sơ trên giao diện Web (cả Lĩnh vực và Loại công trình)."""
    res = predict_phobert_batch(model, tokenizer, [segmented_text], max_len=max_len)
    f_probs = res["field_probs"][0]
    t_probs = res["work_type_probs"][0]
    f_id = int(res["field_preds"][0])
    t_id = int(res["work_type_preds"][0])
    return {
        "field_id": f_id,
        "field_confidence": round(float(f_probs[f_id - 1]) * 100.0, 1),
        "field_probabilities": {idx + 1: round(float(p) * 100.0, 1) for idx, p in enumerate(f_probs)},
        "work_type_id": t_id,
        "work_type_name": WORK_TYPE_GROUPS.get(t_id, "Bài báo khoa học"),
        "work_type_confidence": round(float(t_probs[t_id - 1]) * 100.0, 1),
        "infer_ms": res["avg_infer_ms"],
    }
