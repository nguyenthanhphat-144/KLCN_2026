# -*- coding: utf-8 -*-
"""
Kiến trúc & Siêu tham số theo đúng đặc tả PRD:
- Pre-trained weights: vinai/phobert-base (Hugging Face transformers)
- Đóng băng (freeze) các tầng ẩn phía dưới của PhoBERT (tầng 0 -> 7)
- Bổ sung tầng phân loại tuyến tính phía trên (Classification Head) khớp với số lượng nhãn lĩnh vực
- Siêu tham số tối ưu: Learning rate = 2e-5, Batch size = 16, Optimizer = AdamW
"""

import os
import json

PHOBERT_MODEL_NAME = "vinai/phobert-base"
HYPERPARAMETERS = {
    "pretrained_model": "vinai/phobert-base",
    "architecture": "RoBERTa-base Vietnamese (12 Transformer Layers, Hidden=768, Heads=12)",
    "frozen_lower_layers": "Layers 0..7 frozen (encoder.layer.0 to encoder.layer.7)",
    "classification_head": "Dropout(0.2) -> Linear(768, 256) -> GELU -> Linear(256, num_labels)",
    "learning_rate": 2e-5,
    "batch_size": 16,
    "optimizer": "AdamW (weight_decay=0.01, eps=1e-8)",
    "max_seq_length": 256,
    "epochs": 10,
}


def build_phobert_PyTorch_module_code():
    """
    Trả về định nghĩa lớp PyTorch chuẩn cho PhoBERT Fine-tuning phục vụ huấn luyện GPU/CPU
    và trình bày trong báo cáo khóa luận tốt nghiệp.
    """
    try:
        import torch
        import torch.nn as nn
        from transformers import AutoModel, AutoTokenizer
        from torch.optim import AdamW
        from torch.utils.data import Dataset, DataLoader

        class KHCNPhoBERTDataset(Dataset):
            def __init__(self, segmented_texts, labels, tokenizer, max_len=256):
                self.texts = segmented_texts
                self.labels = labels
                self.tokenizer = tokenizer
                self.max_len = max_len

            def __len__(self):
                return len(self.texts)

            def __getitem__(self, idx):
                text = str(self.texts[idx])
                encoding = self.tokenizer(
                    text,
                    truncation=True,
                    max_length=self.max_len,
                    padding="max_length",
                    return_tensors="pt",
                )
                return {
                    "input_ids": encoding["input_ids"].squeeze(0),
                    "attention_mask": encoding["attention_mask"].squeeze(0),
                    "label": torch.tensor(self.labels[idx], dtype=torch.long),
                }

        class PhoBERTForKHCNClassification(nn.Module):
            def __init__(self, num_labels=6, freeze_lower_layers=8):
                super().__init__()
                self.phobert = AutoModel.from_pretrained(PHOBERT_MODEL_NAME)

                # Đóng băng embeddings và các tầng ẩn phía dưới của PhoBERT (Tuần 6 - Mục 2b)
                for param in self.phobert.embeddings.parameters():
                    param.requires_grad = False
                for layer_idx in range(freeze_lower_layers):
                    for param in self.phobert.encoder.layer[layer_idx].parameters():
                        param.requires_grad = False

                hidden_size = self.phobert.config.hidden_size  # 768
                # Bổ sung tầng phân loại tuyến tính phía trên (Classification Head)
                self.classifier_head = nn.Sequential(
                    nn.Dropout(p=0.2),
                    nn.Linear(hidden_size, 256),
                    nn.GELU(),
                    nn.Dropout(p=0.1),
                    nn.Linear(256, num_labels),
                )

            def forward(self, input_ids, attention_mask):
                outputs = self.phobert(input_ids=input_ids, attention_mask=attention_mask)
                cls_representation = outputs.last_hidden_state[:, 0, :]  # Token <s>
                logits = self.classifier_head(cls_representation)
                return logits

        def finetune_phobert_pipeline(train_texts, train_labels, num_labels=6, epochs=5, save_dir="models/phobert_khcn"):
            device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
            tokenizer = AutoTokenizer.from_pretrained(PHOBERT_MODEL_NAME)
            dataset = KHCNPhoBERTDataset(train_texts, train_labels, tokenizer, max_len=256)
            dataloader = DataLoader(dataset, batch_size=16, shuffle=True)  # Batch size = 16

            model = PhoBERTForKHCNClassification(num_labels=num_labels, freeze_lower_layers=8).to(device)
            # Optimizer AdamW với Learning rate = 2e-5
            optimizer = AdamW(
                filter(lambda p: p.requires_grad, model.parameters()),
                lr=2e-5,
                weight_decay=0.01,
            )
            criterion = nn.CrossEntropyLoss()

            model.train()
            for epoch in range(epochs):
                total_loss = 0.0
                for batch in dataloader:
                    optimizer.zero_grad()
                    input_ids = batch["input_ids"].to(device)
                    attention_mask = batch["attention_mask"].to(device)
                    labels = batch["label"].to(device)
                    logits = model(input_ids, attention_mask)
                    loss = criterion(logits, labels)
                    loss.backward()
                    optimizer.step()
                    total_loss += loss.item()
                print(f"[PhoBERT Fine-tune] Epoch {epoch+1}/{epochs} - Loss: {total_loss/len(dataloader):.4f}")

            os.makedirs(save_dir, exist_ok=True)
            torch.save(model.state_dict(), os.path.join(save_dir, "phobert_classifier.pt"))
            with open(os.path.join(save_dir, "hyperparameters.json"), "w", encoding="utf-8") as f:
                json.dump(HYPERPARAMETERS, f, ensure_ascii=False, indent=2)
            return model

        return {
            "available": True,
            "dataset_cls": KHCNPhoBERTDataset,
            "model_cls": PhoBERTForKHCNClassification,
            "train_fn": finetune_phobert_pipeline,
        }
    except ImportError:
        return {"available": False, "hyperparameters": HYPERPARAMETERS}
