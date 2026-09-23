"""
finetune_v2.py
==============
V2 fine-tuning: bge-pharma-v1 -> bge-pharma-v2

Training data : v2_hard_negative_mining_clean.csv  (108,831 rows)
Start model   : models/bge-pharma-v1/              (V1 fine-tuned checkpoint)
Output        : models/bge-pharma-v2/
Loss          : MultipleNegativesRankingLoss(scale=20)
Dataset cols  : anchor, positive, negative_1, negative_2, negative_3

Hyperparameters (all per V2_TRAINING_PLAN.md):
  epochs            : 1
  per-device batch  : 4
  grad_accumulation : 4   -> effective batch 16
  learning rate     : 2e-5
  warmup            : 0.1 (float -> treated as ratio by Trainer)
  seed              : 42
  device            : auto-detect (CUDA if available, else CPU)
  fp16              : True on CUDA, False on CPU

Safety guards:
  - V1 model directory is never written to
  - evaluation_set.csv is never loaded for training
  - raw mining CSV is never touched
  - skipped-row count asserted == 0 (validated dataset)
"""

import csv
import os
import sys
import time
from pathlib import Path

import torch
import datasets

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ============================================================================
# CONFIG
# ============================================================================

PROJECT      = Path(r"F:\Vintyaa\projects\Product Name Automation")
CLEAN_CSV    = PROJECT / "v2_hard_negative_mining_clean.csv"
V1_MODEL_DIR = PROJECT / "models" / "bge-pharma-v1"
V2_MODEL_DIR = PROJECT / "models" / "bge-pharma-v2"
EVAL_CSV     = PROJECT / "evaluation_set.csv"   # never used for training

EPOCHS        = 1
BATCH_SIZE    = 4
GRAD_ACCUM    = 4           # effective batch = 4 * 4 = 16
LR            = 2e-5
WARMUP_RATIO  = 0.1         # passed as warmup_steps float -> ratio
SCALE         = 20.0
SEED          = 42
LOGGING_STEPS = 100         # ~6801 total steps; log every 100

USE_CUDA = torch.cuda.is_available()
DEVICE   = "cuda" if USE_CUDA else "cpu"
USE_CPU  = not USE_CUDA
FP16     = USE_CUDA         # enable fp16 on GPU only

# ============================================================================
# IMPORTS
# ============================================================================

from sentence_transformers import (
    SentenceTransformer,
    SentenceTransformerTrainer,
    SentenceTransformerTrainingArguments,
)
from sentence_transformers.sentence_transformer.losses import MultipleNegativesRankingLoss

try:
    from sentence_transformers.sentence_transformer.training_args import BatchSamplers
    HAS_BATCH_SAMPLERS = True
except ImportError:
    try:
        from sentence_transformers.training_args import BatchSamplers
        HAS_BATCH_SAMPLERS = True
    except ImportError:
        HAS_BATCH_SAMPLERS = False

# ============================================================================
# BANNER
# ============================================================================

print("=" * 68)
print("  V2 FINE-TUNING  --  bge-pharma-v2")
print("=" * 68)
print(f"  Start model   : {V1_MODEL_DIR}")
print(f"  Output        : {V2_MODEL_DIR}")
print(f"  Training data : {CLEAN_CSV}")
print(f"  Device        : {DEVICE}  (fp16={FP16})")
print(f"  Epochs        : {EPOCHS}")
print(f"  Batch / accum : {BATCH_SIZE} / {GRAD_ACCUM}  (eff. {BATCH_SIZE * GRAD_ACCUM})")
print(f"  LR            : {LR}")
print(f"  Scale (tau)   : {SCALE}")
print(f"  Warmup ratio  : {WARMUP_RATIO}")
print(f"  Seed          : {SEED}")
print()

# ============================================================================
# SAFETY CHECKS
# ============================================================================

assert CLEAN_CSV.exists(),    f"MISSING: {CLEAN_CSV}"
assert V1_MODEL_DIR.exists(), f"MISSING: {V1_MODEL_DIR}"
assert EVAL_CSV.exists(),     f"MISSING: {EVAL_CSV}"
assert V2_MODEL_DIR != V1_MODEL_DIR, "V2 output dir must differ from V1 -- aborting"

os.makedirs(str(V2_MODEL_DIR), exist_ok=True)
print(f"  Output directory confirmed: {V2_MODEL_DIR}")

# ============================================================================
# 1. LOAD TRAINING DATA
# ============================================================================

print("\nLoading training data ...", flush=True)
t0 = time.perf_counter()

with open(CLEAN_CSV, "r", encoding="utf-8") as fh:
    raw_rows = list(csv.DictReader(fh))

print(f"  {len(raw_rows):,} rows loaded in {time.perf_counter() - t0:.1f}s.", flush=True)

anchors   = []
positives = []
neg1s     = []
neg2s     = []
neg3s     = []
skipped   = 0

for r in raw_rows:
    a  = r["Input_Column"].strip()
    p  = r["Positive_Product"].strip()
    n1 = r["Mined_Negative_1"].strip()
    n2 = r["Mined_Negative_2"].strip()
    n3 = r["Mined_Negative_3"].strip()
    if not all([a, p, n1, n2, n3]):
        skipped += 1
        continue
    anchors.append(a)
    positives.append(p)
    neg1s.append(n1)
    neg2s.append(n2)
    neg3s.append(n3)

print(f"  {len(anchors):,} examples ready  ({skipped} skipped -- expected 0).", flush=True)
assert skipped == 0, f"{skipped} rows have blank fields -- validation should have caught this."

# ============================================================================
# 2. BUILD HUGGINGFACE DATASET
# ============================================================================

# Column order matters for MultipleNegativesRankingLoss:
#   column 0 = anchor, column 1 = positive, columns 2+ = negatives
train_ds = datasets.Dataset.from_dict({
    "anchor":     anchors,
    "positive":   positives,
    "negative_1": neg1s,
    "negative_2": neg2s,
    "negative_3": neg3s,
})
print(f"  Dataset columns : {train_ds.column_names}")
print(f"  Dataset size    : {len(train_ds):,}")

steps_per_epoch = len(anchors) // (BATCH_SIZE * GRAD_ACCUM)   # drop_last=True
total_steps     = steps_per_epoch * EPOCHS
warmup_n        = int(total_steps * WARMUP_RATIO)
print(f"  Steps per epoch : {steps_per_epoch:,}")
print(f"  Total steps     : {total_steps:,}")
print(f"  Warmup steps    : ~{warmup_n:,}  ({WARMUP_RATIO * 100:.0f}% of total)")

# ============================================================================
# 3. LOAD V1 START MODEL
# ============================================================================

print(f"\nLoading V1 start model: {V1_MODEL_DIR} ...", flush=True)
t0 = time.perf_counter()
model = SentenceTransformer(str(V1_MODEL_DIR), device=DEVICE)
print(f"  Loaded in {time.perf_counter() - t0:.1f}s  on {DEVICE}.", flush=True)

probe = model.encode(["probe"], normalize_embeddings=True, show_progress_bar=False)
assert probe.shape[1] == 768, f"Unexpected embedding dim: {probe.shape[1]}"
print(f"  Embedding dim   : {probe.shape[1]}  (expected 768)  OK")

# ============================================================================
# 4. LOSS
# ============================================================================

loss = MultipleNegativesRankingLoss(model, scale=SCALE)
print(f"\n  Loss: MultipleNegativesRankingLoss(scale={SCALE})")
print("  NOTE: each anchor faces 3 explicit hard negatives + (B-1) in-batch negatives.")
print(f"        With eff. batch {BATCH_SIZE * GRAD_ACCUM}: ~{BATCH_SIZE * GRAD_ACCUM - 1 + 3} negatives/anchor/step.")

# ============================================================================
# 5. TRAINING ARGUMENTS
# ============================================================================

trainer_kwargs = dict(
    output_dir                  = str(V2_MODEL_DIR),
    num_train_epochs            = EPOCHS,
    per_device_train_batch_size = BATCH_SIZE,
    gradient_accumulation_steps = GRAD_ACCUM,
    learning_rate               = LR,
    warmup_steps                = WARMUP_RATIO,   # float -> ratio (same as V1)
    fp16                        = FP16,
    bf16                        = False,
    use_cpu                     = USE_CPU,
    logging_steps               = LOGGING_STEPS,
    save_strategy               = "epoch",
    report_to                   = "none",
    dataloader_drop_last        = True,           # required by MNRL
    seed                        = SEED,
)

if HAS_BATCH_SAMPLERS:
    trainer_kwargs["batch_sampler"] = BatchSamplers.NO_DUPLICATES
    print("  BatchSamplers.NO_DUPLICATES: active")
else:
    print("  BatchSamplers: not available -- proceeding without")

args = SentenceTransformerTrainingArguments(**trainer_kwargs)

# ============================================================================
# 6. TRAINER
# ============================================================================

trainer = SentenceTransformerTrainer(
    model         = model,
    args          = args,
    train_dataset = train_ds,
    loss          = loss,
)

# ============================================================================
# 7. TRAIN
# ============================================================================

print("\n" + "=" * 68)
print("  Starting V2 training")
print(f"  Examples   : {len(anchors):,}")
print(f"  Steps      : ~{total_steps:,}")
print(f"  Device     : {DEVICE}  (fp16={FP16})")
print("=" * 68, flush=True)

t_train_start   = time.perf_counter()
train_result    = trainer.train()
t_train_elapsed = time.perf_counter() - t_train_start

print(f"\nTraining complete in {t_train_elapsed:.0f}s  ({t_train_elapsed / 60:.1f} min).", flush=True)
print(f"  Train result: {train_result}", flush=True)

# ============================================================================
# 8. SAVE V2 MODEL
# ============================================================================

model.save_pretrained(str(V2_MODEL_DIR))
print(f"\n  V2 model saved: {V2_MODEL_DIR}", flush=True)

# ============================================================================
# 9. WRITE TRAINING SUMMARY
# ============================================================================

log_history  = trainer.state.log_history
loss_entries = [e for e in log_history if "loss" in e]

sep = "=" * 68
summary_lines = [
    sep,
    "  V2 TRAINING SUMMARY",
    sep,
    "",
    f"  Start model       : {V1_MODEL_DIR}",
    f"  Output model      : {V2_MODEL_DIR}",
    f"  Training data     : {CLEAN_CSV}",
    f"  Training examples : {len(anchors):,}",
    f"  Negatives/example : 3  (mined hard negatives)",
    f"  Total neg slots   : {len(anchors) * 3:,}",
    "",
    f"  Loss              : MultipleNegativesRankingLoss(scale={SCALE})",
    f"  Epochs            : {EPOCHS}",
    f"  Batch size        : {BATCH_SIZE}  (eff. {BATCH_SIZE * GRAD_ACCUM})",
    f"  Grad accumulation : {GRAD_ACCUM}",
    f"  Learning rate     : {LR}",
    f"  Warmup ratio      : {WARMUP_RATIO}",
    f"  Seed              : {SEED}",
    f"  Device            : {DEVICE}  (fp16={FP16})",
    f"  Optimizer steps   : {trainer.state.global_step}",
    f"  Train time        : {t_train_elapsed:.0f}s  ({t_train_elapsed / 60:.1f} min)",
    "",
    "  Loss log:",
]

for entry in loss_entries:
    step  = entry.get("step",   "?")
    epoch = entry.get("epoch",  "?")
    lv    = entry.get("loss",   "?")
    lr    = entry.get("learning_rate", "?")
    epoch_str = f"{epoch:.3f}" if isinstance(epoch, float) else str(epoch)
    loss_str  = f"{lv:.4f}"    if isinstance(lv, float)    else str(lv)
    lr_str    = f"{lr:.2e}"    if isinstance(lr, float)     else str(lr)
    summary_lines.append(f"    step={step:>6}  epoch={epoch_str}  loss={loss_str}  lr={lr_str}")

summary_lines += [
    "",
    "  NEXT: Run finetune_v2_eval.py to compare V1 vs V2 retrieval metrics.",
    "",
    sep,
    "  END",
    sep,
]

summary_text = "\n".join(summary_lines)
summary_path = PROJECT / "V2_TRAINING_SUMMARY.txt"

with open(summary_path, "w", encoding="utf-8") as fh:
    fh.write(summary_text)

print(f"\n  Summary written: {summary_path}", flush=True)
print()
print(summary_text)
