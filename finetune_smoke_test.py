"""
finetune_smoke_test.py
Smoke test: 8 examples, 1 training step.
Run from project directory.
"""
import sys, csv, copy, time
from pathlib import Path
import torch
import datasets
from sentence_transformers import SentenceTransformer
from sentence_transformers.sentence_transformer.losses import MultipleNegativesRankingLoss

sys.stdout.reconfigure(encoding="utf-8")
PROJECT = Path(r"f:\Vintyaa\projects\Product Name Automation")
TRAIN_CSV = PROJECT / "training_split.csv"

print("=" * 60)
print("  SMOKE TEST — MultipleNegativesRankingLoss + ST 6.0.1")
print("=" * 60)

# ── 1. Check BatchSamplers availability ──────────────────────────────────
try:
    from sentence_transformers.sentence_transformer.training_args import BatchSamplers
    print("BatchSamplers: imported from sentence_transformers.sentence_transformer.training_args ✓")
except ImportError:
    try:
        from sentence_transformers.training_args import BatchSamplers
        print("BatchSamplers: imported from sentence_transformers.training_args ✓")
    except ImportError:
        BatchSamplers = None
        print("BatchSamplers: NOT available — will train without it")

# ── 2. Check SentenceTransformerTrainer ──────────────────────────────────
try:
    from sentence_transformers import SentenceTransformerTrainer, SentenceTransformerTrainingArguments
    print("SentenceTransformerTrainer: available ✓")
except ImportError:
    print("FATAL: SentenceTransformerTrainer not available")
    sys.exit(1)

# ── 3. Load 8 training rows ───────────────────────────────────────────────
print("\nLoading 8 training rows ...", flush=True)
with open(TRAIN_CSV, encoding="utf-8") as f:
    rows = list(csv.DictReader(f))[:8]
print(f"  {len(rows)} rows loaded.")

anchors     = [r["Input_Column"]       for r in rows]
positives   = [r["Positive_Product"]   for r in rows]
neg1_list   = [r["Explicit_Negative"]  for r in rows]
neg2_list   = [r["Mined_Negative_1"]   for r in rows]
neg3_list   = [r["Mined_Negative_2"]   for r in rows]
neg4_list   = [r["Mined_Negative_3"]   for r in rows]

# Verify no blanks in smoke examples
for i, (a, p, n1, n2, n3, n4) in enumerate(
        zip(anchors, positives, neg1_list, neg2_list, neg3_list, neg4_list)):
    assert all([a, p, n1, n2, n3, n4]), f"Row {i} has blank field"
print("  All fields non-blank ✓")

# ── 4. Load model ─────────────────────────────────────────────────────────
print("\nLoading BAAI/bge-base-en-v1.5 ...", flush=True)
t0 = time.perf_counter()
model = SentenceTransformer("BAAI/bge-base-en-v1.5", device="cpu")
print(f"  Loaded in {time.perf_counter()-t0:.1f}s.")

# Snapshot a parameter before training
param_name = "0.auto_model.encoder.layer.0.attention.self.query.weight"
try:
    before_param = model[0].auto_model.encoder.layer[0].attention.self.query.weight.detach().clone()
    snapshot_ok = True
except Exception as e:
    print(f"  Could not snapshot param: {e}")
    snapshot_ok = False

# ── 5. Build dataset ──────────────────────────────────────────────────────
# Column order matters: anchor → positive → negatives
train_ds = datasets.Dataset.from_dict({
    "anchor":     anchors,
    "positive":   positives,
    "negative_1": neg1_list,
    "negative_2": neg2_list,
    "negative_3": neg3_list,
    "negative_4": neg4_list,
})
print(f"\nDataset: {train_ds}")
print(f"  Columns (order): {train_ds.column_names}")

# ── 6. Build loss ─────────────────────────────────────────────────────────
loss = MultipleNegativesRankingLoss(model, scale=20.0)
print(f"\nLoss: {loss.__class__.__name__}(scale=20, similarity=cosine)")
print("  NOTE: 'hardness_mode'/'hardness_strength' are not parameters of")
print("  MultipleNegativesRankingLoss in sentence-transformers 6.0.1.")
print("  Proceeding with standard scale=20, cosine similarity.")

# ── 7. Training args — 1 step only ────────────────────────────────────────
import tempfile, os
smoke_dir = str(PROJECT / "models" / "smoke_test_tmp")
os.makedirs(smoke_dir, exist_ok=True)

trainer_kwargs = dict(
    output_dir=smoke_dir,
    num_train_epochs=1,
    per_device_train_batch_size=4,
    gradient_accumulation_steps=1,   # smoke: accumulate nothing
    max_steps=1,                     # exactly 1 optimizer step
    learning_rate=2e-5,
    warmup_ratio=0.0,
    fp16=False, bf16=False,
    use_cpu=True,                    # no_cuda renamed to use_cpu in newer Trainer
    logging_steps=1,
    save_strategy="no",
    report_to="none",
)
if BatchSamplers is not None:
    trainer_kwargs["batch_sampler"] = BatchSamplers.NO_DUPLICATES

args = SentenceTransformerTrainingArguments(**trainer_kwargs)

# ── 8. Trainer ────────────────────────────────────────────────────────────
trainer = SentenceTransformerTrainer(
    model=model,
    args=args,
    train_dataset=train_ds,
    loss=loss,
)

print("\nRunning 1 training step ...", flush=True)
t0 = time.perf_counter()
result = trainer.train()
elapsed = time.perf_counter() - t0
print(f"  Done in {elapsed:.1f}s.")
print(f"  Training result: {result}")

# ── 9. Verify loss was computed ───────────────────────────────────────────
log_history = trainer.state.log_history
print(f"\nLog history: {log_history}")
if log_history and "loss" in log_history[0]:
    loss_val = log_history[0]["loss"]
    print(f"  Loss: {loss_val:.4f}  ✓")
    assert loss_val > 0, "Loss is zero or negative — unexpected"
else:
    print("  WARNING: no loss in log history")

# ── 10. Verify parameter update ───────────────────────────────────────────
if snapshot_ok:
    after_param = model[0].auto_model.encoder.layer[0].attention.self.query.weight.detach()
    diff = (after_param - before_param).abs().max().item()
    print(f"\nMax weight delta after 1 step: {diff:.6f}")
    if diff > 1e-9:
        print("  Parameters updated ✓")
    else:
        print("  WARNING: parameters did NOT change — gradient may not be flowing")

# ── 11. Verify model can still encode ─────────────────────────────────────
test_vecs = model.encode(["CARDACE 2.5 MG TABLET", "TEMSAN 80 MG TABLET"], normalize_embeddings=True)
print(f"\nPost-training encode test: shape={test_vecs.shape}, dtype={test_vecs.dtype}")
assert test_vecs.shape == (2, 768), f"Unexpected shape: {test_vecs.shape}"
print("  Encode test ✓")

print("\n" + "=" * 60)
print("  SMOKE TEST PASSED — safe to run full training")
print("=" * 60)
