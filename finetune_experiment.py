"""
finetune_experiment.py
======================
Fine-tuning experiment: BAAI/bge-base-en-v1.5 -> models/bge-pharma-v1/
Run from: f:\Vintyaa\projects\Product Name Automation

Does NOT modify:
  transformer.py, PRODUCT_MASTER.xlsx, hard_negative_mining.csv,
  training_split.csv, evaluation_set.csv, baseline_evaluation_results.csv
"""
import csv, sys, time, unicodedata, re, os
from pathlib import Path
import numpy as np
import torch
import openpyxl
import datasets

sys.stdout.reconfigure(encoding="utf-8")

PROJECT     = Path(r"f:\Vintyaa\projects\Product Name Automation")
TRAIN_CSV   = PROJECT / "training_split.csv"
EVAL_CSV    = PROJECT / "evaluation_set.csv"
MASTER_XLSX = PROJECT / "PRODUCT_MASTER.xlsx"
BASELINE_CSV= PROJECT / "baseline_evaluation_results.csv"
MODEL_DIR   = PROJECT / "models" / "bge-pharma-v1"
EMB_DIR     = PROJECT / "embeddings"
FT_EMB_PATH = EMB_DIR / "finetuned_master_embeddings.npy"
FT_META_PATH= EMB_DIR / "finetuned_master_meta.csv"
OUT_RESULTS = PROJECT / "finetuned_evaluation_results.csv"
OUT_REPORT  = PROJECT / "finetuned_evaluation_report.txt"

MODEL_BASE  = "BAAI/bge-base-en-v1.5"
RETRIEVAL_K = 10

# ── hyper-parameters ────────────────────────────────────────────────────
BATCH_SIZE      = 4
GRAD_ACCUM      = 4
EPOCHS          = 1
LR              = 2e-5
WARMUP_RATIO    = 0.1
SCALE           = 20.0
ENCODE_BATCH    = 128

print("=" * 64)
print("  FINE-TUNING EXPERIMENT — bge-pharma-v1")
print("=" * 64)
print(f"  Base model : {MODEL_BASE}")
print(f"  Output     : {MODEL_DIR}")
print(f"  Batch size : {BATCH_SIZE}  (grad_accum={GRAD_ACCUM}, eff={BATCH_SIZE*GRAD_ACCUM})")
print(f"  Epochs     : {EPOCHS}")
print(f"  LR         : {LR}")
print(f"  Scale (τ)  : {SCALE}")
print()

# ── helpers ─────────────────────────────────────────────────────────────
def _norm(s):
    s2 = unicodedata.normalize("NFKC", s)
    return re.sub(r"\s+", " ", s2).strip().upper()

# ── 1. Check API ─────────────────────────────────────────────────────────
from sentence_transformers import SentenceTransformer
from sentence_transformers.sentence_transformer.losses import MultipleNegativesRankingLoss
from sentence_transformers import SentenceTransformerTrainer, SentenceTransformerTrainingArguments

try:
    from sentence_transformers.sentence_transformer.training_args import BatchSamplers
    HAS_BATCH_SAMPLERS = True
    print("BatchSamplers.NO_DUPLICATES: available ✓", flush=True)
except ImportError:
    try:
        from sentence_transformers.training_args import BatchSamplers
        HAS_BATCH_SAMPLERS = True
        print("BatchSamplers.NO_DUPLICATES: available (fallback import) ✓", flush=True)
    except ImportError:
        HAS_BATCH_SAMPLERS = False
        print("BatchSamplers: not available, skipping", flush=True)

# ── 2. Load training data ────────────────────────────────────────────────
print("\nLoading training data ...", flush=True)
with open(TRAIN_CSV, encoding="utf-8") as f:
    raw_rows = list(csv.DictReader(f))

anchors = []; positives = []; neg1s = []; neg2s = []; neg3s = []; neg4s = []
skipped = 0
for r in raw_rows:
    a  = r["Input_Column"].strip()
    p  = r["Positive_Product"].strip()
    n1 = r["Explicit_Negative"].strip()
    n2 = r["Mined_Negative_1"].strip()
    n3 = r["Mined_Negative_2"].strip()
    n4 = r["Mined_Negative_3"].strip()
    if not all([a, p, n1, n2, n3, n4]):
        skipped += 1
        continue
    anchors.append(a); positives.append(p)
    neg1s.append(n1); neg2s.append(n2); neg3s.append(n3); neg4s.append(n4)

print(f"  {len(anchors)} examples loaded  ({skipped} skipped due to blank fields).")

# Dataset: column order defines anchor→positive→neg1→neg2→neg3→neg4
train_ds = datasets.Dataset.from_dict({
    "anchor":     anchors,
    "positive":   positives,
    "negative_1": neg1s,
    "negative_2": neg2s,
    "negative_3": neg3s,
    "negative_4": neg4s,
})
print(f"  Dataset columns: {train_ds.column_names}")

# ── 3. Load base model ───────────────────────────────────────────────────
print("\nLoading base model ...", flush=True)
t0 = time.perf_counter()
model = SentenceTransformer(MODEL_BASE, device="cpu")
print(f"  Loaded in {time.perf_counter()-t0:.1f}s.", flush=True)

# ── 4. Loss ──────────────────────────────────────────────────────────────
loss = MultipleNegativesRankingLoss(model, scale=SCALE)
print(f"  Loss: MultipleNegativesRankingLoss(scale={SCALE})", flush=True)
print("  NOTE: hardness_mode/hardness_strength are NOT parameters of this loss")
print("        in sentence-transformers 6.0.1. Using standard cosine MNRL.", flush=True)

# ── 5. Training arguments ────────────────────────────────────────────────
os.makedirs(str(MODEL_DIR), exist_ok=True)
trainer_kwargs = dict(
    output_dir=str(MODEL_DIR),
    num_train_epochs=EPOCHS,
    per_device_train_batch_size=BATCH_SIZE,
    gradient_accumulation_steps=GRAD_ACCUM,
    learning_rate=LR,
    warmup_steps=WARMUP_RATIO,      # float = ratio; warmup_ratio deprecated in Transformers v5+
    fp16=False, bf16=False,
    use_cpu=True,                    # no_cuda renamed to use_cpu in newer Trainer
    logging_steps=25,
    save_strategy="epoch",
    report_to="none",
    dataloader_drop_last=True,   # MNRL needs consistent batch size
)
if HAS_BATCH_SAMPLERS:
    trainer_kwargs["batch_sampler"] = BatchSamplers.NO_DUPLICATES

args = SentenceTransformerTrainingArguments(**trainer_kwargs)

# ── 6. Train ─────────────────────────────────────────────────────────────
print(f"\nStarting 1-epoch training on {len(anchors)} examples ...", flush=True)
print(f"  Effective batch = {BATCH_SIZE} × {GRAD_ACCUM} = {BATCH_SIZE*GRAD_ACCUM}", flush=True)
n_steps = max(1, len(anchors) // BATCH_SIZE)
n_opt_steps = max(1, n_steps // GRAD_ACCUM)
print(f"  Approx batches: {n_steps}  optimizer steps: {n_opt_steps}", flush=True)

trainer = SentenceTransformerTrainer(
    model=model,
    args=args,
    train_dataset=train_ds,
    loss=loss,
)

t_train_start = time.perf_counter()
train_result = trainer.train()
t_train_elapsed = time.perf_counter() - t_train_start
print(f"\nTraining complete in {t_train_elapsed:.0f}s  ({t_train_elapsed/60:.1f} min).", flush=True)
print(f"  Train result: {train_result}", flush=True)

# ── 7. Save fine-tuned model ─────────────────────────────────────────────
model.save_pretrained(str(MODEL_DIR))
print(f"  Model saved: {MODEL_DIR}", flush=True)

# ── 8. Load fine-tuned model fresh ───────────────────────────────────────
print("\nLoading fine-tuned model from disk ...", flush=True)
ft_model = SentenceTransformer(str(MODEL_DIR), device="cpu")
print("  Fine-tuned model loaded.", flush=True)

# ── 9. Load PRODUCT_MASTER ───────────────────────────────────────────────
print("Loading PRODUCT_MASTER ...", flush=True)
wb = openpyxl.load_workbook(str(MASTER_XLSX), read_only=True, data_only=True)
ws = wb.active
it = ws.iter_rows(values_only=True)
hdr = {str(h).strip().upper(): i for i, h in enumerate(next(it)) if h is not None}
pc_i, pn_i, bn_i = hdr["PRODUCT_CODE"], hdr["PRODUCT_NAME"], hdr["BRAND_NAME"]
master_recs = []; consec = 0
for row in it:
    if all(v is None or str(v).strip() == "" for v in row):
        consec += 1
        if consec >= 3: break
        continue
    consec = 0
    code  = str(row[pc_i]).strip() if row[pc_i] else ""
    name  = str(row[pn_i]).strip() if row[pn_i] else ""
    brand = str(row[bn_i]).strip().upper() if row[bn_i] else ""
    if not code: continue
    master_recs.append({"PRODUCT_CODE": code, "product_name": name, "BRAND_NAME": brand})
wb.close()
print(f"  {len(master_recs)} master records.", flush=True)

exact_name = {}; normed_name = {}; code_to_rec = {}
for rec in master_recs:
    exact_name.setdefault(rec["product_name"], []).append(rec)
    normed_name.setdefault(_norm(rec["product_name"]), []).append(rec)
    code_to_rec[rec["PRODUCT_CODE"]] = rec

def resolve_code(name):
    hits = exact_name.get(name, []) or normed_name.get(_norm(name), [])
    return hits[0]["PRODUCT_CODE"] if hits else ""

# ── 10. Re-embed PRODUCT_MASTER with fine-tuned model ────────────────────
print("\nRe-embedding 2765 master records with fine-tuned model ...", flush=True)
os.makedirs(str(EMB_DIR), exist_ok=True)
master_texts = [r["product_name"] for r in master_recs]
t0 = time.perf_counter()
ft_master_emb = ft_model.encode(
    master_texts,
    batch_size=ENCODE_BATCH,
    normalize_embeddings=True,
    show_progress_bar=True,
    convert_to_numpy=True,
).astype(np.float32)
print(f"  Embedded {len(master_recs)} records in {time.perf_counter()-t0:.1f}s.", flush=True)
print(f"  Shape: {ft_master_emb.shape}  dtype: {ft_master_emb.dtype}")

np.save(str(FT_EMB_PATH), ft_master_emb)
print(f"  Saved: {FT_EMB_PATH}", flush=True)

# Save meta csv (code + name mapping for index alignment)
with open(FT_META_PATH, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["idx", "PRODUCT_CODE", "product_name", "BRAND_NAME"])
    w.writeheader()
    for i, r in enumerate(master_recs):
        w.writerow({"idx": i, **r})
print(f"  Saved meta: {FT_META_PATH}", flush=True)

# ── 11. Build retrieval helpers (same logic as transformer.py) ────────────
def get_brand_indices(records, brand_hint):
    if not brand_hint:
        return list(range(len(records)))
    filtered = [i for i, r in enumerate(records) if r["BRAND_NAME"] == brand_hint.upper()]
    return filtered if filtered else list(range(len(records)))

def retrieve_top_k_local(query_vec, master_matrix, indices, records, k=10):
    if indices:
        sub_matrix = master_matrix[indices]
        scores = sub_matrix @ query_vec
        top_local = np.argsort(scores)[::-1][:k]
        results = []
        for rank, local_idx in enumerate(top_local, 1):
            global_idx = indices[local_idx]
            results.append({
                "rank": rank,
                "PRODUCT_CODE": records[global_idx]["PRODUCT_CODE"],
                "product_name": records[global_idx]["product_name"],
                "similarity_score": float(scores[local_idx]),
            })
        return results
    return []

master_matrix = ft_master_emb  # shape (N, 768), already normalised

# ── 12. Load evaluation set ───────────────────────────────────────────────
print("\nLoading evaluation set ...", flush=True)
with open(EVAL_CSV, encoding="utf-8") as f:
    eval_rows = list(csv.DictReader(f))
for r in eval_rows:
    if not r["Positive_Product_Code"]:
        r["Positive_Product_Code"] = resolve_code(r["Positive_Product"])
print(f"  {len(eval_rows)} eval rows.", flush=True)

# ── 13. Batch encode eval queries with fine-tuned model ──────────────────
unique_inputs = list({r["Input_Column"] for r in eval_rows})
print(f"Batch-encoding {len(unique_inputs)} unique eval queries ...", flush=True)
t0 = time.perf_counter()
ft_vecs = ft_model.encode(
    unique_inputs, batch_size=ENCODE_BATCH,
    normalize_embeddings=True, convert_to_numpy=True
).astype(np.float32)
input_to_vec = {text: ft_vecs[i] for i, text in enumerate(unique_inputs)}
print(f"  Encoded in {time.perf_counter()-t0:.1f}s.", flush=True)

# ── 14. Retrieval ─────────────────────────────────────────────────────────
unique_groups = list({(r["Input_Column"], r["Positive_Product"]) for r in eval_rows})
print(f"Running Top-{RETRIEVAL_K} retrieval for {len(unique_groups)} groups ...", flush=True)
group_ret = {}
for ocr_input, pos_product in unique_groups:
    pos_code   = resolve_code(pos_product)
    query_vec  = input_to_vec[ocr_input]
    pos_rec    = code_to_rec.get(pos_code, {})
    brand_hint = pos_rec.get("BRAND_NAME", "").strip() or None
    indices    = get_brand_indices(master_recs, brand_hint)
    candidates = retrieve_top_k_local(query_vec, master_matrix, indices, master_recs, k=RETRIEVAL_K)
    group_ret[(ocr_input, pos_product)] = (pos_code, candidates)

# ── 15. Build per-row results ─────────────────────────────────────────────
FT_FIELDS = [
    "Input_Column","Positive_Product","Positive_Product_Code",
    "Positive_Rank","Positive_Score",
    "Top1_Product","Top1_Score","Top10_Contains_Positive","Excel_Row",
]
ft_rows = []
for r in eval_rows:
    key = (r["Input_Column"], r["Positive_Product"])
    pos_code, candidates = group_ret[key]
    pos_rank = 0; pos_score = ""
    for c in candidates:
        if c["PRODUCT_CODE"] == pos_code:
            pos_rank = c["rank"]; pos_score = f"{c['similarity_score']:.4f}"; break
    top1 = candidates[0] if candidates else {}
    ft_rows.append({
        "Input_Column":            r["Input_Column"],
        "Positive_Product":        r["Positive_Product"],
        "Positive_Product_Code":   pos_code,
        "Positive_Rank":           pos_rank,
        "Positive_Score":          pos_score,
        "Top1_Product":            top1.get("product_name",""),
        "Top1_Score":              f"{top1['similarity_score']:.4f}" if top1 else "",
        "Top10_Contains_Positive": pos_rank > 0,
        "Excel_Row":               r["Excel_Row"],
    })

with open(OUT_RESULTS, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=FT_FIELDS)
    w.writeheader(); w.writerows(ft_rows)
print(f"Written: {OUT_RESULTS}", flush=True)

# ── 16. Metrics ───────────────────────────────────────────────────────────
seen = set(); gm = []
for br in ft_rows:
    key = (br["Input_Column"], br["Positive_Product"])
    if key not in seen:
        seen.add(key); gm.append(br)
N = len(gm)

def R(k):
    return sum(1 for r in gm if isinstance(r["Positive_Rank"],int) and 0<r["Positive_Rank"]<=k)/N
def MRR():
    return sum(1.0/r["Positive_Rank"] for r in gm
               if isinstance(r["Positive_Rank"],int) and r["Positive_Rank"]>0)/N

r1,r3,r5,r10 = R(1),R(3),R(5),R(10)
mrr_val = MRR()
not_in_top10 = sum(1 for r in gm if not(isinstance(r["Positive_Rank"],int) and r["Positive_Rank"]>0))

from collections import Counter
rank_dist = Counter(r["Positive_Rank"] for r in gm)

# Load baseline metrics from baseline CSV for comparison
base_gm_ranks = {}
with open(BASELINE_CSV, encoding="utf-8") as f:
    for br in csv.DictReader(f):
        key = (br["Input_Column"], br["Positive_Product"])
        if key not in base_gm_ranks:
            base_gm_ranks[key] = int(br["Positive_Rank"])
N_base = len(base_gm_ranks)

def base_R(k):
    return sum(1 for rank in base_gm_ranks.values() if 0<rank<=k)/N_base
def base_MRR():
    return sum(1.0/r for r in base_gm_ranks.values() if r>0)/N_base

b1,b3,b5,b10 = base_R(1),base_R(3),base_R(5),base_R(10)
b_mrr = base_MRR()

# ── 17. Report ────────────────────────────────────────────────────────────
lines = []
def A(s=""): lines.append(s)
A("=" * 64)
A("  FINE-TUNING EXPERIMENT REPORT")
A("=" * 64)
A()
A(f"  Base model   : {MODEL_BASE}")
A(f"  Fine-tuned   : {MODEL_DIR}")
A(f"  Training rows: {len(anchors)}")
A(f"  Epochs       : {EPOCHS}")
A(f"  Batch/eff    : {BATCH_SIZE}/{BATCH_SIZE*GRAD_ACCUM}")
A(f"  LR           : {LR}")
A(f"  Scale (τ)    : {SCALE}")
A(f"  Train time   : {t_train_elapsed:.0f}s  ({t_train_elapsed/60:.1f} min)")
A()
A("─" * 64)
A("  RETRIEVAL METRICS (170 unique eval groups, Top-10)")
A("─" * 64)
A(f"  {'Metric':<12} {'Baseline':>10} {'Fine-tuned':>12} {'Delta':>8}")
A(f"  {'─'*12} {'─'*10} {'─'*12} {'─'*8}")
A(f"  {'Recall@1':<12} {b1*100:>9.1f}%  {r1*100:>11.1f}%  {(r1-b1)*100:>+7.1f}%")
A(f"  {'Recall@3':<12} {b3*100:>9.1f}%  {r3*100:>11.1f}%  {(r3-b3)*100:>+7.1f}%")
A(f"  {'Recall@5':<12} {b5*100:>9.1f}%  {r5*100:>11.1f}%  {(r5-b5)*100:>+7.1f}%")
A(f"  {'Recall@10':<12} {b10*100:>9.1f}%  {r10*100:>11.1f}%  {(r10-b10)*100:>+7.1f}%")
A(f"  {'MRR':<12} {b_mrr:>10.4f}  {mrr_val:>12.4f}  {(mrr_val-b_mrr):>+8.4f}")
A()
A("  Rank distribution — fine-tuned (unique groups):")
for rank in sorted(rank_dist):
    label = f"Rank {rank}" if rank>0 else "NOT_FOUND"
    A(f"    {label:15}: {rank_dist[rank]:4d}  ({100*rank_dist[rank]/N:.1f}%)")
A()
A(f"  Positive NOT in Top-10 (fine-tuned) : {not_in_top10}  ({100*not_in_top10/N:.1f}%)")
A(f"  Positive NOT in Top-10 (baseline)   : 7  (4.1%)")
A()
A("─" * 64)
A("  FILES")
A("─" * 64)
A(f"  Fine-tuned model           : {MODEL_DIR}")
A(f"  Fine-tuned master embeddings: {FT_EMB_PATH}")
A(f"  Evaluation results         : {OUT_RESULTS}")
A()
A("=" * 64)
A("  END OF REPORT")
A("=" * 64)

report_text = "\n".join(lines)
with open(OUT_REPORT, "w", encoding="utf-8") as f:
    f.write(report_text)
print(f"\nWritten: {OUT_REPORT}", flush=True)
print()
print(report_text)
