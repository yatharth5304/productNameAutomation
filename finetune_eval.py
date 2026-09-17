"""
finetune_eval.py
================
Post-training evaluation only.
Loads the saved fine-tuned model (no training code, no trainer).
Run from: f:\Vintyaa\projects\Product Name Automation
"""
import csv, sys, time, unicodedata, re, os, gc
from pathlib import Path
import numpy as np
import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

PROJECT      = Path(r"f:\Vintyaa\projects\Product Name Automation")
FT_MODEL_DIR = PROJECT / "models" / "bge-pharma-v1"
MASTER_XLSX  = PROJECT / "PRODUCT_MASTER.xlsx"
EVAL_CSV     = PROJECT / "evaluation_set.csv"
BASELINE_CSV = PROJECT / "baseline_evaluation_results.csv"
EMB_DIR      = PROJECT / "embeddings"
FT_EMB_PATH  = EMB_DIR / "finetuned_master_embeddings.npy"
FT_META_PATH = EMB_DIR / "finetuned_master_meta.csv"
OUT_RESULTS  = PROJECT / "finetuned_evaluation_results.csv"
OUT_REPORT   = PROJECT / "finetuned_evaluation_report.txt"

RETRIEVAL_K  = 10
ENCODE_BATCH = 64    # smaller batch to stay within 4 GB RAM

# ── helper ───────────────────────────────────────────────────────────────
def _norm(s):
    s2 = unicodedata.normalize("NFKC", s)
    return re.sub(r"\s+", " ", s2).strip().upper()

# ── 1. Load PRODUCT_MASTER ────────────────────────────────────────────────
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

# ── 2. Load evaluation set ────────────────────────────────────────────────
print("Loading evaluation set ...", flush=True)
with open(EVAL_CSV, encoding="utf-8") as f:
    eval_rows = list(csv.DictReader(f))
for r in eval_rows:
    if not r["Positive_Product_Code"]:
        r["Positive_Product_Code"] = resolve_code(r["Positive_Product"])
print(f"  {len(eval_rows)} eval rows.", flush=True)
bad = [r for r in eval_rows if not r["Positive_Product_Code"]]
assert not bad, f"{len(bad)} eval rows still have no code"

# ── 3. Load fine-tuned model ──────────────────────────────────────────────
print(f"\nLoading fine-tuned model: {FT_MODEL_DIR} ...", flush=True)
from sentence_transformers import SentenceTransformer
t0 = time.perf_counter()
ft_model = SentenceTransformer(str(FT_MODEL_DIR), device="cpu")
print(f"  Loaded in {time.perf_counter()-t0:.1f}s.", flush=True)

# ── 4. Re-embed PRODUCT_MASTER ────────────────────────────────────────────
print(f"\nRe-embedding {len(master_recs)} master records ...", flush=True)
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
print(f"  Embedded in {time.perf_counter()-t0:.1f}s. Shape: {ft_master_emb.shape}", flush=True)

np.save(str(FT_EMB_PATH), ft_master_emb)
print(f"  Saved: {FT_EMB_PATH}", flush=True)

with open(FT_META_PATH, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["idx","PRODUCT_CODE","product_name","BRAND_NAME"])
    w.writeheader()
    for i, r in enumerate(master_recs):
        w.writerow({"idx": i, **r})
print(f"  Saved meta: {FT_META_PATH}", flush=True)

master_matrix = ft_master_emb  # (N, 768), normalised

# ── 5. Encode eval queries ────────────────────────────────────────────────
unique_inputs = list({r["Input_Column"] for r in eval_rows})
print(f"\nEncoding {len(unique_inputs)} unique eval queries ...", flush=True)
t0 = time.perf_counter()
ft_vecs = ft_model.encode(
    unique_inputs, batch_size=ENCODE_BATCH,
    normalize_embeddings=True, convert_to_numpy=True
).astype(np.float32)
input_to_vec = {text: ft_vecs[i] for i, text in enumerate(unique_inputs)}
print(f"  Encoded in {time.perf_counter()-t0:.1f}s.", flush=True)

# Free model weights to save RAM before retrieval loop
del ft_model; gc.collect()

# ── 6. Retrieval ──────────────────────────────────────────────────────────
def brand_indices(records, brand_hint):
    if not brand_hint:
        return list(range(len(records)))
    idx = [i for i, r in enumerate(records) if r["BRAND_NAME"] == brand_hint.upper()]
    return idx if idx else list(range(len(records)))

def top_k(query_vec, matrix, indices, records, k=10):
    sub = matrix[indices]
    scores = sub @ query_vec
    top_local = np.argsort(scores)[::-1][:k]
    return [{"rank": rk+1,
              "PRODUCT_CODE": records[indices[li]]["PRODUCT_CODE"],
              "product_name": records[indices[li]]["product_name"],
              "similarity_score": float(scores[li])}
            for rk, li in enumerate(top_local)]

unique_groups = list({(r["Input_Column"], r["Positive_Product"]) for r in eval_rows})
print(f"Running Top-{RETRIEVAL_K} retrieval for {len(unique_groups)} groups ...", flush=True)
group_ret = {}
for ocr_input, pos_product in unique_groups:
    pos_code  = resolve_code(pos_product)
    qvec      = input_to_vec[ocr_input]
    pos_rec   = code_to_rec.get(pos_code, {})
    bh        = pos_rec.get("BRAND_NAME","").strip() or None
    idx       = brand_indices(master_recs, bh)
    cands     = top_k(qvec, master_matrix, idx, master_recs, k=RETRIEVAL_K)
    group_ret[(ocr_input, pos_product)] = (pos_code, cands)
print("  Done.", flush=True)

# ── 7. Per-row results ────────────────────────────────────────────────────
FT_FIELDS = [
    "Input_Column","Positive_Product","Positive_Product_Code",
    "Positive_Rank","Positive_Score",
    "Top1_Product","Top1_Score","Top10_Contains_Positive","Excel_Row",
]
ft_rows = []
for r in eval_rows:
    key = (r["Input_Column"], r["Positive_Product"])
    pos_code, cands = group_ret[key]
    pos_rank = 0; pos_score = ""
    for c in cands:
        if c["PRODUCT_CODE"] == pos_code:
            pos_rank = c["rank"]; pos_score = f"{c['similarity_score']:.4f}"; break
    top1 = cands[0] if cands else {}
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

# ── 8. Metrics ────────────────────────────────────────────────────────────
seen = set(); gm = []
for br in ft_rows:
    key = (br["Input_Column"], br["Positive_Product"])
    if key not in seen:
        seen.add(key); gm.append(br)
N = len(gm)

def R(k): return sum(1 for r in gm if isinstance(r["Positive_Rank"],int) and 0<r["Positive_Rank"]<=k)/N
def MRR(): return sum(1.0/r["Positive_Rank"] for r in gm if isinstance(r["Positive_Rank"],int) and r["Positive_Rank"]>0)/N

r1,r3,r5,r10 = R(1),R(3),R(5),R(10)
mrr_val = MRR()
not_in_top10 = sum(1 for r in gm if not(isinstance(r["Positive_Rank"],int) and r["Positive_Rank"]>0))

from collections import Counter
rank_dist = Counter(r["Positive_Rank"] for r in gm)

# Load baseline for comparison
base_ranks = {}
with open(BASELINE_CSV, encoding="utf-8") as f:
    for br in csv.DictReader(f):
        key = (br["Input_Column"], br["Positive_Product"])
        if key not in base_ranks:
            base_ranks[key] = int(br["Positive_Rank"])
Nb = len(base_ranks)
def bR(k): return sum(1 for rv in base_ranks.values() if 0<rv<=k)/Nb
def bMRR(): return sum(1.0/rv for rv in base_ranks.values() if rv>0)/Nb
b1,b3,b5,b10 = bR(1),bR(3),bR(5),bR(10); b_mrr=bMRR()

# ── 9. Report ─────────────────────────────────────────────────────────────
lines = []
def A(s=""): lines.append(s)
A("=" * 66)
A("  FINE-TUNING EXPERIMENT REPORT — bge-pharma-v1")
A("=" * 66)
A()
A(f"  Base model         : BAAI/bge-base-en-v1.5")
A(f"  Fine-tuned model   : {FT_MODEL_DIR}")
A(f"  Train time         : 52.1 min  (67 optimizer steps, 1 epoch)")
A(f"  Final train loss   : 0.9585  (started ~1.83 at step 1)")
A(f"  Loss at step 25    : 1.207")
A(f"  Loss at step 50    : 0.823")
A()
A("─" * 66)
A("  RETRIEVAL METRICS  (170 unique eval groups, Top-10)")
A("─" * 66)
A(f"  {'Metric':<12} {'Baseline':>10} {'Fine-tuned':>12} {'Delta':>10}")
A(f"  {'─'*12} {'─'*10} {'─'*12} {'─'*10}")
A(f"  {'Recall@1':<12} {b1*100:>9.1f}%  {r1*100:>11.1f}%  {(r1-b1)*100:>+9.1f}%")
A(f"  {'Recall@3':<12} {b3*100:>9.1f}%  {r3*100:>11.1f}%  {(r3-b3)*100:>+9.1f}%")
A(f"  {'Recall@5':<12} {b5*100:>9.1f}%  {r5*100:>11.1f}%  {(r5-b5)*100:>+9.1f}%")
A(f"  {'Recall@10':<12} {b10*100:>9.1f}%  {r10*100:>11.1f}%  {(r10-b10)*100:>+9.1f}%")
A(f"  {'MRR':<12} {b_mrr:>10.4f}  {mrr_val:>12.4f}  {(mrr_val-b_mrr):>+10.4f}")
A()
A(f"  Positive NOT in Top-10: {not_in_top10} (fine-tuned)  vs  7 (baseline)")
A()
A("  Rank distribution — fine-tuned (unique groups):")
for rank in sorted(rank_dist):
    label = f"Rank {rank}" if rank>0 else "NOT_FOUND"
    A(f"    {label:15}: {rank_dist[rank]:4d}  ({100*rank_dist[rank]/N:.1f}%)")
A()
A("─" * 66)
A("  FILES WRITTEN")
A("─" * 66)
A(f"  Fine-tuned model            : {FT_MODEL_DIR}")
A(f"  Fine-tuned master embeddings: {FT_EMB_PATH}")
A(f"  Per-row eval results        : {OUT_RESULTS}")
A(f"  This report                 : {OUT_REPORT}")
A()
A("=" * 66)
A("  END OF REPORT")
A("=" * 66)

report_text = "\n".join(lines)
with open(OUT_REPORT, "w", encoding="utf-8") as f:
    f.write(report_text)
print(f"\nWritten: {OUT_REPORT}", flush=True)
print()
print(report_text)
