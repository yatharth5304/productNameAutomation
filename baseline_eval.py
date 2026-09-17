"""
baseline_eval.py  --  Baseline retrieval evaluation
Run from: f:\Vintyaa\projects\Product Name Automation
"""
import csv, sys, time, unicodedata, re
from pathlib import Path
import numpy as np
import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

PROJECT  = Path(r"f:\Vintyaa\projects\Product Name Automation")
EVAL_CSV = PROJECT / "evaluation_set.csv"
MASTER   = PROJECT / "PRODUCT_MASTER.xlsx"
OUT_BASE = PROJECT / "baseline_evaluation_results.csv"
OUT_META = PROJECT / "evaluation_split_metadata.txt"
RETRIEVAL_K = 10

def _norm(s):
    import unicodedata, re
    s2 = unicodedata.normalize("NFKC", s)
    return re.sub(r"\s+", " ", s2).strip().upper()

print("Loading PRODUCT_MASTER ...", flush=True)
wb = openpyxl.load_workbook(str(MASTER), read_only=True, data_only=True)
ws = wb.active
it = ws.iter_rows(values_only=True)
hdr = {str(h).strip().upper(): i for i, h in enumerate(next(it)) if h is not None}
pc_i, pn_i, bn_i = hdr["PRODUCT_CODE"], hdr["PRODUCT_NAME"], hdr["BRAND_NAME"]
exact_name = {}; normed_name = {}; code_to_rec = {}
consec = 0
for row in it:
    if all(v is None or str(v).strip() == "" for v in row):
        consec += 1
        if consec >= 3: break
        continue
    consec = 0
    code = str(row[pc_i]).strip() if row[pc_i] else ""
    name = str(row[pn_i]).strip() if row[pn_i] else ""
    brand= str(row[bn_i]).strip().upper() if row[bn_i] else ""
    if not code: continue
    rec = {"PRODUCT_CODE": code, "product_name": name, "BRAND_NAME": brand}
    exact_name.setdefault(name, []).append(rec)
    normed_name.setdefault(_norm(name), []).append(rec)
    code_to_rec[code] = rec
wb.close()
print(f"  {len(code_to_rec)} master records.", flush=True)

def resolve_code(name):
    hits = exact_name.get(name, []) or normed_name.get(_norm(name), [])
    return hits[0]["PRODUCT_CODE"] if hits else ""

with open(EVAL_CSV, encoding="utf-8") as f:
    eval_rows = list(csv.DictReader(f))
print(f"Eval rows: {len(eval_rows)}", flush=True)

for r in eval_rows:
    if not r["Positive_Product_Code"]:
        r["Positive_Product_Code"] = resolve_code(r["Positive_Product"])

bad = [r for r in eval_rows if not r["Positive_Product_Code"]]
assert not bad, f"{len(bad)} eval rows have no code"
print("  All positive codes resolved.", flush=True)

print("\nLoading retriever ...", flush=True)
import transformer as t_mod
retriever = t_mod.load_retriever()
print("  Retriever ready.", flush=True)

unique_inputs = list({r["Input_Column"] for r in eval_rows})
print(f"Batch-encoding {len(unique_inputs)} unique eval queries ...", flush=True)
t0 = time.perf_counter()
vecs = t_mod.encode_queries_batch(
    unique_inputs, retriever.model, batch_size=t_mod.ENCODE_BATCH_SIZE
)
input_to_vec = {text: vecs[i].astype(np.float32) for i, text in enumerate(unique_inputs)}
print(f"  Encoded in {time.perf_counter()-t0:.1f}s.", flush=True)

unique_groups = list({(r["Input_Column"], r["Positive_Product"]) for r in eval_rows})
print(f"Running Top-{RETRIEVAL_K} retrieval for {len(unique_groups)} groups ...", flush=True)
t0 = time.perf_counter()
group_retrieval = {}
for ocr_input, pos_product in unique_groups:
    pos_code  = resolve_code(pos_product)
    query_vec = input_to_vec[ocr_input]
    pos_rec   = code_to_rec.get(pos_code, {})
    brand_hint= pos_rec.get("BRAND_NAME", "").strip() or None
    filtered_indices = t_mod.get_brand_filtered_indices(retriever.records, brand_hint)
    candidates = t_mod.retrieve_top_k(
        query_vec, retriever.master_matrix, filtered_indices,
        retriever.records, k=RETRIEVAL_K
    )
    group_retrieval[(ocr_input, pos_product)] = (pos_code, candidates)
print(f"  Done in {time.perf_counter()-t0:.2f}s.", flush=True)

BASE_FIELDS = [
    "Input_Column","Positive_Product","Positive_Product_Code",
    "Positive_Rank","Positive_Score",
    "Top1_Product","Top1_Score","Top10_Contains_Positive","Excel_Row",
]
base_rows = []
for r in eval_rows:
    key = (r["Input_Column"], r["Positive_Product"])
    pos_code, candidates = group_retrieval[key]
    pos_rank = 0; pos_score = ""
    for c in candidates:
        if c["PRODUCT_CODE"] == pos_code:
            pos_rank  = c["rank"]
            pos_score = f"{c['similarity_score']:.4f}"
            break
    top1 = candidates[0] if candidates else {}
    base_rows.append({
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

with open(OUT_BASE, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=BASE_FIELDS)
    w.writeheader(); w.writerows(base_rows)
print(f"\nWritten: {OUT_BASE}", flush=True)

seen = set(); group_metric = []
for br in base_rows:
    key = (br["Input_Column"], br["Positive_Product"])
    if key not in seen:
        seen.add(key); group_metric.append(br)

N = len(group_metric)
def recall_at_k(k):
    return sum(1 for r in group_metric
               if isinstance(r["Positive_Rank"],int) and 0 < r["Positive_Rank"] <= k) / N
def mrr_score():
    return sum(1.0/r["Positive_Rank"] for r in group_metric
               if isinstance(r["Positive_Rank"],int) and r["Positive_Rank"]>0) / N

r1=recall_at_k(1); r3=recall_at_k(3); r5=recall_at_k(5); r10=recall_at_k(10)
mrr_val=mrr_score()
not_in_top10=sum(1 for r in group_metric
                 if not(isinstance(r["Positive_Rank"],int) and r["Positive_Rank"]>0))

from collections import Counter
rank_dist=Counter(r["Positive_Rank"] for r in group_metric)

print("\n" + "="*62)
print("  BASELINE EVALUATION — BAAI/bge-base-en-v1.5 (pretrained)")
print("="*62)
print(f"  Evaluation unique groups   : {N}")
print(f"  Evaluation rows (total)    : {len(eval_rows)}")
print(f"  Positive NOT in Top-10     : {not_in_top10}  ({100*not_in_top10/N:.1f}%)")
print()
print(f"  Recall@1   : {r1:.4f}  ({r1*100:.1f}%)")
print(f"  Recall@3   : {r3:.4f}  ({r3*100:.1f}%)")
print(f"  Recall@5   : {r5:.4f}  ({r5*100:.1f}%)")
print(f"  Recall@10  : {r10:.4f}  ({r10*100:.1f}%)")
print(f"  MRR        : {mrr_val:.4f}")
print()
print("  Rank distribution (per unique group):")
for rank in sorted(rank_dist):
    label = f"Rank {rank}" if rank > 0 else "NOT_FOUND"
    print(f"    {label:15}: {rank_dist[rank]:4d}  ({100*rank_dist[rank]/N:.1f}%)")
print("="*62)

with open(OUT_META,"a",encoding="utf-8") as f:
    f.write("\n")
    f.write("BASELINE RETRIEVAL METRICS (BAAI/bge-base-en-v1.5, pretrained)\n")
    f.write("="*50+"\n")
    f.write(f"Evaluation unique groups   : {N}\n")
    f.write(f"Evaluation rows (total)    : {len(eval_rows)}\n")
    f.write(f"Positive NOT in Top-10     : {not_in_top10}\n")
    f.write(f"Recall@1                   : {r1:.4f}\n")
    f.write(f"Recall@3                   : {r3:.4f}\n")
    f.write(f"Recall@5                   : {r5:.4f}\n")
    f.write(f"Recall@10                  : {r10:.4f}\n")
    f.write(f"MRR                        : {mrr_val:.4f}\n")
    f.write("Rank distribution:\n")
    for rank in sorted(rank_dist):
        label="Rank "+str(rank) if rank>0 else "NOT_FOUND"
        f.write(f"  {label:18}: {rank_dist[rank]}\n")

print(f"Metrics appended to: {OUT_META}", flush=True)
print("Done.", flush=True)
