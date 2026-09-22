"""
V2 Hard-Negative Mining Pipeline
=================================

Mines hard negatives for the V2 positive dataset using the V1 fine-tuned model.
Uses PRODUCT_MASTER.xlsx as the candidate pool.

Output: v2_hard_negative_mining.csv
Report: V2_HARD_NEGATIVE_QUALITY_REPORT.md
"""

import argparse
import csv
import json
import logging
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Set

import numpy as np
import openpyxl
import torch
from sentence_transformers import SentenceTransformer

# ============================================================================
# CONFIG
# ============================================================================

PROJECT_DIR = Path(__file__).resolve().parent

# Inputs
V2_POSITIVE_CSV = str(PROJECT_DIR / "v2_positive_pairs_final.csv")
MASTER_XLSX = str(PROJECT_DIR / "PRODUCT_MASTER.xlsx")
EVALUATION_CSV = str(PROJECT_DIR / "evaluation_set.csv")

# V1 Fine-tuned model
V1_MODEL_DIR = str(PROJECT_DIR / "models" / "bge-pharma-v1")

# Outputs
OUTPUT_CSV = str(PROJECT_DIR / "v2_hard_negative_mining.csv")
REPORT_MD = str(PROJECT_DIR / "V2_HARD_NEGATIVE_QUALITY_REPORT.md")

# Mining parameters
MINING_TOP_K = 10  # Retrieve top-K from master for mining
MAX_MINED = 3      # Keep top-3 hard negatives per positive
ENCODE_BATCH_SIZE = 64
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"

# Column names
COL_INPUT = "PRODUCT_NAME"
COL_POSITIVE = "MATERIAL DESCRIPTION"

# ============================================================================
# LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ============================================================================
# DATA LOADING
# ============================================================================

def load_v2_positives(csv_path: str) -> List[Dict]:
    """Load V2 positive pairs from CSV."""
    logger.info("Loading V2 positive pairs: %s", csv_path)
    rows = []
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append({
                "ocr_input": row[COL_INPUT].strip(),
                "positive": row[COL_POSITIVE].strip(),
            })
    logger.info("  Loaded %d positive pairs", len(rows))
    return rows


def load_master_records(master_path: str) -> List[Dict]:
    """Load PRODUCT_MASTER.xlsx records."""
    logger.info("Loading PRODUCT_MASTER: %s", master_path)
    wb = openpyxl.load_workbook(master_path, read_only=True, data_only=True)
    ws = wb.active

    row_iter = ws.iter_rows(values_only=True)
    try:
        header_row = next(row_iter)
    except StopIteration:
        wb.close()
        raise ValueError("PRODUCT_MASTER is empty")

    col_map = {}
    for idx, h in enumerate(header_row):
        if h is not None:
            col_map[str(h).strip().upper()] = idx

    required = {"PRODUCT_CODE", "PRODUCT_NAME", "BRAND_NAME"}
    missing = required - set(col_map.keys())
    if missing:
        wb.close()
        raise ValueError(f"Missing columns in PRODUCT_MASTER: {missing}")

    pc_idx = col_map["PRODUCT_CODE"]
    pn_idx = col_map["PRODUCT_NAME"]
    bn_idx = col_map["BRAND_NAME"]

    records = []
    for excel_row_num, row in enumerate(row_iter, start=2):
        product_code = row[pc_idx]
        product_name = row[pn_idx]
        brand_name = row[bn_idx]

        if product_name is None or str(product_name).strip() == "":
            continue

        records.append({
            "row_number": excel_row_num,
            "PRODUCT_CODE": str(product_code).strip() if product_code is not None else "",
            "product_name": str(product_name).strip(),
            "BRAND_NAME": str(brand_name).strip().upper() if brand_name is not None else "",
        })

    wb.close()
    logger.info("  Loaded %d master records", len(records))
    return records


def load_evaluation_inputs(eval_path: str) -> Set[str]:
    """Load evaluation Input_Column values to exclude from training."""
    logger.info("Loading evaluation inputs: %s", eval_path)
    eval_inputs = set()
    with open(eval_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            eval_inputs.add(row["Input_Column"].strip())
    logger.info("  Loaded %d evaluation inputs", len(eval_inputs))
    return eval_inputs


def build_name_to_records(records: List[Dict]) -> Dict[str, List[Dict]]:
    """Build mapping from product_name to records."""
    mapping = {}
    for r in records:
        mapping.setdefault(r["product_name"].strip(), []).append(r)
    return mapping


def get_codes_for_name(name: str, name_to_records: Dict[str, List[Dict]]) -> Set[str]:
    """Get PRODUCT_CODEs for a product name."""
    return {r["PRODUCT_CODE"] for r in name_to_records.get(name.strip(), [])}


def extract_brand_from_ocr(ocr_text: str) -> Optional[str]:
    """Extract brand hint from OCR text using first word heuristic."""
    # Simple heuristic: first word is often the brand
    words = ocr_text.strip().split()
    if words:
        return words[0].upper()
    return None


# ============================================================================
# MODEL & EMBEDDING
# ============================================================================

def load_v1_model(model_dir: str) -> SentenceTransformer:
    """Load the V1 fine-tuned model."""
    logger.info("Loading V1 model from: %s", model_dir)
    model = SentenceTransformer(model_dir, device=DEVICE)
    logger.info("  Model loaded on %s", DEVICE)
    return model


def encode_batch(texts: List[str], model: SentenceTransformer, batch_size: int) -> np.ndarray:
    """Encode a batch of texts."""
    matrix = model.encode(
        texts,
        batch_size=batch_size,
        normalize_embeddings=True,
        show_progress_bar=True,
        convert_to_numpy=True,
    )
    return matrix.astype(np.float32)


# ============================================================================
# MINING LOGIC
# ============================================================================

def mine_negatives_for_query(
    query_vec: np.ndarray,
    master_matrix: np.ndarray,
    master_records: List[Dict],
    name_to_records: Dict[str, List[Dict]],
    positive_name: str,
    top_k: int = MINING_TOP_K,
    max_mined: int = MAX_MINED,
) -> List[Dict]:
    """
    Mine hard negatives for a single query.
    
    Returns list of mined negatives with product_name, PRODUCT_CODE, similarity_score.
    """
    # Compute similarities
    scores = master_matrix @ query_vec  # Cosine similarity (dot product on normalized)
    
    # Get top-K indices
    actual_k = min(top_k, len(scores))
    top_indices = np.argsort(scores)[::-1][:actual_k]
    
    # Get positive codes to exclude
    positive_codes = get_codes_for_name(positive_name, name_to_records)
    
    # Filter out positives and collect valid negatives
    mined = []
    for idx in top_indices:
        record = master_records[idx]
        if record["PRODUCT_CODE"] in positive_codes:
            continue
        mined.append({
            "product_name": record["product_name"],
            "PRODUCT_CODE": record["PRODUCT_CODE"],
            "similarity_score": float(scores[idx]),
        })
        if len(mined) >= max_mined:
            break
    
    return mined


# ============================================================================
# MAIN MINING PIPELINE
# ============================================================================

def run_mining(test_only: bool = False, test_n: int = 10):
    """Run the full hard-negative mining pipeline."""
    t_start = time.perf_counter()
    logger.info("=== V2 Hard-Negative Mining Started ===")
    logger.info("Device: %s", DEVICE)
    
    # Load data
    v2_positives = load_v2_positives(V2_POSITIVE_CSV)
    master_records = load_master_records(MASTER_XLSX)
    eval_inputs = load_evaluation_inputs(EVALUATION_CSV)
    
    # Build name-to-records mapping
    name_to_records = build_name_to_records(master_records)
    
    # Filter out evaluation inputs from training
    v2_training = [p for p in v2_positives if p["ocr_input"] not in eval_inputs]
    logger.info("Training positives after eval exclusion: %d", len(v2_training))
    
    if test_only:
        v2_training = v2_training[:test_n]
        logger.info("[TEST MODE] Processing first %d rows", len(v2_training))
    
    # Load V1 model
    model = load_v1_model(V1_MODEL_DIR)
    
    # Prepare master matrix
    master_texts = [r["product_name"] for r in master_records]
    logger.info("Encoding %d master products...", len(master_texts))
    master_matrix = encode_batch(master_texts, model, ENCODE_BATCH_SIZE)
    logger.info("  Master matrix shape: %s", master_matrix.shape)
    
    # Batch encode all queries
    query_texts = [p["ocr_input"] for p in v2_training]
    logger.info("Encoding %d queries...", len(query_texts))
    query_matrix = encode_batch(query_texts, model, ENCODE_BATCH_SIZE)
    logger.info("  Query matrix shape: %s", query_matrix.shape)
    
    # Mining loop
    logger.info("Mining hard negatives (top-%d, keep-%d)...", MINING_TOP_K, MAX_MINED)
    
    output_rows = []
    stats = {
        "total_positives": len(v2_training),
        "full_3_mined": 0,
        "fewer_than_3_mined": 0,
        "zero_mined": 0,
        "positive_in_top10": 0,
        "eval_leakage_attempted": 0,
        "validation_failures": 0,
    }
    
    similarity_scores = {1: [], 2: [], 3: []}
    same_brand_counts = {1: 0, 2: 0, 3: 0}
    
    for idx, pair in enumerate(v2_training):
        ocr_input = pair["ocr_input"]
        positive = pair["positive"]
        query_vec = query_matrix[idx]
        
        # Mine negatives
        mined = mine_negatives_for_query(
            query_vec=query_vec,
            master_matrix=master_matrix,
            master_records=master_records,
            name_to_records=name_to_records,
            positive_name=positive,
            top_k=MINING_TOP_K,
            max_mined=MAX_MINED,
        )
        
        # Check if positive was in top-10 (shouldn't be since we filtered, but verify)
        positive_codes = get_codes_for_name(positive, name_to_records)
        scores = master_matrix @ query_vec
        top10_indices = np.argsort(scores)[::-1][:10]
        positive_in_top10 = any(
            master_records[i]["PRODUCT_CODE"] in positive_codes 
            for i in top10_indices
        )
        if positive_in_top10:
            stats["positive_in_top10"] += 1
        
        # Extract brand for same-brand analysis
        query_brand = extract_brand_from_ocr(ocr_input)
        positive_brand = positive.split()[0].upper() if positive else ""
        
        # Count same-brand negatives
        for slot_idx, neg in enumerate(mined):
            slot = slot_idx + 1
            similarity_scores[slot].append(neg["similarity_score"])
            neg_brand = neg["product_name"].split()[0].upper() if neg["product_name"] else ""
            if query_brand and neg_brand == query_brand:
                same_brand_counts[slot] += 1
            elif positive_brand and neg_brand == positive_brand:
                same_brand_counts[slot] += 1
        
        # Build output row
        row = {
            "Input_Column": ocr_input,
            "Positive_Product": positive,
            "Positive_Product_Code": next(iter(get_codes_for_name(positive, name_to_records)), ""),
        }
        
        # Add mined negatives
        for n in range(1, MAX_MINED + 1):
            if n <= len(mined):
                neg = mined[n-1]
                row[f"Mined_Negative_{n}"] = neg["product_name"]
                row[f"Mined_Negative_{n}_Code"] = neg["PRODUCT_CODE"]
                row[f"Mined_Negative_{n}_Score"] = f"{neg['similarity_score']:.4f}"
            else:
                row[f"Mined_Negative_{n}"] = ""
                row[f"Mined_Negative_{n}_Code"] = ""
                row[f"Mined_Negative_{n}_Score"] = ""
        
        # Update stats
        n_mined = len(mined)
        if n_mined == MAX_MINED:
            stats["full_3_mined"] += 1
        elif n_mined == 0:
            stats["zero_mined"] += 1
        else:
            stats["fewer_than_3_mined"] += 1
        
        output_rows.append(row)
        
        if (idx + 1) % 5000 == 0:
            elapsed = time.perf_counter() - t_start
            logger.info("  Processed %d / %d (%.1fs, %.1f rows/s)",
                        idx + 1, len(v2_training), elapsed, (idx + 1) / elapsed)
    
    # Write output CSV
    if not test_only:
        fieldnames = [
            "Input_Column", "Positive_Product", "Positive_Product_Code",
            "Mined_Negative_1", "Mined_Negative_1_Code", "Mined_Negative_1_Score",
            "Mined_Negative_2", "Mined_Negative_2_Code", "Mined_Negative_2_Score",
            "Mined_Negative_3", "Mined_Negative_3_Code", "Mined_Negative_3_Score",
        ]
        
        with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(output_rows)
        
        logger.info("Written: %s (%d rows)", OUTPUT_CSV, len(output_rows))
    
    # Generate quality report
    generate_quality_report(stats, similarity_scores, same_brand_counts, v2_training, output_rows, master_records, name_to_records, t_start)
    
    logger.info("=== Mining Complete in %.1fs ===", time.perf_counter() - t_start)
    return output_rows, stats


def generate_quality_report(
    stats: Dict,
    similarity_scores: Dict[int, List[float]],
    same_brand_counts: Dict[int, int],
    v2_training: List[Dict],
    output_rows: List[Dict],
    master_records: List[Dict],
    name_to_records: Dict[str, List[Dict]],
    t_start: float,
):
    """Generate detailed mining quality report."""
    
    total = stats["total_positives"]
    
    # Compute similarity stats
    sim_stats = {}
    for slot in [1, 2, 3]:
        scores = similarity_scores[slot]
        if scores:
            sim_stats[slot] = {
                "count": len(scores),
                "min": min(scores),
                "max": max(scores),
                "mean": np.mean(scores),
                "median": np.median(scores),
                "std": np.std(scores),
            }
        else:
            sim_stats[slot] = {"count": 0}
    
    # Validation checks on output
    validation_errors = []
    master_names = {r["product_name"] for r in master_records}
    
    for row in output_rows:
        pos = row["Positive_Product"]
        if pos not in master_names:
            validation_errors.append(f"Positive not in master: {pos}")
        
        for n in [1, 2, 3]:
            neg = row.get(f"Mined_Negative_{n}", "")
            if neg:
                if neg not in master_names:
                    validation_errors.append(f"Negative not in master: {neg}")
                if neg == pos:
                    validation_errors.append(f"Negative equals positive: {neg}")
        
        # Check duplicates within row
        negs = [row.get(f"Mined_Negative_{n}", "") for n in [1, 2, 3] if row.get(f"Mined_Negative_{n}", "")]
        if len(negs) != len(set(negs)):
            validation_errors.append(f"Duplicate negatives in row: {row['Input_Column']}")
    
    # Check evaluation leakage
    eval_inputs = load_evaluation_inputs(EVALUATION_CSV)
    leakage_count = sum(1 for row in output_rows if row["Input_Column"] in eval_inputs)
    if leakage_count > 0:
        validation_errors.append(f"Evaluation leakage: {leakage_count} rows")
    
    # Write report
    with open(REPORT_MD, "w", encoding="utf-8") as f:
        f.write("# V2 HARD-NEGATIVE MINING QUALITY REPORT\n\n")
        f.write(f"Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write("## SUMMARY\n\n")
        f.write(f"- **Total positive pairs processed**: {total:,}\n")
        f.write(f"- **Total negatives generated**: {stats['full_3_mined'] * 3 + stats['fewer_than_3_mined'] * 1 + stats['fewer_than_3_mined'] * 2:,}\n")
        f.write(f"- **Rows with 3 hard negatives**: {stats['full_3_mined']:,} ({stats['full_3_mined']/total*100:.1f}%)\n")
        f.write(f"- **Rows with 1-2 hard negatives**: {stats['fewer_than_3_mined']:,} ({stats['fewer_than_3_mined']/total*100:.1f}%)\n")
        f.write(f"- **Rows with 0 hard negatives**: {stats['zero_mined']:,} ({stats['zero_mined']/total*100:.1f}%)\n")
        f.write(f"- **Positive appeared in top-10**: {stats['positive_in_top10']:,}\n")
        f.write(f"- **Validation errors**: {len(validation_errors):,}\n")
        f.write(f"- **Evaluation leakage**: {leakage_count:,}\n\n")
        
        f.write("## SIMILARITY SCORE DISTRIBUTION\n\n")
        f.write("| Slot | Count | Min | Max | Mean | Median | Std |\n")
        f.write("|------|-------|-----|-----|------|--------|-----|\n")
        for slot in [1, 2, 3]:
            s = sim_stats[slot]
            if s["count"] > 0:
                f.write(f"| {slot} | {s['count']:,} | {s['min']:.4f} | {s['max']:.4f} | {s['mean']:.4f} | {s['median']:.4f} | {s['std']:.4f} |\n")
            else:
                f.write(f"| {slot} | 0 | - | - | - | - | - |\n")
        f.write("\n")
        
        f.write("## SAME-BRAND PERCENTAGE\n\n")
        for slot in [1, 2, 3]:
            count = sim_stats[slot]["count"]
            if count > 0:
                pct = same_brand_counts[slot] / count * 100
                f.write(f"- **Slot {slot}**: {same_brand_counts[slot]:,} / {count:,} = {pct:.1f}%\n")
        f.write("\n")
        
        f.write("## VALIDATION RESULTS\n\n")
        if validation_errors:
            f.write("**FAILURES**:\n")
            for err in validation_errors[:20]:
                f.write(f"- {err}\n")
            if len(validation_errors) > 20:
                f.write(f"- ... and {len(validation_errors) - 20} more\n")
        else:
            f.write("✅ **All validation checks passed**\n\n")
        
        f.write("## REPRESENTATIVE MINING EXAMPLES\n\n")
        # Sample diverse examples
        step = max(1, len(output_rows) // 10)
        samples = [output_rows[i] for i in range(0, len(output_rows), step)][:10]
        
        for s in samples:
            f.write(f"### Input: `{s['Input_Column']}`\n")
            f.write(f"- **Positive**: {s['Positive_Product']} [{s['Positive_Product_Code']}]\n")
            for n in [1, 2, 3]:
                neg = s.get(f"Mined_Negative_{n}", "")
                if neg:
                    f.write(f"- **Mined {n}**: {neg} [{s.get(f'Mined_Negative_{n}_Code', '')}] score={s.get(f'Mined_Negative_{n}_Score', '')}\n")
            f.write("\n")
        
        f.write("## WEAKEST & STRONGEST NEGATIVES\n\n")
        if sim_stats[1]["count"] > 0:
            # Find weakest (lowest score) and strongest (highest score) for slot 1
            all_scores = []
            for row in output_rows:
                if row.get("Mined_Negative_1"):
                    all_scores.append({
                        "input": row["Input_Column"],
                        "positive": row["Positive_Product"],
                        "negative": row["Mined_Negative_1"],
                        "score": float(row["Mined_Negative_1_Score"]),
                    })
            if all_scores:
                all_scores.sort(key=lambda x: x["score"])
                f.write("**Weakest (lowest similarity)**:\n")
                for s in all_scores[:5]:
                    f.write(f"- `{s['input']}` → **{s['positive']}** vs **{s['negative']}** (score={s['score']:.4f})\n")
                f.write("\n**Strongest (highest similarity)**:\n")
                for s in all_scores[-5:]:
                    f.write(f"- `{s['input']}` → **{s['positive']}** vs **{s['negative']}** (score={s['score']:.4f})\n")
        
        f.write(f"\n---\n*Total runtime: {time.perf_counter() - t_start:.1f}s*\n")
    
    logger.info("Quality report written: %s", REPORT_MD)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="V2 Hard-Negative Mining")
    parser.add_argument("--test", action="store_true", help="Run in test mode (first N rows)")
    parser.add_argument("--test-n", type=int, default=10, help="Number of rows in test mode")
    args = parser.parse_args()
    
    run_mining(test_only=args.test, test_n=args.test_n)