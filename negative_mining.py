"""
negative_mining.py -- Hard-negative mining for pharmaceutical product mapping
==============================================================================

Responsibility:
    Read training_data.xlsx, use the existing transformer.py retrieval stage
    (precomputed BAAI/bge-base-en-v1.5 embeddings + brand filter) to mine
    hard negatives, and write hard_negative_mining.csv.

Architecture position:
    training_data.xlsx
         |
    negative_mining.py  <-- THIS SCRIPT
         |              uses transformer.py (read-only, no model change)
    hard_negative_mining.csv
         |
    (future fine-tuning stage)

Usage:
    python negative_mining.py               # full run
    python negative_mining.py --test        # first 10 rows only, verbose
    python negative_mining.py --test-n 25  # first N rows, verbose

DO NOT modify transformer.py from here.
DO NOT regenerate master embeddings.
DO NOT alter training_data.xlsx or PRODUCT_MASTER.xlsx.
"""

from __future__ import annotations

import argparse
import csv
import logging
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

PROJECT_DIR: Path = Path(__file__).resolve().parent

TRAINING_XLSX: str = str(PROJECT_DIR / "training_data.xlsx")
OUTPUT_CSV:    str = str(PROJECT_DIR / "hard_negative_mining.csv")

# Candidates to retrieve per query for MINING ONLY.
# Does NOT change the production TOP_K=10 in transformer.py.
MINING_TOP_K: int = 6

# Maximum mined negatives to keep per row (after removing pos + explicit neg)
MAX_MINED: int = 3

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# VALIDATION
# ---------------------------------------------------------------------------

def _validate_environment() -> None:
    """
    Check all required inputs before loading the model.
    Raises RuntimeError on any problem.
    Does NOT silently regenerate stale embeddings.
    """
    errors: List[str] = []

    if not Path(TRAINING_XLSX).exists():
        errors.append(f"training_data.xlsx not found: {TRAINING_XLSX}")

    try:
        import transformer as _t  # noqa: F401
    except ImportError as e:
        errors.append(f"Cannot import transformer.py: {e}")

    if errors:
        raise RuntimeError("Pre-flight validation failed:\n" + "\n".join(f"  {e}" for e in errors))

    try:
        import transformer as t
        t.load_embeddings()
    except (FileNotFoundError, ValueError) as e:
        raise RuntimeError(
            f"Embedding validation failed: {e}\n"
            "Run: python transformer.py --rebuild"
        ) from e


# ---------------------------------------------------------------------------
# TRAINING DATA I/O
# ---------------------------------------------------------------------------

_COL_INPUT    = "INPUT_COLUMN"
_COL_OUTPUT   = "OUTPUT_COLUMN"
_COL_POSITIVE = "OUTPUT SHOULD BE"   # trailing space in header is stripped


def load_training_rows(path: str = TRAINING_XLSX):
    """
    Load training_data.xlsx using iter_rows (never ws.cell in read_only mode).

    Returns (rows: List[Dict], skip_reasons: List[str]).
    Each dict: excel_row, ocr_input, explicit_neg, positive.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"training_data.xlsx not found: {p}")

    logger.info("Loading training data: %s", p)
    wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    ws = wb.active

    row_iter = ws.iter_rows(values_only=True)

    try:
        header_row = next(row_iter)
    except StopIteration:
        wb.close()
        raise ValueError("training_data.xlsx is empty.")

    col_map: Dict[str, int] = {}
    for idx, h in enumerate(header_row):
        if h is not None:
            col_map[str(h).strip().upper()] = idx

    required = {_COL_INPUT, _COL_OUTPUT, _COL_POSITIVE}
    missing  = required - set(col_map.keys())
    if missing:
        wb.close()
        raise ValueError(
            f"Required column(s) missing from training_data.xlsx: {missing}\n"
            f"Actual (stripped) headers: {list(col_map.keys())}"
        )

    ic_idx  = col_map[_COL_INPUT]
    oc_idx  = col_map[_COL_OUTPUT]
    pos_idx = col_map[_COL_POSITIVE]

    rows: List[Dict] = []
    skip_reasons: List[str] = []
    consecutive_empty: int = 0

    for excel_row, row in enumerate(row_iter, start=2):
        if all(v is None or str(v).strip() == "" for v in row):
            consecutive_empty += 1
            if consecutive_empty >= 3:
                break
            continue
        consecutive_empty = 0

        def _clean(v) -> Optional[str]:
            if v is None: return None
            s = str(v).strip()
            return s if s else None

        ocr_input_s    = _clean(row[ic_idx]  if ic_idx  < len(row) else None)
        explicit_neg_s = _clean(row[oc_idx]  if oc_idx  < len(row) else None)
        positive_s     = _clean(row[pos_idx] if pos_idx < len(row) else None)

        if not ocr_input_s:
            skip_reasons.append(f"Row {excel_row}: blank Input_Column")
            continue
        if not positive_s:
            skip_reasons.append(f"Row {excel_row}: blank Output Should be")
            continue
        if not explicit_neg_s:
            skip_reasons.append(f"Row {excel_row}: blank Output_Column")
            continue

        rows.append({
            "excel_row":    excel_row,
            "ocr_input":    ocr_input_s,
            "explicit_neg": explicit_neg_s,
            "positive":     positive_s,
        })

    wb.close()
    logger.info("  Loaded %d valid training rows (%d skipped).", len(rows), len(skip_reasons))
    return rows, skip_reasons


# ---------------------------------------------------------------------------
# MASTER RECORD LOOKUP
# ---------------------------------------------------------------------------

def _build_name_to_records(records: List[Dict]) -> Dict[str, List[Dict]]:
    mapping: Dict[str, List[Dict]] = {}
    for r in records:
        mapping.setdefault(r["product_name"].strip(), []).append(r)
    return mapping


def _codes_for_name(name: str, name_to_records: Dict[str, List[Dict]]) -> set:
    return {r["PRODUCT_CODE"] for r in name_to_records.get(name.strip(), [])}


def _first_code_for_name(name: str, name_to_records: Dict[str, List[Dict]]) -> Optional[str]:
    matches = name_to_records.get(name.strip(), [])
    return matches[0]["PRODUCT_CODE"] if matches else None


# ---------------------------------------------------------------------------
# BRAND HINT EXTRACTION
# ---------------------------------------------------------------------------

def _extract_brand_hint(ocr_text: str) -> Optional[str]:
    """
    Attempt to get a brand from garbage_check.py.
    Falls back to None (full-master search) if unavailable.
    We do NOT re-implement garbage_check logic here.
    """
    try:
        import garbage_check as gc
        if hasattr(gc, "check_garbage"):
            result = gc.check_garbage(ocr_text)
            if isinstance(result, dict):
                brand = result.get("brand") or result.get("Brand") or result.get("BRAND")
                return brand if brand else None
        return None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# MINING LOGIC
# ---------------------------------------------------------------------------

def mine_negatives_for_row(
    ocr_input:       str,
    positive:        str,
    explicit_neg:    str,
    retriever,
    master_records:  List[Dict],
    name_to_records: Dict[str, List[Dict]],
    top_k:           int = MINING_TOP_K,
    max_mined:       int = MAX_MINED,
    brand_hint:      Optional[str] = None,
    query_vec=None,   # precomputed float32 ndarray; if None, encode on-the-fly
) -> Dict:
    import transformer as t
    import numpy as np

    positive_codes     = _codes_for_name(positive,     name_to_records)
    explicit_neg_codes = _codes_for_name(explicit_neg, name_to_records)
    positive_code      = next(iter(positive_codes),     None)
    explicit_neg_code  = next(iter(explicit_neg_codes), None)

    filtered_indices = t.get_brand_filtered_indices(master_records, brand_hint)

    # Use precomputed vector if provided (batch mode); fall back to on-the-fly.
    if query_vec is None:
        query_vec = t.encode_query(ocr_input, retriever.model)

    candidates       = t.retrieve_top_k(
        query_vec, retriever.master_matrix, filtered_indices, master_records, k=top_k
    )

    positive_in_top_k  = any(c["PRODUCT_CODE"] in positive_codes     for c in candidates)
    explicit_in_top_k  = any(c["PRODUCT_CODE"] in explicit_neg_codes for c in candidates)

    excluded_codes  = positive_codes | explicit_neg_codes
    valid_remaining = [c for c in candidates if c["PRODUCT_CODE"] not in excluded_codes]
    mined           = valid_remaining[:max_mined]

    return {
        "candidates_raw":    candidates,
        "positive_in_top_k": positive_in_top_k,
        "explicit_in_top_k": explicit_in_top_k,
        "mined":             mined,
        "positive_code":     positive_code,
        "explicit_neg_code": explicit_neg_code,
    }


# ---------------------------------------------------------------------------
# OUTPUT
# ---------------------------------------------------------------------------

CSV_FIELDNAMES = [
    "Input_Column",
    "Positive_Product",
    "Positive_Product_Code",
    "Explicit_Negative",
    "Explicit_Negative_Code",
    "Mined_Negative_1",
    "Mined_Negative_1_Code",
    "Mined_Negative_1_Score",
    "Mined_Negative_2",
    "Mined_Negative_2_Code",
    "Mined_Negative_2_Score",
    "Mined_Negative_3",
    "Mined_Negative_3_Code",
    "Mined_Negative_3_Score",
    "Positive_In_Top6",
    "Explicit_In_Top6",
    "Mined_Count",
    "Excel_Row",
]


def _build_csv_row(tr: Dict, result: Dict) -> Dict:
    row = {
        "Input_Column":           tr["ocr_input"],
        "Positive_Product":       tr["positive"],
        "Positive_Product_Code":  result["positive_code"] or "",
        "Explicit_Negative":      tr["explicit_neg"],
        "Explicit_Negative_Code": result["explicit_neg_code"] or "",
        "Positive_In_Top6":       result["positive_in_top_k"],
        "Explicit_In_Top6":       result["explicit_in_top_k"],
        "Mined_Count":            len(result["mined"]),
        "Excel_Row":              tr["excel_row"],
    }
    for i in range(MAX_MINED):
        n    = i + 1
        slot = result["mined"][i] if i < len(result["mined"]) else None
        row[f"Mined_Negative_{n}"]       = slot["product_name"]    if slot else ""
        row[f"Mined_Negative_{n}_Code"]  = slot["PRODUCT_CODE"]    if slot else ""
        row[f"Mined_Negative_{n}_Score"] = f'{slot["similarity_score"]:.4f}' if slot else ""
    return row


# ---------------------------------------------------------------------------
# VERBOSE TEST OUTPUT
# ---------------------------------------------------------------------------

def _print_test_row(tr: Dict, result: Dict, row_num: int) -> None:
    cands = result["candidates_raw"]
    mined = result["mined"]
    print(f"\n{'-' * 70}")
    print(f"  Test row {row_num}  (Excel row {tr['excel_row']})")
    print(f"  Input   : {tr['ocr_input']!r}")
    print(f"  Positive: {tr['positive']!r}  [{result['positive_code']}]")
    print(f"  Explicit: {tr['explicit_neg']!r}  [{result['explicit_neg_code']}]")
    print()
    print(f"  TOP-{MINING_TOP_K} candidates:")
    print(f"  {'Rank':<5} {'Score':>7}  {'Code':<14}  product_name")
    print(f"  {'-----':<5} {'-------':>7}  {'-'*14:<14}  {'-'*40}")
    for c in cands:
        tag = ""
        if c["PRODUCT_CODE"] == result["positive_code"]:     tag = " [POSITIVE -- excluded]"
        elif c["PRODUCT_CODE"] == result["explicit_neg_code"]: tag = " [EXPLICIT NEG -- excluded]"
        print(f"  {c['rank']:<5} {c['similarity_score']:>7.4f}  {c['PRODUCT_CODE']:<14}  {c['product_name']}{tag}")
    print()
    print(f"  Positive in Top-6 : {result['positive_in_top_k']}")
    print(f"  Explicit in Top-6 : {result['explicit_in_top_k']}")
    print()
    if mined:
        print(f"  Final mined negatives ({len(mined)}):")
        for i, m in enumerate(mined, 1):
            print(f"    {i}. [{m['PRODUCT_CODE']}]  {m['product_name']}  (score={m['similarity_score']:.4f})")
    else:
        print("  Final mined negatives: NONE")


# ---------------------------------------------------------------------------
# MAIN PROCESSING
# ---------------------------------------------------------------------------

def run(
    test_only:  bool = False,
    test_n:     int  = 10,
    output_csv: str  = OUTPUT_CSV,
) -> None:
    t_start = time.perf_counter()
    logger.info("=== negative_mining.py starting ===")
    _validate_environment()

    import transformer as t_mod
    logger.info("Loading retriever ...")
    retriever       = t_mod.load_retriever()
    master_records  = retriever.records
    name_to_records = _build_name_to_records(master_records)
    logger.info("  %d master records indexed.", len(master_records))

    training_rows, skip_reasons = load_training_rows(TRAINING_XLSX)

    if test_only:
        training_rows = training_rows[:test_n]
        logger.info("[TEST MODE] Processing first %d rows.", len(training_rows))

    stats = {
        "total_xlsx_rows":    len(training_rows) + len(skip_reasons),
        "valid_rows":         0,
        "skipped_rows":       len(skip_reasons),
        "positive_in_top6":   0,
        "explicit_in_top6":   0,
        "full_3_mined":       0,
        "fewer_than_3_mined": 0,
        "zero_mined":         0,
    }

    csv_rows: List[Dict] = []

    # ------------------------------------------------------------------
    # BATCH ENCODE all query texts in one forward pass.
    # This is identical to how the master rebuild works and reduces
    # total encoding from N individual calls (~11s each on CPU) to a
    # single batched call (~2-3 min for 1429 rows total).
    # ------------------------------------------------------------------
    import transformer as t_enc
    import numpy as np

    all_texts = [tr["ocr_input"] for tr in training_rows]
    logger.info("Batch-encoding %d query texts (batch_size=%d) ...",
                len(all_texts), t_enc.ENCODE_BATCH_SIZE)
    t_enc_start = time.perf_counter()
    all_query_vecs = t_enc.encode_queries_batch(
        all_texts, retriever.model, batch_size=t_enc.ENCODE_BATCH_SIZE
    )
    logger.info("  Encoded %d queries in %.1fs.",
                len(all_texts), time.perf_counter() - t_enc_start)

    for idx, tr in enumerate(training_rows):
        brand_hint = _extract_brand_hint(tr["ocr_input"])
        query_vec  = all_query_vecs[idx].astype(np.float32)

        result = mine_negatives_for_row(
            ocr_input       = tr["ocr_input"],
            positive        = tr["positive"],
            explicit_neg    = tr["explicit_neg"],
            retriever       = retriever,
            master_records  = master_records,
            name_to_records = name_to_records,
            top_k           = MINING_TOP_K,
            max_mined       = MAX_MINED,
            brand_hint      = brand_hint,
            query_vec       = query_vec,
        )

        if test_only:
            _print_test_row(tr, result, idx + 1)

        stats["valid_rows"] += 1
        if result["positive_in_top_k"]:  stats["positive_in_top6"]  += 1
        if result["explicit_in_top_k"]:  stats["explicit_in_top6"]  += 1
        n_mined = len(result["mined"])
        if   n_mined == MAX_MINED: stats["full_3_mined"] += 1
        elif n_mined == 0:         stats["zero_mined"]   += 1
        else:                      stats["fewer_than_3_mined"] += 1

        csv_rows.append(_build_csv_row(tr, result))

        if (idx + 1) % 100 == 0:
            elapsed = time.perf_counter() - t_start
            logger.info("  Processed %d / %d  (%.1fs, %.1f rows/s)",
                        idx + 1, len(training_rows), elapsed, (idx + 1) / elapsed)

    if not test_only:
        out_path = Path(output_csv)
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES)
            writer.writeheader()
            writer.writerows(csv_rows)
        logger.info("Written: %s  (%d rows)", out_path, len(csv_rows))
    else:
        logger.info("[TEST MODE] CSV not written.")

    elapsed_total = time.perf_counter() - t_start
    print("\n" + "=" * 70)
    print("  NEGATIVE MINING REPORT")
    print("=" * 70)
    print(f"  Total rows in training_data.xlsx : {stats['total_xlsx_rows']}")
    print(f"  Valid rows processed             : {stats['valid_rows']}")
    print(f"  Skipped rows                     : {stats['skipped_rows']}")
    if skip_reasons:
        for r in skip_reasons[:10]:
            print(f"    {r}")
        if len(skip_reasons) > 10:
            print(f"    ... and {len(skip_reasons)-10} more")
    print(f"  Positive appeared in Top-6       : {stats['positive_in_top6']}")
    print(f"  Explicit neg appeared in Top-6   : {stats['explicit_in_top6']}")
    print(f"  Rows with exactly 3 mined negs   : {stats['full_3_mined']}")
    print(f"  Rows with 1-2 mined negs         : {stats['fewer_than_3_mined']}")
    print(f"  Rows with 0 mined negs           : {stats['zero_mined']}")
    print(f"  Total runtime                    : {elapsed_total:.1f}s")
    if stats["valid_rows"]:
        print(f"  Avg per row                      : {elapsed_total/stats['valid_rows']*1000:.1f} ms")
    print("=" * 70)

    if not test_only and csv_rows:
        print("\n  10 REPRESENTATIVE EXAMPLES:")
        step = max(1, len(csv_rows) // 10)
        samples = [csv_rows[i] for i in range(0, len(csv_rows), step)][:10]
        for s in samples:
            print(f"\n  Input   : {s['Input_Column']!r}")
            print(f"  Positive: {s['Positive_Product']!r}  [{s['Positive_Product_Code']}]")
            print(f"  Explicit: {s['Explicit_Negative']!r}  [{s['Explicit_Negative_Code']}]")
            for n in range(1, MAX_MINED + 1):
                name  = s.get(f"Mined_Negative_{n}", "")
                code  = s.get(f"Mined_Negative_{n}_Code", "")
                score = s.get(f"Mined_Negative_{n}_Score", "")
                if name:
                    print(f"  Mined {n} : {name!r}  [{code}]  score={score}")
                else:
                    print(f"  Mined {n} : (none)")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="negative_mining.py",
        description=(
            "Hard-negative mining for pharmaceutical product mapping.\n\n"
            "Examples:\n"
            "  python negative_mining.py --test\n"
            "  python negative_mining.py --test-n 25\n"
            "  python negative_mining.py\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--test",   action="store_true",
                   help="Process only the first --test-n rows (verbose).")
    p.add_argument("--test-n", type=int, default=10, metavar="N",
                   help="Number of rows in test mode (default: 10).")
    p.add_argument("--output", type=str, default=OUTPUT_CSV, metavar="PATH",
                   help=f"Output CSV path.")
    return p


def main(argv=None) -> int:
    parser = _build_parser()
    args   = parser.parse_args(argv)
    try:
        run(test_only=args.test, test_n=args.test_n, output_csv=args.output)
        return 0
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        logger.error("%s", exc)
        return 1
    except KeyboardInterrupt:
        print("\nInterrupted.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
