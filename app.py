#!/usr/bin/env python3
"""
Single-command production pipeline for pharma product mapping.

Usage:
    python app.py

Input:  mapping.xlsx (sheet with PRODUCT_NAME column)
Output: mapping.xlsx (updated in-place with two sheets:
         - "mapping"      : PRODUCT_NAME, PRODUCT_CODE, MASTER_PRODUCT_NAME
         - "garbage_check": intermediate classification results)

The pipeline replicates the exact decisions of:
  - garbage_check.py (classification: 0/1/MAYBE_PRODUCT)
  - mapping.py       (mapping confirmed products to master codes)
without modifying either original script.
"""

import os
import sys
import shutil
import time
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook, Workbook

# ---------------------------------------------------------------------------
# CONFIGURATION (overrideable via environment variables)
# ---------------------------------------------------------------------------
INPUT_XLSX = Path(os.environ.get("MAPPING_INPUT_XLSX", r"f:\Vintyaa\projects\Product Name Automation\mapping.xlsx"))
MASTER_XLSX = Path(os.environ.get("MASTER_XLSX", r"f:\Vintyaa\projects\Product Name Automation\PRODUCT_MASTER1.xlsx"))
BRANDS_TXT = Path(os.environ.get("BRANDS_TXT", r"f:\Vintyaa\projects\Product Name Automation\Brand Names.txt"))

# Required for mapping.py import (even in local mode)
os.environ.setdefault("NOVITA_API_KEY", "dummy_local_mode")

# Force local mode for both pipelines (no API calls)
os.environ.setdefault("GARBAGE_PROCESSING_MODE", "local")
os.environ.setdefault("MAPPING_PROCESSING_MODE", "local")

# ---------------------------------------------------------------------------
# IMPORT ORIGINAL PIPELINES (after env vars set)
# ---------------------------------------------------------------------------
import garbage_check as gc
import mapping as mp

# Override config constants to match our paths and desired behavior
gc.EXCEL_FILE = str(INPUT_XLSX)
gc.BRANDS_FILE = str(BRANDS_TXT)
gc.PRODUCT_MASTER_FILE = str(MASTER_XLSX)
gc.PROCESSING_MODE = os.environ["GARBAGE_PROCESSING_MODE"]
gc.CLEAR_EXISTING = True
gc.ROW_LIMIT = None

mp.MASTER_XLSX_PATH = str(MASTER_XLSX)
mp.INPUT_OUTPUT_XLSX_PATH = str(INPUT_XLSX)
mp.PROCESSING_MODE = os.environ["MAPPING_PROCESSING_MODE"]
mp.PROCESS_CONFIRMED_ZERO_ONLY = True
mp.PROCESS_MAYBE_PRODUCT_ONLY = False
mp.ROW_LIMIT = None

# ---------------------------------------------------------------------------
# CONSTANTS FOR OUTPUT FORMAT
# ---------------------------------------------------------------------------
SHEET_GARBAGE = "garbage_check"
SHEET_MAPPING = "mapping"

COL_PRODUCT_NAME = "PRODUCT_NAME"
COL_PRODUCT_CODE = "PRODUCT_CODE"
COL_MASTER_PRODUCT = "MASTER_PRODUCT_NAME"

# Garbage check intermediate columns
GC_COLS = [
    "PRODUCT_NAME",
    "Garbage check",
    "Matched brand",
    "Matched text",
    "Match type",
    "Matched garbage phrase",
    "Garbage match type",
]

# Final mapping columns (3-column public interface)
FINAL_COLS = [COL_PRODUCT_NAME, COL_PRODUCT_CODE, COL_MASTER_PRODUCT]

# ---------------------------------------------------------------------------
# UTILITY FUNCTIONS
# ---------------------------------------------------------------------------
def validate_input_file(xlsx_path: Path) -> pd.DataFrame:
    """Validate input file exists and has required structure."""
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Input file not found: {xlsx_path}")

    # Read first sheet
    try:
        df = pd.read_excel(xlsx_path, sheet_name=0, engine="openpyxl")
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file: {e}")

    if df.empty:
        raise ValueError("Input file has no data rows")

    # Find PRODUCT_NAME column (case-insensitive)
    product_col = None
    for col in df.columns:
        if str(col).strip().upper() == "PRODUCT_NAME":
            product_col = col
            break

    if product_col is None:
        raise ValueError(
            f"Required column 'PRODUCT_NAME' not found. "
            f"Available columns: {list(df.columns)}"
        )

    # Return only the product name column, preserving order and duplicates
    return df[[product_col]].rename(columns={product_col: COL_PRODUCT_NAME})


def create_backup(xlsx_path: Path) -> Path:
    """Create a timestamped backup of the input file."""
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    backup_path = xlsx_path.with_name(f"{xlsx_path.stem}_backup_{timestamp}{xlsx_path.suffix}")
    shutil.copy2(xlsx_path, backup_path)
    return backup_path


def build_garbage_check_resources():
    """Load brands, build indices, and load master evidence ONCE."""
    print("[1/4] Loading garbage_check resources...")
    brands = gc.load_brands()
    brand_index = gc.build_brand_index(brands)
    master_evidence = gc.build_master_product_evidence(brand_index)
    garbage_rows = gc.build_garbage_set()
    print(f"      {len(brands)} brands, {len(master_evidence)} master brand groups, {len(garbage_rows)} exact garbage phrases")
    return brands, brand_index, master_evidence, garbage_rows


def classify_single_product(name: str, brand_index, master_evidence, garbage_rows) -> tuple:
    """
    Replicate garbage_check.py main() classification logic for a single product name.
    Returns: (classification, matched_brand, matched_text, match_type, garbage_phrase, garbage_match_type)
    classification: "0" | "1" | "MAYBE_PRODUCT"
    """
    spans = gc.token_spans(name)
    exact_match = gc.exact_product_match(name, brand_index, precomputed_spans=spans)
    fuzzy_match = None if exact_match else gc.fuzzy_product_match(name, brand_index, precomputed_spans=spans)
    master_match = gc.master_product_match(name, master_evidence, precomputed_spans=spans)
    local_match = exact_match or fuzzy_match

    matched_brand_norm = gc.normalize_text(local_match["brand"]) if local_match else ""
    initial_match_brand_norm = matched_brand_norm

    # ---- EXCEPTION LOGIC (copied verbatim from garbage_check.py main()) ----

    # CELOL + DINNER/SET exception
    if local_match and matched_brand_norm == "CELOL" and gc.re.search(r"\b(MARKER|DINNER|SETS?)\b", name.upper()):
        return ("1", "", "", "", "", "")

    # EXACT_ONLY_BRANDS: invalidate fuzzy matches
    if local_match and matched_brand_norm in gc.EXACT_ONLY_BRANDS:
        if local_match.get("match_type") != "exact":
            local_match = None
            matched_brand_norm = ""

    # TAMLET matching TABLET/TABLETS
    if local_match and matched_brand_norm == "TAMLET":
        if gc.normalize_text(local_match.get("matched_text", "")) in ("TABLET", "TABLETS"):
            local_match = None
            matched_brand_norm = ""

    # COMPANY_BRANDS + COMPANY keyword
    _COMPANY_BRANDS = {"IMPETUS", "VINTOR", "EMNU"}
    if local_match and matched_brand_norm in _COMPANY_BRANDS:
        if gc.re.search(r"\bCOMPANY\b", name, gc.re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # AMARYL + SEMI
    if local_match and matched_brand_norm == "AMARYL":
        if gc.re.search(r"\bSEMI\b", name, gc.re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # EMNU + MANUFACTURER
    if local_match and matched_brand_norm == "EMNU":
        if gc.re.search(r"\bMANUFACTURER\b", name, gc.re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # ZUVENTUS without ORS
    if local_match and matched_brand_norm == "ZUVENTUS":
        if not gc.re.search(r"\bORS\b", name, gc.re.IGNORECASE):
            return ("1", "", "", "", "", "")

    # NEW without NORMET
    if local_match and matched_brand_norm == "NEW":
        if not gc.re.search(r"\bNORMET\b", name, gc.re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # VITAMIN without D3
    if local_match and matched_brand_norm == "VITAMIN":
        if not gc.re.search(r"\bD3\b", name, gc.re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # EMCOR without CREAM/TUBE
    if local_match and matched_brand_norm == "EMCOR":
        if not gc.re.search(r"\b(CREAM|TUBE)\b", name, gc.re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # NUMLO -> S-NUMLO remap
    if local_match and matched_brand_norm == "NUMLO":
        if ("S" + matched_brand_norm) in gc.normalize_text(name):
            local_match["brand"] = "S-NUMLO"
            matched_brand_norm = "S-NUMLO"

    # Master match fallback for invalidated NEW
    if not local_match and master_match and initial_match_brand_norm in ("", "NEW"):
        local_match = master_match
        matched_brand_norm = gc.normalize_text(local_match["brand"])

    # ---- FINAL CLASSIFICATION ----
    if local_match and matched_brand_norm not in gc.LOCAL_REVIEW_BRANDS:
        # Confirmed product
        return (
            "0",
            local_match["brand"],
            local_match["matched_text"],
            local_match["match_type"],
            "",
            "",
        )

    # Exact/pattern/metadata garbage
    if (gc.is_exact_garbage_row(name, garbage_rows)
        or gc.is_pattern_garbage_row(name)
        or gc.is_metadata_garbage_row(name)):
        rule_type = "exact"
        if gc.is_exact_garbage_row(name, garbage_rows):
            rule_type = "exact"
        elif gc.is_pattern_garbage_row(name):
            rule_type = "pattern"
        else:
            rule_type = "metadata"
        return ("1", "", "", "", "", "")

    # Fuzzy garbage match
    garbage_match = gc.fuzzy_garbage_match(name, garbage_rows)
    if garbage_match:
        return (
            "1",
            "",
            "",
            "",
            garbage_match["garbage_phrase"],
            garbage_match["match_type"],
        )

    # Local match but brand in LOCAL_REVIEW_BRANDS (currently empty)
    if local_match:
        # In local mode, this falls through to MAYBE_PRODUCT below
        pass

    # MAYBE_PRODUCT: has pharma form/qty signals but no brand match
    if gc.is_maybe_product_row(name):
        return ("MAYBE_PRODUCT", "pattern-inferred", "form+qty match", "MAYBE_PRODUCT", "", "")

    # No brand, no pharma signals -> MAYBE_PRODUCT (local mode fallback)
    return ("MAYBE_PRODUCT", "pattern-inferred", "no-brand-fallback", "MAYBE_PRODUCT", "", "")


def run_garbage_check_phase(product_names: list, brand_index, master_evidence, garbage_rows) -> list:
    """Classify all product names through garbage_check pipeline."""
    print(f"[2/4] Running garbage_check on {len(product_names)} rows...")
    results = []
    counts = {"0": 0, "1": 0, "MAYBE_PRODUCT": 0}

    for i, name in enumerate(product_names):
        if i % 2000 == 0 and i > 0:
            print(f"      Processed {i}/{len(product_names)}...")

        classification, matched_brand, matched_text, match_type, garbage_phrase, garbage_match_type = \
            classify_single_product(name, brand_index, master_evidence, garbage_rows)

        counts[classification] = counts.get(classification, 0) + 1
        results.append({
            "PRODUCT_NAME": name,
            "Garbage check": classification,
            "Matched brand": matched_brand,
            "Matched text": matched_text,
            "Match type": match_type,
            "Matched garbage phrase": garbage_phrase,
            "Garbage match type": garbage_match_type,
        })

    print(f"      Classification: 0={counts['0']}  1={counts['1']}  MAYBE_PRODUCT={counts['MAYBE_PRODUCT']}")
    return results


def run_mapping_phase(input_xlsx: Path, gc_results: list) -> list:
    """Write garbage_check sheet, then call process_excel_file for exact equivalence."""
    # Write garbage_check sheet first (required by process_excel_file)
    print("[3/4] Writing garbage_check sheet...")
    wb = load_workbook(input_xlsx)
    if "garbage_check" in wb.sheetnames:
        wb.remove(wb["garbage_check"])
    ws_gc = wb.create_sheet("garbage_check")
    ws_gc.append(GC_COLS)
    for row in gc_results:
        ws_gc.append([row[c] for c in GC_COLS])
    wb.save(input_xlsx)

    # Call process_excel_file - this reads garbage_check sheet and writes mapping sheet
    print("[4/4] Running mapping via process_excel_file (exact original pipeline)...")
    mp.INPUT_OUTPUT_XLSX_PATH = str(input_xlsx)
    mp.process_excel_file()

    # Read the mapping sheet and convert to final 3-column format
    print("Reading mapping results...")
    df_map = pd.read_excel(input_xlsx, sheet_name="mapping", engine="openpyxl")

    # Build lookup from mapping results
    map_lookup = {}
    for _, r in df_map.iterrows():
        inp = str(r.get("Input_Column", "")).strip()
        code = r.get("Product_Code", "")
        name = r.get("Output_Column", "")
        status = str(r.get("Match_Status", "")).strip()

        # Normalize product code
        if pd.isna(code) or str(code).strip() in ("", "nan", "NaN", "None", "0", "0.0"):
            norm_code = "0"
        else:
            try:
                norm_code = str(int(float(str(code).strip())))
            except:
                norm_code = str(code).strip()

        # Determine master name based on status
        if norm_code == "0" or status in ("NO_CLEAR_MATCH", "VARIANT_NOT_IN_MASTER", "AMBIGUOUS_STRENGTH", "NO_BRAND", "NO_CANDIDATES", "LOCAL_ONLY_UNRESOLVED", "API_ERROR", "ERROR", "SKIPPED_EMPTY"):
            norm_name = "NO SUGGESTION"
        else:
            norm_name = str(name) if not pd.isna(name) else "NO SUGGESTION"

        map_lookup[inp] = (norm_code, norm_name)

    # Build final results preserving original order
    final_results = []
    for row in gc_results:
        name = row["PRODUCT_NAME"]
        classification = row["Garbage check"]

        if classification == "1":
            final_results.append({COL_PRODUCT_NAME: name, COL_PRODUCT_CODE: "1", COL_MASTER_PRODUCT: "GARBAGE"})
        elif classification == "0":
            code, master = map_lookup.get(name, ("0", "NO SUGGESTION"))
            final_results.append({COL_PRODUCT_NAME: name, COL_PRODUCT_CODE: code, COL_MASTER_PRODUCT: master})
        else:  # MAYBE_PRODUCT
            final_results.append({COL_PRODUCT_NAME: name, COL_PRODUCT_CODE: "0", COL_MASTER_PRODUCT: "NO SUGGESTION"})

    return final_results


def write_output_workbook(input_xlsx: Path, gc_results: list, final_results: list):
    """Write both sheets to the output workbook."""
    print("Writing output workbook...")

    wb = Workbook()

    # Sheet 1: garbage_check (intermediate)
    ws_gc = wb.active
    ws_gc.title = SHEET_GARBAGE
    ws_gc.append(GC_COLS)
    for row in gc_results:
        ws_gc.append([row[c] for c in GC_COLS])

    # Sheet 2: mapping (final 3-column interface)
    ws_map = wb.create_sheet(SHEET_MAPPING)
    ws_map.append(FINAL_COLS)
    for row in final_results:
        ws_map.append([row[c] for c in FINAL_COLS])

    wb.save(input_xlsx)
    print(f"      Saved to {input_xlsx}")


def print_summary(final_results: list):
    """Print final statistics."""
    total = len(final_results)
    mapped = sum(1 for r in final_results if r[COL_PRODUCT_CODE] not in ("0", "1"))
    garbage = sum(1 for r in final_results if r[COL_PRODUCT_CODE] == "1")
    no_suggestion = sum(1 for r in final_results if r[COL_PRODUCT_CODE] == "0")

    print("\n" + "=" * 60)
    print("PIPELINE SUMMARY")
    print("=" * 60)
    print(f"Total rows processed:     {total}")
    print(f"  Successfully mapped:    {mapped}")
    print(f"  Garbage (code=1):       {garbage}")
    print(f"  No suggestion (code=0): {no_suggestion}")
    print("=" * 60)


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------------------------
def main():
    start_time = time.time()

    print("=" * 60)
    print("PHARMA PRODUCT MAPPING PIPELINE (app.py)")
    print("=" * 60)
    print(f"Input:  {INPUT_XLSX}")
    print(f"Master: {MASTER_XLSX}")
    print(f"Brands: {BRANDS_TXT}")
    print(f"Mode:   local (no API calls)")
    print()

    # 1. Validate and load input
    print("Validating input file...")
    df_input = validate_input_file(INPUT_XLSX)
    product_names = df_input[COL_PRODUCT_NAME].astype(str).tolist()
    print(f"  Loaded {len(product_names)} product names")

    # 2. Create backup
    backup = create_backup(INPUT_XLSX)
    print(f"  Backup created: {backup.name}")

    # 3. Run garbage_check phase
    brands, brand_index, master_evidence, garbage_rows = build_garbage_check_resources()
    gc_results = run_garbage_check_phase(product_names, brand_index, master_evidence, garbage_rows)

    # 4. Run mapping phase (uses process_excel_file for exact equivalence)
    final_results = run_mapping_phase(INPUT_XLSX, gc_results)

    # 5. Write final 3-column mapping sheet (overwrites the mapping sheet from process_excel_file)
    write_output_workbook(INPUT_XLSX, gc_results, final_results)

    # 6. Summary
    print_summary(final_results)
    elapsed = time.time() - start_time
    print(f"\nCompleted in {elapsed:.1f} seconds")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user")
        sys.exit(130)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)