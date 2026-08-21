# Project Context & Session History

## Core Operating Principles & Constraints

1. **Accuracy is the Absolute Highest Priority**:
   - No performance optimization, refactoring, or code change may alter classification decisions, matching semantics, brand-specific exceptions, `MAYBE_PRODUCT` logic, OCR handling, edge-case behavior, or output values.
   - Every optimization must be strictly zero-risk and verifiable against original behavior.
2. **Minimal & Surgical Changes**:
   - Avoid broad refactoring, structural redesign, or cosmetic cleanup in unrelated parts of the codebase.
   - Changes must be targeted only to the specific requirement or bottleneck being addressed.
3. **Mandatory Syntax & Compile Checks**:
   - Any code edit to `.py` files must immediately be validated using a Python compile/syntax check (`ast.parse` / `py_compile`) to ensure zero syntax errors or indentation defects before running.

---

## Component Overview

- **`garbage_check.py`**:
  - Initial filtering and classification pipeline for raw OCR product strings.
  - Classifies rows into:
    - `0`: Confirmed genuine product row (local brand match).
    - `1`: Confirmed structural / metadata / non-pharma garbage row.
    - `MAYBE_PRODUCT`: Uncertain row containing dosage form / quantity / strength patterns that local brand matching did not confirm.
    - `REVIEW` / LLM fallback: Ambiguous rows requiring further inspection.
- **`mapping.py`**:
  - Downstream product mapping pipeline matching classified product rows against the master database (`PRODUCT_MASTER.xlsx`).
  - Resolves brand, variant, sub-variant, dosage form, pack size, and product codes.

---

## Detailed Session History & Technical Decisions

### 1. `garbage_check.py`

#### A. Accuracy & Matching Semantics
- **Brand-Specific Exceptions**:
  - **`NUMLO` Exception**: When the matched brand text (normalized) starts with `S` (i.e. raw input span was `SNUMLO`), `matched_brand_norm` is remapped to `S-NUMLO` so downstream logic treats it as `S-NUMLO`. Plain `NUMLO` (without `S` prefix) remains unchanged as `NUMLO`.
  - **`EXACT_ONLY_BRANDS`**: Brand matches for `{"EFCURE", "KEMCURE", "CTAX", "NUNIT", "EMNU", "ACEM"}` must be exact span matches only. Any fuzzy match on these brands is invalidated to prevent false positives from OCR noise.
  - **Other Active Exceptions**:
    - `CELOL` matching rows with `DINNER` or `SET(S)` are classified directly as `1` (non-pharma garbage).
    - `TAMLET` matching `TABLET` / `TABLETS` dosage words is invalidated.
    - `IMPETUS` / `VINTOR` / `EMNU` containing `COMPANY` are invalidated as company header lines.
    - `AMARYL` containing `SEMI` is invalidated (orphan sub-brand not in master).
    - `EMNU` containing `MANUFACTURER` is invalidated.
    - `ZUVENTUS` without `ORS` is classified directly as `1`.
    - `NEW` without `NORMIT` is invalidated.
    - `VITAMIN` without `D3` is invalidated.
    - `EMCOR` without `CREAM` or `TUBE` is invalidated.

- **`brand_norm_set` / `brand_norm_to_brand` Regression & Restoration**:
  - *Attempted Optimization*: An earlier optimization attempted to look up normalized tokens in a dictionary/set (`brand_norm_set` / `brand_norm_to_brand`) for exact matching.
  - *Regression Identified*: Caused correctness issues by altering token boundary semantics and span matching fidelity compared to original brand definitions.
  - *Resolution & Semantics Restored*: Restored strict original raw-span vs. original-brand exact-match semantics using:
    ```python
    original_brand_set = {brand for brand, _ in brand_index}
    ```
    and checking:
    ```python
    span["text"] in original_brand_set
    ```
    ensuring 100% fidelity to the original matching algorithm without semantic drift.

#### B. Performance Optimizations & Caching
- **OCR-Line Single-Normalization**:
  - Normalized OCR input row tokens/text once per row rather than repeatedly executing `normalize_text()` across multiple sub-checks.
- **`token_spans()` & `row_tokens(name)` Caching (OPT-1)**:
  - `token_spans(name)` constructs all consecutive token sub-spans ($O(T^2)$ combinations). Previously, it was called twice per non-exact row: once in `exact_product_match()` and again in `fuzzy_product_match()`.
  - Updated both `exact_product_match(name, brand_index, precomputed_spans=None)` and `fuzzy_product_match(name, brand_index, precomputed_spans=None)` to accept precomputed spans.
  - In `main()`, `spans = token_spans(name)` is computed once per row and passed to both functions via `precomputed_spans=spans`. When `precomputed_spans is None`, each function falls back to computing spans internally for backward compatibility.
- **Full-Script Execution-Time Measurement**:
  - Wrapped entry point execution with `time.time()` to record start and end timestamps, printing total elapsed seconds upon completion.

#### C. Startup Bottleneck Investigation & Startup Fixes
- **Problem**: Running `python garbage_check.py` resulted in an apparent freeze of 4–5+ minutes immediately following the console output:
  ```text
  875 brands loaded
  ```
- **Investigation Findings**:
  - `load_workbook("test.xlsx")` took ~11.3s to parse XML data for 17,654 rows (123k+ cells) in memory.
  - The `CLEAR_EXISTING` loop executed 105,924 individual cell lookups (`ws.cell(row=row, column=column).value = None`) with no intermediate logging.
  - Console stdout buffering on Windows PowerShell added to perceived inactivity before per-row prints began.
- **Implemented Startup Fixes**:
  1. **Progress Indicators**:
     - Added `print("Loading workbook...")` before `load_workbook()`.
     - Added `print("Workbook loaded.")` after loading completes.
     - Added `print(f"Clearing {ws.max_row - 1} rows...")` before clearing.
     - Added `print("Cleared previous results outside column A.")` after clearing.
     - Added `print(f"{len(pending)} rows to process. Starting classification...")` before the row loop.
  2. **`CLEAR_EXISTING` Speed Optimization**:
     - Replaced nested `ws.cell()` coordinate lookups with `ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=ws.max_column)` to iterate pre-fetched `Cell` objects directly, speeding up the clearing phase significantly.

#### D. Code Cleanup & Indentation / NameError Repairs
- Repaired duplicate function definitions, indentation mismatches, and potential `NameError` occurrences caused by earlier tool-assisted multi-replacements.
- Verified syntax with `ast.parse` and confirmed clean execution.

---

### 2. `mapping.py`

#### A. Optimizations & Structural Findings
- **Cleaned Input Single-Normalization**:
  - Single normalization of cleaned input text for exact-match brand lookups to avoid re-normalizing per brand candidate.
- **Auto-Save Frequency Adjustment**:
  - Auto-save frequency in `_save_to_mapping_sheet` adjusted from every 10 rows to every 50 rows (`processed_count % 50 == 0`) to reduce repeated Excel file rewriting overhead during long batch runs.
- **Performance Profiling & Bottleneck Analysis**:
  - **`DELAY_BETWEEN_PRODUCTS = 2.0`**: Identified that `time.sleep(DELAY_BETWEEN_PRODUCTS)` accounts for 2 seconds per row (over 9 hours across 17k rows). When the LLM reranker block is commented out/bypassed, this delay represents the dominant time sink.
  - **`_variant_subvariant_tokens(brand_map)`**: Currently called inside `desegment_ocr_input()` on every row, scanning the entire master catalog repeatedly. Identified as a high-impact zero-risk candidate for pre-computation at startup.
  - **`_pack_form_skip_tokens()`**: Currently rebuilds a `set` union from scratch on each invocation inside brand grouping loops. Identified as a zero-risk candidate for conversion to a module-level constant.

---

## Current Status & Next Actions

- **`garbage_check.py`**:
  - Clean syntax verified (`SYNTAX OK`).
  - Startup progress logging active.
  - `CLEAR_EXISTING` optimized with `ws.iter_rows()`.
  - `token_spans()` caching active and verified.
  - All brand exceptions (`NUMLO`, `ACEM`, `EMNU`, `EXACT_ONLY_BRANDS`, etc.) intact and active.
- **`mapping.py`**:
  - Auto-save updated to every 50 rows.
  - Performance analysis documented with ranked zero-risk opportunities.
- **Pending/Recommended Zero-Risk Optimizations (Ready for future implementation if requested)**:
  1. *`garbage_check.py`*: Pre-normalize `EXACT_GARBAGE_ROWS` once at module load to avoid repeated `normalize_text()` inside `fuzzy_garbage_match()`.
  2. *`mapping.py`*: Pre-compute `_variant_subvariant_tokens(brand_map)` once at startup rather than per-row.
  3. *`mapping.py`*: Replace `_pack_form_skip_tokens()` with a module-level constant (`frozenset`).