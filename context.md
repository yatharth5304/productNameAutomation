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
  - **Reranker-aware delay**: The live LLM reranker remains enabled. `DELAY_BETWEEN_LLM_REQUESTS = 2.0` now runs only after a row actually attempts the reranker, tracked by `LLM_REQUEST_ATTEMPTS`; locally resolved rows no longer incur a 2-second delay. This preserves throttling after both successful and failed external requests without changing request, matching, scoring, retry, or output logic.
  - **`_variant_subvariant_tokens(brand_map)`**: Currently called inside `desegment_ocr_input()` on every row, scanning the entire master catalog repeatedly. Identified as a high-impact zero-risk candidate for pre-computation at startup.
  - **`_pack_form_skip_tokens()`**: Currently rebuilds a `set` union from scratch on each invocation inside brand grouping loops. Identified as a zero-risk candidate for conversion to a module-level constant.

#### B. Accuracy Fixes (Implemented)
- **Fix 2 - narrowed the "glued sub-brand" test in `filter_items_prefer_exact_brand_token` (Step 1.65, `mapping.py:3262-3263`)**:
  - **Change (one line added, nothing else touched)**: an item is now treated as a glued sub-brand extension only when the remainder after the brand prefix is *also* absent from the input tokens:
    ```python
    if (first != brand_first and first.startswith(brand_first)
            and len(first) > len(brand_first) and first not in input_tokens
            and first[len(brand_first):] not in input_tokens):
    ```
  - **Rationale**: the filter existed to stop input `ESLO 2.5` from picking `ESLOMET - 2.5`. Because master product names spell sub-brands glued (`CARDACEH10TABLET10x15T`, `OROFERS100INJECTION`) while the OCR input spells them separated (`CARDACE H 2.5`, `OROFER S 100`), the original test also deleted the *correct* SKUs whenever a standalone sub-brand letter was present, leaving those rows with no candidate and a `VARIANT_NOT_IN_MASTER` blank. The remainder check distinguishes "the input never asked for this extension" (`MET` absent -> still dropped) from "the input did ask for it, just spaced" (`H`, `S` present -> now kept).
  - **Original purpose preserved**: `ESLOMET` is still dropped when the input contains only `ESLO`; `CARDACEH` is still dropped when no standalone `H` appears; single-item lists and lists with no clean exact-brand item are still returned untouched.
  - **Verified impact (1450-row Sheet2 replay, live reranker disabled, original brand hints)**:
    - `MATCHED` 1171 -> 1217, `VARIANT_NOT_IN_MASTER` 166 -> 110, `RECOVERED_API_ERROR` 44 -> 50, `NO_CLEAR_MATCH` 235 -> 183.
    - 52 rows newly mapped, **0 mappings lost**, 3 previously wrong product codes corrected:
      - rows 37 / 437 `OROFER XT SYP 1X200ML`: `424440118 OROFER XT-DHA 7 X 1 KIT` -> `424441225 OROFERXT+SUSPENSION200ML`
      - row 565 `XILIA TRIO 1 TAB 10X10T`: `421111093 XILIA TRIO - 2 TABLETS` -> `421111092 XILIATRIO-1TABLETS`
    - Of the 52 newly mapped rows, 46 are correct. The 6 exceptions are `CARDACE H 5MG` rows that land on `CARDACEH2.5` because `repair_master_variant_fields` leaves `CARDACEH5TABLET10x15T` as VARIANT=`H5` / SUB_VARIANT=`PLAIN` (the separate single-digit-strength split gap). Those rows were blank before, so this is blank -> wrong-strength, never correct -> wrong.
  - **Verified impact (full 12,971-row Source="0" corpus, A/B in one process against the verbatim pre-fix function body)**:
    - `MATCHED` 10935 -> 11310, `VARIANT_NOT_IN_MASTER` 1193 -> 764; rows carrying a product code 11294 -> 11711.
    - 417 rows newly mapped, **0 mappings lost, 0 exceptions**, 48 product codes switched: 33 x `OROFER XT-DHA 7 X 1 KIT` -> `OROFERXT+SUSPENSION200ML`, 13 x `XILIA TRIO - 2 TABLETS` -> `XILIATRIO-1TABLETS`, 1 x `CARDACE2.5MGTABLET10x15T` -> `CARDACEPROTECT2.5TABLET10x10T` (row 7863 `CARDACE 2.5/PROTECT`, a correction), 1 x `CARDACE5MGTABLET10x15T` -> `CARDACEPROTECT2.5TABLET10x10T` (row 7864 `CARDACE 5/PROTECT`, wrong before and after - the `PROTECT` variant is now honoured but the strength should be `421112903`).
    - Of the 417 new mappings, 333 have a strength printed in the input: **287 correct, 46 wrong**. All 46 are `CARDACE H 5MG` rows landing on `CARDACEH2.5` or `CARDACEH10`. The remaining 84 print no strength (or no `100`/`200` for `OROFER S`) and cannot be adjudicated from the input text alone.
  - **Known limitation exposed, not caused, by this fix**: `repair_master_variant_fields` does not split single-digit strengths, so `CARDACEH5TABLET10x15T`, `CARDACEAM5TABLET10x15T`, `CARDACEMETO5TABLET10x10T` and `CARDACEPROTECT5TABLET10x10T` remain VARIANT=`H5`/`AM5`/`METO5`/`PROTECT5` with SUB_VARIANT=`PLAIN`. A detected sub-variant of `5` therefore matches no sibling and ranking falls through to the `2.5` or `10` SKU. These rows were blank before Fix 2, so the transition is blank -> wrong-strength, never correct -> wrong; a sibling-aware single-digit split is the fix that would resolve them.
  - **Validation performed**: 7/7 targeted glued-filter unit cases (including the `ESLOMET` guard); full-file diff shows exactly one hunk with line endings unchanged; `ast.parse` + `py_compile` clean; the on-disk 1450-row run is bit-identical to the pre-measured counterfactual sweep (0 of 1450 rows differ); full-corpus A/B confirms 0 lost mappings outside Sheet2.

- **Fix 1 - forward `forced_brand` / `auto_detect` into the Step 1.7 lower-variant retry (`mapping.py:3439-3440`)**:
  - **Change (two arguments added to the existing recursive call, retry algorithm untouched)**:
    ```python
    res = process_product(stripped, brand_map, all_variants, all_sub_variants,
                          client, verbose=verbose, _lower_variant_retry=True,
                          forced_brand=forced_brand, auto_detect=auto_detect)
    ```
  - **Rationale**: Step 1.7 correctly detects an orphan qualifier the brand does not stock and correctly decides to retry without it, but the recursive call dropped the brand hint. The child invocation therefore entered Step 1's no-hint branch and returned `status="NO_BRAND"`, which the parent propagated via `return res` even though the parent had already resolved the brand. The three `status="NO_BRAND"` returns in the file all sit inside Step 1 and are unreachable once a hint resolves in `brand_map`, so this recursion was the sole source of every `NO_BRAND` row and the recovery path was structurally dead for hinted input. Forwarding the two arguments restores the documented behaviour; `_lower_variant_retry=True` still blocks a second recursion and the existing `if res["status"] == "MATCHED"` gate still governs acceptance and the `RECOVERED_LOWER_VARIANT` / `LOW` relabelling.
  - **Verified impact (1450-row Sheet2 replay)**:
    - `NO_BRAND` **66 -> 0**; `RECOVERED_LOWER_VARIANT` 0 -> 65, `RECOVERED_API_ERROR` 50 -> 51; `MATCHED` unchanged at 1217, `VARIANT_NOT_IN_MASTER` unchanged at 110.
    - 66 rows newly mapped, **0 lost, 0 switched**. All 66 carry `confidence="LOW"`. Bit-identical to the pre-measured counterfactual (0 of 1450 rows differ).
    - All 1217 previously `MATCHED` rows are unchanged in both product code and status, and the `VARIANT_NOT_IN_MASTER` set is identical row-for-row (no row entered or left the refusal set), so the documented correct rejections - AMARYL `MP`, CETAPIN `S` / `P` - still refuse.
  - **Verified impact (full 12,971-row Source="0" corpus, A/B in one process against the pre-Fix1 call semantics)**:
    - `NO_BRAND` **365 -> 0**; `RECOVERED_LOWER_VARIANT` 0 -> 347, `RECOVERED_API_ERROR` 401 -> 416, `AMBIGUOUS_STRENGTH` 132 -> 135; `MATCHED` unchanged at 11310, `VARIANT_NOT_IN_MASTER` unchanged at 764.
    - 362 rows newly mapped (all `confidence="LOW"`), 3 rows moved `NO_BRAND` -> `AMBIGUOUS_STRENGTH` (still a blank refusal), **0 mappings lost, 0 codes switched, 0 exceptions**.
    - All 11310 previously `MATCHED` rows are byte-identical in code and status; the `VARIANT_NOT_IN_MASTER` set of 764 rows is identical row-for-row.
    - Largest recovered groups: VYLDA family (`VYLD A` split by a spurious OCR space) ~144 rows, `POVIZTRA FLEXTOUCH` -> `POVIZTRA FLXT` 44, `PROCTOSEDYL BD CREAM` 18, `CORDAR ONE-X` 5. The known-imperfect groups are `HOSIT INJECTION` 22 and `OROFER FCM 750MG/15ML` 17.
  - **`auto_detect` coverage**: `process_excel_file` concatenates all Source="0" rows ahead of the `MAYBE_PRODUCT` rows and `ROW_LIMIT = 19888` truncates before them, so no corpus row reaches Step 1.7 with `auto_detect=True`. Exercised directly instead: with `forced_brand=None, auto_detect=True`, five orphan-qualifier inputs all returned `NO_BRAND` before the fix; afterwards four recover a product at `LOW` confidence and `CORDAR ONE-X 15TAB` still returns `NO_BRAND` (auto-detect cannot resolve that split brand - unchanged, not a regression).
  - **Known imperfect outcomes left for later fixes** (all blank -> wrong, never correct -> wrong, and all emitted at `LOW` confidence under a distinct status so they stay filterable): `HOSIT FCM INJ` rows down-map to `HOSIT INJECTION` (421110085) because the real SKUs sit under `BRAND_NAME = "HOSIT FCM"` (brand-key fragmentation); `OROFER FCM I` rows pick 750MG/15ML (421112705) over 500MG/10ML (421111114) because 421112705 carries `SUB_VARIANT = PLAIN` instead of its strength; `CARDAC E METO-5MG` picks `METO2.5` (single-digit split gap); `EMSITA TRIO F` picks `EMSITA TRIO` rather than `EMSITA TRIO FORTE` (421112883).
  - **Validation performed**: recursive call inspected before and after; `ast.parse` + `py_compile` clean; 1450-row replay identical to prediction with `NO_BRAND == 0`, `lost == 0`, `switched == 0`; full-corpus A/B confirming 0 lost, 0 switched, 0 exceptions and an unchanged `MATCHED` / `VARIANT_NOT_IN_MASTER` population; synthetic `auto_detect=True` probes.

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
  - Local-only rows no longer sleep; the existing 2-second API throttle is retained only after an LLM reranker attempt.
  - Performance analysis documented with ranked zero-risk opportunities.
  - **Fix 2 implemented and verified**: the Step 1.65 glued-sub-brand test now also requires the post-brand remainder to be absent from the input (`mapping.py:3262-3263`). Syntax verified (`ast.parse` + `py_compile` OK). +52 newly mapped Sheet2 rows, 0 lost, 3 wrong codes corrected; `ESLOMET`-style suppression intact.
  - **Fix 1 implemented and verified**: the Step 1.7 lower-variant retry now forwards `forced_brand` / `auto_detect` (`mapping.py:3439-3440`). Syntax verified (`ast.parse` + `py_compile` OK). `NO_BRAND` 66 -> 0 on Sheet2, +66 newly mapped at `LOW` confidence, 0 lost, 0 switched, refusal set unchanged.
- **Recommended Next Accuracy Fixes (analysed, not implemented)**, in priority order:
  1. **Brand-key fragmentation** - master `BRAND_NAME` holds multi-word keys (`OROFER-XT TAB`, `OROFER-XT DROP`, `FERIUM XT`, `HOSIT FCM`, `ATAZOR-R`) while `garbage_check.py` emits the short root brand, so those SKUs are unreachable. Largest remaining class (~61 Sheet2 rows); needs its own regression sweep because it widens candidate generation.
  2. **Sibling-aware single-digit variant split** in `repair_master_variant_fields` - would fix the 46 `CARDACE H 5MG` rows plus `CARDACE 5/PROTECT`, `CARDAC E METO-5MG`, and the `AM5` cases. Must stay sibling/brand-aware: a blanket relaxation would wrongly split `D3`, `CD3`, `TH4`.
  3. **Master data repair** - `OROFER FCM INJECTION 750MG/15ML` (421112705) carries `SUB_VARIANT = PLAIN` instead of its strength, which is why strength-less `OROFER FCM` input lands on it instead of 500MG/10ML.
- **Pending/Recommended Zero-Risk Optimizations (Ready for future implementation if requested)**:
  1. *`garbage_check.py`*: Pre-normalize `EXACT_GARBAGE_ROWS` once at module load to avoid repeated `normalize_text()` inside `fuzzy_garbage_match()`.
  2. *`mapping.py`*: Pre-compute `_variant_subvariant_tokens(brand_map)` once at startup rather than per-row.
  3. *`mapping.py`*: Replace `_pack_form_skip_tokens()` with a module-level constant (`frozenset`).

## Brand Vocabulary Synchronization (`Brand Names.txt` <- `PRODUCT_MASTER1.xlsx`)

`PRODUCT_MASTER1.xlsx` (2,765 rows, built from `Product.xlsx`) is now the source of
truth for the brand vocabulary. `Brand Names.txt` was resynchronized against its
`BRAND_NAME` column; `garbage_check.py` logic was **not** touched.

- **Before**: 875 quoted entries, 0 exact duplicates, 1 normalization collision
  (`KETOPLAST` / `KETOPLAST+` both normalize to `KETOPLAST`, so the second entry was
  unreachable). Single line, `", "`-separated, no trailing newline.
- **After**: 849 entries == exactly the 849 distinct `BRAND_NAME` values, 0 duplicates,
  0 normalization collisions. Format, ordering convention (plain lexicographic, with the
  pre-existing `SKIZOTUS` / `S-NUMLO` anomaly kept) and the 20 multi-word entries
  preserved. `load_brands()` (`re.findall(r'"([^"]+)"', ...)`) parses it unchanged.
- **Removed 28** entries that are not `PRODUCT_MASTER1` brand keys. Most are brand+variant
  fragments whose base brand *is* a key (`OROFER-XT TAB`, `OROFER-XT DROP`, `FERIUM XT`,
  `FERIUM 1K/D3/INJ`, `CALONAT D3`, `CALONAT XT`, `HOSIT FCM`, `ENCICARB 1K/500/750`,
  `ATAZOR-R`, `EMANZEN TBR`, `MOFILET S`, `TAVIN EM`, `DOCEZAP-20/80/120`, `SPASMOTER-4/-M`,
  `NEXDOL-SR`, `TUSIMOX-DS`, `VOWFLAM-TZ`, `HU-OMEGA-3`, `KETOPLAST+`, `LEMCUREZYM`) -
  27 of the 28 fall in this class. The remaining one, `CIDOFOVIR`, matches no
  `PRODUCT_MASTER1` product name at all and is a pure orphan.
- **Added 2**: `CIDNAVIR`, `FERROUS` - the two `PRODUCT_MASTER1` fallback brands that had
  no vocabulary entry, inserted at their alphabetical positions. Zero corpus effect
  (no input matches them), but every master brand group is now reachable by a hint.
- **This closes recommended next fix #1 (brand-key fragmentation)** from the section above:
  the fragmentation was resolved on the master side by `PRODUCT_MASTER1`, and the
  vocabulary now agrees with it.

**Measured effect** (12,972-row corpus, `garbage_check.exact_product_match` /
`fuzzy_product_match` re-run with both vocabularies):

- 23 of 12,972 rows change their brand hint - and they are exactly the 23 rows that
  previously produced a hint with no brand group in `PRODUCT_MASTER1` (`NO_BRAND`).
  Rows carrying an unresolvable hint: **23 -> 0**. No currently-mapping row is affected.
- 12 rows retarget to a live base brand and **all 12 now resolve to a product code**
  (8 `MATCHED`, 4 `RECOVERED_LOWER_VARIANT`) - e.g. `DROFER XT TAB 1*15` -> `OROFER`
  -> 424442006, `CALONATE-D3 TAB 1X15S` -> `CALONAT` -> 424441177.
- 11 rows lose their brand hint entirely: glued OCR spellings (`OROFERXT TAB 1X15`,
  `FERIUMXT DROP 15ML`, `OROFERD3 TAB 1*10`, `DROPER XT DROP 13ML`, `FERIUXT 10-S`)
  that only the removed multi-word entries could match as a single token. They were
  already unmapped, so mapping output is unchanged; the difference is that
  `garbage_check.py` now routes them to LLM review instead of classifying them locally
  as products. Fixing them properly belongs in glued-token alias handling, not in the
  vocabulary.

**Validation**: `load_brands()` returns 849 entries whose set equals the master's
`BRAND_NAME` set exactly; no empty, lowercase, whitespace-edged or quote-bearing entries;
`build_brand_index()` builds 849 pairs; `python -m py_compile garbage_check.py` OK.
Only `Brand Names.txt` and this file were modified (one cosmetic separator fix:
`"SKIZOTUS","S-NUMLO"` -> `"SKIZOTUS", "S-NUMLO"`, inert for `load_brands()`).
