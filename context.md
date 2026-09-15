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

---

## Performance: `_variant_subvariant_tokens` memoized (mapping.py)

**Change** — `mapping.py:1269-1295`. `_variant_subvariant_tokens(brand_map)` rebuilt the
master's VARIANT / SUB_VARIANT alpha-token vocabulary from scratch on **every input row**,
even though the result is a pure function of a `brand_map` that is frozen for the whole
run. It now carries a single-slot cache:

```python
_VST_CACHE_MAP = None
_VST_CACHE_VAL = None
...
    global _VST_CACHE_MAP, _VST_CACHE_VAL
    if _VST_CACHE_MAP is brand_map:      # same master object -> same vocabulary
        return _VST_CACHE_VAL
    ...                                   # loop body unchanged
    _VST_CACHE_MAP, _VST_CACHE_VAL = brand_map, (var, subvar)
    return var, subvar
```

The key is the **identity** of `brand_map`, and the cache holds a reference to it. Keeping
that reference alive prevents the object from being collected and its `id()` reused by a
different map, so a stale hit is impossible. A value-based key was deliberately avoided
for that reason. On a miss the vocabulary is computed by the original loop and stored; on
a hit the stored sets are returned. Nothing else in the function changed.

**Why it was expensive** — the master holds 2,764 items, so each call issues 5,528
`re.findall` invocations, and the only caller (`desegment_ocr_input`, `mapping.py:1327`)
runs once per row from `process_product` (`mapping.py:3288`). Over the 12,972-row corpus
that was ~71.7M `findall` calls to produce the same two sets 12,972 times. cProfile on a
1,500-row slice put the function at 51.4s of 94.4s total - 54% of runtime, and the largest
`tottime` entry in the program.

**Why accuracy cannot move** — the function is pure (reads only `it["variant"]` /
`it["sub_variant"]`, returns derived sets, no I/O or global state). `brand_map` is written
at exactly one place, `build_brand_product_map` (`mapping.py:1041`), and never mutated
afterwards; no assignment to an item's `variant` / `sub_variant` field exists anywhere in
the file. The returned sets reach only `_split_glued_variant` (`mapping.py:1282`), which
does membership tests and never mutates them, so sharing the cached objects is safe. A
second `brand_map` simply misses the guard and recomputes.

**Validation** — full 12,972-row replay of `hints2.tsv` against `PRODUCT_MASTER1.xlsx`
with the LLM disabled, before and after, comparing **every** result field:

| | before | after |
|---|---|---|
| wall clock | 171.4s | **63.0s** |
| per row | 13.21 ms | **4.86 ms** |

- **differing rows: 0** - `product_code` 0, `status` 0, `output` 0, `confidence` 0,
  `suggestions` 0, `candidate_count` 0. The two result dumps are byte-identical.
- 0 exceptions in either leg.
- **speedup 2.72x, 108.4s of 171.4s removed (63%)**.
- Cache reuse proven by instrumenting `re.findall`: **5,528 calls on the first call, 0
  across the next 500**; the same set objects are returned every time.
- Correctness of the guard: the memoized value equals the un-memoized loop for
  `PRODUCT_MASTER1` (351 variant / 2 sub-variant tokens); a *different* `brand_map`
  (`PRODUCT_MASTER.xlsx`, 269 / 59) misses and returns its own fresh, correct value;
  switching back to the first map is correct again.
- `python -m py_compile mapping.py` OK. Only the lines above changed - verified by diffing
  against a pre-edit copy; original mixed line endings (4,009 CRLF + 27 bare LF) preserved.

**Next performance finding, not implemented** — with this cached, the new top cost is
`desegment_ocr_input` rebuilding `brand_compacts` / `brand_letters` per row
(`mapping.py:1314-1315`, plus the same comprehension in `strip_manufacturer_noise`,
`mapping.py:1225`): ~2.7M calls each to `compact_brand_token` and `_brand_letters`,
~26s cumulative in the 1,500-row profile. Same safety argument, separate change.

## Performance: character-set fuzzy pre-filter (`garbage_check.py`)

**Change** — three edits, no matching logic touched.

1. `garbage_check.py:601-613` - `build_brand_index` now also fills a module-level
   `_BRAND_CHAR_SETS` dict, `brand_norm -> frozenset(brand_norm)`, once per run. The dict
   form was chosen over widening the index tuples so `exact_product_match` stays
   byte-identical; the value is a pure function of the key, so two brands sharing a
   normalized form cannot corrupt an entry.
2. `fuzzy_product_match` - `span_sets = [frozenset(span["norm"]) for span in spans]` is
   built once before the brand loop, and `brand_set = _BRAND_CHAR_SETS[brand_norm]` is
   hoisted next to `max_edits` (one lookup per brand, not per pair). The inner loop is
   `for span, span_set in zip(spans, span_sets)`, which preserves span visitation order.
3. The 10-line `# --- Character-bag pre-filter ---` block (two dict-counting loops plus
   `sum(abs(v) for v in freq.values())`) is replaced by:

```python
if len(span_set - brand_set) > max_edits:
    continue
```

Loop nesting, brand order, the `max_edits` rule, the length filter, the
`edit_distance_leq_one` -> `damerau_distance` sequence and the order-sensitive
rank/tie-break block are unchanged.

**Why it was expensive** — the bag block ran ~2,165 times per fuzzy row and rebuilt the
span's own character counts for every brand, even though the span is fixed across the
whole brand loop. Per fuzzy row: 9,895 (brand, span) pairs examined, 2,165 pass the length
filter and built a bag, 1.2 pass the bag, 1.0 reach the DP - i.e. nearly all the time went
into a pre-filter rejecting 99.94% of what it saw.

**Why accuracy cannot move** — `len(span_set - brand_set) > max_edits` is a *necessary*
condition for Damerau-Levenshtein distance <= `max_edits`: a character present in the span
but absent from the brand must be removed by a deletion or a substitution (transpositions
only reorder, never remove), and distinct characters need distinct edits. The new filter is
therefore weaker, never stronger - every pair the old filter accepted still passes. Pairs
that newly slip through reach the exact distance checks, and an exact result <= `max_edits`
always satisfied the old `char_diff <= 2*max_edits` bound too, so no pair can become a
candidate that was not one before. The candidate sequence handed to the order-sensitive
tie-break is unchanged.

**Measured** (17,654-row `test.xlsx` / sheet `garbage_check`, 849 brands,
`USE_LLM_FALLBACK = False`):

| | old | new |
|---|---|---|
| `fuzzy_product_match` over the 5,404 rows that reach it | 74.2s (13.74 ms/row) | **30.8s (5.70 ms/row)** - 2.41x |
| full `main()` end to end | 62.3s | **29.9s** - 2.09x |

**Verified identical** — unit level: both module versions loaded in one process, same brand
list (849, identical order) and same spans; over all 5,404 fuzzy rows, 0 dict mismatches
and 0 mismatches on `brand` / `matched_text` / `match_type`. End to end: `main()` run for
each version against a separate copy of `test.xlsx` with `CLEAR_EXISTING = True`, then
columns B:G diffed cell by cell across all 17,655 rows - **0 differences**.
`python -m py_compile garbage_check.py` OK; startup still prints `849 brands loaded`;
line endings preserved (1,204 CRLF, 0 bare LF).

**Not implemented** — the analysis' follow-up (indexing spans by normalized length, which
would take the fuzzy path to ~11.9s / 5.0x overall) was deliberately left out: it reorders
span visits within each brand, so its safety rests on a tie-break argument rather than on
the loop being untouched.

**Unrelated pre-existing state** — `EXCEL_FILE` is set to `"test1.xlsx"`, which does not
exist in the project (only `test.xlsx` does). That predates this work and was not changed;
the validation harness overrides `EXCEL_FILE` at runtime instead.

## Configuration: `PROCESSING_MODE` (`garbage_check.py`)

`USE_LLM_FALLBACK` has been replaced by a single named mode switch. Earlier references to
`USE_LLM_FALLBACK = False` in this document therefore describe what is now
`PROCESSING_MODE = "local"`.

```python
PROCESSING_MODE = "local"   # "llm" | "local" | "both"
```

| value | behaviour |
|---|---|
| `"local"` | local classification logic only; no LLM/API call is reachable |
| `"llm"` | local pre-pass plus the LLM fallback that resolves its unconfirmed rows |
| `"both"` | local logic and LLM processing per the current pipeline |

`"llm"` and `"both"` select the same route, because the pipeline is local-first by
construction: the per-row loop is what fills `llm_pending`, so the LLM only ever sees the
rows local classification could not confirm and no row can reach it without the local pass.
Both spellings are accepted so the setting reads the way the caller expects; there is no
LLM-only path to route to, and creating one would have changed classification behaviour.

**Wiring** — three edits, no classification or matching logic touched:

- `garbage_check.py:21-38` — `USE_LLM_FALLBACK = False` replaced by `PROCESSING_MODE`,
  `VALID_PROCESSING_MODES = ("llm", "local", "both")`, and an import-time `ValueError` on
  an unrecognised value (a typo fails loudly instead of silently disabling the LLM).
- `garbage_check.py:938-944` — new `llm_enabled()` returns
  `PROCESSING_MODE in ("llm", "both")`. It reads the module global at call time, so a
  harness can set `m.PROCESSING_MODE` on the imported module and have it take effect.
- `garbage_check.py:948` and `garbage_check.py:1190` — the two LLM gates now call
  `llm_enabled()`. The API-key guard in `main()` is gated too, so `"local"` no longer
  requires a key to be present.

`PROCESSING_MODE` is the only control point: `USE_LLM_FALLBACK` no longer exists, so there
is no second flag that can contradict it. Harnesses that set `USE_LLM_FALLBACK` on the
imported module must set `PROCESSING_MODE` instead.

**Verified** — `python -m py_compile garbage_check.py` OK. `llm_enabled()` returns
`False` / `True` / `True` for `"local"` / `"llm"` / `"both"`; `PROCESSING_MODE = "LOCAL"`
raises `ValueError` at import. AST check: `requests.*` is issued only by `call_model`, the
only `classify_batch` call in `main()` is the one inside `if llm_enabled():`, so no API
call is reachable in `"local"` mode. End to end, the pre-change file with
`USE_LLM_FALLBACK = False` and the new file with `PROCESSING_MODE = "local"` were each run
by `main()` against a separate copy of `test.xlsx` with `CLEAR_EXISTING = True` and
`requests.post` patched to raise: both completed without attempting a call, and columns
B:G diffed cell by cell across all 17,655 rows show 0 differences. Line endings
preserved (1,230 CRLF, 0 bare LF).

## Configuration: `PROCESSING_MODE` (`mapping.py`)

`mapping.py` had no LLM switch at all: the Step 6 reranker was called unconditionally for
every row that Steps 1-5 left with more than one plausible candidate. Runs with the model
disabled were previously simulated by monkeypatching `call_groq_llm` to raise, which worked
but labelled every such row `RECOVERED_API_ERROR`. A named mode switch now exists.

```python
PROCESSING_MODE = "both"   # "llm" | "local" | "both"
```

| value | behaviour |
|---|---|
| `"local"` | local candidate generation, filtering and ranking only; no API call is reachable |
| `"llm"` | local pipeline plus the Step 6 LLM reranker |
| `"both"` | local logic and LLM reranking per the current pipeline (pre-existing behaviour) |

The default is `"both"` because that is what the file did before. Note this differs from
`garbage_check.py`, whose default is `"local"` — there the pre-existing state was
`USE_LLM_FALLBACK = False`. Each file's default preserves its own prior behaviour.

`"llm"` and `"both"` select the same route. The pipeline is local-first by construction:
Steps 1-5 build, filter and rank the candidate set, and the reranker only ever sees rows
those steps could not settle. There is no LLM-only path to route to, and creating one would
mean bypassing candidate generation — a change to matching behaviour, not a config option.

**Wiring** — three edits, no classification, candidate-generation, filtering or ranking
logic touched:

- `mapping.py:103-121` — `PROCESSING_MODE`, `VALID_PROCESSING_MODES = ("llm", "local",
  "both")`, and an import-time `ValueError` on an unrecognised value, so a typo fails
  loudly instead of silently disabling the reranker.
- `mapping.py:2975-2981` — new `llm_enabled()` returns `PROCESSING_MODE in ("llm", "both")`,
  reading the module global at call time so a harness can set `m.PROCESSING_MODE` on the
  imported module and have it take effect.
- `mapping.py:3628` — the gate, placed in `process_product` immediately after
  `rerank_documents = build_rerank_documents(items)` and before the prompt is built. In
  `"local"` mode it returns the row from the ranking Steps 1-5 already produced, reusing the
  same `local_best_item` / `local_suggestions` the reranker's error path falls back to:
  `RECOVERED_LOCAL_ONLY` / `LOW` when there is a local best, `LOCAL_ONLY_UNRESOLVED` /
  `NONE` when there is not. `LLM_REQUEST_ATTEMPTS` is deliberately not incremented, which
  is what keeps the `DELAY_BETWEEN_LLM_REQUESTS` throttle in `main()` from firing — that
  throttle was already conditional on the counter advancing, so it needed no edit.

Gating inside `process_product` covers every entry point: all three `process_product` call
sites (the lower-variant recursive retry at `mapping.py:3482`, the batch loop in `main()` at
`mapping.py:3966`, and the interactive block at `mapping.py:4067`) go through the same Step 6.

**Two new status values.** `RECOVERED_LOCAL_ONLY` mirrors `RECOVERED_API_ERROR` and
`LOCAL_ONLY_UNRESOLVED` mirrors `API_ERROR`, but they say what actually happened instead of
reporting an API error for a call that was never attempted. Only one status consumer exists
in the file — `res["status"] == "MATCHED"` in the lower-variant recovery at
`mapping.py:3488` — and neither new value is `MATCHED`, exactly as neither of the two
statuses they mirror is, so control flow there is unchanged.

**Verified** — `python -m py_compile mapping.py` OK. `llm_enabled()` returns `False` /
`True` / `True` for `"local"` / `"llm"` / `"both"`; `PROCESSING_MODE = "LOCAL"` raises
`ValueError` at import. AST check: `requests.post` and `requests.RequestException` appear
only inside `call_groq_llm`, the sole `call_groq_llm` call sits at `mapping.py:3653`, and the
gate at `mapping.py:3628` precedes it, so no API call is reachable in `"local"` mode.

Full-corpus replay, 12,972 `Source='0'` rows against `PRODUCT_MASTER.xlsx` with
`forced_brand` hints, `requests` replaced by an object that raises `AssertionError` on any
`post` (none was raised in any leg):

| leg | configuration | time |
|---|---|---|
| A | pre-change file, `call_groq_llm` patched to raise | 299s |
| B | new file, `PROCESSING_MODE = "both"`, `call_groq_llm` patched to raise | 270s |
| C | new file, `PROCESSING_MODE = "local"`, nothing patched | 236s |

- **A vs B on `product_code` + `status` + `output` + `confidence`: 0 differences.** The
  default mode reproduces pre-change behaviour exactly.
- **A vs C on `product_code` + `output` + `confidence`: 0 differences.** Every mapping
  decision is identical to the established LLM-off baseline.
- A vs C on `status`: 416 differences, all of them the intended relabel
  `RECOVERED_API_ERROR` -> `RECOVERED_LOCAL_ONLY`. Mapped-row count 12,073 in all three legs.

`LOCAL_ONLY_UNRESOLVED` did not fire on this corpus — no row reaches Step 6 without a local
best. The `API_ERROR` branch it mirrors is unexercised here for the same reason (0
`API_ERROR` rows in leg A), so both remain untested against real data.

Line endings preserved: 4,074 CRLF, 27 bare LF (was 4,025 / 27; +49 lines, no LF drift).
`PROCESSING_MODE` is the only control point — there is no second flag that can contradict it.

## Console output: batch mode reduced to one line per row (`mapping.py`)

Batch runs used to print the whole internal filter trace for every row — the brand
banner, `After variant filter`, `After sub-variant filter`, `After name-strength
filter`, `After form prioritization`, `Detected pack size`, `After pack size filter`,
`Calling reranker...`, then `✓ Matched (...)` and `Code: ...` — followed by a bare
`[9/17094] <input>` progress line. Roughly ten lines per product, and the mapped
product and the input never appeared together.

`process_product` already had a dual-output convention: `if verbose:` prints the
interactive STEP trace, `else:` / `if not verbose:` printed the terse batch lines.
The cleanup deletes the batch branches only. Every `if verbose:` branch is byte
identical, so the interactive single-product mode (menu option 2) is unchanged. Two
unconditional brand-hint routing notes (`[forced_brand] ...`, `[auto_detect] ...`)
were rewrapped as `if verbose:` rather than deleted, so they are still available when
debugging one product. The one print left unguarded inside `process_product` is
`  API Error: {e}` — a real failure, not progress noise.

The per-row line is emitted by `process_excel_file`, which already holds both the
input name and the result dict, so no new output was added inside internal
functions and the processing order is untouched. `format_row_result()` (just above
`process_excel_file`) reads `res["output"]`, `res["product_code"]` and
`res["status"]` verbatim — the console shows the mapping result, never a
reconstructed or guessed name:

```text
[ 1/12] CORDARON 100 TAB 10 TAB                      → CORDARONE 100 MG TABLET 20x15T [421112908]  (MATCHED)
[ 4/12] AMARYL MP1MG TAB BL1X20S                     → NO_CLEAR_MATCH  (VARIANT_NOT_IN_MASTER)
[ 9/12] EPOFER 4000PFS 1X1                           → EPOFER 4000 - PFS [422220109]  (RECOVERED_API_ERROR)
```

The progress counter is kept and right-padded to the width of the total, so the
arrow column stays aligned from `[   1/17094]` to `[17094/17094]`. A row that does
not map prints its real placeholder `NO_CLEAR_MATCH` plus the status that produced
it (`VARIANT_NOT_IN_MASTER`, `AMBIGUOUS_STRENGTH`, `REVIEW`, ...) — it is never
dressed up as a match. Already-mapped rows print
`→ (already mapped, skipped)` and exceptions print `→ ERROR: <msg>` on the same
one-line format instead of a bare `  Error: ...`.

`print(` call count in `mapping.py`: 198 -> 176. No mapping logic, candidate
generation, scoring, filtering, reranking, classification, written output or
processing order was touched; every edit is a print removal, an `if verbose:` guard,
or the new formatting helper.

Verification:
- `python -m py_compile mapping.py` clean; 4,090 CRLF / 0 bare LF (no line-ending drift).
- Replayed 3,532 rows (all 1,662 non-`MATCHED` rows plus a 1-in-6 sweep) against the
  pre-cleanup result cache with the reranker disabled: `product_code`, `status`,
  `output` and `confidence` identical on **every** row, 0 differences. A 927-row
  re-run after the counter-alignment tweak: 0 differences.
- Of those 3,532 rows, exactly 416 still printed anything from `process_product` —
  the 416 `RECOVERED_API_ERROR` rows, i.e. only the genuine `API Error` line.
- A 12-row batch through `process_excel_file` against a temp workbook (real
  `test.xlsx` untouched) produced the lines above, and each line matches the row
  written to the `mapping` sheet field for field.

## Variant-aware parenthetical stripping (`mapping.py`)

`strip_parenthetical_noise` deleted every parenthetical of 1-4 letters as
purchase-source noise (`(AP)`, `(SP)`, `(DPCO)`, `(EMC)`). Some suppliers use the
same notation to carry the sub-brand instead, so the rule was also deleting the
only token that identified the variant:

```text
AMARYL-(M)1 20S            -> AMARYL- 1 20S      -> AMARYL 1MG TABLET      (plain, wrong)
CARDACE (H) 5 TAB 1X15     -> CARDACE 5 TAB 1X15 -> CARDACE 5 MG TABLET    (plain, wrong)
TEMSAN (CT) 40TAB 1X10     -> TEMSAN 40TAB 1X10  -> TEMSAN 40 MG. TABLETS  (plain, wrong)
```

The tag is now looked up against the **resolved brand's own** `VARIANT` values. If
it is one of them the parenthetical is unwrapped (`(X)` -> ` X `) so the variant
survives into matching; if it is not, it is deleted exactly as before. Nothing
about candidate generation, filtering, ranking, reranking, fallback or the
`forced_brand` routing changed — only the string handed to them.

The variant vocabulary comes from `get_brand_variant_tags(brand_map, brand_hint)`,
a new read-only helper: it resolves the hint to a `brand_map` key (exact, then
case-insensitive scan — the same two-step lookup the `forced_brand` branch already
uses), and returns that key's `VARIANT` values restricted to 1-4 letters, with the
`PLAIN` sentinel excluded. It reads the post-`repair_master_variant_fields` values
held in `brand_map`, so it agrees with what the variant filter later sees.

Scope is deliberately limited to rows that arrive with a brand hint (`garbage_check`
column B == `"0"`), because `strip_parenthetical_noise` runs before brand detection.
`MAYBE_PRODUCT` rows pass `forced_brand=None`, get `brand_variant_tags=None`, and
keep the old behaviour byte for byte; the same is true of the second call site in
`suggest_local_match`, which was left untouched and relies on the default argument.
16 `MAYBE_PRODUCT` rows carry a short parenthetical and are therefore not covered —
only one of them (`SEMI AMARYL ( M ) 10 TAB`) looks like a real variant tag.

A tag that already stands as its own token elsewhere in the name is left deleted:
there is no variant to rescue, and re-inserting it duplicated the token. This
guard exists because it prevented a measured regression —
`TEMSAN-H TABLETS (H )15 TABLET 1X15 TAB` already carries `H` in `TEMSAN-H`, and
unwrapping produced `... TABLETS H 15 ...`, which moved the pick from
`424441191 TEMSAN H TABLETS 1X15 T` to `424441192 TEMSAN 80 H TABLETS 1X15T`,
inventing a strength the input never had.

Verification:
- `python -m py_compile mapping.py` clean; 4,122 CRLF / 0 bare LF (no line-ending drift).
- Blast radius measured, not assumed: the preprocessed input string was computed for
  **all 17,093 driver rows** (`"0"` + `MAYBE_PRODUCT`) under both versions.
  Exactly **37 rows** differ. Every other row is byte identical going into the rest
  of the pipeline, and the new helper is pure, so those rows cannot change.
- All 125 rows containing a short parenthetical were fully re-mapped under both
  versions with the reranker disabled: **30 changed, 95 identical**. All 30 move
  from a `VARIANT=PLAIN` SKU (or `NO_CLEAR_MATCH`) to the variant the input names —
  `CARDACE H/AM/METO`, `AMARYL M/MV/M FORTE`, `CORDARONE X`, `TELSITE H`,
  `FERIUM XT +`, `PROXYM ER`, `TEMSAN CT`, `ASOMEX D`, `HOSIT FE`. 0 regressions.
- The tags that are **not** variants of their row's brand are unchanged, confirmed
  row by row: `(SP)`, `(HP)`, `(NR)`, `(DPCO)`, `(EMC)`, `(EMCU)`, `(NRX)`, `(BOX)`,
  `(TAB)`, `(MCP)`, `(PEN)`, `(b)`, `(G)`, `(ND)`, `(NET)`, `(ORSL)`. The four cases
  called out explicitly still map as before: `AMARYL 1 (30) (D)` -> 421113127,
  `CETAPIN XR 1000 (15) (D)` -> `CETAPIN XR 1000MG TABLET 10x15T`,
  `LASIX INJ (D)` and `LASIX INJECTION(4ML)(D)` -> `LASIX 40 MG/4ML INJECTION`.
  `D` is a variant of `ASOMEX` and `PREGANZA` but not of `AMARYL`, `CETAPIN` or
  `LASIX`, which is exactly why brand scoping is required.
- Regression control: 800 rows drawn at random (seed 20260826) from the 17,056
  rows whose preprocessed input did not change were fully re-mapped both ways —
  **800/800 identical** on `product_code`, `output` and `status`.
- One changed row is better but still not exact: `PROXYM 300(ER)TAB 1X15` now maps
  to `424441862 PROXYM ER TABLET 1X15T`. The master has no `PROXYM ER 300` SKU
  (`ER` exists only at `PLAIN` and `200`), so the strength is unrepresentable; the
  previous answer `PROXYM 300 TABLETS` had the strength but the wrong formulation.
- The project has no test suite or validation script; the A/B replay harness above
  (two copies of `mapping.py` imported from separate directories, `call_groq_llm`
  monkeypatched to raise, driven from the real `test.xlsx` `garbage_check` sheet)
  is the comparison that was run. It lives outside the project directory.

## Pipeline rule: only `remark = 0` rows are handed to `mapping.py`

`garbage_check.py` writes its classification into column B of `test.xlsx!garbage_check`
(column A = `PRODUCT_NAME`, column B = remark, column C = matched brand hint). The
remark distribution over the current sheet is `0` 13,065 / `MAYBE_PRODUCT` 4,028 /
`REVIEW` 520 / `1` 41 (17,654 rows).

Rule: **only rows whose remark is exactly `0` (confirmed product) are processed by
`mapping.py`.** `MAYBE_PRODUCT` rows and every other remark are excluded from the
processing input — they are not reclassified, not converted to `NO_CLEAR_MATCH`, and
no status is written for them; they are simply never read into `df`.

Implementation — two additions to `mapping.py`, nothing removed or reordered:

- `mapping.py:110-120` — new module constant beside the existing
  `PROCESS_MAYBE_PRODUCT_ONLY`:

  ```python
  PROCESS_CONFIRMED_ZERO_ONLY = True
  if PROCESS_CONFIRMED_ZERO_ONLY and PROCESS_MAYBE_PRODUCT_ONLY:
      raise ValueError(...)   # the two flags are mutually exclusive
  ```

  Set it `False` to restore the previous behaviour of also processing
  `MAYBE_PRODUCT` rows.

- the row-combination block in the reader (`# Combine rows based on flag`) gains a
  leading branch; the pre-existing `PROCESS_MAYBE_PRODUCT_ONLY` branch becomes `elif`
  and the `else` concat is untouched:

  ```python
  if PROCESS_CONFIRMED_ZERO_ONLY:
      df_combined = df_0.copy().reset_index(drop=True)
  elif PROCESS_MAYBE_PRODUCT_ONLY:
      ...
  else:
      df_combined = pd.concat([df_0, df_mp], ignore_index=True)
  ```

`garbage_check.py` is unchanged — no classification logic was touched. The `mask_0` /
`mask_mp` selection already matched only `"0"` and `"MAYBE_PRODUCT"`, so the 561
`REVIEW`/`1` rows were never processed; the only behavioural change is the exclusion
of the 4,028 `MAYBE_PRODUCT` rows.

Verified results:

- `py_compile` clean on both `mapping.py` and `garbage_check.py`; `mapping.py` remains
  CRLF throughout (4,146 CRLF / 0 bare LF).
- Row selection against the real workbook: **13,065 rows handed to the mapper**, all
  remark `0`; 4,028 `MAYBE_PRODUCT` skipped; 561 other remarks excluded as before
  (previous behaviour was 17,093 rows).
- No other mapping behaviour changed: 400 remark-`0` rows sampled (seed 20260826) and
  replayed through pre-patch and post-patch copies of `mapping.py` imported from
  separate directories with `call_groq_llm` monkeypatched to raise — **400/400
  identical** on `output`, `product_code` and `status`. `process_product` was not
  modified.

Consequence to be aware of: `_save_to_mapping_sheet` deletes and recreates the
`mapping` sheet on each run, so the next run writes 13,065 rows and the previously
written `MAYBE_PRODUCT` result rows will no longer appear there. That follows directly
from excluding them from the processing input.


## Form-aware retry guard in Steps 2 / 3 / 3.5 (`mapping.py`)

### Root cause

`prioritize_by_dosage_form` (Step 4) is a **filter, not a re-orderer** -
`if exact_form_matches: return exact_form_matches`. It therefore cannot rescue a
dosage-form-correct SKU that an *earlier* subtractive filter has already deleted:

* Step 3 `filter_items_by_sub_variant` - its existing form-aware branch is gated on
  `if not specific_detected and detected_forms:` (`mapping.py:2548`), i.e. it fires only
  when *no* specific numeric strength was found, which is the exact inverse of the
  failing case. `PAUSE 500MG INJ` narrows on `SUB_VARIANT=500`, which exists only as
  `PAUSE TABLET 500 MG`, so `PAUSE INJECTION` (`SV=PLAIN`) is gone before Step 4.
* Step 2 `filter_items_by_variant` - its PLAIN fallback collapses the scope to a single
  PLAIN SKU (`VICARE 1-CREAM 1X30GM` -> `VICARE CAPSULES`).

Both then exit through the unfloored sole-survivor short-circuit
(`if len(items) == 1:` -> `MATCHED`/`HIGH`), which is why every one of these rows
reports high confidence while being wrong.

### Implemented change - 54 inserted lines, 0 deleted, 0 reordered

Filter order is **unchanged** and `prioritize_by_dosage_form` is **untouched**. Steps 2,
3 and 3.5 run exactly as before; only on a form-incompatible result is the *same* filter
re-run against the form-compatible pool.

1. New helper `_form_compatible(items, detected_forms)` immediately after
   `prioritize_by_dosage_form`. Compares dosage-form **groups** (`_FORM_GROUPS`).
   **A product with no detectable form counts as COMPATIBLE** - many correct master names
   carry no form token (`C-ZID 1GM`, `MATERNA-HMG 150 I.U.`, `METPURE H 50 TABELTS`), and
   discarding them is what makes a pure form-first ordering destructive (measured
   separately: ~35 regressions).
2. `_input_forms = detect_dosage_form_in_input(input_name, verbose=False)` hoisted above
   Step 2. Step 4 keeps its own call and its own verbose output; the function is pure.
3. After each of Steps 2/3/3.5: if the filter result is form-incompatible **and** a
   form-compatible pool exists in that step's own input set, re-run the same filter on
   the pool and adopt the result only if non-empty.

**Critical detail - `reference_items` is narrowed on the retry.** The PLAIN fallback
inside `filter_items_by_variant` (`filtered = plain_products if plain_products else
ref_plains`) reads from `reference_items`, so an un-narrowed retry re-injects the SKU the
guard just excluded. Verified directly:

```
input forms       : ['CREAM']
normal Step-2 out : ['VICARE CAPSULES']          form-compatible? False
compatible pool   : ['VICARE I-CREAM 1x30GM']
retry ref=FULL    : ['VICARE CAPSULES']          <-- excluded SKU re-injected
retry ref=POOL    : ['VICARE I-CREAM 1x30GM']    <-- what the landed code does
```

### Validation - full 13,065-row A/B replay

`python -m py_compile mapping.py garbage_check.py` -> OK. Line endings preserved:
`mapping.py` CRLF 4200 / bare LF 0; `garbage_check.py` CRLF 1230 / bare LF 0 and
**unmodified** (`git diff --stat garbage_check.py` empty).

Harness: two copies of `mapping.py`, baseline produced by neutralising the single hoisted
line (`_input_forms = set()`), which makes all three guards dead code and reproduces exact
pre-change behaviour. `call_groq_llm` monkeypatched to raise, so **no LLM call occurs**;
every `remark = 0` row of `test.xlsx!garbage_check` replayed through `process_product`
with its column-C brand hint.

```
base 13065 rows | new 13065 rows | row keys identical
rows differing (output / product_code / status): 41
  code -> none : 0        <- no mapping lost
  none -> code : 1
  code switched: 40
```

`code -> none = 0`: the guard never turns a mapped row into `NO_CLEAR_MATCH`. Named cases:

| Input | Before | After |
|---|---|---|
| `PAUSE 500MG INJ` (x6) | `424441657 PAUSE TABLET 500 MG` | `424441951 PAUSE INJECTION` |
| `FERIUM INJ 1K` (x9) | chewable tablets / drops / syrup | `421111885 FERIUM INJECTION 1K (1GM/20ML)` |
| `FERIUM ... INJ` (x6) | chewable tablets / syrup | `421110953 FERIUM INJECTION 500MG` (form right, strength unverified) |
| `VICARE 1-CREAM` (x4) | `421110538 VICARE CAPSULES` | `421113026 VICARE I-CREAM 1x30GM` |
| `CORDERONE INJ.3M X1` | `421112910 CORDARONE X 200 MG TABLET` | `421112909 CORDARONE 150 MG/3ML INJECTION` |
| `ZOSECTA 40MG INJ` (x2) | `424440050 ZOSECTA TABLET` | `421112230 ZOSECTA IV 20Mg.` |
| `OROFER XT SYR/SUSP` (x2) | **correct** | **REGRESSED - see below** |

Also corrected: `PANSALVE 40 INJ VIAL` -> `421111709`; `IMINORAL SYP 50ML` ->
`421111811`; `VITANOVA D3 6L INJ` -> `423330702`; `C-LET 2.0 GM INJECTION IV` ->
`421112471`; `PAUSE 1000 MG TAB` (x2) -> `424442023 PAUSE 1000 TABLETS`;
`FERIUM XT 15\ 1*15TAB` -> `424441955`. The single `none -> code` gain is
`RABIFAST INJ 1 VAIL`: `AMBIGUOUS_STRENGTH`/`NO_CLEAR_MATCH` ->
`423331410 RABIFAST IV 20 MG INJECTION`.

### Open regression - 2 rows, NOT yet fixed

Validation point 5 ("previously correct mappings do not change unexpectedly") **fails on
two rows**:

```
[8908]  'OROFER XT+ SYR 200 ML'  424441225 OROFER XT + SUSPENSION 200 ML -> 424440118 OROFER XT-DHA 7 X 1 KIT
[10424] 'OROFER XT SYR 1X150'    424440590 OROFER XT SUSPENSION 150 ML   -> 424440118 OROFER XT-DHA 7 X 1 KIT
```

Cause is **pre-existing and outside this change**: `\bSYR\.?\b` is listed under INJECTION
in the form table at `mapping.py:2636` (intended as *syringe*), but the OCR corpus writes
**SYRUP** as `SYR`. The guard did not create the mislabel, it only made it consequential.

The remedy is one edit - drop bare `\bSYR\.?\b` from the INJECTION alternation, keeping
the counted form `\d+SYR\.?\b` (`1SYR` = a real syringe). Measured, not applied:

```
guard-only        vs guard+SYRfix : 2 rows differ - both are restorations to the baseline code
base vs guard+SYRfix : 39 differ | code->none 0 | none->code 1 | switched 38
```

**Not applied**, because this task's directive was "Do not make any other changes". Note
`\bSYR\.?\b` also appears under INJECTION in `detect_dosage_form_in_product`
(`mapping.py:2699`); only the input-side occurrence at 2636 was measured, and 2699 should
be left alone unless separately measured.

Honest decomposition of the 41: ~24 confirmed correct, 8 form-correct with strength not
independently verified (the FERIUM 500MG group), 6 reranker-dependent statuses whose final
code still depends on Step 6, 1 neutral (`DROPER XT PLUSE SYP 200ML` moves between two
wrong SKUs - the brand itself is a bad upstream fuzzy match), **2 regressions** above.

MAPPING_MAIN.py created as the new combined garbage-check + product-mapping pipeline; detailed context is maintained in context_combined.md.
