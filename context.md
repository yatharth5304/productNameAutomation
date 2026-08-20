# Product Name Automation — Project Context

Last updated: 2026-08-19 (afternoon session)

---

## Project Overview

This repository contains **two distinct but related scripts** for working with Indian pharmaceutical product data:

1. **`garbage_check.py`** — classifies OCR-extracted lines from pharmaceutical stock-and-sales reports as either a real product row (`0`) or garbage/non-product content (`1`). Output goes into `test.xlsx`.

2. **`mapping.py`** — given a messy, OCR'd, or abbreviated product name (from a distributor or stock report), identifies the single correct matching record in a canonical master product list (`PRODUCT_MASTER.xlsx`) and returns the exact master product name and product code. Output goes into a separate working workbook (`prachi6.xlsx` by default, configurable).

Both scripts are single-file prototypes; neither is a packaged application. They share the same domain context (Emcure pharma product data at Vintyaa) but solve separate stages of the pipeline: `garbage_check.py` filters out non-product lines, and `mapping.py` normalizes the surviving product lines against a master catalog.

---

## Current Repository Structure

```text
Product Name Automation/
|-- garbage_check.py
|-- mapping.py
|-- maybe_product_2_check.py   <- standalone script to test MAYBE_PRODUCT_2 logic before pipeline integration
|-- Brand Names.txt
|-- PRODUCT_MASTER.xlsx
|-- test.xlsx          ← three named sheets: garbage_check (classifier output) + mapping (mapping output) + maybe_product_2 (MP2 candidates for manual review)
|-- context.md
`-- __pycache__/
    `-- garbage_check.cpython-314.pyc
```

---

## File Responsibilities

### `garbage_check.py`

Main OCR classification script. Responsibilities:

- Defines runtime configuration constants
- Loads the local brand list from `Brand Names.txt`
- Performs local product detection by scanning the full normalized row for brand matches
- Performs exact local garbage detection using a maintained garbage phrase list
- Computes fuzzy candidates for unresolved rows using `difflib.SequenceMatcher`
- Builds a simplified LLM prompt only for unresolved rows
- Calls the Novita-compatible chat completions endpoint with retry logic
- Parses model responses into `0` / `1` / `BAD_OUTPUT` / `ERROR`
- Writes results into `test.xlsx`
- Prints token/cost estimates at the end

Main functions:

- `load_brands()`: extracts quoted brand names from the brand text file
- `normalize_text()`: strips case, spacing, and punctuation noise for matching
- `edit_distance_leq_one()`: strict single-edit tolerance (substitution, insertion, deletion); used for first-pass 1-edit check in fuzzy product matching
- `damerau_distance(a, b, max_dist)`: Damerau-Levenshtein distance capped at `max_dist`; adds adjacent-transposition detection; per-row early-exit if no DP cell in a completed row is `<= max_dist`; used for transposition hits and 2-edit long-brand hits
- `build_brand_index()`: precomputes normalized brands
- `row_tokens()`: extracts OCR row tokens used for whole-row span matching
- `token_spans()`: generates consecutive token spans across the full row
- `exact_product_match()`: first pass — exact normalized span match against brand index
- `fuzzy_product_match()`: second pass — adaptive edit-distance span match; `<= 1` edit for brands 4–9 chars, `<= 2` edits for brands `>= 10` chars; transpositions detected via `damerau_distance`; character-bag pre-filter eliminates ~90% of `damerau_distance` calls; returns matched brand, matched text, and match type (`1-substitution`, `1-insertion`, `1-deletion`, `1-transposition`, `2-edit`)
- `build_garbage_set()`: normalizes the explicit garbage phrase list
- `is_exact_garbage_row()`: local exact garbage confirmation
- `is_pattern_garbage_row()`: regex-pattern garbage confirmation; uses `SAFE_GARBAGE_PATTERNS` (~65 patterns covering Bill Nos, Division/Company/Supplier headers, PURC/EXP financial summaries, HTML artifacts, software-generated rows, report footers, distributor/agency names, LTD/LIMITED suffix rows, GSTIN rows, monthly financial rows, column-header artifacts, etc.); formerly-MAYBE_GARBAGE patterns are now included here and classify directly as `1`
- `is_metadata_garbage_row()`: vocabulary-based garbage confirmation; returns True when every alphabetic token belongs to `METADATA_GARBAGE_VOCAB`; catches composite header/summary rows not covered by the exact list
- `is_maybe_product_row()`: returns True when a row contains a pharmaceutical dosage-form token AND a qty/strength signal AND no structural-garbage keyword; used to label uncertain-but-likely-product rows as `MAYBE_PRODUCT`
- `fuzzy_garbage_match()`: fuzzy match against the exact garbage phrase list; marks rows `1` with matched phrase and match type in columns F/G
- `brand_candidates(name, brands)`: tokenizes row text, computes top fuzzy matches, filters by minimum score (used only for LLM prompt enrichment)
- `format_row(i, name, cands)`: converts a row plus candidates into prompt-ready text
- `call_model(prompt, max_tokens)`: sends the HTTP request, handles retries/rate limits, accumulates token usage
- `classify_batch(names, brands)`: builds a batch prompt and parses model output into per-row decisions
- `has_result(value)`: treats blank cells as unclassified
- `main()`: orchestration entry point for workbook read/process/write flow

### `mapping.py`

Pharma product-name matching and normalization engine. Responsibilities:

- Reads a column of messy/OCR'd input product names from an Excel workbook
- Loads the canonical `PRODUCT_MASTER.xlsx` and builds an in-memory brand-keyed index
- Runs each input through a multi-stage deterministic rule-based funnel (brand → variant → sub-variant → dosage form → pack size → local rank) without calling an LLM
- Falls back to a Groq-hosted LLM reranker only when local ranking cannot confidently pick a single candidate
- Writes back matched product name, product code, match status, confidence, candidate count, and top-N suggestions
- Supports two operating modes: **batch** (`process_excel_file`) and **interactive** (`interactive_mode` — REPL for single-input debugging with verbose step-by-step tracing)

Current version tag: `updated_ts_5` (10 fixes on top of `updated_ts_3` baseline).

#### `mapping.py` — file regions

| Lines (approx.) | Region | Contents |
|---|---|---|
| 1–83 | Header changelog | `FIX 1–10` comment block documenting `updated_ts_5` vs `updated_ts_3`. Inline decisions log. |
| 85–175 | CONFIG | API key/URL/model, hardcoded file paths, Excel column names, matching thresholds, token-equivalence dictionaries. |
| 177–208 | PROMPTS | `SYSTEM_PROMPT` and `USER_PROMPT_TEMPLATE` for the Groq reranker call. |
| 210–650 | Utils / normalization | `norm`, `normalize_variant_token`, `compact_*_token`, strength-token expansion, OCR numeric-noise repair, master-field repair, pack-size value parsing, duplicate-word cleanup. |
| 651–740 | Compound-token filtering | `extract_brand_suffix_tokens`, `filter_items_by_compound_tokens`. |
| 740–890 | Pack-position / strength-variant splitting | `get_all_pack_positions`, `split_strength_and_variant`, `detect_compact_variant_strength_combo`. |
| 890–952 | Pack size extraction (FIX 1) | `extract_pack_size_from_input` — regex-order fix; MG/MCG/IU removed from pack extraction (they are dosage strengths, not pack counts). |
| 952–1060 | Master data loading | `load_master`, `build_brand_product_map`, `extract_all_variants_from_data`, `extract_all_sub_variants_from_data`. |
| 1060–1730 | Brand detection | `find_potential_brands`, `group_related_brands`, `strip_manufacturer_noise` (hardcoded to EMCOO/Emcure), `desegment_ocr_input`, abbreviated sub-brand promotion (FIX 4/4b), `find_best_brand_for_input`. |
| 1729–2256 | Variant / sub-variant detection & filtering | `detect_variants_in_input`, `detect_sub_variants_in_input`, `filter_items_by_variant` (FIX 5 — most-specific-variant-wins), `filter_by_subvariant_in_product_name`. |
| 2256–2508 | Product-name scoring | `calculate_product_name_priority_score` (FIX 8 — structural pack/unit-suffix penalty), `filter_items_by_sub_variant` (FIX 6 — combo first-component fallback). |
| 2508–2680 | Dosage form | `detect_dosage_form_in_input` (FIX 7 — expanded OCR-safe tokens), `detect_dosage_form_in_product`, `prioritize_by_dosage_form`. |
| 2680–2736 | Pack size filtering | `filter_items_by_pack_size`. |
| 2736–2890 | Brand-specific override | `apply_galact_granules_rules` — hardcoded special case for brand GALACT (flavor/pack disambiguation). Documented exception to the data-driven principle; see Known Issues. |
| 2892–2919 | Reranker context builders | `build_brand_context_string`, `build_rerank_documents`. |
| 2919–2957 | Groq API call | `call_groq_llm`. |
| 2959–2996 | LLM answer → candidate matching | `find_best_match_with_fuzzy`. |
| 2996–3095 | Local ranking | `rank_local_candidates` (FIX 9 — ghost-token penalty), `choose_best_local_candidate`, `has_ambiguous_missing_strength`. |
| 3095–3260 | Suggestions / recovery | `TOP_N_SUGGESTIONS`, `NO_BRAND_RECOVERY_MAX_EDITS`, `nearest_brand_within_edits`, `make_result`, `suggest_nearest_brands`, `format_suggestions`. |
| 3258–3639 | Core orchestration | `process_product()` (~330 lines); `suggest_local_match()`; `review_highlighted_output_file()` (defined but **not wired into CLI menu**). |
| 3680–3826 | Batch entry point | `process_excel_file()` — resumable, auto-saves every 10 rows, writes a `_output.xlsx` copy, prints status/confidence summary. |
| 3826–3878 | Interactive entry point | `interactive_mode()` — REPL with verbose tracing plus `brands`/`stats` commands. |
| 3878–3895 | `__main__` | Menu: 1 = batch, 2 = interactive. |

#### `mapping.py` — key components

- **Normalization layer** — uppercase/punctuation/OCR cleanup shared by every stage (`norm`, `strip_parenthetical_noise`, `desegment_ocr_input`, `normalize_ocr_numeric_noise`).
- **Master loader** (`load_master`, `build_brand_product_map`) — reads `PRODUCT_MASTER.xlsx`, validates required columns, normalizes text, and builds `brand_map: dict[brand → list[item]]`, the in-memory structure everything else queries.
- **Brand resolver** (`find_best_brand_for_input` and helpers) — identifies which brand (and which of its SKUs) an input plausibly refers to, including typo recovery and abbreviated sub-brand promotion.
- **Variant / sub-variant resolver** — narrows the brand's SKUs by detected variant (e.g., MET, FORTE) and sub-variant (e.g., strength combos).
- **Dosage-form resolver** — prioritizes SKUs whose form (tablet/syrup/injection/etc.) matches OCR-safe tokens found in the input.
- **Pack-size resolver** — extracts a pack count/volume from the input and filters/prioritizes accordingly; falls back to the unpacked base SKU when no pack size is stated.
- **Local scorer/ranker** (`rank_local_candidates`) — final tiebreak among surviving candidates using structural penalties (ghost tokens, pack shapes, unit suffixes).
- **LLM reranker** (`call_groq_llm` + `find_best_match_with_fuzzy`) — only invoked when local ranking is not confident; the LLM's free-text answer is fuzzy-matched back onto a real candidate (never trusted verbatim).
- **Orchestrator** (`process_product`) — strings all of the above together per input row, with multiple recovery/fallback branches so ambiguous or over-filtered rows still return a best-effort LOW-confidence guess plus suggestions rather than nothing.
- **I/O drivers** — `process_excel_file` (batch), `interactive_mode` (REPL), `review_highlighted_output_file` (manual-review suggestion generator, currently unreachable from the menu).

#### `mapping.py` — data models & flow

Master item (one per SKU, inside `brand_map[brand]`):
```
{ product, pack_size, variant, sub_variant, product_code, product_norm }
```

Result dict (`make_result()` — canonical schema written back to Excel):
```
{ output, product_code, status, confidence, candidate_count, suggestions[] }
```

Status vocabulary: `MATCHED`, `NO_BRAND`, `RECOVERED_NO_BRAND`, `RECOVERED_LOWER_VARIANT`, `VARIANT_NOT_IN_MASTER`, `AMBIGUOUS_STRENGTH`, `NO_CANDIDATES`, `RECOVERED_NO_CANDIDATES`, `API_ERROR`, `RECOVERED_API_ERROR`, `LLM_REJECTED`, `RECOVERED_LLM_REJECTED`, `LLM_UNMATCHED`, `RECOVERED_LLM_UNMATCHED`, `SKIPPED_EMPTY`, `ERROR`.

Confidence vocabulary: `HIGH`, `MEDIUM`, `LOW`, `NONE`.

Per-row pipeline: raw text → clean/de-OCR → brand lookup → candidate list narrows through compound-token, variant, sub-variant, name-strength, dosage-form, pack-size filters (+ the GALACT override) → local rank → single/strong candidate returns immediately, else LLM rerank → fuzzy-match LLM answer to a real candidate → result dict → written to DataFrame → autosave every 10 rows → final save (+ a `_output.xlsx` copy).

#### `mapping.py` — Excel I/O schema

Input/output workbook columns (configurable via constants):

- `INPUT_COLUMN = "Input_Column"` (reads input product names)
- `OUTPUT_COLUMN = "Output_Column"` (matched canonical product name)
- `PRODUCT_CODE_COLUMN = "Product_Code"`
- `STATUS_COLUMN = "Match_Status"` (why a row matched or failed)
- `CONFIDENCE_COLUMN = "Confidence"` (HIGH / MEDIUM / LOW / NONE)
- `CANDIDATES_COLUMN = "Candidates_To_Model"` (count fed to model/ranker)
- `SUGGESTIONS_COLUMN = "Suggestions"` (top-N nearest products for manual review)

Master file columns read by `load_master()`:

- `PRODUCT_COL = "product_name"` (required)
- `BRAND_COL = "BRAND_NAME"` (required)
- `PACK_SIZE_COL = "PACK_SIZE"` (optional)
- `VARIANT_COL = "VARIANT"` (optional)
- `SUB_VARIANT_COL = "SUB_VARIANT"` (optional)
- `PRODUCT_CODE_COL = "PRODUCT_CODE"` (optional)

#### `mapping.py` — configuration surface

Key constants (all user-configurable at top of file):

- `GROQ_API_KEY`: reads from `GROQ_API_KEY` env var; falls back to the active API key in source.
- `GROQ_API_URL`: `https://hskyauefqcgbvgvxkluj.supabase.co/functions/v1/gonka/chat/completions`
- `MODEL_NAME`: `moonshotai/Kimi-K2.6`
- `MASTER_XLSX_PATH`: `f:\Vintyaa\projects\Product Name Automation\PRODUCT_MASTER.xlsx`
- `INPUT_OUTPUT_XLSX_PATH`: `f:\Vintyaa\projects\Product Name Automation\test.xlsx` (reads from sheet `garbage_check`, writes to sheet `mapping`)
- `ROW_LIMIT`: `30` (configurable batch limit for testing; set to `None` for all rows)
- `FUZZY_MIN_SCORE`: `80` (relaxed from 85 in FIX 10 to catch brand typos like ASOMAX→ASOMEX)
- `BRAND_FUZZY_MAX_EDITS`: `2`
- `MANUFACTURER_NOISE_NAMES`: `{"EMCOO"}` (OCR noise prefix for Emcure; hard-coded per master)
- `TOP_N_RERANK`: `3`; `REQUEST_TIMEOUT`: `240`; `DELAY_BETWEEN_PRODUCTS`: `2.0`
- `MAX_OUTPUT_TOKENS`: `512`

### `PRODUCT_MASTER.xlsx`

Canonical master product list used exclusively by `mapping.py`.

- Single sheet: `Sheet1`
- 2,766 rows (including header)
- Columns (6 total; `MAIN_BRAND` has been removed):

| Column | Role |
|---|---|
| `PRODUCT_CODE` | Unique SKU identifier (e.g., `420000001`) |
| `product_name` | Full canonical product name (e.g., `ABACUS - 50 TAB`) |
| `BRAND_NAME` | Brand key used to build `brand_map` in memory |
| `PACK_SIZE` | Pack size string (e.g., `30S`, `60T`); may be `None` for base/unpackaged variants |
| `VARIANT` | Variant token (e.g., `300 MG`, `PLAIN`, `L`) |
| `SUB_VARIANT` | Sub-variant token (e.g., `PLAIN`) |

The file is loaded once at startup by `load_master()` / `build_brand_product_map()`. All normalization happens in memory; the source file is never written to. Contains 874 unique `BRAND_NAME`s.

### `Brand Names.txt`

Single-line quoted brand-name master list used by `garbage_check.py` as the fuzzy-match source. Contains 874 medicine brand names (100% synchronized with the 874 unique brands in `PRODUCT_MASTER.xlsx`), including multi-word brands such as `BIC EMTAF`, `CALONAT XT`, `EMANZEN TBR`, `OROFER XT`, and `PRO FERIUM`.

The parser expects the file to remain in quoted-string format because `load_brands()` uses this regex:

```python
re.findall(r'"([^"]+)"', f.read())
```


### `test.xlsx`

Working input/output workbook for `garbage_check.py`. Also acts as the current manual test fixture.

Observed structure (updated 2026-08-19):

- **Sheet `garbage_check`** — where `garbage_check.py` reads and writes. Row 1 headers: `PRODUCT_NAME`, `Garbage check`, `Matched brand`, `Matched text`, `Match type`, `Matched garbage phrase`, `Garbage match type`
- **Sheet `mapping`** — where `mapping.py` writes results; starts empty and is populated by running `mapping.py` in batch mode.

`garbage_check` sheet column contract:

- Column A: OCR-extracted row text (input)
- Column B: classifier result (`0`, `1`, `BAD_OUTPUT`, `ERROR`, `REVIEW`, `MAYBE_PRODUCT`, `MAYBE_PRODUCT_2`, or blank) — written as **text string**, not a number
- Column C: locally matched brand name, or `"pattern-inferred"` for MAYBE_PRODUCT rows
- Column D: matched row fragment/text span, or `"form+qty match"` for MAYBE_PRODUCT rows
- Column E: match type (`exact`, `1-substitution`, `1-insertion`, `1-deletion`, `1-transposition`, `2-edit`, or `"MAYBE_PRODUCT"`)
- Column F: matched garbage phrase (when fuzzy-garbage logic fires)
- Column G: garbage match type

### `__pycache__/`

Generated Python bytecode cache. Not source of truth.

---

## Tech Stack

### `garbage_check.py`

- Python standard library: `difflib`, `re`, `time`
- `requests` for HTTP calls
- `openpyxl` for Excel workbook I/O
- External API: Novita OpenAI-compatible chat completions at `https://api.novita.ai/openai/v1/chat/completions`; model `nvidia/nemotron-3-nano-30b-a3b`

### `mapping.py`

- Python standard library: `os`, `re`, `sys`, `time`, `warnings`, `datetime`, `pathlib`
- `pandas` + `openpyxl` for Excel I/O
- `requests` for HTTP calls
- `rapidfuzz` (`fuzz.token_sort_ratio`, `fuzz.token_set_ratio`, `fuzz.ratio`, `Levenshtein.distance`) for fuzzy matching
- External API: OpenAI-compatible chat completions at `https://hskyauefqcgbvgvxkluj.supabase.co/functions/v1/gonka/chat/completions`; model `moonshotai/Kimi-K2.6`

Neither script has a `requirements.txt`, README, or dependency manifest.

---

## Architecture And Execution Flow

### `garbage_check.py` pipeline

Linear batch-processing pipeline:

1. Load brand names from `Brand Names.txt`.
2. Open `test.xlsx`; ensure column B header is `Garbage check`.
3. Optionally clear prior results if `CLEAR_EXISTING = True`.
4. Collect pending rows from column A, skipping blanks and already-classified rows.
5. Slice work by `ROW_LIMIT` and `BATCH_SIZE`.
6. For each batch:
   - Exact product match → if hit, write `0`, store diagnostics, skip LLM
   - Fuzzy product match → if hit, write `0`, store diagnostics, skip LLM
   - Exact garbage check → if hit, write `1`
   - Safe pattern garbage check → if hit, write `1`
   - Metadata vocabulary garbage check → if hit, write `1`
   - Fuzzy garbage check → if hit, write `1`, store matched phrase/type in F/G
   - MAYBE_PRODUCT check → if hit, write `MAYBE_PRODUCT`, store diagnostics in C/D/E
   - MAYBE_PRODUCT_2 check (planned) → if hit, write `MAYBE_PRODUCT_2`; wider signal: form OR qty (not both required), with guard filter
   - Remaining unresolved → `REVIEW` (current phase) or LLM fallback (when `USE_LLM_FALLBACK = True`)
   - Save workbook after every batch
7. Print cost statistics.

### `mapping.py` pipeline

Per-row deterministic funnel with forced brand scope, then LLM as last resort:

1. Load `PRODUCT_MASTER.xlsx` once; build `brand_map` in memory.
2. Read `test.xlsx` sheet `garbage_check` filtering rows where column B == `"0"`. Extract column A (input name) and column C (`Matched brand` hint as `forced_brand`).
3. For each input row:
   - Normalize / de-OCR input text
   - Strip manufacturer noise prefix (e.g., EMCOO)
   - Scope candidates strictly to `forced_brand` from column C:
     - Exact match in `brand_map` or case-insensitive match -> use brand candidates
     - Unmatched / missing hint -> return `NO_CLEAR_MATCH` (`NO_BRAND`) immediately (auto-detection disabled)
   - Narrow brand candidate SKUs by compound tokens → variant (most-specific-wins, FIX 5) → sub-variant (combo fallback, FIX 6) → dosage form (FIX 7) → pack size (FIX 1 regex-order fix) → local rank (ghost-token penalty, FIX 9)
   - Apply GALACT-granules hardcoded override if brand is GALACT
   - If single/strong local candidate: return HIGH/MEDIUM confidence immediately (no LLM call)
   - Else: invoke LLM reranker (`moonshotai/Kimi-K2.6`) with a short-listed candidate set; fuzzy-match its free-text answer back to a real candidate
   - Produce result dict with status, confidence, candidate count, and top-N suggestions
4. Write results back to `test.xlsx` sheet `mapping` using `_save_to_mapping_sheet()`:
   - Initialized with `None` (`object` dtype) to support mixed string/int columns
   - Autosaves every 10 rows without modifying the `garbage_check` sheet
   - Final save to sheet `mapping`.

---

## Conventions

### Shared conventions

- Constant-driven configuration near the top of each file (ALL-CAPS naming, often with inline rationale comments)
- Snake_case function naming; imperative orchestration in a main entry-point function
- All progress/diagnostics via `print()` — no logging module
- Broad `try/except` around external I/O; per-row errors → status cell, never crash the batch
- No formal automated test suite; validation is manual via workbook inspection

### `garbage_check.py`-specific

- Workbook saved after each batch to reduce data loss risk
- Classification can resume only when `CLEAR_EXISTING = False`
- `CLEAR_EXISTING = True` makes reruns destructive for all columns except column A
- `MAYBE_GARBAGE` label no longer exists; all patterns merged into `SAFE_GARBAGE_PATTERNS` on 2026-08-18

### `mapping.py`-specific

- Each behavioral fix gets a numbered `FIX N` comment at its point of use and a summary in the top-of-file changelog block — the project's existing convention for documenting decisions
- Function naming prefix convention: `extract_`, `detect_`, `filter_`, `calculate_`, `build_`, `strip_`; private helpers prefixed with `_`
- Non-atomic autosave: overwrites the source workbook every 10 rows (no temp-file-then-rename)
- Global mutable counters (`REQ_COUNT`, `SUM_PROMPT_TOKENS`, etc.) are module-level; fine for single-threaded CLI use

---

## Known Issues, Fragile Areas, And Tech Debt

### Security (both scripts)

- `garbage_check.py`: `API_KEY` is hardcoded in source. Highest-priority issue.
- `mapping.py`: `GROQ_API_KEY` falls back to a literal `gsk_...` key baked into the file. **Treat as a live credential; rotate in the Groq console immediately.**
- Both secrets should move to environment variables or a local config file excluded from version control.

### Reliability / Accuracy — `garbage_check.py`

- Model output parsing is brittle; real rows have already landed in `BAD_OUTPUT` (rows 49–51: HDSIT...)
- Depends on prompt compliance rather than structured output or JSON mode
- `SequenceMatcher` over every brand for every row is potentially slow as lists grow
- Tokenization may miss some OCR patterns (only extracts alphabetic substrings of length >= 3)

### Reliability / Accuracy — `mapping.py`

- Hardcoded absolute Windows paths (`MASTER_XLSX_PATH`, `INPUT_OUTPUT_XLSX_PATH`) — not portable without editing source
- Brand-specific hardcoded override (`apply_galact_granules_rules`) for brand GALACT contradicts the stated data-driven design principle
- `MANUFACTURER_NOISE_NAMES = {"EMCOO"}` — single-master hardcode; silently does nothing if reused against a different master
- `review_highlighted_output_file()` is defined but unreachable from the CLI menu — dead code or a missing menu option
- No automated tests for a highly branchy pipeline (`process_product` is ~330 lines with dozens of conditional branches)
- `_lower_variant_retry` boolean flag caps recursion at one level — a fragile pattern if more retry layers are added later
- Non-atomic autosave in batch mode: overwrites the source file every 10 rows without a temp-file-then-rename

### Product / Workflow (both scripts)

- Workbook filenames hardcoded in both scripts
- No separation between library code and CLI/runtime entry point
- No requirements file, README, or setup documentation
- No logging framework beyond `print()`

### Repository hygiene

- The folder was not a Git repository when inspected on 2026-08-17
- `__pycache__` present in working tree; no `.gitignore`, no `requirements.txt`

---

## `garbage_check.py` — Configuration Surface

Top-level constants:

- `API_KEY`, `NOVITA_URL`, `MODEL`
- `EXCEL_FILE`, `BRANDS_FILE`
- `BATCH_SIZE` (current: `2`), `BATCH_WAIT`
- `ROW_LIMIT` (current: `190000`, effectively unlimited)
- `MAX_RETRIES`, `RETRY_WAIT`
- `CLEAR_EXISTING` (current: `True`)
- `USE_LLM_FALLBACK` (current: `False` during local-tuning phase)
- `PRICE_INPUT_PER_M`, `PRICE_OUTPUT_PER_M`
- `TOP_CANDIDATES` (current: `3`), `MIN_SHOW_SCORE` (current: `0.60`)
- `METADATA_GARBAGE_VOCAB`: ~70 document-structure/metadata words (medicine-context tokens deliberately excluded)
- `MAYBE_PRODUCT_FORM_PATTERN`: compiled regex for dosage-form tokens (TAB, CAP, INJ, CREAM, VIAL, SYP, OINT, LOTION, PFS, etc.)
- `MAYBE_PRODUCT_QTY_PATTERN`: compiled regex for qty/strength signals (MG, MCG, IU, ML, GM, NxN, NS, NTAB, NVIAL)
- `MAYBE_PRODUCT_GUARD_PATTERN`: compiled regex for structural-garbage keywords (TOTAL, SALE, PURCHASE, INVOICE, STOCK, VALUE, REPORT, SUMMARY, IMPORT, CALL)
- `MAYBE_GARBAGE_PATTERNS` has been **removed** (2026-08-18); all 14 patterns merged into `SAFE_GARBAGE_PATTERNS`

---

## Notable Decisions Log

### 2026-08-17

- A persistent project memory file (`context.md`) was introduced in the repo root.
- `context.md` should not be treated as a mandatory startup ritual — only read when additional project context is actually needed.
- Future code changes should be discussed before implementation, with trade-offs laid out first, per the user's workflow rule.
- Primary optimization target for `garbage_check.py`: avoid false negatives above all else. A real product row should not be classified as garbage.
- Matching should not assume the brand is always the first word; OCR shifts may place the brand token later in the row.
- Preferred future pipeline is recall-first hybrid: local fuzzy logic scans the full row for brand evidence; rows confidently identified locally bypass the LLM; only uncertain rows go to the LLM.
- Local acceptance rule: at most one character edit allowed; spaces ignored during comparison.
- Brand confirmation must search across the whole row using consecutive token spans.
- 3-letter brand safeguard approved: exact standalone token match only for brands of 3 normalized characters.
- Agreed local-classification policy:
  - Exact product match first → fuzzy product match → garbage (exact → safe pattern → metadata vocabulary) → fuzzy garbage → `MAYBE_PRODUCT` → `REVIEW`
  - For brands 4–9 chars: `<= 1` edit (substitution, insertion, deletion, adjacent transposition via Damerau-Levenshtein)
  - For brands `>= 10` chars: `<= 2` edits via Damerau-Levenshtein; 2-edit matches labelled `2-edit`; threshold raised from 7 to 10 after false positives on mid-length words (e.g., MEDICAL matching TREDICAL)
  - Brands excluded from auto-acceptance: `EFCURE` (likely EMCURE OCR collision), `CTAX` (spurious match against TAX tokens)
- Large explicit-match garbage phrase list supplied by user. Product confirmation takes precedence over garbage matching when there is an exact local-product hit.
- Any row not decisively confirmed as product or garbage should be escalated to the LLM.

### 2026-08-17 (session 2)

- Added `CTAX` to `LOCAL_REVIEW_BRANDS` alongside `EFCURE`.
- Replaced plain Levenshtein with Damerau-Levenshtein (`damerau_distance()`); adjacent transpositions now count as one edit.
- Extended adaptive edit-distance threshold; raised from 7 to 10 chars for 2-edit range.
- Added `METADATA_GARBAGE_VOCAB` (~70 words) and `is_metadata_garbage_row()`; medicine-context tokens deliberately excluded.
- Updated garbage check order: exact phrase → safe regex → metadata vocabulary.
- Added per-row early-exit to `damerau_distance` DP; eliminates 50–80% of DP work for non-matching pairs.
- Added character-bag pre-filter in `fuzzy_product_match`; eliminates ~90% of `damerau_distance` calls.
- Promoted fuzzy garbage matches to direct `1` (was MAYBE_GARBAGE).
- Expanded `SAFE_GARBAGE_PATTERNS` from ~5 to ~50 patterns.
- Added `MAYBE_GARBAGE_PATTERNS` (4 patterns, later grew to 14) and `is_maybe_garbage_row()`.
- Updated `main()` pipeline: exact → safe pattern → metadata vocab → maybe-garbage → fuzzy garbage → REVIEW/LLM.
- Various pattern additions and `EXACT_GARBAGE_ROWS` expanded with 50+ confirmed company/distributor names.

### 2026-08-17 (session 2 — performance + accuracy fixes)

- Raised 2-edit adaptive threshold from `>= 7` to `>= 10` chars to reduce false positives on mid-length common words.
- Added per-row early-exit to `damerau_distance` DP.
- Added character-bag pre-filter in `fuzzy_product_match`.
- Promoted fuzzy garbage from MAYBE_GARBAGE to direct `1`; matched phrase and match type still written to F/G for auditability.
- Analyzed 7,397 REVIEW rows; found major categories: Bill Nos, Division/Company/Supplier headers, PURC/EXP financial rows, MediVision software rows, HTML artifacts, report footers.
- Expanded `SAFE_GARBAGE_PATTERNS` to ~50 patterns.
- Added `MAYBE_GARBAGE_PATTERNS` (4 patterns) and `is_maybe_garbage_row()`.
- Updated `main()` pipeline to include `is_maybe_garbage_row` step.
- Removed `.*\s--$` pattern per user confirmation that product rows can also end with ` --`.
- Various pattern additions and expansions.
- Expanded `MAYBE_GARBAGE_PATTERNS` from 2 to 14 patterns.

### 2026-08-18

- Eliminated `MAYBE_GARBAGE` as a concept entirely: `MAYBE_GARBAGE_PATTERNS` and `is_maybe_garbage_row()` removed. All 14 patterns merged into `SAFE_GARBAGE_PATTERNS`; they now classify as `1` directly.
- `main()` pipeline simplified: separate `is_maybe_garbage_row` branch removed.
- Added `MAYBE_PRODUCT` tier: rows with no brand match and no garbage signal that contain a dosage-form token AND a qty/strength signal are labelled `MAYBE_PRODUCT`.
- Added three new constants: `MAYBE_PRODUCT_FORM_PATTERN`, `MAYBE_PRODUCT_QTY_PATTERN`, `MAYBE_PRODUCT_GUARD_PATTERN`.
- Added `is_maybe_product_row()` implementing the three-signal rule.
- Guard pattern excludes risky rows (e.g., `TOTAL CAL D3 TAB 1X20`) that match form+qty but also contain garbage keywords.
- Expected impact: ~2,511 of 6,231 REVIEW rows promoted to `MAYBE_PRODUCT`.
- Added `^Total\s+Value\s*\(` to `SAFE_GARBAGE_PATTERNS`.
- `context_mapping.md` merged into this file and deleted (documentation consolidation only; no code changes).

### Pre-existing code decisions inferred from the repository

`garbage_check.py`:
- LLM used as decision engine rather than local rules only; fuzzy brand matching is a prompt-enrichment step, not the final classifier
- Batch size deliberately kept small (`2`), suggesting prompt quality or rate-limit sensitivity has been a concern
- Workbook saved after every batch, indicating a preference for crash resilience over throughput

`mapping.py` (`updated_ts_5`):
- 10 fixes integrated over `updated_ts_3`: pack-size regex ordering (biggest fix — affected ~279 of 929 red rows), parenthetical-noise stripping, compact-prefix brand-boundary fix, abbreviated sub-brand promotion (+ product-name fallback), most-specific-variant-wins, combo first-component fallback, expanded OCR dosage-form tokens, structural-only priority penalty, ghost-token penalty extended to product names, fuzzy threshold relaxed 85 → 80
- Design principle: keep matching data-driven — derived from the master file's own brand/variant/sub-variant/product-name data. (The GALACT-granules rule is a documented exception.)
- LLM as tie-breaker only, not primary decision-maker; confident single-candidate rows never call the LLM — a deliberate cost/latency/determinism tradeoff
- Confidence tagging + Suggestions column: uncertain rows flagged for manual review rather than silently guessed

---

## Open Questions

### `garbage_check.py`

- Should `BAD_OUTPUT` rows be retried automatically with a stricter or alternate parsing strategy?
- Should the model be asked for structured JSON output instead of free-form reasoning lines?
- Should brand matching move to a faster or more domain-aware fuzzy matcher as the brand list grows?
- Is `test.xlsx` intended as a disposable fixture, or is it also the operational production input?
- Are there benchmark labels available for measuring precision/recall?

### `mapping.py`

- Should `review_highlighted_output_file()` be wired into the CLI (e.g., menu option 3), or was it superseded?
- Should the GALACT-granules special case be generalized into a data-driven equivalent?
- Is this meant to stay a single portable script, or evolve into a proper package?
- Rotate the hardcoded Groq key now? (flagged as urgent — live credential in source)

### Both scripts

- Should there be a shared config file (`.env` or similar) so file paths and API keys are managed in one place?
- Should the project split into a proper package structure with modules, requirements, and tests?
- Should the workbook schema become explicit and versioned if more columns are added later?

---

## Suggested Areas To Revisit First

- **Immediate**: Rotate the Groq API key hardcoded in `mapping.py` and move it to an env var with no inline fallback
- **Immediate**: Move `garbage_check.py`'s `API_KEY` to an env var
- **High**: Structured output / parser hardening in `garbage_check.py` to reduce `BAD_OUTPUT`
- **High**: Replace hardcoded absolute paths in `mapping.py` with CLI args or a `.env` file
- **Medium**: Resumability defaults and safer rerun behavior in `garbage_check.py`
- **Medium**: Separate sample data from live output files in both scripts
- **Medium**: Add minimal documentation and a `requirements.txt`
- **Low**: Wire `review_highlighted_output_file()` into the `mapping.py` CLI menu or remove as dead code

---

## Changelog

### 2026-08-17

- Created `context.md` with the initial full-project audit and operating notes (covers `garbage_check.py` only at that point).
- Recorded the recall-first hybrid-classification direction.
- Updated `garbage_check.py` to implement the hybrid pipeline.
- Simplified the LLM prompt/output contract.
- Added the 3-letter brand safeguard.
- Replaced arbitrary substring matching with full-row consecutive token-span matching.
- Added local match diagnostics; added brand-specific review exceptions (`EFCURE`, `CTAX`).
- Added local-only testing mode (`USE_LLM_FALLBACK = False`).
- Added workbook audit columns for local matching.
- Updated reset behavior to wipe every column except A.
- Expanded garbage handling (exact rows, safe patterns, metadata vocab).
- Split local product detection into two passes.
- Added Damerau-Levenshtein; extended adaptive edit-distance threshold.
- Added per-row early-exit to `damerau_distance` DP; character-bag pre-filter.
- Expanded `SAFE_GARBAGE_PATTERNS` to ~50 patterns; expanded `EXACT_GARBAGE_ROWS` with 50+ confirmed names.
- Added `MAYBE_GARBAGE_PATTERNS` (4 → 14 patterns) and `is_maybe_garbage_row()`.

### 2026-08-18

- Eliminated `MAYBE_GARBAGE`: merged all 14 patterns into `SAFE_GARBAGE_PATTERNS`.
- Added `MAYBE_PRODUCT` tier with three new constants and `is_maybe_product_row()`.
- Added `^Total\s+Value\s*\(` to `SAFE_GARBAGE_PATTERNS`.
- **`context.md` updated**: documented `mapping.py`, `PRODUCT_MASTER.xlsx`; merged `context_mapping.md` in full. `context_mapping.md` deleted as redundant after this merge.

### 2026-08-19

- Separated `test.xlsx` into two dedicated sheets: `garbage_check` and `mapping`.
- Fixed `garbage_check.py` to explicitly target `wb["garbage_check"]` by name instead of `wb.active`.
- `mapping.py` connected directly to `test.xlsx` to process only confirmed `0` rows from sheet `garbage_check`.
- Added `_save_to_mapping_sheet()` in `mapping.py` using `openpyxl` to safely rewrite only the `mapping` sheet while preserving `garbage_check`.
- Fixed output dataframe column initialization to use `None` (`object` dtype) instead of `""` to prevent pandas `StringDtype` casting errors when setting integer candidate counts.
- Added `ROW_LIMIT = 30` config constant in `mapping.py` for testing batch slices (set to `None` for full runs).
- Passed column C (`Matched brand`) from `garbage_check` to `mapping.py` as `forced_brand` to restrict product mapping candidates to the confirmed brand only.
- Disabled auto-detection fallback in `mapping.py`: unmatched forced brands return `NO_CLEAR_MATCH` directly.
- Added brand-specific false-positive rules in `garbage_check.py`:
  - `CELOL` with `DINNER` or `SET`: classified directly as garbage (`1`).
  - `TAMLET` matched against dosage token `TABLET`/`TABLETS`: invalidated so the row falls through to standard garbage/review checks.
  - `IMPETUS`, `VINTOR`, `EMNU` with `\bCOMPANY\b` (word boundary, case-insensitive): invalidated — these are company/distributor header lines, not product rows. (`EMCU` was in this set but has been removed — brand no longer exists in master.)
  - `AMARYL` with `\bSEMI\b`: invalidated — SEMI AMARYL does not exist in the product master.
  - `EMNU` with `\bMANUFACTURER\b`: invalidated — manufacturer header lines are not product rows.
- Changed handling of `EFCURE`, `KEMCURE`, and `CTAX`:
  - Removed all three from `LOCAL_REVIEW_BRANDS` (now empty `set()`).
  - Added to `EXACT_ONLY_BRANDS = {"EFCURE", "KEMCURE", "CTAX"}`: fuzzy matches invalidated, only exact span matches auto-confirmed as `0`.
- Normalized all brand exception checks in `garbage_check.py` using `normalize_text(local_match["brand"])`, ensuring brands containing spaces/hyphens in `Brand Names.txt` (such as `"C - TAX"`) match their respective rules without string formatting mismatch.

- Configured LLM reranker in `mapping.py` to use `moonshotai/Kimi-K2.6` via the OpenAI-compatible gateway at `https://hskyauefqcgbvgvxkluj.supabase.co/functions/v1/gonka/chat/completions`.
- Installed `rapidfuzz` dependency in the local environment.

### 2026-08-19 (afternoon session)

- Removed `EMCU` from `_COMPANY_BRANDS` exception set in `garbage_check.py` — brand no longer exists in `PRODUCT_MASTER.xlsx`.
- Added `NUNIT` to `EXACT_ONLY_BRANDS` — fuzzy hits risk false positives against common tokens like `UNIT`; only exact span matches accepted.
- Added brand-specific whitelist/invalidation exceptions in `garbage_check.py`:
  - `ZUVENTUS`: if matched row does **not** contain `ORS`, classify directly as `1` (garbage) without further processing. Only real product is `ZUVENTUS ORS ORANGE 21 GM`. All other ZUVENTUS rows (Division headers, Lifestyle, Athena, Gromaxx, Total Value lines) are company/division lines.
  - `NEW`: if matched row does **not** contain `NORMIT`, invalidate the match and fall through to standard garbage/review pipeline.
  - `VITAMIN`: if matched row does **not** contain `D3`, invalidate the match and fall through to standard garbage/review pipeline.
  - `EMCOR`: if matched row does **not** contain `CREAM` or `TUBE`, invalidate the match and fall through to standard garbage/review pipeline.
- Analysed 6,086 unresolved rows (3,570 `REVIEW` + 2,516 `MAYBE_PRODUCT`); identified safe zero-risk garbage patterns.
- Added to `SAFE_GARBAGE_PATTERNS` (~470 rows resolved without LLM):
  - `^MARG\s+ERP\b` — software advertisement rows (362 occurrences with varying phone numbers at end)
  - `\bSANOFI\b` — Sanofi is not an Emcure brand; all rows are company headers / annexure labels (87 rows)
  - `^\s*(TAB(LE)?\s+)*TAB(LE)?\s*$` — rows consisting entirely of repeated TAB/TABLE OCR garbage (6 rows)
  - `\b(BEDSHEET|BEDSEET|HOTPOT|TOWEL|DUFFAL\s+BAG|ELETRAL|WATER\s+JAR)\b` — non-pharma physical items (12 rows)
  - `^Item\s+Name\b` — column-header artifact from OCR of table headers (2 rows)
  - `\bSOFTWARE\s+LICEN` — software licence rows (1 row)
- Added to `EXACT_GARBAGE_ROWS`: `"DEFAULT"`, `"All Marketing Groups"`.
- Added short-row garbage pattern to `SAFE_GARBAGE_PATTERNS`: rows with ≤2 alphanumeric characters (e.g. `hr/`, `CV`, `AV`, `HO`, `PM --`, `MG`, `GM`) — 80 rows. Safe because `MAYBE_PRODUCT` check fires before pattern garbage, so real strength-token fragments from split OCR rows are handled first.
- Fixed pattern gaps identified from REVIEW screenshot:
  - `DUFFAL\s+BAG` → `\bDUFFAL\b` (standalone word match): `BAG22MED` suffix was breaking the word boundary after BAG
  - Added `CHOCOLATE` to non-pharma physical items pattern
  - Added `\bSTELL?\s+BOW[AL]+\b` — OCR corruption of STEEL BOWL / STELL BOWAL
  - Added `\bSTEEL\s+GLASS\b` — non-pharma item
  - Added `^Receipt\s+Value\b` — receipt report lines
  - Added `^Near\s+Expiry\s+Product\b` — near-expiry scheme lines
  - Added `\bMEDICAL\s+STOR\b` — medical store reference/receipt lines (e.g. SAHIL MEDICAL STOR FEROZEPU A001335)
  - Raised short-row threshold from ≤2 to ≤3 alphanumeric characters: catches EMK, EMO, EMX, ESPIN-like fragments; safe because brand matching runs before pattern garbage check
  - Added all 12 month names (`JAN`–`DEC`) to `METADATA_GARBAGE_VOCAB` — fixes `LAST MONTH SALE MAY` not being caught by metadata garbage check
- Designed and prototyped `MAYBE_PRODUCT_2` classification tier:
  - Wider than `MAYBE_PRODUCT`: requires form signal **OR** qty signal (not both); guard filter still applied.
  - Also widens qty pattern to accept `*` as pack separator (catches `TICAGRELOR 1*10`).
  - Implemented as standalone script `maybe_product_2_check.py` for manual validation before pipeline integration.
  - Script writes candidates to sheet `maybe_product_2` in `test.xlsx` with `Signal_Type` column (`form-only` or `qty-only`).
  - From 2,983 current REVIEW rows: 1,752 qualify as MAYBE_PRODUCT_2 (298 form-only + 1,454 qty-only); 1,231 remain pure REVIEW with no pharma signal.
  - Next step: user reviews `maybe_product_2` sheet; if results look good, logic will be integrated into `garbage_check.py` main pipeline.

### 2026-08-20 (afternoon session)

#### `mapping.py` — MAYBE_PRODUCT brand detection integration

- Added `SOURCE_COLUMN = "Source"` constant — written to `mapping` sheet to indicate whether a row came from a confirmed `"0"` classification or a `"MAYBE_PRODUCT"` classification.
- Added `JW_MIN_SCORE = 0.85` constant — Jaro-Winkler minimum similarity threshold (0.0–1.0 scale via `rapidfuzz.distance.JaroWinkler.normalized_similarity`).
- Added `PROCESS_MAYBE_PRODUCT_ONLY = False` constant (was temporarily set to `True` for a MAYBE_PRODUCT-only test run; reverted to `False` so both `"0"` and `MAYBE_PRODUCT` rows are processed together in all future runs).
- Imported `JaroWinkler` from `rapidfuzz.distance` (fix for `AttributeError: module 'rapidfuzz.fuzz' has no attribute 'jaro_winkler_similarity'` — that function does not exist in `rapidfuzz.fuzz`; the correct API is `JaroWinkler.normalized_similarity` from `rapidfuzz.distance`).
- `find_best_brand_for_input()` — added Jaro-Winkler as a 3rd fallback scorer after `fuzz.ratio` (99 cutoff) and `fuzz.token_sort_ratio` (80 cutoff) both fail. JW runs over `[brand_query_norm, brand_root_norm, *brand_tokens_norm]` at 0.85 cutoff. Tagged as `best_source = "jaro-winkler"` for auditability.
- `process_product()` — added `auto_detect: bool = False` parameter:
  - `forced_brand=None, auto_detect=True` → calls `find_best_brand_for_input()` (MAYBE_PRODUCT path)
  - `forced_brand=None, auto_detect=False` → returns `NO_CLEAR_MATCH` immediately (unchanged `"0"` row behavior)
  - `forced_brand` supplied but not in `brand_map`, `auto_detect=True` → also falls back to auto-detect
  - Updated Step 1 comment block to reflect new branching logic; removed stale "AUTO-DETECT fallback (disabled)" comments.
- `process_excel_file()` — restructured reading section:
  - Reads both `"0"` rows (forced brand from col C, `_auto_detect=False`) and `"MAYBE_PRODUCT"` rows (no brand hint, `_auto_detect=True`) from `garbage_check` sheet.
  - Initialises `df = pd.DataFrame()` before `try:` block to prevent `UnboundLocalError` if an exception fires before `df` is assigned.
  - When `PROCESS_MAYBE_PRODUCT_ONLY=True`: only `df_mp` is used; `df_0` is skipped and a message is printed.
  - When `PROCESS_MAYBE_PRODUCT_ONLY=False`: `pd.concat([df_0, df_mp])` — confirmed rows first, then MAYBE_PRODUCT.
  - `ROW_LIMIT` applies to the combined (or filtered) total.
  - `Source` column written per row (`"0"` or `"MAYBE_PRODUCT"`).
  - Internal helper columns (`_brand_hint`, `_auto_detect`, `_source_label`) stripped before every save (autosave + final save) via `df.drop(columns=[c for c in df.columns if c.startswith("_")])`.
- `_save_to_mapping_sheet()` behavior: deletes and fully recreates the `mapping` sheet on every save. Any previously saved `"0"` row results are lost when a MAYBE_PRODUCT-only run saves. Keep this in mind when switching `PROCESS_MAYBE_PRODUCT_ONLY` between runs.

#### Indentation bugs fixed (introduced by manual user edits)

- `_exact_pack_matches()` nested inside `apply_galact_granules_rules()`: `try:` was at 4-space indent (function level) instead of 8-space (inside nested function). Fixed to 8 spaces.
- `process_excel_file()` reading section: `print()` and `df = pd.DataFrame()` before the `try:` block were at 0-space indent (module level), breaking the function scope and making the `except` unmatchable. Fixed to 4-space indent.
- `process_excel_file()` processing loop: `try:` and `except Exception as e:` were at 4-space indent (function level) instead of 8-space (inside the `for` loop). This caused all rows to be skipped by the for loop, and only the last row to be processed once after the loop ended (since `idx`/`row` remained in scope). Fixed to 8 spaces.
