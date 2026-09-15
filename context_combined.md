# context_combined.md — MAPPING_MAIN.py

Detailed context for `MAPPING_MAIN.py`, the combined garbage-check + product-mapping
pipeline. `context.md` remains the context file for `garbage_check.py` and `mapping.py`;
it carries a single pointer to this file.

Everything below is derived from the code as it stands in `MAPPING_MAIN.py`
(4,493 lines) and from the validation run recorded at the end.

---

## 1. Purpose

Produce, in one pass over the input workbook, a single three-column answer per input row:
the raw OCR text, a remark, and a product code. It replaces the previous two-step manual
workflow (run `garbage_check.py`, then run `mapping.py` over the rows it marked `0`) with
one script and one output file, without changing either existing script and without
changing the classification or mapping logic they implement.

## 2. Architecture

Self-contained single module. `garbage_check.py` and `mapping.py` are **not imported**;
their functions are transplanted verbatim so that `MAPPING_MAIN.py` can be run and
versioned independently and cannot be broken by an edit to either source file.

Layout (line numbers as built):

| Lines | Section |
|---|---|
| 1-41 | module docstring: flow, provenance, the two deliberate divergences |
| 43-59 | union of both files' imports |
| 61-122 | `CONFIGURATION` — paths, `PROCESSING_MODE`, `llm_enabled()`, `MAPPING_LLM_PROMPT`, sentinels |
| 125-857 | `SECTION 1` — garbage-check stage, verbatim from `garbage_check.py` |
| 860-863 | `SECTION 2` header — mapping stage |
| 866-929 | `RERANKER TRANSPORT` — `call_groq_llm` commented out + an active raising stub |
| 932-4198 | mapping stage, verbatim from `mapping.py` (two documented splices) |
| 4202-4313 | `classify_row()` — stage-1 driver |
| 4315-4372 | `RERANKER_DEPENDENT_STATUSES`, `decide_output()`, `map_one()` |
| 4375-4493 | `build_output_workbook()`, `main()`, `__main__` guard |

What was carried over, by construction (line-range transplant, so interleaved comments
survive):

* from `garbage_check.py` (1,231 lines): kept 729, dropped 502 — dropped ranges `1-88`
  (header + all LLM/API configuration, including the API key) and `818-1231`
  (`brand_candidates`, `format_row`, `call_model`, `classify_batch`, `has_result`, the
  `write_*`/`clear_*` cell writers, its `llm_enabled`, `main`, `__main__` guard).
* from `mapping.py` (4,201 lines): kept 3,255, dropped 946 — dropped ranges `1-138`
  (header, imports, credential block, model/path/`ROW_LIMIT`/`PROCESSING_MODE` config),
  `1142-1237` + `1555-1584` + `1664-1835` (its own brand matching), `3051-3106`
  (`llm_enabled`, `call_groq_llm`), `3253-3273` + `3303-3336` (nearest-brand recovery),
  `3803-4201` (`suggest_local_match`, `review_highlighted_output_file`,
  `_save_to_mapping_sheet`, `format_row_result`, `process_excel_file`,
  `interactive_mode`, `__main__` guard).

Exactly four top-level names collide between the two files — `PROCESSING_MODE`,
`VALID_PROCESSING_MODES`, `ROW_LIMIT`, `llm_enabled`. Each is declared once in the
`CONFIGURATION` block and excluded from both transplants. A duplicate-definition scan and
a `symtable`-based undefined-global scan both run clean on the built file.

## 3. End-to-end pipeline

```
main()
  load_brands() -> build_brand_index()      # garbage-check brand vocabulary
  build_garbage_set()                       # garbage phrase set
  load_master() -> build_brand_product_map()
  extract_all_variants_from_data() / extract_all_sub_variants_from_data()
  read column A of INPUT_SHEET_NAME (read-only)

  for each input row:
      blank                -> emit ("", "", "")            # not classified
      classify_row(...)
          "GARBAGE"        -> GARBAGE       / 1     (stop)
          "REVIEW"         -> REVIEW        / 2     (stop)
          "MAYBE_PRODUCT"  -> NO_SUGGESTION / 0     (stop; mapping is not entered)
          "PRODUCT"        -> remark/code start at NO_SUGGESTION / 0
                              map_one() -> process_product(forced_brand=<brand>,
                                                           auto_detect=False,
                                                           client=None)
                              decide_output() -> overwrite on local success only

  build_output_workbook()  -> MAPPING_MAIN_OUTPUT.xlsx
```

`map_one()` is wrapped in a `try/except`: a row that raises keeps `NO_SUGGESTION / 0` and
the run continues. An exception can never produce a product code.

## 4. Input / output workbook behavior

* Input `test1.xlsx`, sheet `garbage_check`, is opened with
  `load_workbook(..., read_only=True, data_only=True)`; only column A is read and the
  handle is closed before any writing. **The input workbook is never written to.**
* Output `MAPPING_MAIN_OUTPUT.xlsx` is a brand-new `Workbook()` written at the end of the
  run. It is not a copy of the input and carries none of the intermediate columns B-G that
  `garbage_check.py` writes into its own workbook.
* Output row *N* corresponds to input row *N*. Blank input rows are emitted as a blank
  triple rather than skipped, so that alignment holds; this mirrors `garbage_check.py`,
  which skips blank rows and leaves them unclassified rather than calling them garbage.
* `ROW_LIMIT = None` processes every row. Setting an integer truncates from the top of the
  sheet, as in both source scripts.

## 5. Three-column output schema

Sheet `mapping_main`, header row `Input | Remark | Product Code`, no other sheet and no
other column. Product codes are written as integers when the master value is all digits,
so the column is numeric throughout; the sentinels `0`, `1` and `2` are integers too.
Real master codes are 9 digits (`410000003`-`491110205`), so no sentinel can ever collide
with a genuine product code.

## 6. Garbage-check stage

Fully local. No LLM call exists in this stage and none can be introduced by configuration:
the classifier, the batching helpers and the API configuration of `garbage_check.py` were
all excluded from the transplant, so the only code present is the local matcher.

`classify_row(name, brand_index, garbage_rows)` is `garbage_check.main()`'s per-row body
with the workbook writes replaced by a return value. It returns
`("PRODUCT", brand) | ("GARBAGE", "") | ("MAYBE_PRODUCT", "") | ("REVIEW", brand_or_"")`.

Order preserved exactly:

1. `token_spans` -> `exact_product_match` -> `fuzzy_product_match` (fuzzy only if no exact).
2. Brand-specific exceptions, in the original sequence:
   * `CELOL` + `MARKER|DINNER|SETS?` -> garbage immediately.
   * `EXACT_ONLY_BRANDS` (`EFCURE`, `KEMCURE`, `CTAX`, `NUNIT`, `EMNU`, `ACEM`) —
     non-`exact` match invalidated.
   * `TAMLET` matching the dosage word `TABLET`/`TABLETS` — invalidated.
   * `_COMPANY_BRANDS` (`IMPETUS`, `VINTOR`, `EMNU`) + `COMPANY` — invalidated.
   * `AMARYL` + `SEMI` — invalidated (`SEMI AMARYL` is not in the master).
   * `EMNU` + `MANUFACTURER` — invalidated.
   * `ZUVENTUS` without `ORS` -> garbage immediately (only real product is ORS ORANGE).
   * `NEW` without `NORMET` — invalidated.
   * `VITAMIN` without `D3` — invalidated.
   * `EMCOR` without `CREAM|TUBE` — invalidated.
   * `NUMLO` with an `S` prefix in the normalised row -> brand remapped to `S-NUMLO`.
3. Terminal branches: confirmed match (and brand not in `LOCAL_REVIEW_BRANDS`) -> product;
   else exact/pattern/metadata garbage -> garbage; else fuzzy garbage -> garbage; else a
   surviving local match -> review; else `is_maybe_product_row()` -> maybe product, or
   review.

`LOCAL_REVIEW_BRANDS` is empty, exactly as in the source, so today no confirmed match is
diverted to review. The `S-NUMLO` remap mutates a copy of the match dict rather than the
matcher's own dict; the classification outcome is identical.

Only rows classified `PRODUCT` continue into the mapping stage. `MAYBE_PRODUCT` rows do
**not** enter mapping — under the previous two-step workflow they could be mapped with
`auto_detect=True`, and that path is intentionally gone (see §8).

## 7. Mapping stage

`process_product()` and every function it depends on are transplanted verbatim. Preserved
in full, unchanged: input cleaning (`clean_duplicate_words`, `strip_parenthetical_noise`,
`strip_manufacturer_noise`, `desegment_ocr_input`), normalisation and tokenisation,
candidate generation, the compound-token pre-filter, variant handling
(`filter_items_by_variant` and its `PLAIN` fallback), sub-variant handling, name-strength
matching, dosage-form handling (`prioritize_by_dosage_form`, `_FORM_GROUPS`,
`forms_share_group`), pack-size handling, the GALACT rules, the ambiguous-strength check,
local ranking and scoring, the lower-variant retry, and the recovery paths that do not
depend on removed code.

Filter order is untouched: variant (Step 2) -> sub-variant (Step 3) -> name-strength
(Step 3.5) -> dosage form (Step 4) -> pack size (Step 5) -> GALACT -> ambiguous strength ->
local ranking -> Step 6 reranker (disabled). `prioritize_by_dosage_form` has **not** been
moved.

`map_one()` always calls `process_product(..., client=None, forced_brand=<brand from the
garbage-check stage>, auto_detect=False)`.

## 8. Brand-source architecture

Brand detection and matching come from the **garbage-check stage only**. The brand string
that `classify_row()` returns (`local_match["brand"]`, i.e. the same value
`garbage_check.py` writes into column C) is passed straight through as
`process_product(forced_brand=...)`, which is exactly the contract the two-step workflow
used. `auto_detect` is always `False`.

All the `garbage_check.py` brand machinery is retained: `SHORT_BRANDS`,
`EXACT_ONLY_BRANDS`, `LOCAL_REVIEW_BRANDS`, `_BRAND_CHAR_SETS`, `build_brand_index`,
`exact_product_match`, `fuzzy_product_match`, `token_spans` and the full exception cascade
in §6, including the `NUMLO` / `S-NUMLO` remap and the `NEW`-requires-`NORMET` rule.

## 9. Removed / omitted `mapping.py` brand matching

Deliberately not carried over, so that only one brand matcher exists:

`find_potential_brands`, `group_related_brands`, `find_abbrev_promoted_brand`,
`find_best_brand_for_input`, `suggest_nearest_brands`, `nearest_brand_within_edits`.

They had exactly four call sites, all inside `process_product`'s Step 1. One splice
(`mapping.py` lines 3455-3497) replaces them with the pipeline's own pre-existing result:

* brand hint not in `brand_map` -> `make_result("NO_CLEAR_MATCH", "", status="NO_BRAND")`
  (previously: auto-detect if enabled, else this same result).
* no brand hint at all -> the same `NO_BRAND` result.
* `if not brand:` -> the same `NO_BRAND` result, now unreachable, kept as a strict
  terminator.

Consequence: the `RECOVERED_NO_BRAND` status no longer exists. It was the edit-distance
recovery that mapped a row whose brand was within `NO_BRAND_RECOVERY_MAX_EDITS = 1` of a
master brand; removing it means no product code is ever produced from a brand the
garbage-check stage did not confirm. `NO_BRAND_RECOVERY_MAX_EDITS` and
`build_item_suggestions` remain defined but are now unused.

## 10. Existing accuracy safeguards (inherited unchanged)

* Strict `NO_CLEAR_MATCH` instead of a guessed mapping whenever brand, variant or strength
  cannot be resolved.
* `VARIANT_NOT_IN_MASTER` — an input variant absent from the master yields no code.
* `AMBIGUOUS_STRENGTH` — two candidates equally plausible on strength yields no code.
* `filter_items_by_name_strength` (Step 3.5) prevents a strength-bearing input from
  collapsing onto a differently dosed SKU.
* Sub-variant filtering before dosage form, so a sub-variant mismatch is never rescued by
  form agreement.
* The lower-variant retry accepts its result only when the recursive call returns
  `MATCHED`, then relabels it `RECOVERED_LOWER_VARIANT`.
* `repair_master_variant_fields` normalises master variant/sub-variant fields at load.
* Pack-size filtering is applied last and never introduces candidates.

## 11. Dosage-form-compatible retry behavior

Transplanted verbatim from the current `mapping.py` (the `# FORM-AWARE RETRY GUARD
(Steps 2/3/3.5)` block). Shape, as built:

* `_input_forms = detect_dosage_form_in_input(input_name)` is hoisted **above** Step 2.
* Each of Steps 2, 3 and 3.5 runs exactly as before. Its input pool is captured first
  (`_before_var`, `_before_sv`, `_before_ns`).
* After each step, if the input has a detected form and **every** surviving candidate is
  form-incompatible, the form-compatible subset of the pre-filter pool is computed
  (`_pool = _form_compatible(_before_*, _input_forms)`) and the *same* filter is re-run on
  that pool. The retry is adopted only when it returns a non-empty result.
* Each retry passes `reference_items=_pool`. This is essential and easy to lose:
  `filter_items_by_variant`'s `PLAIN` fallback consults `reference_items`, so passing the
  full original list would re-inject the very SKU the form check just excluded.
* `_form_compatible` treats a product with **no** detectable form as compatible. This is
  deliberate — many correct master names carry no form token, and dropping them is what
  made a pure form-first ordering destructive (~35 regressions when that was measured).
* `prioritize_by_dosage_form` remains a filter at Step 4, in place, not a re-orderer and
  not moved.

Measured effect of this guard in `mapping.py` (recorded in `context.md`): 41 rows changed,
39 improvements, 2 regressions — both `OROFER XT ... SYR` rows, caused by the `SYR` defect
in §18.

## 12. Reranker behavior and current commented-out state

`PROCESSING_MODE = "local"`, so `llm_enabled()` is `False` and **no LLM call is made or
possible**. Three layers enforce this:

1. `main()` raises immediately if `llm_enabled()` is ever true.
2. The reranker transport is commented out. `call_groq_llm` exists only as an active stub
   that raises `RuntimeError`, so any future path reaching for the reranker fails loudly
   instead of silently degrading into an unreranked guess.
3. No credential, endpoint or model name is present in this file at all
   (`HUGGINGFACE_API_KEY`, `HUGGINGFACE_API_URL`, `MAX_OUTPUT_TOKENS`, `MODEL_NAME`,
   `REQUEST_TIMEOUT` are undefined here by design).

In `process_product`, Step 6 keeps its original `if not llm_enabled():` early return
verbatim, and everything after it — prompt assembly, the API call, token accounting, and
the `NO_CLEAR_MATCH` / unmatched-answer handling — is preserved as a commented block,
followed by a strict `NO_CLEAR_MATCH` terminator so the function can never fall through
returning `None`.

The accuracy consequence is implemented in `decide_output()`:

```python
RERANKER_DEPENDENT_STATUSES = {
    "RECOVERED_LOCAL_ONLY", "LOCAL_ONLY_UNRESOLVED",
    "RECOVERED_API_ERROR", "API_ERROR",
    "RECOVERED_LLM_REJECTED", "LLM_REJECTED",
    "RECOVERED_LLM_UNMATCHED", "LLM_UNMATCHED",
}
```

Every one of these statuses means "the reranker was supposed to confirm or reject a
locally chosen candidate, and it did not". All of them map to `NO_SUGGESTION / 0`. In
particular `RECOVERED_LOCAL_ONLY` — the status `mapping.py` returns for a row it resolved
locally *while waiting for* the reranker — is **not** promoted to a product code. With the
reranker off, the only statuses that yield a code are `MATCHED` (which, in local mode, can
only come from local certainty such as the "strong local match — skipping reranker"
branch) and `RECOVERED_LOWER_VARIANT` (accepted only when its recursive call returned
`MATCHED`).

To re-enable: set `PROCESSING_MODE`, uncomment `call_groq_llm` and the Step 6 block, supply
the credential/endpoint/model configuration, and set a non-empty `MAPPING_LLM_PROMPT`.

## 13. `MAPPING_LLM_PROMPT` and its intentionally blank value

`MAPPING_LLM_PROMPT = ""`. The variable exists so there is one obvious, named place for the
reranker prompt when the LLM stage is turned back on. It is blank on purpose: no prompt is
configured because no LLM call is made in this build, and shipping a half-specified prompt
would invite someone to enable the reranker without reviewing it. The original prompt
scaffolding (`SYSTEM_PROMPT`, `USER_PROMPT_TEMPLATE`) is still present from the transplant
but is referenced only by commented-out code.

## 14. Special cases inherited from the existing pipeline

From the garbage-check stage: every brand exception in §6.

From the mapping stage, all retained verbatim: OCR de-segmentation of glued inputs;
manufacturer-prefix noise stripping (e.g. `EMCURE`); parenthetical-noise stripping that
spares genuine brand variant tags; duplicate-word cleanup; the GALACT-specific rules; the
compound-token pre-filter; `_FORM_GROUPS`, which treats SUSPENSION / SYRUP / SOLUTION /
DROPS / TONIC as one `liquid_oral` group so a SYRUP input can match a SUSPENSION SKU; the
`PLAIN` variant fallback; and the dosage-form pattern tables, transplanted exactly as they
stand (including the known `SYR` defect in §18 — not worked around here).

## 15. Current assumptions

* The input workbook has a sheet named `garbage_check` whose column A holds the raw OCR
  row and whose row 1 is a header.
* `Brand Names.txt` supplies the brand vocabulary as quoted strings; the master workbook
  supplies `PRODUCT_CODE, PRODUCT_NAME, BRAND_NAME, PACK_SIZE, VARIANT, SUB_VARIANT`.
* Every real product code is 9 digits, so `0`/`1`/`2` are safe sentinels.
* Paths in the `CONFIGURATION` block are absolute, so the script does not depend on the
  working directory (both source scripts use relative paths for the input and brand files).
* `PROCESSING_MODE` stays `"local"`.

## 16. Design decisions

1. **Transplant, not import.** Importing the two scripts would have executed their
   module-level API configuration and coupled this pipeline to future edits; a verbatim
   line-range transplant keeps it self-contained and keeps every interleaved comment.
2. **Line-range exclusion, not AST re-emission.** Re-emitting parsed nodes would have
   silently dropped the comments that carry most of the accuracy rationale.
3. **One brand matcher.** The garbage-check one, because that is the brand the two-step
   workflow already fed to the mapping stage; keeping a second one would create a second,
   differently-tuned opinion about brands.
4. **Reranker-dependent means unresolved.** Emitting the unreranked local candidate would
   have been more aggressive than the existing pipeline, which defers exactly those rows to
   the reranker. `NO_SUGGESTION / 0` is the honest answer.
5. **Two explicit sentinel scales.** Remark carries the human-readable reason and the code
   column carries a machine value, so a downstream consumer can filter on either.
6. **Fail closed.** Unreachable branches still return `NO_CLEAR_MATCH`; a row that raises
   keeps `NO_SUGGESTION / 0`; `call_groq_llm` raises rather than returning something usable.
7. **Blank rows emitted blank**, preserving row alignment without inventing a
   classification that the garbage-check stage does not make.

## 17. Validation performed

Run on 2026-08-27 against `test1.xlsx` sheet `garbage_check` (9,002 data rows).

* `python -m py_compile` clean on all three files.
* Duplicate top-level definition scan on the built file: none.
* A `symtable`-based undefined-global scan on the built file: clean (197 module-level
  names). The same scan was run against `garbage_check.py` and `mapping.py` as a control;
  both clean. This is the check that proves no kept function references a removed one.
* Output workbook: exactly one sheet (`mapping_main`), header exactly
  `('Input', 'Remark', 'Product Code')`, every row exactly 3 cells, 9,002 data rows.
* Output semantics, 0 violations across all 9,002 rows:

  | Remark | Code | Rows |
  |---|---|---|
  | `GARBAGE` | `1` | 2,902 |
  | `REVIEW` | `2` | 1,107 |
  | `NO_SUGGESTION` | `0` | 4,489 |
  | matched product name | 9-digit code | 504 |

  The 4,489 are 4,385 maybe-product rows plus 104 confirmed-product rows the mapping stage
  could not resolve. Of the 608 confirmed-product rows, 504 were mapped and 104 were not.
* Reranker-dependent rows report `NO_SUGGESTION / 0`. The 104 unresolved rows are
  `VARIANT_NOT_IN_MASTER` 67, `RECOVERED_LOCAL_ONLY` 22, `AMBIGUOUS_STRENGTH` 15. The 22
  `RECOVERED_LOCAL_ONLY` rows are precisely the reranker-dependent ones, and none received
  a code. The 504 codes come from `MATCHED` 493 plus `RECOVERED_LOWER_VARIANT` 11.
* Successful mappings overwrite the initial `NO_SUGGESTION / 0`: every confirmed-product
  row starts at that value in `main()` and is replaced only by `decide_output()`.
* No LLM call occurs. Verified three ways: `PROCESSING_MODE` is `"local"`; the only live
  reference to `call_groq_llm` in the file is the raising stub and no live code path
  touches `requests`; and a harness run with `requests.post`, `requests.get` and
  `socket.socket.connect` all replaced by raising stubs completed with 0 attempts.
* Faithfulness A/B. All 608 confirmed-product rows were run through both
  `MAPPING_MAIN.process_product` and `mapping.process_product` with identical arguments
  (`forced_brand` from `classify_row`, `auto_detect=False`, `client=None`, reranker
  disabled in both). Comparing `(output, product_code, status)`: 0 differences. The
  transplant is behaviourally identical to the mapping engine on every row it can reach.
* Both source scripts are byte-identical to the pre-build snapshots taken from them
  (sha256 prefixes `1a545166919c03e1` and `0bdc886002f6e0ea`). `git status` shows
  `garbage_check.py` unmodified and `mapping.py` at its pre-existing HEAD+54 insertions.
* `test1.xlsx` is unmodified.

## 18. Known limitations

* **SYR misclassified as INJECTION.** In the dosage-form tables, bare `\bSYR\.?\b` sits
  under `INJECTION` (input side, and again on the product side). OCR writes SYRUP as `SYR`,
  so a syrup row is read as an injection. With the form-aware retry active this costs two
  correct `OROFER XT ... SUSPENSION` mappings, which instead become
  `424440118 OROFER XT-DHA 7 X 1 KIT`. This is measured and pending a decision on
  `mapping.py`; it is transplanted verbatim here and deliberately not worked around, so that
  this file stays behaviourally identical to the mapping engine. A fix belongs in
  `mapping.py` first and should then be mirrored here.
* **MAYBE_PRODUCT rows are never mapped.** 4,385 rows, 49% of the input, end at
  `NO_SUGGESTION / 0` without entering the mapping stage. Under the two-step workflow these
  could be attempted with `auto_detect=True`, which used the mapping engine's own brand
  matching -- intentionally absent here.
* **REVIEW rows are terminal.** 1,107 rows return `REVIEW / 2` with no attempt at a
  suggestion.
* **No reranker.** 22 rows per run reach a local candidate the reranker was meant to
  adjudicate and are reported as unresolved.
* **No incremental resume.** The garbage-check script could resume via `has_result`; this
  script always processes the whole sheet and rewrites the output file.
* Recovery that depended on removed code (`RECOVERED_NO_BRAND`) is gone, so a small number
  of single-edit brand typos that previously recovered now do not.
* The output workbook carries no diagnostics -- no status, confidence, candidate count or
  suggestions column -- by design, which makes a disagreement harder to triage from the
  file alone.

## 19. Future work

* Decide the SYR fix in `mapping.py`, then mirror it here.
* Re-enable the reranker (section 12) and re-measure the 22 `RECOVERED_LOCAL_ONLY` rows;
  that is the largest single block of recoverable accuracy in the current run.
* Decide whether MAYBE_PRODUCT rows should get a mapping attempt, and if so what brand
  source is acceptable for them given section 9.
* Consider an optional diagnostics sheet, kept out of the three-column deliverable.

## 20. Relationship between the three files

| File | Role | Status |
|---|---|---|
| `garbage_check.py` | source of truth for classification and brand detection | unchanged |
| `mapping.py` | source of truth for the mapping pipeline | unchanged |
| `MAPPING_MAIN.py` | combined runnable pipeline, transplanted from both | the deliverable |

`MAPPING_MAIN.py` is a derived artifact. Accuracy changes belong in `garbage_check.py` or
`mapping.py` first, where they can be A/B measured against the existing corpus, and are
then mirrored here by re-running the same transplant. The two intentional divergences are
section 9 (one brand matcher) and section 12 (reranker-dependent rows are unresolved);
anything else that differs is a transplant bug and should be reported as such. `context.md`
documents the two source scripts and points here for this file.
