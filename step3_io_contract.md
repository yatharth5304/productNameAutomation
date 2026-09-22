# Step 3 Scorer: Input/Output Contract

## 1. Current Available Inputs from Existing Pipeline

### OCR Input Layer

**Source file:** `test2.xlsx` (Column A, named `Input_Column` in code)

The original OCR-extracted product name string is available as-is, with no cleaning applied at the retrieval stage. `transformer.py` encodes this text exactly as provided (`encode_query()` line 536: "Used exactly as provided -- no cleaning, normalisation, or reconstruction").

### Garbage Check Layer

**Source file:** `garbage_check.py`

`garbage_check.py` currently reads `test2.xlsx` and writes results back to it in a `garbage_check` worksheet with these columns:

| Column | Header | Values |
|--------|--------|--------|
| B | Garbage check | `"0"` (confirmed product), `"1"` (garbage), `"MAYBE_PRODUCT"`, `"REVIEW"` |
| C | Matched brand | Brand name (exact or fuzzy match), or `"pattern-inferred"` for form+qty matches, or empty |
| D | Matched text | The specific text span that matched the brand |
| E | Match type | `"exact"`, `"1-substitution"`, `"1-insertion"`, `"1-deletion"`, `"1-transposition"`, `"2-edit"` |
| F | Matched garbage phrase | For garbage rows |
| G | Garbage match type | For garbage rows |

Only rows classified as `"0"` (confirmed product by `PROCESSING_MODE = "local"` in `mapping.py`: `PROCESS_CONFIRMED_ZERO_ONLY = True`) are handed to the mapper. The brand hint from Column C feeds into the retriever.

### Retrieval Layer

**Source file:** `transformer.py`

The `retrieve()` function (lines 663-722) returns a result dict with:

- `ocr_input` — original OCR text
- `detected_brand` — normalised brand (None if none detected; sentinels `pattern-inferred`, `nan`, `none`, `""` are treated as None)
- `candidate_pool_size` — number of eligible master records
- `filter_case` — `"A"` (brand-filtered), `"B"` (full master, no brand), `"C"` (full master, brand not found), or `"EMPTY_INPUT"`
- `candidates` — list of up to `TOP_K=10` dicts, each containing:
  - `rank` — 1-based position
  - `similarity_score` — cosine similarity (float)
  - `PRODUCT_CODE` — master product code (e.g., `"421112895"`)
  - `product_name` — master product name (e.g., `"CARDACE AM 2.5 TABLET 10x15T"`)
  - `BRAND_NAME` — uppercase brand (e.g., `"CARDACE"`)
  - `master_record_index` — 0-based index into master
  - `master_row_number` — Excel row number

The fine-tuned model produces identical output format via `finetune_eval.py` and `deep_eval_analysis.py`, which re-embed the master with the fine-tuned model and run the same retrieval logic.

### Master Product Record

**Source file:** `PRODUCT_MASTER.xlsx`

Contains columns: `PRODUCT_CODE`, `PRODUCT_NAME`, `BRAND_NAME`. The master has 2,765 records. `transformer.py` loads these as `{row_number, PRODUCT_CODE, product_name, BRAND_NAME}` dicts. The downstream scoring can look up full master record details by `PRODUCT_CODE` if needed, since the codes are unique per `code_to_rec` mapping in `baseline_eval.py` line 45.

### Fine-tuned Model Artifacts

- Model: `models/bge-pharma-v1/` (saved via `model.save_pretrained`)
- Fine-tuned embeddings: `embeddings/finetuned_master_embeddings.npy` (shape `(2765, 768)`, normalised)
- Fine-tuned meta: `embeddings/finetuned_master_meta.csv` (idx + PRODUCT_CODE + product_name + BRAND_NAME)

---

## 2. Proposed Step 3 Input Contract

Step 3 receives a single structured payload per OCR input. The payload is constructed by composing the existing pipeline stages:

```json
{
  "ocr_input": "CARDACE AM 2.25/5MG 1X10 PCS",
  "detected_brand": "CARDACE",
  "filter_case": "A",
  "candidates": [
    {
      "rank": 1,
      "similarity_score": 0.8567,
      "PRODUCT_CODE": "421112895",
      "product_name": "CARDACE AM 2.5 TABLET 10x15T",
      "BRAND_NAME": "CARDACE",
      "master_record_index": 142,
      "master_row_number": 889
    }
    // ... up to 10 candidates
  ]
}
```

### Required Inputs

| Field | Source | Rationale |
|-------|--------|-----------|
| `ocr_input` | `transformer.py` `retrieve()` → `ocr_input` | The complete original OCR text is needed to extract strength, form, and pack size tokens for structured comparison against candidates. The deep error analysis (`deep_eval_analysis.py` lines 54-63) shows all 47 error cases are resolved by comparing structured fields extracted from the OCR input. |
| `detected_brand` | `transformer.py` → `retrieve()` → `detected_brand` | Already normalised (None for `pattern-inferred`/`nan`/`none`/`""` sentinels). Confirms the brand filtering case. |
| `filter_case` | `transformer.py` → `retrieve()` → `filter_case` | Tells the scorer whether candidates were brand-filtered (Case A) or from full master (Case B/C), which affects confidence bounds. |
| `candidates[]` | `transformer.py` → `retrieve()` → `candidates` | The Top-10 from the V1 fine-tuned retriever. Each candidate has `rank`, `similarity_score`, `PRODUCT_CODE`, `product_name`, `BRAND_NAME`. |

### Optional Inputs (not part of core contract, available on-demand)

| Field | Source | When to use |
|-------|--------|-------------|
| `master_record_index` / `master_row_number` | `transformer.py` candidate dict | Debugging/tracing only; not needed for scoring logic |
| Full master record (beyond `product_name`) | `PRODUCT_MASTER.xlsx` via `PRODUCT_CODE` lookup | If the scorer needs additional master columns (e.g., VARIANT, SUB_VARIANT, PACK_SIZE) that are not in the candidate dict. **Note:** `PRODUCT_MASTER.xlsx` only has 3 columns (PRODUCT_CODE, PRODUCT_NAME, BRAND_NAME) per `transformer.py` line 172. VARIANT/SUB_VARIANT/PACK_SIZE are only in the older `PRODUCT_MASTER1.xlsx` used by `mapping.py`. |

---

## 3. Required vs Optional Inputs Summary

### Required (must be passed to Step 3)

1. Original OCR product name — `ocr_input`
2. Detected brand (normalised) — `detected_brand`
3. Filter case — `filter_case`
4. Top-10 candidates with PRODUCT_CODE, product_name, BRAND_NAME, similarity_score, rank

### Optional (available but not part of core contract)

- `candidate_pool_size` — for debugging context
- `master_record_index` / `master_row_number` — for tracing
- Full master record lookup — if extended master schema needed in future

---

## 4. Proposed Step 3 Output Contract

Step 3 returns a single result per OCR input:

```json
{
  "selected_PRODUCT_CODE": "421112895",
  "selected_product_name": "CARDACE AM 2.5 TABLET 10x15T",
  "selected_retriever_rank": 1,
  "scorer_result": "CONFIDENT_MATCH",
  "score_breakdown": {
    "brand_match": true,
    "strength_match": true,
    "form_match": true,
    "pack_match": true,
    "sub_variant_match": true,
    "retriever_score": 0.8567
  },
  "evidence": "All structured fields match: strength=2.5MG, form=TABLET, pack=10x15T"
}
```

### Required Outputs

| Field | Description |
|-------|-------------|
| `selected_PRODUCT_CODE` | The final mapped product code (empty/None if no match) |
| `selected_product_name` | The master product name of the selected candidate |
| `selected_retriever_rank` | The rank (1-10) the selected candidate held in the Top-10 |
| `scorer_result` | Classification of outcome: `"CONFIDENT_MATCH"`, `"LOW_CONFIDENCE"`, `"NO_CLEAR_MATCH"`, `"AMBIGUOUS"` |
| `score_breakdown` | Structured per-field comparison results (see below) |
| `evidence` | Human-readable explanation of why this candidate was selected or why matching failed |

### `score_breakdown` Fields

| Sub-field | Source | Description |
|-----------|--------|-------------|
| `brand_match` | Step 3 computation | Whether extracted brand from OCR matches candidate brand |
| `strength_match` | Step 3 computation | Whether strength tokens match (e.g., `"5MG"` in both) |
| `form_match` | Step 3 computation | Whether dosage form matches (e.g., TABLET vs TABLET) |
| `pack_match` | Step 3 computation | Whether pack size/count matches (e.g., `10x15T`) |
| `sub_variant_match` | Step 3 computation | Whether sub-variant/modifier matches (e.g., `AM`, `FORTE`, `DM`) |
| `retriever_score` | Input (from candidates) | Pass-through of the original similarity score |

---

## 5. Required vs Optional Outputs Summary

### Required Outputs (core contract)

1. `selected_PRODUCT_CODE`
2. `selected_product_name`
3. `selected_retriever_rank`
4. `scorer_result`
5. `score_breakdown` (with brand_match, strength_match, form_match, pack_match, sub_variant_match, retriever_score)
6. `evidence`

### Optional Outputs (for evaluation/debugging, not part of core contract)

| Field | Purpose |
|-------|---------|
| `confidence_score` (0-1) | Numeric confidence for threshold-based filtering in downstream pipeline |
| `all_scores` | Full scoring breakdown for all 10 candidates (for error analysis) |
| `matched_fields_detail` | Exact matched strings per field (e.g., `"strength": {"ocr": "5MG", "candidate": "5MG"}`) |
| `unmatched_fields` | List of fields that did not match and their values |
| `processing_log` | Step-by-step trace of the scoring process |

---

## 6. Data Flow Diagram

```
┌─────────────────────────────────────────────────────────┐
│  OCR Input (test2.xlsx, Column A / Input_Column)       │
└──────────────┬──────────────────────────────────────────┘
               │ Original OCR text (no cleaning)
               ▼
┌─────────────────────────────────────────────────────────┐
│  garbage_check.py                                       │
│  • Classify: 0 (product) / 1 (garbage) / MAYBE_PRODUCT  │
│  • Brand detection: exact + fuzzy                       │
│  Outputs: Column B (status), Column C (brand)           │
└──────────────┬──────────────────────────────────────────┘
               │ Only rows with Column B = "0" proceed
               │ Brand hint (Column C) passed downstream
               ▼
┌─────────────────────────────────────────────────────────┐
│  transformer.py (V1 fine-tuned model)                   │
│  • Brand hard-filter (Case A/B/C)                       │
│  • Cosine similarity Top-K retrieval                    │
│  • Returns Top-10 candidates with rank + score           │
└──────────────┬──────────────────────────────────────────┘
               │ Top-10 candidates + filtered_indices
               ▼
┌─────────────────────────────────────────────────────────┐
│  STEP 3 SCORER (this contract)                          │
│                                                         │
│  INPUT:                                                  │
│    - ocr_input (str)                                     │
│    - detected_brand (str|None)                           │
│    - filter_case (str: "A"|"B"|"C"|"EMPTY_INPUT")        │
│    - candidates[] = [{rank, similarity_score,           │
│                       PRODUCT_CODE, product_name,        │
│                       BRAND_NAME}]                       │
│                                                         │
│  OUTPUT:                                                 │
│    - selected_PRODUCT_CODE (str)                         │
│    - selected_product_name (str)                         │
│    - selected_retriever_rank (int)                       │
│    - scorer_result (str: CONFIDENT_MATCH|...)            │
│    - score_breakdown {brand_match, strength_match, ...}  │
│    - evidence (str)                                      │
└──────────────┬──────────────────────────────────────────┘
               │ Final product mapping decision
               ▼
┌─────────────────────────────────────────────────────────┐
│  Output / downstream consumption                        │
│  • Write PRODUCT_CODE to output spreadsheet             │
│  • Use scorer_result for confidence routing             │
└─────────────────────────────────────────────────────────┘
```

---

## 7. Unresolved Questions

1. **Scorer result vocabulary:** The exact set of `scorer_result` values is proposed above but not finalized. The evaluation artifacts reference confidence levels (HIGH/MEDIUM/LOW) in `mapping.py` line 145, but those are for the LLM reranker, not Step 3. Should Step 3 reuse a different vocabulary?

2. **Empty candidate list handling:** When `filter_case = "EMPTY_INPUT"` or no candidates are returned, `selected_PRODUCT_CODE` should be empty. This needs explicit treatment in the output contract.

3. **No-match fallback:** If Step 3 cannot select any candidate (e.g., all are below a self-imposed threshold, though thresholds are out of scope here), should the output be `"NO_CLEAR_MATCH"` with `selected_PRODUCT_CODE = ""`, or should it default to the highest-ranked candidate? **Proposed answer:** Return `"NO_CLEAR_MATCH"` with empty code; let downstream handle it.

4. **Brand mismatch in candidates:** The deep error analysis shows 0 cases of "brand confusion" (Group B fixability = 0). Since the retriever already brand-filters (Case A), all Top-10 candidates should be same-brand. Should `brand_match` be computed anyway as a safety check, or can it be assumed true? **Proposed answer:** Compute it as a safety check — 0 errors today doesn't mean 0 tomorrow.

5. **Multiple master records per product name:** `baseline_eval.py` line 43 uses `exact_name.setdefault(name, []).append(rec)` — some product names map to multiple PRODUCT_CODEs. In the candidate dict, `product_name` is unique per candidate, so this is handled at the retriever level. Step 3 receives the already-resolved `PRODUCT_CODE` so this is not an issue.

---

## DECISION

**INPUT:**

```
{
  "ocr_input": str,                    // original OCR text, un-cleaned
  "detected_brand": str | None,        // normalised brand from garbage_check, None if sentinel
  "filter_case": str,                  // "A" | "B" | "C" | "EMPTY_INPUT"
  "candidates": [
    {
      "rank": int,                    // 1-based
      "similarity_score": float,       // cosine similarity
      "PRODUCT_CODE": str,             // master product code
      "product_name": str,             // master product name
      "BRAND_NAME": str                // uppercase brand
    }
  ]
}
```

**Required:** `ocr_input`, `detected_brand`, `filter_case`, `candidates[]` (each with `rank`, `similarity_score`, `PRODUCT_CODE`, `product_name`, `BRAND_NAME`).

**Optional (available on-demand, not in core contract):** `candidate_pool_size`, `master_record_index`, `master_row_number`, full master record lookup via `PRODUCT_CODE`.

---

**OUTPUT:**

```
{
  "selected_PRODUCT_CODE": str,        // "" if no match
  "selected_product_name": str,        // "" if no match
  "selected_retriever_rank": int,      // 0 if no selection made
  "scorer_result": str,                // "CONFIDENT_MATCH" | "LOW_CONFIDENCE" | "NO_CLEAR_MATCH" | "AMBIGUOUS"
  "score_breakdown": {
    "brand_match": bool,
    "strength_match": bool,
    "form_match": bool,
    "pack_match": bool,
    "sub_variant_match": bool,
    "retriever_score": float
  },
  "evidence": str                      // human-readable reasoning
}
```

**Required:** `selected_PRODUCT_CODE`, `selected_product_name`, `selected_retriever_rank`, `scorer_result`, `score_breakdown` (all 5 sub-fields), `evidence`.

**Optional (not in core contract):** `confidence_score` (float), `all_scores` (per-candidate), `matched_fields_detail` (exact matched strings), `unmatched_fields` (list), `processing_log` (trace).
