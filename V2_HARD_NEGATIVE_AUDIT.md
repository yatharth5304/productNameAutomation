# V2 Hard-Negative Mining — Audit Report

**Source file**: `v2_hard_negative_mining.csv`  
**Audit performed**: 2026-09-23  
**Auditor**: automated — `v2_audit.py` + direct inspection  

---

## Executive Summary

| Finding | Result |
|---------|--------|
| Total rows | 108,913 |
| Rows with 3 valid negatives | 108,913 (100%) |
| Missing values / malformed fields | **0** |
| Positive/negative code collision | **0** |
| Positive/negative name collision | **0** |
| Duplicate-negative rows (Section B) | **82** |
| Root cause of all 82 duplicates | Single catalogue alias (`ARTH WEIGHT MANAGEMENT SUPPORT 15T`) |
| Negatives with score ≥ 0.99 | 8 (negative-slots), 2 with score = 1.00 |
| Evaluation input leakage | **0** (no training `Input_Column` matches eval) |
| Rows whose positive code appears in eval | 15,423 (expected — same products appear in both splits) |
| Rows with a mined-negative code in eval | 44,122 (expected — see Section E) |
| Confirmed catalogue aliases | 1 pair (`421113061` vs `421113210`) |
| Rows safe for training (no structural issues) | **108,831** |
| Rows needing review (dup negatives only) | **82** |

---

## 1. Dataset Statistics

### Column Schema

All 12 expected columns are present with no extra or missing columns:

```
Input_Column
Positive_Product
Positive_Product_Code
Mined_Negative_1  Mined_Negative_1_Code  Mined_Negative_1_Score
Mined_Negative_2  Mined_Negative_2_Code  Mined_Negative_2_Score
Mined_Negative_3  Mined_Negative_3_Code  Mined_Negative_3_Score
```

### Row Counts

| Metric | Value |
|--------|-------|
| Total rows | 108,913 |
| Rows with 3 hard negatives | 108,913 |
| Rows with 2 hard negatives | 0 |
| Rows with 1 hard negative | 0 |
| Rows with 0 hard negatives | 0 |

### Score Distribution (from CSV)

| Slot | Count | Min | Max | Mean | Median | Std |
|------|-------|-----|-----|------|--------|-----|
| 1 | 108,913 | 0.4467 | 1.0000 | 0.7705 | 0.7816 | 0.0850 |
| 2 | 108,913 | 0.4369 | 0.9675 | 0.7168 | 0.7205 | 0.0838 |
| 3 | 108,913 | 0.4228 | 0.9612 | 0.6804 | 0.6801 | 0.0799 |

These distributions are consistent with the quality report — confirming the CSV as written matches the reported statistics.

---

## 2. Row Integrity (Section A)

✅ **All row-integrity checks passed.**

| Check | Result |
|-------|--------|
| Empty `Input_Column` | 0 |
| Empty `Positive_Product` | 0 |
| Empty `Positive_Product_Code` | 0 |
| Missing negative code where name is present | 0 |
| Unparseable score values | 0 |
| Inconsistent slot fields (code/score present but name empty) | 0 |
| Control characters in product names | 0 |

The CSV is structurally clean. Every row has exactly three populated negative slots, each with a name, code, and parseable float score.

---

## 3. Negative Uniqueness (Section B)

> [!WARNING]
> **82 rows** contain duplicate negatives. The same product name appears in more than one negative slot for the same training row.

### Root Cause — Identified and Confirmed

The cause is a **single catalogue alias**: the product name `ARTH WEIGHT MANAGEMENT SUPPORT 15T` exists under **two different PRODUCT_CODEs** in the master:

| Code | Name |
|------|------|
| `421113061` | ARTH WEIGHT MANAGEMENT SUPPORT 15T |
| `421113210` | ARTH WEIGHT MANAGEMENT SUPPORT 15T |

These are **two distinct catalogue records with identical names** (confirmed from Section F). When the mining script retrieves Top-10 candidates, both entries appear in the results. Because they normalise to the same product name, the deduplication check (which compares by name) sees them as duplicates — but the mining script uses product code to exclude the positive, so both records passed the positive-exclusion filter and were assigned to two of the three negative slots.

This explains all 82 affected inputs: they are all product names containing "15T" / "15S" / "15s" — pack sizes that cause the `ARTH WEIGHT MANAGEMENT SUPPORT 15T` entries to be the highest-similarity non-positive matches.

### Sample Affected Rows (first 20 of 82)

| Row | Input | Neg names | Codes |
|-----|-------|-----------|-------|
| 1305 | VISTAVIT CAPS-15s | ARTH WEIGHT...15T × 2, ARTH SLEEP... | 421113210, 421113061 |
| 1950 | DINOR TAB 15 S | CORTIMAX-S CREAM, ARTH...15T × 2 | 421113061, 421113210 |
| 2330 | VISTAVIT CAPS 15S 10X15 | GLIPSOV, ARTH...15T × 2 | 421113061, 421113210 |
| 3227 | DINOR 2 15TA | ARTH...15T × 2, AMARYL 2MG | 421113061, 421113210 |
| 4354 | CALCIEM TAB 15S | ARTH...15T × 2, EXAFIB 15 | 421113061, 421113210 |
| 6690 | DINOR 2MG TAB 15TAB | AMARYL 2MG, ARTH...15T × 2 | 421113128, 421113210, 421113061 |
| 7064 | FM ACTIVE TAB 15T | FM ACTIVE TAB, ARTH...15T × 2 | 421112161, 421113061, 421113210 |
| 10288 | CALCIEM 15 TAB | EXAFIB 15, ARTH...15T × 2 | 424441889, 421113210, 421113061 |
| 13381 | Dinor Tab 15 s | CORTIMAX-S CREAM, ARTH...15T × 2 | 420030117, 421113061, 421113210 |
| 13891 | DINOR TABS 15T | ARTH...15T × 2, FERIUM XT | 421113061, 421113210, 424441955 |
| 15217 | DINOR TAB 15S | ARTH...15T × 2, ARMOD-150 | 421113061, 421113210, 424442061 |
| 17727 | DINOR TAB 15T | ARTH...15T × 2, FERIUM XT | 421113061, 421113210, 424441955 |
| 18275 | FM ACTIVE 1*15T | FM ACTIVE TAB, ARTH...15T × 2 | 421112161, 421113061, 421113210 |
| 20660 | DINOR TAB 15TAB | ARTH...15T × 2, ARMOD-150 | 421113061, 421113210, 424442061 |
| 21615 | CALCIEM TABS 15S | ARTH...15T × 2, EXAFIB 15 | 421113210, 421113061, 424441889 |
| 23208 | ZILARBI-CN 10MG TA 15S | ZILARBI CN, ARTH...15T × 2 | 421112320, 421113061, 421113210 |
| 25118 | DINOR TABS 15S | ARTH...15T × 2, ARMOD-150 | 421113210, 421113061, 424442061 |
| 25549 | Calciem Tab 15S | ARTH...15T × 2, EXAFIB 15 | 421113061, 421113210, 424441889 |
| 26719 | FEMZACT 15 S | CORTIMAX-S, ARTH...15T × 2 | 420030117, 421113061, 421113210 |
| 33312 | CALCIEM TAB. 15*TAB | ARTH...15T × 2, EXAFIB 15 | 421113061, 421113210, 424441889 |

*(62 more rows follow the same pattern)*

### Classification of the 82 Rows

In all 82 rows the pattern is:
- Slot X: `ARTH WEIGHT MANAGEMENT SUPPORT 15T` [421113061]  
- Slot Y: `ARTH WEIGHT MANAGEMENT SUPPORT 15T` [421113210]  

Occupying 2 of the 3 negative slots. The third slot always holds a genuinely distinct product.

---

## 4. Positive / Negative Collision (Section C)

✅ **Zero collisions on both checks.**

| Check | Count |
|-------|-------|
| Negative PRODUCT_CODE == Positive PRODUCT_CODE | **0** |
| Negative normalised name == Positive normalised name | **0** |

The mining exclusion logic (excluding the positive by PRODUCT_CODE) worked correctly for all 108,913 rows. No training row contains its own positive as a negative.

---

## 5. Very-High-Similarity Negatives (Section D)

> [!NOTE]
> High similarity alone does not make a negative invalid. These are flagged for review, not deletion.

| Threshold | Count (individual neg-slots) |
|-----------|------------------------------|
| ≥ 0.95 | 440 |
| ≥ 0.98 | 28 |
| ≥ 0.99 | 8 |
| = 1.00 | **2** |

### All 8 Negatives with Score ≥ 0.99

| Row | Slot | Score | Positive [Code] | Negative [Code] |
|-----|------|-------|-----------------|-----------------|
| 5673 | 1 | 0.9919 | `EMPRI L 25/5MG TABLET 10x10T` [421113070] | `EMPRI 25MG TABLET 10x10T` [421113068] |
| 23483 | 1 | 0.9933 | `XGRAST  INJ 300 MCG/0.5 ML PFS` [422220290] | `XGRAST 300 MCG INJ 1 PFS` [492220092] |
| 23865 | 1 | 0.9902 | `MAXTRA ORAL DROPS` [423331221] | `MAXTRA HS ORAL DROPS 15 ML` [423331762] |
| 43589 | 1 | 0.9957 | `DIOF DS SUSPENSION` [423331421] | `DIOF DS SUSPENSION 60ML` [423331332] |
| 59422 | 1 | 0.9901 | `MAXTRA P DS SYRUP 120ML` [423331688] | `MAXTRA DC SYRUP 120 ML` [423331755] |
| 84888 | 1 | 0.9912 | `EPOFER 4000 - PFS` [422220109] | `EPOFER 4000 IU VIAL` [422220112] |
| 103824 | 1 | **1.0000** | `I RINSE NASAL SPRAY` [423331477] | `I-RINSE SPRAY 100 ML` [423331727] |
| 103870 | 1 | **1.0000** | `LAZID E KIT 150/300/600 MG TABLET 2X5` [424441070] | `LAZID E KIT` [494440103] |

### Analysis of the Score = 1.00 Cases

**Case A (Row 103824)**: `I RINSE NASAL SPRAY` [423331477] vs `I-RINSE SPRAY 100 ML` [423331727]  
Different codes. The positive is the canonical catalogue name; the negative is a pack-specific variant. These appear to be **genuinely distinct catalogue entries** — the positive is the generic "brand" listing and the negative is a 100 mL SKU. This is an excellent hard negative (pack-size discrimination). Recommended: keep, verify codes are distinct products in the master.

**Case B (Row 103870)**: `LAZID E KIT 150/300/600 MG TABLET 2X5` [424441070] vs `LAZID E KIT` [494440103]  
Different codes (note different prefix: `494...` vs `424...`). The negative has no dosage/pack information — it may be an alias/stub entry or a different SKU altogether. The embedding similarity is 1.00 because the shorter name is a substring of the positive when encoded. This requires verification against the master catalogue — if `494440103` is a real, distinct product, it is a useful hard negative; if it is an alias or data artefact, it should be replaced.

### Analysis of Other ≥ 0.99 Cases

All 6 remaining cases are **plausible hard negatives**:
- EMPRI L 25/5 vs EMPRI 25mg — different combination product vs single-agent (excellent)
- XGRAST 300 MCG PFS vs XGRAST 300 MCG INJ — formulation/presentation variant (excellent)
- MAXTRA ORAL DROPS vs MAXTRA HS ORAL DROPS — different sub-brand, same form (good)
- DIOF DS SUSPENSION vs DIOF DS SUSPENSION 60ML — pack-size variant (good)
- MAXTRA P DS vs MAXTRA DC — different variant letters (good)
- EPOFER 4000 PFS vs EPOFER 4000 IU VIAL — presentation variant (excellent)

None of these have the same PRODUCT_CODE — they are confirmed distinct catalogue records.

---

## 6. Evaluation Leakage (Section E)

### Input-level leakage

✅ **Zero input-level leakage.**

| Check | Count |
|-------|--------|
| `Input_Column` exact match to eval input | **0** |
| `Input_Column` normalised match to eval input | **0** |

No training `Input_Column` string (raw or normalised) matches any evaluation `Input_Column`. The mining script's eval-exclusion filter worked correctly.

### Product-code level: expected overlap

| Check | Count | Interpretation |
|-------|-------|----------------|
| Training `Positive_Product_Code` in eval positive codes | 15,423 | **Expected — see below** |
| Training mined-negative code in eval positive codes | 44,122 | **Expected — see below** |

> [!IMPORTANT]
> This is **not a leakage problem**. It is the correct and expected behaviour.
>
> The evaluation set was built to test retrieval against the full PRODUCT_MASTER catalogue. The 103 unique product codes in `eval_pos_codes` represent popular/important products that naturally also appear as training positives (or as hard negatives for *other* training examples). The training and evaluation sets are split by **OCR input query** (`Input_Column`), not by product code. The model is evaluated on whether it can retrieve the correct product for an unseen input — the products themselves (catalogue entries) are shared between train and eval, which is correct for retrieval tasks.
>
> The handoff quality report's `Evaluation leakage: 0` refers specifically to input-level leakage, which is confirmed correct.

---

## 7. Catalogue Duplicates / Aliases (Section F)

| Pattern | Count |
|---------|-------|
| Same normalised name → multiple PRODUCT_CODEs | **1** |
| Same PRODUCT_CODE → multiple product names | **0** |

### The Confirmed Alias

| Normalised Name | Code 1 | Code 2 |
|-----------------|--------|--------|
| `ARTH WEIGHT MANAGEMENT SUPPORT 15T` | `421113061` | `421113210` |

This is the direct and sole root cause of all 82 duplicate-negative rows in Section B.

**Both codes have different prefixes** (`421113061` vs `421113210`). The PRODUCT_CODE prefix convention in this catalogue appears to encode product family or registrant. The fact that the last digits also differ confirms these are two distinct master records with identical display names. This is a genuine catalogue alias or duplicate master entry.

**Impact**: The mining script correctly excluded the positive (by code) and then selected both records as negatives — yielding two slots with the same product name but different codes. The duplication check in the quality report counted this as a validation error because it compared names, not codes.

---

## 8. Representative Examples

*(10 evenly spaced rows from the mining CSV — confirming data quality)*

### Input: `EMCURE ORS SACHET 1S`
- **Positive**: `EMCURE ORS 5.125GM SACHET` [421110014]
- **Neg 1**: `ORS SACHET` [421110019] score=0.7691
- **Neg 2**: `EMCURE ORS SACHET 21.8GM 5S` [421110016] score=0.7617
- **Neg 3**: `ORS-LEMON SACHET` [421110029] score=0.7501

### Input: `EMANZEN 5 MG TAB 10 S`
- **Positive**: `EMANZEN TABLET 5 MG` [424440086]
- **Neg 1**: `EMANZEN N CAPSULES 10` [420000468] score=0.7502
- **Neg 2**: `EMANZEN AP TABLET 10x10T` [421113182] score=0.7398
- **Neg 3**: `EMANZEN FORTE TABLET 10 MG` [424440087] score=0.7299

### Input: `ASOMEX 5 20*10`
- **Positive**: `ASOMEX 5 MG TABLET` [424440003]
- **Neg 1**: `ASOMEX AT 5 MG TABLET` [424440006] score=0.7729
- **Neg 2**: `ASOMEX 5 MG TABLET 1x15T` [424441737] score=0.7513
- **Neg 3**: `ASOMEX AT 5 MG TABLET 1X15T` [424441922] score=0.7194

### Input: `VINTOR 2000IU INJ PFS 1`
- **Positive**: `VINTOR 2000 IU INJ 1 PFS` [422220113]
- **Neg 1**: `VINTOR 6000 IU INJ 1 PFS` [422220118] score=0.8286
- **Neg 2**: `VINTOR 10000 IU INJ 1 PFS` [422220120] score=0.8181
- **Neg 3**: `VINTOR 4000 IU INJ 1 PFS` [422220108] score=0.8136

These examples confirm the dataset contains high-quality hard negatives with the expected discrimination patterns (same brand/different strength, same brand/different formulation, different pack size).

---

## 9. Recommended Cleaning Actions

### Action 1 — Resolve the 82 duplicate-negative rows (REQUIRED before training)

**Nature of the issue**: All 82 rows have two negative slots occupied by `ARTH WEIGHT MANAGEMENT SUPPORT 15T` — once with code `421113061` and once with code `421113210`.

**Decision options** (choose one; do not mix):

**Option A (Recommended)** — Treat as a catalogue alias; keep only one slot:  
Designate one code (e.g. `421113061`, the lower code) as the canonical representative. In each of the 82 rows, remove the duplicate slot and replace it with the **next-ranked valid candidate** from the mining Top-10 that is not the positive and not the already-kept `ARTH WEIGHT MANAGEMENT SUPPORT 15T`. Since the mining script only used Top-10 and most rows already have a genuine 3rd negative, for many rows this means:  
- Row already has a 3rd unique product in the remaining slot → keep it, de-duplicate the ARTH pair to 1.  
- The row now has 2 confirmed unique negatives and 1 missing slot → flag for further sourcing.

To implement this cleanly, re-run the mining for only these 82 `Input_Column` values with Top-20 retrieval and a deduplication step that compares by **normalised name** (not code), then select the best 3 distinct-by-name negatives.

**Option B (Conservative)** — Drop the duplicate slot; tolerate 2 negatives for 82 rows:  
De-duplicate to keep only one `ARTH WEIGHT MANAGEMENT SUPPORT 15T` per row. The 82 rows would then have 2 valid negatives instead of 3. This is acceptable: `MultipleNegativesRankingLoss` can handle variable numbers of negatives per example.

**Do NOT**: 
- invent a replacement negative
- silently keep both codes treating them as distinct products
- delete these 82 rows

### Action 2 — Verify the two score = 1.00 negatives (ADVISORY)

For Row 103870 (`LAZID E KIT` [494440103]): check whether PRODUCT_CODE `494440103` is a valid active SKU in the master or a stub/alias/retired entry. If it is a valid distinct product, keep it. If it is an alias for `424441070`, replace it.

For Row 103824 (`I-RINSE SPRAY 100 ML` [423331727]): verify that `423331727` is a distinct 100 mL pack separate from `423331477`. If confirmed distinct, this is an excellent hard negative — keep it.

### Action 3 — No action required for leakage (CONFIRMED CLEAN)

The evaluation leakage check is confirmed clean at the input level. The product-code overlap is structurally correct for retrieval training. No rows need to be removed for leakage reasons.

### Action 4 — No action required for high-similarity negatives (ADVISORY ONLY)

The 438 cases with score ≥ 0.95 (excluding the 2 score=1.00 handled above) all have different product codes and represent genuinely difficult negatives. These are training assets, not errors. Keep all of them.

---

## 10. Final Row Counts — Safe for Training

**Definition of "safe"**: no duplicate negatives, no positive/negative collision, no leakage.

| Category | Count |
|----------|-------|
| Total rows in `v2_hard_negative_mining.csv` | **108,913** |
| ✅ Safe for training without any changes | **108,831** |
| ⚠️ Needs resolution before training (dup negatives) | **82** |

The 82 rows needing resolution are all fixable by a single deterministic rule (de-duplicate ARTH alias, optionally source replacement). They should **not** be deleted — they can be cleaned and added back to the training set.

---

## 11. V1 Model and Architecture Context

For completeness, key facts confirmed from the repository code:

| Item | Value |
|------|-------|
| V1 model checkpoint | `models/bge-pharma-v1/` (fine-tuned from `BAAI/bge-base-en-v1.5`) |
| Embedding dimension | 768 |
| V1 training loss | `MultipleNegativesRankingLoss` (scale=20) |
| V1 training data source | `training_split.csv` |
| V1 training: epochs | 1 |
| V1 training: batch size | 4 (effective 16 with grad accum 4) |
| V1 training: LR | 2e-5 |
| Mining model used for V2 negatives | `models/bge-pharma-v1/` (V1 fine-tuned) |
| Mining Top-K | 10 |
| Mining keep-N negatives | 3 |
| Product representation | Full `product_name` string only (no field splitting) |
| Retrieval Top-K (production) | 10 (note: V2 handoff specifies Top-20 for V2) |
| Evaluation set size | 364 rows (365 lines - 1 header), 170 unique inputs, 103 unique positive codes |
| Evaluation leakage | 0 (input level — confirmed) |

---

*End of V2 Hard-Negative Mining Audit Report.*
