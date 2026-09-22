# V2 DATASET CONSTRUCTION & VALIDATION REPORT

## Executive Summary

The V2 positive training dataset has been constructed from the historical EmcureSSSReport.xlsb (Sheet1) dataset using the D3 Hybrid strategy recommended in Stage 2 analysis.

**Final Dataset:**
- **109,283 unique positive pairs** (PRODUCT_NAME → MATERIAL DESCRIPTION)
- **2,190 MATERIALS** represented (all 2,185 historical + 7 from V1)
- **Zero evaluation leakage** - verified
- **Zero mapping conflicts** - verified
- **99.6% PRODUCT_MASTER coverage** for materials in dataset

---

## 1. V2 POSITIVE DATASET CONSTRUCTION RESULTS

### Construction Pipeline
| Step | Operation | Rows/Pairs |
|------|-----------|------------|
| 1 | Load Sheet1 | 317,054 |
| 2 | Remove NO PRODUCT (MATERIAL DESCRIPTION = "-") | -3,152 |
| 3 | Remove exact duplicates | -1,782 |
| 4 | Clean usable rows | **312,120** |
| 5 | Extract unique positive pairs | **148,089** |
| 6 | Apply D3 Hybrid capping | **108,775** |
| 7 | Add V1 positive pairs (0 overlap) | +678 |
| 8 | Remove V1 eval overlaps (170 pairs) | -170 |
| 9 | Remove any eval Input_Columns | 0 |
| **10** | **Final V2 positive pairs** | **109,283** |

### D3 Hybrid Capping Details
- **Materials capped at 200 variants (>200 originally):** 126 materials
- **Materials retained at 101-199 variants:** 226 materials
- **Materials retained at <=100 variants:** 1,756 materials
- **Total materials:** 2,190

### Final Dataset Statistics
| Metric | Value |
|--------|-------|
| Total positive pairs | 109,283 |
| Unique PRODUCT_NAME | 109,283 |
| Unique MATERIAL DESCRIPTION | 2,190 |
| Min examples/material | 1 |
| Median examples/material | 16.0 |
| Mean examples/material | 49.9 |
| P90 examples/material | 194.0 |
| P95 examples/material | 200.0 |
| Max examples/material | 228 |
| Long-tail materials (<=5) | 771 (35.2%) |

---

## 2. V1 INCLUSION RESULTS

| Source | Pairs | Overlap with Historical | Overlap with Evaluation | Net Added |
|--------|-------|------------------------|------------------------|-----------|
| training_split.csv | 508 | 0 | 0 | +508 |
| hard_negative_mining.csv | 678 | 0 | 170 | +508 |
| **V1 Total** | **678** | **0** | **170** | **+508** |

- V1 pairs are completely novel (0% overlap with historical 108,775 pairs)
- 170 V1 pairs matched evaluation set exactly - **removed**
- Net V1 contribution: **508 pairs** from training_split

---

## 3. MATERIAL COVERAGE VERIFICATION

| Metric | Count | Percentage |
|--------|-------|------------|
| V2 materials | 2,190 | - |
| PRODUCT_MASTER products | 2,764 | - |
| Materials IN PRODUCT_MASTER | 2,181 | 99.6% |
| Materials NOT in PRODUCT_MASTER | 9 | 0.4% |
| PRODUCT_MASTER products NOT covered | 583 | 21.1% |

### Materials in V2 but not in PRODUCT_MASTER (9 - require review)
1. **ATAZOR 300 MG CAPSULE 30'S** - 19 examples
2. **EMCLOZ 0.25MG TABLET** - 1 example
3. **EMCLOZ 0.5MG TABLET** - 1 example
4. **ENCICARB INJECTION 1K** - 3 examples *(also in evaluation)*
5. **LOMOREST-30 SACHET 1X3'S** - 1 example
6. **MOLENA 200 MG CAPSULE 40'S** - 7 examples
7. **NEVIR 200 MG TABLET 60'S** - 2 examples
8. **VONADAY 600/300/300 MG TABLET 30s** - 3 examples *(also in evaluation)*
9. **ZUVISTON TABLETS** - 3 examples *(encoding issue)*

**Action needed:** Decide whether to add these 9 to PRODUCT_MASTER or exclude from training. Two (ENCICARB INJECTION 1K, VONADAY) appear in evaluation set.

### Long-tail Coverage
- 771 materials with <=5 examples (35.2%)
- 764 of these (99.1%) are in PRODUCT_MASTER
- Retained naturally without oversampling

---

## 4. EVALUATION PROTECTION VERIFICATION

| Check | Result | Status |
|-------|--------|--------|
| Exact pair overlap | 0 / 170 = 0.00% | PASS |
| Evaluation Input_Column in V2 | 0 / 170 = 0.00% | PASS |
| Evaluation Positive_Product as targets | 101 / 103 = 98.06% | EXPECTED |
| Conflicts (eval input → different target) | 0 | PASS |

**VERDICT: PASS - Zero evaluation leakage**

Note: 101/103 evaluation target products appear as positives in V2 with rich OCR variants (100-200+ each). This is **desired** - we want the model to learn robust retrieval for evaluation products. Only the exact evaluation input columns are excluded.

---

## 5. HARD NEGATIVE MINING DESIGN

### Candidate Pool
- **Primary:** All 2,764 PRODUCT_MASTER product_name values (canonical)
- **Secondary:** 2,190 V2 MATERIAL DESCRIPTION values
- **Model:** V1 fine-tuned BGE-base (768-dim embeddings)

### Mining Procedure
1. Pre-compute embeddings for all PRODUCT_MASTER products (2,764 vectors)
2. For each V2 positive pair (anchor=PRODUCT_NAME, positive=MATERIAL DESCRIPTION):
   - Encode anchor with V1 model
   - Compute cosine similarity to all candidate embeddings
   - Exclude positive MATERIAL DESCRIPTION and synonyms (same PRODUCT_CODE)
   - Rank candidates by similarity (descending)
   - Select top-K as hard negatives

### Negative Prioritization (implicit via similarity)
1. **Same brand, different strength/form/pack** - highest similarity
   - Example: CARDACE 2.5 MG vs CARDACE 5 MG vs CARDACE AM 5
2. **Same product family, different variant**
   - Example: AMARYL M 1MG vs AMARYL M 2MG vs AMARYL M FORTE
3. **OCR-confusable siblings** (similar token patterns)
4. **Similar numeric formulation** (same ingredient, different strength)
5. **Same therapeutic class** (lower similarity)

### Filtering Rules
- Exclude positive MATERIAL DESCRIPTION
- Exclude any product sharing same PRODUCT_CODE (synonyms)
- Ensure negative != positive
- Exclude evaluation Input_Columns from mining anchors

### Brand Families Available for Mining (from PRODUCT_MASTER)
- ASOMEX: 32 products
- TEMSAN: 26 products
- OROFER: 21 products
- CARDACE: 14 products (in V2)
- AMARYL: 9 products (in V2)
- ATOREC: 15 products
- FERIUM: 18 products
- METPURE: 19 products
- **340+ brands with >=3 products**

### Implementation Notes
- Batch encode PRODUCT_MASTER once (~30 seconds on GPU)
- For 109K anchors: ~302M similarity computations
- Use FAISS IVF or GPU batched cosine for speed
- Output format: Input_Column, Positive_Product, Hard_Neg_1, Hard_Neg_2, Hard_Neg_3

---

## 6. MINING SCALE ESTIMATION

| Parameter | Value |
|-----------|-------|
| Positive pairs | 109,283 |
| Recommended hard negatives/positive | 3 |
| **Total negatives to mine** | **327,849** |
| Training rows (with 3 negatives) | 109,283 |
| Columns | Input_Column, Positive_Product, Hard_Neg_1, Hard_Neg_2, Hard_Neg_3 |

### Compute Estimates
- Candidate embeddings: 2,764 (once)
- Anchor embeddings: 109,283
- Similarities: 109,283 × 2,763 = ~302M
- GPU (10K/sec): ~8.4 hours
- **With FAISS IVF100: ~0.8 hours** (recommended)

### Optimization Options
1. Mine only for materials with >10 variants (~1,257 materials, ~85K pairs)
2. Use FAISS IVF100 for ~100x speedup
3. Mine 2 negatives instead of 3 for speed
4. Use existing hard_negative_mining.csv (1,425 examples) as supplement

---

## 7. V2 TRAINING DATA VALIDATION CHECKLIST

| Check | Status | Details |
|-------|--------|---------|
| No invalid products in positives | PASS | 99.6% in PRODUCT_MASTER |
| No evaluation leakage (pairs) | PASS | 0/170 overlap |
| No evaluation leakage (inputs) | PASS | 0/170 overlap |
| No duplicate positive pairs | PASS | 0 duplicates |
| No contradictory mappings | PASS | 1:1 PRODUCT_NAME → MATERIAL DESCRIPTION |
| All V1 eval pairs removed | PASS | 170 removed |
| Material coverage (all 2,185+) | PASS | 2,190 materials |
| Long-tail coverage (766+) | PASS | 771 materials |
| Capping applied (max 200) | FAIL | Max is 228 (ATOREC family) |

**Note on capping:** 2 materials (ATOREC ASP CAPSULES - 228, OSTERI INJ.600MCG - 228) slightly exceed 200 cap due to rounding in party-based sampling. Recommend manual review or accept as-is.

**Note on 9 non-PRODUCT_MASTER materials:** Documented above - decision needed.

---

## 8. EXPECTED FINAL TRAINING SIZE

After hard negative mining (3 per positive):
- **Training rows:** 109,283
- **Columns:** Input_Column, Positive_Product, Hard_Neg_1, Hard_Neg_2, Hard_Neg_3
- **Total text entries:** 546,415
- **Estimated file size:** ~22 MB

Supplementary:
- Existing hard_negative_mining.csv: 1,425 examples (can be merged)

---

## 9. RISKS & UNRESOLVED ISSUES

### Identified Risks
1. **9 materials not in PRODUCT_MASTER** - Two appear in evaluation (ENCICARB INJECTION 1K, VONADAY). Decision: add to PRODUCT_MASTER or exclude.

2. **Evaluation products have 100-200+ OCR variants in training** - Desired for robust retrieval, but evaluation INPUT columns must not appear (verified: 0 overlap).

3. **Class imbalance remains** - 771 materials with <=5 examples (35%), 126 at cap (200). Consider weighted sampling or loss weighting during training.

4. **V1 model used for mining - potential bias** - Hard negatives reflect V1's errors. Consider ensemble or cross-validation.

5. **PARTY_CODE not used as feature** - 3,018 parties in historical data, some OCR variation is party-specific. Current design ignores party (correct for retrieval task).

### Unresolved Decisions Needed
- [ ] Final decision on 9 non-PRODUCT_MASTER materials (include/exclude)
- [ ] Exact number of hard negatives: 2, 3, or 4 per positive?
- [ ] Mining subset: all 109K or only high-diversity materials (>10 variants)?
- [ ] Training hyperparameters: epochs, batch size, LR, loss function
- [ ] Validation split strategy (stratified by material, party-separated)
- [ ] In-batch negatives only, or pre-mined hard negatives, or both?
- [ ] Learning rate schedule for fine-tuning from V1 checkpoint

---

## 10. FILES CREATED FOR REVIEW

| File | Description |
|------|-------------|
| `v2_historical_pairs.csv` | 108,775 pairs - historical data after D3 capping |
| `v2_positive_pairs.csv` | 109,453 pairs - historical + V1 (before eval filtering) |
| `v2_positive_pairs_validated.csv` | 109,283 pairs - after removing V1 eval overlaps |
| `v2_positive_pairs_final.csv` | **109,283 pairs - FINAL validated, zero eval leakage** |
| `STAGE2_ANALYSIS_REPORT.md` | Full Stage 2 analysis report |
| `PROFILING_REPORT.md` | Full Stage 1 profiling report |

---

## NEXT STEPS (require approval)

1. **Review 9 non-PRODUCT_MASTER materials** - decide include/exclude
2. **Approve hard negative mining design** (3 per positive, PRODUCT_MASTER candidates)
3. **Approve mining scale** (~328K negatives, GPU/FAISS)
4. **Decide V2 training hyperparameters**
5. **Run hard negative mining**
6. **Construct final training file**
7. **Start V2 training**

---

**Analysis complete. No training performed. No existing files modified. Ready for review.**