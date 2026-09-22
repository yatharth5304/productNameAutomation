# V2 POSITIVE DATASET — FINAL VALIDATION REPORT

## Executive Summary

**VERDICT: READY FOR HARD-NEGATIVE MINING ✓**

All 13 validation checks pass. The V2 positive dataset has been reconstructed with:
- 9 excluded materials removed completely (0 rows remaining)
- Maximum 200 examples per material enforced (no material exceeds cap)
- 100% PRODUCT_MASTER coverage for all remaining materials
- Zero evaluation leakage (pair and input level)
- Zero mapping conflicts or duplicates

---

## Validation Results

| # | Check | Result | Status |
|---|-------|--------|--------|
| 1 | Final positive pairs | **108,913** | — |
| 2 | Unique materials | **2,181** | — |
| 3 | Examples per material (min/median/mean/P90/P95/max) | **1 / 16.0 / 49.9 / 194.0 / 200.0 / 200** | — |
| 4 | Max examples ≤ 200 | **200** | **PASS** |
| 5 | 9 excluded materials have 0 rows | **0 rows** | **PASS** |
| 6 | Zero duplicate positive pairs | **0** | **PASS** |
| 7 | Zero mapping conflicts (PN → multiple MD) | **0** | **PASS** |
| 8 | Zero evaluation Input_Column leakage | **0 / 170** | **PASS** |
| 9 | Zero exact evaluation pair leakage | **0 / 170** | **PASS** |
| 10 | Evaluation set unmodified | **364 rows** | **PASS** |
| 11 | All targets in PRODUCT_MASTER | **2,181 / 2,181 = 100%** | **PASS** |
| 12 | V1 pairs correctly retained | **486** (from training_split) | **PASS** |
| 13 | Material coverage statistics | See below | **PASS** |

---

## Material Coverage Statistics

| Metric | Value |
|--------|-------|
| V2 materials in PRODUCT_MASTER | 2,181 / 2,181 = **100.0%** |
| PRODUCT_MASTER products covered | 2,181 / 2,764 = **78.9%** |
| PRODUCT_MASTER products NOT covered | 583 |
| Long-tail materials (≤5 examples) | **764** |
| Long-tail in PRODUCT_MASTER | 764 / 764 = **100.0%** |
| Materials at cap (200) | **208** |

### Variant Distribution (unique PRODUCT_NAME per MATERIAL DESCRIPTION)

| Bucket | Materials | Percentage |
|--------|-----------|------------|
| 1 | 265 | 12.1% |
| 2 | 194 | 8.9% |
| 3–5 | 305 | 14.0% |
| 6–10 | 185 | 8.5% |
| 11–25 | 316 | 14.5% |
| 26–50 | 262 | 12.0% |
| 51–100 | 220 | 10.1% |
| 101–200 | 434 | 19.9% |
| **Total** | **2,181** | **100.0%** |

---

## Excluded Materials (9) — Confirmed Zero Rows

1. ATAZOR 300 MG CAPSULE 30'S
2. EMCLOZ 0.25MG TABLET
3. EMCLOZ 0.5MG TABLET
4. ENCICARB INJECTION 1K
5. LOMOREST-30 SACHET 1X3'S
6. MOLENA 200 MG CAPSULE 40'S
7. NEVIR 200 MG TABLET 60'S
8. VONADAY 600/300/300 MG TABLET 30s
9. ZUVISTON TABLETS (encoding variant: `ZUVISTON TABLETS�`)

**Total historical rows removed for exclusions:** 31

---

## V1 Integration Summary

| Source | Pairs Available | Overlap with Historical | Overlap with Evaluation | Excluded (not in PRODUCT_MASTER) | **Added to V2** |
|--------|-----------------|------------------------|------------------------|----------------------------------|----------------|
| training_split.csv | 508 | 0 | 0 | 22 | **486** |
| hard_negative_mining.csv | 678 | 0 | 170 | 22 | **0** |
| **V1 Total** | **678** | **0** | **170** | **22** | **486** |

- V1 training_split contributed 486 novel positive pairs
- V1 hard_negative_mining pairs were all duplicates of training_split (same 678 unique pairs)
- 170 V1 pairs matched evaluation exactly — **excluded**
- 22 V1 pairs had targets not in PRODUCT_MASTER — **excluded**

---

## Dataset Construction Pipeline (Final)

1. **Load Sheet1**: 317,054 rows
2. **Remove NO PRODUCT**: -3,152 rows (MATERIAL DESCRIPTION = "-")
3. **Remove exact duplicates**: -1,782 rows
4. **Clean usable**: 312,120 rows
5. **Exclude 9 materials**: -31 rows
6. **Unique positive pairs**: 148,060
7. **D3 Hybrid capping** (retain ≤100, retain 101–200, cap >200 at 200): **108,746 historical pairs**
8. **Add V1 training_split (filtered)**: +486 pairs
9. **Final recapping** (82 materials exceeded 200 after V1 addition): -319 pairs
10. **Final V2 positive pairs**: **108,913**

---

## Final Dataset Files

| File | Description |
|------|-------------|
| `v2_positive_pairs_final.csv` | **108,913 pairs — FINAL validated dataset** (Input_Column=PRODUCT_NAME, Positive_Product=MATERIAL DESCRIPTION) |
| `v2_historical_pairs_final.csv` | 108,746 pairs — historical only (after D3 capping, before V1 addition) |

---

## Ready for Hard-Negative Mining

**YES** — The V2 positive dataset passes all validation criteria and is ready for the next phase:

- ✅ 108,913 positive anchors
- ✅ 2,181 unique MATERIALS (all in PRODUCT_MASTER)
- ✅ Zero evaluation leakage
- ✅ Zero conflicts/duplicates
- ✅ Max 200 examples/material enforced
- ✅ Long-tail (764 materials) retained naturally
- ✅ OCR diversity preserved (median 16 variants/material, 434 materials with 101–200 variants)

**Next step:** Run hard-negative mining using the approved design (PRODUCT_MASTER candidates, V1 model, FAISS IVF, 3 negatives per positive).

---

*No training performed. No existing files modified. evaluation_set.csv and PRODUCT_MASTER.xlsx unchanged.*