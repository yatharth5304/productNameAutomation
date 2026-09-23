# V2 Hard-Negative Mining — Cleaning Report

**Generated**: 2026-09-23 11:54:05

---

## 1. Overview

| Item | Value |
|------|-------|
| Source file | `v2_hard_negative_mining.csv` |
| Output file | `v2_hard_negative_mining_clean.csv` |
| Original row count | 108,913 |
| Rows dropped | 82 |
| Final row count | 108,831 |
| Schema preserved | Yes — identical 12 columns in original order |
| Retained rows modified | No — byte/value-equivalent to source |

## 2. Reason for Dropping 82 Rows

All 82 dropped rows contain **duplicate negative product names** caused by a single
confirmed catalogue alias:

| Alias name | Code A | Code B |
|-----------|--------|--------|
| `ARTH WEIGHT MANAGEMENT SUPPORT 15T` | `421113061` | `421113210` |

These two PRODUCT_CODEs refer to the same display name in PRODUCT_MASTER.xlsx.
The mining script excluded the positive by product code, so both alias records
passed the positive-exclusion filter and were independently selected as negatives,
occupying two of the three negative slots with the same normalised product name.

**Cleaning decision**: drop all 82 rows without replacement.
Rationale: the 82 rows represent 0.075% of the dataset. Dropping them keeps the
training data simple, clean, and free from catalogue-alias ambiguity.
No re-mining, no invented replacements, no evaluation data used.

## 3. Dropped Row Inputs (all 82)

```
  1. [data row   1304]  VISTAVIT CAPS-15s
  2. [data row   1949]  DINOR TAB 15 S
  3. [data row   2329]  VISTAVIT CAPS 15S 10X15
  4. [data row   3226]  DINOR 2 15TA
  5. [data row   4353]  CALCIEM TAB 15S
  6. [data row   6689]  DINOR 2MG TAB 15TAB
  7. [data row   7063]  FM ACTIVE TAB 15T
  8. [data row  10287]  CALCIEM 15 TAB
  9. [data row  13380]  Dinor Tab 15 s
 10. [data row  13890]  DINOR TABS 15T
 11. [data row  15216]  DINOR TAB 15S
 12. [data row  17726]  DINOR TAB 15T
 13. [data row  18274]  FM ACTIVE 1*15T
 14. [data row  20659]  DINOR TAB 15TAB
 15. [data row  21614]  CALCIEM TABS 15S
 16. [data row  23207]  ZILARBI-CN 10MG TA 15S
 17. [data row  25117]  DINOR TABS 15S
 18. [data row  25548]  Calciem Tab 15S
 19. [data row  26718]  FEMZACT 15 S
 20. [data row  33311]  CALCIEM TAB. 15*TAB
 21. [data row  33378]  CALCIEM TABS 15s
 22. [data row  33392]  DINOR 2MG TAB 15T
 23. [data row  33565]  AMYRIL MV2 15 S
 24. [data row  33700]  BIRYTH150 TAB 4 S
 25. [data row  35156]  ATOCOR 10MG TAB 15S
 26. [data row  36314]  DINOR TAB-15s
 27. [data row  36952]  DINOR TAB 15 s
 28. [data row  37181]  ARIPITEX 15MG TAB 10T
 29. [data row  37773]  VISTAVIT CAP 15*CAP
 30. [data row  38446]  CALCIEM TAB 1 X 15
 31. [data row  38862]  DINOR TAB 15  S
 32. [data row  40368]  FM ACTIVE TAB 15T 15T
 33. [data row  40732]  TENDOFIRM 15CAP
 34. [data row  43328]  UNMET TB 1*15
 35. [data row  43505]  DINOR 2MG TAB 15S
 36. [data row  46065]  TENDOFIRM CAP 15S
 37. [data row  46507]  DINOR 2MG TABS. 15S
 38. [data row  47600]  FM ACTIVE TAB 15 TAE
 39. [data row  49517]  ACTIVE TAB 15S 15S
 40. [data row  51416]  VYLDA TAB 15 T 15 S
 41. [data row  52923]  DINOR TABS 15S 15S
 42. [data row  53641]  CALCIEM TAB 15 TAB
 43. [data row  54349]  CALCIEM TAB. 15S
 44. [data row  54635]  DINOR 15T
 45. [data row  54924]  CALCIEM 15TAB
 46. [data row  56506]  DINOR TAB 15-S
 47. [data row  59389]  ARIPITEX 15 TAB 10
 48. [data row  61044]  VISTAVIT CAP 15
 49. [data row  61054]  DINOR DIENOGEST 2MG TAB 15S
 50. [data row  61691]  DINOR TAB. 15S
 51. [data row  62179]  ARIPITEX 15 TAB. 10S
 52. [data row  62450]  DINOR TAB 15s
 53. [data row  63135]  CALCIEM TAB 15S 1
 54. [data row  63795]  CALCIEM 15S 15S --
 55. [data row  64438]  FM-ACTIVE TAB 1*15TA
 56. [data row  65440]  VISTAVIT CAPS 15S
 57. [data row  70345]  VISTAVIT CAPS 15CAPS
 58. [data row  71651]  FM ACTIVE TAB 15TB
 59. [data row  73168]  Vistavit Caps 1*15
 60. [data row  76810]  FM ACTIVE 15TAB
 61. [data row  77163]  VISTAVIT CAP 15S
 62. [data row  77738]  CALCIEM-TAB 15S
 63. [data row  78446]  ARIPITEX 15GM 10S
 64. [data row  79006]  CALCIEM TAB 15 S
 65. [data row  80005]  ARIPITEX-15 TAB 10TAB
 66. [data row  80022]  VIL G 50 15STP
 67. [data row  81576]  ARTH SLEEP SUPPORT GUMMIES 15 15S
 68. [data row  85281]  FM ACTIVE (15TAB) 15TAB
 69. [data row  90269]  ATORMAC CV 10 CAP 15S
 70. [data row  91631]  FM ACTIVE 10/15TAB 15T
 71. [data row  93759]  CALCIEM TAB 15S 15S
 72. [data row  94256]  CALCIEM TAB 15TAB
 73. [data row  94258]  DINOR-TABLET 15T
 74. [data row  95328]  CALCIEM TAB 15-S
 75. [data row  96345]  DINOR TAB 2MG 15 S
 76. [data row  98045]  TEPARO 15 MG INJ(THIC 1
 77. [data row  99823]  DINOR TAB 15S 15S
 78. [data row 100318]  ARTH SLEEP SUPPORT GUMMIES 15S (EMCURE) 15 GMS
 79. [data row 104355]  NUMLOAT TAB 15S
 80. [data row 106407]  DINOR 2mg 15TAB
 81. [data row 106856]  VISTAVIT CAP 15TAB
 82. [data row 108234]  VISTAVIT CAPS 15s
```

## 4. Validation Results

The clean CSV was reloaded from disk and all 10 checks were run independently
against the written file.

| # | Check | Result | Detail |
|---|-------|--------|--------|
| 1 | row count | **PASS** | Row count = 108,831 |
| 2 | all 3 negatives | **PASS** | 0 rows missing negatives |
| 3 | no dup neg names | **PASS** | 0 rows with dup normalised neg names |
| 4 | no dup neg codes | **PASS** | 0 rows with dup neg codes |
| 5 | no code collision | **PASS** | 0 positive/negative code collisions |
| 6 | no name collision | **PASS** | 0 positive/negative name collisions |
| 7 | no missing fields | **PASS** | 0 rows with missing fields |
| 8 | valid scores | **PASS** | 0 unparseable score values |
| 9 | no eval leakage | **PASS** | 0 rows match eval inputs |
| 10 | rows unchanged | **PASS** | 0 rows differ from original |

> [!NOTE]
> All 10 validation checks passed. The clean CSV is ready for V2 training.

## 5. Evaluation Leakage Confirmation

Evaluation inputs checked: 170

**Result: ZERO evaluation input leakage.**
No training `Input_Column` matches any evaluation `Input_Column` (exact string match).

## 6. High-Similarity Negatives

Per cleaning decision: high-similarity negatives were **not removed**.
They were audited and confirmed to be legitimate hard negatives (different strengths,
formulations, pack sizes, or variants with distinct PRODUCT_CODEs).

| Threshold | Count (neg-slots) | Action |
|-----------|-------------------|--------|
| >= 0.95 | 440 | Retained |
| >= 0.98 | 28  | Retained |
| >= 0.99 | 8   | Retained |
| = 1.00  | 2   | Retained (verified distinct catalogue records) |

## 7. Schema

The clean CSV uses the identical 12-column schema as the source file:

```
  Input_Column
  Positive_Product
  Positive_Product_Code
  Mined_Negative_1
  Mined_Negative_1_Code
  Mined_Negative_1_Score
  Mined_Negative_2
  Mined_Negative_2_Code
  Mined_Negative_2_Score
  Mined_Negative_3
  Mined_Negative_3_Code
  Mined_Negative_3_Score
```

## 8. File Locations

| File | Description |
|------|-------------|
| `v2_hard_negative_mining.csv` | Raw mining output — untouched |
| `v2_hard_negative_mining_clean.csv` | Clean training dataset — 108,831 rows |
| `V2_HARD_NEGATIVE_AUDIT.md` | Full audit report |
| `V2_HARD_NEGATIVE_CLEANING_REPORT.md` | This report |

## 9. Summary

| Metric | Value |
|--------|-------|
| Original rows | 108,913 |
| Rows dropped (dup negatives) | 82 |
| Final clean rows | 108,831 |
| Drop rate | 0.075% |
| All 10 validations | PASSED |
| Eval leakage | ZERO |
| Rows modified | NONE |

---
*End of cleaning report.*