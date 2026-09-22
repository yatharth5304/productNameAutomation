# GARBAGE_REVIEW_ANALYSIS

## Executive Summary

| Metric | Value |
|---|---|
| Total rows in testRaw.xlsx (garbage_check sheet) | 410,298 |
| Rows classified REVIEW | **1,905** |
| Confirmed-garbage candidates (CANDIDATE FOR 1) | **375 (19.7%)** |
| -- STRONG evidence | 140 |
| -- MODERATE evidence | 235 |
| Legitimate Review / ambiguous cases (KEEP REVIEW) | **1,530 (80.3%)** |
| Rows current code would NOW classify as 1 (version lag) | 110 |

> [!IMPORTANT]
> The REVIEW rows in testRaw.xlsx were produced by an **earlier run** of garbage_check.py.
> Some patterns in SAFE_GARBAGE_PATTERNS were added **after** testRaw.xlsx was generated.
> 110 rows carry REVIEW today that the current version would already classify as 1.
> This is the single most important finding.

---

## 1. Remark Distribution in testRaw.xlsx

```
0             -> 397,555   (confirmed product, local brand match)
MAYBE_PRODUCT ->   6,953   (dosage/form/qty pattern, no brand match)
1             ->   3,885   (confirmed garbage)
REVIEW        ->   1,905   (unresolved, pending LLM or manual review)
Total         -> 410,298
```

---

## 2. Classification Logic: Exact Decision Tree

The main() loop in garbage_check.py processes each row in this order:

```
1. exact_product_match()      -> result "0" (LOCAL_PRODUCT_MATCH)
2. fuzzy_product_match()      -> result "0" (LOCAL_PRODUCT_MATCH)
   [brand exception checks run after brand is matched]
3. is_exact_garbage_row()     -> result "1" (GARBAGE:exact_garbage)
4. is_pattern_garbage_row()   -> result "1" (GARBAGE:pattern_garbage)
5. is_metadata_garbage_row()  -> result "1" (GARBAGE:metadata_garbage)
6. fuzzy_garbage_match()      -> result "1" (FUZZY_GARBAGE)
7. if local_match in LOCAL_REVIEW_BRANDS -> REVIEW (LOCAL_REVIEW_BRANDS_PENDING_LLM)
8. is_maybe_product_row()     -> MAYBE_PRODUCT (or REVIEW if False)
   -> if False:               -> REVIEW (NO_MATCH_FALLBACK)
```

In local mode (PROCESSING_MODE = "local"): rows reaching steps 7-8 get REVIEW/MAYBE_PRODUCT directly.

REVIEW rows come from exactly two paths:
- **Path A**: LOCAL_REVIEW_BRANDS_PENDING_LLM -- brand matched, brand is in LOCAL_REVIEW_BRANDS (currently **empty set**)
- **Path B**: NO_MATCH_FALLBACK -- no brand match, no garbage check fired, not MAYBE_PRODUCT

In testRaw.xlsx, **all 1,905 REVIEW rows came through Path B** (NO_MATCH_FALLBACK).
LOCAL_REVIEW_BRANDS is currently empty, so Path A is inert.

---

## 3. The 110-Row Version Discrepancy

**110 REVIEW rows** in testRaw.xlsx are REVIEW but the **current** garbage_check.py would classify
them as 1 (confirmed garbage). These patterns were added after testRaw.xlsx was generated:

- `\b(AGENCIES?|DISTRIBUTORS?|MEDICO[SE]?|CORPORATION|ENTERPRISES?|TRADERS?|TRADING)\b`
- `\b(PVT\.?\s+)?LTD\.?\s*$`
- `\bLIMITED\.?\s*$`
- `\b(BEDSHEET|BEDSEET|HOTPOT|TOWEL|...|WATER\s+JAR)\b`

Affected rows: `HAJI ENTERPRISES` (x5), `BAREILLY PHARMACEUTICALS PVT LTD` (x3),
`M/S ECONOMIC TRADERS` (x4), `MEDICASH AGENCIES`, `BALLABH AGENCIES`,
`AVENTIS PHARMA LIMITED`, `NUCRON PHARMACEUTICALS LIMITED`, `GARHWAL DRUG AGENCIES PVT. LTD.`,
`FDC BEDSEET`, `MILTON WATER JAR`.

> [!NOTE]
> These 110 rows are **already solved**. They become 1 automatically on the next full re-run
> of garbage_check.py with CLEAR_EXISTING = True.

---

## 4. Why REVIEW Rows Reach REVIEW: Pathway Analysis

| Pathway Group | Count | Description |
|---|---|---|
| FALLBACK:AMBIGUOUS_NO_BRAND | 1,512 | No brand match, no garbage match, no pharma signal |
| FALLBACK:STRUCTURAL_KW | 141 | Structural keyword present but MAYBE_PRODUCT_GUARD blocked |
| GARBAGE:pattern_garbage | 110 | Already caught by current patterns (version lag) |
| FALLBACK:DISTRIBUTION_ENTITY | 50 | Distribution/medical entity text |
| FALLBACK:PHARMA_ENTITY | 49 | Pharma company names |
| FALLBACK:VERY_SHORT | 43 | Very short text (<=5 alphanumeric chars) |

### Why 1,512 rows fall through all checks (FALLBACK:AMBIGUOUS_NO_BRAND)

Root causes:

1. **Glued OCR tokens** -- fused without spaces: `OROFERCAPS`, `ENHEMO1000MGINJECTION`, `FEMZACTCAPS`, `VONADAY600/300/3001X`
2. **Non-Emcure product names** -- `CURENAL 4T`, `NICOACE PLUS`, `ZOLINEG 1MIU INJECTION`
3. **Aggregate/financial rows with variable numerics** -- `Totals: 464937.00 --`, `Purchase 0.00 Purchase Return 0.00`
4. **Garbled product codes** -- `DAPAGZAM10/1000 10`, `GALACTGRANULESE1X400 G`
5. **Company names without suffix patterns** -- `KRISHNA PHARMA`, `MAHESH PHARMA`, `NUCRON`, `NUSURGE`
6. **ZZZZZZ placeholder rows** -- `ZZZZZZ 5015`, `ZZZZZZ 4024`, `ZZZZZZ 6162`

---

## 5. Detailed Failure Pattern Analysis

### Pattern 1: INSUFFICIENT_EVIDENCE -- 1,530 rows (80.3%)

Cannot be confidently classified as garbage from existing rules alone.

**Sub-group A: Legitimate products where brand matching failed (~750 rows)**
Real Emcure/pharma products with glued OCR tokens or non-Emcure brands not in vocabulary.
These should remain REVIEW.

**Sub-group B: Financial/aggregate rows with variable numeric content (~200 rows)**
`Totals: 464937.00 --`, `Totals 175972.00 287469.00 152052.00 427.49`,
`Purchase 0.00 Purchase Return 0.00`, `Totals 110196.00 94121.00 3556.10`.

Not caught because:
- EXACT_GARBAGE_ROWS uses exact normalized matching; numerics change every row
- SAFE_GARBAGE_PATTERNS has `^Purchase Return` but NOT `^Purchase \d+` or `^Totals \d+`

**Sub-group C: Emcure division names without existing entries (~100 rows)**
`NUCRON CV`, `NUSURGE`, `INVENTIA(INVENTIA)`, `XENNEX 1 (EMC)` -- division names not in
EXACT_GARBAGE_ROWS as standalone entries.

---

### Pattern 2: STRUCTURAL_KEYWORD -- 101 rows (5.3%) -> CANDIDATE FOR 1

Strong structural keywords (TOTAL, PURCHASE, DIVISION, OPENING) escaped all garbage checks.

Examples:
- `TOTAL CAL D 3MG TAB.20S 1X20S` -- product-like name with TOTAL
- `CAL-123 TOTAL TAB` -- product-like name with TOTAL
- `Purchase Value 0.00 Pur. Return Value` -- financial row
- `VALUE OF OPENING STOCK AS ON -` -- structural row
- `Total Of EMC 1`, `Total Of EMC 8 -` -- company total lines
- `GINNOVA BIOPHARMACEUTICALS LTD. Total` -- supplier total
- `NUCRON DIVISION`, `INVENTIA DIVISION`, `PHARMA DIVISION`, `PHARMA NXT DIVISION`

Why: MAYBE_PRODUCT_GUARD_PATTERN blocks TOTAL/PURCHASE/VALUE from becoming MAYBE_PRODUCT,
but no corresponding garbage rule fires for these. The `^Total\s*[:(]` pattern exists but
`Total Of EMC 1` does not match it (no colon/paren after Total).

> [!WARNING]
> `TOTAL CAL D 3MG TAB.20S 1X20S` and `GLYCIRNOM TOTAL 30 TAB 10 TAB` are ambiguous.
> If TOTAL is part of a product name (not an aggregate), classifying as 1 would be wrong.
> The LLM prompt explicitly warns about this. These two rows should stay REVIEW.

---

### Pattern 3: COMPANY_DIVISION_PREFIX -- 79 rows (4.1%) -> CANDIDATE FOR 1

Rows starting with known Emcure division prefixes: NUCRON, NUSURGE, INVENTIA, XENNEX,
EMCUTIX, AVENTIS, IMPETUS, VINTOR.

Examples: `NUCRON` (x2), `NUSURGE`, `INVENTIA(INVENTIA)`, `XENNEX 1 (EMC)`, `XENNEX --`,
`NUCRON CV`, `NUCRON CV DIVISION`, `NUCRON DIVISION`, `INVENTIA DIVISION`,
`PHARMA DIVISION`, `PHARMA NXT DIVISION`, `AVENTIS PHARMA LIMITED`.

Why: EXACT_GARBAGE_ROWS has `EMCURE INVENTIA`, `EMCURE NUCRON`, `EMCURE XENNEX` but NOT
the bare standalone forms. SAFE_GARBAGE_PATTERNS has `\bSANOFI\b` but not these.

---

### Pattern 4: COMPANY_ENTITY -- 58 rows (3.0%) -> CANDIDATE FOR 1

Company entities with PHARMACEUTICALS/BIOPHARMA/HEALTHCARE keywords:
- `BAREILLY PHARMACEUTICALS PVT LTD` (x3) -- already caught by LTD pattern (version lag)
- `CHHINDWARA PHARMACEUTICALS` (x2) -- no LTD/AGENCIES suffix, escapes all patterns
- `NUCRON PHARMACEUTICALS LIMITED` -- already caught (version lag)
- `Mfg. NUCRON PHARMACEUTICAL LTD` -- `Mfg.` prefix confuses `^`-anchored patterns
- `I am satisfied with Marg ERP Computerise YOUR CHEMIST SHOP Rs.5400 Only` -- software ad

---

### Pattern 5: PHARMA_COMPANY_NAME -- 60 rows (3.2%) -> CANDIDATE FOR 1

Company-like names with PHARMA but no dosage form:
- `KRISHNA PHARMA`, `MAHESH PHARMA`, `PHARMA NXT`, `PHARMA` -- distributor names
- `PHARMA TRADERS` (x5) -- already caught by TRADERS pattern (version lag)
- `MASTER MEDICO` -- already caught by MEDICO pattern (version lag)

Why: Not in EXACT_GARBAGE_ROWS; no AGENCIES/DISTRIBUTORS/ENTERPRISES suffix in
SAFE_GARBAGE_PATTERNS; `PHARMA` alone is not in any existing pattern.

---

### Pattern 6: NO_VOWEL_OCR_NOISE -- 34 rows (1.8%) -> CANDIDATE FOR 1

All 34 are `ZZZZZZ NNNN` rows: `ZZZZZZ 5015`, `ZZZZZZ 4024`, `ZZZZZZ 6162`, etc.
OCR placeholder/error sentinel values used by some report generators.

Why: ZZZZZZ has 6 alphanumeric chars, exceeding the <=3 alnum short-text threshold.
No brand matches. No existing pattern targets ZZZZZZ.

---

### Pattern 7: INVALIDATED_NO_PHARMA -- 33 rows (1.7%) -> CANDIDATE FOR 1

A brand was detected then invalidated by an exception rule. Row fell to REVIEW because
it has no pharma dosage/qty signal.

Key examples:
- `SYLATE 500 NEW 1X10TAB` -- NEW brand matched, invalidated (no NORMET). Has TAB/1X10 signals;
  should have become MAYBE_PRODUCT in current code. Suggests version lag.
- `SUSTIMAX 4T(GKM NEW PHARMA)` -- NEW brand matched in parenthetical, invalidated.
  `4T` not captured by MAYBE_PRODUCT_QTY_PATTERN.
- `EMCOR CREM 15 M` -- EMCOR matched, invalidated (no CREAM/TUBE; `CREM` != `CREAM`).
- `EMCOR` -- bare name, matched then invalidated -> REVIEW.
- `PHARMED ACE` -- no brand match; OCR corruption of a product name.

---

### Pattern 8: SHORT_TEXT -- 3 rows -> CANDIDATE FOR 1

`O.B.` (x3) -- 2 alphanumeric chars. Existing short-text pattern would match this but
rows were processed before the pattern was added (version lag).

---

### Pattern 9: SUPPLY_CHAIN_ENTITY -- 7 rows -> CANDIDATE FOR 1

`OM DRUG HOUSE` (x3), `SPECIALITY DRUG HOUSE`, `SHIPRA MEDICAL STORE` (x2),
`SANJIVANI MEDICINES PRIVATE LIMITEDB/H...` (OCR-glued company + address).

---

## 6. Legitimate REVIEW Rows (Should NOT Become 1)

| Type | Estimated Count | Examples |
|---|---|---|
| Non-Emcure pharmaceutical products | ~500-700 | `CURENAL 4T`, `NICOACE PLUS 10T`, `ZOLINEG 1MIU INJECTION`, `LAXIHUSK GRANULES` |
| Glued Emcure brand+suffix (OCR fusion) | ~300-400 | `OROFERCAPS`, `FEMZACTCAPS`, `ENHEMO1000MGINJECTION`, `DYDROFEMTAB`, `VONADAY600/300/3001X` |
| Variable financial rows (needs new rules) | ~200 | `Totals: 464937.00 --`, `Purchase 0.00 Purchase Return 0.00` |
| Ambiguous short tokens | ~100 | `NUCARE`, `FLORINA`, `EMC IVA` |

---

## 7. Rule/Function Responsibility in garbage_check.py

| Rule / Function | Rows Affected | Nature of Gap |
|---|---|---|
| EXACT_GARBAGE_ROWS (lines 97-310) | ~79 Emcure division names | Missing bare forms: NUCRON, NUSURGE, INVENTIA, XENNEX |
| SAFE_GARBAGE_PATTERNS (lines 312-421) | 110 version-lag + ~180 genuinely missing | No Totals pattern, no ZZZZZZ, no standalone division names |
| is_metadata_garbage_row() | 0 in REVIEW | Working correctly |
| fuzzy_product_match() (lines 658-737) | ~400 glued-token products | Cannot isolate brand span from fused OCR tokens |
| is_maybe_product_row() | 33 invalidated rows | Version lag: should be MAYBE_PRODUCT in current code |
| LOCAL_REVIEW_BRANDS | 0 | Empty set, inert |

**Primary failure point**: The NO_MATCH_FALLBACK path (lines ~1171-1181) is responsible for
all 1,905 rows. It is the final `else` clause: if is_maybe_product_row() returns False and
no garbage rule fired, the row becomes REVIEW. The upstream rules are the actual gaps.

---

## 8. Risk Assessment for Potential Fixes

| Proposed Fix | Rows Addressed | Classification Risk |
|---|---|---|
| `^Totals?\s+[\d,]+...` pattern | ~200 financial rows | **Very LOW** |
| `^Purchase\s+[\d.]+\s+Purchase\s+Return` | ~30 rows | **Very LOW** |
| `\bZZZZZZ\b` pattern | 34 rows | **ZERO** |
| Bare division names (NUCRON, NUSURGE, XENNEX) | ~79 rows | **LOW** (verify first) |
| PHARMACEUTICALS keyword pattern | ~20 rows | **MODERATE** (verify first) |
| PHARMA-suffix company names | ~40 rows | **MODERATE** (verify first) |
| `^[A-Z]+\s+DIVISION$` pattern | ~20 rows | **LOW** |

---

## 9. Quantitative Summary

| Category | Count | % of REVIEW |
|---|---|---|
| **Total REVIEW rows** | 1,905 | 100% |
| Already solved by current code (version lag) | 110 | 5.8% |
| Strong garbage candidates | 140 | 7.4% |
| Moderate garbage candidates | 235 | 12.3% |
| **Total garbage candidates** | **375** | **19.7%** |
| Legitimate review cases (real products / ambiguous) | ~930 | ~48.8% |
| Financial/aggregate rows (need new rules) | ~200 | ~10.5% |
| KEEP REVIEW total | 1,530 | 80.3% |

### Dominant Failure Patterns (Ranked by Impact)

| Rank | Failure Pattern | Count | Responsible Rule |
|---|---|---|---|
| 1 | Version lag (patterns added post-run) | 110 | SAFE_GARBAGE_PATTERNS |
| 2 | Variable financial rows (Totals N, Purchase N Return N) | ~200 | SAFE_GARBAGE_PATTERNS (missing) |
| 3 | Missing bare Emcure division names | 79 | EXACT_GARBAGE_ROWS |
| 4 | PHARMA-suffix distributor names | 60 | SAFE_GARBAGE_PATTERNS (missing) |
| 5 | Missing PHARMACEUTICALS-suffix companies | 58 | SAFE_GARBAGE_PATTERNS (missing) |
| 6 | OCR sentinel values (ZZZZZZ) | 34 | SAFE_GARBAGE_PATTERNS (missing) |
| 7 | Invalidated-brand, no-pharma-signal rows | 33 | is_maybe_product_row() version lag |
| 8 | Supply chain entities | 7 | SAFE_GARBAGE_PATTERNS version lag |
| 9 | Ultra-short text (O.B.) | 3 | SAFE_GARBAGE_PATTERNS version lag |

---

## 10. Recommended Modifications (NOT IMPLEMENTED -- analysis only)

### High Priority, Low Risk

**A. Add to SAFE_GARBAGE_PATTERNS:**

```python
# Variable financial aggregate rows
re.compile(r"^Totals?\s+[\d,]+(?:\.\d+)?(?:\s+[\d,]+(?:\.\d+)?)*\s*$", re.IGNORECASE),
re.compile(r"^Purchase\s+[\d,]+(?:\.\d+)?\s+Purchase\s+Return\s+[\d,]+", re.IGNORECASE),
# OCR sentinel placeholder
re.compile(r"^Z{4,}\b", re.IGNORECASE),
# Bare Emcure division names (verify against PRODUCT_MASTER first)
re.compile(r"^(NUCRON|NUSURGE|INVENTIA|XENNEX|EMCUTIX)(\s+\w+)*$", re.IGNORECASE),
```

**B. Add to EXACT_GARBAGE_ROWS:**

```python
"NUCRON CV", "NUCRON DIVISION", "INVENTIA DIVISION",
"XENNEX DIVISION", "PHARMA DIVISION", "PHARMA NXT DIVISION",
```

**C. Add PHARMACEUTICALS to AGENCIES pattern** (verify against master data first):

```python
re.compile(r"\bPHARMACEUTICALS?\b", re.IGNORECASE),
```

> [!CAUTION]
> PHARMACEUTICALS could appear in legitimate product column text from suppliers.
> Must verify against actual master data before adding.

### Medium Priority

**D. PHARMA-suffix distributor names**: Pattern like `^[A-Z][A-Za-z\s]+\bPHARMA\b$` catches
`KRISHNA PHARMA`, `MAHESH PHARMA` -- requires careful boundary verification.

### Not Recommended (High Risk)

- Broad TOTAL addition to garbage -- `TOTAL CAL D 3MG` could be a real product name
- Broad DIVISION word addition -- appears in legitimate product column context

---

## 11. Per-Row Data

See garbage_review_analysis.csv for full per-row detail.

### High-Confidence Garbage Candidates (Sample)

| Row | OCR Text | Assessment | Failure Pattern |
|---|---|---|---|
| 10396 | ZZZZZZ 5015 | CANDIDATE FOR 1 | NO_VOWEL_OCR_NOISE |
| 11014 | NUSURGE | CANDIDATE FOR 1 | COMPANY_DIVISION_PREFIX |
| 11088 | NUCRON | CANDIDATE FOR 1 | COMPANY_DIVISION_PREFIX |
| 146015 | O.B. | CANDIDATE FOR 1 | SHORT_TEXT |
| 151855 | VALUE OF OPENING STOCK AS ON - | CANDIDATE FOR 1 | STRUCTURAL_KEYWORD |
| 171552 | EMCOR | CANDIDATE FOR 1 | INVALIDATED_NO_PHARMA |
| 38319 | XENNEX 1 (EMC) | CANDIDATE FOR 1 | COMPANY_DIVISION_PREFIX |
| 40502 | XENNEX -- | CANDIDATE FOR 1 | COMPANY_DIVISION_PREFIX |
| 276619 | SHIPRA MEDICAL STORE | CANDIDATE FOR 1 | SUPPLY_CHAIN_ENTITY |
| 107996 | OM DRUG HOUSE | CANDIDATE FOR 1 | SUPPLY_CHAIN_ENTITY |
| 35694 | I am satisfied with Marg ERP... | CANDIDATE FOR 1 | STRUCTURAL_KEYWORD |
| 276614 | Sale of period 50164.15 GST... | CANDIDATE FOR 1 | STRUCTURAL_KEYWORD |

### Legitimate Review Rows (Should Remain REVIEW)

| Row | OCR Text | Why REVIEW |
|---|---|---|
| 719 | CURENAL 4T | Non-Emcure product |
| 832 | NICOACE PLUS 10T | Non-Emcure product |
| 887 | SAV HF100 10T | Non-Emcure product |
| 6679 | DYDROFEMTAB | Glued Emcure token (DYDROFEM+TAB) |
| 6693 | OROFERCAPS | Glued OROFER+CAPS |
| 5638 | ZOLINEG 1MIU INJECTION 1MIU | Non-Emcure pharmaceutical |
| 3821 | VONADAY600/300/3001X | Glued token, VONADAY is Emcure brand |

---

*Analysis performed: 2026-09-22*
*testRaw.xlsx: 410,298 rows | REVIEW rows: 1,905 | Brands: 849 | Mode: local*
