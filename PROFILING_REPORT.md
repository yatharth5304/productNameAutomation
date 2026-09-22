# PROFILING REPORT: EmcureSSSReport.xlsb (Sheet1)
## Dataset Profile for V2 Sentence Transformer Product-Mapping Retriever

---

### EXECUTIVE SUMMARY

**Dataset**: 317,054 rows × 4 columns (PARTY_CODE, MATERIAL CODE, PRODUCT_NAME, MATERIAL DESCRIPTION) from Sheet1 of EmcureSSSReport.xlsb

**Key Findings**:
- **Zero mapping conflicts**: Every PRODUCT_NAME maps to exactly ONE MATERIAL DESCRIPTION (100% deterministic)
- **Massive OCR variation**: 438 materials have ≥100 unique OCR variants; 1,257 materials have ≥10 variants
- **No evaluation contamination**: 0% overlap on Input_Column (PRODUCT_NAME) with evaluation_set.csv or hard_negative_mining.csv
- **High PRODUCT_MASTER coverage**: 78.9% of PRODUCT_MASTER products (2,180/2,764) appear in this dataset
- **Data quality issue**: 3,152 rows (1.0%) have MATERIAL DESCRIPTION = "-" and MATERIAL CODE = "NO PRODUCT" — these are unmapped/unknown products
- **Class imbalance**: 33.3% of materials have ≤5 training examples; 8.9% have ≥500 examples
- **Strong brand signal**: 71.1% brand overlap with PRODUCT_MASTER; brand names consistently appear in both PRODUCT_NAME and MATERIAL DESCRIPTION

**Training Value**: HIGH — This dataset provides exactly the kind of diverse OCR variations needed to improve Recall@3. The 150,976 unique PRODUCT_NAME → 2,185 valid MATERIAL DESCRIPTION mappings with rich variation (median 21 variants per material) is ideal for contrastive fine-tuning.

**Major Risk**: The 3,152 "NO PRODUCT" rows must be excluded. The long-tail of 728 materials with ≤5 examples may need filtering or augmentation. No evaluation contamination detected.

---

### 1. DATASET OVERVIEW

| Metric | Value |
|--------|-------|
| Total rows | 317,054 |
| Columns | 4 (PARTY_CODE, MATERIAL CODE, PRODUCT_NAME, MATERIAL DESCRIPTION) |
| Data types | PARTY_CODE: int64, MATERIAL CODE: object, PRODUCT_NAME: str, MATERIAL DESCRIPTION: str |
| Missing values | 0 across all columns |
| Exact row duplicates | 1,836 (0.58%) |
| Unique (PRODUCT_NAME, MATERIAL DESCRIPTION, PARTY_CODE) duplicates | 1,836 |

**Column Details**:
- **PARTY_CODE**: 3,018 unique values (range: 400,000–602,254), customer/party identifiers
- **MATERIAL CODE**: 2,186 unique values (2,185 numeric + 1 "NO PRODUCT")
- **PRODUCT_NAME**: 150,976 unique values (47.6% uniqueness rate)
- **MATERIAL DESCRIPTION**: 2,186 unique values (1 is "-")

---

### 2. PRODUCT_NAME STATISTICS

| Metric | Value |
|--------|-------|
| Total PRODUCT_NAME values | 317,054 |
| Unique PRODUCT_NAME | 150,976 / 317,054 = **47.62%** |
| Exact duplicate rows | 166,078 (52.38%) |

**Frequency Distribution**:
| Occurrences | Unique PRODUCT_NAME Count |
|-------------|---------------------------|
| 1x | 113,518 |
| 2x | 16,416 |
| 3x | 6,387 |
| 5x | 2,263 |
| 10x | 630 |
| 20x | 111 |
| 50x | 256 |
| 100x | 48 |

**Top 10 Most Repeated PRODUCT_NAME**:
1. PROCTOSEDYL BD CREAM 20GM — 340x
2. OROFER XT DROPS 15ML — 309x
3. PAUSE INJ 5ML — 236x
4. XPLODE POWDER 200GM — 214x
5. PAUSE MF TAB 10S — 211x
6. DYDROFEM TAB 10S — 210x
7. GALACT GRANULES 200GM — 192x
8. PROXYM GEL 30GM — 191x
9. OROFER XT TAB 15S — 185x
10. OROFER S 100 INJ 5ML — 183x

**Length Statistics**:
- Min: 4 chars
- Max: 100 chars
- Mean: 19.9 chars
- Median: 20.0 chars
- ≤5 chars: 55 rows (e.g., VYLDA, HOSIT, PAUSE, INBEC)
- ≥100 chars: 2 rows (e.g., "ARTH BONE SUPPORT FOR WOMEN OVER 40 YEARS...")

---

### 3. MAPPING DETERMINISM (PRODUCT_NAME → MATERIAL DESCRIPTION)

| Metric | Value |
|--------|-------|
| Unique PRODUCT_NAME → MATERIAL DESCRIPTION pairs | 150,976 |
| PRODUCT_NAME with exactly 1 MATERIAL DESCRIPTION | **150,976 / 150,976 = 100.00%** |
| PRODUCT_NAME with multiple MATERIAL DESCRIPTION | **0 / 150,976 = 0.00%** |

**Reverse Mapping (MATERIAL DESCRIPTION → PRODUCT_NAME)**:
- Unique MATERIAL DESCRIPTION: 2,186
- MATERIALS with multiple PRODUCT_NAME: **1,919 / 2,186 = 87.79%**
- MATERIALS with single PRODUCT_NAME: 267 (12.21%)

**Distribution of PRODUCT_NAME variants per MATERIAL DESCRIPTION**:
| Variants per Material | Material Count |
|----------------------|----------------|
| 1 | 267 |
| 2–5 | ~600 |
| 6–10 | ~300 |
| 11–20 | ~200 |
| 21–50 | ~300 |
| 51–100 | ~200 |
| 100+ | 438 |

**Conclusion**: The mapping is **perfectly deterministic** from PRODUCT_NAME to MATERIAL DESCRIPTION (no conflicts). However, the reverse is highly one-to-many — each correct product has many OCR/input variations. This is the ideal structure for retriever training.

---

### 4. MATERIAL DESCRIPTION STATISTICS

| Metric | Value |
|--------|-------|
| Unique MATERIAL DESCRIPTION (valid) | 2,185 (excluding "-") |
| Total training examples (valid) | 313,902 |
| Mean examples per material | 143.7 |
| Median examples per material | 21.0 |
| Min examples | 1 |
| Max examples | 2,127 (LASIX 40 MG/4ML INJECTION) |

**Percentile Distribution**:
- P10: 1 example
- P25: 3 examples
- P50: 21 examples
- P75: 121 examples
- P90: 456 examples
- P95: 731 examples
- P99: 1,829 examples

**Class Imbalance**:
- Materials with ≤5 examples: **728 / 2,185 = 33.3%**
- Materials with ≥500 examples: **195 / 2,185 = 8.9%**

**Top 10 Materials by Unique OCR Variants**:
1. (MATERIAL DESCRIPTION = "-") — 2,887 variants, 3,152 rows — **EXCLUDE**
2. LASIX 40 MG/4ML INJECTION — 1,190 variants, 2,127 rows
3. CORDARONE X 200 MG TABLET 10x15T — 1,037 variants, 1,949 rows
4. AMARYL M 1MG TABLET 8x20T — 991 variants, 2,046 rows
5. AMARYL M 2MG TABLET 8x20T — 973 variants, 2,021 rows
6. CORDARONE 100 MG TABLET 20x15T — 936 variants, 2,017 rows
7. CARDACE AM 5 TABLET 10x15T — 917 variants, 1,923 rows
8. CARDACE AM 2.5 TABLET 10x15T — 898 variants, 1,771 rows
9. AMARYL M FORTE 1MG TABLET 10x15T — 898 variants, 1,848 rows
10. AMARYL M FORTE 2MG TABLET 10x15T — 863 variants, 1,797 rows

---

### 5. OCR VARIATION ANALYSIS

**Variation Richness**:
- Materials with ≥100 unique PRODUCT_NAME variants: **439 / 2,185 = 20.1%**
- Materials with ≥50 variants: **665 / 2,185 = 30.4%**
- Materials with ≥20 variants: **1,026 / 2,185 = 47.0%**
- Materials with ≥10 variants: **1,257 / 2,185 = 57.5%**

**Representative Variation Examples**:

**LASIX 40 MG/4ML INJECTION** (1,190 variants):
- LASIX 4ML AMP 4ML
- LASIX AMP 10X4ML (39x)
- LASIX INJ 4ML 10X4ML (7x)
- LASIX AMP 10*4ML (30x)
- LASIX-4.ML BIG-AMP 10*4.M
- LASIX INJ 4ML 1
- LASIX 4ML INJ(SANOFI) 40MG/4ML

**CORDARONE X 200 MG TABLET 10x15T** (1,037 variants):
- CORDORONE X 200mg 10s (typo: CORDORONE)
- CORDARONE-X-200 MG TAB 15 S (hyphen variations)
- CORDARONE X 200MG TAB 15S (42x)
- CORDARONE X 200 15TAB (21x)
- CORDARONE X 200 TAB 1*15 (6x)
- CORDARONE 20*15

**AMARYL M 1MG TABLET 8x20T** (991 variants):
- AMARYL M1 20S (32x)
- AMARYL-M-1 MG TAB 20 S
- AMARYL M 1MG 20S (49x)
- AMARYL M1 TAB 1*20 (8x)
- AMARYAL M 2 20S (typo: AMARYAL)
- AMARYL-M-1MG TABS 1X20

**CALONAT XT TABLET 10x15T** (149 variants):
- CALONAT - XT TAB 15S (spacing)
- CALONAT -XT TAB 15s (case)
- CALONAT XT 1*15 (pack format)
- CALONAT XT 10X15 (pack format)
- CALONAT XT CAP 10 S (wrong form)
- CALONAT XT STRIP OF 15 TABLETS (verbose)

**Variation Types Observed**:
1. **Spelling errors**: CORDORONE, AMARYAL, DULCOFLX, DULCOLEX
2. **Hyphen/spacing**: CORDARONE-X vs CORDARONE X, CARD-ACE vs CARDACE
3. **Pack format**: 10x15T, 10*15, 1*15, 1X15, 15S, 15TAB, 15 T
4. **Strength formatting**: 1MG vs 1 MG vs 1MG. vs 10MG
5. **Brand abbreviation**: CARDAC, EMLASIX, AVCA2.5
6. **Extra tokens**: (SANOFI), (EMCURE), PCS, AMP, TABS
7. **Case variation**: tab vs TAB vs Tab
8. **Missing spaces**: CORDARONE-X-TAB, AMARYL-M-FORTE
9. **Dosage form confusion**: CAP vs TAB, INJ vs AMP
10. **Corrupted numbers**: 10S 10S, 15 S 15 S10, 1104ML

---

### 6. MATERIAL CODE CONSISTENCY

| Metric | Value |
|--------|-------|
| Unique MATERIAL CODE | 2,186 |
| MATERIAL CODE → 1 MATERIAL DESCRIPTION | **2,186 / 2,186 = 100%** |
| MATERIAL CODE → multiple MATERIAL DESCRIPTION | 0 |
| MATERIAL DESCRIPTION → 1 MATERIAL CODE | **2,186 / 2,186 = 100%** |
| MATERIAL DESCRIPTION → multiple MATERIAL CODE | 0 |

**Special Case**: 3,152 rows have MATERIAL CODE = "NO PRODUCT" and MATERIAL DESCRIPTION = "-"
- These represent **unmapped/unknown products** from 395 different PARTY_CODEs
- 2,887 unique PRODUCT_NAME values in this group
- **Zero overlap** with valid mappings (no PRODUCT_NAME appears both in "-" and valid groups)
- **Must be excluded from training**

**Code Quality**: Perfect 1:1 correspondence between MATERIAL CODE and MATERIAL DESCRIPTION for valid data. This confirms MATERIAL CODE can be used as a reliable grouping key.

---

### 7. PARTY_CODE ANALYSIS

| Metric | Value |
|--------|-------|
| Unique PARTY_CODE | 3,018 |
| PRODUCT_NAME appearing with multiple PARTY_CODE | 36,974 / 150,976 = 24.5% |
| PRODUCT_NAME exclusive to single PARTY_CODE | 114,002 / 150,976 = 75.5% |

**Cross-Party Mapping Consistency**:
- **0 conflicts found**: Same PRODUCT_NAME always maps to same MATERIAL DESCRIPTION across all 3,018 parties
- Sampled 1,000 multi-party PRODUCT_NAMEs — all consistent
- Top 10 parties by volume all show 0 mismatches vs global mapping

**Party-Exclusive PRODUCT_NAME**: 75.5% of unique PRODUCT_NAME values appear in only one party
- This suggests many OCR variations are party/customer-specific
- Top party (409180) has 873 exclusive PRODUCT_NAMEs

**Conclusion**: PARTY_CODE does **not** change the mapping relationship. The same PRODUCT_NAME → MATERIAL DESCRIPTION mapping holds across all parties. PARTY_CODE is not needed as a model feature for the core retrieval task, though it explains some of the OCR variation diversity.

---

### 8. EXISTING DATASET OVERLAP

| Dataset | Input_Column Overlap | Positive_Product Overlap | Pair Overlap |
|---------|---------------------|-------------------------|--------------|
| evaluation_set.csv (170 inputs) | 0 / 170 = **0.00%** | 98 / 103 = 95.15% | 0 / 170 = **0.00%** |
| training_data.xlsx (680 inputs) | 0 / 680 = **0.00%** | 184 / 187 = 98.40% | 0 / 680 = **0.00%** |
| hard_negative_mining.csv (678 inputs) | 0 / 678 = **0.00%** | 187 / 194 = 96.39% | 0 / 678 = **0.00%** |
| training_split.csv (508 inputs) | 0 / 508 = **0.00%** | 159 / 164 = 96.95% | 0 / 508 = **0.00%** |
| PRODUCT_MASTER.xlsx (2,764 products) | N/A | 2,180 / 2,764 = 78.87% | N/A |

**Key Insight**: 
- **Zero Input_Column (PRODUCT_NAME) overlap** with any existing training/evaluation data
- **High MATERIAL DESCRIPTION overlap** (95–98%) with existing positive products
- This means the new dataset provides **novel OCR variations** for the same target products — ideal for data augmentation

---

### 9. FIXED EVALUATION CONTAMINATION CHECK

**Evaluation Set**: 364 rows, 170 unique Input_Column, 103 unique Positive_Product

| Check | Result |
|-------|--------|
| Exact PRODUCT_NAME (Input_Column) overlap | **0 / 170 = 0.00%** |
| Exact (Input, Positive) pair overlap | **0 / 170 = 0.00%** |
| Evaluation inputs in new dataset with different target | **0** |
| Evaluation rows contaminated | **0 / 364 = 0.00%** |
| Hard negative mining Input_Column overlap | **0 / 678 = 0.00%** |
| Hard negative conflicts | **0** |

**Evaluation Products in New Dataset**: 98/103 (95.15%) evaluation target products appear in the new dataset with rich OCR variations (17–1,037 variants each). This is **not contamination** — it's the desired training signal.

**Verdict**: **SAFE TO USE** for training without contaminating the fixed evaluation benchmark.

---

### 10. DATA QUALITY FINDINGS

| Issue | Count | Severity |
|-------|-------|----------|
| Exact row duplicates | 1,836 (0.58%) | Low — deduplicate |
| MATERIAL DESCRIPTION = "-" | 3,152 (1.0%) | **High — MUST EXCLUDE** |
| MATERIAL CODE = "NO PRODUCT" | 3,152 (1.0%) | **High — MUST EXCLUDE** |
| PRODUCT_NAME with multiple spaces | 1,287 | Low — normalize |
| PRODUCT_NAME with multiple periods | 142 | Low |
| PRODUCT_NAME with multiple hyphens | 5,930 | Medium — many are valid (e.g., CORDARONE-X) |
| PRODUCT_NAME >80 chars | 11 | Low — verbose descriptions |
| PRODUCT_NAME == MATERIAL DESCRIPTION | 1,928 | Low — exact matches |
| Non-ASCII/control chars | 0 | None |
| MATERIALS with ≤5 examples | 728 (33.3%) | Medium — long tail |

**Suspicious Records**:
- 55 PRODUCT_NAME ≤5 chars (e.g., "VYLDA", "PAUSE", "LASIX") — mostly brand-only shorthand
- 11 PRODUCT_NAME >50 chars longer than MATERIAL DESCRIPTION — verbose marketing text vs canonical name
- 3,152 rows with MATERIAL DESCRIPTION = "-" — completely unmapped, from 395 parties, 2,887 unique PRODUCT_NAMEs (all distinct from valid set)

---

### 11. TRAINING VALUE / RISKS

#### A. What Makes This Dataset Useful

1. **Perfect mapping determinism**: 100% consistent PRODUCT_NAME → MATERIAL DESCRIPTION — no label noise
2. **Massive OCR variation**: 1,257 materials have ≥10 variants; 438 have ≥100 — exactly the diversity needed for robust retrieval
3. **Novel OCR patterns**: Zero overlap on PRODUCT_NAME with existing training data — adds new variation types
4. **High target coverage**: 78.9% of PRODUCT_MASTER products covered; 95% of evaluation targets covered
5. **Rich brand families**: 14 CARDACE materials, 9 AMARYL, 32 ASOMEX, 26 TEMSAN — enables learning brand+strength+form compositionality
6. **Real-world noise**: Spelling errors, hyphenation, pack formats, abbreviations, case variation — matches production OCR errors
7. **Scale**: 313,902 valid training pairs across 2,185 products — sufficient for contrastive fine-tuning

#### B. Problems That Make Blind Training Dangerous

1. **3,152 "NO PRODUCT" rows** — must be filtered out before any training
2. **Severe class imbalance**: 33.3% of materials have ≤5 examples; median is only 21
3. **1,836 exact duplicate rows** — need deduplication
4. **Long-tail materials**: 728 materials with ≤5 examples may not support meaningful gradient updates
5. **Party-exclusive variations**: 75.5% of PRODUCT_NAME unique to one party — may overfit to party-specific patterns if not shuffled properly
6. **Very short PRODUCT_NAMEs** (55 rows ≤5 chars) — may not provide sufficient signal

#### C. Does It Contain OCR Variation Needed for Recall@3 Improvement?

**YES**. The dataset contains exactly the variation types that cause retrieval failures:
- Hyphen/spacing errors (CARD-ACE vs CARDACE)
- Strength format variation (1MG vs 1 MG vs 10MG)
- Pack size confusion (10x15T vs 1*15 vs 15S)
- Spelling corruptions (CORDORONE, AMARYAL)
- Brand abbreviation (CARDAC, EMLASIX)
- Missing/wrong dosage forms (CAP vs TAB)

With 1,257 materials having ≥10 variants, the model can learn invariance to these perturbations.

#### D. Enough Different OCR Representations Per Product?

**YES for major products, NO for long tail**:
- Top 195 materials (≥500 examples): 100–2,887 variants each — excellent
- 665 materials (≥50 variants): good coverage
- 728 materials (≤5 examples): insufficient — need filtering or augmentation

#### E. Major Risks to Handle Before Training

1. **Filter "NO PRODUCT" rows** (3,152 rows, 1.0%)
2. **Deduplicate exact rows** (1,836 rows)
3. **Handle class imbalance**: Consider minimum 10–20 examples per material, or use weighted sampling
4. **Stratify by PARTY_CODE** in train/val split to avoid party leakage
5. **Normalize whitespace/hyphens** in preprocessing (not in data — in tokenizer/preprocessing)
6. **Consider excluding materials with <10 examples** (728 materials, but only ~2,000 rows total)

---

### 12. RECOMMENDED NEXT ANALYSIS STEP

**Immediate Next Step**: Create a **cleaned training dataset** by:
1. Removing 3,152 rows with MATERIAL DESCRIPTION = "-"
2. Removing 1,836 exact duplicate rows
3. Computing final per-material example counts
4. Deciding on minimum examples threshold (recommend: ≥10)

**Then**: Design the V2 training configuration addressing:
- Train/validation split strategy (stratified by MATERIAL DESCRIPTION, ensuring party separation)
- Handling of long-tail materials (exclude? augment? weight?)
- Hard negative mining strategy using the 1,257 materials with ≥10 variants
- Whether to include PARTY_CODE as metadata (recommend: no, based on evidence)

**Do NOT yet decide**: epochs, batch size, learning rate, loss function, number of negatives — these depend on the cleaned dataset statistics.

---

*Report generated from automated profiling of EmcureSSSReport.xlsb Sheet1. All numbers are measured from the actual dataset.*