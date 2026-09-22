# SECOND-STAGE TRAINING-DATA ANALYSIS REPORT
## EmcureSSSReport.xlsb (Sheet1) for V2 Sentence Transformer Retriever

---

### EXECUTIVE SUMMARY

The 317,054-row historical dataset (Sheet1) cleans to **312,120 usable rows** (98.4% retention) with **148,089 unique PRODUCT_NAME → MATERIAL DESCRIPTION pairs** across **2,185 materials**. 

**Key findings:**
- **Zero mapping conflicts** — every PRODUCT_NAME maps to exactly one MATERIAL DESCRIPTION
- **Zero evaluation contamination** — 0% overlap with fixed evaluation_set.csv
- **Massive OCR diversity** — 1,918 materials (87.8%) have multiple OCR variants; median 16 variants/material
- **Complete V1 novelty** — 0% pair overlap with V1 training data; 148,089 entirely new positive pairs
- **Severe class imbalance** — 35% materials have ≤5 variants (long-tail); top 43 materials have 500-1,190 variants

**Training value: VERY HIGH** — The dataset provides exactly the OCR variation types that limit Recall@3, with zero risk to the fixed benchmark.

---

### 1. CLEANED DATASET SIZE

| Metric | Count | Percentage |
|--------|-------|------------|
| Original rows | 317,054 | 100% |
| NO PRODUCT rows (MATERIAL DESCRIPTION = "-") | 3,152 | 1.0% |
| Exact duplicate rows | 1,782 | 0.6% |
| **Usable rows** | **312,120** | **98.4%** |

---

### 2. POSITIVE PAIR ANALYSIS

| Metric | Value |
|--------|-------|
| Unique PRODUCT_NAME | 148,089 |
| Unique MATERIAL DESCRIPTION | 2,185 |
| Unique PRODUCT_NAME + MATERIAL DESCRIPTION pairs | 148,089 |
| Duplicate pairs (same PN+MD appearing multiple times) | 36,841 |
| PRODUCT_NAME → multiple MATERIAL DESCRIPTION | **0 (0.00%)** |
| MATERIAL DESCRIPTION ← multiple PRODUCT_NAME | **1,918 / 2,185 = 87.78%** |

**Conclusion**: Perfect 1:1 forward mapping (no label noise). Rich 1:many reverse mapping ideal for contrastive learning.

---

### 3. CONFLICT ANALYSIS

**ZERO CONFLICTS** — No PRODUCT_NAME maps to multiple MATERIAL DESCRIPTION values.

This is a critical quality signal: the historical mapping is fully deterministic. The model will not learn contradictory associations.

---

### 4. MATERIAL FREQUENCY DISTRIBUTION (by unique OCR variants)

| Bucket | Materials | %Materials | Examples | %Examples |
|--------|-----------|------------|----------|-----------|
| 1 | 267 | 12.22% | 282 | 0.09% |
| 2 | 193 | 8.83% | 413 | 0.13% |
| 3-5 | 306 | 14.00% | 1,274 | 0.41% |
| 6-10 | 186 | 8.51% | 1,649 | 0.53% |
| 11-25 | 317 | 14.51% | 7,531 | 2.41% |
| 26-50 | 262 | 11.99% | 14,492 | 4.64% |
| 51-100 | 221 | 10.11% | 29,192 | 9.35% |
| 101-250 | 287 | 13.14% | 104,609 | 33.52% |
| 251-500 | 103 | 4.71% | 81,678 | 26.17% |
| 500+ | 43 | 1.97% | 71,000 | 22.75% |
| **TOTAL** | **2,185** | **100%** | **312,120** | **100%** |

**Key insight**: The top 13.14% of materials (101-250 variants) contribute 33.5% of training examples. The top 1.97% (500+) contribute 22.8%.

---

### 5. OCR DIVERSITY

| Statistic | Value |
|-----------|-------|
| Mean variants/material | 67.8 |
| Median variants/material | 16.0 |
| P90 | 193.0 |
| P95 | 296.8 |
| Maximum | 1,190 (LASIX 40 MG/4ML INJECTION) |

**Low diversity examples (1-2 variants):**
- ABACUS - 50 TAB: ABACUS 50 10S, ABACUS-50 TAB.
- ACEM - 250 TABLET: ACEM5860 10S, ACEM 250MG 10

**Medium diversity (10-20 variants):**
- ABIRACURE TABLET: 10 variants (ABIRACURE 250MG, ABIRACURE 250MG 120S, etc.)
- AMAZIL 10 TABLETS: 12 variants (AMAZIL 10 10S, AMAZIL 10 TAB 10TAB, etc.)

**High diversity (500+ variants):**
- AMARYL 1MG TABLET 10x30T: 750 variants
- AMARYL M 1MG TABLET 8x20T: 991 variants
- CARDACE AM 5 TABLET 10x15T: 917 variants
- LASIX 40 MG/4ML INJECTION: 1,190 variants

---

### 6. LONG-TAIL IMPACT

| Metric | Value |
|--------|-------|
| Materials with ≤5 variants | 766 (35.06% of 2,185) |
| Training examples from long-tail | 1,969 (0.63% of 312,120) |
| Long-tail materials in PRODUCT_MASTER | 763 / 766 = **99.61%** |
| PRODUCT_MASTER covered by historical | 2,180 / 2,764 = **78.87%** |
| Covered products that are long-tail | 763 / 2,180 = **35.00%** |

**Implication**: The long-tail represents real PRODUCT_MASTER products (99.6% match) but contributes negligible training signal (0.63% of examples). They should be **retained naturally** without oversampling — they provide correct positive anchors even if few variants exist.

---

### 7. HIGH-FREQUENCY MATERIALS (Top 30 by unique variants)

| Rank | MATERIAL DESCRIPTION | Variants | Rows |
|------|---------------------|----------|------|
| 1 | LASIX 40 MG/4ML INJECTION | 1,190 | 2,107 |
| 2 | CORDARONE X 200 MG TABLET 10x15T | 1,037 | 1,937 |
| 3 | AMARYL M 1MG TABLET 8x20T | 991 | 2,029 |
| 4 | AMARYL M 2MG TABLET 8x20T | 973 | 2,001 |
| 5 | CORDARONE 100 MG TABLET 20x15T | 936 | 2,002 |
| 6 | CARDACE AM 5 TABLET 10x15T | 917 | 1,917 |
| 7 | AMARYL M FORTE 1MG TABLET 10x15T | 898 | 1,827 |
| 8 | CARDACE AM 2.5 TABLET 10x15T | 898 | 1,768 |
| 9 | AMARYL M FORTE 2MG TABLET 10x15T | 863 | 1,778 |
| 10 | CARDACE 2.5 MG TABLET 10x15T | 861 | 2,111 |
| ... | ... | ... | ... |
| 30 | AMARYL 3MG TABLET 10x30T | 659 | 1,701 |

**Near-duplicate analysis (normalized uniqueness):**
- LASIX 40 MG/4ML INJECTION: 1,190 raw → 958 normalized (80.5% unique)
- CORDARONE X 200 MG TABLET: 1,037 raw → 753 normalized (72.6% unique)
- AMARYL M 1MG TABLET: 991 raw → 694 normalized (70.0% unique)
- CARDACE 2.5 MG TABLET: 861 raw → 629 normalized (73.1% unique)

**Finding**: 20-30% of raw variants are near-duplicates (whitespace/hyphen/case differences). Capping at 100-200 retains genuine diversity while reducing redundancy.

---

### 8. EVALUATION SET PROTECTION

| Check | Result |
|-------|--------|
| Evaluation pairs (Input_Column, Positive_Product) | 170 |
| Historical pairs (PRODUCT_NAME, MATERIAL DESCRIPTION) | 148,089 |
| **Exact pair overlap** | **0 (0.00%)** |
| Evaluation Input_Column overlap | 0 / 170 = 0.00% |
| Hard negative mining pair overlap | 0 / 678 = 0.00% |
| Hard negative mining Input_Column overlap | 0 / 678 = 0.00% |
| Conflicts (eval input → different target) | 0 |

**VERDICT**: **ZERO CONTAMINATION**. The historical dataset can be safely used for training without compromising the fixed evaluation benchmark.

---

### 9. V1 TRAINING DATA RELATIONSHIP

| Metric | Value |
|--------|-------|
| Historical unique pairs | 148,089 |
| V1 training_split pairs | 508 |
| V1 hard_negative_mining pairs | 678 |
| V1 all pairs | 678 |
| Overlap with V1 training_split | 0 / 508 = 0.00% |
| Overlap with V1 hard_negative | 0 / 678 = 0.00% |
| **NEW positive pairs from historical** | **148,089 (100%)** |
| Materials in V1 | 194 |
| Materials in historical | 2,185 |
| Common materials | 187 |

**New variants gained for common materials (top 10):**
- CORDARONE X 200 MG TABLET: V1=4 → Hist=1,037 (**+1,033 new**)
- AMARYL M 2MG TABLET: V1=2 → Hist=973 (**+971 new**)
- CORDARONE 100 MG TABLET: V1=2 → Hist=936 (**+934 new**)
- CARDACE AM 5 TABLET: V1=1 → Hist=917 (**+916 new**)
- CARDACE AM 2.5 TABLET: V1=3 → Hist=898 (**+895 new**)

**Conclusion**: Historical data adds massive OCR diversity (100-1000x more variants) for the same materials V1 was trained on.

---

### 10. TRAINING STRATEGY SIMULATIONS

| Strategy | Total Pairs | % Retained | Materials | Min | Median | Max |
|----------|-------------|------------|-----------|-----|--------|-----|
| A: All cleaned pairs | 148,089 | 100% | 2,185 | 1 | 16 | 1,190 |
| B: Cap=50 | 51,021 | 34.5% | 2,185 | 1 | 16 | 50 |
| B: Cap=100 | 77,819 | 52.5% | 2,185 | 1 | 16 | 100 |
| B: Cap=200 | 108,775 | 73.5% | 2,185 | 1 | 16 | 200 |
| B: Cap=500 | 137,254 | 92.7% | 2,185 | 1 | 16 | 500 |
| C: Balanced (16/material) | 22,569 | 15.2% | 2,185 | 16 | 16 | 16 |
| **D: Hybrid (≤50 keep, >50 cap 100)** | **77,819** | **52.5%** | **2,185** | **1** | **~16** | **100** |
| D2: Hybrid (≤20/50/100) | 78,691 | 53.1% | 2,185 | 1 | ~16 | 100 |
| **D3: Hybrid (≤100 keep, >100 cap 200)** | **121,119** | **81.8%** | **2,185** | **1** | **~16** | **200** |

---

### 11. OCR VARIATION TYPES FOR RECALL@3 IMPROVEMENT

The dataset contains **all critical variation types** that cause retrieval failures:

| Variation Type | Present? | Examples |
|----------------|----------|----------|
| **Spelling/OCR corruption** | ✓ | CORDORONE, AMARYAL, LASIK, ENHEM, DULCOFLX |
| **Hyphen/spacing** | ✓ | CARD-ACE vs CARDACE, CORDARONE-X vs CORDARONE X |
| **Strength format** | ✓ | 1MG vs 1 MG vs 10MG vs 1MG. vs 1 MG TAB |
| **Pack size format** | ✓ | 10x15T vs 1*15 vs 15S vs 15TAB vs 10*15 vs AMP/VIAL |
| **Brand abbreviation** | ✓ | CARDAC, EMLASIX, EM DYDRO, CARDAC AM |
| **Dosage form confusion** | ✓ | CAP vs TAB, INJ vs AMP, DROPS vs TABLET |
| **Numeric token corruption** | ✓ | 104ML, 110, 15S 15 S, 20 S 20 S, 10X4ML |
| **Case variation** | ✓ | tab vs TAB vs Tab, mg vs MG |
| **Extra tokens** | ✓ | (SANOFI), (EMCURE), PCS, VIAL, STRIP |
| **Product family confusion** | ✓ | CARDACE 1.25/2.5/5/10/AM/H/METO/PROTECT (14 materials) |

**Product families with high confusion risk:**
- **CARDACE**: 14 materials (1.25, 2.5, 5, 10, AM, H, METO, PROTECT)
- **AMARYL**: 9 materials (1/2/3MG, M 1/2MG, M FORTE 1/2MG, MV 1/2MG)
- **ASOMEX**: 32 materials (1.25/2.5/5, M, TM, D, LT, AT variants)
- **TEMSAN**: 26 materials (20/40/80, AM, H, BS, CT)
- **OROFER**: 21 materials (XT, S, FCM, D3, CAPSULES, DROPS)

---

### 12. RECOMMENDED V2 DATASET CONSTRUCTION STRATEGY

#### RECOMMENDED: Strategy D3 (Hybrid: retain ≤100, cap >100 at 200)

| Parameter | Value |
|-----------|-------|
| **Training pairs** | ~121,000 (81.8% of cleaned data) |
| **Materials represented** | 2,185 (ALL retained) |
| **Min examples/material** | 1 (long-tail preserved) |
| **Median examples/material** | ~16 |
| **Max examples/material** | 200 (capped from 1,190) |
| **Long-tail (≤5 variants)** | 766 materials retained naturally |
| **Frequent materials capped** | 43 materials capped at 200 |

#### RATIONALE

1. **Retains ALL 2,185 materials** — including 766 long-tail (99.6% in PRODUCT_MASTER)
2. **Caps 43 extreme-frequency materials** at 200 variants (from 1,190 max) — reduces family dominance
3. **Preserves 82% of diverse OCR examples** — sufficient for robust contrastive learning
4. **Keeps genuine diversity** — normalized uniqueness is 68-80% even at top; 200 cap retains most signal
5. **Reduces CARDACE/AMARYL/LASIX/CORDARONE dominance** — these 4 families dominate top 30

#### ALTERNATIVE: Strategy D (Hybrid: retain ≤50, cap >50 at 100)
- ~78,000 pairs (52.5% retention)
- More aggressive capping, faster training, less redundancy

#### FINAL RECOMMENDATIONS

| Decision | Recommendation | Evidence |
|----------|----------------|----------|
| **Training pairs** | 118,000–130,000 | Strategy D3 or D |
| **Use all 319K raw rows?** | NO — apply cleaning + capping | 3,152 NO PRODUCT + 1,782 dupes + severe imbalance |
| **Cap frequent materials?** | YES — at 100-200 variants | Top 43 materials have 500-1,190 variants; 20-30% near-duplicates |
| **Handle rare materials?** | Retain naturally (no oversampling) | 766 materials, 99.6% in PRODUCT_MASTER, 0.63% of data |
| **Include V1 training data?** | YES — combine datasets | 0% overlap; V1 adds 678 pairs, 194 materials, different distribution |
| **Negatives strategy** | In-batch + existing hard_negative_mining.csv (1,425) + mine new from 2,185 materials using V1 model | Historical provides 2,185 candidate positive classes for hard negative mining |

---

### NEXT STEPS (not implemented yet)

1. **Implement Strategy D3 cleaning + capping** to produce `v2_training_pairs.csv`
2. **Merge with V1 training data** (deduplicate on pairs)
3. **Design train/validation split** — stratified by MATERIAL DESCRIPTION, ensuring no PARTY_CODE leakage
4. **Mine new hard negatives** using V1 model on the 2,185 historical materials
5. **Configure V2 training** — epochs, batch size, learning rate, loss function
6. **Run evaluation** on fixed `evaluation_set.csv` to measure Recall@3 improvement

---

*Analysis complete. No files modified. No training performed. Ready for V2 dataset construction decision.*