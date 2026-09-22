import pandas as pd
import numpy as np

# Load data
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
train_split = pd.read_csv('training_split.csv')
hard_neg = pd.read_csv('hard_negative_mining.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

# Load constructed V2 dataset
v2_df = pd.read_csv('v2_positive_pairs.csv')
print(f"Loaded V2 positive pairs: {len(v2_df):,}")
print(f"Materials: {v2_df['MATERIAL DESCRIPTION'].nunique():,}")

print("=" * 70)
print("TASK 3: VERIFY MATERIAL COVERAGE")
print("=" * 70)

# Historical materials
hist_materials = set(v2_df['MATERIAL DESCRIPTION'].unique())
print(f"V2 materials: {len(hist_materials):,}")

# PRODUCT_MASTER products
master_products = set(product_master['product_name'].unique())
print(f"PRODUCT_MASTER products: {len(master_products):,}")

# Coverage
covered = hist_materials & master_products
missing_from_master = hist_materials - master_products
master_missing = master_products - hist_materials

print(f"Materials in PRODUCT_MASTER: {len(covered):,} / {len(hist_materials):,} = {len(covered)/len(hist_materials)*100:.2f}%")
print(f"Materials NOT in PRODUCT_MASTER: {len(missing_from_master):,}")
print(f"PRODUCT_MASTER products NOT covered: {len(master_missing):,} / {len(master_products):,} = {len(master_missing)/len(master_products)*100:.2f}%")

if len(missing_from_master) > 0:
    print(f"\nMaterials in V2 but not in PRODUCT_MASTER:")
    for m in sorted(missing_from_master):
        print(f"  {m}")

# Examples per material stats
md_counts = v2_df.groupby('MATERIAL DESCRIPTION').size()
print(f"\nExamples per material:")
print(f"  Min: {md_counts.min()}")
print(f"  Median: {md_counts.median():.1f}")
print(f"  Mean: {md_counts.mean():.1f}")
print(f"  P90: {md_counts.quantile(0.90):.1f}")
print(f"  P95: {md_counts.quantile(0.95):.1f}")
print(f"  Max: {md_counts.max()}")

# Long-tail check
long_tail = md_counts[md_counts <= 5]
print(f"\nMaterials with <=5 examples: {len(long_tail):,} ({len(long_tail)/len(md_counts)*100:.2f}%)")
long_tail_in_master = [m for m in long_tail.index if m in master_products]
print(f"  In PRODUCT_MASTER: {len(long_tail_in_master):,} ({len(long_tail_in_master)/len(long_tail)*100:.2f}%)")

# Capped materials
capped = md_counts[md_counts == 200]
print(f"\nMaterials at cap (200): {len(capped):,}")
for m in sorted(capped.index)[:10]:
    print(f"  {m}")

print("=" * 70)
print("TASK 4: EVALUATION PROTECTION")
print("=" * 70)

eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
v2_pairs = set(zip(v2_df['PRODUCT_NAME'], v2_df['MATERIAL DESCRIPTION']))

# Exact pair overlap
pair_overlap = eval_pairs & v2_pairs
print(f"Evaluation pairs: {len(eval_pairs):,}")
print(f"V2 pairs: {len(v2_pairs):,}")
print(f"Exact pair overlap: {len(pair_overlap):,}")

# Input_Column overlap
eval_inputs = set(eval_set['Input_Column'].unique())
v2_inputs = set(v2_df['PRODUCT_NAME'].unique())
input_overlap = eval_inputs & v2_inputs
print(f"Evaluation Input_Column overlap: {len(input_overlap):,} / {len(eval_inputs):,} = {len(input_overlap)/len(eval_inputs)*100:.2f}%")

# Check conflicts: eval input in V2 with different target
conflicts = 0
for inp in input_overlap:
    eval_target = eval_set[eval_set['Input_Column'] == inp]['Positive_Product'].values[0]
    v2_targets = set(v2_df[v2_df['PRODUCT_NAME'] == inp]['MATERIAL DESCRIPTION'].unique())
    if eval_target not in v2_targets:
        conflicts += 1
        print(f"  CONFLICT: {inp} -> EVAL:{eval_target} vs V2:{v2_targets}")
print(f"Conflicts (eval input -> different target): {conflicts}")

# Also check: V2 pairs where positive product is in evaluation but input is different
# This is not contamination but worth noting
eval_products = set(eval_set['Positive_Product'].unique())
v2_products = set(v2_df['MATERIAL DESCRIPTION'].unique())
product_overlap = eval_products & v2_products
print(f"\nEvaluation Positive_Product in V2: {len(product_overlap):,} / {len(eval_products):,} = {len(product_overlap)/len(eval_products)*100:.2f}%")
print("(This is expected and NOT contamination - V2 provides OCR variants for eval targets)")

print("=" * 70)
print("TASK 5: HARD NEGATIVE MINING DESIGN")
print("=" * 70)

print("""
HARD NEGATIVE MINING DESIGN FOR V2
===================================

Objective: Mine hard negatives for ~109K positive pairs that are 
pharmaceutically relevant and similar to the positive.

CANDIDATE POOL:
- All 2,764 PRODUCT_MASTER products (canonical names)
- 2,192 V2 MATERIALS (which map to PRODUCT_MASTER)
- Use the V1 fine-tuned model to score candidates

MINING PROCEDURE PER POSITIVE PAIR:
1. Anchor: PRODUCT_NAME (OCR input)
2. Positive: MATERIAL DESCRIPTION (canonical)
3. Candidate negatives: All PRODUCT_MASTER products EXCEPT the positive
4. Score candidates using V1 model (embed anchor, embed candidates, cosine similarity)
5. Select top-K most similar (but incorrect) products as hard negatives

NEGATIVE PRIORITIZATION RULES:
1. Same brand, different strength/form/pack (highest priority)
   - e.g., CARDACE 2.5 MG vs CARDACE 5 MG vs CARDACE AM 5
2. Same product family, different variant
   - e.g., AMARYL M 1MG vs AMARYL M 2MG vs AMARYL M FORTE
3. OCR-confusable siblings
   - Products with similar spelling/token patterns
4. Similar numeric formulation
   - Same active ingredient, different strength
5. Same therapeutic class (lower priority)
   - Different brand, similar indication

FILTERING RULES:
- Exclude the positive MATERIAL DESCRIPTION
- Exclude any MATERIAL DESCRIPTION that maps to same PRODUCT_CODE (synonyms)
- Exclude obvious non-products
- Ensure negative != positive

IMPLEMENTATION:
- Batch encode all PRODUCT_MASTER products once (2,764 embeddings)
- For each anchor, compute similarities to all candidates
- Select top 3-5 hard negatives per positive
- Store as (anchor, positive, hard_neg_1, hard_neg_2, hard_neg_3)

CONSTRAINTS FROM PRODUCTION CANDIDATE POOL:
- The production pipeline uses a candidate pool of valid products
- Mining should respect the same product universe
- Use PRODUCT_MASTER as the authoritative product list
- V1 model embeddings are 768-dim (BGE-base)
""")

# Analyze brand families for negative mining
print("\nBRAND FAMILIES FOR NEGATIVE MINING (from PRODUCT_MASTER):")
master_brands = product_master['BRAND_NAME'].dropna().unique()
print(f"Unique brands in PRODUCT_MASTER: {len(master_brands)}")

# Products per brand
brand_products = product_master.groupby('BRAND_NAME')['product_name'].apply(list)
large_families = {b: p for b, p in brand_products.items() if len(p) >= 3}
print(f"Brands with >=3 products: {len(large_families)}")
for b, products in sorted(large_families.items(), key=lambda x: -len(x[1]))[:15]:
    print(f"  {b}: {len(products)} products - {products[:5]}...")

print("=" * 70)
print("TASK 6: MINING SCALE ESTIMATION")
print("=" * 70)

n_positives = len(v2_df)
print(f"Positive pairs: {n_positives:,}")

# Storage per negative
# Each negative: (anchor_idx, positive_idx, neg_text, neg_score) ~ 100 bytes
# Or in training format: Input_Column, Positive_Product, Hard_Neg_1, Hard_Neg_2, Hard_Neg_3

print(f"\nStorage estimates:")
for k in [1, 2, 3, 4, 5]:
    total_negs = n_positives * k
    print(f"  {k} negatives/positive: {total_negs:,} total negatives")

print(f"\nCPU mining time estimates (V1 model, batch size 64):")
print(f"  Embed all 2,764 PRODUCT_MASTER products: ~30 seconds")
print(f"  For each positive, score against 2,763 candidates:")
print(f"    109K * 2,763 = ~300M similarity computations")
print(f"    At ~10K/sec (GPU): ~8 hours")
print(f"    At ~1K/sec (CPU): ~83 hours")

print(f"\nRecommendation:")
print(f"  - Mine 3 hard negatives per positive")
print(f"  - Total negatives: ~328K")
print(f"  - Use GPU for reasonable time")
print(f"  - Can subset to high-value positives (e.g., materials with >10 variants)")
print(f"  - Or use approximate nearest neighbor (FAISS) for speed")

print("=" * 70)
print("TASK 7: V2 TRAINING DATA VALIDATION PLAN")
print("=" * 70)

print("""
VALIDATION CHECKLIST FOR V2 DATASET
====================================

PRE-TRAINING CHECKS:
--------------------
[ ] No invalid products in positives (all MATERIAL DESCRIPTION in PRODUCT_MASTER or documented exception)
[ ] No evaluation leakage (0 pair overlap, 0 input overlap with conflicts)
[ ] No duplicate positive pairs
[ ] No contradictory mappings (PRODUCT_NAME -> single MATERIAL DESCRIPTION)
[ ] All positives exist in PRODUCT_MASTER (or documented reason)
[ ] All negatives differ from positives
[ ] No accidental positive-as-negative in hard negative mining
[ ] Material coverage: all 2,185+ historical materials represented
[ ] Long-tail coverage: 766 materials with <=5 examples retained
[ ] Hard negative quality: sampled from same brand family where possible
[ ] Class distribution: min=1, median=16, max=200 (capped)

POST-MINING CHECKS:
-------------------
[ ] Each positive has K hard negatives
[ ] No hard negative equals its positive
[ ] No hard negative equals another positive for same anchor
[ ] Hard negatives are valid PRODUCT_MASTER products
[ ] Brand-family negatives prioritized for same-brand materials
[ ] No data leakage from evaluation set into negatives

FORMAT CHECKS:
--------------
[ ] Consistent column names: Input_Column, Positive_Product, Hard_Neg_1, Hard_Neg_2, Hard_Neg_3
[ ] No null/empty values
[ ] Correct data types
[ ] Compatible with existing training pipeline

EVALUATION CHECKS:
------------------
[ ] evaluation_set.csv completely untouched
[ ] No V2 training pair appears in evaluation
[ ] No V2 hard negative mining uses evaluation anchors
""")

# Run actual validation checks
print("=" * 70)
print("RUNNING VALIDATION CHECKS")
print("=" * 70)

# Check 1: No duplicate pairs
dup_pairs = v2_df.duplicated(subset=['PRODUCT_NAME', 'MATERIAL DESCRIPTION']).sum()
print(f"1. Duplicate positive pairs: {dup_pairs}")

# Check 2: Contradictory mappings
pn_to_md = v2_df.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
conflicts = (pn_to_md > 1).sum()
print(f"2. PRODUCT_NAME -> multiple MATERIAL DESCRIPTION: {conflicts}")

# Check 3: All materials in PRODUCT_MASTER or documented
v2_materials = set(v2_df['MATERIAL DESCRIPTION'].unique())
master_products = set(product_master['product_name'].unique())
not_in_master = v2_materials - master_products
print(f"3. Materials NOT in PRODUCT_MASTER: {len(not_in_master):,}")
if len(not_in_master) <= 20:
    for m in sorted(not_in_master):
        print(f"   {m}")

# Check 4: Evaluation protection (already done)
print(f"4. Evaluation pair overlap: {len(pair_overlap):,}")
print(f"   Evaluation input overlap: {len(input_overlap):,}")
print(f"   Conflicts: {conflicts}")

# Check 5: V1 pairs in evaluation (need to filter)
v1_all_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product'])) | \
               set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
v1_eval_overlap = v1_all_pairs & eval_pairs
print(f"5. V1 pairs in evaluation: {len(v1_eval_overlap):,} (MUST FILTER OUT)")

# Check 6: Long-tail retention
long_tail_v2 = md_counts[md_counts <= 5]
print(f"6. Long-tail materials retained: {len(long_tail_v2):,}")

# Check 7: Capped materials
capped_v2 = md_counts[md_counts == 200]
print(f"7. Materials at cap (200): {len(capped_v2):,}")

# Check 8: Unique PRODUCT_NAME
print(f"8. Unique PRODUCT_NAME: {v2_df['PRODUCT_NAME'].nunique():,}")
print(f"   Unique MATERIAL DESCRIPTION: {v2_df['MATERIAL DESCRIPTION'].nunique():,}")

# Save final validated dataset (with V1 eval pairs removed)
v2_clean_pairs = v2_pairs - v1_eval_overlap
print(f"\nAfter removing V1 eval overlaps: {len(v2_clean_pairs):,} pairs")

# Convert back to dataframe
v2_clean_df = pd.DataFrame(list(v2_clean_pairs), columns=['PRODUCT_NAME', 'MATERIAL DESCRIPTION'])
v2_clean_df.to_csv('v2_positive_pairs_validated.csv', index=False)
print(f"Saved v2_positive_pairs_validated.csv")