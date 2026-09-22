import pandas as pd
import numpy as np

# Load all data
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
train_split = pd.read_csv('training_split.csv')
hard_neg = pd.read_csv('hard_negative_mining.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 70)
print("RECONSTRUCTING V2 POSITIVE DATASET WITH EXCLUSIONS & FIXED CAPPING")
print("=" * 70)

# ============================================================
# STEP 1: CLEAN HISTORICAL DATA
# ============================================================
original_rows = len(df)
no_product_mask = (df['MATERIAL DESCRIPTION'] == '-') | (df['MATERIAL CODE'] == 'NO PRODUCT')
df_clean = df[~no_product_mask].copy()
df_clean = df_clean.drop_duplicates().reset_index(drop=True)
print(f"Clean historical rows: {len(df_clean):,}")

# ============================================================
# STEP 2: EXCLUDE 9 MATERIALS NOT IN PRODUCT_MASTER
# ============================================================
excluded_materials = {
    'ATAZOR 300 MG CAPSULE 30\'S',
    'EMCLOZ 0.25MG TABLET',
    'EMCLOZ 0.5MG TABLET',
    'ENCICARB INJECTION 1K',
    'LOMOREST-30 SACHET 1X3\'S',
    'MOLENA 200 MG CAPSULE 40\'S',
    'NEVIR 200 MG TABLET 60\'S',
    'VONADAY 600/300/300 MG TABLET 30s',
    'ZUVISTON TABLETS'
}

before_exclude = len(df_clean)
df_clean = df_clean[~df_clean['MATERIAL DESCRIPTION'].isin(excluded_materials)].copy()
after_exclude = len(df_clean)
print(f"Excluded {len(excluded_materials)} materials: removed {before_exclude - after_exclude:,} rows")

# Verify exclusion
remaining_excluded = df_clean[df_clean['MATERIAL DESCRIPTION'].isin(excluded_materials)]
print(f"Remaining excluded materials: {len(remaining_excluded)} (should be 0)")

# ============================================================
# STEP 3: GET UNIQUE POSITIVE PAIRS
# ============================================================
pair_df = df_clean[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates().reset_index(drop=True)
print(f"Unique positive pairs after exclusion: {len(pair_df):,}")

# ============================================================
# STEP 4: APPLY D3 HYBRID CAPPING WITH PROPER 200 MAX
# ============================================================
# Group by MATERIAL DESCRIPTION, get unique PRODUCT_NAME variants
md_variants = pair_df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].apply(set).to_dict()

print(f"\nApplying D3 Hybrid capping:")
print(f"  Retain <=100 variants: keep all")
print(f"  Retain 101-200 variants: keep all")
print(f"  Cap >200 variants at 200")

np.random.seed(42)

final_pairs = []
capped_count = 0
retained_101_200 = 0
retained_le_100 = 0

for md, variants in md_variants.items():
    n = len(variants)
    if n > 200:
        # Sample exactly 200
        selected = np.random.choice(list(variants), size=200, replace=False)
        capped_count += 1
    elif n > 100:
        # Keep all 101-200
        selected = list(variants)
        retained_101_200 += 1
    else:
        # Keep all <=100
        selected = list(variants)
        retained_le_100 += 1
    
    for pn in selected:
        final_pairs.append({'PRODUCT_NAME': pn, 'MATERIAL DESCRIPTION': md})

v2_hist_df = pd.DataFrame(final_pairs)
print(f"\nHistorical pairs after capping: {len(v2_hist_df):,}")
print(f"  Materials capped (>200): {capped_count}")
print(f"  Materials retained 101-200: {retained_101_200}")
print(f"  Materials retained <=100: {retained_le_100}")
print(f"  Total materials: {capped_count + retained_101_200 + retained_le_100}")

# Verify max is exactly 200
v2_variant_counts = v2_hist_df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
print(f"\nVariant stats after capping:")
print(f"  Min: {v2_variant_counts.min()}")
print(f"  Median: {v2_variant_counts.median():.1f}")
print(f"  Mean: {v2_variant_counts.mean():.1f}")
print(f"  P90: {v2_variant_counts.quantile(0.90):.1f}")
print(f"  P95: {v2_variant_counts.quantile(0.95):.1f}")
print(f"  Max: {v2_variant_counts.max()}")
print(f"  Materials at 200: {len(v2_variant_counts[v2_variant_counts == 200])}")

# ============================================================
# STEP 5: INCLUDE V1 POSITIVES
# ============================================================
v1_train_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product']))
v1_hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
v1_all_pairs = v1_train_pairs | v1_hn_pairs

hist_pairs = set(zip(v2_hist_df['PRODUCT_NAME'], v2_hist_df['MATERIAL DESCRIPTION']))

# V1 pairs not in historical
v1_new = v1_all_pairs - hist_pairs
print(f"\nV1 all pairs: {len(v1_all_pairs)}")
print(f"V1 new pairs (not in historical): {len(v1_new)}")

# Check V1 against evaluation
eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
v1_eval_overlap = v1_all_pairs & eval_pairs
print(f"V1 pairs overlapping with evaluation: {len(v1_eval_overlap)} (will remove)")

# Remove V1 eval overlaps from v1_new
v1_new_clean = v1_new - eval_pairs
print(f"V1 pairs to add after eval removal: {len(v1_new_clean)}")

# Combine
all_pairs = hist_pairs | v1_new_clean
v2_final_df = pd.DataFrame(list(all_pairs), columns=['PRODUCT_NAME', 'MATERIAL DESCRIPTION'])
print(f"\nFinal V2 positive pairs: {len(v2_final_df):,}")

# ============================================================
# STEP 6: FINAL VERIFICATION - NO MATERIAL > 200
# ============================================================
# Re-check after V1 addition
md_counts = v2_final_df.groupby('MATERIAL DESCRIPTION').size()
over_200 = md_counts[md_counts > 200]
if len(over_200) > 0:
    print(f"\nWARNING: {len(over_200)} materials exceed 200 after V1 addition!")
    for md, cnt in over_200.items():
        print(f"  {md}: {cnt}")
    # Need to cap these too
    for md in over_200.index:
        # Get all pairs for this material
        mat_pairs = v2_final_df[v2_final_df['MATERIAL DESCRIPTION'] == md]
        # Keep 200
        kept = mat_pairs.sample(n=200, random_state=42)
        # Remove excess
        v2_final_df = v2_final_df.drop(mat_pairs.index.difference(kept.index))
    print(f"After re-capping: {len(v2_final_df):,} pairs")
else:
    print(f"\nNo materials exceed 200 after V1 addition.")

# Final recount
md_counts = v2_final_df.groupby('MATERIAL DESCRIPTION').size()
v2_variant_counts = v2_final_df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()

# ============================================================
# STEP 7: SAVE FINAL DATASET
# ============================================================
v2_final_df.to_csv('v2_positive_pairs_final.csv', index=False)
print(f"\nSaved v2_positive_pairs_final.csv")

# Also save historical-only for reference
v2_hist_df.to_csv('v2_historical_pairs_final.csv', index=False)
print(f"Saved v2_historical_pairs_final.csv")

print("\n" + "=" * 70)
print("COMPLETE VALIDATION")
print("=" * 70)

# ============================================================
# VALIDATION CHECKS
# ============================================================

# 1. Final number of positive pairs
print(f"\n1. Final positive pairs: {len(v2_final_df):,}")

# 2. Number of unique materials
n_materials = v2_final_df['MATERIAL DESCRIPTION'].nunique()
print(f"2. Unique materials: {n_materials:,}")

# 3. Examples per material
print(f"3. Examples per material:")
print(f"   Min: {md_counts.min()}")
print(f"   Median: {md_counts.median():.1f}")
print(f"   Mean: {md_counts.mean():.1f}")
print(f"   P90: {md_counts.quantile(0.90):.1f}")
print(f"   P95: {md_counts.quantile(0.95):.1f}")
print(f"   Max: {md_counts.max()}")

# 4. Maximum examples <= 200
max_examples = md_counts.max()
print(f"4. Maximum examples per material: {max_examples} {'<= 200 ✓' if max_examples <= 200 else '> 200 ✗'}")

# 5. All 9 excluded materials have zero rows
excluded_check = v2_final_df[v2_final_df['MATERIAL DESCRIPTION'].isin(excluded_materials)]
print(f"5. Excluded materials rows: {len(excluded_check)} {'= 0 ✓' if len(excluded_check) == 0 else '> 0 ✗'}")

# 6. Zero duplicate positive pairs
dup_pairs = v2_final_df.duplicated(subset=['PRODUCT_NAME', 'MATERIAL DESCRIPTION']).sum()
print(f"6. Duplicate positive pairs: {dup_pairs} {'= 0 ✓' if dup_pairs == 0 else '> 0 ✗'}")

# 7. Zero mapping conflicts
pn_to_md = v2_final_df.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
conflicts = (pn_to_md > 1).sum()
print(f"7. Mapping conflicts (PN -> multiple MD): {conflicts} {'= 0 ✓' if conflicts == 0 else '> 0 ✗'}")

# 8. Zero evaluation Input_Column leakage
eval_inputs = set(eval_set['Input_Column'].unique())
v2_inputs = set(v2_final_df['PRODUCT_NAME'].unique())
input_overlap = eval_inputs & v2_inputs
print(f"8. Evaluation Input_Column leakage: {len(input_overlap)} {'= 0 ✓' if len(input_overlap) == 0 else '> 0 ✗'}")

# 9. Zero exact evaluation pair leakage
eval_pairs_set = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
v2_pairs_set = set(zip(v2_final_df['PRODUCT_NAME'], v2_final_df['MATERIAL DESCRIPTION']))
pair_overlap = eval_pairs_set & v2_pairs_set
print(f"9. Exact evaluation pair leakage: {len(pair_overlap)} {'= 0 ✓' if len(pair_overlap) == 0 else '> 0 ✗'}")

# 10. Evaluation set NOT modified
print(f"10. Evaluation set rows: {len(eval_set)} (unchanged)")

# 11. All remaining positive targets exist in PRODUCT_MASTER
master_products = set(product_master['product_name'].unique())
v2_materials = set(v2_final_df['MATERIAL DESCRIPTION'].unique())
not_in_master = v2_materials - master_products
print(f"11. Materials NOT in PRODUCT_MASTER: {len(not_in_master)} {'= 0 ✓' if len(not_in_master) == 0 else f'> 0 ✗: {not_in_master}'}")

# 12. V1 training pairs handling
v1_retained = v2_pairs_set & v1_all_pairs
print(f"12. V1 pairs retained in V2: {len(v1_retained)}")
print(f"    V1 training_split pairs in V2: {len(v2_pairs_set & v1_train_pairs)}")
print(f"    V1 hard_negative pairs in V2: {len(v2_pairs_set & v1_hn_pairs)}")

# 13. Material coverage statistics
print(f"\n13. MATERIAL COVERAGE STATISTICS")
print(f"    V2 materials in PRODUCT_MASTER: {len(v2_materials & master_products):,} / {n_materials:,} = {len(v2_materials & master_products)/n_materials*100:.1f}%")
print(f"    PRODUCT_MASTER covered: {len(v2_materials & master_products):,} / {len(master_products):,} = {len(v2_materials & master_products)/len(master_products)*100:.1f}%")
print(f"    PRODUCT_MASTER NOT covered: {len(master_products - v2_materials):,}")

# Long-tail
long_tail = md_counts[md_counts <= 5]
long_tail_in_master = [m for m in long_tail.index if m in master_products]
print(f"    Long-tail materials (<=5): {len(long_tail):,}")
print(f"    Long-tail in PRODUCT_MASTER: {len(long_tail_in_master):,} ({len(long_tail_in_master)/len(long_tail)*100:.1f}%)")

# Capped materials
capped = md_counts[md_counts == 200]
print(f"    Materials at cap (200): {len(capped):,}")

# Variant distribution
print(f"\n    Variant distribution (unique PRODUCT_NAME per MATERIAL):")
buckets = {
    '1': (1, 1),
    '2': (2, 2),
    '3-5': (3, 5),
    '6-10': (6, 10),
    '11-25': (11, 25),
    '26-50': (26, 50),
    '51-100': (51, 100),
    '101-200': (101, 200),
}
for bucket_name, (low, high) in buckets.items():
    if high == 200:
        mask = v2_variant_counts >= low
    else:
        mask = (v2_variant_counts >= low) & (v2_variant_counts <= high)
    mats_in_bucket = mask.sum()
    print(f"      {bucket_name}: {mats_in_bucket:,} materials")

# Final verdict
print("\n" + "=" * 70)
all_checks_pass = (
    max_examples <= 200 and
    len(excluded_check) == 0 and
    dup_pairs == 0 and
    conflicts == 0 and
    len(input_overlap) == 0 and
    len(pair_overlap) == 0 and
    len(not_in_master) == 0
)
print(f"V2 POSITIVE DATASET READY FOR HARD-NEGATIVE MINING: {'YES ✓' if all_checks_pass else 'NO ✗'}")
print("=" * 70)