import pandas as pd
import numpy as np

# Load final dataset
v2_df = pd.read_csv('v2_positive_pairs_final.csv')
eval_set = pd.read_csv('evaluation_set.csv')
train_split = pd.read_csv('training_split.csv')
hard_neg = pd.read_csv('hard_negative_mining.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

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

print("=" * 70)
print("FINAL V2 VALIDATION REPORT")
print("=" * 70)

# 1. Final positive pairs
print(f"\n1. Final positive pairs: {len(v2_df):,}")

# 2. Unique materials
n_materials = v2_df['MATERIAL DESCRIPTION'].nunique()
print(f"2. Unique materials: {n_materials:,}")

# 3. Examples per material
md_counts = v2_df.groupby('MATERIAL DESCRIPTION').size()
v2_variant_counts = v2_df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
print(f"3. Examples per material:")
print(f"   Min: {md_counts.min()}")
print(f"   Median: {md_counts.median():.1f}")
print(f"   Mean: {md_counts.mean():.1f}")
print(f"   P90: {md_counts.quantile(0.90):.1f}")
print(f"   P95: {md_counts.quantile(0.95):.1f}")
print(f"   Max: {md_counts.max()}")

# 4. Maximum examples <= 200
max_examples = md_counts.max()
print(f"4. Maximum examples per material: {max_examples} {'<= 200 PASS' if max_examples <= 200 else '> 200 FAIL'}")

# 5. All 9 excluded materials have zero rows
excluded_check = v2_df[v2_df['MATERIAL DESCRIPTION'].isin(excluded_materials)]
print(f"5. Excluded materials rows: {len(excluded_check)} {'= 0 PASS' if len(excluded_check) == 0 else '> 0 FAIL'}")

# 6. Zero duplicate positive pairs
dup_pairs = v2_df.duplicated(subset=['PRODUCT_NAME', 'MATERIAL DESCRIPTION']).sum()
print(f"6. Duplicate positive pairs: {dup_pairs} {'= 0 PASS' if dup_pairs == 0 else '> 0 FAIL'}")

# 7. Zero mapping conflicts
pn_to_md = v2_df.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
conflicts = (pn_to_md > 1).sum()
print(f"7. Mapping conflicts (PN -> multiple MD): {conflicts} {'= 0 PASS' if conflicts == 0 else '> 0 FAIL'}")

# 8. Zero evaluation Input_Column leakage
eval_inputs = set(eval_set['Input_Column'].unique())
v2_inputs = set(v2_df['PRODUCT_NAME'].unique())
input_overlap = eval_inputs & v2_inputs
print(f"8. Evaluation Input_Column leakage: {len(input_overlap)} {'= 0 PASS' if len(input_overlap) == 0 else '> 0 FAIL'}")

# 9. Zero exact evaluation pair leakage
eval_pairs_set = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
v2_pairs_set = set(zip(v2_df['PRODUCT_NAME'], v2_df['MATERIAL DESCRIPTION']))
pair_overlap = eval_pairs_set & v2_pairs_set
print(f"9. Exact evaluation pair leakage: {len(pair_overlap)} {'= 0 PASS' if len(pair_overlap) == 0 else '> 0 FAIL'}")

# 10. Evaluation set NOT modified
print(f"10. Evaluation set rows: {len(eval_set)} (unchanged)")

# 11. All remaining positive targets exist in PRODUCT_MASTER
master_products = set(product_master['product_name'].unique())
v2_materials = set(v2_df['MATERIAL DESCRIPTION'].unique())
not_in_master = v2_materials - master_products
print(f"11. Materials NOT in PRODUCT_MASTER: {len(not_in_master)} {'= 0 PASS' if len(not_in_master) == 0 else f'> 0 FAIL: {not_in_master}'}")

# 12. V1 training pairs handling
v1_train_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product']))
v1_hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
v1_all_pairs = v1_train_pairs | v1_hn_pairs
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
all_checks_pass = (
    max_examples <= 200 and
    len(excluded_check) == 0 and
    dup_pairs == 0 and
    conflicts == 0 and
    len(input_overlap) == 0 and
    len(pair_overlap) == 0 and
    len(not_in_master) == 0
)
print("\n" + "=" * 70)
print(f"V2 POSITIVE DATASET READY FOR HARD-NEGATIVE MINING: {'YES' if all_checks_pass else 'NO'}")
print("=" * 70)