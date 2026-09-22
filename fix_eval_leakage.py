import pandas as pd
import numpy as np

# Load data
eval_set = pd.read_csv('evaluation_set.csv')
v2_df = pd.read_csv('v2_positive_pairs_validated.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 70)
print("FIXING EVALUATION LEAKAGE")
print("=" * 70)

# The issue: V1 pairs (which include evaluation pairs) were added to V2
# We need to remove ALL pairs where the positive product is in evaluation set
# AND the input column matches, but more broadly, we should not have ANY
# evaluation input columns in training

eval_inputs = set(eval_set['Input_Column'].unique())
eval_products = set(eval_set['Positive_Product'].unique())
eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))

v2_pairs = set(zip(v2_df['PRODUCT_NAME'], v2_df['MATERIAL DESCRIPTION']))

# Check current overlap
overlap = eval_pairs & v2_pairs
print(f"Evaluation pairs in V2: {len(overlap)}")

# The fix: Remove from V2 any pair where:
# 1. The PRODUCT_NAME is an evaluation Input_Column (170 unique inputs)
# 2. The pair exactly matches an evaluation pair

# Strategy: Remove all V2 pairs where PRODUCT_NAME is in eval_inputs
# This is the safest - don't train on any evaluation input
v2_filtered = v2_df[~v2_df['PRODUCT_NAME'].isin(eval_inputs)].copy()

# Also remove any pair that exactly matches an evaluation pair (double safety)
v2_filtered = v2_filtered[~v2_filtered.apply(lambda r: (r['PRODUCT_NAME'], r['MATERIAL DESCRIPTION']) in eval_pairs, axis=1)]

print(f"V2 before filtering: {len(v2_df):,}")
print(f"V2 after filtering: {len(v2_filtered):,}")
print(f"Removed: {len(v2_df) - len(v2_filtered):,}")

# Verify
v2_filtered_pairs = set(zip(v2_filtered['PRODUCT_NAME'], v2_filtered['MATERIAL DESCRIPTION']))
eval_overlap_after = eval_pairs & v2_filtered_pairs
eval_input_overlap_after = eval_inputs & set(v2_filtered['PRODUCT_NAME'].unique())

print(f"\nAfter filtering:")
print(f"  Evaluation pair overlap: {len(eval_overlap_after)}")
print(f"  Evaluation input overlap: {len(eval_input_overlap_after)}")

# Check materials still represented
materials_after = v2_filtered['MATERIAL DESCRIPTION'].nunique()
print(f"  Materials: {materials_after:,}")

# Check long-tail
md_counts = v2_filtered.groupby('MATERIAL DESCRIPTION').size()
long_tail = md_counts[md_counts <= 5]
print(f"  Long-tail materials (<=5): {len(long_tail):,}")

# Save fixed dataset
v2_filtered.to_csv('v2_positive_pairs_final.csv', index=False)
print(f"\nSaved v2_positive_pairs_final.csv")

# Also check the 11 materials not in PRODUCT_MASTER - should we keep them?
v2_materials = set(v2_filtered['MATERIAL DESCRIPTION'].unique())
master_products = set(product_master['product_name'].unique())
not_in_master = v2_materials - master_products
print(f"\nMaterials NOT in PRODUCT_MASTER: {len(not_in_master)}")
for m in sorted(not_in_master):
    count = len(v2_filtered[v2_filtered['MATERIAL DESCRIPTION'] == m])
    print(f"  {m}: {count} examples")

# Check if any of these 11 are in evaluation
eval_products_in_v2 = eval_products & v2_materials
print(f"\nEvaluation products in V2 (as positives): {len(eval_products_in_v2)}")
for p in sorted(eval_products_in_v2):
    count = len(v2_filtered[v2_filtered['MATERIAL DESCRIPTION'] == p])
    print(f"  {p}: {count} OCR variants in training")

# Final stats
print("\n" + "=" * 70)
print("FINAL V2 DATASET STATISTICS")
print("=" * 70)
print(f"Total positive pairs: {len(v2_filtered):,}")
print(f"Unique PRODUCT_NAME: {v2_filtered['PRODUCT_NAME'].nunique():,}")
print(f"Unique MATERIAL DESCRIPTION: {materials_after:,}")
print(f"Examples per material:")
print(f"  Min: {md_counts.min()}")
print(f"  Median: {md_counts.median():.1f}")
print(f"  Mean: {md_counts.mean():.1f}")
print(f"  P90: {md_counts.quantile(0.90):.1f}")
print(f"  P95: {md_counts.quantile(0.95):.1f}")
print(f"  Max: {md_counts.max()}")