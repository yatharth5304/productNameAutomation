import pandas as pd
import numpy as np

# Load data
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
train_split = pd.read_csv('training_split.csv')
hard_neg = pd.read_csv('hard_negative_mining.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

# Clean
no_product_mask = (df['MATERIAL DESCRIPTION'] == '-') | (df['MATERIAL CODE'] == 'NO PRODUCT')
df_clean = df[~no_product_mask].copy()
df_clean = df_clean.drop_duplicates().reset_index(drop=True)

# Get unique pairs
pair_df = df_clean[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates().reset_index(drop=True)

# Get variant counts per material
md_variants = pair_df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].apply(set).to_dict()

print("=" * 70)
print("RECALCULATING D3 HYBRID STRATEGY CORRECTLY")
print("=" * 70)

# D3: retain <=100, cap >200 at 200 (materials with 101-200 keep all)
total_pairs = 0
capped_count = 0
for md, variants in md_variants.items():
    n = len(variants)
    if n > 200:
        total_pairs += 200
        capped_count += 1
    elif n > 100:
        total_pairs += n
        capped_count += 1
    else:
        total_pairs += n

print(f"Expected D3 pairs: {total_pairs:,}")
print(f"Materials capped (>200): {capped_count}")
print(f"Materials with 101-200 (retained): {sum(1 for v in md_variants.values() if 100 < len(v) <= 200)}")
print(f"Materials retained as-is (<=100): {sum(1 for v in md_variants.values() if len(v) <= 100)}")

# Now properly construct the dataset
print("\n" + "=" * 70)
print("CONSTRUCTING V2 DATASET WITH PROPER SAMPLING")
print("=" * 70)

np.random.seed(42)  # Reproducible

final_pairs = []

for md, variants in md_variants.items():
    n = len(variants)
    if n > 200:
        # Sample 200 variants randomly (but reproducibly)
        selected = np.random.choice(list(variants), size=200, replace=False)
    elif n > 100:
        # 100 < n <= 200: keep all
        selected = list(variants)
    else:
        selected = list(variants)
    
    for pn in selected:
        final_pairs.append({'PRODUCT_NAME': pn, 'MATERIAL DESCRIPTION': md})

v2_hist_df = pd.DataFrame(final_pairs)
print(f"V2 historical pairs: {len(v2_hist_df):,}")

# Verify
v2_variant_counts = v2_hist_df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
print(f"\nVariant stats:")
print(f"  Min: {v2_variant_counts.min()}")
print(f"  Median: {v2_variant_counts.median():.1f}")
print(f"  Mean: {v2_variant_counts.mean():.1f}")
print(f"  P90: {v2_variant_counts.quantile(0.90):.1f}")
print(f"  P95: {v2_variant_counts.quantile(0.95):.1f}")
print(f"  Max: {v2_variant_counts.max()}")
print(f"  Materials: {v2_variant_counts.nunique()}")

# Check capped materials
capped = v2_variant_counts[v2_variant_counts == 200]
print(f"\nMaterials at cap (200): {len(capped)}")

# Save
v2_hist_df.to_csv('v2_historical_pairs.csv', index=False)
print(f"\nSaved v2_historical_pairs.csv")

# Now include V1 positives
print("\n" + "=" * 70)
print("TASK 2: INCLUDE V1 POSITIVES")
print("=" * 70)

v1_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product']))
v1_hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
v1_all_pairs = v1_pairs | v1_hn_pairs

hist_pairs = set(zip(v2_hist_df['PRODUCT_NAME'], v2_hist_df['MATERIAL DESCRIPTION']))

print(f"V1 all pairs: {len(v1_all_pairs)}")
print(f"Historical pairs: {len(hist_pairs)}")

# Check overlap
overlap = hist_pairs & v1_all_pairs
print(f"Overlap: {len(overlap)}")

# Check V1 against evaluation
eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
v1_eval_overlap = v1_all_pairs & eval_pairs
print(f"V1 pairs overlapping with evaluation: {len(v1_eval_overlap)}")

# V1 pairs not in historical (to add)
v1_new = v1_all_pairs - hist_pairs
print(f"V1 pairs to add (not in historical): {len(v1_new)}")

# Check if any V1 pairs conflict with historical (same PRODUCT_NAME -> different MATERIAL DESCRIPTION)
v1_pn_to_md = {}
for pn, md in v1_all_pairs:
    if pn not in v1_pn_to_md:
        v1_pn_to_md[pn] = md
    elif v1_pn_to_md[pn] != md:
        print(f"V1 CONFLICT: {pn} -> {v1_pn_to_md[pn]} vs {md}")

hist_pn_to_md = {}
for pn, md in hist_pairs:
    if pn not in hist_pn_to_md:
        hist_pn_to_md[pn] = md
    elif hist_pn_to_md[pn] != md:
        print(f"HIST CONFLICT: {pn} -> {hist_pn_to_md[pn]} vs {md}")

# Check cross conflicts
cross_conflicts = 0
for pn in set(v1_pn_to_md.keys()) & set(hist_pn_to_md.keys()):
    if v1_pn_to_md[pn] != hist_pn_to_md[pn]:
        cross_conflicts += 1
        if cross_conflicts <= 10:
            print(f"CROSS CONFLICT: {pn} -> V1:{v1_pn_to_md[pn]} vs HIST:{hist_pn_to_md[pn]}")
print(f"Total cross conflicts: {cross_conflicts}")

# Combine
all_pairs = hist_pairs | v1_new
v2_final_df = pd.DataFrame(list(all_pairs), columns=['PRODUCT_NAME', 'MATERIAL DESCRIPTION'])
print(f"\nFinal V2 positive pairs: {len(v2_final_df):,}")

v2_final_materials = v2_final_df['MATERIAL DESCRIPTION'].nunique()
print(f"Final materials: {v2_final_materials:,}")

# Save
v2_final_df.to_csv('v2_positive_pairs.csv', index=False)
print(f"Saved v2_positive_pairs.csv")