import pandas as pd
import numpy as np

# Load all data
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
train_split = pd.read_csv('training_split.csv')
hard_neg = pd.read_csv('hard_negative_mining.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 70)
print("TASK 1: CONSTRUCT V2 POSITIVE DATASET (D3 HYBRID)")
print("=" * 70)

# Step 1: Clean the historical dataset
original_rows = len(df)
no_product_mask = (df['MATERIAL DESCRIPTION'] == '-') | (df['MATERIAL CODE'] == 'NO PRODUCT')
df_clean = df[~no_product_mask].copy()
df_clean = df_clean.drop_duplicates().reset_index(drop=True)

print(f"Original rows: {original_rows:,}")
print(f"After NO PRODUCT removal: {len(df_clean):,}")
print(f"After dedup: {len(df_clean):,}")

# Step 2: Get unique positive pairs (PRODUCT_NAME -> MATERIAL DESCRIPTION)
# We need unique pairs with their OCR variants
pair_df = df_clean[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates().reset_index(drop=True)
print(f"Unique positive pairs: {len(pair_df):,}")

# Step 3: Count variants per material
md_variant_counts = df_clean.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
md_variant_counts.name = 'num_variants'

# Step 4: Apply D3 hybrid strategy - cap at 200 for materials with >100 variants
# For materials with >100 variants, select 200 diverse variants
# For materials with <=100 variants, keep all

print(f"\nMaterials with >100 variants: {(md_variant_counts > 100).sum()}")
print(f"Materials with >200 variants: {(md_variant_counts > 200).sum()}")

# For capped materials, we need to select 200 diverse variants
# Strategy: Use the original data order but sample across parties to preserve diversity
# Better: sample proportionally from each PARTY_CODE for that material

def select_diverse_variants(df_material, target_n):
    """Select up to target_n diverse variants from a material's data"""
    if len(df_material) <= target_n:
        return df_material
    
    # Strategy: Sample from each party proportionally to preserve party diversity
    # First, get unique variants per party
    party_variants = df_material.groupby('PARTY_CODE')['PRODUCT_NAME'].apply(lambda x: list(x.unique()))
    
    selected = []
    remaining = target_n
    
    # Round-robin selection from parties
    parties = list(party_variants.keys())
    party_idx = {p: 0 for p in parties}
    
    while remaining > 0 and any(party_idx[p] < len(party_variants[p]) for p in parties):
        for p in parties:
            if remaining <= 0:
                break
            if party_idx[p] < len(party_variants[p]):
                variant = party_variants[p][party_idx[p]]
                party_idx[p] += 1
                # Get one row for this variant
                row = df_material[(df_material['PARTY_CODE'] == p) & (df_material['PRODUCT_NAME'] == variant)].iloc[0]
                selected.append(row)
                remaining -= 1
    
    return pd.DataFrame(selected)

# Apply capping
print(f"\nApplying D3 hybrid capping (max 200 variants for materials with >100)...")
final_rows = []
capped_materials = []

for md in md_variant_counts.index:
    df_mat = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]
    n_variants = md_variant_counts[md]
    
    if n_variants > 100:
        # Cap at 200
        selected = select_diverse_variants(df_mat, 200)
        final_rows.append(selected)
        capped_materials.append((md, n_variants, len(selected['PRODUCT_NAME'].unique())))
    else:
        # Keep all
        final_rows.append(df_mat)

df_v2 = pd.concat(final_rows, ignore_index=True)

# Get unique pairs in final dataset
v2_pairs = df_v2[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates()
print(f"\nV2 historical pairs after capping: {len(v2_pairs):,}")

# Verify material retention
v2_materials = v2_pairs['MATERIAL DESCRIPTION'].nunique()
print(f"Materials retained: {v2_materials:,}")

# Check variant distribution
v2_variant_counts = v2_pairs.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
print(f"Variant stats after capping:")
print(f"  Min: {v2_variant_counts.min()}")
print(f"  Median: {v2_variant_counts.median():.1f}")
print(f"  Mean: {v2_variant_counts.mean():.1f}")
print(f"  P90: {v2_variant_counts.quantile(0.90):.1f}")
print(f"  P95: {v2_variant_counts.quantile(0.95):.1f}")
print(f"  Max: {v2_variant_counts.max()}")

print(f"\nMaterials capped: {len(capped_materials)}")
for md, orig, new in capped_materials[:10]:
    print(f"  {md}: {orig} -> {new} variants")

print(f"\nTotal capped materials: {len(capped_materials)}")

# Save V2 historical pairs
v2_pairs.to_csv('v2_historical_pairs.csv', index=False)
print(f"\nSaved v2_historical_pairs.csv with {len(v2_pairs):,} pairs")