import pandas as pd
import numpy as np

# Load all data
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
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

master_products = set(product_master['product_name'].unique())

# Also need to exclude any V1 materials not in PRODUCT_MASTER
# Let's check V1 materials
v1_train_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product']))
v1_hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
v1_all_pairs = v1_train_pairs | v1_hn_pairs

v1_materials = set(md for _, md in v1_all_pairs)
v1_not_in_master = v1_materials - master_products
print(f"V1 materials not in PRODUCT_MASTER: {len(v1_not_in_master)}")
for m in sorted(v1_not_in_master):
    print(f"  {m}")

print("=" * 70)
print("RECONSTRUCTING V2 WITH COMPLETE EXCLUSIONS")
print("=" * 70)

# STEP 1: Clean historical
no_product_mask = (df['MATERIAL DESCRIPTION'] == '-') | (df['MATERIAL CODE'] == 'NO PRODUCT')
df_clean = df[~no_product_mask].copy()
df_clean = df_clean.drop_duplicates().reset_index(drop=True)

# STEP 2: Exclude 9 materials from historical
df_clean = df_clean[~df_clean['MATERIAL DESCRIPTION'].isin(excluded_materials)].copy()

# STEP 3: Get unique positive pairs
pair_df = df_clean[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates().reset_index(drop=True)

# STEP 4: D3 Hybrid capping
md_variants = pair_df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].apply(set).to_dict()

np.random.seed(42)
final_pairs = []

for md, variants in md_variants.items():
    n = len(variants)
    if n > 200:
        selected = np.random.choice(list(variants), size=200, replace=False)
    else:
        selected = list(variants)
    for pn in selected:
        final_pairs.append({'PRODUCT_NAME': pn, 'MATERIAL DESCRIPTION': md})

v2_hist_df = pd.DataFrame(final_pairs)
hist_pairs = set(zip(v2_hist_df['PRODUCT_NAME'], v2_hist_df['MATERIAL DESCRIPTION']))

# STEP 5: V1 pairs - ONLY add those that:
# - Not in historical
# - Not in evaluation
# - Material IS in PRODUCT_MASTER
# - Material NOT in excluded_materials

eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
eval_inputs = set(eval_set['Input_Column'].unique())

v1_new = v1_all_pairs - hist_pairs
print(f"V1 new pairs: {len(v1_new)}")

# Filter V1 new pairs
v1_filtered = set()
for pn, md in v1_new:
    if (pn, md) in eval_pairs:
        continue  # Skip eval overlap
    if pn in eval_inputs:
        continue  # Skip eval input
    if md in excluded_materials:
        continue  # Skip excluded materials
    if md not in master_products:
        continue  # Skip materials not in PRODUCT_MASTER
    v1_filtered.add((pn, md))

print(f"V1 filtered pairs (added): {len(v1_filtered)}")

# STEP 6: Combine
all_pairs = hist_pairs | v1_filtered
v2_final_df = pd.DataFrame(list(all_pairs), columns=['PRODUCT_NAME', 'MATERIAL DESCRIPTION'])

# STEP 7: Final recapping - ensure no material > 200
md_counts = v2_final_df.groupby('MATERIAL DESCRIPTION').size()
over_200 = md_counts[md_counts > 200]
if len(over_200) > 0:
    print(f"Recapping {len(over_200)} materials exceeding 200...")
    for md in over_200.index:
        mat_rows = v2_final_df[v2_final_df['MATERIAL DESCRIPTION'] == md]
        kept = mat_rows.sample(n=200, random_state=42)
        v2_final_df = v2_final_df.drop(mat_rows.index.difference(kept.index))

# Save
v2_final_df.to_csv('v2_positive_pairs_final.csv', index=False)
print(f"Saved final dataset: {len(v2_final_df):,} pairs")

# Verify exclusions
excluded_check = v2_final_df[v2_final_df['MATERIAL DESCRIPTION'].isin(excluded_materials)]
print(f"Excluded materials in final: {len(excluded_check)}")

# Verify all materials in PRODUCT_MASTER
v2_materials = set(v2_final_df['MATERIAL DESCRIPTION'].unique())
not_in_master = v2_materials - master_products
print(f"Materials not in PRODUCT_MASTER: {len(not_in_master)}")
for m in sorted(not_in_master):
    print(f"  {m}")