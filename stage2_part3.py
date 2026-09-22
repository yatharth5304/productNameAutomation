import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
train_split = pd.read_csv('training_split.csv')
hard_neg = pd.read_csv('hard_negative_mining.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

# Clean
no_product_mask = (df['MATERIAL DESCRIPTION'] == '-') | (df['MATERIAL CODE'] == 'NO PRODUCT')
df_clean = df[~no_product_mask].copy()
df_clean = df_clean.drop_duplicates().reset_index(drop=True)

md_variant_counts = df_clean.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
md_row_counts = df_clean.groupby('MATERIAL DESCRIPTION').size()

print("=" * 70)
print("7. HIGH-FREQUENCY MATERIALS")
print("=" * 70)

top_mats = md_variant_counts.sort_values(ascending=False).head(30)
for i, (md, var_count) in enumerate(top_mats.items(), 1):
    total_rows = md_row_counts[md]
    print(f"{i:2d}. {md} — {var_count:,} variants, {total_rows:,} rows")

# Near-duplicate check
print(f"\nNear-duplicate check for top 10 materials:")
def normalize(s):
    return s.upper().replace('-', ' ').replace('.', '').replace('  ', ' ').strip()

for md, var_count in top_mats.head(10).items():
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    normalized = [normalize(v) for v in variants]
    unique_norm = len(set(normalized))
    print(f"  {md}: {var_count} raw -> {unique_norm} normalized ({unique_norm/var_count*100:.1f}% unique)")

print("\n" + "=" * 70)
print("8. EVALUATION SET PROTECTION")
print("=" * 70)

eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
clean_pairs = set(zip(df_clean['PRODUCT_NAME'], df_clean['MATERIAL DESCRIPTION']))

overlap = eval_pairs & clean_pairs
print(f"Evaluation pairs: {len(eval_pairs):,}")
print(f"Clean dataset pairs: {len(clean_pairs):,}")
print(f"Overlap (exact Input_Column + Positive_Product): {len(overlap):,}")

eval_inputs = set(eval_set['Input_Column'].unique())
clean_inputs = set(df_clean['PRODUCT_NAME'].unique())
input_overlap = eval_inputs & clean_inputs
print(f"Evaluation Input_Column overlap: {len(input_overlap):,} / {len(eval_inputs):,} = {len(input_overlap)/len(eval_inputs)*100:.2f}%")

conflicts = 0
for inp in input_overlap:
    eval_target = eval_set[eval_set['Input_Column'] == inp]['Positive_Product'].values[0]
    clean_targets = set(df_clean[df_clean['PRODUCT_NAME'] == inp]['MATERIAL DESCRIPTION'].unique())
    if eval_target not in clean_targets:
        conflicts += 1
print(f"Conflicts (eval input maps to different target): {conflicts}")

# Also check hard negative mining
hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
hn_overlap = hn_pairs & clean_pairs
hn_inputs = set(hard_neg['Input_Column'].unique())
hn_input_overlap = hn_inputs & clean_inputs
print(f"\nHard negative mining pair overlap: {len(hn_overlap):,} / {len(hn_pairs):,}")
print(f"Hard negative mining input overlap: {len(hn_input_overlap):,} / {len(hn_inputs):,}")

print("\n" + "=" * 70)
print("9. V1 TRAINING DATA RELATIONSHIP")
print("=" * 70)

v1_train_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product']))
v1_hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
v1_all_pairs = v1_train_pairs | v1_hn_pairs
hist_pairs = clean_pairs

overlap_train = hist_pairs & v1_train_pairs
overlap_hn = hist_pairs & v1_hn_pairs
overlap_all = hist_pairs & v1_all_pairs

print(f"Historical unique pairs: {len(hist_pairs):,}")
print(f"V1 training_split pairs: {len(v1_train_pairs):,}")
print(f"V1 hard_negative_mining pairs: {len(v1_hn_pairs):,}")
print(f"V1 all pairs: {len(v1_all_pairs):,}")

print(f"Overlap with V1 training_split: {len(overlap_train):,} / {len(v1_train_pairs):,} = {len(overlap_train)/len(v1_train_pairs)*100:.2f}%")
print(f"Overlap with V1 hard_negative: {len(overlap_hn):,} / {len(v1_hn_pairs):,} = {len(overlap_hn)/len(v1_hn_pairs)*100:.2f}%")
print(f"Overlap with V1 all: {len(overlap_all):,} / {len(v1_all_pairs):,} = {len(overlap_all)/len(v1_all_pairs)*100:.2f}%")

new_pairs = hist_pairs - v1_all_pairs
print(f"New positive pairs from historical (not in V1): {len(new_pairs):,}")

v1_materials = set([p[1] for p in v1_all_pairs])
hist_materials = set([p[1] for p in hist_pairs])
common_materials = v1_materials & hist_materials

print(f"Materials in V1: {len(v1_materials)}")
print(f"Materials in historical: {len(hist_materials)}")
print(f"Common materials: {len(common_materials)}")

print(f"\nNew variants gained for common materials (top 20 by gain):")
gains = []
for md in common_materials:
    v1_variants = set(p[0] for p in v1_all_pairs if p[1] == md)
    hist_variants = set(p[0] for p in hist_pairs if p[1] == md)
    new_variants = hist_variants - v1_variants
    if len(new_variants) > 0:
        gains.append((md, len(v1_variants), len(hist_variants), len(new_variants)))

gains.sort(key=lambda x: -x[3])
for md, v1_v, hist_v, new_v in gains[:20]:
    print(f"  {md}: V1={v1_v}, Hist={hist_v}, New={new_v}")