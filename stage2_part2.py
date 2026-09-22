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

print("=" * 70)
print("4. MATERIAL FREQUENCY DISTRIBUTION")
print("=" * 70)

md_variant_counts = df_clean.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
md_row_counts = df_clean.groupby('MATERIAL DESCRIPTION').size()

total_materials = len(md_variant_counts)
total_examples = df_clean.shape[0]

buckets = {
    '1': (1, 1),
    '2': (2, 2),
    '3-5': (3, 5),
    '6-10': (6, 10),
    '11-25': (11, 25),
    '26-50': (26, 50),
    '51-100': (51, 100),
    '101-250': (101, 250),
    '251-500': (251, 500),
    '500+': (501, float('inf')),
}

print(f"{'Bucket':<12} {'Materials':>10} {'%Mats':>10} {'Examples':>12} {'%Ex':>10}")
print("-" * 55)

for bucket_name, (low, high) in buckets.items():
    if high == float('inf'):
        mask = md_variant_counts >= low
    else:
        mask = (md_variant_counts >= low) & (md_variant_counts <= high)
    mats_in_bucket = mask.sum()
    mats_list = md_variant_counts[mask].index
    examples_in_bucket = md_row_counts[mats_list].sum() if len(mats_list) > 0 else 0
    print(f"{bucket_name:<12} {mats_in_bucket:>10,} {mats_in_bucket/total_materials*100:>9.2f}% {examples_in_bucket:>12,} {examples_in_bucket/total_examples*100:>9.2f}%")

print("\n" + "=" * 70)
print("5. OCR DIVERSITY")
print("=" * 70)

print(f"Unique OCR variants per material:")
print(f"  Mean: {md_variant_counts.mean():.1f}")
print(f"  Median: {md_variant_counts.median():.1f}")
print(f"  P90: {md_variant_counts.quantile(0.90):.1f}")
print(f"  P95: {md_variant_counts.quantile(0.95):.1f}")
print(f"  Max: {md_variant_counts.max():.0f}")

print(f"\nLow diversity (1-2 variants):")
low_div = md_variant_counts[md_variant_counts <= 2].index[:10]
for md in low_div:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"  {md} ({len(variants)}): {list(variants)[:5]}")

print(f"\nMedium diversity (10-20 variants):")
med_div = md_variant_counts[(md_variant_counts >= 10) & (md_variant_counts <= 20)].index[:5]
for md in med_div:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"  {md} ({len(variants)} variants):")
    for v in variants[:8]:
        print(f"    {v}")

print(f"\nHigh diversity (500+ variants):")
high_div = md_variant_counts[md_variant_counts >= 500].index[:5]
for md in high_div:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"  {md} ({len(variants)} variants):")
    for v in variants[:8]:
        print(f"    {v}")

print("\n" + "=" * 70)
print("6. LONG-TAIL IMPACT")
print("=" * 70)

long_tail_mats = md_variant_counts[md_variant_counts <= 5].index
print(f"Materials with <=5 variants: {len(long_tail_mats):,} / {total_materials:,} = {len(long_tail_mats)/total_materials*100:.2f}%")

master_products = set(product_master['product_name'].unique())
long_tail_in_master = [m for m in long_tail_mats if m in master_products]
print(f"Long-tail materials in PRODUCT_MASTER: {len(long_tail_in_master):,} / {len(long_tail_mats):,} = {len(long_tail_in_master)/len(long_tail_mats)*100:.2f}%")

long_tail_examples = df_clean[df_clean['MATERIAL DESCRIPTION'].isin(long_tail_mats)].shape[0]
print(f"Training examples from long-tail: {long_tail_examples:,} / {df_clean.shape[0]:,} = {long_tail_examples/df_clean.shape[0]*100:.2f}%")

master_covered = master_products & set(df_clean['MATERIAL DESCRIPTION'].unique())
print(f"\nPRODUCT_MASTER covered by historical: {len(master_covered):,} / {len(master_products):,} = {len(master_covered)/len(master_products)*100:.2f}%")

covered_long_tail = [m for m in master_covered if m in long_tail_mats]
print(f"Covered products that are long-tail: {len(covered_long_tail):,} / {len(master_covered):,} = {len(covered_long_tail)/len(master_covered)*100:.2f}%")

# Show some long-tail materials in master
print(f"\nLong-tail materials in PRODUCT_MASTER (examples):")
for m in sorted(covered_long_tail)[:20]:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == m]['PRODUCT_NAME'].nunique()
    rows = len(df_clean[df_clean['MATERIAL DESCRIPTION'] == m])
    print(f"  {m}: {variants} variants, {rows} rows")