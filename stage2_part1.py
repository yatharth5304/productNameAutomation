import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
train_split = pd.read_csv('training_split.csv')
hard_neg = pd.read_csv('hard_negative_mining.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 70)
print("1. CLEANED DATASET SIZE")
print("=" * 70)

original_rows = len(df)
print(f"Original rows: {original_rows:,}")

no_product_mask = (df['MATERIAL DESCRIPTION'] == '-') | (df['MATERIAL CODE'] == 'NO PRODUCT')
no_product_rows = no_product_mask.sum()
df_clean = df[~no_product_mask].copy()
print(f"NO PRODUCT rows removed: {no_product_rows:,}")

exact_dupes = df_clean.duplicated().sum()
df_clean = df_clean.drop_duplicates().reset_index(drop=True)
print(f"Exact duplicate rows removed: {exact_dupes:,}")

remaining_rows = len(df_clean)
print(f"Remaining usable rows: {remaining_rows:,}")
print(f"Retention rate: {remaining_rows/original_rows*100:.2f}%")

print("\n" + "=" * 70)
print("2. POSITIVE PAIR ANALYSIS")
print("=" * 70)

unique_pn = df_clean['PRODUCT_NAME'].nunique()
unique_md = df_clean['MATERIAL DESCRIPTION'].nunique()
unique_pairs = df_clean[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates().shape[0]

print(f"Unique PRODUCT_NAME: {unique_pn:,}")
print(f"Unique MATERIAL DESCRIPTION: {unique_md:,}")
print(f"Unique PRODUCT_NAME + MATERIAL DESCRIPTION pairs: {unique_pairs:,}")

pair_counts = df_clean.groupby(['PRODUCT_NAME', 'MATERIAL DESCRIPTION']).size()
dup_pairs = (pair_counts > 1).sum()
print(f"Duplicate positive pairs: {dup_pairs:,}")

pn_to_md = df_clean.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
pn_multi_md = (pn_to_md > 1).sum()
print(f"PRODUCT_NAME -> multiple MATERIAL DESCRIPTION: {pn_multi_md:,} / {len(pn_to_md):,} = {pn_multi_md/len(pn_to_md)*100:.2f}%")

md_to_pn = df_clean.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
md_multi_pn = (md_to_pn > 1).sum()
print(f"MATERIAL DESCRIPTION <- multiple PRODUCT_NAME: {md_multi_pn:,} / {len(md_to_pn):,} = {md_multi_pn/len(md_to_pn)*100:.2f}%")

print("\n" + "=" * 70)
print("3. CONFLICT ANALYSIS")
print("=" * 70)

if pn_multi_md > 0:
    conflicts = pn_to_md[pn_to_md > 1].index
    print(f"Found {len(conflicts)} PRODUCT_NAME with multiple MATERIAL DESCRIPTION")
    for pn in conflicts[:20]:
        mds = df_clean[df_clean['PRODUCT_NAME'] == pn]['MATERIAL DESCRIPTION'].unique()
        print(f"  {pn}")
        for md in mds:
            count = len(df_clean[(df_clean['PRODUCT_NAME'] == pn) & (df_clean['MATERIAL DESCRIPTION'] == md)])
            print(f"    -> {md} ({count}x)")
else:
    print("No PRODUCT_NAME maps to multiple MATERIAL DESCRIPTION (0 conflicts)")