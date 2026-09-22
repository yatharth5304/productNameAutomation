import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("3. CORRECT MAPPING ANALYSIS (PRODUCT_NAME -> MATERIAL DESCRIPTION)")
print("=" * 60)

# Unique relationships
unique_pairs = df[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates()
print(f"Unique PRODUCT_NAME -> MATERIAL DESCRIPTION pairs: {len(unique_pairs):,}")

# How many PRODUCT_NAME map to exactly one MATERIAL DESCRIPTION
pn_to_md = df.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
single_mapping = (pn_to_md == 1).sum()
multi_mapping = (pn_to_md > 1).sum()

print(f"PRODUCT_NAME with exactly 1 MATERIAL DESCRIPTION: {single_mapping:,} / {len(pn_to_md):,} = {single_mapping/len(pn_to_md)*100:.2f}%")
print(f"PRODUCT_NAME with multiple MATERIAL DESCRIPTION: {multi_mapping:,} / {len(pn_to_md):,} = {multi_mapping/len(pn_to_md)*100:.2f}%")

# Distribution of how many MATERIAL DESCRIPTION per PRODUCT_NAME
md_per_pn_dist = pn_to_md.value_counts().sort_index()
print(f"\nDistribution of MATERIAL DESCRIPTION count per PRODUCT_NAME:")
for cnt, num_pn in md_per_pn_dist.head(20).items():
    print(f"  {cnt} MATERIAL DESCRIPTION: {num_pn:,} PRODUCT_NAME")

# Conflicting mappings examples
print(f"\nExamples of PRODUCT_NAME mapping to multiple MATERIAL DESCRIPTION:")
conflicts = pn_to_md[pn_to_md > 1].index
for pn in conflicts[:20]:
    mds = df[df['PRODUCT_NAME'] == pn]['MATERIAL DESCRIPTION'].unique()
    print(f"  {pn}")
    for md in mds:
        count = len(df[(df['PRODUCT_NAME'] == pn) & (df['MATERIAL DESCRIPTION'] == md)])
        print(f"    -> {md} ({count}x)")

# Also check reverse: MATERIAL DESCRIPTION -> PRODUCT_NAME
md_to_pn = df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
multi_pn = (md_to_pn > 1).sum()
print(f"\nMATERIAL DESCRIPTION with multiple PRODUCT_NAME: {multi_pn:,} / {len(md_to_pn):,} = {multi_pn/len(md_to_pn)*100:.2f}%")

# Distribution
pn_per_md_dist = md_to_pn.value_counts().sort_index()
print(f"\nDistribution of PRODUCT_NAME count per MATERIAL DESCRIPTION:")
for cnt, num_md in pn_per_md_dist.head(30).items():
    print(f"  {cnt} PRODUCT_NAME: {num_md:,} MATERIAL DESCRIPTION")