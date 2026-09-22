import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 60)
print("ADDITIONAL ANALYSIS FOR TRAINING VALUE ASSESSMENT")
print("=" * 60)

# 1. How many PRODUCT_MASTER products are covered?
master_products = set(product_master['product_name'].unique())
new_md = set(df['MATERIAL DESCRIPTION'].unique())
covered = master_products & new_md
print(f"PRODUCT_MASTER products covered in new dataset: {len(covered)} / {len(master_products)} = {len(covered)/len(master_products)*100:.1f}%")

# 2. For evaluation set products, how many OCR variants in new dataset?
eval_products = set(eval_set['Positive_Product'].unique())
print(f"\nEvaluation set products: {len(eval_products)}")
for prod in sorted(eval_products):
    if prod in new_md:
        variants = df[df['MATERIAL DESCRIPTION'] == prod]['PRODUCT_NAME'].nunique()
        total_rows = len(df[df['MATERIAL DESCRIPTION'] == prod])
        print(f"  {prod}: {variants} variants, {total_rows} rows")
    else:
        print(f"  {prod}: NOT IN NEW DATASET")

# 3. Distribution of examples per material (for training balance)
md_counts = df[df['MATERIAL DESCRIPTION'] != '-'].groupby('MATERIAL DESCRIPTION').size()
print(f"\nTraining examples per material (excluding '-'):")
print(f"  Count: {len(md_counts)}")
print(f"  Mean: {md_counts.mean():.1f}")
print(f"  Median: {md_counts.median():.1f}")
print(f"  Min: {md_counts.min()}")
print(f"  Max: {md_counts.max()}")
print(f"  Std: {md_counts.std():.1f}")

# Materials with very few examples
few = md_counts[md_counts <= 5]
print(f"\nMaterials with <=5 examples: {len(few)} / {len(md_counts)} = {len(few)/len(md_counts)*100:.1f}%")
if len(few) > 0:
    print(f"  Examples: {few.head(10).to_dict()}")

# Materials with many examples
many = md_counts[md_counts >= 500]
print(f"\nMaterials with >=500 examples: {len(many)} / {len(md_counts)} = {len(many)/len(md_counts)*100:.1f}%")

# 4. OCR variation richness - how many materials have >=10 variants?
md_variants = df[df['MATERIAL DESCRIPTION'] != '-'].groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
rich = (md_variants >= 10).sum()
print(f"\nMaterials with >=10 OCR variants: {rich} / {len(md_variants)} = {rich/len(md_variants)*100:.1f}%")
rich20 = (md_variants >= 20).sum()
print(f"Materials with >=20 OCR variants: {rich20} / {len(md_variants)} = {rich20/len(md_variants)*100:.1f}%")
rich50 = (md_variants >= 50).sum()
print(f"Materials with >=50 OCR variants: {rich50} / {len(md_variants)} = {rich50/len(md_variants)*100:.1f}%")
rich100 = (md_variants >= 100).sum()
print(f"Materials with >=100 OCR variants: {rich100} / {len(md_variants)} = {rich100/len(md_variants)*100:.1f}%")

# 5. Brand coverage
# Check which brands from PRODUCT_MASTER appear in new dataset
master_brands = set(product_master['BRAND_NAME'].dropna().unique())
# Extract brands from MATERIAL DESCRIPTION
new_brands = set(df['MATERIAL DESCRIPTION'].str.split().str[0].unique())
brand_overlap = master_brands & new_brands
print(f"\nBrand overlap with PRODUCT_MASTER: {len(brand_overlap)} / {len(master_brands)} = {len(brand_overlap)/len(master_brands)*100:.1f}%")
print(f"  Common brands: {sorted(brand_overlap)[:30]}")

# 6. Party code distribution per material
# For each material, how many different parties?
md_party = df[df['MATERIAL DESCRIPTION'] != '-'].groupby('MATERIAL DESCRIPTION')['PARTY_CODE'].nunique()
print(f"\nParties per material (excluding '-'):")
print(f"  Mean: {md_party.mean():.1f}")
print(f"  Median: {md_party.median():.1f}")
print(f"  Max: {md_party.max()}")

# 7. Check if the "-" MATERIAL DESCRIPTION rows are from specific parties
dash_rows = df[df['MATERIAL DESCRIPTION'] == '-']
print(f"\n'-' MATERIAL DESCRIPTION by PARTY_CODE (top 10):")
print(dash_rows['PARTY_CODE'].value_counts().head(10))

# 8. Check if "-" rows have PRODUCT_NAME that appear elsewhere with valid MATERIAL DESCRIPTION
dash_pns = set(dash_rows['PRODUCT_NAME'].unique())
valid_rows = df[df['MATERIAL DESCRIPTION'] != '-']
valid_pns = set(valid_rows['PRODUCT_NAME'].unique())
overlap = dash_pns & valid_pns
print(f"\nPRODUCT_NAME in '-' rows that ALSO appear with valid MATERIAL DESCRIPTION: {len(overlap)} / {len(dash_pns)}")
if len(overlap) > 0:
    for pn in list(overlap)[:10]:
        valid_mds = valid_rows[valid_rows['PRODUCT_NAME'] == pn]['MATERIAL DESCRIPTION'].unique()
        print(f"  {pn}: valid mappings -> {valid_mds}")