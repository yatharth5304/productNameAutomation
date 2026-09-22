import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("6. BRAND / PRODUCT FAMILY ANALYSIS")
print("=" * 60)

# Look for recurring prefixes/brands in MATERIAL DESCRIPTION
# Extract first word as potential brand
md_series = df['MATERIAL DESCRIPTION']
first_words = md_series.str.split().str[0]
brand_counts = first_words.value_counts()
print(f"Top 30 first words in MATERIAL DESCRIPTION (potential brands):")
for brand, count in brand_counts.head(30).items():
    unique_mds = df[df['MATERIAL DESCRIPTION'].str.startswith(brand)]['MATERIAL DESCRIPTION'].nunique()
    print(f"  {brand:20s}: {count:6d} rows, {unique_mds:3d} unique MATERIALS")

# Also check PRODUCT_NAME first words
pn_first_words = df['PRODUCT_NAME'].str.split().str[0]
pn_brand_counts = pn_first_words.value_counts()
print(f"\nTop 30 first words in PRODUCT_NAME:")
for brand, count in pn_brand_counts.head(30).items():
    print(f"  {brand:20s}: {count:6d} rows")

# Look for product families - materials sharing same prefix
print(f"\n{'='*60}")
print("PRODUCT FAMILY EXAMPLES (same brand, different strengths/forms)")
print(f"{'='*60}")

for brand in ['CARDACE', 'AMARYL', 'CORDARONE', 'LASIX', 'ENHEMO', 'EMDYDRO', 'DYDROEVA', 'CALONAT']:
    brand_mds = df[df['MATERIAL DESCRIPTION'].str.startswith(brand)]['MATERIAL DESCRIPTION'].unique()
    print(f"\n{brand} family ({len(brand_mds)} materials):")
    for md in sorted(brand_mds)[:15]:
        variant_count = df[df['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].nunique()
        total_rows = len(df[df['MATERIAL DESCRIPTION'] == md])
        print(f"  [{variant_count:3d} variants, {total_rows:4d} rows] {md}")

# Show OCR variants mapping to same material for a brand family
print(f"\n{'='*60}")
print("OCR VARIANTS -> SAME MATERIAL (brand family example)")
print(f"{'='*60}")

# Pick CARDACE family
cardace_mds = df[df['MATERIAL DESCRIPTION'].str.startswith('CARDACE')]['MATERIAL DESCRIPTION'].unique()
for md in sorted(cardace_mds)[:5]:
    variants = df[df['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"\nMATERIAL: {md}")
    print(f"  {len(variants)} unique OCR variants")
    for v in sorted(variants)[:15]:
        print(f"    {v}")

# Check if same brand appears in PRODUCT_NAME consistently
print(f"\n{'='*60}")
print("BRAND CONSISTENCY: PRODUCT_NAME vs MATERIAL DESCRIPTION")
print(f"{'='*60}")

# For each MATERIAL DESCRIPTION, check if its brand appears in PRODUCT_NAME variants
brand_consistency = []
for md in df['MATERIAL DESCRIPTION'].unique()[:100]:
    md_brand = md.split()[0] if md else ''
    variants = df[df['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    matches = sum(1 for v in variants if md_brand.upper() in v.upper())
    brand_consistency.append((md, len(variants), matches, matches/len(variants) if variants else 0))

brand_consistency.sort(key=lambda x: x[3])
print(f"Lowest brand consistency (brand in MD not found in PRODUCT_NAME):")
for md, total, matches, ratio in brand_consistency[:20]:
    print(f"  {ratio:.2f} ({matches}/{total}) : {md}")

print(f"\nHighest brand consistency:")
for md, total, matches, ratio in brand_consistency[-20:]:
    print(f"  {ratio:.2f} ({matches}/{total}) : {md}")