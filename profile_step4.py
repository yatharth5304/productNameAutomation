import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("4. MATERIAL DESCRIPTION ANALYSIS")
print("=" * 60)

unique_md = df['MATERIAL DESCRIPTION'].nunique()
print(f"Unique MATERIAL DESCRIPTION: {unique_md:,}")

# Examples per material
md_counts = df.groupby('MATERIAL DESCRIPTION').size().sort_values(ascending=False)
print(f"\nTraining examples per MATERIAL DESCRIPTION:")
print(f"  Mean: {md_counts.mean():.1f}")
print(f"  Median: {md_counts.median():.1f}")
print(f"  Min: {md_counts.min()}")
print(f"  Max: {md_counts.max()}")

# Percentiles
for p in [10, 25, 50, 75, 90, 95, 99]:
    print(f"  P{p}: {md_counts.quantile(p/100):.1f}")

# Products with highest number of OCR variants (unique PRODUCT_NAME per MATERIAL DESCRIPTION)
md_to_pn_unique = df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique().sort_values(ascending=False)
print(f"\nTop 30 MATERIAL DESCRIPTION by unique PRODUCT_NAME variants:")
for md, cnt in md_to_pn_unique.head(30).items():
    total_examples = md_counts[md]
    print(f"  {cnt:3d} variants, {total_examples:5d} total rows: {md}")

# Also show by total rows
print(f"\nTop 30 MATERIAL DESCRIPTION by total training rows:")
for md, cnt in md_counts.head(30).items():
    unique_variants = md_to_pn_unique[md]
    print(f"  {cnt:5d} rows, {unique_variants:3d} variants: {md}")