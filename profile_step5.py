import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("5. OCR VARIATION ANALYSIS")
print("=" * 60)

# For each MATERIAL DESCRIPTION, look at the PRODUCT_NAME variants
md_groups = df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].apply(list)

# Find materials with many variants
md_variant_counts = df.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique().sort_values(ascending=False)

print(f"Materials with >=100 unique PRODUCT_NAME variants: {(md_variant_counts >= 100).sum()}")
print(f"Materials with >=50 unique PRODUCT_NAME variants: {(md_variant_counts >= 50).sum()}")
print(f"Materials with >=20 unique PRODUCT_NAME variants: {(md_variant_counts >= 20).sum()}")
print(f"Materials with >=10 unique PRODUCT_NAME variants: {(md_variant_counts >= 10).sum()}")
print(f"Materials with >=5 unique PRODUCT_NAME variants: {(md_variant_counts >= 5).sum()}")

# Show representative examples of OCR variation
print(f"\n{'='*60}")
print("REPRESENTATIVE OCR VARIATION EXAMPLES")
print(f"{'='*60}")

# Pick a few materials with many variants
for md in md_variant_counts.head(10).index:
    variants = df[df['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"\nMATERIAL DESCRIPTION: {md}")
    print(f"  Unique PRODUCT_NAME variants: {len(variants)}")
    # Show first 20 variants
    for v in variants[:20]:
        count = len(df[(df['MATERIAL DESCRIPTION'] == md) & (df['PRODUCT_NAME'] == v)])
        print(f"    ({count:3d}x) {v}")

# Analyze specific variation patterns
print(f"\n{'='*60}")
print("VARIATION PATTERN ANALYSIS")
print(f"{'='*60}")

# Look for common variation types
def analyze_variations(variants, target):
    """Analyze variations against a target canonical form"""
    variations = []
    for v in variants:
        if v != target:
            variations.append(v)
    return variations

# Check specific materials for variation types
sample_mds = [
    'CALONAT XT TABLET 10x15T',
    'DYDROEVA TABLET 1X10T',
    'EMDYDRO TABLET 1X10T',
    'ENHEMO 500 INJECTION 5x5ML',
    'LASIX 40 MG TABLET 80x15T',
    'AMARYL 1MG TABLET 10x30T',
]

for md in sample_mds:
    if md in df['MATERIAL DESCRIPTION'].values:
        variants = df[df['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
        print(f"\n--- {md} ---")
        print(f"Total unique variants: {len(variants)}")
        for v in sorted(variants)[:30]:
            print(f"  {v}")

# Look for specific patterns across the dataset
print(f"\n{'='*60}")
print("COMMON VARIATION PATTERNS (sampling)")
print(f"{'='*60}")

# Check for hyphen differences
hyphen_variations = 0
space_variations = 0
case_variations = 0

# Sample some groups
for md in md_variant_counts[md_variant_counts >= 20].index[:50]:
    variants = df[df['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    for v in variants:
        # Compare against MATERIAL DESCRIPTION (normalized)
        md_norm = md.replace('-', ' ').replace('  ', ' ').strip().upper()
        v_norm = v.replace('-', ' ').replace('  ', ' ').strip().upper()
        if md_norm != v_norm:
            if '-' in v and '-' not in md:
                hyphen_variations += 1
            if '  ' in v or v != v.strip():
                space_variations += 1
            if v.upper() != v and v.lower() != v:
                case_variations += 1

print(f"Note: Full pattern analysis requires more detailed comparison logic")
print(f"Above shows representative samples of variation types")