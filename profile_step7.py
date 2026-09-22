import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("7. MATERIAL CODE CONSISTENCY")
print("=" * 60)

# MATERIAL CODE -> MATERIAL DESCRIPTION
mc_to_md = df.groupby('MATERIAL CODE')['MATERIAL DESCRIPTION'].nunique()
print(f"Unique MATERIAL CODE: {df['MATERIAL CODE'].nunique():,}")
print(f"MATERIAL CODE with 1 MATERIAL DESCRIPTION: {(mc_to_md == 1).sum():,}")
print(f"MATERIAL CODE with >1 MATERIAL DESCRIPTION: {(mc_to_md > 1).sum():,}")

if (mc_to_md > 1).sum() > 0:
    print(f"\nMATERIAL CODEs mapping to multiple MATERIAL DESCRIPTION:")
    for mc in mc_to_md[mc_to_md > 1].index[:20]:
        mds = df[df['MATERIAL CODE'] == mc]['MATERIAL DESCRIPTION'].unique()
        print(f"  {mc}: {mds}")

# MATERIAL DESCRIPTION -> MATERIAL CODE
md_to_mc = df.groupby('MATERIAL DESCRIPTION')['MATERIAL CODE'].nunique()
print(f"\nMATERIAL DESCRIPTION with 1 MATERIAL CODE: {(md_to_mc == 1).sum():,}")
print(f"MATERIAL DESCRIPTION with >1 MATERIAL CODE: {(md_to_mc > 1).sum():,}")

if (md_to_mc > 1).sum() > 0:
    print(f"\nMATERIAL DESCRIPTIONs mapping to multiple MATERIAL CODE:")
    for md in md_to_mc[md_to_mc > 1].index[:20]:
        mcs = df[df['MATERIAL DESCRIPTION'] == md]['MATERIAL CODE'].unique()
        print(f"  {md}: {mcs}")

# Check the "-" MATERIAL DESCRIPTION
print(f"\n{'='*60}")
print("SPECIAL CASE: MATERIAL DESCRIPTION = '-'")
print(f"{'='*60}")
dash_rows = df[df['MATERIAL DESCRIPTION'] == '-']
print(f"Rows with '-' MATERIAL DESCRIPTION: {len(dash_rows):,}")
print(f"Unique MATERIAL CODE for '-': {dash_rows['MATERIAL CODE'].nunique()}")
print(f"Unique PRODUCT_NAME for '-': {dash_rows['PRODUCT_NAME'].nunique()}")
print(f"Unique PARTY_CODE for '-': {dash_rows['PARTY_CODE'].nunique()}")
print(f"MATERIAL CODEs: {dash_rows['MATERIAL CODE'].unique()[:20]}")

# Sample of PRODUCT_NAME for "-"
print(f"\nSample PRODUCT_NAME for '-':")
for pn in dash_rows['PRODUCT_NAME'].unique()[:30]:
    print(f"  {pn}")

# Check if MATERIAL CODE is consistent for non-dash materials
print(f"\n{'='*60}")
print("MATERIAL CODE SAMPLE (non-dash)")
print(f"{'='*60}")
non_dash = df[df['MATERIAL DESCRIPTION'] != '-']
mc_md_pairs = non_dash[['MATERIAL CODE', 'MATERIAL DESCRIPTION']].drop_duplicates()
print(f"Unique MATERIAL CODE -> MATERIAL DESCRIPTION pairs (non-dash): {len(mc_md_pairs):,}")

# Show some examples
for mc in non_dash['MATERIAL CODE'].unique()[:10]:
    mds = non_dash[non_dash['MATERIAL CODE'] == mc]['MATERIAL DESCRIPTION'].unique()
    pns = non_dash[non_dash['MATERIAL CODE'] == mc]['PRODUCT_NAME'].nunique()
    print(f"  {mc} -> {mds} ({pns} PRODUCT_NAME variants)")