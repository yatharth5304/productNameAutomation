import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("8. PARTY_CODE ANALYSIS")
print("=" * 60)

unique_parties = df['PARTY_CODE'].nunique()
print(f"Unique PARTY_CODE values: {unique_parties:,}")

# Party distribution
party_counts = df['PARTY_CODE'].value_counts()
print(f"\nTop 20 PARTY_CODE by row count:")
for party, count in party_counts.head(20).items():
    unique_pn = df[df['PARTY_CODE'] == party]['PRODUCT_NAME'].nunique()
    unique_md = df[df['PARTY_CODE'] == party]['MATERIAL DESCRIPTION'].nunique()
    print(f"  {party}: {count:6d} rows, {unique_pn:5d} unique PRODUCT_NAME, {unique_md:4d} unique MATERIAL DESCRIPTION")

# Check if same PRODUCT_NAME maps to different MATERIAL DESCRIPTION across parties
print(f"\n{'='*60}")
print("CROSS-PARTY MAPPING CONSISTENCY")
print(f"{'='*60}")

# For each PRODUCT_NAME, check if it maps to different MATERIAL DESCRIPTION across parties
pn_party_md = df.groupby(['PRODUCT_NAME', 'PARTY_CODE'])['MATERIAL DESCRIPTION'].nunique()
# Actually, let's check: for each PRODUCT_NAME, how many distinct MATERIAL DESCRIPTION across all parties
pn_md_across_parties = df.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
# But we already know each PRODUCT_NAME maps to exactly 1 MATERIAL DESCRIPTION overall
# The question is: does the same PRODUCT_NAME + PARTY_CODE combo behave differently?

# Check if PRODUCT_NAME + PARTY_CODE is more deterministic than PRODUCT_NAME alone
# (We already know PRODUCT_NAME alone is 100% deterministic)
# So let's check: for a given PRODUCT_NAME, does it appear with multiple PARTY_CODEs?
pn_party_count = df.groupby('PRODUCT_NAME')['PARTY_CODE'].nunique()
print(f"PRODUCT_NAME appearing with multiple PARTY_CODE: {(pn_party_count > 1).sum():,} / {len(pn_party_count):,}")

# For PRODUCT_NAME that appear with multiple parties, check if MATERIAL DESCRIPTION differs
multi_party_pns = pn_party_count[pn_party_count > 1].index
conflicts = 0
for pn in multi_party_pns[:1000]:  # Sample first 1000
    md_by_party = df[df['PRODUCT_NAME'] == pn].groupby('PARTY_CODE')['MATERIAL DESCRIPTION'].first()
    if md_by_party.nunique() > 1:
        conflicts += 1
        if conflicts <= 10:
            print(f"  CONFLICT: {pn}")
            for party, md in md_by_party.items():
                print(f"    PARTY {party}: {md}")

print(f"\nConflicts found in sample: {conflicts}")

# Check the reverse: for a given MATERIAL DESCRIPTION, does PARTY_CODE affect which PRODUCT_NAME maps to it?
# This is less relevant since we already know multiple PRODUCT_NAME map to same MATERIAL DESCRIPTION

# Check if certain PARTY_CODEs have systematically different mappings
print(f"\n{'='*60}")
print("PARTY_CODE SPECIFIC MAPPING PATTERNS")
print(f"{'='*60}")

# For each PARTY_CODE, what's the overlap with other parties?
# Sample a few parties
for party in party_counts.head(10).index:
    party_data = df[df['PARTY_CODE'] == party]
    pn_set = set(party_data['PRODUCT_NAME'].unique())
    md_set = set(party_data['MATERIAL DESCRIPTION'].unique())
    
    # Check if this party's PRODUCT_NAMEs map to same MDs globally
    global_md_for_pn = df[df['PRODUCT_NAME'].isin(pn_set)].groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].first()
    party_md_for_pn = party_data.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].first()
    
    mismatches = 0
    for pn in party_md_for_pn.index:
        if pn in global_md_for_pn.index:
            if party_md_for_pn[pn] != global_md_for_pn[pn]:
                mismatches += 1
    
    print(f"  PARTY {party}: {len(pn_set)} PRODUCT_NAME, {len(md_set)} MATERIAL DESCRIPTION, {mismatches} mismatches vs global")

# Check if there are PARTY_CODE specific PRODUCT_NAME that don't appear elsewhere
print(f"\n{'='*60}")
print("PARTY-EXCLUSIVE PRODUCT_NAME")
print(f"{'='*60}")

# Find PRODUCT_NAME that only appear for one PARTY_CODE
pn_party_map = df.groupby('PRODUCT_NAME')['PARTY_CODE'].apply(set)
exclusive_pn = {pn: parties for pn, parties in pn_party_map.items() if len(parties) == 1}
print(f"PRODUCT_NAME exclusive to single PARTY_CODE: {len(exclusive_pn):,} / {len(pn_party_map):,}")

# Distribution of exclusive PN per party
party_exclusive_counts = {}
for pn, parties in exclusive_pn.items():
    party = list(parties)[0]
    party_exclusive_counts[party] = party_exclusive_counts.get(party, 0) + 1

print(f"\nTop 10 parties by exclusive PRODUCT_NAME count:")
for party, count in sorted(party_exclusive_counts.items(), key=lambda x: -x[1])[:10]:
    total_pn = party_counts[party] if party in party_counts.index else 0
    print(f"  PARTY {party}: {count:,} exclusive PRODUCT_NAME")