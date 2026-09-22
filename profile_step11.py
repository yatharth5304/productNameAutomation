import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("11. DATASET QUALITY PROBLEMS")
print("=" * 60)

print(f"Total rows: {len(df):,}")
print(f"Columns: {list(df.columns)}")

# 1. Blank/NULL checks
print(f"\n--- NULL/BLANK CHECKS ---")
for col in df.columns:
    null_count = df[col].isnull().sum()
    blank_count = (df[col].astype(str).str.strip() == '').sum() if df[col].dtype == 'object' else 0
    print(f"  {col}: NULL={null_count}, BLANK={blank_count}")

# 2. Suspicious duplicate records (exact row duplicates)
exact_dupes = df.duplicated().sum()
print(f"\n--- EXACT ROW DUPLICATES ---")
print(f"  Exact duplicate rows: {exact_dupes:,}")

# 3. Conflicting mappings (already checked - 0)
print(f"\n--- CONFLICTING MAPPINGS ---")
pn_to_md = df.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
conflicts = (pn_to_md > 1).sum()
print(f"  PRODUCT_NAME -> multiple MATERIAL DESCRIPTION: {conflicts}")

# 4. Malformed values - unusual characters
print(f"\n--- UNUSUAL CHARACTERS IN PRODUCT_NAME ---")
# Check for non-ASCII, control chars, etc.
def has_unusual_chars(s):
    if pd.isna(s):
        return False
    s = str(s)
    # Check for non-printable chars, excessive special chars
    for ch in s:
        if ord(ch) < 32 or ord(ch) > 126:
            return True
    return False

unusual = df['PRODUCT_NAME'].apply(has_unusual_chars).sum()
print(f"  PRODUCT_NAME with non-ASCII/control chars: {unusual:,}")

# Check for specific problematic patterns
patterns = {
    'multiple_spaces': df['PRODUCT_NAME'].str.contains(r'  +').sum(),
    'leading_trailing_space': (df['PRODUCT_NAME'].str.startswith(' ') | df['PRODUCT_NAME'].str.endswith(' ')).sum(),
    'tabs': df['PRODUCT_NAME'].str.contains(r'\t').sum(),
    'newlines': df['PRODUCT_NAME'].str.contains(r'\n').sum(),
    'multiple_periods': df['PRODUCT_NAME'].str.contains(r'\.{2,}').sum(),
    'multiple_hyphens': df['PRODUCT_NAME'].str.contains(r'-{2,}').sum(),
    'only_numbers': df['PRODUCT_NAME'].str.match(r'^[\d\s\.\-]+$').sum(),
    'very_short': (df['PRODUCT_NAME'].str.len() <= 3).sum(),
    'very_long': (df['PRODUCT_NAME'].str.len() > 80).sum(),
}
for name, count in patterns.items():
    print(f"  {name}: {count:,}")

# 5. Material Description issues
print(f"\n--- MATERIAL DESCRIPTION ISSUES ---")
# Check the "-" case
dash_count = (df['MATERIAL DESCRIPTION'] == '-').sum()
print(f"  MATERIAL DESCRIPTION == '-': {dash_count:,} rows")

# Check for very short material descriptions
md_len = df['MATERIAL DESCRIPTION'].str.len()
short_md = (md_len <= 3).sum()
print(f"  MATERIAL DESCRIPTION <= 3 chars: {short_md:,}")

# 6. MATERIAL CODE issues
print(f"\n--- MATERIAL CODE ISSUES ---")
mc_dash = (df['MATERIAL CODE'] == 'NO PRODUCT').sum()
print(f"  MATERIAL CODE == 'NO PRODUCT': {mc_dash:,} rows")

# Check for non-numeric MATERIAL CODE
mc_numeric = df['MATERIAL CODE'].apply(lambda x: str(x).isdigit() if pd.notna(x) else False)
non_numeric_mc = (~mc_numeric).sum()
print(f"  Non-numeric MATERIAL CODE: {non_numeric_mc:,}")

# 7. PARTY_CODE issues
print(f"\n--- PARTY_CODE ISSUES ---")
party_stats = df['PARTY_CODE'].describe()
print(f"  Min: {party_stats['min']}, Max: {party_stats['max']}")
print(f"  Mean: {party_stats['mean']:.0f}")

# 8. Suspicious records - PRODUCT_NAME == MATERIAL DESCRIPTION
same = (df['PRODUCT_NAME'] == df['MATERIAL DESCRIPTION']).sum()
print(f"\n--- PRODUCT_NAME == MATERIAL DESCRIPTION ---")
print(f"  Identical: {same:,}")

# 9. Show examples of suspicious records
print(f"\n--- SUSPICIOUS RECORD EXAMPLES ---")

# Very short PRODUCT_NAME
short_pn = df[df['PRODUCT_NAME'].str.len() <= 5]
if len(short_pn) > 0:
    print(f"  Short PRODUCT_NAME (<=5 chars):")
    for _, row in short_pn.head(20).iterrows():
        print(f"    '{row['PRODUCT_NAME']}' -> '{row['MATERIAL DESCRIPTION']}'")

# PRODUCT_NAME with only numbers/symbols
num_only = df[df['PRODUCT_NAME'].str.match(r'^[\d\s\.\-]+$')]
if len(num_only) > 0:
    print(f"  Numeric-only PRODUCT_NAME:")
    for _, row in num_only.head(10).iterrows():
        print(f"    '{row['PRODUCT_NAME']}' -> '{row['MATERIAL DESCRIPTION']}'")

# MATERIAL DESCRIPTION "-"
dash_rows = df[df['MATERIAL DESCRIPTION'] == '-']
if len(dash_rows) > 0:
    print(f"  MATERIAL DESCRIPTION '-': (showing first 10)")
    for _, row in dash_rows.head(10).iterrows():
        print(f"    '{row['PRODUCT_NAME']}' -> '{row['MATERIAL DESCRIPTION']}' (CODE: {row['MATERIAL CODE']})")

# 10. Duplicate PRODUCT_NAME + MATERIAL DESCRIPTION + PARTY_CODE combinations
trip_dupes = df.duplicated(subset=['PRODUCT_NAME', 'MATERIAL DESCRIPTION', 'PARTY_CODE']).sum()
print(f"\n  Duplicate (PRODUCT_NAME, MATERIAL DESCRIPTION, PARTY_CODE): {trip_dupes:,}")

# 11. Records where PRODUCT_NAME contains MATERIAL DESCRIPTION (or vice versa) suspiciously
# This might indicate data entry errors
print(f"\n--- POTENTIAL DATA ENTRY ERRORS ---")
# PRODUCT_NAME much longer than MATERIAL DESCRIPTION
df['len_diff'] = df['PRODUCT_NAME'].str.len() - df['MATERIAL DESCRIPTION'].str.len()
very_long_pn = df[df['len_diff'] > 50]
print(f"  PRODUCT_NAME >50 chars longer than MATERIAL DESCRIPTION: {len(very_long_pn):,}")
if len(very_long_pn) > 0:
    for _, row in very_long_pn.head(5).iterrows():
        print(f"    PN({len(row['PRODUCT_NAME'])}): {row['PRODUCT_NAME']}")
        print(f"    MD({len(row['MATERIAL DESCRIPTION'])}): {row['MATERIAL DESCRIPTION']}")

# 12. Check for obvious non-product rows in MATERIAL DESCRIPTION
# The "-" is already identified. Any others?
md_values = df['MATERIAL DESCRIPTION'].value_counts()
print(f"\n--- MATERIAL DESCRIPTION VALUE COUNTS (bottom 20) ---")
for md, cnt in md_values.tail(20).items():
    print(f"  {cnt:4d}: {md}")