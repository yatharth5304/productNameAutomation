import pandas as pd
import numpy as np

df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')

print("=" * 60)
print("2. PRODUCT_NAME ANALYSIS")
print("=" * 60)

total_rows = len(df)
total_product_names = df['PRODUCT_NAME'].nunique()
print(f"Total rows: {total_rows:,}")
print(f"Unique PRODUCT_NAME: {total_product_names:,} / {total_rows:,} = {total_product_names/total_rows*100:.2f}%")

# Duplicate analysis
product_name_counts = df['PRODUCT_NAME'].value_counts()
exact_duplicates = total_rows - total_product_names
print(f"Exact duplicate rows: {exact_duplicates:,} ({exact_duplicates/total_rows*100:.2f}%)")

# Frequency distribution
freq_dist = product_name_counts.value_counts().sort_index()
print(f"\nFrequency distribution of PRODUCT_NAME occurrences:")
for freq, count in freq_dist.head(20).items():
    print(f"  Appears {freq}x: {count:,} unique PRODUCT_NAME values")

# Highly repeated examples
print(f"\nTop 20 most repeated PRODUCT_NAME values:")
for name, count in product_name_counts.head(20).items():
    print(f"  {count:5d}x: {name}")

# Products appearing 2x, 3x, 5x, 10x, etc.
for threshold in [2, 3, 5, 10, 20, 50, 100]:
    count = (product_name_counts >= threshold).sum()
    print(f"PRODUCT_NAME appearing >={threshold}x: {count:,}")

# Length analysis
df['prod_name_len'] = df['PRODUCT_NAME'].str.len()
print(f"\nPRODUCT_NAME length stats:")
print(f"  Min: {df['prod_name_len'].min()}")
print(f"  Max: {df['prod_name_len'].max()}")
print(f"  Mean: {df['prod_name_len'].mean():.1f}")
print(f"  Median: {df['prod_name_len'].median():.1f}")

# Short and long examples
short = df[df['prod_name_len'] <= 5]
long = df[df['prod_name_len'] >= 100]
print(f"  <=5 chars: {len(short):,} rows")
if len(short) > 0:
    print(f"  Examples: {short['PRODUCT_NAME'].unique()[:10].tolist()}")
print(f"  >=100 chars: {len(long):,} rows")
if len(long) > 0:
    print(f"  Examples: {long['PRODUCT_NAME'].unique()[:5].tolist()}")