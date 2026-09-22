import pandas as pd
import numpy as np

# Load existing data files
eval_set = pd.read_csv('evaluation_set.csv')
training_data = pd.read_excel('training_data.xlsx')
hard_neg = pd.read_csv('hard_negative_mining.csv')
train_split = pd.read_csv('training_split.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 60)
print("EXISTING DATA OVERVIEW")
print("=" * 60)

print(f"evaluation_set.csv: {len(eval_set)} rows, columns: {list(eval_set.columns)}")
print(f"training_data.xlsx: {len(training_data)} rows, columns: {list(training_data.columns)}")
print(f"hard_negative_mining.csv: {len(hard_neg)} rows, columns: {list(hard_neg.columns)}")
print(f"training_split.csv: {len(train_split)} rows, columns: {list(train_split.columns)}")
print(f"PRODUCT_MASTER.xlsx: {len(product_master)} rows, columns: {list(product_master.columns)}")

# Load new dataset
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
print(f"\nNew dataset (Sheet1): {len(df)} rows")

print("\n" + "=" * 60)
print("9. OVERLAP WITH EXISTING TRAINING DATA")
print("=" * 60)

# PRODUCT_NAME exact overlap
# Need to identify the correct column names in each file
print(f"\nColumn names check:")
print(f"  evaluation_set: {list(eval_set.columns)}")
print(f"  training_data: {list(training_data.columns)}")
print(f"  hard_negative_mining: {list(hard_neg.columns)}")
print(f"  training_split: {list(train_split.columns)}")
print(f"  PRODUCT_MASTER: {list(product_master.columns)}")