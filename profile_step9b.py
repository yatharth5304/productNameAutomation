import pandas as pd
import numpy as np

# Load all data
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
training_data = pd.read_excel('training_data.xlsx')
hard_neg = pd.read_csv('hard_negative_mining.csv')
train_split = pd.read_csv('training_split.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 60)
print("9. OVERLAP WITH EXISTING TRAINING DATA")
print("=" * 60)

# New dataset PRODUCT_NAME and MATERIAL DESCRIPTION
new_pn = set(df['PRODUCT_NAME'].unique())
new_md = set(df['MATERIAL DESCRIPTION'].unique())

# Evaluation set: Input_Column and Positive_Product
eval_inputs = set(eval_set['Input_Column'].unique())
eval_positives = set(eval_set['Positive_Product'].unique())

# Training data: Input_Column and Output_Column
train_inputs = set(training_data['Input_Column'].unique())
train_outputs = set(training_data['Output_Column'].unique())

# Hard negatives: Input_Column and Positive_Product
hn_inputs = set(hard_neg['Input_Column'].unique())
hn_positives = set(hard_neg['Positive_Product'].unique())

# Training split: Input_Column and Positive_Product
ts_inputs = set(train_split['Input_Column'].unique())
ts_positives = set(train_split['Positive_Product'].unique())

# PRODUCT_MASTER product_name
master_products = set(product_master['product_name'].unique())

print(f"\n--- PRODUCT_NAME (Input_Column) OVERLAP ---")
print(f"New dataset unique PRODUCT_NAME: {len(new_pn):,}")

for name, existing_set in [
    ("evaluation_set Input_Column", eval_inputs),
    ("training_data Input_Column", train_inputs),
    ("hard_negative_mining Input_Column", hn_inputs),
    ("training_split Input_Column", ts_inputs),
]:
    overlap = new_pn & existing_set
    print(f"  Overlap with {name}: {len(overlap):,} / {len(existing_set):,} = {len(overlap)/len(existing_set)*100:.2f}%")
    if len(overlap) > 0 and len(overlap) <= 20:
        print(f"    Examples: {list(overlap)[:20]}")

print(f"\n--- MATERIAL DESCRIPTION / Positive_Product OVERLAP ---")
print(f"New dataset unique MATERIAL DESCRIPTION: {len(new_md):,}")

for name, existing_set in [
    ("evaluation_set Positive_Product", eval_positives),
    ("training_data Output_Column", train_outputs),
    ("hard_negative_mining Positive_Product", hn_positives),
    ("training_split Positive_Product", ts_positives),
    ("PRODUCT_MASTER product_name", master_products),
]:
    overlap = new_md & existing_set
    print(f"  Overlap with {name}: {len(overlap):,} / {len(existing_set):,} = {len(overlap)/len(existing_set)*100:.2f}%")
    if len(overlap) > 0 and len(overlap) <= 20:
        print(f"    Examples: {list(overlap)[:20]}")

# Check full pair overlap (Input_Column -> Positive_Product)
print(f"\n--- FULL MAPPING PAIR OVERLAP ---")

# Build pair sets
new_pairs = set(zip(df['PRODUCT_NAME'], df['MATERIAL DESCRIPTION']))
eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
train_pairs = set(zip(training_data['Input_Column'], training_data['Output_Column']))
hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
ts_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product']))

for name, existing_pairs in [
    ("evaluation_set", eval_pairs),
    ("training_data", train_pairs),
    ("hard_negative_mining", hn_pairs),
    ("training_split", ts_pairs),
]:
    overlap = new_pairs & existing_pairs
    print(f"  Pair overlap with {name}: {len(overlap):,} / {len(existing_pairs):,} = {len(overlap)/len(existing_pairs)*100:.2f}%")
    if len(overlap) > 0 and len(overlap) <= 10:
        for p in list(overlap)[:10]:
            print(f"    {p[0]} -> {p[1]}")