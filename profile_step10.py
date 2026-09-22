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
print("10. FIXED EVALUATION SET CONTAMINATION CHECK")
print("=" * 60)

# Evaluation set has 364 rows, 170 unique Input_Column, 103 unique Positive_Product
print(f"Evaluation set: {len(eval_set)} rows, {eval_set['Input_Column'].nunique()} unique inputs, {eval_set['Positive_Product'].nunique()} unique positives")

# Check 1: Exact Input_Column overlap
eval_inputs = set(eval_set['Input_Column'].unique())
new_pn = set(df['PRODUCT_NAME'].unique())
exact_input_overlap = eval_inputs & new_pn
print(f"\n1. Exact PRODUCT_NAME (Input_Column) overlap: {len(exact_input_overlap)} / {len(eval_inputs)} = {len(exact_input_overlap)/len(eval_inputs)*100:.2f}%")

# Check 2: Input_Column + Positive_Product pair overlap
eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
new_pairs = set(zip(df['PRODUCT_NAME'], df['MATERIAL DESCRIPTION']))
pair_overlap = eval_pairs & new_pairs
print(f"2. Exact (Input_Column, Positive_Product) pair overlap: {len(pair_overlap)} / {len(eval_pairs)} = {len(pair_overlap)/len(eval_pairs)*100:.2f}%")

# Check 3: For each evaluation Input_Column, does it appear in new dataset with DIFFERENT Positive_Product?
# This would be a conflict/contamination
print(f"\n3. CONFLICT CHECK: Evaluation inputs appearing in new dataset with different target")
conflicts = []
for _, row in eval_set.iterrows():
    inp = row['Input_Column']
    pos = row['Positive_Product']
    if inp in new_pn:
        # Get all MATERIAL DESCRIPTION for this PRODUCT_NAME in new dataset
        new_mds = set(df[df['PRODUCT_NAME'] == inp]['MATERIAL DESCRIPTION'].unique())
        if pos not in new_mds:
            conflicts.append((inp, pos, new_mds))
        elif len(new_mds) > 1:
            conflicts.append((inp, pos, new_mds))

print(f"   Evaluation inputs found in new dataset with different/conflicting targets: {len(conflicts)}")
for inp, pos, new_mds in conflicts[:20]:
    print(f"   EVAL: {inp} -> {pos}")
    print(f"   NEW:  {inp} -> {new_mds}")

# Check 4: Evaluation Positive_Product appearing in new dataset with different Input_Column
# This is less concerning but worth noting
print(f"\n4. Evaluation Positive_Product in new dataset:")
eval_positives = set(eval_set['Positive_Product'].unique())
new_md = set(df['MATERIAL DESCRIPTION'].unique())
pos_overlap = eval_positives & new_md
print(f"   Overlap: {len(pos_overlap)} / {len(eval_positives)} = {len(pos_overlap)/len(eval_positives)*100:.2f}%")

# For overlapping positives, show the Input_Column variations in new dataset
print(f"\n   For overlapping Positive_Product, number of PRODUCT_NAME variants in new dataset:")
for pos in sorted(pos_overlap)[:20]:
    variants = df[df['MATERIAL DESCRIPTION'] == pos]['PRODUCT_NAME'].nunique()
    print(f"     {pos}: {variants} variants")

# Check 5: Evaluation groups affected
# Each evaluation row is a test case; check if any are "contaminated"
print(f"\n5. EVALUATION ROW CONTAMINATION SUMMARY")
contaminated_rows = 0
for _, row in eval_set.iterrows():
    inp = row['Input_Column']
    pos = row['Positive_Product']
    if inp in new_pn:
        new_mds = set(df[df['PRODUCT_NAME'] == inp]['MATERIAL DESCRIPTION'].unique())
        if pos not in new_mds or len(new_mds) > 1:
            contaminated_rows += 1

print(f"   Evaluation rows with Input_Column in new dataset: {eval_set['Input_Column'].isin(new_pn).sum()} / {len(eval_set)}")
print(f"   Evaluation rows with conflicting target in new dataset: {contaminated_rows}")

# Check 6: What about the hard_negative_mining.csv (1425 examples)?
print(f"\n6. HARD NEGATIVE MINING DATASET OVERLAP")
hn_inputs = set(hard_neg['Input_Column'].unique())
hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))

hn_input_overlap = hn_inputs & new_pn
hn_pair_overlap = hn_pairs & new_pairs
print(f"   Input_Column overlap: {len(hn_input_overlap)} / {len(hn_inputs)} = {len(hn_input_overlap)/len(hn_inputs)*100:.2f}%")
print(f"   Pair overlap: {len(hn_pair_overlap)} / {len(hn_pairs)} = {len(hn_pair_overlap)/len(hn_pairs)*100:.2f}%")

# Check for conflicts in hard negative data
hn_conflicts = 0
for _, row in hard_neg.iterrows():
    inp = row['Input_Column']
    pos = row['Positive_Product']
    if inp in new_pn:
        new_mds = set(df[df['PRODUCT_NAME'] == inp]['MATERIAL DESCRIPTION'].unique())
        if pos not in new_mds:
            hn_conflicts += 1
            if hn_conflicts <= 10:
                print(f"   HN CONFLICT: {inp} -> {pos} (new: {new_mds})")
print(f"   Hard negative conflicts: {hn_conflicts}")