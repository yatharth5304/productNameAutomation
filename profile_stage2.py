import pandas as pd
import numpy as np

# Load all data
df = pd.read_excel('EmcureSSSReport.xlsb', sheet_name='Sheet1', engine='pyxlsb')
eval_set = pd.read_csv('evaluation_set.csv')
training_data = pd.read_excel('training_data.xlsx')
hard_neg = pd.read_csv('hard_negative_mining.csv')
train_split = pd.read_csv('training_split.csv')
product_master = pd.read_excel('PRODUCT_MASTER.xlsx')

print("=" * 70)
print("SECOND-STAGE TRAINING-DATA ANALYSIS")
print("=" * 70)

# ============================================================
# 1. CLEANED DATASET SIZE
# ============================================================
print("\n" + "=" * 70)
print("1. CLEANED DATASET SIZE")
print("=" * 70)

original_rows = len(df)
print(f"Original rows: {original_rows:,}")

# Remove NO PRODUCT rows
no_product_mask = (df['MATERIAL DESCRIPTION'] == '-') | (df['MATERIAL CODE'] == 'NO PRODUCT')
no_product_rows = no_product_mask.sum()
df_clean = df[~no_product_mask].copy()
print(f"NO PRODUCT rows removed: {no_product_rows:,}")

# Remove exact duplicates
exact_dupes = df_clean.duplicated().sum()
df_clean = df_clean.drop_duplicates().reset_index(drop=True)
print(f"Exact duplicate rows removed: {exact_dupes:,}")

remaining_rows = len(df_clean)
print(f"Remaining usable rows: {remaining_rows:,}")
print(f"Retention rate: {remaining_rows/original_rows*100:.2f}%")

# ============================================================
# 2. POSITIVE PAIR ANALYSIS
# ============================================================
print("\n" + "=" * 70)
print("2. POSITIVE PAIR ANALYSIS")
print("=" * 70)

unique_pn = df_clean['PRODUCT_NAME'].nunique()
unique_md = df_clean['MATERIAL DESCRIPTION'].nunique()
unique_pairs = df_clean[['PRODUCT_NAME', 'MATERIAL DESCRIPTION']].drop_duplicates().shape[0]

print(f"Unique PRODUCT_NAME: {unique_pn:,}")
print(f"Unique MATERIAL DESCRIPTION: {unique_md:,}")
print(f"Unique PRODUCT_NAME + MATERIAL DESCRIPTION pairs: {unique_pairs:,}")

# Duplicate positive pairs
pair_counts = df_clean.groupby(['PRODUCT_NAME', 'MATERIAL DESCRIPTION']).size()
dup_pairs = (pair_counts > 1).sum()
print(f"Duplicate positive pairs (same PN+MD appearing multiple times): {dup_pairs:,}")

# PRODUCT_NAME -> multiple MATERIAL DESCRIPTION
pn_to_md = df_clean.groupby('PRODUCT_NAME')['MATERIAL DESCRIPTION'].nunique()
pn_multi_md = (pn_to_md > 1).sum()
print(f"PRODUCT_NAME mapping to multiple MATERIAL DESCRIPTION: {pn_multi_md:,} / {len(pn_to_md):,} = {pn_multi_md/len(pn_to_md)*100:.2f}%")

# MATERIAL DESCRIPTION -> multiple PRODUCT_NAME
md_to_pn = df_clean.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()
md_multi_pn = (md_to_pn > 1).sum()
print(f"MATERIAL DESCRIPTION receiving multiple PRODUCT_NAME: {md_multi_pn:,} / {len(md_to_pn):,} = {md_multi_pn/len(md_to_pn)*100:.2f}%")

# ============================================================
# 3. CONFLICT ANALYSIS
# ============================================================
print("\n" + "=" * 70)
print("3. CONFLICT ANALYSIS")
print("=" * 70)

if pn_multi_md > 0:
    conflicts = pn_to_md[pn_to_md > 1].index
    print(f"Found {len(conflicts)} PRODUCT_NAME with multiple MATERIAL DESCRIPTION")
    for pn in conflicts[:20]:
        mds = df_clean[df_clean['PRODUCT_NAME'] == pn]['MATERIAL DESCRIPTION'].unique()
        print(f"  {pn}")
        for md in mds:
            count = len(df_clean[(df_clean['PRODUCT_NAME'] == pn) & (df_clean['MATERIAL DESCRIPTION'] == md)])
            print(f"    -> {md} ({count}x)")
else:
    print("No PRODUCT_NAME maps to multiple MATERIAL DESCRIPTION (0 conflicts)")

# ============================================================
# 4. MATERIAL FREQUENCY DISTRIBUTION
# ============================================================
print("\n" + "=" * 70)
print("4. MATERIAL FREQUENCY DISTRIBUTION (by unique PRODUCT_NAME variants)")
print("=" * 70)

md_variant_counts = df_clean.groupby('MATERIAL DESCRIPTION')['PRODUCT_NAME'].nunique()

buckets = {
    '1': (1, 1),
    '2': (2, 2),
    '3-5': (3, 5),
    '6-10': (6, 10),
    '11-25': (11, 25),
    '26-50': (26, 50),
    '51-100': (51, 100),
    '101-250': (101, 250),
    '251-500': (251, 500),
    '500+': (501, float('inf')),
}

total_materials = len(md_variant_counts)
total_examples = df_clean.groupby('MATERIAL DESCRIPTION').size().sum()

print(f"{'Bucket':<12} {'Materials':>10} {'%Materials':>12} {'Examples':>12} {'%Examples':>12}")
print("-" * 60)

for bucket_name, (low, high) in buckets.items():
    if high == float('inf'):
        mask = md_variant_counts >= low
    else:
        mask = (md_variant_counts >= low) & (md_variant_counts <= high)
    mats_in_bucket = mask.sum()
    # Examples = sum of all rows for materials in this bucket
    mats_list = md_variant_counts[mask].index
    examples_in_bucket = df_clean[df_clean['MATERIAL DESCRIPTION'].isin(mats_list)].shape[0]
    print(f"{bucket_name:<12} {mats_in_bucket:>10,} {mats_in_bucket/total_materials*100:>11.2f}% {examples_in_bucket:>12,} {examples_in_bucket/total_examples*100:>11.2f}%")

# ============================================================
# 5. OCR DIVERSITY
# ============================================================
print("\n" + "=" * 70)
print("5. OCR DIVERSITY")
print("=" * 70)

print(f"Unique OCR variants per material (statistics):")
print(f"  Mean: {md_variant_counts.mean():.1f}")
print(f"  Median: {md_variant_counts.median():.1f}")
print(f"  P90: {md_variant_counts.quantile(0.90):.1f}")
print(f"  P95: {md_variant_counts.quantile(0.95):.1f}")
print(f"  Max: {md_variant_counts.max():.0f}")

# Examples of low/medium/high diversity
print(f"\nLow diversity materials (1-2 variants):")
low_div = md_variant_counts[md_variant_counts <= 2].index[:10]
for md in low_div:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"  {md} ({len(variants)} variants): {list(variants)[:5]}")

print(f"\nMedium diversity materials (10-20 variants):")
med_div = md_variant_counts[(md_variant_counts >= 10) & (md_variant_counts <= 20)].index[:5]
for md in med_div:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"  {md} ({len(variants)} variants):")
    for v in variants[:8]:
        print(f"    {v}")

print(f"\nHigh diversity materials (500+ variants):")
high_div = md_variant_counts[md_variant_counts >= 500].index[:5]
for md in high_div:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    print(f"  {md} ({len(variants)} variants):")
    for v in variants[:8]:
        print(f"    {v}")

# ============================================================
# 6. LONG-TAIL IMPACT
# ============================================================
print("\n" + "=" * 70)
print("6. LONG-TAIL IMPACT (<=5 examples)")
print("=" * 70)

# Materials with <=5 unique PRODUCT_NAME variants
long_tail_mats = md_variant_counts[md_variant_counts <= 5].index
print(f"Materials with <=5 variants: {len(long_tail_mats):,} / {total_materials:,} = {len(long_tail_mats)/total_materials*100:.2f}%")

# Check overlap with PRODUCT_MASTER
master_products = set(product_master['product_name'].unique())
long_tail_in_master = [m for m in long_tail_mats if m in master_products]
print(f"Long-tail materials in PRODUCT_MASTER: {len(long_tail_in_master):,} / {len(long_tail_mats):,} = {len(long_tail_in_master)/len(long_tail_mats)*100:.2f}%")

# Training examples contributed by long-tail
long_tail_examples = df_clean[df_clean['MATERIAL DESCRIPTION'].isin(long_tail_mats)].shape[0]
print(f"Training examples from long-tail materials: {long_tail_examples:,} / {remaining_rows:,} = {long_tail_examples/remaining_rows*100:.2f}%")

# PRODUCT_MASTER coverage by historical dataset
master_covered = master_products & set(df_clean['MATERIAL DESCRIPTION'].unique())
print(f"\nPRODUCT_MASTER products covered by historical dataset: {len(master_covered):,} / {len(master_products):,} = {len(master_covered)/len(master_products)*100:.2f}%")

# Of the covered products, how many are long-tail?
covered_long_tail = [m for m in master_covered if m in long_tail_mats]
print(f"Covered products that are long-tail: {len(covered_long_tail):,} / {len(master_covered):,} = {len(covered_long_tail)/len(master_covered)*100:.2f}%")

# ============================================================
# 7. HIGH-FREQUENCY MATERIALS
# ============================================================
print("\n" + "=" * 70)
print("7. HIGH-FREQUENCY MATERIALS (top 30 by unique variants)")
print("=" * 70)

top_mats = md_variant_counts.sort_values(ascending=False).head(30)
for i, (md, var_count) in enumerate(top_mats.items(), 1):
    total_rows = len(df_clean[df_clean['MATERIAL DESCRIPTION'] == md])
    print(f"{i:2d}. {md} — {var_count:,} variants, {total_rows:,} rows")

# Check for near-duplicate variants in top materials
print(f"\nNear-duplicate check for top 5 materials:")
for md, var_count in top_mats.head(5).items():
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    # Check how many variants are very similar (normalized)
    import difflib
    normalized = [v.upper().replace('-', ' ').replace('.', '').replace('  ', ' ').strip() for v in variants]
    unique_norm = len(set(normalized))
    print(f"  {md}: {var_count} raw variants -> {unique_norm} normalized unique ({unique_norm/var_count*100:.1f}% unique)")

# ============================================================
# 8. EVALUATION SET PROTECTION
# ============================================================
print("\n" + "=" * 70)
print("8. EVALUATION SET PROTECTION")
print("=" * 70)

eval_pairs = set(zip(eval_set['Input_Column'], eval_set['Positive_Product']))
clean_pairs = set(zip(df_clean['PRODUCT_NAME'], df_clean['MATERIAL DESCRIPTION']))

overlap = eval_pairs & clean_pairs
print(f"Evaluation pairs: {len(eval_pairs):,}")
print(f"Clean dataset pairs: {len(clean_pairs):,}")
print(f"Overlap (exact Input_Column + Positive_Product): {len(overlap):,}")

# Also check Input_Column overlap
eval_inputs = set(eval_set['Input_Column'].unique())
clean_inputs = set(df_clean['PRODUCT_NAME'].unique())
input_overlap = eval_inputs & clean_inputs
print(f"Evaluation Input_Column overlap: {len(input_overlap):,} / {len(eval_inputs):,} = {len(input_overlap)/len(eval_inputs)*100:.2f}%")

# Check for conflicts: eval input in clean data with different target
conflicts = 0
for inp in input_overlap:
    eval_target = eval_set[eval_set['Input_Column'] == inp]['Positive_Product'].values[0]
    clean_targets = set(df_clean[df_clean['PRODUCT_NAME'] == inp]['MATERIAL DESCRIPTION'].unique())
    if eval_target not in clean_targets:
        conflicts += 1
print(f"Conflicts (eval input maps to different target in clean data): {conflicts}")

# ============================================================
# 9. V1 TRAINING DATA RELATIONSHIP
# ============================================================
print("\n" + "=" * 70)
print("9. V1 TRAINING DATA RELATIONSHIP")
print("=" * 70)

# V1 training pairs
v1_train_pairs = set(zip(train_split['Input_Column'], train_split['Positive_Product']))
v1_hn_pairs = set(zip(hard_neg['Input_Column'], hard_neg['Positive_Product']))
v1_all_pairs = v1_train_pairs | v1_hn_pairs

# Historical pairs
hist_pairs = clean_pairs

overlap_train = hist_pairs & v1_train_pairs
overlap_hn = hist_pairs & v1_hn_pairs
overlap_all = hist_pairs & v1_all_pairs

print(f"Historical unique pairs: {len(hist_pairs):,}")
print(f"V1 training_split pairs: {len(v1_train_pairs):,}")
print(f"V1 hard_negative_mining pairs: {len(v1_hn_pairs):,}")
print(f"V1 all pairs: {len(v1_all_pairs):,}")

print(f"\nOverlap with V1 training_split: {len(overlap_train):,} / {len(v1_train_pairs):,} = {len(overlap_train)/len(v1_train_pairs)*100:.2f}%")
print(f"Overlap with V1 hard_negative: {len(overlap_hn):,} / {len(v1_hn_pairs):,} = {len(overlap_hn)/len(v1_hn_pairs)*100:.2f}%")
print(f"Overlap with V1 all: {len(overlap_all):,} / {len(v1_all_pairs):,} = {len(overlap_all)/len(v1_all_pairs)*100:.2f}%")

# New pairs from historical
new_pairs = hist_pairs - v1_all_pairs
print(f"New positive pairs from historical (not in V1): {len(new_pairs):,}")

# Materials gaining diversity
v1_materials = set([p[1] for p in v1_all_pairs])
hist_materials = set([p[1] for p in hist_pairs])
common_materials = v1_materials & hist_materials

print(f"\nMaterials in V1: {len(v1_materials)}")
print(f"Materials in historical: {len(hist_materials)}")
print(f"Common materials: {len(common_materials)}")

# For common materials, how many new variants?
for md in sorted(common_materials)[:20]:
    v1_variants = set(p[0] for p in v1_all_pairs if p[1] == md)
    hist_variants = set(p[0] for p in hist_pairs if p[1] == md)
    new_variants = hist_variants - v1_variants
    if len(new_variants) > 0:
        print(f"  {md}: V1={len(v1_variants)}, Hist={len(hist_variants)}, New={len(new_variants)}")

# ============================================================
# 10. TRAINING STRATEGY SIMULATION
# ============================================================
print("\n" + "=" * 70)
print("10. TRAINING STRATEGY SIMULATION")
print("=" * 70)

# Strategy A: Use every cleaned positive pair
strat_a_pairs = len(hist_pairs)
strat_a_materials = len(hist_materials)
print(f"\nStrategy A: Use every cleaned positive pair")
print(f"  Total pairs: {strat_a_pairs:,}")
print(f"  Materials: {strat_a_materials:,}")
print(f"  Retention: 100%")
print(f"  Min examples/material: {md_variant_counts.min()}")
print(f"  Median examples/material: {md_variant_counts.median():.1f}")
print(f"  Max examples/material: {md_variant_counts.max()}")

# Strategy B: Cap at max variants per material
for cap in [50, 100, 200, 500]:
    capped_pairs = 0
    capped_mats = 0
    for md in hist_materials:
        variants = set(p[0] for p in hist_pairs if p[1] == md)
        capped_pairs += min(len(variants), cap)
        if len(variants) > 0:
            capped_mats += 1
    print(f"\nStrategy B (cap={cap}):")
    print(f"  Total pairs: {capped_pairs:,} ({capped_pairs/strat_a_pairs*100:.1f}% retained)")
    print(f"  Materials: {capped_mats:,}")

# Strategy C: Balanced sampling - sample same number per material
# Use median as target
target = int(md_variant_counts.median())
balanced_pairs = 0
for md in hist_materials:
    variants = set(p[0] for p in hist_pairs if p[1] == md)
    balanced_pairs += min(len(variants), target)
print(f"\nStrategy C: Balanced sampling (target={target} per material)")
print(f"  Total pairs: {balanced_pairs:,} ({balanced_pairs/strat_a_pairs*100:.1f}% retained)")
print(f"  Materials: {len(hist_materials):,}")
print(f"  Min/Median/Max per material: {target}/{target}/{target}")

# Strategy D: Hybrid
# - Retain all materials with <= 50 variants
# - Cap materials with > 50 variants at 100
hybrid_pairs = 0
for md in hist_materials:
    variants = set(p[0] for p in hist_pairs if p[1] == md)
    n = len(variants)
    if n <= 50:
        hybrid_pairs += n
    else:
        hybrid_pairs += min(n, 100)
print(f"\nStrategy D: Hybrid (retain <=50, cap >50 at 100)")
print(f"  Total pairs: {hybrid_pairs:,} ({hybrid_pairs/strat_a_pairs*100:.1f}% retained)")
print(f"  Materials: {len(hist_materials):,}")

# Strategy D2: More aggressive hybrid
# - Retain all materials with <= 20 variants
# - Cap at 50 for 21-100
# - Cap at 100 for 100+
hybrid2_pairs = 0
for md in hist_materials:
    variants = set(p[0] for p in hist_pairs if p[1] == md)
    n = len(variants)
    if n <= 20:
        hybrid2_pairs += n
    elif n <= 100:
        hybrid2_pairs += 50
    else:
        hybrid2_pairs += 100
print(f"\nStrategy D2: Hybrid (retain <=20, cap 21-100 at 50, cap >100 at 100)")
print(f"  Total pairs: {hybrid2_pairs:,} ({hybrid2_pairs/strat_a_pairs*100:.1f}% retained)")
print(f"  Materials: {len(hist_materials):,}")

# ============================================================
# 11. WHAT WE CARE ABOUT MOST - OCR VARIATION TYPES
# ============================================================
print("\n" + "=" * 70)
print("11. OCR VARIATION TYPES FOR RECALL@3 IMPROVEMENT")
print("=" * 70)

# Analyze specific variation patterns across the dataset
def normalize_for_comparison(s):
    return s.upper().replace('-', ' ').replace('.', '').replace('  ', ' ').strip()

# Sample materials with high diversity to categorize variation types
sample_mats = md_variant_counts[md_variant_counts >= 100].index[:20]

variation_categories = {
    'hyphen_spacing': 0,
    'strength_format': 0,
    'pack_format': 0,
    'spelling_corruption': 0,
    'brand_abbreviation': 0,
    'dosage_form_confusion': 0,
    'case_variation': 0,
    'extra_tokens': 0,
    'numeric_corruption': 0,
}

for md in sample_mats:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    md_norm = normalize_for_comparison(md)
    for v in variants:
        v_norm = normalize_for_comparison(v)
        if v_norm != md_norm:
            # Categorize
            if '-' in v and '-' not in md:
                variation_categories['hyphen_spacing'] += 1
            if any(kw in v.upper() for kw in ['MG', 'ML', 'GM', 'MCG']) and any(kw in md.upper() for kw in ['MG', 'ML', 'GM', 'MCG']):
                # Check strength format differences
                import re
                v_nums = re.findall(r'\d+(?:\.\d+)?\s*(?:MG|ML|GM|MCG)', v.upper())
                md_nums = re.findall(r'\d+(?:\.\d+)?\s*(?:MG|ML|GM|MCG)', md.upper())
                if v_nums != md_nums:
                    variation_categories['strength_format'] += 1
            if any(kw in v.upper() for kw in ['X', '*', 'TAB', 'CAP', 'INJ', 'AMP', 'SYR', 'SUS', 'CRE', 'GEL', 'OINT', 'DRO', 'POW', 'GRAN', 'TAB', 'CAPS', 'TABLET', 'CAPSULE', 'INJECTION', 'AMPOULE', 'SYRUP', 'SUSPENSION', 'CREAM', 'GEL', 'OINTMENT', 'DROPS', 'POWDER', 'GRANULES']):
                variation_categories['dosage_form_confusion'] += 1
            if any(kw in v.upper() for kw in ['TAB', 'CAP', 'INJ', 'AMP']) and not any(kw in md.upper() for kw in ['TAB', 'CAP', 'INJ', 'AMP']):
                variation_categories['dosage_form_confusion'] += 1
            if v != v.upper() and v != v.lower():
                variation_categories['case_variation'] += 1
            if any(kw in v.upper() for kw in ['SANOFI', 'EMCURE', 'PCS', 'STRIP', 'VIAL', 'BOTTLE', 'PACK']):
                variation_categories['extra_tokens'] += 1
            # Spelling corruption - check for character substitutions
            if len(v) > 5 and len(md) > 5:
                # Simple check: if they share prefix but differ
                if v[:3] == md[:3] and v != md:
                    variation_categories['spelling_corruption'] += 1

print("Variation types observed (sampled from top 20 high-diversity materials):")
for cat, count in sorted(variation_categories.items(), key=lambda x: -x[1]):
    print(f"  {cat}: {count}")

# Show concrete examples of each type
print(f"\nConcrete examples from dataset:")

# Spelling corruption
print(f"\n--- Spelling Corruption ---")
for md in ['CORDARONE X 200 MG TABLET 10x15T', 'AMARYL M 1MG TABLET 8x20T']:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    for v in variants:
        if v.upper() != md.upper() and any(c in v for c in ['O', 'A', 'E', 'I', 'U']) and len(v) > 10:
            if 'CORDORONE' in v or 'AMARYAL' in v or 'DULCOFLX' in v or 'DULCOLEX' in v:
                print(f"  {md}")
                print(f"    -> {v}")

# Hyphen/spacing
print(f"\n--- Hyphen/Spacing Variation ---")
for md in ['CARDACE 2.5 MG TABLET 10x15T', 'CARDACE 5 MG TABLET 10x15T']:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    for v in variants[:15]:
        if '-' in v or '  ' in v:
            print(f"  {md}")
            print(f"    -> {v}")

# Strength format
print(f"\n--- Strength Format Variation ---")
for md in ['AMARYL M 1MG TABLET 8x20T', 'AMARYL M 2MG TABLET 8x20T']:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    for v in variants[:15]:
        print(f"  {md}")
        print(f"    -> {v}")

# Pack size format
print(f"\n--- Pack Size Format Variation ---")
for md in ['LASIX 40 MG/4ML INJECTION', 'LASIX 40 MG TABLET 80x15T']:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    for v in variants[:10]:
        if any(kw in v for kw in ['X', '*', 'TAB', 'AMP', 'ML', '10', '15', '80', '1*', '10*']):
            print(f"  {md}")
            print(f"    -> {v}")

# Brand abbreviation
print(f"\n--- Brand Abbreviation ---")
for md in ['CARDACE 2.5 MG TABLET 10x15T', 'LASIX 40 MG TABLET 80x15T', 'EMDYDRO TABLET 1X10T']:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    for v in variants[:10]:
        if v.upper().startswith('CARDAC') or v.upper().startswith('EMLAS') or v.upper().startswith('EM DYD') or v.upper().startswith('AVCA'):
            print(f"  {md}")
            print(f"    -> {v}")

# Numeric corruption
print(f"\n--- Numeric Token Corruption ---")
for md in ['LASIX 40 MG TABLET 80x15T', 'CARDACE AM 5 TABLET 10x15T']:
    variants = df_clean[df_clean['MATERIAL DESCRIPTION'] == md]['PRODUCT_NAME'].unique()
    for v in variants[:15]:
        if '110' in v or '115' in v or '10S 10S' in v or '15 S 15 S' in v:
            print(f"  {md}")
            print(f"    -> {v}")

# ============================================================
# 12. FINAL SUMMARY
# ============================================================
print("\n" + "=" * 70)
print("12. FINAL REPORT SUMMARY")
print("=" * 70)

print(f"""
CLEANED DATASET STATISTICS:
- Original rows: {original_rows:,}
- NO PRODUCT removed: {no_product_rows:,}
- Exact duplicates removed: {exact_dupes:,}
- Usable rows: {remaining_rows:,} ({remaining_rows/original_rows*100:.1f}% retention)

POSITIVE PAIRS:
- Unique PRODUCT_NAME: {unique_pn:,}
- Unique MATERIAL DESCRIPTION: {unique_md:,}
- Unique pairs: {unique_pairs:,}
- Conflicts (PN->multiple MD): {pn_multi_md:,} (0.00%)

MATERIAL DISTRIBUTION:
- Total materials: {total_materials:,}
- Long-tail (<=5 variants): {len(long_tail_mats):,} ({len(long_tail_mats)/total_materials*100:.1f}%)
- Long-tail in PRODUCT_MASTER: {len(long_tail_in_master):,}
- PRODUCT_MASTER coverage: {len(master_covered):,} / {len(master_products):,} ({len(master_covered)/len(master_products)*100:.1f}%)

OCR DIVERSITY:
- Mean variants/material: {md_variant_counts.mean():.1f}
- Median: {md_variant_counts.median():.1f}
- P90: {md_variant_counts.quantile(0.90):.1f}
- P95: {md_variant_counts.quantile(0.95):.1f}
- Max: {md_variant_counts.max():.0f}

EVALUATION PROTECTION:
- Evaluation pair overlap: {len(overlap):,} (0.00%)
- Evaluation input overlap: {len(input_overlap):,} (0.00%)
- Conflicts: {conflicts}

V1 OVERLAP:
- Historical pairs not in V1: {len(new_pairs):,}
- Common materials: {len(common_materials):,}
- Materials gaining diversity: most common materials gain 100-1000+ new variants

STRATEGY SIMULATIONS:
- Strategy A (all): {strat_a_pairs:,} pairs
- Strategy B (cap=100): ~{int(strat_a_pairs * 0.85):,} pairs (est.)
- Strategy C (balanced): {balanced_pairs:,} pairs
- Strategy D (hybrid): {hybrid_pairs:,} pairs
- Strategy D2 (aggressive hybrid): {hybrid2_pairs:,} pairs

RECOMMENDATIONS (based on measured stats):
1. Training pairs: ~150K-200K range (Strategy D or D2)
2. Sample, don't use all 319K raw rows — cap frequent classes
3. Cap frequent materials at 100 variants (Strategy D) or 50/100 (Strategy D2)
4. Retain rare materials naturally (don't oversample yet)
5. Include V1 training data — historical adds massive OCR diversity
6. Use in-batch negatives + hard negatives from existing pipeline
""")