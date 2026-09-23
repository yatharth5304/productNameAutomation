# V2 HARD-NEGATIVE MINING QUALITY REPORT

Generated: 2026-09-23 10:28:10

## SUMMARY

- **Total positive pairs processed**: 108,913
- **Total negatives generated**: 326,739
- **Rows with 3 hard negatives**: 108,913 (100.0%)
- **Rows with 1-2 hard negatives**: 0 (0.0%)
- **Rows with 0 hard negatives**: 0 (0.0%)
- **Positive appeared in top-10**: 108,761
- **Validation errors**: 82
- **Evaluation leakage**: 0

## SIMILARITY SCORE DISTRIBUTION

| Slot | Count | Min | Max | Mean | Median | Std |
|------|-------|-----|-----|------|--------|-----|
| 1 | 108,913 | 0.4467 | 1.0000 | 0.7705 | 0.7816 | 0.0850 |
| 2 | 108,913 | 0.4369 | 0.9675 | 0.7168 | 0.7205 | 0.0838 |
| 3 | 108,913 | 0.4228 | 0.9612 | 0.6804 | 0.6801 | 0.0799 |

## SAME-BRAND PERCENTAGE

- **Slot 1**: 91,723 / 108,913 = 84.2%
- **Slot 2**: 78,996 / 108,913 = 72.5%
- **Slot 3**: 67,250 / 108,913 = 61.7%

## VALIDATION RESULTS

**FAILURES**:
- Duplicate negatives in row: VISTAVIT CAPS-15s
- Duplicate negatives in row: DINOR TAB 15 S
- Duplicate negatives in row: VISTAVIT CAPS 15S 10X15
- Duplicate negatives in row: DINOR 2 15TA
- Duplicate negatives in row: CALCIEM TAB 15S
- Duplicate negatives in row: DINOR 2MG TAB 15TAB
- Duplicate negatives in row: FM ACTIVE TAB 15T
- Duplicate negatives in row: CALCIEM 15 TAB
- Duplicate negatives in row: Dinor Tab 15 s
- Duplicate negatives in row: DINOR TABS 15T
- Duplicate negatives in row: DINOR TAB 15S
- Duplicate negatives in row: DINOR TAB 15T
- Duplicate negatives in row: FM ACTIVE 1*15T
- Duplicate negatives in row: DINOR TAB 15TAB
- Duplicate negatives in row: CALCIEM TABS 15S
- Duplicate negatives in row: ZILARBI-CN 10MG TA 15S
- Duplicate negatives in row: DINOR TABS 15S
- Duplicate negatives in row: Calciem Tab 15S
- Duplicate negatives in row: FEMZACT 15 S
- Duplicate negatives in row: CALCIEM TAB. 15*TAB
- ... and 62 more
## REPRESENTATIVE MINING EXAMPLES

### Input: `EMANZEN 5 MG TAB 10 S`
- **Positive**: EMANZEN TABLET 5 MG [424440086]
- **Mined 1**: EMANZEN N CAPSULES 10 [420000468] score=0.7502
- **Mined 2**: EMANZEN AP TABLET 10x10T [421113182] score=0.7398
- **Mined 3**: EMANZEN FORTE TABLET 10 MG [424440087] score=0.7299

### Input: `Hosit Amp.`
- **Positive**: HOSIT INJECTION [421110085]
- **Mined 1**: HOSIT TABLET [424441222] score=0.7107
- **Mined 2**: HOSIT FORTE INJECTION [421111291] score=0.6423
- **Mined 3**: HOSIT L TABLET [424441221] score=0.6403

### Input: `ASOMEX 5 20*10`
- **Positive**: ASOMEX 5 MG TABLET [424440003]
- **Mined 1**: ASOMEX AT 5 MG TABLET [424440006] score=0.7729
- **Mined 2**: ASOMEX 5 MG TABLET 1x15T [424441737] score=0.7513
- **Mined 3**: ASOMEX AT 5 MG TABLET 1X15T [424441922] score=0.7194

### Input: `XILIA TRIO-1 1*10`
- **Positive**: XILIA TRIO - 1 TABLETS [421111092]
- **Mined 1**: XILIA TRIO - 2 TABLETS [421111093] score=0.7768
- **Mined 2**: XILIA - 1 TABLETS [421111516] score=0.7507
- **Mined 3**: XILIA-MP 1 TABLETS [421110647] score=0.6883

### Input: `DOLOSPAN TAB 10 X10T`
- **Positive**: DOLOSPAN TABLET [421111806]
- **Mined 1**: DOLOSPAN DP TABLET [420001015] score=0.7501
- **Mined 2**: DOLOSPAN P TABLETS [421110042] score=0.7274
- **Mined 3**: DOLOSPAN INJECTION [421111821] score=0.7157

### Input: `VINTOR 2000IU INJ PFS 1`
- **Positive**: VINTOR 2000 IU INJ 1 PFS [422220113]
- **Mined 1**: VINTOR 6000 IU INJ 1 PFS [422220118] score=0.8286
- **Mined 2**: VINTOR 10000 IU INJ 1 PFS [422220120] score=0.8181
- **Mined 3**: VINTOR 4000 IU INJ 1 PFS [422220108] score=0.8136

### Input: `CARDACE AM 10 TABL1X15`
- **Positive**: CARDACE AM 10 TABLET 10x15T [421112894]
- **Mined 1**: CARDACE AM 5 TABLET 10x15T [421112896] score=0.8824
- **Mined 2**: CARDACE 10 MG TABLET 10x15T [421112891] score=0.8400
- **Mined 3**: CARDACE AM 2.5 TABLET 10x15T [421112895] score=0.8230

### Input: `EMANZEN - D Tab`
- **Positive**: EMANZEN D TABLET [424440056]
- **Mined 1**: EMANZEN DP TABLET [421112653] score=0.8379
- **Mined 2**: EMANZEN D OINTMENT 30 GM [421110987] score=0.7896
- **Mined 3**: EMZOSIN - D TABLET [421111726] score=0.7495

### Input: `FERIUM XT 15TA`
- **Positive**: FERIUM XT TABLET 10x15T [424441955]
- **Mined 1**: FERIUM XT DROPS 15 ML [424441719] score=0.8071
- **Mined 2**: FERIUM XT TABLET [424441702] score=0.7979
- **Mined 3**: FERIUM XT SUSPENSION [424440591] score=0.7397

### Input: `XILIA TRIO-1 110TAB`
- **Positive**: XILIA TRIO - 1 TABLETS [421111092]
- **Mined 1**: XILIA TRIO - 2 TABLETS [421111093] score=0.7338
- **Mined 2**: XILIA - 1 TABLETS [421111516] score=0.6851
- **Mined 3**: XILIA-MP 1 TABLETS [421110647] score=0.6396

## WEAKEST & STRONGEST NEGATIVES

**Weakest (lowest similarity)**:
- `UNMET` → **UNMET TABLET 10x10T** vs **G-MET SUSPENSION** (score=0.4467)
- `NUACE (DARUN-RITONO)-800 MG TAB. 1*30` → **NUACE TABLETS 1x30 T** vs **CARDACE 10 MG TABLET 10x15T** (score=0.4628)
- `VONADAY BOTT` → **VONADAY 600/300/300 MG TABLET 30s** vs **VONARIL INJECTION** (score=0.4645)
- `IFEN PLASTER --` → **INFEN PLASTER 1X7** vs **EVANEW GLOVES** (score=0.4671)
- `DINOR TAB 1S` → **DINOR TABLET** vs **TEMNER TABLETS 1X30T** (score=0.4677)

**Strongest (highest similarity)**:
- `EMPRI 25MG TABLET 10X10T 10*10T` → **EMPRI L 25/5MG TABLET 10x10T** vs **EMPRI 25MG TABLET 10x10T** (score=0.9919)
- `XGRAST 300 MCG INJ 1 PFS 1` → **XGRAST  INJ 300 MCG/0.5 ML PFS** vs **XGRAST 300 MCG INJ 1 PFS** (score=0.9933)
- `DIOF DS SUSPENSION 60ML 60ML` → **DIOF DS SUSPENSION** vs **DIOF DS SUSPENSION 60ML** (score=0.9957)
- `I-RINSE SPRAY 100 ML` → **I RINSE NASAL SPRAY** vs **I-RINSE SPRAY 100 ML** (score=1.0000)
- `LAZID E KIT` → **LAZID E KIT 150/300/600 MG TABLET 2X5** vs **LAZID E KIT** (score=1.0000)

---
*Total runtime: 1349.2s*
