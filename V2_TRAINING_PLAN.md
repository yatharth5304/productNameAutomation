# V2 Training Plan

**Status**: AWAITING REVIEW — do not execute until approved.  
**Prepared**: 2026-09-23  

---

## 1. V1 Training Setup (verified from repository code)

All values below were read directly from
[`finetune_experiment.py`](file:///F:/Vintyaa/projects/Product%20Name%20Automation/finetune_experiment.py),
[`models/bge-pharma-v1/README.md`](file:///F:/Vintyaa/projects/Product%20Name%20Automation/models/bge-pharma-v1/README.md),
and
[`models/bge-pharma-v1/checkpoint-67/trainer_state.json`](file:///F:/Vintyaa/projects/Product%20Name%20Automation/models/bge-pharma-v1/checkpoint-67/trainer_state.json).

| Setting | V1 Value | Source |
|---------|----------|--------|
| Start model | `BAAI/bge-base-en-v1.5` (downloaded from HF) | `finetune_experiment.py` L114 |
| Training data | `training_split.csv` — 1,061 examples | README: `dataset_size:1061` |
| Data columns | anchor, positive, neg_1, neg_2, neg_3, neg_4 (6 cols) | `finetune_experiment.py` L83–96 |
| neg_1 source | `Explicit_Negative` (hand-labelled) | `finetune_experiment.py` L88 |
| neg_2/3/4 source | `Mined_Negative_1/2/3` (FAISS-mined) | `finetune_experiment.py` L89–91 |
| Loss | `MultipleNegativesRankingLoss` | `finetune_experiment.py` L118 |
| Loss scale (τ) | 20.0 | `finetune_experiment.py` L42 |
| Similarity function | cosine (`cos_sim`) | README loss params |
| Epochs | 1 | `finetune_experiment.py` L38 |
| Per-device batch size | 4 | `finetune_experiment.py` L36 |
| Gradient accumulation | 4 | `finetune_experiment.py` L37 |
| Effective batch size | **16** (4 × 4) | computed |
| Learning rate | 2e-5 | `finetune_experiment.py` L39 |
| Warmup | 0.1 (float → ratio → ~10% of steps) | `finetune_experiment.py` L131 |
| LR scheduler | linear decay after warmup | Trainer default |
| fp16 / bf16 | False / False | `finetune_experiment.py` L132 |
| Device | CPU (`use_cpu=True`) | `finetune_experiment.py` L133 |
| `dataloader_drop_last` | True | `finetune_experiment.py` L137 |
| Batch sampler | `BatchSamplers.NO_DUPLICATES` (if available) | `finetune_experiment.py` L140 |
| Logging steps | 25 | `finetune_experiment.py` L134 |
| Save strategy | epoch | `finetune_experiment.py` L135 |
| Output dir | `models/bge-pharma-v1/` | `finetune_experiment.py` L25 |
| Actual optimizer steps | 67 | `trainer_state.json` global_step=67 |
| Loss at step 25 | 1.207 | `trainer_state.json` |
| Loss at step 50 | 0.823 | `trainer_state.json` |
| Final training loss | ~0.959 | `finetuned_evaluation_report.txt` |
| Random seed | not set explicitly (HF default = 42) | `finetune_experiment.py` — absent |

### V1 Recorded Evaluation Results (target to beat)

| Metric | Baseline (BAAI) | V1 (bge-pharma-v1) | Delta V1 |
|--------|-----------------|---------------------|----------|
| Recall@1 | 42.9% | **72.4%** | +29.4% |
| Recall@3 | 80.6% | **94.7%** | +14.1% |
| Recall@5 | 92.9% | **98.2%** | +5.3% |
| Recall@10 | 95.9% | **100.0%** | +4.1% |
| MRR | 0.6346 | **0.8373** | +0.2026 |

V2 will be compared against **V1**, not against the BAAI baseline.

---

## 2. V2 Training Data

**Source file**: `v2_hard_negative_mining_clean.csv`  
**Row count**: 108,831 (validated — all 10 checks passed)

### Column mapping

| CSV column | Dataset column | Role |
|-----------|----------------|------|
| `Input_Column` | `anchor` | OCR product name (query) |
| `Positive_Product` | `positive` | Correct catalogue product |
| `Mined_Negative_1` | `negative_1` | Hard negative (highest similarity) |
| `Mined_Negative_2` | `negative_2` | Hard negative (second highest) |
| `Mined_Negative_3` | `negative_3` | Hard negative (third highest) |

Columns NOT used: `Positive_Product_Code`, `*_Code`, `*_Score` — metadata only.

### Dataset construction

```python
import csv, datasets

rows = list(csv.DictReader(open("v2_hard_negative_mining_clean.csv", encoding="utf-8")))

train_ds = datasets.Dataset.from_dict({
    "anchor":     [r["Input_Column"]     for r in rows],
    "positive":   [r["Positive_Product"] for r in rows],
    "negative_1": [r["Mined_Negative_1"] for r in rows],
    "negative_2": [r["Mined_Negative_2"] for r in rows],
    "negative_3": [r["Mined_Negative_3"] for r in rows],
})
```

> [!IMPORTANT]
> **V2 has 5 dataset columns, not 6.** V1 had an `Explicit_Negative` (hand-labelled) as `negative_1`, plus 3 mined negatives as `negative_2/3/4`. V2's mining CSV has **no Explicit_Negative column** — only 3 mined hard negatives. The dataset is therefore `(anchor, positive, neg_1, neg_2, neg_3)`.
>
> `MultipleNegativesRankingLoss` treats every column after `positive` as a negative regardless of count. 3 negatives is fully supported.

---

## 3. How the 3 Hard Negatives Enter Training

`MultipleNegativesRankingLoss` for a batch of N examples with K explicit negatives each:

1. All anchors, positives, and negatives are encoded.
2. For anchor `i`: the correct positive is `positive_i`. All other positives in the batch (**in-batch negatives**) plus the K=3 **explicit negatives** (`neg_1_i`, `neg_2_i`, `neg_3_i`) are treated as negatives.
3. Loss = cross-entropy over softmax of scaled cosine similarities.

With effective batch 16 and 3 negatives per example, each anchor faces:
- `16 - 1 = 15` in-batch negatives (other examples' positives)
- `3` explicit hard negatives
- **Total: 18 negatives per anchor per step**

The hard-mined negatives are guaranteed in the denominator for every anchor at every step — unlike purely in-batch negatives which are random. This is the mechanism by which V2 learns to discriminate similar products.

---

## 4. Loss Function

```python
from sentence_transformers.sentence_transformer.losses import MultipleNegativesRankingLoss

loss = MultipleNegativesRankingLoss(model, scale=20.0)
```

| Parameter | Value | Source |
|-----------|-------|--------|
| Loss class | `MultipleNegativesRankingLoss` | Inherited from V1 — unchanged |
| Scale (τ) | **20.0** | Inherited from V1 — unchanged |
| Similarity | cosine (default `cos_sim`) | Inherited from V1 — unchanged |

No custom loss is required. The existing framework correctly handles explicit negatives.

---

## 5. Batch Construction

| Setting | Value | Source |
|---------|-------|--------|
| Per-device batch size | **4** | Inherited from V1 |
| Gradient accumulation | **4** | Inherited from V1 |
| Effective batch size | **16** | Computed |
| `dataloader_drop_last` | **True** | Inherited from V1 — required by MNRL |
| `batch_sampler` | `BatchSamplers.NO_DUPLICATES` | Inherited from V1 |

---

## 6. Epochs

| Setting | Value | Rationale |
|---------|-------|-----------|
| Epochs | **1** | Inherited from V1. With 108,831 rows the model sees ~6,801 × 16 = 108,816 examples per epoch. One epoch provides complete coverage of all hard-negative pairs without risking overfitting. Can be extended after reviewing the 1-epoch checkpoint. |

---

## 7. Learning Rate

| Setting | Value | Source |
|---------|-------|--------|
| Learning rate | **2e-5** | Inherited from V1 |
| LR scheduler | Linear decay (Trainer default) | Inherited from V1 |

Note: starting from V1 checkpoint rather than the BAAI base means the model is already partially adapted. 2e-5 remains appropriate — it is not so large as to catastrophically forgetting V1's learning.

---

## 8. Gradient Accumulation

**4 steps** — inherited from V1. Achieves effective batch of 16 on CPU without requiring 16 examples simultaneously in memory.

---

## 9. Warmup and Scheduler

| Setting | Value | Source |
|---------|-------|--------|
| `warmup_steps` arg | **0.1** | Inherited from V1 |
| Interpretation | 10% of total steps = ~680 warmup steps | Float → ratio |
| LR peak | 2e-5 | At end of warmup |
| LR decay | Linear to ~0 | After warmup |

> [!NOTE]
> Passing `warmup_steps=0.1` (a float) to `SentenceTransformerTrainingArguments` (which inherits from HuggingFace `TrainingArguments`) triggers warmup-ratio behaviour. This is identical to V1 — confirmed by the V1 README listing `warmup_steps: 0.1`. V2 uses the same argument unchanged.

---

## 10. Evaluation Strategy

- **During training**: no online evaluation (same as V1, no `eval_dataset` argument).
- **After training**: a separate evaluation script (adapted from `finetune_eval.py`) will load `models/bge-pharma-v2/`, re-embed PRODUCT_MASTER, and run retrieval against `evaluation_set.csv`.
- **Metrics**: Recall@1, Recall@3, Recall@5, Recall@10, MRR, rank distribution.
- **Comparison**: V1 vs V2 on the same eval set with the same retrieval logic.

---

## 11. Output Checkpoint

| Item | Value |
|------|-------|
| **Start model** | `models/bge-pharma-v1/` (V1 fine-tuned checkpoint) |
| **Output dir** | `models/bge-pharma-v2/` |
| V1 modified? | **No — never touched** |
| Save strategy | epoch (one save at end of training) |

> [!IMPORTANT]
> **V2 starts from the V1 checkpoint, not from `BAAI/bge-base-en-v1.5`.**
> V1 already adapted the model to pharmaceutical product naming. V2 continues fine-tuning from V1 to add hard-negative discrimination. Starting from BAAI base would discard all of V1's domain adaptation.

---

## 12. Reproducibility Settings

| Setting | Value | Source |
|---------|-------|--------|
| Random seed | **42** | Newly explicit — V1 omitted this (relied on HF default 42). V2 sets `seed=42` in `SentenceTransformerTrainingArguments` to make it unambiguous. |
| Training data | `v2_hard_negative_mining_clean.csv` (108,831 rows) | Validated — cleaning report |
| Model start | `models/bge-pharma-v1/` | Fixed checkpoint |
| sentence-transformers | 6.0.1 | Confirmed from smoke test comments |

---

## 13. Expected Dataset Size

| Metric | Value |
|--------|-------|
| Training examples | 108,831 |
| Negatives per example | 3 |
| Total negative slots | 326,493 |
| Effective batch size | 16 |
| Steps per epoch (`drop_last=True`) | **6,801** |
| Total optimizer steps (1 epoch) | **6,801** |

### Runtime estimate (CPU)

V1 ran 67 steps in 52 minutes = **1.28 steps/min**.  
V2 at the same rate: 6,801 / 1.28 = **5,313 min ≈ 88.5 hours ≈ 3.7 days** on CPU.

> [!WARNING]
> **CPU training will take approximately 3–4 days.** This is because V2 has 101× more optimizer steps than V1. If a GPU is available, training can be completed in a few hours. Confirm hardware availability before starting.

---

## 14. Summary: V1 vs V2 Differences

| Aspect | V1 | V2 | Classification |
|--------|----|----|----------------|
| Start model | `BAAI/bge-base-en-v1.5` | `models/bge-pharma-v1/` | **Required by V2 design** |
| Training data | `training_split.csv` (1,061 rows) | `v2_hard_negative_mining_clean.csv` (108,831 rows) | **Required by V2 design** |
| Dataset columns | anchor, pos, neg×4 | anchor, pos, neg×3 | **Required by V2 data** (no Explicit_Negative) |
| Negative type | 1 explicit + 3 mined | 3 mined hard negatives | **Required by V2 design** |
| Output dir | `models/bge-pharma-v1/` | `models/bge-pharma-v2/` | **Required — must not overwrite V1** |
| Random seed | implicit (42) | explicit `seed=42` | **Newly explicit for reproducibility** |
| Loss | MNRL scale=20 | MNRL scale=20 | Inherited unchanged |
| Epochs | 1 | 1 | Inherited unchanged |
| Batch size | 4 | 4 | Inherited unchanged |
| Grad accum | 4 | 4 | Inherited unchanged |
| LR | 2e-5 | 2e-5 | Inherited unchanged |
| Warmup | 0.1 (ratio) | 0.1 (ratio) | Inherited unchanged |
| Device | CPU | CPU (GPU if available) | Inherited / advisory |

---

## 15. Open Questions for Review

1. **GPU availability** — Is a GPU available? If yes, training drops from ~88h to a few hours. Should `use_cpu=False` and `fp16=True` be set?

2. **Epochs** — Should V2 run for more than 1 epoch? With 6,801 steps, 1 epoch gives complete coverage. V1 also used 1 epoch. Recommend 1 epoch for the first run, evaluate, then decide on additional epochs.

3. **Temperature/scale** — Should scale remain 20? With harder negatives (mean similarity ~0.77 vs V1's ~0.80) the denominator scores are tighter. A higher scale (e.g., 25) would sharpen discrimination further. Recommend keeping 20 to minimise variables for the first run.

4. **Start model** — Confirmed to be `models/bge-pharma-v1/`? If the goal is to measure the pure effect of hard-negative mining independently of V1's training data, starting from `BAAI/bge-base-en-v1.5` would be a cleaner experiment — but would require re-learning domain adaptation from scratch. Assume V1 start unless instructed otherwise.

---

*Awaiting approval before `finetune_v2.py` is written and training is launched.*
