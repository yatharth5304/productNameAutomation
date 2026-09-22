---
tags:
- sentence-transformers
- sentence-similarity
- feature-extraction
- dense
- generated_from_trainer
- dataset_size:1061
- loss:MultipleNegativesRankingLoss
base_model: BAAI/bge-base-en-v1.5
widget:
- source_sentence: HOSIT D TAB S
  sentences:
  - HOSIT L TABLET
  - TRUSTEN AZ TAB
  - HOSIT TABLET
  - HOSIT D3 TABLET
  - ULTIMOX CLAV KID DT TAB
- source_sentence: HOSIT D TAB S
  sentences:
  - TRUSTEN AZ TAB
  - HOSIT L TABLET
  - ULTIMOX CLAV KID DT TAB
  - HOSIT TABLET
  - HOSIT D3 TABLET
- source_sentence: NUMLO AT 2315
  sentences:
  - NUMLO-AT TABLET 1X15 T
  - NUMLO-AT TABLET
  - NUMLO-R TABLETS
  - NUMLO-TM TABLETS
  - NUMLO-D TABLET
- source_sentence: NUMLO AT 2315
  sentences:
  - NUMLO-TM TABLETS
  - NUMLO-AT TABLET 1X15 T
  - NUMLO-D TABLET
  - NUMLO-AT TABLET
  - NUMLO-R TABLETS
- source_sentence: CINEZ D 10X15
  sentences:
  - VOVERAN D 10X10 T
  - CINEZ-D TABLETS
  - TREDICAL FT 10X10
  - IKA 10X10
  - CINEZ-D TABLETS 1 X 15 T
pipeline_tag: sentence-similarity
library_name: sentence-transformers
---

# SentenceTransformer based on BAAI/bge-base-en-v1.5

This is a [sentence-transformers](https://www.SBERT.net) model finetuned from [BAAI/bge-base-en-v1.5](https://huggingface.co/BAAI/bge-base-en-v1.5). It maps inputs to a 768-dimensional dense vector space and can be used for semantic textual similarity, semantic search, paraphrase mining, classification, clustering, and more.

## Model Details

### Model Description
- **Model Type:** Sentence Transformer
- **Base model:** [BAAI/bge-base-en-v1.5](https://huggingface.co/BAAI/bge-base-en-v1.5) <!-- at revision a5beb1e3e68b9ab74eb54cfd186867f64f240e1a -->
- **Maximum Sequence Length:** 512 tokens
- **Output Dimensionality:** 768 dimensions
- **Similarity Function:** Cosine Similarity
- **Supported Modality:** Text
<!-- - **Training Dataset:** Unknown -->
<!-- - **Language:** Unknown -->
<!-- - **License:** Unknown -->

### Model Sources

- **Documentation:** [Sentence Transformers Documentation](https://sbert.net)
- **Repository:** [Sentence Transformers on GitHub](https://github.com/huggingface/sentence-transformers)
- **Hugging Face:** [Sentence Transformers on Hugging Face](https://huggingface.co/models?library=sentence-transformers)

### Full Model Architecture

```
SentenceTransformer(
  (0): Transformer({'transformer_task': 'feature-extraction', 'modality_config': {'text': {'method': 'forward', 'method_output_name': 'last_hidden_state'}}, 'module_output_name': 'token_embeddings', 'architecture': 'BertModel'})
  (1): Pooling({'embedding_dimension': 768, 'pooling_mode': 'cls', 'include_prompt': True})
  (2): Normalize({'module_input_name': 'sentence_embedding', 'module_output_name': 'sentence_embedding'})
)
```

## Usage

### Direct Usage (Sentence Transformers)

First install the Sentence Transformers library:

```bash
pip install -U sentence-transformers
```
Then you can load this model and run inference.
```python
from sentence_transformers import SentenceTransformer

# Download from the 🤗 Hub
model = SentenceTransformer("sentence_transformers_model_id")
# Run inference
sentences = [
    'CINEZ D 10X15',
    'CINEZ-D TABLETS 1 X 15 T',
    'CINEZ-D TABLETS',
]
embeddings = model.encode(sentences)
print(embeddings.shape)
# [3, 768]

# Get the similarity scores for the embeddings
similarities = model.similarity(embeddings, embeddings)
print(similarities)
# tensor([[1.0000, 0.9118, 0.7964],
#         [0.9118, 1.0000, 0.8853],
#         [0.7964, 0.8853, 1.0000]])
```
<!--
### Direct Usage (Transformers)

<details><summary>Click to see the direct usage in Transformers</summary>

</details>
-->

<!--
### Downstream Usage (Sentence Transformers)

You can finetune this model on your own dataset.

<details><summary>Click to expand</summary>

</details>
-->

<!--
### Out-of-Scope Use

*List how the model may foreseeably be misused and address what users ought not to do with the model.*
-->

<!--
## Bias, Risks and Limitations

*What are the known or foreseeable issues stemming from this model? You could also flag here known failure cases or weaknesses of the model.*
-->

<!--
### Recommendations

*What are recommendations with respect to the foreseeable issues? For example, filtering explicit content.*
-->

## Training Details

### Training Dataset

#### Unnamed Dataset

* Size: 1,061 training samples
* Columns: <code>anchor</code>, <code>positive</code>, <code>negative_1</code>, <code>negative_2</code>, <code>negative_3</code>, and <code>negative_4</code>
* Approximate statistics based on the first 100 samples:
  |          | anchor                                                                            | positive                                                                          | negative_1                                                                       | negative_2                                                                        | negative_3                                                                        | negative_4                                                                        |
  |:---------|:----------------------------------------------------------------------------------|:----------------------------------------------------------------------------------|:---------------------------------------------------------------------------------|:----------------------------------------------------------------------------------|:----------------------------------------------------------------------------------|:----------------------------------------------------------------------------------|
  | type     | string                                                                            | string                                                                            | string                                                                           | string                                                                            | string                                                                            | string                                                                            |
  | modality | text                                                                              | text                                                                              | text                                                                             | text                                                                              | text                                                                              | text                                                                              |
  | details  | <ul><li>min: 5 tokens</li><li>mean: 12.46 tokens</li><li>max: 20 tokens</li></ul> | <ul><li>min: 7 tokens</li><li>mean: 11.53 tokens</li><li>max: 20 tokens</li></ul> | <ul><li>min: 5 tokens</li><li>mean: 9.98 tokens</li><li>max: 17 tokens</li></ul> | <ul><li>min: 6 tokens</li><li>mean: 11.93 tokens</li><li>max: 18 tokens</li></ul> | <ul><li>min: 6 tokens</li><li>mean: 11.44 tokens</li><li>max: 20 tokens</li></ul> | <ul><li>min: 6 tokens</li><li>mean: 11.09 tokens</li><li>max: 18 tokens</li></ul> |
* Samples:
  | anchor                                    | positive                                        | negative_1                                            | negative_2                                   | negative_3                                    | negative_4                              |
  |:------------------------------------------|:------------------------------------------------|:------------------------------------------------------|:---------------------------------------------|:----------------------------------------------|:----------------------------------------|
  | <code>GALACT POW 400GM 400GM</code>       | <code>GALACT GRANULES 400GM ELAICHI</code>      | <code>GALACT PLUS NUTRITION POWDER 400GM CHOCO</code> | <code>GALACT GRANULES 400 GM KESAR</code>    | <code>GALACT GRANULES 400 GM CHOCOLATE</code> | <code>PROTOTAL POW.200GM</code>         |
  | <code>CARDACE AM 2.25/5MG 1X10 PCS</code> | <code>CARDACE AM 2.5 TABLET 10x15T</code>       | <code>CARDACE AM 5 TABLET 10x15T</code>               | <code>CARDACE 1.25 MG TABLET 10x15T</code>   | <code>CARDACE 2.5 MG TABLET 10x15T</code>     | <code>CARDACE 5 MG TABLET 10x15T</code> |
  | <code>OSTERI INJ.600MCG/2.4 2.4ML</code>  | <code>OSTERI INJ.600MCG /2.4MLPEN DEVICE</code> | <code>OSTERI INJECTION 750 MCG PFS</code>             | <code>TASPIANCE INJ 180 MCG/0.5ML PFS</code> | <code>PROTOTAL POW.200GM</code>               | <code>MAXTRA  SYP  60 ML</code>         |
* Loss: [<code>MultipleNegativesRankingLoss</code>](https://sbert.net/docs/package_reference/sentence_transformer/losses.html#multiplenegativesrankingloss) with these parameters:
  ```json
  {
      "scale": 20.0,
      "similarity_fct": "cos_sim",
      "gather_across_devices": false,
      "directions": [
          "query_to_doc"
      ],
      "partition_mode": "joint",
      "hardness_mode": null,
      "hardness_strength": 0.0
  }
  ```

### Training Hyperparameters
#### Non-Default Hyperparameters

- `per_device_train_batch_size`: 4
- `num_train_epochs`: 1
- `learning_rate`: 2e-05
- `warmup_steps`: 0.1
- `gradient_accumulation_steps`: 4
- `use_cpu`: True
- `dataloader_drop_last`: True
- `dataloader_pin_memory`: False
- `batch_sampler`: no_duplicates

#### All Hyperparameters
<details><summary>Click to expand</summary>

- `per_device_train_batch_size`: 4
- `num_train_epochs`: 1
- `max_steps`: -1
- `learning_rate`: 2e-05
- `lr_scheduler_type`: linear
- `lr_scheduler_kwargs`: None
- `warmup_steps`: 0.1
- `optim`: adamw_torch_fused
- `optim_args`: None
- `weight_decay`: 0.0
- `adam_beta1`: 0.9
- `adam_beta2`: 0.999
- `adam_epsilon`: 1e-08
- `optim_target_modules`: None
- `gradient_accumulation_steps`: 4
- `average_tokens_across_devices`: True
- `max_grad_norm`: 1.0
- `label_smoothing_factor`: 0.0
- `bf16`: False
- `fp16`: False
- `bf16_full_eval`: False
- `fp16_full_eval`: False
- `tf32`: None
- `gradient_checkpointing`: False
- `gradient_checkpointing_kwargs`: None
- `torch_compile`: False
- `torch_compile_backend`: None
- `torch_compile_mode`: None
- `use_liger_kernel`: False
- `liger_kernel_config`: None
- `use_cache`: False
- `neftune_noise_alpha`: None
- `torch_empty_cache_steps`: None
- `auto_find_batch_size`: False
- `log_on_each_node`: True
- `logging_nan_inf_filter`: True
- `include_num_input_tokens_seen`: no
- `log_level`: passive
- `log_level_replica`: warning
- `disable_tqdm`: False
- `project`: huggingface
- `trackio_space_id`: None
- `trackio_bucket_id`: None
- `trackio_static_space_id`: None
- `per_device_eval_batch_size`: 8
- `prediction_loss_only`: True
- `eval_on_start`: False
- `eval_do_concat_batches`: True
- `eval_use_gather_object`: False
- `eval_accumulation_steps`: None
- `include_for_metrics`: []
- `batch_eval_metrics`: False
- `save_only_model`: False
- `save_on_each_node`: False
- `enable_jit_checkpoint`: False
- `push_to_hub`: False
- `hub_private_repo`: None
- `hub_model_id`: None
- `hub_strategy`: every_save
- `hub_always_push`: False
- `hub_revision`: None
- `load_best_model_at_end`: False
- `ignore_data_skip`: False
- `restore_callback_states_from_checkpoint`: False
- `full_determinism`: False
- `seed`: 42
- `data_seed`: None
- `use_cpu`: True
- `accelerator_config`: {'split_batches': False, 'dispatch_batches': None, 'even_batches': True, 'use_seedable_sampler': True, 'non_blocking': False, 'gradient_accumulation_kwargs': None}
- `parallelism_config`: None
- `dataloader_drop_last`: True
- `dataloader_num_workers`: 0
- `dataloader_pin_memory`: False
- `dataloader_persistent_workers`: False
- `dataloader_prefetch_factor`: None
- `dataloader_multiprocessing_context`: None
- `dataloader_in_order`: True
- `remove_unused_columns`: True
- `label_names`: None
- `train_sampling_strategy`: random
- `length_column_name`: length
- `ddp_find_unused_parameters`: None
- `ddp_bucket_cap_mb`: None
- `ddp_broadcast_buffers`: False
- `ddp_static_graph`: None
- `ddp_backend`: None
- `ddp_timeout`: 1800
- `fsdp`: None
- `fsdp_config`: None
- `deepspeed`: None
- `debug`: []
- `skip_memory_metrics`: True
- `do_predict`: False
- `resume_from_checkpoint`: None
- `local_rank`: -1
- `prompts`: None
- `batch_sampler`: no_duplicates
- `multi_dataset_batch_sampler`: proportional
- `router_mapping`: {}
- `learning_rate_mapping`: {}
- `warmup_ratio`: None

</details>

### Training Logs
| Epoch  | Step | Training Loss |
|:------:|:----:|:-------------:|
| 0.3774 | 25   | 1.2068        |
| 0.7547 | 50   | 0.8227        |


### Training Time
- **Training**: 52.3 minutes

### Framework Versions
- Python: 3.14.4
- Sentence Transformers: 6.0.1
- Transformers: 5.17.0
- PyTorch: 2.14.0+cpu
- Accelerate: 1.15.0
- Datasets: 5.0.1
- Tokenizers: 0.23.2

## Additional Resources

- [Training and Finetuning Embedding Models with Sentence Transformers](https://huggingface.co/blog/train-sentence-transformers): the end-to-end guide for training or finetuning Sentence Transformer models.
- [Introduction to Matryoshka Embedding Models](https://huggingface.co/blog/matryoshka): variable-size embeddings that can be truncated with minimal quality loss.
- [Binary and Scalar Embedding Quantization for Significantly Faster & Cheaper Retrieval](https://huggingface.co/blog/embedding-quantization): post-training compression of embedding vectors.
- [Multimodal Embedding & Reranker Models with Sentence Transformers](https://huggingface.co/blog/multimodal-sentence-transformers): use text, image, audio, and video models through the same API.
- [Training and Finetuning Multimodal Embedding & Reranker Models with Sentence Transformers](https://huggingface.co/blog/train-multimodal-sentence-transformers): train multimodal embedding models, with a Visual Document Retrieval walkthrough.

## Citation

### BibTeX

#### Sentence Transformers
```bibtex
@inproceedings{reimers-2019-sentence-bert,
    title = "Sentence-BERT: Sentence Embeddings using Siamese BERT-Networks",
    author = "Reimers, Nils and Gurevych, Iryna",
    booktitle = "Proceedings of the 2019 Conference on Empirical Methods in Natural Language Processing",
    month = "11",
    year = "2019",
    publisher = "Association for Computational Linguistics",
    url = "https://arxiv.org/abs/1908.10084",
}
```

#### MultipleNegativesRankingLoss
```bibtex
@misc{oord2019representationlearningcontrastivepredictive,
      title={Representation Learning with Contrastive Predictive Coding},
      author={Aaron van den Oord and Yazhe Li and Oriol Vinyals},
      year={2019},
      eprint={1807.03748},
      archivePrefix={arXiv},
      primaryClass={cs.LG},
      url={https://arxiv.org/abs/1807.03748},
}
```

<!--
## Glossary

*Clearly define terms in order to be accessible across audiences.*
-->

<!--
## Model Card Authors

*Lists the people who create the model card, providing recognition and accountability for the detailed work that goes into its construction.*
-->

<!--
## Model Card Contact

*Provides a way for people who have updates to the Model Card, suggestions, or questions, to contact the Model Card authors.*
-->