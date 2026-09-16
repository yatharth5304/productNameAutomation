"""
transformer.py -- Sentence Transformer retrieval stage
=======================================================

Architecture position:

    OCR Extracted Row
          |
    garbage_check.py   ->  product/garbage decision + brand extraction
          |
    transformer.py     ->  brand hard-filter + semantic retrieval -> Top-10 candidates
          |
    (future)           ->  detailed scoring / ranking stage -> final mapping

Responsibility of THIS module:
  - Load BAAI/bge-base-en-v1.5 ONCE per process.
  - Maintain precomputed, normalised 768-d embeddings for every row in
    PRODUCT_MASTER.xlsx (product_name text only).
  - Accept an OCR product name + optional brand hint from garbage_check.py.
  - Hard-filter the master to brand-specific rows (Case A) or fall back to
    the full master (Case B: no brand; Case C: brand not found in master).
  - Return the Top-10 semantically similar candidates using cosine similarity
    (implemented as dot-product on unit-normalised vectors).
  - NEVER make the final product-mapping decision; that is reserved for the
    future scoring/ranking stage.

Usage -- standalone (CLI):
    python transformer.py --rebuild           # generate / refresh master embeddings
    python transformer.py --query "GALACT POW 200GM 1*200G" --brand "GALACT"
    python transformer.py --query "CARDACE AM 2.25/5MG 1X10 PCS"  # no brand
    python transformer.py                     # run built-in test suite

Usage -- importable API:
    from transformer import load_retriever, retrieve

    retriever = load_retriever()              # load once
    result    = retrieve(retriever, "GALACT POW 200GM 1*200G", brand="GALACT")
    for c in result["candidates"]:
        print(c["rank"], c["similarity_score"], c["product_name"])
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional

import numpy as np
import openpyxl

# ============================================================================
# CONFIG -- edit these values here; do NOT scatter magic numbers through code.
# Follows the existing project convention of a flat top-of-file config block.
# ============================================================================

# Path to the authoritative product master (3 cols: PRODUCT_CODE, product_name, BRAND_NAME)
MASTER_XLSX_PATH: str = r"f:\Vintyaa\projects\Product Name Automation\PRODUCT_MASTER.xlsx"

# Subdirectory for all generated embedding artifacts (relative to this file's directory)
EMBEDDINGS_DIR: str = "embeddings"

# Filenames inside EMBEDDINGS_DIR
EMBEDDINGS_NPY_FILE:  str = "master_embeddings.npy"
EMBEDDINGS_META_FILE: str = "master_embeddings_meta.json"
RECORD_INDEX_FILE:    str = "master_record_index.json"

# Sentence Transformer model.
# DO NOT change without explicit instruction; downstream fine-tuning targets this model.
MODEL_NAME: str = "BAAI/bge-base-en-v1.5"

# Number of candidates to return per query (FIXED at 10 per specification)
TOP_K: int = 10

# Batch size used when encoding master product names during --rebuild.
ENCODE_BATCH_SIZE: int = 64

# Device selection. None = auto-detect (CPU if no CUDA, else CUDA).
DEVICE: Optional[str] = None  # auto-detect

# Expected embedding dimension for BGE-base-en-v1.5.
EXPECTED_DIM: int = 768

# Special brand values written by garbage_check.py that do NOT represent a real brand.
_NON_BRAND_SENTINELS: frozenset = frozenset({"pattern-inferred", "nan", "none", ""})

# ============================================================================
# LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ============================================================================
# PATH HELPERS
# ============================================================================

def _project_dir() -> Path:
    return Path(__file__).resolve().parent


def _embeddings_dir() -> Path:
    return _project_dir() / EMBEDDINGS_DIR


def _npy_path() -> Path:
    return _embeddings_dir() / EMBEDDINGS_NPY_FILE


def _meta_path() -> Path:
    return _embeddings_dir() / EMBEDDINGS_META_FILE


def _index_path() -> Path:
    return _embeddings_dir() / RECORD_INDEX_FILE


# ============================================================================
# PRODUCT MASTER I/O
# ============================================================================

def load_master_records(master_path: str = MASTER_XLSX_PATH) -> List[Dict]:
    """
    Load PRODUCT_MASTER.xlsx and return a list of record dicts.

    Verified columns: PRODUCT_CODE, product_name, BRAND_NAME.

    Raises:
        FileNotFoundError  -- if the file does not exist.
        ValueError         -- if required columns are missing or master is empty.
    """
    path = Path(master_path)
    if not path.exists():
        raise FileNotFoundError(
            f"PRODUCT_MASTER not found: {path}\n"
            "Ensure MASTER_XLSX_PATH points to the correct file."
        )

    logger.info("Loading master: %s", path)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active

    # --- stream the sheet once with iter_rows (values_only=True) ------------
    # IMPORTANT: do NOT use ws.cell(row, col) inside a read_only worksheet.
    # That API re-parses the entire XML stream from the start for every call,
    # producing O(N²) runtime.  iter_rows() streams the XML exactly once.

    row_iter = ws.iter_rows(values_only=True)

    # First row = headers
    try:
        header_row = next(row_iter)
    except StopIteration:
        wb.close()
        raise ValueError("PRODUCT_MASTER is empty (no header row).")

    # Build case-insensitive column index map (0-based within the row tuple)
    col_map: Dict[str, int] = {}
    for idx, h in enumerate(header_row):
        if h is not None:
            col_map[str(h).strip().upper()] = idx

    required_cols = {"PRODUCT_CODE", "PRODUCT_NAME", "BRAND_NAME"}
    missing = required_cols - set(col_map.keys())
    if missing:
        wb.close()
        raise ValueError(
            f"Required column(s) {missing} not found in PRODUCT_MASTER.\n"
            f"Actual headers: {list(header_row)}\n"
            "Cannot proceed. Check MASTER_XLSX_PATH and the file contents."
        )

    pc_idx = col_map["PRODUCT_CODE"]
    pn_idx = col_map["PRODUCT_NAME"]
    bn_idx = col_map["BRAND_NAME"]

    records: List[Dict] = []
    for excel_row_num, row in enumerate(row_iter, start=2):   # 1-based, data starts at row 2
        product_code = row[pc_idx]
        product_name = row[pn_idx]
        brand_name   = row[bn_idx]

        if product_name is None or str(product_name).strip() == "":
            continue

        records.append({
            "row_number":   excel_row_num,
            "PRODUCT_CODE": str(product_code).strip() if product_code is not None else "",
            "product_name": str(product_name).strip(),
            "BRAND_NAME":   str(brand_name).strip().upper() if brand_name is not None else "",
        })

    wb.close()

    if not records:
        raise ValueError("PRODUCT_MASTER has no valid rows with a non-empty product_name.")

    logger.info("  Loaded %d master records.", len(records))
    return records


def _master_file_hash(master_path: str = MASTER_XLSX_PATH) -> str:
    """MD5 hash of the master file bytes -- used to detect staleness."""
    h = hashlib.md5()
    with open(master_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


# ============================================================================
# DEVICE SELECTION
# ============================================================================

def _resolve_device(device: Optional[str] = DEVICE) -> str:
    if device is not None:
        return device
    try:
        import torch
        return "cuda" if torch.cuda.is_available() else "cpu"
    except ImportError:
        return "cpu"


# ============================================================================
# MODEL LOADING
# ============================================================================

def load_model(model_name: str = MODEL_NAME, device: Optional[str] = DEVICE):
    """
    Load the SentenceTransformer model ONCE per process.

    Keep the returned object alive and pass it to all encode calls.
    Never reload the model inside a per-row loop.
    """
    resolved = _resolve_device(device)
    logger.info("Loading model '%s' on device '%s' ...", model_name, resolved)
    t0 = time.perf_counter()

    try:
        from sentence_transformers import SentenceTransformer
    except ImportError:
        raise ImportError(
            "sentence-transformers is not installed.\n"
            "Install with: pip install sentence-transformers"
        )

    try:
        model = SentenceTransformer(model_name, device=resolved)
    except Exception as exc:
        raise RuntimeError(
            f"Failed to load model '{model_name}' on device '{resolved}'.\n"
            f"Underlying error: {exc}"
        ) from exc

    logger.info("  Model loaded in %.2fs.", time.perf_counter() - t0)

    # Validate embedding dimension with a probe encoding
    probe = model.encode(["probe"], normalize_embeddings=True, show_progress_bar=False)
    actual_dim = probe.shape[1]
    if actual_dim != EXPECTED_DIM:
        raise ValueError(
            f"Model '{model_name}' returned {actual_dim}-d embeddings "
            f"but EXPECTED_DIM={EXPECTED_DIM}. "
            "Update EXPECTED_DIM in the config block if this is intentional."
        )

    return model


# ============================================================================
# MASTER EMBEDDING GENERATION
# ============================================================================

def compute_master_embeddings(
    records: List[Dict],
    model,
    batch_size: int = ENCODE_BATCH_SIZE,
) -> np.ndarray:
    """
    Encode every master product_name and return a normalised float32 matrix.

    Shape: (len(records), EXPECTED_DIM)
    Normalisation: unit L2 norm per row -> cosine similarity == dot product.
    """
    texts = [r["product_name"] for r in records]
    logger.info("Encoding %d master product names (batch_size=%d) ...", len(texts), batch_size)
    t0 = time.perf_counter()

    matrix = model.encode(
        texts,
        batch_size=batch_size,
        normalize_embeddings=True,
        show_progress_bar=True,
        convert_to_numpy=True,
    )
    matrix = matrix.astype(np.float32)

    elapsed = time.perf_counter() - t0
    logger.info(
        "  Encoded %d records -> shape %s in %.2fs (%.1f rec/s).",
        len(records), matrix.shape, elapsed, len(records) / elapsed,
    )
    return matrix


# ============================================================================
# EMBEDDING PERSISTENCE
# ============================================================================

def save_embeddings(
    matrix:      np.ndarray,
    records:     List[Dict],
    model_name:  str = MODEL_NAME,
    master_path: str = MASTER_XLSX_PATH,
) -> None:
    """Save embedding matrix, metadata, and record index to EMBEDDINGS_DIR."""
    emb_dir = _embeddings_dir()
    emb_dir.mkdir(parents=True, exist_ok=True)

    np.save(str(_npy_path()), matrix)
    logger.info("  Saved embeddings: %s  (shape %s, dtype %s)", _npy_path(), matrix.shape, matrix.dtype)

    master_hash = _master_file_hash(master_path)
    meta = {
        "model_name":    model_name,
        "embedding_dim": int(matrix.shape[1]),
        "n_records":     int(matrix.shape[0]),
        "master_path":   str(Path(master_path).resolve()),
        "master_hash":   master_hash,
        "generated_at":  time.strftime("%Y-%m-%dT%H:%M:%S"),
    }
    with open(_meta_path(), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)
    logger.info("  Saved metadata:   %s", _meta_path())

    with open(_index_path(), "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)
    logger.info("  Saved record index: %s  (%d records)", _index_path(), len(records))


def _load_raw_embeddings() -> tuple:
    npy  = _npy_path()
    meta = _meta_path()
    idx  = _index_path()

    missing = [str(p) for p in (npy, meta, idx) if not p.exists()]
    if missing:
        raise FileNotFoundError(
            "Embedding artifacts not found:\n"
            + "\n".join(f"  {m}" for m in missing)
            + "\n\nGenerate them by running:\n"
            "  python transformer.py --rebuild"
        )

    matrix = np.load(str(npy))
    with open(meta, "r", encoding="utf-8") as f:
        meta_dict = json.load(f)
    with open(idx, "r", encoding="utf-8") as f:
        records = json.load(f)

    return matrix, meta_dict, records


def _validate_embeddings_meta(
    meta:         dict,
    live_records: List[Dict],
    model_name:   str = MODEL_NAME,
    master_path:  str = MASTER_XLSX_PATH,
) -> None:
    """
    Validate stored metadata against the current PRODUCT_MASTER and model.
    Raises ValueError on ANY mismatch -- never silently accepts stale embeddings.
    """
    errors = []

    if meta.get("model_name") != model_name:
        errors.append(
            f"  Model mismatch: stored='{meta.get('model_name')}' vs configured='{model_name}'"
        )

    stored_dim = meta.get("embedding_dim")
    if stored_dim != EXPECTED_DIM:
        errors.append(f"  Dimension mismatch: stored={stored_dim} vs expected={EXPECTED_DIM}")

    stored_n = meta.get("n_records")
    live_n   = len(live_records)
    if stored_n != live_n:
        errors.append(
            f"  Record count mismatch: stored={stored_n} vs current master={live_n}"
        )

    stored_hash = meta.get("master_hash")
    live_hash   = _master_file_hash(master_path)
    if stored_hash != live_hash:
        errors.append(
            f"  PRODUCT_MASTER file has changed (MD5 mismatch).\n"
            f"    stored hash : {stored_hash}\n"
            f"    current hash: {live_hash}"
        )

    if errors:
        raise ValueError(
            "Stored embeddings are stale or incompatible with the current "
            "PRODUCT_MASTER / model configuration:\n"
            + "\n".join(errors)
            + "\n\nDo NOT use stale embeddings. Regenerate by running:\n"
            "  python transformer.py --rebuild"
        )


def load_embeddings(
    model_name:  str = MODEL_NAME,
    master_path: str = MASTER_XLSX_PATH,
) -> tuple:
    """
    Load and validate stored embeddings.

    Returns (matrix: np.ndarray, records: List[Dict]) on success.
    Raises FileNotFoundError if never generated; ValueError if stale.
    """
    matrix, meta, records = _load_raw_embeddings()
    live_records = load_master_records(master_path)
    _validate_embeddings_meta(meta, live_records, model_name, master_path)
    logger.info(
        "Embeddings validated: model='%s', dim=%d, records=%d.",
        meta["model_name"], meta["embedding_dim"], meta["n_records"],
    )
    return matrix, records


# ============================================================================
# END-TO-END REBUILD
# ============================================================================

def rebuild_master_embeddings(
    model_name:  str = MODEL_NAME,
    master_path: str = MASTER_XLSX_PATH,
    batch_size:  int = ENCODE_BATCH_SIZE,
    device:      Optional[str] = DEVICE,
) -> tuple:
    """
    Full rebuild: load master -> load model -> encode -> save.
    Returns (matrix, records, model).
    """
    logger.info("=== Master embedding rebuild started ===")
    t_total = time.perf_counter()

    records = load_master_records(master_path)
    model   = load_model(model_name, device)
    matrix  = compute_master_embeddings(records, model, batch_size)
    save_embeddings(matrix, records, model_name, master_path)

    logger.info("=== Rebuild complete in %.2fs. ===", time.perf_counter() - t_total)
    return matrix, records, model


# ============================================================================
# BRAND FILTER
# ============================================================================

def _normalise_brand(brand: Optional[str]) -> Optional[str]:
    """
    Normalise a brand hint from garbage_check.py.

    Returns None for None, empty, or sentinel values ("pattern-inferred",
    "nan", "none"), which all mean "no real brand -- use full master".

    Otherwise returns UPPERCASE-stripped string for case-insensitive comparison
    against BRAND_NAME in the master.

    Preserves special brands: "S-NUMLO", "EMCURE ORS", compound names, etc.
    """
    if brand is None:
        return None
    s = str(brand).strip()
    if s.upper() in _NON_BRAND_SENTINELS:
        return None
    return s.upper()


def get_brand_filtered_indices(
    records:    List[Dict],
    brand_hint: Optional[str],
) -> List[int]:
    """
    Return the list of record indices eligible for retrieval.

    Case A -- brand detected AND has matching master rows:
        indices where BRAND_NAME == norm_brand (case-insensitive, both upper)

    Case B -- no brand detected (brand_hint is None or a sentinel):
        all indices [0..N-1]

    Case C -- brand detected but ZERO master rows match:
        all indices [0..N-1]  (fall back to full master per specification)
    """
    norm_brand = _normalise_brand(brand_hint)

    if norm_brand is None:
        return list(range(len(records)))          # Case B

    matched = [i for i, r in enumerate(records) if r["BRAND_NAME"] == norm_brand]

    if matched:
        return matched                             # Case A

    logger.debug(
        "Brand '%s' not found in PRODUCT_MASTER -- falling back to full master (%d records).",
        norm_brand, len(records),
    )
    return list(range(len(records)))               # Case C


# ============================================================================
# QUERY ENCODING
# ============================================================================

def encode_query(query_text: str, model) -> np.ndarray:
    """
    Encode ONE OCR product name.

    Used exactly as provided -- no cleaning, normalisation, or reconstruction.

    Returns 1-D float32 ndarray of shape (EXPECTED_DIM,), unit-normalised.
    """
    vec = model.encode(
        [query_text],
        normalize_embeddings=True,
        show_progress_bar=False,
        convert_to_numpy=True,
    )
    return vec[0].astype(np.float32)


def encode_queries_batch(
    query_texts: List[str],
    model,
    batch_size: int = ENCODE_BATCH_SIZE,
) -> np.ndarray:
    """Encode multiple OCR queries in one batched call."""
    matrix = model.encode(
        query_texts,
        batch_size=batch_size,
        normalize_embeddings=True,
        show_progress_bar=False,
        convert_to_numpy=True,
    )
    return matrix.astype(np.float32)


# ============================================================================
# RETRIEVAL
# ============================================================================

def retrieve_top_k(
    query_vec:        np.ndarray,
    master_matrix:    np.ndarray,
    filtered_indices: List[int],
    records:          List[Dict],
    k:                int = TOP_K,
) -> List[Dict]:
    """
    Cosine similarity retrieval (dot product on unit-normalised vectors).

    Parameters
    ----------
    query_vec        : 1-D float32  (EXPECTED_DIM,)
    master_matrix    : 2-D float32  (N_master, EXPECTED_DIM)
    filtered_indices : eligible row indices into records / master_matrix
    records          : master record dicts, parallel to master_matrix
    k                : candidates to return (default TOP_K)

    Returns
    -------
    List[Dict] -- up to k candidates with keys:
        rank, similarity_score, PRODUCT_CODE, product_name, BRAND_NAME,
        master_record_index, master_row_number
    """
    if not filtered_indices:
        return []

    idx_arr    = np.array(filtered_indices, dtype=np.int32)
    sub_matrix = master_matrix[idx_arr]                      # (n_eligible, dim)
    scores     = sub_matrix @ query_vec                      # cosine similarity

    actual_k    = min(k, len(filtered_indices))
    top_sub_pos = np.argsort(scores)[::-1][:actual_k]

    candidates = []
    for rank, sub_pos in enumerate(top_sub_pos, start=1):
        master_idx = filtered_indices[int(sub_pos)]
        record     = records[master_idx]
        candidates.append({
            "rank":                rank,
            "similarity_score":    float(scores[sub_pos]),
            "PRODUCT_CODE":        record["PRODUCT_CODE"],
            "product_name":        record["product_name"],
            "BRAND_NAME":          record["BRAND_NAME"],
            "master_record_index": master_idx,
            "master_row_number":   record["row_number"],
        })

    return candidates


# ============================================================================
# HIGH-LEVEL IMPORTABLE API
# ============================================================================

class Retriever:
    """
    Holds all state needed for retrieval queries.

    Create ONE instance per process with load_retriever() and reuse it.
    Never recreate inside a per-row loop.
    """

    def __init__(self, model, master_matrix: np.ndarray, records: List[Dict]):
        self.model         = model
        self.master_matrix = master_matrix
        self.records       = records

    @property
    def n_master(self) -> int:
        return len(self.records)


def load_retriever(
    model_name:  str = MODEL_NAME,
    master_path: str = MASTER_XLSX_PATH,
    device:      Optional[str] = DEVICE,
) -> "Retriever":
    """
    Validate stored embeddings and load the model.

    Raises FileNotFoundError / ValueError if embeddings are missing or stale.

    Primary entry point for downstream pipeline code:

        from transformer import load_retriever, retrieve
        retriever = load_retriever()
        result    = retrieve(retriever, ocr_text, brand=brand_hint)
    """
    master_matrix, records = load_embeddings(model_name, master_path)
    model = load_model(model_name, device)
    logger.info(
        "Retriever ready: %d master records, dim=%d, device=%s.",
        len(records), master_matrix.shape[1], _resolve_device(device),
    )
    return Retriever(model, master_matrix, records)


def retrieve(
    retriever:  "Retriever",
    ocr_text:   str,
    brand:      Optional[str] = None,
    top_k:      int = TOP_K,
) -> Dict:
    """
    Semantic retrieval for one OCR product name.

    Parameters
    ----------
    retriever : Retriever   from load_retriever()
    ocr_text  : str         original OCR name -- used EXACTLY as-is
    brand     : str|None    brand from garbage_check.py col C (None = full master)
    top_k     : int         candidates to return

    Returns
    -------
    Dict:
        ocr_input           str      -- original OCR text
        detected_brand      str|None -- normalised brand (None if none detected)
        candidate_pool_size int      -- eligible master records before ST
        filter_case         str      -- "A" | "B" | "C" | "EMPTY_INPUT"
        candidates          List[Dict] -- up to top_k ranked candidates

    IMPORTANT: similarity_score is semantic similarity, NOT final match confidence.
    The #1 candidate is NOT the confirmed product. Forward all candidates to
    the downstream scoring/ranking stage.
    """
    if not ocr_text or str(ocr_text).strip() == "":
        return {
            "ocr_input":           ocr_text,
            "detected_brand":      _normalise_brand(brand),
            "candidate_pool_size": 0,
            "filter_case":         "EMPTY_INPUT",
            "candidates":          [],
        }

    norm_brand       = _normalise_brand(brand)
    filtered_indices = get_brand_filtered_indices(retriever.records, brand)

    if norm_brand is None:
        filter_case = "B"
    elif len(filtered_indices) < len(retriever.records):
        filter_case = "A"
    else:
        filter_case = "C"

    query_vec  = encode_query(ocr_text, retriever.model)
    candidates = retrieve_top_k(
        query_vec, retriever.master_matrix, filtered_indices, retriever.records, top_k
    )

    return {
        "ocr_input":           ocr_text,
        "detected_brand":      norm_brand,
        "candidate_pool_size": len(filtered_indices),
        "filter_case":         filter_case,
        "candidates":          candidates,
    }


def retrieve_batch(
    retriever:  "Retriever",
    queries:    List[Dict],
    top_k:      int = TOP_K,
    batch_size: int = ENCODE_BATCH_SIZE,
) -> List[Dict]:
    """
    Batched retrieval for multiple OCR queries.

    queries: list of {"ocr_text": str, "brand": str|None}
    Returns one result dict per query (same schema as retrieve()).
    """
    if not queries:
        return []

    texts          = [str(q.get("ocr_text", "")).strip() for q in queries]
    non_empty_mask = [bool(t) for t in texts]
    ne_texts       = [t for t, ok in zip(texts, non_empty_mask) if ok]

    encoded    = (encode_queries_batch(ne_texts, retriever.model, batch_size)
                  if ne_texts else np.zeros((0, EXPECTED_DIM), dtype=np.float32))
    results    = []
    enc_cursor = 0

    for i, q in enumerate(queries):
        brand = q.get("brand", None)
        if not non_empty_mask[i]:
            results.append({
                "ocr_input": texts[i], "detected_brand": _normalise_brand(brand),
                "candidate_pool_size": 0, "filter_case": "EMPTY_INPUT", "candidates": [],
            })
            continue

        norm_brand       = _normalise_brand(brand)
        filtered_indices = get_brand_filtered_indices(retriever.records, brand)
        filter_case      = ("B" if norm_brand is None
                            else "A" if len(filtered_indices) < len(retriever.records)
                            else "C")
        query_vec        = encoded[enc_cursor].astype(np.float32)
        enc_cursor      += 1
        candidates       = retrieve_top_k(
            query_vec, retriever.master_matrix, filtered_indices, retriever.records, top_k
        )
        results.append({
            "ocr_input": texts[i], "detected_brand": norm_brand,
            "candidate_pool_size": len(filtered_indices),
            "filter_case": filter_case, "candidates": candidates,
        })

    return results


# ============================================================================
# OUTPUT FORMATTING
# ============================================================================

def _format_result(result: Dict) -> str:
    lines    = []
    brand_s  = result["detected_brand"] or "(none)"
    fc_label = {
        "A": "brand-filtered",
        "B": "full master (no brand)",
        "C": "full master (brand not in master)",
        "EMPTY_INPUT": "empty input",
    }.get(result["filter_case"], result["filter_case"])

    lines.append("-" * 72)
    lines.append(f"  Query : {result['ocr_input']!r}")
    lines.append(f"  Brand : {brand_s}  [Case {result['filter_case']}: {fc_label}]")
    lines.append(f"  Pool  : {result['candidate_pool_size']} records eligible")
    lines.append("")

    if not result["candidates"]:
        lines.append("  (no candidates)")
    else:
        lines.append(f"  {'Rank':<5} {'Score':>7}  {'PRODUCT_CODE':<14}  product_name")
        lines.append(f"  {'-----':<5} {'-------':>7}  {'-'*14:<14}  {'-'*40}")
        for c in result["candidates"]:
            lines.append(
                f"  {c['rank']:<5} {c['similarity_score']:>7.4f}  "
                f"{c['PRODUCT_CODE']:<14}  {c['product_name']}"
            )
    return "\n".join(lines)


# ============================================================================
# BUILT-IN TEST SUITE
# ============================================================================

TEST_QUERIES = [
    {"ocr_text": "GALACT POW 200GM 1*200G",                       "brand": "GALACT"},
    {"ocr_text": "CARDACE AM 2.25/5MG 1X10 PCS",                  "brand": "CARDACE"},
    {"ocr_text": "OROFER XT+ SYR 200 ML",                         "brand": "OROFER"},
    {"ocr_text": "OSTERI INJ.600MCG/2.4 2.4ML",                   "brand": "OSTERI"},
    {"ocr_text": "GALACT 1X30 CAP 30 CAP",                        "brand": "GALACT"},
    {"ocr_text": "PROXYM-ER 300MG 15 TAB PLAN 15 TAB",            "brand": "PROXYM"},
    {"ocr_text": "GALACT LACTATION GRANULES 400G CHOCOLATE FLAV",  "brand": "GALACT"},
    {"ocr_text": "GALACT LACTATION GRANULES 400G ELAICHI FLAV",   "brand": "GALACT"},
    {"ocr_text": "GALACT LACTATION GRANULES 400G KESAR FLV",      "brand": "GALACT"},
    {"ocr_text": "POVIZTRA FLXT 05MG 1X1",                        "brand": "POVIZTRA"},
    {"ocr_text": "FLAWLIZO MICRO 0.01 GEL P 15GM",                "brand": "FLAWLIZO"},
    {"ocr_text": "GALACT POW 200GM 1*200G",                       "brand": None},
    {"ocr_text": "CARDACE AM 2.25/5MG 1X10 PCS",                  "brand": "pattern-inferred"},
]


def run_test_suite(retriever: "Retriever") -> None:
    print("\n" + "=" * 72)
    print("  transformer.py -- built-in test suite")
    print(f"  Model  : {MODEL_NAME}")
    print(f"  Master : {MASTER_XLSX_PATH}")
    print(f"  Records: {retriever.n_master}    TOP_K={TOP_K}")
    print("=" * 72)

    t_total = time.perf_counter()
    for q in TEST_QUERIES:
        t0     = time.perf_counter()
        result = retrieve(retriever, q["ocr_text"], brand=q.get("brand"))
        ms     = (time.perf_counter() - t0) * 1000
        print(_format_result(result))
        print(f"  (retrieved in {ms:.1f} ms)\n")

    total_s = time.perf_counter() - t_total
    print("=" * 72)
    print(
        f"  {len(TEST_QUERIES)} queries in {total_s:.2f}s "
        f"({total_s / len(TEST_QUERIES) * 1000:.1f} ms/query avg)\n"
    )


# ============================================================================
# CLI
# ============================================================================

def _build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="transformer.py",
        description=(
            "Sentence Transformer retrieval stage -- pharmaceutical product mapping.\n\n"
            "Examples:\n"
            '  python transformer.py --rebuild\n'
            '  python transformer.py --query "GALACT POW 200GM" --brand "GALACT"\n'
            '  python transformer.py --query "CARDACE AM 2.25/5MG"\n'
            "  python transformer.py                    # run built-in test suite"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--rebuild", action="store_true",
                   help="Regenerate master embeddings from PRODUCT_MASTER.xlsx.")
    p.add_argument("--query", type=str, default=None, metavar="OCR_TEXT",
                   help="A single OCR product name to retrieve candidates for.")
    p.add_argument("--brand", type=str, default=None, metavar="BRAND",
                   help="Brand hint for --query (from garbage_check.py col C).")
    p.add_argument("--top-k", type=int, default=TOP_K, metavar="K",
                   help=f"Candidates to return (default: {TOP_K}).")
    p.add_argument("--batch-size", type=int, default=ENCODE_BATCH_SIZE, metavar="N",
                   help=f"Encoding batch size (default: {ENCODE_BATCH_SIZE}).")
    p.add_argument("--device", type=str, default=None, metavar="DEVICE",
                   help="Device override: 'cpu' or 'cuda'. Default: auto-detect.")
    p.add_argument("--master", type=str, default=MASTER_XLSX_PATH, metavar="PATH",
                   help=f"Path to PRODUCT_MASTER.xlsx.")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    """CLI entry point. Exit 0 = success, 1 = error."""
    parser = _build_arg_parser()
    args   = parser.parse_args(argv)

    try:
        if args.rebuild:
            matrix, records, _model = rebuild_master_embeddings(
                model_name=MODEL_NAME,
                master_path=args.master,
                batch_size=args.batch_size,
                device=args.device,
            )
            print(
                f"\nRebuild complete.\n"
                f"  Records  : {len(records)}\n"
                f"  Dimension: {matrix.shape[1]}\n"
                f"  Model    : {MODEL_NAME}\n"
                f"  Stored in: {_embeddings_dir()}\n"
            )
            return 0

        if args.query:
            retriever = load_retriever(MODEL_NAME, args.master, args.device)
            result = retrieve(retriever, args.query, brand=args.brand, top_k=args.top_k)
            print(_format_result(result))
            return 0

        retriever = load_retriever(MODEL_NAME, args.master, args.device)
        run_test_suite(retriever)
        return 0

    except (FileNotFoundError, ValueError) as exc:
        logger.error("%s", exc)
        return 1
    except RuntimeError as exc:
        logger.error("Runtime error: %s", exc)
        return 1
    except KeyboardInterrupt:
        print("\nInterrupted.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
