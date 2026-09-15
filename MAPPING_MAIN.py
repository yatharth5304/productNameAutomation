"""
MAPPING_MAIN.py — combined garbage-check + product-mapping pipeline.

ONE pass over the input workbook produces ONE output workbook with three columns:

    Input | Remark | Product Code

Flow
----
    input row
       -> garbage-check classification (fully LOCAL, no LLM)
            garbage        -> GARBAGE        / 1
            review         -> REVIEW         / 2
            maybe product  -> NO_SUGGESTION  / 0
            confirmed product (remark "0") -> continue, carrying its matched brand
       -> product-mapping pipeline (local Steps 1-5 + ranking)
            local mapping succeeded          -> <product name> / <9-digit code>
            NO_CLEAR_MATCH                   -> NO_SUGGESTION  / 0
            decision would need the reranker -> NO_SUGGESTION  / 0

Provenance
----------
Every classification and mapping function below is transplanted verbatim from
`garbage_check.py` and `mapping.py`. Those two files are NOT imported and NOT
modified; this script is self-contained so it can be run and versioned on its own.

Two things are deliberately different from the originals, both documented in
`context_combined.md`:

1. Brand detection comes from the garbage-check stage ONLY. `mapping.py`'s own
   independent brand matching (`find_best_brand_for_input`, `find_potential_brands`,
   `group_related_brands`, `find_abbrev_promoted_brand`, `suggest_nearest_brands`,
   `nearest_brand_within_edits`) is omitted, and the branches that called it now
   return the pipeline's existing NO_BRAND / NO_CLEAR_MATCH result.
2. No LLM call can occur. The reranker code is retained but commented out, and any
   row whose final code would have come from the reranker is reported as
   NO_SUGGESTION / 0 rather than falling back to an unreranked local candidate.

Accuracy rule this file is built around: a NO_SUGGESTION is acceptable, a wrong
product code is not.
"""

import difflib
import os
import re
import sys
import time
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook, load_workbook
from rapidfuzz import fuzz, process
from rapidfuzz.distance import Levenshtein, JaroWinkler

warnings.filterwarnings("ignore")


# ============================================================================
# CONFIGURATION
# ----------------------------------------------------------------------------
# The four names that collide between the two source files (PROCESSING_MODE,
# VALID_PROCESSING_MODES, ROW_LIMIT, llm_enabled) are declared ONCE here.
# No API key, endpoint or model credential is stored in this file.
# ============================================================================

# Input: the workbook produced by the garbage-check stage's own source workbook.
# Only column A (the raw OCR row) is read; nothing in it is written back.
INPUT_XLSX_PATH  = r"f:\Vintyaa\projects\Product Name Automation\test1.xlsx"
INPUT_SHEET_NAME = "garbage_check"

# Output: a NEW workbook, one sheet, exactly three columns. Never overwrites the input.
OUTPUT_XLSX_PATH   = r"f:\Vintyaa\projects\Product Name Automation\MAPPING_MAIN_OUTPUT.xlsx"
OUTPUT_SHEET_NAME  = "mapping_main"
OUTPUT_COLUMNS     = ("Input", "Remark", "Product Code")

# Master product data and the brand vocabulary (read-only).
MASTER_XLSX_PATH = r"f:\Vintyaa\projects\Product Name Automation\PRODUCT_MASTER1.xlsx"
BRANDS_FILE      = r"f:\Vintyaa\projects\Product Name Automation\Brand Names.txt"

# Process only this many input rows (None = all rows).
ROW_LIMIT = None

# Single control point for LLM usage in this file.
#   "local" -> local logic only; no LLM/API call is made at any point
#   "llm" / "both" -> would enable the reranker, which is COMMENTED OUT here
# MAPPING_MAIN ships as "local" by design: see the reranker note below.
PROCESSING_MODE = "local"
VALID_PROCESSING_MODES = ("llm", "local", "both")
if PROCESSING_MODE not in VALID_PROCESSING_MODES:
    raise ValueError(
        f"PROCESSING_MODE must be one of {VALID_PROCESSING_MODES}, "
        f"got {PROCESSING_MODE!r}"
    )


def llm_enabled() -> bool:
    """True when PROCESSING_MODE would allow the reranker API to be called.

    Read at call time, not captured at import, so a harness can set the mode on
    the module. While the reranker body is commented out this must stay False:
    see MAPPING_LLM_PROMPT and RERANKER_DEPENDENT_STATUSES below.
    """
    return PROCESSING_MODE in ("llm", "both")


# The reranker prompt is intentionally BLANK. No prompt is configured because no
# LLM call is made in this build; the original prompt scaffolding is preserved
# (commented out) next to call_groq_llm so it can be restored later.
MAPPING_LLM_PROMPT = ""

# Sentinel product codes. Every real master code is 9 digits (410000003-491110205),
# so 0/1/2 can never collide with a genuine product code.
CODE_GARBAGE       = 1     # remark GARBAGE
CODE_REVIEW        = 2     # remark REVIEW
CODE_NO_SUGGESTION = 0     # remark NO_SUGGESTION
REMARK_GARBAGE        = "GARBAGE"
REMARK_REVIEW         = "REVIEW"
REMARK_NO_SUGGESTION  = "NO_SUGGESTION"


# ==========================================================================
# SECTION 1: GARBAGE-CHECK STAGE (local classification)
# Transplanted verbatim from garbage_check.py (logic unchanged).
# ==========================================================================


SHORT_BRANDS = {"ETS", "G3N", "ICL", "IKA", "NDS", "OXA", "PRX", "REE", "VIL"}
LOCAL_REVIEW_BRANDS = set()  # currently no brands require LLM review
# Brands where only exact matches are accepted; fuzzy matches are invalidated.
# EFCURE/KEMCURE/CTAX are too close to common OCR noise to trust a fuzzy hit.
# NUNIT: fuzzy hits risk false positives against common tokens like UNIT.
EXACT_ONLY_BRANDS = {"EFCURE", "KEMCURE", "CTAX", "NUNIT","EMNU","ACEM"}

EXACT_GARBAGE_ROWS = {
    "",
    "--",
    "++",
    "++--",
    "TOTAL",
    "0",
    "0 0",
    "-- --",
    "-- ++",
    "TOTAL --",
    "EMPTY",
    "++ ++",
    "TOTAL (VALUE IN RS.) --",
    "VALUE IN RS.",
    "EMCURE PHARMACEUTICALS LTD",
    "EMCURE PHARMACEUTICALS LTD *",
    "END OF REPORT TOTAL",
    "TOTAL QUANTITY",
    "TOTAL VALUE",
    "TOTALS",
    "TOTALS --",
    "-",
    "QUANTITY",
    "TOTAL VALUE --",
    "LAST MONTH SALE",
    "STOCK & SALES ANALYSIS",
    "EMCURE",
    "GRAND TOTAL --",
    "CONTINUED",
    "EMCURE INVENTIA",
    "EMCURE NUCRON",
    "EMCURE CD",
    "SUPPLIER NAME",
    "QUANTITY --",
    "PAGE NO.",
    "COMPANY TOTAL --",
    "TOTAL SALES IN THIS PERIOD --",
    "TOTAL CLOSING STOCK (SALES VALUE)",
    "DIVISION : 00 --",
    "TOTAL VALUE (00) --",
    "LAST MONTH SALE FEBRUARY",
    "EMCURE NUSURGE",
    "GRAND TOTAL",
    "PARTICULARS PKG.",
    "TOTAL VALUE --",
    "COMPANY TOTAL --",
    "TOTAL 0",
    "CHEMICO MEDICAL AGENCIES",
    "INDIA LIMITED *",
    "PHARMACUTICALS LTD. *",
    "PHARMACUTICALS LTD.",
    "PURCHASE VALUE --",
    "PAGE",
    "INDIA LIMITED",
    "AMOUNT TOTAL",
    "QTY TOTAL",
    "PURCHASE VALUE 0.00 CLOSING",
    "OUR SOFTWARE MARG ERP 9880074116080233041179880427548 --",
    "EMCURE IMPETUS",
    "EMCURE CD --",
    "GRANDTOTAL",
    "GRANDTOTAL --",
    "COMPANY EMCURE PHARMACEUTICA --",
    "SALE VALUE --",
    "OPENING VALUE --",
    "** LIQUDATION IS BASED ON LAST THREE MONTHS SALES --",
    "EMCURE INVENTIA --",
    "EMCURE XENNEX --",
    "THIS PDF REPORT IS CREATED FROM MEDICA ULTIMATE. FOR SOFTWARE ENQUIRY CONTACT +91-022-47474747 9750000648/658 9702074265 --",
    "THIS PDF REPORT IS CREATED FROM MEDICA ULTIMATE. FOR SOFTWARE ENQUIRY CONTACT 91-022-47474747 9750000648/658 9702074265 --",
    "TOTAL QUANTITY --",
    "POWERED BY SWILERP FOR RETAIL DISTRIBUTION & CHAIN STORES --",
    "EMCURE (XENNEX)",
    "EMCURE PHARMACEUTICALS",
    "PURCHASE DETAIL :-",
    "SUPPLIER NAME INVOICE",
    "***** --",
    "TOTAL:",
    "#NAME?",
    "RATE .01",
    "VALUE",
    "EMCURE XENNEX",
    "EMCURE PHARMA --",
    "EMCURE PHARMA --",
    "LAST MONTH",
    "EMCURE PHARMACEUTICALS LTD.",
    "BILL NOS. --",
    "EMCURE NUCRON --",
    "GROUP WISE",
    "PRODUCT NAME STRENGTH",
    "ASHIRWAD ENTERPRISES",
    "EMCURE NUSURGE --",
    "EMCURE --",
    "RS.",
    "VALUE 0.00",
    "MANUFACTURER GROUP: --",
    "VALUE IN",
    "* *",
    "OPENING VALUE --",
    "** - > NOT SOLD FOR 180 DAYS * -> NOT SOLD FOR 90 DAYS #-> 90 DAYS NEAR EXPIRY STOCK. & -> UNSUPPLIED --",
    "DIVISION TOTAL --",
    "GST --",
    "SUB TOTAL --",
    "DSTK : DAYS STOCK ON 3 MONTH SALE : -2 NO SALE NO STOCK -1 NO SALE 0 NO STOCK REST DAYS STOCK# --",
    "NEW PAGE STARTS HERE ++",
    "NEW PAGE STARTS HERE 0",
    "PRODUCT DETAILS",
    "NEW PAGE STARTS HERE",
    "NEW PAGE STARTS HERE --",
    "ITEM DESCRIPTION BLANK_HEADER1",
    "ITEM DESCRIPTION",
    "ITEM NAME UNIT",
    "ITEM NAME ++",
    "ITEM PACK",
    "ITEM NAME PACK",
    "ITEM NAME",
    "ITEM DESCRIPTION COL2",
    "ITEM DESCRIPTION RATE",
    "ITEM",
    "ITEMS PACKING",
    "ITEM NAME UOM",
    "ITEM NAME PACKG",
    "ITEM DESCRIPTION OPENING STOCK (1)",
    "ITEM NAME PACK SCM",
    "NAME OF ITEM",
    "NO/PRODUCT NAME UNIT",
    "PRODUCT NAME PACKING",
    "PRODUCT NAME PACKING",
    "PRODUCT & PACK",
    "PRODUCT PACK",
    "PRODUCT PKG",
    "PRODUCT DESCRIPTION PACKING",
    "PRODUCT NAME PACK",
    "NON MOVING PRODUCTS ABOVE 30 DAYS --",
    "PRODUCT NAME UNIT",
    "PRODUCT NAME",
    "SRNO. PRODUCT",
    "PRODUCTNAME PACK",
    "PRODUCT NAME BLANK_HEADER1",
    "NON PRODUCTS ABOVE 30 BLANK_HEADER1",
    "PRODUCT NAME PACKG",
    "PRODUCTNAME PACK OP",
    "PRODUCT OPENING",
    "PRODUCT NAME AND PACK",
    "PRODUCT PACKING",
    "PRODUCT DESRIPTION",
    "PRODUCT NAME CI VAL",
    "NON MOVING PRODUCTS ABOVE 90 DAYS",
    "PARTICULARS PACK",
    "COL1 COL2",
    "ID",
    "ID --",
    "JHARNA MEDICAL DISTRIBUTOR",
    "PRODUCT NAME SHELF PACKING",
    "PAGE NO.1 EMCURE-CD PRODUCT NAME SHELF PACKING ID",
    "CATEGORY TOTAL --",
    # Company / distributor / agency names confirmed as garbage from test.xlsx analysis
    "ZUVENTUS LIFESTYLE",
    "SS GENNOVA BIOPHARMA LTD",
    "SS GENNOVA BIOPHARMA LTD (INFIUS)",
    "SS GENNOVA BIOPHARMA LTD (INFIOUS)",
    "BHARAT MEDICAL STORES",
    "HEALTHWAYS AGENCIES",
    "UMA MEDICAL AGENCIES",
    "JAI MEDICAL TRADERS",
    "SHUBHAM MEDICAL AGENCIES",
    "MANGLA MEDICAL STORE",
    "SINGLA MEDICAL AGENCIES",
    "INDRALOK MEDICAL AGENCY",
    "JAGDAMBA MEDICAL AGENCY",
    "SOVA AGENCY",
    "GUPTA MEDICAL AGENCIES",
    "BANKEY BIHARI DRUG DISTRIBUTORS",
    "VINAY ENTERPRISES",
    "SANJAY DRUG AGENCIES",
    "SURESH MEDICAL AGENCIES",
    "KOMAL AGENCIES",
    "AGGARWAL MEDICAL STORE",
    "GENNOVA BIOPHARMACEUTICALS LTD",
    "GENNOVA BIOPHARMACEUTICALS LTD.",
    "GENNOVA BIOPHARMA LTD",
    "GENNOVA BIOPHARMA LTD.",
    "GENNOVA BIOPHARMA PVT LTD",
    "GENNOVA BIOPHARMACETICALS LTD",
    "GENNOVA BIOPHARMACUTICALS",
    "GENNOVA BIOPHARMA",
    "GENNOVA TRANSPLANT",
    "PROTECTION HEALTHCARE",
    "NUCRON PHARMA",
    "ZEE PHARMA",
    "SAB PHARMA",
    "SANJIVANI MEDICOS",
    "MUKESH MEDICOS",
    "D.N.DRUG DISTRIBUTORS",
    "COMPANY EMCURE CV DIV",
    "COMPANY EMCURE DERMA",
    "COMPANY SANOFI ORION DIV",
    "EMCUTIX DIVISION",
    "INTAS ALECTA DIVISION",
    "NUCRON PHARMA EMCURE -NUCRO",
    "EMCURE PHARAMA LTD -INVENSIA",
    "LONE MEDICAL AGENCY",
    "KOCHHAR MEDICAL AGENCIES",
    "ARJUN MEDICOSE",
    "H.M SALES CORPORATION",
    "SANOFI EMCURE CV",
    "SANOFI ORION",
    "AVENTIS DIABETES",
    "AVENTIS EMCURE",
    "DISTRIBUTION CHAIN STORES",
    "DEFAULT",
    "All Marketing Groups",
}

SAFE_GARBAGE_PATTERNS = (
    # Original patterns
    re.compile(r"^MANUFACTURER\s+\d+$", re.IGNORECASE),
    re.compile(r"^QUANTITY\s+\d+(?:\.\d+)?$", re.IGNORECASE),
    re.compile(r"^VALUE\s+\d+(?:\.\d+)?$", re.IGNORECASE),
    re.compile(r"^NO\s*/\s*PRODUCT\s+NAME\s+UNIT$", re.IGNORECASE),
    re.compile(r"^PAGE\s+NO\.?\s*\d+.*PRODUCT\s+NAME.*PACKING.*ID$", re.IGNORECASE),
    # Invoice / bill reference lines
    re.compile(r"^Bill", re.IGNORECASE),   # Bill Nos., Bills:, Bills --, Bills
    re.compile(r"^[A-Z]{4}\d{6,}\s*/\s*\d{2}/\d{2}/\d{4}"),
    re.compile(r"^(EIAB|EIPU|ZIMU|EIMU)\d{4,}\s+Dt\.", re.IGNORECASE),
    # Division / company / supplier structural rows
    re.compile(r"^Division\s*:", re.IGNORECASE),
    re.compile(r"^Division\s+\w", re.IGNORECASE),  # Division EMCURE ORION (no colon variant)
    re.compile(r"^Total\s*[:(]", re.IGNORECASE),   # Total (G1), Total :(Opening Val
    re.compile(r"^Total\s+Value\s*\(", re.IGNORECASE),  # Total Value (EMCURE) --, Total Value (CV) --
    re.compile(r"^Company\s*(Group)?\s*:", re.IGNORECASE),  # Company:, COMPANY GROUP:
    re.compile(r"^Com\s*:\s*", re.IGNORECASE),       # Com : EMCURE (PHARMA)
    re.compile(r"^Company\s+\w", re.IGNORECASE),   # Company EMCURE ORION (no colon variant)
    re.compile(r"^Primary Company\s+", re.IGNORECASE),
    re.compile(r"^Company Name\s+", re.IGNORECASE),
    re.compile(r"^Supplier\s+\w", re.IGNORECASE),
    re.compile(r"^Store\s+\w", re.IGNORECASE),
    re.compile(r"^Manufacturer\s+(Group|Name)\s*:", re.IGNORECASE),
    re.compile(r"^For Manufacturer Group\s*:", re.IGNORECASE),
    re.compile(r"^(Firm|Group|Com)[\s\w]*Total", re.IGNORECASE),
    re.compile(r"^COMNAME\s+TOTAL", re.IGNORECASE),
    re.compile(r"^ANNEXURE\s+\d", re.IGNORECASE),
    # Software-generated / system rows
    re.compile(r"^Generated at\s+\d{4}-\d{2}-\d{2}", re.IGNORECASE),
    re.compile(r"^Prepared by\s+\w+\s+on\s+\d", re.IGNORECASE),
    re.compile(r"^Print Date\s+\d", re.IGNORECASE),
    re.compile(r"^Print\s+HEALTH\s+QRCODE", re.IGNORECASE),
    re.compile(r"^Powered\s+By:", re.IGNORECASE),
    re.compile(r"^Medica Ultimate", re.IGNORECASE),
    re.compile(r"^Report End\)", re.IGNORECASE),
    re.compile(r"^Printed on", re.IGNORECASE),
    # HTML artifact rows (OCR of web-generated PDFs)
    re.compile(r"^hr/>"),
    re.compile(r"^u>"),
    re.compile(r"^span\s+style", re.IGNORECASE),
    # Financial summary / aggregate rows
    re.compile(r"^PURC\s*:", re.IGNORECASE),
    re.compile(r"^EXP\d+M\s*:", re.IGNORECASE),
    re.compile(r"^(Sales|Opening|Closing|Open|Close)\s+(Stock|Val|Value)(s)?\s*[:\-]?", re.IGNORECASE),
    re.compile(r"^(Opening|Closing)\s+Val\.", re.IGNORECASE),
    re.compile(r"^Values?\s*:\s*[\d.]", re.IGNORECASE),
    re.compile(r"^(TAX AMOUNT|COMPANY NAME:)", re.IGNORECASE),
    re.compile(r"^(UC SALE|CL\.STK\.|MR\.Balance|Op\.Val\.|Pur\.Val\.)", re.IGNORECASE),
    re.compile(r"^(Prev Month Tot|Value Age)", re.IGNORECASE),
    re.compile(r"^STOCK\s*&\s*SALES STATEMENT", re.IGNORECASE),
    re.compile(r"^Purchase Return", re.IGNORECASE),
    re.compile(r"^\w+\)\s+(Sale|Purchase|Stock|Return)\s+Value", re.IGNORECASE),  # Jun) Sale Value, May) Sale Value
    # Report footer / legend rows
    re.compile(r"^Non Moving Since\s+\d+\s+Days?", re.IGNORECASE),
    re.compile(r"^Sales and Last Month\s+sales based on", re.IGNORECASE),
    re.compile(r"^Last Month\s+sales between", re.IGNORECASE),
    re.compile(r"^Purchase bill included values", re.IGNORECASE),
    re.compile(r"^(PENDING DEBIT NOTES|NO SUGGESTION|SHEDULED DRUG)$", re.IGNORECASE),
    # Distributor / agency / medical store rows (formerly MAYBE_GARBAGE, promoted to direct garbage)
    re.compile(
        r"\b(AGENCIES?|DISTRIBUTORS?|MEDICO[SE]?|CORPORATION|ENTERPRISES?|TRADERS?|TRADING)\b",
        re.IGNORECASE,
    ),
    re.compile(r"\b(PVT\.?\s+)?LTD\.?\s*$", re.IGNORECASE),
    re.compile(r"\bLIMITED\.?\s*$", re.IGNORECASE),
    re.compile(r"^EIKO\d{4,}", re.IGNORECASE),
    re.compile(r"^GSTIN\s*:", re.IGNORECASE),
    re.compile(r"^GST\s*NO\.?\s*:", re.IGNORECASE),
    re.compile(r"^Op\.\s+Amt", re.IGNORECASE),
    re.compile(r"^O\.B\.\s+PUR\.VAL", re.IGNORECASE),
    re.compile(r"^Items\s+\d+\s*$", re.IGNORECASE),
    re.compile(
        r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(Sal|Sale|Stock|Pur)\b",
        re.IGNORECASE,
    ),
    re.compile(r"^Copyright\s+\d{4}", re.IGNORECASE),
    re.compile(r"^Powered by\s", re.IGNORECASE),
    re.compile(r"\bcolumn_\d+\b", re.IGNORECASE),
    re.compile(r"^TEST\s+ITEM\b", re.IGNORECASE),
    re.compile(r"^[A-Z]{4}\d{4,}\s+Dt\.\d{2}/\d{2}/\d{4}", re.IGNORECASE),
    re.compile(r"^Manufacturer\s+\d+\s*~", re.IGNORECASE),
    # MARG ERP software advertisement rows (370+ occurrences with varying phone numbers)
    re.compile(r"^MARG\s+ERP\b", re.IGNORECASE),
    # SANOFI rows — Sanofi is not an Emcure brand; all SANOFI rows in this dataset
    # are company headers, annexure labels, or cross-company division lines.
    re.compile(r"\bSANOFI\b", re.IGNORECASE),
    # OCR garbage: rows consisting entirely of repeated TAB / TABLE tokens
    re.compile(r"^\s*(TAB(LE)?\s+)*TAB(LE)?\s*$", re.IGNORECASE),
    # Non-pharma physical item rows (bedsheets, bags, utensils, food)
    # DUFFAL matched as standalone word — BAG22MED suffix breaks \b after BAG so match on DUFFAL alone.
    re.compile(r"\b(BEDSHEET|BEDSEET|HOTPOT|TOWEL|DUFFAL|ELETRAL|WATER\s+JAR|CHOCOLATE)\b", re.IGNORECASE),
    # OCR corruptions of non-pharma items (STELL BOWAL = STEEL BOWL, STEEL GLASS)
    re.compile(r"\bSTELL?\s+BOW[AL]+\b", re.IGNORECASE),
    re.compile(r"\bSTEEL\s+GLASS\b", re.IGNORECASE),
    # Column-header artifact from OCR of table headers
    re.compile(r"^Item\s+Name\b", re.IGNORECASE),
    # Software licence rows
    re.compile(r"\bSOFTWARE\s+LICEN", re.IGNORECASE),
    # Receipt / near-expiry report lines
    re.compile(r"^Receipt\s+Value\b", re.IGNORECASE),
    re.compile(r"^Near\s+Expiry\s+Product\b", re.IGNORECASE),
    # Medical store reference / receipt lines (e.g. SAHIL MEDICAL STOR FEROZEPU A001335)
    re.compile(r"\bMEDICAL\s+STOR\b", re.IGNORECASE),
    # Rows with 3 or fewer alphanumeric characters are too short to be a product
    # (e.g. hr/, CV, AV, HO, PM --, MG, GM, EMK, EMO, EMX).
    # Brand matching runs before this check, so 3-char brands (VIL, IKA, etc.) in the
    # master are already classified as 0 before reaching pattern garbage.
    re.compile(r"^[^A-Za-z0-9]*[A-Za-z0-9]{0,3}[^A-Za-z0-9]*$"),
)

# Vocabulary of document-structure / metadata words.
# A row whose every alphabetic token belongs to this set — with no local brand
# match — is structural garbage (a header, summary, total line, etc.).
# Medicine-context tokens (TAB, CAP, INJ, MG, ML, …) are deliberately EXCLUDED
# to avoid incorrectly flagging real product lines that carry those suffixes.
METADATA_GARBAGE_VOCAB = {
    # aggregations / totals
    "TOTAL", "SUBTOTAL", "GRAND", "GRANDTOTAL", "VALUE", "QUANTITY", "QTY",
    "AMOUNT", "NET", "GROSS", "SUM",
    # financial
    "MRP", "RATE", "DISCOUNT", "CREDIT", "DEBIT", "BALANCE",
    "PURCHASE", "SALE", "SALES", "INVOICE", "RS",
    # document / report structure
    "PAGE", "REPORT", "SUMMARY", "ANALYSIS", "GROUP", "CATEGORY",
    "DIVISION", "COMPANY", "SUPPLIER", "MANUFACTURER", "DISTRIBUTOR",
    "DATE", "MONTH", "YEAR", "PERIOD", "CONTINUED",
    # stock
    "STOCK", "CLOSING", "OPENING", "CARRY", "FORWARD",
    # column / header labels
    "PRODUCT", "ITEM", "NAME", "PACKING", "PACK", "UNIT", "UOM",
    "PACKG", "DETAIL", "DETAILS", "DESCRIPTION", "DESRIPTION",
    "CODE", "TYPE", "SRNO", "SR", "ID", "NO", "COL", "BLANK", "HEADER",
    "PARTICULARS", "PKG", "STRENGTH", "SCM",
    # movement / ageing
    "LAST", "CURRENT", "NEW", "MOVING", "NON", "ABOVE", "DAYS",
    # month names — needed for rows like "LAST MONTH SALE MAY"
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
    # common prepositions / conjunctions safe to include
    "IN", "OF", "BY", "THE", "AND", "FOR", "TO", "AT", "FROM",
    "THIS", "SUB", "WISE", "BRAND",
    # misc garbage-header words
    "STARTS", "HERE", "SECTION", "SCHEDULE", "SHELF",
}

# ---------------------------------------------------------------------------
# MAYBE_PRODUCT patterns
# A row matches MAYBE_PRODUCT when it contains a pharmaceutical dosage-form
# token AND a quantity/strength signal, AND does NOT contain any structural
# garbage keywords.  These rows are very likely real products that local brand
# matching failed to confirm; they are written as MAYBE_PRODUCT for review.
# ---------------------------------------------------------------------------

# Signal 1: a recognised pharmaceutical dosage form anywhere in the row.
MAYBE_PRODUCT_FORM_PATTERN = re.compile(
    r"\b(TABS?|TABLETS?|CAPS?|CAPSULES?|SYP|SYRUP|INJ|INJN|OINT|OINTMENT|CREAM|GEL|"
    r"DROPS?|LOTION|SUSP|SUSPENSION|SACHET?|SACH|POWDER|SPRAY|INHALER|PATCH|SYRINGE|"
    r"VIAL|AMP|AMPULE|AMPOULE|SOLUTION|SOLN|SYR|SUS|SOL|PFS|LOZENGE|"
    r"SUPPOSITORIES?|SUPPOS?|SUPPO|SUPP|ENEMA|EYE|EAR|NASAL|CHEWABLE|EFFERVESCENT)\b",
    re.IGNORECASE,
)

# Signal 2: a numeric strength, pack count, or volume.
# Includes * as a pack separator (e.g. 1*10) — widened from original MP1 pattern.
MAYBE_PRODUCT_QTY_PATTERN = re.compile(
    r"(\b\d+\.?\d*\s*(MG|MCG|IU|ML|GM|G)\b"      # strength: 500MG, 1.5ML, 100IU
    r"|\b\d+\s*[xX*]\s*\d+\b"                      # pack: 1X10, 10X10, 1*10
    r"|\b\d+['']?\s*[sS]\b"                         # count: 10S, 15s, 10'S, 15's
    r"|\b\d+\s*(TAB|TABS|CAP|CAPS|VIAL|AMP)\b"     # qty+form: 10TAB, 1VIAL
    r"|\b\d+\s*(ML|GM|G)\b"                         # volume/weight: 100ML, 30GM
    r"|\b\d+\s*MD\b"                                # dose unit: 500MD, 1MD
    r"|\bPCS\b)",                                   # pieces: PCS anywhere in row
    re.IGNORECASE,
)

# Guard: if any of these structural keywords are present the row is too risky
# to auto-promote (e.g. 'TOTAL CAL D3 TAB 1X20', 'DOXOLIN M TAB SALE DOM').
MAYBE_PRODUCT_GUARD_PATTERN = re.compile(
    r"\b(TOTAL|SALE|PURCHASE|INVOICE|STOCK|VALUE|REPORT|SUMMARY|IMPORT|CALL)\b",
    re.IGNORECASE,
)

MAYBE_PRODUCT_BRAND_PATTERN = re.compile(
    r"\b(DULCOFLEX)\b",
    re.IGNORECASE,
)


def load_brands():
    with open(BRANDS_FILE, "r", encoding="utf-8") as f:
        return re.findall(r'"([^"]+)"', f.read())


def normalize_text(value):
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())


def row_tokens(value):
    return re.findall(r"[A-Z0-9]+", str(value).upper())


def edit_distance_leq_one(a, b):
    """Return the edit type when Levenshtein distance is <= 1, else None."""
    if a == b:
        return "exact"
    if abs(len(a) - len(b)) > 1:
        return None

    if len(a) > len(b):
        a, b = b, a
        shorter_is_first = False
    else:
        shorter_is_first = True

    i = 0
    j = 0
    edits = 0
    mismatch_type = None
    while i < len(a) and j < len(b):
        if a[i] == b[j]:
            i += 1
            j += 1
            continue
        edits += 1
        if edits > 1:
            return None
        if len(a) == len(b):
            mismatch_type = "1-substitution"
            i += 1
            j += 1
        else:
            mismatch_type = "1-insertion" if shorter_is_first else "1-deletion"
            j += 1

    if j < len(b) or i < len(a):
        edits += 1
        if len(a) == len(b):
            mismatch_type = "1-substitution"
        else:
            mismatch_type = "1-insertion" if shorter_is_first else "1-deletion"

    return mismatch_type if edits <= 1 else None


def damerau_distance(a, b, max_dist=2):
    """Damerau-Levenshtein distance between a and b, capped at max_dist.
    Handles substitution, insertion, deletion, and adjacent transpositions.
    Returns the integer distance if <= max_dist, else None.
    Uses per-row early exit: if the minimum value in a completed DP row already
    exceeds max_dist, the final distance provably exceeds max_dist too.
    """
    if a == b:
        return 0
    la, lb = len(a), len(b)
    if abs(la - lb) > max_dist:
        return None

    # DP table; initialised to a value larger than any real distance we'd accept
    inf = la + lb + 1
    dp = [[inf] * (lb + 1) for _ in range(la + 1)]
    for i in range(la + 1):
        dp[i][0] = i
    for j in range(lb + 1):
        dp[0][j] = j

    for i in range(1, la + 1):
        row_min = inf
        for j in range(1, lb + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            dp[i][j] = min(
                dp[i - 1][j] + 1,        # deletion
                dp[i][j - 1] + 1,        # insertion
                dp[i - 1][j - 1] + cost, # substitution
            )
            # adjacent transposition
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                dp[i][j] = min(dp[i][j], dp[i - 2][j - 2] + 1)
            if dp[i][j] < row_min:
                row_min = dp[i][j]
        # Early exit: if no cell in this row is <= max_dist, no path can succeed
        if row_min > max_dist:
            return None

    dist = dp[la][lb]
    return dist if dist <= max_dist else None


def token_spans(name):
    """Generate normalized consecutive token spans across the full row."""
    tokens = row_tokens(name)
    spans = []
    for start in range(len(tokens)):
        combined = ""
        span_tokens = []
        for end in range(start, len(tokens)):
            span_tokens.append(tokens[end])
            combined += tokens[end]
            spans.append({
                "tokens": tuple(span_tokens),
                "text": " ".join(span_tokens),
                "norm": combined,
            })
    return spans


# Character sets of the normalized brand names, filled in by build_brand_index so
# the fuzzy pre-filter does not rebuild them for every (brand, span) pair.
# frozenset(brand_norm) is a pure function of brand_norm, so an entry stays
# correct no matter how many brands share a normalized form.
_BRAND_CHAR_SETS = {}


def build_brand_index(brands):
    index = [(brand, normalize_text(brand)) for brand in brands]
    for _, brand_norm in index:
        if brand_norm not in _BRAND_CHAR_SETS:
            _BRAND_CHAR_SETS[brand_norm] = frozenset(brand_norm)
    return index


def exact_product_match(name, brand_index, precomputed_spans=None):
    spans = precomputed_spans if precomputed_spans is not None else token_spans(name)
    if not spans:
        return None

    for brand, brand_norm in brand_index:
        if len(brand_norm) <= 3:
            for span in spans:
                if len(span["tokens"]) == 1 and span["norm"] == brand_norm:
                    return {
                        "brand": brand,
                        "matched_text": span["text"],
                        "match_type": "exact",
                    }
            continue
        for span in spans:
            if span["norm"] == brand_norm:
                return {
                    "brand": brand,
                    "matched_text": span["text"],
                    "match_type": "exact",
                }
    return None


def fuzzy_product_match(name, brand_index, precomputed_spans=None):
    spans = precomputed_spans if precomputed_spans is not None else token_spans(name)
    if not spans:
        return None

    # Lower rank = better quality match; 2-edit is least preferred
    rank = {
        "1-substitution": 0,
        "1-insertion": 1,
        "1-deletion": 2,
        "1-transposition": 3,
        "2-edit": 4,
    }
    best_match = None

    # One character set per span instead of one per (brand, span) pair. Parallel
    # to `spans`, so the span visitation order inside the brand loop is unchanged.
    span_sets = [frozenset(span["norm"]) for span in spans]

    for brand, brand_norm in brand_index:
        brand_len = len(brand_norm)
        if brand_len <= 3:
            continue  # 3-char brands handled exclusively by exact_product_match

        # Adaptive threshold: 2 edits only for long brands (>= 10 chars) where
        # accidental collisions with common English words are extremely unlikely.
        # Brands 4-9 chars stay at <= 1 edit to prevent false positives.
        max_edits = 2 if brand_len >= 10 else 1
        brand_set = _BRAND_CHAR_SETS[brand_norm]

        for span, span_set in zip(spans, span_sets):
            span_norm = span["norm"]
            span_len = len(span_norm)
            if abs(span_len - brand_len) > max_edits:
                continue

            # --- Character-set pre-filter (set ops, runs before the O(m*n) DP) ---
            # A character present in the span but absent from the brand has to be
            # removed by a deletion or a substitution -- transpositions only reorder
            # characters, they never remove one -- and distinct characters need
            # distinct edits. So len(span_set - brand_set) is a lower bound on the
            # Damerau-Levenshtein distance: a necessary condition, which means no
            # pair the exact distance checks below would accept is rejected here.
            if len(span_set - brand_set) > max_edits:
                continue

            # --- Try standard single-edit first for precise match type ---
            match_type = edit_distance_leq_one(span_norm, brand_norm)
            if match_type == "exact":
                continue  # exact hits are handled by exact_product_match

            if match_type is None:
                # Could still be a transposition (DL dist == 1) or 2-edit (for long brands)
                dist = damerau_distance(span_norm, brand_norm, max_dist=max_edits)
                if dist is None:
                    continue
                if dist == 1:
                    match_type = "1-transposition"
                elif dist == 2 and brand_len >= 10:
                    match_type = "2-edit"
                else:
                    continue

            candidate = {
                "brand": brand,
                "matched_text": span["text"],
                "match_type": match_type,
            }
            if best_match is None:
                best_match = candidate
                continue
            current_rank = rank.get(candidate["match_type"], 99)
            best_rank = rank.get(best_match["match_type"], 99)
            if current_rank < best_rank:
                best_match = candidate
                continue
            if current_rank == best_rank and len(span["norm"]) > len(normalize_text(best_match["matched_text"])):
                best_match = candidate

    return best_match


def build_garbage_set():
    garbage = {normalize_text(value) for value in EXACT_GARBAGE_ROWS}
    garbage.add("")
    return garbage


def is_exact_garbage_row(name, garbage_rows):
    return normalize_text(name) in garbage_rows


def is_pattern_garbage_row(name):
    text = str(name).strip()
    return any(pattern.match(text) for pattern in SAFE_GARBAGE_PATTERNS)


def is_metadata_garbage_row(name):
    """Return True when every alphabetic token in the row is a known
    document-structure / metadata word, indicating the row is a header,
    summary, or total line rather than a real product entry."""
    tokens = re.findall(r"[A-Z]+", str(name).upper())
    if not tokens:
        return False  # blank / numeric-only rows are handled elsewhere
    return all(t in METADATA_GARBAGE_VOCAB for t in tokens)


def is_maybe_product_row(name):
    """Return True when a row looks like a real pharmaceutical product line
    that local brand matching failed to confirm.

    Requires AT LEAST ONE of:
      - a dosage-form token (TAB, TABLET, SUPP, SYRINGE, INJ, CREAM, VIAL, …)
      - a quantity or strength signal (MG, ML, NxN, N*N, NS, N'S, NMD, …)
      - a known brand name that escaped brand detection (DULCOFLEX, …)
    AND:
      - absence of structural-garbage keywords (TOTAL, SALE, INVOICE, …)
    """
    if MAYBE_PRODUCT_GUARD_PATTERN.search(name):
        return False
    return (
        bool(MAYBE_PRODUCT_FORM_PATTERN.search(name))
        or bool(MAYBE_PRODUCT_QTY_PATTERN.search(name))
        or bool(MAYBE_PRODUCT_BRAND_PATTERN.search(name))
    )


def fuzzy_garbage_match(name, garbage_rows):
    row_norm = normalize_text(name)
    if not row_norm:
        return None

    best_match = None
    rank = {"exact": 0, "1-substitution": 1, "1-insertion": 2, "1-deletion": 3}

    for garbage_phrase in EXACT_GARBAGE_ROWS:
        garbage_norm = normalize_text(garbage_phrase)
        if not garbage_norm:
            continue
        if abs(len(row_norm) - len(garbage_norm)) > 1:
            continue
        match_type = edit_distance_leq_one(row_norm, garbage_norm)
        if not match_type or match_type == "exact":
            continue
        candidate = {
            "garbage_phrase": garbage_phrase,
            "match_type": match_type,
        }
        if best_match is None:
            best_match = candidate
            continue
        current_rank = rank[candidate["match_type"]]
        best_rank = rank[best_match["match_type"]]
        if current_rank < best_rank:
            best_match = candidate
            continue
        if current_rank == best_rank and len(garbage_norm) > len(normalize_text(best_match["garbage_phrase"])):
            best_match = candidate

    return best_match


# ==========================================================================
# SECTION 2: MAPPING STAGE (normalisation -> ranking)
# Transplanted verbatim from mapping.py (logic unchanged).
# ==========================================================================


# ==========================================================================
# RERANKER TRANSPORT -- PRESENT BUT DISABLED
#
# The reranker HTTP call from mapping.py, preserved verbatim and commented out.
# No API key or endpoint is copied into this file, so nothing here can fire.
# Re-enable by uncommenting this function and supplying HUGGINGFACE_API_KEY,
# HUGGINGFACE_API_URL, MAX_OUTPUT_TOKENS, MODEL_NAME and REQUEST_TIMEOUT in
# the CONFIG block (none of them are defined here).
# ==========================================================================
# def call_groq_llm(client, system_prompt: str, user_prompt: str, documents: list, verbose=False):
#     """POST one reranker prompt to the Hugging Face Inference Providers router.
#
#     The historical name is kept deliberately: existing A/B harnesses disable the
#     reranker by monkeypatching ``mapping.call_groq_llm``, and renaming it would
#     make those patches silently ineffective. ``client`` stays unused, as before.
#     """
#     headers = {
#         "Authorization": f"Bearer {HUGGINGFACE_API_KEY}",
#         "Content-Type": "application/json",
#     }
#     payload = {
#         "model": MODEL_NAME,
#         "messages": [
#             {"role": "system", "content": system_prompt},
#             {"role": "user", "content": user_prompt},
#         ],
#         "temperature": 0,
#         "max_tokens": MAX_OUTPUT_TOKENS,
#         "stream": False,
#     }
#     try:
#         resp = requests.post(HUGGINGFACE_API_URL, headers=headers, json=payload,
#                              timeout=REQUEST_TIMEOUT)
#         resp.raise_for_status()
#         data = resp.json()
#     except requests.RequestException as e:
#         detail = ""
#         if e.response is not None:
#             try:
#                 detail = f" — {e.response.text[:300]}"
#             except Exception:
#                 detail = ""
#         raise RuntimeError(f"Hugging Face request failed: {e}{detail}")
#     choices = data.get("choices", []) or []
#     if not choices:
#         raise RuntimeError(f"Empty Hugging Face response: {data}")
#     top_text = str(choices[0].get("message", {}).get("content", "")).strip()
#     usage = data.get("usage")
#     if verbose:
#         print(f"  Top text: {top_text}")
#     return top_text, usage


def call_groq_llm(*args, **kwargs):
    """Hard stop: the reranker is disabled in this pipeline.

    Kept as an active stub (rather than absent) so that any future code path
    that reaches for the reranker fails loudly instead of silently degrading
    into an unreranked guess.
    """
    raise RuntimeError(
        "Reranker disabled in MAPPING_MAIN.py (PROCESSING_MODE='local'). "
        "Rows needing the reranker are reported as NO_SUGGESTION / 0."
    )


INPUT_COLUMN  = "Input_Column"
OUTPUT_COLUMN = "Output_Column"
PRODUCT_CODE_COLUMN = "Product_Code"
# Diagnostic / review columns for NO_CLEAR_MATCH handling.
STATUS_COLUMN      = "Match_Status"          # why a row matched / failed
CONFIDENCE_COLUMN  = "Confidence"            # HIGH / MEDIUM / LOW / NONE
CANDIDATES_COLUMN  = "Candidates_To_Model"   # how many products fed to the model/ranker
SUGGESTIONS_COLUMN = "Suggestions"           # top-N nearest products for manual review
SOURCE_COLUMN      = "Source"                # "0" (confirmed product) or "MAYBE_PRODUCT" (auto-detected brand)

PRODUCT_COL   = "product_name"
BRAND_COL     = "BRAND_NAME"
PACK_SIZE_COL = "PACK_SIZE"
VARIANT_COL   = "VARIANT"
SUB_VARIANT_COL = "SUB_VARIANT"
PRODUCT_CODE_COL = "PRODUCT_CODE"

FUZZY_MIN_SCORE = 80       # FIX 10: relaxed from 85 to catch brand OCR typos
BRAND_NAME_NEXTSTEP_SCORE = 99
JW_MIN_SCORE = 0.85        # Jaro-Winkler minimum score for MAYBE_PRODUCT brand fallback (0.0-1.0 scale)
# Max character edits between the input's brand token and a fuzzy-matched master
# brand for that match to be accepted by brand DETECTION. A wider gap means a
# different brand / generic name (e.g. CETZINE→CETIRIZINE = 3 edits) → reject so
# the row becomes NO_CLEAR_MATCH instead of a confident wrong mapping.
BRAND_FUZZY_MAX_EDITS = 2
# Manufacturer/company names that show up as OCR noise in front of the real
# brand (e.g. 'EMCOO TAPICEM…', 'EMC00CLAMOB…'). This master is Emcure, so
# 'EMCURE' is the company, never a brand. Configure per-master if reused.
MANUFACTURER_NOISE_NAMES = {"EMCOO"}
MANUFACTURER_NOISE_MAX_EDITS = 3
MAX_OUTPUT_TOKENS = 512
TOP_N_RERANK = 3
DOSAGE_UNIT_SUFFIXES = {"MG", "GM", "ML", "MCG", "IU", "UNIT"}
FORM_HINT_TOKENS = {
    "TA", "TAB", "TABS", "TABLET", "TABLETS",
    "CAP", "CAPS", "CAPSULE", "CAPSULES",
    "DROP", "DROPS", "SYP", "SYRUP", "TONIC",
    "SUSPENSION", "INJ", "INJECTION", "CREAM",
    "GEL", "OINTMENT", "LOTION", "PLUS",
    "BAR", "BARS", "GRANUL", "GRANULES",
}

VARIANT_TOKEN_EQUIVALENTS = {
    "TA": "TABLET",
    "TAB": "TABLET",
    "TABS": "TABLET",
    "TABLETS": "TABLET",
    "CAP": "CAPSULE",
    "CAPS": "CAPSULE",
    "CAPSULES": "CAPSULE",
    "INJ": "INJECTION",
    "INJECTIONS": "INJECTION",
    "DROP": "DROPS",
    "SYP": "SYRUP",
    "BARS": "BAR",
    "GRANUL": "GRANULES",
    "LOT": "LOTION",
    "LO": "LOTION",
    "CRM": "CREAM",
    "CR": "CREAM",
}

INPUT_TOKEN_EQUIVALENTS = {
    "FORT": "FORTE",
    "FORTS": "FORTE",
}

GENERIC_VARIANT_MATCH_TOKENS = {
    "TABLET", "CAPSULE", "INJECTION", "SYRUP", "DROPS",
    "CREAM", "LOTION", "GEL", "OINTMENT", "SOLUTION",
    "SUSPENSION", "POWDER", "SACHET", "BAR", "GRANULES",
}

REQ_COUNT = 0
LLM_REQUEST_ATTEMPTS = 0
SUM_PROMPT_TOKENS = 0
SUM_COMPLETION_TOKENS = 0
SUM_TOTAL_TOKENS = 0
df = pd.DataFrame()

# =========================
# PROMPTS
# =========================
SYSTEM_PROMPT = (
    "Role:You are hard core thinker."
    "TASK:Given a product_name that contains variants, sub-variants, and pack sizes, identify the best matching product from a provided product list.All the products in the list belongs to same brand." \
    "You have to select one of the products from the list using your extremely capable reasoning"
)

USER_PROMPT_TEMPLATE = r"""
You are a pharma product normalization engine.
Your task is to match an INPUT PRODUCT NAME to exactly ONE master product name from the given BRAND CONTEXT.

IMPORTANT PACK SIZE RULES:
1. If INPUT has a pack size (e.g., 1X10, 1X15), look for EXACT pack size match first
2. If exact pack size NOT found, choose the product WITHOUT any pack size (base variant)
3. NEVER choose a different pack size - only exact match or base variant

--------------------
INPUT PRODUCT:
{input_name}

MASTER PRODUCTS (same brand only):
{brand_context}
--------------------

OUTPUT FORMAT (IMPORTANT):
- Output ONLY the exact master product name
- Do not explain anything
- Do not output multiple names
- If INPUT has pack size but exact match not found, prefer product WITHOUT [PACK:...] notation
"""

# =========================
# UTILS
# =========================
def norm(s: str) -> str:
    s = str(s or "").upper().strip()
    s = s.replace("+", " PLUS ")
    s = s.replace("-", " ")
    s = s.replace("/", " ")
    s = s.replace("*", " ")
    s = re.sub(r'\bX\b', ' ', s)
    s = s.replace("&", " AND ")
    s = re.sub(r'\b([A-Z])\.([A-Z])\b', r'\1\2', s)
    s = re.sub(r'\b([A-Z])\.([A-Z])\.([A-Z])\b', r'\1\2\3', s)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_variant_token(s: str) -> str:
    s = str(s or "").upper().replace("+", " PLUS ").strip()
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def compact_variant_token(s: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(s or "").upper().replace("+", " PLUS "))

def compact_brand_token(s: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(s or "").upper())

def _expand_strength_token(tok: str) -> str:
    """Expand a magnitude-suffixed strength token to its full numeric value.

    '1K' -> '1000', '1.5K' -> '1500', '2M' -> '2000000'. Any token that is not
    a bare number followed by a single K/M is returned unchanged — crucially
    this leaves dosage units like 'MG', 'MCG', 'ML' alone (they have no leading
    digit) so only true K/M magnitude abbreviations are rewritten.
    """
    m = re.fullmatch(r"(\d+(?:\.\d+)?)([KM])", str(tok or "").upper())
    if not m:
        return tok
    mult = 1000 if m.group(2) == "K" else 1_000_000
    return str(int(float(m.group(1)) * mult))

def _normalize_strength_words(words) -> set:
    """Apply _expand_strength_token across a set/iterable of tokens."""
    return {_expand_strength_token(w) for w in words}

def _input_compact_joins(input_text) -> set:
    """Compact joins of adjacent input tokens, so a brand/variant word written
    solid in the master ('D3') is recognised when the input splits it ('D 3' /
    'D-3' → 'D3'). Returns 2- and 3-token joins."""
    toks = re.findall(r'[A-Z0-9]+', str(input_text or "").upper())
    joins = set()
    for i in range(len(toks) - 1):
        joins.add(toks[i] + toks[i + 1])
    for i in range(len(toks) - 2):
        joins.add(toks[i] + toks[i + 1] + toks[i + 2])
    return joins

def _expand_strength_in_compact(s: str) -> str:
    """Expand K/M magnitude suffixes inside a space-stripped compact string.

    Only a digit run followed by K/M that is itself followed by another digit
    or end-of-string is expanded, so embedded units survive: in 'ENCICARB1K'
    the 'K' ends the string -> 'ENCICARB1000', but in 'ENCICARB1000MGINJ' the
    'M' is followed by 'G' (a letter) and is left untouched.
    """
    return re.sub(
        r"(\d+(?:\.\d+)?)([KM])(?=\d|$)",
        lambda m: str(int(float(m.group(1)) * (1000 if m.group(2) == "K" else 1_000_000))),
        str(s or "").upper(),
    )

def canonical_variant_match_token(token: str) -> str:
    token = str(token or "").upper().strip()
    return VARIANT_TOKEN_EQUIVALENTS.get(token, token)

def extract_variant_match_tokens(text: str) -> set:
    normalized = normalize_ocr_numeric_noise(str(text or "").upper())
    tokens = set()
    for token in re.findall(r"[A-Z0-9]+", normalized):
        if token.isdigit():
            continue
        if re.fullmatch(r"\d+(?:\.\d+)?", token):
            continue
        if re.fullmatch(r"\d+[X*]\d+(?:MG|GM|ML|MCG|IU|UNIT|G|L|KG)?", token):
            continue
        if re.fullmatch(r"\d+(?:MG|GM|ML|MCG|IU|UNIT|G|L|KG)", token):
            continue
        if re.fullmatch(r"\d+[STCP]", token):
            continue
        if token in {"X", "PACK", "OF"}:
            continue
        token = canonical_variant_match_token(token)
        if token and token not in DOSAGE_UNIT_SUFFIXES:
            tokens.add(token)
    return tokens

def build_flexible_brand_pattern(brand_name: str) -> str:
    brand_compact = compact_brand_token(brand_name)
    if not brand_compact:
        return ""
    separator = r"[\s\-]*"
    char_pattern = separator.join(re.escape(ch) for ch in brand_compact)
    return r"(?<![A-Z0-9])" + char_pattern + r"(?![A-Z0-9])"

def get_brand_removal_name(brand_name: str, available_variants: set = None) -> str:
    tokens = re.findall(r"[A-Z0-9]+", str(brand_name or "").upper().replace("+", " PLUS "))
    if len(tokens) <= 1:
        return str(brand_name or "").upper().strip()
    variant_compacts = set()
    if available_variants:
        variant_compacts = {
            compact_variant_token(variant)
            for variant in available_variants
            if variant and str(variant).upper() != "PLAIN"
        }
    removable_tokens = [tokens[0]]
    suffix_tokens = tokens[1:]
    i = 0
    while i < len(suffix_tokens):
        token = suffix_tokens[i]
        two_token_compact = ""
        if i + 1 < len(suffix_tokens):
            two_token_compact = compact_variant_token(token + " " + suffix_tokens[i + 1])
        preserve_suffix = (
            token in FORM_HINT_TOKENS
            or compact_variant_token(token) in variant_compacts
            or (two_token_compact and two_token_compact in variant_compacts)
        )
        if not preserve_suffix:
            removable_tokens.append(token)
        i += 1
    return " ".join(removable_tokens).strip()

def strip_brand_from_input(input_text: str, brand_name: str, available_variants: set = None) -> str:
    if not brand_name:
        return input_text
    input_upper = str(input_text or "").upper().strip()
    removal_brand = get_brand_removal_name(brand_name, available_variants=available_variants)
    brand_normalized = removal_brand.upper().replace("-", " ").strip()
    brand_pattern = r'\b' + re.escape(brand_normalized).replace(r"\ ", r"[-\s]+") + r'\b'
    stripped = re.sub(brand_pattern, "", input_upper, count=1, flags=re.IGNORECASE).strip()
    stripped = re.sub(r'\b' + re.escape(brand_normalized) + r'\b', "", stripped, flags=re.IGNORECASE).strip()
    if stripped != input_upper:
        return re.sub(r"\s+", " ", stripped).strip()
    flexible_pattern = build_flexible_brand_pattern(removal_brand)
    if flexible_pattern:
        stripped = re.sub(flexible_pattern, "", input_upper, count=1, flags=re.IGNORECASE).strip()
        if stripped != input_upper:
            return re.sub(r"\s+", " ", stripped).strip()
    return re.sub(r"\s+", " ", input_upper).strip()

def normalize_master_text(value, default: str = "") -> str:
    if pd.isna(value):
        return default
    value = str(value).strip()
    if value in ("nan", "NaN", "None"):
        return default
    return value

def normalize_master_code(value) -> str:
    return normalize_master_text(value, default="")

def normalize_master_attr(value, default: str = "PLAIN") -> str:
    normalized = normalize_master_text(value, default="").upper()
    return normalized if normalized else default

def alpha_signature(value: str) -> str:
    letters = re.sub(r"[^A-Z]", "", str(value or "").upper())
    consonants = re.sub(r"[AEIOU]", "", letters)
    return consonants or letters

def collapse_spaced_thousands(s: str) -> str:
    """Join a spaced thousands separator into one number: '10 000' -> '10000',
    '2 000' -> '2000'. Only a 1-3 digit group followed by a space and EXACTLY a
    3-digit group is joined, so genuine separate numbers ('10 5', '2 15') are
    left alone. Used to match master strengths like EPOFER 'VARIANT=10 000 - PFS'
    against an input written '10000'."""
    prev = None
    out = str(s or "")
    while prev != out:
        prev = out
        out = re.sub(r'\b(\d{1,3})\s+(\d{3})\b', r'\1\2', out)
    return out

def numeric_strength_variant_value(variant: str) -> str:
    """If a VARIANT field encodes a pure dose (optionally with spaced thousands
    and a trailing form/unit tag) — e.g. '10 000 - PFS', '3000', '40000' — return
    its collapsed numeric value ('10000', '3000', '40000'). Returns '' for normal
    alphabetic variants like 'M', 'XR', 'V 0.2 MG' so this stays EPOFER-specific."""
    text = collapse_spaced_thousands(str(variant or "").upper().strip())
    # strip a trailing form / unit / PFS tag (with optional '-' separator)
    text = re.sub(r'[\s\-]*(?:PFS|IU|UNIT|MG|MCG|ML|GM)\.?$', '', text).strip()
    return text if re.fullmatch(r'\d+(?:\.\d+)?', text) else ""

def normalize_ocr_numeric_noise(value: str) -> str:
    text = str(value or "").upper().replace("+", " PLUS ")
    for source_token, target_token in INPUT_TOKEN_EQUIVALENTS.items():
        text = re.sub(r'\b' + re.escape(source_token) + r'\b', target_token, text)
    # OCR occasionally turns strip-count tokens like "10S" into "10SD".
    text = re.sub(r'\b(\d+)\s*SD\b', r'\1S', text)
    text = re.sub(r'(?<=\d)O(?=\d)', '0', text)
    text = re.sub(
        r'(?<=\d)O(?=\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT|TABS?|TABLETS?|CAPS?|CAPSULES?|ML|$))',
        '0',
        text,
    )
    text = re.sub(
        r'\b(\d+(?:\.\d+)?)\s*K\b',
        lambda m: str(int(float(m.group(1)) * 1000))
        if float(m.group(1)) * 1000 == int(float(m.group(1)) * 1000)
        else str(float(m.group(1)) * 1000),
        text,
    )
    # Normalize a combination-dose separator BETWEEN TWO NUMBERS to a tight slash:
    # 'M 5 / 1000' and 'M 5-1000' both read like '5/1000' (which resolves right).
    # Fires only digit-<sep>-digit, so 'ML / VIAL', '500MG/10', 'ENCICARB -1000'
    # (space-dash, no leading digit) and hyphenated brands ('CARDACE-AM') are all
    # untouched. (No master name contains digit-dash-digit, so this is input-only.)
    text = re.sub(r'(\d)\s*[/-]\s*(\d)', r'\1/\2', text)
    # Split a strength glued to a standard strip-pack token: '510S' -> '5 10S'
    # (5 mg, pack of 10). Only when the whole number is too large to be a real
    # pack count (>100), so legit '10S'/'15S'/'30S'/'100S' are left intact.
    text = re.sub(
        r'\b(\d+)(10|15|20|30|14|28)S\b',
        lambda m: f'{m.group(1)} {m.group(2)}S' if int(m.group(1) + m.group(2)) > 100 else m.group(0),
        text,
    )
    return re.sub(r"\s+", " ", text).strip()

def extract_numeric_strength(value: str) -> str:
    text = normalize_master_attr(value, default="")
    if not text:
        return ""
    cleaned = re.sub(
        r'\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT|TAB|TABS|TABLET|TABLETS|INJ|INJECTION)\s*$',
        '',
        text,
    ).strip()
    if re.fullmatch(r'\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)*', cleaned):
        return cleaned
    return ""

def is_numeric_subvariant_like(value: str) -> bool:
    return bool(re.fullmatch(r'\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)*', str(value or "").strip().upper()))

_ORAL_FORM_TOKENS = {
    "TABLET", "TABLETS", "TAB", "TABS", "CAPSULE", "CAPSULES", "CAP", "CAPS",
    "SYRUP", "SYP", "DROPS", "DROP", "SUSPENSION", "SUSP", "CREAM", "GEL",
    "OINTMENT", "LOTION", "SACHET", "POWDER", "POWD", "GRANULES", "GRANULE",
    "SPRAY", "NASALSPRAY", "SOLUTION", "SOL", "TONIC", "BAR", "KIT",
}


def repair_master_variant_fields(product_name: str, variant: str, sub_variant: str) -> tuple:
    variant = normalize_master_attr(variant)
    sub_variant = normalize_master_attr(sub_variant)
    # Strip ORAL/topical form words baked into the VARIANT (e.g. 'GM 50/2/500
    # TABLET' → 'GM 50/2/500', 'M TABLET' → 'M', 'O DROPS' → 'O'). Injection
    # forms (PFS/VIAL/INJ) are deliberately NOT stripped — the rules below need
    # them (EPOFER '10 000 - PFS', AUGPEN 'I.V.', etc.).
    _vt = [t for t in variant.split() if t.upper() not in _ORAL_FORM_TOKENS]
    variant = " ".join(_vt).strip() or "PLAIN"
    numeric_pattern = r'\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)*'
    numeric_with_unit_pattern = numeric_pattern + r'(?:\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT))?'
    noisy_sub_variant_tokens = {
        "MG", "GM", "ML", "MCG", "G", "M", "IU", "UNIT",
        "TAB", "TABS", "TABLET", "TABLETS", "INJ", "INJECTION",
        # Injection delivery forms — a dose, not these, is the real sub-variant.
        # (PFS = pre-filled syringe ≈ injection.) Keeps EPOFER 4000/6000 as
        # VARIANT=PLAIN, SUB_VARIANT=<strength>, consistent with 3000/40000.
        "PFS", "VIAL", "AMP", "AMPS", "SYR", "SYRINGE", "PEN",
    }
    # VARIANT like '10 000 - PFS' / '2 000 - PFS' / '4000 IU' — a dose glued to
    # an injection-form/unit tag. Collapse spaced thousands, drop the tag, and
    # treat the number as the strength so the whole family is VARIANT=PLAIN,
    # SUB_VARIANT=<strength> (no variant literally carries 'PFS' to mis-match).
    if sub_variant == "PLAIN":
        collapsed_variant = collapse_spaced_thousands(variant)
        form_dose = re.fullmatch(
            r'(\d+(?:\.\d+)?)\s*-?\s*(?:PFS|VIAL|AMP|AMPS|SYR|SYRINGE|PEN|IU|INJ|INJECTION)',
            collapsed_variant,
        )
        if form_dose:
            return "PLAIN", form_dose.group(1)

    variant_numeric = extract_numeric_strength(variant)
    if sub_variant in noisy_sub_variant_tokens and variant_numeric:
        return "PLAIN", variant_numeric
    if (
        re.fullmatch(numeric_pattern, variant)
        and re.search(r'[A-Z]', sub_variant)
        and not is_numeric_subvariant_like(sub_variant)
        and sub_variant not in noisy_sub_variant_tokens
    ):
        return normalize_master_attr(sub_variant), variant
    if sub_variant != "PLAIN":
        return variant, sub_variant
    if re.fullmatch(numeric_pattern, variant):
        return "PLAIN", variant
    pure_numeric_with_unit_match = re.fullmatch(
        r'(' + numeric_pattern + r')\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT)',
        variant,
    )
    if pure_numeric_with_unit_match:
        return "PLAIN", pure_numeric_with_unit_match.group(1)
    split_match = re.fullmatch(r'(.+?)[\s\-]+(' + numeric_pattern + r')', variant)
    if split_match:
        repaired_variant = split_match.group(1).strip()
        repaired_sub_variant = split_match.group(2).strip()
        if repaired_variant:
            return repaired_variant, repaired_sub_variant
    compact_combo_match = re.fullmatch(r'([A-Z]+(?:\s+[A-Z]+)*?)(' + numeric_pattern + r')', variant)
    if compact_combo_match:
        repaired_variant = compact_combo_match.group(1).strip()
        repaired_sub_variant = compact_combo_match.group(2).strip()
        if repaired_variant and repaired_sub_variant and (
            '/' in repaired_sub_variant or len(repaired_sub_variant.replace('.', '')) >= 2
        ):
            return repaired_variant, repaired_sub_variant
    reverse_split_match = re.fullmatch(r'(' + numeric_pattern + r')\s+([A-Z]+(?:\s+[A-Z0-9]+)*)', variant)
    if reverse_split_match:
        repaired_sub_variant = reverse_split_match.group(1).strip()
        repaired_variant = reverse_split_match.group(2).strip()
        if repaired_variant and repaired_variant not in noisy_sub_variant_tokens:
            return repaired_variant, repaired_sub_variant
    unit_split_match = re.fullmatch(r'(.+?)[\s\-]+(' + numeric_with_unit_pattern + r')', variant)
    if unit_split_match:
        repaired_variant = unit_split_match.group(1).strip()
        repaired_sub_variant = unit_split_match.group(2).strip()
        repaired_sub_variant = re.sub(r'\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT)\s*$', '', repaired_sub_variant).strip()
        if repaired_variant and repaired_sub_variant:
            return repaired_variant, repaired_sub_variant
    return variant, sub_variant

def is_reasonable_pack_count(value) -> bool:
    try:
        pack_num = float(str(value).strip())
    except (TypeError, ValueError):
        return False
    return pack_num.is_integer() and 0 < pack_num <= 100

def find_tablike_pack_matches(input_upper: str) -> list:
    matches = []
    pattern = r'(?<![\d.])(\d+)\s*(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\b'
    for m in re.finditer(pattern, input_upper):
        if is_reasonable_pack_count(m.group(1)):
            matches.append(m)
    return matches

def parse_pack_size_value(pack: str):
    pack_upper = str(pack or "").upper().strip()
    if not pack_upper or pack_upper in ("NAN", "NONE"):
        return None
    combo_match = re.search(r'(\d+)\s*[X*]\s*(\d+)', pack_upper)
    if combo_match:
        return float(int(combo_match.group(1)) * int(combo_match.group(2)))
    numeric_match = re.search(r'(\d+(?:\.\d+)?)', pack_upper)
    if numeric_match:
        return float(numeric_match.group(1))
    return None

def simple_normalize_for_llm(s: str) -> str:
    if not s:
        return s
    s = str(s).upper()
    s = re.sub(r'\b([A-Z])\.([A-Z])\b', r'\1\2', s)
    s = re.sub(r'\b([A-Z])\.([A-Z])\.([A-Z])\b', r'\1\2\3', s)
    s = s.replace("-", " ")
    s = s.replace("+", " PLUS ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def extract_combination_dosage(input_name: str) -> str:
    input_name = normalize_ocr_numeric_noise(input_name)
    # Trailing boundary allows a glued unit ('5/1000MG' — '1000MG' is one token so
    # \b never fires before MG); leading lookbehind avoids starting mid-number.
    pattern = r'(?<![\d./])(\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)+)(?![\d./])'
    match = re.search(pattern, input_name)
    return match.group(1) if match else ""

def clean_duplicate_words(input_name: str) -> str:
    words = input_name.split()
    cleaned = []
    for i, word in enumerate(words):
        if i == 0 or word.upper() != words[i-1].upper():
            cleaned.append(word)
    result = cleaned
    for n in (2, 3):
        i = 0
        deduped = []
        while i < len(result):
            chunk = result[i:i+n]
            if len(chunk) == n and result[i+n:i+2*n] == chunk:
                deduped.extend(chunk)
                i += 2 * n
            else:
                deduped.append(result[i])
                i += 1
        result = deduped
    return " ".join(result)


def extract_brand_like_query(input_name: str) -> str:
    text = normalize_ocr_numeric_noise(input_name)
    tokens = re.findall(r'[A-Z0-9.]+', text)
    if not tokens:
        return ""
    skip_tokens = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES)
    skip_tokens |= {"PACK", "OF", "PCS", "PC", "NOS", "NO", "VIAL", "AMP", "AMPS", "PFS", "SYR"}
    brand_tokens = []
    for token in tokens:
        if token in skip_tokens:
            break
        if re.search(r'\d', token):
            # Brand glued to a strength/pack, e.g. 'CLIXANE40MG' → take the
            # leading alpha run (>=2 chars) as the brand, then stop.
            glued = re.match(r'^([A-Z]{2,})\d', token)
            if glued and not brand_tokens:
                brand_tokens.append(glued.group(1))
            break
        brand_tokens.append(token)
    return " ".join(brand_tokens).strip()


# ============================================================================
# FIX 2: STRIP PARENTHETICAL SOURCE-NOISE TOKENS
# ----------------------------------------------------------------------------
# "(AP)", "(AV)", "(SP)", "(ST)" etc. are purchase-source annotations, not
# part of the product name. Strip them anywhere in the input before any
# downstream matching.
# ============================================================================
def get_brand_variant_tags(brand_map: dict, brand_hint: str) -> set:
    """Short (<=4-letter) real VARIANT values belonging to a resolved brand hint.

    Used only by strip_parenthetical_noise, so that a parenthetical tag which is
    one of the brand's own VARIANT values -- "AMARYL-(M)1", "CARDACE (H) 5 TAB"
    -- is preserved rather than deleted as source noise. "PLAIN" is the
    no-variant sentinel and is never a tag.
    """
    if not brand_hint or not brand_map:
        return set()
    hint = str(brand_hint).strip().upper()
    key = hint if hint in brand_map else next((k for k in brand_map if k.upper() == hint), None)
    if key is None:
        return set()
    tags = set()
    for item in brand_map[key]:
        variant = str(item.get("variant", "") or "").strip().upper()
        if variant and variant != "PLAIN" and re.fullmatch(r'[A-Z]{1,4}', variant):
            tags.add(variant)
    return tags


def strip_parenthetical_noise(s: str, brand_variant_tags: set = None) -> str:
    if not s:
        return s
    text = str(s)
    tag_pattern = r'\(\s*([A-Z]{1,4})\s*\)'
    # VARIANT-AWARE EXCEPTION: when the brand for this row is already known and
    # the tag is one of that brand's own VARIANT values, unwrap "(X)" -> " X "
    # so the variant survives into matching ("AMARYL-(M)1" -> "AMARYL- M 1").
    # A tag that already stands as its own token elsewhere in the name needs no
    # rescuing, so it is dropped as before and the string is left untouched.
    outside_tokens = None
    if brand_variant_tags:
        outside_tokens = set(re.findall(
            r'[A-Z0-9]+',
            re.sub(tag_pattern, ' ', text, flags=re.IGNORECASE).upper()))

    def _unwrap_or_drop(m):
        tag = m.group(1).strip().upper()
        if brand_variant_tags and tag in brand_variant_tags and tag not in outside_tokens:
            return ' ' + tag + ' '
        return ' '

    # Remove (X), (XX), (XXX) — short all-letter parenthetical tags (≤4 chars).
    # Leaves legitimate content like "(1.5ML)" intact because it contains digits.
    cleaned = re.sub(tag_pattern, _unwrap_or_drop, text, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', cleaned).strip()


# =========================
# COMPOUND BRAND-SUFFIX TOKEN EXTRACTION
# =========================
def extract_brand_suffix_tokens(input_name: str, brand_name: str) -> list:
    if not brand_name:
        return []
    input_upper = normalize_ocr_numeric_noise(input_name)
    brand_upper = brand_name.upper().strip()
    tokens = []
    skip_tokens = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES)
    skip_tokens |= {
        "PACK", "OF", "PCS", "PC", "NOS", "NO", "VIAL", "AMP", "AMPS", "PFS", "SYR",
        "POWDE", "POWDER", "GRANUL", "GRANULE", "GRANULES", "GRANULS",
    }

    def should_keep_token(token: str) -> bool:
        token_letters = re.sub(r'[^A-Z]', '', token.upper())
        return bool(token_letters) and token_letters not in skip_tokens

    pattern1 = re.escape(brand_upper) + r'-([A-Z]+\d+|\d+[A-Z]+)\b'
    for m in re.findall(pattern1, input_upper):
        if should_keep_token(m):
            tokens.append(m)
    pattern2 = re.escape(brand_upper) + r'\s+([A-Z]+\d+|\d+[A-Z]+)\b'
    for m in re.findall(pattern2, input_upper):
        if should_keep_token(m) and m not in tokens:
            tokens.append(m)

    bm = re.search(r'\b' + re.escape(brand_upper) + r'\b', input_upper)
    if bm:
        tail_match = re.match(r'[\s\-]*([A-Z]{2,})\b', input_upper[bm.end():])
        if tail_match:
            alpha_token = tail_match.group(1)
            if alpha_token not in skip_tokens and alpha_token not in tokens:
                tokens.append(alpha_token)
    return tokens


def filter_items_by_compound_tokens(items: list, compound_tokens: list, verbose=False) -> list:
    if not compound_tokens:
        return items
    if verbose:
        print(f"\n  [COMPOUND TOKEN] Tokens to match: {compound_tokens}")
        print(f"  [COMPOUND TOKEN] Products before filter: {len(items)}")
    normalized_tokens = {compact_variant_token(token) for token in compound_tokens if token}
    matched = []
    for item in items:
        prod_upper = item["product"].upper().replace("-", " ").replace("_", " ")
        prod_compact = compact_variant_token(prod_upper)
        product_tokens = re.findall(r'[A-Z0-9]+', prod_upper)
        product_compact_tokens = {compact_variant_token(tok) for tok in product_tokens if tok}
        for i in range(len(product_tokens) - 1):
            product_compact_tokens.add(compact_variant_token(product_tokens[i] + product_tokens[i + 1]))
        for i in range(len(product_tokens) - 2):
            product_compact_tokens.add(compact_variant_token(product_tokens[i] + product_tokens[i + 1] + product_tokens[i + 2]))
        has_match = False
        for token in compound_tokens:
            normalized = compact_variant_token(token)
            if not token:
                continue
            if re.fullmatch(r'(?:[A-Z]+\d+|\d+[A-Z]+)', token):
                if token in product_tokens or normalized in product_compact_tokens:
                    has_match = True
                    break
                continue
            if token in prod_upper or (normalized and normalized in prod_compact):
                has_match = True
                break
        if has_match:
            matched.append(item)
            if verbose:
                print(f"    ✓ Matched '{item['product']}' on token(s) {compound_tokens}")
    if matched:
        if verbose:
            print(f"  [COMPOUND TOKEN] Products after filter: {len(matched)}")
        return matched
    if verbose:
        print(f"  [COMPOUND TOKEN] No products matched tokens — keeping all {len(items)} products")
    return items


def get_compound_token_digits(compound_tokens: list) -> set:
    digits = set()
    for token in compound_tokens:
        for d in re.findall(r'\d+', token):
            digits.add(d)
    return digits


# =========================
# PACK-POSITION HELPER
# =========================
def get_all_pack_positions(input_upper: str) -> list:
    positions = []
    # NxM — allow optional unit suffix right after (e.g. "1*25ML", "2x15TAB")
    # so "25" in "1*25ML" is recognised as part of the pack, not a sub-variant.
    for m in re.finditer(
        r'\b(\d+)\s*[X*×]\s*(\d+(?:\.\d+)?)(?:\s*(?:ML|GM|MG|MCG|IU|G|L|KG|TA|TABS?|TABLETS?|CAPS?|CAPSULES?|STRIPS?))?',
        input_upper,
    ):
        positions.append((m.start(), m.end()))
    for m in re.finditer(r'\bX\s*(\d+)\b', input_upper):
        positions.append((m.start(), m.end()))
    for m in re.finditer(r'\b(\d+)([STCP])\b', input_upper):
        if m.group(2) in ['S', 'T', 'C', 'P']:
            positions.append((m.start(), m.end()))
    for m in re.finditer(r'\b(\d+)\s+(?:S|T|PCS)\b', input_upper):
        positions.append((m.start(), m.end()))
    tablike_matches = find_tablike_pack_matches(input_upper)
    if tablike_matches and not positions:
        last_match = tablike_matches[-1]
        positions.append((last_match.start(), last_match.end()))
    return positions


def is_position_in_pack_pattern(match_start: int, match_end: int, pack_positions: list) -> bool:
    for pack_start, pack_end in pack_positions:
        if (pack_start <= match_start < pack_end) or (pack_start < match_end <= pack_end):
            return True
    return False


def split_strength_and_variant(input_name: str, available_variants: set, verbose=False) -> tuple:
    detected_variants = set()
    detected_sub_variants = set()
    input_upper = input_name.upper()
    pack_positions = get_all_pack_positions(input_upper)
    pattern = r'\b(\d+(?:\.\d+)?)([A-Z]+)\b'
    for match in re.finditer(pattern, input_upper):
        number_part = match.group(1)
        letter_part = match.group(2)
        found_letter_variant = False
        found_full_letter_variant = False
        is_dosage_suffix = letter_part in DOSAGE_UNIT_SUFFIXES
        if is_position_in_pack_pattern(match.start(), match.end(), pack_positions):
            continue
        if not is_dosage_suffix:
            # Only treat the alpha run as variant LETTERS when EVERY letter is
            # itself an available variant ('5M', '5MV'). Otherwise a real word
            # glued to a number ('1VIL' = 1 VIAL) would spuriously match a buried
            # single-letter variant ('L'). Whole multi-letter variants ('XR') are
            # still caught by the full-token check just below.
            if all(letter in available_variants for letter in letter_part):
                for i, letter in enumerate(letter_part):
                    if letter in available_variants:
                        found_letter_variant = True
                        detected_variants.add(letter)
                        if i == 0:
                            detected_sub_variants.add(number_part)
            if letter_part in available_variants:
                found_full_letter_variant = True
                detected_variants.add(letter_part)
                detected_sub_variants.add(number_part)
        if (
            number_part in available_variants
            and not found_letter_variant
            and not found_full_letter_variant
        ):
            detected_variants.add(number_part)
    return detected_variants, detected_sub_variants


def detect_compact_variant_strength_combo(input_name: str, available_variants: set) -> tuple:
    detected_variants = set()
    detected_sub_variants = set()
    tokens = re.findall(r'[A-Z0-9]+', normalize_ocr_numeric_noise(input_name))
    skip_tokens = _pack_form_skip_tokens() | {"PACK", "OF"}
    for i in range(len(tokens) - 1):
        compact_token = tokens[i]
        next_token = tokens[i + 1]
        m = re.fullmatch(r'([A-Z]+)(\d+(?:\.\d+)?)', compact_token)
        if not m:
            continue
        letter_part, numeric_part = m.groups()
        if letter_part in available_variants:
            detected_variants.add(letter_part)
            detected_sub_variants.add(numeric_part)
        if next_token in skip_tokens:
            continue
        candidate_variant = f"{letter_part} {next_token}"
        if candidate_variant in available_variants:
            detected_variants.add(candidate_variant)
            detected_sub_variants.add(numeric_part)
    return detected_variants, detected_sub_variants


def detect_variant_numeric_pairs(input_name: str, available_variants: set) -> tuple:
    detected_variants = set()
    detected_sub_variants = set()
    normalized = normalize_ocr_numeric_noise(input_name)
    seen = set()
    patterns = [
        r'\b([A-Z]+)(\d+(?:\.\d+)?)\b',
        r'\b([A-Z]+)[\-\s]+(\d+(?:\.\d+)?)\b',
    ]
    for pattern in patterns:
        for m in re.finditer(pattern, normalized):
            letter_part, numeric_part = m.groups()
            if letter_part in DOSAGE_UNIT_SUFFIXES:
                continue
            key = (letter_part, numeric_part)
            if key in seen:
                continue
            seen.add(key)
            if letter_part in available_variants:
                detected_variants.add(letter_part)
                detected_sub_variants.add(numeric_part)
    return detected_variants, detected_sub_variants


def product_contains_combination(product_name: str, combination: str) -> bool:
    product_norm = product_name.upper().replace(" ", "").replace("-", "")
    combo_norm = combination.upper().replace(" ", "").replace("-", "")
    if not combo_norm:
        return False
    # Boundary-aware: the combo's first digit must not sit inside a larger number
    # (so '5/1000' does NOT match the '2.5/1000' tail) and its last digit must not
    # run into more digits. A glued letter prefix ('M5/1000') is still a match.
    pat = r'(?<![\d.])' + re.escape(combo_norm) + r'(?![\d.])'
    return bool(re.search(pat, product_norm))


def extract_available_attributes_from_items(items: list) -> tuple:
    available_variants = {"PLAIN"}
    available_sub_variants = {"PLAIN"}
    available_pack_sizes = set()
    for item in items:
        variant = str(item.get("variant", "PLAIN")).strip().upper()
        sub_variant = str(item.get("sub_variant", "PLAIN")).strip().upper()
        pack_size = str(item.get("pack_size", "")).strip()
        if variant and variant not in ("", "NAN", "NONE"):
            available_variants.add(variant)
        if sub_variant and sub_variant not in ("", "NAN", "NONE"):
            available_sub_variants.add(sub_variant)
        if pack_size and pack_size not in ("", "nan", "NaN", "None"):
            available_pack_sizes.add(pack_size)
    return available_variants, available_sub_variants, available_pack_sizes


# ============================================================================
# FIX 1 (BIGGEST IMPACT): extract_pack_size_from_input — regex-order fix
# ----------------------------------------------------------------------------
# BUG in updated_ts_3: the old order tried `\d+\s*(?:ML|GM|MG|MCG|IU|...)` BEFORE
# the pack-count tokens (10S, 15 TAB, 1X15). This meant "CARDACE 1.25 MG 15 TAB"
# extracted pack=1.25 (the strength!) instead of 15.
#
# FIX: try unambiguous pack-count tokens FIRST. MG/MCG/IU are dosage units,
# never pack counts, so they are removed entirely from pack extraction. Liquid
# volumes (ML/GM/L/KG) remain as a last-resort fallback for syrups/creams.
#
# Verified on 17 tricky cases (17/17 pass), affects ~279 of 929 red rows.
# ============================================================================
def extract_pack_size_from_input(s: str) -> str:
    s_upper = normalize_ocr_numeric_noise(s)

    # 1. NxM / N*M combo notation — highest priority (explicit)
    m = re.search(r'\b(\d+)\s*[X*×]\s*(\d+(?:\.\d+)?)(?:\s*(?:ML|GM|MG|MCG|IU|G|L|KG|TA|TABS?|TABLETS?|CAPS?|CAPSULES?|T|C))?\b', s_upper)
    if m:
        total = float(m.group(1)) * float(m.group(2))
        return str(int(total) if total.is_integer() else total)

    # 2. PACK OF N
    m = re.search(r'\bPACK\s*(?:OF\s*)?(\d+)\b', s_upper)
    if m:
        return m.group(1)

    # 3. UNAMBIGUOUS pack-count suffix tokens FIRST:
    #    10S, 15T glued. "N TAB" is ambiguous (could be strength), so handle
    #    suffix tokens BEFORE "N TAB" matching.
    m = re.search(r'\b(\d+)[STCP]\b', s_upper)
    if m:
        return m.group(1)
    m = re.search(r'\b(\d+)\s+(?:S|T|PCS)\b', s_upper)
    if m:
        return m.group(1)

    # 4. N TAB / N TABLET / N CAP — pack count adjacent to form word.
    #    Prefer the LAST such match (closer to the end is usually pack).
    adjacent_form_matches = []
    for m in re.finditer(
        r'(?<![\d.])(\d+)\s*(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\b', s_upper):
        if is_reasonable_pack_count(m.group(1)):
            adjacent_form_matches.append((m.start(), m.group(1)))
    # 4b. TAB N / CAP N — OCR sometimes swaps the order.
    for m in re.finditer(
        r'\b(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\s*(\d+)\b', s_upper):
        if is_reasonable_pack_count(m.group(1)):
            adjacent_form_matches.append((m.start(), m.group(1)))
    if adjacent_form_matches:
        adjacent_form_matches.sort(key=lambda x: x[0], reverse=True)
        return adjacent_form_matches[0][1]

    # 5. Liquid volume only (ML/GM/G/L/KG) — NEVER MG/MCG/IU.
    #    Used for syrup/injectable fill volumes, cream tube sizes.
    m = re.search(r'\b(\d+(?:\.\d+)?)\s*(?:ML|GM|G|L|KG)\b', s_upper)
    if m:
        return m.group(1)

    return ""


# =========================
# DATA LOADING
# =========================
def load_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    for required in (PRODUCT_COL, BRAND_COL):
        if required not in df.columns:
            raise ValueError(f"Missing column '{required}'. Found: {list(df.columns)}")
    cols = [PRODUCT_COL, BRAND_COL]
    if PACK_SIZE_COL in df.columns:
        cols.append(PACK_SIZE_COL)
        print(f"✓ Found '{PACK_SIZE_COL}' column in master – pack-size matching ENABLED.")
    else:
        print(f"⚠ '{PACK_SIZE_COL}' column NOT found in master – pack sizes will be ignored.")
    if VARIANT_COL in df.columns:
        cols.append(VARIANT_COL)
        print(f"✓ Found '{VARIANT_COL}' column in master – variant-based filtering ENABLED.")
    else:
        print(f"⚠ '{VARIANT_COL}' column NOT found in master – all products will be shown.")
    if SUB_VARIANT_COL in df.columns:
        cols.append(SUB_VARIANT_COL)
        print(f"✓ Found '{SUB_VARIANT_COL}' column in master – sub-variant filtering ENABLED.")
    else:
        print(f"⚠ '{SUB_VARIANT_COL}' column NOT found in master – sub-variant filtering DISABLED.")
    if PRODUCT_CODE_COL in df.columns:
        cols.append(PRODUCT_CODE_COL)
        print(f"✓ Found '{PRODUCT_CODE_COL}' column in master – product codes will be returned.")
    else:
        print(f"⚠ '{PRODUCT_CODE_COL}' column NOT found in master – product codes will not be available.")
    df = df[cols].copy()
    df[PRODUCT_COL] = df[PRODUCT_COL].map(normalize_master_text)
    df[BRAND_COL] = df[BRAND_COL].map(normalize_master_text)
    if PACK_SIZE_COL in df.columns:
        df[PACK_SIZE_COL] = df[PACK_SIZE_COL].map(normalize_master_text)
    else:
        df[PACK_SIZE_COL] = ""
    if VARIANT_COL in df.columns:
        df[VARIANT_COL] = df[VARIANT_COL].map(normalize_master_attr)
    else:
        df[VARIANT_COL] = "PLAIN"
    if SUB_VARIANT_COL in df.columns:
        df[SUB_VARIANT_COL] = df[SUB_VARIANT_COL].map(normalize_master_attr)
    else:
        df[SUB_VARIANT_COL] = "PLAIN"
    if PRODUCT_CODE_COL in df.columns:
        df[PRODUCT_CODE_COL] = df[PRODUCT_CODE_COL].map(normalize_master_code)
    else:
        df[PRODUCT_CODE_COL] = ""
    df = df[(df[PRODUCT_COL] != "") & (df[BRAND_COL] != "")]
    df["product_norm"] = df[PRODUCT_COL].map(norm)
    df["brand_norm"] = df[BRAND_COL].map(norm)
    return df


def build_brand_product_map(df: pd.DataFrame) -> dict:
    brand_map = {}
    for brand_norm, group in df.groupby("brand_norm"):
        original_brand = group.iloc[0][BRAND_COL]
        seen = set()
        items = []
        for _, row in group.iterrows():
            prod = str(row[PRODUCT_COL]).strip()
            pack = str(row.get(PACK_SIZE_COL, "")).strip()
            pack = "" if pack in ("nan", "NaN", "None") else pack
            variant = str(row.get(VARIANT_COL, "PLAIN")).strip().upper()
            variant = "PLAIN" if variant in ("", "NAN", "NONE") else variant
            sub_variant = str(row.get(SUB_VARIANT_COL, "PLAIN")).strip().upper()
            sub_variant = "PLAIN" if sub_variant in ("", "NAN", "NONE") else sub_variant
            variant, sub_variant = repair_master_variant_fields(prod, variant, sub_variant)
            product_code = str(row.get(PRODUCT_CODE_COL, "")).strip()
            product_code = "" if product_code in ("nan", "NaN", "None") else product_code
            key = (prod, pack)
            if prod and key not in seen:
                seen.add(key)
                items.append({
                    "product": prod,
                    "pack_size": pack,
                    "variant": variant,
                    "sub_variant": sub_variant,
                    "product_code": product_code,
                    "product_norm": norm(prod)
                })
        brand_map[original_brand] = items
    return brand_map


def extract_all_variants_from_data(brand_map: dict) -> set:
    all_variants = set()
    for brand, items in brand_map.items():
        for item in items:
            variant = item.get("variant", "PLAIN")
            if variant != "PLAIN":
                all_variants.add(variant.upper())
    all_variants.add("PLAIN")
    return all_variants


def extract_all_sub_variants_from_data(brand_map: dict) -> set:
    all_sub_variants = set()
    for brand, items in brand_map.items():
        for item in items:
            sub_variant = item.get("sub_variant", "PLAIN")
            if sub_variant != "PLAIN":
                all_sub_variants.add(sub_variant.upper())
    all_sub_variants.add("PLAIN")
    return all_sub_variants


# ============================================================================
# FIX 4: abbreviated sub-brand promotion
# ----------------------------------------------------------------------------
# "CARDACE-MET 5 TAB" picks plain CARDACE because METO's 2nd word "METO"
# isn't in the input — only the abbreviation "MET" is. Detect this case
# after normal brand selection and promote to the longer sub-brand when its
# next word has a prefix-match to the token following the matched brand.
# Purely structural; no brand names hardcoded.
# ============================================================================
def _pack_form_skip_tokens():
    """Tokens that should NEVER trigger sub-brand promotion."""
    tokens = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES)
    tokens |= {
        "PCS", "PC", "NOS", "NO", "VIAL", "AMP", "AMPS", "PFS", "SYR",
        "POWDE", "POWDER", "GRANUL", "GRANULE", "GRANULES", "GRANULS",
    }
    return tokens


def _brand_letters(s: str) -> str:
    """Letters-only, upper-cased compaction of a brand-like string."""
    return re.sub(r"[^A-Z]", "", str(s or "").upper())


def _resembles_manufacturer(token: str) -> bool:
    """True if a token is an OCR-garbled manufacturer name (EMCOO/EMC00/EMCU →
    EMCURE): shares the first 3 letters and is within MANUFACTURER_NOISE_MAX_EDITS
    edits of, or a prefix of, a configured manufacturer name."""
    alpha = re.sub(r"[^A-Z]", "", re.sub(r"0", "O", str(token or "").upper()))
    if len(alpha) < 3:
        return False
    for mfr in MANUFACTURER_NOISE_NAMES:
        if alpha[:3] == mfr[:3] and (
            mfr.startswith(alpha) or alpha.startswith(mfr)
            or Levenshtein.distance(alpha, mfr) <= MANUFACTURER_NOISE_MAX_EDITS
        ):
            return True
    return False


def _is_strong_brand_token(token: str, brand_compacts: set) -> bool:
    """True if a token clearly IS a master brand: its compact form equals a brand
    compact, or is a brand compact followed by a digit (OROFER→OROFER150)."""
    c = compact_brand_token(token)
    if not c:
        return False
    if c in brand_compacts:
        return True
    return any(c.startswith(bc) and len(c) > len(bc) and c[len(bc)].isdigit()
               for bc in brand_compacts)


def strip_manufacturer_noise(input_name: str, brand_map: dict) -> str:
    """Remove a leading manufacturer-name prefix that OCR left in front of the
    real brand — but ONLY when that prefix is not itself a brand and the brand
    that follows IS a strong match. So 'EMCOO TAPICEM …'→'TAPICEM …' and the
    glued 'EMC00CLAMOB …'→'CLAMOB …', while real brands 'EMCOR CREAM' and
    'EMCURE ORS …' are left untouched (they resolve on the full input)."""
    text = str(input_name or "").strip()
    tokens = text.split()
    if not tokens:
        return text
    brand_compacts = {compact_brand_token(b) for b in brand_map}
    first = tokens[0].upper()

    # A) manufacturer glued to brand by an OCR digit run: 'EMC00CLAMOB'
    glued = re.match(r'^([A-Z]{2,5}\d+)([A-Z]{3,})$', first)
    if (glued and _resembles_manufacturer(glued.group(1))
            and _is_strong_brand_token(glued.group(2), brand_compacts)):
        return " ".join([glued.group(2)] + tokens[1:]).strip()

    # B) manufacturer as a separate leading token: 'EMCOO TAPICEM …'
    if (len(tokens) >= 2 and _resembles_manufacturer(first)
            and not _is_strong_brand_token(first, brand_compacts)
            and _is_strong_brand_token(tokens[1], brand_compacts)):
        return " ".join(tokens[1:]).strip()

    return text


_DESEG_UNITS = r'MG|MCG|ML|GM|IU'
_DESEG_FORMS = (r'TABLETS|TABLET|TABS|TAB|CAPSULES|CAPSULE|CAPS|CAP|'
                r'INJECTION|INJ|SYRUP|SYP|DROPS|GRANULES|SUSPENSION')


def _split_glued_brand_token(tok: str, brand_letters: set) -> str:
    """Insert a space after the longest known brand that prefixes a glued token
    ('CELOLCAPSULES'→'CELOL CAPSULES', 'CARDACEAM2.5/5'→'CARDACE AM2.5/5',
    'ATOREC10MGTABLETS'→'ATOREC 10MGTABLETS')."""
    m = re.match(r'^[A-Za-z]+', tok)
    if not m:
        return tok
    lead = m.group(0).upper()
    best = ''
    for b in brand_letters:
        if len(best) < len(b) <= len(lead) and lead.startswith(b):
            best = b
    # Drop a hyphen sitting at the brand→variant boundary ('CONSIVAS-20TAB' →
    # 'CONSIVAS 20TAB', 'CARDACE-MET' → 'CARDACE MET') so the variant token is clean.
    if best and best != lead:
        return tok[:len(best)] + ' ' + tok[len(best):].lstrip('-')
    if best == lead and len(tok) > len(lead):
        return lead + ' ' + tok[len(lead):].lstrip('-')
    return tok


# Single-slot cache for _variant_subvariant_tokens. Keyed on the identity of the
# brand_map object and holding a reference to it: keeping the reference alive stops
# the object being collected and its id() reused by a different map, so a stale hit
# is impossible. A value-based key is deliberately avoided for that reason.
_VST_CACHE_MAP = None
_VST_CACHE_VAL = None


def _variant_subvariant_tokens(brand_map: dict) -> tuple:
    """Alpha tokens (len>=2) seen in any VARIANT / SUB_VARIANT across the master,
    used to split a glued variant+sub-variant token ('OHFORTE'→'OH FORTE').

    The result is a pure function of brand_map, which build_brand_product_map
    creates once per run and never mutates, so it is memoized per map object
    instead of being rebuilt for every input row."""
    global _VST_CACHE_MAP, _VST_CACHE_VAL
    if _VST_CACHE_MAP is brand_map:      # same master object -> same vocabulary
        return _VST_CACHE_VAL
    var, subvar = set(), set()
    for items in brand_map.values():
        for it in items:
            for t in re.findall(r'[A-Z]{2,}', str(it.get("variant", "")).upper()):
                var.add(t)
            for t in re.findall(r'[A-Z]{2,}', str(it.get("sub_variant", "")).upper()):
                subvar.add(t)
    _VST_CACHE_MAP, _VST_CACHE_VAL = brand_map, (var, subvar)
    return var, subvar


def _split_glued_variant(tok: str, var_tokens: set, subvar_tokens: set) -> str:
    """'OHFORTE' → 'OH FORTE' when the prefix is a known variant and the suffix a
    known sub-variant. Leaves a token that is itself a known variant/sub-variant
    (e.g. 'METO', 'PROTECT', 'FORTE') intact."""
    u = tok.upper()
    if not tok.isalpha() or len(u) < 4 or u in var_tokens or u in subvar_tokens:
        return tok
    for i in range(2, len(u) - 1):
        if u[:i] in var_tokens and u[i:] in subvar_tokens:
            return tok[:i] + ' ' + tok[i:]
    return tok


def desegment_ocr_input(input_name: str, brand_map: dict) -> str:
    """Repair OCR strings that lost their spaces (or gained noise) so the brand /
    variant / strength / form become separable tokens. Handles: trailing 'X'
    redaction runs and '--', space-split decimals ('12. 5'→'12.5'), a split
    leading letter ('E MSITA'→'EMSITA'), a brand glued to the rest
    ('CELOLCAPSULES'→'CELOL CAPSULES'), and glued strength+unit+form
    ('10MGTABLETS'→'10 MG TABLETS')."""
    s = str(input_name or "")
    # OCR redaction 'X' runs — strip ONLY when glued before '--' ('CAPXXXXX --')
    # or a standalone token ('… XX --'); never 'XX' inside a real brand (DIVEXX).
    s = re.sub(r'[Xx]{2,}(?=\s*-{2,})', '', s)
    s = re.sub(r'(?<!\S)[Xx]{2,}(?!\S)', ' ', s)
    s = re.sub(r'\s*-{1,}\s*$', '', s)                  # trailing '--'
    s = re.sub(r'(\d)\.\s+(\d)', r'\1.\2', s)           # '12. 5' -> '12.5'
    s = re.sub(r'\s+', ' ', s).strip()
    tokens = s.split()
    if not tokens:
        return s

    brand_compacts = {compact_brand_token(b) for b in brand_map}
    brand_letters = {_brand_letters(b) for b in brand_map if len(_brand_letters(b)) >= 4}

    # Rejoin a split leading letter when it forms a real brand: 'E MSITA'→'EMSITA'
    if (len(tokens) >= 2 and len(tokens[0]) == 1 and tokens[0].isalpha()
            and _is_strong_brand_token(tokens[0] + tokens[1], brand_compacts)):
        tokens = [tokens[0] + tokens[1]] + tokens[2:]

    tokens[0] = _split_glued_brand_token(tokens[0], brand_letters)
    s = " ".join(tokens)

    # Split a glued variant+sub-variant token ('OHFORTE'→'OH FORTE') on the
    # tokens after the brand, using the master's variant/sub-variant vocabulary.
    var_tokens, subvar_tokens = _variant_subvariant_tokens(brand_map)
    parts = s.split()
    s = " ".join(parts[:1] + [_split_glued_variant(p, var_tokens, subvar_tokens)
                              for p in parts[1:]])

    # Split glued strength+unit+form: '10MGTABLETS'→'10 MG TABLETS', '10TAB'→'10 TAB'
    s = re.sub(r'(?i)(\d)\s*(' + _DESEG_UNITS + r')(' + _DESEG_FORMS + r')', r'\1 \2 \3', s)
    s = re.sub(r'(?i)(?<=[A-Za-z])(?=\d)', ' ', s)      # variant/brand → strength
    s = re.sub(r'(?i)(\d)(' + _DESEG_FORMS + r')\b', r'\1 \2', s)  # '10TAB'→'10 TAB'
    return re.sub(r'\s+', ' ', s).strip()


def is_substring_brand_reject(query_brand: str, candidate_brand: str) -> bool:
    """Reject a fuzzy brand hit when the candidate (master) brand is a STRICT
    substring of the input's brand-like query — i.e. the input carries extra
    letters that make it a different, more complete brand name.

    Asymmetric by design (per user rule): input 'AVIL' vs master 'VIL' is a
    reject (VIL ⊂ AVIL → different brand), but input 'VIL' vs master 'AVIL' is
    allowed (input is a sub-token of the master brand). Genuine typos such as
    ASOMAX↔ASOMEX are untouched (neither contains the other).
    """
    q = _brand_letters(query_brand)
    c = _brand_letters(candidate_brand)
    if not q or not c:
        return False
    return c != q and c in q


def find_unmatched_qualifier_tokens(input_name: str, brand: str,
                                    available_variants, available_sub_variants,
                                    brand_items=None) -> list:
    """Return alpha qualifier tokens the input places right after the brand that
    are NOT a known variant/sub-variant of that brand (nor a form/pack/unit, nor
    an abbreviation-prefix of a real attribute). A non-empty result means the
    input asked for a sub-brand the master does not carry (e.g. AMARYL 'MP',
    CETAPIN 'P'/'S', ACEGABA 'GRS') — the match should be rejected rather than
    silently collapsed onto the base/PLAIN product.
    """
    text = str(input_name or "").upper().replace("-", " ")
    brand_up = str(brand or "").upper().replace("-", " ").strip()

    def _find_boundary(hay, needle):
        """Find needle in hay only at WORD boundaries, so a brand that is a
        prefix of a longer word ('FERIUM INJ' inside 'FERIUM INJECTION') is not
        matched mid-word — which would leave a partial fragment ('ECTION')."""
        if not needle:
            return -1
        start = 0
        while True:
            i = hay.find(needle, start)
            if i < 0:
                return -1
            end = i + len(needle)
            before_ok = (i == 0) or (not hay[i - 1].isalnum())
            after_ok = (end >= len(hay)) or (not hay[end].isalnum())
            if before_ok and after_ok:
                return i
            start = i + 1

    idx = _find_boundary(text, brand_up)
    if idx >= 0:
        tail = text[idx + len(brand_up):]
    else:
        first = brand_up.split()[0] if brand_up.split() else ""
        pos = _find_boundary(text, first)
        if pos >= 0:
            tail = text[pos + len(first):]
        else:
            # Brand matched only fuzzily (e.g. typo ETOREC→ATOREC) — its name is
            # NOT in the input verbatim, so the input's leading token IS the
            # mis-spelled brand, not a qualifier. The fuzzy brand corresponds to
            # the FIRST token of the brand-like query only ('CETAPHIN'≈'CETAPIN');
            # any following tokens ('P') are qualifiers, so scan after just that
            # first token — scanning past the whole query would swallow the 'P'
            # and hide a missing-variant orphan.
            bq = extract_brand_like_query(input_name).upper()
            bq_first = bq.split()[0] if bq.split() else bq
            bpos = _find_boundary(text, bq_first)
            tail = text[bpos + len(bq_first):] if bpos >= 0 else ""

    # Tokens known to be legitimate (forms, packs, units, declared attributes).
    known = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES) | _pack_form_skip_tokens()
    known |= {"PLAIN"}
    known |= {str(v).upper() for v in (available_variants or set())}
    known |= {str(v).upper() for v in (available_sub_variants or set())}

    # Every alpha token that actually appears in this brand's product names.
    # If the input qualifier is anywhere in this vocabulary it names a real SKU
    # word (descriptive term or variant) — only a token absent from it signals a
    # sub-brand the master does not carry.
    brand_vocab = set()
    for it in (brand_items or []):
        for tok in re.findall(r"[A-Z]+", str(it.get("product", "")).upper()):
            brand_vocab.add(tok)
    # Prefix-expansion targets (FIX 4b: 'PRO' -> 'PROTECT'); exclude PLAIN and
    # form words so a bare 'P' can't masquerade as an abbreviation of 'PLAIN'.
    prefix_targets = (brand_vocab | known) - ({"PLAIN"} | set(FORM_HINT_TOKENS))
    brand_compact = _brand_letters(brand_up)

    orphans = []
    tail_tokens = re.findall(r"[A-Z0-9]+", tail)
    for ti, tok in enumerate(tail_tokens):
        if any(ch.isdigit() for ch in tok):
            break  # reached strength/pack — qualifiers only precede the number
        if tok in FORM_HINT_TOKENS:
            break  # reached the dosage form — variant qualifiers precede it
        canon = canonical_variant_match_token(tok)
        if tok in known or canon in known or tok in brand_vocab:
            continue
        # Glued sub-brand carried inside a product-name token (ESLO+'MET'=ESLOMET).
        if (brand_compact + tok) in brand_vocab:
            continue
        # Compact-join with an ADJACENT token forms a known variant/vocab word
        # ('AMARYL M V' → 'MV'): the space is OCR/spacing noise, not an orphan.
        prev_join = (tail_tokens[ti - 1] + tok) if ti > 0 else ""
        next_join = (tok + tail_tokens[ti + 1]) if ti + 1 < len(tail_tokens) else ""
        if (prev_join and (prev_join in known or prev_join in brand_vocab)) or \
           (next_join and (next_join in known or next_join in brand_vocab)):
            continue
        # Allow only true abbreviation-prefixes (len>=3) of a real attribute.
        if len(tok) >= 3 and any(t != tok and t.startswith(tok) for t in prefix_targets):
            continue
        orphans.append(tok)

    # Leading dose-modifier PREFIX before the brand ('SEMI AMARYL', 'HALF INDERAL')
    # names a distinct half-dose sub-brand. If the master carries no such prefixed
    # SKU for this brand, it's a different product that is simply absent here →
    # flag it (reject) instead of collapsing onto the plain brand. Only fires for
    # a verbatim brand hit (idx>0); manufacturer noise is handled elsewhere.
    if idx > 0:
        for tok in re.findall(r"[A-Z]+", text[:idx]):
            if tok in {"SEMI", "HALF"} and tok not in brand_vocab and tok not in orphans:
                orphans.append(tok)
    return orphans


# ============================================================================
# FIX 4b: filter_items_by_product_name_abbrev — fallback for Scenario B
# ----------------------------------------------------------------------------
# FIX 4 (find_abbrev_promoted_brand) only works when the master stores sub-
# brands as separate BRAND_NAME entries in brand_map ("CARDACE PROTECT" as a
# distinct brand key). Many masters instead keep all family members under
# ONE brand_name "CARDACE" and encode the sub-brand inside the product_name
# ("CARDACE PROTECT 2.5 TABLET"). In that case, brand_map has no "CARDACE
# PROTECT" key and FIX 4 returns None.
#
# This fallback scans the PRODUCT NAMES of the already-selected brand's
# items for an alpha token T (≥2 chars) such that T strictly extends the
# input's abbreviation token (input "PRO" → product "PROTECT"). If exactly
# one such expansion exists, items are filtered to products containing it.
#
# Structural; no brand / sub-brand names hardcoded. Does nothing when:
#   - the input abbrev matches an existing token verbatim (exact variant)
#   - the abbrev is a form/unit word (TAB, MG, …)
#   - the abbrev is <2 chars (handled by single-letter variant logic)
#   - more than one candidate expansion exists (ambiguous — don't guess)
# ============================================================================
def filter_items_by_product_name_abbrev(input_name: str, brand_name: str,
                                         items: list, verbose=False) -> list:
    if not brand_name or not items:
        return items

    input_upper = str(input_name or "").upper()
    matched_upper = str(brand_name).upper().replace("-", " ").strip()
    if not matched_upper:
        return items

    bm = re.search(r'\b' + re.escape(matched_upper) + r'\b', input_upper)
    if not bm:
        return items

    tail_tokens = re.findall(r'[A-Z]{2,}', input_upper[bm.end():])
    if not tail_tokens:
        return items
    brand_tokens = set(matched_upper.split())
    filtered_items = list(items)
    for abbrev in tail_tokens:
        if abbrev in _pack_form_skip_tokens():
            continue

        expansions = set()
        abbrev_as_exact_match = False
        for item in filtered_items:
            prod = str(item["product"]).upper().replace("-", " ")
            for tok in re.findall(r'\b[A-Z]{2,}\b', prod):
                if tok in brand_tokens or tok in _pack_form_skip_tokens():
                    continue
                if tok == abbrev:
                    abbrev_as_exact_match = True
                elif len(tok) > len(abbrev) and tok.startswith(abbrev):
                    expansions.add(tok)

        # If the abbrev already exists as an exact token in the current family,
        # keep scanning later tokens rather than forcing an expansion.
        if abbrev_as_exact_match:
            continue

        if len(expansions) != 1:
            continue

        expansion = next(iter(expansions))
        narrowed = [
            i for i in filtered_items
            if re.search(r'\b' + re.escape(expansion) + r'\b',
                         str(i["product"]).upper().replace("-", " "))
        ]
        if narrowed and len(narrowed) < len(filtered_items):
            if verbose:
                print(f"  [PRODUCT-NAME ABBREV] '{abbrev}' → '{expansion}': "
                      f"filtered {len(filtered_items)} → {len(narrowed)} items")
            filtered_items = narrowed

    return filtered_items


# =========================
# VARIANT DETECTION
# =========================
def detect_variants_in_input(
    input_name: str,
    brand_name: str,
    available_variants: set,
    compound_token_letters: set = None,
    verbose=False,
    brand_items: list = None,
) -> set:
    input_upper = normalize_ocr_numeric_noise(input_name)
    pack_positions = get_all_pack_positions(input_upper)

    def _single_letter_variant_only_in_pack(token: str) -> bool:
        token = str(token or "").strip().upper()
        if token not in {"S", "T", "C", "P"}:
            return False
        token_matches = list(re.finditer(r'\b' + re.escape(token) + r'\b', input_upper))
        if not token_matches:
            return False
        return all(
            is_position_in_pack_pattern(m.start(), m.end(), pack_positions)
            for m in token_matches
        )

    if brand_name:
        input_without_brand = strip_brand_from_input(input_upper, brand_name,
                                                     available_variants=available_variants)
    else:
        input_without_brand = input_upper
    input_without_brand = re.sub(r'\s+', ' ', input_without_brand).strip()
    normalized_input = normalize_variant_token(input_without_brand)
    compact_input_token_list = [
        compact_variant_token(token)
        for token in re.findall(r'[A-Z0-9.\-]+', input_without_brand)
        if compact_variant_token(token)
    ]
    compact_input_tokens = set(compact_input_token_list)
    for i in range(len(compact_input_token_list) - 1):
        compact_input_tokens.add(compact_input_token_list[i] + compact_input_token_list[i + 1])
    for i in range(len(compact_input_token_list) - 2):
        compact_input_tokens.add(compact_input_token_list[i] + compact_input_token_list[i + 1] + compact_input_token_list[i + 2])

    detected_variants = set()
    split_variants, _ = split_strength_and_variant(input_without_brand, available_variants, verbose=verbose)
    if compound_token_letters:
        split_variants -= compound_token_letters
    detected_variants.update(split_variants)
    combo_variants, combo_sub_variants = detect_compact_variant_strength_combo(
        input_without_brand, available_variants
    )
    if compound_token_letters:
        combo_variants -= compound_token_letters
    detected_variants.update(combo_variants)
    pair_variants, _ = detect_variant_numeric_pairs(
        input_without_brand, available_variants
    )
    detected_variants.update(pair_variants)

    words = re.findall(r'[A-Z0-9]+', input_without_brand)
    input_match_tokens = extract_variant_match_tokens(input_without_brand)
    input_descriptor_tokens = {
        token for token in input_match_tokens
        if token not in GENERIC_VARIANT_MATCH_TOKENS
    }
    input_forms_detected = detect_dosage_form_in_input(input_name, verbose=False)

    for variant in available_variants:
        if variant == "PLAIN":
            continue
        variant_upper = variant.upper()
        if len(variant_upper) == 1 and _single_letter_variant_only_in_pack(variant_upper):
            continue
        variant_compact = compact_variant_token(variant_upper)
        if compound_token_letters and variant_upper in compound_token_letters:
            continue
        if variant_compact and variant_compact in compact_input_tokens:
            detected_variants.add(variant)
            continue
        if variant_upper in words:
            detected_variants.add(variant)
            continue
        patterns = [
            r'[\-\s]' + re.escape(variant_upper) + r'[\-\s]',
            r'[\-\s]' + re.escape(variant_upper) + r'\b',
            r'\b' + re.escape(variant_upper) + r'[\-\s]',
        ]
        for pattern in patterns:
            if re.search(pattern, input_without_brand):
                detected_variants.add(variant)
                break

    for variant in available_variants:
        if variant == "PLAIN" or len(variant) <= 1:
            continue
        if compound_token_letters and variant.upper() in compound_token_letters:
            continue
        variant_upper = variant.upper()
        variant_normalized = normalize_variant_token(variant_upper)
        if variant_normalized and re.search(
            r'\b' + re.escape(variant_normalized).replace(r'\ ', r'\s+') + r'\b',
            normalized_input,
        ):
            detected_variants.add(variant)
            continue
        if re.search(r'\b' + re.escape(variant_upper) + r'\b', input_without_brand):
            detected_variants.add(variant)

    for variant in available_variants:
        if variant == "PLAIN":
            continue
        if len(str(variant).strip()) == 1 and _single_letter_variant_only_in_pack(variant):
            continue
        variant_match_tokens = extract_variant_match_tokens(variant)
        if not variant_match_tokens:
            continue
        if variant_match_tokens.issubset(input_match_tokens):
            if variant not in detected_variants:
                detected_variants.add(variant)
            continue
        if (
            len(variant_match_tokens) >= 2
            and input_descriptor_tokens
            and len(input_descriptor_tokens) >= 2
            and input_descriptor_tokens.issubset(variant_match_tokens)
        ):
            detected_variants.add(variant)

        variant_form_tokens = variant_match_tokens & GENERIC_VARIANT_MATCH_TOKENS
        variant_non_form_tokens = variant_match_tokens - GENERIC_VARIANT_MATCH_TOKENS
        if (
            variant not in detected_variants
            and variant_non_form_tokens
            and variant_non_form_tokens.issubset(input_match_tokens)
            and input_forms_detected
            and variant_form_tokens
            and forms_share_group(variant_form_tokens, input_forms_detected)
        ):
            detected_variants.add(variant)

    input_form_tokens = input_match_tokens & GENERIC_VARIANT_MATCH_TOKENS
    for form_token in sorted(input_form_tokens):
        form_variants = []
        for variant in available_variants:
            if variant == "PLAIN":
                continue
            variant_match_tokens = extract_variant_match_tokens(variant)
            if form_token in variant_match_tokens:
                form_variants.append(variant)
        if len(form_variants) == 1:
            variant = form_variants[0]
            # Guard: a variant whose VALUE merely contains the form word (e.g.
            # GLIPSOV 'M TABLET') must not be triggered by the input's form alone
            # — only if its non-form token(s) (the 'M') are also in the input.
            non_form = extract_variant_match_tokens(variant) - GENERIC_VARIANT_MATCH_TOKENS
            if non_form and not non_form.issubset(input_match_tokens):
                continue
            if variant not in detected_variants:
                detected_variants.add(variant)

    # Strength-as-VARIANT (EPOFER: VARIANT='10 000 - PFS', '3000', '40000', …).
    # Match such pure-dose variants by their numeric value, collapsing spaced
    # thousands ('10 000'→'10000') so an input written '10000' resolves to the
    # right SKU. Only fires for numeric-dose variants, leaving alpha variants
    # (M, XR, V 0.2 MG) to the logic above.
    input_collapsed = collapse_spaced_thousands(input_without_brand)
    input_strength_numbers = set(re.findall(r'\d+(?:\.\d+)?', input_collapsed))
    if input_strength_numbers:
        for variant in available_variants:
            if variant in detected_variants:
                continue
            vnum = numeric_strength_variant_value(variant)
            if vnum and vnum in input_strength_numbers:
                detected_variants.add(variant)

    # Injection-form VARIANT (AUGPEN VARIANT='I.V.', or 'VIAL'/'PFS'): when the
    # input carries an injection-group form, match these form-only variants so
    # the strength (sub-variant) can then pick the right SKU.
    _INJ_FORM_VAR = {"IV", "INJ", "INJECTION", "VIAL", "PFS", "AMP", "AMPS", "SYR", "SYRINGE"}
    if input_forms_detected and forms_share_group({"INJECTION"}, input_forms_detected):
        for variant in available_variants:
            if variant == "PLAIN" or variant in detected_variants:
                continue
            vt = {re.sub(r'[^A-Z]', '', t) for t in str(variant).upper().split()}
            vt = {t for t in vt if t}
            if vt and vt <= _INJ_FORM_VAR:
                # A pure injection-form variant (PFS/VIAL/I.V.) carries no dose of
                # its own. If the input names a dose, only adopt this variant when
                # some SKU under it actually has that dose as a sub-variant — else
                # the dose belongs to a different (PLAIN-variant) SKU that we must
                # NOT exclude (CISCURE '0.25 INJ' → the 0.25 INJECTION, not PFS).
                if input_strength_numbers and brand_items is not None:
                    has_dose = any(
                        str(it.get("sub_variant", "")).strip() in input_strength_numbers
                        for it in brand_items
                        if str(it.get("variant", "")).upper() == str(variant).upper()
                    )
                    if not has_dose:
                        continue
                detected_variants.add(variant)

    # Combo-dose → VARIANT inference: a full combination dose in the input that
    # exactly matches a product's combo sub_variant pins that product's variant,
    # even when the variant letter is absent ('EMPRI 12.5/500' → the 'M' SKU,
    # whose SUB_VARIANT is '12.5/500'). Full-combo match only, so a single number
    # never over-infers.
    if brand_items:
        input_combos = set(re.findall(r'(?<![\d./])\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)+(?![\d./])',
                                      input_without_brand))
        if input_combos:
            for it in brand_items:
                ivar = str(it.get("variant", "PLAIN")).strip().upper()
                isub = str(it.get("sub_variant", "")).strip().upper()
                if ivar and ivar != "PLAIN" and "/" in isub and isub in input_combos:
                    detected_variants.add(ivar)

    if not any(variant != 'PLAIN' for variant in detected_variants):
        detected_variants.add('PLAIN')

    if verbose:
        print(f"  Final detected variants: {sorted(detected_variants)}")
    return detected_variants


# =========================
# SUB-VARIANT DETECTION
# =========================
def detect_sub_variants_in_input(
    input_name: str,
    brand_name: str,
    available_sub_variants: set,
    available_variants: set,
    compound_token_digits: set = None,
    verbose=False,
) -> set:
    input_upper = normalize_ocr_numeric_noise(input_name)
    detected_sub_variants = set()
    if brand_name:
        input_without_brand = strip_brand_from_input(input_upper, brand_name,
                                                     available_variants=available_variants)
    else:
        input_without_brand = input_upper

    combo_dosage = extract_combination_dosage(input_upper)
    combo_components: set = set()
    if combo_dosage:
        detected_sub_variants.add(combo_dosage)
        combo_components = set(re.findall(r'\d+(?:\.\d+)?', combo_dosage))

    _, split_sub_variants = split_strength_and_variant(
        input_without_brand, available_variants, verbose=verbose
    )
    if compound_token_digits:
        split_sub_variants -= compound_token_digits
    if combo_components:
        split_sub_variants -= combo_components
    detected_sub_variants.update(split_sub_variants)
    _, combo_sub_variants = detect_compact_variant_strength_combo(
        input_without_brand, available_variants
    )
    if combo_components:
        combo_sub_variants -= combo_components
    detected_sub_variants.update(combo_sub_variants)
    _, pair_sub_variants = detect_variant_numeric_pairs(
        input_without_brand, available_variants
    )
    if combo_components:
        pair_sub_variants -= combo_components
    detected_sub_variants.update(pair_sub_variants)

    pack_positions = get_all_pack_positions(input_upper)
    sorted_sub_variants = sorted(
        [sv for sv in available_sub_variants if sv != "PLAIN"],
        key=lambda x: (len(str(x).replace('.', '')), len(str(x))),
        reverse=True,
    )
    matched_positions = []

    def is_position_matched(start, end):
        for ms, me in matched_positions:
            if (ms <= start < me) or (ms < end <= me):
                return True
        return False

    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants:
            continue
        if compound_token_digits and str(sub_var) in compound_token_digits:
            continue
        if combo_components and str(sub_var) in combo_components:
            continue
        sub_str = extract_numeric_strength(sub_var)
        if not sub_str:
            continue
        # Strength units only — NOT ML/L (those are liquid VOLUMES, not doses):
        # 'SUSPENSION 30 ML' must not match a '30 MG' strength sub-variant.
        unit_pattern = r'(?<![\d.])' + re.escape(sub_str) + r'\s*(?:MG|GM|MCG|G|M|IU|UNIT)(?:\b|/)'
        for m in re.finditer(unit_pattern, input_upper):
            if (not is_position_in_pack_pattern(m.start(), m.end(), pack_positions)
                    and not is_position_matched(m.start(), m.end())):
                detected_sub_variants.add(sub_var)
                matched_positions.append((m.start(), m.end()))
                break

    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants:
            continue
        if compound_token_digits and str(sub_var) in compound_token_digits:
            continue
        if combo_components and str(sub_var) in combo_components:
            continue
        sub_str = str(sub_var)
        pattern = r'\b' + re.escape(sub_str) + r'\b'
        for m in re.finditer(pattern, input_upper):
            context_start = max(0, m.start() - 5)
            context_end = min(len(input_upper), m.end() + 5)
            context = input_upper[context_start:context_end]
            if '/' in context:
                continue
            # Skip a number that is a liquid VOLUME ('30 ML') — not a strength.
            if re.match(r'\s*M?L\b', input_upper[m.end():]):
                continue
            # Skip a digit that is part of a DECIMAL ('5' inside '2.5', '2' before '.5').
            if (m.start() > 0 and input_upper[m.start()-1] == '.') or re.match(r'\.\d', input_upper[m.end():]):
                continue
            if (not is_position_in_pack_pattern(m.start(), m.end(), pack_positions)
                    and not is_position_matched(m.start(), m.end())):
                detected_sub_variants.add(sub_var)
                matched_positions.append((m.start(), m.end()))
                break

    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants:
            continue
        if compound_token_digits and str(sub_var) in compound_token_digits:
            continue
        if combo_components and str(sub_var) in combo_components:
            continue
        sub_str = str(sub_var)
        pattern = r'[-\s/]' + re.escape(sub_str) + r'\b'
        for m in re.finditer(pattern, input_upper):
            if re.match(r'\s*M?L\b', input_upper[m.end():]):
                continue   # liquid volume, not a strength
            if re.match(r'\.\d', input_upper[m.end():]):
                continue   # part of a decimal, not a strength
            if (not is_position_in_pack_pattern(m.start(), m.end(), pack_positions)
                    and not is_position_matched(m.start(), m.end())):
                detected_sub_variants.add(sub_var)
                matched_positions.append((m.start(), m.end()))
                break

    input_alpha_tokens = {t for t in re.findall(r'[A-Z]+', input_upper) if len(t) >= 3}
    input_alpha_signatures = {alpha_signature(t) for t in input_alpha_tokens}
    input_tokens_set_upper = set(re.findall(r'[A-Z]+', input_upper))
    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants or is_numeric_subvariant_like(sub_var):
            continue
        sub_var_upper = str(sub_var).strip().upper()
        if sub_var_upper not in input_tokens_set_upper:
            continue
        sub_signature = alpha_signature(sub_var)
        if len(sub_signature) >= 3 and sub_signature in input_alpha_signatures:
            detected_sub_variants.add(sub_var)

    _SV_ABBREV = {"FORT": "FORTE", "FORTS": "FORTE"}
    for abbrev, expanded in _SV_ABBREV.items():
        if abbrev in input_tokens_set_upper and expanded in available_sub_variants:
            detected_sub_variants.add(expanded)

    # Strength-vs-pack rescue: a dose can look like a pack count ('EXDUO 100 TAB'
    # reads '100' as 100 tablets). If no numeric strength was detected yet but
    # exactly ONE input number equals an available numeric sub-variant (a real
    # strength of this brand), claim it as the strength even though pack-parsing
    # swallowed it.
    if not any(is_numeric_subvariant_like(s) for s in detected_sub_variants if s != "PLAIN"):
        # Collect input numbers but EXCLUDE liquid volumes ('30 ML' is a pack,
        # not a dose) so a suspension volume can't masquerade as a strength.
        rescue_text = normalize_ocr_numeric_noise(input_without_brand)
        input_numbers = set()
        for m in re.finditer(r'(\d+(?:\.\d+)?)\s*([A-Z]*)', rescue_text):
            unit = m.group(2)
            if unit == "L" or unit.startswith("ML"):
                continue
            input_numbers.add(m.group(1))
        strength_subs = {s for s in available_sub_variants
                         if is_numeric_subvariant_like(s) and "/" not in str(s)}
        hits = {s for s in strength_subs if str(s) in input_numbers}
        if len(hits) == 1:
            detected_sub_variants.add(next(iter(hits)))

    detected_sub_variants.add("PLAIN")
    if verbose:
        print(f"  Final detected sub-variants: {sorted(detected_sub_variants)}")
    return detected_sub_variants


# ============================================================================
# FIX 5: filter_items_by_variant — most-specific variant wins
# ----------------------------------------------------------------------------
# When detected variants form a subset chain (e.g. {M, M FORTE}), prefer the
# superset (M FORTE) — fixes AMARYL M FORTE being mapped to plain AMARYL M.
# Structural word-set subset check; no hardcoded variant names.
# Falls back to the old inclusive behavior if specific-only yields nothing.
# ============================================================================
def _specific_detected_variants(detected_variants: set) -> set:
    specific = {v for v in detected_variants if v and v != "PLAIN"}
    if len(specific) <= 1:
        return detected_variants
    word_sets = {v: set(str(v).upper().split()) for v in specific}
    kept = set()
    for v, vset in word_sets.items():
        is_subset_of_other = any(
            v != other and vset < word_sets[other]
            for other in word_sets
        )
        if not is_subset_of_other:
            kept.add(v)
    if "PLAIN" in detected_variants:
        kept.add("PLAIN")
    return kept if kept else detected_variants


def filter_items_by_variant(items: list, detected_variants: set,
                             reference_items: list = None, verbose=False) -> list:
    if verbose:
        print("\n" + "="*80)
        print("STEP 2: VARIANT-BASED FILTERING")
        print("="*80)
        print(f"Detected variants in input: {sorted(detected_variants)}")
        print(f"Total products before filtering: {len(items)}")

    # --- FIX 5: narrow to most-specific variants first
    specific_variants = _specific_detected_variants(detected_variants)
    if verbose and specific_variants != detected_variants:
        print(f"Narrowed to most-specific: {sorted(specific_variants)}")

    def _partition(variant_set):
        vmatches, plains = [], []
        for item in items:
            iv = item.get("variant", "PLAIN").upper()
            if iv == "PLAIN":
                plains.append(item)
            elif iv in variant_set:
                vmatches.append(item)
        return vmatches, plains

    variant_matches, plain_products = _partition(specific_variants)

    # Fallback to broader set if specific-only filter emptied the list
    if not variant_matches and specific_variants != detected_variants:
        if verbose:
            print(f"⚠ No items for specific variants — falling back to full detected set")
        variant_matches, plain_products = _partition(detected_variants)

    if variant_matches:
        filtered = variant_matches
        if verbose:
            print(f"\n✓ Found {len(variant_matches)} specific variant matches")
            print(f"  Excluding {len(plain_products)} PLAIN products")
    elif reference_items and reference_items is not items:
        ref_matches, ref_plains = [], []
        for item in reference_items:
            iv = item.get("variant", "PLAIN").upper()
            if iv == "PLAIN":
                ref_plains.append(item)
            elif iv in specific_variants or iv in detected_variants:
                ref_matches.append(item)
        if ref_matches:
            filtered = ref_matches
            if verbose:
                print(f"\n✓ Recovered {len(ref_matches)} variant matches from full brand list")
        else:
            filtered = plain_products if plain_products else ref_plains
            if verbose:
                print(f"\n⚠ No variant matches; using {len(filtered)} PLAIN as fallback")
    else:
        filtered = plain_products
        if verbose:
            print(f"\n⚠ No variant matches; using {len(plain_products)} PLAIN as fallback")

    if verbose:
        print(f"\nProducts after variant filtering: {len(filtered)}")
        if variant_matches:
            variants_kept = {item.get("variant", "PLAIN") for item in filtered}
            for var in sorted(variants_kept):
                count = sum(1 for i in filtered if i.get("variant", "PLAIN") == var)
                print(f"  - {var}: {count} products")

    if not filtered:
        return items
    return filtered


def filter_by_subvariant_in_product_name(items: list, detected_sub_variants: set, verbose=False) -> list:
    numeric_sub_variants = {sv for sv in detected_sub_variants
                            if sv != 'PLAIN' and '/' not in str(sv)}
    if not numeric_sub_variants:
        return items
    matches_with_scores = []
    for item in items:
        product_name = item["product"].upper()
        best_score = 0
        matched_svs = set()
        for sub_var in numeric_sub_variants:
            sub_str = str(sub_var)
            score = 0
            if re.search(r'\b' + re.escape(sub_str) + r'\s*(?:MG|GM|ML|MCG|IU)', product_name):
                score = 100
            elif re.search(r'\b' + re.escape(sub_str) + r'\b', product_name):
                if not re.search(r'\b' + re.escape(sub_str) + r'\s*[X*×]', product_name):
                    score = 90
            elif re.search(r'[A-Z]' + re.escape(sub_str), product_name):
                score = 75
            elif sub_str in product_name:
                score = 50
            if score > best_score:
                best_score = score
                matched_svs = {sub_var}
            elif score == best_score and score > 0:
                matched_svs.add(sub_var)
        if best_score >= 75:
            matches_with_scores.append((item, best_score, matched_svs))
    if matches_with_scores:
        matches_with_scores.sort(key=lambda x: x[1], reverse=True)
        best_score = matches_with_scores[0][1]
        return [item for item, score, _ in matches_with_scores if score == best_score]
    return items


# ============================================================================
# FIX 8: calculate_product_name_priority_score — structural-only penalty
# ----------------------------------------------------------------------------
# Drops hardcoded _SKIP_PENALTY list. Identifies pack shapes and unit
# suffixes by regex; penalty magnitude is 40 per extra text word (was 10)
# so ghost variants (BETA, AMH, EZ, LAR) actually lose against plain.
# ============================================================================
def calculate_product_name_priority_score(product_name: str, input_name: str) -> int:
    _TOKEN_SPLIT = r'[\s\-+/]+'

    input_upper = input_name.upper().strip()
    input_words = {w for w in re.split(_TOKEN_SPLIT, input_upper) if w}
    input_raw_numbers = set(re.findall(r'\d+(?:\.\d+)?', input_upper))

    product_upper = product_name.upper().strip()
    product_words = {w for w in re.split(_TOKEN_SPLIT, product_upper) if w}

    _PACK_SHAPE = re.compile(
        r'^\d+[Xx*]\d+[A-Z]?$|^\d+[STCP]$|^\d+PCS?$|^\d+NOS?$',
        re.IGNORECASE,
    )
    _UNIT_SUFFIX = re.compile(r'^(?:MG|ML|GM|MCG|IU|KG)\.?$', re.IGNORECASE)

    def _is_structural(token: str) -> bool:
        if _PACK_SHAPE.match(token) or _UNIT_SUFFIX.match(token):
            return True
        bare = token.rstrip('.')
        if bare in FORM_HINT_TOKENS:
            return True
        glued = re.match(r'^\d+([A-Z]+)$', token)
        if glued and glued.group(1) in FORM_HINT_TOKENS:
            return True
        return False

    score = 0
    extra_words = product_words - input_words
    if extra_words:
        extra_words = {w for w in extra_words if not _is_structural(w)}
        numeric_re = re.compile(r'^\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)*$')
        extra_numeric = {w for w in extra_words if numeric_re.match(w)}
        extra_numeric_penalized = {w for w in extra_numeric if w not in input_raw_numbers}
        extra_text = extra_words - extra_numeric
        score -= len(extra_text) * 40
        score -= len(extra_numeric_penalized) * 25

    if "SEMI" not in input_words and "SEMI" in product_upper:
        score -= 20

    if "MP" in input_words:
        if "MP" in product_words:
            product_word_list = product_upper.split()
            try:
                mp_pos_in_product = product_word_list.index("MP")
                if mp_pos_in_product <= 1:
                    score += 20
                else:
                    score -= 10
            except ValueError:
                pass

    return score


def prioritize_by_product_name_match(items: list, input_name: str, verbose=False) -> list:
    if len(items) <= 1:
        return items
    scored_items = [(item, calculate_product_name_priority_score(item["product"], input_name))
                    for item in items]
    scored_items.sort(key=lambda x: x[1], reverse=True)
    if verbose:
        print(f"\n  Product name prioritization scores:")
        for item, score in scored_items:
            print(f"    {score}: {item['product'][:50]}")
    return [item for item, _ in scored_items]


# ============================================================================
# FIX 6: filter_items_by_sub_variant — combo first-component fallback
# ----------------------------------------------------------------------------
# When input has "10/5" combo and master has no literal combo SKU, fall back
# to the FIRST numeric component ("CARDACE AM 10/5" → "CARDACE AM 10").
# Structural split on '/'; master's actual SKUs decide the match.
# ============================================================================
def filter_items_by_sub_variant(items: list, detected_sub_variants: set, input_name: str = "",
                                 reference_items: list = None, verbose=False) -> list:
    if verbose:
        print("\n" + "="*80)
        print("STEP 3: SUB-VARIANT-BASED FILTERING")
        print("="*80)
        print(f"Detected sub-variants: {sorted(detected_sub_variants)}")

    specific_detected = {sv for sv in detected_sub_variants if sv != "PLAIN"}
    match_targets = specific_detected if specific_detected else set(detected_sub_variants)
    detected_forms = detect_dosage_form_in_input(input_name, verbose=False) if input_name else set()
    alpha_detected = {
        sv for sv in specific_detected
        if re.search(r'[A-Z]', str(sv)) and not is_numeric_subvariant_like(str(sv))
    }

    combo_dosages = {sv for sv in detected_sub_variants if '/' in str(sv)}
    if combo_dosages:
        combo_matches = []
        for item in items:
            for combo in combo_dosages:
                if product_contains_combination(item["product"], combo):
                    combo_matches.append(item)
                    break
        if combo_matches:
            if verbose:
                print(f"\n✓ {len(combo_matches)} literal combo matches")
            return combo_matches

        # --- FIX 6: structural component fallback (priority-ordered)
        # When the master has no literal "40/800" combo SKU, split the combo
        # on '/' and try each component as an EXACT sub-variant in priority
        # order: the FIRST part wins; only if it finds nothing do we fall back
        # to the SECOND part (and any further parts). Each part is checked
        # first against the SUB_VARIANT column, then as a whole-word match in
        # the product name.
        ordered_components = []
        for combo in combo_dosages:
            parts = [p.strip() for p in str(combo).split('/') if p.strip()]
            for idx, part in enumerate(parts):
                if idx >= len(ordered_components):
                    ordered_components.append(set())
                ordered_components[idx].add(part.upper())

        for position, components in enumerate(ordered_components):
            if not components:
                continue
            # 1) exact match on the SUB_VARIANT column
            comp_col_matches = [
                item for item in items
                if str(item.get("sub_variant", "PLAIN")).strip().upper() in components
            ]
            if comp_col_matches:
                if verbose:
                    print(f"\n✓ No literal combo — using component #{position + 1} "
                          f"{sorted(components)} (sub_variant column)")
                if input_name:
                    return prioritize_by_product_name_match(comp_col_matches, input_name, verbose=verbose)
                return comp_col_matches
            # 2) exact whole-word match inside the PRODUCT NAME
            comp_name_matches = []
            for item in items:
                pname = item["product"].upper()
                for comp in components:
                    if re.search(r'\b' + re.escape(str(comp)) + r'\b', pname):
                        comp_name_matches.append(item)
                        break
            if comp_name_matches:
                if verbose:
                    print(f"\n✓ No literal combo — using component #{position + 1} "
                          f"{sorted(components)} (product name)")
                if input_name:
                    return prioritize_by_product_name_match(comp_name_matches, input_name, verbose=verbose)
                return comp_name_matches

    if alpha_detected:
        alpha_matches = [i for i in items if i.get("sub_variant", "PLAIN").upper() in alpha_detected]
        if alpha_matches:
            narrowed = filter_by_subvariant_in_product_name(alpha_matches, detected_sub_variants,
                                                            verbose=verbose)
            if input_name:
                return prioritize_by_product_name_match(narrowed, input_name, verbose=verbose)
            return narrowed

    exact_matches = []
    plain_products = []
    numeric_detected = {
        sv for sv in detected_sub_variants
        if sv != 'PLAIN' and is_numeric_subvariant_like(str(sv))
    }
    for item in items:
        item_sv = item.get("sub_variant", "PLAIN").upper()
        if item_sv in match_targets:
            exact_matches.append(item)
        elif item_sv == "PLAIN":
            plain_products.append(item)

    if exact_matches:
        if not specific_detected and detected_forms:
            # Narrow the EXACT sub-variant matches by form — do NOT widen to all
            # items (that let e.g. COBAFORTE 'CD3 XL' outrank the plain 'CD3').
            form_matches = [
                item for item in exact_matches
                if (
                    detect_dosage_form_in_product(item["product"]) & detected_forms
                    or forms_share_group(
                        detect_dosage_form_in_product(item["product"]),
                        detected_forms,
                    )
                )
            ]
            if form_matches:
                if input_name:
                    return prioritize_by_product_name_match(form_matches, input_name, verbose=verbose)
                return form_matches
        # When several numeric strengths matched (e.g. 'AM 2.5 ... 10'), the real
        # strength sits next to the variant and pack/qty trails — keep the
        # candidate(s) whose strength appears EARLIEST in the input.
        numeric_exact = {str(i.get("sub_variant", "")).strip() for i in exact_matches
                         if is_numeric_subvariant_like(str(i.get("sub_variant", "")))}
        if input_name and len(numeric_exact) > 1:
            up = normalize_ocr_numeric_noise(input_name)
            def _pos(sub):
                m = re.search(r'(?<![\d.])' + re.escape(str(sub)) + r'(?![\d.])', up)
                return m.start() if m else 10**6
            best_pos = min(_pos(s) for s in numeric_exact)
            earliest = [i for i in exact_matches
                        if _pos(str(i.get("sub_variant", "")).strip()) == best_pos]
            if earliest and len(earliest) < len(exact_matches):
                exact_matches = earliest
        if input_name:
            return prioritize_by_product_name_match(exact_matches, input_name, verbose=verbose)
        return exact_matches

    # Single numeric input (e.g. 'EMPRI L 10') against COMBO sub-variants
    # ('10/5','25/5'): no exact whole-combo match exists, so keep the combo(s)
    # whose FIRST component equals a detected number (10 → 10/5, not 25/5).
    if numeric_detected:
        first_comp = [
            item for item in items
            if '/' in str(item.get("sub_variant", ""))
            and str(item.get("sub_variant", "")).split('/')[0].strip().upper() in numeric_detected
        ]
        if first_comp and len(first_comp) < len(items):
            if verbose:
                print(f"\n✓ Single number → combo first-component match: "
                      f"{[i['product'] for i in first_comp]}")
            if input_name:
                return prioritize_by_product_name_match(first_comp, input_name, verbose=verbose)
            return first_comp

    if numeric_detected:
        name_matches = filter_by_subvariant_in_product_name(items, detected_sub_variants, verbose=verbose)
        if name_matches and len(name_matches) < len(items):
            if input_name:
                return prioritize_by_product_name_match(name_matches, input_name, verbose=verbose)
            return name_matches

    if numeric_detected and plain_products:
        name_matches = filter_by_subvariant_in_product_name(plain_products, detected_sub_variants,
                                                             verbose=verbose)
        if name_matches and len(name_matches) < len(plain_products):
            return name_matches

    if plain_products:
        return plain_products
    return items


# ============================================================================
# FIX 7: detect_dosage_form_in_input — expanded OCR-safe tokens
# ----------------------------------------------------------------------------
# Accepts SYP., SUSP., INJ., ING (OCR for INJ), VIAL, PFS, AMPS, SYR,
# glued digit+form tokens like 15CAP, 10INJ, and trailing-dot variants.
# Purely pattern-based; no hardcoded brand names.
# ============================================================================
def detect_dosage_form_in_input(input_name: str, verbose=False) -> set:
    input_upper = input_name.upper()
    detected_forms = set()
    form_patterns = [
        # TABLET — glued forms 10TAB/15TABS, standalone TAB/TABS/TA, with optional trailing dot
        (r'\bTABLETS?\b|\bTABS?\.?\b|\bTA\b|\d+TABLETS?|\d+TABS?\b', 'TABLET'),
        # CAPSULE — glued 15CAP/10CAPS, standalone
        (r'\bCAPSULES?\b|\bCAPS?\.?\b|\d+CAPSULES?|\d+CAPS?\b', 'CAPSULE'),
        # INJECTION — INJ, INJ., ING (OCR), VIAL, PFS, AMPS, SYR, I.V./IV, glued \d+INJ/SYR
        (r'\bINJECTIONS?\b|\bINJ\.?\b|\bING\.?\b|\d+INJ\b|\bVIAL\b|\bPFS\b|\bAMPS?\b|\bSYR\.?\b|\d+SYR\.?\b|\bI\.?V\.?\b',
         'INJECTION'),
        # SYP/SYRUP = SUSPENSION (one liquid-oral form per spec), incl glued \d+SYP
        (r'\bSUSPENSIONS?\b|\bSUSP\.?\b|\bSUSPN\b|\bSYRUPS?\b|\bSYP\.?\b|\d+SYP\b', 'SUSPENSION'),
        (r'\bDROPS?\b|\d+DROPS?\b', 'DROPS'),
        (r'\bCREAMS?\b|\bCRM\.?\b', 'CREAM'),
        (r'\bOINTMENTS?\b|\bOI\.?\b|\bOINT\.?\b', 'OINTMENT'),
        (r'\bGELS?\b', 'GEL'),
        (r'\bSOLUTIONS?\b|\bSOL\.?\b', 'SOLUTION'),
        (r'\bPOWDERS?\b|\bPOW\.?\b', 'POWDER'),
        (r'\bBARS?\b|\bSYNDET\b', 'BAR'),
        (r'\bGRANULES?\b|\bGRANUL\b|\bGRANULS\b', 'GRANULES'),
        (r'\bSACHETS?\b', 'SACHET'),
        (r'\bLOTIONS?\b|\bLOT\.?\b', 'LOTION'),
        (r'\bTONICS?\b', 'TONIC'),
    ]
    for pattern, form_name in form_patterns:
        if re.search(pattern, input_upper):
            detected_forms.add(form_name)

    # OCR sometimes truncates syrup/suspension cues to a bare "SY" before an ML
    # pack, e.g. "XT + SY 200 ML". Treat that as a liquid-oral hint without
    # changing any non-liquid families.
    if (
        not ({"SUSPENSION", "DROPS", "SOLUTION", "TONIC"} & detected_forms)
        and re.search(r'\bSY\b', input_upper)
        and re.search(r'\b\d+(?:\.\d+)?\s*ML\b', input_upper)
    ):
        detected_forms.add("SUSPENSION")

    pack_positions = get_all_pack_positions(input_upper)

    def _all_matches_inside_pack(pattern: str) -> bool:
        matches = list(re.finditer(pattern, input_upper))
        if not matches:
            return False
        return all(
            is_position_in_pack_pattern(m.start(), m.end(), pack_positions)
            for m in matches
        )

    if "CAPSULE" in detected_forms and "TABLET" in detected_forms:
        tablet_pack_only = _all_matches_inside_pack(
            r'\bTABLETS?\b|\bTABS?\.?\b|\bTA\b|\d+TABLETS?|\d+TABS?\b'
        )
        capsule_pack_only = _all_matches_inside_pack(
            r'\bCAPSULES?\b|\bCAPS?\.?\b|\d+CAPSULES?|\d+CAPS?\b'
        )
        if tablet_pack_only and not capsule_pack_only:
            detected_forms.discard("TABLET")
        elif capsule_pack_only and not tablet_pack_only:
            detected_forms.discard("CAPSULE")
    return detected_forms


def detect_dosage_form_in_product(product_name: str) -> set:
    product_upper = product_name.upper()
    detected_forms = set()
    form_patterns = [
        # tab/tabs/tab. = tablet, cap/caps/cap. = capsule
        (r'\bTABLETS?\b|\bTABS?\.?\b', 'TABLET'),
        (r'\bCAPSULES?\b|\bCAPS?\.?\b', 'CAPSULE'),
        # inj = vial (and amp/pfs/I.V.) — all injection
        (r'\bINJECTIONS?\b|\bINJ\.?\b|\bVIAL\b|\bPFS\b|\bAMPS?\b|\bSYR\.?\b|\bI\.?V\.?\b', 'INJECTION'),
        # syp/syrup = suspension (treated as one liquid-oral form per spec)
        (r'\bSUSPENSIONS?\b|\bSUSP\.?\b|\bSUSPN\b|\bSYRUPS?\b|\bSYP\.?\b', 'SUSPENSION'),
        (r'\bDROPS?\b', 'DROPS'),
        (r'\bCREAMS?\b', 'CREAM'),
        (r'\bOINTMENTS?\b', 'OINTMENT'),
        (r'\bGELS?\b', 'GEL'),
        (r'\bSOLUTIONS?\b|\bSOL\b', 'SOLUTION'),
        (r'\bPOWDERS?\b', 'POWDER'),
        (r'\bBARS?\b|\bSYNDET\b', 'BAR'),
        (r'\bGRANULES?\b|\bGRANUL\b|\bGRANULS\b', 'GRANULES'),
        (r'\bSACHETS?\b', 'SACHET'),
        (r'\bLOTIONS?\b', 'LOTION'),
        (r'\bTONICS?\b', 'TONIC'),
    ]
    for pattern, form_name in form_patterns:
        if re.search(pattern, product_upper):
            detected_forms.add(form_name)
    return detected_forms


_FORM_GROUPS = {
    'TABLET':     'solid_oral',
    'CAPSULE':    'solid_oral',
    'SUSPENSION': 'liquid_oral',
    'SYRUP':      'liquid_oral',
    'SOLUTION':   'liquid_oral',
    'DROPS':      'liquid_oral',
    'TONIC':      'liquid_oral',
    'INJECTION':  'parenteral',
    'CREAM':      'topical',
    'GEL':        'topical',
    'OINTMENT':   'topical',
    'LOTION':     'topical',
    'POWDER':     'powder',
    'GRANULES':   'powder',
    'SACHET':     'powder',
    'BAR':        'bar',
}


def forms_share_group(product_forms: set, detected_forms: set) -> bool:
    product_groups = {_FORM_GROUPS[f] for f in product_forms if f in _FORM_GROUPS}
    detected_groups = {_FORM_GROUPS[f] for f in detected_forms if f in _FORM_GROUPS}
    return bool(product_groups & detected_groups)


def prioritize_by_dosage_form(items: list, detected_forms: set, verbose=False) -> list:
    if not detected_forms:
        return items
    if verbose:
        print("\n" + "="*80)
        print("STEP 4.5: DOSAGE FORM PRIORITIZATION")
        print("="*80)
        print(f"Detected forms in input: {sorted(detected_forms)}")

    detected_groups = {_FORM_GROUPS[f] for f in detected_forms if f in _FORM_GROUPS}
    exact_form_matches = []
    other_products = []
    for item in items:
        product_forms = detect_dosage_form_in_product(item["product"])
        if product_forms & detected_forms:
            exact_form_matches.append(item)
        else:
            other_products.append(item)

    if exact_form_matches:
        if verbose:
            print(f"\n✓ Found {len(exact_form_matches)} products matching dosage form")
        return exact_form_matches

    if detected_groups:
        non_conflicting = []
        excluded = []
        for item in items:
            product_forms = detect_dosage_form_in_product(item["product"])
            product_groups = {_FORM_GROUPS[f] for f in product_forms if f in _FORM_GROUPS}
            if product_groups and product_groups.isdisjoint(detected_groups):
                excluded.append(item)
            else:
                non_conflicting.append(item)
        if non_conflicting and len(non_conflicting) < len(items):
            if verbose:
                print(f"\n⚠ No exact form match — excluded {len(excluded)} conflicting-form products")
            return non_conflicting

    return items


def _form_compatible(items: list, detected_forms: set) -> list:
    """Items whose dosage-form GROUP does not conflict with the input's.

    Products with NO detectable form count as COMPATIBLE: many correct master
    names carry no form token ('METPURE H 50 TABELTS', 'C-ZID 1GM',
    'MATERNA-HMG 150 I.U.'), so discarding them is what makes a pure
    form-first ordering destructive. Used only by the form-aware retries in
    Steps 2/3/3.5; prioritize_by_dosage_form is untouched.
    """
    detected_groups = {_FORM_GROUPS[f] for f in detected_forms if f in _FORM_GROUPS}
    if not detected_groups:
        return list(items)
    keep = []
    for item in items:
        product_groups = {_FORM_GROUPS[f]
                          for f in detect_dosage_form_in_product(item['product'])
                          if f in _FORM_GROUPS}
        if not product_groups or (product_groups & detected_groups):
            keep.append(item)
    return keep


# ============================================================================
# PACK SIZE FILTER — respecting the RULE:
#   - Input has no pack indicator → return products without pack (BASE).
#   - Input has a pack → prefer exact pack match; fall back to BASE if none.
# (This is the existing behavior from updated_ts_3, which aligns with the rule.)
# ============================================================================
def filter_items_by_pack_size(items: list, input_pack_size: str,
                               reference_items: list = None, verbose=False) -> list:
    if verbose:
        print("\n" + "="*80)
        print("STEP 5: PACK SIZE FILTERING")
        print("="*80)
        print(f"Input pack size: '{input_pack_size or 'NONE'}'")

    products_with_pack = []
    products_without_pack = []
    for item in items:
        pack = item["pack_size"].strip()
        if pack and pack not in ("", "nan", "NaN", "None"):
            pack_num = parse_pack_size_value(pack)
            if pack_num is not None:
                products_with_pack.append((item, pack_num))
            else:
                products_without_pack.append(item)
        else:
            products_without_pack.append(item)

    if verbose:
        print(f"  Products with pack size:    {len(products_with_pack)}")
        print(f"  Products without pack size: {len(products_without_pack)}")

    # RULE: no pack in input → return base (unpacked) products
    if not input_pack_size:
        if products_without_pack:
            if verbose:
                print(f"\n✓ Input has no pack — returning {len(products_without_pack)} BASE products")
            return products_without_pack
        if products_with_pack:
            products_with_pack.sort(key=lambda x: x[1])
            return [item for item, _ in products_with_pack[:1]]
        return items

    try:
        input_pack_num = float(input_pack_size)
    except ValueError:
        return products_without_pack if products_without_pack else items

    exact_matches = [item for item, pn in products_with_pack if pn == input_pack_num]
    if exact_matches:
        return exact_matches

    # No exact pack match → RULE: fall back to base (unpacked)
    if products_without_pack:
        if verbose:
            print(f"\n✓ No exact pack match — falling back to BASE products")
        return products_without_pack

    return items


def apply_galact_granules_rules(
    items: list,
    brand: str,
    brand_scope_items: list,
    input_name: str,
    raw_input_name: str,
    detected_variants: set,
    detected_sub_variants: set,
    input_pack_size: str,
    verbose=False,
) -> list:
    if str(brand or "").strip().upper() != "GALACT":
        return items

    input_upper = normalize_ocr_numeric_noise(input_name)
    raw_input_upper = normalize_ocr_numeric_noise(raw_input_name or input_name)
    specific_variants = {
        str(v).strip().upper() for v in detected_variants
        if str(v).strip().upper() not in ("", "PLAIN")
    }
    specific_sub_variants = {
        str(sv).strip().upper() for sv in detected_sub_variants
        if str(sv).strip().upper() not in ("", "PLAIN")
    }
    powder_like_input = (
        ("POWD" in raw_input_upper or "POWDER" in raw_input_upper)
        and "PLUS" not in raw_input_upper
        and "NUTRI" not in raw_input_upper
        and "BAR" not in raw_input_upper
        and "MOMS" not in raw_input_upper
    )
    granules_mentioned = (
        "GRANUL" in input_upper
        or "GRANUL" in raw_input_upper
        or "GRANULES" in specific_variants
        or powder_like_input
    )

    def is_galact_granules_item(item: dict) -> bool:
        variant_upper = str(item.get("variant", "PLAIN")).strip().upper()
        variant_compact = compact_variant_token(variant_upper)
        product_upper = str(item.get("product", "")).strip().upper()
        return (
            variant_compact in {"GRANULES", "GARNULES"}
            or "GRANUL" in product_upper
        )

    granules_scope = [
        item for item in (brand_scope_items or items)
        if is_galact_granules_item(item)
    ]
    if not granules_scope:
        return items

    flavor_aliases = {
        "CHOCOLATE": {"CHOCOLATE", "CHOCO", "CHOC", "CH"},
        "KESAR": {"KESAR", "KES", "K"},
        "ELAICHI": {"ELAICHI", "ELA", "ILA", "E"},
    }

    explicit_flavor = ""
    flavor_tokens = set(re.findall(r'[A-Z]+', raw_input_upper))
    flavor_tokens.update(re.findall(r'\(([A-Z]{1,4})\)', str(raw_input_name or "").upper()))
    for flavor_name, aliases in flavor_aliases.items():
        if flavor_tokens & aliases:
            explicit_flavor = flavor_name
            break

    effective_pack_size = input_pack_size
    if not effective_pack_size:
        pack_values = {
            int(pack_num)
            for item in granules_scope
            for pack_num in [parse_pack_size_value(item.get("pack_size", ""))]
            if pack_num is not None and float(pack_num).is_integer()
        }
        raw_numbers = {
            int(float(num))
            for num in re.findall(r'\d+(?:\.\d+)?', raw_input_upper)
        }
        inferred_pack_values = sorted(pack_values & raw_numbers)
        if len(inferred_pack_values) == 1:
            effective_pack_size = str(inferred_pack_values[0])

    def _exact_pack_matches(pool: list, pack_size_value: str) -> list:
        if not pack_size_value:
            return []
        try:
            pack_num = float(pack_size_value)
        except ValueError:
            return []
        return [
            item for item in pool
            if parse_pack_size_value(item.get("pack_size", "")) == pack_num
        ]

    if explicit_flavor:
        flavor_matches = [
            item for item in granules_scope
            if str(item.get("sub_variant", "PLAIN")).strip().upper() == explicit_flavor
        ]
        if flavor_matches:
            exact_flavor_pack = _exact_pack_matches(flavor_matches, effective_pack_size)
            if exact_flavor_pack:
                if verbose:
                    print(f"\n✓ GALACT GRANULES rule: explicit flavor {explicit_flavor} with exact pack")
                return exact_flavor_pack
            no_pack_flavor = [
                item for item in flavor_matches
                if not str(item.get("pack_size", "")).strip()
            ]
            if no_pack_flavor:
                if verbose:
                    print(f"\n✓ GALACT GRANULES rule: explicit flavor {explicit_flavor} → unpacked flavor match")
                return no_pack_flavor
            if verbose:
                print(f"\n✓ GALACT GRANULES rule: explicit flavor {explicit_flavor} → flavor family match")
            return flavor_matches

    if not effective_pack_size:
        if not granules_mentioned or specific_sub_variants:
            return items
        base_granules = [
            item for item in granules_scope
            if not str(item.get("pack_size", "")).strip()
            and str(item.get("sub_variant", "PLAIN")).strip().upper() == "PLAIN"
        ]
        if base_granules:
            if verbose:
                print("\n✓ GALACT GRANULES rule: no pack → base unpacked product")
            return base_granules
        return items

    exact_pack_granules = _exact_pack_matches(granules_scope, effective_pack_size)
    if not exact_pack_granules:
        return items

    if verbose:
        print(f"\n✓ GALACT GRANULES rule: exact pack {effective_pack_size} → {len(exact_pack_granules)} candidates")

    if specific_sub_variants:
        return exact_pack_granules

    if not specific_variants or specific_variants == {"GRANULES"}:
        elaichi_matches = [
            item for item in exact_pack_granules
            if str(item.get("sub_variant", "PLAIN")).strip().upper() == "ELAICHI"
        ]
        if elaichi_matches:
            if verbose:
                print("✓ GALACT GRANULES rule: missing flavor → prefer ELAICHI")
            return elaichi_matches

    return exact_pack_granules


# =========================
# BUILD BRAND CONTEXT STRING
# =========================
def build_brand_context_string(items: list, verbose=False) -> str:
    lines = []
    for item in items[:100]:
        prod_norm = simple_normalize_for_llm(item["product"])
        pack = item["pack_size"].strip()
        if pack and pack not in ("", "nan", "NaN"):
            lines.append(f"{prod_norm} [PACK:{pack}]")
        else:
            lines.append(prod_norm)
    return "\n".join(lines)


def build_rerank_documents(items: list) -> list:
    docs = []
    for item in items[:100]:
        prod_norm = simple_normalize_for_llm(item["product"])
        pack = item["pack_size"].strip()
        if pack and pack not in ("", "nan", "NaN"):
            docs.append(f"{prod_norm} [PACK:{pack}]")
        else:
            docs.append(prod_norm)
    return docs


# =========================
# FUZZY MATCHING
# =========================
def find_best_match_with_fuzzy(llm_response: str, items: list, verbose=False) -> tuple:
    llm_response_clean = re.sub(r'\s*\[PACK:.*?\]', '', llm_response).strip()
    response_norm = norm(llm_response_clean)
    for item in items:
        if item["product_norm"] == response_norm:
            return item, "exact"
    scorers = [
        ("token_sort_ratio", fuzz.token_sort_ratio),
        ("token_set_ratio", fuzz.token_set_ratio),
        ("ratio", fuzz.ratio),
    ]
    best_match = None
    best_score = 0
    best_scorer = ""
    candidate_products = [item["product"] for item in items]
    for scorer_name, scorer_func in scorers:
        result = process.extractOne(llm_response_clean, candidate_products, scorer=scorer_func)
        if result:
            match_name, score, _ = result
            if score > best_score:
                best_match = match_name
                best_score = score
                best_scorer = scorer_name
    if best_match and best_score >= 75:
        for item in items:
            if item["product"] == best_match:
                return item, "fuzzy"
    if best_match and best_score >= 65:
        for item in items:
            if item["product"] == best_match:
                return item, "fuzzy_low_confidence"
    return None, "none"


# ============================================================================
# FIX 9: rank_local_candidates — ghost-token penalty on product NAME
# ----------------------------------------------------------------------------
# Extends the existing variant/sub_variant-column ghost penalty to scan the
# product NAME itself, so products like "EPITRIL BETA" or "TELSITE AMH" get
# penalised against plain-variant candidates when those words aren't in input.
# ============================================================================
def rank_local_candidates(items: list, input_name: str) -> list:
    if not items:
        return []

    ranked_items = prioritize_by_product_name_match(items, input_name, verbose=False)
    if len(ranked_items) == 1:
        return [(100.0, ranked_items[0])]

    input_norm = norm(input_name)
    input_tokens_upper = set(input_norm.upper().split())

    # Ghost-token vocabulary (pharma sub-brand / variant suffixes). Used only
    # for PENALTY, not for categorization. If an input doesn't contain the
    # ghost word, any product with that word in its name is deprioritised.
    KNOWN_GHOST_TOKENS = {
        "FORTE", "BETA", "MD", "AMH", "EZ", "LAR", "XR", "SR", "CR",
        "PLUS", "DS", "OD", "MR", "ER", "SEMI",
    }

    scored_items = []
    for item in ranked_items:
        product_norm = norm(item["product"])
        fuzzy_score = max(
            fuzz.token_sort_ratio(input_norm, product_norm),
            fuzz.token_set_ratio(input_norm, product_norm),
            fuzz.ratio(input_norm, product_norm),
        )
        priority_score = calculate_product_name_priority_score(item["product"], input_name)

        alpha_ghost_penalty = 0

        # Existing: alpha sub-variant column
        item_sv = str(item.get("sub_variant", "PLAIN")).strip().upper() or "PLAIN"
        if item_sv != "PLAIN" and re.match(r"^[A-Z]+$", item_sv) and item_sv not in input_tokens_upper:
            alpha_ghost_penalty = -200

        # Existing: multi-word variant column
        item_v = str(item.get("variant", "PLAIN")).strip().upper() or "PLAIN"
        if item_v != "PLAIN" and re.match(r"^[A-Z ]+$", item_v):
            v_words = set(item_v.split())
            if not v_words.issubset(input_tokens_upper):
                alpha_ghost_penalty -= 150

        # --- FIX 9: scan PRODUCT NAME for known ghost tokens not in input
        product_tokens = set(re.findall(r'[A-Z]+', item["product"].upper()))
        ghost_in_product = (product_tokens & KNOWN_GHOST_TOKENS) - input_tokens_upper
        if ghost_in_product:
            alpha_ghost_penalty -= 120 * len(ghost_in_product)

        score = fuzzy_score + priority_score + alpha_ghost_penalty
        scored_items.append((score, item))

    scored_items.sort(key=lambda x: x[0], reverse=True)
    return scored_items


def choose_best_local_candidate(items: list, input_name: str) -> dict:
    ranked = rank_local_candidates(items, input_name)
    return ranked[0][1] if ranked else None


def has_ambiguous_missing_strength(items: list, detected_sub_variants: set) -> bool:
    specific_detected = {sv for sv in detected_sub_variants if sv != "PLAIN"}
    if specific_detected or len(items) <= 1:
        return False
    item_sub_variants = {
        str(item.get("sub_variant", "PLAIN")).strip().upper()
        for item in items
        if str(item.get("sub_variant", "PLAIN")).strip().upper() not in ("", "PLAIN")
    }
    if len(item_sub_variants) <= 1:
        return False
    if not all(is_numeric_subvariant_like(sub_variant) for sub_variant in item_sub_variants):
        return False
    item_variants = {
        str(item.get("variant", "PLAIN")).strip().upper() for item in items
    }
    return len(item_variants) == 1


# ============================================================================
# RESULT RECORD + SUGGESTIONS — NO_CLEAR_MATCH handling
# ----------------------------------------------------------------------------
# process_product now returns a structured dict instead of a bare (name, code)
# tuple so every row carries:
#   - status:          why it matched / failed (reason code)
#   - confidence:      HIGH / MEDIUM / LOW / NONE
#   - candidate_count: how many products were handed to the model / ranker
#   - suggestions:     top-N nearest products for manual review
# When a hard match isn't found we AUTO-RECOVER to the nearest plausible
# product (marked LOW confidence) instead of giving up, so the unmatched
# count shrinks while staying auditable via the status/confidence columns.
# ============================================================================
TOP_N_SUGGESTIONS = 3

# Max character edits (Levenshtein distance) allowed between the input's brand
# token and a master brand for the no-brand auto-recovery to assign that brand.
# Beyond this the brands are treated as genuinely different → NO_CLEAR_MATCH.
# Set to 1: only single-character typos recover (AVIL↔LAVIR is 2 edits → rejected).
NO_BRAND_RECOVERY_MAX_EDITS = 1


def make_result(output, product_code="", status="MATCHED", confidence="HIGH",
                candidate_count=0, suggestions=None):
    return {
        "output": output,
        "product_code": product_code,
        "status": status,
        "confidence": confidence,
        "candidate_count": int(candidate_count or 0),
        "suggestions": suggestions or [],
    }


def _suggestions_from_ranked(ranked, top_n=TOP_N_SUGGESTIONS):
    out = []
    for score, item in ranked[:top_n]:
        out.append({
            "product": item.get("product", ""),
            "product_code": item.get("product_code", ""),
            "score": round(float(score), 1),
        })
    return out


def build_item_suggestions(items, input_name, top_n=TOP_N_SUGGESTIONS):
    if not items:
        return []
    return _suggestions_from_ranked(rank_local_candidates(items, input_name), top_n)


def format_suggestions(suggestions):
    if not suggestions:
        return ""
    parts = []
    for s in suggestions:
        code = f" ({s['product_code']})" if s.get("product_code") else ""
        parts.append(f"{s.get('product', '')}{code} [{s.get('score', '')}]")
    return " | ".join(parts)


def filter_items_by_name_strength(items: list, input_name: str, verbose=False) -> list:
    """Disambiguate SKUs that share variant/sub-variant and differ only by a
    dose carried in the PRODUCT NAME / PACK_SIZE (e.g. OROFER FCM '1K' vs
    '500MG/10ML' vs '750MG/15ML', all VARIANT=FCM, SUB_VARIANT=PLAIN). Match the
    input's explicit dose-with-unit (500MG) against the product name. Only acts
    when it narrows a multi-item set, so it never fires on already-resolved rows.
    """
    if len(items) <= 1:
        return items
    txt = collapse_spaced_thousands(normalize_ocr_numeric_noise(input_name))
    dose_pairs = re.findall(r'(\d+(?:\.\d+)?)\s*(MG|MCG|GM|IU|ML|UNIT)\b', txt)
    if not dose_pairs:
        return items
    matched = []
    for it in items:
        name = collapse_spaced_thousands(str(it.get("product", "")).upper())
        if any(re.search(r'(?<![\d.])' + re.escape(num) + r'\s*' + unit, name)
               for num, unit in dose_pairs):
            matched.append(it)
    if matched and len(matched) < len(items):
        if verbose:
            print(f"  [NAME-STRENGTH] dose {dose_pairs} → {len(items)}→{len(matched)} items")
        return matched
    return items


def filter_items_prefer_exact_brand_token(items: list, brand: str, input_name: str,
                                          verbose=False) -> list:
    """Prefer products whose name STARTS with the matched brand exactly over a
    glued sub-brand extension the input did not ask for (e.g. input 'ESLO 2.5'
    must pick 'ESLO - 2.5', not 'ESLOMET - 2.5'; both sit under BRAND_NAME=ESLO).
    Only drops glued-extension items when a clean exact-brand item also exists.
    """
    if len(items) <= 1 or not brand:
        return items
    brand_first = brand.upper().replace("-", " ").split()[0] if brand.split() else ""
    if not brand_first:
        return items
    input_tokens = set(re.findall(r'[A-Z]+', normalize_ocr_numeric_noise(input_name)))
    exact, glued = [], []
    for it in items:
        toks = re.findall(r'[A-Z]+', str(it.get("product", "")).upper())
        first = toks[0] if toks else ""
        if (first != brand_first and first.startswith(brand_first)
                and len(first) > len(brand_first) and first not in input_tokens
                and first[len(brand_first):] not in input_tokens):
            glued.append(it)          # e.g. 'ESLOMET' when input has only 'ESLO'
        else:
            exact.append(it)
    if exact and glued:
        if verbose:
            print(f"  [EXACT-BRAND] dropping glued sub-brand items "
                  f"({[i['product'] for i in glued]})")
        return exact
    return items


# =========================
# MAIN PRODUCT PROCESSOR
# =========================
def process_product(input_name: str, brand_map: dict, all_variants: set, all_sub_variants: set,
                    client, verbose=False, _lower_variant_retry=False, forced_brand: str = None,
                    auto_detect: bool = False) -> tuple:
    global REQ_COUNT, LLM_REQUEST_ATTEMPTS, SUM_PROMPT_TOKENS, SUM_COMPLETION_TOKENS, SUM_TOTAL_TOKENS

    # Preprocessing
    raw_input_name = input_name
    input_name = clean_duplicate_words(input_name)
    forced_brand_variant_tags = get_brand_variant_tags(brand_map, forced_brand) if forced_brand else None
    input_name = strip_parenthetical_noise(input_name, forced_brand_variant_tags)      # FIX 2
    input_name = strip_manufacturer_noise(input_name, brand_map)  # EMCURE-prefix noise
    input_name = desegment_ocr_input(input_name, brand_map)  # un-glue OCR strings

    if verbose:
        print("\n" + "="*80)
        print("PRODUCT MATCHING")
        print("="*80)
        print(f"Input: '{input_name}'")

    # Step 1: Brand detection
    # For rows with a forced_brand hint (column B == "0"): look up that brand directly.
    # For MAYBE_PRODUCT rows (auto_detect=True, forced_brand=None): run find_best_brand_for_input().
    # If forced_brand is provided but not in brand_map: fall back to auto_detect if enabled,
    # otherwise return NO_CLEAR_MATCH.

    brand = None
    items = []

    if forced_brand:
        forced_upper = forced_brand.strip().upper()
        # Try exact match first, then case-insensitive scan
        if forced_upper in brand_map:
            brand = forced_upper
            items = list(brand_map[brand])
            if verbose:
                print(f"\n  [forced_brand] Using hint '{brand}' — {len(items)} candidates")
        else:
            matched_key = next(
                (k for k in brand_map if k.upper() == forced_upper), None
            )
            if matched_key:
                brand = matched_key
                items = list(brand_map[brand])
                if verbose:
                    print(f"\n  [forced_brand] Using hint '{brand}' (normalised) — {len(items)} candidates")
            else:
                # Hint supplied but not in brand_map.
                # BRAND SOURCE: garbage_check only.  mapping.py's own brand
                # detection (find_best_brand_for_input / find_potential_brands /
                # group_related_brands / find_abbrev_promoted_brand) and its
                # nearest-brand recovery (suggest_nearest_brands /
                # nearest_brand_within_edits) are deliberately NOT carried into
                # this combined pipeline, so there is no second, independent
                # brand matcher. A brand hint that is not in the master is a
                # dead end -> NO_CLEAR_MATCH, never a guessed brand.
                if verbose:
                    print(f"  [forced_brand] Hint '{forced_brand}' not in brand_map → NO_CLEAR_MATCH")
                return make_result("NO_CLEAR_MATCH", "",
                                   status="NO_BRAND", confidence="NONE",
                                   candidate_count=0, suggestions=[])
    else:
        # No brand hint. Only garbage_check-confirmed product rows reach this
        # function, and those always carry a brand, so this is the defensive
        # path. auto_detect is unsupported here (see BRAND SOURCE note above).
        if verbose:
            print(f"  [forced_brand] No brand hint for '{input_name[:50]}' → NO_CLEAR_MATCH")
        return make_result("NO_CLEAR_MATCH", "",
                           status="NO_BRAND", confidence="NONE",
                           candidate_count=0, suggestions=[])

    if not brand:
        # Unreachable with the branches above (brand is set or the function has
        # already returned), kept as a strict terminator. The RECOVERED_NO_BRAND
        # edit-distance recovery that used to live here is removed together with
        # mapping.py's brand matching: no product code is ever produced from a
        # brand that garbage_check did not confirm.
        return make_result("NO_CLEAR_MATCH", "", status="NO_BRAND",
                           confidence="NONE", candidate_count=0,
                           suggestions=[])

    # Step 1.5: Compound token pre-filter
    compound_tokens = extract_brand_suffix_tokens(input_name, brand)
    compound_token_digits = get_compound_token_digits(compound_tokens)
    compound_token_letters = set()
    for token in compound_tokens:
        letters = re.sub(r'\d', '', token)
        if letters:
            compound_token_letters.add(letters)
    if compound_tokens:
        if verbose:
            print(f"\n  Compound tokens: {compound_tokens}")
        items = filter_items_by_compound_tokens(items, compound_tokens, verbose=verbose)

    # Step 1.6: Product-name abbreviation filter (FIX 4b) — fires when sub-
    # brand is embedded in product name rather than in brand_map keys.
    items = filter_items_by_product_name_abbrev(input_name, brand, items, verbose=verbose)

    # Step 1.65: prefer exact brand token over a glued sub-brand the input did
    # not ask for (input 'ESLO' must not map to 'ESLOMET').
    items = filter_items_prefer_exact_brand_token(items, brand, input_name, verbose=verbose)

    # Step 1.66: SEMI/HALF brand-prefix discriminator. When the master carries
    # BOTH 'SEMI X …' and plain 'X …' SKUs (a half-dose sub-brand), the input's
    # SEMI prefix decides which set to keep — and its ABSENCE excludes the SEMI
    # set (so 'XILIA MP2' → 'XILIA-MP 2', 'SEMI XILIA MP2' → 'SEMI XILIA MP 2').
    # If only one set exists, nothing changes (the SEMI-absent orphan rule below
    # still rejects 'SEMI AMARYL' where no SEMI SKU exists).
    def _is_semi_item(it):
        return bool(re.search(r'\b(?:SEMI|HALF)\b', str(it.get("product", "")).upper()))
    _input_has_semi = bool(re.search(r'\b(?:SEMI|HALF)\b', normalize_ocr_numeric_noise(input_name)))
    _semi_items = [it for it in items if _is_semi_item(it)]
    _plain_items = [it for it in items if not _is_semi_item(it)]
    if _semi_items and _plain_items:
        items = _semi_items if _input_has_semi else _plain_items
        if verbose:
            print(f"\n  SEMI/HALF prefix discriminator: input_has_semi={_input_has_semi} "
                  f"→ kept {len(items)} of {len(_semi_items) + len(_plain_items)} SKUs")

    brand_scope_items = list(items)
    available_variants, available_sub_variants, _ = extract_available_attributes_from_items(brand_scope_items)

    # Step 1.7: Reject inputs naming a variant/sub-brand the master lacks.
    # e.g. AMARYL 'MP', CETAPIN 'P'/'S', ACEGABA 'GRS' — collapsing these onto
    # the base/PLAIN product produced confident-but-wrong matches.
    orphan_quals = find_unmatched_qualifier_tokens(
        input_name, brand, available_variants, available_sub_variants,
        brand_items=brand_scope_items)
    if orphan_quals:
        # Lower-variant fallback: if the orphan is an EXTRA non-numeric qualifier
        # on top of a sub-brand variant the brand DOES carry (e.g. input 'X M PLUS'
        # where only 'X M' exists), drop the orphan and map to that lower variant
        # at LOW confidence. A missing numeric strength is never down-mapped.
        numeric_orphan = any(any(ch.isdigit() for ch in o) for o in orphan_quals)
        if not _lower_variant_retry and not numeric_orphan:
            stripped = input_name
            for o in orphan_quals:
                stripped = re.sub(r'\b' + re.escape(o) + r'\b', ' ', stripped)
            stripped = re.sub(r'\s+', ' ', stripped).strip()
            lower_vars = detect_variants_in_input(
                stripped, brand, available_variants, verbose=False)
            if stripped and stripped != input_name and any(v != "PLAIN" for v in lower_vars):
                res = process_product(stripped, brand_map, all_variants, all_sub_variants,
                                      client, verbose=verbose, _lower_variant_retry=True,
                                      forced_brand=forced_brand, auto_detect=auto_detect)
                if res["status"] == "MATCHED" and res["output"] != "NO_CLEAR_MATCH":
                    if verbose:
                        print(f"\nRECOVERED (lower variant): '{brand}' lacks "
                              f"{orphan_quals}; mapped to lower variant '{res['output']}'")
                    res["status"] = "RECOVERED_LOWER_VARIANT"
                    res["confidence"] = "LOW"
                return res
        suggestions = build_item_suggestions(brand_scope_items, input_name)
        if verbose:
            print(f"\nNO_CLEAR_MATCH: qualifier {orphan_quals} is not a known "
                  f"variant/sub-brand of '{brand}' in the master")
        return make_result("NO_CLEAR_MATCH", "", status="VARIANT_NOT_IN_MASTER",
                           confidence="NONE", candidate_count=len(brand_scope_items),
                           suggestions=suggestions)

    # FORM-AWARE RETRY GUARD (Steps 2/3/3.5)
    # Detected once up-front so the subtractive filters below can consult it.
    # Step 4 keeps its own detect_dosage_form_in_input call and verbose output;
    # the function is pure, so calling it twice changes nothing.
    _input_forms = detect_dosage_form_in_input(input_name, verbose=False)

    # Step 2: Variant filtering (FIX 5)
    detected_variants = detect_variants_in_input(
        input_name, brand, available_variants,
        compound_token_letters=compound_token_letters, verbose=verbose,
        brand_items=brand_scope_items)
    _before_var = items
    items = filter_items_by_variant(items, detected_variants,
                                     reference_items=brand_scope_items, verbose=verbose)
    if _input_forms and items and not _form_compatible(items, _input_forms):
        _pool = _form_compatible(_before_var, _input_forms)
        if _pool:
            # reference_items MUST be narrowed to the compatible pool: the PLAIN
            # fallback inside filter_items_by_variant otherwise re-injects the
            # conflicting-form SKU straight back out of the full brand scope.
            _retry = filter_items_by_variant(_pool, detected_variants,
                                             reference_items=_pool, verbose=verbose)
            if _retry:
                items = _retry

    # Step 3: Sub-variant filtering (FIX 6)
    detected_sub_variants = detect_sub_variants_in_input(
        input_name, brand, available_sub_variants, available_variants,
        compound_token_digits=compound_token_digits, verbose=verbose)
    _before_sv = items
    items = filter_items_by_sub_variant(items, detected_sub_variants, input_name,
                                         reference_items=brand_scope_items, verbose=verbose)
    if _input_forms and items and not _form_compatible(items, _input_forms):
        _pool = _form_compatible(_before_sv, _input_forms)
        if _pool:
            _retry = filter_items_by_sub_variant(_pool, detected_sub_variants, input_name,
                                                 reference_items=_pool, verbose=verbose)
            if _retry:
                items = _retry

    # Step 3.5: dose-in-name disambiguation — when SKUs share variant/sub-variant
    # and differ only by a strength in the product name (OROFER FCM 500MG/10ML).
    _before_ns = items
    items = filter_items_by_name_strength(items, input_name, verbose=verbose)
    if _input_forms and items and not _form_compatible(items, _input_forms):
        _pool = _form_compatible(_before_ns, _input_forms)
        if _pool:
            _retry = filter_items_by_name_strength(_pool, input_name, verbose=verbose)
            if _retry:
                items = _retry

    # Step 4: Dosage form prioritization (FIX 7)
    detected_forms = detect_dosage_form_in_input(input_name, verbose=verbose)
    items = prioritize_by_dosage_form(items, detected_forms, verbose=verbose)

    # Step 5: Pack size filtering (FIX 1 + RULE: no-pack → base)
    input_pack_size = extract_pack_size_from_input(input_name)
    items = filter_items_by_pack_size(items, input_pack_size,
                                       reference_items=brand_scope_items, verbose=verbose)
    items = apply_galact_granules_rules(
        items,
        brand,
        brand_scope_items,
        input_name,
        raw_input_name,
        detected_variants,
        detected_sub_variants,
        input_pack_size,
        verbose=verbose,
    )

    if has_ambiguous_missing_strength(items, detected_sub_variants):
        # Input strength is unresolvable against multiple numeric SKUs — do NOT
        # guess a base variant; return NO_CLEAR_MATCH (suggestions kept for review).
        suggestions = build_item_suggestions(items, input_name)
        if verbose:
            print("\nNO_CLEAR_MATCH: ambiguous strength (input does not resolve to one SKU)")
        return make_result("NO_CLEAR_MATCH", "", status="AMBIGUOUS_STRENGTH",
                           confidence="NONE", candidate_count=len(items),
                           suggestions=suggestions)

    if not items:
        suggestions = build_item_suggestions(brand_scope_items, input_name)
        recovered = choose_best_local_candidate(brand_scope_items, input_name)
        if recovered:
            if verbose:
                print(f"\nRECOVERED (over-filtered) → '{recovered['product']}'")
            return make_result(recovered["product"], recovered.get("product_code", ""),
                               status="RECOVERED_NO_CANDIDATES", confidence="LOW",
                               candidate_count=len(brand_scope_items),
                               suggestions=suggestions)
        if verbose:
            print("\n✗ No products left after filtering")
        return make_result("NO_CLEAR_MATCH", "", status="NO_CANDIDATES",
                           confidence="NONE", candidate_count=0,
                           suggestions=suggestions)

    # Local ranking (FIXES 8 + 9)
    local_ranked = rank_local_candidates(items, input_name)
    local_best_item = local_ranked[0][1] if local_ranked else None
    local_best_score = local_ranked[0][0] if local_ranked else 0
    local_second_score = local_ranked[1][0] if len(local_ranked) > 1 else None

    if verbose and local_ranked:
        print("\n" + "="*80)
        print("LOCAL FINAL CANDIDATE RANKING")
        print("="*80)
        for score, item in local_ranked[:5]:
            print(f"  {score:5.1f} -> {item['product']}")

    local_suggestions = _suggestions_from_ranked(local_ranked)
    if local_best_item:
        if len(items) == 1:
            if verbose:
                print("\n✓ Single candidate — skipping reranker")
            return make_result(local_best_item["product"],
                               local_best_item.get("product_code", ""),
                               status="MATCHED", confidence="HIGH",
                               candidate_count=len(items), suggestions=local_suggestions)
        if local_best_score >= 90 and (local_second_score is None or
                                         (local_best_score - local_second_score) >= 8):
            if verbose:
                print("\n✓ Strong local match — skipping reranker")
            return make_result(local_best_item["product"],
                               local_best_item.get("product_code", ""),
                               status="MATCHED", confidence="HIGH",
                               candidate_count=len(items), suggestions=local_suggestions)

    # Step 6: Reranker
    input_name_llm = simple_normalize_for_llm(input_name)
    brand_context = build_brand_context_string(items, verbose=verbose)
    rerank_documents = build_rerank_documents(items)
    # PROCESSING_MODE == "local": no API call is made. The row is resolved from
    # the ranking Steps 1-5 already produced, reusing the same local_best_item and
    # local_suggestions the reranker error path falls back to. Candidate
    # generation, filtering and ranking are untouched -- only the reranker call is
    # skipped. LLM_REQUEST_ATTEMPTS is deliberately not incremented, which also
    # keeps the DELAY_BETWEEN_LLM_REQUESTS throttle in main() from firing.
    if not llm_enabled():
        candidate_count = len(rerank_documents)
        if local_best_item:
            if verbose:
                print(f"\nLOCAL ONLY (reranker disabled) → '{local_best_item['product']}'")
            return make_result(local_best_item["product"],
                               local_best_item.get("product_code", ""),
                               status="RECOVERED_LOCAL_ONLY", confidence="LOW",
                               candidate_count=candidate_count,
                               suggestions=local_suggestions)
        return make_result("NO_CLEAR_MATCH", "", status="LOCAL_ONLY_UNRESOLVED",
                           confidence="NONE", candidate_count=candidate_count,
                           suggestions=local_suggestions)

    # ======================================================================
    # Step 6 (continued): reranker / LLM path -- PRESENT BUT DISABLED.
    #
    # PROCESSING_MODE is pinned to "local" in this file, so llm_enabled() is
    # always False and the branch above always returns. The original reranker
    # code is preserved verbatim below, commented out, so it can be re-enabled
    # without re-deriving it. Re-enabling requires all of:
    #   * PROCESSING_MODE = "llm" (or "both") in the CONFIG block
    #   * the reranker credentials/endpoint config (HUGGINGFACE_API_KEY,
    #     HUGGINGFACE_API_URL, MAX_OUTPUT_TOKENS, MODEL_NAME, REQUEST_TIMEOUT)
    #     -- deliberately NOT copied
    #     into this file
    #   * the commented-out call_groq_llm() definition above
    #   * a non-empty MAPPING_LLM_PROMPT (currently blank by design)
    # Until then, every row that would need the reranker to decide is reported
    # as NO_SUGGESTION / 0 by decide_output(); an unreranked local candidate is
    # never promoted to a product code.
    # ======================================================================
    # final_prompt = USER_PROMPT_TEMPLATE.format(
        # input_name=input_name_llm,
        # brand_context=brand_context,
    # )
    # candidate_count = len(rerank_documents)
    # LLM_REQUEST_ATTEMPTS += 1
    # try:
        # answer, usage = call_groq_llm(client, SYSTEM_PROMPT, final_prompt,
                                              # rerank_documents, verbose=verbose)
    # except Exception as e:
        # print(f"  API Error: {e}")
        # if local_best_item:
            # return make_result(local_best_item["product"],
                               # local_best_item.get("product_code", ""),
                               # status="RECOVERED_API_ERROR", confidence="LOW",
                               # candidate_count=candidate_count,
                               # suggestions=local_suggestions)
        # return make_result(f"API_ERROR: {str(e)[:50]}", "", status="API_ERROR",
                           # confidence="NONE", candidate_count=candidate_count,
                           # suggestions=local_suggestions)
    #
    # REQ_COUNT += 1
    # if usage:
        # pt = (usage.get("prompt_tokens") if isinstance(usage, dict)
              # else getattr(usage, "prompt_tokens", 0)) or 0
        # ct = (usage.get("completion_tokens") if isinstance(usage, dict)
              # else getattr(usage, "completion_tokens", 0)) or 0
        # tt = (usage.get("total_tokens") if isinstance(usage, dict)
              # else getattr(usage, "total_tokens", 0)) or (pt + ct)
        # SUM_PROMPT_TOKENS += pt
        # SUM_COMPLETION_TOKENS += ct
        # SUM_TOTAL_TOKENS += tt
    #
    # answer = answer.strip()
    # if "NO_CLEAR_MATCH" in answer.upper():
        # if local_best_item:
            # if verbose:
                # print(f"\nRECOVERED (LLM said NO_CLEAR_MATCH) → '{local_best_item['product']}'")
            # return make_result(local_best_item["product"],
                               # local_best_item.get("product_code", ""),
                               # status="RECOVERED_LLM_REJECTED", confidence="LOW",
                               # candidate_count=candidate_count,
                               # suggestions=local_suggestions)
        # return make_result("NO_CLEAR_MATCH", "", status="LLM_REJECTED",
                           # confidence="NONE", candidate_count=candidate_count,
                           # suggestions=local_suggestions)
    #
    # matched_item, match_type = find_best_match_with_fuzzy(answer, items, verbose=verbose)
    # if not matched_item:
        # if local_best_item:
            # if verbose:
                # print(f"\nRECOVERED (reranker answer unmatched) → '{local_best_item['product']}'")
            # return make_result(local_best_item["product"],
                               # local_best_item.get("product_code", ""),
                               # status="RECOVERED_LLM_UNMATCHED", confidence="LOW",
                               # candidate_count=candidate_count,
                               # suggestions=local_suggestions)
        # return make_result("NO_CLEAR_MATCH", "", status="LLM_UNMATCHED",
                           # confidence="NONE", candidate_count=candidate_count,
                           # suggestions=local_suggestions)
    #
    # matched_product_name = matched_item["product"]
    # product_code = matched_item.get("product_code", "")
    # matched_pack_size = matched_item.get("pack_size", "")
    #
    # if verbose:
        # print("\n" + "="*80)
        # print(f"FINAL: {matched_product_name}  (via {match_type})")
        # print("="*80)
    # match_confidence = "HIGH" if match_type == "exact" else "MEDIUM"
    # return make_result(matched_product_name, product_code,
                       # status="MATCHED", confidence=match_confidence,
                       # candidate_count=candidate_count, suggestions=local_suggestions)
    # Strict terminator. Not reachable while PROCESSING_MODE == "local"; it
    # exists so process_product() can never fall off the end returning None.
    return make_result("NO_CLEAR_MATCH", "", status="LOCAL_ONLY_UNRESOLVED",
                       confidence="NONE", candidate_count=len(rerank_documents),
                       suggestions=local_suggestions)


# ==========================================================================
# STAGE 1 DRIVER: garbage-check classification for a single row
#
# Refactored from garbage_check.main()'s per-row body. Fully local: no LLM call
# is made or possible in this stage. The exception cascade below is reproduced
# in the original order; the only change is that results are returned instead of
# written into columns B-G of the input workbook.
#
# Returns (kind, brand) where kind is one of:
#   "PRODUCT"        -> confirmed product row, brand hint is the second element
#   "GARBAGE"        -> garbage row
#   "MAYBE_PRODUCT"  -> form+qty pattern only, no confirmed brand
#   "REVIEW"         -> unresolved; garbage_check defers these to review
# ==========================================================================
def classify_row(name, brand_index, garbage_rows, verbose=False):
    spans = token_spans(name)
    exact_match = exact_product_match(name, brand_index, precomputed_spans=spans)
    fuzzy_match = None if exact_match else fuzzy_product_match(name, brand_index,
                                                               precomputed_spans=spans)
    local_match = exact_match or fuzzy_match

    matched_brand_norm = normalize_text(local_match["brand"]) if local_match else ""

    # Exception: CELOL matching rows with DINNER or SET are non-pharma garbage
    if local_match and matched_brand_norm == "CELOL" and re.search(
            r"\b(MARKER|DINNER|SETS?)\b", str(name).upper()):
        return "GARBAGE", ""

    # EXACT_ONLY_BRANDS: fuzzy matches are too risky (e.g. CTAX vs TAX,
    # EFCURE vs EMCURE). Only exact span matches are trusted; fuzzy hits are
    # invalidated so the row falls through to garbage / review checks.
    if local_match and matched_brand_norm in EXACT_ONLY_BRANDS:
        if local_match.get("match_type") != "exact":
            local_match = None
            matched_brand_norm = ""

    # Exception: TAMLET matching the common dosage token TABLET/TABLETS
    if local_match and matched_brand_norm == "TAMLET":
        if normalize_text(local_match.get("matched_text", "")) in ("TABLET", "TABLETS"):
            local_match = None
            matched_brand_norm = ""

    # Exception: IMPETUS / VINTOR / EMNU rows containing "COMPANY" are
    # company/distributor header lines, not product rows.
    _COMPANY_BRANDS = {"IMPETUS", "VINTOR", "EMNU"}
    if local_match and matched_brand_norm in _COMPANY_BRANDS:
        if re.search(r"\bCOMPANY\b", str(name), re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # Exception: AMARYL + SEMI is not a valid product (SEMI AMARYL is not in
    # the master list).
    if local_match and matched_brand_norm == "AMARYL":
        if re.search(r"\bSEMI\b", str(name), re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # Exception: EMNU + MANUFACTURER is a manufacturer/company header line.
    if local_match and matched_brand_norm == "EMNU":
        if re.search(r"\bMANUFACTURER\b", str(name), re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # Exception: ZUVENTUS has only one real product (ORS ORANGE); every other
    # ZUVENTUS row is a company/division line -> garbage directly.
    if local_match and matched_brand_norm == "ZUVENTUS":
        if not re.search(r"\bORS\b", str(name), re.IGNORECASE):
            return "GARBAGE", ""

    # Exception: NEW brand -- only rows containing NORMET are real products.
    if local_match and matched_brand_norm == "NEW":
        if not re.search(r"\bNORMET\b", str(name), re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # Exception: VITAMIN brand -- only rows containing D3 are real products.
    if local_match and matched_brand_norm == "VITAMIN":
        if not re.search(r"\bD3\b", str(name), re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # Exception: EMCOR brand -- real product rows must contain CREAM or TUBE.
    if local_match and matched_brand_norm == "EMCOR":
        if not re.search(r"\b(CREAM|TUBE)\b", str(name), re.IGNORECASE):
            local_match = None
            matched_brand_norm = ""

    # Exception: NUMLO -- when the raw text is SNUMLO the real product is
    # S-NUMLO. A plain NUMLO match (no S prefix) is left untouched.
    if local_match and matched_brand_norm == "NUMLO":
        if ("S" + matched_brand_norm) in normalize_text(name):
            local_match = dict(local_match)
            local_match["brand"] = "S-NUMLO"

    if local_match and matched_brand_norm not in LOCAL_REVIEW_BRANDS:
        return "PRODUCT", local_match["brand"]
    elif (
        is_exact_garbage_row(name, garbage_rows)
        or is_pattern_garbage_row(name)
        or is_metadata_garbage_row(name)
    ):
        return "GARBAGE", ""
    else:
        if fuzzy_garbage_match(name, garbage_rows):
            return "GARBAGE", ""
        elif local_match:
            # Brand requires review -> garbage_check defers this row.
            return "REVIEW", local_match["brand"]
        else:
            if is_maybe_product_row(name):
                return "MAYBE_PRODUCT", ""
            return "REVIEW", ""


# ==========================================================================
# STAGE 2 DRIVER + OUTPUT DECISION
# ==========================================================================

# Statuses that process_product() only ever returns because the reranker was
# unavailable or its answer could not be used. Every one of them carries a
# LOCALLY chosen candidate (or none at all) that the reranker was supposed to
# confirm or reject. With the reranker disabled, such a row is NOT resolved, so
# it is reported as NO_SUGGESTION / 0 rather than mapped to the local guess.
#
#   RECOVERED_LOCAL_ONLY    reranker skipped (PROCESSING_MODE == "local")
#   LOCAL_ONLY_UNRESOLVED   reranker skipped, no local candidate either
#   RECOVERED_API_ERROR     reranker call failed, fell back to local candidate
#   API_ERROR               reranker call failed, no local candidate
#   RECOVERED_LLM_REJECTED  reranker said NO_CLEAR_MATCH, fell back to local
#   LLM_REJECTED            reranker said NO_CLEAR_MATCH, no local candidate
#   RECOVERED_LLM_UNMATCHED reranker answer matched nothing, fell back to local
#   LLM_UNMATCHED           reranker answer matched nothing, no local candidate
RERANKER_DEPENDENT_STATUSES = {
    "RECOVERED_LOCAL_ONLY",
    "LOCAL_ONLY_UNRESOLVED",
    "RECOVERED_API_ERROR",
    "API_ERROR",
    "RECOVERED_LLM_REJECTED",
    "LLM_REJECTED",
    "RECOVERED_LLM_UNMATCHED",
    "LLM_UNMATCHED",
}


def decide_output(res):
    """Translate a process_product() result into (remark, product_code).

    Accuracy rule: a product code is emitted only when the LOCAL pipeline
    resolved the row on its own. NO_CLEAR_MATCH, a missing code, and every
    reranker-dependent status all collapse to NO_SUGGESTION / 0.
    """
    status = res.get("status", "")
    output = res.get("output", "")
    code = str(res.get("product_code", "") or "").strip()

    if status in RERANKER_DEPENDENT_STATUSES:
        return REMARK_NO_SUGGESTION, CODE_NO_SUGGESTION
    if not code or output == "NO_CLEAR_MATCH" or output.startswith("API_ERROR:"):
        return REMARK_NO_SUGGESTION, CODE_NO_SUGGESTION
    return output, (int(code) if code.isdigit() else code)


def map_one(input_name, brand, brand_map, all_variants, all_sub_variants):
    """Run the mapping pipeline for one garbage_check-confirmed product row.

    auto_detect is always False: the brand comes from garbage_check only.
    client is always None: the reranker is disabled.
    """
    res = process_product(input_name, brand_map, all_variants, all_sub_variants,
                          None, verbose=False, forced_brand=brand,
                          auto_detect=False)
    return res


# ==========================================================================
# PIPELINE DRIVER
# ==========================================================================
def build_output_workbook(rows):
    """Write the single-sheet, three-column output workbook.

    The input workbook is never opened for writing; a new file is created at
    OUTPUT_XLSX_PATH with exactly one sheet and exactly the columns
    Input | Remark | Product Code.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = OUTPUT_SHEET_NAME
    ws.append(list(OUTPUT_COLUMNS))
    for input_value, remark, code in rows:
        ws.append([input_value, remark, code])
    wb.save(OUTPUT_XLSX_PATH)
    return len(rows)


def main():
    if llm_enabled():
        raise RuntimeError(
            "PROCESSING_MODE must stay 'local' in MAPPING_MAIN.py: the reranker "
            "is commented out and no credentials are configured here."
        )

    # ---- stage 1 resources (garbage_check) ----
    brands = load_brands()
    brand_index = build_brand_index(brands)
    garbage_rows = build_garbage_set()
    print(f"{len(brands)} brands loaded (garbage-check stage)")

    # ---- stage 2 resources (mapping) ----
    df_master = load_master(MASTER_XLSX_PATH)
    brand_map = build_brand_product_map(df_master)
    all_variants = extract_all_variants_from_data(brand_map)
    all_sub_variants = extract_all_sub_variants_from_data(brand_map)
    print(f"{len(brand_map)} master brands, {len(all_variants)} variants, "
          f"{len(all_sub_variants)} sub-variants (mapping stage)")

    # ---- read the input column (read-only; the input file is never written) ----
    print(f"Reading input: {INPUT_XLSX_PATH} [{INPUT_SHEET_NAME}]")
    wb_in = load_workbook(INPUT_XLSX_PATH, read_only=True, data_only=True)
    ws_in = wb_in[INPUT_SHEET_NAME]
    raw = []
    for row in ws_in.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        value = row[0]
        raw.append("" if value is None else str(value).strip())
    wb_in.close()
    if ROW_LIMIT is not None:
        raw = raw[:ROW_LIMIT]
    print(f"{len(raw)} input rows read. Starting combined pipeline...")

    counts = {"blank": 0, "garbage": 0, "review": 0, "maybe_product": 0,
              "product": 0, "mapped": 0, "no_suggestion": 0}
    out_rows = []

    for i, name in enumerate(raw, start=2):
        # Blank rows are not classified, mirroring garbage_check, which skips
        # them entirely. An empty output row keeps output row N aligned with
        # input row N instead of inventing a classification.
        if name == "":
            counts["blank"] += 1
            out_rows.append(("", "", ""))
            continue

        kind, brand = classify_row(name, brand_index, garbage_rows)

        if kind == "GARBAGE":
            counts["garbage"] += 1
            out_rows.append((name, REMARK_GARBAGE, CODE_GARBAGE))
            continue
        if kind == "REVIEW":
            counts["review"] += 1
            out_rows.append((name, REMARK_REVIEW, CODE_REVIEW))
            continue
        if kind == "MAYBE_PRODUCT":
            # Not a confirmed product: the mapping stage is not entered at all.
            counts["maybe_product"] += 1
            out_rows.append((name, REMARK_NO_SUGGESTION, CODE_NO_SUGGESTION))
            continue

        # Confirmed product row: starts at NO_SUGGESTION / 0 and is overwritten
        # only if the local mapping pipeline resolves it on its own.
        counts["product"] += 1
        remark, code = REMARK_NO_SUGGESTION, CODE_NO_SUGGESTION
        try:
            res = map_one(name, brand, brand_map, all_variants, all_sub_variants)
            remark, code = decide_output(res)
        except Exception as e:
            # Never let one bad row abort the run or fabricate a code.
            print(f"row {i}: {name!r} -> mapping error, kept NO_SUGGESTION ({e})")
            remark, code = REMARK_NO_SUGGESTION, CODE_NO_SUGGESTION
        if code == CODE_NO_SUGGESTION and remark == REMARK_NO_SUGGESTION:
            counts["no_suggestion"] += 1
        else:
            counts["mapped"] += 1
        out_rows.append((name, remark, code))

        if counts["product"] % 500 == 0:
            print(f"  ... {i - 1} rows processed "
                  f"({counts['mapped']} mapped, {counts['no_suggestion']} no-suggestion)")

    written = build_output_workbook(out_rows)
    print(f"\nWrote {written} rows to {OUTPUT_XLSX_PATH} "
          f"[sheet: {OUTPUT_SHEET_NAME}]")
    print(f"  blank            : {counts['blank']}")
    print(f"  garbage (1)      : {counts['garbage']}")
    print(f"  review (2)       : {counts['review']}")
    print(f"  maybe product (0): {counts['maybe_product']}")
    print(f"  confirmed product: {counts['product']}")
    print(f"    mapped         : {counts['mapped']}")
    print(f"    no suggestion  : {counts['no_suggestion']}")
    print(f"  LLM calls made   : 0 (PROCESSING_MODE='local')")


if __name__ == "__main__":
    main()
