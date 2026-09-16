import difflib
import re
import time

import requests
from openpyxl import load_workbook

# ---------------- CONFIG ----------------
API_KEY = "sk_dpI1XRuvqvfeLamnRTwKCnkikDY1dDe4aF2hTfv0aT8"
NOVITA_URL = "https://api.novita.ai/openai/v1/chat/completions"
MODEL = "nvidia/nemotron-3-nano-30b-a3b"

EXCEL_FILE = "test2.xlsx"
BRANDS_FILE = "Brand Names.txt"

BATCH_SIZE = 2    # rows per API call
BATCH_WAIT = 2       # seconds between API calls, keeps rate limits away
ROW_LIMIT = 19000000     # only process this many rows (None = all rows)
MAX_RETRIES = 5
RETRY_WAIT = 10      # base seconds between retries (doubles each retry)
# Classification mode for a run. This is the single control point for LLM usage;
# nothing else in this file decides whether the API is called.
#   "local" -> local classification logic only, no LLM/API calls at all
#   "llm"   -> local pre-pass plus the LLM fallback that resolves the rows the
#              local pass could not confirm
#   "both"  -> local logic and LLM processing per the current pipeline
# The pipeline is local-first by construction: the local pass is what produces
# the uncertain-row set the LLM consumes, so no row can reach the LLM without
# it. "llm" and "both" therefore select the same route; both spellings are
# accepted so the setting reads the way you expect.
PROCESSING_MODE = "local"

VALID_PROCESSING_MODES = ("llm", "local", "both")
if PROCESSING_MODE not in VALID_PROCESSING_MODES:
    raise ValueError(
        f"PROCESSING_MODE must be one of {VALID_PROCESSING_MODES}, "
        f"got {PROCESSING_MODE!r}"
    )

# Wipe column B and redo everything. Set to False so a crashed run
# can resume where it stopped instead of restarting.
CLEAR_EXISTING = True

# Novita serverless pricing for this model, USD per million tokens
PRICE_INPUT_PER_M = 0.02
PRICE_OUTPUT_PER_M = 0.05

TOP_CANDIDATES = 3
MIN_SHOW_SCORE = 0.60  # below this, report "no close brand"
# ----------------------------------------

TOKENS = {"input": 0, "output": 0}

PROMPT_TEMPLATE = """You are a data-cleaning classifier for a pharmaceutical stock-and-sales dataset. The rows below were extracted by OCR from PDFs and scanned images, so real product names may contain spelling errors.

Each row has already been pre-checked locally. These are only the uncertain rows that could not be confirmed as a real product or as an exact garbage row by strict local rules.

BRAND LIST RULE:
- Every real product row contains exactly one of the known brand names, possibly with small OCR errors.
- Match against the brand list only. Do not use outside knowledge of medicine names.
- The brand may appear anywhere in the row, not only at the beginning.

DECISION RULE:
- Answer 0 if the row is a real product row.
- Answer 1 if the row is garbage.
- Do not assume medicine-like tokens such as MG, TAB, CAP, INJ prove it is a product.
- The word TOTAL does not automatically make a row garbage if a real brand is present.

OUTPUT RULE:
- Reply with exactly one line per row.
- Each line must be exactly: <row number>=<0 or 1>
- No extra words. No explanation. No punctuation other than the equals sign.

EXAMPLES:
1. OROFER TOTAL || candidates: OROFER (word "OROFER", 1.00)
2. TOTAL SALE || candidates: PROTOTAL (word "TOTAL", 0.77)
3. CARDIAC DIVISION || candidates: no close brand
4. WY LDA 50MG || candidates: VYLDA (word "WYLDA", 0.80)

NOW CLASSIFY THESE ROWS:
{rows}

ANSWERS:
1=0
2=1
3=1
4=0
"""

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


def brand_candidates(name, brands):
    """Top fuzzy-matched brands for a row: list of (brand, matched_word, score)."""
    tokens = re.findall(r"[A-Za-z]{3,}", name.upper())
    words = tokens + [tokens[i] + tokens[i + 1] for i in range(len(tokens) - 1)]
    if not words:
        return []
    best = []
    for brand in brands:
        b = brand.upper()
        score, word = max(
            (difflib.SequenceMatcher(None, w, b).ratio(), w) for w in words
        )
        best.append((brand, word, score))
    best.sort(key=lambda x: -x[2])
    return [c for c in best[:TOP_CANDIDATES] if c[2] >= MIN_SHOW_SCORE]


def format_row(i, name, cands):
    if cands:
        cand_text = ", ".join(f'{b} (word "{w}", {s:.2f})' for b, w, s in cands)
    else:
        cand_text = "no close brand"
    return f"{i}. {name} || candidates: {cand_text}"


def call_model(prompt, max_tokens):
    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0,
        "max_tokens": max_tokens,
    }
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }

    wait = RETRY_WAIT
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(NOVITA_URL, json=payload, headers=headers, timeout=120)
            if resp.status_code == 429:
                retry_after = resp.headers.get("Retry-After")
                pause = int(retry_after) if retry_after and retry_after.isdigit() else wait
                print(f"    rate limited (attempt {attempt}), waiting {pause}s")
                time.sleep(pause)
                wait *= 2
                continue
            resp.raise_for_status()
            data = resp.json()
            usage = data.get("usage", {})
            TOKENS["input"] += usage.get("prompt_tokens", 0)
            TOKENS["output"] += usage.get("completion_tokens", 0)
            return data["choices"][0]["message"]["content"]
        except Exception as e:
            print(f"    attempt {attempt} failed: {e}")
            if attempt < MAX_RETRIES:
                time.sleep(wait)
                wait *= 2
    return None


def classify_batch(names, brands):
    """Returns a list of results ('0'/'1'/'BAD_OUTPUT'/'ERROR'), one per name."""
    rows_text = "\n".join(
        format_row(i, name, brand_candidates(name, brands))
        for i, name in enumerate(names, start=1)
    )
    prompt = PROMPT_TEMPLATE.format(rows=rows_text, n=len(names))

    content = call_model(prompt, max_tokens=60 * len(names) + 100)
    if content is None:
        return ["ERROR"] * len(names)

    answers = {}
    for line in content.splitlines():
        m = re.match(r"\s*(\d+)\s*=\s*([01])\s*$", line)
        if m:
            answers[int(m.group(1))] = m.group(2)

    # single-row fallback: some responses still return only the digit
    if len(names) == 1 and 1 not in answers:
        m = re.search(r"\b([01])\b", content.strip())
        if m:
            answers[1] = m.group(1)

    results = [answers.get(i, "BAD_OUTPUT") for i in range(1, len(names) + 1)]
    if "BAD_OUTPUT" in results:
        print(f"    unparsed model output:\n{content}")
    return results


def has_result(value):
    """Treat None/blank strings as not-yet-classified."""
    return value is not None and str(value).strip() != ""


def write_local_match(ws, row, match):
    ws.cell(row=row, column=3, value=match["brand"])
    ws.cell(row=row, column=4, value=match["matched_text"])
    ws.cell(row=row, column=5, value=match["match_type"])


def clear_match_details(ws, row):
    for column in (3, 4, 5, 6, 7):
        ws.cell(row=row, column=column).value = None


def write_garbage_match(ws, row, match):
    ws.cell(row=row, column=6, value=match["garbage_phrase"])
    ws.cell(row=row, column=7, value=match["match_type"])


def clear_garbage_match(ws, row):
    for column in (6, 7):
        ws.cell(row=row, column=column).value = None


def llm_enabled():
    """True when PROCESSING_MODE allows LLM/API calls.

    PROCESSING_MODE is read at call time so a test harness can set the mode on
    the imported module and have it take effect.
    """
    return PROCESSING_MODE in ("llm", "both")


def main():
    if llm_enabled() and not API_KEY:
        raise RuntimeError(
            "API key not found. Paste your Novita API key into API_KEY at the top of this file."
        )

    brands = load_brands()
    brand_index = build_brand_index(brands)
    garbage_rows = build_garbage_set()
    print(f"{len(brands)} brands loaded")

    print("Loading workbook...")
    wb = load_workbook(EXCEL_FILE)
    ws = wb["garbage_check"]
    print("Workbook loaded.")
    ws["B1"] = "Garbage check"
    ws["C1"] = "Matched brand"
    ws["D1"] = "Matched text"
    ws["E1"] = "Match type"
    ws["F1"] = "Matched garbage phrase"
    ws["G1"] = "Garbage match type"

    if CLEAR_EXISTING:
        print(f"Clearing {ws.max_row - 1} rows...")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                 min_col=2, max_col=ws.max_column):
            for cell in row:
                cell.value = None
        print("Cleared previous results outside column A.")

    pending = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=1).value
        if name is None or str(name).strip() == "":
            continue
        if has_result(ws.cell(row=row, column=2).value):
            continue  # already done, allows resume when CLEAR_EXISTING is False
        pending.append((row, str(name).strip()))

    if ROW_LIMIT is not None:
        pending = pending[:ROW_LIMIT]
    print(f"{len(pending)} rows to process. Starting classification...")

    local_product = 0
    local_garbage = 0
    llm_pending = []
    for row, name in pending:
        spans = token_spans(name)
        exact_match = exact_product_match(name, brand_index, precomputed_spans=spans)
        fuzzy_match = None if exact_match else fuzzy_product_match(name, brand_index, precomputed_spans=spans)
        local_match = exact_match or fuzzy_match

        matched_brand_norm = normalize_text(local_match["brand"]) if local_match else ""

        # Exception: CELOL matching rows with DINNER or SET are non-pharma garbage
        if local_match and matched_brand_norm == "CELOL" and re.search(r"\b(MARKER|DINNER|SETS?)\b", str(name).upper()):
            ws.cell(row=row, column=2, value="1")
            clear_match_details(ws, row)
            clear_garbage_match(ws, row)
            local_garbage += 1
            print(f"row {row}: {name!r} -> 1 (brand exception: CELOL + DINNER/SET)")
            continue

        # EXACT_ONLY_BRANDS: fuzzy matches are too risky (e.g. CTAX vs TAX, EFCURE vs EMCURE).
        # Only exact span matches are trusted; fuzzy hits are invalidated so the row
        # falls through to garbage / review checks.
        if local_match and matched_brand_norm in EXACT_ONLY_BRANDS:
            if local_match.get("match_type") != "exact":
                print(
                    f"row {row}: {name!r} -> invalidated {local_match['brand']} "
                    f"fuzzy match ({local_match.get('match_type')}) — exact only"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: TAMLET matching the common dosage token TABLET/TABLETS is a false positive
        if local_match and matched_brand_norm == "TAMLET":
            if normalize_text(local_match.get("matched_text", "")) in ("TABLET", "TABLETS"):
                print(
                    f"row {row}: {name!r} -> invalidated TAMLET match on dosage word "
                    f"{local_match['matched_text']!r}"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: IMPETUS / VINTOR / EMNU rows containing "COMPANY" are
        # company/distributor header lines, not product rows — invalidate the brand match.
        _COMPANY_BRANDS = {"IMPETUS", "VINTOR", "EMNU"}
        if local_match and matched_brand_norm in _COMPANY_BRANDS:
            if re.search(r"\bCOMPANY\b", str(name), re.IGNORECASE):
                print(
                    f"row {row}: {name!r} -> invalidated {local_match['brand']} match "
                    f"(row contains COMPANY)"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: AMARYL + SEMI is not a valid product (SEMI AMARYL does not exist
        # in the master list) — invalidate so the row falls through to garbage/review.
        if local_match and matched_brand_norm == "AMARYL":
            if re.search(r"\bSEMI\b", str(name), re.IGNORECASE):
                print(
                    f"row {row}: {name!r} -> invalidated AMARYL match "
                    f"(row contains SEMI — not in master)"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: EMNU + MANUFACTURER is a manufacturer/company header line, not a product.
        if local_match and matched_brand_norm == "EMNU":
            if re.search(r"\bMANUFACTURER\b", str(name), re.IGNORECASE):
                print(
                    f"row {row}: {name!r} -> invalidated EMNU match "
                    f"(row contains MANUFACTURER)"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: ZUVENTUS has only one real product (ORS ORANGE).
        # All other ZUVENTUS rows (Division headers, Lifestyle, Athena, Gromaxx,
        # Total Value lines, etc.) are company/division lines — classify as
        # garbage directly without going through the full garbage-check funnel.
        if local_match and matched_brand_norm == "ZUVENTUS":
            if not re.search(r"\bORS\b", str(name), re.IGNORECASE):
                ws.cell(row=row, column=2, value="1")
                clear_match_details(ws, row)
                clear_garbage_match(ws, row)
                local_garbage += 1
                print(f"row {row}: {name!r} -> 1 (brand exception: ZUVENTUS without ORS)")
                continue

        # Exception: NEW brand — only rows containing the exact word NORMIT are
        # real product rows. Any other NEW match (e.g. "NEW BATCH", "NEW ARRIVAL",
        # structural headers) is invalidated and falls through to garbage/review.
        if local_match and matched_brand_norm == "NEW":
            if not re.search(r"\bNORMET\b", str(name), re.IGNORECASE):
                print(
                    f"row {row}: {name!r} -> invalidated NEW match "
                    f"(no NORMET signal — not a real product row)"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: VITAMIN brand — only rows containing D3 are real product rows.
        # Generic "VITAMIN" rows without D3 are not valid SKUs in the master;
        # invalidate and let the row fall through to garbage/review.
        if local_match and matched_brand_norm == "VITAMIN":
            if not re.search(r"\bD3\b", str(name), re.IGNORECASE):
                print(
                    f"row {row}: {name!r} -> invalidated VITAMIN match "
                    f"(no D3 signal — not a real product row)"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: EMCOR brand — real product rows must contain CREAM or TUBE.
        # Any other EMCOR match is invalidated and falls through to garbage/review.
        if local_match and matched_brand_norm == "EMCOR":
            if not re.search(r"\b(CREAM|TUBE)\b", str(name), re.IGNORECASE):
                print(
                    f"row {row}: {name!r} -> invalidated EMCOR match "
                    f"(no CREAM/TUBE signal — not a real product row)"
                )
                local_match = None
                matched_brand_norm = ""

        # Exception: NUMLO — when the matched span starts with 'S' (i.e. the raw
        # text is SNUMLO), the real product is S-NUMLO, not NUMLO.
        # Remap matched_brand_norm so downstream classification uses S-NUMLO.
        # A plain NUMLO match (no S prefix) is left untouched.
        if local_match and matched_brand_norm == "NUMLO":
            if ("S" + matched_brand_norm) in normalize_text(name):
                local_match["brand"] = "S-NUMLO"

        if local_match and matched_brand_norm not in LOCAL_REVIEW_BRANDS:
            ws.cell(row=row, column=2, value="0")
            write_local_match(ws, row, local_match)
            local_product += 1
            print(
                f"row {row}: {name!r} -> 0 "
                f"(local brand match: {local_match['brand']} via "
                f"{local_match['matched_text']!r}, {local_match['match_type']})"
            )
        elif (
            is_exact_garbage_row(name, garbage_rows)
            or is_pattern_garbage_row(name)
            or is_metadata_garbage_row(name)
        ):
            ws.cell(row=row, column=2, value="1")
            clear_match_details(ws, row)
            clear_garbage_match(ws, row)
            local_garbage += 1
            if is_exact_garbage_row(name, garbage_rows):
                rule_type = "exact garbage rule"
            elif is_pattern_garbage_row(name):
                rule_type = "safe garbage pattern"
            else:
                rule_type = "metadata garbage rule"
            print(f"row {row}: {name!r} -> 1 ({rule_type})")
        else:
            garbage_match = fuzzy_garbage_match(name, garbage_rows)
            if garbage_match:
                ws.cell(row=row, column=2, value="1")
                if local_match:
                    write_local_match(ws, row, local_match)
                else:
                    clear_match_details(ws, row)
                write_garbage_match(ws, row, garbage_match)
                local_garbage += 1
                print(
                    f"row {row}: {name!r} -> 1 "
                    f"(fuzzy garbage: {garbage_match['garbage_phrase']!r}, "
                    f"{garbage_match['match_type']})"
                )
            elif local_match:
                write_local_match(ws, row, local_match)
                clear_garbage_match(ws, row)
                print(
                    f"row {row}: {name!r} -> REVIEW "
                    f"(local match {local_match['brand']} via "
                    f"{local_match['matched_text']!r}, {local_match['match_type']}; "
                    f"brand requires review)"
                )
                llm_pending.append((row, name))
            else:
                if is_maybe_product_row(name):
                    ws.cell(row=row, column=3, value="pattern-inferred")
                    ws.cell(row=row, column=4, value="form+qty match")
                    ws.cell(row=row, column=5, value="MAYBE_PRODUCT")
                    clear_garbage_match(ws, row)
                    print(f"row {row}: {name!r} -> MAYBE_PRODUCT (form+qty pattern)")
                else:
                    clear_match_details(ws, row)
                    clear_garbage_match(ws, row)
                llm_pending.append((row, name))

    wb.save(EXCEL_FILE)
    print(
        f"{len(pending)} rows checked: "
        f"{local_product} local product, {local_garbage} local garbage, "
        f"{len(llm_pending)} unresolved"
    )

    if llm_enabled():
        for start in range(0, len(llm_pending), BATCH_SIZE):
            batch = llm_pending[start:start + BATCH_SIZE]
            results = classify_batch([name for _, name in batch], brands)

            for (row, name), result in zip(batch, results):
                ws.cell(row=row, column=2, value=result)
                print(f"row {row}: {name!r} -> {result}")

            wb.save(EXCEL_FILE)
            print(f"  saved ({start + len(batch)}/{len(llm_pending)})")
            time.sleep(BATCH_WAIT)
    else:
        for row, name in llm_pending:
            current = ws.cell(row=row, column=2).value
            if current is None or str(current).strip() == "":
                if ws.cell(row=row, column=5).value == "MAYBE_PRODUCT":
                    ws.cell(row=row, column=2, value="MAYBE_PRODUCT")
                    print(f"row {row}: {name!r} -> MAYBE_PRODUCT (form+qty pattern)")
                else:
                    ws.cell(row=row, column=2, value="REVIEW")
                    print(f"row {row}: {name!r} -> REVIEW (local unresolved)")
        wb.save(EXCEL_FILE)

    print("Done.")

    cost_in = TOKENS["input"] / 1_000_000 * PRICE_INPUT_PER_M
    cost_out = TOKENS["output"] / 1_000_000 * PRICE_OUTPUT_PER_M
    total_cost = cost_in + cost_out
    print("\n---------- COST ----------")
    print(f"rows classified : {len(pending)}")
    print(f"input tokens    : {TOKENS['input']:,}  (${cost_in:.6f})")
    print(f"output tokens   : {TOKENS['output']:,}  (${cost_out:.6f})")
    print(f"total cost      : ${total_cost:.6f}")
    if pending:
        print(f"cost per 100 rows: ${total_cost / len(pending) * 100:.6f}")
    print("--------------------------")


if __name__ == "__main__":
    main()
