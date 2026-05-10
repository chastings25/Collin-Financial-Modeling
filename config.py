import os
from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
EDGAR_USER_AGENT = os.getenv("EDGAR_USER_AGENT", "10k-model-builder contact@example.com")

CLAUDE_MODEL = "claude-sonnet-4-6"

# ── EDGAR endpoints ──────────────────────────────────────────────────────────
EDGAR_BASE_URL = "https://data.sec.gov"
EDGAR_ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data"
EDGAR_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
EDGAR_RATE_LIMIT_SLEEP = 0.15  # seconds between requests (SEC allows 10 req/s)

# ── Excel cell font colors (ARGB hex) ────────────────────────────────────────
COLOR_INPUT = "FF0000CD"       # blue  — hardcoded inputs
COLOR_FORMULA = "FF000000"     # black — calculated cells
COLOR_LINK = "FF006400"        # green — cross-sheet formula links
COLOR_HEADER_TEXT = "FFFFFFFF" # white — text on dark header rows
COLOR_NEGATIVE = "FF8B0000"    # dark red — negative variance cells

# ── Excel fill colors (ARGB hex) ─────────────────────────────────────────────
FILL_SECTION_HEADER = "FF1F3864"   # dark navy
FILL_SUBSECTION = "FFD6E4F0"       # light blue
FILL_TOTAL = "FFE2EFDA"            # light green
FILL_GRAND_TOTAL = "FFC6E0B4"      # slightly deeper green
FILL_INPUT = "FFFFFF00"            # yellow
FILL_CHECK_OK = "FF92D050"         # bright green — balance check pass
FILL_CHECK_FAIL = "FFFF0000"       # red — balance check fail
FILL_WHITE = "FFFFFFFF"
FILL_ALT_ROW = "FFF2F2F2"          # light grey — alternate row shading

# ── Tab colors ────────────────────────────────────────────────────────────────
TAB_IS = "00366092"
TAB_BS = "00375623"
TAB_CF = "00215868"
TAB_RATIOS = "003B1F6A"
TAB_BRIDGE = "007F6000"

# ── Number formats ────────────────────────────────────────────────────────────
FMT_CURRENCY = '#,##0;(#,##0);"-"'
FMT_CURRENCY_DEC = '#,##0.0;(#,##0.0);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_PCT_1 = '0.1%;(0.1%);"-"'
FMT_MULTIPLE = '0.0"x"'
FMT_INTEGER = '#,##0;(#,##0);"-"'
FMT_TEXT = "@"

# ── Font ──────────────────────────────────────────────────────────────────────
FONT_NAME = "Calibri"
FONT_SIZE_NORMAL = 10
FONT_SIZE_HEADER = 11
FONT_SIZE_TITLE = 13

# ── Column widths ─────────────────────────────────────────────────────────────
COL_WIDTH_LABEL = 34
COL_WIDTH_YEAR = 14
COL_WIDTH_SPACER = 2

# ── Row heights ───────────────────────────────────────────────────────────────
ROW_HEIGHT_NORMAL = 15
ROW_HEIGHT_HEADER = 18
ROW_HEIGHT_SPACER = 6
ROW_HEIGHT_TITLE = 20

# ── Sheet names (locked — referenced by cross-sheet formulas) ────────────────
SHEET_IS = "Income Statement"
SHEET_BS = "Balance Sheet"
SHEET_CF = "Cash Flow"
SHEET_RATIOS = "Ratios & KPIs"
SHEET_BRIDGE = "EBITDA Bridge"
