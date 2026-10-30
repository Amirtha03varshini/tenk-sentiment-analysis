"""
tenk_analyzer.py
================
SEC 10-K Sentiment & Financial Analysis Toolkit
===============================================

This module downloads a company's 10-K filings from the SEC's EDGAR
database, extracts the MD&A narrative and the three core financial
statements, scores the MD&A text with the Loughran-McDonald sentiment
dictionary, pulls stock prices at each balance-sheet date, computes
four analyst ratios, and writes everything into a formatted Excel
workbook with charts.

Usage
-----
    python tenk_analyzer.py AAPL 10

    The first argument is the ticker symbol, the second (optional)
    argument is the number of most-recent 10-Ks to analyze (default 10).

Design notes
------------
* PEP 8 style, type hints, and docstrings are used throughout.
* Every network call has retry logic and error handling.
* The code was developed in stages with an AI coding assistant
  (prompt history is included in the lab document).
"""

from __future__ import annotations

import datetime as dt
import json
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
from bs4 import BeautifulSoup
from lxml import etree
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

#: SEC requires a descriptive User-Agent that identifies the caller.
#: Set the SEC_USER_AGENT environment variable to your own contact
#: information before running (SEC fair-access policy).
import os

SEC_USER_AGENT = os.environ.get(
    "SEC_USER_AGENT",
    "10-K Research Project (contact: set SEC_USER_AGENT env var)",
)
SEC_HEADERS = {"User-Agent": SEC_USER_AGENT}

#: Base directories for all outputs (created relative to this file).
BASE_DIR = Path(__file__).resolve().parent
FILINGS_DIR = BASE_DIR / "filings"        # downloaded 10-K HTML files
MDNA_DIR = BASE_DIR / "mdna_text"         # extracted MD&A text files
DATA_DIR = BASE_DIR / "data"              # Loughran-McDonald dictionary
OUTPUT_DIR = BASE_DIR / "output"          # Excel workbooks

#: Path to the Loughran-McDonald master dictionary (CSV).
LM_DICTIONARY_PATH = (
    DATA_DIR / "Loughran-McDonald_MasterDictionary_1993-2025.csv"
)

#: Sentiment categories tracked (columns of the LM master dictionary).
LM_CATEGORIES = [
    "Negative", "Positive", "Uncertainty", "Litigious",
    "Strong_Modal", "Weak_Modal", "Constraining",
]

#: SEC allows at most 10 requests/second; we stay well below that.
REQUEST_DELAY_SECONDS = 0.15


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _http_get(url: str, retries: int = 4, timeout: int = 60) -> requests.Response:
    """GET a URL with retry/back-off and SEC-compliant headers.

    Args:
        url: The URL to fetch.
        retries: Number of attempts before giving up.
        timeout: Per-request timeout in seconds.

    Returns:
        A requests.Response object with status_code 200.

    Raises:
        RuntimeError: If all retries fail.
    """
    for attempt in range(1, retries + 1):
        try:
            time.sleep(REQUEST_DELAY_SECONDS)  # be polite to EDGAR
            response = requests.get(url, headers=SEC_HEADERS, timeout=timeout)
            if response.status_code == 200:
                return response
            # 404 means the resource genuinely does not exist.
            if response.status_code == 404:
                return response
            # SEC throttles (403/429) — back off substantially.
            if response.status_code in (403, 429):
                wait = min(10 * attempt, 30)
                print(
                    f"    [throttled {response.status_code}] backing "
                    f"off {wait}s ({url})"
                )
                time.sleep(wait)
                continue
        except requests.RequestException as exc:
            print(f"    [retry {attempt}/{retries}] {url} -> {exc}")
        time.sleep(1.5 * attempt)  # exponential-ish back-off
    raise RuntimeError(f"GET failed after {retries} attempts: {url}")


def _to_number(cell) -> Optional[float]:
    """Convert an SEC table cell to a float, handling $, commas, ().

    Parenthesized values are treated as negative, per accounting
    convention.
    """
    if cell is None:
        return None
    s = str(cell).strip()
    if not s or s.lower() in {"nan", "—", "-", "$", "%"}:
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("$", "").replace(",", "").replace(" ", "")
    s = s.split(" ")[0]
    try:
        value = float(s)
    except ValueError:
        return None
    return -value if negative else value


def _fmt_date(d: str) -> str:
    """Normalize a YYYY-MM-DD date string; return '' if invalid."""
    try:
        return dt.date.fromisoformat(d[:10]).isoformat()
    except (ValueError, TypeError):
        return ""


# ---------------------------------------------------------------------------
# Stage 1: Identify the company and list its 10-K filings
# ---------------------------------------------------------------------------

def get_cik(ticker: str) -> Tuple[str, str]:
    """Look up the SEC CIK number and company name for a ticker.

    Args:
        ticker: Stock ticker, e.g. ``AAPL``.

    Returns:
        Tuple (cik, company_name). The CIK is zero-padded to 10 digits
        where relevant for EDGAR URLs.

    Raises:
        ValueError: If the ticker cannot be resolved.
    """
    response = _http_get("https://www.sec.gov/files/company_tickers.json")
    mapping = response.json()
    ticker_upper = ticker.strip().upper()
    for entry in mapping.values():
        if entry["ticker"] == ticker_upper:
            return str(entry["cik_str"]).zfill(10), entry["title"]
    raise ValueError(f"Ticker {ticker!r} not found in SEC company list")


def _parse_filing_array(data: Dict, filings: List[Dict]) -> None:
    """Append 10-K entries from a submissions JSON fragment.

    Handles both shapes: the main submissions file (arrays under
    'filings.recent') and archived fragments (arrays at the top level).
    """
    if "filings" in data and "recent" in data.get("filings", {}):
        recent = data["filings"]["recent"]
    else:
        recent = data  # archived fragment: arrays at the top level
    n = len(recent["form"])
    for i in range(n):
        if recent["form"][i] != "10-K":
            continue
        report_date = _fmt_date(recent["reportDate"][i])
        filings.append(
            {
                "accession": recent["accessionNumber"][i],
                "filed": _fmt_date(recent["filingDate"][i]),
                "report_date": report_date,
                "primary_doc": recent["primaryDocument"][i],
                "fiscal_year": (
                    int(report_date[:4]) if report_date else None
                ),
                "index_url": "",  # filled in by get_10k_filings
            }
        )


def get_10k_filings(cik: str, count: int) -> List[Dict]:
    """Return the ``count`` most recent 10-K filings for a CIK.

    Each item is a dict with keys: accession, filed, report_date,
    primary_doc, fiscal_year, and index_url.

    Note: amended filings (10-K/A) are skipped so that each fiscal
    year appears exactly once. Filers with heavy filing volumes (e.g.
    Intel) push older 10-Ks into EDGAR's archived submission files,
    which are loaded on demand until ``count`` filings are found.
    """
    response = _http_get(
        f"https://data.sec.gov/submissions/CIK{int(cik):010d}.json"
    )
    data = response.json()
    filings: List[Dict] = []
    _parse_filing_array(data, filings)
    archive_files = list(data["filings"].get("files", []))
    while len(filings) < count and archive_files:
        archive = archive_files.pop(0)
        archive_url = (
            f"https://data.sec.gov/submissions/"
            f"{archive['name']}"
        )
        try:
            archive_response = _http_get(archive_url)
            _parse_filing_array(archive_response.json(), filings)
        except (RuntimeError, json.JSONDecodeError, ValueError):
            break
    for filing in filings:
        filing["index_url"] = (
            f"https://www.sec.gov/Archives/edgar/data/"
            f"{int(cik):d}/{filing['accession'].replace('-', '')}"
        )
    filings.sort(key=lambda f: f["report_date"], reverse=True)
    return filings[:count]


# ---------------------------------------------------------------------------
# Stage 2: Download and label the 10-K documents
# ---------------------------------------------------------------------------

def download_filings(ticker: str, filings: List[Dict]) -> List[Path]:
    """Download each 10-K's primary document with a descriptive name.

    Files are saved as ``filings/{TICKER}/{TICKER}_10K_FY{year}_{filed}.htm``.
    Returns the list of local file paths (oldest first).
    """
    company_dir = FILINGS_DIR / ticker.upper()
    company_dir.mkdir(parents=True, exist_ok=True)
    paths: List[Path] = []
    for filing in sorted(filings, key=lambda f: f["report_date"]):
        url = f"{filing['index_url']}/{filing['primary_doc']}"
        response = _http_get(url)
        if response.status_code != 200:
            print(f"    !! could not download {url} (skipped)")
            continue
        name = (
            f"{ticker.upper()}_10K_FY{filing['fiscal_year']}_"
            f"{filing['filed']}.htm"
        )
        path = company_dir / name
        path.write_bytes(response.content)
        paths.append(path)
        print(f"    saved {path.relative_to(BASE_DIR)}")
    return paths


# ---------------------------------------------------------------------------
# Stage 3: Extract the MD&A section (Item 7)
# ---------------------------------------------------------------------------

def _strip_tags(html: str) -> str:
    """Convert filing HTML to plain text, dropping tables' noise gently."""
    import warnings

    try:
        from bs4 import XMLParsedAsHTMLWarning

        warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
    except ImportError:
        pass
    soup = BeautifulSoup(html, "lxml")
    for node in soup(["script", "style"]):
        node.decompose()
    return soup.get_text("\n")


def _extract_mda_by_item7(text: str) -> Optional[str]:
    """Primary MD&A extraction: locate 'Item 7 ... Management's' headings.

    The terminator must be the Item 7A *heading* (always followed by
    'Quantitative ...'), never an inline cross-reference such as
    'Item 8 of this Form 10-K' inside the MD&A narrative. The
    table-of-contents occurrence is rejected because it yields only a
    few words before the next Item 7A line.
    """
    start_pat = re.compile(
        r"item\s+7\.?\s*[\.\:\-]*\s*(management|management\u2019s|"
        r"management's)",
        re.IGNORECASE,
    )
    end_pat = re.compile(
        r"item\s+7a\.?[\s\xa0]*(quantitative|qualitative)",
        re.IGNORECASE,
    )
    end_pat_loose = re.compile(r"item\s+7a", re.IGNORECASE)
    for match in start_pat.finditer(text):
        for pattern in (end_pat, end_pat_loose):
            end = pattern.search(text, match.start() + 50)
            if end is None:
                continue
            segment = text[match.start():end.start()]
            if len(segment.split()) > 500:  # substantive MD&A
                return segment
    return None


def _extract_mda_fallback(text: str) -> Optional[str]:
    """Fallback MD&A extraction for headings without an 'Item 7' prefix.

    Some filers (e.g. Intel) label the section simply as
    'Management's Discussion and Analysis' with the item number only
    in the table of contents. Candidates in the first 12% or last 8%
    of the text are skipped — those are the front table of contents,
    the back-of-document index, and the glossary of abbreviations.
    """
    heading_pat = re.compile(
        r"management.{0,3}s discussion and analysis", re.IGNORECASE
    )
    end_pat = re.compile(
        r"(quantitative and qualitative disclosures about market risk"
        r"|item\s+7a\.?[\s\xa0]*(quantitative|qualitative)"
        r"|item\s+8\.?[\s\xa0]*financial statements)",
        re.IGNORECASE,
    )
    n = len(text)
    for match in heading_pat.finditer(text):
        if match.start() < 0.08 * n or match.start() > 0.92 * n:
            continue
        end = end_pat.search(text, match.start() + 50)
        end_pos = end.start() if end else min(match.start() + 150000, n)
        segment = text[match.start():end_pos]
        if len(segment.split()) > 500:  # substantive MD&A
            return segment
    return None


def extract_mda(ticker: str, filing: Dict, html: str) -> Optional[str]:
    """Extract the Item 7 MD&A narrative from 10-K HTML.

    Strategy 1: locate every occurrence of an "Item 7" heading, cut the
    text at the following "Item 7A" heading, and keep the first
    candidate that yields a substantive passage. This rejects the
    table-of-contents occurrence, which produces only a few words.

    Strategy 2 (fallback): find standalone 'Management's Discussion
    and Analysis' headings (used by filers such as Intel).

    Returns:
        The MD&A text, or None if it cannot be located.
    """
    text = _strip_tags(html)
    mda = _extract_mda_by_item7(text)
    if mda is not None:
        return mda
    mda = _extract_mda_fallback(text)
    if mda is not None:
        return mda
    print(f"    !! MD&A not found for {filing['accession']}")
    return None


def save_mda(ticker: str, filing: Dict, mda_text: str) -> Path:
    """Write the MD&A text to a labeled file in mdna_text/."""
    MDNA_DIR.mkdir(parents=True, exist_ok=True)
    name = (
        f"{ticker.upper()}_MDA_FY{filing['fiscal_year']}_"
        f"{filing['filed']}.txt"
    )
    path = MDNA_DIR / name
    path.write_text(mda_text, encoding="utf-8")
    return path


# ---------------------------------------------------------------------------
# Stage 4: Extract the financial statements (R-files)
# ---------------------------------------------------------------------------

#: Patterns used to locate the main statement R-files by report name.
STATEMENT_PATTERNS = {
    "Income Statement": [
        r"CONSOLIDATED STATEMENTS? OF OPERATIONS",
        r"CONSOLIDATED STATEMENTS? OF INCOME",
    ],
    "Balance Sheet": [r"CONSOLIDATED BALANCE SHEETS"],
    "Cash Flows": [
        r"CONSOLIDATED STATEMENTS? OF CASH FLOWS",
    ],
}


#: A period header cell looks like 'Sep. 27, 2025' / 'Sep. 27, 2025 '.
_DATE_LIKE = re.compile(r"^[A-Za-z]{3,9}\.?\s+\d{1,2},\s+\d{4}$")


def _parse_rfile(url: str) -> pd.DataFrame:
    """Parse an SEC 'R' financial-report file into a DataFrame.

    R-files have a stable structure across filing vintages: a title
    row, a row of ``<th class="th">`` period headers (dates), and data
    rows whose first ``<td class="pl">`` is the line-item label and
    whose remaining cells are values. This parser uses that structure
    directly instead of generic HTML-table flattening.

    Returns a DataFrame indexed by line-item label whose columns are
    the period headers (dates) found in the file.
    """
    response = _http_get(url)
    if response.status_code != 200:
        raise RuntimeError(f"R-file not found: {url}")
    soup = BeautifulSoup(response.content, "lxml")
    table = soup.find("table", class_="report")
    if table is None:
        raise RuntimeError(f"No report table in {url}")
    rows = table.find_all("tr")
    # Statement title, e.g. 'CONSOLIDATED BALANCE SHEETS - USD ($)
    # shares in Thousands, $ in Millions' — kept so the workbook can
    # show the source units of each statement.
    title_cell = table.find("th", class_="tl")
    title = (
        re.sub(r"\s+", " ", title_cell.get_text(" ", strip=True))
        if title_cell is not None
        else ""
    )
    # Locate the header row: every th.th cell holds a date.
    header_row = None
    for tr in rows:
        ths = tr.find_all("th", class_="th")
        if ths and all(
            _DATE_LIKE.match(th.get_text(strip=True)) for th in ths
        ):
            header_row = tr
            break
    if header_row is None:
        raise RuntimeError(f"No period header row in {url}")
    periods = [
        th.get_text(strip=True)
        for th in header_row.find_all("th", class_="th")
    ]
    labels: List[str] = []
    data_rows: List[List[Optional[float]]] = []
    seen: Dict[str, int] = {}
    for tr in rows:
        if tr is header_row:
            continue
        tds = tr.find_all("td")
        if not tds:
            continue
        label = tds[0].get_text(" ", strip=True)
        if not label:
            continue
        values = [
            _to_number(td.get_text(" ", strip=True)) for td in tds[1:]
        ]
        # Pad rows that omit some period columns.
        values += [None] * (len(periods) - len(values))
        # Balance sheets repeat line labels across sections (e.g.
        # 'Marketable securities' in current and non-current assets).
        # Suffix repeats so labels stay unique for alignment.
        count = seen.get(label, 0) + 1
        seen[label] = count
        labels.append(label if count == 1 else f"{label} ({count})")
        data_rows.append(values[: len(periods)])
    frame = pd.DataFrame(data_rows, columns=periods, index=labels)
    frame.attrs["title"] = title
    return frame


def _normalize_period(col_name: str) -> str:
    """Normalize an R-file period header to YYYY-MM-DD where possible."""
    cleaned = re.sub(r"(12 Months Ended|3 Months Ended)", "", col_name)
    cleaned = cleaned.strip()
    for fmt in ("%b. %d, %Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            return dt.datetime.strptime(cleaned, fmt).date().isoformat()
        except ValueError:
            continue
    return col_name


def extract_statements(filing: Dict) -> Dict[str, pd.DataFrame]:
    """Extract IS / BS / CF DataFrames for a single 10-K via R-files.

    Returns a dict mapping statement name -> DataFrame with normalized
    ISO period columns. Missing statements are simply omitted.
    """
    base = filing["index_url"]
    summary = _http_get(f"{base}/FilingSummary.xml")
    if summary.status_code != 200:
        return {}
    root = etree.fromstring(summary.content)
    # <Report> nodes are nested under <Reports>, so search recursively.
    reports = list(root.iter("Report"))
    found: Dict[str, str] = {}
    for report in reports:
        short_name = (report.findtext("ShortName") or "").strip()
        file_name = (report.findtext("HtmlFileName") or "").strip()
        if "parenthetical" in short_name.lower():
            continue
        for statement, patterns in STATEMENT_PATTERNS.items():
            if statement in found:
                continue
            if any(re.search(p, short_name, re.IGNORECASE) for p in patterns):
                found[statement] = file_name
    statements: Dict[str, pd.DataFrame] = {}
    for statement, file_name in found.items():
        try:
            frame = _parse_rfile(f"{base}/{file_name}")
            frame.columns = [_normalize_period(c) for c in frame.columns]
            # Two header cells can normalize to the same period label;
            # keep only the first such column.
            frame = frame.loc[
                :, ~pd.Index(frame.columns).duplicated(keep="first")
            ]
            statements[statement] = frame
        except (RuntimeError, ValueError) as exc:
            print(f"    !! {statement} parse failed for {filing['accession']}: {exc}")
    return statements


def combine_statements(
    per_filing: List[Tuple[Dict, Dict[str, pd.DataFrame]]]
) -> Dict[str, pd.DataFrame]:
    """Merge each filing's statements into one wide table per statement.

    Columns are period dates (deduplicated, most recent filing wins for
    overlaps) and sorted oldest -> newest.
    """
    merged: Dict[str, Dict[str, pd.DataFrame]] = {}
    label_order: Dict[str, List[str]] = {}
    # Pass 1 (newest filing first): column values — on duplicate
    # periods, the most recent presentation of the numbers is kept.
    for filing, statements in sorted(
        per_filing, key=lambda t: t[0]["report_date"], reverse=True
    ):
        for statement, frame in statements.items():
            bucket = merged.setdefault(statement, {})
            for j, col in enumerate(frame.columns):
                if col not in bucket:  # newest filing wins
                    bucket[col] = frame.iloc[:, j]
    # Pass 2 (oldest filing first): row order follows the statement's
    # natural presentation order rather than pandas' sorted union.
    for filing, statements in sorted(
        per_filing, key=lambda t: t[0]["report_date"]
    ):
        for statement, frame in statements.items():
            order = label_order.setdefault(statement, [])
            for label in frame.index:
                if label not in order:
                    order.append(label)
    result: Dict[str, pd.DataFrame] = {}
    for statement, bucket in merged.items():
        frame = pd.DataFrame(bucket)
        frame = frame[sorted(frame.columns)]
        frame = frame.reindex(
            [l for l in label_order.get(statement, []) if l in frame.index]
        )
        # Carry the statement title (with units) of the most recent
        # filing forward onto the merged table.
        for filing, statements in sorted(
            per_filing, key=lambda t: t[0]["report_date"], reverse=True
        ):
            if statement in statements:
                frame.attrs["title"] = statements[statement].attrs.get(
                    "title", ""
                )
                break
        result[statement] = frame
    return result


# ---------------------------------------------------------------------------
# Stage 5: Clean the MD&A text for textual analysis
# ---------------------------------------------------------------------------

def clean_mda_text(raw_text: str) -> str:
    """Clean MD&A text for Loughran-McDonald scoring.

    Following the LM methodology (Loughran & McDonald 2011, JF):
      * lowercase everything,
      * remove XBRL/HTML artifacts and page furniture,
      * drop tokens containing digits,
      * keep alphabetic tokens (apostrophes allowed, then stripped).
    """
    text = raw_text.lower()
    # Remove common filing furniture that survives tag stripping.
    for junk in [
        r"table of contents", r"\f", r"item\s+\d+[a-z]*\.",
        r"sec\.?\s*file number.*",
    ]:
        text = re.sub(junk, " ", text)
    # Keep only word characters/apostrophes/whitespace.
    text = re.sub(r"[^a-z\s']", " ", text)
    text = re.sub(r"'", "", text)
    # Drop tokens with no alphabetic characters (pure numbers).
    tokens = [t for t in text.split() if re.search(r"[a-z]", t)]
    return " ".join(tokens)


# ---------------------------------------------------------------------------
# Stage 6: Loughran-McDonald sentiment scoring
# ---------------------------------------------------------------------------

def load_lm_dictionary(path: Path = LM_DICTIONARY_PATH) -> Dict[str, set]:
    """Load the LM master dictionary into category -> set-of-words maps.

    Returns:
        Dict with one entry per category in LM_CATEGORIES plus a
        'word_count' entry with the total dictionary size.
    """
    frame = pd.read_csv(path, low_memory=False)
    categories: Dict[str, set] = {}
    for category in LM_CATEGORIES:
        words = frame.loc[
            frame[category] > 0, "Word"
        ].astype(str)
        categories[category] = set(words.str.lower())
    return categories


def score_mda(text: str, dictionary: Dict[str, set]) -> Dict[str, float]:
    """Score cleaned MD&A text with the LM dictionary.

    Returns total word count, raw counts per category, the share of
    total words per category (in percent), and the P/N ratio
    (positive count / negative count; negative when negative == 0).
    """
    tokens = text.split()
    total = len(tokens)
    result: Dict[str, float] = {"Total Words": total}
    for category in LM_CATEGORIES:
        hits = sum(1 for t in tokens if t in dictionary[category])
        result[f"{category} Count"] = hits
        result[f"{category} %"] = (
            100.0 * hits / total if total else 0.0
        )
    positive = result["Positive Count"]
    negative = result["Negative Count"]
    result["P/N Ratio"] = (
        positive / negative if negative else float("nan")
    )
    return result


# ---------------------------------------------------------------------------
# Stage 7: XBRL financial data for ratios
# ---------------------------------------------------------------------------

#: us-gaap concept fallbacks per ratio input, tried in order.
CONCEPT_FALLBACKS = {
    "Revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
    ],
    "NetIncome": ["NetIncomeLoss", "ProfitLoss"],
    "Assets": ["Assets"],
    "AssetsCurrent": ["AssetsCurrent"],
    "LiabilitiesCurrent": ["LiabilitiesCurrent"],
    "Liabilities": ["Liabilities"],
    #: Total equity including any noncontrolling-interest portion.
    "EquityTotal": [
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
        "StockholdersEquity",
    ],
}


def _concept_series(cik: str, concept: str) -> Dict[str, float]:
    """Fetch annual (10-K, FY) values for one us-gaap concept.

    Values reported inside 10-Ks that cover only a quarter (Apple used
    to tag unaudited quarterly data) are excluded by requiring the
    duration between 'start' and 'end' to be roughly a full year.

    Returns a dict mapping period end date (YYYY-MM-DD) -> value.
    """
    url = (
        f"https://data.sec.gov/api/xbrl/companyconcept/"
        f"CIK{int(cik):010d}/us-gaap/{concept}.json"
    )
    try:
        response = _http_get(url)
    except RuntimeError:
        return {}
    if response.status_code != 200:
        return {}
    try:
        data = response.json()
    except json.JSONDecodeError:
        return {}
    unit_key = next(iter(data.get("units", {})), None)
    if unit_key is None:
        return {}
    series: Dict[str, float] = {}
    for item in data["units"][unit_key]:
        if item.get("form") != "10-K" or item.get("fp") != "FY":
            continue
        end = _fmt_date(item.get("end", ""))
        if not end:
            continue
        start = _fmt_date(item.get("start", ""))
        if start:
            try:
                duration_days = (
                    dt.date.fromisoformat(end)
                    - dt.date.fromisoformat(start)
                ).days
            except ValueError:
                continue
            if duration_days < 300:  # drop quarterly/interim periods
                continue
        # Keep the most recently filed value per period end.
        series[end] = item["val"]
    return series


def fetch_fundamentals(cik: str) -> Dict[str, Dict[str, float]]:
    """Fetch all ratio input series for a company from XBRL.

    For each input, the concept fallback list is merged: the first
    concept in the list wins for any period it covers, and later
    concepts only fill periods the earlier ones miss.
    """
    fundamentals: Dict[str, Dict[str, float]] = {}
    for label, concepts in CONCEPT_FALLBACKS.items():
        merged: Dict[str, float] = {}
        for concept in concepts:
            series = _concept_series(cik, concept)
            for end, value in series.items():
                if end not in merged:
                    merged[end] = value
        fundamentals[label] = merged
        print(
            f"    {label}: {len(fundamentals[label])} annual periods"
        )
    return fundamentals


def compute_ratios(
    fundamentals: Dict[str, Dict[str, float]],
    balance_dates: List[str],
) -> pd.DataFrame:
    """Compute the four required ratios for each fiscal year.

    Liquidity    : Current ratio = current assets / current liabilities
    Solvency     : Liabilities-to-equity = total liabilities / total
                   equity. Total liabilities come from the us-gaap
                   'Liabilities' tag when the filer reports it, and
                   otherwise are derived as total assets minus total
                   equity (including noncontrolling interests) so the
                   balance-sheet identity holds.
    Efficiency   : Asset turnover = revenue / average total assets
    Profitability: Net profit margin = net income / revenue

    For the earliest year in the window, average assets fall back to
    ending assets (the prior-year balance is outside the window).
    """
    rows = []
    for date in sorted(balance_dates):
        year = date[:4]
        assets = fundamentals["Assets"].get(date)
        assets_prev = None
        earlier = [d for d in fundamentals["Assets"] if d < date]
        if earlier:
            assets_prev = fundamentals["Assets"][max(earlier)]
        revenue = next(
            (
                fundamentals["Revenue"][d]
                for d in [date]
                if d in fundamentals["Revenue"]
            ),
            None,
        )
        if revenue is None:  # fiscal year ends may differ slightly
            candidates = [
                d for d in fundamentals["Revenue"]
                if d[:4] == year
            ]
            revenue = (
                fundamentals["Revenue"][candidates[0]] if candidates else None
            )
        net_income = next(
            (
                fundamentals["NetIncome"][d]
                for d in [date] if d in fundamentals["NetIncome"]
            ),
            None,
        )
        if net_income is None:
            candidates = [
                d for d in fundamentals["NetIncome"] if d[:4] == year
            ]
            net_income = (
                fundamentals["NetIncome"][candidates[0]] if candidates else None
            )
        current_assets = fundamentals["AssetsCurrent"].get(date)
        current_liab = fundamentals["LiabilitiesCurrent"].get(date)
        equity = fundamentals["EquityTotal"].get(date)
        assets_total = assets
        # Total liabilities: reported tag when available, else derived
        # from the balance-sheet identity (assets = liabilities +
        # equity) so the two sides always use matching definitions.
        liabilities = fundamentals["Liabilities"].get(date)
        if liabilities is None and assets_total is not None and equity:
            liabilities = assets_total - equity
        avg_assets = (
            (assets + assets_prev) / 2
            if assets is not None and assets_prev is not None
            else assets
        )
        rows.append(
            {
                "Fiscal Year": f"FY{year}",
                "Balance Sheet Date": date,
                "Revenue ($)": revenue,
                "Net Income ($)": net_income,
                "Total Assets ($)": assets,
                "Current Ratio": (
                    current_assets / current_liab
                    if current_assets and current_liab
                    else None
                ),
                "Liabilities-to-Equity": (
                    liabilities / equity
                    if liabilities is not None and equity
                    else None
                ),
                "Asset Turnover": (
                    revenue / avg_assets
                    if revenue is not None and avg_assets
                    else None
                ),
                "Net Profit Margin": (
                    net_income / revenue
                    if net_income is not None and revenue
                    else None
                ),
            }
        )
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Stage 8: Stock prices at balance-sheet dates
# ---------------------------------------------------------------------------

def fetch_stock_prices(ticker: str, dates: List[str]) -> Dict[str, float]:
    """Fetch each balance-sheet date's closing stock price (Yahoo).

    If the exact date is not a trading day, the nearest available
    trading day is used.
    """
    first = min(dates)
    last = max(dates)
    def _to_epoch(d: dt.date, offset_days: int = 0) -> int:
        """Convert a date to a UTC epoch timestamp."""
        moment = dt.datetime.combine(
            d + dt.timedelta(days=offset_days),
            dt.time(12),
            tzinfo=dt.timezone.utc,
        )
        return int(moment.timestamp())

    t0 = _to_epoch(dt.date.fromisoformat(first), -10)
    t1 = _to_epoch(dt.date.fromisoformat(last), +10)
    url = (
        f"https://query1.finance.yahoo.com/v8/finance/chart/"
        f"{ticker.upper()}?period1={t0}&period2={t1}&interval=1d"
    )
    response = requests.get(
        url, headers={"User-Agent": "Mozilla/5.0"}, timeout=30
    )
    response.raise_for_status()
    result = response.json()["chart"]["result"][0]
    timestamps = result["timestamp"]
    closes = result["indicators"]["quote"][0]["close"]
    prices: Dict[str, float] = {}
    for target in dates:
        target_date = dt.date.fromisoformat(target)
        # Prefer the last trading close on or before the balance-sheet
        # date; fall back to the earliest close after it.
        best_before: Optional[int] = None
        best_after: Optional[int] = None
        for i, ts in enumerate(timestamps):
            if closes[i] is None:
                continue
            d = dt.datetime.fromtimestamp(ts, dt.timezone.utc).date()
            if d <= target_date:
                if best_before is None or d >= dt.datetime.fromtimestamp(
                    timestamps[best_before], dt.timezone.utc
                ).date():
                    best_before = i
            elif best_after is None:
                best_after = i
        best_idx = best_before if best_before is not None else best_after
        if best_idx is not None:
            prices[target] = round(closes[best_idx], 2)
    return prices


# ---------------------------------------------------------------------------
# Stage 9: Build the Excel workbook
# ---------------------------------------------------------------------------

#: Shared style objects for the workbook.
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
THIN = Side(style="thin", color="B0B0B0")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _write_frame(ws, frame: pd.DataFrame, start_row: int = 1) -> None:
    """Write a DataFrame to a worksheet with styled headers."""
    columns = list(frame.columns)
    for j, column in enumerate(columns, start=1):
        cell = ws.cell(row=start_row, column=j, value=str(column))
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = BOX
    for i, (_, row) in enumerate(frame.iterrows(), start=start_row + 1):
        for j, column in enumerate(columns, start=1):
            value = row[column]
            if pd.isna(value):
                value = None
            elif isinstance(value, (int, float)):
                value = float(value)
            cell = ws.cell(row=i, column=j, value=value)
            cell.border = BOX
            if isinstance(value, float):
                cell.number_format = "#,##0.00"
    ws.freeze_panes = ws.cell(row=start_row + 1, column=2)


def _write_metadata(ws, company: Dict, filings: List[Dict]) -> None:
    """Write the README/metadata tab."""
    ws["A1"] = f"{company['name']} ({company['ticker']}) - 10-K Analysis"
    ws["A1"].font = TITLE_FONT
    rows = [
        ("CIK", company["cik"]),
        ("Generated (UTC)", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Sentiment dictionary", "Loughran-McDonald Master Dictionary "
         "1993-2025 (sraf.nd.edu, free for academic research)"),
        ("Price source", "Yahoo Finance daily closes; last trading close "
         "on or before each balance-sheet date"),
        ("Ratio inputs", "SEC XBRL company-concept API (us-gaap, 10-K "
         "annual values)"),
        ("Solvency ratio", "Total liabilities / total equity. Total "
         "liabilities use the us-gaap 'Liabilities' tag when reported, "
         "else total assets minus total equity (balance-sheet identity)."),
        ("Efficiency ratio", "Asset turnover = revenue / average of "
         "beginning and ending total assets."),
        ("MD&A scope note", "MD&A length varies by filer and year; "
         "compare category percentages (share of total words), not raw "
         "counts, across filings."),
    ]
    for i, (label, value) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        ws.cell(row=i, column=2, value=value)
    start = 10
    ws.cell(row=start - 1, column=1, value="Filings used").font = Font(
        bold=True, size=12
    )
    headers = ["Fiscal Year", "Filed", "Balance Sheet Date",
               "Accession Number", "EDGAR URL"]
    for j, header in enumerate(headers, start=1):
        cell = ws.cell(row=start, column=j, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BOX
    for i, filing in enumerate(
        sorted(filings, key=lambda f: f["report_date"]), start=start + 1
    ):
        for j, value in enumerate(
            [
                f"FY{filing['fiscal_year']}",
                filing["filed"],
                filing["report_date"],
                filing["accession"],
                filing["index_url"],
            ],
            start=1,
        ):
            cell = ws.cell(row=i, column=j, value=value)
            cell.border = BOX
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 70


def build_workbook(
    company: Dict,
    filings: List[Dict],
    statements: Dict[str, pd.DataFrame],
    sentiment: pd.DataFrame,
    prices: Dict[str, float],
    ratios: pd.DataFrame,
) -> Path:
    """Assemble the final Excel workbook with all required tabs.

    Tabs:
      1. README (metadata & filing index)
      2. Income Statement (all periods side by side)
      3. Balance Sheet
      4. Cash Flows
      5. Sentiment Analysis (LM dictionary scores)
      6. Stock Prices (balance-sheet dates)
      7. Financial Ratios (liquidity/solvency/efficiency/profitability)
      8. Chart (stock price vs P/N ratio)
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    _write_metadata(wb.create_sheet("README"), company, filings)

    for statement_name in ["Income Statement", "Balance Sheet", "Cash Flows"]:
        ws = wb.create_sheet(statement_name[:31])
        frame = statements.get(statement_name)
        if frame is not None:
            # Show the source title/units above the table.
            title = frame.attrs.get("title", "")
            if title:
                cell = ws.cell(row=1, column=1, value=title)
                cell.font = Font(bold=True, size=11, color="1F4E79")
            frame = frame.copy()
            # Present only the fiscal-year window covered by the
            # downloaded filings (e.g. FY2016-FY2025) so each tab
            # shows the ten periods side by side.
            fiscal_years = {
                f["fiscal_year"] for f in filings if f["fiscal_year"]
            }
            keep = [
                c for c in frame.columns
                if str(c)[:4].isdigit() and int(str(c)[:4]) in fiscal_years
            ]
            frame = frame[keep]
            frame_out = frame.reset_index().rename(
                columns={"index": "Line Item"}
            )
            frame_out.columns = [str(c) for c in frame_out.columns]
            _write_frame(ws, frame_out, start_row=3)
            ws.column_dimensions["A"].width = 45
            for j in range(2, frame_out.shape[1] + 1):
                ws.column_dimensions[
                    get_column_letter(j)
                ].width = 16

    # Sentiment tab.
    ws = wb.create_sheet("Sentiment Analysis")
    _write_frame(ws, sentiment)
    for j in range(1, sentiment.shape[1] + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14

    # Stock price tab.
    ws = wb.create_sheet("Stock Prices")
    price_frame = pd.DataFrame(
        [
            {
                "Fiscal Year": f"FY{d[:4]}",
                "Balance Sheet Date": d,
                "Closing Price ($)": p,
            }
            for d, p in sorted(prices.items())
        ]
    )
    _write_frame(ws, price_frame)
    ws.column_dimensions["A"].width = 14

    # Ratios tab.
    ws = wb.create_sheet("Ratios")
    _write_frame(ws, ratios)
    for j in range(1, ratios.shape[1] + 1):
        ws.column_dimensions[get_column_letter(j)].width = 16

    # Chart tab: stock price vs P/N ratio.
    ws = wb.create_sheet("Chart")
    ws["A1"] = (
        f"{company['ticker']}: Stock Price vs MD&A Sentiment (P/N Ratio)"
    )
    ws["A1"].font = TITLE_FONT
    chart_frame = ratios[["Fiscal Year", "Balance Sheet Date"]].copy()
    chart_frame["P/N Ratio"] = [
        sentiment.loc[
            sentiment["Fiscal Year"] == fy, "P/N Ratio"
        ].iloc[0]
        if (sentiment["Fiscal Year"] == fy).any()
        else None
        for fy in chart_frame["Fiscal Year"]
    ]
    chart_frame["Stock Price ($)"] = [
        prices.get(d) for d in chart_frame["Balance Sheet Date"]
    ]
    chart_frame = chart_frame.dropna(subset=["P/N Ratio"])
    _write_frame(ws, chart_frame, start_row=3)
    n_rows = chart_frame.shape[0]
    price_col = chart_frame.columns.get_loc("Stock Price ($)") + 1
    pn_col = chart_frame.columns.get_loc("P/N Ratio") + 1

    # Primary-axis line chart: stock price.
    chart_price = LineChart()
    chart_price.title = (
        f"{company['ticker']} - Stock Price vs MD&A Sentiment P/N Ratio"
    )
    chart_price.style = 12
    chart_price.y_axis.title = "Stock Price ($)"
    chart_price.x_axis.title = "Fiscal Year"
    price_ref = Reference(
        ws,
        min_col=price_col,
        min_row=3,
        max_row=3 + n_rows,
    )
    cats = Reference(
        ws, min_col=1, min_row=4, max_row=3 + n_rows
    )
    chart_price.add_data(price_ref, titles_from_data=True)
    chart_price.set_categories(cats)
    chart_price.height = 10
    chart_price.width = 24

    # Secondary-axis line chart: P/N ratio.
    chart_pn = LineChart()
    chart_pn.y_axis.axId = 200
    chart_pn.y_axis.title = "P/N Ratio"
    chart_pn.y_axis.crosses = "max"
    pn_ref = Reference(
        ws, min_col=pn_col, min_row=3, max_row=3 + n_rows
    )
    chart_pn.add_data(pn_ref, titles_from_data=True)
    chart_pn.y_axis.majorGridlines = None
    chart_price += chart_pn
    ws.add_chart(chart_price, "H3")

    output_path = (
        OUTPUT_DIR / f"{company['ticker']}_10K_Analysis.xlsx"
    )
    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def analyze(ticker: str, count: int = 10) -> Path:
    """Run the full pipeline for one company; return the workbook path."""
    print(f"\n=== {ticker.upper()} : resolving company ===")
    cik, name = get_cik(ticker)
    company = {"ticker": ticker.upper(), "cik": cik, "name": name}
    print(f"    {name} (CIK {cik})")

    print(f"=== {ticker.upper()} : listing {count} most recent 10-Ks ===")
    filings = get_10k_filings(cik, count)
    for filing in filings:
        print(
            f"    FY{filing['fiscal_year']} 10-K filed "
            f"{filing['filed']} (B/S date {filing['report_date']})"
        )

    print(f"=== {ticker.upper()} : downloading filings ===")
    download_filings(ticker, filings)

    print(f"=== {ticker.upper()} : extracting MD&A & statements ===")
    mda_results: List[Tuple[Dict, str]] = []
    statement_results: List[Tuple[Dict, Dict[str, pd.DataFrame]]] = []
    for filing in filings:
        path = (
            FILINGS_DIR / ticker.upper()
            / (
                f"{ticker.upper()}_10K_FY{filing['fiscal_year']}_"
                f"{filing['filed']}.htm"
            )
        )
        if not path.exists():
            continue
        html = path.read_text(encoding="utf-8", errors="ignore")
        mda = extract_mda(ticker, filing, html)
        if mda:
            saved = save_mda(ticker, filing, mda)
            print(f"    MD&A -> {saved.relative_to(BASE_DIR)}")
            mda_results.append((filing, mda))
        statements = extract_statements(filing)
        print(
            f"    FY{filing['fiscal_year']}: "
            f"{list(statements.keys()) or 'no statements parsed'}"
        )
        statement_results.append((filing, statements))

    print(f"=== {ticker.upper()} : cleaning text & scoring sentiment ===")
    dictionary = load_lm_dictionary()
    sentiment_rows = []
    for filing, mda in mda_results:
        cleaned = clean_mda_text(mda)
        # Persist the cleaned text that is actually scored, so the
        # scored input can be inspected and audited.
        cleaned_path = (
            MDNA_DIR
            / (
                f"{ticker.upper()}_MDA_FY{filing['fiscal_year']}_"
                f"{filing['filed']}_cleaned.txt"
            )
        )
        cleaned_path.write_text(cleaned, encoding="utf-8")
        scores = score_mda(cleaned, dictionary)
        scores = {"Fiscal Year": f"FY{filing['fiscal_year']}", **scores}
        sentiment_rows.append(scores)
    sentiment = pd.DataFrame(sentiment_rows)
    if sentiment.empty:
        raise RuntimeError(
            "No MD&A sections were extracted; cannot score sentiment."
        )
    sentiment = sentiment.sort_values("Fiscal Year")
    print(sentiment.to_string(index=False))

    print(f"=== {ticker.upper()} : fetching fundamentals (XBRL) ===")
    fundamentals = fetch_fundamentals(cik)

    balance_dates = [f["report_date"] for f in filings]
    print(f"=== {ticker.upper()} : computing ratios ===")
    ratios = compute_ratios(fundamentals, balance_dates)
    print(ratios.to_string(index=False))

    print(f"=== {ticker.upper()} : fetching stock prices ===")
    prices = fetch_stock_prices(ticker, balance_dates)
    for date, price in sorted(prices.items()):
        print(f"    {date}: ${price:,.2f}")

    # ---- Completion checks: every requested year must be present. --
    problems = []
    if len(filings) < count:
        problems.append(
            f"only {len(filings)} 10-Ks found (requested {count})"
        )
    if len(mda_results) < len(filings):
        problems.append(
            f"MD&A extracted for {len(mda_results)} of "
            f"{len(filings)} filings"
        )
    if len(sentiment_rows) < len(filings):
        problems.append(
            f"sentiment scored for {len(sentiment_rows)} of "
            f"{len(filings)} filings"
        )
    missing_prices = [f for f in filings if f["report_date"] not in prices]
    if missing_prices:
        problems.append(
            "missing stock prices for "
            + ", ".join(f["report_date"] for f in missing_prices)
        )
    if len(ratios) < len(filings) or ratios.iloc[:, 3:].isna().any().any():
        problems.append("ratio table incomplete (missing years or values)")
    if problems:
        raise RuntimeError(
            "Completion check failed: " + "; ".join(problems)
        )

    print(f"=== {ticker.upper()} : building Excel workbook ===")
    statements = combine_statements(statement_results)
    workbook_path = build_workbook(
        company, filings, statements, sentiment, prices, ratios
    )
    print(f"    workbook -> {workbook_path}")
    return workbook_path


def main(argv: Optional[List[str]] = None) -> int:
    """Command-line entry point."""
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print(__doc__)
        return 1
    ticker = argv[0]
    count = int(argv[1]) if len(argv) > 1 else 10
    analyze(ticker, count)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
