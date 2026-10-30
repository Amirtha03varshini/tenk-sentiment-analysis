# SEC 10-K Sentiment & Financial Analysis Toolkit

Vibe-coded final project: downloads a company's last ten 10-K filings
from SEC EDGAR, extracts the MD&A narrative and the three core
financial statements, scores the MD&A text with the Loughran-McDonald
sentiment dictionary, pulls stock prices at each balance-sheet date,
computes four analyst ratios, and writes a formatted Excel workbook
with charts — for **Apple (AAPL)** and **Intel (INTC)**.

## Quick start

```bash
pip install -r requirements.txt

# SEC fair-access policy: identify yourself (name + email)
export SEC_USER_AGENT="Your Name (your-email@example.com)"   # macOS/Linux
set SEC_USER_AGENT=Your Name (your-email@example.com)        # Windows

# One-time: download the Loughran-McDonald master dictionary (~9 MB)
python download_dictionary.py

# Run the pipeline (CLI)
python tenk_analyzer.py AAPL 10
python tenk_analyzer.py INTC 10

# Or use the desktop app (GUI)
python tenk_desktop_app.py

# Optional: comparison figures + dashboard dataset
python build_comparison.py

# Verification checks (completeness, balance-sheet identity, P/N math)
python verify_outputs.py
```

## What you get

For each company, `output/{TICKER}_10K_Analysis.xlsx` with:

| Tab | Contents |
|---|---|
| README | Metadata, methodology notes, filing index with EDGAR links |
| Income Statement | All periods side by side (from the 10-K R-files) |
| Balance Sheet | All periods side by side |
| Cash Flows | All periods side by side |
| Sentiment Analysis | LM dictionary word counts, category shares, P/N ratio |
| Stock Prices | Close on each balance-sheet date (Yahoo Finance) |
| Ratios | Current ratio, liabilities-to-equity, asset turnover, net margin |
| Chart | Native Excel chart: stock price vs MD&A P/N ratio |

Supporting folders:

- `filings/` — downloaded 10-K documents, labeled `TICKER_10K_FY{year}_{filed}.htm`
- `mdna_text/` — extracted MD&A text (raw + cleaned, labeled by fiscal year)
- `output/figures/` — comparison charts; `output/comparison_data.json` feeds the HTML dashboard

## Files

| File | Purpose |
|---|---|
| `tenk_analyzer.py` | Core pipeline (EDGAR download, MD&A extraction, R-file statement parsing, LM sentiment, XBRL ratios, stock prices, Excel workbook) |
| `tenk_desktop_app.py` | Tkinter desktop app wrapping the pipeline |
| `download_dictionary.py` | Fetches the LM master dictionary from SRAF (Google Drive) |
| `build_comparison.py` | Cross-company figures + dashboard dataset |
| `verify_outputs.py` | Reconciliation checks (completeness, balance-sheet identity, statement-vs-XBRL, P/N arithmetic) |
| `audit_mda.py` | MD&A boundary audit (start/end positions per filing) |
| `dashboard/index.html` | Self-contained HTML dashboard (Chart.js) |

## Data sources

- SEC EDGAR: submissions JSON, filing documents, R-file financial
  statements, XBRL company-concept API
- Loughran-McDonald Master Dictionary 1993-2025 (SRAF, Notre Dame) —
  free for academic research
- Yahoo Finance chart API for daily closes

## Notes

- The Loughran-McDonald dictionary CSV is NOT committed to GitHub
  (check the license); run `download_dictionary.py` instead.
- MD&A length varies by filer and year — compare category
  percentages (share of total words) rather than raw counts.
- Prices use the last trading close on or before each balance-sheet
  date.
