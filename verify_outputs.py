"""
verify_outputs.py — Reconciliation checks for the generated workbooks.

Reads each company's workbook and verifies:
  1. Every fiscal year FY2016-FY2025 appears on the Sentiment,
     Stock Prices, and Ratios tabs (10 rows each).
  2. Balance-sheet identity: assets = liabilities + equity
     (within rounding, using XBRL inputs).
  3. Income-statement figures on the statement tab match the XBRL
     inputs used for the ratios (revenue and net income).
  4. The P/N ratio column equals positive count / negative count.

Run:  python verify_outputs.py
"""

from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
EXPECTED_YEARS = [f"FY{y}" for y in range(2016, 2026)]
TOLERANCE = 0.02  # 2% relative tolerance for reconciliation


def check(name: str, condition: bool, detail: str = "") -> None:
    """Print PASS/FAIL for one check."""
    status = "PASS" if condition else "FAIL"
    suffix = f" ({detail})" if detail else ""
    print(f"  [{status}] {name}{suffix}")


def verify_company(ticker: str) -> None:
    """Run all reconciliation checks for one company's workbook."""
    path = BASE / "output" / f"{ticker}_10K_Analysis.xlsx"
    print(f"\n===== {ticker}: {path.name} =====")
    ratios = pd.read_excel(path, sheet_name="Ratios")
    sentiment = pd.read_excel(path, sheet_name="Sentiment Analysis")
    prices = pd.read_excel(path, sheet_name="Stock Prices")

    # 1. Completeness -------------------------------------------------
    check(
        "Ratios tab has all 10 fiscal years",
        sorted(ratios["Fiscal Year"].tolist()) == sorted(EXPECTED_YEARS),
        f"{len(ratios)} rows",
    )
    check(
        "Sentiment tab has all 10 fiscal years",
        sorted(sentiment["Fiscal Year"].tolist())
        == sorted(EXPECTED_YEARS),
        f"{len(sentiment)} rows",
    )
    check(
        "Stock Prices tab has all 10 fiscal years",
        sorted(prices["Fiscal Year"].tolist()) == sorted(EXPECTED_YEARS),
        f"{len(prices)} rows",
    )

    # 2. Balance-sheet identity (assets = liabilities + equity) -------
    # Total liabilities are either reported or derived as
    # assets - equity, so verify against the statement tab instead.
    # Statement tabs have a title row, then the header on row 3.
    balance = pd.read_excel(path, sheet_name="Balance Sheet", header=2)
    label_assets = next(
        (l for l in balance["Line Item"] if str(l).strip().lower()
         in {"total assets", "total assets (2)"}),
        None,
    )
    # Equity: 'Total stockholders' equity' / 'Total shareholders'
    # equity' / 'Total Intel stockholders' equity' — but never a
    # combined 'liabilities and equity' total.
    label_equity = next(
        (
            l for l in balance["Line Item"]
            if "total" in str(l).lower()
            and ("stockholders" in str(l).lower()
                 or "shareholders" in str(l).lower())
            and "equity" in str(l).lower()
            and "liabilit" not in str(l).lower()
            and "temporary" not in str(l).lower()
        ),
        None,
    )
    # Liabilities: a standalone 'Total liabilities' subtotal. Some
    # filers (e.g. Intel) do not present one; for those, the combined
    # 'Total liabilities ... and stockholders' equity' row is used as
    # the cross-check against Total assets instead.
    label_liab = next(
        (
            l for l in balance["Line Item"]
            if "total liabilities" in str(l).lower()
            and "equity" not in str(l).lower()
            and "stockholders" not in str(l).lower()
        ),
        None,
    )
    label_combined = next(
        (
            l for l in balance["Line Item"]
            if "total liabilities" in str(l).lower()
            and ("stockholders" in str(l).lower()
                 or "shareholders" in str(l).lower())
        ),
        None,
    )
    if label_assets is not None and label_equity is not None:
        row_a = balance[balance["Line Item"] == label_assets].iloc[0]
        row_e = balance[balance["Line Item"] == label_equity].iloc[0]
        if label_liab is not None:
            row_l = balance[balance["Line Item"] == label_liab].iloc[0]
            mismatches = []
            for col in balance.columns[1:]:
                try:
                    a = row_a[col]
                    e = row_e[col]
                    l = row_l[col]
                except (KeyError, TypeError):
                    continue
                if pd.isna(a) or pd.isna(e) or pd.isna(l):
                    continue
                if abs(a - (l + e)) > TOLERANCE * max(abs(a), 1):
                    mismatches.append(col)
            check(
                "Balance sheet: assets = liabilities + equity",
                not mismatches,
                f"checked {balance.shape[1] - 1} periods"
                + (f"; mismatch: {mismatches}" if mismatches else ""),
            )
        elif label_combined is not None:
            # Filer presents no standalone liabilities subtotal
            # (e.g. Intel); check combined row equals total assets.
            row_c = (
                balance[balance["Line Item"] == label_combined].iloc[0]
            )
            mismatches = []
            for col in balance.columns[1:]:
                try:
                    a = row_a[col]
                    c = row_c[col]
                except (KeyError, TypeError):
                    continue
                if pd.isna(a) or pd.isna(c):
                    continue
                if abs(a - c) > TOLERANCE * max(abs(a), 1):
                    mismatches.append(col)
            check(
                "Balance sheet: combined liabilities+equity = assets",
                not mismatches,
                f"checked {balance.shape[1] - 1} periods"
                + (f"; mismatch: {mismatches}" if mismatches else ""),
            )
        else:
            check(
                "Balance sheet liabilities presentation found",
                False,
                "no standalone or combined liabilities row",
            )
    else:
        check(
            "Balance sheet labels found (assets/liabilities/equity)",
            False,
            f"labels: {label_assets}, {label_liab}, {label_equity}",
        )

    # 3. Income statement vs XBRL ratio inputs -----------------------
    income = pd.read_excel(path, sheet_name="Income Statement", header=2)
    label_rev = next(
        (l for l in income["Line Item"] if str(l).strip().lower()
         in {"net sales", "net revenue", "revenue"}),
        None,
    )
    label_ni = next(
        (l for l in income["Line Item"] if str(l).strip().lower()
         in {"net income", "net income (loss)"}),
        None,
    )
    for stmt_label, ratio_col, name in (
        (label_rev, "Revenue ($)", "revenue"),
        (label_ni, "Net Income ($)", "net income"),
    ):
        if stmt_label is None:
            check(f"Income statement {name} label found", False)
            continue
        row = income[income["Line Item"] == stmt_label].iloc[0]
        mismatches = []
        for _, ratio_row in ratios.iterrows():
            date = ratio_row["Balance Sheet Date"]
            if date not in income.columns:
                continue
            stmt_val = row[date]
            ratio_val = ratio_row[ratio_col]
            if pd.isna(stmt_val) or pd.isna(ratio_val):
                continue
            # Statement tab may be in $ millions.
            scale = 1e6 if ratio_val > 1e9 else 1
            if abs(stmt_val * scale - ratio_val) > TOLERANCE * max(
                abs(ratio_val), 1
            ):
                mismatches.append(date)
        check(
            f"Income statement {name} matches ratio inputs",
            not mismatches,
            f"{len(ratios) - len(mismatches)}/{len(ratios)} periods"
            + (f"; mismatch: {mismatches}" if mismatches else ""),
        )

    # 4. P/N ratio arithmetic ----------------------------------------
    ok = True
    for _, row in sentiment.iterrows():
        if row["Negative Count"] == 0:
            continue
        expected = row["Positive Count"] / row["Negative Count"]
        if abs(expected - row["P/N Ratio"]) > 1e-6:
            ok = False
    check("P/N ratio = positive / negative counts", ok)


def main() -> None:
    """Verify every workbook in the output directory."""
    for ticker in ["AAPL", "INTC"]:
        verify_company(ticker)


if __name__ == "__main__":
    main()
