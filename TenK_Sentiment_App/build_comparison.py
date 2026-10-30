"""
build_comparison.py — Cross-company comparison layer.

Reads both companies' Excel workbooks, pulls additional XBRL series
needed for the graduate analysis questions (R&D, gross profit, capex,
buybacks, dividends), renders the matplotlib figures used in the lab
document, and writes a JSON dataset consumed by the HTML dashboard.

Run:  python build_comparison.py
"""

from pathlib import Path
import json
import os
import sys

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from tenk_analyzer import (  # noqa: E402
    _concept_series,
    SEC_USER_AGENT,
)

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "output"
FIGURES = OUTPUT / "figures"

CIKS = {"AAPL": "0000320193", "INTC": "0000050863"}

EXTRA_CONCEPTS = {
    "R&D Expense": ["ResearchAndDevelopmentExpense"],
    "Gross Profit": ["GrossProfit"],
    "CapEx": ["PaymentsToAcquirePropertyPlantAndEquipment"],
    "Buybacks": [
        "PaymentsForRepurchaseOfCommonStock",
        "PaymentsForRepurchaseOfCommonStockAndCashSettlementsOfDerivative",
    ],
    "Dividends": [
        "PaymentsOfDividends",
        "PaymentsOfDividendsCommonStock",
        "PaymentsOfDividendsAndDividendEquivalentsDeclaredOrPaid",
    ],
}


def load_workbook_data(ticker: str) -> dict:
    """Load the sentiment / price / ratio tabs for one company."""
    path = OUTPUT / f"{ticker}_10K_Analysis.xlsx"
    ratios = pd.read_excel(path, sheet_name="Ratios")
    sentiment = pd.read_excel(path, sheet_name="Sentiment Analysis")
    prices = pd.read_excel(path, sheet_name="Stock Prices")
    data = ratios.merge(sentiment, on="Fiscal Year")
    data = data.merge(
        prices[["Fiscal Year", "Closing Price ($)"]],
        on="Fiscal Year",
    )
    return {
        "data": data,
        "fiscal_years": ratios["Fiscal Year"].tolist(),
        "dates": ratios["Balance Sheet Date"].tolist(),
    }


def fetch_extras(ticker: str, dates: list) -> pd.DataFrame:
    """Fetch the extra XBRL concepts aligned to the fiscal years."""
    frame = pd.DataFrame({"Balance Sheet Date": dates})
    for label, concepts in EXTRA_CONCEPTS.items():
        merged: dict = {}
        for concept in concepts:
            series = _concept_series(CIKS[ticker], concept)
            for end, value in series.items():
                if end not in merged:
                    merged[end] = value
        frame[label] = [
            merged.get(date) for date in frame["Balance Sheet Date"]
        ]
    return frame


def build_dataset() -> dict:
    """Build the combined JSON dataset for both companies."""
    dataset = {"companies": {}}
    for ticker in ["AAPL", "INTC"]:
        loaded = load_workbook_data(ticker)
        data = loaded["data"].merge(
            fetch_extras(ticker, loaded["dates"]),
            on="Balance Sheet Date",
        )
        # Derived metrics used by the analysis questions.
        data["R&D Intensity"] = data["R&D Expense"] / data["Revenue ($)"]
        data["Gross Margin"] = (
            data["Gross Profit"] / data["Revenue ($)"]
        )
        records = data.to_dict(orient="records")

        def _json_safe(value):
            """Convert pandas/numpy scalars to plain JSON types."""
            if value is None:
                return None
            try:
                if pd.isna(value):
                    return None
            except (TypeError, ValueError):
                pass
            if hasattr(value, "item"):  # numpy scalar
                return value.item()
            if isinstance(value, (int, float, str, bool)):
                return value
            return str(value)

        dataset["companies"][ticker] = [
            {key: _json_safe(value) for key, value in record.items()}
            for record in records
        ]
    return dataset


def render_figures() -> None:
    """Render the matplotlib figures used in the lab document."""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    FIGURES.mkdir(parents=True, exist_ok=True)
    dataset = build_dataset()
    aapl = pd.DataFrame(dataset["companies"]["AAPL"])
    intc = pd.DataFrame(dataset["companies"]["INTC"])

    # --- Figure 1 & 2: sentiment (P/N) vs stock price, per company --
    for ticker, frame in [("AAPL", aapl), ("INTC", intc)]:
        fig, ax1 = plt.subplots(figsize=(9, 4.5))
        ax2 = ax1.twinx()
        ax1.plot(
            frame["Fiscal Year"], frame["P/N Ratio"],
            color="#1F4E79", marker="o", label="MD&A P/N ratio",
        )
        ax2.plot(
            frame["Fiscal Year"], frame["Closing Price ($)"],
            color="#C0504D", marker="s", label="Stock price ($)",
        )
        ax1.set_ylabel("Positive/Negative ratio", color="#1F4E79")
        ax2.set_ylabel("Stock price ($)", color="#C0504D")
        ax1.set_title(
            f"{ticker}: MD&A sentiment (P/N ratio) vs stock price, "
            "FY2016-FY2025"
        )
        ax1.tick_params(axis="x", rotation=45)
        ax1.grid(alpha=0.3)
        fig.tight_layout()
        fig.savefig(
            FIGURES / f"fig_sentiment_price_{ticker.lower()}.png",
            dpi=150,
        )
        plt.close(fig)

    # --- Figure 3: the four ratios side by side ----------------------
    fig, axes = plt.subplots(2, 2, figsize=(11, 7))
    panels = [
        ("Current Ratio", "Liquidity - Current ratio"),
        ("Liabilities-to-Equity", "Solvency - Liabilities to equity"),
        ("Asset Turnover", "Efficiency - Asset turnover"),
        ("Net Profit Margin", "Profitability - Net profit margin"),
    ]
    for ax, (column, title) in zip(axes.flat, panels):
        ax.plot(aapl["Fiscal Year"], aapl[column], marker="o",
                label="AAPL", color="#1F4E79")
        ax.plot(intc["Fiscal Year"], intc[column], marker="s",
                label="INTC", color="#C0504D")
        ax.set_title(title)
        ax.tick_params(axis="x", rotation=45)
        ax.legend()
        ax.grid(alpha=0.3)
    fig.suptitle("AAPL vs INTC - financial ratios, FY2016-FY2025")
    fig.tight_layout()
    fig.savefig(FIGURES / "fig_ratios.png", dpi=150)
    plt.close(fig)

    # --- Figure 4: R&D intensity ------------------------------------
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.plot(aapl["Fiscal Year"], 100 * aapl["R&D Intensity"],
            marker="o", label="AAPL", color="#1F4E79")
    ax.plot(intc["Fiscal Year"], 100 * intc["R&D Intensity"],
            marker="s", label="INTC", color="#C0504D")
    ax.set_title("R&D intensity (R&D expense / revenue), % of revenue")
    ax.set_ylabel("% of revenue")
    ax.tick_params(axis="x", rotation=45)
    ax.legend()
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(FIGURES / "fig_rnd.png", dpi=150)
    plt.close(fig)

    # --- Figure 5: gross margin -------------------------------------
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.plot(aapl["Fiscal Year"], 100 * aapl["Gross Margin"],
            marker="o", label="AAPL", color="#1F4E79")
    ax.plot(intc["Fiscal Year"], 100 * intc["Gross Margin"],
            marker="s", label="INTC", color="#C0504D")
    ax.set_title("Gross margin, % of revenue")
    ax.set_ylabel("% of revenue")
    ax.tick_params(axis="x", rotation=45)
    ax.legend()
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(FIGURES / "fig_margin.png", dpi=150)
    plt.close(fig)

    # --- Figure 6: capital allocation --------------------------------
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5), sharey=False)
    for ax, (ticker, frame) in zip(
        axes, [("AAPL", aapl), ("INTC", intc)]
    ):
        revenue = frame["Revenue ($)"] / 1e9
        capex = frame["CapEx"].fillna(0) / 1e9
        buybacks = frame["Buybacks"].fillna(0) / 1e9
        dividends = frame["Dividends"].fillna(0) / 1e9
        positions = range(len(frame))
        ax.bar(positions, capex, label="CapEx", color="#1F4E79")
        ax.bar(positions, buybacks, bottom=capex,
               label="Buybacks", color="#9BBB59")
        ax.bar(positions, dividends, bottom=capex + buybacks,
               label="Dividends", color="#C0504D")
        ax.plot(positions, revenue, color="black",
                marker="o", linewidth=1.5, label="Revenue")
        ax.set_xticks(list(positions))
        ax.set_xticklabels(frame["Fiscal Year"], rotation=45)
        ax.set_title(f"{ticker}: capital allocation ($B)")
        ax.legend(fontsize=8)
    fig.tight_layout()
    fig.savefig(FIGURES / "fig_capital.png", dpi=150)
    plt.close(fig)

    # Persist the dataset for the HTML dashboard.
    with open(OUTPUT / "comparison_data.json", "w") as handle:
        json.dump(dataset, handle, indent=1)
    print(f"figures written to {FIGURES}")
    print(f"dataset written to {OUTPUT / 'comparison_data.json'}")


if __name__ == "__main__":
    render_figures()
