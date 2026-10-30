"""
make_dashboard.py — Build the self-contained HTML dashboard.

Reads output/comparison_data.json and renders dashboard/index.html —
a single self-contained page (Chart.js via CDN) that can be hosted
on GitHub Pages or opened locally. No build server needed.

Run:  python make_dashboard.py
"""

import json
from pathlib import Path

BASE = Path(__file__).resolve().parent
DATA_PATH = BASE / "output" / "comparison_data.json"
OUT_PATH = BASE / "dashboard" / "index.html"

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>10-K Sentiment & Financial Analysis — AAPL vs INTC</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {
    --navy: #1F4E79; --red: #C0504D; --green: #9BBB59;
    --ink: #222; --muted: #667; --bg: #f6f8fa; --card: #fff;
  }
  * { box-sizing: border-box; }
  body {
    font-family: Georgia, 'Times New Roman', serif;
    margin: 0; background: var(--bg); color: var(--ink);
  }
  header {
    background: var(--navy); color: #fff; padding: 28px 20px;
  }
  header h1 { margin: 0 0 6px; font-size: 26px; }
  header p { margin: 0; opacity: .85; font-size: 14px; }
  main { max-width: 1080px; margin: 0 auto; padding: 20px; }
  .kpi-row { display: flex; gap: 14px; flex-wrap: wrap; margin: 18px 0; }
  .kpi {
    flex: 1 1 200px; background: var(--card); border-radius: 8px;
    padding: 14px 16px; box-shadow: 0 1px 4px rgba(0,0,0,.08);
  }
  .kpi .label { font-size: 12px; color: var(--muted); text-transform: uppercase; letter-spacing: .04em; }
  .kpi .value { font-size: 24px; font-weight: bold; margin-top: 2px; }
  .kpi .sub { font-size: 12px; color: var(--muted); }
  .card {
    background: var(--card); border-radius: 8px; padding: 18px;
    margin: 16px 0; box-shadow: 0 1px 4px rgba(0,0,0,.08);
  }
  .card h2 { margin: 0 0 4px; font-size: 19px; color: var(--navy); }
  .card p.note { margin: 0 0 12px; font-size: 13px; color: var(--muted); }
  .ct { margin: 0 0 6px; font-size: 15px; color: var(--navy); }
  .chart-box { position: relative; height: 340px; }
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .swipe-hint { display: none; font-size: 12px; color: var(--muted); margin: 0 0 6px; }
  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  th, td { border: 1px solid #e0e4e8; padding: 6px 8px; text-align: right; }
  th { background: var(--navy); color: #fff; text-align: right; }
  th:first-child, td:first-child { text-align: left; }
  tr:nth-child(even) td { background: #fafbfc; }
  @media (max-width: 800px) {
    .grid2 { grid-template-columns: 1fr; }
    .swipe-hint { display: block; }
    table { font-size: 11px; }
    th, td { padding: 4px 6px; }
  }
  footer {
    padding: 22px 20px; color: var(--muted); font-size: 12px;
    max-width: 1080px; margin: 0 auto;
  }
  footer a { color: var(--navy); }
</style>
</head>
<body>
<header>
  <h1>SEC 10-K Sentiment & Financial Analysis</h1>
  <p>Apple (AAPL) vs Intel (INTC) · Fiscal years 2016–2025 ·
     MD&A scored with the Loughran-McDonald Master Dictionary</p>
</header>
<main>

<div class="kpi-row" id="kpis"></div>

<div class="card">
  <h2>MD&A sentiment vs stock price</h2>
  <p class="note">Positive/negative word ratio from each year's MD&A
    (left axis) against the stock price at each balance-sheet date
    (right axis).</p>
  <div class="grid2">
    <div><h3 class="ct">Apple (AAPL)</h3><div class="chart-box"><canvas id="chartAapl"></canvas></div></div>
    <div><h3 class="ct">Intel (INTC)</h3><div class="chart-box"><canvas id="chartIntc"></canvas></div></div>
  </div>
</div>

<div class="card">
  <h2>Financial ratios</h2>
  <p class="note">Liquidity, solvency, efficiency and profitability —
    one ratio each, computed from SEC XBRL data (10-K annual values).</p>
  <div class="grid2">
    <div class="chart-box"><canvas id="chartLiq"></canvas></div>
    <div class="chart-box"><canvas id="chartSolv"></canvas></div>
    <div class="chart-box"><canvas id="chartEff"></canvas></div>
    <div class="chart-box"><canvas id="chartProf"></canvas></div>
  </div>
</div>

<div class="card">
  <h2>Strategy signals from the 10-Ks</h2>
  <p class="note">R&D intensity, gross margin, and capital
    allocation (capex / buybacks / dividends) extracted from the
    filings.</p>
  <div class="grid2">
    <div class="chart-box"><canvas id="chartRnd"></canvas></div>
    <div class="chart-box"><canvas id="chartMargin"></canvas></div>
  </div>
  <div class="chart-box" style="height:360px"><canvas id="chartCapital"></canvas></div>
</div>

<div class="card">
  <h2>Sentiment detail</h2>
  <p class="note">Loughran-McDonald category counts as a share of
    total MD&A words.</p>
  <div class="grid2">
    <div class="chart-box"><canvas id="chartSentAapl"></canvas></div>
    <div class="chart-box"><canvas id="chartSentIntc"></canvas></div>
  </div>
  <div style="overflow-x:auto">
    <p class="swipe-hint">Swipe the table sideways to see all columns →</p>
    <table id="sentimentTable"></table>
  </div>
</div>

</main>
<footer>
  <p>Built from SEC EDGAR filings (10-K), the Loughran-McDonald Master
  Dictionary 1993-2025 (SRAF, Notre Dame — free for academic research),
  and Yahoo Finance daily closes. Sources:
  <a href="https://www.sec.gov/edgar/search/">SEC EDGAR</a> ·
  <a href="https://sraf.nd.edu/loughranmcdonald-master-dictionary/">SRAF Master Dictionary</a> ·
  <a href="https://finance.yahoo.com/">Yahoo Finance</a>.
  MD&A length varies by filer and year; compare percentage shares,
  not raw counts. Stock prices are split-adjusted closes on or before
  each balance-sheet date.</p>
</footer>
<script>
const DATA = __DATA__;

function years(t) { return DATA.companies[t].map(r => r["Fiscal Year"]); }
function col(t, name) { return DATA.companies[t].map(r => r[name]); }

function dualAxisChart(id, ticker, label) {
  new Chart(document.getElementById(id), {
    type: "line",
    data: { labels: years(ticker), datasets: [
      { label: "P/N ratio", data: col(ticker, "P/N Ratio"),
        borderColor: "#1F4E79", backgroundColor: "#1F4E79",
        yAxisID: "y", tension: .25 },
      { label: "Stock price ($)", data: col(ticker, "Closing Price ($)"),
        borderColor: "#C0504D", backgroundColor: "#C0504D",
        yAxisID: "y1", tension: .25 },
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { title: { display: true, text: label } },
      scales: {
        y: { position: "left", title: { display: true, text: "P/N ratio" }},
        y1: { position: "right", title: { display: true, text: "$" },
              grid: { drawOnChartArea: false }},
      },
    },
  });
}

function compareChart(id, columnName, title, yAxisLabel, scale) {
  const s = (scale === undefined) ? 1 : scale;
  new Chart(document.getElementById(id), {
    type: "line",
    data: { labels: years("AAPL"), datasets: [
      { label: "AAPL", data: col("AAPL", columnName).map(v => v === null ? null : v * s),
        borderColor: "#1F4E79", tension: .25 },
      { label: "INTC", data: col("INTC", columnName).map(v => v === null ? null : v * s),
        borderColor: "#C0504D", tension: .25 },
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { title: { display: true, text: title } },
      scales: { y: { title: { display: true, text: yAxisLabel } } },
    },
  });
}

function kpi(label, value, sub) {
  return '<div class="kpi"><div class="label">' + label +
         '</div><div class="value">' + value +
         '</div><div class="sub">' + sub + '</div></div>';
}

function buildKpis() {
  const a = DATA.companies["AAPL"], i = DATA.companies["INTC"];
  const fmtB = v => "$" + (v / 1e9).toFixed(1) + "B";
  const priceChg = c => {
    const first = c[0]["Closing Price ($)"], last = c[c.length-1]["Closing Price ($)"];
    const v = (last / first - 1) * 100;
    return (v >= 0 ? "+" : "") + v.toFixed(1) + "%";
  };
  const revChg = c => {
    const first = c[0]["Revenue ($)"], last = c[c.length-1]["Revenue ($)"];
    const v = (last / first - 1) * 100;
    return (v >= 0 ? "+" : "") + v.toFixed(1) + "%";
  };
  document.getElementById("kpis").innerHTML =
    kpi("AAPL revenue", revChg(a), "FY2016 → FY2025") +
    kpi("INTC revenue", revChg(i), "FY2016 → FY2025") +
    kpi("AAPL stock", priceChg(a), "FY2016 → FY2025 close") +
    kpi("INTC stock", priceChg(i), "FY2016 → FY2025 close") +
    kpi("AAPL net margin FY2025",
        (a[a.length-1]["Net Profit Margin"] * 100).toFixed(1) + "%",
        "FY2016: " + (a[0]["Net Profit Margin"] * 100).toFixed(1) + "%") +
    kpi("INTC net margin FY2025",
        (i[i.length-1]["Net Profit Margin"] * 100).toFixed(1) + "%",
        "FY2016: " + (i[0]["Net Profit Margin"] * 100).toFixed(1) + "%");
}

function buildTable() {
  const rows = ["Fiscal Year", "Total Words", "Negative %", "Positive %",
                "P/N Ratio", "Closing Price ($)"];
  let html = "<tr>" + rows.map(r => "<th>" + r + "</th><th></th>").join("") + "</tr>";
  html += "<tr>" + rows.map(() => "<th>AAPL</th><th>INTC</th>").join("") + "</tr>";
  for (let k = 0; k < DATA.companies["AAPL"].length; k++) {
    html += "<tr>";
    for (const name of rows) {
      const av = DATA.companies["AAPL"][k][name];
      const iv = DATA.companies["INTC"][k][name];
      const f = v => (v === null || v === undefined) ? "—" :
        (typeof v === "number" ? (v > 1000 ? v.toLocaleString(undefined,
          {maximumFractionDigits: 0}) : v.toFixed(2)) : v);
      html += "<td>" + f(av) + "</td><td>" + f(iv) + "</td>";
    }
    html += "</tr>";
  }
  document.getElementById("sentimentTable").innerHTML = html;
}

function buildCapitalChart() {
  new Chart(document.getElementById("chartCapital"), {
    type: "bar",
    data: { labels: years("AAPL"), datasets: [
      { label: "AAPL capex ($B)", data: col("AAPL","CapEx").map(v => v===null?0:v/1e9), backgroundColor: "#7d9cc4" },
      { label: "AAPL buybacks ($B)", data: col("AAPL","Buybacks").map(v => v===null?0:v/1e9), backgroundColor: "#9BBB59" },
      { label: "INTC capex ($B)", data: col("INTC","CapEx").map(v => v===null?0:v/1e9), backgroundColor: "#d9a5a3" },
      { label: "INTC buybacks ($B)", data: col("INTC","Buybacks").map(v => v===null?0:v/1e9), backgroundColor: "#C0504D" },
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { title: { display: true, text: "Capital allocation: capex vs buybacks ($B)" } },
      scales: { y: { title: { display: true, text: "$B" } } },
    },
  });
}

function buildSentimentShares() {
  const cats = ["Negative %", "Positive %", "Uncertainty %", "Litigious %"];
  for (const [id, ticker] of [["chartSentAapl","AAPL"], ["chartSentIntc","INTC"]]) {
    new Chart(document.getElementById(id), {
      type: "line",
      data: { labels: years(ticker), datasets: cats.map((c, idx) => ({
        label: c, data: col(ticker, c),
        borderColor: ["#C0504D", "#9BBB59", "#8064A2", "#4F81BD"][idx],
        tension: .25,
      }))},
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { title: { display: true, text: ticker + " — LM categories (% of MD&A words)" } },
        scales: { y: { title: { display: true, text: "% of words" } } },
      },
    });
  }
}

buildKpis();
dualAxisChart("chartAapl", "AAPL", "AAPL: sentiment vs stock price");
dualAxisChart("chartIntc", "INTC", "INTC: sentiment vs stock price");
compareChart("chartLiq", "Current Ratio", "Liquidity — current ratio", "x");
compareChart("chartSolv", "Liabilities-to-Equity", "Solvency — liabilities to equity", "x");
compareChart("chartEff", "Asset Turnover", "Efficiency — asset turnover", "x");
compareChart("chartProf", "Net Profit Margin", "Profitability — net profit margin", "%", 100);
compareChart("chartRnd", "R&D Intensity", "R&D intensity", "% of revenue", 100);
compareChart("chartMargin", "Gross Margin", "Gross margin", "% of revenue", 100);
buildCapitalChart();
buildSentimentShares();
buildTable();
</script>
</body>
</html>
"""


def main() -> None:
    """Inject the comparison dataset into the template and write it."""
    dataset = json.loads(DATA_PATH.read_text())
    OUT_PATH.parent.mkdir(exist_ok=True)
    html = TEMPLATE.replace(
        "__DATA__", json.dumps(dataset, indent=1)
    )
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"dashboard written to {OUT_PATH}")


if __name__ == "__main__":
    main()
