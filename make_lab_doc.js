/**
 * make_lab_doc.js — Generate the Word lab document.
 *
 * Builds "Lab_Document_10K_Sentiment_Analysis.docx" from the verified
 * pipeline outputs. All figures are read from output/figures/.
 *
 * Run:  node make_lab_doc.js
 */

const fs = require("fs");
const path = require("path");
const docx = require("docx");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  ExternalHyperlink, TableOfContents, BorderStyle, WidthType, TabStopType, LeaderType,
  ShadingType, PageNumber, PageBreak, FootnoteReferenceRun,
} = docx;

const BASE = __dirname;
const PROMPT1_FULL = require("fs").readFileSync(BASE + "/user_prompt_1.txt", "utf8").trim();
const FIG = (name) => path.join(BASE, "output", "figures", name);

// ---------------------------------------------------------------- palette
const NAVY = "1F4E79";
const LIGHT = "EDF2F7";
const MUTED = "666666";

// ---------------------------------------------------------------- helpers
const P = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: 140, line: 300 },
    alignment: AlignmentType.JUSTIFIED,
    ...opts,
    children: [new TextRun({ text, ...opts.run })],
  });

const H1 = (text) =>
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text })] });
const H2 = (text) =>
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text })] });

const B = (text) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80, line: 300 },
    children: [new TextRun({ text })],
  });

const B2 = (text) =>
  new Paragraph({
    numbering: { reference: "steps", level: 0 },
    spacing: { after: 80, line: 300 },
    children: [new TextRun({ text })],
  });

const thin = { style: BorderStyle.SINGLE, size: 1, color: "B0B0B0" };
const allB = { top: thin, bottom: thin, left: thin, right: thin };
const cellMargins = { top: 50, bottom: 50, left: 90, right: 90 };

function cell(text, width, opts = {}) {
  return new TableCell({
    borders: allB,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    children: [
      new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text), bold: !!opts.bold, size: opts.size || 20 })],
      }),
    ],
  });
}

function table(headers, rows, widths, opts = {}) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      cell(h, widths[i], { bold: true, fill: NAVY, align: opts.aligns ? opts.aligns[i] : undefined, size: 20 })
    ),
  });
  // header text should be white
  const bodyRows = rows.map((r, ri) =>
    new TableRow({
      children: r.map((c, i) =>
        cell(c, widths[i], {
          align: opts.aligns ? opts.aligns[i] : undefined,
          fill: ri % 2 === 1 ? "F6F8FA" : undefined,
        })
      ),
    })
  );
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...bodyRows],
  });
}

function figure(name, caption, widthPx) {
  const file = FIG(name);
  const dims = require("image-size").imageSize ? require("image-size").imageSize(file) : { width: 1350, height: 675 };
  const w = widthPx || 620;
  const h = Math.round((dims.height / dims.width) * w);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 60 },
      children: [
        new ImageRun({
          type: "png",
          data: fs.readFileSync(file),
          transformation: { width: w, height: h },
          altText: { title: caption, description: caption, name: name },
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: caption, italics: true, color: MUTED, size: 18 })],
    }),
  ];
}

const CITE = (n) => new FootnoteReferenceRun(n);

// ---------------------------------------------------------------- data
const AAPL_ROWS = [
  ["FY2016", "8,356", "1.28", "0.30", "0.234", "28.18"],
  ["FY2017", "8,556", "1.29", "0.29", "0.227", "38.53"],
  ["FY2018", "8,182", "1.03", "0.43", "0.417", "56.44"],
  ["FY2019", "4,718", "1.00", "0.19", "0.191", "54.71"],
  ["FY2020", "4,465", "1.25", "0.13", "0.107", "112.28"],
  ["FY2021", "2,610", "0.54", "0.50", "0.929", "146.92"],
  ["FY2022", "2,425", "1.07", "0.25", "0.231", "150.43"],
  ["FY2023", "2,139", "0.75", "0.05", "0.062", "171.21"],
  ["FY2024", "2,083", "0.82", "0.14", "0.176", "227.79"],
  ["FY2025", "2,451", "0.82", "0.20", "0.250", "255.46"],
];

const INTC_ROWS = [
  ["FY2016", "7,039", "0.80", "0.75", "0.946", "36.27"],
  ["FY2017", "6,466", "0.73", "1.22", "1.681", "46.16"],
  ["FY2018", "8,035", "0.90", "1.62", "1.806", "46.75"],
  ["FY2019", "10,102", "0.73", "1.75", "2.392", "60.08"],
  ["FY2020", "11,037", "0.89", "1.69", "1.908", "47.07"],
  ["FY2021", "13,132", "0.91", "1.47", "1.622", "51.31"],
  ["FY2022", "13,946", "1.03", "1.30", "1.266", "26.43"],
  ["FY2023", "28,284", "2.90", "1.02", "0.351", "50.25"],
  ["FY2024", "25,948", "3.14", "1.04", "0.331", "20.30"],
  ["FY2025", "7,886", "1.94", "0.91", "0.471", "36.20"],
];

const AAPL_RATIO_ROWS = [
  ["FY2016", "215.6", "45.7", "39.1", "1.35", "1.51", "0.70", "21.2"],
  ["FY2017", "229.2", "48.4", "38.5", "1.28", "1.80", "0.66", "21.1"],
  ["FY2018", "265.6", "59.5", "38.3", "1.13", "2.41", "0.72", "22.4"],
  ["FY2019", "260.2", "55.3", "37.8", "1.54", "2.74", "0.74", "21.2"],
  ["FY2020", "274.5", "57.4", "38.2", "1.36", "3.96", "0.83", "20.9"],
  ["FY2021", "365.8", "94.7", "41.8", "1.07", "4.56", "1.08", "25.9"],
  ["FY2022", "394.3", "99.8", "43.3", "0.88", "5.96", "1.12", "25.3"],
  ["FY2023", "383.3", "97.0", "44.1", "0.99", "4.67", "1.09", "25.3"],
  ["FY2024", "391.0", "93.7", "46.2", "0.87", "5.41", "1.09", "24.0"],
  ["FY2025", "416.2", "112.0", "46.9", "0.89", "3.87", "1.15", "26.9"],
];

const INTC_RATIO_ROWS = [
  ["FY2016", "59.4", "10.3", "61.0", "1.75", "0.71", "0.55", "17.4"],
  ["FY2017", "62.8", "9.6", "62.3", "1.69", "0.77", "0.53", "15.3"],
  ["FY2018", "70.8", "21.1", "61.7", "1.73", "0.72", "0.56", "29.7"],
  ["FY2019", "72.0", "21.0", "58.6", "1.40", "0.76", "0.54", "29.2"],
  ["FY2020", "77.9", "20.9", "56.0", "1.91", "0.89", "0.54", "26.8"],
  ["FY2021", "79.0", "19.9", "55.4", "2.13", "0.77", "0.49", "25.1"],
  ["FY2022", "63.1", "8.0", "42.6", "1.57", "0.76", "0.36", "12.7"],
  ["FY2023", "54.2", "1.7", "40.0", "1.54", "0.74", "0.29", "3.1"],
  ["FY2024", "53.1", "-18.8", "32.7", "1.33", "0.87", "0.27", "-35.3"],
  ["FY2025", "52.9", "-0.3", "34.8", "2.02", "0.67", "0.26", "-0.5"],
];

const RATIO_HEADERS = [
  "Fiscal Year", "Revenue ($B)", "Net Income ($B)", "Gross Margin (%)",
  "Current Ratio", "Liab./Equity", "Asset Turnover", "Net Margin (%)",
];

const SENT_HEADERS = [
  "Fiscal Year", "MD&A Words", "Negative %", "Positive %", "P/N Ratio", "Stock Price ($)",
];

const RND_ROWS = [
  ["FY2016", "10.0", "4.6", "12.7", "21.4", "215.6", "59.4"],
  ["FY2021", "21.9", "6.0", "15.2", "19.3", "365.8", "79.0"],
  ["FY2024", "31.4", "8.0", "16.5", "31.1", "391.0", "53.1"],
  ["FY2025", "34.5", "8.3", "13.8", "26.1", "416.2", "52.9"],
];

const CAP_ROWS = [
  ["FY2016", "12.7", "29.7", "12.2", "9.6", "2.6", "4.9"],
  ["FY2021", "11.1", "86.0", "14.5", "18.7", "2.4", "5.6"],
  ["FY2023", "11.0", "77.5", "15.0", "25.8", "0.0", "3.1"],
  ["FY2025", "12.7", "90.7", "15.4", "14.6", "0.0", "0.0"],
];

// ---------------------------------------------------------------- content
const children = [];

// ----------------------------------------------------------- title page
children.push(
  new Paragraph({ spacing: { before: 2400, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "SEC 10-K Sentiment and Financial Analysis", bold: true, size: 52, color: NAVY })] }),
  new Paragraph({ spacing: { before: 120, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "A Gen AI “Vibe Coding” Final Project", size: 28, color: MUTED })] }),
  new Paragraph({ spacing: { before: 480, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Apple Inc. (AAPL) and Intel Corporation (INTC)", bold: true, size: 26 })] }),
  new Paragraph({ spacing: { before: 80, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Fiscal Years 2016–2025", size: 26 })] }),
  new Paragraph({ spacing: { before: 2400, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Amirtha Varshini V", bold: true, size: 24 })] }),
  new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "University of New Haven — Finance", size: 22 })] }),
  new Paragraph({ spacing: { before: 60, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "October 2026", size: 22 })] }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ---------------------------------------------------------------- TOC
// Static table of contents with page numbers taken from the rendered
// document (LibreOffice PDF render; Carlito is metric-compatible with
// Calibri, so Microsoft Word pagination matches). Kept static so the
// contents are populated in every viewer without field updates.
const TOC_ENTRIES = [
  [1, "1. Objective and Scope", "3"],
  [1, "2. Requirement Checklist", "3"],
  [1, "3. What Is Vibe Coding, and How AI Was Used", "4"],
  [2, "3.1 The approach", "4"],
  [2, "3.2 AI models compared", "4"],
  [2, "3.3 Human oversight and verification", "5"],
  [1, "4. Methodology", "5"],
  [2, "4.1 Data sources", "5"],
  [2, "4.2 Pipeline architecture", "5"],
  [2, "4.3 MD&A extraction and boundary verification", "6"],
  [2, "4.4 Sentiment scoring", "6"],
  [2, "4.5 Ratio definitions", "6"],
  [2, "4.6 Limitations", "7"],
  [1, "5. Results: Apple (AAPL)", "8"],
  [2, "5.1 Apple data", "9"],
  [1, "6. Results: Intel (INTC)", "10"],
  [2, "6.1 Intel data", "11"],
  [1, "7. Company Comparison and Ten-Year Relationships", "12"],
  [2, "7.1 Relationships observed over the ten years", "12"],
  [1, "8. Graduate Analysis Questions (3)", "14"],
  [2, "Question 1. R&D intensity and revenue growth — did each company’s R&D spending buy growth?", "14"],
  [2, "Question 2. Gross margin trajectory — what drove the FY2022 crossover, in management’s own words?", "15"],
  [2, "Question 3. Capital allocation — what does the cash-flow statement reveal about strategy and risk?", "15"],
  [1, "9. How to Reproduce This Project — Step-by-Step Guide", "17"],
  [2, "9.1 One-time setup", "17"],
  [2, "9.2 Run the pipeline", "17"],
  [2, "9.3 Expected outputs", "17"],
  [2, "9.4 Host the HTML dashboard on GitHub Pages", "18"],
  [2, "9.5 Common errors and their fixes (all encountered during development)", "18"],
  [1, "10. Deliverables Index", "18"],
  [1, "References", "19"],
  [1, "Appendix A. AI Prompts Used (Vibe-Coding History)", "20"],
];
children.push(
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "Table of Contents" })] }),
  ...TOC_ENTRIES.map(([lvl, t, pg]) =>
    new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: 9360, leader: LeaderType.DOT }],
      indent: { left: lvl === 2 ? 360 : 0 },
      spacing: { after: lvl === 1 ? 60 : 30, line: 276 },
      children: [
        new TextRun({ text: t, bold: lvl === 1, size: lvl === 1 ? 22 : 20 }),
        new TextRun({ text: "\t" + pg, size: lvl === 1 ? 22 : 20 }),
      ],
    })
  ),
  new Paragraph({ children: [new PageBreak()] }),
);

// ------------------------------------------------------------- section 1
children.push(
  H1("1. Objective and Scope"),
  P(
    "This project demonstrates how an AI coding assistant can generate, test, and refine a complete Python research pipeline for SEC filings analysis. The program takes a stock ticker and a filing count as inputs, downloads that company’s Form 10-K annual reports from the SEC’s EDGAR database, extracts the Management’s Discussion and Analysis (MD&A) narrative and the three core financial statements, scores the MD&A text with the Loughran-McDonald (LM) financial sentiment dictionary, pulls the stock price at each balance-sheet date, computes the four analyst ratios (liquidity, solvency, efficiency, profitability), and writes everything into a formatted Excel workbook with a chart comparing the stock price against the MD&A positive/negative (P/N) ratio."
  ),
  P(
    "Two companies were analyzed for the ten fiscal years 2016–2025: Apple Inc. (AAPL), a dividend-paying consumer-technology leader, and Intel Corporation (INTC), a semiconductor manufacturer that underwent a major strategic transformation during the decade. Apple and Intel make an instructive pair: they were both Dow Jones constituents over the whole window, they share a long competitive history in personal computing, and they took opposite financial trajectories over the decade — Apple nearly doubled revenue while Intel’s business contracted."
  ),
  P(
    "The deliverables follow the updated assignment instructions: this lab document, a zipped desktop application (Tkinter GUI wrapping the pipeline), and an HTML dashboard ready to host on GitHub Pages. As a graduate student, three additional analysis questions are answered in Section 8, all answerable directly from the downloaded 10-Ks."
  ),
);

// ------------------------------------------------------------- section 2
children.push(
  H1("2. Requirement Checklist"),
  P(
    "Every numbered requirement from the assignment instructions is mapped below to the artifact that satisfies it, so that each item can be verified directly."
  ),
  table(
    ["#", "Requirement (from Instructions)", "Where it is satisfied"],
    [
      ["1", "User enters ticker and number of 10-Ks; query EDGAR and identify 10-K filings from the past ten years", "tenk_analyzer.py CLI (python tenk_analyzer.py AAPL 10); desktop app input form"],
      ["2", "Download, organize, and label the files with appropriate names", "filings/{TICKER}/{TICKER}_10K_FY{year}_{filed}.htm; MD&A saved as {TICKER}_MD_FY{year}_{filed}.txt"],
      ["3", "Extract MD&A into a text file; extract IS, BS, CF into one Excel workbook, side by side, one tab each", "Workbook tabs “Income Statement”, “Balance Sheet”, “Cash Flows” — all periods side by side"],
      ["4", "Clean the text so it is readable for the LM dictionary", "clean_mda_text() — lowercase, punctuation/number removal; cleaned files saved (*_cleaned.txt)"],
      ["5", "Score the MD&A with the LM dictionary; output to a new tab", "“Sentiment Analysis” tab — word counts, category counts and shares, P/N ratio"],
      ["6", "Stock price at each balance-sheet date in a separate tab", "“Stock Prices” tab — close on or before each balance-sheet date"],
      ["7", "Four ratios (liquidity, solvency, efficiency, profitability) for each of ten years", "“Ratios” tab — current ratio, liabilities-to-equity, asset turnover, net profit margin, FY2016–FY2025"],
      ["8", "Sentiment analysis with total word count, total positive/negative, P/N ratio", "“Sentiment Analysis” tab (counts, percentages, and ratio)"],
      ["9", "Chart comparing stock price and P/N ratios for all years", "“Chart” tab — native Excel combined line chart; also Figures 1–2 in this document"],
      ["10", "Lab document for AAPL and one other company; comment on relationships over ten years; compare the companies", "This document — Sections 5–7"],
      ["11", "Additional analyses (graduate: three questions)", "Section 8"],
      ["12", "All user prompts at the end of the lab document", "Appendix A"],
    ],
    [400, 4200, 4760]
  ),
);

// ------------------------------------------------------------- section 3
children.push(
  H1("3. What Is Vibe Coding, and How AI Was Used"),
  H2("3.1 The approach"),
  P(
    "Vibe coding means describing the desired outcome in natural language and letting an AI assistant generate the code, then reviewing, testing, and refining it in an iterative conversation. No code in this project was written line-by-line by hand: the assistant produced the drafts, and each draft was executed against real SEC data, with failures fed back as new prompts. The human role shifted from typing syntax to specifying requirements, reviewing logic, and validating results — which is precisely the collaboration model the assignment emphasizes."
  ),
  P(
    "The development proceeded in stages, one pipeline requirement at a time: (1) resolve ticker to CIK and list 10-K filings; (2) download and label filings; (3) extract the MD&A section; (4) parse the financial statements; (5) clean the MD&A text; (6) score with the LM dictionary; (7) compute ratios; (8) fetch stock prices; (9) build the Excel workbook and chart. Each stage was tested against Apple first, then Intel, and the differences between the two companies’ filing formats drove several important refinements (Section 4.3). The full prompt history is reproduced in Appendix A."
  ),
  H2("3.2 AI models compared"),
  P(
    "The assignment suggests Google Antigravity; several assistants were evaluated for the roles they play in a project like this. The project itself was built with Perplexity Computer, and the prompts in Appendix A are directly reusable in any of these tools."
  ),
  table(
    ["Model", "Best role in this project", "Key strengths", "Watch-outs"],
    [
      ["Google Antigravity", "Agent-first IDE; end-to-end project builds", "Browser-based agent, free tier, strong on multi-file edits", "Newer tool; verify generated code against sources"],
      ["ChatGPT (GPT)", "Explaining concepts, drafting code and doc text", "Strong reasoning, easy iteration", "Can invent API details — always run the code"],
      ["Claude (Anthropic)", "Long code files, careful refactors", "Very good at large-context editing and documentation", "Same verification duty as above"],
      ["Gemini (Google)", "Multimodal research, Google-integrated lookups", "Large context window", "Check EDGAR specifics against the source"],
      ["GitHub Copilot", "Inline completion inside an editor", "Fast autocomplete while coding", "Weaker at whole-project planning"],
      ["Cursor", "AI-native editor for iterating on a repo", "Codebase-aware edits", "Needs local setup"],
      ["Perplexity", "Research + working code with live data", "Web-connected, executes and tests code in a sandbox", "Review outputs like any assistant"],
    ],
    [1500, 2200, 2800, 2860]
  ),
  H2("3.3 Human oversight and verification"),
  P(
    "Three verification layers keep the AI-generated pipeline honest. First, an automated boundary audit (audit_mda.py) records exactly where each MD&A extraction starts and ends in each filing, with excerpts for manual inspection. Second, a reconciliation script (verify_outputs.py) checks completeness (10 years on every tab), the balance-sheet identity (assets = liabilities + equity), statement-versus-XBRL agreement for revenue and net income, and the P/N ratio arithmetic. Third, headline figures were spot-checked against the filings themselves — for example, Apple’s FY2016 revenue of $215.6B and Intel’s FY2024 net loss of $18.8B both tie to the income statements in the workbooks."
  ),
);

// ------------------------------------------------------------- section 4
children.push(
  H1("4. Methodology"),
  H2("4.1 Data sources"),
  B("SEC EDGAR — filing index via the submissions JSON API; primary 10-K documents; “R-file” financial-statement renderings; XBRL company-concept API for ratio inputs."),
  B("Loughran-McDonald Master Dictionary, 1993–2025 edition (SRAF, University of Notre Dame) — 86,553 words with category flags (2,345 negative, 347 positive, and five further categories); free for academic research."),
  B("Yahoo Finance chart API — daily split-adjusted closing prices."),
  H2("4.2 Pipeline architecture"),
  table(
    ["Stage", "Function", "Output"],
    [
      ["1. Identify", "get_cik() / get_10k_filings()", "Ticker → CIK → list of the 10 most recent 10-Ks (accession numbers, filing dates, balance-sheet dates)"],
      ["2. Download", "download_filings()", "Labeled 10-K HTML files in filings/{TICKER}/"],
      ["3. MD&A", "extract_mda()", "Raw MD&A text files in mdna_text/"],
      ["4. Statements", "extract_statements() / combine_statements()", "IS / BS / CF tables, all periods side by side"],
      ["5. Clean", "clean_mda_text()", "Cleaned text (*_cleaned.txt) — lowercase, numbers/punctuation removed"],
      ["6. Score", "score_mda()", "LM category counts, shares of total words, P/N ratio"],
      ["7. Ratios", "fetch_fundamentals() / compute_ratios()", "Current ratio, liabilities-to-equity, asset turnover, net margin (XBRL inputs)"],
      ["8. Prices", "fetch_stock_prices()", "Close on or before each balance-sheet date"],
      ["9. Workbook", "build_workbook()", "Excel file with 8 tabs including the combined chart"],
    ],
    [1400, 2600, 5360]
  ),
  H2("4.3 MD&A extraction and boundary verification"),
  P(
    "Two heading styles appear in the two companies’ filings. Apple labels the section “Item 7. Management’s Discussion and Analysis …” and the code cuts the text at the “Item 7A” heading; the table-of-contents occurrence of the same heading is rejected automatically because it yields only a few words. Intel, in contrast, labels the section simply “Management’s Discussion and Analysis” with the item number only in the index, so a fallback locator handles standalone headings. In both cases the terminator is the market-risk (Item 7A) heading, never an inline cross-reference such as “Item 8 of this Form 10-K,” which appears inside the MD&A narrative itself. The audit script (audit_mda.py) documents the start position, end position, terminator type, and opening/closing excerpt for all twenty filings; all twenty extractions end at a genuine section boundary."
  ),
  H2("4.4 Sentiment scoring"),
  P(
    "The MD&A text is lowercased, stripped of punctuation and number-only tokens, and each remaining word is looked up in the LM master dictionary. For each year the workbook reports total word count; raw counts for seven categories (negative, positive, uncertainty, litigious, strong modal, weak modal, constraining); each category as a percentage of total MD&A words; and the P/N ratio (positive count ÷ negative count). The P/N ratio is the assignment’s headline sentiment measure."
  ),
  H2("4.5 Ratio definitions"),
  table(
    ["Ratio", "Definition", "Notes"],
    [
      ["Liquidity — current ratio", "Current assets ÷ current liabilities", "us-gaap AssetsCurrent / LiabilitiesCurrent"],
      ["Solvency — liabilities-to-equity", "Total liabilities ÷ total equity", "Reported Liabilities tag when available (Apple); otherwise total assets minus total equity so the balance-sheet identity holds (Intel reports no standalone liabilities subtotal)"],
      ["Efficiency — asset turnover", "Revenue ÷ average total assets", "Average of beginning and ending assets; ending assets for the first year"],
      ["Profitability — net profit margin", "Net income ÷ revenue", "us-gaap NetIncomeLoss / revenue"],
    ],
    [2200, 2400, 4760]
  ),
  H2("4.6 Limitations"),
  B("MD&A length varies enormously across filers and years: Apple’s MD&A shrank from 8,356 words (FY2016) to 2,451 (FY2025) as the company deliberately shortened the section, while Intel’s ballooned to 28,284 words in FY2023–24 when it folded the business overview into MD&A. Category percentages and the P/N ratio control for length; raw counts do not, and the change should be kept in mind when reading the trend breaks."),
  B("Statement-tab line items follow each filer’s own presentation; repeated labels are suffixed (e.g., “Cost of sales (2)”) and rows that exist in only some years simply have empty cells."),
  B("Stock prices are split-adjusted closes, so buy-and-hold returns ignore dividends; the ratio tabs are the appropriate place for fundamental comparisons."),
  B("The dictionary is a general-purpose financial lexicon; it does not know that “gross margin decreased” in one sentence is more informative than a list of risk factors, so tone should corroborate — not replace — the financial statements."),
);

// ------------------------------------------------------------- section 5
children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "5. Results: Apple (AAPL)" })] }),
  P(
    "Apple’s decade was one of compounding growth: revenue rose from $215.6B (FY2016) to $416.2B (FY2025), net income from $45.7B to $112.0B, and the stock price from $28.18 to $255.46 — a nine-fold increase. Every ratio tells the same story of a business that became more profitable even as it grew: gross margin expanded from 39.1% to 46.9% as higher-margin Services revenue became a larger share of the mix, and net margin rose from 21.2% to 26.9%."
  ),
  P(
    "Two balance-sheet trends stand out. First, liquidity deliberately fell: the current ratio declined from 1.35 to 0.89 as Apple ran its balance sheet lean and returned cash aggressively. Second, solvency leverage rose sharply through FY2022 (liabilities-to-equity 5.96) — not because the business was risky, but because Apple issued debt to fund massive share repurchases while equity was being shrunk by those same buybacks. The ratio then eased to 3.87 by FY2025 as earnings rebuilt equity. Efficiency improved almost monotonically (asset turnover 0.70 → 1.15), meaning each dollar of assets generated more sales each year."
  ),
  ...figure("fig_sentiment_price_aapl.png", "Figure 1. AAPL — MD&A sentiment (P/N ratio, left axis) vs stock price (right axis), FY2016–FY2025."),
  P(
    "The sentiment picture is noisier. Apple’s MD&A tone is generally guarded — negative words outnumber positive ones in every year (P/N between 0.06 and 0.93) — which is typical of disclosure drafting, where lawyers highlight risks. Over FY2016–FY2025 the P/N ratio is essentially uncorrelated with the stock price (correlation −0.02): the stock rose nine-fold while tone oscillated. There is, however, a clear mechanical driver of the level shift: Apple’s MD&A shortened by roughly two-thirds over the decade (8,356 → 2,451 words), so year-to-year movements in the P/N ratio increasingly reflect small absolute counts. The FY2021 reading (0.93) and the FY2023 reading (0.06) are both better interpreted as short-document noise than as changes in business conditions; the FY2023 filing also coincided with Apple’s revenue dip (FY2023 revenue fell 2.8%), consistent with a slightly more cautious tone."
  ),
  H2("5.1 Apple data"),
  table(SENT_HEADERS, AAPL_ROWS, [900, 1300, 1100, 1100, 1100, 1400], { aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  new Paragraph({ spacing: { before: 100, after: 200 }, children: [new TextRun({ text: "Table 1. AAPL — MD&A sentiment and stock price by fiscal year.", italics: true, color: MUTED, size: 18 })] }),
  table(RATIO_HEADERS, AAPL_RATIO_ROWS, [900, 1100, 1200, 1200, 1000, 1000, 1100, 1100], { aligns: [AlignmentType.LEFT, ...Array(7).fill(AlignmentType.RIGHT)] }),
  new Paragraph({ spacing: { before: 100, after: 200 }, children: [new TextRun({ text: "Table 2. AAPL — financial data and the four ratios by fiscal year (revenue/net income/gross margin from the 10-K statements; ratios from XBRL inputs).", italics: true, color: MUTED, size: 18 })] }),
);

// ------------------------------------------------------------- section 6
children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "6. Results: Intel (INTC)" })] }),
  P(
    "Intel’s decade is almost a mirror image of Apple’s. Revenue peaked at $79.0B in FY2021 and then fell to $52.9B by FY2025 — below its FY2016 level. The profitability collapse was far steeper: net margin went from +17.4% (FY2016) and a peak of +29.7% (FY2018) to −35.3% in FY2024 — an $18.8B net loss, driven by impairment charges and the rising cost of the foundry build-out. Gross margin, Intel’s hallmark, fell from 61.0% to 34.8%. The stock price went full circle — $36.27 at FY2016 year-end, $20.30 at FY2024, $36.20 at FY2025 — a lost decade for shareholders."
  ),
  P(
    "The balance sheet tells the strategy story. Intel’s liquidity remained solid (current ratio 1.75 → 2.02) and leverage modest (liabilities-to-equity around 0.7–0.9) — this was not a balance-sheet crisis but an income-statement one. Efficiency, however, collapsed: asset turnover fell from 0.55 to 0.26 as billions of dollars of new fabs (construction in progress peaked at $50.4B at FY2024 year-end) sat on the balance sheet generating little revenue yet."
  ),
  ...figure("fig_sentiment_price_intc.png", "Figure 2. INTC — MD&A sentiment (P/N ratio, left axis) vs stock price (right axis), FY2016–FY2025."),
  P(
    "For Intel, the sentiment measure is genuinely informative. The P/N ratio rose from 0.95 (FY2016) to a peak of 2.39 (FY2019) — the years of record margins and share-price strength — and then collapsed to 0.33 by FY2024 as the losses mounted. The correlation between Intel’s P/N ratio and its stock price over the decade is +0.64, and the share of negative words in the MD&A correlates −0.85 with net income. The FY2023–24 filings read dramatically darker (2.9–3.1% negative words versus ~0.8% in the early years): the MD&A language on impairments, accelerated depreciation, and restructuring is exactly the sort of vocabulary the LM dictionary flags. For a company in genuine distress, dictionary sentiment tracked both fundamentals and the stock."
  ),
  H2("6.1 Intel data"),
  table(SENT_HEADERS, INTC_ROWS, [900, 1300, 1100, 1100, 1100, 1400], { aligns: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  new Paragraph({ spacing: { before: 100, after: 200 }, children: [new TextRun({ text: "Table 3. INTC — MD&A sentiment and stock price by fiscal year. The FY2023–24 word counts reflect Intel’s expanded MD&A scope (business overview folded in); percentages are the comparable measure.", italics: true, color: MUTED, size: 18 })] }),
  table(RATIO_HEADERS, INTC_RATIO_ROWS, [900, 1100, 1200, 1200, 1000, 1000, 1100, 1100], { aligns: [AlignmentType.LEFT, ...Array(7).fill(AlignmentType.RIGHT)] }),
  new Paragraph({ spacing: { before: 100, after: 200 }, children: [new TextRun({ text: "Table 4. INTC — financial data and the four ratios by fiscal year.", italics: true, color: MUTED, size: 18 })] }),
);

// ------------------------------------------------------------- section 7
children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "7. Company Comparison and Ten-Year Relationships" })] }),
  ...figure("fig_ratios.png", "Figure 3. AAPL vs INTC — liquidity, solvency, efficiency, and profitability, FY2016–FY2025."),
  P(
    "Across the four ratios the two companies diverged on three of four dimensions. Liquidity: Apple intentionally drained liquidity (1.35 → 0.89) to fund buybacks, while Intel stayed liquid (ending at 2.02). Solvency: Apple’s leverage rose to 5.96 by FY2022 for capital-return reasons, while Intel’s stayed near 0.7–0.9. Efficiency: Apple improved every year (0.70 → 1.15) while Intel halved (0.55 → 0.26). Profitability: Apple’s net margin rose five points while Intel’s fell from +29.7% to negative territory. The single most important comparison is gross margin (Figure 5 in Section 8): the lines cross in FY2022 — Apple’s expanding to 46.9% while Intel’s fell to 34.8%."
  ),
  H2("7.1 Relationships observed over the ten years"),
  B("Sentiment vs fundamentals: For both companies, the share of negative words in the MD&A is strongly negatively correlated with net income (AAPL −0.76, INTC −0.85). When the business deteriorated, the language darkened within the same filing year."),
  B("Sentiment vs stock price: The relationship holds for Intel (P/N vs price correlation +0.64 — tone deteriorated alongside the stock) but vanishes for Apple (−0.02). A plausible reading: dictionary tone is informative when a company is under stress and its MD&A is long and substantive, and nearly meaningless when a filing is very short and consistently risk-averse in tone."),
  B("MD&A length as a confounder: Apple’s shortening MD&A (8,356 → 2,451 words) and Intel’s ballooning MD&A (7,039 → 28,284 words) mean raw counts are not comparable across years; percentage shares and the P/N ratio are the right lenses, and even they should be read with the length context in mind."),
  B("Ratios vs strategy: Apple’s rising liabilities-to-equity was a capital-structure choice (debt-funded buybacks against a shrinking equity base), while Intel’s stable leverage masked a deteriorating income statement — a reminder that no single ratio tells the whole story."),
  B("Efficiency as the cleanest separator: asset turnover moved in opposite directions for the whole decade; Apple got more revenue per asset dollar every year while Intel’s fab investment depressed turnover exactly when margins were also compressing."),
);

// ------------------------------------------------------------- section 8
children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "8. Graduate Analysis Questions (3)" })] }),
  P(
    "Each question below is answered entirely from data extracted from the two companies’ 10-K filings — the income statements, cash-flow statements, and MD&A sections in the generated workbooks."
  ),
  H2("Question 1. R&D intensity and revenue growth — did each company’s R&D spending buy growth?"),
  P(
    "Apple and Intel spent very different amounts of R&D relative to their size, and they got very different results. Apple’s R&D expense grew from $10.0B (FY2016) to $34.5B (FY2025) — but its intensity rose only from 4.6% to 8.3% of revenue, because revenue nearly doubled over the same period. Intel spent a much larger share of every revenue dollar on R&D — 21.4% in FY2016, 31.1% at the FY2024 peak — yet revenue ended the decade 11% lower. Apple’s R&D supported an ecosystem (services, wearables, silicon) that converted research into revenue at rising margins; Intel’s R&D was largely defensive, funding process-node roadmap catch-up (Intel 18A and predecessors) whose payback had not arrived by FY2025. The MD&A language corroborates this: Intel’s FY2024–25 filings describe R&D and process-node spending in terms of rebuilding manufacturing leadership, while Apple’s MD&A describes R&D in terms of new products and services. The lesson for a finance reader: R&D intensity alone is not a growth signal — the marginal revenue per R&D dollar is, and on that measure the decade belongs decisively to Apple."
  ),
  ...figure("fig_rnd.png", "Figure 4. R&D intensity (R&D expense ÷ revenue), FY2016–FY2025."),
  table(
    ["Fiscal Year", "AAPL R&D ($B)", "AAPL Intensity", "INTC R&D ($B)", "INTC Intensity", "AAPL Rev ($B)", "INTC Rev ($B)"],
    RND_ROWS,
    [1100, 1200, 1100, 1200, 1100, 1300, 1300],
    { aligns: [AlignmentType.LEFT, ...Array(6).fill(AlignmentType.RIGHT)] }
  ),
  new Paragraph({ spacing: { before: 100, after: 200 }, children: [new TextRun({ text: "Table 5. R&D expense and revenue (selected years), from the 10-K income statements / XBRL.", italics: true, color: MUTED, size: 18 })] }),
  H2("Question 2. Gross margin trajectory — what drove the FY2022 crossover, in management’s own words?"),
  P(
    "In FY2016 Intel’s gross margin (61.0%) was 22 points above Apple’s (39.1%); by FY2025 Apple’s (46.9%) was 12 points above Intel’s (34.8%). The crossover in FY2022 reflects two independent forces. Apple’s margin expanded with mix: as higher-margin Services revenue grew, consolidated gross margin rose every year from FY2020 onward, and Apple’s MD&A attributes the expansion to the growing Services share of net sales. Intel’s margin collapsed for reasons its own MD&A states plainly. The FY2024 filing: “Our consolidated gross margin in 2024 decreased by $4.4 billion, or 20%, compared to 2023 due primarily to higher 2024 impairment charges and accelerated depreciation.” The FY2025 filing adds that gross margin was further offset by “$878 million of higher 2025 inventory reserves, primarily related to … charges for products manufactured on the early ramp of our Intel 18A process node.” In other words, the margin story is the foundry story: carrying underutilized fabs (impairments, accelerated depreciation, early-ramp inventory charges) directly consumed the gross margin that Intel’s product business still generated."
  ),
  ...figure("fig_margin.png", "Figure 5. Gross margin (% of revenue), FY2016–FY2025."),
  H2("Question 3. Capital allocation — what does the cash-flow statement reveal about strategy and risk?"),
  P(
    "The cash-flow statements show two opposite capital-allocation philosophies. Apple returned essentially all free cash flow to shareholders: buybacks grew from $29.7B (FY2016) to $90.7B (FY2025) — over $700B cumulatively across the decade — plus dividends of roughly $12–15B per year, against modest capex of $7–13B. Apple was betting that its best use of capital was shrinking its own share count. Intel did the reverse: capex roughly doubled, peaking at $25.8B in FY2023 as the foundry build-out accelerated, while shareholder returns went to zero — buybacks stopped after FY2021, the dividend fell from $5.6B to $1.6B, and no dividend was paid in FY2025. Intel was betting the company on new fabs. The risk asymmetry is the takeaway: Apple’s shareholder returns are discretionary and readily reversible, while Intel’s fab investments are multi-year commitments that are difficult to stop once started. A finance reader can see the strategy pivot in the cash-flow statement about three years before it showed up fully in the income statement."
  ),
  ...figure("fig_capital.png", "Figure 6. Capital allocation — capex, buybacks, and dividends vs revenue ($B), FY2016–FY2025."),
  table(
    ["Fiscal Year", "AAPL CapEx", "AAPL Buybacks", "AAPL Dividends", "INTC CapEx", "INTC Buybacks", "INTC Dividends"],
    CAP_ROWS,
    [1100, 1200, 1300, 1300, 1200, 1300, 1300],
    { aligns: [AlignmentType.LEFT, ...Array(6).fill(AlignmentType.RIGHT)] }
  ),
  new Paragraph({ spacing: { before: 100, after: 200 }, children: [new TextRun({ text: "Table 6. Capital allocation ($B, selected years), from the 10-K cash-flow statements / XBRL. Intel reports no dividend cash payments for FY2025.", italics: true, color: MUTED, size: 18 })] }),
);

// ------------------------------------------------------------- section 9
children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "9. How to Reproduce This Project — Step-by-Step Guide" })] }),
  P(
    "These instructions reproduce the entire project from a clean machine. No prior coding experience is assumed beyond installing software."
  ),
  H2("9.1 One-time setup"),
  B2("Install Python 3.10 or newer from python.org (check the box “Add Python to PATH” during installation on Windows)."),
  B2("Install the required packages: open a terminal (Command Prompt / PowerShell / Terminal) and run: pip install -r requirements.txt (or install each: requests, beautifulsoup4, lxml, pandas, openpyxl, matplotlib)."),
  B2("Download the Loughran-McDonald dictionary (~9 MB) by running: python download_dictionary.py. This fetches the official 1993–2025 CSV from the SRAF Google Drive link into data/. (The dictionary is free for academic research; it is deliberately not committed to GitHub.)"),
  B2("Identify yourself to the SEC (their fair-access policy): set the environment variable SEC_USER_AGENT to “Your Name (your-email@example.com)”. Without this, EDGAR may return HTTP 403."),
  H2("9.2 Run the pipeline"),
  B2("Command line: python tenk_analyzer.py AAPL 10 (then python tenk_analyzer.py INTC 10). The two arguments are the ticker and the number of 10-Ks."),
  B2("Desktop app: python tenk_desktop_app.py, type the ticker and count, click Analyze, and watch the log; click “Open output folder” when it finishes."),
  B2("Comparison figures and dashboard data: python build_comparison.py, then python make_dashboard.py."),
  B2("Verification: python verify_outputs.py prints PASS/FAIL for completeness, the balance-sheet identity, statement-vs-XBRL agreement, and the P/N arithmetic. All checks pass for both companies as generated."),
  H2("9.3 Expected outputs"),
  B("output/AAPL_10K_Analysis.xlsx and output/INTC_10K_Analysis.xlsx — eight tabs: README, Income Statement, Balance Sheet, Cash Flows, Sentiment Analysis, Stock Prices, Ratios, Chart."),
  B("filings/ — the labeled 10-K HTML files; mdna_text/ — the extracted and cleaned MD&A text files."),
  B("output/figures/ — the six comparison charts used in this document; dashboard/index.html — the HTML dashboard."),
  H2("9.4 Host the HTML dashboard on GitHub Pages"),
  B2("Create a GitHub account, then create a new public repository (for example, tenk-sentiment-analysis)."),
  B2("Upload dashboard/index.html (plus the project’s Python files, README.md, and requirements.txt). Do not upload the dictionary CSV or the filings — add data/Loughran-McDonald_MasterDictionary_1993-2025.csv and filings/ to a .gitignore file."),
  B2("In the repository, open Settings → Pages, set Source to “Deploy from a branch”, choose main and “/(root)”, and save."),
  B2("After a minute or two the site is live at https://yourusername.github.io/tenk-sentiment-analysis/ — that URL is the “link to the final code” deliverable."),
  H2("9.5 Common errors and their fixes (all encountered during development)"),
  table(
    ["Error", "Cause and fix"],
    [
      ["HTTP 403 from sec.gov", "Missing or generic User-Agent. Set SEC_USER_AGENT to include your name and email; the SEC blocks unidentified traffic."],
      ["HTTP 403/429 mid-run", "EDGAR rate limiting. The code backs off automatically; wait a minute and rerun — the pipeline is idempotent."],
      ["MD&A not found for a filing", "Heading style differs by filer (Apple uses “Item 7 …”, Intel uses a standalone heading). The extractor handles both; see Section 4.3 and audit_mda.py."],
      ["Only 7 of 10 Intel filings found", "Heavy filers push older filings into EDGAR’s archived submission files; the loader follows the archive index automatically."],
      ["Google Drive dictionary download fails", "Open the SRAF page (sraf.nd.edu/loughranmcdonald-master-dictionary) and download the CSV manually into data/."],
      ["pandas applymap errors", "pandas 2.1+ renamed it to .map; the code uses the new API."],
    ],
    [2400, 6960]
  ),
);

// ------------------------------------------------------------- section 10
children.push(
  H1("10. Deliverables Index"),
  table(
    ["Deliverable", "Contents"],
    [
      ["Lab_Document_10K_Sentiment_Analysis.docx", "This document"],
      ["TenK_Sentiment_App.zip", "The desktop application: tenk_desktop_app.py + tenk_analyzer.py + helpers + README + requirements"],
      ["dashboard/index.html", "Self-contained HTML dashboard (Chart.js), ready for GitHub Pages"],
      ["output/AAPL_10K_Analysis.xlsx / INTC_10K_Analysis.xlsx", "The Excel workbooks with all eight tabs"],
      ["output/figures/*.png, output/comparison_data.json", "Comparison charts and the dashboard dataset"],
      ["filings/, mdna_text/", "Downloaded 10-Ks and extracted MD&A text (raw and cleaned)"],
    ],
    [3400, 5960]
  ),
);

// ------------------------------------------------------------- references
children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "References" })] }),
  ...[
    ["U.S. Securities and Exchange Commission, EDGAR company filings and structured data APIs. ", "https://www.sec.gov/edgar/"],
    ["Apple Inc., Form 10-K for fiscal year 2025 (filed October 31, 2025). ", "https://www.sec.gov/Archives/edgar/data/320193/000032019325000079/aapl-20250927.htm"],
    ["Intel Corporation, Form 10-K for fiscal year 2025 (filed January 23, 2026). ", "https://www.sec.gov/Archives/edgar/data/50863/000005086326000011/"],
    ["Intel Corporation, Form 10-K for fiscal year 2024 (filed January 31, 2025). ", "https://www.sec.gov/Archives/edgar/data/50863/000005086325000009/"],
    ["Loughran, T., & McDonald, B. (2011). When Is a Liability Not a Liability? Textual Analysis, Dictionaries, and 10-Ks. Journal of Finance 66(1). ", "https://ssrn.com/abstract=1331573"],
    ["Software Repository for Accounting and Finance, Loughran-McDonald Master Dictionary (1993–2025). University of Notre Dame. ", "https://sraf.nd.edu/loughranmcdonald-master-dictionary/"],
    ["Yahoo Finance, historical daily prices. ", "https://finance.yahoo.com/"],
    ["Google Antigravity. ", "https://antigravity.google/"],
  ].map(([label, url], i) =>
    new Paragraph({
      spacing: { after: 120, line: 300 },
      children: [
        new TextRun({ text: `[${i + 1}] ${label}` }),
        new ExternalHyperlink({
          children: [new TextRun({ text: url, style: "Hyperlink" })],
          link: url,
        }),
      ],
    })
  ),
);

// ------------------------------------------------------------- appendix
children.push(
  new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "Appendix A. AI Prompts Used (Vibe-Coding History)" })] }),
  P(
    "This appendix records the AI prompts behind the project. Section A.1 quotes the user prompts exactly as they were issued during development. Section A.2 is a staged prompt sequence reconstructed from the development history: each entry corresponds to one real development stage, including the failures and fixes actually encountered while running the code against SEC data. The staged sequence is the recommended way to reproduce the workflow prompt-by-prompt in an AI coding tool."
  ),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "A.1 Actual user prompts (verbatim)", bold: true })] }),
  new Paragraph({
    spacing: { after: 140, line: 300 },
    children: [new TextRun({ text: "User prompt 1. ", bold: true }), new TextRun({ text: "The complete initial message is reproduced below, unchanged — the assignment brief pasted from the course page, followed by the student\u2019s own request (spelling and spacing included).", italics: true, color: MUTED })],
  }),
  ...PROMPT1_FULL.split(/\n\s*\n/).map(par =>
    new Paragraph({
      spacing: { after: 120, line: 280 },
      indent: { left: 360, right: 360 },
      children: [new TextRun({ text: par.trim(), size: 18, color: "333333" })],
    })
  ),
  ...[
    ["\u201CIntel (INTC)\u201D", "Answer to the setup question selecting the second company to analyze alongside AAPL."],
    ["\u201CGraduate student\u201D", "Answer to the setup question selecting the three-question analysis track."],
    ["\u201CBoth\u201D", "Answer to the setup question selecting both submission routes: the zipped desktop app and the GitHub-hosted HTML page."],
  ].map(([q, note], i) =>
    new Paragraph({
      spacing: { after: 140, line: 300 },
      children: [new TextRun({ text: `User prompt ${i + 2}. `, bold: true }), new TextRun({ text: q }), new TextRun({ text: " " + note, italics: true, color: MUTED })],
    })
  ),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "A.2 Suggested reproduction prompts (reconstructed)", bold: true })] }),
  ...[
    "Write a Python function that takes a stock ticker, resolves it to an SEC CIK number, and returns the ten most recent 10-K filings with their accession numbers, filing dates, and balance-sheet dates from the EDGAR submissions API. Handle amended 10-K/A filings by skipping them, and follow SEC rate limits with a descriptive User-Agent.",
    "Now download each 10-K's primary document and save it under filings/{TICKER}/ with a filename like AAPL_10K_FY2025_2025-10-31.htm. Add retry logic with back-off, and skip cleanly if a document is missing.",
    "Write a function that extracts the Item 7 MD&A section from a 10-K's HTML as plain text. It must find the real MD&A heading, not the table-of-contents entry, and cut the text at the Item 7A heading — but not at inline cross-references like 'Item 8 of this Form 10-K' inside the narrative. Save each MD&A to a labeled text file.",
    "The Intel filings label the section 'Management's Discussion and Analysis' without an Item 7 prefix. Add a fallback that locates standalone MD&A headings while ignoring the front table of contents, the back index, and the glossary, and ends at the market-risk section.",
    "Intel has fewer than 10 filings in the main submissions JSON because older filings live in EDGAR's archived submission files. Extend the loader to follow the archive index files until the requested count is found; note the archived JSON keeps its arrays at the top level.",
    "Extract the consolidated income statement, balance sheet, and cash-flow statement from each filing using the FilingSummary.xml R-files, and merge each statement across all ten filings into one wide table, periods side by side, keeping the most recent filing's presentation for overlapping periods and preserving the natural statement row order.",
    "The balance-sheet parse produced duplicate row labels ('Marketable securities' appears twice). Make labels unique with a numbered suffix so alignment works, and write the statement title with its units above each table in Excel.",
    "Write a text-cleaning function following the Loughran-McDonald methodology: lowercase, remove punctuation and number-only tokens, and keep the cleaned text for scoring. Then write a scorer that loads the LM master dictionary CSV and reports total word count, counts and percentages for negative, positive, uncertainty, litigious, strong modal, weak modal, and constraining categories, and the positive/negative ratio.",
    "Compute four ratios for each of the ten years from SEC XBRL company-concept data: current ratio, liabilities-to-equity, asset turnover on average assets, and net profit margin. Only accept 10-K annual values — filter out the quarterly periods that Apple used to tag — and handle concept fallbacks (RevenueFromContractWithCustomerExcludingAssessedTax → Revenues → SalesRevenueNet). For total liabilities, use the reported Liabilities tag when available and otherwise derive it as assets minus equity so the balance-sheet identity holds.",
    "Intel's 10-K balance sheet presents no standalone 'Total liabilities' subtotal — the liabilities-to-equity ratio was silently computing assets over equity. Fix it to derive liabilities from the identity and relabel the column.",
    "Fetch the stock price at each balance-sheet date from Yahoo Finance's chart API without an API key, preferring the last close on or before the date, and write a Prices tab.",
    "Build the Excel workbook with openpyxl: a README tab with metadata and the filing index with EDGAR links; one tab each for the three statements with all periods side by side; a Sentiment tab; a Stock Prices tab; a Ratios tab; and a Chart tab with a combined line chart of stock price and P/N ratio on two axes. Style the headers, freeze panes, and add number formats.",
    "The chart data table was broken when a year's sentiment row was missing. Add completion checks that fail the run if any of the ten years is missing from the filings, MD&A, sentiment, prices, or ratios, and save the cleaned text files alongside the raw MD&A.",
    "Write a verification script that checks all ten fiscal years appear on every tab, that assets equal liabilities plus equity on the balance-sheet tab for every period, that income-statement revenue and net income match the XBRL inputs used for ratios, and that the P/N ratio equals positive over negative counts. It must handle both Apple's presentation (standalone Total liabilities) and Intel's (combined liabilities-and-equity row).",
    "Write a desktop app: a Tkinter GUI that takes a ticker and filing count, runs the analyzer in a background thread while streaming its log into the window, and has a button that opens the output folder. Cross-platform, no external GUI dependencies.",
    "Generate comparison figures with matplotlib for both companies — sentiment vs price per company, the four ratios side by side, R&D intensity, gross margin, and capital allocation — and export a JSON dataset that a self-contained Chart.js HTML dashboard can render, plus a generator script for that dashboard.",
    "Finally, write the lab document: an objective section, a requirement checklist mapping every instruction item to its artifact, a vibe-coding narrative with a model comparison, methodology with the MD&A boundary audit, results for both companies with the tables and charts, the three graduate questions answered from the 10-K data, a step-by-step reproduction guide with common errors, references with URLs, and this prompt history in an appendix.",
  ].map((p, i) =>
    new Paragraph({
      spacing: { after: 140, line: 300 },
      children: [new TextRun({ text: `Prompt ${i + 1}. `, bold: true }), new TextRun({ text: p })],
    })
  ),
);

// ---------------------------------------------------------------- footnotes
const footnotes = {
  1: { children: [new Paragraph({ children: [new TextRun({ text: "U.S. Securities and Exchange Commission, EDGAR. ", }), new ExternalHyperlink({ children: [new TextRun({ text: "https://www.sec.gov/edgar/", style: "Hyperlink" })], link: "https://www.sec.gov/edgar/" })] })] },
  2: { children: [new Paragraph({ children: [new TextRun({ text: "Software Repository for Accounting and Finance, Loughran-McDonald Master Dictionary 1993–2025. ", }), new ExternalHyperlink({ children: [new TextRun({ text: "https://sraf.nd.edu/loughranmcdonald-master-dictionary/", style: "Hyperlink" })], link: "https://sraf.nd.edu/loughranmcdonald-master-dictionary/" })] })] },
  3: { children: [new Paragraph({ children: [new TextRun({ text: "Loughran, T., & McDonald, B. (2011). When Is a Liability Not a Liability? Textual Analysis, Dictionaries, and 10-Ks. Journal of Finance 66(1). ", }), new ExternalHyperlink({ children: [new TextRun({ text: "https://ssrn.com/abstract=1331573", style: "Hyperlink" })], link: "https://ssrn.com/abstract=1331573" })] })] },
  4: { children: [new Paragraph({ children: [new TextRun({ text: "Yahoo Finance historical prices. ", }), new ExternalHyperlink({ children: [new TextRun({ text: "https://finance.yahoo.com/", style: "Hyperlink" })], link: "https://finance.yahoo.com/" })] })] },
  5: { children: [new Paragraph({ children: [new TextRun({ text: "Apple Inc. Form 10-K FY2025. ", }), new ExternalHyperlink({ children: [new TextRun({ text: "https://www.sec.gov/Archives/edgar/data/320193/000032019325000079/aapl-20250927.htm", style: "Hyperlink" })], link: "https://www.sec.gov/Archives/edgar/data/320193/000032019325000079/aapl-20250927.htm" })] })] },
  6: { children: [new Paragraph({ children: [new TextRun({ text: "Intel Corporation Form 10-K FY2025. ", }), new ExternalHyperlink({ children: [new TextRun({ text: "https://www.sec.gov/Archives/edgar/data/50863/000005086326000011/", style: "Hyperlink" })], link: "https://www.sec.gov/Archives/edgar/data/50863/000005086326000011/" })] })] },
  7: { children: [new Paragraph({ children: [new TextRun({ text: "Intel Corporation Form 10-K FY2024. ", }), new ExternalHyperlink({ children: [new TextRun({ text: "https://www.sec.gov/Archives/edgar/data/50863/000005086325000009/", style: "Hyperlink" })], link: "https://www.sec.gov/Archives/edgar/data/50863/000005086325000009/" })] })] },
};

// ---------------------------------------------------------------- document
const doc = new Document({
  creator: "Amirtha Varshini V",
  title: "SEC 10-K Sentiment and Financial Analysis — AAPL and INTC",
  description: "Gen AI vibe-coding final project lab document",
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Calibri", color: NAVY },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Calibri", color: NAVY },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022",
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  footnotes,
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "10-K Sentiment & Financial Analysis — AAPL vs INTC", italics: true, color: "999999", size: 16 })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Page ", size: 18 }), new TextRun({ children: [PageNumber.CURRENT], size: 18 }), new TextRun({ text: " of ", size: 18 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18 })],
          })],
        }),
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(BASE, "output", "Lab_Document_10K_Sentiment_Analysis.docx");
  fs.writeFileSync(out, buf);
  console.log("written:", out, buf.length, "bytes");
});
