"""
audit_mda.py — Verification audit of MD&A extraction boundaries.

For every saved 10-K, re-runs the extraction logic and reports:
  * which strategy matched (Item-7 heading vs standalone heading),
  * which terminator ended the segment (Item 7A heading, loose Item 7A,
    market-risk heading, or the character-limit fallback),
  * the position (as % of document) of the start and end,
  * opening and closing excerpts for manual inspection.

Run:  python audit_mda.py
"""

from pathlib import Path
import re
import sys

sys.path.insert(0, str(Path(__file__).parent))
from tenk_analyzer import _strip_tags  # noqa: E402


def audit_document(html_path: Path) -> None:
    """Print the extraction-boundary report for one 10-K."""
    html = html_path.read_text(encoding="utf-8", errors="ignore")
    text = _strip_tags(html)
    n = len(text)

    # ---- Strategy 1: 'Item 7 ... Management's' headings -------------
    start_pat = re.compile(
        r"item\s+7\.?\s*[\.\:\-]*\s*(management|management\u2019s|"
        r"management's)",
        re.IGNORECASE,
    )
    end_strict = re.compile(
        r"item\s+7a\.?[\s\xa0]*(quantitative|qualitative)",
        re.IGNORECASE,
    )
    end_loose = re.compile(r"item\s+7a", re.IGNORECASE)

    def _words(seg: str) -> int:
        return len(seg.split())

    chosen = None
    for match in start_pat.finditer(text):
        for label, pattern in (
            ("Item 7A heading (strict)", end_strict),
            ("Item 7A (loose)", end_loose),
        ):
            end = pattern.search(text, match.start() + 50)
            if end is None:
                continue
            segment = text[match.start():end.start()]
            if _words(segment) > 500:
                chosen = (
                    f"Item-7 heading @ {match.start() / n:.1%}",
                    f"{label} @ {end.start() / n:.1%}",
                    segment,
                )
                break
        if chosen:
            break

    # ---- Strategy 2: standalone MD&A headings ----------------------
    if chosen is None:
        heading_pat = re.compile(
            r"management.{0,3}s discussion and analysis", re.IGNORECASE
        )
        end_pat = re.compile(
            r"(quantitative and qualitative disclosures about market risk"
            r"|item\s+7a\.?[\s\xa0]*(quantitative|qualitative)"
            r"|item\s+8\.?[\s\xa0]*financial statements)",
            re.IGNORECASE,
        )
        for match in heading_pat.finditer(text):
            if match.start() < 0.08 * n or match.start() > 0.92 * n:
                continue
            end = end_pat.search(text, match.start() + 50)
            if end is None:
                end_pos = min(match.start() + 150000, n)
                terminator = "150k-character cap (FALLBACK)"
            else:
                end_pos = end.start()
                terminator = f"market-risk/7A/8 pattern @ {end.start() / n:.1%}"
            segment = text[match.start():end_pos]
            if _words(segment) > 500:
                chosen = (
                    f"standalone heading @ {match.start() / n:.1%}",
                    terminator,
                    segment,
                )
                break

    name = html_path.name
    if chosen is None:
        print(f"{name}: EXTRACTION FAILED")
        return
    start_desc, end_desc, segment = chosen
    print(f"{name}")
    print(f"  start: {start_desc}   end: {end_desc}")
    print(f"  words: {_words(segment):,}")
    print(f"  opening: {segment[:90].strip()!r}")
    print(f"  closing: {segment[-90:].strip()!r}")
    print()


def main() -> None:
    """Audit every saved 10-K for both companies."""
    base = Path(__file__).parent / "filings"
    for company_dir in sorted(base.iterdir()):
        if not company_dir.is_dir():
            continue
        print(f"================ {company_dir.name} ================")
        for html_path in sorted(company_dir.glob("*.htm")):
            audit_document(html_path)


if __name__ == "__main__":
    main()
