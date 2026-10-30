"""
download_dictionary.py — Fetch the Loughran-McDonald master dictionary.

Downloads the official Loughran-McDonald Master Dictionary CSV
(1993-2025 edition) from the Notre Dame Software Repository for
Accounting and Finance (SRAF) into data/. The dictionary is free for
use in academic research; for commercial licenses contact the SRAF.

The official distribution link (a Google Drive file shared on
https://sraf.nd.edu/loughranmcdonald-master-dictionary/) is:

    https://drive.google.com/uc?export=download&id=1iq2RUf8qGFEAk1g8wQntP3habOnR3fXF

Run:  python download_dictionary.py
"""

import re
from pathlib import Path

import requests

DRIVE_ID = "1iq2RUf8qGFEAk1g8wQntP3habOnR3fXF"
DATA_DIR = Path(__file__).resolve().parent / "data"
TARGET = DATA_DIR / "Loughran-McDonald_MasterDictionary_1993-2025.csv"


def main() -> None:
    """Download the dictionary CSV to data/ (about 9 MB)."""
    DATA_DIR.mkdir(exist_ok=True)
    if TARGET.exists() and TARGET.stat().st_size > 5_000_000:
        print(f"Dictionary already present: {TARGET}")
        return
    session = requests.Session()
    url = f"https://drive.google.com/uc?export=download&id={DRIVE_ID}"
    print(f"Downloading Loughran-McDonald master dictionary from {url}")
    response = session.get(url, timeout=120)
    content_type = response.headers.get("Content-Type", "")
    if "text/html" in content_type and len(response.content) < 200_000:
        # Google Drive interstitial for large files: extract the
        # direct-download URL from the confirmation page.
        match = re.search(
            r'https://drive\.usercontent\.google\.com/download\?'
            r'[^"\'<>]+',
            response.text,
        )
        if match is None:
            raise RuntimeError(
                "Could not resolve the Google Drive download link; "
                "visit https://sraf.nd.edu/loughranmcdonald-master-dictionary/ "
                "and download the CSV manually into data/."
            )
        response = session.get(match.group(0), timeout=300)
    if len(response.content) < 5_000_000:
        raise RuntimeError(
            f"Downloaded file is unexpectedly small "
            f"({len(response.content)} bytes); aborting."
        )
    TARGET.write_bytes(response.content)
    print(f"Saved {TARGET} ({len(response.content):,} bytes)")


if __name__ == "__main__":
    main()
