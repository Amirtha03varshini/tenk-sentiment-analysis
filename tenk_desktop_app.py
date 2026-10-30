"""
tenk_desktop_app.py — Desktop GUI for the 10-K analysis pipeline.

A small Tkinter front end around tenk_analyzer.analyze(): the user
types a ticker and the number of 10-Ks to analyze, clicks 'Analyze',
and watches the pipeline log stream by. When the run completes, the
'Open output folder' button reveals the generated Excel workbook.

Run:  python tenk_desktop_app.py
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import scrolledtext, ttk
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent


class AnalyzerApp:
    """Main application window."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("SEC 10-K Sentiment Analyzer")
        self.root.geometry("860x560")
        self.queue: "queue.Queue[str]" = queue.Queue()
        self.process: subprocess.Popen | None = None

        # --- Input row -------------------------------------------------
        frame = ttk.Frame(root, padding=10)
        frame.pack(fill=tk.X)

        ttk.Label(frame, text="Ticker:").grid(row=0, column=0, padx=4)
        self.ticker_var = tk.StringVar(value="AAPL")
        ttk.Entry(frame, textvariable=self.ticker_var, width=10).grid(
            row=0, column=1, padx=4
        )

        ttk.Label(frame, text="Number of 10-Ks:").grid(
            row=0, column=2, padx=4
        )
        self.count_var = tk.StringVar(value="10")
        ttk.Entry(frame, textvariable=self.count_var, width=6).grid(
            row=0, column=3, padx=4
        )

        self.run_button = ttk.Button(
            frame, text="Analyze", command=self.start_analysis
        )
        self.run_button.grid(row=0, column=4, padx=8)

        self.open_button = ttk.Button(
            frame, text="Open output folder", command=self.open_output
        )
        self.open_button.grid(row=0, column=5, padx=8)

        # --- Log area ---------------------------------------------------
        self.log = scrolledtext.ScrolledText(
            root, wrap=tk.WORD, font=("Consolas", 10)
        )
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        self.root.after(100, self.poll_queue)
        self.log.insert(
            tk.END,
            "Enter a ticker symbol (e.g. AAPL) and the number of "
            "10-K filings to analyze, then click Analyze.\n\n"
            "Note: first-time users should run "
            "'python download_dictionary.py' once so the "
            "Loughran-McDonald dictionary is available, and set the "
            "SEC_USER_AGENT environment variable to their name and "
            "email (SEC fair-access policy).\n\n",
        )
        self.log.configure(state=tk.DISABLED)

    # ------------------------------------------------------------------
    def poll_queue(self) -> None:
        """Move pending log lines from the queue into the text area."""
        try:
            while True:
                line = self.queue.get_nowait()
                self.log.configure(state=tk.NORMAL)
                self.log.insert(tk.END, line)
                self.log.see(tk.END)
                self.log.configure(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)

    def start_analysis(self) -> None:
        """Launch the analysis pipeline in a background thread."""
        if self.process is not None and self.process.poll() is None:
            self.queue.put("A run is already in progress; please wait.\n")
            return
        ticker = self.ticker_var.get().strip().upper()
        try:
            count = int(self.count_var.get())
        except ValueError:
            self.queue.put("Number of 10-Ks must be an integer.\n")
            return
        if not ticker:
            self.queue.put("Please enter a ticker symbol.\n")
            return

        self.queue.put(f"Starting analysis of {ticker} ({count} 10-Ks)...\n")
        self.run_button.configure(state=tk.DISABLED)
        thread = threading.Thread(
            target=self.run_pipeline, args=(ticker, count), daemon=True
        )
        thread.start()

    def run_pipeline(self, ticker: str, count: int) -> None:
        """Run tenk_analyzer as a subprocess and stream its output."""
        command = [
            sys.executable,
            str(BASE_DIR / "tenk_analyzer.py"),
            ticker,
            str(count),
        ]
        self.process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
        assert self.process.stdout is not None
        for line in self.process.stdout:
            self.queue.put(line)
        code = self.process.wait()
        if code == 0:
            self.queue.put(
                f"\nDone. Workbook: output/{ticker}_10K_Analysis.xlsx\n"
            )
        else:
            self.queue.put(f"\nAnalysis exited with code {code}.\n")
        self.process = None
        self.root.after(0, lambda: self.run_button.configure(state=tk.NORMAL))

    def open_output(self) -> None:
        """Open the output folder in the system file explorer."""
        path = BASE_DIR / "output"
        path.mkdir(exist_ok=True)
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])


def main() -> None:
    """Launch the application."""
    root = tk.Tk()
    AnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
