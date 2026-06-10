"""
BIST100 Stock Scanner GUI

This script provides a Tkinter-based stock screening dashboard for BIST100 stocks.
It loads company and ticker data from data/bist100_companies.csv, analyzes each stock,
and ranks them using a rule-based scoring model.

Scoring factors:
- Trend strength: price compared with SMA50 and SMA200
- Momentum: 3-month price performance
- Relative strength: 3-month performance compared with BIST100
- Valuation: trailing P/E ratio
- Profitability: return on equity
- Liquidity: trading volume
- Volatility: daily price fluctuation risk
- Volume activity: current volume compared with recent average volume

Output:
- Interactive GUI table
- Timestamped Excel report in the outputs/ folder

This project is for educational and research purposes only.
It does not provide financial advice.
"""

from datetime import datetime
from pathlib import Path
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd
import yfinance as yf


COMPANY_DATA_PATH = "data/bist100_companies.csv"
OUTPUT_DIR = "outputs"


class BISTStockScannerApp:
    """Tkinter GUI application for scanning and ranking BIST100 stocks."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("BIST100 Stock Screening Dashboard")
        self.root.geometry("1420x840")
        self.root.minsize(1200, 760)

        self.colors = {
            "bg": "#f4f6f8",
            "header": "#111827",
            "header_2": "#1f2937",
            "card": "#ffffff",
            "border": "#d1d5db",
            "text": "#111827",
            "muted": "#6b7280",
            "green": "#16a34a",
            "blue": "#2563eb",
            "red": "#dc2626",
            "yellow": "#f59e0b",
        }

        self.results = []
        self.bist_3_month_return = None
        self.company_data = self.load_company_database()

        self.configure_style()
        self.build_layout()

    def configure_style(self) -> None:
        """Configure the GUI style."""
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#e5e7eb",
            foreground="#111827",
            relief="flat",
        )
        style.configure(
            "Treeview",
            font=("Segoe UI", 9),
            rowheight=29,
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#111827",
        )
        style.map(
            "Treeview",
            background=[("selected", "#2563eb")],
            foreground=[("selected", "#ffffff")],
        )

        style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#e5e7eb",
            background="#2563eb",
            thickness=10,
        )

    def build_layout(self) -> None:
        """Build the GUI layout."""
        self.root.configure(bg=self.colors["bg"])

        header_frame = tk.Frame(self.root, bg=self.colors["header"], pady=18)
        header_frame.pack(fill="x")

        title_block = tk.Frame(header_frame, bg=self.colors["header"])
        title_block.pack(side="left", padx=24)

        title_label = tk.Label(
            title_block,
            text="BIST100 Stock Screening Dashboard",
            font=("Segoe UI", 18, "bold"),
            bg=self.colors["header"],
            fg="white",
            anchor="w",
        )
        title_label.pack(anchor="w")

        subtitle_label = tk.Label(
            title_block,
            text="Rule-based screening using trend, momentum, relative strength, valuation, profitability, liquidity, and risk metrics.",
            font=("Segoe UI", 9),
            bg=self.colors["header"],
            fg="#cbd5e1",
            anchor="w",
        )
        subtitle_label.pack(anchor="w", pady=(4, 0))

        button_frame = tk.Frame(header_frame, bg=self.colors["header"])
        button_frame.pack(side="right", padx=24)

        self.export_button = tk.Button(
            button_frame,
            text="Export Excel",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors["blue"],
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            padx=18,
            pady=8,
            cursor="hand2",
            relief="flat",
            command=self.export_to_excel,
        )
        self.export_button.pack(side="left", padx=(0, 10))

        self.scan_button = tk.Button(
            button_frame,
            text="Start Scan",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors["green"],
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            padx=22,
            pady=8,
            cursor="hand2",
            relief="flat",
            command=self.start_scan,
        )
        self.scan_button.pack(side="left")

        info_frame = tk.Frame(self.root, bg=self.colors["bg"], pady=14)
        info_frame.pack(fill="x", padx=16)

        metric_cards = [
            ("Trend", "SMA50 / SMA200 structure"),
            ("Momentum", "3-month price return"),
            ("Relative Strength", "Performance vs BIST100"),
            ("Valuation", "Trailing P/E ratio"),
            ("Profitability", "Return on equity"),
            ("Liquidity", "Volume and volume activity"),
            ("Risk", "Daily volatility penalty"),
        ]

        for title, description in metric_cards:
            card = tk.Frame(
                info_frame,
                bg=self.colors["card"],
                highlightbackground=self.colors["border"],
                highlightthickness=1,
                padx=12,
                pady=9,
            )
            card.pack(side="left", fill="both", expand=True, padx=5)

            tk.Label(
                card,
                text=title,
                font=("Segoe UI", 9, "bold"),
                bg=self.colors["card"],
                fg=self.colors["text"],
                anchor="w",
            ).pack(anchor="w")

            tk.Label(
                card,
                text=description,
                font=("Segoe UI", 8),
                bg=self.colors["card"],
                fg=self.colors["muted"],
                anchor="w",
                wraplength=150,
                justify="left",
            ).pack(anchor="w", pady=(3, 0))

        status_frame = tk.Frame(self.root, bg=self.colors["bg"])
        status_frame.pack(fill="x", padx=22, pady=(0, 8))

        self.status_label = tk.Label(
            status_frame,
            text="Ready. Press Start Scan to analyze BIST100 stocks.",
            font=("Segoe UI", 10),
            bg=self.colors["bg"],
            fg=self.colors["text"],
            anchor="w",
        )
        self.status_label.pack(side="left")

        self.time_label = tk.Label(
            status_frame,
            text="",
            font=("Segoe UI", 9),
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            anchor="e",
        )
        self.time_label.pack(side="right")

        self.progress = ttk.Progressbar(
            self.root,
            orient="horizontal",
            mode="determinate",
            style="Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", padx=22, pady=(0, 12))

        table_frame = tk.Frame(
            self.root,
            bg=self.colors["card"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
        )
        table_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        columns = (
            "symbol",
            "company",
            "score",
            "price",
            "three_month_return",
            "relative_strength",
            "pe_ratio",
            "roe",
            "volume",
            "risk",
            "trend",
        )

        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=20,
        )

        headings = {
            "symbol": "Ticker",
            "company": "Company",
            "score": "Score",
            "price": "Price",
            "three_month_return": "3M Return",
            "relative_strength": "vs BIST100",
            "pe_ratio": "P/E",
            "roe": "ROE",
            "volume": "Volume",
            "risk": "Risk",
            "trend": "Trend",
        }

        widths = {
            "symbol": 80,
            "company": 190,
            "score": 70,
            "price": 80,
            "three_month_return": 100,
            "relative_strength": 110,
            "pe_ratio": 80,
            "roe": 80,
            "volume": 100,
            "risk": 80,
            "trend": 130,
        }

        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="center")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.tag_configure("high_score", background="#dcfce7", foreground="#111827")
        self.tree.tag_configure("medium_score", background="#eff6ff", foreground="#111827")
        self.tree.tag_configure("low_score", background="#fee2e2", foreground="#111827")

    def load_company_database(self) -> pd.DataFrame:
        """Load BIST100 company ticker database."""
        company_data = pd.read_csv(COMPANY_DATA_PATH)

        required_columns = {"ticker", "company_name", "aliases"}
        missing_columns = required_columns - set(company_data.columns)

        if missing_columns:
            raise ValueError(f"Missing columns in company database: {missing_columns}")

        return company_data

    def get_tickers(self) -> list[str]:
        """Return ticker list from the company database."""
        return self.company_data["ticker"].dropna().astype(str).str.upper().tolist()

    def get_company_name(self, ticker: str) -> str:
        """Return company name for a given ticker."""
        match = self.company_data[self.company_data["ticker"].str.upper() == ticker.upper()]

        if match.empty:
            return "-"

        return str(match.iloc[0]["company_name"])

    def calculate_bist100_performance(self) -> float:
        """Calculate the 3-month performance of the BIST100 index."""
        if self.bist_3_month_return is not None:
            return self.bist_3_month_return

        try:
            bist_data = yf.Ticker("XU100.IS").history(period="1y")

            if bist_data.empty or len(bist_data) < 60:
                self.bist_3_month_return = 0.0
                return self.bist_3_month_return

            latest_close = bist_data["Close"].iloc[-1]
            previous_close = bist_data["Close"].iloc[-60]

            self.bist_3_month_return = ((latest_close - previous_close) / previous_close) * 100
            return self.bist_3_month_return

        except Exception as error:
            print(f"BIST100 performance could not be calculated: {error}")
            self.bist_3_month_return = 0.0
            return self.bist_3_month_return

    def start_scan(self) -> None:
        """Start the scanning process in a background thread."""
        self.scan_button.config(state="disabled", text="Scanning...")
        self.start_time = datetime.now()

        scan_thread = threading.Thread(target=self.scan_and_rank_stocks, daemon=True)
        scan_thread.start()

    def scan_and_rank_stocks(self) -> None:
        """Scan all BIST100 stocks and rank them by score."""
        self.results = []
        tickers = self.get_tickers()
        total_tickers = len(tickers)

        successful = 0
        failed = 0

        self.clear_table()
        self.calculate_bist100_performance()

        for index, ticker in enumerate(tickers):
            progress_value = (index / total_tickers) * 100
            self.update_progress(progress_value)
            self.update_status(f"Analyzing {ticker} ({index + 1}/{total_tickers})")

            result = self.analyze_stock(ticker)

            if result is not None:
                self.results.append(result)
                successful += 1
            else:
                failed += 1

        self.results.sort(key=lambda item: item["raw_score"], reverse=True)

        self.root.after(0, lambda: self.fill_table(self.results))

        elapsed = (datetime.now() - self.start_time).total_seconds()
        self.update_status(f"Scan completed. {successful} successful, {failed} failed.")
        self.root.after(
            0,
            lambda: self.time_label.config(
                text=f"Elapsed: {elapsed:.1f}s | BIST100 3M: {self.bist_3_month_return:.1f}%"
            ),
        )
        self.root.after(0, lambda: self.scan_button.config(state="normal", text="Start Scan"))
        self.update_progress(100)

    def analyze_stock(self, ticker: str) -> dict | None:
        """Analyze a single stock and return scoring details."""
        try:
            stock = yf.Ticker(ticker)
            info = stock.info

            price = info.get("currentPrice") or info.get("regularMarketPrice")
            historical_data = stock.history(period="1y")

            if historical_data.empty or len(historical_data) < 200:
                print(f"Skipped {ticker}: not enough historical data.")
                return None

            close = historical_data["Close"]
            volume = historical_data["Volume"]

            latest_close = close.iloc[-1]
            price = price or latest_close

            sma50 = close.rolling(window=50).mean().iloc[-1]
            sma200 = close.rolling(window=200).mean().iloc[-1]

            three_month_previous_close = close.iloc[-60] if len(close) >= 60 else close.iloc[0]
            three_month_return = ((latest_close - three_month_previous_close) / three_month_previous_close) * 100

            daily_volatility = close.pct_change().std() * 100

            average_volume_20 = volume.rolling(window=20).mean().iloc[-1] if len(volume) >= 20 else volume.mean()
            latest_volume = volume.iloc[-1]
            volume_spike = latest_volume > average_volume_20 * 1.5 if average_volume_20 > 0 else False

            pe_ratio = info.get("trailingPE", None)
            roe_raw = info.get("returnOnEquity", None)
            roe = roe_raw * 100 if roe_raw is not None else None

            bist_return = self.calculate_bist100_performance()
            relative_strength = three_month_return - bist_return

            score = self.calculate_score(
                price=price,
                sma50=sma50,
                sma200=sma200,
                three_month_return=three_month_return,
                relative_strength=relative_strength,
                pe_ratio=pe_ratio,
                roe=roe,
                volume=latest_volume,
                volume_spike=volume_spike,
                daily_volatility=daily_volatility,
            )

            trend = self.determine_trend(price, sma50, sma200)

            return {
                "symbol": ticker.replace(".IS", ""),
                "company": self.get_company_name(ticker),
                "raw_score": score,
                "score": f"{score:.1f}",
                "price": f"{price:.2f}",
                "three_month_return": f"{three_month_return:.1f}%",
                "relative_strength": f"{relative_strength:+.1f}%",
                "pe_ratio": f"{pe_ratio:.1f}" if pe_ratio and pe_ratio > 0 else "-",
                "roe": f"{roe:.1f}%" if roe is not None else "-",
                "volume": f"{latest_volume / 1000:.0f}K",
                "risk": f"{daily_volatility:.1f}%",
                "trend": trend,
            }

        except Exception as error:
            print(f"Error while analyzing {ticker}: {str(error)[:100]}")
            return None

    def calculate_score(
        self,
        price: float,
        sma50: float,
        sma200: float,
        three_month_return: float,
        relative_strength: float,
        pe_ratio: float | None,
        roe: float | None,
        volume: float,
        volume_spike: bool,
        daily_volatility: float,
    ) -> float:
        """Calculate a rule-based score between 0 and 10."""
        score = 5.0

        if price > sma50 > sma200:
            score += 3.0
        elif price > sma200:
            score += 1.5
        elif price < sma200:
            score -= 1.0

        if three_month_return > 30:
            score += 2.0
        elif three_month_return > 15:
            score += 1.0
        elif three_month_return < -10:
            score -= 1.0

        if relative_strength > 15:
            score += 1.0
        elif relative_strength > 5:
            score += 0.5
        elif relative_strength < -10:
            score -= 0.5

        if pe_ratio and pe_ratio > 0:
            if pe_ratio < 8:
                score += 2.0
            elif pe_ratio < 15:
                score += 1.0
            elif pe_ratio > 30:
                score -= 0.5

        if roe is not None:
            if roe > 25:
                score += 1.5
            elif roe > 15:
                score += 0.5

        if volume > 1_000_000:
            score += 1.0
        elif volume < 100_000:
            score -= 1.0

        if volume_spike:
            score += 0.5

        if daily_volatility > 5:
            score -= 0.5

        return max(0.0, min(10.0, score))

    def determine_trend(self, price: float, sma50: float, sma200: float) -> str:
        """Determine basic trend status."""
        if price > sma50 > sma200:
            return "Strong uptrend"
        if price > sma200:
            return "Uptrend"
        return "Downtrend"

    def fill_table(self, data: list[dict]) -> None:
        """Fill the GUI table with scan results."""
        for item in data:
            score = item["raw_score"]

            if score >= 7.5:
                tag = "high_score"
            elif score >= 5.5:
                tag = "medium_score"
            else:
                tag = "low_score"

            values = (
                item["symbol"],
                item["company"],
                item["score"],
                item["price"],
                item["three_month_return"],
                item["relative_strength"],
                item["pe_ratio"],
                item["roe"],
                item["volume"],
                item["risk"],
                item["trend"],
            )

            self.tree.insert("", "end", values=values, tags=(tag,))

    def export_to_excel(self) -> None:
        """Export scan results to a timestamped Excel file."""
        if not self.results:
            messagebox.showwarning("Warning", "Please run a scan before exporting.")
            return

        try:
            output_folder = Path(OUTPUT_DIR)
            output_folder.mkdir(parents=True, exist_ok=True)

            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            output_path = output_folder / f"bist100_stock_scan_{timestamp}.xlsx"

            export_data = pd.DataFrame(self.results)

            export_data = export_data.rename(
                columns={
                    "symbol": "Ticker",
                    "company": "Company",
                    "raw_score": "Raw Score",
                    "score": "Score",
                    "price": "Price",
                    "three_month_return": "3M Return",
                    "relative_strength": "Relative Strength vs BIST100",
                    "pe_ratio": "P/E",
                    "roe": "ROE",
                    "volume": "Volume",
                    "risk": "Risk",
                    "trend": "Trend",
                }
            )

            export_data.to_excel(output_path, index=False, engine="openpyxl")

            messagebox.showinfo("Export Completed", f"Excel report saved to:\n{output_path}")
            print(f"Excel report saved to: {output_path}")

        except Exception as error:
            messagebox.showerror("Export Error", f"Excel export failed:\n{error}")
            print(f"Excel export failed: {error}")

    def clear_table(self) -> None:
        """Clear the GUI table."""
        self.root.after(0, lambda: [self.tree.delete(item) for item in self.tree.get_children()])

    def update_status(self, text: str) -> None:
        """Update status label from worker thread."""
        self.root.after(0, lambda: self.status_label.config(text=text))

    def update_progress(self, value: float) -> None:
        """Update progress bar from worker thread."""
        self.root.after(0, lambda: self.progress.config(value=value))


if __name__ == "__main__":
    try:
        root_window = tk.Tk()
        app = BISTStockScannerApp(root_window)
        root_window.mainloop()

    except FileNotFoundError:
        print(f"Company database not found: {COMPANY_DATA_PATH}")
    except ValueError as error:
        print(f"Configuration error: {error}")