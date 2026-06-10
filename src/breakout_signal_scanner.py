"""
Breakout Signal Scanner

This script scans selected BIST stocks using a simple rule-based breakout strategy.
It prints a clean terminal report and exports the full scanner results to a timestamped Excel file.

Strategy logic:
- Trend filter: Close price must be above the 200-day simple moving average.
- Buy signal: Close price breaks above the previous 20-day high.
- Sell signal: Close price falls below the previous 10-day low.
- Otherwise, the script returns a neutral status.

Output:
- Terminal summary table
- Timestamped Excel report in the outputs/ folder

Example output file:
- outputs/breakout_signals_2026-06-10_15-58-22.xlsx

This project is for educational and research purposes only.
It does not provide financial advice.
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
import yfinance as yf


SYMBOLS = [
    "THYAO.IS",
    "ASELS.IS",
    "TUPRS.IS",
    "KCHOL.IS",
    "BIMAS.IS",
    "FROTO.IS",
    "AKBNK.IS",
]

ENTRY_LOOKBACK = 20
EXIT_LOOKBACK = 10
SMA_PERIOD = 200
OUTPUT_DIR = "outputs"


def download_price_data(symbol: str, period: str = "18mo") -> pd.DataFrame:
    """Download historical OHLCV data from Yahoo Finance."""
    data = yf.download(symbol, period=period, interval="1d", progress=False)

    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)

    return data.dropna()


def analyze_symbol(symbol: str) -> dict | None:
    """Analyze one stock and return its signal report."""
    data = download_price_data(symbol)

    if data.empty or len(data) < SMA_PERIOD:
        return None

    close = data["Close"]
    high = data["High"]
    low = data["Low"]

    latest_close = close.iloc[-1]
    latest_date = data.index[-1].date()

    sma_200 = close.rolling(window=SMA_PERIOD).mean().iloc[-1]
    previous_20_day_high = high.rolling(window=ENTRY_LOOKBACK).max().shift(1).iloc[-1]
    previous_10_day_low = low.rolling(window=EXIT_LOOKBACK).min().shift(1).iloc[-1]

    trend_is_positive = latest_close > sma_200

    if trend_is_positive and latest_close > previous_20_day_high:
        signal = "BUY"
        description = "Price broke above the previous 20-day high while above SMA200."
    elif latest_close < previous_10_day_low:
        signal = "SELL"
        description = "Price fell below the previous 10-day low."
    elif trend_is_positive:
        signal = "HOLD"
        description = "Price remains above SMA200, but there is no new breakout signal."
    else:
        signal = "AVOID"
        description = "Price is below SMA200 trend filter."

    distance_to_breakout = ((previous_20_day_high - latest_close) / latest_close) * 100
    distance_to_stop = ((latest_close - previous_10_day_low) / latest_close) * 100

    return {
        "Date": latest_date,
        "Symbol": symbol.replace(".IS", ""),
        "Close": round(latest_close, 2),
        "Signal": signal,
        "SMA200": round(sma_200, 2),
        "Breakout Level": round(previous_20_day_high, 2),
        "Stop Level": round(previous_10_day_low, 2),
        "Distance to Breakout (%)": round(distance_to_breakout, 2),
        "Distance to Stop (%)": round(distance_to_stop, 2),
        "Description": description,
    }


def run_scanner(symbols: list[str]) -> pd.DataFrame:
    """Run the breakout scanner for multiple symbols."""
    reports = []

    for symbol in symbols:
        try:
            result = analyze_symbol(symbol)

            if result is not None:
                reports.append(result)
            else:
                print(f"Skipped {symbol}: not enough data.")

        except Exception as error:
            print(f"Error while analyzing {symbol}: {error}")

    return pd.DataFrame(reports)


def create_timestamped_output_path(output_dir: str = OUTPUT_DIR) -> Path:
    """Create a unique timestamped Excel output path."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_folder = Path(output_dir)
    output_folder.mkdir(parents=True, exist_ok=True)

    return output_folder / f"breakout_signals_{timestamp}.xlsx"


def export_to_excel(report: pd.DataFrame, output_path: Path) -> None:
    """Export the scanner report to an Excel file."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        report.to_excel(writer, index=False, sheet_name="Breakout Signals")

        worksheet = writer.sheets["Breakout Signals"]

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))

            worksheet.column_dimensions[column_letter].width = max_length + 2


def print_terminal_report(report: pd.DataFrame) -> None:
    """Print a compact terminal summary."""
    display_columns = [
        "Date",
        "Symbol",
        "Close",
        "Signal",
        "SMA200",
        "Breakout Level",
        "Stop Level",
        "Distance to Breakout (%)",
        "Distance to Stop (%)",
    ]

    print("\nBIST Breakout Signal Scanner")
    print("=" * 110)
    print(report[display_columns].to_string(index=False))
    print("=" * 110)


if __name__ == "__main__":
    scanner_report = run_scanner(SYMBOLS)

    if scanner_report.empty:
        print("No valid results were generated.")
    else:
        print_terminal_report(scanner_report)

        excel_output_path = create_timestamped_output_path()
        export_to_excel(scanner_report, excel_output_path)

        print(f"\nExcel report saved to: {excel_output_path}")