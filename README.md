# Trading Analysis Lab

Python-based tools for technical analysis, rule-based stock screening, RSI threshold optimization, and simple signal generation on BIST stocks.

This repository contains early-stage algorithmic trading research scripts developed for educational purposes. The main focus is not to provide financial advice, but to study systematic decision-making, technical indicators, scoring models, and risk-aware market analysis.

## Features

- BIST100 stock screening dashboard with a Tkinter GUI
- RSI buy/sell threshold optimization with heatmap output
- Breakout signal scanner using SMA200, 20-day high, and 10-day low rules
- Timestamped Excel and image outputs
- BIST100 company-name-to-ticker mapping using a CSV database

## Project Structure

```text
trading-analysis-lab/
│
├── data/
│   └── bist100_companies.csv
│
├── src/
│   ├── bist_stock_scanner_gui.py
│   ├── breakout_signal_scanner.py
│   └── rsi_threshold_optimizer.py
│
├── outputs/
│   └── generated locally, ignored by Git
│
├── requirements.txt
├── .gitignore
└── README.md