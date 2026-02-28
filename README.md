# EV Bet Calculator

A Python desktop application for tracking and analyzing sports & esports bets. Uses a local SQLite database to store all bet data and calculates **Expected Value (EV)** using Bayesian shrinkage to help decide which odds are worth betting on.

## Features

- **Local SQLite database** — all bet data stored locally, no external API or credentials needed
- **Add & manage bets** — record sport, tournament, matchup, odds, stake and result
- **Pending bets** — track unsettled bets and settle them with one click (profit auto-calculated)
- **Data Table** — sortable, filterable view of settled bets with EV calculations
- **Statistics** — visual charts and win-rate breakdowns (WIP)
- **Data migration** — import existing bet history from `.xlsx` spreadsheets
- **Dark / Light theme** — Catppuccin-inspired themes, switchable from settings
- **Database management** — reset / delete all data from the settings dialog

## Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt`

## Installation

```bash
# Clone the repository
git clone <repo-url>
cd "EV BET calculator"

# Create a virtual environment (recommended)
python -m venv venv
venv\Scripts\activate   # Windows
# source venv/bin/activate  # macOS / Linux

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
python src/main.py
```

On first launch the database is created automatically. To import data from an existing `.xlsx` spreadsheet, go to **Settings → Migration** and select your file. The expected column layout is:

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| Sport | Tournament | Matchup | Bet | Live Status | Odds | Bet Amount (€) | Result (Win/Lose) | Profit |
