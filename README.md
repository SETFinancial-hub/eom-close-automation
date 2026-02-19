# SET Financial — EOM Close Automation & QuickBooks MCP

Automated end-of-month close process for SET Financial Corporation, with QuickBooks Online integration via MCP (Model Context Protocol) for forecasting and financial modeling.

## Overview

This project automates SET Financial's month-end close workflow:

1. **Parse** Nortridge loan/collection/charge-off registers, unearned revenue reports, and funding statements (DPV, Pier/MetaCorp)
2. **Reconcile** portfolio totals across systems
3. **Generate** journal entries for QuickBooks Online
4. **Forecast** using QBO actuals + portfolio data via MCP

## Project Structure

```
├── src/
│   ├── config/          # Account mappings & portfolio definitions
│   ├── parsers/         # Input file parsers (Nortridge, DPV, Pier, etc.)
│   ├── journal_entries/ # JE generation engine
│   ├── reconciliation/  # Cross-system reconciliation
│   ├── qbo/             # QuickBooks Online API integration (MCP)
│   ├── reports/         # Output report generation
│   └── run_close.py     # Main orchestrator
├── data/
│   ├── sample/          # Sample input files (gitignored)
│   └── templates/       # Report templates
├── output/              # Generated reports & JEs (gitignored)
├── tests/
├── EULA.md
├── PRIVACY.md
└── docs/                # GitHub Pages (Intuit compliance)
```

## QuickBooks MCP Integration

The `src/qbo/` module provides MCP tools for:
- Pulling P&L, Balance Sheet, and Cash Flow reports
- Comparing QBO actuals vs. Nortridge portfolio metrics
- Forecasting revenue, charge-offs, and covenant compliance

### Intuit App Setup
- App Name: MCP Server
- App ID: 6e1ffcdf-5...
- Environment: Production
- Company: SET Financial Corporation

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env  # Add your QBO OAuth credentials
```

## Usage

```bash
python src/run_close.py --month 2026-01
```
