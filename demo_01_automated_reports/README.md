# Automated Reports

A portfolio-ready automation demo that transforms raw sales CSV data into executive-style reports, charts and structured business outputs.

![Automated Reports Hero](screenshots/hero.png)

---

## What it does

A single Python script reads a raw sales CSV file, validates the data, computes key business metrics and automatically generates three types of output:

- A formatted **multi-sheet Excel workbook** with bold headers, currency formatting and clean layout
- A short **plain-text executive summary** with the most important KPIs
- Two **PNG charts** ready to embed in presentations or dashboards

Everything runs with one command. No manual steps, no copy-paste.

---

## Business use case

This demo simulates a real operational workflow where a company receives raw sales data in CSV format and needs an automated process to generate clean, decision-ready reporting outputs.

It models the kind of task that operations, finance and back-office teams handle on a weekly or monthly basis — pulling data from a source, cleaning it, computing metrics and preparing reports for stakeholders.

---

## Features

- Loads and validates raw CSV data using pandas
- Recalculates `total_value` to ensure data consistency
- Detects and handles missing values with clear console logging
- Computes seven business metrics: total revenue, units sold, average order value, revenue by region, revenue by category, top 5 products and top 5 clients
- Exports a 6-sheet Excel report with styled headers and number formatting
- Writes a concise plain-text business summary
- Generates two production-quality bar and line charts as PNG files
- Includes an interactive HTML dashboard for visual presentation

---

## Tech stack

| Tool | Purpose |
|---|---|
| Python 3.10+ | Core scripting language |
| pandas | Data loading, cleaning and aggregation |
| openpyxl | Excel report creation and styling |
| matplotlib | Chart generation |
| Chart.js | Interactive dashboard charts |

---

## Project structure

```
demo_01_automated_reports/
├── input/
│   └── sales_data.csv          # Raw fictional sales data (80 rows)
├── output/
│   ├── sales_summary.xlsx      # Multi-sheet Excel report
│   ├── summary.txt             # Plain-text executive summary
│   ├── revenue_by_region.png   # Bar chart by region
│   └── top_products.png        # Bar chart of top 5 products
├── script/
│   └── generate_report.py      # Main automation script
├── screenshots/
│   └── hero.png
├── dashboard.html              # Interactive visual dashboard
├── README.md
└── requirements.txt
```

---

## How to run

**1. Clone or download the project**

```bash
cd demo_01_automated_reports
```

**2. Install dependencies**

```bash
pip install -r requirements.txt
```

**3. Run the report**

```bash
python script/generate_report.py
```

All output files are saved to the `output/` folder automatically.

**Expected console output:**

```
========================================================
  Automated Sales Report Generator — Starting
========================================================

[INFO] Loaded 80 rows from 'sales_data.csv'.
[INFO] Data validation passed — no rows removed.
[INFO] Metrics calculated successfully.
[INFO] Excel report saved → sales_summary.xlsx
[INFO] Text summary saved  → summary.txt
[INFO] Chart saved         → revenue_by_region.png
[INFO] Chart saved         → top_products.png

========================================================
  All outputs saved to: output/
========================================================
```

**4. Open the dashboard (optional)**

Open `dashboard.html` directly in any browser to see an interactive visual version of the report.

---

## Outputs generated

### `sales_summary.xlsx`

A clean Excel workbook with six sheets:

| Sheet | Contents |
|---|---|
| Raw_Data | Full cleaned dataset |
| Summary | High-level KPI overview |
| Revenue_By_Region | Revenue per sales region |
| Revenue_By_Category | Revenue per product category |
| Top_Products | Top 5 products by revenue |
| Top_Clients | Top 5 clients by revenue |

### `summary.txt`

```
========================================================
         AUTOMATED SALES REPORT — SUMMARY
========================================================

  Total Revenue        : $   135,634.86
  Total Units Sold     :           754
  Average Order Value  : $     1,695.44

  --- Regional Performance ---
  Best Region          : North  ($38,350.56)

  --- Product Performance ---
  Best-Selling Product : Laptop Pro 15  ($32,499.75)

  --- Client Performance ---
  Top Client           : Apex Solutions  ($17,023.96)

========================================================
```

### Charts

Two PNG charts are generated automatically:

- `revenue_by_region.png` — Horizontal bar chart comparing revenue across the four sales regions
- `top_products.png` — Vertical bar chart showing the five highest-revenue products

---

## Screenshots

**Dashboard — KPIs and automation pipeline**

![Hero Screenshot](screenshots/hero.png)

---

## Why this matters for businesses

Businesses often waste hours manually preparing reports from spreadsheets. A team member downloads data, opens Excel, builds pivot tables, formats cells, copies values into a summary slide and sends it by email — every single week.

This demo shows how automation can reduce repetitive work, improve reporting consistency and provide faster visibility into revenue, clients, products and regional performance.

The same approach scales to real scenarios:

- Weekly sales reports emailed automatically to management
- Monthly finance summaries pulled from ERP exports
- Regional performance dashboards refreshed on a schedule
- Any recurring report built from flat files or database exports

The script is simple, readable and easy to adapt. No infrastructure required. No database. No web server. Just Python and a CSV file.

---

*Built as a portfolio project · Demo 01 · Automated Reports*
