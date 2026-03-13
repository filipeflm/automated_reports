# Automated Sales Report Generator

A practical Python automation demo that reads raw sales data, validates it, computes key business metrics, and delivers a formatted Excel report, a plain-text summary, and two professional charts — all with a single command.

Built as a portfolio demonstration of back-office reporting automation for operations and analytics teams.

---

## What It Does

| Step | Description |
|------|-------------|
| 1 | Reads a CSV file containing fictional sales transactions |
| 2 | Validates and cleans the data (type enforcement, missing values) |
| 3 | Recalculates `total_value` to guarantee consistency |
| 4 | Computes business metrics: revenue totals, top products, top clients, and more |
| 5 | Exports a multi-sheet **Excel workbook** with formatted headers and currency values |
| 6 | Writes a concise **plain-text summary** for quick stakeholder review |
| 7 | Generates **two PNG charts**: revenue by region and top 5 products |

---

## Project Structure

```
demo_01_automated_reports/
├── input/
│   └── sales_data.csv          # Raw fictional sales data (80 rows)
├── output/                     # All generated files land here
│   ├── sales_summary.xlsx
│   ├── summary.txt
│   ├── revenue_by_region.png
│   └── top_products.png
├── script/
│   └── generate_report.py      # Main automation script
├── screenshots/                # Optional: add demo screenshots here
├── README.md
└── requirements.txt
```

---

## Input Data

**File:** `input/sales_data.csv`

| Column | Type | Description |
|-------------|---------|----------------------------------|
| date | date | Transaction date (YYYY-MM-DD) |
| client_name | string | Client / company name |
| product | string | Product name |
| category | string | Product category |
| region | string | Sales region |
| quantity | integer | Units sold |
| unit_price | float | Price per unit (USD) |
| total_value | float | quantity × unit_price (USD) |

---

## Output Files

### `sales_summary.xlsx`
A multi-sheet Excel workbook with the following sheets:

| Sheet | Contents |
|----------------------|-------------------------------------|
| Raw_Data | Full cleaned dataset |
| Summary | High-level KPI overview |
| Revenue_By_Region | Revenue breakdown per region |
| Revenue_By_Category | Revenue breakdown per category |
| Top_Products | Top 5 products by revenue |
| Top_Clients | Top 5 clients by revenue |

Formatting: bold headers, adjusted column widths, currency number format.

### `summary.txt`
A short business-style text report with total revenue, best region, best-selling product, and top client.

### `revenue_by_region.png`
Horizontal bar chart showing revenue per sales region.

### `top_products.png`
Vertical bar chart showing the top 5 products by total revenue.

---

## How to Run

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the script

```bash
python script/generate_report.py
```

All output files are saved to the `output/` folder automatically.

### Expected console output

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

---

## Business Value

This demo represents a common automation opportunity found in many organisations:

- **Operations teams** that manually copy-paste data into Excel each week
- **Finance and sales teams** that need recurring reports without engineering effort
- **Small and mid-size businesses** that need structured analytics from flat files

By automating this workflow, companies can reduce reporting time from hours to seconds, eliminate manual errors, and deliver consistent, audit-ready outputs on demand.

---

## Tech Stack

| Library | Purpose |
|------------|-------------------------------|
| pandas | Data loading, cleaning, metrics |
| openpyxl | Excel export and formatting |
| matplotlib | Chart generation |

Python 3.10+ recommended.

---

## Author

Built as a portfolio project to demonstrate practical data automation skills.
