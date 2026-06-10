# SAM Factsheet Updater — Web App

A simple local web tool for updating quarterly factsheets.
Open it in your browser, upload three files, enter the new date, done.

---

## Setup (one time only)

**1. Install Python**
Download from https://python.org if you don't have it.

**2. Install required packages**
Open Terminal (Mac) or Command Prompt (Windows) and run:

```
pip install flask python-pptx pandas openpyxl lxml numpy
```

---

## Running the app

**Mac / Linux:**
Double-click `start.command`
OR open Terminal, navigate to this folder, and run:
```
python3 app.py
```

**Windows:**
Double-click `start.bat`
OR open Command Prompt, navigate to this folder, and run:
```
python app.py
```

Then open your browser to: **http://localhost:5000**

---

## What files to upload each quarter

For each of the five strategies, run the app once with that strategy's files:

| Field | File to select |
|---|---|
| **Factsheet (.pptx)** | The PowerPoint factsheet for that strategy |
| **Holdings (.xlsx)** | The holdings export for that strategy from your portfolio system |
| **Model (.xlsx)** | The model file listing approved tickers for that strategy |

Then enter:
- **Current date in file** — the quarter-end date currently shown in the factsheet (e.g. `December 31, 2025`)
- **New date** — the new quarter-end date (e.g. `March 31, 2026`)

Click **Update Factsheet**, then **Download** when it finishes.

---

## What gets updated automatically

- All three date fields in the factsheet (header, "Period Ending", "Sector Percentages as of")
- Top 10 Holdings (all 10 names)
- Pie chart sector weights and legend

**What is NOT changed:** The performance table — update that manually each quarter.

---

## Required file formats

**Holdings file** must have these columns:
- `Ticker`
- `Asset Category` (values: `Equity` or `Money Market`)
- `Total Value`
- `ProductName`

**Model file** must have a `Ticker` column listing the approved stocks.

---

## Adding new tickers

If a new stock is added to a model, open `app.py` in any text editor and find the `SECTOR_MAP` section near the top. Add a line:

```python
"TICKER": "Sector Name",
```

Sector names used: Technology, Industrials, Financials, Communications, Healthcare,
Utilities, Consumer Staples, Energy, Materials, Cons. Discretionary

---

## Troubleshooting

**"No matching stocks found"** — Check that the holdings file has the correct column names listed above.

**"Could not find embedded chart workbook"** — The factsheet PPTX may be structured differently from the WWE template. The script expects a pie chart with an embedded Excel workbook inside.

**"No date replacements found"** — Make sure the "Current date in file" field exactly matches what's written in the PowerPoint (e.g. spelling, spacing, year).
