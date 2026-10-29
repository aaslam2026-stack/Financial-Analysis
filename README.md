# Glass Finance Studio

A simple, Apple-style "glass" web app for building financial models from your own statements.
Plain HTML, CSS and JavaScript, with no build step.

## Features

- **Upload** financial statements as CSV or Excel (.xlsx). Line items go in the first column and years across the top row. A template is in `samples/sample-financials.csv`.
- **Power BI-style dashboard** with year and metric slicers, KPI cards, a trend chart with YoY growth (click a bar to focus on that year), margins, a profit bridge (waterfall), a balance sheet chart, a financial matrix and ratios.
  - **Export for Power BI**: downloads a tidy Excel dataset (`FactFinancials` and `Ratios` sheets) that loads straight into Power BI Desktop (Home → Get data → Excel workbook).
- **Models**, picked from a dropdown, with inputs pre-filled from your data:
  - Valuation: DCF (with sensitivity and scenarios), Comparables, WACC, Monte Carlo
  - Deals & financing: LBO, Loan Schedule
  - Projects & investments: NPV / IRR / MIRR / Payback, Break-even
  - Performance: Ratio Analysis, 3-Statement Forecast
  - **Download Excel**: exports the model's inputs, key results and every table (as real numbers) plus your source data.
- **Feasibility Report**: NPV, IRR, payback, sensitivity analysis, a verdict and recommendations. You can download it to Excel or print it / save it as a PDF.

Data stays in the browser (localStorage). Nothing is uploaded to a server.

## Deploy to Netlify

1. In Netlify, choose **Add new site → Import from Git**, and pick this repo (branch `main`).
2. Leave the build command empty. The publish directory is `.` (already set in `netlify.toml`).
3. Deploy.

Or drag and drop the project folder onto <https://app.netlify.com/drop>.

## Run locally

```bash
python3 -m http.server 8000
```

Then open http://localhost:8000.
