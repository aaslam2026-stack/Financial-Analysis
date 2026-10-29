# Finance Studio

A corporate finance and valuation web app with a clean, institutional banking design. Build financial models from your own statements.
Plain HTML, CSS and JavaScript, with no build step.

## How it works

1. **Model Library** (home page): browse 12 detailed model cards. Each one shows example results for the sample company, the key outputs, the data needed, what it is best for, and a sample Excel file.
2. **Click a model** to open its workspace:
   - **Step 1:** upload your financial statements as **PDF**, Excel (.xlsx/.xls) or CSV, use the sample company, or continue with manual assumptions.
   - **Step 2:** review the assumptions (pre-filled from your statements) and the results.
   - **Step 3:** download everything to Excel.

## Features

- **Models:** DCF (with sensitivity and scenarios), Comparables, WACC, Monte Carlo, LBO, Loan Schedule, NPV / IRR / MIRR / Payback, Break-even, Ratio Analysis, 3-Statement Forecast.
- **Feasibility Report:** NPV, IRR, payback, sensitivity, a verdict and recommendations. You can export it to Excel or print it / save it as a PDF.
- **Power BI-style dashboard** with filters, KPI cards, a trend chart, margins, a profit bridge and a financial matrix. **Export for Power BI** downloads a tidy dataset for Power BI Desktop.
- **PDF reading:** text-based PDFs (exported from Excel, Word or accounting software) are read in the browser. Years are detected from the header row, negatives in (brackets) are handled, and "in thousands / millions" is detected automatically. Scanned image PDFs are not supported.
- **Samples:** `samples/sample-financials.pdf`, `samples/sample-financials.csv` and one Excel output per model in `samples/models/`.

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
