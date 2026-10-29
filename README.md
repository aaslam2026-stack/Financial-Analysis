# Finance Studio

A corporate finance and valuation web app with a clean, institutional banking design. Build financial models from your own statements.
Plain HTML, CSS and JavaScript, with no build step.

## How it works

**Build Excel Model** (main flow):
1. Upload your financial statements as **PDF**, Excel (.xlsx/.xls) or CSV.
2. Choose the models you want: DCF, Comparables, WACC, Monte Carlo, LBO, Loan Schedule, NPV / IRR, Break-even, Feasibility, Ratio Analysis, 3-Statement Forecast (or select all).
3. Download **one Excel workbook**: a Cover sheet, a Historicals sheet with your statements, and one sheet per model.

The workbook is a real, formula-driven model:
- **Blue** cells are inputs. Change one and the whole model recalculates.
- **Green** cells link to the Historicals sheet.
- **Black** cells are formulas. Shaded cells are key outputs.

The **Model Library** shows a detailed card for each model with example results. Clicking a model opens a step-by-step workspace (upload → assumptions & results → Excel download).

## Features

- **Models:** DCF (with sensitivity and scenarios), Comparables, WACC, Monte Carlo, LBO, Loan Schedule, NPV / IRR / MIRR / Payback, Break-even, Ratio Analysis, 3-Statement Forecast.
- **Feasibility Report:** NPV, IRR, payback, sensitivity, a verdict and recommendations. You can export it to Excel or print it / save it as a PDF.
- **Power BI-style dashboard** with filters, KPI cards, a trend chart, margins, a profit bridge and a financial matrix. **Export for Power BI** downloads a tidy dataset for Power BI Desktop.
- **PDF reading:** text-based PDFs (exported from Excel, Word or accounting software) are read in the browser. Years are detected from the header row, negatives in (brackets) are handled, and "in thousands / millions" is detected automatically. Scanned image PDFs are not supported.
- **Samples:** `samples/sample-financials.pdf`, `samples/sample-financials.csv`, a full workbook `samples/models/full-financial-model.xlsx` and one workbook per model in `samples/models/`.

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
