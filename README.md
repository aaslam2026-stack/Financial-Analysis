# Finance Studio

A corporate finance and valuation web app with a clean, editorial design. Build financial models from your own statements.
Plain HTML, CSS and JavaScript, with no build step.

## How it works

**Build Excel Model** (main flow):
1. Upload your financial statements as **PDF**, Excel (.xlsx/.xls) or CSV.
2. Choose the models you want: DCF, Comparables, WACC, Monte Carlo, LBO, Loan Schedule, NPV / IRR, Break-even, Feasibility, Ratio Analysis, 3-Statement Forecast (or select all).
3. Download **one Excel workbook**: a Cover sheet, a Historicals sheet with your statements, and one sheet per model.

The workbook is a formatted, formula-driven financial model:

| Sheet | What it contains |
|---|---|
| Cover | Contents and colour legend |
| Dashboard | KPI tiles, 4 native Excel charts, financial summary and valuation summary |
| Assumptions | Forecast drivers by year, starting from 3-year historical averages |
| Income Statement | Revenue to net income, actual and forecast, with margins |
| Balance Sheet | Assets, liabilities and equity with a balance check |
| Cash Flow | Operating, investing and financing cash flows and free cash flow |
| DCF | Unlevered FCF from the statements, valuation bridge, implied multiples, EV and per-share sensitivities |
| Ratios | Profitability, liquidity, leverage, returns and efficiency (actual and forecast) |
| Model sheets | Scenarios, WACC, Comps, Monte Carlo, LBO, M&A, Loan, NPV / IRR, Break-even, Feasibility, Working Capital, Unit Economics, Runway (as chosen) |
| Accuracy | Forecast-vs-actual backtest with accuracy by metric |
| Formula Guide | Every key formula in plain English |
| Historicals | Your uploaded statements (the source for green links) |

- **Blue** cells are inputs. Change one and the whole model recalculates.
- **Green** cells link to the source data. **Black** cells are formulas. Shaded cells are key outputs.
- Every forecast line and key output has a **“How it’s calculated”** note beside it in plain English.

The **home page** has a search bar, model cards with spreadsheet previews and example results, sample companies for 8 sectors (Software, Retail, Industrial, Healthcare, Real Estate, Energy, Restaurants, Logistics), and models grouped by business stage. Each model card opens a step-by-step workspace (upload → assumptions & results → Excel download).

## Features

- **Models (18):** 3-Statement Model, DCF, Scenario Analysis, Comparables, WACC, Monte Carlo, LBO, M&A Accretion/Dilution, Loan Schedule, NPV / IRR, Break-even, Feasibility, Working Capital, Unit Economics & SaaS, Cash Runway, Ratio Analysis, Forecast Accuracy (backtest), Dashboard & Power BI.
- **Forecast accuracy:** every company gets a backtest: each recent year is forecast from the years before it and compared with the actual result (accuracy = 100% – average error; target 90%).
- **Indicators tab:** live interest rates, yield curve and economic indicators with what each means for a model, plus indicative benchmarks for 26 industries and a company-vs-industry comparison.
- **AI financial analyst:** a chat panel that answers using your company data, open model and market inputs.
- **Real company models:** 45 large companies across industries (Apple, NVIDIA, Amazon, Walmart, ExxonMobil, ...) or any US ticker, loaded from SEC filings; plus 16 industry sample companies.
- **Colour schemes:** Classic, Navy, Emerald, Burgundy, Slate and Dark.
- **Feasibility Report:** NPV, IRR, payback, sensitivity, a verdict and recommendations. You can export it to Excel or print it / save it as a PDF.
- **Power BI-style dashboard** with filters, KPI cards, a trend chart, margins, a profit bridge and a financial matrix. **Export for Power BI** downloads a tidy dataset for Power BI Desktop.
- **PDF reading:** text-based PDFs (exported from Excel, Word or accounting software) are read in the browser. Years are detected from the header row, negatives in (brackets) are handled, and "in thousands / millions" is detected automatically. Scanned image PDFs are not supported.
- **Samples:** `samples/sample-financials.pdf`, `samples/sample-financials.csv`, a full workbook `samples/models/full-financial-model.xlsx` and one workbook per model in `samples/models/`.

Data stays in the browser (localStorage). Nothing is uploaded to a server.

## Live features (Netlify Functions)

Three features run as Netlify Functions in `netlify/functions/`. They work when the site is deployed **from GitHub** (Netlify builds the functions); a drag-and-drop deploy publishes only the static pages.

| Feature | Function | Data source | Setting needed |
|---|---|---|---|
| AI financial analyst chat | `/api/chat` | Claude (`claude-opus-5-5`) via the Anthropic API | `ANTHROPIC_API_KEY` |
| Live rates & indicators | `/api/market` | FRED, Federal Reserve Bank of St. Louis (public CSV) | none |
| Real company models | `/api/company?ticker=AAPL` | SEC EDGAR XBRL company facts (10-K / 20-F) | `SEC_USER_AGENT` recommended, e.g. `Finance Studio you@example.com` |

Without the functions the site still works: the chat answers from a built-in guide, the Indicators tab shows industry benchmarks and lets you enter a risk-free rate, and models use uploaded or sample data.

The chat endpoint only accepts requests from the site itself and caps message sizes. Set a monthly spend limit for the API key in the Anthropic Console.

## Deploy to Netlify

1. In Netlify, choose **Add new site → Import an existing project**, connect GitHub and pick this repo (branch `main`).
2. Leave the build command empty. The publish directory (`.`) and functions directory are already set in `netlify.toml`.
3. Under **Site configuration → Environment variables**, add `ANTHROPIC_API_KEY` (for the chat) and `SEC_USER_AGENT` (for SEC data).
4. Deploy. Netlify installs `@anthropic-ai/sdk` from `package.json` and bundles the functions.

## Run locally

```bash
python3 -m http.server 8000
```

Then open http://localhost:8000.
