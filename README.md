# Glass Finance Studio

A simple, Apple-style "glass" web app for building financial models from your own statements.
Plain HTML, CSS and JavaScript, with no build step.

## Features

- **Upload** financial statements as CSV or Excel (.xlsx). Line items go in the first column and years across the top row. A template is in `samples/sample-financials.csv`.
- **Dashboard** with KPIs, revenue, EBITDA and net income trends, margins, balance sheet mix and key ratios.
- **Models** (inputs pre-filled from your data, and grouped by goal):
  DCF (with sensitivity and scenarios), NPV / IRR / MIRR / Payback, LBO, WACC, Comparables,
  Monte Carlo, Break-even, Ratio Analysis, 3-Statement Forecast, Loan Schedule.
- **Feasibility Report**: NPV, IRR, payback, sensitivity, a climate-risk score for the location, a verdict and recommendations. You can print it or save it as a PDF.
- **Weather** from the [Open-Meteo API](https://open-meteo.com/) (free, no API key): a 7-day outlook on the dashboard and 12 months of weather history in the feasibility report.

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
