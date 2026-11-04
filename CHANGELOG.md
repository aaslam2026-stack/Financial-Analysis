# Changelog

## 2026-11-04 — Model Builder, tested finance engine, integrity checks

### Added
- **`lib/finance/finance.js`**: one module for the core finance formulas (NPV, IRR, PMT, WACC, linked 3-statement projection, DCF, integrity checks). It runs in the browser and in Node. The sign convention is documented at the top of the file.
- **Unit tests** (`tests/finance.test.js`, run with `npm test`): 17 tests with hand-checked answers, covering:
  - NPV, IRR, PMT and WACC;
  - balancing statements in annual, quarterly and monthly mode;
  - the revolver;
  - average-balance interest iteration;
  - Gordon growth and exit-multiple terminal values;
  - the mid-year convention;
  - valuation sanity warnings.
- **Guided Model Builder** (`#builder`, `builder.js`): Setup → Assumptions → Statements → Valuation → Checks → Export.
  - Annual, quarterly or monthly periods over 3–10 years.
  - Per-year drivers, including DSO, DIO and DPO, CapEx, D&A, tax and debt terms.
  - An optional revolver with minimum cash.
  - An interest switch between opening and average balances.
  - A WACC calculator (CAPM).
  - Gordon or exit-multiple terminal value, with a mid-period toggle.
  - An EV → equity bridge, a sensitivity grid and charts.
  - Click any statement figure to see its formula and inputs.
  - Thousands separators, negatives in parentheses, and units in actual, 000s or millions.
- **Nine integrity checks**, each shown as ✅ or ❌ in the app and on an Excel **Checks** sheet:
  - the balance sheet balances;
  - cash ties;
  - retained earnings roll forward;
  - net income links;
  - PP&E rolls forward;
  - debt rolls forward;
  - no negative cash;
  - circularity is controlled;
  - valuation sanity (terminal growth below WACC, terminal value at most 85% of EV, WACC within 4–20%).
- **Builder Excel export**: Cover, Assumptions, Income Statement, Balance Sheet, Cash Flow, DCF and Checks sheets.
  - Live formulas: blue inputs, black formulas, green links.
  - When the revolver is on, Excel's iterative calculation is switched on, so the average-balance interest loop solves.
- **Saved models** in the browser: up to 20 versions each, restore, JSON export and import (validated), and confirmation before any delete or overwrite.
- **Database schema for accounts** (`supabase/migrations/0001_initial_schema.sql`): users, models, assumptions, scenarios, model_versions and templates, all with row-level security. Not connected yet.
- Sample workbook: `samples/models/guided-builder-3-statement-dcf.xlsx`.

### Changed
- The disclaimer now reads "Models are for informational purposes only and are not financial advice." in the site footer and in Excel.
- The Excel formula evaluator supports circular references, but only when a workbook explicitly turns on iteration (at most 100 passes, tolerance 1e-9). Otherwise a circular reference is still an error.
- The site is easier to use on phones: the Model Builder is view-only there.
- Navigation: "Build Excel Model" is now "Excel Models"; "Model Builder" has been added.

### Verified
- `npm test`: 17 of 17 pass.
- The builder's Excel matches the app to within 0.00000002, across annual, quarterly and monthly models, with and without the revolver.
- LibreOffice recalculation of the builder workbooks: 6,504 formulas, 0 mismatches. The existing full workbook: 2,336 formulas, 0 mismatches.
- No console errors. On a 390 px phone screen the page has no horizontal scroll.
