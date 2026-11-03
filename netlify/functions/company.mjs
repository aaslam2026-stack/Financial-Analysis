// Real company financials from SEC EDGAR (annual 10-K / 20-F XBRL data), for any US-listed company.
// GET /api/company?ticker=AAPL -> { name, ticker, cik, years, rows, source, notes }
// rows use the same "Line item, year, year..." layout as an uploaded spreadsheet.
// The SEC asks automated clients to identify themselves: set SEC_USER_AGENT, e.g. "Finance Studio you@example.com".

const UA = process.env.SEC_USER_AGENT || 'FinanceStudio/1.0 (financial-model-builder; set SEC_USER_AGENT)';
const ANNUAL_FORMS = new Set(['10-K', '10-K/A', '10-KT', '20-F', '20-F/A', '40-F']);

// XBRL tags in priority order; the first tag with a value for a period wins
export const TAGS = {
  revenue: ['RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues', 'SalesRevenueNet', 'RevenueFromContractWithCustomerIncludingAssessedTax', 'SalesRevenueGoodsNet', 'RevenuesNetOfInterestExpense'],
  cogs: ['CostOfGoodsAndServicesSold', 'CostOfRevenue', 'CostOfGoodsSold', 'CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization', 'CostOfServices'],
  opInc: ['OperatingIncomeLoss'],
  da: ['DepreciationDepletionAndAmortization', 'DepreciationAndAmortization', 'DepreciationAmortizationAndAccretionNet', 'Depreciation'],
  interest: ['InterestExpense', 'InterestExpenseNonoperating', 'InterestExpenseDebt', 'InterestPaidNet'],
  ni: ['NetIncomeLoss', 'ProfitLoss', 'NetIncomeLossAvailableToCommonStockholdersBasic'],
  capex: ['PaymentsToAcquirePropertyPlantAndEquipment', 'PaymentsToAcquireProductiveAssets', 'PaymentsForCapitalImprovements'],
  ocf: ['NetCashProvidedByUsedInOperatingActivities', 'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations'],
  cash: ['CashAndCashEquivalentsAtCarryingValue', 'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents', 'Cash'],
  ca: ['AssetsCurrent'],
  ta: ['Assets'],
  cl: ['LiabilitiesCurrent'],
  tl: ['Liabilities'],
  eq: ['StockholdersEquity', 'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest'],
  ltdNon: ['LongTermDebtNoncurrent', 'LongTermDebtAndCapitalLeaseObligations'],
  ltdTotal: ['LongTermDebt', 'DebtLongtermAndShorttermCombinedAmount'],
  ltdCur: ['LongTermDebtCurrent', 'DebtCurrent', 'LongTermDebtAndCapitalLeaseObligationsCurrent'],
  stDebt: ['CommercialPaper', 'ShortTermBorrowings', 'OtherShortTermBorrowings'],
};
const DURATION = new Set(['revenue', 'cogs', 'opInc', 'da', 'interest', 'ni', 'capex', 'ocf']);
const days = (a, b) => (new Date(b) - new Date(a)) / 864e5;

// Map of period end date -> value for one concept, from annual filings only.
// Durations must cover about a year; balance-sheet items are point-in-time (no start date).
export function annualSeries(facts, tags, duration) {
  const out = new Map();
  for (const tag of tags) {
    const units = facts?.['us-gaap']?.[tag]?.units?.USD || facts?.['ifrs-full']?.[tag]?.units?.USD;
    if (!units) continue;
    const best = new Map();
    for (const f of units) {
      if (!ANNUAL_FORMS.has(f.form) || f.val == null) continue;
      if (duration ? !f.start || days(f.start, f.end) < 350 || days(f.start, f.end) > 380 : !!f.start) continue;
      const cur = best.get(f.end);
      if (!cur || f.filed > cur.filed) best.set(f.end, f);
    }
    for (const [end, f] of best) if (!out.has(end)) out.set(end, f.val);
  }
  return out;
}

// Value at a period end, allowing balance-sheet dates a few days apart
function at(series, end) {
  if (series.has(end)) return series.get(end);
  for (const [d, v] of series) if (Math.abs(days(d, end)) <= 10) return v;
  return null;
}

export function extractStatements(facts, years = 5) {
  const S = Object.fromEntries(Object.entries(TAGS).map(([k, tags]) => [k, annualSeries(facts, tags, DURATION.has(k))]));
  const ends = [...S.revenue.keys()].sort().slice(-years);
  if (!ends.length) throw new Error('No annual revenue was found in this company’s filings.');
  const col = (k) => ends.map((e) => at(S[k], e));
  const rev = col('revenue'), cogs = col('cogs'), opInc = col('opInc'), da = col('da').map((v) => (v == null ? null : Math.abs(v)));
  const ta = col('ta'), eq = col('eq'), tl = col('tl');
  // Operating expenses excluding D&A, so the model can rebuild EBITDA and EBIT exactly:
  // revenue – COGS – opex – D&A = operating income
  const opex = ends.map((_, i) => (rev[i] != null && opInc[i] != null ? rev[i] - (cogs[i] ?? 0) - opInc[i] - (da[i] ?? 0) : null));
  const debt = ends.map((e) => {
    const non = at(S.ltdNon, e), total = at(S.ltdTotal, e), cur = at(S.ltdCur, e), st = at(S.stDebt, e);
    if (non == null && total == null && cur == null && st == null) return null;
    return (non != null ? non + (cur ?? 0) : total ?? cur ?? 0) + (st ?? 0);
  });
  const label = (e) => String(new Date(e).getUTCFullYear());
  const rows = [
    ['Line Item', ...ends.map(label)],
    ['Revenue', ...rev], ['Cost of Goods Sold', ...cogs], ['Operating Expenses', ...opex], ['Depreciation & Amortization', ...da],
    ['Operating Income', ...opInc], ['Interest Expense', ...col('interest')], ['Net Income', ...col('ni')],
    ['Cash', ...col('cash')], ['Current Assets', ...col('ca')], ['Total Assets', ...ta], ['Current Liabilities', ...col('cl')],
    ['Total Liabilities', ...ends.map((_, i) => tl[i] ?? (ta[i] != null && eq[i] != null ? ta[i] - eq[i] : null))],
    ['Total Debt', ...debt], ['Equity', ...eq], ['CapEx', ...col('capex').map((v) => (v == null ? null : Math.abs(v)))], ['Operating Cash Flow', ...col('ocf')],
  ].map((r) => r.map((v) => (v == null ? '' : v)));
  return { years: ends.map(label), periodEnds: ends, rows };
}

let tickerMap = null;
async function lookupCik(ticker) {
  if (!tickerMap) {
    const res = await fetch('https://www.sec.gov/files/company_tickers.json', { headers: { 'User-Agent': UA }, signal: AbortSignal.timeout(8000) });
    if (!res.ok) throw new Error(`SEC ticker list unavailable (HTTP ${res.status})`);
    tickerMap = new Map(Object.values(await res.json()).map((c) => [c.ticker.toUpperCase(), c]));
  }
  return tickerMap.get(ticker) || tickerMap.get(ticker.replace('.', '-')) || null;
}

export default async (req) => {
  const ticker = (new URL(req.url).searchParams.get('ticker') || '').trim().toUpperCase();
  if (!/^[A-Z0-9.\-]{1,10}$/.test(ticker)) return Response.json({ error: 'Enter a valid ticker symbol, for example AAPL.' }, { status: 400 });
  try {
    const c = await lookupCik(ticker);
    if (!c) return Response.json({ error: `No SEC filer found for ticker ${ticker}.` }, { status: 404 });
    const cik = String(c.cik_str).padStart(10, '0');
    const res = await fetch(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, { headers: { 'User-Agent': UA }, signal: AbortSignal.timeout(12000) });
    if (!res.ok) throw new Error(`SEC data unavailable (HTTP ${res.status})`);
    const data = await res.json();
    const st = extractStatements(data.facts);
    return Response.json({
      name: data.entityName || c.title, ticker, cik, years: st.years, periodEnds: st.periodEnds, rows: st.rows,
      source: `SEC EDGAR XBRL company facts (annual 10-K/20-F filings), CIK ${cik}`,
      notes: 'Figures are as reported in annual filings, in US dollars. Fiscal years are labelled by the calendar year they end in. Operating expenses exclude D&A so EBITDA rebuilds exactly; total debt combines long-term, current and short-term borrowings (leases may be excluded).',
    }, { headers: { 'Cache-Control': 'public, max-age=3600', 'Netlify-CDN-Cache-Control': 'public, durable, max-age=86400, stale-while-revalidate=604800' } });
  } catch (err) {
    return Response.json({ error: err.message || 'SEC data unavailable.' }, { status: 502 });
  }
};

export const config = { path: '/api/company' };
