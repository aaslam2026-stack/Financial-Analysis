/* Finance Studio — guided Model Builder.
   Setup → Assumptions → Statements → Valuation → Checks → Export, built entirely on the tested lib/finance
   module (window.Finance). Models are saved in this browser with version history and can be exported as JSON.
   The Excel export writes the same logic as live formulas, including a Checks sheet. Depends on app.js
   (state, S, esc, makeChart, ds, COLORS, marketRf) and excel.js (addSheet, evaluateBook, bookToBuffer). */
(function () {
  'use strict';
  const F = window.Finance;
  const MODELS_KEY = 'fs-models-v1';
  const MAX_VERSIONS = 20;
  const STEPS = ['Setup', 'Assumptions', 'Statements', 'Valuation', 'Checks', 'Export'];
  const UNITS = { actual: { div: 1, dp: 0, label: '' }, thousands: { div: 1e3, dp: 0, label: '000s' }, millions: { div: 1e6, dp: 1, label: 'millions' } };

  // Per-year drivers: key, label, kind (pct | days | amount), allowed range in display units, help text
  const DRIVERS = [
    ['growth', 'Revenue growth', 'pct', [-90, 300], 'Annual growth; quarterly and monthly periods compound to this rate'],
    ['grossMargin', 'Gross margin', 'pct', [-100, 100], 'Gross profit ÷ revenue'],
    ['opexPct', 'Operating expenses % of revenue', 'pct', [0, 200], 'SG&A and other operating costs, excluding D&A'],
    ['daPct', 'Depreciation % of revenue', 'pct', [0, 100], 'D&A charged against PP&E'],
    ['capexPct', 'CapEx % of revenue', 'pct', [0, 100], 'Cash spent on property, plant and equipment'],
    ['taxRate', 'Tax rate', 'pct', [0, 60], 'Applied to profit before tax; no tax credit on losses'],
    ['dso', 'Receivable days (DSO)', 'days', [0, 730], 'Receivables = revenue × DSO ÷ days in period'],
    ['dio', 'Inventory days (DIO)', 'days', [0, 730], 'Inventory = COGS × DIO ÷ days in period'],
    ['dpo', 'Payable days (DPO)', 'days', [0, 730], 'Payables = COGS × DPO ÷ days in period'],
    ['otherCurrentAssetsPct', 'Other current assets % of revenue', 'pct', [0, 100], 'Prepayments and other operating assets'],
    ['otherCurrentLiabilitiesPct', 'Other current liabilities % of revenue', 'pct', [0, 100], 'Accruals and other operating liabilities'],
    ['payoutRatio', 'Dividend payout % of net income', 'pct', [0, 100], 'Paid only when net income is positive'],
    ['interestRate', 'Interest rate on term debt', 'pct', [0, 40], 'Annual rate'],
    ['debtDrawdown', 'New term debt drawn', 'amount', [0, 1e15], 'Per year'],
    ['debtRepayment', 'Scheduled debt repayment', 'amount', [0, 1e15], 'Per year, capped at the outstanding balance'],
  ];
  const OPENING = [
    ['revenue', 'Last actual annual revenue', 'Income statement'],
    ['cash', 'Cash', 'Assets'], ['receivables', 'Accounts receivable', 'Assets'], ['inventory', 'Inventory', 'Assets'],
    ['otherCurrentAssets', 'Other current assets', 'Assets'], ['ppe', 'PP&E and other non-current assets', 'Assets'],
    ['payables', 'Accounts payable', 'Liabilities'], ['otherCurrentLiabilities', 'Other current liabilities', 'Liabilities'],
    ['termDebt', 'Term debt', 'Liabilities'], ['revolver', 'Revolver drawn', 'Liabilities'], ['otherLiabilities', 'Other non-current liabilities', 'Liabilities'],
    ['contributedCapital', 'Contributed capital', 'Equity'], ['retainedEarnings', 'Retained earnings', 'Equity'],
  ];

  let bm = null;          // the model being edited
  let result = null;      // { proj, annual, val, wacc, checks }
  let step = 0, statementView = 'annual', savedList = [];
  const readOnly = () => window.matchMedia('(max-width: 700px)').matches;

  /* ---------- Model creation ---------- */
  const fill = (v, n) => Array.from({ length: n }, () => v);
  function blankModel() {
    const n = 5;
    return {
      schema: 'finance-studio-model', version: 1, id: 'm' + Date.now().toString(36), name: 'Example Co forecast', company: 'Example Co',
      currency: typeof currency === 'string' ? currency : 'USD', units: 'thousands', dataKey: null, periodType: 'annual', years: n, firstYear: new Date().getFullYear(),
      assumptions: { growth: fill(0.08, n), grossMargin: fill(0.42, n), opexPct: fill(0.22, n), daPct: fill(0.03, n), capexPct: fill(0.04, n), taxRate: fill(0.25, n),
        dso: fill(45, n), dio: fill(50, n), dpo: fill(40, n), otherCurrentAssetsPct: fill(0.02, n), otherCurrentLiabilitiesPct: fill(0.04, n), payoutRatio: fill(0, n),
        interestRate: fill(0.06, n), debtDrawdown: fill(0, n), debtRepayment: fill(2e6, n) },
      opening: { revenue: 50e6, cash: 5e6, receivables: 6e6, inventory: 4e6, otherCurrentAssets: 1e6, ppe: 20e6, payables: 3e6, otherCurrentLiabilities: 2e6,
        termDebt: 12e6, revolver: 0, otherLiabilities: 1e6, contributedCapital: 10e6, retainedEarnings: 8e6 },
      settings: { interestMethod: 'opening', revolver: { enabled: false, limit: 10e6, interestRate: 0.075, minimumCash: 2e6 } },
      valuation: { useCalculator: true, wacc: 0.09, riskFree: (window.marketRf?.()?.value ?? 4.2) / 100, beta: 1.1, marketPremium: 0.055, sizePremium: 0,
        preTaxCostOfDebt: 0.065, equityValue: 60e6, debtValue: 12e6, terminalGrowth: 0.025, method: 'gordon', exitMultiple: 10, midYear: true, shares: 10e6 },
      source: 'Example figures',
    };
  }

  // Start from the statements loaded in the app (last actual year as the opening balance sheet)
  function modelFromData() {
    const m = blankModel();
    if (!(typeof hasData === "function" && hasData())) return m;
    const li = state.years.length - 1, at = (k) => { const v = S(k)?.[li]; return v == null ? null : Math.abs(v); };
    const dr = deriveDrivers(), op = openingBS(), n = m.years;
    const rev = op.rev, cogs = at('cogs') ?? rev * (1 - dr.gm);
    const ar = at('receivables'), inv = at('inventory'), ap = at('payables');
    const nonCashCA = Math.max(0, op.ca - op.cash);
    const receivables = ar ?? nonCashCA * 0.55, inventory = inv ?? nonCashCA * 0.35;
    const otherCurrentAssets = Math.max(0, nonCashCA - receivables - inventory);
    const payables = ap ?? op.cl * 0.5, otherCurrentLiabilities = Math.max(0, op.cl - payables);
    const ppe = Math.max(0, op.ta - op.cash - receivables - inventory - otherCurrentAssets);
    const re = S('retainedEarnings')?.[li] ?? null;
    const equity = op.equity, contributedCapital = re != null ? equity - re : equity, retainedEarnings = re ?? 0;
    const totalAssets = op.cash + receivables + inventory + otherCurrentAssets + ppe;
    // Other liabilities is the balancing item, so the opening balance sheet balances exactly
    const otherLiabilities = totalAssets - payables - otherCurrentLiabilities - op.debt - equity;
    const yr = parseInt(String(state.years[li]).match(/(19|20)\d{2}/)?.[0]) || new Date().getFullYear() - 1;
    Object.assign(m, { dataKey: dataKey(), name: `${state.company} forecast`, company: state.company, firstYear: yr + 1, source: `${state.company} statements, ${state.years[0]}–${state.years[li]}` });
    m.opening = { revenue: rev, cash: op.cash, receivables, inventory, otherCurrentAssets, ppe, payables, otherCurrentLiabilities, termDebt: op.debt, revolver: 0,
      otherLiabilities, contributedCapital, retainedEarnings };
    const days = (bal, base) => (bal != null && base ? Math.round((bal / base) * 365) : null);
    Object.assign(m.assumptions, {
      growth: fill(dr.growth, n), grossMargin: fill(dr.gm, n), opexPct: fill(dr.opex, n), daPct: fill(dr.da, n), capexPct: fill(dr.capex, n), taxRate: fill(dr.tax, n),
      dso: fill(days(receivables, rev) ?? 45, n), dio: fill(days(inventory, cogs) ?? 0, n), dpo: fill(days(payables, cogs) ?? 30, n),
      otherCurrentAssetsPct: fill(+(otherCurrentAssets / rev).toFixed(3), n), otherCurrentLiabilitiesPct: fill(+(otherCurrentLiabilities / rev).toFixed(3), n),
      interestRate: fill(dr.rate, n), debtRepayment: fill(dr.repay, n), debtDrawdown: fill(0, n), payoutRatio: fill(0, n),
    });
    m.units = rev >= 5e9 ? 'millions' : rev >= 5e6 ? 'thousands' : 'actual';
    m.settings.revolver = { enabled: false, limit: Math.round(rev * 0.1), interestRate: 0.075, minimumCash: Math.round(rev * 0.03) };
    m.valuation.debtValue = op.debt;
    m.valuation.equityValue = Math.max(op.equity, 1);
    return m;
  }

  // Keep per-year arrays the same length as the forecast (extend with the last value)
  function resize(m) {
    m.years = Math.min(10, Math.max(3, Math.round(m.years) || 5));
    Object.keys(m.assumptions).forEach((k) => {
      const a = Array.isArray(m.assumptions[k]) ? m.assumptions[k] : [m.assumptions[k]];
      m.assumptions[k] = Array.from({ length: m.years }, (_, i) => a[Math.min(i, a.length - 1)]);
    });
  }

  /* ---------- Calculation ---------- */
  function waccOf(v, tax) {
    if (!v.useCalculator) return { wacc: v.wacc, calc: null };
    const calc = F.calculateWACC({ riskFree: v.riskFree, beta: v.beta, marketPremium: v.marketPremium, sizePremium: v.sizePremium, preTaxCostOfDebt: v.preTaxCostOfDebt,
      taxRate: tax, equityValue: v.equityValue, debtValue: v.debtValue });
    return { wacc: calc.wacc, calc };
  }
  function valuationFor(proj, m, wacc, overrides = {}) {
    const o = m.opening, v = m.valuation;
    return F.valueDCF(proj, { wacc, terminalGrowth: v.terminalGrowth, method: v.method, exitMultiple: v.exitMultiple, midYear: v.midYear,
      netDebt: o.termDebt + (o.revolver || 0) - o.cash, shares: v.shares, taxRates: m.assumptions.taxRate, ...overrides });
  }
  function compute() {
    resize(bm);
    const proj = F.projectStatements(bm.assumptions, bm.opening, { periodType: bm.periodType, years: bm.years, interestMethod: bm.settings.interestMethod, revolver: bm.settings.revolver });
    const w = waccOf(bm.valuation, bm.assumptions.taxRate[0]);
    const val = valuationFor(proj, bm, w.wacc);
    const checks = F.runIntegrityChecks(proj, val.error ? { tvShare: null } : val, w.wacc, bm.valuation.terminalGrowth);
    result = { proj, annual: F.annualise(proj), val, wacc: w.wacc, waccCalc: w.calc, checks };
    return result;
  }

  /* ---------- Formatting: thousands separators, negatives in parentheses, chosen units ---------- */
  const unit = () => UNITS[bm.units] || UNITS.actual;
  function num(v, dp = unit().dp, div = unit().div) {
    if (v == null || !isFinite(v)) return '—';
    const x = v / div;
    if (Math.abs(x) < 0.5 * Math.pow(10, -dp)) return '–';
    const s = Math.abs(x).toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp });
    return x < 0 ? `(${s})` : s;
  }
  const money = (v) => num(v);
  const full = (v) => `${bm.currency} ${num(v, 0, 1)}`;
  const pc = (v, d = 1) => (v == null || !isFinite(v) ? '—' : v < 0 ? `(${Math.abs(v * 100).toFixed(d)}%)` : `${(v * 100).toFixed(d)}%`);
  const unitLabel = () => `${bm.currency}${unit().label ? ' ' + unit().label : ''}`;
  const periodLabel = (i) => {
    const k = result.proj.periodsPerYear, y = bm.firstYear + Math.floor(i / k), q = (i % k) + 1;
    return k === 1 ? `${y}E` : k === 4 ? `Q${q} ${y}` : `${new Date(Date.UTC(y, q - 1, 1)).toLocaleString('en-US', { month: 'short', timeZone: 'UTC' })} ${y}`;
  };
  const yearLabel = (y) => `${bm.firstYear + y}E`;

  /* ---------- Traceability: every statement figure explains its formula and inputs ---------- */
  function prevOf(i) {
    const p = result.proj, o = bm.opening;
    return i > 0 ? p.periods[i - 1] : { ...o, revenue: o.revenue / p.periodsPerYear, revolver: o.revolver || 0 };
  }
  const drv = (key, i) => { const k = result.proj.periodsPerYear, a = bm.assumptions[key]; return a[Math.min(Math.floor(i / k), a.length - 1)]; };
  function trace(key, i) {
    const p = result.proj.periods[i], q = prevOf(i), k = result.proj.periodsPerYear, d = 365 / k, m = money, rv = result.proj.revolver;
    const T = {
      revenue: ['Prior-period revenue × (1 + growth for the period)', [['Prior-period revenue', m(q.revenue)], ['Annual growth', pc(drv('growth', i))], ['Growth per period', pc(F.periodRate(drv('growth', i), k), 3)]]],
      cogs: ['Revenue × (1 – gross margin)', [['Revenue', m(p.revenue)], ['Gross margin', pc(drv('grossMargin', i))]]],
      grossProfit: ['Revenue – cost of goods sold', [['Revenue', m(p.revenue)], ['COGS', m(p.cogs)]]],
      opex: ['Revenue × operating expenses %', [['Revenue', m(p.revenue)], ['Opex %', pc(drv('opexPct', i))]]],
      ebitda: ['Gross profit – operating expenses', [['Gross profit', m(p.grossProfit)], ['Operating expenses', m(p.opex)]]],
      da: ['Revenue × D&A %', [['Revenue', m(p.revenue)], ['D&A %', pc(drv('daPct', i))]]],
      ebit: ['EBITDA – D&A', [['EBITDA', m(p.ebitda)], ['D&A', m(p.da)]]],
      interest: [result.proj.interestMethod === 'average' ? 'Average (opening + closing) debt × rate ÷ periods per year' : 'Opening debt × rate ÷ periods per year',
        [['Opening term debt', m(q.termDebt)], ['Closing term debt', m(p.termDebt)], ['Term debt rate', pc(drv('interestRate', i))], ['Opening revolver', m(q.revolver)], ['Closing revolver', m(p.revolver)], ['Revolver rate', pc(rv.interestRate)], ['Periods per year', k]]],
      pbt: ['EBIT – interest', [['EBIT', m(p.ebit)], ['Interest', m(p.interest)]]],
      tax: ['MAX(0, profit before tax × tax rate)', [['Profit before tax', m(p.pbt)], ['Tax rate', pc(drv('taxRate', i))]]],
      netIncome: ['Profit before tax – tax', [['Profit before tax', m(p.pbt)], ['Tax', m(p.tax)]]],
      dividends: ['MAX(0, net income) × payout ratio', [['Net income', m(p.netIncome)], ['Payout', pc(drv('payoutRatio', i))]]],
      cash: ['Closing cash from the cash flow statement', [['Opening cash', m(p.openingCash)], ['Net change in cash', m(p.netChangeInCash)]]],
      receivables: ['Revenue × DSO ÷ days in period', [['Revenue', m(p.revenue)], ['DSO', drv('dso', i)], ['Days in period', d.toFixed(2)]]],
      inventory: ['COGS × DIO ÷ days in period', [['COGS', m(p.cogs)], ['DIO', drv('dio', i)], ['Days in period', d.toFixed(2)]]],
      otherCurrentAssets: ['Annualised revenue × other current assets %', [['Annualised revenue', m(p.revenue * k)], ['Rate', pc(drv('otherCurrentAssetsPct', i))]]],
      currentAssets: ['Cash + receivables + inventory + other current assets', [['Cash', m(p.cash)], ['Receivables', m(p.receivables)], ['Inventory', m(p.inventory)], ['Other', m(p.otherCurrentAssets)]]],
      ppe: ['Opening PP&E + CapEx – D&A', [['Opening PP&E', m(q.ppe)], ['CapEx', m(p.capex)], ['D&A', m(p.da)]]],
      totalAssets: ['Current assets + PP&E', [['Current assets', m(p.currentAssets)], ['PP&E', m(p.ppe)]]],
      payables: ['COGS × DPO ÷ days in period', [['COGS', m(p.cogs)], ['DPO', drv('dpo', i)], ['Days in period', d.toFixed(2)]]],
      otherCurrentLiabilities: ['Annualised revenue × other current liabilities %', [['Annualised revenue', m(p.revenue * k)], ['Rate', pc(drv('otherCurrentLiabilitiesPct', i))]]],
      currentLiabilities: ['Payables + other current liabilities', [['Payables', m(p.payables)], ['Other', m(p.otherCurrentLiabilities)]]],
      termDebt: ['Opening term debt + drawdowns – repayments', [['Opening', m(q.termDebt)], ['Drawn', m(p.debtDrawdown)], ['Repaid', m(p.debtRepayment)]]],
      revolver: ['Opening revolver + draws – repayments', [['Opening', m(q.revolver)], ['Drawn', m(p.revolverDraw)], ['Repaid', m(p.revolverRepay)], ['Limit', m(rv.limit)], ['Minimum cash', m(rv.minimumCash)]]],
      otherLiabilities: ['Held at the opening level', [['Opening', m(q.otherLiabilities)]]],
      totalLiabilities: ['Current liabilities + term debt + revolver + other', [['Current liabilities', m(p.currentLiabilities)], ['Total debt', m(p.totalDebt)], ['Other', m(p.otherLiabilities)]]],
      contributedCapital: ['Held at the opening level', [['Opening', m(q.contributedCapital)]]],
      retainedEarnings: ['Opening retained earnings + net income – dividends', [['Opening', m(q.retainedEarnings)], ['Net income', m(p.netIncome)], ['Dividends', m(p.dividends)]]],
      equity: ['Contributed capital + retained earnings', [['Contributed capital', m(p.contributedCapital)], ['Retained earnings', m(p.retainedEarnings)]]],
      liabEquity: ['Total liabilities + equity', [['Liabilities', m(p.totalLiabilities)], ['Equity', m(p.equity)]]],
      balanceCheck: ['Total assets – (liabilities + equity); must be zero', [['Total assets', m(p.totalAssets)], ['Liabilities + equity', m(p.totalLiabilities + p.equity)]]],
      cfNetIncome: ['Net income from the income statement', [['Net income', m(p.netIncome)]]],
      cfDa: ['D&A added back (non-cash)', [['D&A', m(p.da)]]],
      changeReceivables: ['Opening – closing receivables', [['Opening', m(q.receivables)], ['Closing', m(p.receivables)]]],
      changeInventory: ['Opening – closing inventory', [['Opening', m(q.inventory)], ['Closing', m(p.inventory)]]],
      changeOtherAssets: ['Opening – closing other current assets', [['Opening', m(q.otherCurrentAssets)], ['Closing', m(p.otherCurrentAssets)]]],
      changePayables: ['Closing – opening payables', [['Opening', m(q.payables)], ['Closing', m(p.payables)]]],
      changeOtherLiabilities: ['Closing – opening other current liabilities', [['Opening', m(q.otherCurrentLiabilities)], ['Closing', m(p.otherCurrentLiabilities)]]],
      cashFromOperations: ['Net income + D&A + working-capital changes', [['Net income', m(p.netIncome)], ['D&A', m(p.da)], ['Working-capital changes', m(p.changeReceivables + p.changeInventory + p.changeOtherAssets + p.changePayables + p.changeOtherLiabilities)]]],
      cfCapex: ['–(Revenue × CapEx %)', [['Revenue', m(p.revenue)], ['CapEx %', pc(drv('capexPct', i))]]],
      cfDrawdown: ['Annual drawdown ÷ periods per year', [['Annual drawdown', m(drv('debtDrawdown', i))], ['Periods per year', k]]],
      cfRepayment: ['–MIN(opening debt + drawdown, annual repayment ÷ periods per year)', [['Opening term debt', m(q.termDebt)], ['Annual repayment', m(drv('debtRepayment', i))]]],
      cfRevolver: ['Revolver draw – repayment, keeping cash at the minimum', [['Draw', m(p.revolverDraw)], ['Repay', m(p.revolverRepay)]]],
      cfDividends: ['–Dividends', [['Dividends', m(p.dividends)]]],
      cashFromFinancing: ['Drawdowns – repayments + revolver – dividends', [['Term debt', m(p.debtDrawdown - p.debtRepayment)], ['Revolver', m(p.revolverDraw - p.revolverRepay)], ['Dividends', m(-p.dividends)]]],
      netChangeInCash: ['Operating + investing + financing', [['Operating', m(p.cashFromOperations)], ['Investing', m(p.cashFromInvesting)], ['Financing', m(p.cashFromFinancing)]]],
      openingCash: ['Prior-period closing cash', [['Opening cash', m(p.openingCash)]]],
      closingCash: ['Opening cash + net change in cash', [['Opening cash', m(p.openingCash)], ['Net change', m(p.netChangeInCash)]]],
      freeCashFlow: ['Cash from operations – CapEx', [['Cash from operations', m(p.cashFromOperations)], ['CapEx', m(p.capex)]]],
    };
    return T[key] || null;
  }

  /* ---------- Statement definitions: [label, value getter, trace key, style, isBalance] ---------- */
  const IS_ROWS = [
    ['Revenue', (p) => p.revenue, 'revenue', 'b'], ['Cost of goods sold', (p) => -p.cogs, 'cogs'], ['Gross profit', (p) => p.grossProfit, 'grossProfit', 't'],
    ['Operating expenses', (p) => -p.opex, 'opex'], ['EBITDA', (p) => p.ebitda, 'ebitda', 't'], ['Depreciation & amortization', (p) => -p.da, 'da'],
    ['EBIT', (p) => p.ebit, 'ebit', 't'], ['Interest expense', (p) => -p.interest, 'interest'], ['Profit before tax', (p) => p.pbt, 'pbt', 't'],
    ['Income tax', (p) => -p.tax, 'tax'], ['Net income', (p) => p.netIncome, 'netIncome', 'g'], ['Dividends', (p) => p.dividends, 'dividends', 'm'],
  ];
  const BS_ROWS = [
    ['Cash', (p) => p.cash, 'cash', '', 1], ['Accounts receivable', (p) => p.receivables, 'receivables', '', 1], ['Inventory', (p) => p.inventory, 'inventory', '', 1],
    ['Other current assets', (p) => p.otherCurrentAssets, 'otherCurrentAssets', '', 1], ['Total current assets', (p) => p.currentAssets, 'currentAssets', 't', 1],
    ['PP&E and other non-current assets', (p) => p.ppe, 'ppe', '', 1], ['Total assets', (p) => p.totalAssets, 'totalAssets', 'g', 1],
    ['Accounts payable', (p) => p.payables, 'payables', '', 1], ['Other current liabilities', (p) => p.otherCurrentLiabilities, 'otherCurrentLiabilities', '', 1],
    ['Total current liabilities', (p) => p.currentLiabilities, 'currentLiabilities', 't', 1], ['Term debt', (p) => p.termDebt, 'termDebt', '', 1], ['Revolver', (p) => p.revolver, 'revolver', '', 1],
    ['Other non-current liabilities', (p) => p.otherLiabilities, 'otherLiabilities', '', 1], ['Total liabilities', (p) => p.totalLiabilities, 'totalLiabilities', 't', 1],
    ['Contributed capital', (p) => p.contributedCapital, 'contributedCapital', '', 1], ['Retained earnings', (p) => p.retainedEarnings, 'retainedEarnings', '', 1],
    ['Total equity', (p) => p.equity, 'equity', 't', 1], ['Total liabilities & equity', (p) => p.totalLiabilities + p.equity, 'liabEquity', 'g', 1],
    ['Balance check (must be zero)', (p) => p.checks.balance, 'balanceCheck', 'm', 1],
  ];
  const CF_ROWS = [
    ['Net income', (p) => p.netIncome, 'cfNetIncome'], ['Depreciation & amortization', (p) => p.da, 'cfDa'], ['(Increase) / decrease in receivables', (p) => p.changeReceivables, 'changeReceivables'],
    ['(Increase) / decrease in inventory', (p) => p.changeInventory, 'changeInventory'], ['(Increase) / decrease in other current assets', (p) => p.changeOtherAssets, 'changeOtherAssets'],
    ['Increase / (decrease) in payables', (p) => p.changePayables, 'changePayables'], ['Increase / (decrease) in other current liabilities', (p) => p.changeOtherLiabilities, 'changeOtherLiabilities'],
    ['Cash from operations', (p) => p.cashFromOperations, 'cashFromOperations', 't'], ['Capital expenditure', (p) => -p.capex, 'cfCapex'], ['Cash from investing', (p) => p.cashFromInvesting, 'cfCapex', 't'],
    ['Term debt drawn', (p) => p.debtDrawdown, 'cfDrawdown'], ['Term debt repaid', (p) => -p.debtRepayment, 'cfRepayment'], ['Revolver draw / (repayment)', (p) => p.revolverDraw - p.revolverRepay, 'cfRevolver'],
    ['Dividends paid', (p) => -p.dividends, 'cfDividends'], ['Cash from financing', (p) => p.cashFromFinancing, 'cashFromFinancing', 't'],
    ['Net change in cash', (p) => p.netChangeInCash, 'netChangeInCash', 't'], ['Opening cash', (p) => p.openingCash, 'openingCash', '', 1], ['Closing cash', (p) => p.closingCashCF, 'closingCash', 'g', 1],
    ['Free cash flow (CFO – CapEx)', (p) => p.freeCashFlow, 'freeCashFlow', 'm'],
  ];

  function statementTable(title, rows) {
    const per = statementView === 'periods' || result.proj.periodsPerYear === 1;
    const k = result.proj.periodsPerYear, ps = result.proj.periods;
    // Annual view of a quarterly/monthly model: flows are summed, balances are closing values; opening balance flows are the year's first
    const cols = per ? ps.map((p, i) => ({ label: periodLabel(i), i, group: [i] })) : result.annual.map((_, y) => ({ label: yearLabel(y), i: (y + 1) * k - 1, group: Array.from({ length: k }, (_, j) => y * k + j) }));
    const val = (r, c) => {
      if (r[4]) return r[2] === 'openingCash' ? r[1](ps[c.group[0]]) : r[1](ps[c.i]);
      return c.group.reduce((s, j) => s + r[1](ps[j]), 0);
    };
    const head = `<tr><th>${title} <span class="b-unit">${esc(unitLabel())}</span></th>${cols.map((c) => `<th>${c.label}</th>`).join('')}</tr>`;
    const body = rows.map((r) => `<tr class="${r[3] || ''}"><td>${r[0]}</td>${cols.map((c) => {
      const idx = r[4] && r[2] === 'openingCash' ? c.group[0] : c.i, sum = !r[4] && c.group.length > 1 ? c.group.join(',') : '';
      return `<td class="tr" tabindex="0" data-trace="${r[2]}" data-i="${idx}" data-sum="${sum}">${money(val(r, c))}</td>`;
    }).join('')}</tr>`).join('');
    return `<div class="table-wrap b-stmt"><table><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
  }

  function showTrace(td) {
    $('#bTrace')?.remove();
    const key = td.dataset.trace, i = +td.dataset.i, sum = td.dataset.sum ? td.dataset.sum.split(',').map(Number) : null;
    const row = td.closest('tr').firstElementChild.textContent;
    let html;
    if (sum && sum.length > 1) {
      const getter = [...IS_ROWS, ...BS_ROWS, ...CF_ROWS].find((r) => r[2] === key)?.[1];
      html = `<p class="b-f">Sum of the ${sum.length} periods in the year</p><table>${sum.map((j) => `<tr><td>${periodLabel(j)}</td><td>${money(getter(result.proj.periods[j]))}</td></tr>`).join('')}</table>`;
    } else {
      const t = trace(key, i);
      if (!t) return;
      html = `<p class="b-f">${esc(t[0])}</p><table>${t[1].map(([l, v]) => `<tr><td>${esc(l)}</td><td>${esc(String(v))}</td></tr>`).join('')}</table>`;
    }
    const pop = document.createElement('div');
    pop.id = 'bTrace';
    pop.className = 'b-trace';
    pop.setAttribute('role', 'dialog');
    const when = sum && sum.length > 1 ? yearLabel(Math.floor(i / result.proj.periodsPerYear)) : periodLabel(i);
    pop.innerHTML = `<div class="row spread"><strong>${esc(row)} · ${when}</strong><button class="chat-x" type="button" aria-label="Close">×</button></div>${html}<p class="b-u">${esc(unitLabel())} unless stated</p>`;
    document.body.appendChild(pop);
    const r = td.getBoundingClientRect(), w = Math.min(340, window.innerWidth - 24);
    pop.style.width = w + 'px';
    pop.style.left = Math.max(12, Math.min(window.innerWidth - w - 12, r.left + window.scrollX - w / 2 + r.width / 2)) + 'px';
    pop.style.top = r.bottom + window.scrollY + 6 + 'px';
    pop.querySelector('button').addEventListener('click', () => pop.remove());
  }

  /* ---------- Inputs ---------- */
  const dis = () => (readOnly() ? ' disabled' : '');
  const toDisplay = (kind, v) => (kind === 'pct' ? +(v * 100).toFixed(4) : kind === 'amount' ? +(v / unit().div).toFixed(4) : v);
  const fromDisplay = (kind, x) => (kind === 'pct' ? x / 100 : kind === 'amount' ? x * unit().div : x);
  const field = (path, label, kind, value, range, note = '') => `<div class="field"><label>${label}</label><input type="number" step="any" data-path="${path}" data-kind="${kind}" data-min="${range[0]}" data-max="${range[1]}" value="${toDisplay(kind, value)}"${dis()}>${note ? `<span class="b-note">${note}</span>` : ''}</div>`;
  const setPath = (path, v) => { const ks = path.split('.'); let o = bm; ks.slice(0, -1).forEach((k) => { o = o[k]; }); o[ks.at(-1)] = v; };

  function onInput(el) {
    const kind = el.dataset.kind, x = parseFloat(el.value), min = +el.dataset.min, max = +el.dataset.max;
    // Validate every input: numbers only, inside a sensible range; invalid entries are rejected, never used
    if (!isFinite(x) || x < min || x > max) { el.classList.add('bad'); el.title = `Enter a number from ${min} to ${max}`; return false; }
    el.classList.remove('bad'); el.title = '';
    setPath(el.dataset.path, fromDisplay(kind, x));
    return true;
  }

  /* ---------- Steps ---------- */
  function stepSetup() {
    const sel = (path, label, opts, v) => `<div class="field"><label>${label}</label><select data-path="${path}"${dis()}>${opts.map(([k, l]) => `<option value="${k}"${String(v) === String(k) ? ' selected' : ''}>${l}</option>`).join('')}</select></div>`;
    const saved = savedList.length ? `<div class="table-wrap"><table><thead><tr><th>Saved model</th><th>Last saved</th><th>Versions</th><th></th></tr></thead><tbody>${savedList.map((s) =>
      `<tr><td>${esc(s.name)}</td><td>${new Date(s.updated).toLocaleString()}</td><td>${s.versions.length}</td><td><button class="link" data-load="${s.id}">Open</button> · <button class="link danger" data-del="${s.id}">Delete</button></td></tr>`).join('')}</tbody></table></div>` : '<p class="muted small">No saved models yet. Use “Save version” to keep this model in this browser.</p>';
    return `<div class="glass card"><h3>Model setup</h3><div class="form-grid">
      <div class="field"><label>Model name</label><input type="text" maxlength="80" data-text="name" value="${esc(bm.name)}"${dis()}></div>
      <div class="field"><label>Company</label><input type="text" maxlength="80" data-text="company" value="${esc(bm.company)}"${dis()}></div>
      <div class="field"><label>Currency</label><select data-path="currency"${dis()}>${currencyOptions(bm.currency)}</select></div>
      ${sel('units', 'Display units', [['actual', 'Actual'], ['thousands', 'Thousands (000s)'], ['millions', 'Millions']], bm.units)}
      ${sel('periodType', 'Periods', [['annual', 'Annual'], ['quarterly', 'Quarterly'], ['monthly', 'Monthly']], bm.periodType)}
      ${sel('years', 'Forecast length', [3, 4, 5, 6, 7, 8, 9, 10].map((n) => [n, `${n} years`]), bm.years)}
      ${field('firstYear', 'First forecast year', 'days', bm.firstYear, [1990, 2100])}
      </div><p class="muted small">Data source: ${esc(bm.source || 'Manual')}. ${(typeof hasData === "function" && hasData()) ? `<button class="link" data-act="fromData">Rebuild from ${esc(state.company)}’s statements</button>` : '<a class="link" href="#upload">Upload statements</a> to start from your own numbers.'} · <button class="link" data-act="new">Start a blank example</button></p></div>
      <div class="glass card"><h3>Opening balance sheet <span class="b-unit">${esc(unitLabel())}</span></h3><p class="muted small">The last actual balance sheet. It must balance before you forecast.</p>
      <div class="form-grid">${OPENING.map(([k, l]) => field(`opening.${k}`, l, 'amount', bm.opening[k] ?? 0, [-1e15, 1e15])).join('')}</div>
      ${openingCheck()}</div>
      <div class="glass card"><h3>Saved models</h3>${saved}</div>`;
  }
  function openingCheck() {
    const o = bm.opening, a = o.cash + o.receivables + o.inventory + o.otherCurrentAssets + o.ppe, le = o.payables + o.otherCurrentLiabilities + o.termDebt + (o.revolver || 0) + o.otherLiabilities + o.contributedCapital + o.retainedEarnings;
    const ok = Math.abs(a - le) <= F.TOLERANCE;
    return `<p class="b-check ${ok ? 'ok' : 'no'}">${ok ? '✅' : '❌'} Opening assets ${full(a)} ${ok ? '=' : '≠'} liabilities + equity ${full(le)}${ok ? '' : ` (difference ${full(a - le)}; adjust other liabilities or equity)`}</p>`;
  }
  function stepAssumptions() {
    const n = bm.years, rv = bm.settings.revolver;
    const rows = DRIVERS.map(([k, l, kind, range, help]) => `<tr><td><span title="${esc(help)}">${l}</span><span class="b-kind">${kind === 'pct' ? '%' : kind === 'days' ? 'days' : unitLabel()}</span></td>${bm.assumptions[k].map((v, y) =>
      `<td><input type="number" step="any" class="b-in" data-path="assumptions.${k}.${y}" data-kind="${kind}" data-min="${range[0]}" data-max="${range[1]}" value="${toDisplay(kind, v)}" aria-label="${esc(l)} ${yearLabel(y)}"${dis()}></td>`).join('')}<td><button class="link small" data-fill="${k}" title="Copy year 1 to every year"${dis()}>Fill →</button></td></tr>`).join('');
    return `<div class="glass card"><h3>Forecast drivers by year</h3><p class="muted small">Blue cells are inputs. Hover a driver for how it is used.</p>
      <div class="table-wrap"><table class="b-grid"><thead><tr><th>Driver</th>${Array.from({ length: n }, (_, y) => `<th>${yearLabel(y)}</th>`).join('')}<th></th></tr></thead><tbody>${rows}</tbody></table></div></div>
      <div class="glass card"><h3>Debt and interest</h3><div class="form-grid">
      <div class="field"><label>Interest calculation</label><select data-path="settings.interestMethod"${dis()}><option value="opening"${bm.settings.interestMethod === 'opening' ? ' selected' : ''}>Opening balances (no circularity)</option><option value="average"${bm.settings.interestMethod === 'average' ? ' selected' : ''}>Average balances (iterative, circular switch on)</option></select></div>
      <div class="field"><label>Revolving credit facility</label><select data-path="settings.revolver.enabled" data-bool="1"${dis()}><option value="0"${rv.enabled ? '' : ' selected'}>Off</option><option value="1"${rv.enabled ? ' selected' : ''}>On (draws to keep minimum cash)</option></select></div>
      ${field('settings.revolver.limit', 'Revolver limit', 'amount', rv.limit, [0, 1e15])}
      ${field('settings.revolver.interestRate', 'Revolver interest rate %', 'pct', rv.interestRate, [0, 40])}
      ${field('settings.revolver.minimumCash', 'Minimum cash balance', 'amount', rv.minimumCash, [0, 1e15])}
      </div><p class="muted small">Average-balance interest makes interest depend on cash, and cash depend on interest. The model solves it by explicit iteration (up to 100 passes, tolerance 0.0000001) and reports convergence on the Checks step.</p></div>`;
  }
  function stepStatements() {
    const toggle = result.proj.periodsPerYear > 1 ? `<div class="row"><span class="muted small">Show</span><button class="btn ${statementView === 'annual' ? '' : 'ghost'}" data-view-st="annual">Annual</button><button class="btn ${statementView === 'periods' ? '' : 'ghost'}" data-view-st="periods">${bm.periodType === 'monthly' ? 'Months' : 'Quarters'}</button></div>` : '';
    return `<div class="glass card"><div class="row spread"><div><h3>Financial statements</h3><p class="muted small">Click any figure to see its formula and inputs. Negatives are in parentheses.</p></div>${toggle}</div></div>
      <div class="glass card">${statementTable('Income statement', IS_ROWS)}</div>
      <div class="glass card">${statementTable('Balance sheet', BS_ROWS)}</div>
      <div class="glass card">${statementTable('Cash flow statement', CF_ROWS)}</div>`;
  }
  function stepValuation() {
    const v = bm.valuation, val = result.val, c = result.waccCalc;
    const waccBlock = `<div class="form-grid">
      <div class="field"><label>WACC source</label><select data-path="valuation.useCalculator" data-bool="1"${dis()}><option value="1"${v.useCalculator ? ' selected' : ''}>Calculate (CAPM)</option><option value="0"${v.useCalculator ? '' : ' selected'}>Enter directly</option></select></div>
      ${v.useCalculator ? [field('valuation.riskFree', 'Risk-free rate %', 'pct', v.riskFree, [-5, 30]), field('valuation.beta', 'Beta', 'days', v.beta, [-5, 10]), field('valuation.marketPremium', 'Equity risk premium %', 'pct', v.marketPremium, [0, 30]),
        field('valuation.sizePremium', 'Size / specific premium %', 'pct', v.sizePremium, [-10, 30]), field('valuation.preTaxCostOfDebt', 'Pre-tax cost of debt %', 'pct', v.preTaxCostOfDebt, [0, 40]),
        field('valuation.equityValue', 'Equity value (market)', 'amount', v.equityValue, [0, 1e16]), field('valuation.debtValue', 'Debt value', 'amount', v.debtValue, [0, 1e16])].join('')
        : field('valuation.wacc', 'WACC %', 'pct', v.wacc, [0, 50])}</div>
      ${c ? `<p class="muted small">Cost of equity ${pc(c.costOfEquity, 2)} = ${pc(v.riskFree, 2)} + ${v.beta} × ${pc(v.marketPremium, 2)}${v.sizePremium ? ` + ${pc(v.sizePremium, 2)}` : ''} · After-tax cost of debt ${pc(c.afterTaxCostOfDebt, 2)} · Weights ${pc(c.equityWeight)} equity / ${pc(c.debtWeight)} debt · <strong>WACC ${pc(c.wacc, 2)}</strong></p>` : ''}`;
    const tvBlock = `<div class="form-grid">
      <div class="field"><label>Terminal value method</label><select data-path="valuation.method"${dis()}><option value="gordon"${v.method === 'gordon' ? ' selected' : ''}>Gordon growth</option><option value="exit"${v.method === 'exit' ? ' selected' : ''}>Exit multiple (EV / EBITDA)</option></select></div>
      ${field('valuation.terminalGrowth', 'Terminal growth %', 'pct', v.terminalGrowth, [-5, 10])}
      ${field('valuation.exitMultiple', 'Exit multiple (x EBITDA)', 'days', v.exitMultiple, [0, 100])}
      <div class="field"><label>Discounting</label><select data-path="valuation.midYear" data-bool="1"${dis()}><option value="1"${v.midYear ? ' selected' : ''}>Mid-period convention</option><option value="0"${v.midYear ? '' : ' selected'}>End of period</option></select></div>
      ${field('valuation.shares', 'Diluted shares outstanding', 'days', v.shares, [1, 1e13])}</div>`;
    if (val.error) return `<div class="glass card"><h3>Cost of capital</h3>${waccBlock}</div><div class="glass card"><h3>Terminal value</h3>${tvBlock}<p class="b-check no">❌ ${esc(val.error)}</p></div>`;
    const k = result.proj.periodsPerYear, years = result.annual.length;
    const fy = Array.from({ length: years }, (_, y) => val.rows.slice(y * k, (y + 1) * k));
    const sumY = (y, f) => fy[y].reduce((s, r) => s + r[f], 0);
    const fRow = (l, f, cls = '') => `<tr class="${cls}"><td>${l}</td>${fy.map((_, y) => `<td>${money(f(y))}</td>`).join('')}</tr>`;
    const fcffTable = `<div class="table-wrap"><table><thead><tr><th>Free cash flow to the firm <span class="b-unit">${esc(unitLabel())}</span></th>${fy.map((_, y) => `<th>${yearLabel(y)}</th>`).join('')}</tr></thead><tbody>
      ${fRow('EBIT', (y) => sumY(y, 'ebit'))}${fRow('Less: taxes on EBIT', (y) => -(sumY(y, 'ebit') - sumY(y, 'nopat')))}${fRow('NOPAT', (y) => sumY(y, 'nopat'), 't')}
      ${fRow('Plus: D&A', (y) => sumY(y, 'da'))}${fRow('Less: CapEx', (y) => -sumY(y, 'capex'))}${fRow('Less: increase in net working capital', (y) => -sumY(y, 'increaseNWC'))}
      ${fRow('FCFF', (y) => sumY(y, 'fcff'), 'g')}${fRow('Present value', (y) => sumY(y, 'pv'))}</tbody></table></div>`;
    const bridge = [['PV of forecast FCFF', val.sumPV], [`PV of terminal value (${v.method === 'exit' ? `${v.exitMultiple}x LTM EBITDA` : `Gordon, g = ${pc(v.terminalGrowth)}`})`, val.pvTerminal], ['Enterprise value', val.enterpriseValue, 'g'],
      [val.netDebt >= 0 ? 'Less: net debt (debt – cash at the valuation date)' : 'Plus: net cash (cash – debt at the valuation date)', -val.netDebt], ['Equity value', val.equityValue, 'g']];
    const sens = sensitivity();
    return `<div class="kpis">${[['Enterprise value', full(val.enterpriseValue)], ['Equity value', full(val.equityValue)], ['Value per share', `${bm.currency} ${val.perShare.toFixed(2)}`], ['WACC', pc(result.wacc, 2)], ['Terminal value share of EV', pc(val.tvShare)]]
      .map(([l, x]) => `<div class="glass kpi"><div class="label">${l}</div><div class="value b-kv">${x}</div></div>`).join('')}</div>
      ${val.warnings.length ? `<div class="glass card b-warn">${val.warnings.map((w) => `<p>⚠️ ${esc(w.text)}</p>`).join('')}</div>` : ''}
      <div class="glass card"><h3>Cost of capital</h3>${waccBlock}</div>
      <div class="glass card"><h3>Terminal value and discounting</h3>${tvBlock}<p class="muted small">Implied ${v.method === 'exit' ? `perpetual growth ${pc(val.impliedGrowth, 2)}` : `exit multiple ${val.impliedExitMultiple?.toFixed(1)}x LTM EBITDA`}. FCFF = EBIT × (1 – tax) + D&A – CapEx – increase in NWC. ${v.midYear ? 'Cash flows are discounted from the middle of each period; the terminal value from the end of the forecast.' : 'Cash flows are discounted from the end of each period.'}</p></div>
      <div class="glass card">${fcffTable}</div>
      <div class="grid two"><div class="glass card"><h3>Enterprise value to equity value</h3><div class="table-wrap"><table><tbody>${bridge.map(([l, x, c]) => `<tr class="${c || ''}"><td>${l}</td><td>${full(x)}</td></tr>`).join('')}<tr><td>Diluted shares</td><td>${num(v.shares, 0, 1)}</td></tr><tr class="g"><td>Value per share</td><td>${bm.currency} ${val.perShare.toFixed(2)}</td></tr></tbody></table></div></div>
      <div class="glass card"><h3>Sensitivity: value per share</h3>${sens}</div></div>
      <div class="grid two"><div class="glass card"><h3>Revenue and EBITDA</h3><div class="chart"><canvas id="bChRev"></canvas></div></div><div class="glass card"><h3>Free cash flow and cash</h3><div class="chart"><canvas id="bChCash"></canvas></div></div></div>`;
  }
  function sensitivity() {
    const v = bm.valuation, w = result.wacc, ws = [-0.01, -0.005, 0, 0.005, 0.01].map((d) => w + d);
    const exit = v.method === 'exit', xs = exit ? [-2, -1, 0, 1, 2].map((d) => v.exitMultiple + d) : [-0.005, -0.0025, 0, 0.0025, 0.005].map((d) => v.terminalGrowth + d);
    const cell = (wi, x) => {
      const r = valuationFor(result.proj, bm, wi, exit ? { exitMultiple: x } : { terminalGrowth: x });
      return r.error ? '—' : r.perShare.toFixed(2);
    };
    return `<div class="table-wrap"><table><thead><tr><th>WACC ↓ / ${exit ? 'exit multiple' : 'terminal growth'} →</th>${xs.map((x) => `<th>${exit ? x.toFixed(1) + 'x' : pc(x, 2)}</th>`).join('')}</tr></thead><tbody>${ws.map((wi, a) => `<tr><td>${pc(wi, 2)}</td>${xs.map((x, b) => `<td class="${a === 2 && b === 2 ? 'hl' : ''}">${cell(wi, x)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  }
  function stepChecks() {
    const passed = result.checks.filter((c) => c.pass).length;
    return `<div class="glass card"><h3>Integrity checks · ${passed} of ${result.checks.length} pass</h3><p class="muted small">These rules are non-negotiable. A failing check is shown here and on the Checks sheet of the Excel file, never hidden.</p>
      <ul class="b-checks">${result.checks.map((c) => `<li class="${c.pass ? 'ok' : 'no'}"><span class="b-ico">${c.pass ? '✅' : '❌'}</span><div><strong>${esc(c.label)}</strong><p>${esc(c.rule)}</p><p class="b-det">${esc(c.detail)}</p></div></li>`).join('')}</ul></div>`;
  }
  function stepExport() {
    return `<div class="glass card"><h3>Export</h3><p class="muted small">The Excel workbook contains Assumptions, Income Statement, Balance Sheet, Cash Flow, DCF and Checks sheets with live formulas (blue inputs, black formulas, green links) in ${esc(unitLabel())}.</p>
      <div class="row"><button class="btn btn-lg" data-act="excel">Download Excel model</button><span class="upload-msg small" id="bMsg" role="status"></span></div></div>
      <div class="glass card"><h3>Save and share</h3><p class="muted small">Saved models stay in this browser with up to ${MAX_VERSIONS} versions each. Export a JSON file to back up a model or move it to another computer.</p>
      <div class="row"><button class="btn" data-act="save">Save version</button><button class="btn ghost" data-act="json">Export JSON</button><label class="btn ghost">Import JSON<input type="file" accept="application/json,.json" id="bImport" hidden></label></div>
      ${versionList()}</div>`;
  }
  function versionList() {
    const s = savedList.find((x) => x.id === bm.id);
    if (!s) return '<p class="muted small">This model has not been saved yet.</p>';
    return `<div class="table-wrap"><table><thead><tr><th>Version</th><th>Saved</th><th></th></tr></thead><tbody>${s.versions.map((v, i) => `<tr><td>v${s.versions.length - i}</td><td>${new Date(v.at).toLocaleString()}</td><td><button class="link" data-restore="${i}">Restore</button></td></tr>`).join('')}</tbody></table></div>`;
  }

  /* ---------- Render ---------- */
  // Identifies the uploaded statements, so a new upload (or removing them) never mixes with the previous company's model
  function dataKey() { return (typeof hasData === 'function' && hasData()) ? `${state.company}|${state.years.join(',')}|${S('revenue')?.join(',')}` : null; }
  function render() {
    if (!bm) bm = loadDraft() || modelFromData();
    // Rebuild only when the uploaded statements change (new upload or removal), not when a saved/imported model is opened
    const cur = dataKey() || '', seen = (() => { try { return localStorage.getItem('fs-builder-data'); } catch (e) { return null; } })();
    if (seen != null && seen !== cur) { bm = modelFromData(); step = 0; }
    try { localStorage.setItem('fs-builder-data', cur); } catch (e) {}
    compute();
    savedList = readSaved();
    const fails = result.checks.filter((c) => !c.pass).length;
    $('#bSteps').innerHTML = STEPS.map((s, i) => `<li class="${i === step ? 'active' : ''}"><button type="button" data-step="${i}"><span>${i + 1}</span>${s}${i === 4 ? ` <em class="${fails ? 'no' : 'ok'}">${fails ? `${fails} ✕` : '✓'}</em>` : ''}</button></li>`).join('');
    $('#bStatus').innerHTML = `${esc(bm.name)} · ${bm.years}-year ${bm.periodType} forecast · ${esc(unitLabel())} · ${fails ? `<span class="no">❌ ${fails} check${fails > 1 ? 's' : ''} failing</span>` : '<span class="ok">✅ all checks pass</span>'}${readOnly() ? ' · <em>View only on small screens</em>' : ''}`;
    $('#bBody').innerHTML = [stepSetup, stepAssumptions, stepStatements, stepValuation, stepChecks, stepExport][step]();
    $('#bPrev').disabled = step === 0;
    $('#bNext').disabled = step === STEPS.length - 1;
    $('#bNext').textContent = step < STEPS.length - 1 ? `Next: ${STEPS[step + 1]} →` : 'Done';
    if (step === 3 && !result.val.error) drawCharts();
    saveDraft();
  }
  function drawCharts() {
    const a = result.annual, labels = a.map((_, y) => yearLabel(y)), d = unit().div;
    const sc = (arr) => arr.map((x) => x / d);
    const axis = { ticks: { callback: (x) => num(x * d) } };
    makeChart('bChRev', { type: 'bar', data: { labels, datasets: [ds('Revenue', sc(a.map((p) => p.revenue)), COLORS.blue), ds('EBITDA', sc(a.map((p) => p.ebitda)), COLORS.green)] }, options: { scales: { y: axis } } });
    makeChart('bChCash', { type: 'line', data: { labels, datasets: [ds('Free cash flow', sc(a.map((p) => p.freeCashFlow)), COLORS.orange), ds('Closing cash', sc(a.map((p) => p.cash)), COLORS.teal)] }, options: { scales: { y: axis } } });
  }

  /* ---------- Persistence: draft, saved models with versions, JSON ---------- */
  const DRAFT_KEY = 'fs-builder-draft';
  function saveDraft() { try { localStorage.setItem(DRAFT_KEY, JSON.stringify(bm)); } catch (e) {} }
  function loadDraft() { try { return validModel(JSON.parse(localStorage.getItem(DRAFT_KEY))); } catch (e) { return null; } }
  function readSaved() { try { const s = JSON.parse(localStorage.getItem(MODELS_KEY)); return Array.isArray(s) ? s : []; } catch (e) { return []; } }
  function writeSaved(list) {
    try { localStorage.setItem(MODELS_KEY, JSON.stringify(list)); return true; } catch (e) { console.warn('Finance Studio: could not save the model (browser storage full or blocked).'); return false; }
  }
  // Accept only well-formed models, so a damaged file or draft can't break the page
  function validModel(m) {
    if (!m || m.schema !== 'finance-studio-model' || typeof m.assumptions !== 'object' || typeof m.opening !== 'object') return null;
    const base = blankModel();
    const out = { ...base, ...m, assumptions: { ...base.assumptions }, opening: { ...base.opening }, settings: { ...base.settings, ...(m.settings || {}), revolver: { ...base.settings.revolver, ...(m.settings?.revolver || {}) } }, valuation: { ...base.valuation, ...(m.valuation || {}) } };
    for (const [k] of DRIVERS) {
      const a = m.assumptions[k];
      if (a != null) out.assumptions[k] = (Array.isArray(a) ? a : [a]).map(Number);
      if (out.assumptions[k].some((x) => !isFinite(x))) return null;
    }
    for (const [k] of OPENING) { const x = Number(m.opening[k] ?? 0); if (!isFinite(x)) return null; out.opening[k] = x; }
    if (!PERIODS[out.periodType]) out.periodType = 'annual';
    if (!UNITS[out.units]) out.units = 'thousands';
    out.name = String(out.name || 'Imported model').slice(0, 80);
    out.company = String(out.company || '').slice(0, 80);
    out.firstYear = Math.round(Number(out.firstYear)) || new Date().getFullYear();
    resize(out);
    return out;
  }
  const PERIODS = F.PERIODS_PER_YEAR;
  function saveVersion() {
    const list = readSaved();
    let s = list.find((x) => x.id === bm.id);
    if (!s) { s = { id: bm.id, name: bm.name, versions: [] }; list.unshift(s); }
    s.name = bm.name;
    s.updated = Date.now();
    s.versions.unshift({ at: s.updated, data: JSON.parse(JSON.stringify(bm)) });
    s.versions = s.versions.slice(0, MAX_VERSIONS);
    if (writeSaved(list)) msg('Saved. Versions are kept in this browser; export JSON for a backup.');
    else msg('Could not save: browser storage is full or blocked. Export JSON instead.');
    render();
  }
  function msg(t) { const el = $('#bMsg') || $('#bStatus'); if (el) el.insertAdjacentHTML('beforeend', ` <span class="b-flash">${esc(t)}</span>`); }
  function exportJSON() {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(bm, null, 2)], { type: 'application/json' }));
    a.download = `${bm.name.replace(/[^\w-]+/g, '-').toLowerCase() || 'model'}.json`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  }
  async function importJSON(file) {
    if (!file || file.size > 2e6) { alert('Choose a Finance Studio model file (.json, under 2 MB).'); return; }
    try {
      const m = validModel(JSON.parse(await file.text()));
      if (!m) throw new Error('bad');
      bm = { ...m, id: 'm' + Date.now().toString(36) };
      step = 0; render();
    } catch (e) { alert('That file is not a valid Finance Studio model.'); }
  }

  /* ---------- Excel export: the same logic as live formulas ---------- */
  function builderBook() {
    compute();
    const book = newBook(), m = bm, k = PERIODS[m.periodType], n = m.years, P = n * k, u = unit();
    const Z = u.div === 1e6 ? '#,##0.0,,;(#,##0.0,,);"–"' : u.div === 1e3 ? '#,##0,;(#,##0,);"–"' : '#,##0;(#,##0);"–"';
    const ZP = '0.0%;(0.0%);"–"', ZC = '#,##0.00;(#,##0.00);"–"';
    const IS = 'Income Statement', BSn = 'Balance Sheet', CFn = 'Cash Flow', A = 'Assumptions';
    const pc = (j) => colL(2 + j);                         // period j (0-based) column; column B holds opening values
    const qr = (s, a) => `${Q(s)}!${a}`;
    const subt = `${m.company || m.name} · ${m.years}-year ${m.periodType} forecast · figures in ${unitLabel()}`;

    // --- Assumptions ---
    const As = addSheet(book, A, 'Assumptions', subt);
    const asm = (...x) => qr(A, As.assume(...x));
    As.section(5, 'Settings', 3);
    const S_K = asm(6, 'Periods per year', k, '0', 'Structural: rebuild the model to change the period type');
    const S_SW = asm(7, 'Interest on average balances (1 = on, 0 = opening balances)', m.settings.interestMethod === 'average' ? 1 : 0, '0', 'Circularity switch; iterative calculation is enabled when the revolver is on');
    const S_RV = asm(8, 'Revolver on (1 = yes, 0 = no)', m.settings.revolver.enabled ? 1 : 0, '0');
    const S_RL = asm(9, 'Revolver limit', m.settings.revolver.limit, Z);
    const S_RR = asm(10, 'Revolver interest rate', m.settings.revolver.interestRate, ZP);
    const S_RM = asm(11, 'Minimum cash balance', m.settings.revolver.minimumCash, Z);
    As.section(13, 'Forecast drivers by year', n + 1);
    As.header(14, ['Driver', ...Array.from({ length: n }, (_, y) => yearLabel(y))]);
    const DROW = {};
    DRIVERS.forEach(([key, label, kind], j) => {
      const r = 15 + j;
      DROW[key] = r;
      As.put(`A${r}`, label, null, 'label');
      m.assumptions[key].forEach((v, y) => As.put(`${colL(1 + y)}${r}`, v, kind === 'pct' ? ZP : kind === 'days' ? '0.0' : Z, 'input'));
    });
    const drvRef = (key, j) => qr(A, `$${colL(1 + Math.floor(j / k))}$${DROW[key]}`);
    const r0 = 16 + DRIVERS.length;
    As.section(r0, 'Opening balance sheet (last actual)', 2);
    const OROW = {};
    OPENING.forEach(([key, label], j) => { const r = r0 + 1 + j; OROW[key] = r; As.assume(r, label, m.opening[key] ?? 0, Z); });
    const op = (key) => qr(A, `$B$${OROW[key]}`);
    As.sh.widths = { A: 52, ...Object.fromEntries(Array.from({ length: n }, (_, y) => [colL(1 + y), 13])), [colL(1 + n)]: 13 };
    As.sh.widths.C = Math.max(As.sh.widths.C || 13, 13);

    // Statement sheet helper: header row, labels, opening column B, period columns
    const statement = (name, title) => {
      const s = addSheet(book, name, title, subt);
      s.put('A5', `${unitLabel()}`, null, 'headL'); s.put('B5', 'Opening', null, 'headA');
      for (let j = 0; j < P; j++) s.put(`${pc(j)}5`, periodLabel(j), null, 'headF');
      s.sh.widths = { A: 46, B: 13, ...Object.fromEntries(Array.from({ length: P }, (_, j) => [pc(j), 12])) };
      s.sh.freeze = { x: 2, y: 5 };
      return s;
    };
    const ISs = statement(IS, 'Income Statement'), BSs = statement(BSn, 'Balance Sheet'), CFs = statement(CFn, 'Cash Flow Statement');
    const row = (s, r, label, f, style = 'calc', open) => {
      s.put(`A${r}`, label, null, style === 'total' || style === 'output' ? 'labelB' : 'label');
      if (open !== undefined) s.put(`B${r}`, open, Z, typeof open === 'string' && open[0] === '=' ? 'link' : 'input');
      for (let j = 0; j < P; j++) s.put(`${pc(j)}${r}`, '=' + f(j, pc(j), j ? pc(j - 1) : 'B'), Z, style);
    };
    // Row numbers
    const I = { rev: 6, cogs: 7, gp: 8, opex: 9, ebitda: 10, da: 11, ebit: 12, int: 13, pbt: 14, tax: 15, ni: 16, div: 17 };
    const B = { cash: 6, ar: 7, inv: 8, oca: 9, tca: 10, ppe: 11, ta: 12, ap: 13, ocl: 14, tcl: 15, td: 16, rv: 17, ol: 18, tl: 19, cc: 20, re: 21, eq: 22, tle: 23, chk: 24, nwc: 26 };
    const C = { ni: 6, da: 7, ar: 8, inv: 9, oca: 10, ap: 11, ocl: 12, cfo: 13, capex: 14, cfi: 15, draw: 16, repay: 17, rvd: 18, rvr: 19, div: 20, cff: 21, net: 22, open: 23, close: 24, cbr: 26, fcf: 27 };
    const days = `(365/${S_K})`, rate = (key, j) => `((1+${drvRef(key, j)})^(1/${S_K})-1)`;
    const is = (r, c) => qr(IS, `${c}${r}`), bs = (r, c) => qr(BSn, `${c}${r}`), cf = (r, c) => qr(CFn, `${c}${r}`);

    // Income statement
    ISs.put(`B${I.rev}`, `=${op('revenue')}/${S_K}`, Z, 'link');
    row(ISs, I.rev, 'Revenue', (j, c, p) => `${p}${I.rev}*(1+${rate('growth', j)})`, 'total');
    row(ISs, I.cogs, 'Cost of goods sold', (j, c) => `${c}${I.rev}*(1-${drvRef('grossMargin', j)})`);
    row(ISs, I.gp, 'Gross profit', (j, c) => `${c}${I.rev}-${c}${I.cogs}`, 'total');
    row(ISs, I.opex, 'Operating expenses', (j, c) => `${c}${I.rev}*${drvRef('opexPct', j)}`);
    row(ISs, I.ebitda, 'EBITDA', (j, c) => `${c}${I.gp}-${c}${I.opex}`, 'total');
    row(ISs, I.da, 'Depreciation & amortization', (j, c) => `${c}${I.rev}*${drvRef('daPct', j)}`);
    row(ISs, I.ebit, 'EBIT', (j, c) => `${c}${I.ebitda}-${c}${I.da}`, 'total');
    const revOn = m.settings.revolver.enabled;
    row(ISs, I.int, 'Interest expense', (j, c, p) => {
      const term = `IF(${S_SW}=1,(${bs(B.td, p)}+${bs(B.td, c)})/2,${bs(B.td, p)})*${drvRef('interestRate', j)}/${S_K}`;
      // The revolver term is circular under average balances, so it only appears when the revolver is on
      return revOn ? `${term}+IF(${S_SW}=1,(${bs(B.rv, p)}+${bs(B.rv, c)})/2,${bs(B.rv, p)})*${S_RR}/${S_K}` : term;
    });
    row(ISs, I.pbt, 'Profit before tax', (j, c) => `${c}${I.ebit}-${c}${I.int}`, 'total');
    row(ISs, I.tax, 'Income tax', (j, c) => `MAX(0,${c}${I.pbt}*${drvRef('taxRate', j)})`);
    row(ISs, I.ni, 'Net income', (j, c) => `${c}${I.pbt}-${c}${I.tax}`, 'output');
    row(ISs, I.div, 'Dividends declared', (j, c) => `MAX(0,${c}${I.ni})*${drvRef('payoutRatio', j)}`);

    // Balance sheet
    row(BSs, B.cash, 'Cash', (j, c) => `${cf(C.close, c)}`, 'link', `=${op('cash')}`);
    row(BSs, B.ar, 'Accounts receivable', (j, c) => `${is(I.rev, c)}*${drvRef('dso', j)}/${days}`, 'calc', `=${op('receivables')}`);
    row(BSs, B.inv, 'Inventory', (j, c) => `${is(I.cogs, c)}*${drvRef('dio', j)}/${days}`, 'calc', `=${op('inventory')}`);
    row(BSs, B.oca, 'Other current assets', (j, c) => `${is(I.rev, c)}*${S_K}*${drvRef('otherCurrentAssetsPct', j)}`, 'calc', `=${op('otherCurrentAssets')}`);
    BSs.put(`B${B.tca}`, `=SUM(B${B.cash}:B${B.oca})`, Z, 'total');
    row(BSs, B.tca, 'Total current assets', (j, c) => `SUM(${c}${B.cash}:${c}${B.oca})`, 'total');
    row(BSs, B.ppe, 'PP&E and other non-current assets', (j, c, p) => `${p}${B.ppe}-${cf(C.capex, c)}-${is(I.da, c)}`, 'calc', `=${op('ppe')}`);
    BSs.put(`B${B.ta}`, `=B${B.tca}+B${B.ppe}`, Z, 'output');
    row(BSs, B.ta, 'Total assets', (j, c) => `${c}${B.tca}+${c}${B.ppe}`, 'output');
    row(BSs, B.ap, 'Accounts payable', (j, c) => `${is(I.cogs, c)}*${drvRef('dpo', j)}/${days}`, 'calc', `=${op('payables')}`);
    row(BSs, B.ocl, 'Other current liabilities', (j, c) => `${is(I.rev, c)}*${S_K}*${drvRef('otherCurrentLiabilitiesPct', j)}`, 'calc', `=${op('otherCurrentLiabilities')}`);
    BSs.put(`B${B.tcl}`, `=B${B.ap}+B${B.ocl}`, Z, 'total');
    row(BSs, B.tcl, 'Total current liabilities', (j, c) => `${c}${B.ap}+${c}${B.ocl}`, 'total');
    row(BSs, B.td, 'Term debt', (j, c, p) => `${p}${B.td}+${cf(C.draw, c)}+${cf(C.repay, c)}`, 'calc', `=${op('termDebt')}`);
    row(BSs, B.rv, 'Revolver', (j, c, p) => `${p}${B.rv}+${cf(C.rvd, c)}+${cf(C.rvr, c)}`, 'calc', `=${op('revolver')}`);
    row(BSs, B.ol, 'Other non-current liabilities', (j, c, p) => `${p}${B.ol}`, 'calc', `=${op('otherLiabilities')}`);
    BSs.put(`B${B.tl}`, `=B${B.tcl}+B${B.td}+B${B.rv}+B${B.ol}`, Z, 'total');
    row(BSs, B.tl, 'Total liabilities', (j, c) => `${c}${B.tcl}+${c}${B.td}+${c}${B.rv}+${c}${B.ol}`, 'total');
    row(BSs, B.cc, 'Contributed capital', (j, c, p) => `${p}${B.cc}`, 'calc', `=${op('contributedCapital')}`);
    row(BSs, B.re, 'Retained earnings', (j, c, p) => `${p}${B.re}+${is(I.ni, c)}-${is(I.div, c)}`, 'calc', `=${op('retainedEarnings')}`);
    BSs.put(`B${B.eq}`, `=B${B.cc}+B${B.re}`, Z, 'total');
    row(BSs, B.eq, 'Total equity', (j, c) => `${c}${B.cc}+${c}${B.re}`, 'total');
    BSs.put(`B${B.tle}`, `=B${B.tl}+B${B.eq}`, Z, 'output');
    row(BSs, B.tle, 'Total liabilities & equity', (j, c) => `${c}${B.tl}+${c}${B.eq}`, 'output');
    BSs.put(`B${B.chk}`, `=B${B.ta}-B${B.tle}`, ZC, 'calc');
    row(BSs, B.chk, 'Balance check (must be zero)', (j, c) => `${c}${B.ta}-${c}${B.tle}`);
    BSs.put(`B${B.nwc}`, `=B${B.ar}+B${B.inv}+B${B.oca}-B${B.ap}-B${B.ocl}`, Z, 'calc');
    row(BSs, B.nwc, 'Memo: net working capital', (j, c) => `${c}${B.ar}+${c}${B.inv}+${c}${B.oca}-${c}${B.ap}-${c}${B.ocl}`);
    for (let j = 0; j < P; j++) BSs.sh.cells[`${pc(j)}${B.chk}`].z = ZC;

    // Cash flow statement (inflows positive, outflows negative)
    row(CFs, C.ni, 'Net income', (j, c) => is(I.ni, c), 'link');
    row(CFs, C.da, 'Depreciation & amortization', (j, c) => is(I.da, c), 'link');
    row(CFs, C.ar, '(Increase) / decrease in receivables', (j, c, p) => `${bs(B.ar, p)}-${bs(B.ar, c)}`);
    row(CFs, C.inv, '(Increase) / decrease in inventory', (j, c, p) => `${bs(B.inv, p)}-${bs(B.inv, c)}`);
    row(CFs, C.oca, '(Increase) / decrease in other current assets', (j, c, p) => `${bs(B.oca, p)}-${bs(B.oca, c)}`);
    row(CFs, C.ap, 'Increase / (decrease) in payables', (j, c, p) => `${bs(B.ap, c)}-${bs(B.ap, p)}`);
    row(CFs, C.ocl, 'Increase / (decrease) in other current liabilities', (j, c, p) => `${bs(B.ocl, c)}-${bs(B.ocl, p)}`);
    row(CFs, C.cfo, 'Cash from operations', (j, c) => `SUM(${c}${C.ni}:${c}${C.ocl})`, 'total');
    row(CFs, C.capex, 'Capital expenditure', (j, c) => `-${is(I.rev, c)}*${drvRef('capexPct', j)}`);
    row(CFs, C.cfi, 'Cash from investing', (j, c) => `${c}${C.capex}`, 'total');
    row(CFs, C.draw, 'Term debt drawn', (j) => `${drvRef('debtDrawdown', j)}/${S_K}`);
    row(CFs, C.repay, 'Term debt repaid', (j, c, p) => `-MIN(${bs(B.td, p)}+${c}${C.draw},${drvRef('debtRepayment', j)}/${S_K})`);
    row(CFs, C.rvd, 'Revolver drawn', (j, c, p) => `IF(${S_RV}=1,IF(${c}${C.cbr}<${S_RM},MIN(MAX(0,${S_RL}-${bs(B.rv, p)}),${S_RM}-${c}${C.cbr}),0),0)`);
    row(CFs, C.rvr, 'Revolver repaid', (j, c, p) => `-IF(${S_RV}=1,IF(${c}${C.cbr}<${S_RM},0,MIN(${bs(B.rv, p)},${c}${C.cbr}-${S_RM})),0)`);
    row(CFs, C.div, 'Dividends paid', (j, c) => `-${is(I.div, c)}`);
    row(CFs, C.cff, 'Cash from financing', (j, c) => `SUM(${c}${C.draw}:${c}${C.div})`, 'total');
    row(CFs, C.net, 'Net change in cash', (j, c) => `${c}${C.cfo}+${c}${C.cfi}+${c}${C.cff}`, 'total');
    row(CFs, C.open, 'Opening cash', (j, c, p) => bs(B.cash, p), 'link');
    row(CFs, C.close, 'Closing cash', (j, c) => `${c}${C.open}+${c}${C.net}`, 'output');
    row(CFs, C.cbr, 'Memo: cash before revolver', (j, c) => `${c}${C.open}+${c}${C.cfo}+${c}${C.cfi}+${c}${C.draw}+${c}${C.repay}+${c}${C.div}`);
    row(CFs, C.fcf, 'Memo: free cash flow (CFO – CapEx)', (j, c) => `${c}${C.cfo}+${c}${C.capex}`);

    // --- DCF ---
    const v = m.valuation, D = addSheet(book, 'DCF', 'Discounted Cash Flow Valuation', subt);
    const dsm = (...x) => qr('DCF', D.assume(...x));
    D.section(5, 'Cost of capital (CAPM)', 3);
    const dRf = dsm(6, 'Risk-free rate', v.riskFree, ZP), dB = dsm(7, 'Beta', v.beta, '0.00'), dM = dsm(8, 'Equity risk premium', v.marketPremium, ZP), dS = dsm(9, 'Size / specific premium', v.sizePremium, ZP);
    const dKd = dsm(10, 'Pre-tax cost of debt', v.preTaxCostOfDebt, ZP), dT = dsm(11, 'Tax rate', `=${drvRef('taxRate', 0)}`, ZP), dE = dsm(12, 'Equity value', v.equityValue, Z), dD = dsm(13, 'Debt value', v.debtValue, Z);
    const dKe = D.output(14, 'Cost of equity', `=${dRf}+${dB}*${dM}+${dS}`, ZP, 'calc', 'Risk-free + beta × premium + size premium');
    const dKdA = D.output(15, 'After-tax cost of debt', `=${dKd}*(1-${dT})`, ZP, 'calc');
    const dWc = D.output(16, 'WACC (calculated)', `=IF(${dE}+${dD}>0,${dE}/(${dE}+${dD})*${dKe}+${dD}/(${dE}+${dD})*${dKdA},${dKe})`, ZP, 'calc', 'E/(D+E) × Ke + D/(D+E) × Kd × (1 – t)');
    const dUse = dsm(17, 'Use calculated WACC (1 = yes, 0 = manual)', v.useCalculator ? 1 : 0, '0');
    const dMan = dsm(18, 'Manual WACC', v.wacc, ZP);
    const W = D.output(19, 'WACC used', `=IF(${dUse}=1,${dWc},${dMan})`, ZP);
    D.section(21, 'Terminal value and discounting', 3);
    const dMeth = dsm(22, 'Terminal method (1 = Gordon growth, 2 = exit multiple)', v.method === 'exit' ? 2 : 1, '0');
    const dG = dsm(23, 'Terminal growth', v.terminalGrowth, ZP), dX = dsm(24, 'Exit multiple (x LTM EBITDA)', v.exitMultiple, '0.0"x"');
    const dMid = dsm(25, 'Mid-period discounting (1 = on, 0 = end of period)', v.midYear ? 1 : 0, '0');
    const dSh = dsm(26, 'Diluted shares outstanding', v.shares, '#,##0');
    const R = { ebit: 30, tax: 31, nopat: 32, da: 33, capex: 34, dnwc: 35, fcff: 36, t: 37, df: 38, pv: 39 };
    D.section(28, 'Free cash flow to the firm', P + 1);
    D.put('A29', unitLabel(), null, 'headL');
    for (let j = 0; j < P; j++) D.put(`${pc(j)}29`, periodLabel(j), null, 'headF');
    const drow = (r, label, f, z = Z, s = 'calc') => { D.put(`A${r}`, label, null, s === 'output' ? 'labelB' : 'label'); for (let j = 0; j < P; j++) D.put(`${pc(j)}${r}`, '=' + f(j, pc(j), j ? pc(j - 1) : 'B'), z, s); };
    drow(R.ebit, 'EBIT', (j, c) => is(I.ebit, c), Z, 'link');
    drow(R.tax, 'Tax rate', (j) => drvRef('taxRate', j), ZP, 'link');
    drow(R.nopat, 'NOPAT = EBIT × (1 – tax rate)', (j, c) => `${c}${R.ebit}*(1-${c}${R.tax})`);
    drow(R.da, 'Plus: D&A', (j, c) => is(I.da, c), Z, 'link');
    drow(R.capex, 'Less: CapEx', (j, c) => cf(C.capex, c), Z, 'link');
    drow(R.dnwc, 'Less: increase in net working capital', (j, c, p) => `-(${bs(B.nwc, c)}-${bs(B.nwc, p)})`);
    drow(R.fcff, 'FCFF', (j, c) => `${c}${R.nopat}+${c}${R.da}+${c}${R.capex}+${c}${R.dnwc}`, Z, 'output');
    drow(R.t, 'Discount period (years)', (j) => `${j + 1}/${S_K}-${dMid}*0.5/${S_K}`, '0.000');
    drow(R.df, 'Discount factor', (j, c) => `1/(1+${W})^${c}${R.t}`, '0.0000');
    drow(R.pv, 'Present value', (j, c) => `${c}${R.fcff}*${c}${R.df}`);
    const L0 = pc(P - k), L1 = pc(P - 1), lastP = pc(P - 1);
    D.section(41, 'Valuation', 3);
    const vSum = D.output(42, 'PV of forecast FCFF', `=SUM(${pc(0)}${R.pv}:${lastP}${R.pv})`, Z, 'calc');
    const vLf = D.output(43, 'LTM FCFF (last 12 months)', `=SUM(${L0}${R.fcff}:${L1}${R.fcff})`, Z, 'calc');
    const vLe = D.output(44, 'LTM EBITDA (last 12 months)', `=SUM(${is(I.ebitda, L0)}:${L1}${I.ebitda})`, Z, 'calc');
    const vTV = D.output(45, 'Terminal value', `=IF(${dMeth}=2,${vLe}*${dX},IF(${dG}<${W},${vLf}*(1+${dG})/(${W}-${dG}),0))`, Z, 'calc', 'Gordon: LTM FCFF × (1 + g) ÷ (WACC – g); exit: LTM EBITDA × multiple');
    const vPTV = D.output(46, 'PV of terminal value', `=${vTV}/(1+${W})^${n}`, Z, 'calc', 'Discounted from the end of the forecast');
    const vEV = D.output(47, 'Enterprise value', `=${vSum}+${vPTV}`, Z);
    const vND = D.output(48, 'Less: net debt (opening debt – cash)', `=${op('termDebt')}+${op('revolver')}-${op('cash')}`, Z, 'calc');
    const vEq = D.output(49, 'Equity value', `=${vEV}-${vND}`, Z);
    const vPS = D.output(50, 'Value per share', `=${vEq}/${dSh}`, '#,##0.00');
    const vTS = D.output(51, 'Terminal value share of EV', `=IF(${vEV}<>0,${vPTV}/${vEV},0)`, ZP, 'calc');
    D.sh.widths = { A: 52, B: 14, C: 14, ...Object.fromEntries(Array.from({ length: P }, (_, j) => [pc(j), 12])) };

    // --- Checks ---
    const K = addSheet(book, 'Checks', 'Integrity Checks', `${subt} · every check must pass`, '2E7D4F');
    K.put('A5', 'Check', null, 'headL'); K.put('B5', 'Status', null, 'headL'); K.put('C5', 'Rule', null, 'headL');
    for (let j = 0; j < P; j++) K.put(`${colL(3 + j)}5`, periodLabel(j), null, 'headF');
    const pcol = (j) => colL(3 + j), first = pcol(0), lastC = pcol(P - 1);
    const kRow = (r, label, rule, f) => {
      K.put(`A${r}`, label, null, 'labelB'); K.put(`C${r}`, rule, null, 'how');
      for (let j = 0; j < P; j++) K.put(`${pcol(j)}${r}`, '=' + f(j, pc(j), j ? pc(j - 1) : 'B'), ZC, 'calc');
      K.put(`B${r}`, `=IF(AND(MAX(${first}${r}:${lastC}${r})<=0.01,MIN(${first}${r}:${lastC}${r})>=-0.01),"✅ PASS","❌ FAIL")`, null, 'output');
    };
    kRow(6, 'Balance sheet balances', 'Assets – (liabilities + equity) = 0 (tolerance 0.01)', (j, c) => bs(B.chk, c));
    kRow(7, 'Cash ties', 'Cash flow closing cash – balance sheet cash = 0', (j, c) => `${cf(C.close, c)}-${bs(B.cash, c)}`);
    kRow(8, 'Retained earnings roll forward', 'Opening RE + net income – dividends – closing RE = 0', (j, c, p) => `${bs(B.re, p)}+${is(I.ni, c)}-${is(I.div, c)}-${bs(B.re, c)}`);
    kRow(9, 'Net income links', 'Income statement NI – cash flow NI = 0', (j, c) => `${is(I.ni, c)}-${cf(C.ni, c)}`);
    kRow(10, 'Fixed assets roll forward', 'Opening PP&E + CapEx – D&A – closing PP&E = 0', (j, c, p) => `${bs(B.ppe, p)}-${cf(C.capex, c)}-${is(I.da, c)}-${bs(B.ppe, c)}`);
    kRow(11, 'Debt rolls forward', 'Opening debt + draws – repayments – closing debt = 0', (j, c, p) => `${bs(B.td, p)}+${bs(B.rv, p)}+${cf(C.draw, c)}+${cf(C.repay, c)}+${cf(C.rvd, c)}+${cf(C.rvr, c)}-${bs(B.td, c)}-${bs(B.rv, c)}`);
    K.put('A12', 'No negative cash', null, 'labelB'); K.put('C12', 'Cash never below zero (turn on the revolver or add funding)', null, 'how');
    for (let j = 0; j < P; j++) K.put(`${pcol(j)}12`, `=${bs(B.cash, pc(j))}`, Z, 'calc');
    K.put('B12', `=IF(MIN(${first}12:${lastC}12)>=-0.01,"✅ PASS","❌ WARNING")`, null, 'output');
    K.put('A13', 'Circular references controlled', null, 'labelB');
    K.put('C13', revOn ? 'Revolver interest on average balances is circular: iterative calculation is on (100 iterations, 0.000001)' : 'No circular references: term debt follows its schedule', null, 'how');
    K.put('B13', `=IF(${S_SW}=1,"✅ ITERATING (switch on)","✅ OFF (opening balances)")`, null, 'output');
    K.put('A14', 'Valuation sanity', null, 'labelB'); K.put('C14', 'Terminal growth < WACC, terminal value ≤ 85% of EV, WACC within 4–20%', null, 'how');
    K.put('B14', `=IF(AND(OR(${dMeth}=2,${dG}<${W}),${vTS}<=0.85,${W}>=0.04,${W}<=0.2),"✅ PASS","❌ WARNING")`, null, 'output');
    K.sh.widths = { A: 34, B: 26, C: 70, ...Object.fromEntries(Array.from({ length: P }, (_, j) => [pcol(j), 11])) };
    K.sh.freeze = { x: 3, y: 5 };

    // --- Cover ---
    const Cv = addSheet(book, 'Cover', `Financial Model: ${m.company || m.name}`, `Prepared with Finance Studio on ${new Date().toLocaleDateString()}`);
    let r = 5;
    [['Model', m.name], ['Forecast', `${m.years} years, ${m.periodType}, ${periodLabel(0)} to ${periodLabel(P - 1)}`], ['Units', `All figures in ${unitLabel()} (number formats; cells hold full values)`], ['Source', m.source || 'Manual']]
      .forEach(([a, b]) => { Cv.put(`A${r}`, a, null, 'labelB'); Cv.put(`B${r++}`, b, null, 'label'); });
    r++; Cv.section(r++, 'Key results', 2);
    [['Enterprise value', `=${vEV}`, Z], ['Equity value', `=${vEq}`, Z], ['Value per share', `=${vPS}`, '#,##0.00'], ['WACC', `=${W}`, ZP], ['Balance sheet check', `=Checks!B6`, null], ['Valuation sanity', `=Checks!B14`, null]]
      .forEach(([a, f, z]) => { Cv.put(`A${r}`, a, null, 'labelB'); Cv.put(`B${r++}`, f, z, 'output'); });
    r++; Cv.section(r++, 'Colour code and sign convention', 2);
    [['input', 'Input', 'Blue: change these'], ['calc', 'Formula', 'Black: calculated'], ['link', 'Link', 'Green: pulled from another sheet'],
      ['label', 'Signs', 'Costs on the income statement are positive and subtracted; on the cash flow statement outflows are negative.']].forEach(([s, a, b]) => { Cv.put(`A${r}`, a, null, s); Cv.put(`B${r++}`, b, null, 'label'); });
    r++; Cv.put(`A${r}`, 'Models are for informational purposes only and are not financial advice.', null, 'note');
    Cv.sh.widths = { A: 30, B: 90 };

    const order = ['Cover', A, IS, BSn, CFn, 'DCF', 'Checks'];
    book.sheets.sort((a, b) => order.indexOf(a.name) - order.indexOf(b.name));
    book.iterative = revOn;
    evaluateBook(book);
    book.refs = { ev: vEV, eq: vEq, ps: vPS, wacc: W, rows: { I, B, C }, pc };
    return book;
  }
  async function downloadExcel() {
    const el = $('#bMsg');
    try {
      el.textContent = 'Building your workbook…';
      const buf = await bookToBuffer(builderBook());
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      a.download = `${bm.name.replace(/[^\w-]+/g, '-').toLowerCase() || 'model'}.xlsx`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      el.textContent = 'Downloaded.';
    } catch (e) {
      console.error('Finance Studio: Excel export failed', e.message);
      el.textContent = e.message || 'The workbook could not be built.';
    }
  }

  /* ---------- Events ---------- */
  function init() {
    const root = $('#view-builder');
    if (!root) return;
    let timer;
    root.addEventListener('input', (e) => {
      const el = e.target;
      if (el.dataset.text) { bm[el.dataset.text] = el.value.slice(0, 80); saveDraft(); return; }
      if (!el.dataset.path || el.tagName !== 'INPUT') return;
      if (!onInput(el)) return;
      clearTimeout(timer);
      timer = setTimeout(() => { const id = el.dataset.path, pos = el.selectionStart; render(); const again = root.querySelector(`[data-path="${CSS.escape(id)}"]`); if (again) { again.focus(); try { again.setSelectionRange(pos, pos); } catch (x) {} } }, 450);
    });
    root.addEventListener('change', (e) => {
      const el = e.target;
      if (el.id === 'bImport') { importJSON(el.files[0]); el.value = ''; return; }
      if (el.tagName !== 'SELECT' || !el.dataset.path) return;
      const x = el.dataset.bool ? el.value === '1' : ['years'].includes(el.dataset.path) ? +el.value : el.value;
      setPath(el.dataset.path, x);
      if (el.dataset.path === 'periodType') statementView = 'annual';
      render();
    });
    root.addEventListener('click', (e) => {
      const t = e.target.closest('button, td.tr');
      if (!t) return;
      if (t.matches('td.tr')) { showTrace(t); return; }
      if (t.dataset.step != null) { step = +t.dataset.step; render(); return; }
      if (t.id === 'bPrev' && step > 0) { step--; render(); window.scrollTo(0, root.offsetTop); }
      if (t.id === 'bNext' && step < STEPS.length - 1) { step++; render(); window.scrollTo(0, root.offsetTop); }
      if (t.dataset.viewSt) { statementView = t.dataset.viewSt; render(); }
      if (t.dataset.fill) { const a = bm.assumptions[t.dataset.fill]; a.fill(a[0]); render(); }
      if (t.dataset.load) { const s = readSaved().find((x) => x.id === t.dataset.load); const mm = s && validModel(s.versions[0]?.data); if (mm) { bm = mm; render(); } }
      if (t.dataset.del) {
        const s = readSaved().find((x) => x.id === t.dataset.del);
        // Never delete without confirmation
        if (s && confirm(`Delete “${s.name}” and all ${s.versions.length} saved versions? This cannot be undone.`)) { writeSaved(readSaved().filter((x) => x.id !== s.id)); render(); }
      }
      if (t.dataset.restore != null) {
        const s = readSaved().find((x) => x.id === bm.id), mm = s && validModel(s.versions[+t.dataset.restore]?.data);
        if (mm && confirm('Replace the current inputs with this saved version? Save first if you want to keep them.')) { bm = mm; render(); }
      }
      const act = t.dataset.act;
      if (act === 'save') saveVersion();
      if (act === 'json') exportJSON();
      if (act === 'excel') downloadExcel();
      if (act === 'fromData' && confirm('Replace the current inputs with figures from the loaded statements?')) { bm = modelFromData(); render(); }
      if (act === 'new' && confirm('Start a new blank example model? Unsaved changes to this model will be lost.')) { bm = blankModel(); step = 0; render(); }
    });
    root.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.matches('td.tr')) showTrace(e.target); });
    document.addEventListener('click', (e) => { if (!e.target.closest('#bTrace, td.tr')) $('#bTrace')?.remove(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') $('#bTrace')?.remove(); });
  }

  window.renderBuilder = () => { try { render(); } catch (e) { console.error('Finance Studio: builder failed to render', e.message); } };
  window.initBuilder = init;
  window.builderState = () => ({ model: bm, result });
  window.builderBook = builderBook;
  window.builderSetModel = (m) => { const v = validModel(m); if (v) { bm = v; render(); } return !!v; };
})();
