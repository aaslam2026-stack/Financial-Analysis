/* Finance Studio — calculation module (lib/finance).
 *
 * Every financial formula in the app lives here and is covered by tests in tests/finance.test.js.
 * The module runs unchanged in the browser (as window.Finance) and in Node (require).
 *
 * SIGN CONVENTION
 *   Income statement: revenue and profits positive; costs are stored as positive amounts and subtracted.
 *   Balance sheet: assets, liabilities and equity positive.
 *   Cash flow statement: inflows positive, outflows negative. CapEx, debt repayments, revolver repayments,
 *   dividends and increases in working-capital assets are negative; borrowings and decreases are positive.
 *
 * PERIODS
 *   Forecasts run annually, quarterly or monthly for 3–10 years. Annual drivers are converted per period:
 *   growth and discount rates compound ((1 + r)^(1/k) – 1), interest rates and amounts per year divide by k,
 *   and working-capital days use the days in the period (365 / k).
 *
 * PRECISION
 *   Values are IEEE-754 doubles (about 15 significant digits). Integrity checks use a 0.01 tolerance, which is
 *   many orders of magnitude above the rounding error of the sums involved.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.Finance = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const PERIODS_PER_YEAR = { annual: 1, quarterly: 4, monthly: 12 };
  const TOLERANCE = 0.01;
  const MAX_ITERATIONS = 100;

  /* ---------- Basic time-value functions ---------- */
  function periodRate(annualRate, periodsPerYear) {
    return Math.pow(1 + annualRate, 1 / periodsPerYear) - 1;
  }
  // Net present value: cashFlows[0] occurs today (undiscounted), cashFlows[t] at the end of period t
  function calculateNPV(rate, cashFlows) {
    return cashFlows.reduce((sum, cf, t) => sum + cf / Math.pow(1 + rate, t), 0);
  }
  // Internal rate of return by bisection; null when the cash flows never change sign
  function calculateIRR(cashFlows) {
    let lo = -0.9999, hi = 10, fLo = calculateNPV(lo, cashFlows);
    if (fLo * calculateNPV(hi, cashFlows) > 0) return null;
    for (let i = 0; i < 300; i++) {
      const mid = (lo + hi) / 2, f = calculateNPV(mid, cashFlows);
      if (Math.abs(f) < 1e-9) return mid;
      if (f * fLo < 0) hi = mid; else { lo = mid; fLo = f; }
    }
    return (lo + hi) / 2;
  }
  // Level payment on an amortizing loan (positive number = payment amount)
  function calculatePayment(ratePerPeriod, periods, principal) {
    if (ratePerPeriod === 0) return principal / periods;
    return (principal * ratePerPeriod) / (1 - Math.pow(1 + ratePerPeriod, -periods));
  }

  /* ---------- Cost of capital ---------- */
  // CAPM cost of equity and weighted average cost of capital
  function calculateWACC({ riskFree, beta, marketPremium, sizePremium = 0, preTaxCostOfDebt, taxRate, equityValue, debtValue }) {
    const costOfEquity = riskFree + beta * marketPremium + sizePremium;
    const afterTaxCostOfDebt = preTaxCostOfDebt * (1 - taxRate);
    const total = equityValue + debtValue;
    const equityWeight = total > 0 ? equityValue / total : 1, debtWeight = total > 0 ? debtValue / total : 0;
    return { costOfEquity, afterTaxCostOfDebt, equityWeight, debtWeight, wacc: equityWeight * costOfEquity + debtWeight * afterTaxCostOfDebt };
  }

  /* ---------- Three-statement projection ---------- */
  const valueFor = (assumption, yearIndex) => (Array.isArray(assumption) ? assumption[Math.min(yearIndex, assumption.length - 1)] : assumption);

  /**
   * Project linked income statement, balance sheet and cash flow statement.
   * @param assumptions per forecast year (number or array): growth, grossMargin, opexPct, daPct, capexPct, taxRate,
   *        dso, dio, dpo, otherCurrentAssetsPct, otherCurrentLiabilitiesPct, payoutRatio,
   *        interestRate, debtDrawdown (per year), debtRepayment (per year)
   * @param opening last actual balance sheet: revenue (annual), cash, receivables, inventory, otherCurrentAssets, ppe,
   *        payables, otherCurrentLiabilities, termDebt, revolver, otherLiabilities, contributedCapital, retainedEarnings
   * @param settings periodType ('annual' | 'quarterly' | 'monthly'), years (3–10), interestMethod ('opening' | 'average'),
   *        revolver { enabled, limit, interestRate, minimumCash }
   */
  function projectStatements(assumptions, opening, settings = {}) {
    const k = PERIODS_PER_YEAR[settings.periodType || 'annual'];
    const years = Math.min(10, Math.max(1, Math.round(settings.years || 5)));
    const method = settings.interestMethod === 'average' ? 'average' : 'opening';
    const rv = { enabled: false, limit: 0, interestRate: 0, minimumCash: 0, ...(settings.revolver || {}) };
    const daysInPeriod = 365 / k;
    const periods = [];
    let maxIterations = 0, converged = true;
    let prev = { revenue: opening.revenue / k, cash: opening.cash, receivables: opening.receivables, inventory: opening.inventory, otherCurrentAssets: opening.otherCurrentAssets,
      ppe: opening.ppe, payables: opening.payables, otherCurrentLiabilities: opening.otherCurrentLiabilities, termDebt: opening.termDebt, revolver: opening.revolver || 0,
      otherLiabilities: opening.otherLiabilities, contributedCapital: opening.contributedCapital, retainedEarnings: opening.retainedEarnings };

    for (let p = 0; p < years * k; p++) {
      const y = Math.floor(p / k), a = (key) => valueFor(assumptions[key], y);
      const revenue = prev.revenue * (1 + periodRate(a('growth'), k));
      const cogs = revenue * (1 - a('grossMargin'));
      const opex = revenue * a('opexPct');
      const da = revenue * a('daPct');
      const capex = revenue * a('capexPct');
      // Working capital: days-based balances plus other items as a % of annualised revenue
      const receivables = revenue * a('dso') / daysInPeriod;
      const inventory = cogs * a('dio') / daysInPeriod;
      const payables = cogs * a('dpo') / daysInPeriod;
      const otherCurrentAssets = revenue * k * a('otherCurrentAssetsPct');
      const otherCurrentLiabilities = revenue * k * a('otherCurrentLiabilitiesPct');
      // Term debt moves on its schedule only, so its average balance never creates a circular reference
      const debtDrawdown = a('debtDrawdown') / k;
      const debtRepayment = Math.min(prev.termDebt + debtDrawdown, a('debtRepayment') / k);
      const termDebt = prev.termDebt + debtDrawdown - debtRepayment;

      // Interest on the revolver's average balance depends on cash, which depends on interest: iterate to a fixed point
      let revolverClose = prev.revolver, row = null, iterations = 0;
      for (; iterations < (method === 'average' ? MAX_ITERATIONS : 1); iterations++) {
        const termBase = method === 'average' ? (prev.termDebt + termDebt) / 2 : prev.termDebt;
        const revolverBase = method === 'average' ? (prev.revolver + revolverClose) / 2 : prev.revolver;
        const interest = termBase * a('interestRate') / k + revolverBase * rv.interestRate / k;
        const ebitda = revenue - cogs - opex, ebit = ebitda - da, pbt = ebit - interest;
        const tax = Math.max(0, pbt * a('taxRate')), netIncome = pbt - tax;
        const dividends = Math.max(0, netIncome) * a('payoutRatio');
        // Cash flow statement (inflows positive, outflows negative)
        const changeReceivables = prev.receivables - receivables, changeInventory = prev.inventory - inventory, changeOtherAssets = prev.otherCurrentAssets - otherCurrentAssets;
        const changePayables = payables - prev.payables, changeOtherLiabilities = otherCurrentLiabilities - prev.otherCurrentLiabilities;
        const cashFromOperations = netIncome + da + changeReceivables + changeInventory + changeOtherAssets + changePayables + changeOtherLiabilities;
        const cashFromInvesting = -capex;
        const cashBeforeRevolver = prev.cash + cashFromOperations + cashFromInvesting + debtDrawdown - debtRepayment - dividends;
        let revolverDraw = 0, revolverRepay = 0;
        if (rv.enabled) {
          if (cashBeforeRevolver < rv.minimumCash) revolverDraw = Math.min(Math.max(0, rv.limit - prev.revolver), rv.minimumCash - cashBeforeRevolver);
          else revolverRepay = Math.min(prev.revolver, cashBeforeRevolver - rv.minimumCash);
        }
        const newRevolver = prev.revolver + revolverDraw - revolverRepay;
        row = { revenue, cogs, grossProfit: revenue - cogs, opex, ebitda, da, ebit, interest, pbt, tax, netIncome, dividends,
          receivables, inventory, otherCurrentAssets, payables, otherCurrentLiabilities, changeReceivables, changeInventory, changeOtherAssets, changePayables, changeOtherLiabilities,
          cashFromOperations, capex, cashFromInvesting, debtDrawdown, debtRepayment, revolverDraw, revolverRepay,
          cashFromFinancing: debtDrawdown - debtRepayment + revolverDraw - revolverRepay - dividends, termDebt, revolver: newRevolver };
        if (method === 'opening' || Math.abs(newRevolver - revolverClose) < 1e-7) { revolverClose = newRevolver; iterations++; break; }
        revolverClose = newRevolver;
      }
      if (method === 'average' && iterations >= MAX_ITERATIONS) converged = false;
      maxIterations = Math.max(maxIterations, iterations);

      const r = row;
      r.netChangeInCash = r.cashFromOperations + r.cashFromInvesting + r.cashFromFinancing;
      r.openingCash = prev.cash;
      r.closingCashCF = prev.cash + r.netChangeInCash;   // ending cash per the cash flow statement
      r.cash = r.closingCashCF;                          // balance sheet cash
      r.ppe = prev.ppe + r.capex - r.da;
      r.currentAssets = r.cash + r.receivables + r.inventory + r.otherCurrentAssets;
      r.totalAssets = r.currentAssets + r.ppe;
      r.currentLiabilities = r.payables + r.otherCurrentLiabilities;
      r.otherLiabilities = prev.otherLiabilities;
      r.totalDebt = r.termDebt + r.revolver;
      r.totalLiabilities = r.currentLiabilities + r.totalDebt + r.otherLiabilities;
      r.contributedCapital = prev.contributedCapital;
      r.retainedEarnings = prev.retainedEarnings + r.netIncome - r.dividends;
      r.equity = r.contributedCapital + r.retainedEarnings;
      r.netWorkingCapital = r.receivables + r.inventory + r.otherCurrentAssets - r.payables - r.otherCurrentLiabilities;
      r.freeCashFlow = r.cashFromOperations - r.capex;
      // Roll-forward differences (each should be zero)
      r.checks = {
        balance: r.totalAssets - (r.totalLiabilities + r.equity),
        cashTie: r.closingCashCF - r.cash,
        retainedEarnings: prev.retainedEarnings + r.netIncome - r.dividends - r.retainedEarnings,
        netIncomeLink: r.netIncome - (r.cashFromOperations - r.da - r.changeReceivables - r.changeInventory - r.changeOtherAssets - r.changePayables - r.changeOtherLiabilities),
        ppe: prev.ppe + r.capex - r.da - r.ppe,
        debt: prev.termDebt + prev.revolver + r.debtDrawdown + r.revolverDraw - r.debtRepayment - r.revolverRepay - r.totalDebt,
      };
      r.year = y + 1; r.periodInYear = (p % k) + 1; r.index = p;
      periods.push(r);
      prev = { revenue: r.revenue, cash: r.cash, receivables: r.receivables, inventory: r.inventory, otherCurrentAssets: r.otherCurrentAssets, ppe: r.ppe, payables: r.payables,
        otherCurrentLiabilities: r.otherCurrentLiabilities, termDebt: r.termDebt, revolver: r.revolver, otherLiabilities: r.otherLiabilities, contributedCapital: r.contributedCapital,
        retainedEarnings: r.retainedEarnings };
    }
    const openingNWC = opening.receivables + opening.inventory + opening.otherCurrentAssets - opening.payables - opening.otherCurrentLiabilities;
    return { periodsPerYear: k, years, interestMethod: method, revolver: rv, periods, iterations: maxIterations, converged, opening, openingNWC };
  }

  // Sum flows (or take closing balances) of each forecast year, so monthly/quarterly results can be shown annually
  function annualise(projection) {
    const k = projection.periodsPerYear, out = [];
    const flows = ['revenue', 'cogs', 'grossProfit', 'opex', 'ebitda', 'da', 'ebit', 'interest', 'pbt', 'tax', 'netIncome', 'dividends', 'cashFromOperations', 'capex', 'cashFromInvesting',
      'cashFromFinancing', 'netChangeInCash', 'freeCashFlow', 'debtDrawdown', 'debtRepayment', 'revolverDraw', 'revolverRepay'];
    for (let y = 0; y < projection.years; y++) {
      const ps = projection.periods.slice(y * k, (y + 1) * k), last = ps.at(-1), row = { ...last };
      flows.forEach((f) => { row[f] = ps.reduce((s, p) => s + p[f], 0); });
      out.push(row);
    }
    return out;
  }

  /* ---------- DCF valuation ---------- */
  /**
   * Unlevered free cash flow to the firm (FCFF) = EBIT × (1 – tax rate) + D&A – CapEx – increase in net working capital.
   * Discounting: end of period, or mid-period when midYear is true. The terminal value is discounted from the end of
   * the forecast. Terminal value: Gordon growth on the last 12 months' FCFF, or exit multiple × last 12 months' EBITDA.
   */
  function valueDCF(projection, { wacc, terminalGrowth, method = 'gordon', exitMultiple = 10, midYear = false, netDebt = 0, shares = 1, taxRates }) {
    const k = projection.periodsPerYear, ps = projection.periods, n = ps.length, years = n / k;
    let prevNWC = projection.openingNWC, sumPV = 0;
    const rows = ps.map((p, i) => {
      const taxRate = Array.isArray(taxRates) ? taxRates[Math.min(Math.floor(i / k), taxRates.length - 1)] : taxRates;
      const nopat = p.ebit * (1 - taxRate), increaseNWC = p.netWorkingCapital - prevNWC;
      prevNWC = p.netWorkingCapital;
      const fcff = nopat + p.da - p.capex - increaseNWC;
      const time = (i + 1) / k - (midYear ? 0.5 / k : 0);
      const discountFactor = 1 / Math.pow(1 + wacc, time), pv = fcff * discountFactor;
      sumPV += pv;
      return { ebit: p.ebit, taxRate, nopat, da: p.da, capex: p.capex, increaseNWC, fcff, time, discountFactor, pv };
    });
    const ltmFCFF = rows.slice(-k).reduce((s, r) => s + r.fcff, 0), ltmEBITDA = ps.slice(-k).reduce((s, p) => s + p.ebitda, 0);
    const warnings = [];
    if (terminalGrowth >= wacc) warnings.push({ id: 'growthAboveWacc', text: 'Terminal growth must be below WACC.' });
    if (wacc < 0.04 || wacc > 0.2) warnings.push({ id: 'waccRange', text: `WACC of ${(wacc * 100).toFixed(1)}% is outside the usual 4–20% range.` });
    const gordonTV = terminalGrowth < wacc ? (ltmFCFF * (1 + terminalGrowth)) / (wacc - terminalGrowth) : null;
    const exitTV = ltmEBITDA * exitMultiple;
    const terminalValue = method === 'exit' ? exitTV : gordonTV;
    if (terminalValue == null) return { rows, warnings, error: 'Terminal growth must be below WACC.' };
    const pvTerminal = terminalValue / Math.pow(1 + wacc, years);
    const enterpriseValue = sumPV + pvTerminal, equityValue = enterpriseValue - netDebt;
    const tvShare = enterpriseValue ? pvTerminal / enterpriseValue : null;
    if (tvShare != null && tvShare > 0.85) warnings.push({ id: 'tvShare', text: `Terminal value is ${(tvShare * 100).toFixed(0)}% of enterprise value (above 85%).` });
    return {
      rows, sumPV, ltmFCFF, ltmEBITDA, terminalValue, pvTerminal, enterpriseValue, netDebt, equityValue, perShare: equityValue / shares, tvShare, warnings,
      impliedExitMultiple: gordonTV != null && ltmEBITDA ? gordonTV / ltmEBITDA : null,
      impliedGrowth: method === 'exit' ? (exitTV * wacc - ltmFCFF) / (exitTV + ltmFCFF) : terminalGrowth,
    };
  }

  /* ---------- Integrity checks (non-negotiable rules) ---------- */
  function runIntegrityChecks(projection, valuation, wacc, terminalGrowth) {
    const ps = projection.periods, worst = (key) => ps.reduce((m, p) => Math.max(m, Math.abs(p.checks[key])), 0);
    const check = (id, label, diff, rule) => ({ id, label, rule, pass: diff <= TOLERANCE, detail: diff <= TOLERANCE ? 'Holds in every period' : `Out by up to ${diff.toFixed(2)}` });
    const list = [
      check('balance', 'Balance sheet balances', worst('balance'), 'Total assets = total liabilities + equity, every period (tolerance 0.01)'),
      check('cashTie', 'Cash ties', worst('cashTie'), 'Ending cash on the cash flow statement = cash on the balance sheet'),
      check('retainedEarnings', 'Retained earnings roll forward', worst('retainedEarnings'), 'Opening RE + net income – dividends = closing RE'),
      check('netIncomeLink', 'Net income links', worst('netIncomeLink'), 'Income statement net income = first line of the cash flow statement'),
      check('ppe', 'Fixed assets roll forward', worst('ppe'), 'Opening PP&E + CapEx – depreciation = closing PP&E'),
      check('debt', 'Debt rolls forward', worst('debt'), 'Opening debt + drawdowns – repayments = closing debt'),
    ];
    const negative = ps.find((p) => p.cash < -TOLERANCE);
    const rv = projection.revolver;
    list.push({ id: 'negativeCash', label: 'No negative cash', rule: 'Cash never goes below zero unless a revolver funds it (otherwise a visible warning)',
      pass: !negative, detail: !negative ? (rv.enabled ? 'Revolver keeps cash at or above the minimum' : 'Cash stays positive in every period')
        : `Cash falls to ${negative.cash.toFixed(0)} in forecast period ${negative.index + 1}${rv.enabled ? ' even with the revolver at its limit' : '; turn on the revolver or add funding'}` });
    list.push({ id: 'circularity', label: 'Circular references controlled', rule: 'Interest on average debt uses an explicit iteration switch, never a silent loop',
      pass: projection.converged, detail: projection.interestMethod === 'average' ? (projection.converged ? `Average-balance interest converged (up to ${projection.iterations} iterations)` : 'Iteration did not converge; switch interest to opening balances') : 'Interest on opening balances: no circularity' });
    if (valuation) {
      const issues = [];
      if (terminalGrowth >= wacc) issues.push('terminal growth is not below WACC');
      if (valuation.tvShare != null && valuation.tvShare > 0.85) issues.push(`terminal value is ${(valuation.tvShare * 100).toFixed(0)}% of EV`);
      if (wacc < 0.04 || wacc > 0.2) issues.push(`WACC ${(wacc * 100).toFixed(1)}% is outside 4–20%`);
      list.push({ id: 'valuation', label: 'Valuation sanity', rule: 'Terminal growth < WACC, terminal value ≤ 85% of EV, WACC within 4–20%', pass: !issues.length,
        detail: issues.length ? `Warning: ${issues.join('; ')}` : 'All valuation sanity tests pass' });
    }
    return list;
  }

  return { PERIODS_PER_YEAR, TOLERANCE, periodRate, calculateNPV, calculateIRR, calculatePayment, calculateWACC, projectStatements, annualise, valueDCF, runIntegrityChecks };
});
