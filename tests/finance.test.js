// Unit tests for lib/finance with hand-checked answers. Run: npm test  (node --test tests/)
const test = require('node:test');
const assert = require('node:assert/strict');
const F = require('../lib/finance/finance.js');

const close = (actual, expected, tol = 1e-6, msg) => assert.ok(Math.abs(actual - expected) <= tol, `${msg || ''} expected ${expected}, got ${actual}`);

// A simple balanced opening balance sheet: assets 1,000 = liabilities 600 + equity 400
const opening = { revenue: 1000, cash: 100, receivables: 100, inventory: 50, otherCurrentAssets: 0, ppe: 750,
  payables: 50, otherCurrentLiabilities: 0, termDebt: 500, revolver: 0, otherLiabilities: 50, contributedCapital: 300, retainedEarnings: 100 };
const base = { growth: 0.1, grossMargin: 0.4, opexPct: 0.15, daPct: 0.05, capexPct: 0.06, taxRate: 0.25, dso: 36.5, dio: 30.4167, dpo: 30.4167,
  otherCurrentAssetsPct: 0, otherCurrentLiabilitiesPct: 0, payoutRatio: 0, interestRate: 0.06, debtDrawdown: 0, debtRepayment: 100 };

test('NPV: known values', () => {
  close(F.calculateNPV(0.1, [-100, 110]), 0);                       // 110 / 1.1 = 100
  close(F.calculateNPV(0.1, [0, 100, 100]), 173.553719, 1e-6);       // 100/1.1 + 100/1.21
  close(F.calculateNPV(0, [-50, 20, 30]), 0);
});

test('IRR: known values', () => {
  close(F.calculateIRR([-100, 110]), 0.1, 1e-7);
  close(F.calculateIRR([-100, 0, 121]), 0.1, 1e-7);                 // 100 × 1.1² = 121
  close(F.calculateIRR([-1000, 300, 400, 500]), 0.0889633947, 1e-6);  // Excel IRR = 8.896%
  assert.equal(F.calculateIRR([100, 50]), null);                     // no sign change
});

test('PMT: level payment', () => {
  close(F.calculatePayment(0.05, 10, 1000), 129.5045750, 1e-6);      // Excel -PMT(5%,10,1000)
  close(F.calculatePayment(0, 4, 1000), 250);
});

test('periodRate compounds back to the annual rate', () => {
  close(Math.pow(1 + F.periodRate(0.12, 12), 12) - 1, 0.12, 1e-12);
  close(F.periodRate(0.21, 2), 0.1, 1e-12);                          // 1.1² = 1.21
});

test('WACC: CAPM and weights', () => {
  const w = F.calculateWACC({ riskFree: 0.04, beta: 1.2, marketPremium: 0.05, preTaxCostOfDebt: 0.06, taxRate: 0.25, equityValue: 600, debtValue: 400 });
  close(w.costOfEquity, 0.10, 1e-12);                                // 4% + 1.2 × 5%
  close(w.afterTaxCostOfDebt, 0.045, 1e-12);
  close(w.wacc, 0.6 * 0.10 + 0.4 * 0.045, 1e-12);                    // 7.8%
});

test('Annual projection: first year by hand', () => {
  const p = F.projectStatements(base, opening, { years: 3 }).periods[0];
  close(p.revenue, 1100);
  close(p.cogs, 660);
  close(p.ebitda, 1100 - 660 - 165);                                 // 275
  close(p.interest, 500 * 0.06);                                     // opening-balance interest = 30
  close(p.netIncome, (275 - 55 - 30) * 0.75);                        // 142.5
  close(p.receivables, 110);                                         // 1,100 × 36.5 / 365
  close(p.ppe, 750 + 66 - 55);
  close(p.termDebt, 400);
});

const allChecksPass = (proj, val, wacc = 0.09, g = 0.025) => {
  const list = F.runIntegrityChecks(proj, val, wacc, g);
  list.forEach((c) => assert.ok(c.pass, `${c.label}: ${c.detail}`));
  return list;
};

for (const periodType of ['annual', 'quarterly', 'monthly']) {
  test(`All integrity checks pass (${periodType}, opening and average interest)`, () => {
    for (const interestMethod of ['opening', 'average']) {
      const proj = F.projectStatements(base, opening, { periodType, years: 5, interestMethod });
      assert.equal(proj.periods.length, 5 * F.PERIODS_PER_YEAR[periodType]);
      proj.periods.forEach((p) => Object.values(p.checks).forEach((d) => assert.ok(Math.abs(d) <= F.TOLERANCE)));
      const list = allChecksPass(proj, null);
      assert.equal(list.length, 8);
    }
  });
}

test('Quarterly revenue compounds to the annual growth rate', () => {
  const yrs = F.annualise(F.projectStatements(base, opening, { periodType: 'quarterly', years: 3 }));
  close(yrs[1].revenue / yrs[0].revenue, 1.1, 1e-9);
});

test('Revolver draws to keep minimum cash, then repays', () => {
  const heavy = { ...base, debtRepayment: [600, 0, 0], growth: 0.05 };
  const proj = F.projectStatements(heavy, opening, { years: 3, revolver: { enabled: true, limit: 1000, interestRate: 0.08, minimumCash: 50 } });
  const [y1, y2] = proj.periods;
  assert.ok(y1.revolverDraw > 0, 'draws in year 1');
  close(y1.cash, 50, 1e-9, 'cash held at the minimum');
  assert.ok(y2.revolverRepay > 0, 'repays from year 2 cash');
  allChecksPass(proj, null);
});

test('Without a revolver, negative cash is flagged, never hidden', () => {
  const heavy = { ...base, debtRepayment: [600, 0, 0] };
  const list = F.runIntegrityChecks(F.projectStatements(heavy, opening, { years: 3 }), null, 0.09, 0.025);
  const neg = list.find((c) => c.id === 'negativeCash');
  assert.equal(neg.pass, false);
  assert.match(neg.detail, /turn on the revolver/);
  assert.ok(list.find((c) => c.id === 'balance').pass, 'still balances');
});

test('Average-balance interest converges and differs from opening-balance interest', () => {
  const heavy = { ...base, debtRepayment: [600, 0, 0] };
  const rv = { enabled: true, limit: 1000, interestRate: 0.08, minimumCash: 50 };
  const avg = F.projectStatements(heavy, opening, { years: 3, interestMethod: 'average', revolver: rv });
  const opn = F.projectStatements(heavy, opening, { years: 3, interestMethod: 'opening', revolver: rv });
  assert.equal(avg.converged, true);
  assert.ok(avg.iterations > 1, 'iterated');
  // Year 1 term debt goes 500 → 0, so average interest on term debt = 250 × 6% = 15 (plus revolver), opening = 30
  close(opn.periods[0].interest, 30);
  const y1 = avg.periods[0];
  close(y1.interest, 250 * 0.06 + (0 + y1.revolver) / 2 * 0.08, 1e-6);
  allChecksPass(avg, null);
});

// Constant flows: zero growth, capex = D&A, no working-capital change → FCFF = EBIT × (1 – t) every year
const flat = { ...base, growth: 0, capexPct: 0.05, debtRepayment: 0, dso: 36.5, dio: 0, dpo: 0 };
const flatOpening = { ...opening, receivables: 100, inventory: 0, payables: 0, otherLiabilities: 100 };

test('DCF Gordon growth on a level FCFF matches the perpetuity formula', () => {
  const proj = F.projectStatements(flat, flatOpening, { years: 5 });
  const v = F.valueDCF(proj, { wacc: 0.1, terminalGrowth: 0, taxRates: 0.25 });
  const fcff = (1000 - 600 - 150 - 50) * 0.75;                       // EBIT 200 → 150
  v.rows.forEach((r) => close(r.fcff, fcff));
  close(v.enterpriseValue, fcff / 0.1, 1e-6);                        // annuity + discounted perpetuity = perpetuity
  close(v.equityValue, v.enterpriseValue);                           // netDebt 0
});

test('Mid-year convention lifts each PV by (1 + WACC)^0.5', () => {
  const proj = F.projectStatements(flat, flatOpening, { years: 5 });
  const end = F.valueDCF(proj, { wacc: 0.1, terminalGrowth: 0.02, taxRates: 0.25 });
  const mid = F.valueDCF(proj, { wacc: 0.1, terminalGrowth: 0.02, taxRates: 0.25, midYear: true });
  close(mid.sumPV / end.sumPV, Math.sqrt(1.1), 1e-9);
  close(mid.pvTerminal, end.pvTerminal, 1e-9);                       // terminal value discounted from year end
});

test('Exit multiple and EV → equity bridge', () => {
  const proj = F.projectStatements(flat, flatOpening, { years: 5 });
  const v = F.valueDCF(proj, { wacc: 0.1, terminalGrowth: 0.02, method: 'exit', exitMultiple: 8, netDebt: 400, shares: 10, taxRates: 0.25 });
  close(v.terminalValue, 250 * 8);                                   // LTM EBITDA 250 × 8
  close(v.equityValue, v.enterpriseValue - 400);
  close(v.perShare, v.equityValue / 10);
  // implied perpetual growth reproduces the exit value
  close(v.ltmFCFF * (1 + v.impliedGrowth) / (0.1 - v.impliedGrowth), v.terminalValue, 1e-6);
});

test('Valuation sanity warnings', () => {
  const proj = F.projectStatements(flat, flatOpening, { years: 5 });
  const bad = F.valueDCF(proj, { wacc: 0.03, terminalGrowth: 0.04, taxRates: 0.25 });
  assert.ok(bad.error);
  assert.deepEqual(bad.warnings.map((w) => w.id).sort(), ['growthAboveWacc', 'waccRange']);
  const tv = F.valueDCF(proj, { wacc: 0.06, terminalGrowth: 0.04, taxRates: 0.25 });
  assert.ok(tv.tvShare > 0.85 && tv.warnings.some((w) => w.id === 'tvShare'));
  const checks = F.runIntegrityChecks(proj, tv, 0.06, 0.04);
  assert.equal(checks.find((c) => c.id === 'valuation').pass, false);
  const good = F.valueDCF(proj, { wacc: 0.1, terminalGrowth: 0.02, taxRates: 0.25 });
  assert.equal(F.runIntegrityChecks(proj, good, 0.1, 0.02).find((c) => c.id === 'valuation').pass, true);
});
