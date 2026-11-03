/* Finance Studio — shared forecasting engine.
   Drivers come from the last three years of history (falling back to sensible defaults), the opening balance
   sheet comes from the last actual year, and the projection links the income statement, balance sheet and
   cash flow statement so the balance sheet always balances. The Excel builder writes the same logic as
   formulas, so the app and the workbook agree. Depends on state/S/cagr from app.js. */

const r3 = (x) => Math.round(x * 1000) / 1000; // drivers are shown to 0.1%
const avgLast = (arr, n = 3) => {
  const v = (arr || []).filter((x) => x != null && isFinite(x)).slice(-n);
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};
const absS = (k) => (S(k) ? S(k).map((v) => (v == null ? null : Math.abs(v))) : null);
const perRev = (vals) => (vals ? vals.map((v, i) => { const r = S('revenue')?.[i]; return v == null || !r ? null : v / r; }) : null);

// Historical income statement lines, cost lines as positive numbers
function histIS() {
  const n = state.years.length, at = (a, i) => (a ? a[i] : null);
  const rev = S('revenue'), cogs = absS('cogs'), opex = absS('opex'), da = absS('da'), int = absS('interest'), ni = S('netIncome');
  return Array.from({ length: n }, (_, i) => {
    // When interest isn't reported separately, treat it as zero so the effective tax rate can still be measured
    const ebit = at(S('ebit'), i), pbt = ebit != null && at(ni, i) != null ? ebit - (at(int, i) ?? 0) : null;
    return { rev: at(rev, i), cogs: at(cogs, i), gp: at(S('grossProfit'), i), opex: at(opex, i), ebitda: at(S('ebitda'), i), da: at(da, i),
      ebit, int: at(int, i), pbt, tax: pbt != null && at(ni, i) != null ? pbt - at(ni, i) : null, ni: at(ni, i) };
  });
}

// Opening balance sheet = last actual year; anything missing is estimated (and flagged)
function openingBS() {
  const li = state.years.length - 1, at = (k) => (li >= 0 ? S(k)?.[li] ?? null : null), est = {};
  const pick = (k, fallback) => { const v = at(k); if (v != null) return v; est[k] = true; return fallback; };
  const rev = at('revenue') ?? 1_000_000;
  const ta = pick('totalAssets', rev * 0.9);
  const ca = pick('currentAssets', ta * 0.45);
  const cash = pick('cash', Math.min(ca * 0.3, rev * 0.08));
  const cl = pick('currentLiabilities', rev * 0.15);
  const debt = pick('debt', 0);
  const tl = at('totalLiabilities');
  const equity = pick('equity', tl != null ? ta - tl : ta - cl - debt);
  return { rev, ta, ca, cash, cl, debt, equity, nca: ca - cash, ppe: ta - ca, other: ta - cl - debt - equity, nwc: ca - cash - cl, est };
}

// Forecast drivers (fractions), rounded to 0.1% so the app and Excel show identical numbers
function deriveDrivers() {
  const H = histIS(), op = openingBS();
  const ratio = (f) => avgLast(H.map(f));
  const em = ratio((h) => (h.ebitda != null && h.rev ? h.ebitda / h.rev : null));
  let gm = ratio((h) => (h.gp != null && h.rev ? h.gp / h.rev : null));
  let opex = ratio((h) => (h.opex != null && h.rev ? h.opex / h.rev : null));
  if (gm == null && opex == null) { opex = 0.2; gm = (em ?? 0.2) + opex; }
  else if (gm == null) gm = (em ?? 0.2) + opex;
  else if (opex == null) opex = em != null ? gm - em : 0.2;
  const tax = ratio((h) => (h.pbt > 0 && h.tax != null ? h.tax / h.pbt : null));
  const rate = ratio((h, i) => { const d = S('debt')?.[i]; return h.int != null && d > 0 ? h.int / d : null; });
  const nca = avgLast(perRev(S('currentAssets')?.map((v, i) => (v == null || S('cash')?.[i] == null ? null : v - S('cash')[i]))));
  const cl = avgLast(perRev(S('currentLiabilities')));
  return {
    growth: r3(cagr(S('revenue')) ?? 0.05),
    gm: r3(gm),
    opex: r3(opex),
    da: r3(avgLast(perRev(absS('da'))) ?? 0.03),
    capex: r3(avgLast(perRev(absS('capex'))) ?? 0.04),
    tax: r3(Math.min(0.4, Math.max(0, tax ?? 0.25))),
    rate: r3(Math.min(0.25, Math.max(0, rate ?? 0.06))),
    nca: r3(nca ?? op.nca / op.rev),
    cl: r3(cl ?? op.cl / op.rev),
    repay: Math.round(op.debt * 0.1),
    payout: 0,
  };
}

// Integrated projection. Drivers may be numbers (constant) or arrays (one value per forecast year).
function project3(dr, op, n) {
  const d = (k, t) => (Array.isArray(dr[k]) ? dr[k][t] : dr[k]);
  const Y = [];
  let p = { rev: op.rev, cash: op.cash, noncash: op.nca, ppe: op.ppe, cl: op.cl, debt: op.debt, other: op.other, eq: op.equity };
  for (let t = 0; t < n; t++) {
    const y = {};
    y.rev = p.rev * (1 + d('growth', t));
    y.cogs = y.rev * (1 - d('gm', t));
    y.gp = y.rev - y.cogs;
    y.opex = y.rev * d('opex', t);
    y.ebitda = y.gp - y.opex;
    y.da = y.rev * d('da', t);
    y.ebit = y.ebitda - y.da;
    y.int = p.debt * d('rate', t);
    y.pbt = y.ebit - y.int;
    y.tax = Math.max(0, y.pbt * d('tax', t));
    y.ni = y.pbt - y.tax;
    y.noncash = y.rev * d('nca', t);
    y.cl = y.rev * d('cl', t);
    y.dnca = p.noncash - y.noncash;
    y.dcl = y.cl - p.cl;
    y.cfo = y.ni + y.da + y.dnca + y.dcl;
    y.capex = -y.rev * d('capex', t);
    y.cfi = y.capex;
    y.repay = -Math.min(p.debt, d('repay', t));
    y.div = -Math.max(0, y.ni * d('payout', t));
    y.cff = y.repay + y.div;
    y.net = y.cfo + y.cfi + y.cff;
    y.openCash = p.cash;
    y.cash = p.cash + y.net;
    y.fcf = y.cfo + y.capex;
    y.ppe = p.ppe - y.capex - y.da;
    y.tca = y.cash + y.noncash;
    y.ta = y.tca + y.ppe;
    y.debt = p.debt + y.repay;
    y.other = p.other;
    y.tl = y.cl + y.debt + y.other;
    y.eq = p.eq + y.ni + y.div;
    y.tle = y.tl + y.eq;
    y.check = y.ta - y.tle;
    y.nwc = y.noncash - y.cl;
    y.dnwc = y.nwc - (p.noncash - p.cl);
    y.ufcf = y.ebit * (1 - d('tax', t)) + y.da + y.capex - y.dnwc;
    Y.push(y);
    p = { rev: y.rev, cash: y.cash, noncash: y.noncash, ppe: y.ppe, cl: y.cl, debt: y.debt, other: y.other, eq: y.eq };
  }
  return Y;
}

// Forecast accuracy backtest: for each of the last (up to) three years, rebuild the drivers from the years
// before it only, forecast that year with the same engine, and compare with what actually happened.
// Accuracy = 100% – mean absolute percentage error.
function backtest() {
  const n = state.years.length;
  if (!hasData() || n < 3 || !S('revenue')) return null;
  const full = state, rows = [];
  try {
    for (let t = Math.max(2, n - 3); t < n; t++) {
      state = { ...full, years: full.years.slice(0, t), items: Object.fromEntries(Object.entries(full.items).map(([k, a]) => [k, a.slice(0, t)])) };
      const y = project3(deriveDrivers(), openingBS(), 1)[0];
      const act = (k) => full.items[k]?.[t] ?? null;
      [['Revenue', y.rev, act('revenue')], ['EBITDA', y.ebitda, act('ebitda')], ['Net income', y.ni, act('netIncome')]].forEach(([metric, forecast, actual]) => {
        if (actual != null && actual !== 0) rows.push({ year: full.years[t], metric, forecast, actual, err: Math.abs(forecast - actual) / Math.abs(actual) });
      });
    }
  } finally {
    state = full;
  }
  if (!rows.length) return null;
  const acc = (list) => Math.max(0, 1 - list.reduce((s, r) => s + r.err, 0) / list.length);
  const byMetric = Object.fromEntries(['Revenue', 'EBITDA', 'Net income'].map((m) => [m, rows.some((r) => r.metric === m) ? acc(rows.filter((r) => r.metric === m)) : null]));
  const vals = Object.values(byMetric).filter((v) => v != null);
  return { rows, byMetric, overall: vals.reduce((a, b) => a + b, 0) / vals.length, tests: new Set(rows.map((r) => r.year)).size };
}
