/* Finance Studio — ratio analysis toolkit for analysts.
   Each ratio is defined once as a small formula tree. The same tree is evaluated in the browser and written
   out as a live Excel formula, so the website and the workbook always agree.
   Also: DuPont analysis, Altman Z''-score, Piotroski F-score, traffic-light assessment and commentary.
   Depends on app.js (state, S, helpers) and engine.js (deriveDrivers, openingBS, project3). */

// Formula tree: a key ('rev', 'ca', ...), a number, ['prev', key] (prior year), or [op, a, b] with op in + - * /
const RATIO_DEFS = [
  // Liquidity
  { id: 'current', cat: 'Liquidity', name: 'Current ratio', f: ['/', 'ca', 'cl'], fmt: 'x', dir: 'high', good: 1.5, weak: 1, how: 'Current assets ÷ current liabilities' },
  { id: 'quick', cat: 'Liquidity', name: 'Quick ratio', f: ['/', ['-', 'ca', 'inv'], 'cl'], fmt: 'x', dir: 'high', good: 1, weak: 0.7, how: '(Current assets – inventory) ÷ current liabilities' },
  { id: 'cashr', cat: 'Liquidity', name: 'Cash ratio', f: ['/', 'cash', 'cl'], fmt: 'x', dir: 'high', good: 0.5, weak: 0.2, how: 'Cash ÷ current liabilities' },
  { id: 'ocfr', cat: 'Liquidity', name: 'Operating cash flow ratio', f: ['/', 'ocf', 'cl'], fmt: 'x', dir: 'high', good: 0.4, weak: 0.2, how: 'Cash from operations ÷ current liabilities' },
  { id: 'nwc', cat: 'Liquidity', name: 'Net working capital', f: ['-', 'ca', 'cl'], fmt: 'money', dir: 'high', good: 0, weak: 0, how: 'Current assets – current liabilities' },
  // Profitability
  { id: 'gm', cat: 'Profitability', name: 'Gross margin', f: ['/', 'gp', 'rev'], fmt: 'pct', dir: 'high', good: 0.4, weak: 0.2, bench: 'gm', how: 'Gross profit ÷ revenue' },
  { id: 'em', cat: 'Profitability', name: 'EBITDA margin', f: ['/', 'ebitda', 'rev'], fmt: 'pct', dir: 'high', good: 0.2, weak: 0.1, bench: 'em', how: 'EBITDA ÷ revenue' },
  { id: 'ebm', cat: 'Profitability', name: 'Operating (EBIT) margin', f: ['/', 'ebit', 'rev'], fmt: 'pct', dir: 'high', good: 0.15, weak: 0.05, how: 'EBIT ÷ revenue' },
  { id: 'nm', cat: 'Profitability', name: 'Net margin', f: ['/', 'ni', 'rev'], fmt: 'pct', dir: 'high', good: 0.1, weak: 0.03, bench: 'nm', how: 'Net income ÷ revenue' },
  { id: 'roa', cat: 'Profitability', name: 'Return on assets (ROA)', f: ['/', 'ni', 'ta'], fmt: 'pct', dir: 'high', good: 0.08, weak: 0.03, how: 'Net income ÷ total assets' },
  { id: 'roe', cat: 'Profitability', name: 'Return on equity (ROE)', f: ['/', 'ni', 'eq'], fmt: 'pct', dir: 'high', good: 0.15, weak: 0.08, bench: 'roe', how: 'Net income ÷ shareholders’ equity' },
  { id: 'roic', cat: 'Profitability', name: 'Return on invested capital (ROIC)', f: ['/', ['*', 'ebit', ['-', 1, 'tax']], ['-', ['+', 'debt', 'eq'], 'cash']], fmt: 'pct', dir: 'high', good: 0.12, weak: 0.06, how: 'EBIT × (1 – tax rate) ÷ (debt + equity – cash)' },
  { id: 'roce', cat: 'Profitability', name: 'Return on capital employed (ROCE)', f: ['/', 'ebit', ['-', 'ta', 'cl']], fmt: 'pct', dir: 'high', good: 0.15, weak: 0.08, how: 'EBIT ÷ (total assets – current liabilities)' },
  // Leverage & solvency
  { id: 'de', cat: 'Leverage', name: 'Debt / equity', f: ['/', 'debt', 'eq'], fmt: 'x', dir: 'low', good: 0.5, weak: 1.5, bench: 'de', how: 'Total debt ÷ shareholders’ equity' },
  { id: 'dta', cat: 'Leverage', name: 'Debt / assets', f: ['/', 'debt', 'ta'], fmt: 'pct', dir: 'low', good: 0.3, weak: 0.6, how: 'Total debt ÷ total assets' },
  { id: 'dcap', cat: 'Leverage', name: 'Debt / capital', f: ['/', 'debt', ['+', 'debt', 'eq']], fmt: 'pct', dir: 'low', good: 0.35, weak: 0.6, how: 'Debt ÷ (debt + equity)' },
  { id: 'nde', cat: 'Leverage', name: 'Net debt / EBITDA', f: ['/', ['-', 'debt', 'cash'], 'ebitda'], fmt: 'x', dir: 'low', good: 2, weak: 4, how: '(Debt – cash) ÷ EBITDA' },
  { id: 'icov', cat: 'Leverage', name: 'Interest coverage', f: ['/', 'ebit', 'int'], fmt: 'x', dir: 'high', good: 5, weak: 2, how: 'EBIT ÷ interest expense' },
  { id: 'emult', cat: 'Leverage', name: 'Equity multiplier', f: ['/', 'ta', 'eq'], fmt: 'x', dir: 'low', good: 2, weak: 4, how: 'Total assets ÷ equity' },
  // Efficiency
  { id: 'at', cat: 'Efficiency', name: 'Asset turnover', f: ['/', 'rev', 'ta'], fmt: 'x', dir: 'high', good: 1, weak: 0.5, how: 'Revenue ÷ total assets' },
  { id: 'fat', cat: 'Efficiency', name: 'Fixed asset turnover', f: ['/', 'rev', ['-', 'ta', 'ca']], fmt: 'x', dir: 'high', good: 2, weak: 1, how: 'Revenue ÷ non-current assets' },
  { id: 'dso', cat: 'Efficiency', name: 'Days sales outstanding (DSO)', f: ['*', ['/', 'ar', 'rev'], 365], fmt: 'days', dir: 'low', good: 45, weak: 75, how: 'Receivables ÷ revenue × 365' },
  { id: 'dio', cat: 'Efficiency', name: 'Days inventory outstanding (DIO)', f: ['*', ['/', 'inv', 'cogs'], 365], fmt: 'days', dir: 'low', good: 60, weak: 120, how: 'Inventory ÷ cost of goods sold × 365' },
  { id: 'dpo', cat: 'Efficiency', name: 'Days payables outstanding (DPO)', f: ['*', ['/', 'ap', 'cogs'], 365], fmt: 'days', dir: 'none', how: 'Payables ÷ cost of goods sold × 365' },
  { id: 'ccc', cat: 'Efficiency', name: 'Cash conversion cycle', f: ['-', ['+', ['*', ['/', 'ar', 'rev'], 365], ['*', ['/', 'inv', 'cogs'], 365]], ['*', ['/', 'ap', 'cogs'], 365]], fmt: 'days', dir: 'low', good: 60, weak: 120, how: 'DSO + DIO – DPO' },
  // Cash flow
  { id: 'fcfm', cat: 'Cash flow', name: 'Free cash flow margin', f: ['/', ['-', 'ocf', 'capex'], 'rev'], fmt: 'pct', dir: 'high', good: 0.08, weak: 0.02, how: '(Cash from operations – CapEx) ÷ revenue' },
  { id: 'fcfconv', cat: 'Cash flow', name: 'Free cash flow conversion', f: ['/', ['-', 'ocf', 'capex'], 'ni'], fmt: 'pct', dir: 'high', good: 0.9, weak: 0.5, how: 'Free cash flow ÷ net income' },
  { id: 'cashconv', cat: 'Cash flow', name: 'Cash conversion (CFO / net income)', f: ['/', 'ocf', 'ni'], fmt: 'x', dir: 'high', good: 1, weak: 0.7, how: 'Cash from operations ÷ net income' },
  { id: 'capint', cat: 'Cash flow', name: 'CapEx intensity', f: ['/', 'capex', 'rev'], fmt: 'pct', dir: 'none', how: 'CapEx ÷ revenue' },
  { id: 'capda', cat: 'Cash flow', name: 'CapEx / depreciation', f: ['/', 'capex', 'da'], fmt: 'x', dir: 'none', how: 'CapEx ÷ D&A (above 1 = investing for growth)' },
  // Growth
  { id: 'revg', cat: 'Growth', name: 'Revenue growth', f: ['-', ['/', 'rev', ['prev', 'rev']], 1], fmt: 'pct', dir: 'high', good: 0.08, weak: 0, bench: 'growth', how: 'Revenue ÷ prior-year revenue – 1' },
  { id: 'ebitdag', cat: 'Growth', name: 'EBITDA growth', f: ['-', ['/', 'ebitda', ['prev', 'ebitda']], 1], fmt: 'pct', dir: 'high', good: 0.08, weak: 0, how: 'EBITDA ÷ prior-year EBITDA – 1' },
  { id: 'nig', cat: 'Growth', name: 'Net income growth', f: ['-', ['/', 'ni', ['prev', 'ni']], 1], fmt: 'pct', dir: 'high', good: 0.08, weak: 0, how: 'Net income ÷ prior-year net income – 1' },
  { id: 'eqg', cat: 'Growth', name: 'Equity growth', f: ['-', ['/', 'eq', ['prev', 'eq']], 1], fmt: 'pct', dir: 'high', good: 0.05, weak: 0, how: 'Equity ÷ prior-year equity – 1' },
];
const RATIO_CATS = ['All', ...new Set(RATIO_DEFS.map((d) => d.cat))];
let ratioUI = { cat: 'All', forecast: false, formulas: false, chart: 'roe' };

// Evaluate a formula tree against one year's data (p = prior year); null when any input is missing
function evalTree(t, d, p) {
  if (typeof t === 'number') return t;
  if (typeof t === 'string') return d[t] ?? null;
  if (t[0] === 'prev') return p ? p[t[1]] ?? null : null;
  const a = evalTree(t[1], d, p), b = evalTree(t[2], d, p);
  if (a == null || b == null) return null;
  if (t[0] === '/') return b === 0 ? null : a / b;
  return t[0] === '+' ? a + b : t[0] === '-' ? a - b : a * b;
}
// The same tree as an Excel formula; ref(key, prior) returns a cell reference or null when the input is missing
function treeToExcel(t, ref) {
  if (typeof t === 'number') return String(t);
  if (typeof t === 'string') return ref(t, false);
  if (t[0] === 'prev') return ref(t[1], true);
  const a = treeToExcel(t[1], ref), b = treeToExcel(t[2], ref);
  return a == null || b == null ? null : `(${a}${t[0]}${b})`;
}

// One object per year: actual years from the uploaded statements, then (optionally) forecast years from the engine
function ratioYears(withForecast) {
  const n = state.years.length, abs = (k, i) => { const v = S(k)?.[i]; return v == null ? null : Math.abs(v); }, at = (k, i) => S(k)?.[i] ?? null;
  const tax = deriveDrivers().tax;
  const years = Array.from({ length: n }, (_, i) => ({
    label: state.years[i], forecast: false, tax,
    rev: at('revenue', i), cogs: abs('cogs', i), gp: at('grossProfit', i), ebitda: at('ebitda', i), ebit: at('ebit', i), da: abs('da', i), int: abs('interest', i), ni: at('netIncome', i),
    cash: at('cash', i), ca: at('currentAssets', i), ta: at('totalAssets', i), cl: at('currentLiabilities', i), tl: at('totalLiabilities', i) ?? (at('totalAssets', i) != null && at('equity', i) != null ? at('totalAssets', i) - at('equity', i) : null),
    debt: at('debt', i), eq: at('equity', i), capex: abs('capex', i), ocf: at('ocf', i),
    ar: at('receivables', i), inv: at('inventory', i), ap: abs('payables', i), re: at('retainedEarnings', i),
  }));
  if (withForecast) {
    const d = defaults();
    project3(d.dr, d.op, 5).forEach((y, t) => years.push({ label: fyLabel(t + 1), forecast: true, tax,
      rev: y.rev, cogs: y.cogs, gp: y.gp, ebitda: y.ebitda, ebit: y.ebit, da: y.da, int: y.int, ni: y.ni, cash: y.cash, ca: y.tca, ta: y.ta, cl: y.cl, tl: y.tl,
      debt: y.debt, eq: y.eq, capex: -y.capex, ocf: y.cfo, ar: null, inv: null, ap: null, re: null }));
  }
  return years;
}

function ratioValue(def, years, i) { return evalTree(def.f, years[i], years[i - 1]); }
function assess(def, v) {
  if (v == null || def.dir === 'none' || !isFinite(v)) return null;
  if (def.dir === 'high') return v >= def.good ? 'good' : v < def.weak ? 'weak' : 'mid';
  return v <= def.good ? 'good' : v > def.weak ? 'weak' : 'mid';
}
function fmtRatio(def, v) {
  if (v == null || !isFinite(v)) return '—';
  return def.fmt === 'pct' ? pct(v) : def.fmt === 'x' ? xf(v) : def.fmt === 'days' ? tag(`${v.toFixed(0)} days`, v, '0 "days"') : fmt(v);
}

// Altman Z''-score (version for private and non-manufacturing companies) and its zones
function altmanZ(d) {
  if ([d.ca, d.cl, d.ta, d.ebit, d.eq, d.tl].some((x) => x == null) || !d.ta || !d.tl) return null;
  const re = d.re ?? d.eq; // retained earnings; falls back to book equity when not reported
  const x1 = (d.ca - d.cl) / d.ta, x2 = re / d.ta, x3 = d.ebit / d.ta, x4 = d.eq / d.tl;
  const z = 6.56 * x1 + 3.26 * x2 + 6.72 * x3 + 1.05 * x4;
  return { z, x1, x2, x3, x4, proxy: d.re == null, zone: z > 2.6 ? 'Safe' : z >= 1.1 ? 'Grey zone' : 'Distress' };
}
// Piotroski F-score: up to 9 pass/fail tests of profitability, leverage/liquidity and efficiency
function piotroski(d, p) {
  if (!p) return null;
  const r = (a, b) => (a == null || b == null || !b ? null : a / b);
  const tests = [
    ['Positive return on assets', r(d.ni, d.ta) != null ? r(d.ni, d.ta) > 0 : null],
    ['Positive operating cash flow', d.ocf != null ? d.ocf > 0 : null],
    ['Return on assets improved', r(d.ni, d.ta) != null && r(p.ni, p.ta) != null ? r(d.ni, d.ta) > r(p.ni, p.ta) : null],
    ['Cash flow exceeds net income (quality of earnings)', d.ocf != null && d.ni != null ? d.ocf > d.ni : null],
    ['Leverage (debt / assets) fell', r(d.debt, d.ta) != null && r(p.debt, p.ta) != null ? r(d.debt, d.ta) < r(p.debt, p.ta) : null],
    ['Current ratio improved', r(d.ca, d.cl) != null && r(p.ca, p.cl) != null ? r(d.ca, d.cl) > r(p.ca, p.cl) : null],
    ['No new shares issued', null],
    ['Gross margin improved', r(d.gp, d.rev) != null && r(p.gp, p.rev) != null ? r(d.gp, d.rev) > r(p.gp, p.rev) : null],
    ['Asset turnover improved', r(d.rev, d.ta) != null && r(p.rev, p.ta) != null ? r(d.rev, d.ta) > r(p.rev, p.ta) : null],
  ];
  const known = tests.filter((t) => t[1] != null);
  return { tests, score: known.filter((t) => t[1]).length, of: known.length };
}

function ratioCommentary(years, li) {
  const good = [], bad = [], b = chosenBenchmark();
  RATIO_DEFS.forEach((def) => {
    const v = ratioValue(def, years, li), pv = li > 0 ? ratioValue(def, years, li - 1) : null, a = assess(def, v);
    if (a == null) return;
    const trend = pv == null ? '' : (def.dir === 'high' ? v > pv : v < pv) ? ', improving' : ', weakening';
    const vs = def.bench && b ? ` (industry ${fmtRatio(def, b[def.bench])})` : '';
    if (a === 'good') good.push(`<b>${def.name}</b> is strong at ${fmtRatio(def, v)}${vs}${trend}.`);
    if (a === 'weak') bad.push(`<b>${def.name}</b> is weak at ${fmtRatio(def, v)}${vs}${trend}.`);
  });
  return { good, bad };
}

/* ---------- Ratio Analysis page ---------- */
function ratioAnalysisView() {
  if (!hasData()) return { html: '<div class="glass card empty"><h3>Upload statements (or load a sample or real company) to run the ratio analysis</h3><a class="btn" href="#upload">Upload data</a></div>' };
  const years = ratioYears(ratioUI.forecast), li = state.years.length - 1, b = chosenBenchmark();
  const defs = RATIO_DEFS.filter((d) => ratioUI.cat === 'All' || d.cat === ratioUI.cat);
  const latest = RATIO_DEFS.map((d) => assess(d, ratioValue(d, years, li))).filter(Boolean);
  const nGood = latest.filter((a) => a === 'good').length, nWeak = latest.filter((a) => a === 'weak').length;
  const z = altmanZ(years[li]), f = piotroski(years[li], years[li - 1]);
  const roeDu = years[li].ni / years[li].eq;
  const com = ratioCommentary(years, li);
  const light = { good: '<span class="pill-s good">Strong</span>', mid: '<span class="pill-s mid">Fair</span>', weak: '<span class="pill-s bad">Weak</span>' };
  const head = ['Ratio', ...(ratioUI.formulas ? ['Formula'] : []), ...years.map((y) => esc(y.label)), ...(b ? ['Industry'] : []), 'Trend', `Rating ${esc(state.years[li])}`];
  const rows = [];
  RATIO_CATS.slice(1).filter((c) => ratioUI.cat === 'All' || c === ratioUI.cat).forEach((cat) => {
    rows.push([`<strong class="rcat">${cat}</strong>`, ...Array(head.length - 1).fill('')]);
    defs.filter((d) => d.cat === cat).forEach((d) => {
      const vals = years.map((_, i) => ratioValue(d, years, i)), v = vals[li], pv = vals[li - 1];
      const tr = v == null || pv == null || d.dir === 'none' ? '—' : (d.dir === 'high' ? v >= pv : v <= pv) ? '<span class="ok">▲ better</span>' : '<span class="bad-t">▼ worse</span>';
      rows.push([d.name, ...(ratioUI.formulas ? [`<span class="muted small">${d.how}</span>`] : []), ...vals.map((x) => fmtRatio(d, x)), ...(b ? [d.bench ? fmtRatio(d, b[d.bench]) : '—'] : []), tr, light[assess(d, v)] || '—']);
    });
  });
  const chartDef = RATIO_DEFS.find((d) => d.id === ratioUI.chart) || RATIO_DEFS[0];
  const scale = chartDef.fmt === 'pct' ? 100 : 1;
  return {
    html: `<div class="ratio-opts glass">
        <div class="cat-tabs">${RATIO_CATS.map((c) => `<button type="button" data-rcat="${c}" class="${c === ratioUI.cat ? 'on' : ''}">${c}</button>`).join('')}</div>
        <div class="row tight-row">
          <button type="button" class="btn ghost btn-sm" data-rtoggle="forecast" aria-pressed="${ratioUI.forecast}">${ratioUI.forecast ? '✓ ' : ''}Include forecast years</button>
          <button type="button" class="btn ghost btn-sm" data-rtoggle="formulas" aria-pressed="${ratioUI.formulas}">${ratioUI.formulas ? '✓ ' : ''}Show formulas</button>
          ${b ? `<span class="muted small">Compared with: <b>${esc(b.name)}</b></span>` : '<a class="link small" href="#indicators">Choose an industry to compare</a>'}
        </div>
      </div>
      <div class="kpis">${kpi('Ratio health score', pct(nGood / (latest.length || 1), 0), `${nGood} strong · ${latest.length - nGood - nWeak} fair · ${nWeak} weak`, nGood >= nWeak ? 'up' : 'down')}
        ${kpi('Altman Z″-score', z ? z.z.toFixed(2) : '—', z ? `${z.zone}${z.proxy ? ' (equity used for retained earnings)' : ''}` : 'needs balance sheet', z ? (z.zone === 'Safe' ? 'up' : 'down') : '')}
        ${kpi('Piotroski F-score', f ? `${f.score} / ${f.of}` : '—', f ? (f.score / f.of >= 0.7 ? 'Financially strong' : f.score / f.of >= 0.4 ? 'Average' : 'Weak') : 'needs two years', f ? (f.score / f.of >= 0.7 ? 'up' : f.score / f.of < 0.4 ? 'down' : '') : '')}
        ${kpi('Return on equity', pct(roeDu), 'DuPont below')}${kpi('Current ratio', xf(years[li].ca / years[li].cl))}${kpi('Net debt / EBITDA', xf((years[li].debt - (years[li].cash || 0)) / years[li].ebitda))}</div>
      <div class="glass card"><h3>Ratio analysis${ratioUI.cat === 'All' ? '' : `: ${ratioUI.cat}`}</h3>${table(head, rows)}
        <p class="muted small">Ratings use common rule-of-thumb thresholds; what is “strong” differs by industry, so compare with the industry column too. Days ratios need receivables, inventory and payables in your upload.${ratioUI.forecast ? ' Columns marked E are forecast years from the 3-statement model.' : ''}</p></div>
      <div class="grid two">
        <div class="glass card"><div class="row spread"><h3>Trend</h3><select id="rChart" aria-label="Ratio to chart">${RATIO_DEFS.map((d) => `<option value="${d.id}"${d.id === chartDef.id ? ' selected' : ''}>${d.name}</option>`).join('')}</select></div><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Commentary</h3>
          ${com.good.length ? `<h4 class="com-h ok">Strengths</h4><ul class="tips">${com.good.slice(0, 6).map((x) => `<li>${x}</li>`).join('')}</ul>` : ''}
          ${com.bad.length ? `<h4 class="com-h bad-t">Watch items</h4><ul class="tips">${com.bad.slice(0, 6).map((x) => `<li>${x}</li>`).join('')}</ul>` : '<p class="muted">No weak ratios on the latest year.</p>'}
        </div>
      </div>
      <div class="grid two">
        <div class="glass card"><h3>DuPont analysis of ROE</h3>${table(['Component', ...years.map((y) => esc(y.label))], [
          ['Net margin', ...years.map((y) => pct(y.ni / y.rev))], ['× Asset turnover', ...years.map((y) => xf(y.rev / y.ta))], ['× Equity multiplier', ...years.map((y) => xf(y.ta / y.eq))], ['<strong>= Return on equity</strong>', ...years.map((y) => pct(y.ni / y.eq))]])}
          <p class="muted small">ROE = (net income ÷ revenue) × (revenue ÷ assets) × (assets ÷ equity). Shows whether returns come from margins, efficiency or leverage.</p></div>
        <div class="glass card"><h3>Credit &amp; quality scores (${esc(state.years[li])})</h3>
          ${z ? table(['Altman Z″ component', 'Value', 'Weight'], [['Working capital ÷ assets', z.x1.toFixed(3), '6.56'], [`Retained earnings ÷ assets${z.proxy ? ' (equity proxy)' : ''}`, z.x2.toFixed(3), '3.26'], ['EBIT ÷ assets', z.x3.toFixed(3), '6.72'], ['Equity ÷ liabilities', z.x4.toFixed(3), '1.05'], ['<strong>Z″-score</strong>', `<strong>${z.z.toFixed(2)}</strong>`, z.zone]]) : '<p class="muted">Needs a full balance sheet.</p>'}
          <p class="muted small">Zones: above 2.6 safe · 1.1 to 2.6 grey · below 1.1 distress.</p>
          ${f ? `<ul class="fscore">${f.tests.map(([t, ok]) => `<li class="${ok == null ? 'na' : ok ? 'pass' : 'fail'}">${ok == null ? '–' : ok ? '✓' : '✗'} ${t}</li>`).join('')}</ul>` : ''}
        </div>
      </div>`,
    charts: [['mc1', { data: { labels: years.map((y) => y.label), datasets: [
      ds(chartDef.name, years.map((_, i) => { const v = ratioValue(chartDef, years, i); return v == null ? null : +(v * scale).toFixed(3); }), COLORS.blue, { type: 'bar' }),
      ...(b && chartDef.bench ? [ds('Industry', years.map(() => +(b[chartDef.bench] * scale).toFixed(3)), COLORS.orange, { type: 'line', borderDash: [6, 4], pointRadius: 0 })] : []),
    ] }, options: { scales: { y: { ticks: { callback: (v) => (chartDef.fmt === 'pct' ? `${v}%` : chartDef.fmt === 'x' ? `${v}x` : v.toLocaleString()) } } } } }]],
  };
}

function initRatios() {
  document.addEventListener('click', (e) => {
    const c = e.target.closest('[data-rcat]'), t = e.target.closest('[data-rtoggle]');
    if (c) { ratioUI.cat = c.dataset.rcat; runModel(); }
    if (t) { ratioUI[t.dataset.rtoggle] = !ratioUI[t.dataset.rtoggle]; runModel(); }
  });
  document.addEventListener('change', (e) => { if (e.target.id === 'rChart') { ratioUI.chart = e.target.value; runModel(); } });
}
