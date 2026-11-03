/* Finance Studio — Indicators tab: live interest rates, economic indicators and industry benchmarks.
   Live data comes from /api/market (Netlify function reading FRED). Depends on app.js and benchmarks.js. */

const IND = {
  DFF: { label: 'Fed funds rate', kind: 'rate', group: 'rates', why: 'The Federal Reserve’s policy rate. Floating-rate loans reset off it, so it drives the interest rate on debt in your forecasts.' },
  DPRIME: { label: 'Bank prime rate', kind: 'rate', group: 'rates', why: 'What banks charge their best borrowers. Small-business loans are usually priced at prime plus a margin.' },
  DGS3MO: { label: '3-month Treasury', kind: 'rate', group: 'rates', why: 'Short-term risk-free rate. Use for cash yield and short projects.' },
  DGS2: { label: '2-year Treasury', kind: 'rate', group: 'rates', why: 'Reflects where markets expect policy rates to go over the next two years.' },
  DGS5: { label: '5-year Treasury', kind: 'rate', group: 'rates', why: 'Benchmark for 5-year term loans and medium-length projects.' },
  DGS10: { label: '10-year Treasury', kind: 'rate', group: 'rates', why: 'The standard risk-free rate in WACC and DCF valuations.' },
  DGS30: { label: '30-year Treasury', kind: 'rate', group: 'rates', why: 'Long-term risk-free rate for very long-lived assets such as utilities and real estate.' },
  MORTGAGE30US: { label: '30-year mortgage', kind: 'rate', group: 'rates', why: 'Drives housing demand and real-estate valuations.' },
  BAMLC0A0CM: { label: 'Investment-grade credit spread', kind: 'rate', group: 'rates', why: 'Extra yield strong companies pay over Treasuries. Risk-free rate + this spread ≈ a healthy company’s pre-tax cost of debt.' },
  BAMLH0A0HYM2: { label: 'High-yield credit spread', kind: 'rate', group: 'rates', why: 'Extra yield riskier borrowers pay. Use for leveraged companies and LBO debt.' },
  T10Y2Y: { label: 'Yield curve (10y – 2y)', kind: 'rate', group: 'rates', why: 'A negative (inverted) curve has historically come before recessions. Consider a cautious downside scenario.' },
  T10YIE: { label: '10-year inflation expectation', kind: 'rate', group: 'economy', why: 'Market-implied inflation. Real growth plus this is a sensible ceiling for long-term (terminal) growth.' },
  CPIAUCSL: { label: 'Inflation (CPI, year on year)', kind: 'yoy', group: 'economy', why: 'Sense-check price-driven revenue growth and cost inflation in your assumptions.' },
  CPILFESL: { label: 'Core inflation (ex food & energy)', kind: 'yoy', group: 'economy', why: 'Underlying inflation trend, which is what central banks watch most closely.' },
  UNRATE: { label: 'Unemployment rate', kind: 'rate', group: 'economy', why: 'Labour-market strength. Rising unemployment usually means slower consumer demand and wage pressure easing.' },
  A191RL1Q225SBEA: { label: 'Real GDP growth (annualised)', kind: 'rate', group: 'economy', why: 'Overall economic growth. Long-run revenue growth for a mature business rarely beats nominal GDP for long.' },
  PAYEMS: { label: 'Jobs added (latest month)', kind: 'diff', group: 'economy', why: 'Monthly change in US payrolls, in thousands. A quick read on economic momentum.' },
  UMCSENT: { label: 'Consumer sentiment', kind: 'level', group: 'economy', why: 'How confident households feel. Leads spending in retail, travel and restaurants.' },
  VIXCLS: { label: 'Volatility index (VIX)', kind: 'level', group: 'markets', why: 'Market fear gauge. When it is high, investors demand a higher equity risk premium, which lowers valuations.' },
  SP500: { label: 'S&P 500 index', kind: 'level', group: 'markets', why: 'Broad US equity market level, the backdrop for trading multiples in comparables.' },
  DCOILWTICO: { label: 'Oil price (WTI, $/barrel)', kind: 'level', group: 'markets', why: 'Key cost driver for transport, airlines and chemicals, and a revenue driver for energy companies.' },
  DTWEXBGS: { label: 'US dollar index', kind: 'level', group: 'markets', why: 'A stronger dollar reduces the US-dollar value of overseas sales.' },
};

let market = null;          // data from /api/market, or { error } when unavailable
let marketLoading = false;
let benchQuery = '', benchGroup = 'All';

function loadPref(key) { try { return JSON.parse(localStorage.getItem(key)); } catch (e) { return null; } }
function savePref(key, v) { try { v == null ? localStorage.removeItem(key) : localStorage.setItem(key, JSON.stringify(v)); } catch (e) {} }
// Risk-free rate chosen on the Indicators tab ({ value: 4.15, date, source }), used as the WACC default
const marketRf = () => loadPref('fs-rf');
// Industry benchmark chosen on the Indicators tab, used for comparisons, comps and WACC defaults
const chosenBenchmark = () => BENCHMARKS.find((b) => b.id === loadPref('fs-bench')) || null;

async function loadMarket(force) {
  if ((market && !force) || marketLoading) return;
  marketLoading = true;
  try {
    const r = await fetch('/api/market', { headers: { Accept: 'application/json' } });
    const j = r.ok && /json/.test(r.headers.get('content-type') || '') ? await r.json() : null;
    market = j && j.series ? j : { error: true };
  } catch (e) {
    market = { error: true };
  }
  marketLoading = false;
  if ($('#view-indicators').classList.contains('active')) renderIndicators();
}

function indValue(id) {
  const s = market?.series?.[id];
  if (!s) return null;
  const d = IND[id];
  if (d.kind === 'yoy') {
    const h = s.history, last = h.at(-1), target = new Date(last[0]); target.setFullYear(target.getFullYear() - 1);
    const ago = [...h].reverse().find(([dt]) => new Date(dt) <= target);
    const prevIdx = h.length - 2, prevTarget = prevIdx >= 0 ? new Date(h[prevIdx][0]) : null;
    if (prevTarget) prevTarget.setFullYear(prevTarget.getFullYear() - 1);
    const prevAgo = prevTarget ? [...h].reverse().find(([dt]) => new Date(dt) <= prevTarget) : null;
    return { value: ago ? (last[1] / ago[1] - 1) * 100 : null, prev: prevAgo ? (h[prevIdx][1] / prevAgo[1] - 1) * 100 : null, date: s.date, history: s.history };
  }
  if (d.kind === 'diff') return { value: s.prev != null ? s.value - s.prev : null, prev: null, date: s.date, history: s.history };
  return { value: s.value, prev: s.prev, yearAgo: s.yearAgo, date: s.date, history: s.history };
}

function spark(history) {
  if (!history || history.length < 2) return '';
  const v = history.map((p) => p[1]), lo = Math.min(...v), hi = Math.max(...v), w = 120, h = 32;
  const pts = v.map((x, i) => `${(i / (v.length - 1)) * w},${h - 2 - ((x - lo) / (hi - lo || 1)) * (h - 4)}`).join(' ');
  return `<svg class="spark" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" aria-hidden="true"><polyline points="${pts}" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>`;
}

function indCard(id) {
  const d = IND[id], x = indValue(id);
  const fmtV = (v) => (v == null ? '—' : d.kind === 'rate' || d.kind === 'yoy' ? `${v.toFixed(2)}%` : d.kind === 'diff' ? `${v >= 0 ? '+' : ''}${Math.round(v).toLocaleString()}k` : v.toLocaleString(undefined, { maximumFractionDigits: 2 }));
  let chg = '';
  if (x && x.value != null && x.prev != null && (d.kind === 'rate' || d.kind === 'yoy')) {
    const bp = Math.round((x.value - x.prev) * 100);
    chg = `<span class="${bp > 0 ? 'up-n' : bp < 0 ? 'down-n' : ''}">${bp > 0 ? '▲' : bp < 0 ? '▼' : '■'} ${Math.abs(bp)} bp</span>`;
  } else if (x && x.value != null && x.yearAgo) {
    const p = x.value / x.yearAgo - 1;
    chg = `<span class="${p > 0 ? 'up-n' : 'down-n'}">${p >= 0 ? '▲' : '▼'} ${Math.abs(p * 100).toFixed(1)}% in 1 yr</span>`;
  }
  return `<div class="ind-card">
    <div class="ind-top"><span class="ind-label">${d.label}</span>${x ? `<span class="ind-date">${x.date}</span>` : ''}</div>
    <div class="ind-value">${fmtV(x?.value)}</div>
    <div class="ind-chg">${chg || '&nbsp;'}</div>
    ${x ? spark(x.history) : ''}
    <p class="ind-why">${d.why}</p>
  </div>`;
}

function renderIndicators() {
  const live = market && !market.error, rf = marketRf();
  $('#indStatus').innerHTML = !market
    ? '<span class="dot-wait"></span> Loading live data…'
    : live
      ? `<span class="dot-live"></span> Live data · ${esc(market.source)} · fetched ${new Date(market.asOf).toLocaleString()}${market.failed ? ` · ${market.failed} series unavailable` : ''} <button class="link" data-act="refresh-market">Refresh</button>`
      : `<span class="dot-off"></span> Live data isn’t available on this copy of the site. It needs the site’s Netlify Functions (deploy from GitHub). You can still enter a risk-free rate yourself below, and the industry benchmarks work offline.`;

  const ids = (g) => Object.keys(IND).filter((k) => IND[k].group === g);
  $('#indRates').innerHTML = ids('rates').map(indCard).join('');
  $('#indEconomy').innerHTML = ids('economy').map(indCard).join('');
  $('#indMarkets').innerHTML = ids('markets').map(indCard).join('');

  const ten = live ? market.series.DGS10 : null, ig = live ? market.series.BAMLC0A0CM : null;
  $('#indUse').innerHTML = `
    <div class="use-row">
      <div>
        <strong>Risk-free rate used in your models</strong>
        <div class="muted small">${rf ? `${rf.value.toFixed(2)}% · ${esc(rf.source)}${rf.date ? ` (${rf.date})` : ''}` : 'Default 4.20% (not set from market data)'}</div>
      </div>
      <div class="row tight-row">
        ${ten ? `<button class="btn" data-act="use-rf">Use live 10-year yield (${ten.value.toFixed(2)}%)</button>` : ''}
        <label class="rf-manual">or enter <input id="rfManual" type="number" step="0.01" placeholder="4.20" value="${rf ? rf.value : ''}">%</label>
        <button class="btn ghost" data-act="set-rf">Save</button>
        ${rf ? '<button class="link" data-act="clear-rf">Reset</button>' : ''}
      </div>
    </div>
    ${ten && ig ? `<p class="muted small">Indicative pre-tax cost of debt for an investment-grade company: 10-year yield ${ten.value.toFixed(2)}% + credit spread ${ig.value.toFixed(2)}% = <b>${(ten.value + ig.value).toFixed(2)}%</b>.</p>` : ''}`;

  if (live && window.Chart) {
    const mats = [['3M', 'DGS3MO'], ['2Y', 'DGS2'], ['5Y', 'DGS5'], ['10Y', 'DGS10'], ['30Y', 'DGS30']].filter(([, id]) => market.series[id]);
    makeChart('chYield', {
      type: 'line',
      data: { labels: mats.map((m) => m[0]), datasets: [
        ds(`Today (${market.series.DGS10?.date || ''})`, mats.map(([, id]) => market.series[id].value), COLORS.blue),
        ds('One year ago', mats.map(([, id]) => market.series[id].yearAgo), COLORS.orange, { borderDash: [6, 4] }),
      ] },
      options: { scales: { y: { ticks: { callback: (v) => `${(+v).toFixed(2)}%` } } } },
    });
  }
  $('#yieldWrap').hidden = !live;
  renderBenchmarks();
}

/* ---------- Industry benchmarks ---------- */
function companyMetrics() {
  if (!hasData()) return null;
  const i = state.years.length - 1, g = (k) => S(k)?.[i] ?? null, rev = g('revenue');
  const div = (a, b) => (a == null || b == null || !b ? null : a / b);
  const prev = i > 0 ? S('revenue')?.[i - 1] : null;
  return {
    growth: cagr(S('revenue')) ?? (prev ? rev / prev - 1 : null), gm: div(g('grossProfit'), rev), em: div(g('ebitda'), rev), nm: div(g('netIncome'), rev),
    roe: div(g('netIncome'), g('equity')), de: div(g('debt'), g('equity')),
  };
}

const BENCH_COLS = [['growth', 'Revenue growth', 'pct'], ['gm', 'Gross margin', 'pct'], ['em', 'EBITDA margin', 'pct'], ['nm', 'Net margin', 'pct'], ['roe', 'Return on equity', 'pct'],
  ['de', 'Debt / equity', 'x'], ['evEbitda', 'EV / EBITDA', 'x'], ['evSales', 'EV / Sales', 'x'], ['pe', 'P / E', 'x'], ['beta', 'Beta', 'n'], ['wacc', 'WACC', 'pct']];
const bf = (v, t) => (v == null ? '—' : t === 'pct' ? `${(v * 100).toFixed(1)}%` : t === 'x' ? `${v.toFixed(1)}x` : v.toFixed(2));

function renderBenchmarks() {
  const groups = ['All', ...new Set(BENCHMARKS.map((b) => b.group))], pick = chosenBenchmark();
  $('#benchGroups').innerHTML = groups.map((g) => `<button data-bgroup="${g}" class="${g === benchGroup ? 'on' : ''}">${g}</button>`).join('');
  const rows = BENCHMARKS.filter((b) => (benchGroup === 'All' || b.group === benchGroup) && (!benchQuery || `${b.name} ${b.group}`.toLowerCase().includes(benchQuery.toLowerCase())));
  $('#benchTable').innerHTML = `<div class="table-wrap"><table class="bench"><thead><tr><th>Industry</th>${BENCH_COLS.map((c) => `<th>${c[1]}</th>`).join('')}<th></th></tr></thead><tbody>
    ${rows.map((b) => `<tr class="${pick?.id === b.id ? 'picked' : ''}"><td><strong>${b.name}</strong><div class="muted small">${b.group}</div></td>${BENCH_COLS.map(([k, , t]) => `<td>${bf(b[k], t)}</td>`).join('')}
      <td><button class="btn ${pick?.id === b.id ? '' : 'ghost'} btn-sm" data-bench="${b.id}">${pick?.id === b.id ? 'Selected' : 'Use'}</button></td></tr>`).join('')}
  </tbody></table></div>`;

  const cm = companyMetrics();
  $('#benchCompare').innerHTML = !pick
    ? '<p class="muted">Choose an industry with <b>Use</b> to compare your company against it. The choice also sets the peer multiples in Comparables and the beta in WACC.</p>'
    : `<div class="row spread"><h3>${hasData() ? esc(state.company) : 'Your company'} vs ${pick.name}</h3><button class="link" data-act="clear-bench">Clear industry</button></div>
      ${!cm ? '<p class="muted">Upload statements (or load a sample company) to compare.</p>' : table(['Metric', 'Your company', 'Industry', 'Difference', 'Assessment'],
        BENCH_COLS.filter(([k]) => k in cm).map(([k, label, t]) => {
          const you = cm[k], ind = pick[k], diff = you == null ? null : you - ind;
          const better = k === 'de' ? diff < 0 : diff > 0, near = you != null && Math.abs(diff) <= Math.abs(ind) * 0.1;
          const tag = you == null ? '—' : near ? '<span class="pill-s mid">In line</span>' : better ? '<span class="pill-s good">Stronger</span>' : '<span class="pill-s bad">Weaker</span>';
          return [label, bf(you, t), bf(ind, t), diff == null ? '—' : t === 'pct' ? `${diff >= 0 ? '+' : ''}${(diff * 100).toFixed(1)} pts` : `${diff >= 0 ? '+' : ''}${diff.toFixed(2)}x`, tag];
        }))}`;
}

function initIndicators() {
  $('#view-indicators').addEventListener('click', (e) => {
    const b = e.target.closest('[data-act],[data-bench],[data-bgroup]');
    if (!b) return;
    if (b.dataset.bench) { savePref('fs-bench', b.dataset.bench); renderBenchmarks(); return; }
    if (b.dataset.bgroup) { benchGroup = b.dataset.bgroup; renderBenchmarks(); return; }
    const act = b.dataset.act;
    if (act === 'refresh-market') { market = null; renderIndicators(); loadMarket(true); }
    if (act === 'use-rf') { const s = market.series.DGS10; savePref('fs-rf', { value: s.value, date: s.date, source: 'Live 10-year Treasury yield (FRED)' }); renderIndicators(); }
    if (act === 'set-rf') { const v = parseFloat($('#rfManual').value); if (isFinite(v)) { savePref('fs-rf', { value: v, date: new Date().toISOString().slice(0, 10), source: 'Entered manually' }); renderIndicators(); } }
    if (act === 'clear-rf') { savePref('fs-rf', null); renderIndicators(); }
    if (act === 'clear-bench') { savePref('fs-bench', null); renderBenchmarks(); }
  });
  $('#benchSearch').addEventListener('input', (e) => { benchQuery = e.target.value.trim(); renderBenchmarks(); });
}
