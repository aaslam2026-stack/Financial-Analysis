/* Glass Finance Studio — plain JavaScript, no build step.
   Libraries: Chart.js (charts), SheetJS (Excel parsing). Weather: Open-Meteo (free, no key). */

const $ = (s) => document.querySelector(s);
const $$ = (s) => [...document.querySelectorAll(s)];
const STORE = 'gfs-data-v1';
const COLORS = { blue: '#0a84ff', green: '#30d158', orange: '#ff9f0a', purple: '#bf5af2', red: '#ff453a', teal: '#64d2ff' };

let state = { company: '', years: [], items: {} };
const charts = {};

/* ---------- Line-item recognition ---------- */
const ALIASES = {
  revenue: ['revenue', 'revenues', 'total revenue', 'sales', 'net sales', 'turnover', 'total sales'],
  cogs: ['cost of goods sold', 'cogs', 'cost of sales', 'cost of revenue'],
  grossProfit: ['gross profit', 'gross margin'],
  opex: ['operating expenses', 'opex', 'sg&a', 'selling general and administrative', 'total operating expenses'],
  ebitda: ['ebitda'],
  da: ['depreciation & amortization', 'depreciation and amortization', 'd&a', 'depreciation', 'amortization'],
  ebit: ['ebit', 'operating income', 'operating profit'],
  interest: ['interest expense', 'interest', 'finance costs'],
  netIncome: ['net income', 'net profit', 'profit after tax', 'net earnings'],
  cash: ['cash', 'cash and equivalents', 'cash & equivalents', 'cash and cash equivalents'],
  currentAssets: ['current assets', 'total current assets'],
  totalAssets: ['total assets'],
  currentLiabilities: ['current liabilities', 'total current liabilities'],
  totalLiabilities: ['total liabilities'],
  debt: ['total debt', 'debt', 'long-term debt', 'long term debt', 'borrowings'],
  equity: ['equity', 'total equity', 'shareholders equity', "shareholders' equity", 'stockholders equity'],
  capex: ['capex', 'capital expenditure', 'capital expenditures'],
  ocf: ['operating cash flow', 'cash from operations', 'net cash from operating activities'],
};
const LABELS = {
  revenue: 'Revenue', cogs: 'Cost of Goods Sold', grossProfit: 'Gross Profit', opex: 'Operating Expenses',
  ebitda: 'EBITDA', da: 'D&A', ebit: 'EBIT', interest: 'Interest Expense', netIncome: 'Net Income',
  cash: 'Cash', currentAssets: 'Current Assets', totalAssets: 'Total Assets', currentLiabilities: 'Current Liabilities',
  totalLiabilities: 'Total Liabilities', debt: 'Total Debt', equity: 'Equity', capex: 'CapEx', ocf: 'Operating Cash Flow',
};
const ALIAS_LIST = Object.entries(ALIASES)
  .flatMap(([key, arr]) => arr.map((a) => [a, key]))
  .sort((a, b) => b[0].length - a[0].length);

function matchKey(label) {
  const l = String(label || '').toLowerCase().replace(/[’']/g, "'").replace(/\s+/g, ' ').trim();
  if (!l) return null;
  const exact = ALIAS_LIST.find(([a]) => a === l);
  if (exact) return exact[1];
  const partial = ALIAS_LIST.find(([a]) => new RegExp(`(^|[^a-z])${a.replace(/[.*+?^${}()|[\]\\&]/g, '\\$&')}($|[^a-z])`).test(l));
  return partial ? partial[1] : null;
}

function parseNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v == null) return null;
  let s = String(v).trim();
  if (!s || s === '-' || s === '—') return null;
  const neg = /^\(.*\)$/.test(s) || s.startsWith('-');
  s = s.replace(/[^0-9.]/g, '');
  if (!s) return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : neg ? -n : n;
}

function parseRows(rows) {
  rows = rows.filter((r) => r && r.some((c) => String(c ?? '').trim() !== ''));
  let hIdx = rows.findIndex((r) => r.slice(1).some((c) => /(19|20)\d{2}/.test(String(c ?? ''))));
  let years;
  if (hIdx === -1) {
    const width = Math.max(...rows.map((r) => r.length)) - 1;
    years = Array.from({ length: width }, (_, i) => `P${i + 1}`);
    hIdx = -1;
  } else {
    years = rows[hIdx].slice(1).map((c) => String(c ?? '').trim()).filter(Boolean);
  }
  const items = {};
  rows.slice(hIdx + 1).forEach((r) => {
    const key = matchKey(r[0]);
    if (!key || items[key]) return;
    const vals = years.map((_, i) => parseNum(r[i + 1]));
    if (vals.some((v) => v != null)) items[key] = vals;
  });
  // Put years in ascending order if they were given newest-first
  const yNum = years.map((y) => parseInt(String(y).match(/(19|20)\d{2}/)?.[0]));
  if (yNum.length > 1 && yNum[0] > yNum[yNum.length - 1]) {
    years.reverse();
    Object.values(items).forEach((a) => a.reverse());
  }
  return { years, items };
}

function parseCSV(text) {
  const rows = [];
  let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"' && text[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') q = false;
      else cur += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cur); cur = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cur); rows.push(row); row = []; cur = '';
    } else cur += c;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

const op = (a, b, fn) => a && b ? a.map((x, i) => (x == null || b[i] == null ? null : fn(x, b[i]))) : null;

function derive(items) {
  const it = { ...items };
  if (it.capex) it.capex = it.capex.map((v) => (v == null ? null : Math.abs(v)));
  if (!it.grossProfit && it.revenue && it.cogs) it.grossProfit = op(it.revenue, it.cogs, (a, b) => a - Math.abs(b));
  if (!it.ebitda) {
    if (it.ebit && it.da) it.ebitda = op(it.ebit, it.da, (a, b) => a + Math.abs(b));
    else if (it.grossProfit && it.opex) it.ebitda = op(it.grossProfit, it.opex, (a, b) => a - Math.abs(b));
  }
  if (!it.ebit && it.ebitda) it.ebit = it.da ? op(it.ebitda, it.da, (a, b) => a - Math.abs(b)) : it.ebitda;
  if (!it.equity && it.totalAssets && it.totalLiabilities) it.equity = op(it.totalAssets, it.totalLiabilities, (a, b) => a - b);
  return it;
}

/* ---------- Helpers ---------- */
const S = (k) => state.items[k] || null;
const last = (arr) => (arr ? [...arr].reverse().find((v) => v != null) ?? null : null);
const L = (k) => last(S(k));
const hasData = () => !!S('revenue');
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function fmt(n, money = true) {
  if (n == null || !isFinite(n)) return '—';
  const a = Math.abs(n), sign = n < 0 ? '-' : '', p = money ? '$' : '';
  if (a >= 1e9) return `${sign}${p}${(a / 1e9).toFixed(2)}B`;
  if (a >= 1e6) return `${sign}${p}${(a / 1e6).toFixed(2)}M`;
  if (a >= 1e3) return `${sign}${p}${(a / 1e3).toFixed(1)}K`;
  return `${sign}${p}${a.toFixed(a < 10 && a % 1 ? 2 : 0)}`;
}
const pct = (n, d = 1) => (n == null || !isFinite(n) ? '—' : `${(n * 100).toFixed(d)}%`);
const xf = (n) => (n == null || !isFinite(n) ? '—' : `${n.toFixed(2)}x`);
const yrs = (n) => (n == null || !isFinite(n) ? 'Never' : `${n.toFixed(1)} yrs`);

const kpi = (label, value, delta = '', cls = '') =>
  `<div class="glass kpi"><div class="label">${label}</div><div class="value">${value}</div>${delta ? `<div class="delta ${cls}">${delta}</div>` : ''}</div>`;
const table = (head, rows, hl) =>
  `<div class="table-wrap"><table><thead><tr>${head.map((h) => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows
    .map((r, ri) => `<tr>${r.map((c, ci) => `<td class="${hl && hl(ri, ci) ? 'hl' : ''}">${c}</td>`).join('')}</tr>`)
    .join('')}</tbody></table></div>`;
const list = (s) => String(s).split(/[,;\s]+/).map(parseNum).filter((v) => v != null);
const median = (a) => { const s = [...a].sort((x, y) => x - y), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
const cagr = (arr) => {
  const v = (arr || []).filter((x) => x != null && x > 0);
  return v.length > 1 ? Math.pow(v[v.length - 1] / v[0], 1 / (v.length - 1)) - 1 : null;
};

/* ---------- Finance math ---------- */
const npv = (r, flows) => flows.reduce((s, cf, t) => s + cf / Math.pow(1 + r, t), 0);
function irr(flows) {
  let lo = -0.99, hi = 10, flo = npv(lo, flows), fhi = npv(hi, flows);
  if (flo * fhi > 0) return null;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2, f = npv(mid, flows);
    if (Math.abs(f) < 1e-7) return mid;
    if (f * flo < 0) hi = mid; else { lo = mid; flo = f; }
  }
  return (lo + hi) / 2;
}
function payback(flows, r = 0) {
  let cum = 0;
  for (let t = 0; t < flows.length; t++) {
    const cf = flows[t] / Math.pow(1 + r, t), prev = cum;
    cum += cf;
    if (t > 0 && prev < 0 && cum >= 0) return t - 1 + -prev / cf;
  }
  return null;
}
function mirr(flows, fin, reinv) {
  const n = flows.length - 1;
  const fv = flows.reduce((s, cf, t) => s + (cf > 0 ? cf * Math.pow(1 + reinv, n - t) : 0), 0);
  const pv = flows.reduce((s, cf, t) => s + (cf < 0 ? cf / Math.pow(1 + fin, t) : 0), 0);
  return pv < 0 && fv > 0 ? Math.pow(fv / -pv, 1 / n) - 1 : null;
}
function randn() {
  let u = 0, v = 0;
  while (!u) u = Math.random();
  while (!v) v = Math.random();
  return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
}

/* ---------- Charts ---------- */
function themeCharts() {
  if (!window.Chart) return;
  const cs = getComputedStyle(document.documentElement);
  Chart.defaults.color = cs.getPropertyValue('--muted').trim();
  Chart.defaults.borderColor = cs.getPropertyValue('--line').trim();
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.maintainAspectRatio = false;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
}
function makeChart(id, config) {
  const el = document.getElementById(id);
  if (!el || !window.Chart) return;
  charts[id]?.destroy();
  charts[id] = new Chart(el, config);
}
const moneyAxis = { ticks: { callback: (v) => fmt(v) } };
const ds = (label, data, color, extra = {}) => ({ label, data, borderColor: color, backgroundColor: color + 'cc', borderRadius: 6, tension: 0.35, ...extra });

/* ---------- Storage ---------- */
function save() { try { localStorage.setItem(STORE, JSON.stringify(state)); } catch (e) {} }
function load() {
  try { const s = JSON.parse(localStorage.getItem(STORE)); if (s && s.items) state = s; } catch (e) {}
}

/* ---------- Upload ---------- */
const SAMPLE_CSV = `Line Item,2020,2021,2022,2023,2024
Revenue,8200000,9100000,10400000,11800000,13100000
Cost of Goods Sold,4900000,5400000,6100000,6800000,7500000
Operating Expenses,1800000,1950000,2150000,2400000,2600000
Depreciation & Amortization,300000,330000,360000,400000,440000
Interest Expense,180000,170000,160000,150000,140000
Net Income,765000,937500,1222500,1537500,1815000
Cash,900000,1100000,1400000,1800000,2300000
Current Assets,3100000,3500000,4000000,4600000,5300000
Total Assets,9500000,10300000,11400000,12600000,14000000
Current Liabilities,1900000,2000000,2200000,2400000,2600000
Total Debt,3000000,2800000,2600000,2400000,2200000
Total Liabilities,5400000,5400000,5500000,5600000,5800000
Equity,4100000,4900000,5900000,7000000,8200000
CapEx,450000,500000,600000,700000,800000
Operating Cash Flow,1200000,1400000,1700000,2000000,2400000`;

function applyRows(rows, company) {
  const parsed = parseRows(rows);
  if (!Object.keys(parsed.items).length) {
    msg('Could not find any known line items. Check the first column has labels like "Revenue" and the first row has years.', true);
    return;
  }
  state = { company: company || $('#companyName').value.trim() || 'My Company', years: parsed.years, items: derive(parsed.items) };
  save();
  const found = Object.keys(parsed.items).length;
  msg(`Loaded ${found} line items across ${parsed.years.length} periods.${S('revenue') ? '' : ' Note: no Revenue row found, so most models will use default values.'}`);
  refreshAll();
}
function msg(t, err) { const el = $('#uploadMsg'); el.textContent = t; el.style.color = err ? 'var(--red)' : 'var(--green)'; }

function handleFile(file) {
  if (!file) return;
  const name = file.name.replace(/\.[^.]+$/, '');
  const reader = new FileReader();
  if (/\.csv$/i.test(file.name)) {
    reader.onload = () => applyRows(parseCSV(reader.result), $('#companyName').value.trim() || name);
    reader.readAsText(file);
  } else {
    if (!window.XLSX) return msg('Excel reader failed to load. Please check your connection or upload a CSV.', true);
    reader.onload = () => {
      const wb = XLSX.read(new Uint8Array(reader.result), { type: 'array' });
      // Combine all sheets so separate P&L / balance sheet tabs both get read
      const rows = wb.SheetNames.flatMap((n) => XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }));
      applyRows(rows, $('#companyName').value.trim() || name);
    };
    reader.readAsArrayBuffer(file);
  }
}

function renderPreview() {
  const card = $('#previewCard');
  card.hidden = !Object.keys(state.items).length;
  if (card.hidden) return;
  $('#preview').innerHTML = table(['Line item', ...state.years.map(esc)],
    Object.keys(LABELS).filter((k) => S(k)).map((k) => [LABELS[k], ...S(k).map((v) => fmt(v))]));
}

/* ---------- Dashboard ---------- */
function renderDashboard() {
  const on = hasData();
  $('#dashEmpty').hidden = on;
  $('#dashContent').hidden = !on;
  $('#companyPill').textContent = on ? state.company : 'No data loaded';
  if (!on) return;
  const rev = S('revenue'), n = rev.length;
  const growth = n > 1 && rev[n - 2] ? rev[n - 1] / rev[n - 2] - 1 : null;
  const em = L('ebitda') / L('revenue'), nm = L('netIncome') / L('revenue');
  const d = (g) => (g == null ? '' : `${g >= 0 ? '▲' : '▼'} ${pct(Math.abs(g))} vs prior year`);
  $('#kpis').innerHTML = [
    kpi('Revenue', fmt(L('revenue')), d(growth), growth >= 0 ? 'up' : 'down'),
    kpi('EBITDA', fmt(L('ebitda')), `${pct(em)} margin`),
    kpi('Net Income', fmt(L('netIncome')), `${pct(nm)} margin`),
    kpi('Cash', fmt(L('cash'))),
    kpi('Revenue CAGR', pct(cagr(rev)), `${state.years[0]}–${state.years[n - 1]}`),
  ].join('');

  makeChart('chTrend', {
    type: 'bar',
    data: { labels: state.years, datasets: [ds('Revenue', rev, COLORS.blue), ds('EBITDA', S('ebitda') || [], COLORS.green), ds('Net Income', S('netIncome') || [], COLORS.purple)] },
    options: { scales: { y: moneyAxis } },
  });
  const m = (k) => (S(k) ? op(S(k), rev, (a, b) => (b ? (a / b) * 100 : null)) : []);
  makeChart('chMargin', {
    type: 'line',
    data: { labels: state.years, datasets: [ds('Gross', m('grossProfit'), COLORS.blue), ds('EBITDA', m('ebitda'), COLORS.green), ds('Net', m('netIncome'), COLORS.purple)] },
    options: { scales: { y: { ticks: { callback: (v) => v + '%' } } } },
  });
  const bs = [['Equity', L('equity'), COLORS.green], ['Debt', L('debt'), COLORS.orange], ['Other liabilities', L('totalLiabilities') != null && L('debt') != null ? L('totalLiabilities') - L('debt') : L('totalLiabilities'), COLORS.purple]].filter((x) => x[1] > 0);
  makeChart('chBalance', {
    type: 'doughnut',
    data: { labels: bs.map((x) => x[0]), datasets: [{ data: bs.map((x) => x[1]), backgroundColor: bs.map((x) => x[2]), borderWidth: 0 }] },
    options: { cutout: '65%', plugins: { legend: { position: 'right' } } },
  });
  const r = ratios().at(-1);
  $('#dashRatios').innerHTML = table(['Ratio', 'Value'], r.map(([k, v]) => [k, v]));
}

function ratios() {
  return state.years.map((_, i) => {
    const g = (k) => (S(k) ? S(k)[i] : null);
    const div = (a, b) => (a == null || b == null || b === 0 ? null : a / b);
    const prevRev = i > 0 && S('revenue') ? S('revenue')[i - 1] : null;
    return [
      ['Revenue growth', pct(prevRev ? g('revenue') / prevRev - 1 : null)],
      ['Gross margin', pct(div(g('grossProfit'), g('revenue')))],
      ['EBITDA margin', pct(div(g('ebitda'), g('revenue')))],
      ['Net margin', pct(div(g('netIncome'), g('revenue')))],
      ['Current ratio', xf(div(g('currentAssets'), g('currentLiabilities')))],
      ['Debt / Equity', xf(div(g('debt'), g('equity')))],
      ['Debt / EBITDA', xf(div(g('debt'), g('ebitda')))],
      ['Interest coverage', xf(div(g('ebit'), g('interest')))],
      ['Return on assets', pct(div(g('netIncome'), g('totalAssets')))],
      ['Return on equity', pct(div(g('netIncome'), g('equity')))],
      ['Cash conversion (OCF / NI)', xf(div(g('ocf'), g('netIncome')))],
    ];
  });
}

/* ---------- Weather (Open-Meteo) ---------- */
const WMO = { 0: 'Clear', 1: 'Mainly clear', 2: 'Partly cloudy', 3: 'Overcast', 45: 'Fog', 48: 'Fog', 51: 'Drizzle', 53: 'Drizzle', 55: 'Drizzle', 61: 'Rain', 63: 'Rain', 65: 'Heavy rain', 71: 'Snow', 73: 'Snow', 75: 'Heavy snow', 80: 'Showers', 81: 'Showers', 82: 'Heavy showers', 95: 'Thunderstorm', 96: 'Thunderstorm', 99: 'Thunderstorm' };

async function geocode(city) {
  const r = await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(city)}&count=1&language=en&format=json`);
  const j = await r.json();
  if (!j.results?.length) throw new Error(`Location "${city}" not found`);
  const p = j.results[0];
  return { lat: p.latitude, lon: p.longitude, name: [p.name, p.admin1, p.country].filter(Boolean).join(', ') };
}

async function loadWeather(city) {
  const out = $('#weatherOut');
  out.innerHTML = '<p class="muted">Loading…</p>';
  try {
    const g = await geocode(city);
    const r = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${g.lat}&longitude=${g.lon}&current=temperature_2m,weather_code,wind_speed_10m&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,wind_speed_10m_max&timezone=auto&forecast_days=7`);
    const w = await r.json();
    const dl = w.daily;
    const rainDays = dl.precipitation_sum.filter((p) => p >= 10).length;
    const hot = Math.max(...dl.temperature_2m_max), windy = Math.max(...dl.wind_speed_10m_max);
    const alerts = [rainDays && `${rainDays} heavy-rain day(s)`, hot >= 35 && `heat up to ${hot}°C`, windy >= 50 && `wind up to ${windy} km/h`].filter(Boolean);
    out.innerHTML = `<div class="kpis">${[
      kpi(esc(g.name), `${w.current.temperature_2m}°C`, WMO[w.current.weather_code] || ''),
      kpi('7-day rain total', `${dl.precipitation_sum.reduce((a, b) => a + b, 0).toFixed(1)} mm`),
      kpi('Max wind', `${windy} km/h`),
      kpi('Business risk', alerts.length ? 'Watch' : 'Low', alerts.join(', ') || 'No severe weather expected', alerts.length ? 'down' : 'up'),
    ].join('')}</div>`;
    makeChart('chWeather', {
      data: {
        labels: dl.time.map((t) => new Date(t + 'T00:00').toLocaleDateString(undefined, { weekday: 'short' })),
        datasets: [
          ds('Max °C', dl.temperature_2m_max, COLORS.orange, { type: 'line', yAxisID: 'y' }),
          ds('Min °C', dl.temperature_2m_min, COLORS.teal, { type: 'line', yAxisID: 'y' }),
          ds('Rain mm', dl.precipitation_sum, COLORS.blue, { type: 'bar', yAxisID: 'y1' }),
        ],
      },
      options: { scales: { y: { position: 'left' }, y1: { position: 'right', grid: { display: false } } } },
    });
  } catch (e) {
    out.innerHTML = `<p style="color:var(--red)">Weather unavailable: ${esc(e.message)}</p>`;
  }
}

/* ---------- Models ---------- */
function defaults() {
  const rev = L('revenue') || 1_000_000;
  const ebitda = L('ebitda') || rev * 0.2;
  return {
    rev, ebitda,
    g: +((cagr(S('revenue')) ?? 0.05) * 100).toFixed(1),
    m: +((ebitda / rev) * 100).toFixed(1),
    da: +(((L('da') ?? rev * 0.03) / rev) * 100).toFixed(1),
    capex: +(((L('capex') ?? rev * 0.04) / rev) * 100).toFixed(1),
    netDebt: (L('debt') || 0) - (L('cash') || 0),
    ni: L('netIncome') || ebitda * 0.5,
    equity: L('equity') || rev,
    debt: L('debt') || rev * 0.3,
    interest: L('interest') && L('debt') ? +((L('interest') / L('debt')) * 100).toFixed(1) : 6,
  };
}

function dcfCore(v) {
  const g = v.g / 100, m = v.m / 100, w = v.wacc / 100, tg = v.tg / 100, t = v.tax / 100;
  let rev = v.rev, pv = 0, fcf = 0;
  const rows = [];
  for (let y = 1; y <= v.years; y++) {
    const prev = rev;
    rev *= 1 + g;
    const ebitda = rev * m, da = rev * (v.da / 100), ebit = ebitda - da;
    const capex = rev * (v.capex / 100), dnwc = (rev - prev) * (v.nwc / 100);
    fcf = ebit * (1 - t) + da - capex - dnwc;
    const p = fcf / Math.pow(1 + w, y);
    pv += p;
    rows.push({ y, rev, ebitda, ebit, capex, fcf, pv: p });
  }
  if (w <= tg) return { error: 'WACC must be higher than terminal growth.' };
  const tv = (fcf * (1 + tg)) / (w - tg), pvTv = tv / Math.pow(1 + w, v.years), ev = pv + pvTv;
  return { rows, pv, tv, pvTv, ev, equity: ev - v.netDebt, perShare: (ev - v.netDebt) / v.shares };
}

const MODELS = [
  {
    id: 'dcf', name: 'DCF Valuation', goals: ['valuation', 'acquisition'],
    desc: 'Discounted cash flow: projects free cash flow, then discounts it back at WACC to find what the business is worth today.',
    inputs: (d) => [['rev', 'Base revenue ($)', d.rev], ['g', 'Revenue growth %', d.g], ['m', 'EBITDA margin %', d.m], ['da', 'D&A % of revenue', d.da],
      ['capex', 'CapEx % of revenue', d.capex], ['nwc', 'Working capital % of new revenue', 10], ['tax', 'Tax rate %', 25], ['wacc', 'WACC %', 10],
      ['tg', 'Terminal growth %', 2.5], ['years', 'Forecast years', 5], ['netDebt', 'Net debt ($)', d.netDebt], ['shares', 'Shares outstanding', 1_000_000]],
    run(v) {
      const r = dcfCore(v);
      if (r.error) return { html: `<div class="glass card">${r.error}</div>` };
      const ws = [-2, -1, 0, 1, 2].map((x) => v.wacc + x), gs = [-1, -0.5, 0, 0.5, 1].map((x) => v.tg + x);
      const sens = ws.map((w) => [`${w.toFixed(1)}%`, ...gs.map((g) => { const s = dcfCore({ ...v, wacc: w, tg: g }); return s.error ? 'n/a' : fmt(s.ev); })]);
      const scen = [['Bear', -3, -3], ['Base', 0, 0], ['Bull', 3, 3]].map(([n, dg, dm]) => {
        const s = dcfCore({ ...v, g: v.g + dg, m: v.m + dm });
        return [n, `${(v.g + dg).toFixed(1)}%`, `${(v.m + dm).toFixed(1)}%`, fmt(s.ev), fmt(s.equity), fmt(s.perShare)];
      });
      return {
        html: `<div class="kpis">${kpi('Enterprise value', fmt(r.ev))}${kpi('Equity value', fmt(r.equity))}${kpi('Value per share', fmt(r.perShare))}${kpi('Terminal value share', pct(r.pvTv / r.ev), 'of enterprise value')}</div>
        <div class="glass card"><h3>Free cash flow projection</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Projection table</h3>${table(['Year', 'Revenue', 'EBITDA', 'EBIT', 'CapEx', 'Free cash flow', 'PV of FCF'], r.rows.map((x) => [`Y${x.y}`, fmt(x.rev), fmt(x.ebitda), fmt(x.ebit), fmt(x.capex), fmt(x.fcf), fmt(x.pv)]))}</div>
        <div class="glass card"><h3>Sensitivity: enterprise value (WACC ↓ × terminal growth →)</h3>${table(['WACC \\ g', ...gs.map((g) => g.toFixed(1) + '%')], sens, (ri, ci) => ri === 2 && ci === 3)}</div>
        <div class="glass card"><h3>Scenarios</h3>${table(['Scenario', 'Growth', 'Margin', 'EV', 'Equity', 'Per share'], scen)}</div>`,
        charts: [['mc1', { type: 'bar', data: { labels: r.rows.map((x) => `Y${x.y}`), datasets: [ds('Free cash flow', r.rows.map((x) => x.fcf), COLORS.blue), ds('Present value', r.rows.map((x) => x.pv), COLORS.green)] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
  {
    id: 'npv', name: 'NPV / IRR', goals: ['investment'],
    desc: 'Checks whether an investment pays off. Enter the upfront cost and the expected cash flows for each year.',
    inputs: (d) => [['inv', 'Initial investment ($)', Math.round(d.ebitda * 2)], ['flows', 'Yearly cash flows (comma separated)', [1, 2, 3, 4, 5].map((y) => Math.round(d.ebitda * 0.5 * Math.pow(1.05, y))).join(', '), 'text'],
      ['rate', 'Discount rate %', 10], ['reinv', 'Reinvestment rate % (MIRR)', 8]],
    run(v) {
      const flows = [-Math.abs(v.inv), ...list(v.flows)], r = v.rate / 100;
      const n = npv(r, flows), i = irr(flows), pv = npv(r, [0, ...flows.slice(1)]);
      let cum = 0;
      const cumArr = flows.map((f) => (cum += f));
      return {
        html: `<div class="kpis">${kpi('NPV', fmt(n), n >= 0 ? 'Creates value ✓' : 'Destroys value ✗', n >= 0 ? 'up' : 'down')}${kpi('IRR', pct(i), i != null && i > r ? 'Above hurdle rate' : 'Below hurdle rate', i != null && i > r ? 'up' : 'down')}${kpi('MIRR', pct(mirr(flows, r, v.reinv / 100)))}${kpi('Payback', yrs(payback(flows)))}${kpi('Discounted payback', yrs(payback(flows, r)))}${kpi('Profitability index', xf(pv / Math.abs(v.inv)))}</div>
        <div class="glass card"><h3>Cumulative cash flow</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card">${table(['Year', 'Cash flow', 'Discounted', 'Cumulative'], flows.map((f, t) => [t, fmt(f), fmt(f / Math.pow(1 + r, t)), fmt(cumArr[t])]))}</div>`,
        charts: [['mc1', { data: { labels: flows.map((_, t) => `Y${t}`), datasets: [ds('Cash flow', flows, COLORS.blue, { type: 'bar' }), ds('Cumulative', cumArr, COLORS.green, { type: 'line' })] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
  {
    id: 'lbo', name: 'LBO', goals: ['acquisition'],
    desc: 'Leveraged buyout: buy the company mostly with debt, pay the debt down with cash flow, then sell. Shows the investor return.',
    inputs: (d) => [['ebitda', 'Entry EBITDA ($)', Math.round(d.ebitda)], ['entry', 'Entry multiple (x EBITDA)', 8], ['lev', 'Debt (x EBITDA)', 4.5], ['rate', 'Interest rate %', 8],
      ['g', 'EBITDA growth %', Math.max(3, d.g)], ['exit', 'Exit multiple (x EBITDA)', 8], ['years', 'Holding period (years)', 5], ['tax', 'Tax rate %', 25],
      ['capex', 'CapEx % of EBITDA', 15], ['sweep', 'Cash used to repay debt %', 100], ['fees', 'Transaction fees % of EV', 2]],
    run(v) {
      const ev0 = v.ebitda * v.entry, debt0 = v.ebitda * v.lev, eq0 = ev0 * (1 + v.fees / 100) - debt0;
      if (eq0 <= 0) return { html: '<div class="glass card">Debt is larger than the purchase price. Lower the debt multiple.</div>' };
      let e = v.ebitda, debt = debt0, cash = 0;
      const rows = [];
      for (let y = 1; y <= v.years; y++) {
        e *= 1 + v.g / 100;
        const int = debt * (v.rate / 100), capex = e * (v.capex / 100);
        const tax = Math.max(0, (e - int - capex) * (v.tax / 100));
        const fcf = e - int - capex - tax;
        const pay = Math.min(debt, Math.max(0, fcf * (v.sweep / 100)));
        debt -= pay; cash += fcf - pay;
        rows.push({ y, e, int, fcf, pay, debt, eqv: e * v.exit - debt + cash });
      }
      const exitEq = rows.at(-1).eqv, moic = exitEq / eq0, ir = Math.pow(Math.max(moic, 0), 1 / v.years) - 1;
      return {
        html: `<div class="kpis">${kpi('Purchase price (EV)', fmt(ev0))}${kpi('Equity invested', fmt(eq0))}${kpi('Exit equity', fmt(exitEq))}${kpi('IRR', pct(ir), ir >= 0.2 ? 'Meets typical 20% PE target' : 'Below typical 20% PE target', ir >= 0.2 ? 'up' : 'down')}${kpi('MOIC', xf(moic))}</div>
        <div class="glass card"><h3>Debt paydown vs equity value</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Debt schedule</h3>${table(['Year', 'EBITDA', 'Interest', 'Free cash flow', 'Debt repaid', 'Debt balance', 'Equity value'], rows.map((x) => [`Y${x.y}`, fmt(x.e), fmt(x.int), fmt(x.fcf), fmt(x.pay), fmt(x.debt), fmt(x.eqv)]))}</div>`,
        charts: [['mc1', { type: 'bar', data: { labels: ['Entry', ...rows.map((x) => `Y${x.y}`)], datasets: [ds('Debt', [debt0, ...rows.map((x) => x.debt)], COLORS.orange, { stack: 's' }), ds('Equity value', [eq0, ...rows.map((x) => x.eqv)], COLORS.green, { stack: 's' })] }, options: { scales: { y: { ...moneyAxis, stacked: true }, x: { stacked: true } } } }]],
      };
    },
  },
  {
    id: 'wacc', name: 'WACC', goals: ['valuation', 'acquisition'],
    desc: 'Weighted average cost of capital: the blended return that lenders and shareholders expect. Use it as the discount rate in the DCF.',
    inputs: (d) => [['E', 'Equity value ($)', Math.round(d.equity)], ['D', 'Debt value ($)', Math.round(d.debt)], ['rf', 'Risk-free rate %', 4.2], ['beta', 'Beta', 1.1],
      ['mrp', 'Market risk premium %', 5.5], ['sp', 'Size / specific premium %', 1], ['kd', 'Pre-tax cost of debt %', d.interest], ['tax', 'Tax rate %', 25]],
    run(v) {
      const ke = (v.rf + v.beta * v.mrp + v.sp) / 100, kd = (v.kd / 100) * (1 - v.tax / 100), V = v.E + v.D;
      const w = (v.E / V) * ke + (v.D / V) * kd;
      return {
        html: `<div class="kpis">${kpi('WACC', pct(w, 2))}${kpi('Cost of equity (CAPM)', pct(ke, 2))}${kpi('After-tax cost of debt', pct(kd, 2))}${kpi('Debt weight', pct(v.D / V))}</div>
        <div class="glass card"><h3>Capital structure</h3><div class="chart"><canvas id="mc1"></canvas></div></div>`,
        charts: [['mc1', { type: 'doughnut', data: { labels: ['Equity', 'Debt'], datasets: [{ data: [v.E, v.D], backgroundColor: [COLORS.green, COLORS.orange], borderWidth: 0 }] }, options: { cutout: '65%' } }]],
      };
    },
  },
  {
    id: 'comps', name: 'Comparables', goals: ['valuation', 'acquisition'],
    desc: 'Values the company using the multiples that similar listed companies trade at. Enter peer multiples separated by commas.',
    inputs: (d) => [['rev', 'Company revenue ($)', Math.round(d.rev)], ['ebitda', 'Company EBITDA ($)', Math.round(d.ebitda)], ['ni', 'Company net income ($)', Math.round(d.ni)],
      ['netDebt', 'Net debt ($)', Math.round(d.netDebt)], ['evE', 'Peer EV/EBITDA', '7.5, 8.2, 9.0, 10.1, 11.4', 'text'], ['evS', 'Peer EV/Sales', '1.1, 1.4, 1.6, 1.9', 'text'], ['pe', 'Peer P/E', '14, 16.5, 18, 21', 'text']],
    run(v) {
      const rows = [['EV / EBITDA', list(v.evE), v.ebitda, false], ['EV / Sales', list(v.evS), v.rev, false], ['P / E', list(v.pe), v.ni, true]]
        .filter((r) => r[1].length)
        .map(([n, m, base, isEq]) => {
          const toEq = (x) => (isEq ? x * base : x * base - v.netDebt);
          return { n, lo: toEq(Math.min(...m)), med: toEq(median(m)), hi: toEq(Math.max(...m)), mm: median(m) };
        });
      const mid = rows.length ? median(rows.map((r) => r.med)) : null;
      return {
        html: `<div class="kpis">${kpi('Implied equity value (median)', fmt(mid))}${rows.map((r) => kpi(r.n + ' median', xf(r.mm))).join('')}</div>
        <div class="glass card"><h3>Valuation range (football field)</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card">${table(['Method', 'Low', 'Median', 'High'], rows.map((r) => [r.n, fmt(r.lo), fmt(r.med), fmt(r.hi)]))}</div>`,
        charts: [['mc1', { type: 'bar', data: { labels: rows.map((r) => r.n), datasets: [ds('Equity value range', rows.map((r) => [r.lo, r.hi]), COLORS.purple)] }, options: { indexAxis: 'y', scales: { x: moneyAxis }, plugins: { legend: { display: false } } } }]],
      };
    },
  },
  {
    id: 'montecarlo', name: 'Monte Carlo', goals: ['valuation', 'investment'],
    desc: 'Runs the DCF many times with random growth and margins to show the range of likely values, not just one number.',
    inputs: (d) => [['rev', 'Base revenue ($)', Math.round(d.rev)], ['g', 'Growth mean %', d.g], ['gsd', 'Growth std. dev. %', 3], ['m', 'Margin mean %', d.m], ['msd', 'Margin std. dev. %', 3],
      ['wacc', 'WACC %', 10], ['tg', 'Terminal growth %', 2.5], ['netDebt', 'Net debt ($)', Math.round(d.netDebt)], ['sims', 'Simulations', 2000]],
    run(v, d) {
      const n = Math.min(Math.max(100, v.sims), 20000), out = [];
      for (let i = 0; i < n; i++) {
        const r = dcfCore({ rev: v.rev, g: v.g + randn() * v.gsd, m: v.m + randn() * v.msd, da: d.da, capex: d.capex, nwc: 10, tax: 25, wacc: v.wacc, tg: v.tg, years: 5, netDebt: v.netDebt, shares: 1 });
        if (!r.error) out.push(r.ev);
      }
      if (!out.length) return { html: '<div class="glass card">WACC must be higher than terminal growth.</div>' };
      out.sort((a, b) => a - b);
      const q = (p) => out[Math.floor(p * (out.length - 1))];
      const bins = 24, lo = q(0.01), hi = q(0.99), w = (hi - lo) / bins || 1, hist = Array(bins).fill(0);
      out.forEach((x) => { const b = Math.floor((x - lo) / w); if (b >= 0 && b < bins) hist[b]++; });
      const base = dcfCore({ ...v, da: d.da, capex: d.capex, nwc: 10, tax: 25, years: 5, shares: 1 }).ev;
      return {
        html: `<div class="kpis">${kpi('P10 (downside)', fmt(q(0.1)))}${kpi('P50 (median)', fmt(q(0.5)))}${kpi('P90 (upside)', fmt(q(0.9)))}${kpi('Chance above base case', pct(out.filter((x) => x > base).length / out.length, 0), `Base EV ${fmt(base)}`)}</div>
        <div class="glass card"><h3>Distribution of enterprise value (${out.length} runs)</h3><div class="chart"><canvas id="mc1"></canvas></div></div>`,
        charts: [['mc1', { type: 'bar', data: { labels: hist.map((_, i) => fmt(lo + w * (i + 0.5))), datasets: [ds('Simulations', hist, COLORS.teal, { barPercentage: 1, categoryPercentage: 1 })] }, options: { plugins: { legend: { display: false } } } }]],
      };
    },
  },
  {
    id: 'breakeven', name: 'Break-even', goals: ['investment', 'performance'],
    desc: 'How many units you must sell to cover all costs, and how much safety margin your sales plan has.',
    inputs: () => [['fixed', 'Fixed costs per year ($)', 500000], ['price', 'Price per unit ($)', 50], ['vc', 'Variable cost per unit ($)', 30], ['units', 'Expected units sold', 40000]],
    run(v) {
      const cm = v.price - v.vc;
      if (cm <= 0) return { html: '<div class="glass card">Price must be higher than variable cost.</div>' };
      const be = v.fixed / cm, profit = v.units * cm - v.fixed, mos = (v.units - be) / v.units;
      const max = Math.max(be * 2, v.units * 1.2), pts = Array.from({ length: 11 }, (_, i) => (max / 10) * i);
      return {
        html: `<div class="kpis">${kpi('Break-even units', Math.ceil(be).toLocaleString())}${kpi('Break-even revenue', fmt(be * v.price))}${kpi('Expected profit', fmt(profit), '', profit >= 0 ? 'up' : 'down')}${kpi('Margin of safety', pct(mos), '', mos >= 0 ? 'up' : 'down')}${kpi('Operating leverage', profit > 0 ? xf((v.units * cm) / profit) : '—')}</div>
        <div class="glass card"><h3>Revenue vs total cost</h3><div class="chart"><canvas id="mc1"></canvas></div></div>`,
        charts: [['mc1', { type: 'line', data: { labels: pts.map((u) => Math.round(u).toLocaleString()), datasets: [ds('Revenue', pts.map((u) => u * v.price), COLORS.green), ds('Total cost', pts.map((u) => v.fixed + u * v.vc), COLORS.red), ds('Fixed cost', pts.map(() => v.fixed), COLORS.orange, { borderDash: [6, 4] })] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
  {
    id: 'ratios', name: 'Ratio Analysis', goals: ['performance'],
    desc: 'Profitability, liquidity, leverage and return ratios calculated from your uploaded statements for every year.',
    inputs: () => [],
    run() {
      if (!hasData()) return { html: '<div class="glass card empty"><h3>Upload statements to see ratios</h3><a class="btn" href="#upload">Upload data</a></div>' };
      const all = ratios(), names = all[0].map((r) => r[0]);
      return { html: `<div class="glass card">${table(['Ratio', ...state.years.map(esc)], names.map((n, i) => [n, ...all.map((y) => y[i][1])]))}</div>` };
    },
  },
  {
    id: 'forecast', name: '3-Statement Forecast', goals: ['performance', 'valuation'],
    desc: 'Projects the income statement forward from your history, with simple cash and debt roll-forward.',
    inputs: (d) => [['rev', 'Base revenue ($)', Math.round(d.rev)], ['g', 'Revenue growth %', d.g], ['gm', 'Gross margin %', hasData() && L('grossProfit') ? +((L('grossProfit') / L('revenue')) * 100).toFixed(1) : 40],
      ['opex', 'Operating expenses % of revenue', hasData() && L('opex') ? +((L('opex') / L('revenue')) * 100).toFixed(1) : 20], ['da', 'D&A % of revenue', d.da], ['capex', 'CapEx % of revenue', d.capex],
      ['debt', 'Opening debt ($)', Math.round(d.debt)], ['repay', 'Debt repaid per year ($)', Math.round(d.debt / 10)], ['rate', 'Interest rate %', d.interest], ['cash', 'Opening cash ($)', Math.round(L('cash') || 0)], ['tax', 'Tax rate %', 25], ['years', 'Years', 5]],
    run(v) {
      let rev = v.rev, debt = v.debt, cash = v.cash;
      const rows = [];
      const y0 = parseInt(String(state.years.at(-1) || '').match(/\d{4}/)?.[0]) || new Date().getFullYear();
      for (let y = 1; y <= v.years; y++) {
        rev *= 1 + v.g / 100;
        const gp = rev * (v.gm / 100), ebitda = gp - rev * (v.opex / 100), da = rev * (v.da / 100), int = debt * (v.rate / 100);
        const ebt = ebitda - da - int, ni = ebt - Math.max(0, ebt * (v.tax / 100));
        const pay = Math.min(debt, v.repay);
        cash += ni + da - rev * (v.capex / 100) - pay;
        debt -= pay;
        rows.push([`${y0 + y}E`, rev, gp, ebitda, ni, cash, debt]);
      }
      const hist = hasData() ? state.years.map((y, i) => [y, S('revenue')[i]]) : [];
      return {
        html: `<div class="glass card"><h3>Revenue: history and forecast</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card">${table(['Year', 'Revenue', 'Gross profit', 'EBITDA', 'Net income', 'Cash', 'Debt'], rows.map((r) => [r[0], ...r.slice(1).map((x) => fmt(x))]))}</div>`,
        charts: [['mc1', { type: 'bar', data: { labels: [...hist.map((h) => h[0]), ...rows.map((r) => r[0])], datasets: [ds('Historical', [...hist.map((h) => h[1]), ...rows.map(() => null)], COLORS.blue), ds('Forecast', [...hist.map(() => null), ...rows.map((r) => r[1])], COLORS.purple), ds('Net income', [...hist.map((_, i) => S('netIncome')?.[i] ?? null), ...rows.map((r) => r[4])], COLORS.green, { type: 'line' })] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
  {
    id: 'loan', name: 'Loan Schedule', goals: ['investment'],
    desc: 'Amortizing loan: monthly payment, total interest and the balance over time.',
    inputs: (d) => [['P', 'Loan amount ($)', Math.round(d.debt)], ['rate', 'Annual interest rate %', d.interest], ['years', 'Term (years)', 7], ['ppy', 'Payments per year', 12]],
    run(v) {
      const n = v.years * v.ppy, r = v.rate / 100 / v.ppy, pmt = r ? (v.P * r) / (1 - Math.pow(1 + r, -n)) : v.P / n;
      let bal = v.P;
      const rows = [];
      for (let y = 1; y <= v.years; y++) {
        let int = 0, prin = 0;
        for (let k = 0; k < v.ppy; k++) { const i = bal * r; int += i; prin += pmt - i; bal -= pmt - i; }
        rows.push([`Y${y}`, pmt * v.ppy, int, prin, Math.max(0, bal)]);
      }
      return {
        html: `<div class="kpis">${kpi('Payment', fmt(pmt), `${v.ppy}× per year`)}${kpi('Total interest', fmt(pmt * n - v.P))}${kpi('Total paid', fmt(pmt * n))}</div>
        <div class="glass card"><h3>Balance and interest</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card">${table(['Year', 'Paid', 'Interest', 'Principal', 'End balance'], rows.map((r) => [r[0], ...r.slice(1).map((x) => fmt(x))]))}</div>`,
        charts: [['mc1', { data: { labels: rows.map((r) => r[0]), datasets: [ds('Principal', rows.map((r) => r[3]), COLORS.blue, { type: 'bar', stack: 'p' }), ds('Interest', rows.map((r) => r[2]), COLORS.orange, { type: 'bar', stack: 'p' }), ds('Balance', rows.map((r) => r[4]), COLORS.green, { type: 'line' })] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
];

let currentModel = 'dcf', currentGoal = 'all';

function renderModelTabs() {
  const visible = MODELS.filter((m) => currentGoal === 'all' || m.goals.includes(currentGoal));
  if (!visible.some((m) => m.id === currentModel)) currentModel = visible[0].id;
  $('#modelTabs').innerHTML = visible.map((m) => `<button data-model="${m.id}" class="${m.id === currentModel ? 'on' : ''}">${m.name}</button>`).join('');
  renderModel();
}

function renderModel() {
  const m = MODELS.find((x) => x.id === currentModel), d = defaults();
  const fields = m.inputs(d).map(([k, label, val, type]) =>
    `<div class="field"><label for="in-${k}">${label}</label><input id="in-${k}" data-k="${k}" ${type === 'text' ? 'type="text"' : 'type="number" step="any"'} value="${esc(val)}"></div>`).join('');
  $('#modelArea').innerHTML = `<div class="model"><div class="glass inputs"><h3>${m.name}</h3><p class="muted small">${m.desc}</p>${fields || `<p class="small">${hasData() ? `Using data for <strong>${esc(state.company)}</strong>.` : ''}</p>`}${hasData() ? '' : '<p class="small muted">Tip: upload statements to pre-fill these inputs.</p>'}</div><div class="results" id="results"></div></div>`;
  $$('#modelArea input').forEach((i) => i.addEventListener('input', runModel));
  runModel();
}

function runModel() {
  const m = MODELS.find((x) => x.id === currentModel), v = {};
  $$('#modelArea input').forEach((i) => { v[i.dataset.k] = i.type === 'text' ? i.value : parseFloat(i.value) || 0; });
  const res = m.run(v, defaults());
  $('#results').innerHTML = res.html;
  (res.charts || []).forEach(([id, cfg]) => makeChart(id, cfg));
}

/* ---------- Feasibility report ---------- */
const FEAS = [['name', 'Project name', 'New Store Expansion', 'text'], ['city', 'Location (city)', 'Miami', 'text'], ['inv', 'Initial investment ($)', 2_000_000], ['rev', 'Year-1 revenue ($)', 1_200_000],
  ['g', 'Revenue growth %', 6], ['var', 'Variable costs % of revenue', 45], ['fixed', 'Fixed costs per year ($)', 250_000], ['life', 'Project life (years)', 10],
  ['rate', 'Discount rate / hurdle %', 10], ['tax', 'Tax rate %', 25], ['salvage', 'Salvage value at end ($)', 200_000]];

function feasFlows(v, revMult = 1, costMult = 1) {
  const dep = (v.inv - v.salvage) / v.life, flows = [-v.inv], rows = [];
  for (let t = 1; t <= v.life; t++) {
    const rev = v.rev * revMult * Math.pow(1 + v.g / 100, t - 1);
    const ebitda = rev * (1 - (v.var / 100) * costMult) - v.fixed * costMult;
    const tax = Math.max(0, (ebitda - dep) * (v.tax / 100));
    const cf = ebitda - tax + (t === v.life ? v.salvage : 0);
    flows.push(cf);
    rows.push([t, rev, ebitda, tax, cf]);
  }
  return { flows, rows };
}

async function climateRisk(city) {
  const g = await geocode(city);
  const end = new Date(Date.now() - 7 * 864e5), start = new Date(end.getTime() - 364 * 864e5), iso = (d) => d.toISOString().slice(0, 10);
  const r = await fetch(`https://archive-api.open-meteo.com/v1/archive?latitude=${g.lat}&longitude=${g.lon}&start_date=${iso(start)}&end_date=${iso(end)}&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,wind_speed_10m_max&timezone=auto`);
  const j = await r.json();
  if (!j.daily) throw new Error(j.reason || 'No climate data');
  const dl = j.daily, cnt = (arr, f) => arr.filter((x) => x != null && f(x)).length;
  const heat = cnt(dl.temperature_2m_max, (x) => x >= 35), freeze = cnt(dl.temperature_2m_min, (x) => x <= 0);
  const rain = cnt(dl.precipitation_sum, (x) => x >= 20), wind = cnt(dl.wind_speed_10m_max, (x) => x >= 50);
  const score = Math.min(100, Math.round(heat * 0.5 + freeze * 0.3 + rain * 1.5 + wind * 1.5));
  const months = {};
  dl.time.forEach((t, i) => {
    const k = t.slice(0, 7);
    months[k] ??= { p: 0, t: 0, n: 0 };
    months[k].p += dl.precipitation_sum[i] || 0;
    months[k].t += ((dl.temperature_2m_max[i] ?? 0) + (dl.temperature_2m_min[i] ?? 0)) / 2;
    months[k].n++;
  });
  return { place: g.name, heat, freeze, rain, wind, score, level: score < 20 ? 'Low' : score < 45 ? 'Moderate' : 'High', months, totalRain: dl.precipitation_sum.reduce((a, b) => a + (b || 0), 0), period: `${iso(start)} to ${iso(end)}` };
}

async function generateReport() {
  const v = {};
  FEAS.forEach(([k, , , t]) => { const val = $(`#f-${k}`).value; v[k] = t === 'text' ? val.trim() : parseFloat(val) || 0; });
  v.life = Math.max(1, Math.round(v.life));
  const out = $('#report');
  out.innerHTML = '<div class="glass card"><p class="muted">Building report and fetching climate data…</p></div>';
  const r = v.rate / 100, { flows, rows } = feasFlows(v);
  const n = npv(r, flows), i = irr(flows), pb = payback(flows), dpb = payback(flows, r), pi = npv(r, [0, ...flows.slice(1)]) / v.inv;
  const beRev = v.fixed / (1 - v.var / 100);
  const sens = [['Revenue −20%', 0.8, 1], ['Revenue −10%', 0.9, 1], ['Base case', 1, 1], ['Revenue +10%', 1.1, 1], ['Costs +10%', 1, 1.1]].map(([lab, rm, cm]) => {
    const f = feasFlows(v, rm, cm).flows; return [lab, fmt(npv(r, f)), pct(irr(f)), yrs(payback(f))];
  });

  let climate = null, climateErr = '';
  try { climate = await climateRisk(v.city); } catch (e) { climateErr = e.message; }

  const checks = [
    [n > 0, `NPV is ${fmt(n)} at a ${v.rate}% hurdle rate`],
    [i != null && i > r, `IRR of ${pct(i)} ${i != null && i > r ? 'beats' : 'is below'} the ${v.rate}% hurdle`],
    [pb != null && pb <= v.life * 0.6, `Payback in ${yrs(pb)} out of a ${v.life}-year life`],
    [!climate || climate.level !== 'High', climate ? `Climate risk at ${esc(climate.place)} is ${climate.level}` : 'Climate risk not assessed'],
  ];
  const passed = checks.filter((c) => c[0]).length;
  const [cls, title] = n <= 0 ? ['bad', 'Not feasible'] : passed === 4 ? ['good', 'Feasible'] : ['mid', 'Feasible with conditions'];

  const recs = [];
  if (n <= 0) recs.push('Re-scope the project: reduce the upfront investment or fixed costs, or find ways to raise prices.');
  if (sens[0][1].startsWith('-')) recs.push('The project loses value if revenue comes in 20% below plan. Validate demand (pilot, pre-orders, market study) before committing.');
  if (pb == null || pb > v.life * 0.6) recs.push('Payback is slow. Consider phasing the investment or leasing assets instead of buying them.');
  if (climate?.level === 'High') recs.push('Weather is a material risk at this location. Budget for insurance, backup power and supply-chain buffers.');
  else if (climate?.level === 'Moderate') recs.push('Some severe-weather days occur at this location. Review insurance coverage and business continuity plans.');
  if (v.var >= 60) recs.push('Variable costs are high relative to revenue. Negotiate supplier terms to widen the contribution margin.');
  if (!recs.length) recs.push('Proceed to detailed planning. Track actual revenue against the base case each quarter.');

  const mk = climate ? Object.keys(climate.months) : [];
  out.innerHTML = `<div class="glass card report">
    <h2>${esc(v.name)}: Feasibility Report</h2>
    <p class="muted small">${esc(v.city)} · Prepared ${new Date().toLocaleDateString()}${hasData() ? ` · Company: ${esc(state.company)}` : ''}</p>
    <div class="verdict ${cls}"><strong>${title}</strong>${passed} of 4 checks passed.</div>
    <h3>Executive summary</h3>
    <ul>${checks.map(([ok, t]) => `<li>${ok ? '✅' : '⚠️'} ${t}</li>`).join('')}</ul>
    <div class="kpis">${kpi('NPV', fmt(n), '', n >= 0 ? 'up' : 'down')}${kpi('IRR', pct(i))}${kpi('Payback', yrs(pb))}${kpi('Discounted payback', yrs(dpb))}${kpi('Profitability index', xf(pi))}${kpi('Break-even revenue', fmt(beRev), 'per year')}</div>
  </div>
  <div class="glass card"><h3>Projected cash flows</h3><div class="chart"><canvas id="fc1"></canvas></div>
    ${table(['Year', 'Revenue', 'EBITDA', 'Tax', 'Cash flow'], [[0, '—', '—', '—', fmt(-v.inv)], ...rows.map((x) => [x[0], fmt(x[1]), fmt(x[2]), fmt(x[3]), fmt(x[4])])])}</div>
  <div class="glass card"><h3>Sensitivity analysis</h3>${table(['Scenario', 'NPV', 'IRR', 'Payback'], sens, (ri) => ri === 2)}</div>
  <div class="glass card"><h3>Climate &amp; weather risk <span class="tag">Open-Meteo</span></h3>
    ${climate ? `<p class="muted small">${esc(climate.place)} · last 12 months (${climate.period})</p>
    <div class="kpis">${kpi('Climate risk score', `${climate.score}/100`, climate.level, climate.level === 'Low' ? 'up' : 'down')}${kpi('Heavy rain days (≥20 mm)', climate.rain)}${kpi('Extreme heat days (≥35°C)', climate.heat)}${kpi('Freezing days', climate.freeze)}${kpi('High wind days (≥50 km/h)', climate.wind)}${kpi('Annual rainfall', `${climate.totalRain.toFixed(0)} mm`)}</div>
    <div class="chart" style="margin-top:14px"><canvas id="fc2"></canvas></div>` : `<p style="color:var(--red)">Climate data unavailable: ${esc(climateErr)}</p>`}
  </div>
  <div class="glass card report"><h3>Recommendations</h3><ul>${recs.map((x) => `<li>${x}</li>`).join('')}</ul>
    <p class="muted small">This report is based on the assumptions entered and historical weather data. It is a planning tool, not financial advice.</p></div>`;

  let cum = 0;
  makeChart('fc1', { data: { labels: flows.map((_, t) => `Y${t}`), datasets: [ds('Cash flow', flows, COLORS.blue, { type: 'bar' }), ds('Cumulative', flows.map((f) => (cum += f)), COLORS.green, { type: 'line' })] }, options: { scales: { y: moneyAxis } } });
  if (climate) makeChart('fc2', {
    data: { labels: mk.map((k) => new Date(k + '-01T00:00').toLocaleDateString(undefined, { month: 'short', year: '2-digit' })), datasets: [
      ds('Rain (mm)', mk.map((k) => +climate.months[k].p.toFixed(1)), COLORS.blue, { type: 'bar', yAxisID: 'y' }),
      ds('Avg temp (°C)', mk.map((k) => +(climate.months[k].t / climate.months[k].n).toFixed(1)), COLORS.orange, { type: 'line', yAxisID: 'y1' })] },
    options: { scales: { y: { position: 'left' }, y1: { position: 'right', grid: { display: false } } } },
  });
  $('#printReport').hidden = false;
}

/* ---------- Routing & init ---------- */
function route() {
  const view = (location.hash || '#dashboard').slice(1);
  const valid = ['dashboard', 'upload', 'models', 'feasibility'].includes(view) ? view : 'dashboard';
  $$('.view').forEach((v) => v.classList.toggle('active', v.id === `view-${valid}`));
  $$('nav a').forEach((a) => a.classList.toggle('active', a.dataset.view === valid));
  if (valid === 'dashboard') renderDashboard();
  if (valid === 'models') renderModelTabs();
  if (valid === 'upload') renderPreview();
}
function refreshAll() { $('#companyPill').textContent = hasData() ? state.company : 'No data loaded'; route(); }

function init() {
  themeCharts();
  load();
  $('#companyName').value = state.company || '';

  const drop = $('#drop'), fi = $('#fileInput');
  fi.addEventListener('change', () => { handleFile(fi.files[0]); fi.value = ''; });
  ['dragover', 'dragenter'].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (ev) => handleFile(ev.dataTransfer.files[0]));

  $$('[data-action="sample"]').forEach((b) => b.addEventListener('click', () => applyRows(parseCSV(SAMPLE_CSV), 'Aurora Retail Co. (sample)')));
  $('#dlTemplate').addEventListener('click', () => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([SAMPLE_CSV], { type: 'text/csv' }));
    a.download = 'financial-statements-template.csv';
    a.click();
  });
  $('#clearData').addEventListener('click', () => {
    state = { company: '', years: [], items: {} };
    try { localStorage.removeItem(STORE); } catch (e) {}
    $('#uploadMsg').textContent = 'Data cleared.';
    refreshAll();
  });

  $('#goalChips').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    currentGoal = b.dataset.goal;
    $$('#goalChips button').forEach((x) => x.classList.toggle('on', x === b));
    renderModelTabs();
  });
  $('#modelTabs').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    currentModel = b.dataset.model;
    renderModelTabs();
  });

  $('#feasForm').innerHTML = FEAS.map(([k, label, val, t]) =>
    `<div class="field"><label for="f-${k}">${label}</label><input id="f-${k}" ${t === 'text' ? 'type="text"' : 'type="number" step="any"'} value="${esc(val)}"></div>`).join('');
  $('#genReport').addEventListener('click', generateReport);
  $('#printReport').addEventListener('click', () => window.print());

  $('#weatherForm').addEventListener('submit', (e) => { e.preventDefault(); loadWeather($('#weatherCity').value.trim() || 'Miami'); });

  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { themeCharts(); route(); });
  window.addEventListener('hashchange', route);
  route();
  loadWeather('Miami');
}

document.addEventListener('DOMContentLoaded', init);
