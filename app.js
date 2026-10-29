/* Glass Finance Studio — plain JavaScript, no build step.
   Libraries: Chart.js (charts), SheetJS (Excel import and export). */

const $ = (s) => document.querySelector(s);
const $$ = (s) => [...document.querySelectorAll(s)];
const STORE = 'gfs-data-v1';
// Institutional palette: steel blue, green, gold, slate, red, teal (readable on light and dark)
const COLORS = { blue: '#2f6db5', green: '#2e8b57', orange: '#c19a3f', purple: '#6b7fa6', red: '#c0392b', teal: '#3a9ca6' };

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
  // Header = the row with the most year-like cells (e.g. 2024, FY2024, 2024A), so titles such as
  // "Year ended December 31, 2024" above the real header row are skipped
  const isYear = (c) => /^\D{0,4}(19|20)\d{2}\D{0,3}$/.test(String(c ?? '').trim());
  let hIdx = -1, best = 0;
  rows.forEach((r, i) => { const n = r.slice(1).filter(isYear).length; if (n > best) { best = n; hIdx = i; } });
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

// Formatted values keep their raw number (and an Excel number format) so exports stay numeric
const tag = (s, raw, z) => Object.assign(new String(s), { raw, z });
const bad = (n) => n == null || !isFinite(n);
function fmt(n, money = true) {
  if (bad(n)) return '—';
  const a = Math.abs(n), sign = n < 0 ? '-' : '', p = money ? '$' : '';
  const z = money ? '$#,##0;-$#,##0' : '#,##0';
  if (a >= 1e9) return tag(`${sign}${p}${(a / 1e9).toFixed(2)}B`, n, z);
  if (a >= 1e6) return tag(`${sign}${p}${(a / 1e6).toFixed(2)}M`, n, z);
  if (a >= 1e3) return tag(`${sign}${p}${(a / 1e3).toFixed(1)}K`, n, z);
  return tag(`${sign}${p}${a.toFixed(a < 10 && a % 1 ? 2 : 0)}`, n, money ? '$#,##0.00' : '#,##0.00');
}
const pct = (n, d = 1) => (bad(n) ? '—' : tag(`${(n * 100).toFixed(d)}%`, n, d > 1 ? '0.00%' : '0.0%'));
const xf = (n) => (bad(n) ? '—' : tag(`${n.toFixed(2)}x`, n, '0.00"x"'));
const yrs = (n) => (bad(n) ? 'Never' : tag(`${n.toFixed(1)} yrs`, n, '0.0 "yrs"'));
const rawAttr = (c) => (c != null && c.raw != null ? ` data-v="${c.raw}" data-z="${esc(c.z)}"` : typeof c === 'number' ? ` data-v="${c}"` : '');

const kpi = (label, value, delta = '', cls = '') =>
  `<div class="glass kpi"><div class="label">${label}</div><div class="value"${rawAttr(value)}>${value}</div>${delta ? `<div class="delta ${cls}">${delta}</div>` : ''}</div>`;
const table = (head, rows, hl) =>
  `<div class="table-wrap"><table><thead><tr>${head.map((h) => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows
    .map((r, ri) => `<tr>${r.map((c, ci) => `<td class="${hl && hl(ri, ci) ? 'hl' : ''}"${rawAttr(c)}>${c}</td>`).join('')}</tr>`)
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
const moneyAxis = { ticks: { callback: (v) => String(fmt(v)) } };
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

const SAMPLE_COMPANY = 'Aurora Retail Co. (sample)';
let uploadNote = '';

// Reusable upload panel, placed into every .uploader-slot (workspace, dashboard, My Data)
function uploaderHTML() {
  return `<div class="uploader">
    <label class="drop">
      <input type="file" class="file-in" accept=".pdf,.xlsx,.xls,.csv" hidden>
      <svg class="drop-icon" viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M12 18v-6M9.5 14.5 12 12l2.5 2.5"/></svg>
      <strong>Drag and drop your financial statements</strong>
      <span class="muted small">or <u>browse files</u></span>
      <span class="formats"><b>PDF</b><b>XLSX</b><b>XLS</b><b>CSV</b></span>
    </label>
    <div class="form-grid tight">
      <div class="field"><label>Company name</label><input class="co-name" placeholder="e.g. Acme Holdings" value="${esc(state.company === SAMPLE_COMPANY ? '' : state.company || '')}"></div>
      <div class="field"><label>Figures are in</label><select class="units">
        <option value="auto">Auto-detect</option><option value="1">Actual amounts</option><option value="1000">Thousands</option><option value="1000000">Millions</option>
      </select></div>
    </div>
    <div class="row">
      <button class="btn ghost" data-action="sample">Use sample company</button>
      <a class="link" href="samples/sample-financials.pdf" download>Sample PDF</a>
      <a class="link" href="samples/sample-financials.csv" download>CSV template</a>
    </div>
    <p class="upload-msg small" role="status">${uploadNote}</p>
  </div>`;
}
function renderUploaders() { $$('.uploader-slot').forEach((s) => { s.innerHTML = uploaderHTML(); }); }
function msg(t, err, root) {
  (root ? [root.querySelector('.upload-msg')] : $$('.upload-msg')).forEach((el) => { if (el) { el.textContent = t; el.className = `upload-msg small ${err ? 'err' : 'ok'}`; } });
}

function applyRows(rows, company, mult = 'auto', root, source = '') {
  const parsed = parseRows(rows);
  if (!Object.keys(parsed.items).length) {
    msg('No recognisable line items were found. Check that line items (e.g. "Revenue") are in the first column and years are in a header row.', true, root);
    return false;
  }
  if (mult === 'auto') {
    const text = rows.slice(0, 40).map((r) => r.join(' ')).join(' ').toLowerCase();
    // Matches headers such as "(in USD thousands)", "in millions of dollars", "$000s", "€m"
    mult = /\b(in|of)\b[^,.;]{0,20}\bmillions?\b|\(millions?\)|[$€£]\s?m\b|\busd ?m\b/.test(text) ? 1e6
      : /\b(in|of)\b[^,.;]{0,20}\bthousands?\b|\(thousands?\)|\(000s?\)|[$€£]\s?000|'000|\busd ?000/.test(text) ? 1e3 : 1;
  }
  const items = Object.fromEntries(Object.entries(parsed.items).map(([k, a]) => [k, a.map((v) => (v == null ? null : v * +mult))]));
  state = { company: company || 'My Company', years: parsed.years, items: derive(items) };
  dash = { from: 0, to: null, metric: 'revenue' };
  save();
  const found = Object.keys(parsed.items).length;
  uploadNote = `${source ? esc(source) + ': ' : ''}found ${found} line items across ${parsed.years.length} periods${+mult > 1 ? ` (scaled from ${+mult === 1e6 ? 'millions' : 'thousands'})` : ''}.${S('revenue') ? '' : ' No revenue line was found, so some models will use default values.'}`;
  refreshAll();
  return true;
}
const loadSample = () => applyRows(parseCSV(SAMPLE_CSV), SAMPLE_COMPANY, 1, null, 'Sample company');

/* PDF reading (pdf.js is loaded only when a PDF is uploaded) */
const PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/';
let pdfjsReady = null;
function loadPdfJs() {
  return (pdfjsReady ??= new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = PDFJS + 'pdf.min.js';
    s.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS + 'pdf.worker.min.js'; res(window.pdfjsLib); };
    s.onerror = () => { pdfjsReady = null; rej(new Error('The PDF reader failed to load. Check your internet connection.')); };
    document.head.appendChild(s);
  }));
}
const NUM_TOKEN = /^\(?[-–]?\$?\(?\d[\d,]*(\.\d+)?\)?%?$|^[—–-]$/;
function lineToRow(text) {
  const toks = text.replace(/\$\s+/g, '$').replace(/\(\s+/g, '(').replace(/\s+\)/g, ')').split(/\s+/).filter(Boolean);
  let i = toks.length;
  while (i > 0 && NUM_TOKEN.test(toks[i - 1])) i--;
  // Header lines like "FY2024 FY2023" have no numbers but several years
  const yearToks = toks.filter((t) => /^\D{0,4}(19|20)\d{2}\D{0,3}$/.test(t));
  if (i === toks.length && yearToks.length >= 2) return ['', ...yearToks];
  return [toks.slice(0, i).join(' ').replace(/[.:\s]+$/, ''), ...toks.slice(i).map((t) => (/^[—–-]$/.test(t) ? '' : t))];
}
async function pdfToRows(buf) {
  const lib = await loadPdfJs();
  const pdf = await lib.getDocument({ data: buf }).promise;
  const rows = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const { items } = await (await pdf.getPage(n)).getTextContent();
    const lines = [];
    items.filter((it) => it.str.trim()).forEach((it) => {
      const x = it.transform[4], y = it.transform[5];
      let line = lines.find((l) => Math.abs(l.y - y) < 3);
      if (!line) lines.push((line = { y, parts: [] }));
      line.parts.push({ x, s: it.str.trim() });
    });
    lines.sort((a, b) => b.y - a.y).forEach((l) => rows.push(lineToRow(l.parts.sort((a, b) => a.x - b.x).map((p) => p.s).join(' '))));
  }
  return rows;
}

async function handleFile(file, root) {
  if (!file) return;
  const company = root?.querySelector('.co-name')?.value.trim() || file.name.replace(/\.[^.]+$/, '').replace(/[-_]+/g, ' ');
  const mult = root?.querySelector('.units')?.value || 'auto';
  msg(`Reading ${file.name}…`, false, root);
  try {
    let rows;
    if (/\.csv$/i.test(file.name)) rows = parseCSV(await file.text());
    else if (/\.pdf$/i.test(file.name)) {
      rows = await pdfToRows(await file.arrayBuffer());
      if (!rows.length) throw new Error('No text found in this PDF. It may be a scanned image. Please upload the Excel or CSV version instead.');
    } else {
      if (!window.XLSX) throw new Error('The Excel reader failed to load. Check your internet connection, or upload a CSV.');
      const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
      // Combine all sheets so separate P&L / balance sheet tabs are all read
      rows = wb.SheetNames.flatMap((n) => XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }));
    }
    applyRows(rows, company, mult, root, file.name);
  } catch (e) {
    msg(e.message || String(e), true, root);
  }
}

function renderPreview() {
  const card = $('#previewCard');
  card.hidden = !Object.keys(state.items).length;
  if (card.hidden) return;
  $('#preview').innerHTML = table(['Line item', ...state.years.map(esc)],
    Object.keys(LABELS).filter((k) => S(k)).map((k) => [LABELS[k], ...S(k).map((v) => fmt(v))]));
}

/* ---------- Dashboard (Power BI style) ---------- */
let dash = { from: 0, to: null, metric: 'revenue' };

function renderDashboard() {
  const on = hasData();
  $('#dashEmpty').hidden = on;
  $('#dashContent').hidden = !on;
  $('#dlPowerBI').hidden = !on;
  $('#companyPill').textContent = on ? state.company : 'No data loaded';
  if (!on) return;

  const n = state.years.length;
  if (dash.to == null || dash.to >= n) dash.to = n - 1;
  if (dash.from > dash.to) dash.from = 0;
  if (!S(dash.metric)) dash.metric = 'revenue';
  const yearOpts = (sel) => state.years.map((y, i) => `<option value="${i}"${i === sel ? ' selected' : ''}>${esc(y)}</option>`).join('');
  $('#slFrom').innerHTML = yearOpts(dash.from);
  $('#slTo').innerHTML = yearOpts(dash.to);
  $('#slMetric').innerHTML = Object.keys(LABELS).filter((k) => S(k)).map((k) => `<option value="${k}"${k === dash.metric ? ' selected' : ''}>${LABELS[k]}</option>`).join('');

  const b = dash.from, e = dash.to;
  const idx = Array.from({ length: e - b + 1 }, (_, i) => b + i);
  const years = idx.map((i) => state.years[i]);
  const at = (k, i) => S(k)?.[i] ?? null;
  const sl = (k) => (S(k) ? idx.map((i) => S(k)[i]) : []);
  const yoy = (k, i) => (i > 0 && at(k, i - 1) ? at(k, i) / at(k, i - 1) - 1 : null);
  const ye = esc(state.years[e]);

  const g = yoy('revenue', e);
  const span = `${esc(state.years[b])}–${ye}`;
  const netDebt = at('debt', e) != null ? at('debt', e) - (at('cash', e) || 0) : null;
  $('#kpis').innerHTML = [
    kpi(`Revenue ${ye}`, fmt(at('revenue', e)), g == null ? '' : `${g >= 0 ? '▲' : '▼'} ${pct(Math.abs(g))} vs prior year`, g >= 0 ? 'up' : 'down'),
    kpi('EBITDA', fmt(at('ebitda', e)), `${pct(at('ebitda', e) / at('revenue', e))} margin`),
    kpi('Net Income', fmt(at('netIncome', e)), `${pct(at('netIncome', e) / at('revenue', e))} margin`),
    netDebt != null && netDebt < 0
      ? kpi('Net cash', fmt(-netDebt), 'More cash than debt', 'up')
      : kpi('Net debt', fmt(netDebt), netDebt != null && at('ebitda', e) ? `${xf(netDebt / at('ebitda', e))} EBITDA` : ''),
    kpi('Revenue CAGR', pct(cagr(sl('revenue'))), span),
  ].join('');

  // Trend: selected metric (columns) + YoY growth (line); click a column to focus that year
  $('#trendTitle').textContent = `${LABELS[dash.metric]} trend, ${state.years[b]}–${state.years[e]}`;
  makeChart('chTrend', {
    data: { labels: years, datasets: [
      ds(LABELS[dash.metric], sl(dash.metric), COLORS.blue, { type: 'bar', yAxisID: 'y', order: 2 }),
      ds('YoY growth %', idx.map((i) => (yoy(dash.metric, i) == null ? null : +(yoy(dash.metric, i) * 100).toFixed(1))), COLORS.orange, { type: 'line', yAxisID: 'y1', order: 1 }),
    ] },
    options: {
      scales: { y: moneyAxis, y1: { position: 'right', grid: { display: false }, ticks: { callback: (v) => v + '%' } } },
      onClick: (_, els) => { if (els.length) { dash.to = idx[els[0].index]; renderDashboard(); } },
    },
  });

  const m = (k) => (S(k) ? idx.map((i) => (at(k, i) != null && at('revenue', i) ? +((at(k, i) / at('revenue', i)) * 100).toFixed(2) : null)) : []);
  makeChart('chMargin', {
    type: 'line',
    data: { labels: years, datasets: [ds('Gross', m('grossProfit'), COLORS.blue), ds('EBITDA', m('ebitda'), COLORS.green), ds('Net', m('netIncome'), COLORS.purple)] },
    options: { scales: { y: { ticks: { callback: (v) => v + '%' } } } },
  });

  // Profit bridge (waterfall) for the focus year: revenue down to net income
  $('#bridgeTitle').textContent = `Profit bridge, ${state.years[e]}`;
  const rev = at('revenue', e), ni = at('netIncome', e);
  if (rev != null && ni != null) {
    const abs = (k) => Math.abs(at(k, e) || 0);
    const cogs = abs('cogs'), da = abs('da'), int = abs('interest');
    const opex = at('opex', e) != null ? abs('opex') : at('ebitda', e) != null ? rev - cogs - at('ebitda', e) : 0;
    const steps = [['Revenue', rev, 'total'], ['COGS', -cogs], ['Opex', -opex], ['D&A', -da], ['Interest', -int], ['Tax & other', -(rev - cogs - opex - da - int - ni)], ['Net income', ni, 'total']];
    let run = 0;
    const bars = steps.map(([, v, t]) => { if (t) { run = v; return [0, v]; } const r = [run, run + v]; run += v; return r; });
    makeChart('chBridge', {
      type: 'bar',
      data: { labels: steps.map((s) => s[0]), datasets: [{ label: 'Amount', data: bars, borderRadius: 6,
        backgroundColor: steps.map(([, v, t]) => (t ? COLORS.blue : v < 0 ? COLORS.red : COLORS.green) + 'cc') }] },
      options: { plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => String(fmt(steps[c.dataIndex][1])) } } }, scales: { y: moneyAxis } },
    });
  }

  $('#balanceTitle').textContent = `Balance sheet, ${state.years[e]}`;
  const tl = at('totalLiabilities', e), debt = at('debt', e);
  const bs = [['Equity', at('equity', e), COLORS.green], ['Debt', debt, COLORS.orange], ['Other liabilities', tl != null && debt != null ? tl - debt : tl, COLORS.purple]].filter((x) => x[1] > 0);
  makeChart('chBalance', {
    type: 'doughnut',
    data: { labels: bs.map((x) => x[0]), datasets: [{ data: bs.map((x) => x[1]), backgroundColor: bs.map((x) => x[2]), borderWidth: 0 }] },
    options: { cutout: '65%', plugins: { legend: { position: 'right' } } },
  });

  // Matrix: every line item for the selected years, with change over the period
  $('#dashMatrix').innerHTML = table(['Line item', ...years.map(esc), `Change ${span}`],
    Object.keys(LABELS).filter((k) => S(k)).map((k) => {
      const s = at(k, b), f = at(k, e);
      return [LABELS[k], ...sl(k).map((v) => fmt(v)), s ? pct(f / s - 1) : '—'];
    }));

  $('#ratiosTitle').textContent = `Key ratios, ${state.years[b]} vs ${state.years[e]}`;
  const all = ratios();
  $('#dashRatios').innerHTML = table(['Ratio', esc(state.years[b]), ye], all[e].map(([k, v], i) => [k, all[b][i][1], v]));
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

/* ---------- Excel & Power BI export ---------- */
const slug = (s) => String(s || 'model').replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase();
const cellOf = (el) => (el.dataset.v != null && el.dataset.v !== '' ? { t: 'n', v: +el.dataset.v, z: el.dataset.z || 'General' } : el.textContent.trim());

function dataSheet() {
  return XLSX.utils.aoa_to_sheet([['Line item', ...state.years], ...Object.keys(LABELS).filter((k) => S(k)).map((k) => [LABELS[k], ...S(k)])]);
}

// Builds a workbook from what is on screen: inputs, KPI tiles and every results table
function exportXlsx(container, inputs, title, filename) {
  if (!window.XLSX) return alert('The Excel library failed to load. Please check your internet connection.');
  const aoa = [[title], [hasData() ? state.company : '', `Generated ${new Date().toLocaleDateString()}`], []];
  if (inputs.length) {
    aoa.push(['Inputs', 'Value']);
    inputs.forEach(([l, v]) => aoa.push([l, v]));
    aoa.push([]);
  }
  const tiles = container.querySelectorAll('.kpi');
  if (tiles.length) {
    aoa.push(['Key results', 'Value', 'Note']);
    tiles.forEach((k) => aoa.push([k.querySelector('.label').textContent, cellOf(k.querySelector('.value')), k.querySelector('.delta')?.textContent || '']));
    aoa.push([]);
  }
  container.querySelectorAll('.card').forEach((card) => {
    const t = card.querySelector('table');
    if (!t) return;
    aoa.push([card.querySelector('h3')?.textContent || 'Table']);
    t.querySelectorAll('tr').forEach((tr) => aoa.push([...tr.children].map(cellOf)));
    aoa.push([]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 36 }, ...Array(12).fill({ wch: 16 })];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Model');
  if (hasData()) XLSX.utils.book_append_sheet(wb, dataSheet(), 'Source data');
  XLSX.writeFile(wb, filename);
}

const CATEGORY = { revenue: 'Income Statement', cogs: 'Income Statement', grossProfit: 'Income Statement', opex: 'Income Statement', ebitda: 'Income Statement', da: 'Income Statement', ebit: 'Income Statement', interest: 'Income Statement', netIncome: 'Income Statement', capex: 'Cash Flow', ocf: 'Cash Flow' };

// Tidy (long) tables are the shape Power BI works with best: one row per company / year / metric
function exportPowerBI() {
  if (!window.XLSX) return alert('The Excel library failed to load. Please check your internet connection.');
  const yearNum = (y) => parseInt(String(y).match(/(19|20)\d{2}/)?.[0]) || String(y);
  const fact = [['Company', 'Year', 'Category', 'Metric', 'Value']];
  Object.keys(LABELS).filter((k) => S(k)).forEach((k) => S(k).forEach((v, i) => {
    if (v != null) fact.push([state.company, yearNum(state.years[i]), CATEGORY[k] || 'Balance Sheet', LABELS[k], v]);
  }));
  const rat = [['Company', 'Year', 'Ratio', 'Value', 'Format']];
  ratios().forEach((rows, i) => rows.forEach(([name, v]) => {
    if (v && v.raw != null) rat.push([state.company, yearNum(state.years[i]), name, v.raw, v.z.includes('%') ? 'Percent' : 'Multiple']);
  }));
  const readme = [
    ['How to use this file in Power BI Desktop'],
    ['1. Home > Get data > Excel workbook, and choose this file.'],
    ['2. Tick FactFinancials and Ratios, then click Load.'],
    ['3. Build visuals: put Year on the X-axis, Value in Values, and Metric (or Ratio) in Legend or a slicer.'],
    ['4. Format the Ratios Value column as a percentage where Format = Percent.'],
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fact), 'FactFinancials');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rat), 'Ratios');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readme), 'README');
  XLSX.writeFile(wb, `${slug(state.company)}-powerbi-dataset.xlsx`);
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
    id: 'dcf', name: 'DCF Valuation', group: 'Valuation',
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
    id: 'npv', name: 'NPV / IRR', group: 'Projects & investments',
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
    id: 'lbo', name: 'LBO', group: 'Deals & financing',
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
    id: 'wacc', name: 'WACC', group: 'Valuation',
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
    id: 'comps', name: 'Comparables', group: 'Valuation',
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
    id: 'montecarlo', name: 'Monte Carlo', group: 'Valuation',
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
        charts: [['mc1', { type: 'bar', data: { labels: hist.map((_, i) => String(fmt(lo + w * (i + 0.5)))), datasets: [ds('Simulations', hist, COLORS.teal, { barPercentage: 1, categoryPercentage: 1 })] }, options: { plugins: { legend: { display: false } } } }]],
      };
    },
  },
  {
    id: 'breakeven', name: 'Break-even', group: 'Projects & investments',
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
    id: 'ratios', name: 'Ratio Analysis', group: 'Performance',
    desc: 'Profitability, liquidity, leverage and return ratios calculated from your uploaded statements for every year.',
    inputs: () => [],
    run() {
      if (!hasData()) return { html: '<div class="glass card empty"><h3>Upload statements to see ratios</h3><a class="btn" href="#upload">Upload data</a></div>' };
      const all = ratios(), names = all[0].map((r) => r[0]);
      return { html: `<div class="glass card">${table(['Ratio', ...state.years.map(esc)], names.map((n, i) => [n, ...all.map((y) => y[i][1])]))}</div>` };
    },
  },
  {
    id: 'forecast', name: '3-Statement Forecast', group: 'Performance',
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
    id: 'loan', name: 'Loan Schedule', group: 'Deals & financing',
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

let currentModel = 'dcf';

function downloadModel() {
  const m = MODELS.find((x) => x.id === currentModel);
  const inputs = $$('#modelArea input').map((i) => [i.labels[0]?.textContent || i.dataset.k, i.type === 'text' ? i.value : parseFloat(i.value) || 0]);
  exportXlsx($('#results'), inputs, m.name, `${slug(hasData() ? state.company : 'model')}-${m.id}.xlsx`);
}

function renderModel() {
  const m = MODELS.find((x) => x.id === currentModel), d = defaults();
  const fields = m.inputs(d).map(([k, label, val, type]) =>
    `<div class="field"><label for="in-${k}">${label}</label><input id="in-${k}" data-k="${k}" ${type === 'text' ? 'type="text"' : 'type="number" step="any"'} value="${esc(val)}"></div>`).join('');
  $('#modelArea').innerHTML = `<div class="model"><div class="glass inputs"><h3>Assumptions</h3>${fields || `<p class="small">${hasData() ? `Using data for <strong>${esc(state.company)}</strong>.` : ''}</p>`}${hasData() ? '' : '<p class="small muted">Tip: upload statements to pre-fill these inputs.</p>'}</div><div class="results" id="results"></div></div>`;
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
const FEAS = [['name', 'Project name', 'New Store Expansion', 'text'], ['inv', 'Initial investment ($)', 2_000_000], ['rev', 'Year-1 revenue ($)', 1_200_000],
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

function generateReport() {
  const v = {};
  FEAS.forEach(([k, , , t]) => { const val = $(`#f-${k}`).value; v[k] = t === 'text' ? val.trim() : parseFloat(val) || 0; });
  v.life = Math.max(1, Math.round(v.life));
  const out = $('#report');
  const r = v.rate / 100, { flows, rows } = feasFlows(v);
  const n = npv(r, flows), i = irr(flows), pb = payback(flows), dpb = payback(flows, r), pi = npv(r, [0, ...flows.slice(1)]) / v.inv;
  const beRev = v.fixed / (1 - v.var / 100);
  const sensRaw = [['Revenue −20%', 0.8, 1], ['Revenue −10%', 0.9, 1], ['Base case', 1, 1], ['Revenue +10%', 1.1, 1], ['Costs +10%', 1, 1.1]].map(([lab, rm, cm]) => {
    const f = feasFlows(v, rm, cm).flows; return { lab, npv: npv(r, f), irr: irr(f), pb: payback(f) };
  });
  const sens = sensRaw.map((x) => [x.lab, fmt(x.npv), pct(x.irr), yrs(x.pb)]);

  const checks = [
    [n > 0, `NPV is ${fmt(n)} at a ${v.rate}% hurdle rate`],
    [i != null && i > r, `IRR of ${pct(i)} ${i != null && i > r ? 'beats' : 'is below'} the ${v.rate}% hurdle`],
    [pb != null && pb <= v.life * 0.6, `Payback in ${yrs(pb)} out of a ${v.life}-year life`],
    [sensRaw[1].npv > 0, `NPV ${sensRaw[1].npv > 0 ? 'stays positive' : 'turns negative'} if revenue is 10% below plan`],
  ];
  const passed = checks.filter((c) => c[0]).length;
  const [cls, title] = n <= 0 ? ['bad', 'Not feasible'] : passed === 4 ? ['good', 'Feasible'] : ['mid', 'Feasible with conditions'];

  const recs = [];
  if (n <= 0) recs.push('Re-scope the project: reduce the upfront investment or fixed costs, or find ways to raise prices.');
  if (sensRaw[0].npv < 0) recs.push('The project loses value if revenue comes in 20% below plan. Validate demand (pilot, pre-orders, market study) before committing.');
  if (pb == null || pb > v.life * 0.6) recs.push('Payback is slow. Consider phasing the investment or leasing assets instead of buying them.');
  if (sensRaw[4].npv < 0) recs.push('A 10% rise in costs wipes out the value. Lock in supplier prices or build in a cost contingency.');
  if (v.var >= 60) recs.push('Variable costs are high relative to revenue. Negotiate supplier terms to widen the contribution margin.');
  if (!recs.length) recs.push('Proceed to detailed planning. Track actual revenue against the base case each quarter.');

  out.innerHTML = `<div class="glass card report">
    <h2>${esc(v.name)}: Feasibility Report</h2>
    <p class="muted small">Prepared ${new Date().toLocaleDateString()}${hasData() ? ` · Company: ${esc(state.company)}` : ''}</p>
    <div class="verdict ${cls}"><strong>${title}</strong>${passed} of 4 checks passed.</div>
    <h3>Executive summary</h3>
    <ul>${checks.map(([ok, t]) => `<li>${ok ? '✅' : '⚠️'} ${t}</li>`).join('')}</ul>
    <div class="kpis">${kpi('NPV', fmt(n), '', n >= 0 ? 'up' : 'down')}${kpi('IRR', pct(i))}${kpi('Payback', yrs(pb))}${kpi('Discounted payback', yrs(dpb))}${kpi('Profitability index', xf(pi))}${kpi('Break-even revenue', fmt(beRev), 'per year')}</div>
  </div>
  <div class="glass card"><h3>Projected cash flows</h3><div class="chart"><canvas id="fc1"></canvas></div>
    ${table(['Year', 'Revenue', 'EBITDA', 'Tax', 'Cash flow'], [[0, '—', '—', '—', fmt(-v.inv)], ...rows.map((x) => [x[0], fmt(x[1]), fmt(x[2]), fmt(x[3]), fmt(x[4])])])}</div>
  <div class="glass card"><h3>Sensitivity analysis</h3>${table(['Scenario', 'NPV', 'IRR', 'Payback'], sens, (ri) => ri === 2)}</div>
  <div class="glass card report"><h3>Recommendations</h3><ul>${recs.map((x) => `<li>${x}</li>`).join('')}</ul>
    <p class="muted small">This report is based on the assumptions entered. It is a planning tool, not financial advice.</p></div>`;

  let cum = 0;
  makeChart('fc1', { data: { labels: flows.map((_, t) => `Y${t}`), datasets: [ds('Cash flow', flows, COLORS.blue, { type: 'bar' }), ds('Cumulative', flows.map((f) => (cum += f)), COLORS.green, { type: 'line' })] }, options: { scales: { y: moneyAxis } } });
  $('#printReport').hidden = false;
  $('#dlReport').hidden = false;
}

function downloadReport() {
  const inputs = FEAS.map(([k, label, , t]) => [label, t === 'text' ? $(`#f-${k}`).value : parseFloat($(`#f-${k}`).value) || 0]);
  exportXlsx($('#report'), inputs, `${$('#f-name').value || 'Project'}: Feasibility Report`, `${slug($('#f-name').value)}-feasibility.xlsx`);
}

/* ---------- Model library (catalog) ---------- */
const CATALOG = [
  { id: 'dcf', group: 'Valuation', file: 'dcf-valuation.xlsx',
    outputs: ['Enterprise value, equity value and value per share', '5-year free cash flow projection', 'WACC × terminal growth sensitivity; bear, base and bull cases'],
    needs: 'Income statement, balance sheet (cash and debt) and CapEx', best: 'Valuing a whole business from its future cash flows.' },
  { id: 'comps', group: 'Valuation', file: 'comparables.xlsx',
    outputs: ['Implied value from EV/EBITDA, EV/Sales and P/E', 'Low, median and high valuation range', 'Football-field chart'],
    needs: 'Revenue, EBITDA and net income, plus peer multiples', best: 'A market cross-check on your DCF.' },
  { id: 'wacc', group: 'Valuation', file: 'wacc.xlsx',
    outputs: ['WACC and cost of equity (CAPM)', 'After-tax cost of debt', 'Capital structure weights'],
    needs: 'Balance sheet (debt and equity) and market inputs', best: 'Setting the discount rate for valuations and projects.' },
  { id: 'montecarlo', group: 'Valuation', file: 'monte-carlo.xlsx',
    outputs: ['P10 / P50 / P90 enterprise value', 'Probability of beating the base case', 'Distribution of thousands of outcomes'],
    needs: 'Income statement (revenue and EBITDA)', best: 'Showing valuation risk to investors or a credit committee.' },
  { id: 'lbo', group: 'Deals & financing', file: 'lbo.xlsx',
    outputs: ['Sponsor IRR and MOIC', 'Year-by-year debt schedule with cash sweep', 'Entry vs exit equity value'],
    needs: 'Income statement (EBITDA)', best: 'Testing a private equity or management buyout.' },
  { id: 'loan', group: 'Deals & financing', file: 'loan-schedule.xlsx',
    outputs: ['Payment per period and total interest', 'Annual principal and interest split', 'Outstanding balance over time'],
    needs: 'Loan amount and terms (pre-filled from total debt)', best: 'Term loans, equipment finance and refinancing.' },
  { id: 'npv', group: 'Projects & investments', file: 'npv-irr.xlsx',
    outputs: ['NPV, IRR, MIRR and profitability index', 'Simple and discounted payback', 'Cumulative cash flow profile'],
    needs: 'Project cash flows (pre-filled from EBITDA)', best: 'Go / no-go decisions on capital projects.' },
  { id: 'breakeven', group: 'Projects & investments', file: 'break-even.xlsx',
    outputs: ['Break-even units and revenue', 'Margin of safety and operating leverage', 'Revenue vs cost chart'],
    needs: 'Fixed costs, price and variable cost per unit', best: 'Pricing, new products and cost planning.' },
  { id: 'feasibility', group: 'Projects & investments', file: 'feasibility-report.xlsx', href: '#feasibility', cta: 'Build this report',
    title: 'Feasibility Report', desc: 'A complete investment appraisal of a new project with a clear verdict and recommendations.',
    outputs: ['Verdict: Feasible, Feasible with conditions or Not feasible', 'Cash flows, NPV, IRR, payback and break-even', 'Sensitivity to revenue and cost changes'],
    needs: 'Project assumptions only; no statements needed', best: 'Business cases for lenders, investors or the board.' },
  { id: 'ratios', group: 'Performance', file: 'ratio-analysis.xlsx',
    outputs: ['Growth, margin, liquidity and leverage ratios', 'Return on assets and equity', 'Every year side by side'],
    needs: 'Income statement and balance sheet (cash flow optional)', best: 'Credit reviews and benchmarking.' },
  { id: 'forecast', group: 'Performance', file: '3-statement-forecast.xlsx',
    outputs: ['Projected revenue, EBITDA and net income', 'Cash and debt roll-forward', 'History vs forecast chart'],
    needs: 'Income statement, plus opening cash and debt', best: 'Budgets and business plans.' },
  { id: 'dashboard', group: 'Reporting', file: 'powerbi-dataset.xlsx', href: '#dashboard', cta: 'Open dashboard',
    title: 'Dashboard & Power BI Export', desc: 'An interactive, Power BI-style view of your statements, with a one-click dataset for Power BI Desktop.',
    outputs: ['KPI cards, trends, margins and a profit bridge', 'Year and metric filters, plus a full financial matrix', 'Power BI dataset: FactFinancials and Ratios tables'],
    needs: 'Income statement and balance sheet (cash flow optional)', best: 'Board packs and monthly management reporting.' },
];
const CATEGORIES = ['All models', ...new Set(CATALOG.map((c) => c.group))];
let currentCat = 'All models', examples = null;

function sampleState() {
  const p = parseRows(parseCSV(SAMPLE_CSV));
  return { company: SAMPLE_COMPANY, years: p.years, items: derive(p.items) };
}

// Runs every model on the sample company once, so each card shows real example figures
function computeExamples() {
  if (examples) return examples;
  const saved = state;
  state = sampleState();
  examples = {};
  try {
    const d = defaults();
    MODELS.forEach((m) => {
      const v = Object.fromEntries(m.inputs(d).map(([k, , val, t]) => [k, t === 'text' ? val : +val]));
      const tpl = document.createElement('template');
      tpl.innerHTML = m.run(v, d).html;
      let pairs = [...tpl.content.querySelectorAll('.kpi')].map((k) => [k.querySelector('.label').textContent, k.querySelector('.value').textContent]);
      if (!pairs.length) pairs = [...tpl.content.querySelectorAll('tbody tr')].map((tr) => [tr.cells[0].textContent, tr.cells[tr.cells.length - 1].textContent]).filter((p) => p[1] !== '—');
      examples[m.id] = pairs.slice(0, 3);
    });
    const fv = Object.fromEntries(FEAS.map(([k, , val]) => [k, val])), { flows } = feasFlows(fv);
    examples.feasibility = [['NPV', fmt(npv(fv.rate / 100, flows))], ['IRR', pct(irr(flows))], ['Payback', yrs(payback(flows))]];
    examples.dashboard = [[`Revenue ${state.years.at(-1)}`, fmt(L('revenue'))], ['EBITDA margin', pct(L('ebitda') / L('revenue'))], ['Revenue CAGR', pct(cagr(S('revenue')))]];
  } finally {
    state = saved;
  }
  return examples;
}

function cardHTML(c, ex) {
  const m = MODELS.find((x) => x.id === c.id), href = c.href || `#model/${c.id}`;
  return `<article class="model-card">
    <div class="mc-top"><span class="mc-cat">${c.group}</span><span class="mc-fmt">Excel export</span></div>
    <h3><a href="${href}">${c.title || m.name}</a></h3>
    <p class="mc-desc">${c.desc || m.desc}</p>
    <div class="mc-example">
      <div class="mc-ex-head">Example output · Aurora Retail Co.</div>
      <dl>${(ex || []).map(([k, v]) => `<div><dt>${esc(k)}</dt><dd>${esc(v)}</dd></div>`).join('')}</dl>
    </div>
    <div class="mc-section"><h4>Key outputs</h4><ul>${c.outputs.map((o) => `<li>${o}</li>`).join('')}</ul></div>
    <div class="mc-section"><h4>Data needed</h4><p>${c.needs}</p></div>
    <p class="mc-best"><strong>Best for:</strong> ${c.best}</p>
    <div class="mc-foot">
      <a class="btn" href="${href}">${c.cta || 'Build this model'} <span aria-hidden="true">→</span></a>
      <a class="link" href="samples/models/${c.file}" download>Sample .xlsx</a>
    </div>
  </article>`;
}

function renderLibrary() {
  const ex = computeExamples();
  $('#catTabs').innerHTML = CATEGORIES.map((c) => `<button data-cat="${c}" class="${c === currentCat ? 'on' : ''}">${c}<span>${c === 'All models' ? CATALOG.length : CATALOG.filter((x) => x.group === c).length}</span></button>`).join('');
  $('#catalog').innerHTML = CATALOG.filter((c) => currentCat === 'All models' || c.group === currentCat).map((c) => cardHTML(c, ex[c.id])).join('');
}

/* ---------- Model workspace ---------- */
let manualMode = false;

function renderWorkspace(id) {
  const m = MODELS.find((x) => x.id === id) || MODELS[0];
  currentModel = m.id;
  $('#wsCrumb').textContent = m.name;
  $('#wsGroup').textContent = m.group;
  $('#wsTitle').textContent = m.name;
  $('#wsDesc').textContent = m.desc;
  const groups = [...new Set(MODELS.map((x) => x.group))];
  $('#modelSelect').innerHTML = groups.map((g) => `<optgroup label="${g}">${MODELS.filter((x) => x.group === g)
    .map((x) => `<option value="${x.id}"${x.id === m.id ? ' selected' : ''}>${x.name}</option>`).join('')}</optgroup>`).join('');

  const has = hasData();
  const covered = [['Income statement', ['revenue', 'netIncome', 'ebitda']], ['Balance sheet', ['totalAssets', 'equity', 'debt', 'cash']], ['Cash flow', ['ocf', 'capex']]]
    .map(([n, ks]) => `<span class="${ks.some((k) => state.items[k]) ? 'ok' : 'miss'}">${ks.some((k) => state.items[k]) ? '✓' : '–'} ${n}</span>`).join('');
  $('#wsData').innerHTML = has
    ? `<div class="data-summary">
        <div class="ds-check" aria-hidden="true">✓</div>
        <div class="grow">
          <strong>${esc(state.company)}</strong>
          <div class="muted small">${esc(state.years[0])}–${esc(state.years.at(-1))} · ${Object.keys(state.items).length} line items</div>
          <div class="coverage">${covered}</div>
        </div>
        <a class="btn ghost" href="#upload">View data</a>
        <button class="btn ghost" data-action="change-data">Upload different statements</button>
      </div>
      <div class="uploader-slot" hidden></div>`
    : `<div class="uploader-slot"></div>
      <p class="manual small muted">No statements to hand? <button class="link" data-action="manual">Continue with manual assumptions →</button></p>`;

  const show = has || manualMode;
  $('#wsModelStep').hidden = !show;
  $$('#stepper li').forEach((li) => {
    const s = +li.dataset.step;
    li.className = s === 1 ? (has ? 'done' : 'active') : show ? 'active' : '';
  });
  if (show) renderModel();
}

/* ---------- Routing & init ---------- */
function route() {
  const [view, arg] = (location.hash || '#library').slice(1).split('/');
  const v = view === 'model' ? 'workspace' : ['library', 'dashboard', 'upload', 'feasibility'].includes(view) ? view : 'library';
  $$('.view').forEach((el) => el.classList.toggle('active', el.id === `view-${v}`));
  $$('nav a').forEach((a) => a.classList.toggle('active', a.dataset.view === (v === 'workspace' ? 'library' : v)));
  $('#companyPill').textContent = hasData() ? state.company : 'No data loaded';
  if (v === 'library') renderLibrary();
  if (v === 'workspace') renderWorkspace(arg);
  if (v === 'dashboard') renderDashboard();
  if (v === 'upload') renderPreview();
  renderUploaders();
}
function refreshAll() {
  route();
  // After an upload inside a model, jump straight to the results
  if ($('#view-workspace').classList.contains('active') && hasData()) $('#wsModelStep').scrollIntoView({ behavior: 'smooth' });
}

function init() {
  themeCharts();
  load();

  // Uploaders appear in several places, so their events are delegated
  document.addEventListener('change', (e) => {
    if (e.target.matches('.file-in')) { handleFile(e.target.files[0], e.target.closest('.uploader')); e.target.value = ''; }
  });
  ['dragover', 'drop'].forEach((t) => window.addEventListener(t, (e) => e.preventDefault()));
  document.addEventListener('dragover', (e) => e.target.closest?.('.drop')?.classList.add('over'));
  document.addEventListener('dragleave', (e) => e.target.closest?.('.drop')?.classList.remove('over'));
  document.addEventListener('drop', (e) => {
    const d = e.target.closest?.('.drop');
    if (!d) return;
    d.classList.remove('over');
    handleFile(e.dataTransfer.files[0], d.closest('.uploader'));
  });
  document.addEventListener('click', (e) => {
    const a = e.target.closest('[data-action]');
    if (!a) return;
    if (a.dataset.action === 'sample') loadSample();
    if (a.dataset.action === 'manual') { manualMode = true; route(); }
    if (a.dataset.action === 'change-data') { a.closest('.data-summary').hidden = true; $('#wsData .uploader-slot').hidden = false; }
  });

  $('#clearData').addEventListener('click', () => {
    state = { company: '', years: [], items: {} };
    uploadNote = '';
    manualMode = false;
    try { localStorage.removeItem(STORE); } catch (e) {}
    route();
  });

  $('#catTabs').addEventListener('click', (e) => {
    const b = e.target.closest('[data-cat]');
    if (b) { currentCat = b.dataset.cat; renderLibrary(); }
  });
  $('#modelSelect').addEventListener('change', (e) => { location.hash = `#model/${e.target.value}`; });
  $('#dlModel').addEventListener('click', downloadModel);

  ['slFrom', 'slTo', 'slMetric'].forEach((id) => $('#' + id).addEventListener('change', () => {
    dash.from = +$('#slFrom').value;
    dash.to = +$('#slTo').value;
    dash.metric = $('#slMetric').value;
    if (dash.from > dash.to) [dash.from, dash.to] = [dash.to, dash.from];
    renderDashboard();
  }));
  $('#slReset').addEventListener('click', () => { dash = { from: 0, to: null, metric: 'revenue' }; renderDashboard(); });
  $('#dlPowerBI').addEventListener('click', exportPowerBI);

  $('#feasForm').innerHTML = FEAS.map(([k, label, val, t]) =>
    `<div class="field"><label for="f-${k}">${label}</label><input id="f-${k}" ${t === 'text' ? 'type="text"' : 'type="number" step="any"'} value="${esc(val)}"></div>`).join('');
  $('#genReport').addEventListener('click', generateReport);
  $('#printReport').addEventListener('click', () => window.print());
  $('#dlReport').addEventListener('click', downloadReport);

  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { themeCharts(); route(); });
  window.addEventListener('hashchange', () => { route(); window.scrollTo(0, 0); });
  route();
}

document.addEventListener('DOMContentLoaded', init);
