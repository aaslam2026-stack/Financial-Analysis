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
      ['EBIT margin', pct(div(g('ebit'), g('revenue')))],
      ['Net margin', pct(div(g('netIncome'), g('revenue')))],
      ['Current ratio', xf(div(g('currentAssets'), g('currentLiabilities')))],
      ['Debt / Equity', xf(div(g('debt'), g('equity')))],
      ['Debt / EBITDA', xf(div(g('debt'), g('ebitda')))],
      ['Net debt / EBITDA', xf(g('debt') != null ? div(g('debt') - (g('cash') || 0), g('ebitda')) : null)],
      ['Interest coverage', xf(div(g('ebit'), g('interest')))],
      ['Return on assets', pct(div(g('netIncome'), g('totalAssets')))],
      ['Return on equity', pct(div(g('netIncome'), g('equity')))],
      ['Asset turnover', xf(div(g('revenue'), g('totalAssets')))],
      ['Free cash flow margin', pct(g('ocf') != null && g('capex') != null ? div(g('ocf') - Math.abs(g('capex')), g('revenue')) : null)],
      ['Cash conversion (OCF / NI)', xf(div(g('ocf'), g('netIncome')))],
    ];
  });
}

/* ---------- Excel & Power BI export ---------- */
const slug = (s) => String(s || 'model').replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase();
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
  const dr = deriveDrivers(), op = openingBS(), rev = op.rev;
  const p1 = (x) => +(x * 100).toFixed(1);
  const ebitda = L('ebitda') || rev * (dr.gm - dr.opex);
  const ly = state.years.at(-1);
  return {
    rev, ebitda, dr, op,
    g: p1(dr.growth), m: p1(dr.gm - dr.opex), da: p1(dr.da), capex: p1(dr.capex), nwc: p1(dr.nca - dr.cl), tax: p1(dr.tax),
    gm: p1(dr.gm), opexPct: p1(dr.opex), nca: p1(dr.nca), cl: p1(dr.cl), interest: p1(dr.rate), repay: dr.repay,
    netDebt: op.debt - op.cash, nwc0: op.nwc,
    ni: L('netIncome') || ebitda * 0.5,
    equity: op.equity, debt: op.debt, cash: op.cash, loanAmt: op.debt || Math.round(rev * 0.3),
    histLabel: ly ? `${String(ly).match(/(19|20)\d{2}/)?.[0] || ly}A` : 'Base',
    rf: marketRf()?.value ?? 4.2,
    bench: chosenBenchmark(),
  };
}

// Map model inputs (percent numbers) back to forecast drivers (fractions) for the 3-statement engine and Excel
const driversFromForecast = (v) => ({ growth: v.g / 100, gm: v.gm / 100, opex: v.opex / 100, da: v.da / 100, capex: v.capex / 100, tax: v.tax / 100,
  rate: v.rate / 100, nca: v.nca / 100, cl: v.cl / 100, repay: v.repay, payout: v.payout / 100 });
const driversFromDCF = (v, base) => ({ ...base, growth: v.g / 100, opex: base.gm - v.m / 100, da: v.da / 100, capex: v.capex / 100, tax: v.tax / 100, nca: v.nwc / 100 + base.cl });

// Unlevered DCF. Working capital is a % of revenue; the first year's change is measured against the actual
// opening working capital (nwc0) when it is known.
function dcfCore(v) {
  const g = v.g / 100, m = v.m / 100, w = v.wacc / 100, tg = v.tg / 100, t = v.tax / 100, nwcP = v.nwc / 100, n = Math.min(10, Math.max(1, Math.round(v.years)));
  let rev = v.rev, nwcPrev = v.nwc0 ?? v.rev * nwcP, pv = 0, fcf = 0;
  const rows = [];
  for (let y = 1; y <= n; y++) {
    rev *= 1 + g;
    const ebitda = rev * m, da = rev * (v.da / 100), ebit = ebitda - da, tax = ebit * t, nopat = ebit - tax;
    const capex = rev * (v.capex / 100), nwc = rev * nwcP, dnwc = nwc - nwcPrev;
    nwcPrev = nwc;
    fcf = nopat + da - capex - dnwc;
    const df = 1 / Math.pow(1 + w, y), p = fcf * df;
    pv += p;
    rows.push({ y, rev, ebitda, da, ebit, tax, nopat, capex, dnwc, fcf, df, pv: p });
  }
  if (w <= tg) return { error: 'WACC must be higher than terminal growth.' };
  const tv = (fcf * (1 + tg)) / (w - tg), pvTv = tv / Math.pow(1 + w, n), ev = pv + pvTv;
  return { rows, pv, tv, pvTv, ev, equity: ev - v.netDebt, perShare: (ev - v.netDebt) / v.shares, exitMult: tv / rows.at(-1).ebitda };
}

// Shared by the Monte Carlo model and its Excel sheet
function monteCarlo(v, d) {
  const n = Math.min(Math.max(100, v.sims), 20000), out = [];
  for (let i = 0; i < n; i++) {
    const r = dcfCore({ rev: v.rev, g: v.g + randn() * v.gsd, m: v.m + randn() * v.msd, da: d.da, capex: d.capex, nwc: d.nwc, nwc0: d.nwc0, tax: d.tax, wacc: v.wacc, tg: v.tg, years: 5, netDebt: v.netDebt, shares: 1 });
    if (!r.error) out.push(r.ev);
  }
  if (!out.length) return null;
  out.sort((a, b) => a - b);
  const q = (p) => out[Math.floor(p * (out.length - 1))];
  const bins = 24, lo = q(0.01), hi = q(0.99), w = (hi - lo) / bins || 1, hist = Array(bins).fill(0);
  out.forEach((x) => { const b = Math.floor((x - lo) / w); if (b >= 0 && b < bins) hist[b]++; });
  const base = dcfCore({ ...v, da: d.da, capex: d.capex, nwc: d.nwc, nwc0: d.nwc0, tax: d.tax, years: 5, shares: 1 }).ev;
  const mean = out.reduce((a, b) => a + b, 0) / out.length, sd = Math.sqrt(out.reduce((a, b) => a + (b - mean) ** 2, 0) / out.length);
  return { out, q, hist, lo, w, base, mean, sd };
}

const fyLabel = (t) => { const y = parseInt(String(state.years.at(-1) || '').match(/(19|20)\d{2}/)?.[0]); return y ? `${y + t}E` : `Y${t}`; };
// Line items as rows, years as columns: [label, values, formatter?, bold?]
const tableT = (cols, rows, hl) => table(['', ...cols], rows.map(([label, vals, f = fmt, bold]) => [bold ? `<strong>${label}</strong>` : label, ...vals.map((x) => (x == null ? '—' : f(x)))]), hl);
const pctPts = (x) => `${x.toFixed(1)}%`;

function lboCore(v) {
  const ev0 = v.ebitda * v.entry, fees = ev0 * (v.fees / 100), debt0 = v.ebitda * v.lev, eq0 = ev0 + fees - debt0;
  if (eq0 <= 0) return null;
  let e = v.ebitda, debt = debt0, cash = 0;
  const rows = [];
  for (let y = 1; y <= v.years; y++) {
    e *= 1 + v.g / 100;
    const int = debt * (v.rate / 100), capex = e * (v.capex / 100);
    const tax = Math.max(0, (e - int - capex) * (v.tax / 100));
    const fcf = e - int - capex - tax;
    const pay = Math.min(debt, Math.max(0, fcf * (v.sweep / 100)));
    debt -= pay; cash += fcf - pay;
    rows.push({ y, e, int, capex, tax, fcf, pay, debt, cash, eqv: e * v.exit - debt + cash, lev: debt / e, cov: int ? e / int : null });
  }
  const L1 = rows.at(-1), exitEq = L1.eqv, moic = exitEq / eq0;
  const irrAt = (entry, exit) => { const q = (L1.e * exit - L1.debt + L1.cash) / (v.ebitda * entry * (1 + v.fees / 100) - debt0); return q > 0 && v.ebitda * entry * (1 + v.fees / 100) - debt0 > 0 ? Math.pow(q, 1 / v.years) - 1 : null; };
  return { ev0, fees, debt0, eq0, rows, exitEq, moic, irr: Math.pow(Math.max(moic, 0), 1 / v.years) - 1, irrAt,
    attr: { growth: (L1.e - v.ebitda) * v.entry, multiple: L1.e * (v.exit - v.entry), delever: debt0 - L1.debt + L1.cash, fees: -fees } };
}

// Five illustrative peer multiples spread around an industry benchmark (or the default list)
const peerList = (m, fallback) => (m ? [0.8, 0.9, 1, 1.1, 1.2].map((k) => (m * k).toFixed(1)).join(', ') : fallback);

const MODELS = [
  {
    id: 'dcf', name: 'DCF Valuation', group: 'Valuation',
    desc: 'Discounted cash flow: projects unlevered free cash flow, then discounts it at WACC to find what the business is worth today.',
    inputs: (d) => [['rev', 'Base revenue ($)', d.rev], ['g', 'Revenue growth %', d.g], ['m', 'EBITDA margin %', d.m], ['da', 'D&A % of revenue', d.da],
      ['capex', 'CapEx % of revenue', d.capex], ['nwc', 'Net working capital % of revenue', d.nwc], ['tax', 'Tax rate %', d.tax], ['wacc', 'WACC %', 10],
      ['tg', 'Terminal growth %', 2.5], ['years', 'Forecast years', 5], ['netDebt', 'Net debt ($)', d.netDebt], ['shares', 'Shares outstanding', 1_000_000], ['price', 'Current share price ($, optional)', 0]],
    run(v, d) {
      v = { ...v, nwc0: d.nwc0 };
      const r = dcfCore(v);
      if (r.error) return { html: `<div class="glass card">${r.error}</div>` };
      const ws = [-2, -1, 0, 1, 2].map((x) => v.wacc + x), gs = [-1, -0.5, 0, 0.5, 1].map((x) => v.tg + x);
      const grid = (f) => ws.map((w) => [pctPts(w), ...gs.map((g) => { const s = dcfCore({ ...v, wacc: w, tg: g }); return s.error ? 'n/a' : f(s); })]);
      const scen = [['Bear', -3, -3], ['Base', 0, 0], ['Bull', 3, 3]].map(([n, dg, dm]) => {
        const s = dcfCore({ ...v, g: v.g + dg, m: v.m + dm });
        return [n, pctPts(v.g + dg), pctPts(v.m + dm), fmt(s.ev), fmt(s.equity), fmt(s.perShare)];
      });
      const up = v.price > 0 ? r.perShare / v.price - 1 : null, cols = r.rows.map((x) => fyLabel(x.y)), col = (k) => r.rows.map((x) => x[k]);
      const center = (ri, ci) => ri === 2 && ci === 3;
      return {
        html: `<div class="kpis">${kpi('Enterprise value', fmt(r.ev))}${kpi('Equity value', fmt(r.equity))}${kpi('Value per share', fmt(r.perShare), up == null ? 'Add a share price to see upside' : `${up >= 0 ? '▲' : '▼'} ${pct(Math.abs(up))} vs current price`, up == null ? '' : up >= 0 ? 'up' : 'down')}
          ${kpi('Implied EV / EBITDA', xf(d.ebitda ? r.ev / d.ebitda : null), 'on last actual EBITDA')}${kpi('Implied exit multiple', xf(r.exitMult), 'terminal value / final-year EBITDA')}${kpi('Terminal value share', pct(r.pvTv / r.ev), 'of enterprise value')}</div>
        <div class="glass card"><h3>Free cash flow projection</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Unlevered free cash flow build</h3>${tableT(cols, [['Revenue', col('rev')], ['EBITDA', col('ebitda')], ['Less: D&A', col('da')], ['EBIT', col('ebit'), fmt, true], ['Less: taxes on EBIT', col('tax')], ['NOPAT', col('nopat'), fmt, true], ['Add: D&A', col('da')], ['Less: CapEx', col('capex')], ['Less: increase in working capital', col('dnwc')], ['Unlevered free cash flow', col('fcf'), fmt, true], ['Discount factor', col('df'), (x) => x.toFixed(4)], ['Present value', col('pv'), fmt, true]])}</div>
        <div class="grid two">
          <div class="glass card"><h3>Valuation bridge</h3>${table(['Item', 'Value'], [['Sum of PV of free cash flows', fmt(r.pv)], ['PV of terminal value', fmt(r.pvTv)], ['<strong>Enterprise value</strong>', fmt(r.ev)], ['Less: net debt', fmt(v.netDebt)], ['<strong>Equity value</strong>', fmt(r.equity)], ['Shares outstanding', fmt(v.shares, false)], ['<strong>Value per share</strong>', fmt(r.perShare)], ['Terminal value (undiscounted)', fmt(r.tv)]])}</div>
          <div class="glass card"><h3>Scenarios</h3>${table(['Scenario', 'Growth', 'Margin', 'EV', 'Equity', 'Per share'], scen)}</div>
        </div>
        <div class="grid two">
          <div class="glass card"><h3>Enterprise value: WACC ↓ × terminal growth →</h3>${table(['WACC \\ g', ...gs.map(pctPts)], grid((s) => fmt(s.ev)), center)}</div>
          <div class="glass card"><h3>Value per share: WACC ↓ × terminal growth →</h3>${table(['WACC \\ g', ...gs.map(pctPts)], grid((s) => fmt(s.perShare)), center)}</div>
        </div>`,
        charts: [['mc1', { type: 'bar', data: { labels: cols, datasets: [ds('Unlevered FCF', col('fcf'), COLORS.blue), ds('Present value', col('pv'), COLORS.green)] }, options: { scales: { y: moneyAxis } } }]],
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
      let cum = 0, dcum = 0;
      const cumArr = flows.map((f) => (cum += f)), dArr = flows.map((f, t) => (dcum += f / Math.pow(1 + r, t)));
      const rates = [0, 5, 10, 15, 20, 25];
      return {
        html: `<div class="kpis">${kpi('NPV', fmt(n), n >= 0 ? 'Creates value ✓' : 'Destroys value ✗', n >= 0 ? 'up' : 'down')}${kpi('IRR', pct(i), i != null && i > r ? 'Above hurdle rate' : 'Below hurdle rate', i != null && i > r ? 'up' : 'down')}${kpi('MIRR', pct(mirr(flows, r, v.reinv / 100)))}${kpi('Payback', yrs(payback(flows)))}${kpi('Discounted payback', yrs(payback(flows, r)))}${kpi('Profitability index', xf(pv / Math.abs(v.inv)))}${kpi('Total net cash', fmt(cum), 'undiscounted')}</div>
        <div class="grid two">
          <div class="glass card"><h3>Cumulative cash flow</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
          <div class="glass card"><h3>NPV profile</h3><div class="chart"><canvas id="mc2"></canvas></div></div>
        </div>
        <div class="glass card"><h3>Cash flow schedule</h3>${tableT(flows.map((_, t) => `Year ${t}`), [['Cash flow', flows, fmt, true], ['Discount factor', flows.map((_, t) => 1 / Math.pow(1 + r, t)), (x) => x.toFixed(4)], ['Discounted cash flow', flows.map((f, t) => f / Math.pow(1 + r, t))], ['Cumulative cash flow', cumArr], ['Cumulative discounted', dArr]])}</div>
        <div class="glass card"><h3>NPV at different discount rates</h3>${table(['Discount rate', ...rates.map((x) => x + '%')], [['NPV', ...rates.map((x) => fmt(npv(x / 100, flows)))]])}</div>`,
        charts: [['mc1', { data: { labels: flows.map((_, t) => `Y${t}`), datasets: [ds('Cash flow', flows, COLORS.blue, { type: 'bar' }), ds('Cumulative', cumArr, COLORS.green, { type: 'line' })] }, options: { scales: { y: moneyAxis } } }],
          ['mc2', { type: 'line', data: { labels: Array.from({ length: 31 }, (_, k) => k + '%'), datasets: [ds('NPV', Array.from({ length: 31 }, (_, k) => npv(k / 100, flows)), COLORS.purple)] }, options: { scales: { y: moneyAxis }, plugins: { legend: { display: false } } } }]],
      };
    },
  },
  {
    id: 'lbo', name: 'LBO', group: 'Deals & financing',
    desc: 'Leveraged buyout: buy the company mostly with debt, pay the debt down with cash flow, then sell. Shows the investor return.',
    inputs: (d) => [['ebitda', 'Entry EBITDA ($)', Math.round(d.ebitda)], ['entry', 'Entry multiple (x EBITDA)', 8], ['lev', 'Debt (x EBITDA)', 4.5], ['rate', 'Interest rate %', 8],
      ['g', 'EBITDA growth %', Math.max(3, d.g)], ['exit', 'Exit multiple (x EBITDA)', 8], ['years', 'Holding period (years)', 5], ['tax', 'Tax rate %', d.tax],
      ['capex', 'CapEx % of EBITDA', 15], ['sweep', 'Cash used to repay debt %', 100], ['fees', 'Transaction fees % of EV', 2]],
    run(v) {
      const L = lboCore(v);
      if (!L) return { html: '<div class="glass card">Debt is larger than the purchase price. Lower the debt multiple.</div>' };
      const cols = ['Entry', ...L.rows.map((x) => fyLabel(x.y))], c = (k, first = null) => [first, ...L.rows.map((x) => x[k])];
      const entries = [-1, -0.5, 0, 0.5, 1].map((x) => v.entry + x), exits = [-2, -1, 0, 1, 2].map((x) => v.exit + x);
      const total = L.attr.growth + L.attr.multiple + L.attr.delever + L.attr.fees;
      return {
        html: `<div class="kpis">${kpi('Purchase price (EV)', fmt(L.ev0))}${kpi('Equity invested', fmt(L.eq0))}${kpi('Exit equity', fmt(L.exitEq))}${kpi('IRR', pct(L.irr), L.irr >= 0.2 ? 'Meets typical 20% PE target' : 'Below typical 20% PE target', L.irr >= 0.2 ? 'up' : 'down')}${kpi('MOIC', xf(L.moic))}${kpi('Debt repaid', fmt(L.debt0 - L.rows.at(-1).debt), `${pct((L.debt0 - L.rows.at(-1).debt) / L.debt0)} of entry debt`)}</div>
        <div class="grid two">
          <div class="glass card"><h3>Sources &amp; uses</h3>${table(['Uses', 'Amount', 'Sources', 'Amount'], [['Purchase price', fmt(L.ev0), 'Debt', fmt(L.debt0)], ['Transaction fees', fmt(L.fees), 'Sponsor equity', fmt(L.eq0)], ['<strong>Total uses</strong>', fmt(L.ev0 + L.fees), '<strong>Total sources</strong>', fmt(L.debt0 + L.eq0)]])}</div>
          <div class="glass card"><h3>Returns attribution</h3>${table(['Value driver', 'Amount', 'Share'], [['EBITDA growth', fmt(L.attr.growth), pct(L.attr.growth / total)], ['Multiple expansion', fmt(L.attr.multiple), pct(L.attr.multiple / total)], ['Debt paydown and cash', fmt(L.attr.delever), pct(L.attr.delever / total)], ['Transaction fees', fmt(L.attr.fees), pct(L.attr.fees / total)], ['<strong>Equity value created</strong>', fmt(total), '100.0%']])}</div>
        </div>
        <div class="glass card"><h3>Debt paydown vs equity value</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Debt schedule and credit statistics</h3>${tableT(cols, [['EBITDA', c('e', v.ebitda)], ['Interest', c('int')], ['CapEx', c('capex')], ['Taxes', c('tax')], ['Free cash flow', c('fcf'), fmt, true], ['Debt repaid', c('pay')], ['Debt balance', c('debt', L.debt0), fmt, true], ['Cash balance', c('cash', 0)], ['Debt / EBITDA', c('lev', v.lev), xf], ['EBITDA / interest', c('cov'), xf], ['Equity value at exit multiple', c('eqv', L.eq0), fmt, true]])}</div>
        <div class="glass card"><h3>IRR sensitivity: entry multiple ↓ × exit multiple →</h3>${table(['Entry \\ Exit', ...exits.map((x) => x.toFixed(1) + 'x')], entries.map((en) => [en.toFixed(1) + 'x', ...exits.map((ex) => pct(L.irrAt(en, ex)))]), (ri, ci) => ri === 2 && ci === 3)}</div>`,
        charts: [['mc1', { type: 'bar', data: { labels: cols, datasets: [ds('Debt', [L.debt0, ...L.rows.map((x) => x.debt)], COLORS.orange, { stack: 's' }), ds('Equity value', [L.eq0, ...L.rows.map((x) => x.eqv)], COLORS.green, { stack: 's' })] }, options: { scales: { y: { ...moneyAxis, stacked: true }, x: { stacked: true } } } }]],
      };
    },
  },
  {
    id: 'wacc', name: 'WACC', group: 'Valuation',
    desc: 'Weighted average cost of capital: the blended return that lenders and shareholders expect. Use it as the discount rate in the DCF.',
    inputs: (d) => [['E', 'Equity value ($)', Math.round(d.equity)], ['D', 'Debt value ($)', Math.round(d.debt)], ['rf', 'Risk-free rate %', +d.rf.toFixed(2)], ['beta', 'Levered beta', d.bench?.beta ?? 1.1],
      ['mrp', 'Market risk premium %', 5.5], ['sp', 'Size / specific premium %', 1], ['kd', 'Pre-tax cost of debt %', d.interest], ['tax', 'Tax rate %', d.tax]],
    run(v) {
      const V = v.E + v.D, kd = (v.kd / 100) * (1 - v.tax / 100);
      const wacc = (beta, mrp) => (v.E / V) * ((v.rf + beta * mrp + v.sp) / 100) + (v.D / V) * kd;
      const ke = (v.rf + v.beta * v.mrp + v.sp) / 100, w = wacc(v.beta, v.mrp), bu = v.E ? v.beta / (1 + (1 - v.tax / 100) * (v.D / v.E)) : null;
      const betas = [-0.2, -0.1, 0, 0.1, 0.2].map((x) => v.beta + x), mrps = [-1, -0.5, 0, 0.5, 1].map((x) => v.mrp + x);
      return {
        html: `<div class="kpis">${kpi('WACC', pct(w, 2))}${kpi('Cost of equity (CAPM)', pct(ke, 2))}${kpi('After-tax cost of debt', pct(kd, 2))}${kpi('Debt weight (D / V)', pct(v.D / V))}${kpi('Unlevered beta', bu == null ? '—' : bu.toFixed(2), 'asset beta, Hamada formula')}${kpi('Debt / equity', xf(v.E ? v.D / v.E : null))}</div>
        <div class="grid two">
          <div class="glass card"><h3>Capital structure</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
          <div class="glass card"><h3>Build-up</h3>${table(['Component', 'Value'], [['Risk-free rate', pctPts(v.rf)], ['Beta × market risk premium', pctPts(v.beta * v.mrp)], ['Size / specific premium', pctPts(v.sp)], ['<strong>Cost of equity</strong>', pct(ke, 2)], ['Pre-tax cost of debt', pctPts(v.kd)], ['Tax shield', pct(-(v.kd / 100) * (v.tax / 100), 2)], ['<strong>After-tax cost of debt</strong>', pct(kd, 2)], ['<strong>WACC</strong>', pct(w, 2)]])}</div>
        </div>
        <div class="glass card"><h3>WACC sensitivity: beta ↓ × market risk premium →</h3>${table(['Beta \\ MRP', ...mrps.map(pctPts)], betas.map((b) => [b.toFixed(2), ...mrps.map((m) => pct(wacc(b, m), 2))]), (ri, ci) => ri === 2 && ci === 3)}</div>`,
        charts: [['mc1', { type: 'doughnut', data: { labels: ['Equity', 'Debt'], datasets: [{ data: [v.E, v.D], backgroundColor: [COLORS.green, COLORS.orange], borderWidth: 0 }] }, options: { cutout: '65%' } }]],
      };
    },
  },
  {
    id: 'comps', name: 'Comparables', group: 'Valuation',
    desc: 'Values the company using the multiples that similar listed companies trade at. Enter peer multiples separated by commas.',
    inputs: (d) => [['rev', 'Company revenue ($)', Math.round(d.rev)], ['ebitda', 'Company EBITDA ($)', Math.round(d.ebitda)], ['ni', 'Company net income ($)', Math.round(d.ni)],
      ['netDebt', 'Net debt ($)', Math.round(d.netDebt)], ['evE', 'Peer EV/EBITDA', peerList(d.bench?.evEbitda, '7.5, 8.2, 9.0, 10.1, 11.4'), 'text'], ['evS', 'Peer EV/Sales', peerList(d.bench?.evSales, '1.1, 1.4, 1.6, 1.9'), 'text'], ['pe', 'Peer P/E', peerList(d.bench?.pe, '14, 16.5, 18, 21'), 'text']],
    run(v) {
      const mean = (a) => a.reduce((s, x) => s + x, 0) / a.length;
      const rows = [['EV / EBITDA', list(v.evE), v.ebitda, false], ['EV / Sales', list(v.evS), v.rev, false], ['P / E', list(v.pe), v.ni, true]]
        .filter((r) => r[1].length)
        .map(([n, m, base, isEq]) => {
          const toEq = (x) => (isEq ? x * base : x * base - v.netDebt);
          return { n, m: [Math.min(...m), median(m), mean(m), Math.max(...m)], eq: [Math.min(...m), median(m), mean(m), Math.max(...m)].map(toEq), count: m.length };
        });
      const mid = rows.length ? median(rows.map((r) => r.eq[1])) : null, avg = rows.length ? mean(rows.map((r) => r.eq[2])) : null;
      return {
        html: `<div class="kpis">${kpi('Implied equity value (median)', fmt(mid))}${kpi('Implied equity value (mean)', fmt(avg))}${rows.map((r) => kpi(r.n + ' median', xf(r.m[1]), `${r.count} peers`)).join('')}</div>
        <div class="glass card"><h3>Valuation range (football field)</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="grid two">
          <div class="glass card"><h3>Peer multiples</h3>${table(['Method', 'Low', 'Median', 'Mean', 'High'], rows.map((r) => [r.n, ...r.m.map(xf)]))}</div>
          <div class="glass card"><h3>Implied equity value</h3>${table(['Method', 'Low', 'Median', 'Mean', 'High'], rows.map((r) => [r.n, ...r.eq.map((x) => fmt(x))]))}</div>
        </div>`,
        charts: [['mc1', { type: 'bar', data: { labels: rows.map((r) => r.n), datasets: [ds('Equity value range', rows.map((r) => [r.eq[0], r.eq[3]]), COLORS.purple)] }, options: { indexAxis: 'y', scales: { x: moneyAxis }, plugins: { legend: { display: false } } } }]],
      };
    },
  },
  {
    id: 'montecarlo', name: 'Monte Carlo', group: 'Valuation',
    desc: 'Runs the DCF many times with random growth and margins to show the range of likely values, not just one number.',
    inputs: (d) => [['rev', 'Base revenue ($)', Math.round(d.rev)], ['g', 'Growth mean %', d.g], ['gsd', 'Growth std. dev. %', 3], ['m', 'Margin mean %', d.m], ['msd', 'Margin std. dev. %', 3],
      ['wacc', 'WACC %', 10], ['tg', 'Terminal growth %', 2.5], ['netDebt', 'Net debt ($)', Math.round(d.netDebt)], ['sims', 'Simulations', 2000]],
    run(v, d) {
      const mc = monteCarlo(v, d);
      if (!mc) return { html: '<div class="glass card">WACC must be higher than terminal growth.</div>' };
      const { out, q, hist, lo, w, base, mean, sd } = mc;
      return {
        html: `<div class="kpis">${kpi('P10 (downside)', fmt(q(0.1)))}${kpi('P50 (median)', fmt(q(0.5)))}${kpi('P90 (upside)', fmt(q(0.9)))}${kpi('Mean', fmt(mean), `std. dev. ${fmt(sd)}`)}${kpi('Chance above base case', pct(out.filter((x) => x > base).length / out.length, 0), `Base EV ${fmt(base)}`)}</div>
        <div class="glass card"><h3>Distribution of enterprise value (${out.length} runs)</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Percentiles of enterprise value</h3>${table(['Statistic', 'P5', 'P10', 'P25', 'P50', 'P75', 'P90', 'P95'], [['Enterprise value', ...[0.05, 0.1, 0.25, 0.5, 0.75, 0.9, 0.95].map((p) => fmt(q(p)))], ['Equity value', ...[0.05, 0.1, 0.25, 0.5, 0.75, 0.9, 0.95].map((p) => fmt(q(p) - v.netDebt))]])}</div>`,
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
      const vols = [0.5, 0.75, 1, 1.25, 1.5].map((k) => v.units * k);
      return {
        html: `<div class="kpis">${kpi('Break-even units', Math.ceil(be).toLocaleString())}${kpi('Break-even revenue', fmt(be * v.price))}${kpi('Expected profit', fmt(profit), '', profit >= 0 ? 'up' : 'down')}${kpi('Margin of safety', pct(mos), '', mos >= 0 ? 'up' : 'down')}${kpi('Contribution margin', pct(cm / v.price), `${fmt(cm)} per unit`)}${kpi('Operating leverage', profit > 0 ? xf((v.units * cm) / profit) : '—')}</div>
        <div class="glass card"><h3>Revenue vs total cost</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Profit at different sales volumes</h3>${tableT(['50%', '75%', 'Plan', '125%', '150%'], [['Units sold', vols, (x) => Math.round(x).toLocaleString()], ['Revenue', vols.map((u) => u * v.price)], ['Variable costs', vols.map((u) => u * v.vc)], ['Fixed costs', vols.map(() => v.fixed)], ['Profit', vols.map((u) => u * cm - v.fixed), fmt, true], ['Profit margin', vols.map((u) => (u * cm - v.fixed) / (u * v.price)), pct]], (ri, ci) => ci === 3)}</div>`,
        charts: [['mc1', { type: 'line', data: { labels: pts.map((u) => Math.round(u).toLocaleString()), datasets: [ds('Revenue', pts.map((u) => u * v.price), COLORS.green), ds('Total cost', pts.map((u) => v.fixed + u * v.vc), COLORS.red), ds('Fixed cost', pts.map(() => v.fixed), COLORS.orange, { borderDash: [6, 4] })] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
  {
    id: 'ratios', name: 'Ratio Analysis', group: 'Performance',
    desc: 'Profitability, liquidity, leverage, efficiency and return ratios calculated from your uploaded statements for every year.',
    inputs: () => [],
    run() {
      if (!hasData()) return { html: '<div class="glass card empty"><h3>Upload statements to see ratios</h3><a class="btn" href="#upload">Upload data</a></div>' };
      const all = ratios(), names = all[0].map((r) => r[0]);
      return { html: `<div class="glass card">${table(['Ratio', ...state.years.map(esc)], names.map((n, i) => [n, ...all.map((y) => y[i][1])]))}</div>` };
    },
  },
  {
    id: 'forecast', name: '3-Statement Model', group: 'Performance',
    desc: 'Linked income statement, balance sheet and cash flow forecast. Drivers start from your 3-year historical averages, and the balance sheet always balances.',
    inputs: (d) => [['years', 'Forecast years', 5], ['g', 'Revenue growth %', d.g], ['gm', 'Gross margin %', d.gm], ['opex', 'Operating expenses % of revenue', d.opexPct], ['da', 'D&A % of revenue', d.da],
      ['capex', 'CapEx % of revenue', d.capex], ['tax', 'Tax rate %', d.tax], ['rate', 'Interest rate on debt %', d.interest], ['nca', 'Non-cash current assets % of revenue', d.nca],
      ['cl', 'Current liabilities % of revenue', d.cl], ['repay', 'Debt repaid per year ($)', d.repay], ['payout', 'Dividend payout % of net income', 0]],
    run(v, d) {
      const n = Math.min(10, Math.max(1, Math.round(v.years))), Y = project3(driversFromForecast(v), d.op, n);
      const cols = [d.histLabel, ...Y.map((_, t) => fyLabel(t + 1))], H = hasData() ? histIS().at(-1) : {}, op = d.op;
      const c = (k, first = null) => [first, ...Y.map((y) => y[k])];
      const bal = Math.max(...Y.map((y) => Math.abs(y.check)));
      const hy = hasData() ? state.years : [], hv = (k) => (hasData() ? S(k) : []);
      return {
        html: `<div class="kpis">${kpi(`Revenue ${cols.at(-1)}`, fmt(Y.at(-1).rev), `${pct(Math.pow(Y.at(-1).rev / op.rev, 1 / n) - 1)} CAGR`)}${kpi('EBITDA', fmt(Y.at(-1).ebitda), `${pct(Y.at(-1).ebitda / Y.at(-1).rev)} margin`)}${kpi('Net income', fmt(Y.at(-1).ni), `${pct(Y.at(-1).ni / Y.at(-1).rev)} margin`)}${kpi('Closing cash', fmt(Y.at(-1).cash))}${kpi('Closing debt', fmt(Y.at(-1).debt))}${kpi('Balance check', bal < 1 ? 'Balanced ✓' : fmt(bal), 'assets = liabilities + equity', bal < 1 ? 'up' : 'down')}</div>
        <div class="grid two">
          <div class="glass card"><h3>Revenue and profit</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
          <div class="glass card"><h3>Cash and debt</h3><div class="chart"><canvas id="mc2"></canvas></div></div>
        </div>
        <div class="glass card"><h3>Income statement</h3>${tableT(cols, [['Revenue', c('rev', H.rev)], ['Cost of goods sold', c('cogs', H.cogs)], ['Gross profit', c('gp', H.gp), fmt, true], ['Operating expenses', c('opex', H.opex)], ['EBITDA', c('ebitda', H.ebitda), fmt, true], ['Depreciation & amortization', c('da', H.da)], ['EBIT', c('ebit', H.ebit), fmt, true], ['Interest expense', c('int', H.int)], ['Profit before tax', c('pbt', H.pbt)], ['Taxes', c('tax', H.tax)], ['Net income', c('ni', H.ni), fmt, true]])}</div>
        <div class="glass card"><h3>Balance sheet</h3>${tableT(cols, [['Cash', c('cash', op.cash)], ['Non-cash current assets', c('noncash', op.nca)], ['Total current assets', c('tca', op.ca), fmt, true], ['Non-current assets', c('ppe', op.ppe)], ['Total assets', c('ta', op.ta), fmt, true], ['Current liabilities', c('cl', op.cl)], ['Total debt', c('debt', op.debt)], ['Other liabilities', c('other', op.other)], ['Total liabilities', c('tl', op.cl + op.debt + op.other), fmt, true], ["Shareholders' equity", c('eq', op.equity)], ['Total liabilities & equity', c('tle', op.ta), fmt, true], ['Balance check', c('check', 0)]])}</div>
        <div class="glass card"><h3>Cash flow statement</h3>${tableT(cols, [['Net income', c('ni')], ['Depreciation & amortization', c('da')], ['Change in non-cash current assets', c('dnca')], ['Change in current liabilities', c('dcl')], ['Cash from operations', c('cfo', hasData() ? L('ocf') : null), fmt, true], ['Capital expenditure', c('capex', hasData() && L('capex') != null ? -L('capex') : null)], ['Cash from investing', c('cfi'), fmt, true], ['Debt repayment', c('repay')], ['Dividends', c('div')], ['Cash from financing', c('cff'), fmt, true], ['Net change in cash', c('net')], ['Closing cash', c('cash', op.cash), fmt, true], ['Free cash flow', c('fcf'), fmt, true]])}</div>`,
        charts: [['mc1', { data: { labels: [...hy, ...Y.map((_, t) => fyLabel(t + 1))], datasets: [ds('Revenue', [...hv('revenue'), ...Y.map((y) => y.rev)], COLORS.blue, { type: 'bar' }), ds('EBITDA', [...hv('ebitda'), ...Y.map((y) => y.ebitda)], COLORS.green, { type: 'bar' }), ds('Net income', [...hv('netIncome'), ...Y.map((y) => y.ni)], COLORS.purple, { type: 'line' })] }, options: { scales: { y: moneyAxis } } }],
          ['mc2', { type: 'line', data: { labels: cols, datasets: [ds('Cash', c('cash', op.cash), COLORS.green), ds('Debt', c('debt', op.debt), COLORS.orange)] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
  {
    id: 'loan', name: 'Loan Schedule', group: 'Deals & financing',
    desc: 'Amortizing loan: payment per period, total interest, effective rate and the balance over time.',
    inputs: (d) => [['P', 'Loan amount ($)', Math.round(d.loanAmt)], ['rate', 'Annual interest rate %', d.interest], ['years', 'Term (years)', 7], ['ppy', 'Payments per year', 12]],
    run(v) {
      const n = v.years * v.ppy, r = v.rate / 100 / v.ppy, pmt = r ? (v.P * r) / (1 - Math.pow(1 + r, -n)) : v.P / n;
      let bal = v.P;
      const rows = [];
      for (let y = 1; y <= v.years; y++) {
        const open = bal;
        let int = 0, prin = 0;
        for (let k = 0; k < v.ppy; k++) { const i = bal * r; int += i; prin += pmt - i; bal -= pmt - i; }
        rows.push([`Y${y}`, open, pmt * v.ppy, int, prin, Math.max(0, bal)]);
      }
      return {
        html: `<div class="kpis">${kpi('Payment per period', fmt(pmt), `${v.ppy}× per year`)}${kpi('Annual debt service', fmt(pmt * v.ppy))}${kpi('Total interest', fmt(pmt * n - v.P), `${pct((pmt * n - v.P) / v.P)} of the loan`)}${kpi('Total paid', fmt(pmt * n))}${kpi('Effective annual rate', pct(Math.pow(1 + r, v.ppy) - 1, 2), `${n} payments`)}</div>
        <div class="glass card"><h3>Balance and interest</h3><div class="chart"><canvas id="mc1"></canvas></div></div>
        <div class="glass card"><h3>Annual schedule</h3>${table(['Year', 'Opening balance', 'Paid', 'Interest', 'Principal', 'Closing balance'], rows.map((x) => [x[0], ...x.slice(1).map((y) => fmt(y))]))}</div>`,
        charts: [['mc1', { data: { labels: rows.map((x) => x[0]), datasets: [ds('Principal', rows.map((x) => x[4]), COLORS.blue, { type: 'bar', stack: 'p' }), ds('Interest', rows.map((x) => x[3]), COLORS.orange, { type: 'bar', stack: 'p' }), ds('Balance', rows.map((x) => x[5]), COLORS.green, { type: 'line' })] }, options: { scales: { y: moneyAxis } } }]],
      };
    },
  },
];

let currentModel = 'dcf';

async function downloadModel() {
  const v = {};
  $$('#modelArea input').forEach((i) => { v[i.dataset.k] = i.type === 'text' ? i.value : parseFloat(i.value) || 0; });
  try {
    await downloadModelWorkbook([currentModel], `${slug(hasData() ? state.company : 'model')}-${currentModel}.xlsx`, { [currentModel]: v });
  } catch (e) { alert(e.message || e); }
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

function feasInputs() {
  const v = {};
  FEAS.forEach(([k, , , t]) => { const val = $(`#f-${k}`).value; v[k] = t === 'text' ? val.trim() : parseFloat(val) || 0; });
  v.life = Math.min(40, Math.max(1, Math.round(v.life)));
  return v;
}

// Shared by the on-screen report and the Excel workbook
function feasAnalysis(v) {
  const r = v.rate / 100, { flows, rows } = feasFlows(v);
  const n = npv(r, flows), i = irr(flows), pb = payback(flows), dpb = payback(flows, r), pi = npv(r, [0, ...flows.slice(1)]) / v.inv;
  const beRev = v.fixed / (1 - v.var / 100);
  const sensRaw = [['Revenue −20%', 0.8, 1], ['Revenue −10%', 0.9, 1], ['Base case', 1, 1], ['Revenue +10%', 1.1, 1], ['Costs +10%', 1, 1.1]].map(([lab, rm, cm]) => {
    const f = feasFlows(v, rm, cm).flows; return { lab, npv: npv(r, f), irr: irr(f), pb: payback(f) };
  });
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
  return { r, flows, rows, n, i, pb, dpb, pi, beRev, sensRaw, checks, passed, cls, title, recs };
}

function generateReport() {
  const v = feasInputs(), out = $('#report');
  const { flows, rows, n, i, pb, dpb, pi, beRev, sensRaw, checks, passed, cls, title, recs } = feasAnalysis(v);
  const sens = sensRaw.map((x) => [x.lab, fmt(x.npv), pct(x.irr), yrs(x.pb)]);

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

async function downloadReport() {
  try {
    await downloadModelWorkbook(['feasibility'], `${slug($('#f-name').value || 'project')}-feasibility.xlsx`, { feasibility: feasInputs() });
  } catch (e) { alert(e.message || e); }
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
    outputs: ['Linked income statement, balance sheet and cash flow', 'Drivers from 3-year historical averages', 'Balance check, free cash flow, closing cash and debt'],
    needs: 'Income statement and balance sheet (cash flow optional)', best: 'Budgets, business plans and lender models.' },
  { id: 'dashboard', group: 'Reporting', file: 'powerbi-dataset.xlsx', href: '#dashboard', cta: 'Open dashboard',
    title: 'Dashboard & Power BI Export', desc: 'An interactive, Power BI-style view of your statements, with a one-click dataset for Power BI Desktop.',
    outputs: ['KPI cards, trends, margins and a profit bridge', 'Year and metric filters, plus a full financial matrix', 'Power BI dataset: FactFinancials and Ratios tables'],
    needs: 'Income statement and balance sheet (cash flow optional)', best: 'Board packs and monthly management reporting.' },
];
const CATEGORIES = ['All models', ...new Set(CATALOG.map((c) => c.group))];
const KEYWORDS = { dcf: 'discounted cash flow intrinsic valuation', lbo: 'leveraged buyout private equity pe sponsor', forecast: '3 statement three statement income statement balance sheet cash flow budget projection',
  npv: 'irr net present value capital budgeting payback', wacc: 'discount rate cost of capital capm beta', comps: 'comparable companies trading multiples peers', montecarlo: 'simulation risk scenario probability',
  breakeven: 'break even cvp pricing units', loan: 'debt amortization mortgage repayment', feasibility: 'project business case appraisal', ratios: 'ratio analysis credit liquidity leverage', dashboard: 'power bi kpi report charts' };
const BAND = { Valuation: '#3b5b8a', 'Deals & financing': '#2f5d50', 'Projects & investments': '#6b2e3e', Performance: '#36707c', Reporting: '#4a4f6a' };
let currentCat = 'All models', examples = null, thumbs = null, query = '', stageFilter = null, showAll = false;

function sampleState() {
  const p = parseRows(parseCSV(SAMPLE_CSV));
  return { company: SAMPLE_COMPANY, years: p.years, items: derive(p.items) };
}

// Small spreadsheet preview for each card, filled with the sample company's real numbers
function thumbData(id, v, d) {
  const k = (x) => (x == null || !isFinite(x) ? '–' : x < 0 ? `(${Math.round(-x / 1000).toLocaleString()})` : Math.round(x / 1000).toLocaleString());
  const yrsC = (n) => Array.from({ length: n }, (_, t) => fyLabel(t + 1));
  const p1 = (x) => (x == null || !isFinite(x) ? '–' : `${(x * 100).toFixed(1)}%`);
  if (id === 'dcf') {
    const r = dcfCore({ ...v, nwc0: d.nwc0 }), c = (f) => r.rows.map(f);
    return { cols: yrsC(r.rows.length), rows: [['Revenue', ...c((x) => k(x.rev))], ['EBITDA', ...c((x) => k(x.ebitda))], ['EBIT', ...c((x) => k(x.ebit))], ['NOPAT', ...c((x) => k(x.nopat))], ['CapEx', ...c((x) => k(x.capex))], ['*Unlevered FCF', ...c((x) => k(x.fcf))], ['Discount factor', ...c((x) => x.df.toFixed(3))], ['*PV of FCF', ...c((x) => k(x.pv))]] };
  }
  if (id === 'forecast') {
    const Y = project3(driversFromForecast(v), d.op, 5), c = (f) => Y.map(f);
    return { cols: yrsC(5), rows: [['Revenue', ...c((y) => k(y.rev))], ['Cost of goods sold', ...c((y) => k(y.cogs))], ['*Gross profit', ...c((y) => k(y.gp))], ['Operating expenses', ...c((y) => k(y.opex))], ['*EBITDA', ...c((y) => k(y.ebitda))], ['D&A', ...c((y) => k(y.da))], ['*Net income', ...c((y) => k(y.ni))], ['Closing cash', ...c((y) => k(y.cash))]] };
  }
  if (id === 'lbo') {
    const L = lboCore(v), c = (f) => L.rows.map(f);
    return { cols: yrsC(L.rows.length), rows: [['EBITDA', ...c((x) => k(x.e))], ['Interest', ...c((x) => k(x.int))], ['*Free cash flow', ...c((x) => k(x.fcf))], ['Debt repaid', ...c((x) => k(x.pay))], ['*Debt balance', ...c((x) => k(x.debt))], ['Debt / EBITDA', ...c((x) => x.lev.toFixed(2) + 'x')], ['*Equity value', ...c((x) => k(x.eqv))]] };
  }
  if (id === 'comps') {
    const L = [list(v.evE), list(v.evS), list(v.pe)], n = Math.max(...L.map((a) => a.length));
    return { cols: ['EV/EBITDA', 'EV/Sales', 'P/E'], rows: [...Array.from({ length: n }, (_, i) => [`Peer ${i + 1}`, ...L.map((a) => (a[i] == null ? '' : a[i].toFixed(1) + 'x'))]), ['*Median', ...L.map((a) => median(a).toFixed(2) + 'x')]] };
  }
  if (id === 'wacc') {
    const V = v.E + v.D, ke = (v.rf + v.beta * v.mrp + v.sp) / 100, kd = (v.kd / 100) * (1 - v.tax / 100);
    return { cols: ['Value'], rows: [['Risk-free rate', p1(v.rf / 100)], ['Levered beta', v.beta.toFixed(2)], ['Market risk premium', p1(v.mrp / 100)], ['*Cost of equity', p1(ke)], ['After-tax cost of debt', p1(kd)], ['Equity weight', p1(v.E / V)], ['Debt weight', p1(v.D / V)], ['*WACC', p1((v.E / V) * ke + (v.D / V) * kd)]] };
  }
  if (id === 'montecarlo') {
    const mc = monteCarlo({ ...v, sims: 600 }, d), ps = [0.05, 0.1, 0.25, 0.5, 0.75, 0.9, 0.95];
    return { cols: ['Enterprise value', 'Equity value'], rows: mc ? ps.map((p) => [`${p === 0.5 ? '*' : ''}P${Math.round(p * 100)}`, k(mc.q(p)), k(mc.q(p) - v.netDebt)]) : [] };
  }
  if (id === 'npv') {
    const f = [-Math.abs(v.inv), ...list(v.flows)], r = v.rate / 100;
    let cum = 0;
    return { cols: f.map((_, t) => `Yr ${t}`), rows: [['*Cash flow', ...f.map(k)], ['Discount factor', ...f.map((_, t) => (1 / Math.pow(1 + r, t)).toFixed(3))], ['Discounted CF', ...f.map((x, t) => k(x / Math.pow(1 + r, t)))], ['*Cumulative', ...f.map((x) => k((cum += x)))]] };
  }
  if (id === 'breakeven') {
    const cm = v.price - v.vc, vols = [0.5, 0.75, 1, 1.25, 1.5].map((m) => v.units * m);
    return { cols: ['50%', '75%', 'Plan', '125%', '150%'], rows: [['Units sold', ...vols.map((u) => Math.round(u).toLocaleString())], ['Revenue', ...vols.map((u) => k(u * v.price))], ['Variable costs', ...vols.map((u) => k(u * v.vc))], ['Fixed costs', ...vols.map(() => k(v.fixed))], ['*Profit', ...vols.map((u) => k(u * cm - v.fixed))]] };
  }
  if (id === 'loan') {
    const n = v.years * v.ppy, r = v.rate / 100 / v.ppy, pmt = r ? (v.P * r) / (1 - Math.pow(1 + r, -n)) : v.P / n;
    let bal = v.P;
    const rows = [];
    for (let y = 1; y <= Math.min(v.years, 7); y++) {
      const open = bal; let int = 0;
      for (let q = 0; q < v.ppy; q++) { const i = bal * r; int += i; bal -= pmt - i; }
      rows.push([`Year ${y}`, k(open), k(pmt * v.ppy), k(int), k(Math.max(0, bal))]);
    }
    return { cols: ['Opening', 'Payments', 'Interest', 'Closing'], rows };
  }
  if (id === 'feasibility') {
    const fv = Object.fromEntries(FEAS.map(([key, , val]) => [key, val])), F = feasFlows(fv).rows.slice(0, 5);
    return { cols: F.map((x) => `Yr ${x[0]}`), rows: [['Revenue', ...F.map((x) => k(x[1]))], ['*EBITDA', ...F.map((x) => k(x[2]))], ['Taxes', ...F.map((x) => k(x[3]))], ['*Net cash flow', ...F.map((x) => k(x[4]))]] };
  }
  if (id === 'ratios') {
    const all = ratios(), yrs = state.years.slice(-5), idx = yrs.map((_, i) => all.length - yrs.length + i);
    return { cols: yrs.map(String), rows: all[0].slice(0, 8).map(([name], j) => [name, ...idx.map((i) => String(all[i][j][1]))]) };
  }
  const yrs = state.years.slice(-5), at = (key) => yrs.map((_, i) => k(S(key)[state.years.length - yrs.length + i]));
  return { cols: yrs.map(String), rows: [['*Revenue', ...at('revenue')], ['EBITDA', ...at('ebitda')], ['Net income', ...at('netIncome')], ['Cash', ...at('cash')], ['Total debt', ...at('debt')], ['Equity', ...at('equity')]] };
}

// Runs every model on the sample company once, so each card shows real example figures
function computeExamples() {
  if (examples) return examples;
  const saved = state;
  state = sampleState();
  examples = {};
  thumbs = {};
  try {
    const d = defaults();
    MODELS.forEach((m) => {
      const v = Object.fromEntries(m.inputs(d).map(([key, , val, t]) => [key, t === 'text' ? val : +val]));
      const tpl = document.createElement('template');
      tpl.innerHTML = m.run(v, d).html;
      let pairs = [...tpl.content.querySelectorAll('.kpi')].map((x) => [x.querySelector('.label').textContent, x.querySelector('.value').textContent]);
      if (!pairs.length) pairs = [...tpl.content.querySelectorAll('tbody tr')].map((tr) => [tr.cells[0].textContent, tr.cells[tr.cells.length - 1].textContent]).filter((p) => p[1] !== '—');
      examples[m.id] = pairs.slice(0, 3);
      thumbs[m.id] = thumbData(m.id, v, d);
    });
    const fv = Object.fromEntries(FEAS.map(([key, , val]) => [key, val])), { flows } = feasFlows(fv);
    examples.feasibility = [['NPV', fmt(npv(fv.rate / 100, flows))], ['IRR', pct(irr(flows))], ['Payback', yrs(payback(flows))]];
    examples.dashboard = [[`Revenue ${state.years.at(-1)}`, fmt(L('revenue'))], ['EBITDA margin', pct(L('ebitda') / L('revenue'))], ['Revenue CAGR', pct(cagr(S('revenue')))]];
    thumbs.feasibility = thumbData('feasibility', {}, d);
    thumbs.dashboard = thumbData('dashboard', {}, d);
  } finally {
    state = saved;
  }
  return examples;
}

function sheetHTML(title, subtitle, color, t) {
  return `<div class="sheet"><div class="sheet-band" style="background:${color}"><b>${esc(title)}</b><span>${esc(subtitle)}</span></div>
    <table><thead><tr><th>Item</th>${t.cols.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead>
    <tbody>${t.rows.map((r) => { const b = r[0][0] === '*'; return `<tr${b ? ' class="b"' : ''}><td>${esc(b ? r[0].slice(1) : r[0])}</td>${r.slice(1).map((x) => `<td>${esc(x)}</td>`).join('')}</tr>`; }).join('')}</tbody></table></div>`;
}

function cardHTML(c, ex, th) {
  const m = MODELS.find((x) => x.id === c.id), href = c.href || `#model/${c.id}`, title = c.title || m.name;
  return `<article class="model-card">
    <a class="mc-thumb" href="${href}" tabindex="-1" aria-hidden="true">${sheetHTML(title, c.outputs[0], BAND[c.group], th || { cols: [], rows: [] })}</a>
    <div class="mc-body">
      <div class="mc-top"><span>${c.group}</span><span>Excel model</span></div>
      <h3><a href="${href}">${title}</a></h3>
      <p class="mc-desc">${c.desc || m.desc}</p>
      <div class="mc-example">${(ex || []).map(([key, val]) => `<span>${esc(key)} <b>${esc(val)}</b></span>`).join('')}</div>
      <div class="mc-foot">
        <a class="btn" href="${href}">${c.cta || 'Build this model'} <span aria-hidden="true">→</span></a>
        <a class="link" href="samples/models/${c.file}" download>Sample .xlsx</a>
      </div>
    </div>
  </article>`;
}

function matches(c) {
  if (!query) return true;
  const m = MODELS.find((x) => x.id === c.id), hay = [c.title, m?.name, c.desc, m?.desc, c.group, c.needs, c.best, ...c.outputs, KEYWORDS[c.id]].join(' ').toLowerCase();
  return query.toLowerCase().split(/\s+/).filter(Boolean).every((w) => hay.includes(w));
}

function renderLibrary() {
  const ex = computeExamples();
  $('#catTabs').innerHTML = CATEGORIES.map((c) => `<button data-cat="${c}" class="${c === currentCat ? 'on' : ''}">${c}<span>${c === 'All models' ? CATALOG.length : CATALOG.filter((x) => x.group === c).length}</span></button>`).join('');
  let list_ = CATALOG.filter((c) => (currentCat === 'All models' || c.group === currentCat) && (!stageFilter || stageFilter.ids.includes(c.id)) && matches(c));
  const filtered = !!(query || stageFilter);
  const note = $('#filterNote');
  note.hidden = !filtered;
  if (filtered) note.innerHTML = `${list_.length} model${list_.length === 1 ? '' : 's'} ${stageFilter ? `for <b>${esc(stageFilter.title)}</b>` : ''}${query ? ` matching “${esc(query)}”` : ''} <button class="link" data-clear-filter>Clear</button>`;
  const total = list_.length;
  if (!filtered && !showAll && currentCat === 'All models') list_ = list_.slice(0, 6);
  $('#catalog').innerHTML = list_.length ? list_.map((c) => cardHTML(c, ex[c.id], thumbs[c.id])).join('') : '<p class="muted">No models match your search. Try “valuation”, “cash flow” or “loan”.</p>';
  $('#viewAll').hidden = list_.length >= total;
  if (!$('#quickTags').innerHTML) $('#quickTags').innerHTML = ['DCF', 'LBO', '3 statement', 'WACC', 'Break-even', 'Feasibility'].map((t) => `<button type="button" data-q="${t}">${t}</button>`).join('');
  if (!$('#sectors').innerHTML) renderSectors();
  if (!$('#stages').innerHTML) renderStages();
  if (!$('#companies').innerHTML) renderCompanies();
}

/* ---------- Sector sample companies ---------- */
const SECTORS = [
  { name: 'Software & SaaS', company: 'Northwind Cloud', text: 'Recurring subscription revenue, high gross margins and heavy sales and marketing spend.', rev: 48e6, g: 0.22, gm: 0.76, opex: 0.54, da: 0.03, capex: 0.03, ta: 1.1, ca: 0.55, debt: 0.1, cl: 0.25, rate: 0.07 },
  { name: 'Retail & Consumer', company: 'Harbor Goods', text: 'Steady store and online sales, thin margins and fast-moving inventory.', rev: 85e6, g: 0.06, gm: 0.38, opex: 0.29, da: 0.025, capex: 0.03, ta: 0.6, ca: 0.6, debt: 0.25, cl: 0.18, rate: 0.065 },
  { name: 'Industrial & Manufacturing', company: 'Ironline Industries', text: 'Capital-intensive production with moderate growth and solid cash generation.', rev: 140e6, g: 0.05, gm: 0.31, opex: 0.16, da: 0.04, capex: 0.05, ta: 0.9, ca: 0.45, debt: 0.3, cl: 0.15, rate: 0.06 },
  { name: 'Healthcare & Medical Devices', company: 'Clearwell Medical', text: 'Premium products with high gross margins and significant R&D.', rev: 64e6, g: 0.09, gm: 0.63, opex: 0.41, da: 0.035, capex: 0.04, ta: 1.2, ca: 0.5, debt: 0.2, cl: 0.14, rate: 0.06 },
  { name: 'Real Estate', company: 'Stonegate Properties', text: 'Rental income from a property portfolio, high leverage and large depreciation.', rev: 38e6, g: 0.04, gm: 0.68, opex: 0.12, da: 0.18, capex: 0.12, ta: 6, ca: 0.08, debt: 0.55, cl: 0.05, rate: 0.055 },
  { name: 'Energy & Utilities', company: 'Brightvale Energy', text: 'Regulated, predictable revenue with heavy investment in long-lived assets.', rev: 210e6, g: 0.03, gm: 0.42, opex: 0.14, da: 0.08, capex: 0.11, ta: 2.2, ca: 0.2, debt: 0.45, cl: 0.1, rate: 0.055 },
  { name: 'Restaurants & Hospitality', company: 'Copper Table Group', text: 'Multi-site food and drink business with high labour costs and lease commitments.', rev: 52e6, g: 0.07, gm: 0.66, opex: 0.53, da: 0.04, capex: 0.05, ta: 0.7, ca: 0.2, debt: 0.35, cl: 0.12, rate: 0.07 },
  { name: 'Logistics & Transport', company: 'Swiftlane Logistics', text: 'Fleet-based delivery with low gross margins and steady volume growth.', rev: 95e6, g: 0.08, gm: 0.24, opex: 0.13, da: 0.05, capex: 0.06, ta: 0.8, ca: 0.35, debt: 0.35, cl: 0.14, rate: 0.065 },
  { name: 'Semiconductors', company: 'Quanta Silicon', text: 'Chip design with very high gross margins and heavy R&D spending.', rev: 120e6, g: 0.16, gm: 0.6, opex: 0.32, da: 0.05, capex: 0.08, ta: 1.4, ca: 0.55, debt: 0.08, cl: 0.12, rate: 0.055, bench: 'semis' },
  { name: 'Pharmaceuticals', company: 'Halcyon Therapeutics', text: 'Patented medicines with high margins and large research budgets.', rev: 160e6, g: 0.07, gm: 0.72, opex: 0.4, da: 0.04, capex: 0.04, ta: 1.6, ca: 0.4, debt: 0.2, cl: 0.15, rate: 0.055, bench: 'pharma' },
  { name: 'Automotive', company: 'Redline Motors', text: 'Vehicle manufacturing with thin margins and heavy capital spending.', rev: 420e6, g: 0.06, gm: 0.17, opex: 0.07, da: 0.04, capex: 0.06, ta: 1.1, ca: 0.4, debt: 0.35, cl: 0.2, rate: 0.065, bench: 'auto' },
  { name: 'Telecom', company: 'Northline Telecom', text: 'Subscription network services, stable cash flows and high debt.', rev: 300e6, g: 0.02, gm: 0.58, opex: 0.21, da: 0.15, capex: 0.16, ta: 2.6, ca: 0.12, debt: 0.45, cl: 0.12, rate: 0.055, bench: 'telecom' },
  { name: 'Media & Entertainment', company: 'Brightframe Studios', text: 'Content and streaming revenue with lumpy production costs.', rev: 140e6, g: 0.06, gm: 0.4, opex: 0.22, da: 0.06, capex: 0.04, ta: 1.5, ca: 0.3, debt: 0.3, cl: 0.18, rate: 0.065, bench: 'media' },
  { name: 'Food & Beverages', company: 'Meadowbrook Foods', text: 'Branded staples with steady demand and strong pricing power.', rev: 260e6, g: 0.05, gm: 0.45, opex: 0.27, da: 0.035, capex: 0.04, ta: 1.0, ca: 0.3, debt: 0.3, cl: 0.16, rate: 0.055, bench: 'beverages' },
  { name: 'Aerospace & Defense', company: 'Vantage Aerospace', text: 'Long-cycle contracts, large backlogs and moderate margins.', rev: 230e6, g: 0.05, gm: 0.2, opex: 0.07, da: 0.03, capex: 0.03, ta: 1.2, ca: 0.5, debt: 0.25, cl: 0.25, rate: 0.06, bench: 'aerospace' },
  { name: 'Construction & Engineering', company: 'Keystone Builders', text: 'Project-based revenue with low margins and working-capital swings.', rev: 180e6, g: 0.07, gm: 0.14, opex: 0.05, da: 0.015, capex: 0.02, ta: 0.6, ca: 0.6, debt: 0.2, cl: 0.3, rate: 0.065, bench: 'construction' },
];

// Five years of statements consistent with the sector's typical economics
function sectorRows(p) {
  const n = 5, y1 = new Date().getFullYear() - 1, years = Array.from({ length: n }, (_, i) => y1 - n + 1 + i);
  const wob = (i, s) => 1 + 0.015 * Math.sin(i * 2.1 + s), r1k = (x) => Math.round(x / 1000) * 1000;
  const L = { Revenue: [], 'Cost of Goods Sold': [], 'Operating Expenses': [], 'Depreciation & Amortization': [], 'Interest Expense': [], 'Net Income': [], Cash: [], 'Current Assets': [], 'Total Assets': [],
    'Current Liabilities': [], 'Total Debt': [], 'Total Liabilities': [], Equity: [], CapEx: [], 'Operating Cash Flow': [] };
  let prev = null;
  years.forEach((_, i) => {
    const rev = p.rev / Math.pow(1 + p.g, n - 1 - i) * wob(i, 0.4), cogs = rev * (1 - p.gm * wob(i, 1)), opex = rev * p.opex * wob(i, 2), da = rev * p.da;
    const ta = rev * p.ta, debt = ta * p.debt, int = debt * p.rate, pbt = rev - cogs - opex - da - int, ni = pbt > 0 ? pbt * 0.75 : pbt;
    const ca = ta * p.ca, cash = ca * 0.3, cl = rev * p.cl, equity = ta - cl - debt - ta * 0.04;
    const ocf = ni + da - (prev ? (rev - prev) * 0.06 : 0);
    const row = { Revenue: rev, 'Cost of Goods Sold': cogs, 'Operating Expenses': opex, 'Depreciation & Amortization': da, 'Interest Expense': int, 'Net Income': ni, Cash: cash, 'Current Assets': ca, 'Total Assets': ta,
      'Current Liabilities': cl, 'Total Debt': debt, 'Total Liabilities': ta - equity, Equity: equity, CapEx: rev * p.capex, 'Operating Cash Flow': ocf };
    Object.keys(L).forEach((k) => L[k].push(r1k(row[k])));
    prev = rev;
  });
  return [['Line Item', ...years], ...Object.entries(L).map(([k, v]) => [k, ...v])];
}

function renderSectors() {
  $('#sectors').innerHTML = SECTORS.map((p, i) => {
    const em = p.gm - p.opex;
    return `<button class="sector" type="button" data-sector="${i}">
      <div class="s-top"><h3>${p.name}</h3><span class="s-count">Sample: ${p.company}</span></div>
      <p>${p.text}</p>
      <div class="s-metrics"><span>Revenue <b>${fmt(p.rev)}</b></span><span>Growth <b>${(p.g * 100).toFixed(0)}%</b></span><span>Gross margin <b>${(p.gm * 100).toFixed(0)}%</b></span><span>EBITDA margin <b>${(em * 100).toFixed(0)}%</b></span></div>
    </button>`;
  }).join('');
}

function loadSector(i) {
  const p = SECTORS[i];
  applyRows(sectorRows(p), `${p.company} (sample)`, 1, null, `${p.name} sample`);
  const bench = p.bench || SECTOR_BENCHMARK[p.name];
  if (bench) savePref('fs-bench', bench);
  location.hash = '#build';
}

/* ---------- Real companies (SEC EDGAR via /api/company) ---------- */
const COMPANIES = [
  ['AAPL', 'Apple', 'Technology', 'hardware'], ['MSFT', 'Microsoft', 'Technology', 'software'], ['NVDA', 'NVIDIA', 'Technology', 'semis'], ['GOOGL', 'Alphabet', 'Technology', 'internet'],
  ['META', 'Meta Platforms', 'Technology', 'internet'], ['AVGO', 'Broadcom', 'Technology', 'semis'], ['ORCL', 'Oracle', 'Technology', 'software'], ['ADBE', 'Adobe', 'Technology', 'software'],
  ['CRM', 'Salesforce', 'Technology', 'software'], ['AMD', 'AMD', 'Technology', 'semis'], ['INTC', 'Intel', 'Technology', 'semis'],
  ['AMZN', 'Amazon', 'Consumer', 'online-retail'], ['WMT', 'Walmart', 'Consumer', 'retail'], ['COST', 'Costco', 'Consumer', 'retail'], ['HD', 'Home Depot', 'Consumer', 'retail'],
  ['NKE', 'Nike', 'Consumer', 'apparel'], ['MCD', 'McDonald’s', 'Consumer', 'restaurants'], ['SBUX', 'Starbucks', 'Consumer', 'restaurants'], ['MAR', 'Marriott', 'Consumer', 'hotels'],
  ['KO', 'Coca-Cola', 'Consumer staples', 'beverages'], ['PEP', 'PepsiCo', 'Consumer staples', 'beverages'], ['PG', 'Procter & Gamble', 'Consumer staples', 'household'],
  ['JNJ', 'Johnson & Johnson', 'Healthcare', 'pharma'], ['LLY', 'Eli Lilly', 'Healthcare', 'pharma'], ['PFE', 'Pfizer', 'Healthcare', 'pharma'], ['ABT', 'Abbott Laboratories', 'Healthcare', 'meddevices'], ['UNH', 'UnitedHealth Group', 'Healthcare', 'healthcare-services'],
  ['XOM', 'ExxonMobil', 'Energy', 'oilgas'], ['CVX', 'Chevron', 'Energy', 'oilgas'], ['NEE', 'NextEra Energy', 'Energy', 'utilities'], ['DUK', 'Duke Energy', 'Energy', 'utilities'],
  ['CAT', 'Caterpillar', 'Industrials', 'industrial'], ['DE', 'Deere & Company', 'Industrials', 'industrial'], ['BA', 'Boeing', 'Industrials', 'aerospace'], ['LMT', 'Lockheed Martin', 'Industrials', 'aerospace'],
  ['TSLA', 'Tesla', 'Industrials', 'auto'], ['GM', 'General Motors', 'Industrials', 'auto'], ['UPS', 'UPS', 'Industrials', 'logistics'], ['DOW', 'Dow', 'Materials', 'chemicals'],
  ['DIS', 'Disney', 'Communication', 'media'], ['NFLX', 'Netflix', 'Communication', 'media'], ['VZ', 'Verizon', 'Communication', 'telecom'], ['T', 'AT&T', 'Communication', 'telecom'],
  ['PLD', 'Prologis', 'Real estate', 'reit'], ['O', 'Realty Income', 'Real estate', 'reit'],
];
let coGroup = 'All';
function renderCompanies() {
  const groups = ['All', ...new Set(COMPANIES.map((c) => c[2]))];
  $('#coTabs').innerHTML = groups.map((g) => `<button data-cogroup="${g}" class="${g === coGroup ? 'on' : ''}">${g}<span>${g === 'All' ? COMPANIES.length : COMPANIES.filter((c) => c[2] === g).length}</span></button>`).join('');
  $('#companies').innerHTML = COMPANIES.filter((c) => coGroup === 'All' || c[2] === coGroup).map(([t, name, group, bench]) => `<button class="company" type="button" data-ticker="${t}" data-bench="${bench}">
      <span class="co-mark">${t}</span><span class="co-name"><strong>${esc(name)}</strong><span>${esc(BENCHMARKS.find((b) => b.id === bench)?.name || group)}</span></span><span class="co-go" aria-hidden="true">→</span></button>`).join('');
}

async function loadCompany(ticker, bench) {
  ticker = ticker.trim().toUpperCase();
  if (!ticker) return;
  const m = $('#tickerMsg');
  m.className = 'upload-msg small';
  m.textContent = `Loading ${ticker} from SEC filings…`;
  $$('.company').forEach((b) => b.classList.toggle('loading', b.dataset.ticker === ticker));
  try {
    const res = await fetch(`/api/company?ticker=${encodeURIComponent(ticker)}`, { headers: { Accept: 'application/json' } });
    const isJson = /json/.test(res.headers.get('content-type') || '');
    const j = isJson ? await res.json() : null;
    if (!j) throw new Error('Real company data needs the site’s Netlify Functions, so it works on the deployed site (deploy from GitHub), not on a static copy.');
    if (!res.ok) throw new Error(j.error || 'Company data unavailable.');
    if (applyRows(j.rows, `${j.name} (${j.ticker})`, 1, null, `SEC 10-K filings ${j.years[0]}–${j.years.at(-1)}`)) {
      if (bench) savePref('fs-bench', bench);
      m.className = 'upload-msg small ok';
      m.textContent = `Loaded ${j.name}: ${j.years.length} fiscal years from SEC filings.`;
      location.hash = '#build';
    }
  } catch (e) {
    m.className = 'upload-msg small err';
    m.textContent = e.message;
  }
  $$('.company').forEach((b) => b.classList.remove('loading'));
}

/* ---------- Business stages ---------- */
const ICONS = {
  seed: '<svg viewBox="0 0 100 100" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"><path d="M50 88c-16 0-26-12-26-28 0-20 14-34 26-34s26 14 26 34c0 16-10 28-26 28z"/><path d="M50 26c0-8 4-14 10-16"/><path d="M28 44c8 6 14 0 22 6 8-6 14 0 22-6"/></svg>',
  sprout: '<svg viewBox="0 0 100 100" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"><path d="M50 92V46"/><path d="M50 58c-22 2-30-12-30-26 16 0 30 8 30 26z"/><path d="M50 50c0-18 12-30 32-30 0 18-12 30-32 30z"/><path d="M30 92h40"/></svg>',
  building: '<svg viewBox="0 0 100 100" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 92V30l32-16v78"/><path d="M50 92V40l32 12v40"/><path d="M28 42h10M28 56h10M28 70h10M60 60h10M60 74h10"/><path d="M10 92h80"/></svg>',
  handshake: '<svg viewBox="0 0 100 100" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"><path d="M8 40l16-10 18 6 12-8 12 4 26 8"/><path d="M8 40l14 22 14 12c4 4 10 4 14 0l22-20"/><path d="M42 36l-10 12c-2 4 4 8 8 4l10-8 16 14"/><path d="M92 40L78 62"/></svg>',
};
const STAGES = [
  { title: 'Startup', sub: 'New ventures and projects', text: 'Check an idea pays off before you commit money.', ids: ['breakeven', 'npv', 'feasibility', 'loan'], icon: 'seed' },
  { title: 'Growing company', sub: 'Scaling revenue and planning cash', text: 'Forecast the statements and see what the business is worth.', ids: ['forecast', 'dcf', 'ratios', 'dashboard'], icon: 'sprout' },
  { title: 'Established company', sub: 'Valuing a mature business', text: 'Cross-check value with cash flows, peers and risk.', ids: ['dcf', 'comps', 'wacc', 'montecarlo'], icon: 'building' },
  { title: 'Deals & financing', sub: 'Buyouts, loans and refinancing', text: 'Test returns, debt capacity and repayment.', ids: ['lbo', 'loan', 'wacc', 'comps'], icon: 'handshake' },
];
function renderStages() {
  $('#stages').innerHTML = STAGES.map((s, i) => `<button class="stage" type="button" data-stage="${i}">
    <h3>${s.title}</h3><p>${s.sub}</p><p>${s.text}</p>
    <p><b>${s.ids.length} models:</b> ${s.ids.map((id) => CATALOG.find((c) => c.id === id)?.title || MODELS.find((m) => m.id === id).name).join(', ')}</p>
    ${ICONS[s.icon]}</button>`).join('');
}

/* ---------- Model workspace ---------- */
let manualMode = false;

function dataSummaryHTML() {
  const covered = [['Income statement', ['revenue', 'netIncome', 'ebitda']], ['Balance sheet', ['totalAssets', 'equity', 'debt', 'cash']], ['Cash flow', ['ocf', 'capex']]]
    .map(([n, ks]) => { const ok = ks.some((k) => state.items[k]); return `<span class="${ok ? 'ok' : 'miss'}">${ok ? '✓' : '–'} ${n}</span>`; }).join('');
  return `<div class="data-summary">
      <div class="ds-check" aria-hidden="true">✓</div>
      <div class="grow">
        <strong>${esc(state.company)}</strong>
        <div class="muted small">${esc(state.years[0])}–${esc(state.years.at(-1))} · ${Object.keys(state.items).length} line items</div>
        <div class="coverage">${covered}</div>
      </div>
      <a class="btn ghost" href="#upload">View data</a>
      <button class="btn ghost" data-action="change-data">Upload different statements</button>
    </div>
    <div class="uploader-slot" hidden></div>`;
}

function renderWorkspace(id) {
  const m = MODELS.find((x) => x.id === id) || MODELS[0];
  currentModel = m.id;
  $('#wsCrumb').textContent = m.name;
  $('#wsGroup').textContent = m.group;
  $('#wsTitle').textContent = m.name;
  $('#wsDesc').textContent = m.desc;
  const cat = CATALOG.find((c) => c.id === m.id);
  $('#wsInfo').innerHTML = cat ? `<span><b>You get:</b> ${cat.outputs.join(' · ')}</span><span><b>Data needed:</b> ${cat.needs}</span>` : '';
  const groups = [...new Set(MODELS.map((x) => x.group))];
  $('#modelSelect').innerHTML = groups.map((g) => `<optgroup label="${g}">${MODELS.filter((x) => x.group === g)
    .map((x) => `<option value="${x.id}"${x.id === m.id ? ' selected' : ''}>${x.name}</option>`).join('')}</optgroup>`).join('');

  const has = hasData();
  $('#wsData').innerHTML = has
    ? dataSummaryHTML()
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

/* ---------- Excel model builder page ---------- */
const BUILD_IDS = ['dcf', 'comps', 'wacc', 'montecarlo', 'lbo', 'loan', 'npv', 'breakeven', 'feasibility'];
let buildSel = new Set(BUILD_IDS);

function renderBuild() {
  const has = hasData();
  $('#buildData').innerHTML = has ? dataSummaryHTML() : '<div class="uploader-slot"></div>';
  $('#buildChoose').hidden = $('#buildExport').hidden = !has;
  $$('#buildStepper li').forEach((li) => { li.className = +li.dataset.step === 1 ? (has ? 'done' : 'active') : has ? 'active' : ''; });
  if (!has) return;
  const groups = [...new Set(CATALOG.filter((c) => BUILD_IDS.includes(c.id)).map((c) => c.group))];
  $('#modelChecks').innerHTML = groups.map((g) => `<div class="check-group"><h4>${g}</h4><div class="check-grid">${CATALOG.filter((c) => c.group === g && BUILD_IDS.includes(c.id)).map((c) => {
    const m = MODELS.find((x) => x.id === c.id), x = XL_MODELS[c.id];
    return `<label class="check"><input type="checkbox" value="${c.id}"${buildSel.has(c.id) ? ' checked' : ''}>
      <span class="box" aria-hidden="true"></span>
      <span class="ct"><strong>${c.title || m.name}</strong><span>${x.desc}</span><em>Sheet: ${x.sheet}</em></span></label>`;
  }).join('')}</div></div>`).join('');
  updateBuildSummary();
}

function updateBuildSummary() {
  const n = buildSel.size;
  $('#buildSummary').textContent = n
    ? `Always included: Cover, Dashboard (with charts), Assumptions, Income Statement, Balance Sheet, Cash Flow, Ratios and Historicals. Plus ${n} model sheet${n > 1 ? 's' : ''}: ${BUILD_IDS.filter((id) => buildSel.has(id)).map((id) => XL_MODELS[id].sheet).join(', ')}.`
    : 'Choose at least one model in step 2.';
  $('#buildBtn').disabled = !n;
}

async function buildExcel() {
  const btn = $('#buildBtn'), m = $('#buildMsg');
  btn.disabled = true;
  m.className = 'upload-msg small';
  m.textContent = 'Building your workbook…';
  try {
    const ids = BUILD_IDS.filter((id) => buildSel.has(id));
    await downloadModelWorkbook(ids, `${slug(state.company)}-financial-model.xlsx`);
    m.className = 'upload-msg small ok';
    m.textContent = `Done. Your Excel model with ${ids.length} model sheet${ids.length > 1 ? 's' : ''} has been downloaded.`;
  } catch (e) {
    m.className = 'upload-msg small err';
    m.textContent = e.message || String(e);
  } finally {
    btn.disabled = false;
  }
}

/* ---------- Colour schemes ---------- */
const SCHEMES = [['classic', 'Classic ivory', '#1c1c1a'], ['navy', 'Navy', '#14305a'], ['emerald', 'Emerald', '#1d5c3f'], ['burgundy', 'Burgundy', '#6b1f2e'], ['slate', 'Slate blue', '#3a4a6b'], ['dark', 'Dark', '#121211']];
const currentScheme = () => document.documentElement.dataset.scheme || 'classic';
function renderThemeMenu() {
  const cur = currentScheme();
  $('#themeMenu').innerHTML = `<p>Colour scheme</p>${SCHEMES.map(([id, name, c]) => `<button type="button" data-scheme="${id}" class="${id === cur ? 'on' : ''}"><i style="background:${c}"></i>${name}</button>`).join('')}`;
  $('#themeSwatch').style.background = SCHEMES.find((x) => x[0] === cur)[2];
}
function setScheme(id) {
  if (id === 'classic') delete document.documentElement.dataset.scheme; else document.documentElement.dataset.scheme = id;
  try { localStorage.setItem('fs-scheme', id); } catch (e) {}
  $('#themeMenu').hidden = true;
  renderThemeMenu();
  themeCharts();
  route();
}

/* ---------- Routing & init ---------- */
function route() {
  const [view, arg] = (location.hash || '#library').slice(1).split('/');
  const v = view === 'model' ? 'workspace' : ['build', 'library', 'dashboard', 'indicators', 'upload', 'feasibility'].includes(view) ? view : 'library';
  $$('.view').forEach((el) => el.classList.toggle('active', el.id === `view-${v}`));
  $$('nav a').forEach((a) => a.classList.toggle('active', a.dataset.view === (v === 'workspace' ? 'library' : v)));
  $('#companyPill').textContent = hasData() ? state.company : 'No data loaded';
  if (v === 'build') renderBuild();
  if (v === 'indicators') { renderIndicators(); loadMarket(); }
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
  if ($('#view-build').classList.contains('active') && hasData()) $('#buildChoose').scrollIntoView({ behavior: 'smooth' });
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
    if (a.dataset.action === 'change-data') { a.closest('.data-summary').hidden = true; a.closest('.step').querySelector('.uploader-slot').hidden = false; }
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
  const toModels = () => $('#modelsSec').scrollIntoView({ behavior: 'smooth' });
  $('#searchInput').addEventListener('input', (e) => { query = e.target.value.trim(); renderLibrary(); });
  $('#searchForm').addEventListener('submit', (e) => { e.preventDefault(); query = $('#searchInput').value.trim(); renderLibrary(); toModels(); });
  $('#quickTags').addEventListener('click', (e) => {
    const b = e.target.closest('[data-q]');
    if (!b) return;
    $('#searchInput').value = query = b.dataset.q; stageFilter = null; currentCat = 'All models'; renderLibrary(); toModels();
  });
  $('#viewAll').addEventListener('click', () => { showAll = true; renderLibrary(); });
  $('#filterNote').addEventListener('click', (e) => {
    if (!e.target.closest('[data-clear-filter]')) return;
    query = ''; stageFilter = null; $('#searchInput').value = ''; renderLibrary();
  });
  $('#coTabs').addEventListener('click', (e) => { const b = e.target.closest('[data-cogroup]'); if (b) { coGroup = b.dataset.cogroup; renderCompanies(); } });
  $('#companies').addEventListener('click', (e) => { const b = e.target.closest('[data-ticker]'); if (b) loadCompany(b.dataset.ticker, b.dataset.bench); });
  $('#tickerForm').addEventListener('submit', (e) => { e.preventDefault(); const t = $('#tickerInput').value.trim().toUpperCase(); loadCompany(t, COMPANIES.find((c) => c[0] === t)?.[3]); });
  $('#sectors').addEventListener('click', (e) => { const b = e.target.closest('[data-sector]'); if (b) loadSector(+b.dataset.sector); });
  $('#stages').addEventListener('click', (e) => {
    const b = e.target.closest('[data-stage]');
    if (!b) return;
    stageFilter = STAGES[+b.dataset.stage]; currentCat = 'All models'; renderLibrary(); toModels();
  });
  document.querySelectorAll('[data-car]').forEach((b) => b.addEventListener('click', () => {
    const c = $('#stages'), w = c.querySelector('.stage')?.offsetWidth || 320;
    c.scrollBy({ left: +b.dataset.car * (w + 28), behavior: 'smooth' });
  }));
  $('#menuBtn').addEventListener('click', () => {
    const open = $('#topbar').classList.toggle('open');
    $('#menuBtn').setAttribute('aria-expanded', open);
  });
  $('#mainNav').addEventListener('click', () => { $('#topbar').classList.remove('open'); $('#menuBtn').setAttribute('aria-expanded', 'false'); });
  $('#modelChecks').addEventListener('change', (e) => {
    if (!e.target.matches('input[type=checkbox]')) return;
    e.target.checked ? buildSel.add(e.target.value) : buildSel.delete(e.target.value);
    updateBuildSummary();
  });
  $('#buildChoose').addEventListener('click', (e) => {
    const b = e.target.closest('[data-sel]');
    if (!b) return;
    buildSel = new Set(b.dataset.sel === 'all' ? BUILD_IDS : []);
    renderBuild();
  });
  $('#buildBtn').addEventListener('click', buildExcel);

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

  renderThemeMenu();
  initIndicators();
  initChat();
  $('#themeBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    const m = $('#themeMenu'); m.hidden = !m.hidden;
    $('#themeBtn').setAttribute('aria-expanded', String(!m.hidden));
  });
  $('#themeMenu').addEventListener('click', (e) => { const b = e.target.closest('[data-scheme]'); if (b) setScheme(b.dataset.scheme); });
  document.addEventListener('click', (e) => { if (!e.target.closest('.theme-pick')) $('#themeMenu').hidden = true; });
  window.addEventListener('hashchange', () => { route(); window.scrollTo(0, 0); });
  route();
}

document.addEventListener('DOMContentLoaded', init);
