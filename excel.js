/* Finance Studio — Excel model builder.
   Builds a formatted, formula-driven workbook: Cover, Dashboard (KPI tiles and native Excel charts),
   Assumptions, Income Statement, Balance Sheet, Cash Flow, statement-driven DCF and Ratios, any other chosen
   models, and the uploaded Historicals. Inputs are blue, links to Historicals are green, formulas are black.
   A small evaluator computes every formula so the file opens with correct values in any viewer; Excel also
   recalculates on load. ExcelJS and JSZip load only when a workbook is built. Depends on app.js and engine.js. */

const EXCELJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
const JSZIP_URL = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
const scriptCache = {};
function loadScript(url, global) {
  return (scriptCache[url] ??= new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = url;
    s.onload = () => res(window[global]);
    s.onerror = () => { delete scriptCache[url]; rej(new Error('The Excel builder failed to load. Check your internet connection.')); };
    document.head.appendChild(s);
  }));
}

const XF = {
  money: '#,##0;(#,##0);"–"', money2: '$#,##0.00;($#,##0.00)', pct: '0.0%', pctS: '0.0%;(0.0%);"–"', pct2: '0.00%', mult: '0.00"x"',
  num: '#,##0', dec: '0.00', df: '0.0000', yrs: '0.0 "yrs"', check: '#,##0.00;(#,##0.00);"✓ balanced"',
};
const colL = (i) => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; };
const colI = (l) => [...l].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
const Q = (sheet) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheet) ? sheet : `'${sheet}'`);
const IS_ = 'Income Statement', BS_ = 'Balance Sheet', CF_ = 'Cash Flow';
const ref = (sheet, addr) => `${Q(sheet)}!${addr}`;

/* ---------- Formula evaluator (covers the functions these models use) ---------- */
const TOKEN_RE = /\s*(?:(\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)|("(?:[^"]|"")*")|((?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|([A-Z][A-Z0-9.]*)(?=\()|(<=|>=|<>|[-+*/^(),=<>&]))/y;
function tokenize(f) {
  const out = [];
  TOKEN_RE.lastIndex = 0;
  while (TOKEN_RE.lastIndex < f.length) {
    const start = TOKEN_RE.lastIndex, m = TOKEN_RE.exec(f);
    if (!m) { if (/^\s*$/.test(f.slice(start))) break; throw new Error(`Cannot parse formula: ${f}`); }
    if (m[1]) out.push({ t: 'num', v: +m[1] });
    else if (m[2]) out.push({ t: 'str', v: m[2].slice(1, -1).replace(/""/g, '"') });
    else if (m[3]) out.push({ t: 'ref', v: m[3] });
    else if (m[4]) out.push({ t: 'fn', v: m[4] });
    else out.push({ t: 'op', v: m[5] });
  }
  return out;
}
const isErr = (x) => typeof x === 'number' && !isFinite(x);
const N = (x) => (x == null || x === '' ? 0 : typeof x === 'boolean' ? +x : typeof x === 'string' ? (x.trim() === '' || isNaN(+x) ? NaN : +x) : x);
const flat = (args) => args.flat(Infinity);
const nums = (args) => flat(args).filter((x) => typeof x === 'number');
const XL_FN = {
  SUM: (a) => nums(a).reduce((s, x) => s + x, 0),
  COUNT: (a) => nums(a).length,
  MIN: (a) => (nums(a).length ? Math.min(...nums(a)) : 0),
  MAX: (a) => (nums(a).length ? Math.max(...nums(a)) : 0),
  AVERAGE: (a) => { const n = nums(a); return n.length ? n.reduce((s, x) => s + x, 0) / n.length : NaN; },
  MEDIAN: (a) => (nums(a).length ? median(nums(a)) : NaN),
  ABS: ([x]) => Math.abs(N(x)),
  ROUND: ([x, d]) => { const k = Math.pow(10, N(d)); return Math.round(N(x) * k) / k; },
  NPV: ([r, ...v]) => nums(v).reduce((s, x, i) => s + x / Math.pow(1 + N(r), i + 1), 0),
  IRR: ([v]) => irr(nums([v])) ?? NaN,
  MIRR: ([v, fin, re]) => mirr(nums([v]), N(fin), N(re)) ?? NaN,
  PMT: ([r, n, pv]) => { r = N(r); n = N(n); pv = N(pv); return r === 0 ? -pv / n : -(pv * r * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1); },
  IF: ([c, a, b]) => (N(c) ? a : b ?? false),
  AND: (a) => flat(a).every((x) => N(x)),
  OR: (a) => flat(a).some((x) => N(x)),
  IFERROR: ([a, b]) => (typeof a === 'number' && !isFinite(a) ? b : a),
};

// Circular references are an error unless the book opts in with book.iterative (an explicit switch, as in Excel's
// iterative calculation): then a cell met again mid-calculation returns its previous pass's value and the whole
// book is recalculated until no value moves by more than 1e-9 (at most 100 passes).
function evaluateBook(book) {
  const byName = Object.fromEntries(book.sheets.map((s) => [s.name, s]));
  let circular = false;
  const cellVal = (sheet, addr) => {
    const c = byName[sheet]?.cells[addr];
    if (!c) return null;
    if (c.f == null) return c.v ?? null;
    if (c.r === undefined) {
      if (c.busy) {
        if (!book.iterative) throw new Error(`Circular reference at ${sheet}!${addr}`);
        circular = true;
        return c.last ?? 0;
      }
      c.busy = true;
      c.r = evalFormula(c.f, sheet);
      c.busy = false;
    }
    return c.r;
  };
  const refVal = (rf, sheet) => {
    const k = rf.lastIndexOf('!');
    const sh = k < 0 ? sheet : rf.slice(0, k).replace(/^'|'$/g, ''), a = (k < 0 ? rf : rf.slice(k + 1)).replace(/\$/g, '');
    if (!a.includes(':')) return cellVal(sh, a);
    const [p, q] = a.split(':').map((x) => x.match(/([A-Z]+)(\d+)/));
    const out = [];
    for (let r = +p[2]; r <= +q[2]; r++) for (let c = colI(p[1]); c <= colI(q[1]); c++) out.push(cellVal(sh, colL(c) + r));
    return out;
  };
  function evalFormula(f, sheet) {
    const t = tokenize(f);
    let i = 0;
    const isOp = (...o) => t[i] && t[i].t === 'op' && o.includes(t[i].v);
    const expect = (v) => { if (!isOp(v)) throw new Error(`Expected ${v} in ${f}`); i++; };
    function cmp() {
      let a = cat();
      while (isOp('=', '<>', '<', '>', '<=', '>=')) {
        const op = t[i++].v, b = cat();
        const [x, y] = typeof a === 'string' || typeof b === 'string' ? [String(a ?? ''), String(b ?? '')] : [N(a), N(b)];
        a = { '=': x === y, '<>': x !== y, '<': x < y, '>': x > y, '<=': x <= y, '>=': x >= y }[op];
      }
      return a;
    }
    function cat() { let a = add(); while (isOp('&')) { i++; a = String(a ?? '') + String(add() ?? ''); } return a; }
    function add() { let a = mul(); while (isOp('+', '-')) { const op = t[i++].v, b = mul(); a = op === '+' ? N(a) + N(b) : N(a) - N(b); } return a; }
    function mul() {
      let a = pow();
      while (isOp('*', '/')) { const op = t[i++].v, b = pow(); a = op === '*' ? N(a) * N(b) : N(b) === 0 ? Infinity : N(a) / N(b); }
      return a;
    }
    function pow() { let a = unary(); while (isOp('^')) { i++; a = Math.pow(N(a), N(unary())); } return a; }
    function unary() { if (isOp('-')) { i++; return -N(unary()); } if (isOp('+')) { i++; return N(unary()); } return primary(); }
    function primary() {
      const tk = t[i++];
      if (!tk) throw new Error(`Unexpected end of ${f}`);
      if (tk.t === 'num' || tk.t === 'str') return tk.v;
      if (tk.t === 'ref') return refVal(tk.v, sheet);
      if (tk.t === 'fn') {
        expect('(');
        const args = [];
        if (!isOp(')')) { args.push(cmp()); while (isOp(',')) { i++; args.push(cmp()); } }
        expect(')');
        if (!XL_FN[tk.v]) throw new Error(`Unsupported function ${tk.v}`);
        return XL_FN[tk.v](args);
      }
      if (tk.v === '(') { const v = cmp(); expect(')'); return v; }
      throw new Error(`Unexpected ${tk.v} in ${f}`);
    }
    const v = cmp();
    if (i < t.length) throw new Error(`Unexpected ${t[i].v} in ${f}`);
    return v;
  }
  book.sheets.forEach((s) => Object.keys(s.cells).forEach((a) => cellVal(s.name, a)));
  book.passes = 1;
  while (book.iterative && circular && book.passes < 100) {
    let moved = 0;
    const formulaCells = book.sheets.flatMap((s) => Object.values(s.cells).filter((c) => c.f != null));
    formulaCells.forEach((c) => { c.last = c.r; c.r = undefined; });
    circular = false;
    book.sheets.forEach((s) => Object.keys(s.cells).forEach((a) => cellVal(s.name, a)));
    formulaCells.forEach((c) => { if (typeof c.r === 'number' && typeof c.last === 'number') moved = Math.max(moved, Math.abs(c.r - c.last)); });
    book.passes++;
    if (moved <= 1e-9) break;
  }
  book.value = (sheet, addr) => cellVal(sheet, addr.replace(/\$/g, ''));
}

/* ---------- Workbook model ---------- */
function newBook() { return { sheets: [], hist: null, out: {}, charts: [], built: [] }; }

function addSheet(book, name, title, subtitle, tab = '0B2545') {
  const sh = { name, cells: {}, widths: { A: 40 }, sections: [], merges: [], heights: {}, freeze: null, tab };
  book.sheets.push(sh);
  const put = (addr, x, z, s) => {
    sh.cells[addr] = typeof x === 'string' && x[0] === '=' ? { f: x.slice(1), z, s } : { v: x, z, s };
    return addr;
  };
  const abs = (a) => a.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  const api = {
    sh, put, abs,
    has: (addr) => !!sh.cells[addr],
    section: (row, text, lastCol = 6) => { put(`A${row}`, text, null, 'section'); sh.sections.push([row, lastCol]); },
    header: (row, labels, startCol = 0) => labels.forEach((l, i) => put(`${colL(startCol + i)}${row}`, l, null, i + startCol === 0 ? 'headL' : 'head')),
    // Assumption row: label in A, value in B (input or link), optional note in C. Returns the absolute ref.
    assume: (row, text, x, z, note) => {
      put(`A${row}`, text, null, 'label');
      put(`B${row}`, x, z, typeof x === 'string' && x[0] === '=' ? 'link' : 'input');
      if (note) put(`C${row}`, note, null, 'note');
      return abs(`B${row}`);
    },
    // Output row: bold label in A, formula in B. Returns the absolute ref (sheet-qualified).
    output: (row, text, f, z, s = 'output', how) => { put(`A${row}`, text, null, s === 'output' ? 'labelB' : 'label'); put(`B${row}`, f, z, s); if (how) put(`C${row}`, how, null, 'how'); return ref(name, abs(`B${row}`)); },
  };
  if (title) {
    put('A1', title, null, 'title');
    put('A2', subtitle, null, 'subtitle');
    put('A3', 'Blue = input you can change · Green = linked to source data · Black = formula · Shaded = key output', null, 'note');
  }
  return api;
}

const baseYear = () => parseInt(String(state.years.at(-1) || '').match(/(19|20)\d{2}/)?.[0]) || null;
const yearLabel = (t) => (baseYear() ? `${baseYear() + t}E` : `Year ${t}`);
const sub = (desc) => `${state.company || 'Manual assumptions'} · ${desc}`;
const near = (a, b) => a != null && b != null && Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b));

// Historical statements exactly as uploaded; remembers where each line item lives
function sheetHistoricals(book) {
  const { put, header, section, sh } = addSheet(book, 'Historicals', 'Historical Financial Statements (as uploaded)', `${state.company} · source data for the model`, '7A8699');
  const n = state.years.length, keys = Object.keys(LABELS).filter((k) => S(k));
  section(5, 'Uploaded line items', n);
  header(6, ['Line item', ...state.years.map(String)]);
  const refs = {};
  keys.forEach((k, j) => {
    const row = 7 + j;
    put(`A${row}`, LABELS[k], null, 'label');
    S(k).forEach((v, i) => { if (v != null) put(`${colL(i + 1)}${row}`, v, XF.money, 'input'); });
    const li = S(k).map((v, i) => (v == null ? -1 : i)).filter((i) => i >= 0).pop();
    refs[k] = { row, latest: `Historicals!$${colL(li + 1)}$${row}`, latestVal: S(k)[li] };
  });
  sh.widths = { A: 34, ...Object.fromEntries(state.years.map((_, i) => [colL(i + 1), 14])) };
  sh.freeze = { x: 1, y: 6 };
  book.hist = refs;
}
function linkOr(book, key, val) {
  const h = book.hist?.[key];
  return h && near(h.latestVal, val) ? `=${h.latest}` : val;
}

/* ---------- Core: Assumptions + linked Income Statement, Balance Sheet, Cash Flow ---------- */
// Plain-English formula for each forecast line, shown beside the numbers so clients can follow the maths
const HOW = Object.fromEntries(Object.entries({
  'Income Statement': { rev: 'Prior-year revenue × (1 + revenue growth)', growth: 'Revenue ÷ prior-year revenue – 1', cogs: 'Revenue × (1 – gross margin)', gp: 'Revenue – cost of goods sold', gm: 'Gross profit ÷ revenue',
    opex: 'Revenue × operating expenses %', ebitda: 'Gross profit – operating expenses', em: 'EBITDA ÷ revenue', da: 'Revenue × D&A %', ebit: 'EBITDA – depreciation & amortization', ebm: 'EBIT ÷ revenue',
    int: 'Opening debt (prior-year balance sheet) × interest rate', pbt: 'EBIT – interest expense', tax: 'Profit before tax × tax rate (zero when there is a loss)', etr: 'Income tax ÷ profit before tax', ni: 'Profit before tax – income tax', nm: 'Net income ÷ revenue' },
  'Balance Sheet': { cash: 'Closing cash from the cash flow statement', noncash: 'Revenue × non-cash current assets %', tca: 'Cash + non-cash current assets', ppe: 'Prior year + capital expenditure – D&A', ta: 'Current assets + non-current assets',
    cl: 'Revenue × current liabilities %', debt: 'Prior-year debt – debt repaid this year', other: 'Held at the last actual level', tl: 'Current liabilities + debt + other liabilities', eq: 'Prior-year equity + net income – dividends',
    tle: 'Total liabilities + equity', check: 'Total assets – (liabilities + equity); must be zero' },
  'Cash Flow': { ni: 'From the income statement', da: 'Non-cash expense added back', dnca: 'Prior-year – this-year non-cash current assets', dcl: 'This-year – prior-year current liabilities', cfo: 'Net income + D&A + working-capital changes',
    capex: 'Revenue × CapEx % (cash paid out)', cfi: 'Capital expenditure', repay: 'Lower of opening debt and the planned repayment', div: 'Net income × dividend payout %', cff: 'Debt repayment + dividends',
    net: 'Operating + investing + financing cash flows', open: 'Prior-year closing cash', close: 'Opening cash + net change in cash', fcf: 'Cash from operations – capital expenditure', fcfm: 'Free cash flow ÷ revenue' },
}).flatMap(([sheet, m]) => Object.entries(m).map(([k, v]) => [`${k}@${sheet}`, v])));

function buildCore(book, dr, n) {
  const withData = hasData(), H = withData ? state.years.length : 1, last = H - 1, op = openingBS();
  const aCol = (i) => colL(1 + i), fCol = (t) => colL(H + t), lastA = aCol(last), allCols = Array.from({ length: H + n }, (_, j) => colL(1 + j));
  const hist = (k, i) => (withData && book.hist?.[k] && S(k)[i] != null ? `Historicals!${colL(i + 1)}${book.hist[k].row}` : null);
  const actualLabel = (i) => {
    if (!withData) return 'Base';
    const y = String(state.years[i]).match(/(19|20)\d{2}/)?.[0];
    return y ? `${y}A` : String(state.years[i]);
  };
  const colLabel = (j) => (j < H ? actualLabel(j) : yearLabel(j - H + 1));

  // --- Assumptions ---
  const As = addSheet(book, 'Assumptions', 'Forecast Assumptions', sub('drivers behind the three financial statements'));
  As.section(5, 'Forecast drivers', n + 1);
  As.put('A6', 'Driver', null, 'headL'); As.put('B6', 'Basis', null, 'headL');
  for (let t = 1; t <= n; t++) As.put(`${colL(1 + t)}6`, yearLabel(t), null, 'headF');
  const anyHist = (...ks) => withData && ks.every((k) => S(k)?.some((v) => v != null));
  const avgNote = (ok) => (ok ? '3-year historical average' : 'Default estimate');
  const DR = [
    ['growth', 'Revenue growth', XF.pct, withData && S('revenue')?.filter((v) => v != null).length > 1 ? `Revenue CAGR ${state.years[0]}–${state.years[last]}` : 'Default estimate'],
    ['gm', 'Gross margin', XF.pct, avgNote(anyHist('grossProfit'))],
    ['opex', 'Operating expenses % of revenue', XF.pct, avgNote(anyHist('opex'))],
    ['da', 'D&A % of revenue', XF.pct, avgNote(anyHist('da'))],
    ['capex', 'CapEx % of revenue', XF.pct, avgNote(anyHist('capex'))],
    ['tax', 'Tax rate', XF.pct, anyHist('netIncome', 'interest') ? 'Historical effective rate' : 'Default estimate'],
    ['rate', 'Interest rate on opening debt', XF.pct, anyHist('interest', 'debt') ? 'Historical interest / debt' : 'Default estimate'],
    ['nca', 'Non-cash current assets % of revenue', XF.pct, avgNote(anyHist('currentAssets', 'cash'))],
    ['cl', 'Current liabilities % of revenue', XF.pct, avgNote(anyHist('currentLiabilities'))],
    ['repay', 'Debt repayment per year', XF.money, '10% of opening debt'],
    ['payout', 'Dividend payout % of net income', XF.pct, 'Your policy'],
  ];
  const AR = {};
  DR.forEach(([k, label, z, basis], j) => {
    const row = 7 + j;
    AR[k] = row;
    As.put(`A${row}`, label, null, 'label');
    As.put(`B${row}`, basis, null, 'note');
    for (let t = 1; t <= n; t++) {
      const val = Array.isArray(dr[k]) ? dr[k][t - 1] : dr[k];
      As.put(`${colL(1 + t)}${row}`, t === 1 || Array.isArray(dr[k]) ? val : `=${colL(t)}${row}`, z, t === 1 || Array.isArray(dr[k]) ? 'input' : undefined);
    }
  });
  As.put(`A${8 + DR.length}`, 'Year 1 is the input and later years copy it. Type over any later year to change that year only.', null, 'note');
  As.sh.widths = { A: 40, B: 30, ...Object.fromEntries(Array.from({ length: n }, (_, t) => [colL(2 + t), 12])) };
  As.sh.freeze = { x: 2, y: 6 };
  const A = (k, t) => `Assumptions!${colL(1 + t)}${AR[k]}`;

  // --- Row maps (assigned before writing so the statements can reference each other) ---
  const rows = (keys, start = 6) => { const R = {}; keys.forEach((k, j) => { if (k) R[k] = start + j; }); return R; };
  const I = rows(['rev', 'growth', 'cogs', 'gp', 'gm', 'opex', 'ebitda', 'em', 'da', 'ebit', 'ebm', 'int', 'pbt', 'tax', 'etr', 'ni', 'nm']);
  const B = rows(['cash', 'noncash', 'tca', 'ppe', 'ta', null, 'cl', 'debt', 'other', 'tl', 'eq', 'tle', null, 'check']);
  const C = rows(['ni', 'da', 'dnca', 'dcl', 'cfo', null, 'capex', 'cfi', null, 'repay', 'div', 'cff', null, 'net', 'open', 'close', null, 'fcf', 'fcfm']);
  const book_ = { IS: addSheet(book, IS_, 'Income Statement', sub('actual and forecast')), BS: addSheet(book, BS_, 'Balance Sheet', sub('actual and forecast')), CF: addSheet(book, CF_, 'Cash Flow Statement', sub('actual and forecast')) };

  const head = (api) => {
    api.put('A5', 'USD', null, 'headL');
    allCols.forEach((c, j) => api.put(`${c}5`, colLabel(j), null, j < H ? 'headA' : 'headF'));
    api.put(`${aCol(0)}4`, withData ? 'Actual' : 'Base year', null, 'bandA');
    api.put(`${fCol(1)}4`, 'Forecast', null, 'bandF');
    api.sh.widths = { A: 42, ...Object.fromEntries(allCols.map((c) => [c, 13])) };
    api.sh.freeze = { x: 1, y: 5 };
  };
  const howCol = colL(H + n + 2); // one narrow spacer column after the last forecast year
  const write = (api, lines, R) => {
    head(api);
    api.put(`${howCol}5`, 'How the forecast is calculated', null, 'headL');
    api.sh.widths[howCol] = 58;
    api.sh.widths[colL(H + n + 1)] = 2;
    lines.forEach((L) => {
      const row = R[L.k];
      api.put(`A${row}`, L.label, null, L.bold ? 'labelB' : L.pct ? 'labelP' : 'label');
      for (let i = 0; i < H; i++) {
        const c = aCol(i), p = i ? aCol(i - 1) : null, x = L.a ? L.a(i, c, p) : null;
        if (x == null) continue;
        const s = typeof x === 'number' ? 'input' : /Historicals!/.test(x) ? 'link' : L.bold ? 'total' : L.pct ? 'pctRow' : undefined;
        api.put(`${c}${row}`, x, L.z || XF.money, s);
      }
      for (let t = 1; t <= n; t++) {
        const c = fCol(t), p = t === 1 ? lastA : fCol(t - 1);
        api.put(`${c}${row}`, L.f(t, c, p), L.z || XF.money, L.bold ? 'total' : L.pct ? 'pctRow' : undefined);
      }
      if (HOW[L.k + '@' + api.sh.name]) api.put(`${howCol}${row}`, HOW[L.k + '@' + api.sh.name], null, 'how');
    });
  };
  const IS = book_.IS, BS = book_.BS, CF = book_.CF;
  const has = (api, c, row) => api.has(`${c}${row}`);
  const div = (c, a, b) => `=IFERROR(${c}${a}/${c}${b},"")`;

  // --- Income statement ---
  write(IS, [
    { k: 'rev', label: 'Revenue', bold: true, a: (i) => (hist('revenue', i) ? `=${hist('revenue', i)}` : i === last ? op.rev : null), f: (t, c, p) => `=${p}${I.rev}*(1+${A('growth', t)})` },
    { k: 'growth', label: 'Revenue growth', pct: true, z: XF.pctS, a: (i, c, p) => (p && has(IS, c, I.rev) && has(IS, p, I.rev) ? `=IFERROR(${c}${I.rev}/${p}${I.rev}-1,"")` : null), f: (t, c, p) => `=${c}${I.rev}/${p}${I.rev}-1` },
    { k: 'cogs', label: 'Cost of goods sold', a: (i) => (hist('cogs', i) ? `=ABS(${hist('cogs', i)})` : null), f: (t, c) => `=${c}${I.rev}*(1-${A('gm', t)})` },
    { k: 'gp', label: 'Gross profit', bold: true, a: (i, c) => (has(IS, c, I.cogs) && has(IS, c, I.rev) ? `=${c}${I.rev}-${c}${I.cogs}` : hist('grossProfit', i) ? `=${hist('grossProfit', i)}` : null), f: (t, c) => `=${c}${I.rev}-${c}${I.cogs}` },
    { k: 'gm', label: 'Gross margin', pct: true, z: XF.pctS, a: (i, c) => (has(IS, c, I.gp) ? div(c, I.gp, I.rev) : null), f: (t, c) => div(c, I.gp, I.rev) },
    { k: 'opex', label: 'Operating expenses', a: (i) => (hist('opex', i) ? `=ABS(${hist('opex', i)})` : null), f: (t, c) => `=${c}${I.rev}*${A('opex', t)}` },
    { k: 'ebitda', label: 'EBITDA', bold: true, a: (i, c) => (has(IS, c, I.gp) && has(IS, c, I.opex) ? `=${c}${I.gp}-${c}${I.opex}` : hist('ebitda', i) ? `=${hist('ebitda', i)}` : null), f: (t, c) => `=${c}${I.gp}-${c}${I.opex}` },
    { k: 'em', label: 'EBITDA margin', pct: true, z: XF.pctS, a: (i, c) => (has(IS, c, I.ebitda) ? div(c, I.ebitda, I.rev) : null), f: (t, c) => div(c, I.ebitda, I.rev) },
    { k: 'da', label: 'Depreciation & amortization', a: (i) => (hist('da', i) ? `=ABS(${hist('da', i)})` : null), f: (t, c) => `=${c}${I.rev}*${A('da', t)}` },
    { k: 'ebit', label: 'EBIT (operating profit)', bold: true, a: (i, c) => (has(IS, c, I.ebitda) && has(IS, c, I.da) ? `=${c}${I.ebitda}-${c}${I.da}` : hist('ebit', i) ? `=${hist('ebit', i)}` : null), f: (t, c) => `=${c}${I.ebitda}-${c}${I.da}` },
    { k: 'ebm', label: 'EBIT margin', pct: true, z: XF.pctS, a: (i, c) => (has(IS, c, I.ebit) ? div(c, I.ebit, I.rev) : null), f: (t, c) => div(c, I.ebit, I.rev) },
    { k: 'int', label: 'Interest expense', a: (i) => (hist('interest', i) ? `=ABS(${hist('interest', i)})` : null), f: (t, c, p) => `=${ref(BS_, p + B.debt)}*${A('rate', t)}` },
    { k: 'pbt', label: 'Profit before tax', bold: true, a: (i, c) => (has(IS, c, I.ebit) && has(IS, c, I.int) ? `=${c}${I.ebit}-${c}${I.int}` : null), f: (t, c) => `=${c}${I.ebit}-${c}${I.int}` },
    { k: 'tax', label: 'Income tax', a: (i, c) => (has(IS, c, I.pbt) && hist('netIncome', i) ? `=${c}${I.pbt}-${c}${I.ni}` : null), f: (t, c) => `=MAX(0,${c}${I.pbt}*${A('tax', t)})` },
    { k: 'etr', label: 'Effective tax rate', pct: true, z: XF.pctS, a: (i, c) => (has(IS, c, I.tax) ? div(c, I.tax, I.pbt) : null), f: (t, c) => div(c, I.tax, I.pbt) },
    { k: 'ni', label: 'Net income', bold: true, a: (i) => (hist('netIncome', i) ? `=${hist('netIncome', i)}` : null), f: (t, c) => `=${c}${I.pbt}-${c}${I.tax}` },
    { k: 'nm', label: 'Net margin', pct: true, z: XF.pctS, a: (i, c) => (hist('netIncome', i) && has(IS, c, I.rev) ? div(c, I.ni, I.rev) : null), f: (t, c) => div(c, I.ni, I.rev) },
  ], I);

  // --- Balance sheet (last actual year completed with estimates when the upload is missing items) ---
  const OPK = { cash: 'cash', currentAssets: 'ca', totalAssets: 'ta', currentLiabilities: 'cl', debt: 'debt', equity: 'equity' };
  const base = (k, i) => (hist(k, i) ? `=${hist(k, i)}` : i === last ? Math.round(op[OPK[k]]) : null);
  const av = (i, ...ks) => ks.every((k) => base(k, i) != null);
  write(BS, [
    { k: 'cash', label: 'Cash and equivalents', a: (i) => base('cash', i), f: (t, c) => `=${ref(CF_, c + C.close)}` },
    { k: 'noncash', label: 'Non-cash current assets', a: (i, c) => (av(i, 'cash', 'currentAssets') ? `=${c}${B.tca}-${c}${B.cash}` : null), f: (t, c) => `=${ref(IS_, c + I.rev)}*${A('nca', t)}` },
    { k: 'tca', label: 'Total current assets', bold: true, a: (i) => base('currentAssets', i), f: (t, c) => `=${c}${B.cash}+${c}${B.noncash}` },
    { k: 'ppe', label: 'Non-current assets (PP&E and other)', a: (i, c) => (av(i, 'totalAssets', 'currentAssets') ? `=${c}${B.ta}-${c}${B.tca}` : null), f: (t, c, p) => `=${p}${B.ppe}-${ref(CF_, c + C.capex)}-${ref(IS_, c + I.da)}` },
    { k: 'ta', label: 'Total assets', bold: true, a: (i) => base('totalAssets', i), f: (t, c) => `=${c}${B.tca}+${c}${B.ppe}` },
    { k: 'cl', label: 'Current liabilities', a: (i) => base('currentLiabilities', i), f: (t, c) => `=${ref(IS_, c + I.rev)}*${A('cl', t)}` },
    { k: 'debt', label: 'Total debt', a: (i) => base('debt', i), f: (t, c, p) => `=${p}${B.debt}+${ref(CF_, c + C.repay)}` },
    { k: 'other', label: 'Other liabilities', a: (i, c) => (av(i, 'totalAssets', 'currentLiabilities', 'debt', 'equity') ? `=${c}${B.ta}-${c}${B.cl}-${c}${B.debt}-${c}${B.eq}` : null), f: (t, c, p) => `=${p}${B.other}` },
    { k: 'tl', label: 'Total liabilities', bold: true, a: (i, c) => (has(BS, c, B.other) ? `=${c}${B.cl}+${c}${B.debt}+${c}${B.other}` : null), f: (t, c) => `=${c}${B.cl}+${c}${B.debt}+${c}${B.other}` },
    { k: 'eq', label: "Shareholders' equity", bold: true, a: (i) => base('equity', i), f: (t, c, p) => `=${p}${B.eq}+${ref(IS_, c + I.ni)}+${ref(CF_, c + C.div)}` },
    { k: 'tle', label: 'Total liabilities & equity', bold: true, a: (i, c) => (has(BS, c, B.tl) ? `=${c}${B.tl}+${c}${B.eq}` : null), f: (t, c) => `=${c}${B.tl}+${c}${B.eq}` },
    { k: 'check', label: 'Balance check (assets – liabilities – equity)', pct: true, z: XF.check, a: (i, c) => (has(BS, c, B.tle) ? `=ROUND(${c}${B.ta}-${c}${B.tle},2)` : null), f: (t, c) => `=ROUND(${c}${B.ta}-${c}${B.tle},2)` },
  ], B);
  if (Object.keys(op.est).length) BS.put(`A${B.check + 2}`, `Blue figures in ${actualLabel(last)} are estimates because they were missing from the upload (${Object.keys(op.est).map((k) => LABELS[k]).join(', ')}).`, null, 'note');

  // --- Cash flow statement ---
  write(CF, [
    { k: 'ni', label: 'Net income', a: (i, c) => (has(IS, c, I.ni) ? `=${ref(IS_, c + I.ni)}` : null), f: (t, c) => `=${ref(IS_, c + I.ni)}` },
    { k: 'da', label: 'Depreciation & amortization', a: (i, c) => (has(IS, c, I.da) ? `=${ref(IS_, c + I.da)}` : null), f: (t, c) => `=${ref(IS_, c + I.da)}` },
    { k: 'dnca', label: '(Increase) / decrease in non-cash current assets', f: (t, c, p) => `=${ref(BS_, p + B.noncash)}-${ref(BS_, c + B.noncash)}` },
    { k: 'dcl', label: 'Increase / (decrease) in current liabilities', f: (t, c, p) => `=${ref(BS_, c + B.cl)}-${ref(BS_, p + B.cl)}` },
    { k: 'cfo', label: 'Cash from operating activities', bold: true, a: (i) => (hist('ocf', i) ? `=${hist('ocf', i)}` : null), f: (t, c) => `=SUM(${c}${C.ni}:${c}${C.dcl})` },
    { k: 'capex', label: 'Capital expenditure', a: (i) => (hist('capex', i) ? `=-ABS(${hist('capex', i)})` : null), f: (t, c) => `=-${ref(IS_, c + I.rev)}*${A('capex', t)}` },
    { k: 'cfi', label: 'Cash from investing activities', bold: true, a: (i, c) => (has(CF, c, C.capex) ? `=${c}${C.capex}` : null), f: (t, c) => `=${c}${C.capex}` },
    { k: 'repay', label: 'Debt repayment', f: (t, c, p) => `=-MIN(${ref(BS_, p + B.debt)},${A('repay', t)})` },
    { k: 'div', label: 'Dividends paid', f: (t, c) => `=-MAX(0,${c}${C.ni}*${A('payout', t)})` },
    { k: 'cff', label: 'Cash from financing activities', bold: true, f: (t, c) => `=${c}${C.repay}+${c}${C.div}` },
    { k: 'net', label: 'Net change in cash', bold: true, f: (t, c) => `=${c}${C.cfo}+${c}${C.cfi}+${c}${C.cff}` },
    { k: 'open', label: 'Opening cash', f: (t, c, p) => `=${ref(BS_, p + B.cash)}` },
    { k: 'close', label: 'Closing cash', bold: true, a: (i, c) => (has(BS, c, B.cash) ? `=${ref(BS_, c + B.cash)}` : null), f: (t, c) => `=${c}${C.open}+${c}${C.net}` },
    { k: 'fcf', label: 'Free cash flow (CFO – CapEx)', bold: true, a: (i, c) => (has(CF, c, C.cfo) && has(CF, c, C.capex) ? `=${c}${C.cfo}+${c}${C.capex}` : null), f: (t, c) => `=${c}${C.cfo}+${c}${C.capex}` },
    { k: 'fcfm', label: 'Free cash flow margin', pct: true, z: XF.pctS, a: (i, c) => (has(CF, c, C.fcf) && has(IS, c, I.rev) ? `=IFERROR(${c}${C.fcf}/${ref(IS_, c + I.rev)},"")` : null), f: (t, c) => `=IFERROR(${c}${C.fcf}/${ref(IS_, c + I.rev)},"")` },
  ], C);

  book.core = { H, n, last, aCol, fCol, lastA, allCols, colLabel, A, I, B, C, IS, BS, CF, op };
  return book.core;
}

/* ---------- Statement-driven sheets ---------- */
const DCF_HOW = { rev: 'From the income statement forecast', ebitda: 'From the income statement', ebit: 'EBITDA – D&A', tax: 'EBIT × tax rate', nopat: 'EBIT – taxes on EBIT', capex: 'From the cash flow statement',
  dnwc: 'Change in (non-cash current assets – current liabilities)', fcf: 'NOPAT + D&A – CapEx – increase in working capital', df: '1 ÷ (1 + WACC) ^ year', pv: 'Free cash flow × discount factor' };
function sheetDCF(book, v) {
  const K = book.core, { n, fCol, lastA, I, B, C } = K;
  const { put, section, header, assume, output, sh } = addSheet(book, 'DCF', 'DCF Valuation', sub('unlevered free cash flow from the linked statements'));
  const dc = (t) => colL(t), last = dc(n);
  let r = 5; section(r++, 'Valuation inputs', Math.max(n, 5));
  const W = assume(r++, 'WACC (discount rate)', v.wacc / 100, XF.pct, book.out.wacc ? 'WACC sheet result:' : null);
  if (book.out.wacc) put(`D${r - 1}`, `=${book.out.wacc.wacc}`, XF.pct2, 'link');
  const TG = assume(r++, 'Terminal growth', v.tg / 100, XF.pct);
  const ndLink = near(K.op.debt - K.op.cash, v.netDebt) ? `=${ref(BS_, lastA + B.debt)}-${ref(BS_, lastA + B.cash)}` : v.netDebt;
  const ND = assume(r++, 'Net debt (last actual year)', ndLink, XF.money, 'Total debt less cash');
  const SH = assume(r++, 'Shares outstanding', v.shares, XF.num), PR = assume(r++, 'Current share price (optional)', v.price || 0, XF.money2);
  r++; section(r++, 'Unlevered free cash flow (from the statements)', Math.max(n, 5));
  header(r++, ['Fiscal year', ...Array.from({ length: n }, (_, t) => yearLabel(t + 1))]);
  const R = {}; ['rev', 'ebitda', 'da', 'ebit', 'trate', 'tax', 'nopat', 'add', 'capex', 'dnwc', 'fcf', 'df', 'pv'].forEach((k) => (R[k] = r++));
  const lab = { rev: 'Revenue', ebitda: 'EBITDA', da: 'Depreciation & amortization', ebit: 'EBIT', trate: 'Tax rate', tax: 'Less: taxes on EBIT', nopat: 'NOPAT', add: 'Add: D&A', capex: 'Less: capital expenditure', dnwc: 'Less: increase in net working capital', fcf: 'Unlevered free cash flow', df: 'Discount factor', pv: 'Present value of FCF' };
  Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['ebit', 'nopat', 'fcf', 'pv'].includes(k) ? 'labelB' : 'label'));
  for (let t = 1; t <= n; t++) {
    const c = dc(t), s = fCol(t), sp = t === 1 ? lastA : fCol(t - 1);
    put(`${c}${R.rev}`, `=${ref(IS_, s + I.rev)}`, XF.money);
    put(`${c}${R.ebitda}`, `=${ref(IS_, s + I.ebitda)}`, XF.money);
    put(`${c}${R.da}`, `=${ref(IS_, s + I.da)}`, XF.money);
    put(`${c}${R.ebit}`, `=${ref(IS_, s + I.ebit)}`, XF.money, 'total');
    put(`${c}${R.trate}`, `=${K.A('tax', t)}`, XF.pct, 'pctRow');
    put(`${c}${R.tax}`, `=${c}${R.ebit}*${c}${R.trate}`, XF.money);
    put(`${c}${R.nopat}`, `=${c}${R.ebit}-${c}${R.tax}`, XF.money, 'total');
    put(`${c}${R.add}`, `=${c}${R.da}`, XF.money);
    put(`${c}${R.capex}`, `=-${ref(CF_, s + C.capex)}`, XF.money);
    put(`${c}${R.dnwc}`, `=(${ref(BS_, s + B.noncash)}-${ref(BS_, s + B.cl)})-(${ref(BS_, sp + B.noncash)}-${ref(BS_, sp + B.cl)})`, XF.money);
    put(`${c}${R.fcf}`, `=${c}${R.nopat}+${c}${R.add}-${c}${R.capex}-${c}${R.dnwc}`, XF.money, 'total');
    put(`${c}${R.df}`, `=1/(1+${W})^${t}`, XF.df);
    if (t === n) Object.entries(DCF_HOW).forEach(([k, txt]) => put(`${colL(n + 1)}${R[k]}`, txt, null, 'how'));
    put(`${c}${R.pv}`, `=${c}${R.fcf}*${c}${R.df}`, XF.money, 'total');
  }
  r++; section(r++, 'Valuation', Math.max(n, 5));
  const V = {};
  const sumRow = r; V.sum = output(r++, 'Sum of PV of free cash flows', `=SUM(B${R.pv}:${last}${R.pv})`, XF.money, 'calc', 'Add up the present value of each forecast year');
  const tvRow = r; V.tv = output(r++, 'Terminal value (Gordon growth, undiscounted)', `=${last}${R.fcf}*(1+${TG})/(${W}-${TG})`, XF.money, 'calc', 'Final-year FCF × (1 + terminal growth) ÷ (WACC – terminal growth)');
  const pvtvRow = r; V.pvtv = output(r++, 'PV of terminal value', `=B${tvRow}*${last}${R.df}`, XF.money, 'calc', 'Terminal value × final-year discount factor');
  const evRow = r; V.ev = output(r++, 'Enterprise value', `=B${sumRow}+B${pvtvRow}`, XF.money, 'output', 'PV of forecast cash flows + PV of terminal value');
  const ndRow = r; output(r++, 'Less: net debt', `=${ND}`, XF.money, 'calc');
  const eqRow = r; V.eq = output(r++, 'Equity value', `=B${evRow}-B${ndRow}`, XF.money, 'output', 'Enterprise value – net debt');
  const psRow = r; V.ps = output(r++, 'Value per share', `=B${eqRow}/${SH}`, XF.money2, 'output', 'Equity value ÷ shares outstanding');
  V.up = output(r++, 'Upside / (downside) vs current price', `=IF(${PR}>0,B${psRow}/${PR}-1,"n/a")`, XF.pctS, 'calc');
  output(r++, 'Terminal value as % of EV', `=B${pvtvRow}/B${evRow}`, XF.pct, 'calc');
  output(r++, 'Implied EV / EBITDA (last actual year)', `=IFERROR(B${evRow}/${ref(IS_, lastA + I.ebitda)},"n/a")`, XF.mult, 'calc');
  output(r++, 'Implied EV / EBITDA (first forecast year)', `=IFERROR(B${evRow}/B${R.ebitda},"n/a")`, XF.mult, 'calc');
  output(r++, 'Implied exit multiple (TV / final-year EBITDA)', `=IFERROR(B${tvRow}/${last}${R.ebitda},"n/a")`, XF.mult, 'calc');
  const grid = (title, perShare) => {
    r++; section(r++, title, 5);
    const hr = r++;
    put(`A${hr}`, 'WACC ↓   Terminal growth →', null, 'headL');
    [-0.01, -0.005, 0, 0.005, 0.01].forEach((d, j) => put(`${colL(j + 1)}${hr}`, `=${TG}+${d}`, XF.pct, 'head'));
    [-0.02, -0.01, 0, 0.01, 0.02].forEach((d) => {
      const row = r++;
      put(`A${row}`, `=${W}+${d}`, XF.pct, 'headL');
      for (let j = 1; j <= 5; j++) {
        const c = colL(j), ev = `NPV($A${row},$B$${R.fcf}:$${last}$${R.fcf})+$${last}$${R.fcf}*(1+${c}$${hr})/($A${row}-${c}$${hr})/(1+$A${row})^${n}`;
        put(`${c}${row}`, perShare ? `=(${ev}-${ND})/${SH}` : `=${ev}`, perShare ? XF.money2 : XF.money, d === 0 && j === 3 ? 'output' : undefined);
      }
    });
  };
  grid('Sensitivity: enterprise value', false);
  grid('Sensitivity: value per share', true);
  sh.widths = { A: 46, ...Object.fromEntries(Array.from({ length: Math.max(n, 5) }, (_, i) => [colL(i + 1), 14])), [colL(n + 1)]: 50 };
  book.out.dcf = V;
}

function sheetRatios(book) {
  const K = book.core, { I, B, C, allCols, colLabel, H, IS, BS, CF } = K;
  const { put, section, sh } = addSheet(book, 'Ratios', 'Ratio Analysis', sub('30+ ratios, DuPont, Altman Z″ and Piotroski F-score (actual and forecast)'));
  const histKey = { ar: 'receivables', inv: 'inventory', ap: 'payables', re: 'retainedEarnings' };
  const stmt = { rev: [IS, IS_, I.rev], cogs: [IS, IS_, I.cogs], gp: [IS, IS_, I.gp], ebitda: [IS, IS_, I.ebitda], ebit: [IS, IS_, I.ebit], da: [IS, IS_, I.da], int: [IS, IS_, I.int], ni: [IS, IS_, I.ni],
    cash: [BS, BS_, B.cash], ca: [BS, BS_, B.tca], ta: [BS, BS_, B.ta], cl: [BS, BS_, B.cl], tl: [BS, BS_, B.tl], debt: [BS, BS_, B.debt], eq: [BS, BS_, B.eq], ocf: [CF, CF_, C.cfo] };
  // Cell reference for an input in a given column (j = index into allCols); null when the figure isn't available
  const refAt = (key, j) => {
    if (j < 0) return null;
    const c = allCols[j];
    if (key === 'tax') return K.A('tax', 1);
    if (key === 'capex') return CF.has(`${c}${C.capex}`) ? `(-${ref(CF_, c + C.capex)})` : null;
    if (histKey[key]) { const h = book.hist?.[histKey[key]]; return h && j < H && S(histKey[key])[j] != null ? `Historicals!${c}${h.row}` : null; }
    const [api, name, row] = stmt[key];
    return api.has(`${c}${row}`) ? ref(name, c + row) : null;
  };
  const last = allCols.length, bmCol = colL(last + 2), howCol = colL(last + 3), bm = chosenBenchmark();
  section(5, 'Key ratios', last + 3);
  put('A6', 'Ratio', null, 'headL');
  allCols.forEach((c, j) => put(`${c}6`, colLabel(j), null, j < H ? 'headA' : 'headF'));
  put(`${bmCol}6`, 'Industry', null, 'head'); put(`${howCol}6`, 'How it’s calculated', null, 'headL');
  let r = 7;
  RATIO_CATS.slice(1).forEach((cat) => {
    put(`A${r++}`, cat, null, 'labelB');
    RATIO_DEFS.filter((d) => d.cat === cat).forEach((d) => {
      const z = d.fmt === 'pct' ? XF.pctS : d.fmt === 'x' ? XF.mult : d.fmt === 'days' ? '0 "days"' : XF.money;
      put(`A${r}`, d.name, null, 'label');
      allCols.forEach((c, j) => { const f = treeToExcel(d.f, (k, prior) => refAt(k, prior ? j - 1 : j)); if (f) put(`${c}${r}`, `=IFERROR(${f},"")`, z); });
      if (bm && d.bench) put(`${bmCol}${r}`, bm[d.bench], z, 'input');
      put(`${howCol}${r}`, d.how, null, 'how');
      r++;
    });
  });
  if (bm) put(`A${r++}`, `Industry column: ${bm.name} (indicative figures; confirm against the latest source).`, null, 'note');

  // DuPont analysis
  r++; section(r++, 'DuPont analysis of return on equity', last + 3);
  const du = [['Net margin', ['/', 'ni', 'rev'], XF.pctS, 'Net income ÷ revenue'], ['× Asset turnover', ['/', 'rev', 'ta'], XF.mult, 'Revenue ÷ total assets'], ['× Equity multiplier', ['/', 'ta', 'eq'], XF.mult, 'Total assets ÷ equity']];
  const r0 = r;
  du.forEach(([label, t, z, how]) => {
    put(`A${r}`, label, null, 'label');
    allCols.forEach((c, j) => { const f = treeToExcel(t, (k, prior) => refAt(k, prior ? j - 1 : j)); if (f) put(`${c}${r}`, `=IFERROR(${f},"")`, z); });
    put(`${howCol}${r}`, how, null, 'how'); r++;
  });
  put(`A${r}`, 'Return on equity (product of the three)', null, 'labelB');
  allCols.forEach((c) => put(`${c}${r}`, `=IFERROR(${c}${r0}*${c}${r0 + 1}*${c}${r0 + 2},"")`, XF.pctS, 'total'));
  put(`${howCol}${r}`, 'Net margin × asset turnover × equity multiplier', null, 'how'); r++;

  // Altman Z''-score (actual years)
  const actual = allCols.slice(0, H);
  r++; section(r++, 'Altman Z″-score (private and non-manufacturing companies)', last + 3);
  const zr = r, zParts = [['Working capital ÷ total assets', ['/', ['-', 'ca', 'cl'], 'ta'], 6.56], ['Retained earnings ÷ total assets', ['/', 're', 'ta'], 3.26], ['EBIT ÷ total assets', ['/', 'ebit', 'ta'], 6.72], ['Equity ÷ total liabilities', ['/', 'eq', 'tl'], 1.05]];
  zParts.forEach(([label, t, w], k) => {
    put(`A${r}`, `${label} (× ${w})`, null, 'label');
    actual.forEach((c, j) => {
      let f = treeToExcel(t, (key, prior) => refAt(key, prior ? j - 1 : j));
      if (!f && k === 1) f = treeToExcel(['/', 'eq', 'ta'], (key) => refAt(key, j)); // equity stands in when retained earnings aren't reported
      if (f) put(`${c}${r}`, `=IFERROR(${f},"")`, '0.000');
    });
    r++;
  });
  put(`A${r}`, 'Z″-score', null, 'labelB');
  actual.forEach((c) => put(`${c}${r}`, `=IFERROR(6.56*${c}${zr}+3.26*${c}${zr + 1}+6.72*${c}${zr + 2}+1.05*${c}${zr + 3},"")`, XF.dec, 'output'));
  put(`${howCol}${r}`, 'Above 2.6 safe · 1.1 to 2.6 grey zone · below 1.1 distress', null, 'how'); const zRow = r++;
  put(`A${r}`, 'Zone', null, 'label');
  actual.forEach((c) => put(`${c}${r}`, `=IF(${c}${zRow}="","",IF(${c}${zRow}>2.6,"Safe",IF(${c}${zRow}>=1.1,"Grey zone","Distress")))`, null));
  r++;
  if (!book.hist?.retainedEarnings) put(`A${r++}`, 'Retained earnings were not in the upload, so book equity is used in their place (a common approximation).', null, 'note');

  // Piotroski F-score (actual years from the second one)
  r++; section(r++, 'Piotroski F-score (1 = pass, 0 = fail)', last + 3);
  const fr = r;
  const tests = [
    ['Positive return on assets', (g) => `IF(${g('ni')}/${g('ta')}>0,1,0)`, ['ni', 'ta']],
    ['Positive operating cash flow', (g) => `IF(${g('ocf')}>0,1,0)`, ['ocf']],
    ['Return on assets improved', (g, p) => `IF(${g('ni')}/${g('ta')}>${p('ni')}/${p('ta')},1,0)`, ['ni', 'ta'], true],
    ['Cash flow exceeds net income', (g) => `IF(${g('ocf')}>${g('ni')},1,0)`, ['ocf', 'ni']],
    ['Leverage (debt ÷ assets) fell', (g, p) => `IF(${g('debt')}/${g('ta')}<${p('debt')}/${p('ta')},1,0)`, ['debt', 'ta'], true],
    ['Current ratio improved', (g, p) => `IF(${g('ca')}/${g('cl')}>${p('ca')}/${p('cl')},1,0)`, ['ca', 'cl'], true],
    ['Gross margin improved', (g, p) => `IF(${g('gp')}/${g('rev')}>${p('gp')}/${p('rev')},1,0)`, ['gp', 'rev'], true],
    ['Asset turnover improved', (g, p) => `IF(${g('rev')}/${g('ta')}>${p('rev')}/${p('ta')},1,0)`, ['rev', 'ta'], true],
  ];
  tests.forEach(([label, f, keys, needsPrior]) => {
    put(`A${r}`, label, null, 'label');
    actual.forEach((c, j) => {
      if (j === 0) return;
      const g = (k) => refAt(k, j), p = (k) => refAt(k, j - 1);
      if (keys.every((k) => g(k) && (!needsPrior || p(k)))) put(`${c}${r}`, `=IFERROR(${f(g, p)},"")`, '0');
    });
    r++;
  });
  put(`A${r}`, 'F-score (tests passed)', null, 'labelB');
  actual.forEach((c, j) => { if (j > 0) put(`${c}${r}`, `=SUM(${c}${fr}:${c}${r - 1})`, '0', 'output'); });
  put(`${howCol}${r}`, `Out of ${tests.length} tests (share issuance isn’t in the statements). 7+ strong, 3 or less weak.`, null, 'how'); r++;

  sh.widths = { A: 40, ...Object.fromEntries(allCols.map((c) => [c, 12])), [colL(last + 1)]: 2, [bmCol]: 12, [howCol]: 52 };
  sh.freeze = { x: 1, y: 6 };
}

/* ---------- Dashboard: KPI tiles, native charts, summaries ---------- */
function sheetDashboard(dash, book) {
  const K = book.core, { I, B, C, allCols, colLabel, H, n, lastA, IS, CF } = K, { put, sh } = dash;
  const lastF = K.fCol(n);
  const tiles = [];
  const fcfCol = CF.has(`${lastA}${C.fcf}`) ? lastA : K.fCol(1);
  tiles.push([`Revenue ${colLabel(H - 1)}`, `=${ref(IS_, lastA + I.rev)}`, XF.money, IS.has(`${lastA}${I.growth}`) ? `=${ref(IS_, lastA + I.growth)}` : null, '"YoY growth "0.0%;"YoY growth "(0.0%)']);
  tiles.push([`EBITDA ${colLabel(H - 1)}`, IS.has(`${lastA}${I.ebitda}`) ? `=${ref(IS_, lastA + I.ebitda)}` : `=${ref(IS_, K.fCol(1) + I.ebitda)}`, XF.money, `=${ref(IS_, (IS.has(`${lastA}${I.em}`) ? lastA : K.fCol(1)) + I.em)}`, '"Margin "0.0%']);
  tiles.push([`Net income ${colLabel(H - 1)}`, IS.has(`${lastA}${I.ni}`) ? `=${ref(IS_, lastA + I.ni)}` : `=${ref(IS_, K.fCol(1) + I.ni)}`, XF.money, `=${ref(IS_, (IS.has(`${lastA}${I.nm}`) ? lastA : K.fCol(1)) + I.nm)}`, '"Net margin "0.0%']);
  tiles.push([`Free cash flow ${colLabel(allCols.indexOf(fcfCol))}`, `=${ref(CF_, fcfCol + C.fcf)}`, XF.money, `=${ref(CF_, fcfCol + C.fcfm)}`, '"FCF margin "0.0%']);
  tiles.push([`Revenue ${colLabel(H + n - 1)} (forecast)`, `=${ref(IS_, lastF + I.rev)}`, XF.money, `=(${ref(IS_, lastF + I.rev)}/${ref(IS_, lastA + I.rev)})^(1/${n})-1`, '"CAGR "0.0%']);
  tiles.push([`Net debt ${colLabel(H - 1)}`, `=${ref(BS_, lastA + B.debt)}-${ref(BS_, lastA + B.cash)}`, XF.money, IS.has(`${lastA}${I.ebitda}`) ? `=IFERROR((${ref(BS_, lastA + B.debt)}-${ref(BS_, lastA + B.cash)})/${ref(IS_, lastA + I.ebitda)},0)` : null, '0.00"x EBITDA"']);
  if (book.out.dcf) tiles.push(['DCF enterprise value', `=${book.out.dcf.ev}`, XF.money, `=${book.out.dcf.ps}`, '"Per share $"#,##0.00']);
  else tiles.push([`Closing cash ${colLabel(H + n - 1)}`, `=${ref(BS_, lastF + B.cash)}`, XF.money, null, null]);
  if (book.out.lbo) tiles.push(['LBO sponsor IRR', `=${book.out.lbo.irr}`, XF.pct, `=${book.out.lbo.moic}`, '0.00"x MOIC"']);
  else if (book.out.wacc) tiles.push(['WACC', `=${book.out.wacc.wacc}`, XF.pct2, `=${book.out.wacc.ke}`, '"Cost of equity "0.0%']);
  else tiles.push([`Equity ${colLabel(H + n - 1)}`, `=${ref(BS_, lastF + B.eq)}`, XF.money, null, null]);

  // 4 tiles per row, each 3 columns wide (B:D, E:G, H:J, K:M)
  tiles.forEach(([label, f, z, subF, subZ], k) => {
    const row = 5 + Math.floor(k / 4) * 4, c0 = 1 + (k % 4) * 3, c1 = c0 + 2, L = colL(c0), R = colL(c1);
    put(`${L}${row}`, label, null, 'tileLabel');
    put(`${L}${row + 1}`, f, z, 'tileValue');
    put(`${L}${row + 2}`, subF ?? '', subZ, 'tileSub');
    [row, row + 1, row + 2].forEach((rr) => { sh.merges.push(`${L}${rr}:${R}${rr}`); for (let c = c0 + 1; c <= c1; c++) put(`${colL(c)}${rr}`, '', null, rr === row ? 'tileLabel' : rr === row + 1 ? 'tileValue' : 'tileSub'); });
  });
  sh.heights[6] = sh.heights[10] = 30;

  // Financial summary table (also the data behind the charts)
  const tRow = 48;
  put(`B${tRow - 1}`, 'Financial summary', null, 'section');
  sh.sections.push([tRow - 1, allCols.length + 2]);
  put(`B${tRow}`, 'USD', null, 'headL');
  const sc = (j) => colL(2 + j);
  allCols.forEach((c, j) => put(`${sc(j)}${tRow}`, colLabel(j), null, j < H ? 'headA' : 'headF'));
  const lines = [['Revenue', IS_, I.rev, XF.money], ['EBITDA', IS_, I.ebitda, XF.money], ['Net income', IS_, I.ni, XF.money], ['Free cash flow', CF_, C.fcf, XF.money], ['Cash', BS_, B.cash, XF.money], ['Total debt', BS_, B.debt, XF.money], ['Gross margin', IS_, I.gm, XF.pctS], ['EBITDA margin', IS_, I.em, XF.pctS], ['Net margin', IS_, I.nm, XF.pctS]];
  const byName = Object.fromEntries(book.sheets.map((s) => [s.name, s]));
  const rowOf = {};
  lines.forEach(([label, s, row, z], k) => {
    const rr = tRow + 1 + k;
    rowOf[label] = rr;
    put(`B${rr}`, label, null, k < 6 ? 'label' : 'labelP');
    allCols.forEach((c, j) => { if (byName[s].cells[`${c}${row}`]) put(`${sc(j)}${rr}`, `=${ref(s, c + row)}`, z, k < 6 ? undefined : 'pctRow'); });
  });

  // Valuation summary (headline outputs of the models in this workbook)
  const vals = [];
  const O = book.out;
  if (O.dcf) vals.push(['DCF enterprise value', O.dcf.ev, XF.money], ['DCF equity value', O.dcf.eq, XF.money], ['DCF value per share', O.dcf.ps, XF.money2]);
  if (O.comps) vals.push(['Comparables: implied equity value (median)', O.comps.median, XF.money]);
  if (O.wacc) vals.push(['WACC', O.wacc.wacc, XF.pct2]);
  if (O.lbo) vals.push(['LBO sponsor IRR', O.lbo.irr, XF.pct], ['LBO MOIC', O.lbo.moic, XF.mult]);
  if (O.mc) vals.push(['Monte Carlo P50 enterprise value', O.mc.p50, XF.money]);
  if (O.npv) vals.push(['Project NPV (NPV / IRR sheet)', O.npv.npv, XF.money], ['Project IRR', O.npv.irr, XF.pct]);
  if (O.feas) vals.push(['Feasibility NPV', O.feas.npv, XF.money], ['Feasibility IRR', O.feas.irr, XF.pct]);
  if (O.be) vals.push(['Break-even units', O.be.units, XF.num]);
  if (O.loan) vals.push(['Loan payment per period', O.loan.pmt, XF.money2]);
  if (O.scenario) vals.push(['Scenario probability-weighted EV', O.scenario.wev, XF.money]);
  let vr = tRow + lines.length + 3;
  if (vals.length) {
    put(`B${vr}`, 'Valuation & returns summary', null, 'section');
    sh.sections.push([vr, 6]); vr++;
    vals.forEach(([label, rf, z]) => { put(`B${vr}`, label, null, 'label'); sh.merges.push(`B${vr}:E${vr}`); put(`F${vr}`, `=${rf}`, z, 'output'); vr++; });
  }

  // Model health: accuracy, balance check and the market inputs used
  vr += 1;
  put(`B${vr}`, 'Model health & inputs', null, 'section'); sh.sections.push([vr, 6]); vr++;
  const chk = `${ref(BS_, `${K.fCol(1)}${B.check}`)}:${ref(BS_, `${K.fCol(n)}${B.check}`).split('!')[1]}`;
  const health = [['Balance sheet balances in every forecast year', `=IF(AND(MAX(${chk})<1,MIN(${chk})>-1),"Yes","Check")`, null]];
  if (O.accuracy) health.push(['Forecast accuracy (backtest on your history)', `=${O.accuracy.overall}`, XF.pct]);
  const rf = marketRf();
  health.push(['Risk-free rate source', rf ? `${rf.value.toFixed(2)}% · ${rf.source} (${rf.date})` : '4.20% default (set a live rate on the Indicators tab)', null]);
  const bm = chosenBenchmark();
  if (bm) health.push(['Industry benchmark', `${bm.name} (indicative)`, null]);
  health.forEach(([label, f, z]) => { put(`B${vr}`, label, null, 'label'); sh.merges.push(`B${vr}:E${vr}`); put(`F${vr}`, f, z, typeof f === 'string' && f[0] === '=' ? 'output' : 'label'); vr++; });

  // Company vs industry benchmark (latest actual year)
  if (bm && hasData()) {
    vr += 1;
    put(`B${vr}`, `Company vs industry: ${bm.name}`, null, 'section'); sh.sections.push([vr, 6]); vr++;
    ['Metric', '', '', 'Company', 'Industry', 'Difference'].forEach((h, j) => { if (h || j === 0) put(`${colL(1 + j)}${vr}`, h, null, j === 0 ? 'headL' : 'head'); });
    sh.merges.push(`B${vr}:D${vr}`); vr++;
    const la = lastA, cmp = [['Gross margin', `=${ref(IS_, la + I.gm)}`, bm.gm], ['EBITDA margin', `=${ref(IS_, la + I.em)}`, bm.em], ['Net margin', `=${ref(IS_, la + I.nm)}`, bm.nm],
      ['Return on equity', `=IFERROR(${ref(IS_, la + I.ni)}/${ref(BS_, la + B.eq)},"")`, bm.roe], ['Debt / equity', `=IFERROR(${ref(BS_, la + B.debt)}/${ref(BS_, la + B.eq)},"")`, bm.de]];
    cmp.forEach(([label, f, ind], k) => {
      const z = k < 4 ? XF.pctS : XF.mult;
      put(`B${vr}`, label, null, 'label'); sh.merges.push(`B${vr}:D${vr}`);
      put(`E${vr}`, f, z); put(`F${vr}`, ind, z, 'input'); put(`G${vr}`, `=IFERROR(E${vr}-F${vr},"")`, k < 4 ? '+0.0%;-0.0%' : '+0.00"x";-0.00"x"', 'output'); vr++;
    });
    put(`B${vr}`, 'Industry figures are indicative (rounded from published industry datasets). Confirm against the latest source before relying on them.', null, 'note');
  }

  // Native charts (anchored in the 2×2 grid between the tiles and the tables)
  const span = `$${sc(0)}$`, endC = `$${sc(allCols.length - 1)}$`;
  const rng = (label) => `Dashboard!${span}${rowOf[label]}:${endC}${rowOf[label]}`;
  const cats = `Dashboard!${span}${tRow}:${endC}${tRow}`;
  const ser = (label, color) => ({ name: label, ref: rng(label), color });
  book.charts.push(
    { type: 'bar', title: 'Revenue, EBITDA and net income', cats, series: [ser('Revenue', '2F6DB5'), ser('EBITDA', '2E8B57'), ser('Net income', 'C19A3F')], fmt: '#,##0', at: [1, 13, 7, 30] },
    { type: 'line', title: 'Profit margins', cats, series: [ser('Gross margin', '2F6DB5'), ser('EBITDA margin', '2E8B57'), ser('Net margin', 'C19A3F')], fmt: '0%', at: [7, 13, 13, 30] },
    { type: 'bar', title: 'Free cash flow', cats, series: [ser('Free cash flow', '3A9CA6')], fmt: '#,##0', at: [1, 30, 7, 46] },
    { type: 'line', title: 'Cash vs total debt', cats, series: [ser('Cash', '2E8B57'), ser('Total debt', 'C0392B')], fmt: '#,##0', at: [7, 30, 13, 46] },
  );
  put('B12', 'Charts update automatically from the statements. Forecast columns follow the Assumptions sheet.', null, 'note');
  sh.widths = { A: 2, B: 24, ...Object.fromEntries(Array.from({ length: Math.max(14, allCols.length + 2) }, (_, i) => [colL(i + 2), 12])) };
}

/* ---------- Other model sheets ---------- */
const XL_MODELS = {
  dcf: { sheet: 'DCF', desc: 'DCF valuation from the linked statements: FCF build, valuation bridge, implied multiples, two sensitivity tables', build: (book, v) => sheetDCF(book, v) },
  ratios: { sheet: 'Ratios', desc: '30+ ratios by category with industry column, DuPont, Altman Z″ and Piotroski F-score (actual and forecast)', build: (book) => sheetRatios(book) },
  forecast: { sheet: 'Income Statement', desc: 'Linked 3-statement model (always included)', core: true, build() {} },

  wacc: { sheet: 'WACC', desc: 'Cost of capital: CAPM build-up, unlevered beta and a beta × market risk premium sensitivity', build(book, v) {
    const { put, section, assume, output, sh } = addSheet(book, 'WACC', 'Weighted Average Cost of Capital', sub('discount rate'));
    const K = book.core;
    let r = 5; section(r++, 'Inputs', 5);
    const E = assume(r++, 'Equity value', near(K.op.equity, v.E) ? `=${ref(BS_, K.lastA + K.B.eq)}` : v.E, XF.money, 'Book value from the balance sheet; replace with market value if listed');
    const D = assume(r++, 'Debt value', near(K.op.debt, v.D) ? `=${ref(BS_, K.lastA + K.B.debt)}` : v.D, XF.money);
    const rfSrc = marketRf(), RF = assume(r++, 'Risk-free rate', v.rf / 100, XF.pct2, rfSrc ? `${rfSrc.source}, ${rfSrc.date}` : 'Default; set a live rate on the Indicators tab'), BE = assume(r++, 'Levered beta', v.beta, XF.dec), MRP = assume(r++, 'Market risk premium', v.mrp / 100, XF.pct2);
    const SP = assume(r++, 'Size / specific risk premium', v.sp / 100, XF.pct2), KD = assume(r++, 'Pre-tax cost of debt', v.kd / 100, XF.pct2), T = assume(r++, 'Tax rate', v.tax / 100, XF.pct);
    r++; section(r++, 'Calculation', 5);
    const ke = output(r++, 'Cost of equity (CAPM)', `=${RF}+${BE}*${MRP}+${SP}`, XF.pct2, 'calc', 'Risk-free rate + beta × market risk premium + size premium');
    const kd = output(r++, 'After-tax cost of debt', `=${KD}*(1-${T})`, XF.pct2, 'calc', 'Pre-tax cost of debt × (1 – tax rate)');
    const we = output(r++, 'Equity weight (E / V)', `=${E}/(${E}+${D})`, XF.pct, 'calc');
    const wd = output(r++, 'Debt weight (D / V)', `=${D}/(${E}+${D})`, XF.pct, 'calc');
    output(r++, 'Debt / equity', `=IFERROR(${D}/${E},0)`, XF.mult, 'calc');
    output(r++, 'Unlevered (asset) beta', `=IFERROR(${BE}/(1+(1-${T})*${D}/${E}),"")`, XF.dec, 'calc');
    const w = output(r++, 'WACC', `=${we}*${ke}+${wd}*${kd}`, XF.pct2, 'output', 'E/V × cost of equity + D/V × after-tax cost of debt');
    r++; section(r++, 'Sensitivity: WACC by beta ↓ and market risk premium →', 5);
    const hr = r++;
    put(`A${hr}`, 'Beta \\ MRP', null, 'headL');
    [-0.01, -0.005, 0, 0.005, 0.01].forEach((d, j) => put(`${colL(j + 1)}${hr}`, `=${MRP}+${d}`, XF.pct, 'head'));
    [-0.2, -0.1, 0, 0.1, 0.2].forEach((d) => {
      const row = r++;
      put(`A${row}`, `=${BE}+${d}`, XF.dec, 'headL');
      for (let j = 1; j <= 5; j++) put(`${colL(j)}${row}`, `=${we}*(${RF}+$A${row}*${colL(j)}$${hr}+${SP})+${wd}*${kd}`, XF.pct2, d === 0 && j === 3 ? 'output' : undefined);
    });
    sh.widths = { A: 40, B: 16, C: 14, D: 14, E: 14, F: 14 };
    book.out.wacc = { wacc: w, ke };
  } },

  lbo: { sheet: 'LBO', desc: 'Leveraged buyout: sources & uses, debt schedule, credit stats, returns attribution, IRR sensitivity', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'LBO', 'Leveraged Buyout', sub('sponsor returns'));
    const K = book.core, n = Math.min(10, Math.max(1, Math.round(v.years))), C = (t) => colL(t + 1), last = C(n);
    let r = 5; section(r++, 'Assumptions', n + 1);
    const EB = assume(r++, 'Entry EBITDA (last actual year)', linkOr(book, 'ebitda', v.ebitda), XF.money), ENTRY = assume(r++, 'Entry multiple (x EBITDA)', v.entry, XF.mult);
    const LEV = assume(r++, 'Debt (x EBITDA)', v.lev, XF.mult), RATE = assume(r++, 'Interest rate', v.rate / 100, XF.pct), G = assume(r++, 'EBITDA growth', v.g / 100, XF.pct);
    const EXIT = assume(r++, 'Exit multiple (x EBITDA)', v.exit, XF.mult), TAX = assume(r++, 'Tax rate', v.tax / 100, XF.pct);
    const CX = assume(r++, 'CapEx % of EBITDA', v.capex / 100, XF.pct), SW = assume(r++, 'Cash used to repay debt', v.sweep / 100, XF.pct), FEES = assume(r++, 'Transaction fees % of EV', v.fees / 100, XF.pct);
    r++; section(r++, 'Sources & uses', n + 1);
    header(r++, ['Uses', 'Amount', 'Sources', 'Amount']);
    const su = r, tot = su + 2;
    put(`A${su}`, 'Purchase price (EV)', null, 'label'); put(`B${su}`, `=${EB}*${ENTRY}`, XF.money); put(`C${su}`, 'Debt', null, 'labelI'); put(`D${su}`, `=${EB}*${LEV}`, XF.money);
    put(`A${su + 1}`, 'Transaction fees', null, 'label'); put(`B${su + 1}`, `=B${su}*${FEES}`, XF.money); put(`C${su + 1}`, 'Sponsor equity', null, 'labelI'); put(`D${su + 1}`, `=B${tot}-D${su}`, XF.money, 'output');
    put(`A${tot}`, 'Total uses', null, 'labelB'); put(`B${tot}`, `=B${su}+B${su + 1}`, XF.money, 'total'); put(`C${tot}`, 'Total sources', null, 'labelBI'); put(`D${tot}`, `=D${su}+D${su + 1}`, XF.money, 'total');
    const EQ0 = `$D$${su + 1}`, DEBT0 = `$D$${su}`;
    r = tot + 1;
    r++; section(r++, 'Debt schedule and credit statistics', n + 1);
    header(r++, ['', 'Entry', ...Array.from({ length: n }, (_, t) => yearLabel(t + 1))]);
    const R = {}; ['eb', 'int', 'cx', 'tax', 'fcf', 'rep', 'debt', 'cash', 'lev', 'cov', 'eqv'].forEach((k) => (R[k] = r++));
    const lab = { eb: 'EBITDA', int: 'Interest', cx: 'Capital expenditure', tax: 'Taxes', fcf: 'Free cash flow', rep: 'Debt repayment', debt: 'Debt balance (end of year)', cash: 'Cash balance', lev: 'Debt / EBITDA', cov: 'EBITDA / interest', eqv: 'Equity value at exit multiple' };
    Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['fcf', 'debt', 'eqv'].includes(k) ? 'labelB' : ['lev', 'cov'].includes(k) ? 'labelP' : 'label'));
    put(`B${R.eb}`, `=${EB}`, XF.money); put(`B${R.debt}`, `=${DEBT0}`, XF.money); put(`B${R.cash}`, 0, XF.money); put(`B${R.eqv}`, `=${EQ0}`, XF.money); put(`B${R.lev}`, `=B${R.debt}/B${R.eb}`, XF.mult, 'pctRow');
    for (let t = 1; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      put(`${c}${R.eb}`, `=${p}${R.eb}*(1+${G})`, XF.money);
      put(`${c}${R.int}`, `=${p}${R.debt}*${RATE}`, XF.money);
      put(`${c}${R.cx}`, `=${c}${R.eb}*${CX}`, XF.money);
      put(`${c}${R.tax}`, `=MAX(0,(${c}${R.eb}-${c}${R.int}-${c}${R.cx})*${TAX})`, XF.money);
      put(`${c}${R.fcf}`, `=${c}${R.eb}-${c}${R.int}-${c}${R.cx}-${c}${R.tax}`, XF.money, 'total');
      put(`${c}${R.rep}`, `=MIN(${p}${R.debt},MAX(0,${c}${R.fcf}*${SW}))`, XF.money);
      put(`${c}${R.debt}`, `=${p}${R.debt}-${c}${R.rep}`, XF.money, 'total');
      put(`${c}${R.cash}`, `=${p}${R.cash}+${c}${R.fcf}-${c}${R.rep}`, XF.money);
      put(`${c}${R.lev}`, `=IFERROR(${c}${R.debt}/${c}${R.eb},0)`, XF.mult, 'pctRow');
      put(`${c}${R.cov}`, `=IFERROR(${c}${R.eb}/${c}${R.int},"")`, XF.mult, 'pctRow');
      put(`${c}${R.eqv}`, `=${c}${R.eb}*${EXIT}-${c}${R.debt}+${c}${R.cash}`, XF.money, 'total');
    }
    r++; section(r++, 'Returns', n + 1);
    const X = output(r++, 'Exit equity value', `=${last}${R.eqv}`, XF.money);
    const MO = output(r++, 'MOIC (multiple of invested capital)', `=${X}/${EQ0}`, XF.mult, 'output', 'Exit equity value ÷ equity invested');
    const IR = output(r++, 'IRR', `=MAX(0,${MO})^(1/${n})-1`, XF.pct, 'output', 'MOIC ^ (1 ÷ holding years) – 1');
    output(r++, 'Debt repaid over the holding period', `=${DEBT0}-${last}${R.debt}`, XF.money, 'calc');
    r++; section(r++, 'Returns attribution (equity value created)', n + 1);
    const a0 = r;
    output(r++, 'EBITDA growth', `=(${last}${R.eb}-${EB})*${ENTRY}`, XF.money, 'calc');
    output(r++, 'Multiple expansion / (contraction)', `=${last}${R.eb}*(${EXIT}-${ENTRY})`, XF.money, 'calc');
    output(r++, 'Debt paydown and cash build', `=${DEBT0}-${last}${R.debt}+${last}${R.cash}`, XF.money, 'calc');
    output(r++, 'Transaction fees', `=-B${su + 1}`, XF.money, 'calc');
    output(r++, 'Total equity value created', `=SUM(B${a0}:B${r - 2})`, XF.money);
    r++; section(r++, 'IRR sensitivity: entry multiple ↓ × exit multiple →', n + 1);
    const hr = r++;
    put(`A${hr}`, 'Entry \\ Exit', null, 'headL');
    [-2, -1, 0, 1, 2].forEach((d, j) => put(`${colL(j + 1)}${hr}`, `=${EXIT}+${d}`, XF.mult, 'head'));
    [-1, -0.5, 0, 0.5, 1].forEach((d) => {
      const row = r++;
      put(`A${row}`, `=${ENTRY}+${d}`, XF.mult, 'headL');
      for (let j = 1; j <= 5; j++) {
        const eq0 = `(${EB}*$A${row}*(1+${FEES})-${DEBT0})`;
        put(`${colL(j)}${row}`, `=IFERROR(MAX(0,(${last}$${R.eb}*${colL(j)}$${hr}-${last}$${R.debt}+${last}$${R.cash})/${eq0})^(1/${n})-1,"")`, XF.pct, d === 0 && j === 3 ? 'output' : undefined);
      }
    });
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: n + 1 }, (_, i) => [C(i), 14])) };
    book.out.lbo = { irr: IR, moic: MO };
  } },

  comps: { sheet: 'Comps', desc: 'Trading comparables: peer multiples (low, median, mean, high) and implied valuation range', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'Comps', 'Comparable Companies', sub('trading multiples'));
    let r = 5; section(r++, 'Company metrics (last actual year)', 8);
    const REV = assume(r++, 'Revenue', linkOr(book, 'revenue', v.rev), XF.money), EB = assume(r++, 'EBITDA', linkOr(book, 'ebitda', v.ebitda), XF.money);
    const NI = assume(r++, 'Net income', linkOr(book, 'netIncome', v.ni), XF.money);
    const K = book.core, ND = assume(r++, 'Net debt', near(K.op.debt - K.op.cash, v.netDebt) ? `=${ref(BS_, K.lastA + K.B.debt)}-${ref(BS_, K.lastA + K.B.cash)}` : v.netDebt, XF.money);
    const lists = [['EV / EBITDA', list(v.evE), EB, false], ['EV / Sales', list(v.evS), REV, false], ['P / E', list(v.pe), NI, true]];
    const peers = Math.max(1, ...lists.map((l) => l[1].length));
    r++; section(r++, 'Peer multiples (enter your peer group in blue)', 8);
    header(r++, ['Peer', ...lists.map((l) => l[0])]);
    const p0 = r;
    for (let k = 0; k < peers; k++, r++) {
      put(`A${r}`, `Peer ${k + 1}`, null, 'label');
      lists.forEach((l, j) => { if (l[1][k] != null) put(`${colL(j + 1)}${r}`, l[1][k], XF.mult, 'input'); });
    }
    const p1 = r - 1;
    r++; section(r++, 'Implied equity value', 8);
    header(r++, ['Method', 'Low', 'Median', 'Mean', 'High', 'Low value', 'Median value', 'Mean value', 'High value']);
    const medRows = [];
    lists.forEach(([name, vals, b, isEq], j) => {
      if (!vals.length) return;
      const rg = `${colL(j + 1)}$${p0}:${colL(j + 1)}$${p1}`;
      put(`A${r}`, name, null, 'label');
      ['MIN', 'MEDIAN', 'AVERAGE', 'MAX'].forEach((fn, k) => {
        put(`${colL(1 + k)}${r}`, `=${fn}(${rg})`, XF.mult);
        put(`${colL(5 + k)}${r}`, isEq ? `=${colL(1 + k)}${r}*${b}` : `=${colL(1 + k)}${r}*${b}-${ND}`, XF.money);
      });
      medRows.push(r); r++;
    });
    r++;
    const med = output(r++, 'Implied equity value (median of methods)', `=MEDIAN(${medRows.map((x) => `G${x}`).join(',')})`, XF.money);
    output(r++, 'Implied equity value (average of methods)', `=AVERAGE(${medRows.map((x) => `H${x}`).join(',')})`, XF.money, 'calc');
    output(r++, 'Valuation range: low', `=MIN(${medRows.map((x) => `F${x}`).join(',')})`, XF.money, 'calc');
    output(r++, 'Valuation range: high', `=MAX(${medRows.map((x) => `I${x}`).join(',')})`, XF.money, 'calc');
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: 8 }, (_, i) => [colL(i + 1), 14])) };
    book.out.comps = { median: med };
  } },

  npv: { sheet: 'NPV_IRR', desc: 'Investment appraisal: NPV, IRR, MIRR, payback, profitability index and NPV profile', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'NPV_IRR', 'NPV / IRR Investment Appraisal', sub('project returns'));
    const flows = list(v.flows), n = flows.length, C = (t) => colL(t + 1), last = C(n);
    let r = 5; section(r++, 'Assumptions', n + 1);
    const INV = assume(r++, 'Initial investment', Math.abs(v.inv), XF.money), R = assume(r++, 'Discount rate', v.rate / 100, XF.pct), RE = assume(r++, 'Reinvestment rate (MIRR)', v.reinv / 100, XF.pct);
    r++; section(r++, 'Cash flows (enter yearly cash flows in blue)', n + 1);
    header(r++, ['', ...Array.from({ length: n + 1 }, (_, t) => `Year ${t}`)]);
    const X = {}; ['cf', 'df', 'dcf', 'cum', 'dcum', 'pb', 'dpb'].forEach((k) => (X[k] = r++));
    const lab = { cf: 'Cash flow', df: 'Discount factor', dcf: 'Discounted cash flow', cum: 'Cumulative cash flow', dcum: 'Cumulative discounted cash flow', pb: 'Payback (helper)', dpb: 'Discounted payback (helper)' };
    Object.entries(X).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['pb', 'dpb'].includes(k) ? 'note' : k === 'cf' ? 'labelB' : 'label'));
    put(`B${X.cf}`, `=-${INV}`, XF.money, 'total');
    flows.forEach((f, t) => put(`${C(t + 1)}${X.cf}`, f, XF.money, 'input'));
    for (let t = 0; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      put(`${c}${X.df}`, `=1/(1+${R})^${t}`, XF.df);
      put(`${c}${X.dcf}`, `=${c}${X.cf}*${c}${X.df}`, XF.money);
      put(`${c}${X.cum}`, t ? `=${p}${X.cum}+${c}${X.cf}` : `=${c}${X.cf}`, XF.money);
      put(`${c}${X.dcum}`, t ? `=${p}${X.dcum}+${c}${X.dcf}` : `=${c}${X.dcf}`, XF.money);
      if (t) {
        put(`${c}${X.pb}`, `=IF(AND(${p}${X.cum}<0,${c}${X.cum}>=0),${t - 1}-${p}${X.cum}/${c}${X.cf},0)`, XF.dec, 'note');
        put(`${c}${X.dpb}`, `=IF(AND(${p}${X.dcum}<0,${c}${X.dcum}>=0),${t - 1}-${p}${X.dcum}/${c}${X.dcf},0)`, XF.dec, 'note');
      }
    }
    r++; section(r++, 'Results', n + 1);
    const NPVr = output(r++, 'Net present value (NPV)', `=B${X.cf}+NPV(${R},C${X.cf}:${last}${X.cf})`, XF.money);
    const IRRr = output(r++, 'Internal rate of return (IRR)', `=IRR(B${X.cf}:${last}${X.cf})`, XF.pct);
    output(r++, 'Modified IRR (MIRR)', `=MIRR(B${X.cf}:${last}${X.cf},${R},${RE})`, XF.pct);
    output(r++, 'Payback period', `=IF(SUM(C${X.pb}:${last}${X.pb})=0,"Never",SUM(C${X.pb}:${last}${X.pb}))`, XF.yrs);
    output(r++, 'Discounted payback period', `=IF(SUM(C${X.dpb}:${last}${X.dpb})=0,"Never",SUM(C${X.dpb}:${last}${X.dpb}))`, XF.yrs);
    output(r++, 'Profitability index', `=NPV(${R},C${X.cf}:${last}${X.cf})/${INV}`, XF.mult);
    output(r++, 'Total net cash (undiscounted)', `=SUM(B${X.cf}:${last}${X.cf})`, XF.money, 'calc');
    r++; section(r++, 'NPV profile', n + 1);
    header(r++, ['Discount rate', 'NPV']);
    [0, 0.05, 0.1, 0.15, 0.2, 0.25].forEach((d) => { put(`A${r}`, d, XF.pct, 'label'); put(`B${r++}`, `=B${X.cf}+NPV(A${r - 1},C${X.cf}:${last}${X.cf})`, XF.money); });
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: n + 1 }, (_, i) => [C(i), 14])) };
    book.out.npv = { npv: NPVr, irr: IRRr };
  } },

  breakeven: { sheet: 'Breakeven', desc: 'Break-even units and revenue, margin of safety, operating leverage and a volume sensitivity', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'Breakeven', 'Break-even Analysis', sub('cost-volume-profit'));
    let r = 5; section(r++, 'Assumptions', 5);
    const F = assume(r++, 'Fixed costs per year', v.fixed, XF.money), P = assume(r++, 'Price per unit', v.price, XF.money2), VC = assume(r++, 'Variable cost per unit', v.vc, XF.money2), U = assume(r++, 'Expected units sold', v.units, XF.num);
    r++; section(r++, 'Results', 5);
    const CM = output(r++, 'Contribution margin per unit', `=${P}-${VC}`, XF.money2, 'calc');
    output(r++, 'Contribution margin ratio', `=${CM}/${P}`, XF.pct, 'calc');
    const BEU = output(r++, 'Break-even units', `=${F}/${CM}`, XF.num);
    output(r++, 'Break-even revenue', `=${BEU}*${P}`, XF.money);
    const PRF = output(r++, 'Expected profit', `=${U}*${CM}-${F}`, XF.money);
    output(r++, 'Margin of safety', `=(${U}-${BEU})/${U}`, XF.pct, 'calc');
    output(r++, 'Operating leverage', `=IFERROR(${U}*${CM}/${PRF},"")`, XF.mult, 'calc');
    r++; section(r++, 'Profit at different sales volumes', 5);
    header(r++, ['', '50%', '75%', 'Plan', '125%', '150%']);
    const vr = r;
    const rowsDef = [['Units sold', (c) => `=${U}*${c}`, XF.num], ['Revenue', (c, col) => `=${col}${vr}*${P}`, XF.money], ['Variable costs', (c, col) => `=${col}${vr}*${VC}`, XF.money], ['Fixed costs', () => `=${F}`, XF.money], ['Profit', (c, col) => `=${col}${vr + 1}-${col}${vr + 2}-${col}${vr + 3}`, XF.money], ['Profit margin', (c, col) => `=IFERROR(${col}${vr + 4}/${col}${vr + 1},"")`, XF.pctS]];
    rowsDef.forEach(([label, f, z], k) => {
      put(`A${r}`, label, null, k === 4 ? 'labelB' : 'label');
      [0.5, 0.75, 1, 1.25, 1.5].forEach((m, j) => put(`${colL(j + 1)}${r}`, f(m, colL(j + 1)), z, k === 4 ? 'total' : undefined));
      r++;
    });
    sh.widths = { A: 40, B: 15, C: 15, D: 15, E: 15, F: 15 };
    book.out.be = { units: BEU };
  } },

  loan: { sheet: 'Loan', desc: 'Amortizing loan: payment, total interest, effective annual rate and annual schedule', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'Loan', 'Loan Amortization Schedule', sub('debt service'));
    const years = Math.min(40, Math.max(1, Math.round(v.years))), K = book.core;
    let r = 5; section(r++, 'Assumptions', 6);
    const P = assume(r++, 'Loan amount', near(K.op.debt, v.P) && K.op.debt ? `=${ref(BS_, K.lastA + K.B.debt)}` : v.P, XF.money), RATE = assume(r++, 'Annual interest rate', v.rate / 100, XF.pct2);
    const Y = assume(r++, 'Term (years)', years, XF.num), PPY = assume(r++, 'Payments per year', v.ppy, XF.num);
    r++; section(r++, 'Payment', 6);
    put(`A${r}`, 'Periodic interest rate', null, 'label'); put(`B${r}`, `=${RATE}/${PPY}`, '0.0000%'); const PR = `$B$${r++}`;
    put(`A${r}`, 'Number of payments', null, 'label'); put(`B${r}`, `=${Y}*${PPY}`, XF.num); const NP = `$B$${r++}`;
    const PMTr = output(r++, 'Payment per period', `=PMT(${PR},${NP},-${P})`, XF.money2); const PMTc = PMTr.split('!')[1];
    output(r++, 'Annual debt service', `=${PMTc}*${PPY}`, XF.money, 'calc');
    const TP = output(r++, 'Total paid', `=${PMTc}*${NP}`, XF.money, 'calc').split('!')[1];
    output(r++, 'Total interest', `=${TP}-${P}`, XF.money);
    output(r++, 'Effective annual rate', `=(1+${PR})^${PPY}-1`, XF.pct2, 'calc');
    r++; section(r++, 'Annual schedule', 6);
    header(r++, ['Year', 'Opening balance', 'Payments', 'Interest', 'Principal', 'Closing balance']);
    for (let y = 1; y <= years; y++, r++) {
      put(`A${r}`, y, '0', 'label');
      put(`B${r}`, y === 1 ? `=${P}` : `=F${r - 1}`, XF.money);
      put(`C${r}`, `=${PMTc}*${PPY}`, XF.money);
      put(`F${r}`, `=MAX(0,${P}*(1+${PR})^(${y}*${PPY})-${PMTc}*((1+${PR})^(${y}*${PPY})-1)/${PR})`, XF.money);
      put(`E${r}`, `=B${r}-F${r}`, XF.money);
      put(`D${r}`, `=C${r}-E${r}`, XF.money);
    }
    sh.widths = { A: 40, B: 16, C: 16, D: 16, E: 16, F: 16 };
    book.out.loan = { pmt: PMTr };
  } },

  montecarlo: { sheet: 'MonteCarlo', desc: 'Monte Carlo valuation: percentiles, mean, standard deviation and distribution', build(book, v, d) {
    const { put, section, header, output, sh } = addSheet(book, 'MonteCarlo', 'Monte Carlo Valuation', sub('simulated DCF outcomes'));
    let r = 5; section(r++, 'Assumptions used in the simulation', 4);
    [['Base revenue', v.rev, XF.money], ['Growth mean', v.g / 100, XF.pct], ['Growth std. dev.', v.gsd / 100, XF.pct], ['EBITDA margin mean', v.m / 100, XF.pct], ['Margin std. dev.', v.msd / 100, XF.pct],
      ['WACC', v.wacc / 100, XF.pct], ['Terminal growth', v.tg / 100, XF.pct], ['Net debt', v.netDebt, XF.money], ['Simulations', Math.min(Math.max(100, v.sims), 20000), XF.num]]
      .forEach(([l, x, z]) => { put(`A${r}`, l, null, 'label'); put(`B${r++}`, x, z); });
    put(`A${r++}`, 'Results were simulated in Finance Studio. Change the assumptions in the app and export again to re-run.', null, 'note');
    const mc = monteCarlo(v, d);
    r++; section(r++, 'Results: enterprise value', 4);
    if (!mc) { put(`A${r}`, 'WACC must be higher than terminal growth.', null, 'note'); return; }
    header(r++, ['Statistic', 'Enterprise value', 'Equity value']);
    let p50 = null;
    [['P5', 0.05], ['P10 (downside)', 0.1], ['P25', 0.25], ['P50 (median)', 0.5], ['P75', 0.75], ['P90 (upside)', 0.9], ['P95', 0.95]].forEach(([l, p]) => {
      put(`A${r}`, l, null, p === 0.5 ? 'labelB' : 'label'); put(`B${r}`, mc.q(p), XF.money, p === 0.5 ? 'output' : undefined); put(`C${r}`, mc.q(p) - v.netDebt, XF.money);
      if (p === 0.5) p50 = ref('MonteCarlo', `$B$${r}`);
      r++;
    });
    [['Mean', mc.mean], ['Standard deviation', mc.sd], ['Base case (no randomness)', mc.base]].forEach(([l, x]) => { put(`A${r}`, l, null, 'label'); put(`B${r++}`, x, XF.money); });
    output(r++, 'Probability above base case', mc.out.filter((x) => x > mc.base).length / mc.out.length, XF.pct, 'calc');
    r++; section(r++, 'Distribution', 4);
    header(r++, ['Bucket', 'From', 'To', 'Simulations']);
    mc.hist.forEach((cnt, i) => { put(`A${r}`, `Bucket ${i + 1}`, null, 'label'); put(`B${r}`, mc.lo + mc.w * i, XF.money); put(`C${r}`, mc.lo + mc.w * (i + 1), XF.money); put(`D${r++}`, cnt, XF.num); });
    sh.widths = { A: 40, B: 18, C: 18, D: 14 };
    book.out.mc = { p50 };
  } },

  feasibility: { sheet: 'Feasibility', desc: 'Project feasibility: cash flows, NPV, IRR, payback, NPV profile, sensitivity and verdict', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'Feasibility', `Feasibility Report: ${v.name || 'Project'}`, 'Investment appraisal of a new project');
    const n = v.life, C = (t) => colL(t + 1), last = C(n), wide = Math.min(n + 1, 12);
    let r = 5; section(r++, 'Assumptions', wide);
    const INV = assume(r++, 'Initial investment', v.inv, XF.money), REV = assume(r++, 'Year-1 revenue', v.rev, XF.money), G = assume(r++, 'Revenue growth', v.g / 100, XF.pct);
    const VAR = assume(r++, 'Variable costs % of revenue', v.var / 100, XF.pct), FIX = assume(r++, 'Fixed costs per year', v.fixed, XF.money);
    const LIFE = assume(r++, 'Project life (years)', n, XF.num, 'Changing the life needs a new export (columns are fixed)');
    const RT = assume(r++, 'Discount rate / hurdle', v.rate / 100, XF.pct), TAX = assume(r++, 'Tax rate', v.tax / 100, XF.pct), SAL = assume(r++, 'Salvage value at end', v.salvage, XF.money);
    r++; section(r++, 'Projected cash flows', wide);
    header(r++, ['', ...Array.from({ length: n + 1 }, (_, t) => `Year ${t}`)]);
    const R = {}; ['rev', 'var', 'fix', 'eb', 'dep', 'tax', 'sal', 'cf', 'cum', 'pb', 'dcf', 'dcum', 'dpb'].forEach((k) => (R[k] = r++));
    const lab = { rev: 'Revenue', var: 'Variable costs', fix: 'Fixed costs', eb: 'EBITDA', dep: 'Depreciation (tax)', tax: 'Taxes', sal: 'Salvage value', cf: 'Net cash flow', cum: 'Cumulative cash flow', pb: 'Payback (helper)', dcf: 'Discounted cash flow', dcum: 'Cumulative discounted cash flow', dpb: 'Discounted payback (helper)' };
    Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['pb', 'dpb'].includes(k) ? 'note' : ['eb', 'cf'].includes(k) ? 'labelB' : 'label'));
    put(`B${R.cf}`, `=-${INV}`, XF.money, 'total'); put(`B${R.cum}`, `=B${R.cf}`, XF.money); put(`B${R.dcf}`, `=B${R.cf}`, XF.money); put(`B${R.dcum}`, `=B${R.dcf}`, XF.money);
    for (let t = 1; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      put(`${c}${R.rev}`, `=${REV}*(1+${G})^${t - 1}`, XF.money);
      put(`${c}${R.var}`, `=${c}${R.rev}*${VAR}`, XF.money);
      put(`${c}${R.fix}`, `=${FIX}`, XF.money);
      put(`${c}${R.eb}`, `=${c}${R.rev}-${c}${R.var}-${c}${R.fix}`, XF.money, 'total');
      put(`${c}${R.dep}`, `=(${INV}-${SAL})/${LIFE}`, XF.money);
      put(`${c}${R.tax}`, `=MAX(0,(${c}${R.eb}-${c}${R.dep})*${TAX})`, XF.money);
      put(`${c}${R.sal}`, t === n ? `=${SAL}` : 0, XF.money);
      put(`${c}${R.cf}`, `=${c}${R.eb}-${c}${R.tax}+${c}${R.sal}`, XF.money, 'total');
      put(`${c}${R.cum}`, `=${p}${R.cum}+${c}${R.cf}`, XF.money);
      put(`${c}${R.pb}`, `=IF(AND(${p}${R.cum}<0,${c}${R.cum}>=0),${t - 1}-${p}${R.cum}/${c}${R.cf},0)`, XF.dec, 'note');
      put(`${c}${R.dcf}`, `=${c}${R.cf}/(1+${RT})^${t}`, XF.money);
      put(`${c}${R.dcum}`, `=${p}${R.dcum}+${c}${R.dcf}`, XF.money);
      put(`${c}${R.dpb}`, `=IF(AND(${p}${R.dcum}<0,${c}${R.dcum}>=0),${t - 1}-${p}${R.dcum}/${c}${R.dcf},0)`, XF.dec, 'note');
    }
    r++; section(r++, 'Results', wide);
    const NPVr = output(r++, 'Net present value (NPV)', `=B${R.cf}+NPV(${RT},C${R.cf}:${last}${R.cf})`, XF.money);
    const IRRr = output(r++, 'Internal rate of return (IRR)', `=IRR(B${R.cf}:${last}${R.cf})`, XF.pct);
    output(r++, 'Payback period', `=IF(SUM(C${R.pb}:${last}${R.pb})=0,"Never",SUM(C${R.pb}:${last}${R.pb}))`, XF.yrs);
    output(r++, 'Discounted payback period', `=IF(SUM(C${R.dpb}:${last}${R.dpb})=0,"Never",SUM(C${R.dpb}:${last}${R.dpb}))`, XF.yrs);
    output(r++, 'Profitability index', `=NPV(${RT},C${R.cf}:${last}${R.cf})/${INV}`, XF.mult);
    output(r++, 'Break-even revenue per year', `=${FIX}/(1-${VAR})`, XF.money);
    r++; section(r++, 'NPV profile', wide);
    header(r++, ['Discount rate', 'NPV']);
    [0, 0.05, 0.1, 0.15, 0.2, 0.25].forEach((d) => { put(`A${r}`, d, XF.pct, 'label'); put(`B${r++}`, `=B${R.cf}+NPV(A${r - 1},C${R.cf}:${last}${R.cf})`, XF.money); });
    const fa = feasAnalysis(v);
    r++; section(r++, 'Sensitivity (calculated at export)', wide);
    header(r++, ['Scenario', 'NPV', 'IRR', 'Payback (yrs)']);
    fa.sensRaw.forEach((x) => { put(`A${r}`, x.lab, null, 'label'); put(`B${r}`, x.npv, XF.money); put(`C${r}`, x.irr ?? 'n/a', XF.pct); put(`D${r++}`, x.pb ?? 'Never', XF.dec); });
    r++; section(r++, `Verdict: ${fa.title} (${fa.passed} of 4 checks passed)`, wide);
    fa.checks.forEach(([ok, t]) => put(`A${r++}`, `${ok ? '✓' : '✗'} ${t}`, null, 'label'));
    r++; put(`A${r++}`, 'Recommendations', null, 'labelB');
    fa.recs.forEach((t) => put(`A${r++}`, `• ${t}`, null, 'label'));
    sh.widths = { A: 46, ...Object.fromEntries(Array.from({ length: n + 1 }, (_, i) => [C(i), 14])) };
    book.out.feas = { npv: NPVr, irr: IRRr };
  } },
};

/* ---------- Consulting models added in this release ---------- */
Object.assign(XL_MODELS, {
  scenario: { sheet: 'Scenarios', desc: 'Bull, base and bear DCF valuations with probabilities and a weighted value', build(book, v, d) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'Scenarios', 'Scenario Analysis', sub('bull, base and bear cases valued with a DCF'));
    const K = book.core, n = 5, C = (t) => colL(t);
    let r = 5; section(r++, 'Shared assumptions', 6);
    const REV = assume(r++, 'Base revenue (last actual year)', linkOr(book, 'revenue', v.rev), XF.money);
    const DA = assume(r++, 'D&A % of revenue', d.da / 100, XF.pct), CX = assume(r++, 'CapEx % of revenue', d.capex / 100, XF.pct), NWC = assume(r++, 'Net working capital % of revenue', d.nwc / 100, XF.pct);
    const NWC0 = assume(r++, 'Opening net working capital', `=(${ref(BS_, K.lastA + K.B.tca)}-${ref(BS_, K.lastA + K.B.cash)})-${ref(BS_, K.lastA + K.B.cl)}`, XF.money, 'Last actual year from the balance sheet');
    const TAX = assume(r++, 'Tax rate', d.tax / 100, XF.pct), W = assume(r++, 'WACC', v.wacc / 100, XF.pct), TG = assume(r++, 'Terminal growth', v.tg / 100, XF.pct);
    const ND = assume(r++, 'Net debt', near(K.op.debt - K.op.cash, v.netDebt) ? `=${ref(BS_, K.lastA + K.B.debt)}-${ref(BS_, K.lastA + K.B.cash)}` : v.netDebt, XF.money), SH = assume(r++, 'Shares outstanding', v.shares, XF.num);
    r++; section(r++, 'Scenario drivers', 6);
    header(r++, ['', 'Bull', 'Base', 'Bear']);
    const gRow = r++, mRow = r++, pRow = r++;
    put(`A${gRow}`, 'Revenue growth', null, 'label'); put(`A${mRow}`, 'EBITDA margin', null, 'label'); put(`A${pRow}`, 'Probability', null, 'label');
    [[v.bullG, v.bullM, v.pBull], [0, 0, v.pBase], [v.bearG, v.bearM, v.pBear]].forEach(([dg, dm, p], j) => {
      put(`${colL(j + 1)}${gRow}`, (v.g + dg) / 100, XF.pct, 'input'); put(`${colL(j + 1)}${mRow}`, (v.m + dm) / 100, XF.pct, 'input'); put(`${colL(j + 1)}${pRow}`, p / 100, XF.pct, 'input');
    });
    put(`E${pRow}`, `=SUM(B${pRow}:D${pRow})`, XF.pct, 'total'); put(`F${pRow}`, 'Should total 100%', null, 'how');
    const evRefs = [];
    ['Bull', 'Base', 'Bear'].forEach((name, j) => {
      const G = `$${colL(j + 1)}$${gRow}`, M = `$${colL(j + 1)}$${mRow}`;
      r++; section(r++, `${name} case: free cash flow and value`, 6);
      header(r++, ['', ...Array.from({ length: n }, (_, t) => yearLabel(t + 1))]);
      const R = {}; ['rev', 'eb', 'da', 'ebit', 'tax', 'cx', 'nwc', 'dnwc', 'fcf', 'df', 'pv'].forEach((k) => (R[k] = r++));
      const lab = { rev: 'Revenue', eb: 'EBITDA', da: 'D&A', ebit: 'EBIT', tax: 'Taxes on EBIT', cx: 'CapEx', nwc: 'Net working capital', dnwc: 'Increase in working capital', fcf: 'Unlevered free cash flow', df: 'Discount factor', pv: 'Present value' };
      Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['fcf', 'pv'].includes(k) ? 'labelB' : 'label'));
      for (let t = 1; t <= n; t++) {
        const c = C(t), p = C(t - 1);
        put(`${c}${R.rev}`, t === 1 ? `=${REV}*(1+${G})` : `=${p}${R.rev}*(1+${G})`, XF.money);
        put(`${c}${R.eb}`, `=${c}${R.rev}*${M}`, XF.money); put(`${c}${R.da}`, `=${c}${R.rev}*${DA}`, XF.money); put(`${c}${R.ebit}`, `=${c}${R.eb}-${c}${R.da}`, XF.money);
        put(`${c}${R.tax}`, `=${c}${R.ebit}*${TAX}`, XF.money); put(`${c}${R.cx}`, `=${c}${R.rev}*${CX}`, XF.money); put(`${c}${R.nwc}`, `=${c}${R.rev}*${NWC}`, XF.money);
        put(`${c}${R.dnwc}`, t === 1 ? `=${c}${R.nwc}-${NWC0}` : `=${c}${R.nwc}-${p}${R.nwc}`, XF.money);
        put(`${c}${R.fcf}`, `=${c}${R.ebit}-${c}${R.tax}+${c}${R.da}-${c}${R.cx}-${c}${R.dnwc}`, XF.money, 'total');
        put(`${c}${R.df}`, `=1/(1+${W})^${t}`, XF.df); put(`${c}${R.pv}`, `=${c}${R.fcf}*${c}${R.df}`, XF.money);
      }
      const ev = output(r++, `${name} enterprise value`, `=SUM(B${R.pv}:${C(n)}${R.pv})+${C(n)}${R.fcf}*(1+${TG})/(${W}-${TG})*${C(n)}${R.df}`, XF.money, 'output', 'PV of FCF + PV of terminal value');
      output(r++, `${name} value per share`, `=(${ev.split('!')[1]}-${ND})/${SH}`, XF.money2, 'calc', '(Enterprise value – net debt) ÷ shares');
      evRefs.push(ev.split('!')[1]);
    });
    r++; section(r++, 'Result', 6);
    const wev = output(r++, 'Probability-weighted enterprise value', `=(B${pRow}*${evRefs[0]}+C${pRow}*${evRefs[1]}+D${pRow}*${evRefs[2]})/E${pRow}`, XF.money, 'output', 'Sum of (probability × scenario value) ÷ total probability');
    output(r++, 'Probability-weighted equity value', `=${wev.split('!')[1]}-${ND}`, XF.money, 'calc', 'Weighted enterprise value – net debt');
    output(r++, 'Range: bull minus bear', `=${evRefs[0]}-${evRefs[2]}`, XF.money, 'calc', 'Bull value – bear value');
    sh.widths = { A: 40, B: 15, C: 15, D: 15, E: 15, F: 15, G: 40 };
    book.out.scenario = { wev };
  } },

  workingcap: { sheet: 'WorkingCapital', desc: 'DSO, DIO, DPO, cash conversion cycle and cash released by better terms', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'WorkingCapital', 'Working Capital & Cash Conversion', sub('receivables, inventory and payables'));
    let r = 5; section(r++, 'Inputs', 4);
    const REV = assume(r++, 'Annual revenue', linkOr(book, 'revenue', v.rev), XF.money), COGS = assume(r++, 'Annual cost of goods sold', v.cogs, XF.money);
    r++; section(r++, 'Days and balances', 4);
    header(r++, ['', 'Current', 'Target', 'Change']);
    const days = r; [['Receivable days (DSO)', v.dso, v.tdso], ['Inventory days (DIO)', v.dio, v.tdio], ['Payable days (DPO)', v.dpo, v.tdpo]].forEach(([l, a, b]) => { put(`A${r}`, l, null, 'label'); put(`B${r}`, a, XF.num, 'input'); put(`C${r}`, b, XF.num, 'input'); put(`D${r}`, `=C${r}-B${r}`, XF.num); r++; });
    const bal = r;
    [['Receivables', (c) => `=${REV}*${c}${days}/365`, 'Revenue × DSO ÷ 365'], ['Inventory', (c) => `=${COGS}*${c}${days + 1}/365`, 'COGS × DIO ÷ 365'], ['Payables', (c) => `=${COGS}*${c}${days + 2}/365`, 'COGS × DPO ÷ 365'],
      ['Net working capital', (c) => `=${c}${bal}+${c}${bal + 1}-${c}${bal + 2}`, 'Receivables + inventory – payables'], ['Cash conversion cycle (days)', (c) => `=${c}${days}+${c}${days + 1}-${c}${days + 2}`, 'DSO + DIO – DPO']]
      .forEach(([l, f, how], k) => { put(`A${r}`, l, null, k >= 3 ? 'labelB' : 'label'); ['B', 'C'].forEach((c) => put(`${c}${r}`, f(c), k === 4 ? XF.num : XF.money, k >= 3 ? 'total' : undefined)); put(`D${r}`, `=C${r}-B${r}`, k === 4 ? XF.num : XF.money); put(`E${r}`, how, null, 'how'); r++; });
    r++; section(r++, 'Result', 4);
    output(r++, 'Cash released by reaching the targets', `=B${bal + 3}-C${bal + 3}`, XF.money, 'output', 'Current working capital – target working capital');
    output(r++, 'Cash value of one day of DSO', `=${REV}/365`, XF.money, 'calc', 'Revenue ÷ 365');
    sh.widths = { A: 40, B: 16, C: 16, D: 16, E: 40 };
  } },

  unitecon: { sheet: 'UnitEconomics', desc: 'LTV, CAC, payback and a monthly MRR/ARR build with churn', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'UnitEconomics', 'Unit Economics & SaaS Metrics', sub('customer value and monthly recurring revenue'));
    const n = Math.min(60, Math.max(6, Math.round(v.months))), C = (t) => colL(t);
    let r = 5; section(r++, 'Inputs', 8);
    const CU = assume(r++, 'Customers today', v.cust, XF.num), AR = assume(r++, 'Revenue per customer per month', v.arpu, XF.money2), GM = assume(r++, 'Gross margin', v.gm / 100, XF.pct);
    const CH = assume(r++, 'Monthly churn', v.churn / 100, XF.pct), AD = assume(r++, 'New customers per month', v.adds, XF.num), AG = assume(r++, 'Growth in new customers per month', v.addg / 100, XF.pct);
    const CAC = assume(r++, 'Customer acquisition cost', v.cac, XF.money2), FX = assume(r++, 'Fixed costs per month', v.fixed, XF.money);
    r++; section(r++, 'Unit economics', 8);
    const ltv = output(r++, 'Customer lifetime value (LTV)', `=IFERROR(${AR}*${GM}/${CH},0)`, XF.money, 'output', 'Monthly revenue per customer × gross margin ÷ monthly churn');
    output(r++, 'LTV / CAC', `=IFERROR(${ltv.split('!')[1]}/${CAC},0)`, XF.mult, 'output', 'Lifetime value ÷ acquisition cost (3x or more is healthy)');
    output(r++, 'CAC payback (months)', `=IFERROR(${CAC}/(${AR}*${GM}),0)`, XF.dec, 'calc', 'Acquisition cost ÷ monthly gross profit per customer');
    output(r++, 'Average customer life (months)', `=IFERROR(1/${CH},0)`, XF.dec, 'calc', '1 ÷ monthly churn');
    r++; section(r++, 'Monthly build', Math.min(n, 30));
    header(r++, ['', ...Array.from({ length: n }, (_, t) => `M${t + 1}`)]);
    const R = {}; ['start', 'lost', 'adds', 'end', 'mrr', 'gp', 'cac', 'fixed', 'profit', 'be'].forEach((k) => (R[k] = r++));
    const lab = { start: 'Customers at start', lost: 'Churned customers', adds: 'New customers', end: 'Customers at end', mrr: 'MRR', gp: 'Gross profit', cac: 'Acquisition spend', fixed: 'Fixed costs', profit: 'Monthly profit', be: 'Break-even month (helper)' };
    Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, k === 'be' ? 'note' : ['end', 'mrr', 'profit'].includes(k) ? 'labelB' : 'label'));
    for (let t = 1; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      put(`${c}${R.start}`, t === 1 ? `=${CU}` : `=${p}${R.end}`, XF.num);
      put(`${c}${R.lost}`, `=${c}${R.start}*${CH}`, XF.num);
      put(`${c}${R.adds}`, t === 1 ? `=${AD}` : `=${p}${R.adds}*(1+${AG})`, XF.num);
      put(`${c}${R.end}`, `=${c}${R.start}-${c}${R.lost}+${c}${R.adds}`, XF.num, 'total');
      put(`${c}${R.mrr}`, `=${c}${R.end}*${AR}`, XF.money, 'total');
      put(`${c}${R.gp}`, `=${c}${R.mrr}*${GM}`, XF.money);
      put(`${c}${R.cac}`, `=${c}${R.adds}*${CAC}`, XF.money);
      put(`${c}${R.fixed}`, `=${FX}`, XF.money);
      put(`${c}${R.profit}`, `=${c}${R.gp}-${c}${R.cac}-${c}${R.fixed}`, XF.money, 'total');
      put(`${c}${R.be}`, `=IF(${c}${R.profit}>=0,${t},"")`, XF.num, 'note');
    }
    r++; section(r++, 'Result', 8);
    output(r++, `MRR in month ${n}`, `=${C(n)}${R.mrr}`, XF.money, 'output', 'Customers at end × revenue per customer');
    output(r++, `ARR in month ${n}`, `=${C(n)}${R.mrr}*12`, XF.money, 'output', 'MRR × 12');
    output(r++, 'Profit break-even month', `=IF(COUNT(B${R.be}:${C(n)}${R.be})=0,"Not reached",MIN(B${R.be}:${C(n)}${R.be}))`, XF.num, 'output', 'First month with monthly profit ≥ 0');
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: n }, (_, i) => [C(i + 1), 12])) };
  } },

  runway: { sheet: 'Runway', desc: 'Monthly burn, runway in months and cash-flow break-even, with funding', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'Runway', 'Cash Runway', sub('monthly cash projection'));
    const n = Math.min(60, Math.max(6, Math.round(v.months))), C = (t) => colL(t);
    let r = 5; section(r++, 'Inputs', 8);
    const CASH = assume(r++, 'Cash today', v.cash, XF.money), RV = assume(r++, 'Monthly revenue', v.rev, XF.money), RG = assume(r++, 'Monthly revenue growth', v.revg / 100, XF.pct);
    const CO = assume(r++, 'Monthly costs', v.cost, XF.money), CG = assume(r++, 'Monthly cost growth', v.costg / 100, XF.pct), FU = assume(r++, 'New funding', v.fund, XF.money), FM = assume(r++, 'Funding arrives in month', Math.round(v.fundM), XF.num);
    r++; section(r++, 'Monthly cash flow', Math.min(n, 30));
    header(r++, ['', ...Array.from({ length: n }, (_, t) => `M${t + 1}`)]);
    const R = {}; ['rev', 'cost', 'net', 'fund', 'cash', 'out', 'be'].forEach((k) => (R[k] = r++));
    const lab = { rev: 'Revenue', cost: 'Costs', net: 'Net cash flow', fund: 'Funding', cash: 'Closing cash', out: 'Cash-out month (helper)', be: 'Break-even month (helper)' };
    Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['out', 'be'].includes(k) ? 'note' : ['net', 'cash'].includes(k) ? 'labelB' : 'label'));
    for (let t = 1; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      put(`${c}${R.rev}`, t === 1 ? `=${RV}` : `=${p}${R.rev}*(1+${RG})`, XF.money);
      put(`${c}${R.cost}`, t === 1 ? `=${CO}` : `=${p}${R.cost}*(1+${CG})`, XF.money);
      put(`${c}${R.net}`, `=${c}${R.rev}-${c}${R.cost}`, XF.money, 'total');
      put(`${c}${R.fund}`, `=IF(${t}=${FM},${FU},0)`, XF.money);
      put(`${c}${R.cash}`, t === 1 ? `=${CASH}+${c}${R.net}+${c}${R.fund}` : `=${p}${R.cash}+${c}${R.net}+${c}${R.fund}`, XF.money, 'total');
      put(`${c}${R.out}`, `=IF(${c}${R.cash}<0,${t},"")`, XF.num, 'note');
      put(`${c}${R.be}`, `=IF(${c}${R.net}>=0,${t},"")`, XF.num, 'note');
    }
    r++; section(r++, 'Result', 8);
    output(r++, 'Current monthly burn', `=MAX(0,-B${R.net})`, XF.money, 'output', 'Costs – revenue in month 1 (zero if cash-flow positive)');
    output(r++, 'Runway (months)', `=IF(COUNT(B${R.out}:${C(n)}${R.out})=0,"${n}+",MIN(B${R.out}:${C(n)}${R.out})-1)`, XF.num, 'output', 'Months before closing cash first falls below zero');
    output(r++, 'Cash-flow break-even month', `=IF(COUNT(B${R.be}:${C(n)}${R.be})=0,"Not reached",MIN(B${R.be}:${C(n)}${R.be}))`, XF.num, 'output', 'First month where revenue covers costs');
    output(r++, 'Lowest cash balance', `=MIN(B${R.cash}:${C(n)}${R.cash})`, XF.money, 'calc', 'Smallest closing cash in the period');
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: n }, (_, i) => [C(i + 1), 12])) };
  } },

  mna: { sheet: 'MandA', desc: 'M&A accretion/dilution: EPS impact, breakeven synergies and sensitivity', build(book, v) {
    const { put, section, header, assume, output, sh } = addSheet(book, 'MandA', 'M&A Accretion / Dilution', sub('earnings-per-share impact of an acquisition'));
    let r = 5; section(r++, 'Inputs', 6);
    const ANI = assume(r++, 'Buyer net income', v.ani, XF.money), ASH = assume(r++, 'Buyer shares outstanding', v.ash, XF.num), APX = assume(r++, 'Buyer share price', v.apx, XF.money2);
    const TNI = assume(r++, 'Target net income', v.tni, XF.money), PR = assume(r++, 'Purchase price (target equity)', v.price, XF.money), STK = assume(r++, 'Paid in new shares', v.stock / 100, XF.pct);
    const DBT = assume(r++, 'Paid with new debt', v.debt / 100, XF.pct), RT = assume(r++, 'Interest rate on new debt', v.rate / 100, XF.pct), CY = assume(r++, 'Interest earned on cash used', v.cashY / 100, XF.pct);
    const SYN = assume(r++, 'Pre-tax synergies per year', v.syn, XF.money), T = assume(r++, 'Tax rate', v.tax / 100, XF.pct), FE = assume(r++, 'Transaction fees', v.fees, XF.money);
    r++; section(r++, 'Calculation', 6);
    const o = (l, f, z, s, how) => output(r++, l, f, z, s, how).split('!')[1];
    const sa = o('Paid in shares', `=${PR}*${STK}`, XF.money, 'calc', 'Price × % in shares'), da = o('Paid with debt', `=${PR}*${DBT}`, XF.money, 'calc', 'Price × % in debt');
    const ca = o('Cash used (incl. fees)', `=MAX(0,${PR}-${sa}-${da})+${FE}`, XF.money, 'calc', 'Price – shares – debt + fees');
    const ns = o('New shares issued', `=${sa}/${APX}`, XF.num, 'calc', 'Paid in shares ÷ buyer share price');
    const pni = o('Pro forma net income', `=${ANI}+${TNI}+${SYN}*(1-${T})-${da}*${RT}*(1-${T})-${ca}*${CY}*(1-${T})`, XF.money, 'calc', 'Buyer + target + after-tax synergies – after-tax new interest – lost interest on cash');
    const e0 = o('Standalone EPS', `=${ANI}/${ASH}`, XF.money2, 'calc', 'Buyer net income ÷ buyer shares');
    const e1 = o('Pro forma EPS', `=${pni}/(${ASH}+${ns})`, XF.money2, 'output', 'Pro forma net income ÷ (buyer shares + new shares)');
    output(r++, 'Accretion / (dilution)', `=${e1}/${e0}-1`, XF.pctS, 'output', 'Pro forma EPS ÷ standalone EPS – 1');
    output(r++, 'Synergies needed to break even (pre-tax)', `=MAX(0,(${e0}*(${ASH}+${ns})-(${pni}-${SYN}*(1-${T})))/(1-${T}))`, XF.money, 'output', 'Synergies at which pro forma EPS equals standalone EPS');
    r++; section(r++, 'Sensitivity: accretion by % paid in shares ↓ and synergies →', 6);
    const hr = r++;
    put(`A${hr}`, 'Shares \\ Synergies', null, 'headL');
    [0, 0.5, 1, 1.5, 2].forEach((k, j) => put(`${colL(j + 1)}${hr}`, `=${SYN}*${k}`, XF.money, 'head'));
    [0, 0.25, 0.5, 0.75, 1].forEach((st) => {
      const row = r++;
      put(`A${row}`, st, XF.pct, 'headL');
      for (let j = 1; j <= 5; j++) {
        const c = colL(j), syn = `${c}$${hr}`, stp = `$A${row}`;
        const ni = `(${ANI}+${TNI}+${syn}*(1-${T})-${PR}*${DBT}*${RT}*(1-${T})-(MAX(0,${PR}-${PR}*${stp}-${PR}*${DBT})+${FE})*${CY}*(1-${T}))`;
        put(`${c}${row}`, `=${ni}/(${ASH}+${PR}*${stp}/${APX})/${e0}-1`, XF.pctS);
      }
    });
    sh.widths = { A: 42, B: 16, C: 16, D: 16, E: 16, F: 16 };
  } },

  accuracy: { sheet: 'Accuracy', desc: 'Forecast accuracy backtest: forecast vs actual for recent years', needsData: true, build(book) {
    const { put, section, header, output, sh } = addSheet(book, 'Accuracy', 'Forecast Accuracy (Backtest)', sub('how well the forecasting method predicts your own history'));
    const bt = backtest();
    let r = 5;
    put(`A${r++}`, 'Each recent year is forecast using only the years before it, with the same drivers and engine as this model, then compared with the actual result.', null, 'note');
    if (!bt) { put(`A${r}`, 'At least 3 years of statements are needed for a backtest.', null, 'note'); return; }
    r++; section(r++, 'Forecast vs actual', 6);
    header(r++, ['Year / metric', 'Forecast', 'Actual', 'Error', 'Accuracy']);
    const keyOf = { Revenue: 'revenue', EBITDA: 'ebitda', 'Net income': 'netIncome' }, ranges = {};
    ['Revenue', 'EBITDA', 'Net income'].forEach((m) => {
      const rows = bt.rows.filter((x) => x.metric === m);
      if (!rows.length) return;
      const r0 = r;
      rows.forEach((x) => {
        const yi = state.years.indexOf(x.year), h = book.hist?.[keyOf[m]];
        put(`A${r}`, `${x.year} ${m}`, null, 'label');
        put(`B${r}`, x.forecast, XF.money);
        put(`C${r}`, h && yi >= 0 ? `=Historicals!${colL(yi + 1)}${h.row}` : x.actual, XF.money, h ? 'link' : 'input');
        put(`D${r}`, `=ABS(B${r}-C${r})/ABS(C${r})`, XF.pct); put(`E${r}`, `=MAX(0,1-D${r})`, XF.pct); r++;
      });
      ranges[m] = `D${r0}:D${r - 1}`;
    });
    r++; section(r++, 'Accuracy', 6);
    const accs = Object.entries(ranges).map(([m, rg]) => output(r++, `${m} accuracy`, `=MAX(0,1-AVERAGE(${rg}))`, XF.pct, 'calc', '100% – average percentage error').split('!')[1]);
    const ov = output(r++, 'Overall forecast accuracy', `=AVERAGE(${accs.join(',')})`, XF.pct, 'output', 'Average of the metric accuracies (target 90%)');
    output(r++, 'Meets the 90% target?', `=IF(${ov.split('!')[1]}>=0.9,"Yes","Not yet")`, null, 'calc');
    sh.widths = { A: 34, B: 16, C: 16, D: 12, E: 12 };
    book.out.accuracy = { overall: ov };
  } },
});

const FORMULA_GUIDE = [
  ['Forecasting', 'Revenue', 'Prior-year revenue × (1 + growth rate)'], ['Forecasting', 'Gross profit', 'Revenue – cost of goods sold'], ['Forecasting', 'EBITDA', 'Gross profit – operating expenses'],
  ['Forecasting', 'EBIT (operating profit)', 'EBITDA – depreciation & amortization'], ['Forecasting', 'Net income', '(EBIT – interest) × (1 – tax rate)'], ['Forecasting', 'Interest expense', 'Opening debt × interest rate'],
  ['Balance sheet', 'Balance check', 'Total assets – total liabilities – equity = 0'], ['Balance sheet', 'Net working capital', 'Non-cash current assets – current liabilities'], ['Balance sheet', 'Equity roll-forward', 'Opening equity + net income – dividends'],
  ['Cash flow', 'Cash from operations', 'Net income + D&A – increase in working capital'], ['Cash flow', 'Free cash flow', 'Cash from operations – capital expenditure'],
  ['Valuation', 'Unlevered free cash flow', 'EBIT × (1 – tax rate) + D&A – CapEx – increase in working capital'], ['Valuation', 'Discount factor', '1 ÷ (1 + WACC) ^ year'],
  ['Valuation', 'Terminal value', 'Final-year FCF × (1 + g) ÷ (WACC – g)'], ['Valuation', 'Enterprise value', 'Sum of discounted FCF + discounted terminal value'], ['Valuation', 'Equity value', 'Enterprise value – net debt (debt – cash)'],
  ['Cost of capital', 'Cost of equity (CAPM)', 'Risk-free rate + beta × market risk premium (+ size premium)'], ['Cost of capital', 'WACC', 'E/V × cost of equity + D/V × pre-tax cost of debt × (1 – tax rate)'],
  ['Cost of capital', 'Unlevered beta', 'Levered beta ÷ (1 + (1 – tax rate) × debt ÷ equity)'], ['Returns', 'NPV', 'Σ cash flow ÷ (1 + rate) ^ year – initial investment'], ['Returns', 'IRR', 'Discount rate at which NPV = 0'],
  ['Returns', 'MOIC', 'Exit equity value ÷ equity invested'], ['Returns', 'Payback period', 'Years until cumulative cash flow turns positive'], ['Ratios', 'Current ratio', 'Current assets ÷ current liabilities'],
  ['Ratios', 'Net debt / EBITDA', '(Debt – cash) ÷ EBITDA'], ['Ratios', 'Return on equity', 'Net income ÷ shareholders’ equity'], ['Working capital', 'Cash conversion cycle', 'DSO + DIO – DPO'],
  ['Unit economics', 'Customer lifetime value', 'Monthly revenue per customer × gross margin ÷ monthly churn'], ['Unit economics', 'CAC payback', 'Acquisition cost ÷ monthly gross profit per customer'],
  ['M&A', 'Accretion / dilution', 'Pro forma EPS ÷ standalone EPS – 1'], ['Accuracy', 'Forecast accuracy', '100% – average |forecast – actual| ÷ |actual|'],
];
function sheetFormulaGuide(book) {
  const { put, section, sh } = addSheet(book, 'Formula Guide', 'Formula Guide', 'How every key number in this workbook is calculated', '7A8699');
  section(5, 'Formulas', 3);
  put('A6', 'Area', null, 'headL'); put('B6', 'Measure', null, 'headL'); put('C6', 'Formula', null, 'headL');
  FORMULA_GUIDE.forEach(([area, m, f], i) => { put(`A${7 + i}`, area, null, 'label'); put(`B${7 + i}`, m, null, 'labelB'); put(`C${7 + i}`, f, null, 'label'); });
  const r = 8 + FORMULA_GUIDE.length;
  put(`A${r}`, 'Tip: click any black cell in another sheet to see its live Excel formula in the formula bar. Blue cells are inputs you can change.', null, 'note');
  sh.widths = { A: 20, B: 30, C: 80 };
}

/* ---------- Assemble ---------- */
function modelInputs(id) {
  if (id === 'feasibility') return document.querySelector('#f-inv') ? feasInputs() : Object.fromEntries(FEAS.map(([k, , val]) => [k, val]));
  const m = MODELS.find((x) => x.id === id), d = defaults();
  return Object.fromEntries(m.inputs(d).map(([k, , val, t]) => [k, t === 'text' ? val : +val]));
}

const SHEET_ORDER = ['Cover', 'Dashboard', 'Assumptions', IS_, BS_, CF_, 'DCF', 'Scenarios', 'Ratios', 'Accuracy', 'WACC', 'Comps', 'MonteCarlo', 'LBO', 'MandA', 'Loan', 'NPV_IRR', 'Breakeven', 'Feasibility', 'WorkingCapital', 'UnitEconomics', 'Runway', 'Formula Guide', 'Historicals'];

function buildBook(ids, overrides = {}) {
  const book = newBook(), d = defaults(), withData = hasData();
  // Forecast drivers: from the 3-statement inputs or DCF inputs when given, otherwise from history
  let dr = d.dr, n = 5;
  if (overrides.forecast) { dr = driversFromForecast(overrides.forecast); n = overrides.forecast.years; }
  else if (ids.includes('dcf')) { const v = overrides.dcf || modelInputs('dcf'); dr = driversFromDCF(v, d.dr); n = v.years; }
  n = Math.min(10, Math.max(1, Math.round(n)));

  const cover = addSheet(book, 'Cover', `Financial Model: ${state.company || 'Manual assumptions'}`, `Prepared with Finance Studio on ${new Date().toLocaleDateString()}`);
  const dash = addSheet(book, 'Dashboard', `Dashboard: ${state.company || 'Manual assumptions'}`, `Key figures and charts · ${n}-year forecast`, 'C19A3F');
  if (withData) sheetHistoricals(book);
  buildCore(book, dr, n);
  const built = [];
  ['wacc', ...ids.filter((x) => x !== 'wacc')].filter((id) => ids.includes(id)).forEach((id) => {
    const x = XL_MODELS[id];
    if (!x || x.core) return;
    x.build(book, overrides[id] || modelInputs(id), d);
    built.push(x);
  });
  if (!ids.includes('ratios')) { /* ratios are cheap and always useful */ sheetRatios(book); built.push(XL_MODELS.ratios); }
  if (withData && !ids.includes('accuracy') && state.years.length >= 3) { XL_MODELS.accuracy.build(book); built.push(XL_MODELS.accuracy); }
  sheetFormulaGuide(book);
  sheetDashboard(dash, book);

  // Cover: contents and colour legend
  const { put, section, sh } = cover;
  let r = 5;
  put(`A${r}`, 'Source data', null, 'label'); put(`B${r++}`, withData ? `${state.company} · ${state.years[0]}–${state.years.at(-1)} (${state.years.length} periods), forecast ${yearLabel(1)}–${yearLabel(n)}` : 'No statements uploaded (manual assumptions)', null, 'label');
  r++; section(r++, 'Contents', 2);
  put(`A${r}`, 'Sheet', null, 'headL'); put(`B${r++}`, 'What it contains', null, 'headL');
  const contents = [['Dashboard', 'KPI tiles, charts, financial summary and valuation summary'], ['Assumptions', 'Forecast drivers by year (change the blue cells)'],
    [IS_, 'Revenue to net income, actual and forecast, with margins'], [BS_, 'Assets, liabilities and equity, with a balance check'], [CF_, 'Operating, investing and financing cash flows, free cash flow'],
    ...built.map((x) => [x.sheet, x.desc]), ['Formula Guide', 'Every key formula in plain English'], ...(withData ? [['Historicals', 'Your uploaded statements (the source for green links)']] : [])];
  contents.sort((a, b) => SHEET_ORDER.indexOf(a[0]) - SHEET_ORDER.indexOf(b[0])).forEach(([s, t]) => { put(`A${r}`, s, null, 'labelB'); put(`B${r++}`, t, null, 'label'); });
  r++; section(r++, 'How to use this model', 2);
  [['input', 'Blue cells are inputs. Change them and every formula updates.'], ['link', 'Green cells link to the source data.'], ['calc', 'Black cells are formulas. Avoid typing over them.'], ['output', 'Shaded cells are key outputs.']]
    .forEach(([s, t]) => { put(`A${r}`, { input: 'Input', link: 'Linked', calc: 'Formula', output: 'Output' }[s], null, s); put(`B${r++}`, t, null, 'label'); });
  r++; put(`A${r}`, 'Models are for informational purposes only and are not financial advice.', null, 'note');
  sh.widths = { A: 30, B: 90 };

  book.sheets.sort((a, b) => SHEET_ORDER.indexOf(a.name) - SHEET_ORDER.indexOf(b.name));
  evaluateBook(book);
  return book;
}

/* ---------- Write (ExcelJS) + native charts (DrawingML via JSZip) ---------- */
const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const thin = (argb = 'FF9AA5B5') => ({ style: 'thin', color: { argb } });
const XL_STYLE = {
  title: { font: { bold: true, size: 16, color: { argb: 'FF0B2545' } } },
  subtitle: { font: { italic: true, size: 10, color: { argb: 'FF5C6B80' } } },
  note: { font: { italic: true, size: 9, color: { argb: 'FF7A8699' } } },
  how: { font: { italic: true, size: 9, color: { argb: 'FF5C6B80' } }, alignment: { wrapText: false } },
  section: { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: fill('FF0B2545') },
  head: { font: { bold: true, color: { argb: 'FF14213D' } }, fill: fill('FFE9EDF3'), alignment: { horizontal: 'right' }, border: { bottom: thin() } },
  headL: { font: { bold: true, color: { argb: 'FF14213D' } }, fill: fill('FFE9EDF3'), border: { bottom: thin() } },
  headA: { font: { bold: true, color: { argb: 'FF14213D' } }, fill: fill('FFE9EDF3'), alignment: { horizontal: 'right' }, border: { bottom: thin() } },
  headF: { font: { bold: true, color: { argb: 'FF0B2545' } }, fill: fill('FFDCE6F2'), alignment: { horizontal: 'right' }, border: { bottom: thin('FF0B2545') } },
  bandA: { font: { italic: true, size: 9, color: { argb: 'FF5C6B80' } } },
  bandF: { font: { italic: true, bold: true, size: 9, color: { argb: 'FF0B2545' } } },
  label: { font: { color: { argb: 'FF14213D' } } },
  labelB: { font: { bold: true, color: { argb: 'FF14213D' } } },
  labelI: { font: { color: { argb: 'FF14213D' } }, alignment: { indent: 2 } },
  labelBI: { font: { bold: true, color: { argb: 'FF14213D' } }, alignment: { indent: 2 } },
  labelP: { font: { italic: true, size: 9, color: { argb: 'FF5C6B80' } }, alignment: { indent: 1 } },
  pctRow: { font: { italic: true, size: 9, color: { argb: 'FF5C6B80' } } },
  input: { font: { color: { argb: 'FF0000FF' } }, fill: fill('FFFFF8E1') },
  link: { font: { color: { argb: 'FF00804A' } } },
  calc: { font: { color: { argb: 'FF000000' } } },
  total: { font: { bold: true }, border: { top: thin() } },
  output: { font: { bold: true, color: { argb: 'FF0B2545' } }, fill: fill('FFDCE6F2'), border: { top: thin('FF0B2545'), bottom: thin('FF0B2545') } },
  tileLabel: { font: { bold: true, size: 9, color: { argb: 'FF5C6B80' } }, fill: fill('FFF3F5F9'), border: { top: { style: 'medium', color: { argb: 'FF0B2545' } } } },
  tileValue: { font: { bold: true, size: 18, color: { argb: 'FF0B2545' } }, fill: fill('FFF3F5F9'), alignment: { horizontal: 'left', vertical: 'middle' } },
  tileSub: { font: { size: 9, color: { argb: 'FF2E8B57' } }, fill: fill('FFF3F5F9'), alignment: { horizontal: 'left' } },
};

const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
// Cached values for a chart range, so viewers that don't recalculate still draw the chart
function rangeVals(book, rf) {
  const k = rf.lastIndexOf('!'), sh = rf.slice(0, k).replace(/^'|'$/g, ''), [p, q] = rf.slice(k + 1).replace(/\$/g, '').split(':').map((x) => x.match(/([A-Z]+)(\d+)/));
  const out = [];
  for (let c = colI(p[1]); c <= colI(q[1]); c++) out.push(book.value(sh, colL(c) + p[2]));
  return out;
}
function chartXml(spec, book) {
  const cats = rangeVals(book, spec.cats), A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const strCache = (vals) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${xmlEsc(v ?? '')}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (vals) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (typeof v === 'number' && isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
  const ser = spec.series.map((s, i) => {
    const color = `<a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill>`;
    const sp = spec.type === 'line' ? `<c:spPr><a:ln w="28575" cap="rnd">${color}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${color}<a:ln>${color}</a:ln></c:spPr></c:marker>` : `<c:spPr>${color}</c:spPr><c:invertIfNegative val="0"/>`;
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xmlEsc(s.name)}</c:v></c:tx>${sp}<c:cat><c:strRef><c:f>${xmlEsc(spec.cats)}</c:f>${strCache(cats)}</c:strRef></c:cat><c:val><c:numRef><c:f>${xmlEsc(s.ref)}</c:f>${numCache(rangeVals(book, s.ref))}</c:numRef></c:val>${spec.type === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
  }).join('');
  const txt = (sz, color = '5C6B80') => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
  const grid = '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E1E5EC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>';
  const noLine = '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>';
  const body = spec.type === 'line'
    ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/><c:axId val="5001"/><c:axId val="5002"/></c:lineChart>`
    : `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="60"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="${A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1"><a:solidFill><a:srgbClr val="0B2545"/></a:solidFill></a:rPr><a:t>${xmlEsc(spec.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${body}<c:catAx><c:axId val="5001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>${noLine}${txt(900)}<c:crossAx val="5002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="5002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${grid}<c:numFmt formatCode="${xmlEsc(spec.fmt)}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${noLine}${txt(900)}<c:crossAx val="5001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx><c:spPr><a:noFill/></c:spPr></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>${txt(900)}</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="DFE3EA"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`;
}
function drawingXml(charts) {
  const anchors = charts.map((c, i) => {
    const [c0, r0, c1, r1] = c.at;
    return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${c0}</xdr:col><xdr:colOff>76200</xdr:colOff><xdr:row>${r0}</xdr:row><xdr:rowOff>76200</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}</xdr:wsDr>`;
}
async function injectCharts(buf, sheetNo, charts, book) {
  if (!charts.length) return buf;
  const JSZip = await loadScript(JSZIP_URL, 'JSZip');
  const zip = await JSZip.loadAsync(buf);
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  charts.forEach((c, i) => zip.file(`xl/charts/chart${i + 1}.xml`, chartXml(c, book)));
  zip.file('xl/drawings/drawing1.xml', drawingXml(charts));
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${charts.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/chart" Target="../charts/chart${i + 1}.xml"/>`).join('')}</Relationships>`);
  const relPath = `xl/worksheets/_rels/sheet${sheetNo}.xml.rels`, rel = `<Relationship Id="rIdFsDrawing" Type="${R}/drawing" Target="../drawings/drawing1.xml"/>`;
  const existing = zip.file(relPath) ? await zip.file(relPath).async('string') : null;
  zip.file(relPath, existing ? existing.replace('</Relationships>', `${rel}</Relationships>`) : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`);
  const sp = `xl/worksheets/sheet${sheetNo}.xml`;
  let sx = await zip.file(sp).async('string');
  if (!/xmlns:r=/.test(sx)) sx = sx.replace('<worksheet ', `<worksheet xmlns:r="${R}" `);
  const at = ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst', '</worksheet>'].map((t) => sx.indexOf(t)).filter((i) => i >= 0).sort((a, b) => a - b)[0];
  sx = sx.slice(0, at) + '<drawing r:id="rIdFsDrawing"/>' + sx.slice(at);
  zip.file(sp, sx);
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('</Types>', `${charts.map((_, i) => `<Override PartName="/xl/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('')}<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function bookToBuffer(book) {
  const ExcelJS = await loadScript(EXCELJS_URL, 'ExcelJS');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Finance Studio';
  wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  book.sheets.forEach((sh) => {
    const view = sh.freeze ? { state: 'frozen', xSplit: sh.freeze.x, ySplit: sh.freeze.y, showGridLines: false } : { showGridLines: false };
    const ws = wb.addWorksheet(sh.name, {
      views: [view],
      properties: { tabColor: { argb: 'FF' + sh.tab } },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } },
    });
    ws.headerFooter.oddFooter = `&L&8${xmlEsc(state.company || 'Finance Studio')} · ${xmlEsc(sh.name)}&R&8Page &P of &N`;
    Object.entries(sh.widths).forEach(([c, w]) => { ws.getColumn(c).width = w; });
    Object.entries(sh.heights).forEach(([r, h]) => { ws.getRow(+r).height = h; });
    sh.sections.forEach(([row, lastCol]) => { for (let c = 0; c <= lastCol; c++) Object.assign(ws.getCell(`${colL(c)}${row}`), XL_STYLE.section); });
    Object.entries(sh.cells).forEach(([addr, c]) => {
      const cell = ws.getCell(addr);
      if (c.f != null) {
        const res = c.r;
        cell.value = { formula: c.f, result: typeof res === 'number' && isNaN(res) ? { error: '#NUM!' } : typeof res === 'number' && !isFinite(res) ? { error: '#DIV/0!' } : res ?? '' };
      } else cell.value = c.v;
      if (c.z) cell.numFmt = c.z;
      const st = XL_STYLE[c.s];
      if (st) Object.entries(st).forEach(([k, val]) => { cell[k] = val; });
    });
    sh.merges.forEach((m) => ws.mergeCells(m));
  });
  let buf = await wb.xlsx.writeBuffer();
  if (book.iterative) {
    // Turn on Excel's iterative calculation so the controlled circular reference (interest on average revolver balances) solves
    const JSZip = await loadScript(JSZIP_URL, 'JSZip');
    const zip = await JSZip.loadAsync(buf);
    let wx = await zip.file('xl/workbook.xml').async('string');
    wx = /<calcPr\b/.test(wx) ? wx.replace(/<calcPr\b/, '<calcPr iterate="1" iterateCount="100" iterateDelta="0.000001"') : wx.replace('</workbook>', '<calcPr iterate="1" iterateCount="100" iterateDelta="0.000001" fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/workbook.xml', wx);
    buf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  return injectCharts(buf, book.sheets.findIndex((s) => s.name === 'Dashboard') + 1, book.charts, book);
}

async function downloadModelWorkbook(ids, filename, overrides) {
  const buf = await bookToBuffer(buildBook(ids, overrides));
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
