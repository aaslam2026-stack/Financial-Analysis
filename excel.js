/* Finance Studio — Excel model builder.
   Builds a real, formula-driven workbook: Cover, Historicals (the uploaded statements) and one sheet per
   chosen model. Inputs are blue, links to Historicals are green, formulas are black. A small evaluator
   computes every formula so the file opens with correct values in any viewer; Excel also recalculates on load.
   ExcelJS is loaded only when a workbook is built. Depends on helpers defined in app.js. */

const EXCELJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
let excelJsReady = null;
function loadExcelJS() {
  return (excelJsReady ??= new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = EXCELJS_URL;
    s.onload = () => res(window.ExcelJS);
    s.onerror = () => { excelJsReady = null; rej(new Error('The Excel builder failed to load. Check your internet connection.')); };
    document.head.appendChild(s);
  }));
}

const XF = { money: '#,##0;(#,##0);"–"', money2: '$#,##0.00;($#,##0.00)', pct: '0.0%', pct2: '0.00%', mult: '0.00"x"', num: '#,##0', dec: '0.00', df: '0.0000', yrs: '0.0 "yrs"' };
const colL = (i) => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; };
const colI = (l) => [...l].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

/* ---------- Formula evaluator (covers the functions these models use) ---------- */
const TOKEN_RE = /\s*(?:(\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)|("(?:[^"]|"")*")|((?:[A-Za-z_][A-Za-z0-9_]*!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|([A-Z][A-Z0-9.]*)(?=\()|(<=|>=|<>|[-+*/^(),=<>&]))/y;
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
  MIN: (a) => (nums(a).length ? Math.min(...nums(a)) : 0),
  MAX: (a) => (nums(a).length ? Math.max(...nums(a)) : 0),
  AVERAGE: (a) => { const n = nums(a); return n.length ? n.reduce((s, x) => s + x, 0) / n.length : NaN; },
  MEDIAN: (a) => (nums(a).length ? median(nums(a)) : NaN),
  ABS: ([x]) => Math.abs(N(x)),
  NPV: ([r, ...v]) => nums(v).reduce((s, x, i) => s + x / Math.pow(1 + N(r), i + 1), 0),
  IRR: ([v]) => irr(nums([v])) ?? NaN,
  MIRR: ([v, fin, re]) => mirr(nums([v]), N(fin), N(re)) ?? NaN,
  PMT: ([r, n, pv]) => { r = N(r); n = N(n); pv = N(pv); return r === 0 ? -pv / n : -(pv * r * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1); },
  IF: ([c, a, b]) => (N(c) ? a : b ?? false),
  AND: (a) => flat(a).every((x) => N(x)),
  OR: (a) => flat(a).some((x) => N(x)),
  IFERROR: ([a, b]) => (isErr(a) || (typeof a === 'number' && isNaN(a)) ? b : a),
};

function evaluateBook(book) {
  const byName = Object.fromEntries(book.sheets.map((s) => [s.name, s]));
  const cellVal = (sheet, addr) => {
    const c = byName[sheet]?.cells[addr];
    if (!c) return null;
    if (c.f == null) return c.v ?? null;
    if (c.r === undefined) {
      if (c.busy) throw new Error(`Circular reference at ${sheet}!${addr}`);
      c.busy = true;
      c.r = evalFormula(c.f, sheet);
      c.busy = false;
    }
    return c.r;
  };
  const refVal = (ref, sheet) => {
    let [sh, a] = ref.includes('!') ? ref.split('!') : [sheet, ref];
    a = a.replace(/\$/g, '');
    if (!a.includes(':')) return cellVal(sh, a);
    const [p, q] = a.split(':').map((x) => x.match(/([A-Z]+)(\d+)/));
    const out = [];
    for (let r = +p[2]; r <= +q[2]; r++) for (let c = colI(p[1]); c <= colI(q[1]); c++) out.push(cellVal(sh, colL(c) + r));
    return out;
  };
  function evalFormula(f, sheet) {
    const t = tokenize(f);
    let i = 0;
    const peek = () => t[i], isOp = (...o) => t[i] && t[i].t === 'op' && o.includes(t[i].v);
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
    if (i < t.length) throw new Error(`Unexpected ${peek().v} in ${f}`);
    return v;
  }
  book.sheets.forEach((s) => Object.keys(s.cells).forEach((a) => cellVal(s.name, a)));
}

/* ---------- Workbook model ---------- */
function newBook() { return { sheets: [], hist: null }; }

function addSheet(book, name, title, subtitle) {
  const sh = { name, cells: {}, widths: { A: 40 }, sections: [], tab: '0B2545' };
  book.sheets.push(sh);
  const put = (addr, x, z, s) => {
    sh.cells[addr] = typeof x === 'string' && x[0] === '=' ? { f: x.slice(1), z, s } : { v: x, z, s };
    return addr;
  };
  const abs = (a) => a.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  const api = {
    sh, put, abs,
    section: (row, text, lastCol = 6) => { put(`A${row}`, text, null, 'section'); sh.sections.push([row, lastCol]); },
    header: (row, labels, startCol = 0) => labels.forEach((l, i) => put(`${colL(startCol + i)}${row}`, l, null, i + startCol === 0 ? 'headL' : 'head')),
    label: (row, text, s = 'label') => put(`A${row}`, text, null, s),
    // Assumption row: label in A, value in B (input or link), optional note in C. Returns the absolute ref.
    assume: (row, text, x, z, note) => {
      put(`A${row}`, text, null, 'label');
      const isLink = typeof x === 'string' && x[0] === '=';
      put(`B${row}`, x, z, isLink ? 'link' : 'input');
      if (note) put(`C${row}`, note, null, 'note');
      return abs(`B${row}`);
    },
  };
  put('A1', title, null, 'title');
  put('A2', subtitle, null, 'subtitle');
  put('A3', 'Blue = input you can change · Green = linked to Historicals · Black = formula', null, 'note');
  return api;
}

// Historical statements sheet; remembers where each line item lives so models can link to it
function sheetHistoricals(book) {
  const { put, header, section } = addSheet(book, 'Historicals', 'Historical Financial Statements', `${state.company} · as uploaded`);
  const n = state.years.length, keys = Object.keys(LABELS).filter((k) => S(k));
  section(5, 'Financial statements', n);
  header(6, ['Line item', ...state.years.map(String)]);
  const refs = {};
  keys.forEach((k, j) => {
    const row = 7 + j;
    put(`A${row}`, LABELS[k], null, 'label');
    S(k).forEach((v, i) => { if (v != null) put(`${colL(i + 1)}${row}`, v, XF.money, 'input'); });
    const li = S(k).map((v, i) => (v == null ? -1 : i)).filter((i) => i >= 0).pop();
    refs[k] = { row, latest: `Historicals!$${colL(li + 1)}$${row}`, latestVal: S(k)[li], col: (i) => `Historicals!${colL(i + 1)}${row}` };
  });
  book.sheets.at(-1).widths = { A: 34, ...Object.fromEntries(state.years.map((_, i) => [colL(i + 1), 15])) };
  book.hist = refs;
}

// Use a green link to the latest historical value when the model input still equals it, otherwise a blue input
function linkOr(book, key, val) {
  const h = book.hist?.[key];
  return h && Math.abs(h.latestVal - val) <= 1e-6 * Math.max(1, Math.abs(val)) ? `=${h.latest}` : val;
}
function netDebtLink(book, val) {
  const d = book.hist?.debt, c = book.hist?.cash;
  if (d && Math.abs(d.latestVal - (c?.latestVal || 0) - val) <= 1e-6 * Math.max(1, Math.abs(val))) return c ? `=${d.latest}-${c.latest}` : `=${d.latest}`;
  return val;
}
const baseYear = () => parseInt(String(state.years.at(-1) || '').match(/(19|20)\d{2}/)?.[0]) || null;
const yearLabel = (t) => (baseYear() ? `${baseYear() + t}E` : `Year ${t}`);
const sub = (desc) => `${state.company || 'Manual assumptions'} · ${desc}`;

/* ---------- Model sheets (formulas mirror the app's calculations) ---------- */
const XL_MODELS = {
  dcf: { sheet: 'DCF', desc: 'Discounted cash flow valuation with sensitivity table', build(book, v) {
    const { put, section, header, assume, sh } = addSheet(book, 'DCF', 'DCF Valuation', sub('discounted cash flow'));
    const n = Math.min(10, Math.max(1, Math.round(v.years))), C = (t) => colL(t), last = C(n);
    let r = 5; section(r++, 'Assumptions', n);
    const REV = assume(r++, 'Base revenue (last actual year)', linkOr(book, 'revenue', v.rev), XF.money);
    const G = assume(r++, 'Revenue growth', v.g / 100, XF.pct), M = assume(r++, 'EBITDA margin', v.m / 100, XF.pct);
    const DA = assume(r++, 'D&A % of revenue', v.da / 100, XF.pct), CX = assume(r++, 'CapEx % of revenue', v.capex / 100, XF.pct);
    const NWC = assume(r++, 'Working capital % of new revenue', v.nwc / 100, XF.pct), TAX = assume(r++, 'Tax rate', v.tax / 100, XF.pct);
    const W = assume(r++, 'WACC (discount rate)', v.wacc / 100, XF.pct), TG = assume(r++, 'Terminal growth', v.tg / 100, XF.pct);
    const ND = assume(r++, 'Net debt', netDebtLink(book, v.netDebt), XF.money, 'Total debt less cash');
    const SH = assume(r++, 'Shares outstanding', v.shares, XF.num);
    r++; section(r++, 'Free cash flow projection', n);
    header(r++, ['Fiscal year', ...Array.from({ length: n }, (_, t) => yearLabel(t + 1))]);
    const R = {}; ['rev', 'ebitda', 'da', 'ebit', 'tax', 'nopat', 'capex', 'nwc', 'fcf', 'df', 'pv'].forEach((k) => (R[k] = r++));
    const lab = { rev: 'Revenue', ebitda: 'EBITDA', da: 'Depreciation & amortization', ebit: 'EBIT', tax: 'Taxes on EBIT', nopat: 'NOPAT', capex: 'Capital expenditure', nwc: 'Increase in working capital', fcf: 'Unlevered free cash flow', df: 'Discount factor', pv: 'Present value of FCF' };
    Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, k === 'fcf' ? 'labelB' : 'label'));
    for (let t = 1; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      put(`${c}${R.rev}`, t === 1 ? `=${REV}*(1+${G})` : `=${p}${R.rev}*(1+${G})`, XF.money);
      put(`${c}${R.ebitda}`, `=${c}${R.rev}*${M}`, XF.money);
      put(`${c}${R.da}`, `=${c}${R.rev}*${DA}`, XF.money);
      put(`${c}${R.ebit}`, `=${c}${R.ebitda}-${c}${R.da}`, XF.money);
      put(`${c}${R.tax}`, `=${c}${R.ebit}*${TAX}`, XF.money);
      put(`${c}${R.nopat}`, `=${c}${R.ebit}-${c}${R.tax}`, XF.money);
      put(`${c}${R.capex}`, `=${c}${R.rev}*${CX}`, XF.money);
      put(`${c}${R.nwc}`, t === 1 ? `=(${c}${R.rev}-${REV})*${NWC}` : `=(${c}${R.rev}-${p}${R.rev})*${NWC}`, XF.money);
      put(`${c}${R.fcf}`, `=${c}${R.nopat}+${c}${R.da}-${c}${R.capex}-${c}${R.nwc}`, XF.money, 'total');
      put(`${c}${R.df}`, `=1/(1+${W})^${t}`, XF.df);
      put(`${c}${R.pv}`, `=${c}${R.fcf}*${c}${R.df}`, XF.money);
    }
    r++; section(r++, 'Valuation', n);
    const V = {}; ['sum', 'tv', 'pvtv', 'ev', 'nd', 'eq', 'ps', 'tvpct'].forEach((k) => (V[k] = r++));
    const out = [['sum', 'Sum of PV of free cash flows', `=SUM(B${R.pv}:${last}${R.pv})`, XF.money],
      ['tv', 'Terminal value (Gordon growth)', `=${last}${R.fcf}*(1+${TG})/(${W}-${TG})`, XF.money],
      ['pvtv', 'PV of terminal value', `=B${V.tv}*${last}${R.df}`, XF.money],
      ['ev', 'Enterprise value', `=B${V.sum}+B${V.pvtv}`, XF.money, 'output'],
      ['nd', 'Less: net debt', `=${ND}`, XF.money],
      ['eq', 'Equity value', `=B${V.ev}-B${V.nd}`, XF.money, 'output'],
      ['ps', 'Value per share', `=B${V.eq}/${SH}`, XF.money2, 'output'],
      ['tvpct', 'Terminal value as % of EV', `=B${V.pvtv}/B${V.ev}`, XF.pct]];
    out.forEach(([k, l, f, z, s]) => { put(`A${V[k]}`, l, null, s ? 'labelB' : 'label'); put(`B${V[k]}`, f, z, s); });
    r++; section(r++, 'Sensitivity: enterprise value', 5);
    const hr = r++;
    put(`A${hr}`, 'WACC ↓   Terminal growth →', null, 'headL');
    [-0.01, -0.005, 0, 0.005, 0.01].forEach((d, j) => put(`${C(j + 1)}${hr}`, `=${TG}+${d}`, XF.pct, 'head'));
    [-0.02, -0.01, 0, 0.01, 0.02].forEach((d) => {
      const row = r++;
      put(`A${row}`, `=${W}+${d}`, XF.pct, 'headL');
      for (let j = 1; j <= 5; j++) {
        const c = C(j);
        put(`${c}${row}`, `=NPV($A${row},$B$${R.fcf}:$${last}$${R.fcf})+$${last}$${R.fcf}*(1+${c}$${hr})/($A${row}-${c}$${hr})/(1+$A${row})^${n}`, XF.money, d === 0 && j === 3 ? 'output' : undefined);
      }
    });
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: Math.max(n, 5) }, (_, i) => [C(i + 1), 15])) };
  } },

  wacc: { sheet: 'WACC', desc: 'Weighted average cost of capital (CAPM)', build(book, v) {
    const { put, section, assume, sh } = addSheet(book, 'WACC', 'Weighted Average Cost of Capital', sub('discount rate'));
    let r = 5; section(r++, 'Assumptions', 2);
    const E = assume(r++, 'Equity value', linkOr(book, 'equity', v.E), XF.money), D = assume(r++, 'Debt value', linkOr(book, 'debt', v.D), XF.money);
    const RF = assume(r++, 'Risk-free rate', v.rf / 100, XF.pct2), B = assume(r++, 'Beta', v.beta, XF.dec), MRP = assume(r++, 'Market risk premium', v.mrp / 100, XF.pct2);
    const SP = assume(r++, 'Size / specific risk premium', v.sp / 100, XF.pct2), KD = assume(r++, 'Pre-tax cost of debt', v.kd / 100, XF.pct2), T = assume(r++, 'Tax rate', v.tax / 100, XF.pct);
    r++; section(r++, 'Calculation', 2);
    const rows = [['Cost of equity (CAPM)', `=${RF}+${B}*${MRP}+${SP}`], ['After-tax cost of debt', `=${KD}*(1-${T})`], ['Equity weight (E / V)', `=${E}/(${E}+${D})`], ['Debt weight (D / V)', `=${D}/(${E}+${D})`]];
    const at = rows.map(([l, f]) => { put(`A${r}`, l, null, 'label'); put(`B${r}`, f, XF.pct2); return `B${r++}`; });
    put(`A${r}`, 'WACC', null, 'labelB'); put(`B${r}`, `=${at[2]}*${at[0]}+${at[3]}*${at[1]}`, XF.pct2, 'output');
    sh.widths = { A: 40, B: 16 };
  } },

  lbo: { sheet: 'LBO', desc: 'Leveraged buyout: sources & uses, debt schedule, IRR and MOIC', build(book, v) {
    const { put, section, header, assume, sh } = addSheet(book, 'LBO', 'Leveraged Buyout', sub('sponsor returns'));
    const n = Math.min(10, Math.max(1, Math.round(v.years))), C = (t) => colL(t + 1), last = C(n);
    let r = 5; section(r++, 'Assumptions', n + 1);
    const EB = assume(r++, 'Entry EBITDA', linkOr(book, 'ebitda', v.ebitda), XF.money), ENTRY = assume(r++, 'Entry multiple (x EBITDA)', v.entry, XF.mult);
    const LEV = assume(r++, 'Debt (x EBITDA)', v.lev, XF.mult), RATE = assume(r++, 'Interest rate', v.rate / 100, XF.pct), G = assume(r++, 'EBITDA growth', v.g / 100, XF.pct);
    const EXIT = assume(r++, 'Exit multiple (x EBITDA)', v.exit, XF.mult), TAX = assume(r++, 'Tax rate', v.tax / 100, XF.pct);
    const CX = assume(r++, 'CapEx % of EBITDA', v.capex / 100, XF.pct), SW = assume(r++, 'Cash used to repay debt', v.sweep / 100, XF.pct), FEES = assume(r++, 'Transaction fees % of EV', v.fees / 100, XF.pct);
    r++; section(r++, 'Sources & uses', n + 1);
    const U = {}; ['ev', 'fees', 'uses', 'debt', 'eq'].forEach((k) => (U[k] = r++));
    [['ev', 'Purchase price (enterprise value)', `=${EB}*${ENTRY}`], ['fees', 'Transaction fees', `=B${U.ev}*${FEES}`], ['uses', 'Total uses', `=B${U.ev}+B${U.fees}`, 'total'],
      ['debt', 'Debt raised', `=${EB}*${LEV}`], ['eq', 'Sponsor equity', `=B${U.uses}-B${U.debt}`, 'output']]
      .forEach(([k, l, f, s]) => { put(`A${U[k]}`, l, null, s ? 'labelB' : 'label'); put(`B${U[k]}`, f, XF.money, s); });
    r++; section(r++, 'Debt schedule', n + 1);
    header(r++, ['', 'Entry', ...Array.from({ length: n }, (_, t) => yearLabel(t + 1))]);
    const R = {}; ['eb', 'int', 'cx', 'tax', 'fcf', 'rep', 'debt', 'cash', 'eqv'].forEach((k) => (R[k] = r++));
    const lab = { eb: 'EBITDA', int: 'Interest', cx: 'Capital expenditure', tax: 'Taxes', fcf: 'Free cash flow', rep: 'Debt repayment', debt: 'Debt balance (end of year)', cash: 'Cash balance', eqv: 'Equity value at exit multiple' };
    Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['fcf', 'eqv'].includes(k) ? 'labelB' : 'label'));
    put(`B${R.eb}`, `=${EB}`, XF.money); put(`B${R.debt}`, `=B${U.debt}`, XF.money); put(`B${R.cash}`, 0, XF.money); put(`B${R.eqv}`, `=B${U.eq}`, XF.money);
    for (let t = 1; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      put(`${c}${R.eb}`, `=${p}${R.eb}*(1+${G})`, XF.money);
      put(`${c}${R.int}`, `=${p}${R.debt}*${RATE}`, XF.money);
      put(`${c}${R.cx}`, `=${c}${R.eb}*${CX}`, XF.money);
      put(`${c}${R.tax}`, `=MAX(0,(${c}${R.eb}-${c}${R.int}-${c}${R.cx})*${TAX})`, XF.money);
      put(`${c}${R.fcf}`, `=${c}${R.eb}-${c}${R.int}-${c}${R.cx}-${c}${R.tax}`, XF.money, 'total');
      put(`${c}${R.rep}`, `=MIN(${p}${R.debt},MAX(0,${c}${R.fcf}*${SW}))`, XF.money);
      put(`${c}${R.debt}`, `=${p}${R.debt}-${c}${R.rep}`, XF.money);
      put(`${c}${R.cash}`, `=${p}${R.cash}+${c}${R.fcf}-${c}${R.rep}`, XF.money);
      put(`${c}${R.eqv}`, `=${c}${R.eb}*${EXIT}-${c}${R.debt}+${c}${R.cash}`, XF.money);
    }
    r++; section(r++, 'Returns', n + 1);
    put(`A${r}`, 'Exit equity value', null, 'labelB'); put(`B${r}`, `=${last}${R.eqv}`, XF.money, 'output'); const X = r++;
    put(`A${r}`, 'MOIC (multiple of invested capital)', null, 'labelB'); put(`B${r}`, `=B${X}/B${U.eq}`, XF.mult, 'output'); const MO = r++;
    put(`A${r}`, 'IRR', null, 'labelB'); put(`B${r}`, `=MAX(0,B${MO})^(1/${n})-1`, XF.pct, 'output');
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: n + 1 }, (_, i) => [C(i), 15])) };
  } },

  comps: { sheet: 'Comps', desc: 'Trading comparables: implied valuation range', build(book, v) {
    const { put, section, header, assume, sh } = addSheet(book, 'Comps', 'Comparable Companies', sub('trading multiples'));
    let r = 5; section(r++, 'Company metrics', 6);
    const REV = assume(r++, 'Revenue', linkOr(book, 'revenue', v.rev), XF.money), EB = assume(r++, 'EBITDA', linkOr(book, 'ebitda', v.ebitda), XF.money);
    const NI = assume(r++, 'Net income', linkOr(book, 'netIncome', v.ni), XF.money), ND = assume(r++, 'Net debt', netDebtLink(book, v.netDebt), XF.money);
    const lists = [['EV / EBITDA', list(v.evE), EB, false], ['EV / Sales', list(v.evS), REV, false], ['P / E', list(v.pe), NI, true]];
    const peers = Math.max(1, ...lists.map((l) => l[1].length));
    r++; section(r++, 'Peer multiples (enter your peer group)', 6);
    header(r++, ['Peer', ...lists.map((l) => l[0])]);
    const p0 = r;
    for (let k = 0; k < peers; k++, r++) {
      put(`A${r}`, `Peer ${k + 1}`, null, 'label');
      lists.forEach((l, j) => { if (l[1][k] != null) put(`${colL(j + 1)}${r}`, l[1][k], XF.mult, 'input'); });
    }
    const p1 = r - 1;
    r++; section(r++, 'Implied equity value', 6);
    header(r++, ['Method', 'Low multiple', 'Median multiple', 'High multiple', 'Low value', 'Median value', 'High value']);
    const medRows = [];
    lists.forEach(([name, vals, base, isEq], j) => {
      if (!vals.length) return;
      const rg = `${colL(j + 1)}$${p0}:${colL(j + 1)}$${p1}`;
      put(`A${r}`, name, null, 'label');
      ['MIN', 'MEDIAN', 'MAX'].forEach((fn, k) => {
        put(`${colL(1 + k)}${r}`, `=${fn}(${rg})`, XF.mult);
        put(`${colL(4 + k)}${r}`, isEq ? `=${colL(1 + k)}${r}*${base}` : `=${colL(1 + k)}${r}*${base}-${ND}`, XF.money);
      });
      medRows.push(`F${r}`); r++;
    });
    r++;
    put(`A${r}`, 'Implied equity value (median of methods)', null, 'labelB'); put(`B${r}`, `=MEDIAN(${medRows.join(',')})`, XF.money, 'output');
    sh.widths = { A: 40, B: 16, C: 16, D: 16, E: 16, F: 16, G: 16 };
  } },

  npv: { sheet: 'NPV_IRR', desc: 'Investment appraisal: NPV, IRR, MIRR, payback', build(book, v) {
    const { put, section, header, assume, sh } = addSheet(book, 'NPV_IRR', 'NPV / IRR Investment Appraisal', sub('project returns'));
    const flows = list(v.flows), n = flows.length, C = (t) => colL(t + 1), last = C(n);
    let r = 5; section(r++, 'Assumptions', n + 1);
    const INV = assume(r++, 'Initial investment', Math.abs(v.inv), XF.money), R = assume(r++, 'Discount rate', v.rate / 100, XF.pct), RE = assume(r++, 'Reinvestment rate (MIRR)', v.reinv / 100, XF.pct);
    r++; section(r++, 'Cash flows (enter yearly cash flows in blue)', n + 1);
    header(r++, ['', ...Array.from({ length: n + 1 }, (_, t) => `Year ${t}`)]);
    const X = {}; ['cf', 'df', 'dcf', 'cum', 'dcum', 'pb', 'dpb'].forEach((k) => (X[k] = r++));
    const lab = { cf: 'Cash flow', df: 'Discount factor', dcf: 'Discounted cash flow', cum: 'Cumulative cash flow', dcum: 'Cumulative discounted cash flow', pb: 'Payback (helper)', dpb: 'Discounted payback (helper)' };
    Object.entries(X).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['pb', 'dpb'].includes(k) ? 'note' : 'label'));
    put(`B${X.cf}`, `=-${INV}`, XF.money);
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
    const res = [['Net present value (NPV)', `=B${X.cf}+NPV(${R},C${X.cf}:${last}${X.cf})`, XF.money],
      ['Internal rate of return (IRR)', `=IRR(B${X.cf}:${last}${X.cf})`, XF.pct],
      ['Modified IRR (MIRR)', `=MIRR(B${X.cf}:${last}${X.cf},${R},${RE})`, XF.pct],
      ['Payback period', `=IF(SUM(C${X.pb}:${last}${X.pb})=0,"Never",SUM(C${X.pb}:${last}${X.pb}))`, XF.yrs],
      ['Discounted payback period', `=IF(SUM(C${X.dpb}:${last}${X.dpb})=0,"Never",SUM(C${X.dpb}:${last}${X.dpb}))`, XF.yrs],
      ['Profitability index', `=NPV(${R},C${X.cf}:${last}${X.cf})/${INV}`, XF.mult]];
    res.forEach(([l, f, z]) => { put(`A${r}`, l, null, 'labelB'); put(`B${r++}`, f, z, 'output'); });
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: n + 1 }, (_, i) => [C(i), 14])) };
  } },

  breakeven: { sheet: 'Breakeven', desc: 'Break-even units, revenue and margin of safety', build(book, v) {
    const { put, section, assume, sh } = addSheet(book, 'Breakeven', 'Break-even Analysis', sub('cost-volume-profit'));
    let r = 5; section(r++, 'Assumptions', 2);
    const F = assume(r++, 'Fixed costs per year', v.fixed, XF.money), P = assume(r++, 'Price per unit', v.price, XF.money2), VC = assume(r++, 'Variable cost per unit', v.vc, XF.money2), U = assume(r++, 'Expected units sold', v.units, XF.num);
    r++; section(r++, 'Results', 2);
    const rows = [['Contribution margin per unit', `=${P}-${VC}`, XF.money2], ['Break-even units', `=${F}/B${r}`, XF.num, 'output'], ['Break-even revenue', `=B${r + 1}*${P}`, XF.money, 'output'],
      ['Expected profit', `=${U}*B${r}-${F}`, XF.money, 'output'], ['Margin of safety', `=(${U}-B${r + 1})/${U}`, XF.pct], ['Operating leverage', `=${U}*B${r}/B${r + 3}`, XF.mult]];
    rows.forEach(([l, f, z, s]) => { put(`A${r}`, l, null, s ? 'labelB' : 'label'); put(`B${r++}`, f, z, s); });
    sh.widths = { A: 40, B: 16 };
  } },

  ratios: { sheet: 'Ratios', desc: 'Ratio analysis linked to Historicals', needsData: true, build(book) {
    const { put, section, header, sh } = addSheet(book, 'Ratios', 'Ratio Analysis', sub('calculated from Historicals'));
    const H = book.hist, n = state.years.length;
    section(5, 'Ratios by year', n);
    header(6, ['Ratio', ...state.years.map(String)]);
    const defs = [['Revenue growth', ['revenue'], (i) => (i ? `=${H.revenue.col(i)}/${H.revenue.col(i - 1)}-1` : null), XF.pct],
      ['Gross margin', ['grossProfit', 'revenue'], (i) => `=${H.grossProfit.col(i)}/${H.revenue.col(i)}`, XF.pct],
      ['EBITDA margin', ['ebitda', 'revenue'], (i) => `=${H.ebitda.col(i)}/${H.revenue.col(i)}`, XF.pct],
      ['Net margin', ['netIncome', 'revenue'], (i) => `=${H.netIncome.col(i)}/${H.revenue.col(i)}`, XF.pct],
      ['Current ratio', ['currentAssets', 'currentLiabilities'], (i) => `=${H.currentAssets.col(i)}/${H.currentLiabilities.col(i)}`, XF.mult],
      ['Debt / Equity', ['debt', 'equity'], (i) => `=${H.debt.col(i)}/${H.equity.col(i)}`, XF.mult],
      ['Debt / EBITDA', ['debt', 'ebitda'], (i) => `=${H.debt.col(i)}/${H.ebitda.col(i)}`, XF.mult],
      ['Interest coverage (EBIT / interest)', ['ebit', 'interest'], (i) => `=${H.ebit.col(i)}/${H.interest.col(i)}`, XF.mult],
      ['Return on assets', ['netIncome', 'totalAssets'], (i) => `=${H.netIncome.col(i)}/${H.totalAssets.col(i)}`, XF.pct],
      ['Return on equity', ['netIncome', 'equity'], (i) => `=${H.netIncome.col(i)}/${H.equity.col(i)}`, XF.pct],
      ['Cash conversion (OCF / net income)', ['ocf', 'netIncome'], (i) => `=${H.ocf.col(i)}/${H.netIncome.col(i)}`, XF.mult]];
    let r = 7;
    defs.filter(([, keys]) => keys.every((k) => H[k])).forEach(([l, keys, f, z]) => {
      put(`A${r}`, l, null, 'label');
      for (let i = 0; i < n; i++) {
        const has = keys.every((k) => S(k)[i] != null) && (l !== 'Revenue growth' || (i > 0 && S('revenue')[i - 1] != null));
        const fx = has ? f(i) : null;
        if (fx) put(`${colL(i + 1)}${r}`, `=IFERROR(${fx.slice(1)},"")`, z);
      }
      r++;
    });
    sh.widths = { A: 40, ...Object.fromEntries(state.years.map((_, i) => [colL(i + 1), 14])) };
  } },

  forecast: { sheet: 'Forecast', desc: '3-statement style forecast: P&L, cash and debt', build(book, v) {
    const { put, section, header, assume, sh } = addSheet(book, 'Forecast', '3-Statement Forecast', sub('projected financials'));
    const n = Math.min(10, Math.max(1, Math.round(v.years))), C = (t) => colL(t + 1);
    let r = 5; section(r++, 'Assumptions', n + 1);
    const REV = assume(r++, 'Base revenue (last actual year)', linkOr(book, 'revenue', v.rev), XF.money), G = assume(r++, 'Revenue growth', v.g / 100, XF.pct);
    const GM = assume(r++, 'Gross margin', v.gm / 100, XF.pct), OP = assume(r++, 'Operating expenses % of revenue', v.opex / 100, XF.pct), DA = assume(r++, 'D&A % of revenue', v.da / 100, XF.pct);
    const CX = assume(r++, 'CapEx % of revenue', v.capex / 100, XF.pct), DEBT = assume(r++, 'Opening debt', linkOr(book, 'debt', v.debt), XF.money), REP = assume(r++, 'Debt repaid per year', v.repay, XF.money);
    const RATE = assume(r++, 'Interest rate', v.rate / 100, XF.pct), CASH = assume(r++, 'Opening cash', linkOr(book, 'cash', v.cash), XF.money), TAX = assume(r++, 'Tax rate', v.tax / 100, XF.pct);
    r++; section(r++, 'Projected financial statements', n + 1);
    header(r++, ['', baseYear() ? `${baseYear()}A` : 'Base', ...Array.from({ length: n }, (_, t) => yearLabel(t + 1))]);
    const R = {}; ['rev', 'gp', 'op', 'eb', 'da', 'ebit', 'int', 'ebt', 'tax', 'ni', 'cx', 'rep', 'debt', 'cash'].forEach((k) => (R[k] = r++));
    const lab = { rev: 'Revenue', gp: 'Gross profit', op: 'Operating expenses', eb: 'EBITDA', da: 'Depreciation & amortization', ebit: 'EBIT', int: 'Interest expense', ebt: 'Profit before tax', tax: 'Taxes', ni: 'Net income', cx: 'Capital expenditure', rep: 'Debt repayment', debt: 'Closing debt', cash: 'Closing cash' };
    Object.entries(R).forEach(([k, row]) => put(`A${row}`, lab[k], null, ['eb', 'ni', 'cash'].includes(k) ? 'labelB' : 'label'));
    put(`B${R.rev}`, `=${REV}`, XF.money); put(`B${R.debt}`, `=${DEBT}`, XF.money); put(`B${R.cash}`, `=${CASH}`, XF.money);
    for (let t = 1; t <= n; t++) {
      const c = C(t), p = C(t - 1);
      const f = { rev: `=${p}${R.rev}*(1+${G})`, gp: `=${c}${R.rev}*${GM}`, op: `=${c}${R.rev}*${OP}`, eb: `=${c}${R.gp}-${c}${R.op}`, da: `=${c}${R.rev}*${DA}`,
        ebit: `=${c}${R.eb}-${c}${R.da}`, int: `=${p}${R.debt}*${RATE}`, ebt: `=${c}${R.ebit}-${c}${R.int}`, tax: `=MAX(0,${c}${R.ebt}*${TAX})`, ni: `=${c}${R.ebt}-${c}${R.tax}`,
        cx: `=${c}${R.rev}*${CX}`, rep: `=MIN(${p}${R.debt},${REP})`, debt: `=${p}${R.debt}-${c}${R.rep}`, cash: `=${p}${R.cash}+${c}${R.ni}+${c}${R.da}-${c}${R.cx}-${c}${R.rep}` };
      Object.entries(f).forEach(([k, fx]) => put(`${c}${R[k]}`, fx, XF.money, ['eb', 'ni', 'cash'].includes(k) ? 'total' : undefined));
    }
    sh.widths = { A: 40, ...Object.fromEntries(Array.from({ length: n + 1 }, (_, i) => [C(i), 15])) };
  } },

  loan: { sheet: 'Loan', desc: 'Amortizing loan schedule', build(book, v) {
    const { put, section, header, assume, sh } = addSheet(book, 'Loan', 'Loan Amortization Schedule', sub('debt service'));
    const years = Math.min(40, Math.max(1, Math.round(v.years)));
    let r = 5; section(r++, 'Assumptions', 6);
    const P = assume(r++, 'Loan amount', linkOr(book, 'debt', v.P), XF.money), RATE = assume(r++, 'Annual interest rate', v.rate / 100, XF.pct2);
    const Y = assume(r++, 'Term (years)', years, XF.num), PPY = assume(r++, 'Payments per year', v.ppy, XF.num);
    r++; section(r++, 'Payment', 6);
    put(`A${r}`, 'Periodic interest rate', null, 'label'); put(`B${r}`, `=${RATE}/${PPY}`, '0.0000%'); const PR = `$B$${r++}`;
    put(`A${r}`, 'Number of payments', null, 'label'); put(`B${r}`, `=${Y}*${PPY}`, XF.num); const NP = `$B$${r++}`;
    put(`A${r}`, 'Payment per period', null, 'labelB'); put(`B${r}`, `=PMT(${PR},${NP},-${P})`, XF.money2, 'output'); const PMTc = `$B$${r++}`;
    put(`A${r}`, 'Total paid', null, 'label'); put(`B${r}`, `=${PMTc}*${NP}`, XF.money); const TP = `B${r++}`;
    put(`A${r}`, 'Total interest', null, 'labelB'); put(`B${r}`, `=${TP}-${P}`, XF.money, 'output'); r++;
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
  } },

  montecarlo: { sheet: 'MonteCarlo', desc: 'Monte Carlo valuation range (simulation results)', build(book, v, d) {
    const { put, section, header, assume, sh } = addSheet(book, 'MonteCarlo', 'Monte Carlo Valuation', sub('simulated DCF outcomes'));
    let r = 5; section(r++, 'Assumptions used in the simulation', 4);
    [['Base revenue', v.rev, XF.money], ['Growth mean', v.g / 100, XF.pct], ['Growth std. dev.', v.gsd / 100, XF.pct], ['EBITDA margin mean', v.m / 100, XF.pct], ['Margin std. dev.', v.msd / 100, XF.pct],
      ['WACC', v.wacc / 100, XF.pct], ['Terminal growth', v.tg / 100, XF.pct], ['Net debt', v.netDebt, XF.money], ['Simulations', Math.min(Math.max(100, v.sims), 20000), XF.num]]
      .forEach(([l, x, z]) => { put(`A${r}`, l, null, 'label'); put(`B${r++}`, x, z); });
    put(`A${r++}`, 'Results below were simulated in Finance Studio. Change the assumptions in the app and export again to re-run.', null, 'note');
    const mc = monteCarlo(v, d);
    r++; section(r++, 'Results: enterprise value', 4);
    if (!mc) { put(`A${r}`, 'WACC must be higher than terminal growth.', null, 'note'); return; }
    [['P10 (downside)', mc.q(0.1)], ['P50 (median)', mc.q(0.5)], ['P90 (upside)', mc.q(0.9)], ['Base case (no randomness)', mc.base]]
      .forEach(([l, x]) => { put(`A${r}`, l, null, 'labelB'); put(`B${r++}`, x, XF.money, 'output'); });
    put(`A${r}`, 'Probability above base case', null, 'labelB'); put(`B${r++}`, mc.out.filter((x) => x > mc.base).length / mc.out.length, XF.pct, 'output');
    r++; section(r++, 'Distribution', 4);
    header(r++, ['Bucket', 'From', 'To', 'Simulations']);
    mc.hist.forEach((cnt, i) => { put(`A${r}`, `Bucket ${i + 1}`, null, 'label'); put(`B${r}`, mc.lo + mc.w * i, XF.money); put(`C${r}`, mc.lo + mc.w * (i + 1), XF.money); put(`D${r++}`, cnt, XF.num); });
    sh.widths = { A: 40, B: 16, C: 16, D: 14 };
  } },

  feasibility: { sheet: 'Feasibility', desc: 'Project feasibility: cash flows, NPV, IRR, payback, verdict', build(book, v) {
    const { put, section, header, assume, sh } = addSheet(book, 'Feasibility', `Feasibility Report: ${v.name || 'Project'}`, 'Investment appraisal of a new project');
    const n = v.life, C = (t) => colL(t + 1), last = C(n);
    let r = 5; section(r++, 'Assumptions', Math.min(n + 1, 12));
    const INV = assume(r++, 'Initial investment', v.inv, XF.money), REV = assume(r++, 'Year-1 revenue', v.rev, XF.money), G = assume(r++, 'Revenue growth', v.g / 100, XF.pct);
    const VAR = assume(r++, 'Variable costs % of revenue', v.var / 100, XF.pct), FIX = assume(r++, 'Fixed costs per year', v.fixed, XF.money);
    const LIFE = assume(r++, 'Project life (years)', n, XF.num, 'Changing the life needs a new export (columns are fixed)');
    const RT = assume(r++, 'Discount rate / hurdle', v.rate / 100, XF.pct), TAX = assume(r++, 'Tax rate', v.tax / 100, XF.pct), SAL = assume(r++, 'Salvage value at end', v.salvage, XF.money);
    r++; section(r++, 'Projected cash flows', Math.min(n + 1, 12));
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
      put(`${c}${R.eb}`, `=${c}${R.rev}-${c}${R.var}-${c}${R.fix}`, XF.money);
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
    r++; section(r++, 'Results', Math.min(n + 1, 12));
    [['Net present value (NPV)', `=B${R.cf}+NPV(${RT},C${R.cf}:${last}${R.cf})`, XF.money],
      ['Internal rate of return (IRR)', `=IRR(B${R.cf}:${last}${R.cf})`, XF.pct],
      ['Payback period', `=IF(SUM(C${R.pb}:${last}${R.pb})=0,"Never",SUM(C${R.pb}:${last}${R.pb}))`, XF.yrs],
      ['Discounted payback period', `=IF(SUM(C${R.dpb}:${last}${R.dpb})=0,"Never",SUM(C${R.dpb}:${last}${R.dpb}))`, XF.yrs],
      ['Profitability index', `=NPV(${RT},C${R.cf}:${last}${R.cf})/${INV}`, XF.mult],
      ['Break-even revenue per year', `=${FIX}/(1-${VAR})`, XF.money]]
      .forEach(([l, f, z]) => { put(`A${r}`, l, null, 'labelB'); put(`B${r++}`, f, z, 'output'); });
    const fa = feasAnalysis(v);
    r++; section(r++, 'Sensitivity (calculated at export)', Math.min(n + 1, 12));
    header(r++, ['Scenario', 'NPV', 'IRR', 'Payback (yrs)']);
    fa.sensRaw.forEach((x) => { put(`A${r}`, x.lab, null, 'label'); put(`B${r}`, x.npv, XF.money); put(`C${r}`, x.irr ?? 'n/a', XF.pct); put(`D${r++}`, x.pb ?? 'Never', XF.dec); });
    r++; section(r++, `Verdict: ${fa.title} (${fa.passed} of 4 checks passed)`, Math.min(n + 1, 12));
    fa.checks.forEach(([ok, t]) => put(`A${r++}`, `${ok ? '✓' : '✗'} ${t}`, null, 'label'));
    r++; put(`A${r++}`, 'Recommendations', null, 'labelB');
    fa.recs.forEach((t) => put(`A${r++}`, `• ${t}`, null, 'label'));
    sh.widths = { A: 46, ...Object.fromEntries(Array.from({ length: n + 1 }, (_, i) => [C(i), 14])) };
  } },
};

/* ---------- Assemble and write ---------- */
function modelInputs(id) {
  if (id === 'feasibility') return document.querySelector('#f-inv') ? feasInputs() : Object.fromEntries(FEAS.map(([k, , val]) => [k, val]));
  const m = MODELS.find((x) => x.id === id), d = defaults();
  return Object.fromEntries(m.inputs(d).map(([k, , val, t]) => [k, t === 'text' ? val : +val]));
}

function buildBook(ids, overrides = {}) {
  const book = newBook(), d = defaults(), withData = hasData();
  const cover = addSheet(book, 'Cover', `Financial Model: ${state.company || 'Manual assumptions'}`, `Prepared with Finance Studio on ${new Date().toLocaleDateString()}`);
  if (withData) sheetHistoricals(book);
  const built = [];
  ids.forEach((id) => {
    const x = XL_MODELS[id];
    if (!x || (x.needsData && !withData)) return;
    x.build(book, overrides[id] || modelInputs(id), d);
    built.push(x);
  });
  // Cover sheet: contents and colour legend
  const { put, section, sh } = cover;
  let r = 5;
  put(`A${r}`, 'Source data', null, 'label'); put(`B${r++}`, withData ? `${state.company} · ${state.years[0]}–${state.years.at(-1)} (${state.years.length} periods)` : 'No statements uploaded (manual assumptions)', null, 'label');
  r++; section(r++, 'Contents', 2);
  put(`A${r}`, 'Sheet', null, 'headL'); put(`B${r++}`, 'What it contains', null, 'headL');
  [...(withData ? [{ sheet: 'Historicals', desc: 'Your uploaded financial statements (the source for linked figures)' }] : []), ...built]
    .forEach((x) => { put(`A${r}`, x.sheet, null, 'labelB'); put(`B${r++}`, x.desc, null, 'label'); });
  r++; section(r++, 'How to use this model', 2);
  [['input', 'Blue cells are inputs. Change them and every formula updates.'], ['link', 'Green cells link to the Historicals sheet.'], ['calc', 'Black cells are formulas. Avoid typing over them.'], ['output', 'Shaded cells are key outputs.']]
    .forEach(([s, t]) => { put(`A${r}`, { input: 'Input', link: 'Linked', calc: 'Formula', output: 'Output' }[s], null, s); put(`B${r++}`, t, null, 'label'); });
  r++; put(`A${r}`, 'Models are planning tools based on the data and assumptions provided. They are not investment advice.', null, 'note');
  sh.widths = { A: 30, B: 80 };
  evaluateBook(book);
  return book;
}

const XL_STYLE = {
  title: { font: { bold: true, size: 16, color: { argb: 'FF0B2545' } } },
  subtitle: { font: { italic: true, size: 10, color: { argb: 'FF5C6B80' } } },
  note: { font: { italic: true, size: 9, color: { argb: 'FF7A8699' } } },
  section: { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B2545' } } },
  head: { font: { bold: true, color: { argb: 'FF14213D' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9EDF3' } }, alignment: { horizontal: 'right' }, border: { bottom: { style: 'thin', color: { argb: 'FF9AA5B5' } } } },
  headL: { font: { bold: true, color: { argb: 'FF14213D' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9EDF3' } }, border: { bottom: { style: 'thin', color: { argb: 'FF9AA5B5' } } } },
  label: { font: { color: { argb: 'FF14213D' } } },
  labelB: { font: { bold: true, color: { argb: 'FF14213D' } } },
  input: { font: { color: { argb: 'FF0000FF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF8E1' } } },
  link: { font: { color: { argb: 'FF00804A' } } },
  calc: { font: { color: { argb: 'FF000000' } } },
  total: { font: { bold: true }, border: { top: { style: 'thin', color: { argb: 'FF9AA5B5' } } } },
  output: { font: { bold: true, color: { argb: 'FF0B2545' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F2' } }, border: { top: { style: 'thin', color: { argb: 'FF0B2545' } }, bottom: { style: 'thin', color: { argb: 'FF0B2545' } } } },
};

async function writeBook(book, filename) {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Finance Studio';
  wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  book.sheets.forEach((sh) => {
    const ws = wb.addWorksheet(sh.name, {
      views: [{ showGridLines: false }],
      properties: { tabColor: { argb: 'FF' + sh.tab } },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    Object.entries(sh.widths).forEach(([c, w]) => { ws.getColumn(c).width = w; });
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
  });
  const buf = await wb.xlsx.writeBuffer();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

async function downloadModelWorkbook(ids, filename, overrides) {
  await writeBook(buildBook(ids, overrides), filename);
}
