/* Finance Studio — AI analyst chat widget.
   Streams answers from /api/chat (Netlify function calling Claude). Each question carries a summary of the
   user's workspace (company data, drivers, open model results, market inputs) so answers use their numbers.
   When the function isn't available, a built-in guide answers common questions. Depends on app.js. */

const chat = { open: false, busy: false, history: [], offline: false };

const CHAT_STARTERS = ['Summarise my company’s financial health', 'Explain my DCF result and its key drivers', 'Is my WACC reasonable?', 'How do current interest rates affect my model?', 'Which assumptions should I stress-test?'];

/* ---------- Workspace context sent with each question ---------- */
function chatContext() {
  const out = [`Current page: ${location.hash || '#library'}`];
  if (hasData()) {
    out.push(`Company: ${state.company.replace(/\.$/, '')}. Periods: ${state.years.join(', ')}. Figures in USD.`);
    const keys = Object.keys(LABELS).filter((k) => S(k)), yrs = state.years.slice(-4), off = state.years.length - yrs.length;
    out.push(`Financial statements (last ${yrs.length} years):\nLine item | ${yrs.join(' | ')}\n${keys.map((k) => `${LABELS[k]} | ${S(k).slice(off).map((v) => (v == null ? '-' : Math.round(v).toLocaleString('en-US'))).join(' | ')}`).join('\n')}`);
    const dr = deriveDrivers(), p = (x) => `${(x * 100).toFixed(1)}%`;
    out.push(`Forecast drivers (3-yr historical averages): revenue growth ${p(dr.growth)}, gross margin ${p(dr.gm)}, opex ${p(dr.opex)} of revenue, D&A ${p(dr.da)}, CapEx ${p(dr.capex)}, tax ${p(dr.tax)}, interest rate ${p(dr.rate)}.`);
    const r = ratios().at(-1);
    out.push(`Latest-year ratios: ${r.map(([k, v]) => `${k} ${v}`).join('; ')}.`);
    if (typeof backtest === 'function') { const bt = backtest(); if (bt) out.push(`Forecast accuracy backtest on own history: ${(bt.overall * 100).toFixed(1)}% (${bt.tests} one-year-ahead tests).`); }
  } else out.push('No financial statements loaded yet.');
  if ($('#view-workspace').classList.contains('active')) {
    const m = MODELS.find((x) => x.id === currentModel);
    const inputs = $$('#modelArea input').map((i) => `${i.labels[0]?.textContent || i.dataset.k}: ${i.value}`).join('; ');
    const kpis = $$('#results .kpi').map((k) => `${k.querySelector('.label').textContent}: ${k.querySelector('.value').textContent}`).join('; ');
    out.push(`Open model: ${m?.name}. Inputs: ${inputs}. Results: ${kpis}.`);
  }
  const rf = marketRf(), b = chosenBenchmark();
  out.push(`Risk-free rate used: ${rf ? `${rf.value}% (${rf.source}, ${rf.date})` : '4.20% default'}.`);
  if (b) out.push(`Chosen industry benchmark: ${b.name} (indicative): growth ${(b.growth * 100).toFixed(0)}%, gross margin ${(b.gm * 100).toFixed(0)}%, EBITDA margin ${(b.em * 100).toFixed(0)}%, net margin ${(b.nm * 100).toFixed(0)}%, EV/EBITDA ${b.evEbitda}x, P/E ${b.pe}x, beta ${b.beta}, WACC ${(b.wacc * 100).toFixed(1)}%.`);
  if (market && !market.error) {
    const s = market.series, f = (id) => (s[id] ? `${s[id].value}% (${s[id].date})` : 'n/a');
    out.push(`Live market data: Fed funds ${f('DFF')}, 10-year Treasury ${f('DGS10')}, 2-year ${f('DGS2')}, IG credit spread ${f('BAMLC0A0CM')}, unemployment ${f('UNRATE')}.`);
  }
  return out.join('\n\n');
}

/* ---------- Small, safe Markdown renderer (escape first, then format) ---------- */
function md(text) {
  const inline = (s) => esc(s).replace(/`([^`]+)`/g, '<code>$1</code>').replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>').replace(/(^|[^*])\*([^*\s][^*]*)\*/g, '$1<em>$2</em>');
  const lines = String(text).split('\n'), html = [];
  let i = 0;
  while (i < lines.length) {
    const l = lines[i];
    if (/^```/.test(l)) { const buf = []; i++; while (i < lines.length && !/^```/.test(lines[i])) buf.push(esc(lines[i++])); i++; html.push(`<pre>${buf.join('\n')}</pre>`); continue; }
    if (/^\s*\|.*\|\s*$/.test(l)) {
      const rows = []; while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) rows.push(lines[i++]);
      const cells = (r) => r.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
      const body = rows.filter((r) => !/^\s*\|[\s:|-]+\|\s*$/.test(r));
      html.push(`<div class="table-wrap"><table><thead><tr>${cells(body[0]).map((c) => `<th>${inline(c)}</th>`).join('')}</tr></thead><tbody>${body.slice(1).map((r) => `<tr>${cells(r).map((c) => `<td>${inline(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`);
      continue;
    }
    if (/^#{1,4}\s/.test(l)) { html.push(`<h4>${inline(l.replace(/^#+\s/, ''))}</h4>`); i++; continue; }
    if (/^\s*([-*]|\d+\.)\s/.test(l)) {
      const ordered = /^\s*\d+\.\s/.test(l), items = [];
      while (i < lines.length && /^\s*([-*]|\d+\.)\s/.test(lines[i])) items.push(lines[i++].replace(/^\s*([-*]|\d+\.)\s/, ''));
      html.push(`<${ordered ? 'ol' : 'ul'}>${items.map((x) => `<li>${inline(x)}</li>`).join('')}</${ordered ? 'ol' : 'ul'}>`);
      continue;
    }
    if (!l.trim()) { i++; continue; }
    const para = []; while (i < lines.length && lines[i].trim() && !/^(#{1,4}\s|```|\s*([-*]|\d+\.)\s|\s*\|)/.test(lines[i])) para.push(lines[i++]);
    html.push(`<p>${inline(para.join(' '))}</p>`);
  }
  return html.join('');
}

/* ---------- Built-in guide, used when the AI service isn't connected ---------- */
const GUIDE = [
  [/dcf|discounted cash|intrinsic/, '**DCF valuation** values a business as the present value of its future free cash flows.\n\n- Unlevered FCF = EBIT × (1 – tax rate) + D&A – CapEx – increase in working capital\n- Terminal value = final-year FCF × (1 + g) ÷ (WACC – g)\n- Enterprise value = sum of discounted FCF + discounted terminal value; equity value = EV – net debt\n\nThe biggest drivers are WACC and terminal growth, so check the sensitivity tables. Keep terminal growth at or below long-run nominal GDP (about 2–3%).'],
  [/wacc|cost of capital|discount rate/, '**WACC** is the blended return lenders and shareholders expect.\n\n- Cost of equity (CAPM) = risk-free rate + beta × market risk premium (+ any size premium)\n- After-tax cost of debt = pre-tax rate × (1 – tax rate)\n- WACC = E/V × cost of equity + D/V × after-tax cost of debt\n\nTypical ranges: 6–8% for stable utilities and staples, 8–10% for most companies, 10–13%+ for small or high-growth firms. Set a current risk-free rate on the **Indicators** tab.'],
  [/terminal|perpetuity|long.?term growth/, '**Terminal growth** is the growth rate assumed forever after the forecast. It usually makes up 60–80% of a DCF value, so keep it modest: at or below long-run inflation plus real GDP growth (about 2–3% in the US). It must be below WACC.'],
  [/ebitda/, '**EBITDA** = revenue – cost of goods sold – operating expenses (before depreciation and amortization). It approximates operating cash profit. EBITDA margin = EBITDA ÷ revenue. Compare yours with your industry on the **Indicators** tab.'],
  [/lbo|buyout|private equity/, '**LBO**: a buyer funds a purchase mostly with debt, repays it from cash flow, then sells.\n\n- Equity invested = purchase price + fees – debt\n- MOIC = exit equity ÷ equity invested; IRR = MOIC^(1/years) – 1\n\nPrivate equity usually targets 20%+ IRR and 2.0–3.0x MOIC. Watch debt/EBITDA (often 4–6x at entry) and interest cover.'],
  [/irr|npv|payback|hurdle/, '**NPV** = sum of each year’s cash flow ÷ (1 + discount rate)^year, minus the investment. Positive NPV creates value.\n\n**IRR** is the discount rate that makes NPV zero. Accept when IRR is above your hurdle rate.\n\n**Payback** is how many years until cumulative cash flow turns positive.'],
  [/working capital|dso|dio|dpo|cash conversion/, '**Working capital** = receivables + inventory – payables.\n\n- DSO = receivables ÷ revenue × 365\n- DIO = inventory ÷ cost of goods sold × 365\n- DPO = payables ÷ cost of goods sold × 365\n- Cash conversion cycle = DSO + DIO – DPO\n\nA shorter cycle frees up cash. Use the Working Capital model to test changes.'],
  [/margin|profitab/, '**Margins** show how much of each dollar of revenue is kept: gross margin (after cost of goods sold), EBITDA margin (after operating costs) and net margin (after everything). Compare with your industry on the Indicators tab; a gap of more than about 5 points is worth explaining.'],
  [/accura|backtest|reliab/, '**Forecast accuracy**: Finance Studio backtests its forecasting method on your own history. It hides each recent year, forecasts it from the earlier years, and compares the forecast with what actually happened. Accuracy = 100% – the average percentage error. Above 90% is strong. No model can guarantee future accuracy, so use scenarios and sensitivities too.'],
  [/ratio|dupont|altman|z.?score|piotroski|f.?score|liquidity|leverage/, 'The **Ratio Analysis** model covers 30+ ratios in six groups (liquidity, profitability, leverage, efficiency, cash flow and growth), each rated strong, fair or weak, with trends and an industry column.\n\n- **DuPont**: ROE = net margin × asset turnover × equity multiplier\n- **Altman Z″**: 6.56 × working capital/assets + 3.26 × retained earnings/assets + 6.72 × EBIT/assets + 1.05 × equity/liabilities. Above 2.6 is safe, below 1.1 is distress.\n- **Piotroski F-score**: pass/fail tests of profitability, leverage, liquidity and efficiency. 7+ is strong.\n\nOpen it from **Models → Ratio Analysis**.'],
  [/beta/, '**Beta** measures how much a company’s shares move with the market (1.0 = market average). Higher beta means a higher cost of equity. Use an industry beta from the Indicators tab if your company isn’t listed.'],
  [/rate|interest|fed|treasury/, 'Interest rates feed your model in three places: the **risk-free rate** in WACC (usually the 10-year Treasury), the **cost of debt** (risk-free rate + credit spread), and the **interest expense** in the forecast. See live values on the **Indicators** tab and click “Use live 10-year yield”.'],
  [/upload|pdf|excel|csv|statement/, 'Upload statements on **Build Excel Model** or **My Data**: PDF (with selectable text), Excel or CSV. Put line items in the first column and years in a header row. The app detects “in thousands / millions” and fills the models automatically.'],
];
function guideAnswer(q) {
  const t = q.toLowerCase();
  if (/summar|health|how is my|overview/.test(t) && hasData()) {
    const i = state.years.length - 1, g = (k) => S(k)?.[i], p = (x) => (x == null || !isFinite(x) ? 'n/a' : `${(x * 100).toFixed(1)}%`);
    const b = chosenBenchmark();
    return `**${state.company}, ${state.years[i]}**\n\n- Revenue: ${fmt(g('revenue'))} (growth ${p(cagr(S('revenue')))} a year on average)\n- EBITDA margin: ${p(g('ebitda') / g('revenue'))}${b ? ` (industry ${p(b.em)})` : ''}\n- Net margin: ${p(g('netIncome') / g('revenue'))}${b ? ` (industry ${p(b.nm)})` : ''}\n- Net debt: ${fmt((g('debt') || 0) - (g('cash') || 0))}\n- Current ratio: ${g('currentLiabilities') ? (g('currentAssets') / g('currentLiabilities')).toFixed(2) + 'x' : 'n/a'}\n\nOpen **Dashboard** for trends, and pick your industry on **Indicators** to compare.`;
  }
  const hit = GUIDE.find(([re]) => re.test(t));
  return hit ? hit[1] : 'I can explain DCF, WACC, LBO, NPV/IRR, working capital, margins, ratios, interest rates, forecast accuracy and how to use each model. Try one of the suggested questions.';
}

/* ---------- UI ---------- */
function renderChat() {
  const log = $('#chatLog');
  log.innerHTML = (chat.history.length ? '' : `<div class="chat-hello"><strong>Hi, I’m your AI financial analyst.</strong><p>Ask about your numbers, any model, interest rates or industry benchmarks. I can see the company and model you have open.</p>
      <div class="chat-starters">${CHAT_STARTERS.map((q) => `<button type="button" data-q="${esc(q)}">${esc(q)}</button>`).join('')}</div></div>`)
    + chat.history.map((m) => `<div class="msg ${m.role}">${m.role === 'user' ? esc(m.content) : md(m.content) || '<span class="typing"><i></i><i></i><i></i></span>'}</div>`).join('');
  log.scrollTop = log.scrollHeight;
  $('#chatNote').innerHTML = chat.offline ? 'Built-in guide mode: the live AI analyst isn’t connected on this site yet.' : 'AI answers can be wrong. Check important figures.';
}

async function askChat(text) {
  text = text.trim();
  if (!text || chat.busy) return;
  chat.busy = true;
  $('#chatSend').disabled = true;
  chat.history.push({ role: 'user', content: text }, { role: 'assistant', content: '' });
  const answer = chat.history.at(-1);
  renderChat();
  try {
    if (chat.offline) throw new Error('offline');
    const res = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: chat.history.slice(0, -1).filter((m) => m.content), context: chatContext() }),
    });
    // No function on this deploy (static hosting) or no API key: fall back to the built-in guide
    if (!res.ok || !/event-stream/.test(res.headers.get('content-type') || '')) throw new Error(res.status === 429 ? 'busy' : 'offline');
    const reader = res.body.getReader(), dec = new TextDecoder();
    let buf = '';
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });
      const parts = buf.split('\n\n');
      buf = parts.pop();
      for (const part of parts) {
        if (!part.startsWith('data: ')) continue;
        const ev = JSON.parse(part.slice(6));
        if (ev.reset) answer.content = '';
        if (ev.t) answer.content += ev.t;
        if (ev.refusal) answer.content += '\n\n_I can’t help with that request._';
        if (ev.error) answer.content += ev.error === 'busy' ? '\n\n_The AI service is busy. Please try again in a moment._' : ev.error === 'not_configured' ? guideAnswer(text) : '\n\n_Something went wrong. Please try again._';
        renderChat();
      }
    }
  } catch (e) {
    if (e.message === 'offline' || e instanceof TypeError) chat.offline = true;
    answer.content = chat.offline ? guideAnswer(text) : e.message === 'busy' ? 'The AI service is busy. Please try again in a moment.' : 'Sorry, I couldn’t reach the AI service. Please try again.';
  }
  chat.busy = false;
  $('#chatSend').disabled = false;
  renderChat();
}

function initChat() {
  const toggle = (open) => {
    chat.open = open ?? !chat.open;
    $('#chatPanel').hidden = !chat.open;
    $('#chatFab').setAttribute('aria-expanded', String(chat.open));
    if (chat.open) { renderChat(); $('#chatInput').focus(); }
  };
  $('#chatFab').addEventListener('click', () => toggle());
  $('#chatClose').addEventListener('click', () => toggle(false));
  $('#chatClear').addEventListener('click', () => { chat.history = []; renderChat(); });
  $('#chatLog').addEventListener('click', (e) => { const b = e.target.closest('[data-q]'); if (b) askChat(b.dataset.q); });
  $('#chatForm').addEventListener('submit', (e) => { e.preventDefault(); const v = $('#chatInput').value; $('#chatInput').value = ''; askChat(v); });
  $('#chatInput').addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); $('#chatForm').requestSubmit(); } });
}
