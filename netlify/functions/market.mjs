// Live market and economic data for the Indicators tab.
// Reads public CSV series from FRED (Federal Reserve Bank of St. Louis); no API key needed.
// GET /api/market -> { asOf, source, series: { ID: { value, date, prev, prevDate, yearAgo, history } } }

export const SERIES = [
  'DFF', 'DPRIME', 'DGS3MO', 'DGS2', 'DGS5', 'DGS10', 'DGS30', 'MORTGAGE30US', 'BAMLC0A0CM', 'BAMLH0A0HYM2', 'T10Y2Y', 'T10YIE',
  'CPIAUCSL', 'CPILFESL', 'UNRATE', 'A191RL1Q225SBEA', 'PAYEMS', 'UMCSENT', 'VIXCLS', 'DCOILWTICO', 'DTWEXBGS', 'SP500',
];

// "observation_date,DGS10\n2024-01-02,3.95\n2024-01-03,." -> [[date, number], ...] (missing values dropped)
export function parseFredCsv(text) {
  return text.trim().split(/\r?\n/).slice(1)
    .map((line) => line.split(','))
    .filter(([d, v]) => /^\d{4}-\d{2}-\d{2}$/.test(d) && v !== '.' && v !== '' && !isNaN(+v))
    .map(([d, v]) => [d, +v]);
}

export function summarize(points) {
  if (!points.length) return null;
  const [date, value] = points.at(-1);
  const prev = points.length > 1 ? points.at(-2) : null;
  const target = new Date(date);
  target.setFullYear(target.getFullYear() - 1);
  const yearAgo = [...points].reverse().find(([d]) => new Date(d) <= target) || null;
  // Thin the history to about 60 points for sparklines and charts
  const step = Math.max(1, Math.ceil(points.length / 60));
  const history = points.filter((_, i) => i % step === 0 || i === points.length - 1);
  return { value, date, prev: prev?.[1] ?? null, prevDate: prev?.[0] ?? null, yearAgo: yearAgo?.[1] ?? null, yearAgoDate: yearAgo?.[0] ?? null, history };
}

export default async () => {
  const start = new Date();
  start.setFullYear(start.getFullYear() - 3);
  const cosd = start.toISOString().slice(0, 10);
  const results = await Promise.allSettled(SERIES.map(async (id) => {
    const res = await fetch(`https://fred.stlouisfed.org/graph/fredgraph.csv?id=${id}&cosd=${cosd}`, {
      headers: { 'User-Agent': 'FinanceStudio/1.0 (market indicators)' },
      signal: AbortSignal.timeout(8000),
    });
    if (!res.ok) throw new Error(`${id}: HTTP ${res.status}`);
    return [id, summarize(parseFredCsv(await res.text()))];
  }));
  const series = Object.fromEntries(results.filter((r) => r.status === 'fulfilled' && r.value[1]).map((r) => r.value));
  const failed = results.filter((r) => r.status === 'rejected').length;
  if (!Object.keys(series).length) {
    return Response.json({ error: 'Market data is temporarily unavailable.' }, { status: 502 });
  }
  return Response.json(
    { asOf: new Date().toISOString(), source: 'FRED, Federal Reserve Bank of St. Louis', failed, series },
    { headers: { 'Cache-Control': 'public, max-age=900', 'Netlify-CDN-Cache-Control': 'public, durable, max-age=3600, stale-while-revalidate=86400' } },
  );
};

export const config = { path: '/api/market' };
