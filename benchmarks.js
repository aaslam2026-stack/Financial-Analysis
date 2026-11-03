/* Finance Studio — indicative industry benchmarks (US).
   Rounded approximations of published industry aggregates (for example the NYU Stern / Damodaran
   industry datasets, January 2025). They are a sense-check for your assumptions, not precise statistics;
   check the latest source before relying on a specific figure.
   Fields: growth = revenue growth (5-yr, annual), gm = gross margin, em = EBITDA margin, nm = net margin,
   roe = return on equity, de = debt / equity (market), evEbitda, evSales, pe, beta (levered), wacc. */
const BENCHMARKS = [
  { id: 'software', name: 'Software (system & application)', group: 'Technology', growth: 0.12, gm: 0.71, em: 0.31, nm: 0.19, roe: 0.28, de: 0.07, evEbitda: 26, evSales: 8.5, pe: 38, beta: 1.2, wacc: 0.095 },
  { id: 'semis', name: 'Semiconductors', group: 'Technology', growth: 0.14, gm: 0.59, em: 0.39, nm: 0.26, roe: 0.3, de: 0.04, evEbitda: 30, evSales: 12, pe: 42, beta: 1.45, wacc: 0.11 },
  { id: 'hardware', name: 'Computers & hardware', group: 'Technology', growth: 0.06, gm: 0.39, em: 0.26, nm: 0.2, roe: 0.6, de: 0.05, evEbitda: 22, evSales: 6.5, pe: 33, beta: 1.1, wacc: 0.096 },
  { id: 'internet', name: 'Internet & online services', group: 'Technology', growth: 0.13, gm: 0.58, em: 0.36, nm: 0.24, roe: 0.25, de: 0.03, evEbitda: 17, evSales: 6, pe: 26, beta: 1.15, wacc: 0.1 },
  { id: 'online-retail', name: 'Retail (online)', group: 'Consumer', growth: 0.11, gm: 0.48, em: 0.18, nm: 0.09, roe: 0.2, de: 0.06, evEbitda: 19, evSales: 3.4, pe: 38, beta: 1.15, wacc: 0.095 },
  { id: 'retail', name: 'Retail (general & big-box)', group: 'Consumer', growth: 0.05, gm: 0.26, em: 0.075, nm: 0.03, roe: 0.22, de: 0.12, evEbitda: 15, evSales: 1.1, pe: 30, beta: 0.85, wacc: 0.078 },
  { id: 'apparel', name: 'Apparel & footwear', group: 'Consumer', growth: 0.05, gm: 0.5, em: 0.15, nm: 0.09, roe: 0.25, de: 0.1, evEbitda: 16, evSales: 2.4, pe: 25, beta: 1.05, wacc: 0.086 },
  { id: 'restaurants', name: 'Restaurants & dining', group: 'Consumer', growth: 0.06, gm: 0.34, em: 0.22, nm: 0.12, roe: 0.4, de: 0.18, evEbitda: 20, evSales: 4.2, pe: 28, beta: 1.0, wacc: 0.083 },
  { id: 'beverages', name: 'Beverages (soft drinks)', group: 'Consumer staples', growth: 0.05, gm: 0.57, em: 0.3, nm: 0.21, roe: 0.4, de: 0.12, evEbitda: 20, evSales: 6, pe: 26, beta: 0.6, wacc: 0.072 },
  { id: 'household', name: 'Household & personal products', group: 'Consumer staples', growth: 0.04, gm: 0.5, em: 0.26, nm: 0.15, roe: 0.3, de: 0.1, evEbitda: 18, evSales: 4.4, pe: 27, beta: 0.75, wacc: 0.076 },
  { id: 'food', name: 'Food processing', group: 'Consumer staples', growth: 0.04, gm: 0.32, em: 0.17, nm: 0.08, roe: 0.18, de: 0.3, evEbitda: 13, evSales: 2.2, pe: 20, beta: 0.6, wacc: 0.068 },
  { id: 'pharma', name: 'Pharmaceuticals', group: 'Healthcare', growth: 0.07, gm: 0.7, em: 0.33, nm: 0.15, roe: 0.22, de: 0.13, evEbitda: 15, evSales: 5, pe: 24, beta: 0.95, wacc: 0.083 },
  { id: 'meddevices', name: 'Medical devices & products', group: 'Healthcare', growth: 0.07, gm: 0.58, em: 0.26, nm: 0.13, roe: 0.14, de: 0.12, evEbitda: 22, evSales: 5.6, pe: 38, beta: 0.95, wacc: 0.085 },
  { id: 'healthcare-services', name: 'Healthcare services & facilities', group: 'Healthcare', growth: 0.07, gm: 0.2, em: 0.1, nm: 0.04, roe: 0.18, de: 0.3, evEbitda: 11, evSales: 1, pe: 18, beta: 0.9, wacc: 0.072 },
  { id: 'oilgas', name: 'Oil & gas (integrated)', group: 'Energy', growth: 0.04, gm: 0.33, em: 0.22, nm: 0.1, roe: 0.15, de: 0.15, evEbitda: 6.5, evSales: 1.4, pe: 13, beta: 0.75, wacc: 0.08 },
  { id: 'utilities', name: 'Utilities (electric & gas)', group: 'Energy', growth: 0.04, gm: 0.45, em: 0.38, nm: 0.13, roe: 0.1, de: 0.7, evEbitda: 12.5, evSales: 4.6, pe: 20, beta: 0.45, wacc: 0.058 },
  { id: 'industrial', name: 'Industrial machinery', group: 'Industrials', growth: 0.06, gm: 0.35, em: 0.2, nm: 0.11, roe: 0.24, de: 0.13, evEbitda: 17, evSales: 3.4, pe: 25, beta: 1.0, wacc: 0.088 },
  { id: 'aerospace', name: 'Aerospace & defense', group: 'Industrials', growth: 0.05, gm: 0.19, em: 0.13, nm: 0.06, roe: 0.18, de: 0.2, evEbitda: 18, evSales: 2.4, pe: 32, beta: 0.95, wacc: 0.083 },
  { id: 'auto', name: 'Automotive', group: 'Industrials', growth: 0.07, gm: 0.17, em: 0.1, nm: 0.05, roe: 0.13, de: 0.35, evEbitda: 13, evSales: 1.3, pe: 25, beta: 1.3, wacc: 0.09 },
  { id: 'construction', name: 'Engineering & construction', group: 'Industrials', growth: 0.07, gm: 0.14, em: 0.09, nm: 0.05, roe: 0.16, de: 0.14, evEbitda: 14, evSales: 1.2, pe: 25, beta: 1.05, wacc: 0.088 },
  { id: 'logistics', name: 'Transportation & logistics', group: 'Industrials', growth: 0.05, gm: 0.22, em: 0.14, nm: 0.06, roe: 0.2, de: 0.25, evEbitda: 11, evSales: 1.5, pe: 21, beta: 1.0, wacc: 0.08 },
  { id: 'chemicals', name: 'Chemicals', group: 'Materials', growth: 0.04, gm: 0.3, em: 0.18, nm: 0.07, roe: 0.13, de: 0.3, evEbitda: 12, evSales: 2.1, pe: 22, beta: 1.0, wacc: 0.08 },
  { id: 'media', name: 'Media & entertainment', group: 'Communication', growth: 0.05, gm: 0.4, em: 0.18, nm: 0.06, roe: 0.08, de: 0.3, evEbitda: 14, evSales: 2.6, pe: 30, beta: 1.0, wacc: 0.084 },
  { id: 'telecom', name: 'Telecom services', group: 'Communication', growth: 0.02, gm: 0.58, em: 0.37, nm: 0.09, roe: 0.12, de: 0.65, evEbitda: 7, evSales: 2.5, pe: 15, beta: 0.7, wacc: 0.063 },
  { id: 'reit', name: 'Real estate (REITs)', group: 'Real estate', growth: 0.05, gm: 0.65, em: 0.55, nm: 0.22, roe: 0.07, de: 0.45, evEbitda: 20, evSales: 11, pe: 35, beta: 0.85, wacc: 0.068 },
  { id: 'hotels', name: 'Hotels & leisure', group: 'Consumer', growth: 0.07, gm: 0.38, em: 0.24, nm: 0.1, roe: 0.25, de: 0.25, evEbitda: 15, evSales: 3.6, pe: 26, beta: 1.15, wacc: 0.088 },
];

// Which benchmark best matches each industry sample company on the home page
const SECTOR_BENCHMARK = { 'Software & SaaS': 'software', 'Retail & Consumer': 'retail', 'Industrial & Manufacturing': 'industrial', 'Healthcare & Medical Devices': 'meddevices',
  'Real Estate': 'reit', 'Energy & Utilities': 'utilities', 'Restaurants & Hospitality': 'restaurants', 'Logistics & Transport': 'logistics' };
