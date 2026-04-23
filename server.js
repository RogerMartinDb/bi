require('dotenv').config();
const express = require('express');
const session = require('express-session');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { registerAuthRoutes, requireAuth } = require('./auth');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(session({
  secret: process.env.SESSION_SECRET || 'dev-secret-change-in-production',
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: process.env.NODE_ENV === 'production',
    httpOnly: true,
    maxAge: 8 * 60 * 60 * 1000, // 8 hours
  },
}));

const basePath = process.env.BASE_PATH || '';

// Auth routes (/auth/login, /auth/callback, /auth/logout, /auth/me)
registerAuthRoutes(app);

// Static files served only to authenticated users
app.use(basePath, requireAuth, express.static(__dirname));

const SKIP = new Set(['_', 'Activity by Vertical', "Other KPI's", 'Date Ending', 'Week']);

const GROUPS = {
  'GGR': 'Main KPIs',
  'NGR': 'Main KPIs',
  'Active': 'Main KPIs',
  'ARPU': 'Main KPIs',
  'Registrations': 'Main KPIs',
  "FTD's": 'Main KPIs',
  'Conversion %': 'Engagement',
  'Retention %': 'Engagement',
  'Retained # Week': 'Engagement',
  'Reactivations %': 'Engagement',
  'Reactivations # Week': 'Engagement',
  'Churn Rate %': 'Engagement',
  'Churn Overall # Week': 'Engagement',
  'Total Deposits': 'Deposits & Turnover',
  'Deposits #': 'Deposits & Turnover',
  'Turnover': 'Deposits & Turnover',
  'Hold %': 'Deposits & Turnover',
  'Casino Actives': 'Casino',
  'Casino GGR': 'Casino',
  'Casino Wagered': 'Casino',
  'Casino Win': 'Casino',
  'Casino Hold %': 'Casino',
  'Lotto Actives': 'Lotto',
  'Lotto GGR': 'Lotto',
  'Lotto Wagered': 'Lotto',
  'Lotto Win': 'Lotto',
  'Lotto Hold %': 'Lotto',
  'Sports Actives': 'Sports',
  'Sports GGR': 'Sports',
  'Sports Wagered': 'Sports',
  'Sports Win': 'Sports',
  'Sports Hold': 'Sports',
  'Horses Actives': 'Horses',
  'Horses GGR': 'Horses',
  'Horses Wager': 'Horses',
  'Horses Win': 'Horses',
  'Horses Hold': 'Horses',
  'LTV Wk': 'Other',
  'Adjustments': 'Other',
  'Sports Refund': 'Other',
  'Horses Refund': 'Other',
  'Tips': 'Other',
  'Chargebacks': 'Other',
  'FeesCredit': 'Other',
  'Freeplays Granted': 'Other',
  'Withdrawals': 'Other',
  'NetCash': 'Other',
  'Reinvestment': 'Other',
};

const GROUP_ORDER = [
  'Main KPIs', 'Engagement', 'Deposits & Turnover',
  'Casino', 'Lotto', 'Sports', 'Horses', 'Other',
];

function loadData() {
  const dataDir = path.join(__dirname, 'data');
  const result = { years: [], metrics: [], data: {}, dateEndings: {} };
  const metricsOrder = new Map();

  const files = fs.readdirSync(dataDir)
    .filter(f => /kpi \d{4}\.xlsx/i.test(f))
    .sort();

  for (const file of files) {
    const match = file.match(/kpi (\d{4})\.xlsx/i);
    if (!match) continue;
    const year = parseInt(match[1]);

    const wb = XLSX.readFile(path.join(dataDir, file));
    const ws = wb.Sheets['Export'];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (rows.length < 2) continue;

    const headers = rows[0].map(h => (typeof h === 'string' ? h.trim() : h));

    result.years.push(year);
    result.data[year] = {};

    const metricCols = [];
    let dateEndingCol = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (h === 'Date Ending') { dateEndingCol = i; continue; }
      if (!h || SKIP.has(h)) continue;
      metricCols.push({ i, name: h });
      if (!metricsOrder.has(h)) metricsOrder.set(h, GROUPS[h] || 'Other');
      result.data[year][h] = new Array(52).fill(null);
    }

    let lastDateEnding = null;
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      const week = row[0];
      if (typeof week !== 'number' || week < 1 || week > 52) continue;
      const wi = Math.round(week) - 1;
      if (dateEndingCol >= 0 && row[dateEndingCol] != null) {
        lastDateEnding = row[dateEndingCol];
      }
      for (const { i, name } of metricCols) {
        const val = row[i];
        if (typeof val === 'number' && isFinite(val)) {
          result.data[year][name][wi] = val;
        }
      }
    }

    if (lastDateEnding != null) {
      const d = typeof lastDateEnding === 'number'
        ? new Date(Date.UTC(1899, 11, 30) + lastDateEnding * 86400000)
        : lastDateEnding;
      result.dateEndings[year] = d instanceof Date
        ? d.toISOString().split('T')[0]
        : String(lastDateEnding);
    }
  }

  const byGroup = new Map();
  for (const [name, group] of metricsOrder) {
    if (!byGroup.has(group)) byGroup.set(group, []);
    byGroup.get(group).push(name);
  }

  for (const group of GROUP_ORDER) {
    if (byGroup.has(group)) {
      for (const name of byGroup.get(group)) {
        result.metrics.push({ key: name, group });
      }
    }
  }

  result.years.sort((a, b) => a - b);
  return result;
}

let data = loadData();

const noCache = (req, res, next) => {
  res.setHeader('Cache-Control', 'no-store');
  next();
};

app.get(basePath + '/api/refresh', noCache, requireAuth, (req, res) => {
  data = loadData();
  res.json({ ok: true });
});

app.get(basePath + '/api/data', noCache, requireAuth, (req, res) => res.json(data));

app.listen(PORT, '0.0.0.0', () => {
  console.log(`KPI Dashboard running at http://0.0.0.0:${PORT}`);
});
