const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.static(__dirname));

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
  const result = { years: [], metrics: [], data: {} };
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
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (!h || SKIP.has(h)) continue;
      metricCols.push({ i, name: h });
      if (!metricsOrder.has(h)) metricsOrder.set(h, GROUPS[h] || 'Other');
      result.data[year][h] = new Array(52).fill(null);
    }

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      const week = row[0];
      if (typeof week !== 'number' || week < 1 || week > 52) continue;
      const wi = Math.round(week) - 1;
      for (const { i, name } of metricCols) {
        const val = row[i];
        if (typeof val === 'number' && isFinite(val)) {
          result.data[year][name][wi] = val;
        }
      }
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

app.get('/api/refresh', noCache, (req, res) => {
  data = loadData();
  res.json({ ok: true });
});

app.get('/api/data', noCache, (req, res) => res.json(data));

app.listen(PORT, '0.0.0.0', () => {
  console.log(`KPI Dashboard running at http://0.0.0.0:${PORT}`);
});
