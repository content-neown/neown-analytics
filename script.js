/* ============================================================
   neOwn Analytics — script.js
   Views: Per Day | Monthly | MoM | DoD | YoY
   Interactive tooltips with crosshair on all charts
   ============================================================ */

/* ── Colours ── */
const PINK   = '#F0008C';
const INDIGO = '#3C28B4';
const GRAY   = '#9CA3AF';
const PINK_L = '#FFD6EE';
const IND_L  = '#C7C2F0';
const GREEN  = '#10B981';
const AMBER  = '#F59E0B';
const RED    = '#EF4444';

/* ── State ── */
let currentSheet     = '';
let currentKPIs      = null;
let currentAiCtx     = '';
let chatHistory      = [];
let chartInstances   = {};
let currentView      = 'perday';
let allSheetsCache   = {};   // { name: { rows, kpis, chartData } }
let allSheetNames    = [];
let currentChartData = [];

/* ── View Definitions ── */
const VIEWS = [
  { id: 'perday',  label: 'Per Day',  desc: 'Daily trend within selected month' },
  { id: 'monthly', label: 'Monthly',  desc: 'All months aggregated'             },
  { id: 'mom',     label: 'MoM',      desc: 'Current vs previous month'         },
  { id: 'dod',     label: 'DoD',      desc: 'Day-over-day % change'             },
  { id: 'yoy',     label: 'YoY',      desc: 'Year-over-year comparison'         },
];

/* ══════════════════════════════════════════════════════════════
   CROSSHAIR PLUGIN — vertical dashed line on hover
   ══════════════════════════════════════════════════════════════ */
const CrosshairPlugin = {
  id: 'crosshair',
  afterDraw(chart) {
    if (!chart.tooltip?._active?.length) return;
    const { ctx } = chart;
    const x = chart.tooltip._active[0].element.x;
    const yKey = Object.keys(chart.scales).find(k => chart.scales[k].axis === 'y');
    if (!yKey) return;
    const { top, bottom } = chart.scales[yKey];
    ctx.save();
    ctx.beginPath();
    ctx.moveTo(x, top);
    ctx.lineTo(x, bottom);
    ctx.lineWidth = 1;
    ctx.strokeStyle = 'rgba(156,163,175,0.5)';
    ctx.setLineDash([4, 4]);
    ctx.stroke();
    ctx.restore();
  }
};
Chart.register(CrosshairPlugin);

/* ══════════════════════════════════════════════════════════════
   HELPERS
   ══════════════════════════════════════════════════════════════ */
function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  const s = String(v).replace(/[₹,%\s]/g, '').replace(/,/g, '').trim();
  if (['-','N/A','#DIV/0!','#VALUE!','#REF!','#N/A'].includes(s)) return 0;
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function fINR(v, compact = false) {
  const n = Number(v) || 0;
  if (compact) {
    if (Math.abs(n) >= 10000000) return '₹' + (n / 10000000).toFixed(2) + ' Cr';
    if (Math.abs(n) >= 100000)   return '₹' + (n / 100000).toFixed(2) + ' L';
    if (Math.abs(n) >= 1000)     return '₹' + (n / 1000).toFixed(1) + 'K';
  }
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

function fNum(v, d = 0) {
  return (Number(v) || 0).toLocaleString('en-IN', { maximumFractionDigits: d });
}

function shortDate(d) {
  if (!d) return '';
  const s = String(d);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})/);
  if (m) return m[1] + '/' + m[2];
  const w = s.match(/^(\d{1,2})[\s\-]([A-Za-z]{3})/);
  if (w) return w[1] + ' ' + w[2];
  return s.slice(0, 5);
}

function pct(a, b) { return b > 0 ? Math.round(a / b * 100) : 0; }

function pctChange(curr, prev) {
  if (!prev) return 0;
  return ((curr - prev) / Math.abs(prev)) * 100;
}

function setStatus(state, text) {
  const dot  = document.getElementById('statusDot').querySelector('.dot');
  const span = document.getElementById('statusText');
  dot.className  = 'dot dot--' + state;
  span.textContent = text;
}

function showError(msg) {
  const bar = document.getElementById('errorBar');
  document.getElementById('errorMsg').textContent = msg;
  bar.hidden = !msg;
}

/* Parse "Apr 25", "April 2025", "2025-04" → { month: 3, year: 2025, label } */
function parseSheetDate(name) {
  const MONTHS = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  let m = String(name).trim().match(/^([A-Za-z]+)\s+(\d{2,4})$/);
  if (m) {
    const idx  = MONTHS.indexOf(m[1].slice(0, 3).toLowerCase());
    const year = m[2].length === 2 ? 2000 + parseInt(m[2]) : parseInt(m[2]);
    return idx >= 0 ? { month: idx, year, label: name } : null;
  }
  m = String(name).match(/^(\d{4})[-\/](\d{1,2})$/);
  if (m) return { month: parseInt(m[2]) - 1, year: parseInt(m[1]), label: name };
  return null;
}

/* ══════════════════════════════════════════════════════════════
   TOOLTIP FACTORY
   Dark, styled tooltips with per-chart value formatting
   ══════════════════════════════════════════════════════════════ */
function makeTooltip(labelFn, titleLabels) {
  return {
    mode: 'index',
    intersect: false,
    backgroundColor: 'rgba(17,24,39,0.93)',
    titleColor: '#F9FAFB',
    bodyColor: '#D1D5DB',
    borderColor: 'rgba(255,255,255,0.08)',
    borderWidth: 1,
    padding: 12,
    cornerRadius: 10,
    titleFont: { family: 'Inter', size: 11, weight: '600' },
    bodyFont:  { family: 'Inter', size: 11 },
    callbacks: {
      title: titleLabels
        ? items => titleLabels[items[0].dataIndex] ?? items[0].label
        : undefined,
      label: ctx => labelFn
        ? labelFn(ctx.raw, ctx)
        : `  ${ctx.dataset.label}: ${ctx.formattedValue}`,
      labelColor: ctx => ({
        backgroundColor: ctx.dataset.borderColor || ctx.dataset.backgroundColor || GRAY,
        borderColor: 'transparent',
        borderRadius: 2,
      }),
    },
  };
}

/* Tooltip value formatters */
const ttINR  = (v, c) => `  ${c.dataset.label}: ${fINR(v, true)}`;
const ttX    = (v, c) => `  ${c.dataset.label}: ${(+v || 0).toFixed(2)}x`;
const ttPct  = (v, c) => `  ${c.dataset.label}: ${(+v || 0).toFixed(2)}%`;
const ttNum  = (v, c) => `  ${c.dataset.label}: ${fNum(v)}`;
const ttDPct = (v, c) => `  ${c.dataset.label}: ${v >= 0 ? '+' : ''}${(+v || 0).toFixed(1)}%`;

/* Mixed tooltip for charts with 2 y-axes */
function mixedTooltip(fns, titleLabels) {
  return {
    mode: 'index',
    intersect: false,
    backgroundColor: 'rgba(17,24,39,0.93)',
    titleColor: '#F9FAFB',
    bodyColor: '#D1D5DB',
    borderColor: 'rgba(255,255,255,0.08)',
    borderWidth: 1,
    padding: 12,
    cornerRadius: 10,
    titleFont: { family: 'Inter', size: 11, weight: '600' },
    bodyFont:  { family: 'Inter', size: 11 },
    callbacks: {
      title: titleLabels
        ? items => titleLabels[items[0].dataIndex] ?? items[0].label
        : undefined,
      label: ctx => {
        const fn = fns[ctx.datasetIndex] || ((v, c) => `  ${c.dataset.label}: ${c.formattedValue}`);
        return fn(ctx.raw, ctx);
      },
      labelColor: ctx => ({
        backgroundColor: ctx.dataset.borderColor || ctx.dataset.backgroundColor || GRAY,
        borderColor: 'transparent',
        borderRadius: 2,
      }),
    },
  };
}

/* ══════════════════════════════════════════════════════════════
   SCALE HELPERS
   ══════════════════════════════════════════════════════════════ */
const yINR  = v => v >= 100000 ? '₹' + (v / 100000).toFixed(1) + 'L' : v >= 1000 ? '₹' + (v / 1000).toFixed(0) + 'K' : '₹' + v;
const yX    = v => v.toFixed(1) + 'x';
const yRs   = v => '₹' + v;
const yPct  = v => v.toFixed(1) + '%';
const yDPct = v => (v >= 0 ? '+' : '') + v.toFixed(0) + '%';
const yNum  = v => fNum(v);

function xAxis(labels) {
  const len = labels.length;
  const interval = len <= 8 ? 0 : len <= 16 ? 1 : len <= 31 ? 2 : Math.floor(len / 10);
  return {
    grid: { color: '#F3F4F6' },
    ticks: {
      font: { family: 'Inter', size: 9 },
      maxRotation: 0,
      maxTicksLimit: 12,
      callback: (v, i) => i % (interval + 1) === 0 ? labels[i] : '',
    },
  };
}

function xAxisFull() {
  return {
    grid: { color: '#F3F4F6' },
    ticks: { font: { family: 'Inter', size: 9 }, maxRotation: 30 },
  };
}

function yAxis(fmt) {
  return {
    grid: { color: '#F3F4F6' },
    ticks: { font: { family: 'Inter', size: 9 }, callback: fmt },
  };
}

function baseLegend() {
  return { labels: { font: { family: 'Inter', size: 10 }, boxWidth: 10, padding: 12 } };
}

/* ══════════════════════════════════════════════════════════════
   DATA PROCESSING
   ══════════════════════════════════════════════════════════════ */
function processRows(rawValues) {
  if (!rawValues || rawValues.length < 2) return [];
  const headers = rawValues[0].map(h => String(h).trim());
  return rawValues.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
      return obj;
    })
    .filter(r => {
      const d = String(r['Date'] || '').trim().toLowerCase();
      return d && !d.includes('total') && !d.includes('average') && !d.includes('avg') && d !== 'date';
    });
}

function computeKPIs(rows) {
  if (!rows.length) return {};
  const sum = k => rows.reduce((a, r) => a + parseNum(r[k]), 0);
  const avg = k => {
    const vals = rows.map(r => parseNum(r[k])).filter(v => v > 0);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
  };
  return {
    days:             rows.length,
    totalSpendsGST:   sum('Spends + GST'),
    fbSpendsGST:      sum('FB Spends with GST'),
    googleSpendsGST:  sum('Google Spends with GST'),
    totalDBPurchase:  sum('Total DB Purchase'),
    fbDBPurchase:     sum('FB DB Purchase'),
    googleDBPurchase: sum('Google DB Purchase'),
    totalDBCPS:       avg('Total DB CPS'),
    fbDBCPS:          avg('FB DB CPS'),
    googleDBCPS:      avg('Google DB CPS'),
    revenueTotal:     sum('Revenue (FB + Unknown + Google)'),
    revenueFB:        sum('Revenue (FB UTM)'),
    revenueGoogle:    sum('Revenue (Google UTM)'),
    revenueUnknown:   sum('Revenue (Unknown)'),
    roasGST:          avg('ROAS (FB + Unknown + Google) WITH GST'),
    fbDBROAS:         avg('FB DB ROAS'),
    googleDBROAS:     avg('Google DB ROAS'),
    shopifyRevenue:   sum('Shopify Daily Revenue (TOTAL)'),
    shopifyCount:     sum('Shopify Daily Sales Count (TOTAL)'),
    fbCPM:            avg('FB CPM'),
    fbCTR:            avg('FB CTR'),
    googleCTR:        avg('Google CTR'),
    fbCPC:            avg('FB CPC'),
    googleCPC:        avg('Google CPC'),
    fbClicks:         sum('FB Clicks'),
    googleClicks:     sum('Google Clicks'),
    totalClicks:      sum('Total Clicks'),
  };
}

function computeAggregateKPIs(kpisArray) {
  if (!kpisArray.length) return {};
  const sum = k => kpisArray.reduce((a, kp) => a + (kp[k] || 0), 0);
  const avg = k => {
    const vals = kpisArray.map(kp => kp[k] || 0).filter(v => v > 0);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
  };
  return {
    days: sum('days'), totalSpendsGST: sum('totalSpendsGST'), fbSpendsGST: sum('fbSpendsGST'),
    googleSpendsGST: sum('googleSpendsGST'), totalDBPurchase: sum('totalDBPurchase'),
    fbDBPurchase: sum('fbDBPurchase'), googleDBPurchase: sum('googleDBPurchase'),
    totalDBCPS: avg('totalDBCPS'), fbDBCPS: avg('fbDBCPS'), googleDBCPS: avg('googleDBCPS'),
    revenueTotal: sum('revenueTotal'), revenueFB: sum('revenueFB'),
    revenueGoogle: sum('revenueGoogle'), revenueUnknown: sum('revenueUnknown'),
    roasGST: avg('roasGST'), fbDBROAS: avg('fbDBROAS'), googleDBROAS: avg('googleDBROAS'),
    shopifyRevenue: sum('shopifyRevenue'), shopifyCount: sum('shopifyCount'),
    fbCPM: avg('fbCPM'), fbCTR: avg('fbCTR'), googleCTR: avg('googleCTR'),
    fbCPC: avg('fbCPC'), googleCPC: avg('googleCPC'),
    fbClicks: sum('fbClicks'), googleClicks: sum('googleClicks'), totalClicks: sum('totalClicks'),
  };
}

function buildChartData(rows) {
  return rows.map(r => ({
    date:             String(r['Date'] || '').trim(),
    totalSpends:      parseNum(r['Spends + GST']),
    fbSpends:         parseNum(r['FB Spends with GST']),
    googleSpends:     parseNum(r['Google Spends with GST']),
    revenue:          parseNum(r['Revenue (FB + Unknown + Google)']),
    shopifyRevenue:   parseNum(r['Shopify Daily Revenue (TOTAL)']),
    fbDBPurchase:     parseNum(r['FB DB Purchase']),
    googleDBPurchase: parseNum(r['Google DB Purchase']),
    roas:             parseNum(r['ROAS (FB + Unknown + Google) WITH GST']),
    fbROAS:           parseNum(r['FB DB ROAS']),
    googleROAS:       parseNum(r['Google DB ROAS']),
    totalDBCPS:       parseNum(r['Total DB CPS']),
    fbDBCPS:          parseNum(r['FB DB CPS']),
    googleDBCPS:      parseNum(r['Google DB CPS']),
    fbCTR:            parseNum(r['FB CTR']),
    fbCPM:            parseNum(r['FB CPM']),
  }));
}

function buildAiContext(sheet, kpis, chartData) {
  if (!kpis || !chartData.length) return '';
  const top = [...chartData].sort((a, b) => b.revenue - a.revenue).slice(0, 3)
    .map(d => `${d.date} (Rev ${fINR(d.revenue, true)}, ROAS ${d.roas.toFixed(2)}x)`).join('; ');
  const bot = [...chartData].filter(d => d.totalSpends > 0).sort((a, b) => a.roas - b.roas).slice(0, 3)
    .map(d => `${d.date} (ROAS ${d.roas.toFixed(2)}x, CPS ${fINR(d.totalDBCPS, true)})`).join('; ');
  return `neOwn Performance Marketing — ${sheet} (${kpis.days} days)
SPENDS: Total ${fINR(kpis.totalSpendsGST, true)} | FB ${fINR(kpis.fbSpendsGST, true)} (${pct(kpis.fbSpendsGST, kpis.totalSpendsGST)}%) | Google ${fINR(kpis.googleSpendsGST, true)} (${pct(kpis.googleSpendsGST, kpis.totalSpendsGST)}%)
PURCHASES: Total ${Math.round(kpis.totalDBPurchase)} | FB ${Math.round(kpis.fbDBPurchase)} | Google ${Math.round(kpis.googleDBPurchase)}
AVG CPS: Total ${fINR(kpis.totalDBCPS, true)} | FB ${fINR(kpis.fbDBCPS, true)} | Google ${fINR(kpis.googleDBCPS, true)}
REVENUE: Attribution ${fINR(kpis.revenueTotal, true)} | Shopify ${fINR(kpis.shopifyRevenue, true)} (${Math.round(kpis.shopifyCount)} orders)
ROAS (GST): ${kpis.roasGST.toFixed(2)}x | FB DB ${kpis.fbDBROAS.toFixed(2)}x | Google DB ${kpis.googleDBROAS.toFixed(2)}x
FB ENGAGEMENT: CPM ${fINR(kpis.fbCPM, true)} | CTR ${kpis.fbCTR.toFixed(2)}% | CPC ${fINR(kpis.fbCPC, true)} | Clicks ${fNum(kpis.fbClicks)}
BEST DAYS: ${top}
WORST ROAS DAYS: ${bot}`;
}

/* ══════════════════════════════════════════════════════════════
   KPI RENDERING
   ══════════════════════════════════════════════════════════════ */
function deltaHtml(curr, prev) {
  if (!prev) return '';
  const d   = pctChange(curr, prev);
  const cls = d >= 0 ? 'kpi-delta-up' : 'kpi-delta-down';
  return `<span class="${cls}">${d >= 0 ? '▲' : '▼'} ${Math.abs(d).toFixed(1)}%</span>`;
}

function renderKPIs(kpis, prevKpis) {
  const avgOrder = kpis.shopifyCount > 0 ? kpis.shopifyRevenue / kpis.shopifyCount : 0;
  const d = prevKpis ? (k) => deltaHtml(kpis[k], prevKpis[k]) : () => '';
  const cards = [
    { label: 'Total Spends (GST)',  value: fINR(kpis.totalSpendsGST, true),  sub: `FB ${fINR(kpis.fbSpendsGST, true)} · G ${fINR(kpis.googleSpendsGST, true)}`, delta: d('totalSpendsGST'), accent: true },
    { label: 'Total DB Purchases',  value: fNum(kpis.totalDBPurchase),        sub: `FB ${fNum(kpis.fbDBPurchase)} · G ${fNum(kpis.googleDBPurchase)}`,            delta: d('totalDBPurchase') },
    { label: 'Avg DB CPS',          value: fINR(kpis.totalDBCPS, true),       sub: `FB ${fINR(kpis.fbDBCPS, true)} · G ${fINR(kpis.googleDBCPS, true)}`,          delta: d('totalDBCPS') },
    { label: 'Attribution Revenue', value: fINR(kpis.revenueTotal, true),     sub: `FB ${fINR(kpis.revenueFB, true)} · G ${fINR(kpis.revenueGoogle, true)}`,       delta: d('revenueTotal') },
    { label: 'ROAS (with GST)',     value: kpis.roasGST.toFixed(2) + 'x',    sub: `FB ${kpis.fbDBROAS.toFixed(2)}x · G ${kpis.googleDBROAS.toFixed(2)}x`,         delta: d('roasGST'), accent: true },
    { label: 'Shopify Revenue',     value: fINR(kpis.shopifyRevenue, true),   sub: `${fNum(kpis.shopifyCount)} orders · ₹${fNum(avgOrder)} AOV`,                   delta: d('shopifyRevenue') },
    { label: 'FB CTR',              value: kpis.fbCTR.toFixed(2) + '%',       sub: `CPM ${fINR(kpis.fbCPM, true)} · CPC ${fINR(kpis.fbCPC, true)}`,               delta: d('fbCTR') },
    { label: 'Total Clicks',        value: fNum(kpis.totalClicks),            sub: `FB ${fNum(kpis.fbClicks)} · G ${fNum(kpis.googleClicks)}`,                      delta: d('totalClicks') },
  ];
  document.getElementById('kpiGrid').innerHTML = cards.map(c => `
    <div class="kpi-card ${c.accent ? 'accent' : ''}">
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value-row">
        <div class="kpi-value">${c.value}</div>
        ${c.delta ? `<div class="kpi-delta">${c.delta}</div>` : ''}
      </div>
      <div class="kpi-sub">${c.sub}</div>
    </div>
  `).join('');
}

function showKPISkeleton() {
  document.getElementById('kpiGrid').innerHTML = Array(8).fill('<div class="kpi-skeleton"></div>').join('');
}

/* ══════════════════════════════════════════════════════════════
   CHART HELPERS
   ══════════════════════════════════════════════════════════════ */
function makeChart(id, config) {
  if (chartInstances[id]) { chartInstances[id].destroy(); }
  const canvas = document.getElementById(id);
  if (!canvas) return;
  // Remove any message overlay
  const wrap = canvas.closest('.chart-wrap');
  if (wrap) { const msg = wrap.querySelector('.chart-msg'); if (msg) msg.remove(); }
  chartInstances[id] = new Chart(canvas, config);
}

function setChartTitles(titles) {
  const labels = document.querySelectorAll('#chartsGrid .chart-label');
  titles.forEach((t, i) => { if (labels[i] && t) labels[i].textContent = t; });
}

function showChartMsg(msg) {
  ['cSpends','cRevenue','cPurchases','cROAS','cCPS','cCTR'].forEach(id => {
    if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const wrap = canvas.closest('.chart-wrap');
    if (!wrap) return;
    const existing = wrap.querySelector('.chart-msg');
    if (existing) existing.remove();
    const el = document.createElement('div');
    el.className = 'chart-msg';
    el.textContent = msg;
    wrap.appendChild(el);
  });
}

/* DoD bar dataset factory — green/red based on sign */
function dodBar(data, key, label) {
  return {
    label,
    data: data.map(d => +d[key].toFixed(1)),
    backgroundColor: data.map(d => d[key] >= 0 ? '#10B98166' : '#EF444466'),
    borderColor:     data.map(d => d[key] >= 0 ? '#10B981'   : '#EF4444'),
    borderWidth: 1.5,
    borderRadius: 3,
  };
}

/* ══════════════════════════════════════════════════════════════
   VIEW SWITCHER
   ══════════════════════════════════════════════════════════════ */
function renderViewBar() {
  const bar = document.getElementById('viewBar');
  if (!bar) return;
  bar.innerHTML = VIEWS.map(v => `
    <button
      class="view-btn ${v.id === currentView ? 'active' : ''}"
      onclick="switchView('${v.id}')"
      title="${v.desc}">
      ${v.label}
    </button>
  `).join('');
}

async function switchView(viewId) {
  currentView = viewId;
  renderViewBar();
  await renderView();
}

/* ══════════════════════════════════════════════════════════════
   VIEW DISPATCHER
   ══════════════════════════════════════════════════════════════ */
async function renderView() {
  switch (currentView) {
    case 'perday':  renderPerDayView();      break;
    case 'monthly': await renderMonthlyView(); break;
    case 'mom':     await renderMoMView();     break;
    case 'dod':     renderDoDView();           break;
    case 'yoy':     await renderYoYView();     break;
  }
}

/* ══════════════════════════════════════════════════════════════
   VIEW 1 — PER DAY
   Daily time-series within the selected month
   ══════════════════════════════════════════════════════════════ */
function renderPerDayView() {
  if (!currentChartData.length) return;
  if (currentKPIs) renderKPIs(currentKPIs);
  setChartTitles([
    'Daily Spends incl. GST',
    'Daily Revenue',
    'Daily DB Purchases',
    'ROAS Trend',
    'CPS Trend (Cost per Subscription)',
    'FB CTR % & CPM ₹',
  ]);
  renderPerDayCharts(currentChartData);
}

function renderPerDayCharts(data) {
  const labels    = data.map(d => shortDate(d.date));
  const fullDates = data.map(d => String(d.date));   // full date shown in tooltip title
  const xA = xAxis(labels);

  // 1. Spends
  makeChart('cSpends', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total',    data: data.map(d => d.totalSpends),  borderColor: '#374151', borderWidth: 2,   pointRadius: 0, tension: .3 },
        { label: 'Facebook', data: data.map(d => d.fbSpends),     borderColor: PINK,      borderWidth: 1.5, pointRadius: 0, tension: .3 },
        { label: 'Google',   data: data.map(d => d.googleSpends), borderColor: INDIGO,    borderWidth: 1.5, pointRadius: 0, tension: .3 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR, fullDates) }, scales: { x: xA, y: yAxis(yINR) } },
  });

  // 2. Revenue
  makeChart('cRevenue', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Shopify Revenue',     data: data.map(d => d.shopifyRevenue), borderColor: INDIGO, backgroundColor: 'rgba(60,40,180,.08)', fill: true, borderWidth: 2, pointRadius: 0, tension: .3 },
        { label: 'Attribution Revenue', data: data.map(d => d.revenue),        borderColor: PINK,   borderWidth: 2, pointRadius: 0, tension: .3 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR, fullDates) }, scales: { x: xA, y: yAxis(yINR) } },
  });

  // 3. Purchases
  makeChart('cPurchases', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Facebook', data: data.map(d => d.fbDBPurchase),     backgroundColor: PINK,   borderRadius: 3 },
        { label: 'Google',   data: data.map(d => d.googleDBPurchase), backgroundColor: INDIGO, borderRadius: 3 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttNum, fullDates) }, scales: { x: xA, y: yAxis(yNum) } },
  });

  // 4. ROAS
  makeChart('cROAS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total ROAS',  data: data.map(d => +d.roas.toFixed(2)),       borderColor: '#374151', borderWidth: 2,   pointRadius: 0, tension: .3 },
        { label: 'FB ROAS',     data: data.map(d => +d.fbROAS.toFixed(2)),     borderColor: PINK,      borderWidth: 1.5, pointRadius: 0, tension: .3, borderDash: [5,3] },
        { label: 'Google ROAS', data: data.map(d => +d.googleROAS.toFixed(2)), borderColor: INDIGO,    borderWidth: 1.5, pointRadius: 0, tension: .3, borderDash: [5,3] },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttX, fullDates) }, scales: { x: xA, y: yAxis(yX) } },
  });

  // 5. CPS
  makeChart('cCPS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total CPS',  data: data.map(d => Math.round(d.totalDBCPS)),  borderColor: '#374151', borderWidth: 2,   pointRadius: 0, tension: .3 },
        { label: 'FB CPS',     data: data.map(d => Math.round(d.fbDBCPS)),     borderColor: PINK,      borderWidth: 1.5, pointRadius: 0, tension: .3 },
        { label: 'Google CPS', data: data.map(d => Math.round(d.googleDBCPS)), borderColor: INDIGO,    borderWidth: 1.5, pointRadius: 0, tension: .3 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR, fullDates) }, scales: { x: xA, y: yAxis(yRs) } },
  });

  // 6. CTR + CPM (dual axis)
  makeChart('cCTR', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'FB CTR %', data: data.map(d => +d.fbCTR.toFixed(2)), backgroundColor: PINK_L, borderRadius: 3, yAxisID: 'y' },
        { label: 'FB CPM ₹', data: data.map(d => Math.round(d.fbCPM)),  type: 'line', borderColor: PINK, borderWidth: 2, pointRadius: 0, tension: .3, yAxisID: 'y2' },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: baseLegend(), tooltip: mixedTooltip([ttPct, ttINR], fullDates) },
      scales: {
        x: xA,
        y:  { ...yAxis(yPct), position: 'left' },
        y2: { ...yAxis(yRs),  position: 'right', grid: { drawOnChartArea: false } },
      },
    },
  });
}

/* ══════════════════════════════════════════════════════════════
   VIEW 2 — MONTHLY SUMMARY
   All months aggregated, each month = one data point
   ══════════════════════════════════════════════════════════════ */
async function renderMonthlyView() {
  await ensureAllLoaded();
  const months = allSheetNames
    .map(name => ({ name, kpis: allSheetsCache[name]?.kpis }))
    .filter(m => m.kpis);

  if (!months.length) { showChartMsg('Loading monthly data…'); return; }

  const ytd = computeAggregateKPIs(months.map(m => m.kpis));
  renderKPIs(ytd);

  setChartTitles([
    'Monthly Total Spends (FB + Google)',
    'Monthly Revenue + ROAS',
    'Monthly DB Purchases',
    'Monthly Avg ROAS',
    'Monthly Avg CPS',
    'Monthly FB CTR % & CPM ₹',
  ]);

  const labels = months.map(m => m.name);
  const xA = xAxisFull();

  // 1. Monthly Spends (stacked bars)
  makeChart('cSpends', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Facebook', data: months.map(m => Math.round(m.kpis.fbSpendsGST)),     backgroundColor: PINK + 'CC', borderRadius: 4 },
        { label: 'Google',   data: months.map(m => Math.round(m.kpis.googleSpendsGST)), backgroundColor: INDIGO + 'CC', borderRadius: 4 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR) }, scales: { x: xA, y: yAxis(yINR) } },
  });

  // 2. Monthly Revenue + ROAS (dual axis)
  makeChart('cRevenue', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Shopify Revenue',     data: months.map(m => Math.round(m.kpis.shopifyRevenue)), backgroundColor: IND_L + 'CC', borderRadius: 4 },
        { label: 'Attribution Revenue', data: months.map(m => Math.round(m.kpis.revenueTotal)),   backgroundColor: PINK + 'CC',  borderRadius: 4 },
        { label: 'ROAS',                data: months.map(m => +m.kpis.roasGST.toFixed(2)),        type: 'line', borderColor: INDIGO, backgroundColor: 'transparent', borderWidth: 2.5, pointRadius: 5, pointBackgroundColor: INDIGO, tension: .3, yAxisID: 'y2' },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: baseLegend(), tooltip: mixedTooltip([ttINR, ttINR, ttX], labels) },
      scales: {
        x: xA,
        y:  { ...yAxis(yINR), position: 'left' },
        y2: { ...yAxis(yX),   position: 'right', grid: { drawOnChartArea: false } },
      },
    },
  });

  // 3. Monthly Purchases
  makeChart('cPurchases', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Facebook', data: months.map(m => Math.round(m.kpis.fbDBPurchase)),     backgroundColor: PINK + 'CC',   borderRadius: 4 },
        { label: 'Google',   data: months.map(m => Math.round(m.kpis.googleDBPurchase)), backgroundColor: INDIGO + 'CC', borderRadius: 4 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttNum) }, scales: { x: xA, y: yAxis(yNum) } },
  });

  // 4. Monthly ROAS trend
  makeChart('cROAS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total ROAS',  data: months.map(m => +m.kpis.roasGST.toFixed(2)),     borderColor: '#374151', backgroundColor: 'rgba(55,65,81,.06)', fill: true, borderWidth: 2.5, pointRadius: 5, pointBackgroundColor: '#374151', tension: .3 },
        { label: 'FB ROAS',     data: months.map(m => +m.kpis.fbDBROAS.toFixed(2)),    borderColor: PINK,   borderWidth: 1.5, pointRadius: 3, tension: .3 },
        { label: 'Google ROAS', data: months.map(m => +m.kpis.googleDBROAS.toFixed(2)), borderColor: INDIGO, borderWidth: 1.5, pointRadius: 3, tension: .3 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttX) }, scales: { x: xA, y: yAxis(yX) } },
  });

  // 5. Monthly CPS trend
  makeChart('cCPS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total CPS',  data: months.map(m => Math.round(m.kpis.totalDBCPS)),  borderColor: '#374151', backgroundColor: 'rgba(55,65,81,.06)', fill: true, borderWidth: 2.5, pointRadius: 5, pointBackgroundColor: '#374151', tension: .3 },
        { label: 'FB CPS',     data: months.map(m => Math.round(m.kpis.fbDBCPS)),     borderColor: PINK,   borderWidth: 1.5, pointRadius: 3, tension: .3 },
        { label: 'Google CPS', data: months.map(m => Math.round(m.kpis.googleDBCPS)), borderColor: INDIGO, borderWidth: 1.5, pointRadius: 3, tension: .3 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR) }, scales: { x: xA, y: yAxis(yRs) } },
  });

  // 6. Monthly CTR + CPM
  makeChart('cCTR', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'FB CTR %', data: months.map(m => +m.kpis.fbCTR.toFixed(2)), backgroundColor: PINK_L + 'CC', borderRadius: 4, yAxisID: 'y' },
        { label: 'FB CPM ₹', data: months.map(m => Math.round(m.kpis.fbCPM)),  type: 'line', borderColor: PINK, borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: PINK, tension: .3, yAxisID: 'y2' },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: baseLegend(), tooltip: mixedTooltip([ttPct, ttINR]) },
      scales: {
        x: xA,
        y:  { ...yAxis(yPct), position: 'left' },
        y2: { ...yAxis(yRs),  position: 'right', grid: { drawOnChartArea: false } },
      },
    },
  });
}

/* ══════════════════════════════════════════════════════════════
   VIEW 3 — MONTH ON MONTH (MoM)
   Current month vs previous month, aligned by day index
   ══════════════════════════════════════════════════════════════ */
async function renderMoMView() {
  const currIdx = allSheetNames.indexOf(currentSheet);
  if (currIdx <= 0) {
    showChartMsg('No previous month available for comparison.');
    if (currentKPIs) renderKPIs(currentKPIs);
    return;
  }

  const prevName = allSheetNames[currIdx - 1];
  await ensureSheetLoaded(prevName);

  const curr = allSheetsCache[currentSheet];
  const prev = allSheetsCache[prevName];
  if (!curr || !prev) { showChartMsg('Data loading…'); return; }

  // KPIs with delta vs previous month
  renderKPIs(curr.kpis, prev.kpis);

  setChartTitles([
    `Spends: ${currentSheet} vs ${prevName}`,
    `Revenue: ${currentSheet} vs ${prevName}`,
    `DB Purchases: ${currentSheet} vs ${prevName}`,
    `ROAS: ${currentSheet} vs ${prevName}`,
    `CPS: ${currentSheet} vs ${prevName}`,
    `FB CTR: ${currentSheet} vs ${prevName}`,
  ]);

  const cD = curr.chartData;
  const pD = prev.chartData;
  const n  = Math.max(cD.length, pD.length);
  const labels = Array.from({ length: n }, (_, i) => `Day ${i + 1}`);
  // Tooltip title shows actual dates from both months
  const fullDates = Array.from({ length: n }, (_, i) => {
    const cd = cD[i] ? cD[i].date : '—';
    const pd = pD[i] ? pD[i].date : '—';
    return `Day ${i + 1}  ·  ${cd}  /  ${pd}`;
  });
  const xA = xAxis(labels);

  const currColor = PINK;
  const prevColor = GRAY;

  function moMLine(data, key, label, color, dashed = false) {
    return {
      label,
      data: data.map(d => d ? d[key] : null),
      borderColor: color,
      borderWidth: dashed ? 2 : 2.5,
      pointRadius: 0,
      tension: .3,
      spanGaps: true,
      ...(dashed ? { borderDash: [5, 4] } : {}),
    };
  }

  // 1. Spends MoM
  makeChart('cSpends', {
    type: 'line',
    data: {
      labels,
      datasets: [
        moMLine(cD, 'totalSpends',  currentSheet, currColor),
        moMLine(pD, 'totalSpends',  prevName,     prevColor, true),
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR, fullDates) }, scales: { x: xA, y: yAxis(yINR) } },
  });

  // 2. Shopify Revenue MoM
  makeChart('cRevenue', {
    type: 'line',
    data: {
      labels,
      datasets: [
        moMLine(cD, 'shopifyRevenue', currentSheet, INDIGO),
        moMLine(pD, 'shopifyRevenue', prevName,     IND_L,   true),
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR, fullDates) }, scales: { x: xA, y: yAxis(yINR) } },
  });

  // 3. Purchases MoM (FB + Google both months)
  makeChart('cPurchases', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: `${currentSheet} FB`,     data: cD.map(d => d.fbDBPurchase),     borderColor: PINK,   borderWidth: 2,   pointRadius: 0, tension: .3 },
        { label: `${currentSheet} Google`, data: cD.map(d => d.googleDBPurchase), borderColor: INDIGO, borderWidth: 2,   pointRadius: 0, tension: .3 },
        { label: `${prevName} FB`,         data: pD.map(d => d.fbDBPurchase),     borderColor: PINK_L, borderWidth: 1.5, pointRadius: 0, tension: .3, borderDash: [5,4] },
        { label: `${prevName} Google`,     data: pD.map(d => d.googleDBPurchase), borderColor: IND_L,  borderWidth: 1.5, pointRadius: 0, tension: .3, borderDash: [5,4] },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttNum, fullDates) }, scales: { x: xA, y: yAxis(yNum) } },
  });

  // 4. ROAS MoM
  makeChart('cROAS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        moMLine(cD.map(d => ({ roas: +d.roas.toFixed(2) })), 'roas', currentSheet, '#374151'),
        moMLine(pD.map(d => ({ roas: +d.roas.toFixed(2) })), 'roas', prevName,     GRAY, true),
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttX, fullDates) }, scales: { x: xA, y: yAxis(yX) } },
  });

  // 5. CPS MoM
  makeChart('cCPS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        moMLine(cD.map(d => ({ cps: Math.round(d.totalDBCPS) })), 'cps', currentSheet, '#374151'),
        moMLine(pD.map(d => ({ cps: Math.round(d.totalDBCPS) })), 'cps', prevName,     GRAY, true),
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttINR, fullDates) }, scales: { x: xA, y: yAxis(yRs) } },
  });

  // 6. CTR MoM
  makeChart('cCTR', {
    type: 'line',
    data: {
      labels,
      datasets: [
        moMLine(cD.map(d => ({ ctr: +d.fbCTR.toFixed(2) })), 'ctr', currentSheet, PINK),
        moMLine(pD.map(d => ({ ctr: +d.fbCTR.toFixed(2) })), 'ctr', prevName,     PINK_L, true),
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: baseLegend(), tooltip: makeTooltip(ttPct, fullDates) }, scales: { x: xA, y: yAxis(yPct) } },
  });
}

/* ══════════════════════════════════════════════════════════════
   VIEW 4 — DAY ON DAY (DoD)
   % change vs previous day within the selected month
   Green bar = positive change, red = negative
   ══════════════════════════════════════════════════════════════ */
function renderDoDView() {
  if (!currentChartData.length) return;
  if (currentKPIs) renderKPIs(currentKPIs);

  const data = currentChartData;
  if (data.length < 2) { showChartMsg('Need at least 2 days for DoD view.'); return; }

  // Compute day-over-day % changes (skip first day - no previous)
  const dod = data.slice(1).map((d, i) => {
    const p = data[i];
    return {
      date:           d.date,
      fbSpends:       pctChange(d.fbSpends,      p.fbSpends),
      googleSpends:   pctChange(d.googleSpends,  p.googleSpends),
      revenue:        pctChange(d.shopifyRevenue, p.shopifyRevenue),
      attrRevenue:    pctChange(d.revenue,        p.revenue),
      fbPurchases:    pctChange(d.fbDBPurchase,   p.fbDBPurchase),
      googlePurchases:pctChange(d.googleDBPurchase, p.googleDBPurchase),
      roas:           pctChange(d.roas,           p.roas),
      fbROAS:         pctChange(d.fbROAS,         p.fbROAS),
      cps:            pctChange(d.totalDBCPS,     p.totalDBCPS),
      fbCPS:          pctChange(d.fbDBCPS,        p.fbDBCPS),
      ctr:            pctChange(d.fbCTR,          p.fbCTR),
      cpm:            pctChange(d.fbCPM,          p.fbCPM),
    };
  });

  const labels    = dod.map(d => shortDate(d.date));
  const fullDates = dod.map(d => String(d.date));
  const xA = xAxis(labels);
  const opts = () => ({
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: baseLegend(), tooltip: makeTooltip(ttDPct, fullDates) },
    scales: { x: xA, y: yAxis(yDPct) },
  });

  setChartTitles([
    'Spends DoD Change %',
    'Revenue DoD Change %',
    'DB Purchases DoD Change %',
    'ROAS DoD Change %',
    'CPS DoD Change %',
    'FB CTR & CPM DoD Change %',
  ]);

  // 1. Spends DoD
  makeChart('cSpends', {
    type: 'bar',
    data: { labels, datasets: [ dodBar(dod, 'fbSpends', 'FB Spends'), dodBar(dod, 'googleSpends', 'Google Spends') ] },
    options: opts(),
  });

  // 2. Revenue DoD
  makeChart('cRevenue', {
    type: 'bar',
    data: { labels, datasets: [ dodBar(dod, 'revenue', 'Shopify Revenue'), dodBar(dod, 'attrRevenue', 'Attribution Revenue') ] },
    options: opts(),
  });

  // 3. Purchases DoD
  makeChart('cPurchases', {
    type: 'bar',
    data: { labels, datasets: [ dodBar(dod, 'fbPurchases', 'FB Purchases'), dodBar(dod, 'googlePurchases', 'Google Purchases') ] },
    options: opts(),
  });

  // 4. ROAS DoD
  makeChart('cROAS', {
    type: 'bar',
    data: { labels, datasets: [ dodBar(dod, 'roas', 'Total ROAS'), dodBar(dod, 'fbROAS', 'FB ROAS') ] },
    options: opts(),
  });

  // 5. CPS DoD
  makeChart('cCPS', {
    type: 'bar',
    data: { labels, datasets: [ dodBar(dod, 'cps', 'Total CPS'), dodBar(dod, 'fbCPS', 'FB CPS') ] },
    options: opts(),
  });

  // 6. CTR + CPM DoD
  makeChart('cCTR', {
    type: 'bar',
    data: { labels, datasets: [ dodBar(dod, 'ctr', 'FB CTR'), dodBar(dod, 'cpm', 'FB CPM') ] },
    options: opts(),
  });
}

/* ══════════════════════════════════════════════════════════════
   VIEW 5 — YEAR ON YEAR (YoY)
   Monthly data grouped and coloured by year
   Falls back to Monthly view if sheet names can't be parsed
   ══════════════════════════════════════════════════════════════ */
async function renderYoYView() {
  await ensureAllLoaded();

  const parsed = allSheetNames
    .map(name => ({ name, p: parseSheetDate(name) }))
    .filter(s => s.p && allSheetsCache[s.name]?.kpis);

  if (!parsed.length) {
    // Names don't have year/month — fall back to monthly summary
    await renderMonthlyView();
    return;
  }

  // Group by year
  const yearGroups = {};
  parsed.forEach(({ name, p }) => {
    const yr = String(p.year);
    if (!yearGroups[yr]) yearGroups[yr] = {};
    yearGroups[yr][p.month] = allSheetsCache[name].kpis;
  });

  const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const years = Object.keys(yearGroups).sort();
  const allMonths = [...new Set(parsed.map(s => s.p.month))].sort((a, b) => a - b);
  const labels = allMonths.map(m => MONTH_NAMES[m]);
  const YEAR_COLORS = [INDIGO, PINK, GREEN, AMBER, '#8B5CF6', '#06B6D4'];

  // KPIs: most recent year's aggregate
  const latestYear = years[years.length - 1];
  const latestKpis = Object.values(yearGroups[latestYear]);
  renderKPIs(computeAggregateKPIs(latestKpis));

  setChartTitles([
    'Monthly Spends by Year',
    'Monthly Revenue by Year',
    'Monthly DB Purchases by Year',
    'Monthly Avg ROAS by Year',
    'Monthly Avg CPS by Year',
    'Monthly FB CTR by Year',
  ]);

  const xA = xAxisFull();

  function yoyDatasets(fn) {
    return years.map((yr, i) => ({
      label: yr,
      data: allMonths.map(m => {
        const kp = yearGroups[yr][m];
        return kp ? fn(kp) : null;
      }),
      borderColor:          YEAR_COLORS[i % YEAR_COLORS.length],
      backgroundColor:      YEAR_COLORS[i % YEAR_COLORS.length] + '22',
      borderWidth: 2.5,
      pointRadius: 5,
      pointBackgroundColor: YEAR_COLORS[i % YEAR_COLORS.length],
      tension: .3,
      spanGaps: true,
    }));
  }

  const lineOpts = (tt, yFmt) => ({
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: baseLegend(), tooltip: makeTooltip(tt) },
    scales: { x: xA, y: yAxis(yFmt) },
  });

  makeChart('cSpends',    { type: 'line', data: { labels, datasets: yoyDatasets(k => Math.round(k.totalSpendsGST))   }, options: lineOpts(ttINR, yINR) });
  makeChart('cRevenue',   { type: 'line', data: { labels, datasets: yoyDatasets(k => Math.round(k.revenueTotal))     }, options: lineOpts(ttINR, yINR) });
  makeChart('cPurchases', { type: 'line', data: { labels, datasets: yoyDatasets(k => Math.round(k.totalDBPurchase))  }, options: lineOpts(ttNum,  yNum) });
  makeChart('cROAS',      { type: 'line', data: { labels, datasets: yoyDatasets(k => +k.roasGST.toFixed(2))         }, options: lineOpts(ttX,   yX)   });
  makeChart('cCPS',       { type: 'line', data: { labels, datasets: yoyDatasets(k => Math.round(k.totalDBCPS))       }, options: lineOpts(ttINR, yRs)  });
  makeChart('cCTR',       { type: 'line', data: { labels, datasets: yoyDatasets(k => +k.fbCTR.toFixed(2))           }, options: lineOpts(ttPct, yPct) });
}

/* ══════════════════════════════════════════════════════════════
   DATA LOADING
   ══════════════════════════════════════════════════════════════ */
async function ensureSheetLoaded(name) {
  if (allSheetsCache[name]) return;
  try {
    const res  = await fetch('/api/sheets?sheet=' + encodeURIComponent(name));
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    const rows      = processRows(data.values || []);
    const kpis      = computeKPIs(rows);
    const chartData = buildChartData(rows);
    allSheetsCache[name] = { rows, kpis, chartData };
  } catch (err) {
    console.warn('Could not load sheet:', name, err.message);
  }
}

async function ensureAllLoaded() {
  const toLoad = allSheetNames.filter(n => !allSheetsCache[n]);
  if (!toLoad.length) return;
  setStatus('loading', 'Loading all months…');
  await Promise.allSettled(toLoad.map(ensureSheetLoaded));
  setStatus('ok', 'All months loaded');
}

async function loadSheet(sheetName) {
  currentSheet = sheetName;
  setStatus('loading', 'Loading…');
  showError('');
  showKPISkeleton();
  document.getElementById('insightsBody').innerHTML = '<div class="insights-empty">Select a month and click Generate Insights.</div>';
  document.getElementById('insightsSub').textContent = 'Claude-powered analysis · ' + sheetName;

  try {
    await ensureSheetLoaded(sheetName);
    const cached = allSheetsCache[sheetName];
    if (!cached) throw new Error('Failed to load sheet data');

    currentKPIs      = cached.kpis;
    currentChartData = cached.chartData;
    currentAiCtx     = buildAiContext(sheetName, currentKPIs, currentChartData);

    setStatus('ok', cached.rows.length + ' days loaded');
    await renderView();

    // Preload all other sheets in background for aggregate views
    setTimeout(() => ensureAllLoaded(), 600);
  } catch (err) {
    showError('Error loading ' + sheetName + ': ' + err.message);
    setStatus('error', 'Error');
  }
}

async function loadSheetList() {
  setStatus('loading', 'Connecting…');
  try {
    const res  = await fetch('/api/sheets?action=list');
    const data = await res.json();

    if (data.error) {
      showError(data.error);
      document.getElementById('setupHint').hidden = false;
      document.getElementById('monthTabs').innerHTML = '';
      setStatus('error', 'Config error');
      return;
    }

    allSheetNames = data.sheets || [];
    if (!allSheetNames.length) {
      showError('No sheets found in the spreadsheet.');
      setStatus('error', 'No sheets');
      return;
    }

    const tabsEl = document.getElementById('monthTabs');
    tabsEl.innerHTML = allSheetNames.map(s =>
      `<button class="tab" onclick="switchTab(this,'${s}')">${s}</button>`
    ).join('');

    const lastTab = tabsEl.lastElementChild;
    lastTab.classList.add('active');
    loadSheet(lastTab.textContent);
  } catch (err) {
    showError('Could not reach /api/sheets — check GOOGLE_SHEETS_API_KEY in Vercel Environment Variables.');
    document.getElementById('setupHint').hidden = false;
    setStatus('error', 'Error');
  }
}

function switchTab(btn, sheetName) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  loadSheet(sheetName);
}

/* ══════════════════════════════════════════════════════════════
   AI INSIGHTS
   ══════════════════════════════════════════════════════════════ */
async function generateInsights() {
  if (!currentAiCtx) return;
  const btn  = document.getElementById('insightsBtn');
  const body = document.getElementById('insightsBody');
  btn.disabled = true;
  btn.textContent = 'Analyzing…';
  body.innerHTML = `<div class="insights-skeleton">${Array(5).fill('<div style="width:75%"></div>').join('')}</div>`;

  try {
    const res = await fetch('/api/insights', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ summary: currentAiCtx, sheetName: currentSheet }),
    });
    const ins = await res.json();
    if (ins.error) throw new Error(ins.error);
    renderInsights(ins);
  } catch (err) {
    body.innerHTML = `<div class="insights-empty" style="color:#B91C1C">Error: ${err.message}</div>`;
  } finally {
    btn.disabled = false;
    btn.textContent = '↺ Refresh';
  }
}

function renderInsights(ins) {
  const body = document.getElementById('insightsBody');
  const li   = arr => (arr || []).map(t => `<li><span class="insights-dot">›</span>${t}</li>`).join('');
  const extra = (ins.aiCharts || []).map(c => `
    <div class="insights-extra-item">
      <span style="font-size:20px">📈</span>
      <div>
        <div class="insights-extra-title">${c.title}</div>
        <div class="insights-extra-body">${c.insight}</div>
      </div>
    </div>`).join('');
  body.innerHTML = `
    ${ins.summary ? `<div class="insights-summary">${ins.summary}</div>` : ''}
    <div class="insights-cols">
      <div class="insights-col green"><div class="insights-col-title">🔥 Highlights</div><ul>${li(ins.highlights)}</ul></div>
      <div class="insights-col orange"><div class="insights-col-title">⚠️ Watch out</div><ul>${li(ins.concerns)}</ul></div>
      <div class="insights-col indigo"><div class="insights-col-title">💡 Actions</div><ul>${li(ins.recommendations)}</ul></div>
    </div>
    ${extra ? `<div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#9CA3AF;margin:14px 0 8px">📊 Additional angles to explore</div><div class="insights-extra">${extra}</div></div>` : ''}
  `;
}

/* ══════════════════════════════════════════════════════════════
   CHATBOT
   ══════════════════════════════════════════════════════════════ */
function toggleChat() {
  const win = document.getElementById('chatWindow');
  win.classList.toggle('open');
  if (win.classList.contains('open')) document.getElementById('chatInput').focus();
}

async function sendChat(preset) {
  const input = document.getElementById('chatInput');
  const text  = preset || input.value.trim();
  if (!text) return;
  input.value = '';

  if (!preset || chatHistory.length > 1) {
    document.getElementById('chatStarters').style.display = 'none';
  }

  appendMsg('user', text);
  chatHistory.push({ role: 'user', content: text });

  const typingId = 'typing-' + Date.now();
  appendTyping(typingId);

  try {
    const res  = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: chatHistory.slice(-12), context: currentAiCtx }),
    });
    const data = await res.json();
    removeTyping(typingId);
    const reply = data.reply || 'Sorry, something went wrong.';
    appendMsg('bot', reply);
    chatHistory.push({ role: 'assistant', content: reply });
  } catch {
    removeTyping(typingId);
    appendMsg('bot', 'Network error. Please try again.');
  }
}

function appendMsg(role, text) {
  const msgs = document.getElementById('chatMessages');
  const div  = document.createElement('div');
  div.className = 'msg msg--' + (role === 'user' ? 'user' : 'bot');
  div.textContent = text;
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
}

function appendTyping(id) {
  const msgs = document.getElementById('chatMessages');
  const div  = document.createElement('div');
  div.id = id;
  div.className = 'msg msg--bot';
  div.innerHTML = '<div class="typing-dots"><span></span><span></span><span></span></div>';
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
}

function removeTyping(id) {
  const el = document.getElementById(id);
  if (el) el.remove();
}

/* ══════════════════════════════════════════════════════════════
   INIT
   ══════════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  renderViewBar();
  loadSheetList();
});
