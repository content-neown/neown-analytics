/* ============================================================
   neOwn Analytics — script.js
   Dashboards: Performance Marketing | Revenue Achievement
   Views: Per Day | Monthly | MoM | DoD | YoY
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
const TEAL   = '#06B6D4';

/* ── State ── */
let currentDashboard = 'ads';          // 'ads' | 'revenue'
let currentSheet     = '';
let currentKPIs      = null;
let currentAiCtx     = '';
let chatHistory      = [];
let chartInstances   = {};
let currentView      = 'perday';
let currentChartData = [];             // ads daily rows
let currentRevData   = null;           // { daily, target, kpis }

// Separate caches per source so switching dashboards keeps data
const cacheBySource = { ads: {}, revenue: {} };
const namesBySource = { ads: [], revenue: [] };

const cache      = () => cacheBySource[currentDashboard];
const sheetNames = () => namesBySource[currentDashboard];

/* ── View definitions (shared) ── */
const VIEWS = [
  { id: 'perday',  label: 'Per Day'  },
  { id: 'monthly', label: 'Monthly'  },
  { id: 'mom',     label: 'MoM'      },
  { id: 'dod',     label: 'DoD'      },
  { id: 'yoy',     label: 'YoY'      },
];

/* ══════════════════════════════════════════════════════════════
   CROSSHAIR PLUGIN
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
    if (Math.abs(n) >= 10000000) return '₹' + (n/10000000).toFixed(2) + ' Cr';
    if (Math.abs(n) >= 100000)   return '₹' + (n/100000).toFixed(2) + ' L';
    if (Math.abs(n) >= 1000)     return '₹' + (n/1000).toFixed(1) + 'K';
  }
  return '₹' + Math.round(n).toLocaleString('en-IN');
}
function fNum(v, d = 0) { return (Number(v)||0).toLocaleString('en-IN', { maximumFractionDigits: d }); }
function fPct(v, d = 1) { return (Number(v)||0).toFixed(d) + '%'; }
function shortDate(d) {
  if (!d) return '';
  const s = String(d);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})/);
  if (m) return m[1] + '/' + m[2];
  const w = s.match(/^(\d{1,2})[\s\-]([A-Za-z]{3})/);
  if (w) return w[1] + ' ' + w[2];
  return s.slice(0, 6);
}
function pct(a, b) { return b > 0 ? Math.round(a/b*100) : 0; }
function pctChange(curr, prev) { return prev ? ((curr-prev)/Math.abs(prev))*100 : 0; }
function setStatus(state, text) {
  document.getElementById('statusDot').querySelector('.dot').className = 'dot dot--' + state;
  document.getElementById('statusText').textContent = text;
}
function showError(msg) {
  const bar = document.getElementById('errorBar');
  document.getElementById('errorMsg').textContent = msg;
  bar.hidden = !msg;
}
function parseSheetDate(name) {
  const MONTHS = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  let m = String(name).trim().match(/^([A-Za-z]+)\s+(\d{2,4})$/);
  if (m) {
    const idx  = MONTHS.indexOf(m[1].slice(0,3).toLowerCase());
    const year = m[2].length === 2 ? 2000 + parseInt(m[2]) : parseInt(m[2]);
    return idx >= 0 ? { month: idx, year, label: name } : null;
  }
  m = String(name).match(/^(\d{4})[-\/](\d{1,2})$/);
  if (m) return { month: parseInt(m[2])-1, year: parseInt(m[1]), label: name };
  return null;
}

/* ══════════════════════════════════════════════════════════════
   TOOLTIP FACTORY
   ══════════════════════════════════════════════════════════════ */
function makeTooltip(labelFn, titleLabels) {
  return {
    mode: 'index', intersect: false,
    backgroundColor: 'rgba(17,24,39,0.93)',
    titleColor: '#F9FAFB', bodyColor: '#D1D5DB',
    borderColor: 'rgba(255,255,255,0.08)', borderWidth: 1,
    padding: 12, cornerRadius: 10,
    titleFont: { family: 'Inter', size: 11, weight: '600' },
    bodyFont:  { family: 'Inter', size: 11 },
    callbacks: {
      title: titleLabels ? items => titleLabels[items[0].dataIndex] ?? items[0].label : undefined,
      label: ctx => labelFn ? labelFn(ctx.raw, ctx) : `  ${ctx.dataset.label}: ${ctx.formattedValue}`,
      labelColor: ctx => ({ backgroundColor: ctx.dataset.borderColor || ctx.dataset.backgroundColor || GRAY, borderColor: 'transparent', borderRadius: 2 }),
    },
  };
}
function mixedTooltip(fns, titleLabels) {
  return {
    mode: 'index', intersect: false,
    backgroundColor: 'rgba(17,24,39,0.93)',
    titleColor: '#F9FAFB', bodyColor: '#D1D5DB',
    borderColor: 'rgba(255,255,255,0.08)', borderWidth: 1,
    padding: 12, cornerRadius: 10,
    titleFont: { family: 'Inter', size: 11, weight: '600' },
    bodyFont:  { family: 'Inter', size: 11 },
    callbacks: {
      title: titleLabels ? items => titleLabels[items[0].dataIndex] ?? items[0].label : undefined,
      label: ctx => { const fn = fns[ctx.datasetIndex] || ((v,c) => `  ${c.dataset.label}: ${c.formattedValue}`); return fn(ctx.raw, ctx); },
      labelColor: ctx => ({ backgroundColor: ctx.dataset.borderColor || ctx.dataset.backgroundColor || GRAY, borderColor: 'transparent', borderRadius: 2 }),
    },
  };
}
const ttINR  = (v,c) => `  ${c.dataset.label}: ${fINR(v,true)}`;
const ttX    = (v,c) => `  ${c.dataset.label}: ${(+v||0).toFixed(2)}x`;
const ttPct  = (v,c) => `  ${c.dataset.label}: ${(+v||0).toFixed(2)}%`;
const ttNum  = (v,c) => `  ${c.dataset.label}: ${fNum(v)}`;
const ttDPct = (v,c) => `  ${c.dataset.label}: ${v>=0?'+':''}${(+v||0).toFixed(1)}%`;

/* ── Scale helpers ── */
const yINR  = v => v>=100000?'₹'+(v/100000).toFixed(1)+'L':v>=1000?'₹'+(v/1000).toFixed(0)+'K':'₹'+v;
const yX    = v => v.toFixed(1)+'x';
const yRs   = v => '₹'+v;
const yPct  = v => v.toFixed(1)+'%';
const yDPct = v => (v>=0?'+':'')+v.toFixed(0)+'%';
const yNum  = v => fNum(v);
function xAxis(labels) {
  const len = labels.length;
  const interval = len<=8?0:len<=16?1:len<=31?2:Math.floor(len/10);
  return { grid:{color:'#F3F4F6'}, ticks:{ font:{family:'Inter',size:9}, maxRotation:0, maxTicksLimit:12, callback:(v,i)=>i%(interval+1)===0?labels[i]:'' } };
}
function xAxisFull() { return { grid:{color:'#F3F4F6'}, ticks:{ font:{family:'Inter',size:9}, maxRotation:30 } }; }
function yAxis(fmt) { return { grid:{color:'#F3F4F6'}, ticks:{ font:{family:'Inter',size:9}, callback:fmt } }; }
function baseLegend() { return { labels:{ font:{family:'Inter',size:10}, boxWidth:10, padding:12 } }; }

/* ══════════════════════════════════════════════════════════════
   ADS DATA PROCESSING
   ══════════════════════════════════════════════════════════════ */
function processRows(rawValues) {
  if (!rawValues||rawValues.length<2) return [];
  const headers = rawValues[0].map(h=>String(h).trim());
  return rawValues.slice(1).map(row=>{
    const obj={}; headers.forEach((h,i)=>{obj[h]=row[i]??'';}); return obj;
  }).filter(r=>{ const d=String(r['Date']||'').trim().toLowerCase(); return d&&!d.includes('total')&&!d.includes('average')&&!d.includes('avg')&&d!=='date'; });
}
function computeKPIs(rows) {
  if (!rows.length) return {};
  const sum = k => rows.reduce((a,r)=>a+parseNum(r[k]),0);
  const avg = k => { const v=rows.map(r=>parseNum(r[k])).filter(v=>v>0); return v.length?v.reduce((a,b)=>a+b,0)/v.length:0; };
  return {
    days:rows.length, totalSpendsGST:sum('Spends + GST'), fbSpendsGST:sum('FB Spends with GST'), googleSpendsGST:sum('Google Spends with GST'),
    totalDBPurchase:sum('Total DB Purchase'), fbDBPurchase:sum('FB DB Purchase'), googleDBPurchase:sum('Google DB Purchase'),
    totalDBCPS:avg('Total DB CPS'), fbDBCPS:avg('FB DB CPS'), googleDBCPS:avg('Google DB CPS'),
    revenueTotal:sum('Revenue (FB + Unknown + Google)'), revenueFB:sum('Revenue (FB UTM)'), revenueGoogle:sum('Revenue (Google UTM)'), revenueUnknown:sum('Revenue (Unknown)'),
    roasGST:avg('ROAS (FB + Unknown + Google) WITH GST'), fbDBROAS:avg('FB DB ROAS'), googleDBROAS:avg('Google DB ROAS'),
    shopifyRevenue:sum('Shopify Daily Revenue (TOTAL)'), shopifyCount:sum('Shopify Daily Sales Count (TOTAL)'),
    fbCPM:avg('FB CPM'), fbCTR:avg('FB CTR'), googleCTR:avg('Google CTR'), fbCPC:avg('FB CPC'), googleCPC:avg('Google CPC'),
    fbClicks:sum('FB Clicks'), googleClicks:sum('Google Clicks'), totalClicks:sum('Total Clicks'),
  };
}
function computeAggregateKPIs(arr) {
  if (!arr.length) return {};
  const sum=k=>arr.reduce((a,k2)=>a+(k2[k]||0),0);
  const avg=k=>{const v=arr.map(k2=>k2[k]||0).filter(v=>v>0);return v.length?v.reduce((a,b)=>a+b,0)/v.length:0;};
  return { days:sum('days'), totalSpendsGST:sum('totalSpendsGST'), fbSpendsGST:sum('fbSpendsGST'), googleSpendsGST:sum('googleSpendsGST'),
    totalDBPurchase:sum('totalDBPurchase'), fbDBPurchase:sum('fbDBPurchase'), googleDBPurchase:sum('googleDBPurchase'),
    totalDBCPS:avg('totalDBCPS'), fbDBCPS:avg('fbDBCPS'), googleDBCPS:avg('googleDBCPS'),
    revenueTotal:sum('revenueTotal'), revenueFB:sum('revenueFB'), revenueGoogle:sum('revenueGoogle'), revenueUnknown:sum('revenueUnknown'),
    roasGST:avg('roasGST'), fbDBROAS:avg('fbDBROAS'), googleDBROAS:avg('googleDBROAS'),
    shopifyRevenue:sum('shopifyRevenue'), shopifyCount:sum('shopifyCount'),
    fbCPM:avg('fbCPM'), fbCTR:avg('fbCTR'), googleCTR:avg('googleCTR'), fbCPC:avg('fbCPC'), googleCPC:avg('googleCPC'),
    fbClicks:sum('fbClicks'), googleClicks:sum('googleClicks'), totalClicks:sum('totalClicks'),
  };
}
function buildChartData(rows) {
  return rows.map(r=>({
    date:String(r['Date']||'').trim(), totalSpends:parseNum(r['Spends + GST']),
    fbSpends:parseNum(r['FB Spends with GST']), googleSpends:parseNum(r['Google Spends with GST']),
    revenue:parseNum(r['Revenue (FB + Unknown + Google)']), shopifyRevenue:parseNum(r['Shopify Daily Revenue (TOTAL)']),
    fbDBPurchase:parseNum(r['FB DB Purchase']), googleDBPurchase:parseNum(r['Google DB Purchase']),
    roas:parseNum(r['ROAS (FB + Unknown + Google) WITH GST']), fbROAS:parseNum(r['FB DB ROAS']), googleROAS:parseNum(r['Google DB ROAS']),
    totalDBCPS:parseNum(r['Total DB CPS']), fbDBCPS:parseNum(r['FB DB CPS']), googleDBCPS:parseNum(r['Google DB CPS']),
    fbCTR:parseNum(r['FB CTR']), fbCPM:parseNum(r['FB CPM']),
  }));
}
function buildAiContext(sheet, kpis, chartData) {
  if (!kpis||!chartData.length) return '';
  const top=[...chartData].sort((a,b)=>b.revenue-a.revenue).slice(0,3).map(d=>`${d.date} (Rev ${fINR(d.revenue,true)}, ROAS ${d.roas.toFixed(2)}x)`).join('; ');
  return `neOwn Performance Marketing — ${sheet} (${kpis.days} days)
SPENDS: Total ${fINR(kpis.totalSpendsGST,true)} | FB ${fINR(kpis.fbSpendsGST,true)} (${pct(kpis.fbSpendsGST,kpis.totalSpendsGST)}%) | Google ${fINR(kpis.googleSpendsGST,true)}
PURCHASES: Total ${Math.round(kpis.totalDBPurchase)} | FB ${Math.round(kpis.fbDBPurchase)} | Google ${Math.round(kpis.googleDBPurchase)}
AVG CPS: Total ${fINR(kpis.totalDBCPS,true)} | FB ${fINR(kpis.fbDBCPS,true)} | Google ${fINR(kpis.googleDBCPS,true)}
REVENUE: Attribution ${fINR(kpis.revenueTotal,true)} | Shopify ${fINR(kpis.shopifyRevenue,true)} (${Math.round(kpis.shopifyCount)} orders)
ROAS: ${kpis.roasGST.toFixed(2)}x | FB ${kpis.fbDBROAS.toFixed(2)}x | Google ${kpis.googleDBROAS.toFixed(2)}x
FB: CPM ${fINR(kpis.fbCPM,true)} | CTR ${kpis.fbCTR.toFixed(2)}% | CPC ${fINR(kpis.fbCPC,true)}
BEST DAYS: ${top}`;
}

/* ══════════════════════════════════════════════════════════════
   REVENUE DATA PROCESSING
   Sheet structure: rows with daily data (col0=day#, col1=dayOfWeek,
   col2=Total, col3=SalesTeam, col4=ManualRenewal, col5=AutoRenewal, col6=SelfServe)
   Special rows: "Monthly Target" and "Monthly Achievement"
   ══════════════════════════════════════════════════════════════ */
function processRevenueSheet(rawValues) {
  if (!rawValues || rawValues.length < 2) return { daily: [], target: null, kpis: {} };
  let target = null, daily = [];
  for (const row of rawValues) {
    const c0 = String(row[0] || '').trim();
    const c0L = c0.toLowerCase();
    if (c0L.includes('monthly target')) {
      target = { total: parseNum(row[2]), sales: parseNum(row[3]), renewals: parseNum(row[4]), selfServe: parseNum(row[5]) };
    }
    // Daily rows: col0 is a pure integer 1-31 (not a range like "1-7")
    if (/^\d{1,2}$/.test(c0)) {
      const day = parseInt(c0);
      if (day >= 1 && day <= 31 && parseNum(row[2]) > 0) {
        daily.push({
          day, dayOfWeek: String(row[1]||'').trim(),
          total:         parseNum(row[2]),
          salesTeam:     parseNum(row[3]),
          manualRenewal: parseNum(row[4]),
          autoRenewal:   parseNum(row[5]),
          selfServe:     parseNum(row[6]),
        });
      }
    }
  }
  daily.sort((a,b) => a.day - b.day);
  const kpis = computeRevenueKPIs(daily, target);
  return { daily, target, kpis };
}

function computeRevenueKPIs(daily, target) {
  const sum = k => daily.reduce((a,d)=>a+(d[k]||0), 0);
  const totalRevenue   = sum('total');
  const salesTotal     = sum('salesTeam');
  const manualTotal    = sum('manualRenewal');
  const autoTotal      = sum('autoRenewal');
  const selfServeTotal = sum('selfServe');
  const renewalsTotal  = manualTotal + autoTotal;
  const days           = daily.length;
  const avgDaily       = days > 0 ? totalRevenue / days : 0;
  const bestDayVal     = daily.length > 0 ? Math.max(...daily.map(d=>d.total)) : 0;
  const targetTotal    = target?.total    || 0;
  const targetSales    = target?.sales    || 0;
  const targetRenewals = target?.renewals || 0;
  const targetSelfServe= target?.selfServe|| 0;
  const achievementPct = targetTotal > 0 ? (totalRevenue / targetTotal) * 100 : 0;
  const remaining      = Math.max(0, targetTotal - totalRevenue);
  const lastDay        = daily.length > 0 ? daily[daily.length-1].day : 0;
  const monthElapsed   = (lastDay / 31) * 100;
  const paceGap        = achievementPct - monthElapsed;
  return {
    days, totalRevenue, salesTotal, manualTotal, autoTotal, selfServeTotal, renewalsTotal,
    avgDaily, bestDayVal, targetTotal, targetSales, targetRenewals, targetSelfServe,
    achievementPct, remaining, lastDay, monthElapsed, paceGap,
    salesAchPct:    targetSales     > 0 ? (salesTotal     / targetSales)     * 100 : 0,
    renewalsAchPct: targetRenewals  > 0 ? (renewalsTotal  / targetRenewals)  * 100 : 0,
    selfServeAchPct:targetSelfServe > 0 ? (selfServeTotal / targetSelfServe) * 100 : 0,
  };
}

function computeAggregateRevKPIs(arr) {
  const sum = k => arr.reduce((a,kp)=>a+(kp[k]||0), 0);
  const totalRevenue = sum('totalRevenue'), targetTotal = sum('targetTotal');
  return {
    days: sum('days'), totalRevenue, salesTotal: sum('salesTotal'),
    manualTotal: sum('manualTotal'), autoTotal: sum('autoTotal'),
    selfServeTotal: sum('selfServeTotal'), renewalsTotal: sum('renewalsTotal'),
    avgDaily: arr.length ? totalRevenue / arr.length : 0,
    targetTotal, achievementPct: targetTotal > 0 ? (totalRevenue/targetTotal)*100 : 0,
    remaining: Math.max(0, targetTotal - totalRevenue),
    paceGap: 0, lastDay: 0, monthElapsed: 0,
    salesAchPct: 0, renewalsAchPct: 0, selfServeAchPct: 0,
  };
}

function buildRevenueAiContext(sheet, kpis, daily, target) {
  if (!kpis || !daily.length) return '';
  const best = daily.reduce((a,b)=>b.total>a.total?b:a, daily[0]);
  return `neOwn Revenue Achievement — ${sheet} (${daily.length} days of data)
MONTHLY TARGET: Total ${fINR(target?.total||0,true)} | Sales ${fINR(target?.sales||0,true)} | Renewals ${fINR(target?.renewals||0,true)} | Self-Serve ${fINR(target?.selfServe||0,true)}
ACHIEVEMENT: ${kpis.achievementPct.toFixed(1)}% — ${fINR(kpis.totalRevenue,true)} of ${fINR(kpis.targetTotal,true)} target
REMAINING: ${fINR(kpis.remaining,true)} | Pace vs Month Progress: ${kpis.paceGap>=0?'+':''}${kpis.paceGap.toFixed(1)}%
CHANNEL BREAKDOWN:
  Sales Team:  ${fINR(kpis.salesTotal,true)} (${kpis.salesAchPct.toFixed(1)}% of target)
  Manual Renewal: ${fINR(kpis.manualTotal,true)}
  Auto Renewal:   ${fINR(kpis.autoTotal,true)}
  Renewals Total: ${fINR(kpis.renewalsTotal,true)} (${kpis.renewalsAchPct.toFixed(1)}% of target)
  Self Serve:  ${fINR(kpis.selfServeTotal,true)} (${kpis.selfServeAchPct.toFixed(1)}% of target)
DAILY AVG: ${fINR(kpis.avgDaily,true)} | Best Day: Day ${best.day} (${best.dayOfWeek}) ${fINR(best.total,true)}`;
}

/* ══════════════════════════════════════════════════════════════
   KPI RENDERING — ADS
   ══════════════════════════════════════════════════════════════ */
function deltaHtml(curr, prev) {
  if (!prev) return '';
  const d = pctChange(curr, prev);
  const cls = d >= 0 ? 'kpi-delta-up' : 'kpi-delta-down';
  return `<span class="${cls}">${d>=0?'▲':'▼'} ${Math.abs(d).toFixed(1)}%</span>`;
}
function renderKPIs(kpis, prevKpis) {
  const avgOrder = kpis.shopifyCount > 0 ? kpis.shopifyRevenue / kpis.shopifyCount : 0;
  const d = prevKpis ? k => deltaHtml(kpis[k], prevKpis[k]) : () => '';
  const cards = [
    { label:'Total Spends (GST)',  value:fINR(kpis.totalSpendsGST,true),  sub:`FB ${fINR(kpis.fbSpendsGST,true)} · G ${fINR(kpis.googleSpendsGST,true)}`,    delta:d('totalSpendsGST'),  accent:true },
    { label:'Total DB Purchases',  value:fNum(kpis.totalDBPurchase),       sub:`FB ${fNum(kpis.fbDBPurchase)} · G ${fNum(kpis.googleDBPurchase)}`,              delta:d('totalDBPurchase') },
    { label:'Avg DB CPS',          value:fINR(kpis.totalDBCPS,true),       sub:`FB ${fINR(kpis.fbDBCPS,true)} · G ${fINR(kpis.googleDBCPS,true)}`,             delta:d('totalDBCPS') },
    { label:'Attribution Revenue', value:fINR(kpis.revenueTotal,true),     sub:`FB ${fINR(kpis.revenueFB,true)} · G ${fINR(kpis.revenueGoogle,true)}`,          delta:d('revenueTotal') },
    { label:'ROAS (with GST)',     value:kpis.roasGST.toFixed(2)+'x',      sub:`FB ${kpis.fbDBROAS.toFixed(2)}x · G ${kpis.googleDBROAS.toFixed(2)}x`,          delta:d('roasGST'),        accent:true },
    { label:'Shopify Revenue',     value:fINR(kpis.shopifyRevenue,true),   sub:`${fNum(kpis.shopifyCount)} orders · ₹${fNum(avgOrder)} AOV`,                    delta:d('shopifyRevenue') },
    { label:'FB CTR',              value:kpis.fbCTR.toFixed(2)+'%',        sub:`CPM ${fINR(kpis.fbCPM,true)} · CPC ${fINR(kpis.fbCPC,true)}`,                  delta:d('fbCTR') },
    { label:'Total Clicks',        value:fNum(kpis.totalClicks),           sub:`FB ${fNum(kpis.fbClicks)} · G ${fNum(kpis.googleClicks)}`,                       delta:d('totalClicks') },
  ];
  document.getElementById('kpiGrid').innerHTML = cards.map(c=>`
    <div class="kpi-card ${c.accent?'accent':''}">
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value-row"><div class="kpi-value">${c.value}</div>${c.delta?`<div class="kpi-delta">${c.delta}</div>`:''}</div>
      <div class="kpi-sub">${c.sub}</div>
    </div>`).join('');
}

/* ══════════════════════════════════════════════════════════════
   KPI RENDERING — REVENUE
   ══════════════════════════════════════════════════════════════ */
function renderRevenueKPIs(kpis, prevKpis) {
  const d = prevKpis ? k => deltaHtml(kpis[k], prevKpis[k]) : () => '';
  const paceLabel = kpis.paceGap >= 0
    ? `<span class="kpi-delta-up">▲ ${kpis.paceGap.toFixed(1)}% ahead of pace</span>`
    : `<span class="kpi-delta-down">▼ ${Math.abs(kpis.paceGap).toFixed(1)}% behind pace</span>`;
  const cards = [
    { label:'Monthly Achievement', value:kpis.achievementPct.toFixed(1)+'%', sub:`Day ${kpis.lastDay} · ${fINR(kpis.totalRevenue,true)} of ${fINR(kpis.targetTotal,true)}`, extraHtml:paceLabel, accent:true },
    { label:'Total Revenue',       value:fINR(kpis.totalRevenue,true),        sub:`Target ${fINR(kpis.targetTotal,true)} · Remaining ${fINR(kpis.remaining,true)}`,           delta:d('totalRevenue') },
    { label:'Sales Team',          value:fINR(kpis.salesTotal,true),          sub:`${kpis.salesAchPct.toFixed(1)}% of ₹${fINR(kpis.targetSales,true)} target`,               delta:d('salesTotal') },
    { label:'Manual Renewal',      value:fINR(kpis.manualTotal,true),         sub:`Part of ${fINR(kpis.renewalsTotal,true)} total renewals`,                                  delta:d('manualTotal') },
    { label:'Auto Renewal',        value:fINR(kpis.autoTotal,true),           sub:`${kpis.renewalsAchPct.toFixed(1)}% of ₹${fINR(kpis.targetRenewals,true)} renewals target`, delta:d('autoTotal'),   accent:true },
    { label:'Self Serve',          value:fINR(kpis.selfServeTotal,true),      sub:`${kpis.selfServeAchPct.toFixed(1)}% of ₹${fINR(kpis.targetSelfServe,true)} target`,        delta:d('selfServeTotal') },
    { label:'Daily Average',       value:fINR(kpis.avgDaily,true),            sub:`Best day ₹${fINR(kpis.bestDayVal,true)} · ${kpis.days} days recorded`,                     delta:d('avgDaily') },
    { label:'Renewals Total',      value:fINR(kpis.renewalsTotal,true),       sub:`Manual ${fINR(kpis.manualTotal,true)} + Auto ${fINR(kpis.autoTotal,true)}`,                delta:d('renewalsTotal') },
  ];
  document.getElementById('kpiGrid').innerHTML = cards.map(c=>`
    <div class="kpi-card ${c.accent?'accent':''}">
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value-row"><div class="kpi-value">${c.value}</div>${c.delta?`<div class="kpi-delta">${c.delta}</div>`:''}</div>
      <div class="kpi-sub">${c.sub}${c.extraHtml?' · '+c.extraHtml:''}</div>
    </div>`).join('');
}

function showKPISkeleton() {
  document.getElementById('kpiGrid').innerHTML = Array(8).fill('<div class="kpi-skeleton"></div>').join('');
}

/* ══════════════════════════════════════════════════════════════
   CHART HELPERS
   ══════════════════════════════════════════════════════════════ */
function makeChart(id, config) {
  if (chartInstances[id]) chartInstances[id].destroy();
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const wrap = canvas.closest('.chart-wrap');
  if (wrap) { const msg = wrap.querySelector('.chart-msg'); if (msg) msg.remove(); }
  chartInstances[id] = new Chart(canvas, config);
}
function setChartTitles(titles) {
  document.querySelectorAll('#chartsGrid .chart-label').forEach((el,i) => { if (titles[i]) el.textContent = titles[i]; });
}
function showChartMsg(msg) {
  ['cSpends','cRevenue','cPurchases','cROAS','cCPS','cCTR'].forEach(id => {
    if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const wrap = canvas.closest('.chart-wrap');
    if (!wrap) return;
    const ex = wrap.querySelector('.chart-msg'); if (ex) ex.remove();
    const el = document.createElement('div');
    el.className = 'chart-msg'; el.textContent = msg;
    wrap.appendChild(el);
  });
}
function dodBar(data, key, label) {
  return {
    label, data: data.map(d=>+d[key].toFixed(1)),
    backgroundColor: data.map(d=>d[key]>=0?'#10B98166':'#EF444466'),
    borderColor:     data.map(d=>d[key]>=0?'#10B981':'#EF4444'),
    borderWidth: 1.5, borderRadius: 3,
  };
}

/* ══════════════════════════════════════════════════════════════
   DASHBOARD BAR
   ══════════════════════════════════════════════════════════════ */
function renderDashBar() {
  const bar = document.getElementById('dashBar');
  if (!bar) return;
  bar.innerHTML = `
    <button class="dash-btn ${currentDashboard==='ads'?'active':''}" onclick="switchDashboard('ads')">📈 Performance Marketing</button>
    <button class="dash-btn ${currentDashboard==='revenue'?'active':''}" onclick="switchDashboard('revenue')">💰 Revenue Achievement</button>
  `;
}
async function switchDashboard(dash) {
  if (currentDashboard === dash) return;
  currentDashboard = dash;
  currentSheet = '';
  currentKPIs = null;
  currentAiCtx = '';
  currentChartData = [];
  currentRevData = null;
  chatHistory = [];
  renderDashBar();
  renderViewBar();
  document.getElementById('insightsBody').innerHTML = '<div class="insights-empty">Click "Generate Insights" to get AI-powered analysis of the selected month.</div>';
  document.getElementById('insightsSub').textContent = 'Claude-powered analysis';
  showKPISkeleton();
  showChartMsg('Loading…');
  await loadSheetList();
}

/* ══════════════════════════════════════════════════════════════
   VIEW BAR
   ══════════════════════════════════════════════════════════════ */
function renderViewBar() {
  const bar = document.getElementById('viewBar');
  if (!bar) return;
  bar.innerHTML = VIEWS.map(v=>`
    <button class="view-btn ${v.id===currentView?'active':''}" onclick="switchView('${v.id}')" title="${v.id}">${v.label}</button>
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
  if (currentDashboard === 'ads') {
    switch (currentView) {
      case 'perday':  renderPerDayView();       break;
      case 'monthly': await renderMonthlyView(); break;
      case 'mom':     await renderMoMView();     break;
      case 'dod':     renderDoDView();           break;
      case 'yoy':     await renderYoYView();     break;
    }
  } else {
    switch (currentView) {
      case 'perday':  renderRevPerDayView();       break;
      case 'monthly': await renderRevMonthlyView(); break;
      case 'mom':     await renderRevMoMView();     break;
      case 'dod':     renderRevDoDView();           break;
      case 'yoy':     await renderRevYoYView();     break;
    }
  }
}

/* ══════════════════════════════════════════════════════════════
   ADS — PER DAY VIEW
   ══════════════════════════════════════════════════════════════ */
function renderPerDayView() {
  if (!currentChartData.length) return;
  if (currentKPIs) renderKPIs(currentKPIs);
  setChartTitles(['Daily Spends incl. GST','Daily Revenue','Daily DB Purchases','ROAS Trend','CPS Trend','FB CTR % & CPM ₹']);
  const data = currentChartData;
  const labels = data.map(d=>shortDate(d.date));
  const fd = data.map(d=>String(d.date));
  const xA = xAxis(labels);
  makeChart('cSpends',{type:'line',data:{labels,datasets:[
    {label:'Total',data:data.map(d=>d.totalSpends),borderColor:'#374151',borderWidth:2,pointRadius:0,tension:.3},
    {label:'Facebook',data:data.map(d=>d.fbSpends),borderColor:PINK,borderWidth:1.5,pointRadius:0,tension:.3},
    {label:'Google',data:data.map(d=>d.googleSpends),borderColor:INDIGO,borderWidth:1.5,pointRadius:0,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cRevenue',{type:'line',data:{labels,datasets:[
    {label:'Shopify Revenue',data:data.map(d=>d.shopifyRevenue),borderColor:INDIGO,backgroundColor:'rgba(60,40,180,.08)',fill:true,borderWidth:2,pointRadius:0,tension:.3},
    {label:'Attribution Revenue',data:data.map(d=>d.revenue),borderColor:PINK,borderWidth:2,pointRadius:0,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cPurchases',{type:'bar',data:{labels,datasets:[
    {label:'Facebook',data:data.map(d=>d.fbDBPurchase),backgroundColor:PINK,borderRadius:3},
    {label:'Google',data:data.map(d=>d.googleDBPurchase),backgroundColor:INDIGO,borderRadius:3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttNum,fd)},scales:{x:xA,y:yAxis(yNum)}}});
  makeChart('cROAS',{type:'line',data:{labels,datasets:[
    {label:'Total ROAS',data:data.map(d=>+d.roas.toFixed(2)),borderColor:'#374151',borderWidth:2,pointRadius:0,tension:.3},
    {label:'FB ROAS',data:data.map(d=>+d.fbROAS.toFixed(2)),borderColor:PINK,borderWidth:1.5,pointRadius:0,tension:.3,borderDash:[5,3]},
    {label:'Google ROAS',data:data.map(d=>+d.googleROAS.toFixed(2)),borderColor:INDIGO,borderWidth:1.5,pointRadius:0,tension:.3,borderDash:[5,3]},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttX,fd)},scales:{x:xA,y:yAxis(yX)}}});
  makeChart('cCPS',{type:'line',data:{labels,datasets:[
    {label:'Total CPS',data:data.map(d=>Math.round(d.totalDBCPS)),borderColor:'#374151',borderWidth:2,pointRadius:0,tension:.3},
    {label:'FB CPS',data:data.map(d=>Math.round(d.fbDBCPS)),borderColor:PINK,borderWidth:1.5,pointRadius:0,tension:.3},
    {label:'Google CPS',data:data.map(d=>Math.round(d.googleDBCPS)),borderColor:INDIGO,borderWidth:1.5,pointRadius:0,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yRs)}}});
  makeChart('cCTR',{type:'bar',data:{labels,datasets:[
    {label:'FB CTR %',data:data.map(d=>+d.fbCTR.toFixed(2)),backgroundColor:PINK_L,borderRadius:3,yAxisID:'y'},
    {label:'FB CPM ₹',data:data.map(d=>Math.round(d.fbCPM)),type:'line',borderColor:PINK,borderWidth:2,pointRadius:0,tension:.3,yAxisID:'y2'},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:mixedTooltip([ttPct,ttINR],fd)},
    scales:{x:xA,y:{...yAxis(yPct),position:'left'},y2:{...yAxis(yRs),position:'right',grid:{drawOnChartArea:false}}}}});
}

/* ── ADS MONTHLY VIEW ── */
async function renderMonthlyView() {
  await ensureAllLoaded();
  const months = namesBySource.ads.map(n=>({name:n,kpis:cacheBySource.ads[n]?.kpis})).filter(m=>m.kpis);
  if (!months.length){showChartMsg('Loading monthly data…');return;}
  renderKPIs(computeAggregateKPIs(months.map(m=>m.kpis)));
  setChartTitles(['Monthly Total Spends','Monthly Revenue + ROAS','Monthly DB Purchases','Monthly Avg ROAS','Monthly Avg CPS','Monthly FB CTR & CPM']);
  const labels=months.map(m=>m.name), xA=xAxisFull();
  makeChart('cSpends',{type:'bar',data:{labels,datasets:[
    {label:'Facebook',data:months.map(m=>Math.round(m.kpis.fbSpendsGST)),backgroundColor:PINK+'CC',borderRadius:4},
    {label:'Google',data:months.map(m=>Math.round(m.kpis.googleSpendsGST)),backgroundColor:INDIGO+'CC',borderRadius:4},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cRevenue',{type:'bar',data:{labels,datasets:[
    {label:'Shopify Revenue',data:months.map(m=>Math.round(m.kpis.shopifyRevenue)),backgroundColor:IND_L+'CC',borderRadius:4},
    {label:'Attribution Revenue',data:months.map(m=>Math.round(m.kpis.revenueTotal)),backgroundColor:PINK+'CC',borderRadius:4},
    {label:'ROAS',data:months.map(m=>+m.kpis.roasGST.toFixed(2)),type:'line',borderColor:INDIGO,backgroundColor:'transparent',borderWidth:2.5,pointRadius:5,pointBackgroundColor:INDIGO,tension:.3,yAxisID:'y2'},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:mixedTooltip([ttINR,ttINR,ttX])},
    scales:{x:xA,y:{...yAxis(yINR),position:'left'},y2:{...yAxis(yX),position:'right',grid:{drawOnChartArea:false}}}}});
  makeChart('cPurchases',{type:'bar',data:{labels,datasets:[
    {label:'Facebook',data:months.map(m=>Math.round(m.kpis.fbDBPurchase)),backgroundColor:PINK+'CC',borderRadius:4},
    {label:'Google',data:months.map(m=>Math.round(m.kpis.googleDBPurchase)),backgroundColor:INDIGO+'CC',borderRadius:4},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttNum)},scales:{x:xA,y:yAxis(yNum)}}});
  makeChart('cROAS',{type:'line',data:{labels,datasets:[
    {label:'Total ROAS',data:months.map(m=>+m.kpis.roasGST.toFixed(2)),borderColor:'#374151',backgroundColor:'rgba(55,65,81,.06)',fill:true,borderWidth:2.5,pointRadius:5,pointBackgroundColor:'#374151',tension:.3},
    {label:'FB ROAS',data:months.map(m=>+m.kpis.fbDBROAS.toFixed(2)),borderColor:PINK,borderWidth:1.5,pointRadius:3,tension:.3},
    {label:'Google ROAS',data:months.map(m=>+m.kpis.googleDBROAS.toFixed(2)),borderColor:INDIGO,borderWidth:1.5,pointRadius:3,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttX)},scales:{x:xA,y:yAxis(yX)}}});
  makeChart('cCPS',{type:'line',data:{labels,datasets:[
    {label:'Total CPS',data:months.map(m=>Math.round(m.kpis.totalDBCPS)),borderColor:'#374151',backgroundColor:'rgba(55,65,81,.06)',fill:true,borderWidth:2.5,pointRadius:5,pointBackgroundColor:'#374151',tension:.3},
    {label:'FB CPS',data:months.map(m=>Math.round(m.kpis.fbDBCPS)),borderColor:PINK,borderWidth:1.5,pointRadius:3,tension:.3},
    {label:'Google CPS',data:months.map(m=>Math.round(m.kpis.googleDBCPS)),borderColor:INDIGO,borderWidth:1.5,pointRadius:3,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR)},scales:{x:xA,y:yAxis(yRs)}}});
  makeChart('cCTR',{type:'bar',data:{labels,datasets:[
    {label:'FB CTR %',data:months.map(m=>+m.kpis.fbCTR.toFixed(2)),backgroundColor:PINK_L+'CC',borderRadius:4,yAxisID:'y'},
    {label:'FB CPM ₹',data:months.map(m=>Math.round(m.kpis.fbCPM)),type:'line',borderColor:PINK,borderWidth:2.5,pointRadius:4,pointBackgroundColor:PINK,tension:.3,yAxisID:'y2'},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:mixedTooltip([ttPct,ttINR])},
    scales:{x:xA,y:{...yAxis(yPct),position:'left'},y2:{...yAxis(yRs),position:'right',grid:{drawOnChartArea:false}}}}});
}

/* ── ADS MoM VIEW ── */
async function renderMoMView() {
  const names = namesBySource.ads, ci = names.indexOf(currentSheet);
  if (ci <= 0){showChartMsg('No previous month available.');if(currentKPIs)renderKPIs(currentKPIs);return;}
  const prevName = names[ci-1];
  await ensureSheetLoaded(prevName,'ads');
  const curr=cacheBySource.ads[currentSheet], prev=cacheBySource.ads[prevName];
  if (!curr||!prev){showChartMsg('Loading data…');return;}
  renderKPIs(curr.kpis, prev.kpis);
  setChartTitles([`Spends: ${currentSheet} vs ${prevName}`,`Revenue: ${currentSheet} vs ${prevName}`,`Purchases: ${currentSheet} vs ${prevName}`,`ROAS: ${currentSheet} vs ${prevName}`,`CPS: ${currentSheet} vs ${prevName}`,`CTR: ${currentSheet} vs ${prevName}`]);
  const cD=curr.chartData, pD=prev.chartData, n=Math.max(cD.length,pD.length);
  const labels=Array.from({length:n},(_,i)=>`Day ${i+1}`);
  const fd=Array.from({length:n},(_,i)=>`Day ${i+1}  ·  ${cD[i]?cD[i].date:'—'}  /  ${pD[i]?pD[i].date:'—'}`);
  const xA=xAxis(labels);
  function ml(data,key,label,color,dash=false){return{label,data:data.map(d=>d?d[key]:null),borderColor:color,borderWidth:dash?2:2.5,pointRadius:0,tension:.3,spanGaps:true,...(dash?{borderDash:[5,4]}:{})};}
  makeChart('cSpends',{type:'line',data:{labels,datasets:[ml(cD,'totalSpends',currentSheet,PINK),ml(pD,'totalSpends',prevName,GRAY,true)]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cRevenue',{type:'line',data:{labels,datasets:[ml(cD,'shopifyRevenue',currentSheet,INDIGO),ml(pD,'shopifyRevenue',prevName,IND_L,true)]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cPurchases',{type:'line',data:{labels,datasets:[
    {label:`${currentSheet} FB`,data:cD.map(d=>d.fbDBPurchase),borderColor:PINK,borderWidth:2,pointRadius:0,tension:.3},
    {label:`${currentSheet} G`,data:cD.map(d=>d.googleDBPurchase),borderColor:INDIGO,borderWidth:2,pointRadius:0,tension:.3},
    {label:`${prevName} FB`,data:pD.map(d=>d.fbDBPurchase),borderColor:PINK_L,borderWidth:1.5,pointRadius:0,tension:.3,borderDash:[5,4]},
    {label:`${prevName} G`,data:pD.map(d=>d.googleDBPurchase),borderColor:IND_L,borderWidth:1.5,pointRadius:0,tension:.3,borderDash:[5,4]},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttNum,fd)},scales:{x:xA,y:yAxis(yNum)}}});
  makeChart('cROAS',{type:'line',data:{labels,datasets:[
    ml(cD.map(d=>({roas:+d.roas.toFixed(2)})),'roas',currentSheet,'#374151'),
    ml(pD.map(d=>({roas:+d.roas.toFixed(2)})),'roas',prevName,GRAY,true),
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttX,fd)},scales:{x:xA,y:yAxis(yX)}}});
  makeChart('cCPS',{type:'line',data:{labels,datasets:[
    ml(cD.map(d=>({cps:Math.round(d.totalDBCPS)})),'cps',currentSheet,'#374151'),
    ml(pD.map(d=>({cps:Math.round(d.totalDBCPS)})),'cps',prevName,GRAY,true),
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yRs)}}});
  makeChart('cCTR',{type:'line',data:{labels,datasets:[
    ml(cD.map(d=>({ctr:+d.fbCTR.toFixed(2)})),'ctr',currentSheet,PINK),
    ml(pD.map(d=>({ctr:+d.fbCTR.toFixed(2)})),'ctr',prevName,PINK_L,true),
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttPct,fd)},scales:{x:xA,y:yAxis(yPct)}}});
}

/* ── ADS DoD VIEW ── */
function renderDoDView() {
  if (!currentChartData.length){return;}
  if (currentKPIs) renderKPIs(currentKPIs);
  const data=currentChartData;
  if (data.length<2){showChartMsg('Need at least 2 days for DoD view.');return;}
  const dod=data.slice(1).map((d,i)=>{const p=data[i];return{date:d.date,fbSpends:pctChange(d.fbSpends,p.fbSpends),googleSpends:pctChange(d.googleSpends,p.googleSpends),revenue:pctChange(d.shopifyRevenue,p.shopifyRevenue),attrRevenue:pctChange(d.revenue,p.revenue),fbPurchases:pctChange(d.fbDBPurchase,p.fbDBPurchase),googlePurchases:pctChange(d.googleDBPurchase,p.googleDBPurchase),roas:pctChange(d.roas,p.roas),fbROAS:pctChange(d.fbROAS,p.fbROAS),cps:pctChange(d.totalDBCPS,p.totalDBCPS),fbCPS:pctChange(d.fbDBCPS,p.fbDBCPS),ctr:pctChange(d.fbCTR,p.fbCTR),cpm:pctChange(d.fbCPM,p.fbCPM)};});
  const labels=dod.map(d=>shortDate(d.date)), fd=dod.map(d=>String(d.date)), xA=xAxis(labels);
  const opts=()=>({responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttDPct,fd)},scales:{x:xA,y:yAxis(yDPct)}});
  setChartTitles(['Spends DoD %','Revenue DoD %','Purchases DoD %','ROAS DoD %','CPS DoD %','CTR & CPM DoD %']);
  makeChart('cSpends',{type:'bar',data:{labels,datasets:[dodBar(dod,'fbSpends','FB Spends'),dodBar(dod,'googleSpends','Google Spends')]},options:opts()});
  makeChart('cRevenue',{type:'bar',data:{labels,datasets:[dodBar(dod,'revenue','Shopify Revenue'),dodBar(dod,'attrRevenue','Attribution Revenue')]},options:opts()});
  makeChart('cPurchases',{type:'bar',data:{labels,datasets:[dodBar(dod,'fbPurchases','FB Purchases'),dodBar(dod,'googlePurchases','Google Purchases')]},options:opts()});
  makeChart('cROAS',{type:'bar',data:{labels,datasets:[dodBar(dod,'roas','Total ROAS'),dodBar(dod,'fbROAS','FB ROAS')]},options:opts()});
  makeChart('cCPS',{type:'bar',data:{labels,datasets:[dodBar(dod,'cps','Total CPS'),dodBar(dod,'fbCPS','FB CPS')]},options:opts()});
  makeChart('cCTR',{type:'bar',data:{labels,datasets:[dodBar(dod,'ctr','FB CTR'),dodBar(dod,'cpm','FB CPM')]},options:opts()});
}

/* ── ADS YoY VIEW ── */
async function renderYoYView() {
  await ensureAllLoaded();
  const parsed=namesBySource.ads.map(n=>({name:n,p:parseSheetDate(n)})).filter(s=>s.p&&cacheBySource.ads[s.name]?.kpis);
  if (!parsed.length){await renderMonthlyView();return;}
  const yg={};
  parsed.forEach(({name,p})=>{const yr=String(p.year);if(!yg[yr])yg[yr]={};yg[yr][p.month]=cacheBySource.ads[name].kpis;});
  const MNAMES=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const years=Object.keys(yg).sort(), allMonths=[...new Set(parsed.map(s=>s.p.month))].sort((a,b)=>a-b);
  const labels=allMonths.map(m=>MNAMES[m]), YCOLS=[INDIGO,PINK,GREEN,AMBER,'#8B5CF6',TEAL];
  renderKPIs(computeAggregateKPIs(Object.values(yg[years[years.length-1]]).filter(Boolean)));
  setChartTitles(['Spends by Year','Revenue by Year','Purchases by Year','ROAS by Year','CPS by Year','CTR by Year']);
  const xA=xAxisFull();
  function yoyDs(fn){return years.map((yr,i)=>({label:yr,data:allMonths.map(m=>{const k=yg[yr][m];return k?fn(k):null;}),borderColor:YCOLS[i%YCOLS.length],backgroundColor:YCOLS[i%YCOLS.length]+'22',borderWidth:2.5,pointRadius:5,pointBackgroundColor:YCOLS[i%YCOLS.length],tension:.3,spanGaps:true}));}
  const lo=(tt,yf)=>({responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(tt)},scales:{x:xA,y:yAxis(yf)}});
  makeChart('cSpends',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.totalSpendsGST))},options:lo(ttINR,yINR)});
  makeChart('cRevenue',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.revenueTotal))},options:lo(ttINR,yINR)});
  makeChart('cPurchases',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.totalDBPurchase))},options:lo(ttNum,yNum)});
  makeChart('cROAS',{type:'line',data:{labels,datasets:yoyDs(k=>+k.roasGST.toFixed(2))},options:lo(ttX,yX)});
  makeChart('cCPS',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.totalDBCPS))},options:lo(ttINR,yRs)});
  makeChart('cCTR',{type:'line',data:{labels,datasets:yoyDs(k=>+k.fbCTR.toFixed(2))},options:lo(ttPct,yPct)});
}

/* ══════════════════════════════════════════════════════════════
   REVENUE — PER DAY VIEW
   ══════════════════════════════════════════════════════════════ */
function renderRevPerDayView() {
  if (!currentRevData?.daily.length) return;
  renderRevenueKPIs(currentRevData.kpis);
  setChartTitles(['Revenue by Channel','Cumulative vs Target Pace','Daily Revenue Trend','Sales Team','Renewals (Manual + Auto)','Self Serve']);
  const { daily, target } = currentRevData;
  const labels = daily.map(d=>`${d.day} ${d.dayOfWeek.slice(0,3)}`);
  const fd = daily.map(d=>`Day ${d.day} — ${d.dayOfWeek}`);
  const xA = xAxis(labels);
  // 1. Stacked channel breakdown
  makeChart('cSpends',{type:'bar',data:{labels,datasets:[
    {label:'Sales Team',     data:daily.map(d=>d.salesTeam),     backgroundColor:PINK+'CC',  stack:'rev',borderRadius:0},
    {label:'Manual Renewal', data:daily.map(d=>d.manualRenewal), backgroundColor:INDIGO+'CC',stack:'rev'},
    {label:'Auto Renewal',   data:daily.map(d=>d.autoRenewal),   backgroundColor:IND_L+'CC', stack:'rev'},
    {label:'Self Serve',     data:daily.map(d=>d.selfServe),     backgroundColor:AMBER+'CC', stack:'rev'},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:{...yAxis(yINR),stacked:true}}}});
  // 2. Cumulative vs target pace
  let cum=0;
  const cumulativeData=daily.map(d=>{cum+=d.total;return cum;});
  const targetPace=target?daily.map(d=>Math.round(d.day*(target.total/31))):[];
  const cumDs=[{label:'Cumulative Revenue',data:cumulativeData,borderColor:INDIGO,backgroundColor:'rgba(60,40,180,.08)',fill:true,borderWidth:2.5,pointRadius:0,tension:.3}];
  if(targetPace.length) cumDs.push({label:'Target Pace',data:targetPace,borderColor:GRAY,borderWidth:1.5,pointRadius:0,tension:.3,borderDash:[5,3]});
  makeChart('cRevenue',{type:'line',data:{labels,datasets:cumDs},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  // 3. Daily total + 7-day MA
  const ma7=daily.map((_,i)=>{const sl=daily.slice(Math.max(0,i-6),i+1);return Math.round(sl.reduce((a,d)=>a+d.total,0)/sl.length);});
  makeChart('cPurchases',{type:'bar',data:{labels,datasets:[
    {label:'Daily Revenue',data:daily.map(d=>d.total),backgroundColor:INDIGO+'55',borderRadius:3},
    {label:'7-Day MA',data:ma7,type:'line',borderColor:PINK,borderWidth:2,pointRadius:0,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  // 4. Sales Team
  makeChart('cROAS',{type:'line',data:{labels,datasets:[{label:'Sales Team',data:daily.map(d=>d.salesTeam),borderColor:PINK,backgroundColor:'rgba(240,0,140,.06)',fill:true,borderWidth:2,pointRadius:0,tension:.3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  // 5. Renewals
  makeChart('cCPS',{type:'line',data:{labels,datasets:[
    {label:'Manual Renewal',data:daily.map(d=>d.manualRenewal),borderColor:INDIGO,borderWidth:2,pointRadius:0,tension:.3},
    {label:'Auto Renewal',data:daily.map(d=>d.autoRenewal),borderColor:TEAL,borderWidth:2,pointRadius:0,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  // 6. Self Serve
  makeChart('cCTR',{type:'line',data:{labels,datasets:[{label:'Self Serve',data:daily.map(d=>d.selfServe),borderColor:AMBER,backgroundColor:'rgba(245,158,11,.06)',fill:true,borderWidth:2,pointRadius:0,tension:.3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
}

/* ── REVENUE MONTHLY VIEW ── */
async function renderRevMonthlyView() {
  await ensureAllLoaded();
  const months=namesBySource.revenue.map(n=>({name:n,data:cacheBySource.revenue[n]})).filter(m=>m.data?.kpis);
  if (!months.length){showChartMsg('Loading monthly data…');return;}
  renderRevenueKPIs(computeAggregateRevKPIs(months.map(m=>m.data.kpis)));
  setChartTitles(['Monthly Revenue vs Target','Monthly Channel Breakdown','Monthly Achievement %','Monthly Sales Team','Monthly Renewals','Monthly Self Serve']);
  const labels=months.map(m=>m.name), xA=xAxisFull();
  makeChart('cSpends',{type:'bar',data:{labels,datasets:[
    {label:'Actual Revenue',data:months.map(m=>Math.round(m.data.kpis.totalRevenue)),backgroundColor:INDIGO+'CC',borderRadius:4},
    {label:'Target',data:months.map(m=>Math.round(m.data.kpis.targetTotal||0)),backgroundColor:'rgba(156,163,175,.3)',borderRadius:4},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cRevenue',{type:'bar',data:{labels,datasets:[
    {label:'Sales Team',data:months.map(m=>Math.round(m.data.kpis.salesTotal)),backgroundColor:PINK+'CC',stack:'ch',borderRadius:0},
    {label:'Manual Renewal',data:months.map(m=>Math.round(m.data.kpis.manualTotal)),backgroundColor:INDIGO+'CC',stack:'ch'},
    {label:'Auto Renewal',data:months.map(m=>Math.round(m.data.kpis.autoTotal)),backgroundColor:IND_L+'CC',stack:'ch'},
    {label:'Self Serve',data:months.map(m=>Math.round(m.data.kpis.selfServeTotal)),backgroundColor:AMBER+'CC',stack:'ch'},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR)},scales:{x:xA,y:{...yAxis(yINR),stacked:true}}}});
  makeChart('cPurchases',{type:'line',data:{labels,datasets:[
    {label:'Achievement %',data:months.map(m=>+m.data.kpis.achievementPct.toFixed(1)),borderColor:INDIGO,backgroundColor:'rgba(60,40,180,.06)',fill:true,borderWidth:2.5,pointRadius:5,pointBackgroundColor:INDIGO,tension:.3},
    {label:'Target 100%',data:months.map(()=>100),borderColor:GRAY,borderWidth:1.5,borderDash:[5,3],pointRadius:0},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttPct)},scales:{x:xA,y:yAxis(yPct)}}});
  makeChart('cROAS',{type:'line',data:{labels,datasets:[{label:'Sales Team',data:months.map(m=>Math.round(m.data.kpis.salesTotal)),borderColor:PINK,backgroundColor:'rgba(240,0,140,.06)',fill:true,borderWidth:2.5,pointRadius:5,pointBackgroundColor:PINK,tension:.3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cCPS',{type:'line',data:{labels,datasets:[
    {label:'Manual Renewal',data:months.map(m=>Math.round(m.data.kpis.manualTotal)),borderColor:INDIGO,borderWidth:2,pointRadius:3,tension:.3},
    {label:'Auto Renewal',data:months.map(m=>Math.round(m.data.kpis.autoTotal)),borderColor:TEAL,borderWidth:2,pointRadius:3,tension:.3},
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cCTR',{type:'line',data:{labels,datasets:[{label:'Self Serve',data:months.map(m=>Math.round(m.data.kpis.selfServeTotal)),borderColor:AMBER,backgroundColor:'rgba(245,158,11,.06)',fill:true,borderWidth:2.5,pointRadius:5,pointBackgroundColor:AMBER,tension:.3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR)},scales:{x:xA,y:yAxis(yINR)}}});
}

/* ── REVENUE MoM VIEW ── */
async function renderRevMoMView() {
  const names=namesBySource.revenue, ci=names.indexOf(currentSheet);
  if (ci<=0){showChartMsg('No previous month available.');if(currentRevData)renderRevenueKPIs(currentRevData.kpis);return;}
  const prevName=names[ci-1];
  await ensureSheetLoaded(prevName,'revenue');
  const curr=cacheBySource.revenue[currentSheet], prev=cacheBySource.revenue[prevName];
  if (!curr||!prev){showChartMsg('Loading data…');return;}
  renderRevenueKPIs(curr.kpis, prev.kpis);
  setChartTitles([`Total: ${currentSheet} vs ${prevName}`,`Channels: ${currentSheet} vs ${prevName}`,`Sales Team: ${currentSheet} vs ${prevName}`,`Manual Renewal: ${currentSheet} vs ${prevName}`,`Auto Renewal: ${currentSheet} vs ${prevName}`,`Self Serve: ${currentSheet} vs ${prevName}`]);
  const cD=curr.daily, pD=prev.daily, n=Math.max(cD.length,pD.length);
  const labels=Array.from({length:n},(_,i)=>`Day ${i+1}`);
  const fd=Array.from({length:n},(_,i)=>`Day ${i+1}  ·  ${cD[i]?`${cD[i].day} ${cD[i].dayOfWeek.slice(0,3)}`:'—'}  /  ${pD[i]?`${pD[i].day} ${pD[i].dayOfWeek.slice(0,3)}`:'—'}`);
  const xA=xAxis(labels);
  function rl(data,key,label,color,dash=false){return{label,data:data.map(d=>d?d[key]:null),borderColor:color,borderWidth:dash?2:2.5,pointRadius:0,tension:.3,spanGaps:true,...(dash?{borderDash:[5,4]}:{})};}
  makeChart('cSpends',{type:'line',data:{labels,datasets:[rl(cD,'total',currentSheet,INDIGO),rl(pD,'total',prevName,GRAY,true)]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cRevenue',{type:'line',data:{labels,datasets:[
    rl(cD,'salesTeam',`${currentSheet} Sales`,PINK), rl(cD,'selfServe',`${currentSheet} SS`,AMBER),
    rl(pD,'salesTeam',`${prevName} Sales`,PINK_L,true), rl(pD,'selfServe',`${prevName} SS`,GRAY,true),
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cPurchases',{type:'line',data:{labels,datasets:[rl(cD,'salesTeam',currentSheet,PINK),rl(pD,'salesTeam',prevName,PINK_L,true)]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cROAS',{type:'line',data:{labels,datasets:[rl(cD,'manualRenewal',currentSheet,INDIGO),rl(pD,'manualRenewal',prevName,IND_L,true)]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cCPS',{type:'line',data:{labels,datasets:[rl(cD,'autoRenewal',currentSheet,TEAL),rl(pD,'autoRenewal',prevName,GRAY,true)]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
  makeChart('cCTR',{type:'line',data:{labels,datasets:[rl(cD,'selfServe',currentSheet,AMBER),rl(pD,'selfServe',prevName,GRAY,true)]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttINR,fd)},scales:{x:xA,y:yAxis(yINR)}}});
}

/* ── REVENUE DoD VIEW ── */
function renderRevDoDView() {
  if (!currentRevData?.daily.length) return;
  renderRevenueKPIs(currentRevData.kpis);
  const daily=currentRevData.daily;
  if (daily.length<2){showChartMsg('Need at least 2 days for DoD view.');return;}
  const dod=daily.slice(1).map((d,i)=>{const p=daily[i];return{day:d.day,dayOfWeek:d.dayOfWeek,total:pctChange(d.total,p.total),salesTeam:pctChange(d.salesTeam,p.salesTeam),manualRenewal:pctChange(d.manualRenewal,p.manualRenewal),autoRenewal:pctChange(d.autoRenewal,p.autoRenewal),selfServe:pctChange(d.selfServe,p.selfServe),renewals:pctChange(d.manualRenewal+d.autoRenewal,p.manualRenewal+p.autoRenewal)};});
  const labels=dod.map(d=>`${d.day} ${d.dayOfWeek.slice(0,3)}`), fd=dod.map(d=>`Day ${d.day} — ${d.dayOfWeek}`), xA=xAxis(labels);
  const opts=()=>({responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(ttDPct,fd)},scales:{x:xA,y:yAxis(yDPct)}});
  setChartTitles(['Total Revenue DoD %','Channel Mix DoD %','Sales Team DoD %','Manual Renewal DoD %','Auto Renewal DoD %','Self Serve DoD %']);
  makeChart('cSpends',{type:'bar',data:{labels,datasets:[dodBar(dod,'total','Total Revenue')]},options:opts()});
  makeChart('cRevenue',{type:'bar',data:{labels,datasets:[dodBar(dod,'salesTeam','Sales Team'),dodBar(dod,'renewals','Renewals')]},options:opts()});
  makeChart('cPurchases',{type:'bar',data:{labels,datasets:[dodBar(dod,'salesTeam','Sales Team')]},options:opts()});
  makeChart('cROAS',{type:'bar',data:{labels,datasets:[dodBar(dod,'manualRenewal','Manual Renewal')]},options:opts()});
  makeChart('cCPS',{type:'bar',data:{labels,datasets:[dodBar(dod,'autoRenewal','Auto Renewal')]},options:opts()});
  makeChart('cCTR',{type:'bar',data:{labels,datasets:[dodBar(dod,'selfServe','Self Serve')]},options:opts()});
}

/* ── REVENUE YoY VIEW ── */
async function renderRevYoYView() {
  await ensureAllLoaded();
  const parsed=namesBySource.revenue.map(n=>({name:n,p:parseSheetDate(n)})).filter(s=>s.p&&cacheBySource.revenue[s.name]?.kpis);
  if (!parsed.length){await renderRevMonthlyView();return;}
  const yg={};
  parsed.forEach(({name,p})=>{const yr=String(p.year);if(!yg[yr])yg[yr]={};yg[yr][p.month]=cacheBySource.revenue[name].kpis;});
  const MNAMES=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const years=Object.keys(yg).sort(), allMonths=[...new Set(parsed.map(s=>s.p.month))].sort((a,b)=>a-b);
  const labels=allMonths.map(m=>MNAMES[m]), YCOLS=[INDIGO,PINK,GREEN,AMBER,'#8B5CF6',TEAL];
  renderRevenueKPIs(computeAggregateRevKPIs(Object.values(yg[years[years.length-1]]).filter(Boolean)));
  setChartTitles(['Total Revenue by Year','Sales Team by Year','Manual Renewal by Year','Auto Renewal by Year','Self Serve by Year','Achievement % by Year']);
  const xA=xAxisFull();
  function yoyDs(fn){return years.map((yr,i)=>({label:yr,data:allMonths.map(m=>{const k=yg[yr][m];return k?fn(k):null;}),borderColor:YCOLS[i%YCOLS.length],backgroundColor:YCOLS[i%YCOLS.length]+'22',borderWidth:2.5,pointRadius:5,pointBackgroundColor:YCOLS[i%YCOLS.length],tension:.3,spanGaps:true}));}
  const lo=(tt,yf)=>({responsive:true,maintainAspectRatio:false,plugins:{legend:baseLegend(),tooltip:makeTooltip(tt)},scales:{x:xA,y:yAxis(yf)}});
  makeChart('cSpends',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.totalRevenue))},options:lo(ttINR,yINR)});
  makeChart('cRevenue',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.salesTotal))},options:lo(ttINR,yINR)});
  makeChart('cPurchases',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.manualTotal))},options:lo(ttINR,yINR)});
  makeChart('cROAS',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.autoTotal))},options:lo(ttINR,yINR)});
  makeChart('cCPS',{type:'line',data:{labels,datasets:yoyDs(k=>Math.round(k.selfServeTotal))},options:lo(ttINR,yINR)});
  makeChart('cCTR',{type:'line',data:{labels,datasets:yoyDs(k=>+k.achievementPct.toFixed(1))},options:lo(ttPct,yPct)});
}

/* ══════════════════════════════════════════════════════════════
   DATA LOADING
   ══════════════════════════════════════════════════════════════ */
async function ensureSheetLoaded(name, source) {
  const src = source || currentDashboard;
  if (cacheBySource[src][name]) return;
  try {
    const res  = await fetch(`/api/sheets?sheet=${encodeURIComponent(name)}&source=${src}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    if (src === 'revenue') {
      cacheBySource.revenue[name] = processRevenueSheet(data.values || []);
    } else {
      const rows = processRows(data.values || []);
      cacheBySource.ads[name] = { rows, kpis: computeKPIs(rows), chartData: buildChartData(rows) };
    }
  } catch (err) { console.warn('Could not load sheet:', name, err.message); }
}

async function ensureAllLoaded() {
  const src = currentDashboard;
  const toLoad = namesBySource[src].filter(n => !cacheBySource[src][n]);
  if (!toLoad.length) return;
  setStatus('loading', 'Loading all months…');
  await Promise.allSettled(toLoad.map(n => ensureSheetLoaded(n, src)));
  setStatus('ok', 'All months loaded');
}

async function loadSheet(sheetName) {
  currentSheet = sheetName;
  setStatus('loading', 'Loading…');
  showError('');
  showKPISkeleton();
  document.getElementById('insightsBody').innerHTML = '<div class="insights-empty">Click "Generate Insights" for AI analysis.</div>';
  document.getElementById('insightsSub').textContent = 'Claude-powered analysis · ' + sheetName;
  try {
    await ensureSheetLoaded(sheetName, currentDashboard);
    if (currentDashboard === 'revenue') {
      const cached = cacheBySource.revenue[sheetName];
      if (!cached) throw new Error('Failed to load sheet');
      currentRevData = cached;
      currentKPIs    = cached.kpis;
      currentAiCtx   = buildRevenueAiContext(sheetName, cached.kpis, cached.daily, cached.target);
      setStatus('ok', cached.daily.length + ' days loaded');
    } else {
      const cached = cacheBySource.ads[sheetName];
      if (!cached) throw new Error('Failed to load sheet');
      currentKPIs      = cached.kpis;
      currentChartData = cached.chartData;
      currentAiCtx     = buildAiContext(sheetName, currentKPIs, currentChartData);
      setStatus('ok', cached.rows.length + ' days loaded');
    }
    await renderView();
    setTimeout(() => ensureAllLoaded(), 600);
  } catch (err) {
    showError('Error loading ' + sheetName + ': ' + err.message);
    setStatus('error', 'Error');
  }
}

async function loadSheetList() {
  setStatus('loading', 'Connecting…');
  try {
    const res  = await fetch(`/api/sheets?action=list&source=${currentDashboard}`);
    const data = await res.json();
    if (data.error) {
      showError(data.error);
      document.getElementById('setupHint').hidden = false;
      document.getElementById('monthTabs').innerHTML = '';
      setStatus('error', 'Config error');
      return;
    }
    namesBySource[currentDashboard] = data.sheets || [];
    const names = namesBySource[currentDashboard];
    if (!names.length) { showError('No sheets found.'); setStatus('error', 'No sheets'); return; }
    const tabsEl = document.getElementById('monthTabs');
    tabsEl.innerHTML = names.map(s=>`<button class="tab" onclick="switchTab(this,'${s}')">${s}</button>`).join('');
    const lastTab = tabsEl.lastElementChild;
    lastTab.classList.add('active');
    loadSheet(lastTab.textContent);
  } catch (err) {
    showError('Could not reach /api/sheets — check environment variables in Vercel.');
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
  const btn = document.getElementById('insightsBtn'), body = document.getElementById('insightsBody');
  btn.disabled = true; btn.textContent = 'Analyzing…';
  body.innerHTML = `<div class="insights-skeleton">${Array(5).fill('<div style="width:75%"></div>').join('')}</div>`;
  try {
    const res = await fetch('/api/insights', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ summary: currentAiCtx, sheetName: currentSheet, type: currentDashboard }),
    });
    const ins = await res.json();
    if (ins.error) throw new Error(ins.error);
    renderInsights(ins);
  } catch (err) {
    body.innerHTML = `<div class="insights-empty" style="color:#B91C1C">Error: ${err.message}</div>`;
  } finally { btn.disabled = false; btn.textContent = '↺ Refresh'; }
}

function renderInsights(ins) {
  const li = arr => (arr||[]).map(t=>`<li><span class="insights-dot">›</span>${t}</li>`).join('');
  const extra = (ins.aiCharts||[]).map(c=>`<div class="insights-extra-item"><span style="font-size:20px">📈</span><div><div class="insights-extra-title">${c.title}</div><div class="insights-extra-body">${c.insight}</div></div></div>`).join('');
  document.getElementById('insightsBody').innerHTML = `
    ${ins.summary?`<div class="insights-summary">${ins.summary}</div>`:''}
    <div class="insights-cols">
      <div class="insights-col green"><div class="insights-col-title">🔥 Highlights</div><ul>${li(ins.highlights)}</ul></div>
      <div class="insights-col orange"><div class="insights-col-title">⚠️ Watch out</div><ul>${li(ins.concerns)}</ul></div>
      <div class="insights-col indigo"><div class="insights-col-title">💡 Actions</div><ul>${li(ins.recommendations)}</ul></div>
    </div>
    ${extra?`<div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#9CA3AF;margin:14px 0 8px">📊 Additional angles</div><div class="insights-extra">${extra}</div></div>`:''}
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
  if (!preset || chatHistory.length > 1) document.getElementById('chatStarters').style.display = 'none';
  appendMsg('user', text);
  chatHistory.push({ role: 'user', content: text });
  const tid = 'typing-' + Date.now();
  appendTyping(tid);
  try {
    const res  = await fetch('/api/chat', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: chatHistory.slice(-12), context: currentAiCtx }),
    });
    const data = await res.json();
    removeTyping(tid);
    const reply = data.reply || 'Sorry, something went wrong.';
    appendMsg('bot', reply);
    chatHistory.push({ role: 'assistant', content: reply });
  } catch { removeTyping(tid); appendMsg('bot', 'Network error. Please try again.'); }
}
function appendMsg(role, text) {
  const msgs=document.getElementById('chatMessages'), div=document.createElement('div');
  div.className='msg msg--'+(role==='user'?'user':'bot'); div.textContent=text;
  msgs.appendChild(div); msgs.scrollTop=msgs.scrollHeight;
}
function appendTyping(id) {
  const msgs=document.getElementById('chatMessages'), div=document.createElement('div');
  div.id=id; div.className='msg msg--bot';
  div.innerHTML='<div class="typing-dots"><span></span><span></span><span></span></div>';
  msgs.appendChild(div); msgs.scrollTop=msgs.scrollHeight;
}
function removeTyping(id) { const el=document.getElementById(id); if(el)el.remove(); }

/* ══════════════════════════════════════════════════════════════
   INIT
   ══════════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  renderDashBar();
  renderViewBar();
  loadSheetList();
});
