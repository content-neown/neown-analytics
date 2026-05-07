/* ── Config ── */
const PINK   = '#F0008C';
const INDIGO = '#3C28B4';
const GRAY   = '#9CA3AF';
const PINK_L = '#FFD6EE';
const IND_L  = '#C7C2F0';

/* ── State ── */
let currentSheet  = '';
let currentKPIs   = null;
let currentAiCtx  = '';
let chatHistory   = [];
let chartInstances = {};

/* ── Helpers ── */
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
    if (Math.abs(n) >= 100000)  return '₹' + (n/100000).toFixed(2) + ' L';
    if (Math.abs(n) >= 1000)    return '₹' + (n/1000).toFixed(1) + 'K';
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

function setStatus(state, text) {
  const dot  = document.getElementById('statusDot').querySelector('.dot');
  const span = document.getElementById('statusText');
  dot.className = 'dot dot--' + state;
  span.textContent = text;
}

function showError(msg) {
  const bar = document.getElementById('errorBar');
  document.getElementById('errorMsg').textContent = msg;
  bar.hidden = !msg;
}

/* ── Data processing ── */
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
    days:            rows.length,
    totalSpendsGST:  sum('Spends + GST'),
    fbSpendsGST:     sum('FB Spends with GST'),
    googleSpendsGST: sum('Google Spends with GST'),
    totalDBPurchase: sum('Total DB Purchase'),
    fbDBPurchase:    sum('FB DB Purchase'),
    googleDBPurchase:sum('Google DB Purchase'),
    totalDBCPS:      avg('Total DB CPS'),
    fbDBCPS:         avg('FB DB CPS'),
    googleDBCPS:     avg('Google DB CPS'),
    revenueTotal:    sum('Revenue (FB + Unknown + Google)'),
    revenueFB:       sum('Revenue (FB UTM)'),
    revenueGoogle:   sum('Revenue (Google UTM)'),
    revenueUnknown:  sum('Revenue (Unknown)'),
    roasGST:         avg('ROAS (FB + Unknown + Google) WITH GST'),
    fbDBROAS:        avg('FB DB ROAS'),
    googleDBROAS:    avg('Google DB ROAS'),
    shopifyRevenue:  sum('Shopify Daily Revenue (TOTAL)'),
    shopifyCount:    sum('Shopify Daily Sales Count (TOTAL)'),
    fbCPM:           avg('FB CPM'),
    fbCTR:           avg('FB CTR'),
    googleCTR:       avg('Google CTR'),
    fbCPC:           avg('FB CPC'),
    googleCPC:       avg('Google CPC'),
    fbClicks:        sum('FB Clicks'),
    googleClicks:    sum('Google Clicks'),
    totalClicks:     sum('Total Clicks'),
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
  const top = [...chartData].sort((a,b) => b.revenue - a.revenue).slice(0,3)
    .map(d => `${d.date} (Rev ${fINR(d.revenue,true)}, ROAS ${d.roas.toFixed(2)}x)`).join('; ');
  const bot = [...chartData].filter(d => d.totalSpends > 0).sort((a,b) => a.roas - b.roas).slice(0,3)
    .map(d => `${d.date} (ROAS ${d.roas.toFixed(2)}x, CPS ${fINR(d.totalDBCPS,true)})`).join('; ');
  return `neOwn Performance Marketing — ${sheet} (${kpis.days} days)
SPENDS: Total ${fINR(kpis.totalSpendsGST,true)} | FB ${fINR(kpis.fbSpendsGST,true)} (${pct(kpis.fbSpendsGST,kpis.totalSpendsGST)}%) | Google ${fINR(kpis.googleSpendsGST,true)} (${pct(kpis.googleSpendsGST,kpis.totalSpendsGST)}%)
PURCHASES: Total ${Math.round(kpis.totalDBPurchase)} | FB ${Math.round(kpis.fbDBPurchase)} | Google ${Math.round(kpis.googleDBPurchase)}
AVG CPS: Total ${fINR(kpis.totalDBCPS,true)} | FB ${fINR(kpis.fbDBCPS,true)} | Google ${fINR(kpis.googleDBCPS,true)}
REVENUE: Attribution ${fINR(kpis.revenueTotal,true)} | Shopify ${fINR(kpis.shopifyRevenue,true)} (${Math.round(kpis.shopifyCount)} orders)
ROAS (GST): ${kpis.roasGST.toFixed(2)}x | FB DB ${kpis.fbDBROAS.toFixed(2)}x | Google DB ${kpis.googleDBROAS.toFixed(2)}x
FB ENGAGEMENT: CPM ${fINR(kpis.fbCPM,true)} | CTR ${kpis.fbCTR.toFixed(2)}% | CPC ${fINR(kpis.fbCPC,true)} | Clicks ${fNum(kpis.fbClicks)}
BEST DAYS: ${top}
WORST ROAS DAYS: ${bot}`;
}

/* ── Render KPI cards ── */
function renderKPIs(kpis) {
  const avgOrder = kpis.shopifyCount > 0 ? kpis.shopifyRevenue / kpis.shopifyCount : 0;
  const cards = [
    { label: 'Total Spends (GST)',     value: fINR(kpis.totalSpendsGST,true),    sub: `FB ${fINR(kpis.fbSpendsGST,true)} · G ${fINR(kpis.googleSpendsGST,true)}`, accent: true },
    { label: 'Total DB Purchases',     value: fNum(kpis.totalDBPurchase),         sub: `FB ${fNum(kpis.fbDBPurchase)} · G ${fNum(kpis.googleDBPurchase)}` },
    { label: 'Avg DB CPS',             value: fINR(kpis.totalDBCPS,true),         sub: `FB ${fINR(kpis.fbDBCPS,true)} · G ${fINR(kpis.googleDBCPS,true)}` },
    { label: 'Attribution Revenue',    value: fINR(kpis.revenueTotal,true),       sub: `FB ${fINR(kpis.revenueFB,true)} · G ${fINR(kpis.revenueGoogle,true)}` },
    { label: 'ROAS (with GST)',        value: kpis.roasGST.toFixed(2) + 'x',     sub: `FB ${kpis.fbDBROAS.toFixed(2)}x · G ${kpis.googleDBROAS.toFixed(2)}x`, accent: true },
    { label: 'Shopify Revenue',        value: fINR(kpis.shopifyRevenue,true),     sub: `${fNum(kpis.shopifyCount)} orders · ₹${fNum(avgOrder)} AOV` },
    { label: 'FB CTR',                 value: kpis.fbCTR.toFixed(2) + '%',        sub: `CPM ${fINR(kpis.fbCPM,true)} · CPC ${fINR(kpis.fbCPC,true)}` },
    { label: 'Total Clicks',           value: fNum(kpis.totalClicks),             sub: `FB ${fNum(kpis.fbClicks)} · G ${fNum(kpis.googleClicks)}` },
  ];
  document.getElementById('kpiGrid').innerHTML = cards.map(c => `
    <div class="kpi-card ${c.accent ? 'accent' : ''}">
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-sub">${c.sub}</div>
    </div>
  `).join('');
}

function showKPISkeleton() {
  document.getElementById('kpiGrid').innerHTML = Array(8).fill('<div class="kpi-skeleton"></div>').join('');
}

/* ── Charts ── */
const chartDefaults = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: { labels: { font: { family: 'Inter', size: 10 }, boxWidth: 10, padding: 12 } },
    tooltip: { bodyFont: { family: 'Inter', size: 11 }, titleFont: { family: 'Inter', size: 11 } },
  },
  scales: {
    x: {
      grid: { color: '#F3F4F6' },
      ticks: { font: { family: 'Inter', size: 9 }, maxRotation: 0 },
    },
    y: {
      grid: { color: '#F3F4F6' },
      ticks: { font: { family: 'Inter', size: 9 } },
    },
  },
};

function makeChart(id, config) {
  if (chartInstances[id]) { chartInstances[id].destroy(); }
  const ctx = document.getElementById(id);
  if (!ctx) return;
  chartInstances[id] = new Chart(ctx, config);
}

function tickInterval(len) {
  if (len <= 8)  return 0;
  if (len <= 16) return 1;
  if (len <= 31) return 2;
  return Math.floor(len / 10);
}

function renderCharts(data) {
  if (!data.length) return;
  const labels   = data.map(d => shortDate(d.date));
  const interval = tickInterval(labels.length);
  const xOpts    = { ...chartDefaults.scales.x, ticks: { ...chartDefaults.scales.x.ticks, maxTicksLimit: 10, callback: (v, i) => i % (interval+1) === 0 ? labels[i] : '' }};

  // ── 1. Spends ──
  makeChart('cSpends', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total', data: data.map(d => d.totalSpends), borderColor: '#374151', borderWidth: 2, pointRadius: 0, tension: .3 },
        { label: 'Facebook', data: data.map(d => d.fbSpends), borderColor: PINK, borderWidth: 1.5, pointRadius: 0, tension: .3 },
        { label: 'Google',   data: data.map(d => d.googleSpends), borderColor: INDIGO, borderWidth: 1.5, pointRadius: 0, tension: .3 },
      ],
    },
    options: { ...chartDefaults, scales: { x: xOpts, y: { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => v >= 100000 ? '₹'+(v/100000).toFixed(1)+'L' : v >= 1000 ? '₹'+(v/1000).toFixed(0)+'K' : '₹'+v } } } },
  });

  // ── 2. Revenue ──
  makeChart('cRevenue', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Shopify Revenue', data: data.map(d => d.shopifyRevenue), borderColor: INDIGO, backgroundColor: 'rgba(60,40,180,.08)', fill: true, borderWidth: 2, pointRadius: 0, tension: .3 },
        { label: 'Attribution Revenue', data: data.map(d => d.revenue), borderColor: PINK, borderWidth: 2, pointRadius: 0, tension: .3 },
      ],
    },
    options: { ...chartDefaults, scales: { x: xOpts, y: { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => v >= 100000 ? '₹'+(v/100000).toFixed(1)+'L' : v >= 1000 ? '₹'+(v/1000).toFixed(0)+'K' : '₹'+v } } } },
  });

  // ── 3. DB Purchases ──
  makeChart('cPurchases', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Facebook', data: data.map(d => d.fbDBPurchase),     backgroundColor: PINK,  borderRadius: 3 },
        { label: 'Google',   data: data.map(d => d.googleDBPurchase), backgroundColor: INDIGO, borderRadius: 3 },
      ],
    },
    options: { ...chartDefaults, scales: { x: { ...xOpts, stacked: false }, y: { ...chartDefaults.scales.y, stacked: false } } },
  });

  // ── 4. ROAS ──
  makeChart('cROAS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total ROAS', data: data.map(d => +d.roas.toFixed(2)),       borderColor: '#374151', borderWidth: 2, pointRadius: 0, tension: .3 },
        { label: 'FB ROAS',    data: data.map(d => +d.fbROAS.toFixed(2)),     borderColor: PINK,   borderWidth: 1.5, pointRadius: 0, tension: .3, borderDash: [5,3] },
        { label: 'Google ROAS',data: data.map(d => +d.googleROAS.toFixed(2)), borderColor: INDIGO, borderWidth: 1.5, pointRadius: 0, tension: .3, borderDash: [5,3] },
      ],
    },
    options: { ...chartDefaults, scales: { x: xOpts, y: { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => v.toFixed(1)+'x' } } } },
  });

  // ── 5. CPS ──
  makeChart('cCPS', {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label: 'Total CPS',  data: data.map(d => Math.round(d.totalDBCPS)),  borderColor: '#374151', borderWidth: 2, pointRadius: 0, tension: .3 },
        { label: 'FB CPS',     data: data.map(d => Math.round(d.fbDBCPS)),     borderColor: PINK,   borderWidth: 1.5, pointRadius: 0, tension: .3 },
        { label: 'Google CPS', data: data.map(d => Math.round(d.googleDBCPS)), borderColor: INDIGO, borderWidth: 1.5, pointRadius: 0, tension: .3 },
      ],
    },
    options: { ...chartDefaults, scales: { x: xOpts, y: { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => '₹'+v } } } },
  });

  // ── 6. CTR & CPM ──
  makeChart('cCTR', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'FB CTR %',  data: data.map(d => +d.fbCTR.toFixed(2)),  backgroundColor: PINK_L, borderRadius: 3, yAxisID: 'y' },
        { label: 'FB CPM ₹',  data: data.map(d => Math.round(d.fbCPM)),   type: 'line', borderColor: PINK, borderWidth: 2, pointRadius: 0, tension: .3, yAxisID: 'y2' },
      ],
    },
    options: {
      ...chartDefaults,
      scales: {
        x: xOpts,
        y:  { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => v.toFixed(1)+'%' }, position: 'left' },
        y2: { ...chartDefaults.scales.y, ticks: { ...chartDefaults.scales.y.ticks, callback: v => '₹'+v },             position: 'right', grid: { drawOnChartArea: false } },
      },
    },
  });
}

/* ── Load sheet data ── */
async function loadSheet(sheetName) {
  currentSheet = sheetName;
  setStatus('loading', 'Loading…');
  showError('');
  showKPISkeleton();
  document.getElementById('insightsBody').innerHTML = '<div class="insights-empty">Select a month and click Generate Insights.</div>';
  document.getElementById('insightsSub').textContent = 'Claude-powered analysis · ' + sheetName;

  try {
    const res  = await fetch('/api/sheets?sheet=' + encodeURIComponent(sheetName));
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    const rows      = processRows(data.values || []);
    const kpis      = computeKPIs(rows);
    const chartData = buildChartData(rows);
    currentKPIs     = kpis;
    currentAiCtx    = buildAiContext(sheetName, kpis, chartData);

    renderKPIs(kpis);
    renderCharts(chartData);
    setStatus('ok', rows.length + ' days loaded');
  } catch (err) {
    showError('Error loading ' + sheetName + ': ' + err.message);
    setStatus('error', 'Error');
  }
}

/* ── Load sheet list ── */
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

    const sheets = data.sheets || [];
    if (!sheets.length) {
      showError('No sheets found in the spreadsheet.');
      setStatus('error', 'No sheets');
      return;
    }

    const tabsEl = document.getElementById('monthTabs');
    tabsEl.innerHTML = sheets.map(s => `<button class="tab" onclick="switchTab(this, '${s}')">${s}</button>`).join('');

    // Auto-select last sheet
    const lastTab = tabsEl.lastElementChild;
    lastTab.classList.add('active');
    loadSheet(lastTab.textContent);

  } catch (err) {
    showError('Could not reach /api/sheets. Make sure GOOGLE_SHEETS_API_KEY is set in Vercel Environment Variables.');
    document.getElementById('setupHint').hidden = false;
    setStatus('error', 'Error');
  }
}

function switchTab(btn, sheetName) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  loadSheet(sheetName);
}

/* ── AI Insights ── */
async function generateInsights() {
  if (!currentAiCtx) return;
  const btn  = document.getElementById('insightsBtn');
  const body = document.getElementById('insightsBody');
  btn.disabled = true;
  btn.textContent = 'Analyzing…';
  body.innerHTML = `<div class="insights-skeleton">${Array(5).fill('<div style="width:${70+Math.random()*25}%"></div>').join('')}</div>`;

  try {
    const res  = await fetch('/api/insights', {
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
      <div class="insights-col green">
        <div class="insights-col-title">🔥 Highlights</div>
        <ul>${li(ins.highlights)}</ul>
      </div>
      <div class="insights-col orange">
        <div class="insights-col-title">⚠️ Watch out</div>
        <ul>${li(ins.concerns)}</ul>
      </div>
      <div class="insights-col indigo">
        <div class="insights-col-title">💡 Actions</div>
        <ul>${li(ins.recommendations)}</ul>
      </div>
    </div>
    ${extra ? `<div><div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#9CA3AF;margin:14px 0 8px">📊 Additional angles to explore</div><div class="insights-extra">${extra}</div></div>` : ''}
  `;
}

/* ── Chatbot ── */
function toggleChat() {
  const win = document.getElementById('chatWindow');
  win.classList.toggle('open');
  if (win.classList.contains('open')) {
    document.getElementById('chatInput').focus();
  }
}

async function sendChat(preset) {
  const input = document.getElementById('chatInput');
  const text  = preset || input.value.trim();
  if (!text) return;
  input.value = '';

  // Hide starters after first real message
  if (!preset || chatHistory.length > 1) {
    document.getElementById('chatStarters').style.display = 'none';
  }

  appendMsg('user', text);
  chatHistory.push({ role: 'user', content: text });

  // Typing indicator
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

/* ── Init ── */
document.addEventListener('DOMContentLoaded', loadSheetList);
