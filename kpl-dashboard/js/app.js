// ═══════════════════════════════════════════════════════
// app.js · 頁面設定 / 導覽 / 初始化
// ═══════════════════════════════════════════════════════

const PAGES = [
  {
    group: '📊 儀表板',
    items: [
      { id:'daily',        icon:'📅', label:'每日動支監控', status:'ready' },
      { id:'dispatch',     icon:'💼', label:'總費用動支率', status:'ready' },
      { id:'freight',      icon:'🚚', label:'運費損益分析', status:'ready' },
      { id:'productivity', icon:'⚡', label:'人效監控',     status:'wip'   },
      { id:'monthly',      icon:'📆', label:'月度結算',     status:'wip'   },
    ]
  },
  {
    group: '📁 資料管理',
    items: [
      { id:'import',       icon:'📤', label:'資料匯入',     status:'placeholder' },
      { id:'manual-input', icon:'✏️', label:'手動輸入',     status:'ready' },
    ]
  },
];

let currentPageId = 'daily';

// ── 渲染左側選單 ──
function renderSidebar() {
  const sb = document.getElementById('sidebar');
  let html = '<div class="sb-nav">';
  PAGES.forEach(group => {
    html += `<div class="sb-group"><div class="sb-group-label">${group.group}</div>`;
    group.items.forEach(item => {
      const active = item.id === currentPageId ? 'active' : '';
      let badge = '';
      if (item.status === 'wip') badge = '<span class="sb-item-badge wip">WIP</span>';
      else if (item.status === 'placeholder') badge = '<span class="sb-item-badge soon">TBD</span>';
      html += `
      <a href="#${item.id}" class="sb-item ${active}" onclick="navigate(event, '${item.id}')">
        <span class="sb-item-icon">${item.icon}</span>
        <span class="sb-item-label">${item.label}</span>
        ${badge}
      </a>`;
    });
    html += '</div>';
  });
  html += '</div>';

  // 底部使用者資訊 + 登出
  const userId = sessionStorage.getItem('kpl_user') || '使用者';
  html += `
  <div class="sb-user">
    <div class="sb-user-avatar">👤</div>
    <div class="sb-user-info">
      <div class="sb-user-name">${userId}</div>
      <div class="sb-user-role">日翊文化行銷</div>
    </div>
    <button class="sb-logout" onclick="logout()" title="登出">⏻</button>
  </div>`;

  sb.innerHTML = html;
}

// ── 登出 ──
function logout() {
  sessionStorage.removeItem('kpl_user');
  sessionStorage.removeItem('kpl_auth');
  location.href = 'login.html';
}

// ── 切換頁面 ──
function navigate(event, pageId) {
  if (event) event.preventDefault();
  currentPageId = pageId;
  renderSidebar();
  document.getElementById('sidebar').classList.remove('open');
  loadPage(pageId);
}

// ── 載入頁面 ──
function loadPage(pageId) {
  const main = document.getElementById('main');
  const tpl = PAGE_TEMPLATES[pageId];
  if (!tpl) {
    main.innerHTML = '<div class="wip-page"><div class="wip-icon">⚠️</div><div class="wip-title">找不到頁面</div></div>';
    return;
  }
  main.innerHTML = tpl;

  if (pageId === 'daily')        initDailyPage();
  else if (pageId === 'dispatch') initDispatchPage();
  else if (pageId === 'freight') initFreightPage();
  else if (pageId === 'import')  initImportPage();
  else if (pageId === 'manual-input') initManualPage();
}

// ── 各頁面初始化 ──
function initDailyPage() {
  renderDailyPage();
  const modal = document.getElementById('daily-modal');
  if (modal) modal.addEventListener('click', e => {
    if (e.target.id === 'daily-modal') closeDailyModal();
  });
}

function initDispatchPage() { renderDispatchPage(); }
function initFreightPage()  { renderFreightPage(); }
function initImportPage()   { if (typeof updateStatus === 'function') updateStatus(); }
function initManualPage()   { if (typeof renderManualForm === 'function') renderManualForm(); }

// ── 手機側欄切換 ──
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

// ── 頂部時間 ──
function updateTime() {
  const el = document.getElementById('nav-time');
  if (el) el.textContent = new Date().toLocaleString('zh-TW', { hour12:false });
}

function checkMobile() {
  const toggle = document.getElementById('menu-toggle');
  if (!toggle) return;
  toggle.style.display = window.innerWidth <= 768 ? 'inline-flex' : 'none';
}

// ── 初始化 ──
document.addEventListener('DOMContentLoaded', () => {
  // 驗證登入
  if (!sessionStorage.getItem('kpl_auth')) {
    location.href = 'login.html';
    return;
  }

  renderSidebar();

  const hash = location.hash.slice(1);
  if (hash) {
    for (const group of PAGES) {
      if (group.items.find(i => i.id === hash)) {
        currentPageId = hash;
        break;
      }
    }
  }
  loadPage(currentPageId);

  updateTime();
  setInterval(updateTime, 60000);
  checkMobile();
  window.addEventListener('resize', checkMobile);
});

// ════════════════════════════════════════════
// Daily Page 邏輯
// ════════════════════════════════════════════
function renderDailyPage() {
  const grid = document.getElementById('daily-grid');
  grid.innerHTML = [renderM012(), renderM015()].join('');

  const now = new Date();
  document.getElementById('page-meta').textContent =
    `更新時間：${now.toLocaleString('zh-TW')} · 資料區間：${DATA.dateFrom} ~ ${DATA.dateTo}`;
  document.getElementById('filter-meta').textContent = `資料截至：${DATA.dateTo}`;
  document.getElementById('filter-from').value = DATA.dateFrom;
  document.getElementById('filter-to').value   = DATA.dateTo;
}

function applyFilter() {
  const from = document.getElementById('filter-from').value;
  const to   = document.getElementById('filter-to').value;
  if (from) DATA.dateFrom = from;
  if (to)   DATA.dateTo = to;
  renderDailyPage();
  toast('🔄 已更新');
}

function openDailyModal() {
  document.getElementById('in-budget').value     = DATA.budget;
  document.getElementById('in-actual').value     = DATA.actual;
  document.getElementById('in-day').value        = DATA.dayOfMonth;
  document.getElementById('in-total-days').value = DATA.totalDays;
  DATA.units.forEach((u, i) => {
    document.getElementById(`in-u${i}-fee`).value = u.fee;
    document.getElementById(`in-u${i}-hr`).value  = u.hr;
  });
  document.getElementById('in-peak').value   = DATA.thresholdPeak;
  document.getElementById('in-stable').value = DATA.thresholdStable;
  document.getElementById('daily-modal').classList.add('on');
}

function closeDailyModal() {
  document.getElementById('daily-modal').classList.remove('on');
}

function saveDailyData() {
  const g = id => { const v = document.getElementById(id).value; return v === '' ? null : Number(v); };
  DATA.budget     = g('in-budget')     ?? DATA.budget;
  DATA.actual     = g('in-actual')     ?? DATA.actual;
  DATA.dayOfMonth = g('in-day')        ?? DATA.dayOfMonth;
  DATA.totalDays  = g('in-total-days') ?? DATA.totalDays;
  DATA.units.forEach((u, i) => {
    u.fee = g(`in-u${i}-fee`) ?? u.fee;
    u.hr  = g(`in-u${i}-hr`)  ?? u.hr;
  });
  DATA.thresholdPeak   = g('in-peak')   ?? DATA.thresholdPeak;
  DATA.thresholdStable = g('in-stable') ?? DATA.thresholdStable;
  closeDailyModal();
  renderDailyPage();
  toast('✓ 資料已更新');
}

// ════════════════════════════════════════════
// Dispatch Page 邏輯
// ════════════════════════════════════════════
function renderDispatchPage() {
  const grid = document.getElementById('dispatch-grid');
  grid.innerHTML = [renderT001(), renderT002(), renderT003()].join('');

  document.getElementById('dispatch-from').value = DATA.dateFrom;
  document.getElementById('dispatch-to').value   = DATA.dateTo;
  const days = getDispatchDailyFiltered().length;
  document.getElementById('dispatch-meta').textContent =
    `資料區間：${DATA.dateFrom} ~ ${DATA.dateTo} · 共 ${days} 天 · 含人力+運務總覽`;
}

function applyDispatchFilter() {
  const from = document.getElementById('dispatch-from').value;
  const to   = document.getElementById('dispatch-to').value;
  if (from) DATA.dateFrom = from;
  if (to)   DATA.dateTo = to;
  renderDispatchPage();
  toast('🔄 已更新');
}

// ════════════════════════════════════════════
// Freight Page 邏輯
// ════════════════════════════════════════════
function renderFreightPage() {
  const grid = document.getElementById('freight-grid');
  grid.innerHTML = [renderF001(), renderF002(), renderF003(), renderF009(), renderF010()].join('');

  document.getElementById('freight-from').value = DATA.dateFrom;
  document.getElementById('freight-to').value   = DATA.dateTo;
  document.getElementById('freight-meta').textContent =
    `資料區間：${DATA.dateFrom} ~ ${DATA.dateTo} · 共 ${DATA.freight.totalOrders.toLocaleString()} 筆配送`;
}

function applyFreightFilter() {
  const from = document.getElementById('freight-from').value;
  const to   = document.getElementById('freight-to').value;
  if (from) DATA.dateFrom = from;
  if (to)   DATA.dateTo = to;
  renderFreightPage();
  toast('🔄 已更新');
}

// ════════════════════════════════════════════
// Import Page 邏輯
// ════════════════════════════════════════════
let parsedFreight = null;

function onDragOver(e, id) {
  e.preventDefault();
  const el = document.getElementById(id);
  el.style.borderColor = 'var(--ry-blue)';
  el.style.background = 'var(--ry-blue-pale)';
}
function onDragLeave(id) {
  const el = document.getElementById(id);
  el.style.borderColor = 'var(--ry-line)';
  el.style.background = 'var(--ry-bg)';
}
function onDrop(e, type) {
  e.preventDefault();
  onDragLeave(type + '-drop');
  if (e.dataTransfer.files[0]) parseExcel(e.dataTransfer.files[0], type);
}
function onFileSelect(e, type) {
  if (e.target.files[0]) parseExcel(e.target.files[0], type);
}

function parseExcel(file, type) {
  document.getElementById(type + '-status').textContent = '解析中…';
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array' });
      if (type === 'freight') parseFreight(wb, file.name);
    } catch(err) {
      document.getElementById(type + '-status').textContent = '❌ 解析失敗';
      toast('❌ ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseFreight(wb, fileName) {
  const SHEET = '進貨日與計價費用';
  if (!wb.SheetNames.includes(SHEET)) {
    toast('❌ 找不到「' + SHEET + '」分頁');
    document.getElementById('freight-status').textContent = '❌ 分頁不存在';
    return;
  }
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[SHEET], { defval:0 });
  const rows = [];
  raw.forEach(r => {
    const d = String(r['列標籤'] || '');
    if (!d || d.includes('總計')) return;
    const parts = d.split('/');
    if (parts.length < 3) return;
    const date = parts[1].padStart(2,'0') + '/' + parts[2].padStart(2,'0');
    rows.push({ date, daxi: Number(r['大溪倉'] || 0), dadu: Number(r['大肚倉'] || 0), gangshan: Number(r['岡山倉'] || 0) });
  });
  if (!rows.length) { toast('❌ 找不到有效資料'); return; }

  const totals = {
    daxi:     rows.reduce((s,r)=>s+r.daxi,0),
    dadu:     rows.reduce((s,r)=>s+r.dadu,0),
    gangshan: rows.reduce((s,r)=>s+r.gangshan,0),
  };
  parsedFreight = { rows, fileName, totals, at: new Date() };
  document.getElementById('freight-status').textContent = `✅ ${rows.length} 天`;
  showFreightPreview(rows, totals);
  const btns = document.getElementById('freight-btns');
  btns.style.display = 'flex';
  toast(`✅ 解析完成：${rows.length} 天資料`);
}

function showFreightPreview(rows, totals) {
  const show = rows.length <= 8 ? rows :
    [...rows.slice(0,4), {date:'…',daxi:'…',dadu:'…',gangshan:'…'}, ...rows.slice(-3)];

  const trs = show.map(r => {
    const fmt = v => v === '…' ? '…' : '$' + Number(v).toLocaleString();
    return `<tr>
      <td style="padding:6px 10px;font-weight:700">${r.date}</td>
      <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);color:#0E7BAD">${fmt(r.daxi)}</td>
      <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);color:#2DA870">${fmt(r.dadu)}</td>
      <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);color:#E07855">${fmt(r.gangshan)}</td>
    </tr>`;
  }).join('');

  document.getElementById('freight-preview').innerHTML = `
    <div style="font-size:var(--fs-xs);font-weight:700;color:var(--ry-muted);margin-bottom:6px">📋 資料預覽（共 ${rows.length} 天）</div>
    <div style="max-height:200px;overflow-y:auto;border:1px solid var(--ry-line);border-radius:3px;font-size:var(--fs-xs)">
      <table style="width:100%;border-collapse:collapse">
        <thead style="position:sticky;top:0;background:var(--ry-blue-dark)">
          <tr>
            <th style="padding:6px 10px;color:white;text-align:left">日期</th>
            <th style="padding:6px 10px;color:#7ec8e3;text-align:right">大溪倉</th>
            <th style="padding:6px 10px;color:#7ed4a0;text-align:right">大肚倉</th>
            <th style="padding:6px 10px;color:#f0b090;text-align:right">岡山倉</th>
          </tr>
        </thead>
        <tbody>${trs}</tbody>
        <tfoot style="background:var(--ry-blue-pale);border-top:2px solid var(--ry-blue-dark)">
          <tr>
            <td style="padding:6px 10px;font-weight:800;color:var(--ry-blue-dark)">合計</td>
            <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);font-weight:800;color:#0E7BAD">$${totals.daxi.toLocaleString()}</td>
            <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);font-weight:800;color:#2DA870">$${totals.dadu.toLocaleString()}</td>
            <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);font-weight:800;color:#E07855">$${totals.gangshan.toLocaleString()}</td>
          </tr>
        </tfoot>
      </table>
    </div>`;
  document.getElementById('freight-preview').style.display = 'block';
}

function applyFreight() {
  if (!parsedFreight) return;
  DATA.freight.dailyByWarehouse = parsedFreight.rows.map(r => [r.date, r.daxi, r.gangshan, r.dadu]);
  const map = {};
  parsedFreight.rows.forEach(r => { map[r.date] = r; });
  DATA.dispatch.daily = DATA.dispatch.daily.map(row => {
    const f = map[row[0]];
    return f ? [row[0], row[1], f.daxi, row[3], f.dadu, row[5], f.gangshan] : row;
  });
  updateStatus();
  toast('✅ 運務資料已套用！切換至「總費用動支率」頁查看');
}

function downloadFreight() {
  if (!parsedFreight) return;
  const { rows, totals, fileName, at } = parsedFreight;
  const headers = ['日期', '大溪倉', '大肚倉', '岡山倉', '當日合計'];
  const data = rows.map(r => [r.date, r.daxi, r.dadu, r.gangshan, r.daxi+r.dadu+r.gangshan]);
  const total = ['月合計', totals.daxi, totals.dadu, totals.gangshan, totals.daxi+totals.dadu+totals.gangshan];
  const ws = XLSX.utils.aoa_to_sheet([headers, ...data, total]);
  ws['!cols'] = [{wch:10},{wch:14},{wch:14},{wch:14},{wch:14}];
  const info = XLSX.utils.aoa_to_sheet([
    ['KPL 儀表板 · 運費彙總確認'],[''],
    ['來源檔案', fileName],['解析時間', at.toLocaleString('zh-TW')],
    ['資料天數', rows.length],['大溪倉合計', totals.daxi],
    ['大肚倉合計', totals.dadu],['岡山倉合計', totals.gangshan],
    ['三倉總計', totals.daxi+totals.dadu+totals.gangshan],
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '運費彙總');
  XLSX.utils.book_append_sheet(wb, info, '資料說明');
  const n = at;
  const name = `運費彙總_${n.getFullYear()}${String(n.getMonth()+1).padStart(2,'0')}${String(n.getDate()).padStart(2,'0')}.xlsx`;
  XLSX.writeFile(wb, name);
  toast('⬇️ 已下載：' + name);
}

function resetFreight() {
  parsedFreight = null;
  document.getElementById('freight-status').textContent = '尚未上傳';
  document.getElementById('freight-preview').style.display = 'none';
  document.getElementById('freight-btns').style.display = 'none';
  document.getElementById('freight-file').value = '';
}

function updateStatus() {
  document.getElementById('status-time').textContent = '更新：' + new Date().toLocaleString('zh-TW');
  let daxiF=0, daduF=0, gsF=0, daxiL=0, daduL=0, gsL=0;
  DATA.dispatch.daily.forEach(r => { daxiL+=r[1]; daxiF+=r[2]; daduL+=r[3]; daduF+=r[4]; gsL+=r[5]; gsF+=r[6]; });
  const rows = [
    { type:'🚚 運務費用', real:!!parsedFreight, daxi:daxiF, dadu:daduF, gs:gsF },
    { type:'💰 人力費用', real:false,           daxi:daxiL, dadu:daduL, gs:gsL },
  ];
  document.getElementById('status-tbody').innerHTML = rows.map(r => {
    const total = r.daxi + r.dadu + r.gs;
    const c = r.real ? '#1b7c33' : '#e07855';
    const s = r.real ? '✅ 已上傳' : '⚠️ 範例資料';
    return `<tr>
      <td style="font-weight:700">${r.type}</td>
      <td><span style="color:${c};font-weight:700">${s}</span></td>
      <td class="mono" style="text-align:right;color:#0E7BAD">$${r.daxi.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:#2DA870">$${r.dadu.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:#E07855">$${r.gs.toLocaleString()}</td>
      <td class="mono" style="text-align:right;font-weight:800">$${total.toLocaleString()}</td>
    </tr>`;
  }).join('');
}

// ════════════════════════════════════════════
// Manual Input Page 邏輯
// ════════════════════════════════════════════
function renderManualForm() {
  document.getElementById('mi-budget').value     = DATA.budget;
  document.getElementById('mi-actual').value     = DATA.actual;
  document.getElementById('mi-day').value        = DATA.dayOfMonth;
  document.getElementById('mi-total-days').value = DATA.totalDays;
  document.getElementById('mi-peak').value       = DATA.thresholdPeak;
  document.getElementById('mi-stable').value     = DATA.thresholdStable;

  const tbody = document.getElementById('mi-units-tbody');
  tbody.innerHTML = DATA.units.map((u, i) => `
    <tr>
      <td>${u.name}${u.type === 'sup' ? ' <span class="tag" style="background:#eff6ff;color:#1e5ca8;border-color:#bfdbfe">SUP</span>' : ''}</td>
      <td><input type="number" class="filter-input" style="width:140px" id="mi-u${i}-fee" value="${u.fee}"></td>
      <td><input type="number" class="filter-input" style="width:100px" id="mi-u${i}-hr" value="${u.hr}"></td>
    </tr>
  `).join('');
}

function saveManualData() {
  const g = id => { const v = document.getElementById(id).value; return v === '' ? null : Number(v); };
  DATA.budget     = g('mi-budget')     ?? DATA.budget;
  DATA.actual     = g('mi-actual')     ?? DATA.actual;
  DATA.dayOfMonth = g('mi-day')        ?? DATA.dayOfMonth;
  DATA.totalDays  = g('mi-total-days') ?? DATA.totalDays;
  DATA.thresholdPeak   = g('mi-peak')   ?? DATA.thresholdPeak;
  DATA.thresholdStable = g('mi-stable') ?? DATA.thresholdStable;
  DATA.units.forEach((u, i) => {
    u.fee = g(`mi-u${i}-fee`) ?? u.fee;
    u.hr  = g(`mi-u${i}-hr`)  ?? u.hr;
  });
  toast('✓ 資料已儲存，切換到儀表板可看到更新');
}

function resetManualForm() {
  renderManualForm();
  toast('↺ 已重新載入');
}
