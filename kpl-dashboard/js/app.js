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
      { id:'picks',        icon:'⚡', label:'揀次分析',     status:'ready' },
      { id:'labor',        icon:'⏱', label:'工時結構分析', status:'ready' },
      { id:'productivity', icon:'📊', label:'人效監控',     status:'wip'   },
      { id:'monthly',      icon:'📆', label:'月度結算',     status:'wip'   },
    ]
  },
  {
    group: '📁 資料管理',
    items: [
      { id:'import',       icon:'📤', label:'資料匯入',     status:'ready' },
      { id:'org',          icon:'🏢', label:'組織設定',     status:'ready' },
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
  else if (pageId === 'picks')   initPicksPage();
  else if (pageId === 'labor')   initLaborPage();
  else if (pageId === 'import')  initImportPage();
  else if (pageId === 'org')     initOrgPage();
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
let parsedLabor   = null;
let parsedPicks   = null;

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
      else if (type === 'labor') parseLabor(wb, file.name);
      else if (type === 'picks') parsePicks(wb, file.name);
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

function parseLabor(wb, fileName) {
  const sheetName = wb.SheetNames[0];
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
  if (!raw.length) { toast('❌ 找不到有效資料'); return; }

  const sample = raw[0];
  const required = ['倉別', '日期', '廠商', '班別', '員編', '作業課別', '姓名', '作業區域', '作業時數', '實際費用'];
  const missing = required.filter(c => !(c in sample));
  if (missing.length) {
    toast('❌ 缺少欄位：' + missing.join('、'));
    document.getElementById('labor-status').textContent = '❌ 格式不符';
    return;
  }

  const records = [];
  raw.forEach(r => {
    const hrs = Number(r['作業時數']) || 0;
    if (hrs <= 0) return;
    const opArea = String(r['作業區域'] || '');
    if (opArea === '午休時間') return;
    const serial = Number(r['日期']);
    let dateStr = '';
    if (serial > 0) {
      const d = new Date(Math.round((serial - 25569) * 86400000));
      dateStr = d.toISOString().slice(0, 10);
    }
    records.push({
      wh:       String(r['倉別'] || ''),
      date:     dateStr,
      vendor:   String(r['廠商'] || ''),
      shift:    String(r['班別'] || ''),
      empId:    String(r['員編'] || ''),
      dept:     String(r['作業課別'] || ''),
      name:     String(r['姓名'] || ''),
      opArea,
      hours:    Math.round(hrs * 100) / 100,
      boxHours: Math.round((Number(r['裝箱時數'])  || 0) * 100) / 100,
      nightHrs: Math.round((Number(r['夜間時數'])  || 0) * 100) / 100,
      normHrs:  Math.round((Number(r['正常時數'])  || 0) * 100) / 100,
      cost:     Number(r['實際費用']) || 0,
    });
  });

  if (!records.length) { toast('❌ 找不到有效工時記錄'); return; }

  const totalHrs  = records.reduce((s, r) => s + r.hours, 0);
  const totalCost = records.reduce((s, r) => s + r.cost,  0);
  parsedLabor = { records, fileName, at: new Date() };
  document.getElementById('labor-status').textContent = `✅ ${records.length} 筆 · ${totalHrs.toFixed(1)}h`;
  showLaborPreview(records, totalHrs, totalCost);
  document.getElementById('labor-btns').style.display = 'flex';
  toast(`✅ 工時解析完成：${records.length} 筆記錄`);
}

function showLaborPreview(records, totalHrs, totalCost) {
  const byOp = {};
  records.forEach(r => {
    if (!byOp[r.opArea]) byOp[r.opArea] = { hrs: 0, cost: 0, count: 0 };
    byOp[r.opArea].hrs   += r.hours;
    byOp[r.opArea].cost  += r.cost;
    byOp[r.opArea].count++;
  });
  const trs = Object.entries(byOp)
    .sort((a, b) => b[1].hrs - a[1].hrs)
    .map(([op, it]) => `<tr>
      <td style="padding:6px 10px;font-weight:700">${op}</td>
      <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono)">${it.hrs.toFixed(1)} h</td>
      <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono)">$${it.cost.toLocaleString()}</td>
      <td style="padding:6px 10px;text-align:right;color:var(--ry-muted)">${it.count}</td>
    </tr>`).join('');

  document.getElementById('labor-preview').innerHTML = `
    <div style="font-size:var(--fs-xs);font-weight:700;color:var(--ry-muted);margin-bottom:6px">
      📋 作業區域摘要（${records.length} 筆 / ${totalHrs.toFixed(1)}h / $${totalCost.toLocaleString()}）
    </div>
    <div style="max-height:180px;overflow-y:auto;border:1px solid var(--ry-line);border-radius:3px;font-size:var(--fs-xs)">
      <table style="width:100%;border-collapse:collapse">
        <thead style="position:sticky;top:0;background:var(--ry-blue-dark)">
          <tr>
            <th style="padding:6px 10px;color:white;text-align:left">作業區域</th>
            <th style="padding:6px 10px;color:#b3d4f5;text-align:right">工時</th>
            <th style="padding:6px 10px;color:#b3d4f5;text-align:right">費用</th>
            <th style="padding:6px 10px;color:#b3d4f5;text-align:right">筆數</th>
          </tr>
        </thead>
        <tbody>${trs}</tbody>
        <tfoot style="background:var(--ry-blue-pale,#eff6ff);border-top:2px solid var(--ry-blue-dark)">
          <tr>
            <td style="padding:6px 10px;font-weight:800;color:var(--ry-blue-dark)">合計</td>
            <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);font-weight:800;color:var(--ry-blue)">${totalHrs.toFixed(1)} h</td>
            <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);font-weight:800;color:var(--ry-blue)">$${totalCost.toLocaleString()}</td>
            <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);font-weight:800;color:var(--ry-blue)">${records.length}</td>
          </tr>
        </tfoot>
      </table>
    </div>`;
  document.getElementById('labor-preview').style.display = 'block';
}

function applyLabor() {
  if (!parsedLabor) return;
  LABOR_RAW = parsedLabor.records;
  if (currentPageId === 'labor') renderLaborPage();
  updateStatus();
  toast('✅ 工時資料已套用！切換至「工時結構分析」頁查看');
}

function resetLabor() {
  parsedLabor = null;
  document.getElementById('labor-status').textContent = '尚未上傳';
  document.getElementById('labor-preview').style.display = 'none';
  document.getElementById('labor-btns').style.display = 'none';
  document.getElementById('labor-file').value = '';
}

function parsePicks(wb, fileName) {
  const sheetName = wb.SheetNames[0];
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
  if (!raw.length) { toast('❌ 找不到有效資料'); return; }

  const sample = raw[0];
  const required = ['倉別', '日期', '業務類別', '作業區', '工時區域', '揀次'];
  const missing = required.filter(c => !(c in sample));
  if (missing.length) {
    toast('❌ 缺少欄位：' + missing.join('、'));
    document.getElementById('picks-import-status').textContent = '❌ 格式不符';
    return;
  }

  const records = [];
  raw.forEach(r => {
    const p = Number(r['揀次']) || 0;
    if (p <= 0) return;
    const serial = Number(r['日期']);
    let dateStr = '';
    if (serial > 0) {
      const d = new Date(Math.round((serial - 25569) * 86400000));
      dateStr = d.toISOString().slice(0, 10);
    }
    records.push({
      date:  dateStr,
      wh:    String(r['倉別'] || ''),
      biz:   String(r['業務類別'] || ''),
      area:  String(r['作業區'] || ''),
      op:    String(r['工時區域'] || ''),
      picks: p,
    });
  });

  if (!records.length) { toast('❌ 找不到有效揀次記錄'); return; }

  const totals = {};
  records.forEach(r => { totals[r.wh] = (totals[r.wh] || 0) + r.picks; });
  const totalPicks = records.reduce((s, r) => s + r.picks, 0);

  parsedPicks = { records, fileName, totals, at: new Date() };
  document.getElementById('picks-import-status').textContent = `✅ ${records.length} 筆 · ${totalPicks.toLocaleString()} 次`;
  showPicksPreview(records, totals, totalPicks);
  document.getElementById('picks-btns').style.display = 'flex';
  toast(`✅ 揀次解析完成：${records.length} 筆記錄`);
}

function showPicksPreview(records, totals, totalPicks) {
  const byOp = {};
  records.forEach(r => { byOp[r.op] = (byOp[r.op] || 0) + r.picks; });
  const trs = Object.entries(byOp)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([op, cnt]) => `<tr>
      <td style="padding:6px 10px;font-weight:700">${op}</td>
      <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono)">${cnt.toLocaleString()}</td>
    </tr>`).join('');

  document.getElementById('picks-preview').innerHTML = `
    <div style="font-size:var(--fs-xs);font-weight:700;color:var(--ry-muted);margin-bottom:6px">📋 工時區域摘要（共 ${totalPicks.toLocaleString()} 揀次）</div>
    <div style="max-height:180px;overflow-y:auto;border:1px solid var(--ry-line);border-radius:3px;font-size:var(--fs-xs)">
      <table style="width:100%;border-collapse:collapse">
        <thead style="position:sticky;top:0;background:var(--ry-blue-dark)">
          <tr>
            <th style="padding:6px 10px;color:white;text-align:left">工時區域</th>
            <th style="padding:6px 10px;color:#b3d4f5;text-align:right">揀次</th>
          </tr>
        </thead>
        <tbody>${trs}</tbody>
        <tfoot style="background:var(--ry-blue-pale,#eff6ff);border-top:2px solid var(--ry-blue-dark)">
          <tr>
            <td style="padding:6px 10px;font-weight:800;color:var(--ry-blue-dark)">合計</td>
            <td style="padding:6px 10px;text-align:right;font-family:var(--f-mono);font-weight:800;color:var(--ry-blue)">${totalPicks.toLocaleString()}</td>
          </tr>
        </tfoot>
      </table>
    </div>`;
  document.getElementById('picks-preview').style.display = 'block';
}

function applyPicks() {
  if (!parsedPicks) return;
  PICKS_RAW = parsedPicks.records;
  if (currentPageId === 'picks') renderPicksPage();
  updateStatus();
  toast('✅ 揀次資料已套用！切換至「揀次分析」頁查看');
}

function resetPicks() {
  parsedPicks = null;
  document.getElementById('picks-import-status').textContent = '尚未上傳';
  document.getElementById('picks-preview').style.display = 'none';
  document.getElementById('picks-btns').style.display = 'none';
  document.getElementById('picks-file').value = '';
}

function updateStatus() {
  document.getElementById('status-time').textContent = '更新：' + new Date().toLocaleString('zh-TW');
  let daxiF=0, daduF=0, gsF=0, daxiL=0, daduL=0, gsL=0;
  DATA.dispatch.daily.forEach(r => { daxiL+=r[1]; daxiF+=r[2]; daduL+=r[3]; daduF+=r[4]; gsL+=r[5]; gsF+=r[6]; });

  const src = (typeof PICKS_RAW !== 'undefined') ? PICKS_RAW : [];
  let daxiP=0, daduP=0, gsP=0;
  src.forEach(r => {
    if (r.wh === '大溪倉') daxiP += r.picks;
    else if (r.wh === '大肚倉') daduP += r.picks;
    else if (r.wh === '岡山倉') gsP   += r.picks;
  });

  const rows = [
    { type:'🚚 運務費用', real:!!parsedFreight, daxi:daxiF,  dadu:daduF,  gs:gsF,  unit:'$' },
    { type:'💰 人力費用', real:!!parsedLabor,   daxi:daxiL,  dadu:daduL,  gs:gsL,  unit:'$' },
    { type:'⚡ 揀次資料', real:!!parsedPicks,   daxi:daxiP,  dadu:daduP,  gs:gsP,  unit:''  },
  ];
  document.getElementById('status-tbody').innerHTML = rows.map(r => {
    const total = r.daxi + r.dadu + r.gs;
    const c = r.real ? '#1b7c33' : '#e07855';
    const s = r.real ? '✅ 已上傳' : '⚠️ 範例資料';
    const fmt = v => r.unit === '$' ? '$' + v.toLocaleString() : v.toLocaleString();
    return `<tr>
      <td style="font-weight:700">${r.type}</td>
      <td><span style="color:${c};font-weight:700">${s}</span></td>
      <td class="mono" style="text-align:right;color:#0E7BAD">${fmt(r.daxi)}</td>
      <td class="mono" style="text-align:right;color:#2DA870">${fmt(r.dadu)}</td>
      <td class="mono" style="text-align:right;color:#E07855">${fmt(r.gs)}</td>
      <td class="mono" style="text-align:right;font-weight:800">${fmt(total)}</td>
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

// ════════════════════════════════════════════
// Picks Page 揀次分析
// ════════════════════════════════════════════
function initPicksPage() { renderPicksPage(); }

function renderPicksPage() {
  const wh = document.getElementById('picks-wh')?.value || '';
  const op = document.getElementById('picks-op')?.value || '';

  let data = (typeof PICKS_RAW !== 'undefined') ? PICKS_RAW : [];
  if (wh) data = data.filter(r => r.wh === wh);
  if (op) data = data.filter(r => r.op === op);

  const total = data.reduce((s, r) => s + r.picks, 0);

  const byDate = {};
  data.forEach(r => { byDate[r.date] = (byDate[r.date] || 0) + r.picks; });
  const dates = Object.keys(byDate).sort();
  const dailyVals = dates.map(d => byDate[d]);
  const avgDaily = dates.length ? Math.round(total / dates.length) : 0;
  const peakIdx = dailyVals.length ? dailyVals.indexOf(Math.max(...dailyVals)) : -1;
  const peakDate = peakIdx >= 0 ? dates[peakIdx] : '';
  const peakVal  = peakIdx >= 0 ? dailyVals[peakIdx] : 0;

  const whTotals = { '大肚倉': 0, '大溪倉': 0, '岡山倉': 0 };
  data.forEach(r => { whTotals[r.wh] = (whTotals[r.wh] || 0) + r.picks; });

  const byOp = {};
  data.forEach(r => {
    if (!byOp[r.op]) byOp[r.op] = { total: 0, wh: {} };
    byOp[r.op].total += r.picks;
    byOp[r.op].wh[r.wh] = (byOp[r.op].wh[r.wh] || 0) + r.picks;
  });
  const ops = Object.keys(byOp).sort((a, b) => byOp[b].total - byOp[a].total);
  const maxOpVal = ops.length ? byOp[ops[0]].total : 1;

  const maxDay = Math.max(...dailyVals, 1);
  let trendHtml = '<div style="display:flex;gap:2px;align-items:flex-end;height:80px;padding:12px 16px 8px;overflow-x:auto">';
  dates.forEach((d, i) => {
    const h = Math.round(dailyVals[i] / maxDay * 64);
    const day = d.slice(8);
    const isPeak = i === peakIdx;
    trendHtml += `<div style="flex:1;min-width:14px;display:flex;flex-direction:column;align-items:center;gap:2px" title="${d}: ${dailyVals[i].toLocaleString()}">
      <div style="width:100%;background:${isPeak ? 'var(--ry-gold)' : 'var(--ry-blue)'};height:${Math.max(h, 2)}px;border-radius:2px 2px 0 0"></div>
      <div style="font-size:9px;color:var(--ry-muted)">${day}</div>
    </div>`;
  });
  trendHtml += '</div>';

  let opRows = ops.map(o => {
    const item = byOp[o];
    const pct  = total ? (item.total / total * 100).toFixed(1) : '0.0';
    const barW = (item.total / maxOpVal * 100).toFixed(1);
    return `<tr>
      <td>${o}</td>
      <td style="text-align:right;font-family:var(--f-mono)">${(item.wh['大肚倉'] || 0).toLocaleString()}</td>
      <td style="text-align:right;font-family:var(--f-mono)">${(item.wh['大溪倉'] || 0).toLocaleString()}</td>
      <td style="text-align:right;font-family:var(--f-mono)">${(item.wh['岡山倉'] || 0).toLocaleString()}</td>
      <td style="text-align:right;font-weight:700;font-family:var(--f-mono)">${item.total.toLocaleString()}</td>
      <td style="min-width:130px">
        <div style="background:var(--ry-line);border-radius:2px;height:8px;margin-bottom:2px">
          <div style="background:var(--ry-blue);height:8px;border-radius:2px;width:${barW}%"></div>
        </div>
        <span style="font-size:10px;color:var(--ry-muted);font-family:var(--f-mono)">${pct}%</span>
      </td>
    </tr>`;
  }).join('');
  opRows += `<tr style="font-weight:700;border-top:2px solid var(--ry-line)">
    <td>合計</td>
    <td style="text-align:right;font-family:var(--f-mono)">${whTotals['大肚倉'].toLocaleString()}</td>
    <td style="text-align:right;font-family:var(--f-mono)">${whTotals['大溪倉'].toLocaleString()}</td>
    <td style="text-align:right;font-family:var(--f-mono)">${whTotals['岡山倉'].toLocaleString()}</td>
    <td style="text-align:right;font-family:var(--f-mono)">${total.toLocaleString()}</td>
    <td></td>
  </tr>`;

  document.getElementById('picks-grid').innerHTML = `
  <div class="w s4">
    <div class="gold-band">TOTAL</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>月總揀次</div></div>
    <div style="padding:20px 16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-blue);line-height:1;margin-bottom:6px">${total.toLocaleString()}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted)">三倉合計 · 2026年3月</div>
    </div>
  </div>
  <div class="w s4">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">DAILY AVG</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>日均揀次</div></div>
    <div style="padding:20px 16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-ink);line-height:1;margin-bottom:6px">${avgDaily.toLocaleString()}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted)">每日平均（${dates.length} 天）</div>
    </div>
  </div>
  <div class="w s4">
    <div class="gold-band" style="background:var(--ry-red);color:white">PEAK DAY</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-red)"></div>峰值日</div></div>
    <div style="padding:20px 16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-red);line-height:1;margin-bottom:6px">${peakDate.slice(5) || '—'}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted)">${peakVal.toLocaleString()} 揀次</div>
    </div>
  </div>
  <div class="w s12">
    <div class="gold-band">📅 每日揀次趨勢</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>每日三倉合計揀次（2026/03）</div><span class="wmeta">金色=峰值</span></div>
    ${trendHtml}
  </div>
  <div class="w s12">
    <div class="gold-band">📊 作業區域分析</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各作業區域揀次量 × 三倉</div><span class="wmeta">單位：次</span></div>
    <table class="tbl">
      <thead><tr><th>作業區域</th><th style="text-align:right">大肚倉</th><th style="text-align:right">大溪倉</th><th style="text-align:right">岡山倉</th><th style="text-align:right">合計</th><th>佔比</th></tr></thead>
      <tbody>${opRows}</tbody>
    </table>
  </div>`;

  const meta = document.getElementById('picks-meta');
  if (meta) meta.textContent = `資料：2026年3月 · 三倉 · ${data.length.toLocaleString()} 筆作業記錄`;
}

// ════════════════════════════════════════════
// Labor Page 工時結構分析
// ════════════════════════════════════════════
function initLaborPage() { renderLaborPage(); }

function renderLaborPage() {
  const shiftFilter  = document.getElementById('labor-shift')?.value  || '';
  const vendorFilter = document.getElementById('labor-vendor')?.value || '';

  let data = (typeof LABOR_RAW !== 'undefined') ? LABOR_RAW : [];
  data = data.filter(r => r.opArea !== '午休時間' && r.hours > 0);
  if (shiftFilter)  data = data.filter(r => r.shift  === shiftFilter);
  if (vendorFilter) data = data.filter(r => r.vendor === vendorFilter);

  const totalHrs  = data.reduce((s, r) => s + r.hours, 0);
  const totalCost = data.reduce((s, r) => s + r.cost,  0);
  const avgRate   = totalHrs > 0 ? Math.round(totalCost / totalHrs) : 0;
  const empCount  = new Set(data.map(r => r.empId)).size;

  const byOp = {};
  data.forEach(r => {
    if (!byOp[r.opArea]) byOp[r.opArea] = { hrs: 0, cost: 0 };
    byOp[r.opArea].hrs  += r.hours;
    byOp[r.opArea].cost += r.cost;
  });
  const ops = Object.keys(byOp).sort((a, b) => byOp[b].hrs - byOp[a].hrs);
  const COLORS = ['#1e5ca8', '#f5c400', '#d9401b', '#2ea85a', '#e07855', '#5a6478'];

  let structHtml = '';
  ops.forEach((o, i) => {
    const pct  = totalHrs > 0 ? (byOp[o].hrs / totalHrs * 100) : 0;
    const rate = byOp[o].hrs > 0 ? Math.round(byOp[o].cost / byOp[o].hrs) : 0;
    structHtml += `
    <div style="margin-bottom:14px">
      <div style="display:flex;justify-content:space-between;margin-bottom:4px">
        <span style="font-size:var(--fs-sm);font-weight:700">${o}</span>
        <span style="font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono)">${byOp[o].hrs.toFixed(1)}h · $${rate}/h · ${pct.toFixed(1)}%</span>
      </div>
      <div style="background:var(--ry-line);border-radius:3px;height:14px;overflow:hidden">
        <div style="background:${COLORS[i % COLORS.length]};height:14px;border-radius:3px;width:${pct.toFixed(1)}%;transition:width .4s"></div>
      </div>
    </div>`;
  });

  const byShift = {};
  data.forEach(r => {
    if (!byShift[r.shift]) byShift[r.shift] = { hrs: 0, cost: 0 };
    byShift[r.shift].hrs  += r.hours;
    byShift[r.shift].cost += r.cost;
  });
  const shiftRows = ['日', '中', '夜'].filter(s => byShift[s]).map(s => {
    const it   = byShift[s];
    const pct  = totalHrs > 0 ? (it.hrs / totalHrs * 100) : 0;
    const rate = it.hrs > 0 ? Math.round(it.cost / it.hrs) : 0;
    return `<tr>
      <td><b>${s}班</b></td>
      <td style="text-align:right;font-family:var(--f-mono)">${it.hrs.toFixed(1)}</td>
      <td style="text-align:right;font-family:var(--f-mono)">$${it.cost.toLocaleString()}</td>
      <td style="text-align:right;font-family:var(--f-mono)">$${rate}</td>
      <td style="min-width:80px"><div style="background:var(--ry-line);height:8px;border-radius:2px">
        <div style="background:var(--ry-blue);height:8px;border-radius:2px;width:${pct.toFixed(1)}%"></div>
      </div></td>
    </tr>`;
  }).join('');

  const byVendor = {};
  data.forEach(r => {
    if (!byVendor[r.vendor]) byVendor[r.vendor] = { hrs: 0, cost: 0, emps: new Set() };
    byVendor[r.vendor].hrs  += r.hours;
    byVendor[r.vendor].cost += r.cost;
    byVendor[r.vendor].emps.add(r.empId);
  });
  const vendorRows = Object.entries(byVendor)
    .sort((a, b) => b[1].hrs - a[1].hrs)
    .map(([v, it]) => {
      const rate = it.hrs > 0 ? Math.round(it.cost / it.hrs) : 0;
      return `<tr>
        <td>${v}</td>
        <td style="text-align:right">${it.emps.size}</td>
        <td style="text-align:right;font-family:var(--f-mono)">${it.hrs.toFixed(1)}</td>
        <td style="text-align:right;font-family:var(--f-mono)">$${it.cost.toLocaleString()}</td>
        <td style="text-align:right;font-family:var(--f-mono)">$${rate}</td>
      </tr>`;
    }).join('');

  const fmeta = document.getElementById('labor-filter-meta');
  if (fmeta) fmeta.textContent = `${data.length} 筆工時記錄`;

  document.getElementById('labor-grid').innerHTML = `
  <div class="w s3">
    <div class="gold-band">HOURS</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>總工時</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-blue);line-height:1">${totalHrs.toFixed(1)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">小時</div>
    </div>
  </div>
  <div class="w s3">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">COST</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>總費用</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">$${totalCost.toLocaleString()}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">元</div>
    </div>
  </div>
  <div class="w s3">
    <div class="gold-band" style="background:#2ea85a;color:white">RATE</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:#2ea85a"></div>平均時薪</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:#2ea85a;line-height:1">$${avgRate}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">元/小時</div>
    </div>
  </div>
  <div class="w s3">
    <div class="gold-band" style="background:var(--ry-muted);color:white">PEOPLE</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-muted)"></div>出勤人次</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">${empCount}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">人</div>
    </div>
  </div>
  <div class="w s6">
    <div class="gold-band">⚡ 工時結構 · 作業區域</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各作業區域工時佔比</div><span class="wmeta">總 ${totalHrs.toFixed(1)} h</span></div>
    <div style="padding:16px">${structHtml || '<div style="color:var(--ry-muted);padding:16px;text-align:center">無資料</div>'}</div>
  </div>
  <div class="w s6">
    <div class="gold-band">🌙 班別工時分析</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>日班 / 中班 / 夜班</div></div>
    <table class="tbl">
      <thead><tr><th>班別</th><th style="text-align:right">工時(h)</th><th style="text-align:right">費用</th><th style="text-align:right">時薪</th><th>佔比</th></tr></thead>
      <tbody>${shiftRows || '<tr><td colspan="5" style="text-align:center;color:var(--ry-muted)">無資料</td></tr>'}</tbody>
    </table>
  </div>
  <div class="w s12">
    <div class="gold-band">🏢 廠商工時彙總</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各派遣廠商工時與費用</div></div>
    <table class="tbl">
      <thead><tr><th>廠商</th><th style="text-align:right">人次</th><th style="text-align:right">工時(h)</th><th style="text-align:right">費用</th><th style="text-align:right">時薪</th></tr></thead>
      <tbody>${vendorRows || '<tr><td colspan="5" style="text-align:center;color:var(--ry-muted)">無資料</td></tr>'}</tbody>
    </table>
  </div>`;

  const meta = document.getElementById('labor-meta');
  if (meta) meta.textContent = `資料：2026年3月 · 全區各課 · ${data.length} 筆工時記錄`;
}

// ════════════════════════════════════════════
// Org Page 組織設定
// ════════════════════════════════════════════
let orgEditWh   = -1;
let orgEditDept = -1;

const ORG_COLORS = [
  '#1e5ca8','#2ea85a','#d9401b','#6366f1',
  '#f59e0b','#7c3aed','#0ea5e9','#e07855',
  '#14b8a6','#ec4899','#9ca3af','#f5c400',
];
const TYPE_COLORS = { '服務EC': '#1e5ca8', '營收EC': '#e07855', '後勤支援': '#2ea85a' };

function initOrgPage() {
  orgEditWh = -1;
  orgEditDept = -1;
  renderOrgPage();
}

function renderOrgPage() {
  const { warehouses, depts } = DATA.org;

  const whSwatches = (i, sel) => ORG_COLORS.map(c =>
    `<div class="org-color-swatch ${c===sel?'selected':''}" style="background:${c}" onclick="setOrgWhColor(${i},'${c}')"></div>`
  ).join('');
  const deptSwatches = (i, sel) => ORG_COLORS.map(c =>
    `<div class="org-color-swatch ${c===sel?'selected':''}" style="background:${c}" onclick="setOrgDeptColor(${i},'${c}')"></div>`
  ).join('');

  const whRows = warehouses.map((w, i) => {
    if (orgEditWh === i) {
      return `<div class="org-row editing">
        <div class="org-row-main" style="flex-wrap:wrap;gap:8px">
          <div class="org-color-picker">${whSwatches(i, w.color)}</div>
          <input class="filter-input" style="flex:1;min-width:90px" id="owh-n-${i}" value="${w.name}" placeholder="倉別名稱">
          <input class="filter-input" style="flex:2;min-width:120px" id="owh-r-${i}" value="${w.region}" placeholder="所屬部門">
          <button class="btn btn-ghost" style="font-size:11px;color:var(--ry-red)" onclick="deleteOrgWh(${i})">刪除</button>
          <button class="btn btn-primary" style="font-size:11px" onclick="saveOrgWhRow(${i})">確認</button>
        </div>
      </div>`;
    }
    return `<div class="org-row" onclick="editOrgWh(${i})">
      <div class="org-row-main">
        <div class="org-color-dot" style="background:${w.color}"></div>
        <div class="org-row-info">
          <div class="org-row-name">${w.name}</div>
          <div class="org-row-sub">${w.region}</div>
        </div>
        <div class="org-row-arrow">›</div>
      </div>
    </div>`;
  }).join('');

  const deptRows = depts.map((d, i) => {
    const tc = TYPE_COLORS[d.type] || '#5a6478';
    if (orgEditDept === i) {
      const whOpts = warehouses.map(w =>
        `<option ${w.name === d.wh ? 'selected' : ''}>${w.name}</option>`).join('');
      const typeOpts = ['服務EC','營收EC','後勤支援'].map(t =>
        `<option ${t === d.type ? 'selected' : ''}>${t}</option>`).join('');
      return `<div class="org-row editing">
        <div class="org-row-main" style="flex-wrap:wrap;gap:8px">
          <div class="org-color-picker">${deptSwatches(i, d.color)}</div>
          <input class="filter-input" style="flex:1;min-width:90px" id="odept-n-${i}" value="${d.name}" placeholder="課別名稱">
          <select class="filter-input" style="flex:1;min-width:80px" id="odept-t-${i}">${typeOpts}</select>
          <select class="filter-input" style="flex:1;min-width:80px" id="odept-w-${i}">${whOpts}</select>
          <button class="btn btn-ghost" style="font-size:11px;color:var(--ry-red)" onclick="deleteOrgDept(${i})">刪除</button>
          <button class="btn btn-primary" style="font-size:11px" onclick="saveOrgDeptRow(${i})">確認</button>
        </div>
      </div>`;
    }
    return `<div class="org-row" onclick="editOrgDept(${i})">
      <div class="org-row-main">
        <div class="org-color-dot" style="background:${d.color}"></div>
        <div class="org-row-info">
          <div class="org-row-name">${d.name} <span class="org-type-tag" style="background:${tc}20;color:${tc}">● ${d.type}</span></div>
          <div class="org-row-sub">${d.wh}</div>
        </div>
        <div class="org-row-arrow">›</div>
      </div>
    </div>`;
  }).join('');

  document.getElementById('org-grid').innerHTML = `
  <div class="w s6">
    <div class="org-card">
      <div class="org-card-head">
        <div>
          <div class="org-card-title">倉別</div>
          <div class="org-card-count">${warehouses.length} 個倉別</div>
        </div>
        <button class="btn btn-outline" style="font-size:12px" onclick="addOrgWh()">+ 新增倉別</button>
      </div>
      ${whRows}
    </div>
  </div>
  <div class="w s6">
    <div class="org-card">
      <div class="org-card-head">
        <div>
          <div class="org-card-title">課別</div>
          <div class="org-card-count">${depts.length} 個課別</div>
        </div>
        <button class="btn btn-outline" style="font-size:12px" onclick="addOrgDept()">+ 新增課別</button>
      </div>
      ${deptRows}
    </div>
  </div>`;
}

function editOrgWh(i)   { orgEditWh = i; orgEditDept = -1; renderOrgPage(); }
function editOrgDept(i) { orgEditDept = i; orgEditWh = -1; renderOrgPage(); }

function setOrgWhColor(i, color) {
  DATA.org.warehouses[i].color = color;
  orgEditWh = i;
  renderOrgPage();
}
function setOrgDeptColor(i, color) {
  DATA.org.depts[i].color = color;
  orgEditDept = i;
  renderOrgPage();
}

function saveOrgWhRow(i) {
  const n = document.getElementById(`owh-n-${i}`)?.value.trim();
  const r = document.getElementById(`owh-r-${i}`)?.value.trim();
  if (n) DATA.org.warehouses[i].name   = n;
  if (r) DATA.org.warehouses[i].region = r;
  orgEditWh = -1;
  renderOrgPage();
}
function saveOrgDeptRow(i) {
  const n = document.getElementById(`odept-n-${i}`)?.value.trim();
  const t = document.getElementById(`odept-t-${i}`)?.value;
  const w = document.getElementById(`odept-w-${i}`)?.value;
  if (n) DATA.org.depts[i].name = n;
  if (t) DATA.org.depts[i].type = t;
  if (w) DATA.org.depts[i].wh   = w;
  orgEditDept = -1;
  renderOrgPage();
}

function deleteOrgWh(i) {
  if (!confirm(`確定刪除「${DATA.org.warehouses[i].name}」？`)) return;
  DATA.org.warehouses.splice(i, 1);
  orgEditWh = -1;
  renderOrgPage();
}
function deleteOrgDept(i) {
  if (!confirm(`確定刪除「${DATA.org.depts[i].name}」？`)) return;
  DATA.org.depts.splice(i, 1);
  orgEditDept = -1;
  renderOrgPage();
}

function addOrgWh() {
  DATA.org.warehouses.push({ name: '新倉別', region: '請輸入部門', color: '#9ca3af' });
  orgEditWh   = DATA.org.warehouses.length - 1;
  orgEditDept = -1;
  renderOrgPage();
}
function addOrgDept() {
  const wh = DATA.org.warehouses[0]?.name || '';
  DATA.org.depts.push({ name: '新課別', type: '服務EC', wh, color: '#9ca3af' });
  orgEditDept = DATA.org.depts.length - 1;
  orgEditWh   = -1;
  renderOrgPage();
}

function saveOrgSettings() {
  orgEditWh = -1;
  orgEditDept = -1;
  renderOrgPage();
  toast('✅ 組織設定已儲存');
}
