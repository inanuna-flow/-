// ═══════════════════════════════════════════════════════
// app.js · 頁面設定 / 導覽 / 初始化
// ═══════════════════════════════════════════════════════

// ── 超級管理員設定 ──
const ADMIN_USER_ID = 'inari';
let pagePermissions = {}; // 從伺服器載入

function isAdmin() {
  return (sessionStorage.getItem('kpl_user') || '').toLowerCase() === ADMIN_USER_ID;
}

async function loadPagePermissions() {
  try {
    const res = await fetch('/api/page-permissions');
    if (res.ok) pagePermissions = await res.json();
  } catch { /* 若載入失敗，管理員帳號仍可看全部 */ }
}

function isPageVisible(pageId) {
  if (isAdmin()) return true; // 管理員永遠看得到所有頁面
  return pagePermissions[pageId] !== false;
}

const PAGES = [
  {
    group: '📊 儀表板',
    items: [
      { id:'daily',        icon:'📅', label:'每日動支監控', status:'ready' },
      { id:'dispatch',     icon:'💼', label:'總費用動支率', status:'ready' },
      { id:'freight',      icon:'🚚', label:'運費損益分析', status:'ready' },
      { id:'labor',        icon:'⏱', label:'人力工時結構', status:'ready' },
      { id:'productivity', icon:'📊', label:'揀次人效分析', status:'ready' },
      { id:'annual',       icon:'📋', label:'年度規劃分析', status:'ready' },
    ]
  },
  {
    group: '📁 資料管理',
    items: [
      { id:'import',       icon:'📤', label:'資料匯入',     status:'ready' },
      { id:'org',          icon:'🏢', label:'組織設定',     status:'ready' },
    ]
  },
];

let currentPageId = 'daily';
let annualViewMode = 'labor';

// ── 渲染左側選單 ──
function renderSidebar() {
  const sb = document.getElementById('sidebar');
  let html = '<div class="sb-nav">';
  PAGES.forEach(group => {
    // 過濾出此群組中有權限的頁面
    const visibleItems = group.items.filter(item => isPageVisible(item.id));
    if (visibleItems.length === 0) return;

    html += `<div class="sb-group"><div class="sb-group-label">${group.group}</div>`;
    visibleItems.forEach(item => {
      const active = item.id === currentPageId ? 'active' : '';
      let badge = '';
      if (item.status === 'wip') badge = '<span class="sb-item-badge wip">WIP</span>';
      else if (item.status === 'placeholder') badge = '<span class="sb-item-badge soon">TBD</span>';
      // 管理員可看到被關閉的頁面但顯示灰色標記
      const hiddenMark = (isAdmin() && pagePermissions[item.id] === false)
        ? '<span class="sb-item-badge" style="background:#aaa;color:#fff">隱藏</span>' : '';
      html += `
      <a href="#${item.id}" class="sb-item ${active}" onclick="navigate(event, '${item.id}')">
        <span class="sb-item-icon">${item.icon}</span>
        <span class="sb-item-label">${item.label}</span>
        ${badge}${hiddenMark}
      </a>`;
    });
    html += '</div>';
  });

  // 管理員專屬：權限設定入口
  if (isAdmin()) {
    const adminActive = currentPageId === 'admin' ? 'active' : '';
    html += `
    <div class="sb-group">
      <div class="sb-group-label" style="color:#f5c400">⚙️ 超級管理員</div>
      <a href="#admin" class="sb-item ${adminActive}" onclick="navigate(event, 'admin')" style="border-left: 3px solid #f5c400">
        <span class="sb-item-icon">🔐</span>
        <span class="sb-item-label">頁面權限設定</span>
      </a>
    </div>`;
  }

  html += '</div>';

  // 底部使用者資訊 + 登出
  const userId = sessionStorage.getItem('kpl_user') || '使用者';
  const adminBadge = isAdmin()
    ? '<span style="font-size:10px;background:#f5c400;color:#123d74;padding:1px 6px;border-radius:99px;font-weight:700;margin-left:4px">管理員</span>'
    : '';
  html += `
  <div class="sb-user">
    <div class="sb-user-avatar">${isAdmin() ? '👑' : '👤'}</div>
    <div class="sb-user-info">
      <div class="sb-user-name">${userId}${adminBadge}</div>
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
  // 管理員頁面不需要 template
  if (pageId === 'admin') {
    if (!isAdmin()) { navigate(null, 'daily'); return; }
    document.getElementById('main').innerHTML = '';
    renderAdminPage();
    return;
  }

  // 非管理員：檢查頁面是否有權限
  if (!isAdmin() && pagePermissions[pageId] === false) {
    navigate(null, 'daily');
    return;
  }

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
  else if (pageId === 'labor')   initLaborPage();
  else if (pageId === 'import')  initImportPage();
  else if (pageId === 'org')     initOrgPage();
  else if (pageId === 'productivity') initProductivityPage();
  else if (pageId === 'annual')  renderAnnualPage();
  else if (pageId === 'admin')   renderAdminPage();
}

// ── 各頁面初始化 ──
function initDailyPage() { renderDailyPage(); }

function initDispatchPage() { renderDispatchPage(); }
function initFreightPage()  { renderFreightPage(); }
function initImportPage()   { if (typeof updateStatus === 'function') updateStatus(); }

const DASHBOARD_DATE_FILTERS = {
  daily:        { from:'filter-from',        to:'filter-to',        meta:'filter-meta' },
  dispatch:     { from:'dispatch-from',     to:'dispatch-to',     meta:null },
  freight:      { from:'freight-from',      to:'freight-to',      meta:null },
  picks:        { from:'picks-from',        to:'picks-to',        meta:'picks-date-meta' },
  labor:        { from:'labor-from',        to:'labor-to',        meta:'labor-date-meta' },
  productivity: { from:'productivity-from', to:'productivity-to', meta:'productivity-date-meta' },
};

function syncDashboardDateInputs(pageId = currentPageId) {
  const cfg = DASHBOARD_DATE_FILTERS[pageId];
  if (!cfg) return;
  const from = document.getElementById(cfg.from);
  const to = document.getElementById(cfg.to);
  if (from) from.value = DATA.dateFrom;
  if (to) to.value = DATA.dateTo;
  if (cfg.meta) {
    const meta = document.getElementById(cfg.meta);
    if (meta) meta.textContent = `🔗 與其他頁共用日期設定`;
  }
}

function setSharedDateRangeFromInputs(pageId = currentPageId) {
  const cfg = DASHBOARD_DATE_FILTERS[pageId];
  if (!cfg) return false;
  const from = document.getElementById(cfg.from)?.value || DATA.dateFrom;
  const to = document.getElementById(cfg.to)?.value || DATA.dateTo;
  if (from && to && from > to) {
    toast('❌ 起始日期不可晚於結束日期');
    return false;
  }
  if (from) DATA.dateFrom = from;
  if (to) DATA.dateTo = to;
  return true;
}

function dateInSelectedRange(dateStr) {
  if (!dateStr) return false;
  return (!DATA.dateFrom || dateStr >= DATA.dateFrom) && (!DATA.dateTo || dateStr <= DATA.dateTo);
}

// ════════════════════════════════════════════
// 🔐 管理員頁面：頁面權限設定
// ════════════════════════════════════════════
function renderAdminPage() {
  if (!isAdmin()) return;
  const main = document.getElementById('main');

  const allItems = PAGES.flatMap(g => g.items);
  const rows = allItems.map(item => {
    const isOn = pagePermissions[item.id] !== false;
    return `
    <tr>
      <td style="padding:12px 16px;font-size:15px">${item.icon} ${item.label}</td>
      <td style="padding:12px 16px;color:#888;font-size:12px;font-family:monospace">${item.id}</td>
      <td style="padding:12px 16px;text-align:center">
        <label class="admin-toggle">
          <input type="checkbox" id="perm-${item.id}" ${isOn ? 'checked' : ''}>
          <span class="admin-toggle-slider"></span>
        </label>
      </td>
    </tr>`;
  }).join('');

  main.innerHTML = `
  <div style="max-width:680px;margin:32px auto;padding:0 16px">
    <div style="margin-bottom:24px">
      <div style="font-size:22px;font-weight:900;color:#1a1d24">🔐 頁面權限設定</div>
      <div style="font-size:13px;color:#888;margin-top:4px">控制一般帳號可以看到哪些頁面。管理員帳號（inari）永遠可以看到全部頁面。</div>
    </div>

    <div style="background:#fff;border-radius:8px;box-shadow:0 2px 12px rgba(0,0,0,.08);overflow:hidden;border:1px solid #eee">
      <table style="width:100%;border-collapse:collapse">
        <thead>
          <tr style="background:#f4f6fb;border-bottom:2px solid #e8eaf0">
            <th style="padding:10px 16px;text-align:left;font-size:12px;color:#666;font-weight:700">頁面名稱</th>
            <th style="padding:10px 16px;text-align:left;font-size:12px;color:#666;font-weight:700">頁面 ID</th>
            <th style="padding:10px 16px;text-align:center;font-size:12px;color:#666;font-weight:700">一般帳號可見</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>

    <div style="margin-top:20px;padding:16px;background:#fff8e6;border-radius:8px;border:1px solid #f5c400;font-size:13px;color:#7a6000">
      ⚠️ 儲存前需輸入您的密碼重新驗證身份，確保安全。
    </div>

    <div id="admin-save-area" style="margin-top:20px">
      <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap">
        <div style="flex:1;min-width:200px">
          <label style="display:block;font-size:11px;font-weight:700;color:#666;margin-bottom:6px;letter-spacing:.08em">確認密碼</label>
          <input type="password" id="admin-psw" placeholder="請輸入您的登入密碼"
            style="width:100%;padding:10px 14px;border:1.5px solid #dde2ec;border-radius:4px;font-size:13px">
        </div>
        <button onclick="savePermissions()"
          style="padding:10px 28px;background:#1e5ca8;color:#fff;border:none;border-radius:4px;font-size:14px;font-weight:700;cursor:pointer;white-space:nowrap">
          💾 儲存設定
        </button>
      </div>
      <div id="admin-msg" style="margin-top:10px;font-size:13px;display:none"></div>
    </div>
  </div>

  <style>
  .admin-toggle { position:relative;display:inline-block;width:44px;height:24px; }
  .admin-toggle input { opacity:0;width:0;height:0; }
  .admin-toggle-slider {
    position:absolute;inset:0;background:#ccc;border-radius:24px;cursor:pointer;transition:.2s;
  }
  .admin-toggle-slider:before {
    content:'';position:absolute;width:18px;height:18px;left:3px;bottom:3px;
    background:#fff;border-radius:50%;transition:.2s;
  }
  .admin-toggle input:checked + .admin-toggle-slider { background:#1e5ca8; }
  .admin-toggle input:checked + .admin-toggle-slider:before { transform:translateX(20px); }
  </style>`;
}

async function savePermissions() {
  const pswEl = document.getElementById('admin-psw');
  const msgEl = document.getElementById('admin-msg');
  const psw = pswEl?.value || '';
  if (!psw) {
    showAdminMsg('❌ 請輸入密碼', 'error');
    pswEl.focus();
    return;
  }

  // 收集所有開關狀態
  const allItems = PAGES.flatMap(g => g.items);
  const permissions = {};
  allItems.forEach(item => {
    const cb = document.getElementById(`perm-${item.id}`);
    permissions[item.id] = cb ? cb.checked : true;
  });

  const btn = document.querySelector('#admin-save-area button');
  btn.disabled = true;
  btn.textContent = '驗證中…';
  showAdminMsg('', '');

  try {
    const res = await fetch('/api/page-permissions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        USER_ID: sessionStorage.getItem('kpl_user'),
        PSW: psw,
        permissions,
      }),
    });
    const data = await res.json();
    if (res.ok && data.ok) {
      pagePermissions = data.permissions;
      pswEl.value = '';
      showAdminMsg('✅ 權限已儲存！其他帳號下次登入時生效。', 'ok');
      renderSidebar(); // 重新渲染側欄（顯示隱藏標記）
    } else {
      showAdminMsg(`❌ ${data.MSG || '儲存失敗'}`, 'error');
    }
  } catch (err) {
    showAdminMsg(`❌ 網路錯誤：${err.message}`, 'error');
  }

  btn.disabled = false;
  btn.innerHTML = '💾 儲存設定';
}

function showAdminMsg(text, type) {
  const el = document.getElementById('admin-msg');
  if (!el) return;
  el.style.display = text ? 'block' : 'none';
  el.style.color = type === 'error' ? '#d9401b' : '#1a7a3f';
  el.textContent = text;
}

function rerenderDashboardPage(pageId = currentPageId) {
  if (pageId === 'daily') renderDailyPage();
  else if (pageId === 'dispatch') renderDispatchPage();
  else if (pageId === 'freight') renderFreightPage();
  else if (pageId === 'labor') renderLaborPage();
  else if (pageId === 'productivity') renderProductivityPage();
  else if (pageId === 'annual') renderAnnualPage();
}

async function applyDashboardDateFilter(pageId = currentPageId) {
  if (!setSharedDateRangeFromInputs(pageId)) return;
  await loadCloudBudgetData();
  await loadCloudLaborData();
  if (pageId === 'dispatch') syncDispatchBudgetForCurrentMonth();
  rerenderDashboardPage(pageId);
  toast('🔄 日期區間已更新');
}

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

function getBudgetYear() {
  const year = Number(String(DATA.dateFrom || '').slice(0, 4));
  return Number.isInteger(year) && year >= 2000 && year <= 2100 ? year : new Date().getFullYear();
}

function applyCloudBudgetRows(rows, year = getBudgetYear()) {
  if (!Array.isArray(rows) || !rows.length) return false;
  const labor = createBudgetBuckets();
  const freight = createBudgetBuckets();
  let latest = '';
  let oldest = '';
  let count = 0;

  rows.forEach(row => {
    const wh = row.warehouseName;
    const type = row.costType;
    const amount = Number(row.budgetAmount) || 0;
    const monthText = String(row.month || '');
    const monthIndex = Number(monthText.slice(5, 7)) - 1;
    if (!wh || monthIndex < 0 || monthIndex > 11) return;
    if (type === 'labor' && labor[wh]) labor[wh][monthIndex] = amount;
    else if (type === 'freight' && freight[wh]) freight[wh][monthIndex] = amount;
    else return;
    const monthValue = `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
    latest = latest ? (monthValue > latest ? monthValue : latest) : monthValue;
    oldest = oldest ? (monthValue < oldest ? monthValue : oldest) : monthValue;
    count++;
  });

  if (!count) return false;
  DATA.annualBudget.labor = labor;
  DATA.annualBudget.freight = freight;
  DATA.dispatch.budget = buildDispatchBudget(labor, freight, getCurrentMonthIndex());
  DATA.dataLatest.budget = latest || `${getCurrentMonthIndex() + 1}月`;
  DATA.dataOldest = DATA.dataOldest || {};
  DATA.dataOldest.budget = oldest;
  return true;
}

async function loadCloudBudgetData() {
  try {
    const res = await fetch(`/api/data/budget?year=${getBudgetYear()}`);
    if (!res.ok) return false;
    const data = await res.json();
    return data.ok ? applyCloudBudgetRows(data.rows, data.year) : false;
  } catch {
    return false;
  }
}

async function syncCloudBudget(parsed) {
  if (!parsed) return false;
  const payload = {
    year: getBudgetYear(),
    monthIndex: parsed.monthIndex,
    labor: parsed.labor,
    freight: parsed.freight,
    fileName: parsed.fileName,
    version: parsed.version,
    importedBy: sessionStorage.getItem('kpl_user') || '',
  };
  const res = await fetch('/api/import/budget', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.ok) throw new Error(data.MSG || `HTTP ${res.status}`);
  return data;
}

function normalizeCloudDate(value) {
  return String(value || '').slice(0, 10);
}

function resetDispatchLaborForRange(from = DATA.dateFrom, to = DATA.dateTo) {
  DATA.dispatch.daily = DATA.dispatch.daily.map(row => {
    const fullDate = dispatchRowFullDate(row);
    if (!fullDate || fullDate < from || fullDate > to) return row;
    return [row[0], 0, row[2], 0, row[4], 0, row[6], row[7]];
  });
}

function applyCloudLaborRows(rows) {
  resetDispatchLaborForRange();
  if (!Array.isArray(rows) || !rows.length) {
    LABOR_RAW = [];
    DATA.dataLatest.labor = '';
    DATA.dataOldest = DATA.dataOldest || {};
    DATA.dataOldest.labor = '';
    return true;
  }
  LABOR_RAW = rows.map(row => ({
    wh: row.wh,
    date: normalizeCloudDate(row.date),
    vendor: row.vendor || '',
    shift: row.shift || '',
    empId: row.empId || '',
    dept: row.dept || '',
    name: '',
    opArea: row.opArea || '',
    hours: Number(row.hours) || 0,
    boxHours: 0,
    nightHrs: 0,
    normHrs: Number(row.hours) || 0,
    cost: Number(row.cost) || 0,
    check: '',
    sourceSheet: 'Cloud SQL',
  })).filter(row => row.date && row.wh && row.opArea);

  const kpiRows = getLaborKpiRecords(LABOR_RAW);
  const daily = summarizeLaborDaily(kpiRows);
  applyLaborToDispatch(daily);
  updateDispatchLatestUploadDate(daily.map(row => row.date));
  const dates = LABOR_RAW.map(row => row.date).filter(Boolean).sort();
  DATA.dataLatest.labor = dates[dates.length - 1] || '';
  DATA.dataOldest = DATA.dataOldest || {};
  DATA.dataOldest.labor = dates[0] || '';
  return true;
}

async function loadCloudLaborData() {
  try {
    const params = new URLSearchParams({
      date_from: DATA.dateFrom,
      date_to: DATA.dateTo,
    });
    const res = await fetch(`/api/data/labor?${params.toString()}`);
    if (!res.ok) return false;
    const data = await res.json();
    return data.ok ? applyCloudLaborRows(data.rows) : false;
  } catch {
    return false;
  }
}

async function syncCloudLabor(parsed) {
  if (!parsed?.records?.length) return false;
  const payload = {
    records: parsed.records,
    fileName: parsed.fileName,
    importedBy: sessionStorage.getItem('kpl_user') || '',
  };
  const res = await fetch('/api/import/labor', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.ok) throw new Error(data.MSG || `HTTP ${res.status}`);
  return data;
}

// ── 初始化 ──
document.addEventListener('DOMContentLoaded', async () => {
  // 驗證登入
  if (!sessionStorage.getItem('kpl_auth')) {
    location.href = 'login.html';
    return;
  }

  // 載入頁面權限（管理員也載入，用於顯示哪些頁面被隱藏）
  await loadPagePermissions();
  await loadCloudBudgetData();
  await loadCloudLaborData();

  renderSidebar();

  const hash = location.hash.slice(1);
  if (hash) {
    const allItems = PAGES.flatMap(g => g.items);
    if (allItems.find(i => i.id === hash)) {
      currentPageId = hash;
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
  applyDashboardDateFilter('daily');
}

// ════════════════════════════════════════════
// Dispatch Page 邏輯
// ════════════════════════════════════════════
function renderDispatchPage() {
  syncDispatchBudgetForCurrentMonth();
  const grid = document.getElementById('dispatch-grid');
  grid.innerHTML = [renderDispatchBudgetNotice(), renderT001(), renderT002(), renderT003()].join('');

  document.getElementById('dispatch-from').value = DATA.dateFrom;
  document.getElementById('dispatch-to').value   = DATA.dateTo;
  const days = getDispatchDailyFiltered().length;
  const latestDate = getDispatchLatestUploadDate();
  document.getElementById('dispatch-meta').textContent =
    `資料區間：${DATA.dateFrom} ~ ${DATA.dateTo} · 資料最新日期：${latestDate || '尚未匯入'} · 共 ${days} 天 · 含人力+運務總覽`;
}

function hasDispatchBudget() {
  if (typeof hasAnnualDispatchBudget === 'function' && hasAnnualDispatchBudget()) return true;
  return Object.values(DATA.dispatch.budget || {}).some(b => (b.labor || 0) > 0 || (b.freight || 0) > 0);
}

function renderDispatchBudgetNotice() {
  if (hasDispatchBudget()) return '';
  const rows = getDispatchDailyFiltered();
  const actual = rows.reduce((sum, row) => sum + row[1] + row[2] + row[3] + row[4] + row[5] + row[6], 0);
  if (!actual) return '';

  return `
  <div style="grid-column:1/-1;background:#fff4e8;border-left:4px solid var(--ry-orange);border-radius:4px;padding:12px 16px;color:var(--ry-ink);font-size:var(--fs-sm);line-height:1.7">
    <b>尚未套用年度預算。</b>
    目前已有實際費用 ${fmtMoney(actual)}，但總費用動支率需要「年度預算」作為分母；請先在資料匯入頁上傳年度預算並按「套用到儀表板」，動支率才會計算。
  </div>`;
}

function applyDispatchFilter() {
  applyDashboardDateFilter('dispatch');
}

// ════════════════════════════════════════════
// Freight Page 邏輯
// ════════════════════════════════════════════
function renderFreightPage() {
  const grid = document.getElementById('freight-grid');
  grid.innerHTML = [
    renderF001(),
    renderF002(),
    renderF003(),
    renderF009(),
    renderF010(),
  ].join('');

  document.getElementById('freight-from').value = DATA.dateFrom;
  document.getElementById('freight-to').value   = DATA.dateTo;
  const summary = typeof getFreightFilteredSummary === 'function'
    ? getFreightFilteredSummary()
    : { totalOrders: DATA.freight.totalOrders };
  const days = typeof getFreightTrendFiltered === 'function'
    ? getFreightTrendFiltered().length
    : DATA.freight.dailyTrend.length;
  document.getElementById('freight-meta').textContent =
    `資料區間：${DATA.dateFrom} ~ ${DATA.dateTo} · ${days} 天 · 共 ${summary.totalOrders.toLocaleString()} 筆配送`;
}

function applyFreightFilter() {
  applyDashboardDateFilter('freight');
}

// ════════════════════════════════════════════
// Import Page 邏輯
// ════════════════════════════════════════════
let parsedFreight = null;
let parsedLabor   = null;
let parsedPicks   = null;
let parsedBudget  = null;

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
      const detectedType = detectWorkbookType(wb);
      if (detectedType && detectedType !== type) {
        document.getElementById(type + '-status').textContent = `↪ 已改由${importTypeLabel(detectedType)}解析`;
        type = detectedType;
        document.getElementById(type + '-status').textContent = '解析中…';
      }
      if (type === 'freight') parseFreight(wb, file.name);
      else if (type === 'labor') parseLabor(wb, file.name);
      else if (type === 'picks') parsePicks(wb, file.name);
      else if (type === 'budget') parseBudget(wb, file.name);
    } catch(err) {
      document.getElementById(type + '-status').textContent = '❌ 解析失敗';
      toast('❌ ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function detectWorkbookType(wb) {
  const sheets = wb.SheetNames || [];
  if (
    (sheets.includes('人力') && sheets.includes('運務')) ||
    (sheets.includes('人力預算_轉換') && sheets.includes('運費預算_轉換'))
  ) return 'budget';
  return '';
}

function importTypeLabel(type) {
  return {
    budget: '年度預算',
    freight: '運務費用',
    labor: '工時資料',
    picks: '揀次資料',
  }[type] || '正確類型';
}

function parseBudget(wb, fileName) {
  const V2_LABOR_SHEET = '人力';
  const V2_FREIGHT_SHEET = '運務';
  const V1_LABOR_SHEET = '人力預算_轉換';
  const V1_FREIGHT_SHEET = '運費預算_轉換';
  const isV2 = wb.SheetNames.includes(V2_LABOR_SHEET) && wb.SheetNames.includes(V2_FREIGHT_SHEET);
  const isV1 = wb.SheetNames.includes(V1_LABOR_SHEET) && wb.SheetNames.includes(V1_FREIGHT_SHEET);
  if (!isV2 && !isV1) {
    toast('❌ 年度預算缺少必要分頁');
    document.getElementById('budget-status').textContent = '❌ 分頁不存在';
    return;
  }

  const labor = createBudgetBuckets();
  const freight = createBudgetBuckets();
  const warnings = [];

  if (isV2) {
    parseBudgetWideRows(XLSX.utils.sheet_to_json(wb.Sheets[V2_LABOR_SHEET], { defval:'' }), '單位', labor, '人力預算', warnings);
    parseBudgetWideRows(XLSX.utils.sheet_to_json(wb.Sheets[V2_FREIGHT_SHEET], { defval:'' }), '倉別', freight, '運務預算', warnings);
  } else {
    parseBudgetLongRows(XLSX.utils.sheet_to_json(wb.Sheets[V1_LABOR_SHEET], { defval:'' }), '區域', labor, '人力預算', warnings);
    parseBudgetLongRows(XLSX.utils.sheet_to_json(wb.Sheets[V1_FREIGHT_SHEET], { defval:'' }), '倉庫', freight, '運費預算', warnings);
  }

  const monthIndex = getCurrentMonthIndex();
  parsedBudget = {
    labor,
    freight,
    dispatchBudget: buildDispatchBudget(labor, freight, monthIndex),
    monthIndex,
    warnings,
    fileName,
    version: isV2 ? 'V2' : 'V1',
    at: new Date(),
  };
  document.getElementById('budget-status').textContent = `✅ ${parsedBudget.version} · ${monthIndex + 1}月預算`;
  showBudgetPreview(parsedBudget);
  document.getElementById('budget-btns').style.display = 'flex';
  toast(`✅ 年度預算${parsedBudget.version}解析完成：${monthIndex + 1}月`);
}

function createBudgetBuckets() {
  return { '大溪倉': Array(12).fill(0), '大肚倉': Array(12).fill(0), '岡山倉': Array(12).fill(0) };
}

function parseBudgetLongRows(rows, warehouseField, target, label, warnings) {
  rows.forEach((r, i) => {
    const wh = normalizeWarehouseName(r[warehouseField]);
    const monthIndex = parseBudgetMonth(r['月份']);
    const amount = parseMoney(r['金額']);
    if (!wh || monthIndex < 0 || Number.isNaN(amount)) {
      warnings.push(`${label}第 ${i + 2} 列無法辨識`);
      return;
    }
    target[wh][monthIndex] += amount;
  });
}

function parseBudgetWideRows(rows, warehouseField, target, label, warnings) {
  rows.forEach((r, i) => {
    const wh = normalizeBudgetWarehouseName(r[warehouseField], r);
    if (isBudgetTotalRow(r)) return;
    if (!wh && rowBudgetMonthTotal(r) === 0) return;
    if (!wh) {
      warnings.push(`${label}第 ${i + 2} 列無法辨識倉別`);
      return;
    }
    for (let monthIndex = 0; monthIndex < 12; monthIndex++) {
      const monthLabel = `${monthIndex + 1}月`;
      const amount = parseMoney(r[monthLabel]);
      if (Number.isNaN(amount)) {
        warnings.push(`${label}第 ${i + 2} 列 ${monthLabel} 金額不是數字`);
        continue;
      }
      target[wh][monthIndex] += amount;
    }
  });
}

function normalizeBudgetWarehouseName(value, row = {}) {
  const direct = normalizeWarehouseName(value);
  if (direct) return direct;

  const s = String(value || '').trim();
  const region = String(row['區別'] || '').trim();
  if (s.includes('大肚')) return '大肚倉';
  if (s.includes('岡山')) return '岡山倉';
  if (region === '北區' || ['理貨一課', '理貨二課', '倉儲管理課', '北區後勤', '日翊籍'].includes(s)) return '大溪倉';
  return '';
}

function isBudgetTotalRow(row) {
  const label = String(row['區別'] || row['單位'] || row['倉別'] || '').trim();
  return label === '合計';
}

function rowBudgetMonthTotal(row) {
  let total = 0;
  for (let monthIndex = 0; monthIndex < 12; monthIndex++) {
    const amount = parseMoney(row[`${monthIndex + 1}月`]);
    if (Number.isNaN(amount)) return NaN;
    total += amount;
  }
  return total;
}

function parseBudgetMonth(value) {
  const m = String(value || '').match(/(\d{1,2})/);
  if (!m) return -1;
  const n = Number(m[1]);
  return n >= 1 && n <= 12 ? n - 1 : -1;
}

function getCurrentMonthIndex() {
  const m = Number(String(DATA.dateFrom || '').slice(5, 7));
  return m >= 1 && m <= 12 ? m - 1 : 0;
}

function buildDispatchBudget(labor, freight, monthIndex) {
  return {
    '大溪倉': { labor: labor['大溪倉'][monthIndex] || 0, freight: freight['大溪倉'][monthIndex] || 0 },
    '大肚倉': { labor: labor['大肚倉'][monthIndex] || 0, freight: freight['大肚倉'][monthIndex] || 0 },
    '岡山倉': { labor: labor['岡山倉'][monthIndex] || 0, freight: freight['岡山倉'][monthIndex] || 0 },
  };
}

function syncDispatchBudgetForCurrentMonth() {
  if (!DATA.annualBudget?.labor || !DATA.annualBudget?.freight) return;
  DATA.dispatch.budget = buildDispatchBudget(DATA.annualBudget.labor, DATA.annualBudget.freight, getCurrentMonthIndex());
}

function showBudgetPreview(parsed) {
  const rows = Object.entries(parsed.dispatchBudget).map(([wh, b]) => `
    <tr>
      <td class="preview-cell preview-label">${wh}</td>
      <td class="preview-cell mono num-right">${fmtMoney(b.labor)}</td>
      <td class="preview-cell mono num-right">${fmtMoney(b.freight)}</td>
      <td class="preview-cell mono num-right preview-strong">${fmtMoney(b.labor + b.freight)}</td>
    </tr>`).join('');
  const total = Object.values(parsed.dispatchBudget).reduce((s, b) => s + b.labor + b.freight, 0);
  document.getElementById('budget-preview').innerHTML = `
    <div class="import-preview-title">
      📋 年度預算預覽（套用月份：${parsed.monthIndex + 1}月）
    </div>
    ${parsed.warnings.length ? `<div class="import-alert import-alert-warning">
      ⚠️ 警告：${parsed.warnings.length.toLocaleString()} 項（前 5 項：${parsed.warnings.slice(0, 5).join('；')}）
    </div>` : ''}
    <div class="preview-table-wrap">
      <table class="preview-table">
        <thead class="preview-thead">
          <tr>
            <th class="preview-head">倉別</th>
            <th class="preview-head preview-head-num">人力預算</th>
            <th class="preview-head preview-head-num">運費預算</th>
            <th class="preview-head preview-head-num">合計</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
        <tfoot class="preview-total">
          <tr>
            <td class="preview-cell preview-total-label">合計</td>
            <td colspan="3" class="preview-cell mono num-right preview-total-value">${fmtMoney(total)}</td>
          </tr>
        </tfoot>
      </table>
    </div>`;
  document.getElementById('budget-preview').style.display = 'block';
}

async function applyBudget() {
  if (!parsedBudget) return;
  DATA.annualBudget.labor = parsedBudget.labor;
  DATA.annualBudget.freight = parsedBudget.freight;
  DATA.dispatch.budget = parsedBudget.dispatchBudget;
  DATA.dataLatest.budget = `${parsedBudget.monthIndex + 1}月`;
  if (currentPageId === 'dispatch') renderDispatchPage();
  if (currentPageId === 'daily') renderDailyPage();
  if (currentPageId === 'annual') renderAnnualPage();
  updateStatus();
  toast('✅ 年度預算已套用！總費用動支率預算已更新');
  try {
    const result = await syncCloudBudget(parsedBudget);
    toast(`✅ 年度預算已同步雲端（${result.rows} 筆）`);
  } catch (err) {
    console.warn('Budget cloud sync failed:', err);
    toast(`⚠️ 年度預算已套用，但雲端同步失敗：${err.message}`);
  }
}

function resetBudget() {
  parsedBudget = null;
  document.getElementById('budget-status').textContent = '尚未上傳';
  document.getElementById('budget-preview').style.display = 'none';
  document.getElementById('budget-btns').style.display = 'none';
  document.getElementById('budget-file').value = '';
}

function parseFreight(wb, fileName) {
  const SUMMARY_SHEET = '進貨日與計價費用';
  const DETAIL_SHEET = '貨運費用明細總表';
  if (!wb.SheetNames.includes(SUMMARY_SHEET)) {
    toast('❌ 找不到「' + SUMMARY_SHEET + '」分頁');
    document.getElementById('freight-status').textContent = '❌ 分頁不存在';
    return;
  }
  if (!wb.SheetNames.includes(DETAIL_SHEET)) {
    toast('❌ 找不到「' + DETAIL_SHEET + '」分頁');
    document.getElementById('freight-status').textContent = '❌ 明細分頁不存在';
    return;
  }

  const errors = [];
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[SUMMARY_SHEET], { defval:0 });
  const rows = [];
  const seenDates = {};
  raw.forEach((r, i) => {
    const d = String(r['列標籤'] || '');
    if (!d || d.includes('總計')) return;
    const dateInfo = parseFreightDate(d);
    if (!dateInfo) {
      errors.push(`彙總分頁第 ${i + 2} 列日期無法辨識：${d}`);
      return;
    }
    const row = {
      date: dateInfo.short,
      fullDate: dateInfo.full,
      daxi: Number(r['大溪倉'] || 0),
      dadu: Number(r['大肚倉'] || 0),
      gangshan: Number(r['岡山倉'] || 0),
    };
    if ([row.daxi, row.dadu, row.gangshan].some(v => Number.isNaN(v))) {
      errors.push(`彙總分頁第 ${i + 2} 列金額不是數字`);
      return;
    }
    if ([row.daxi, row.dadu, row.gangshan].some(v => v < 0)) {
      errors.push(`彙總分頁第 ${i + 2} 列出現負數，依規則不可匯入`);
      return;
    }
    if (seenDates[row.fullDate]) {
      errors.push(`彙總分頁日期 ${row.fullDate} 重複出現，請先回 Excel 彙總成每日一列`);
      return;
    }
    seenDates[row.fullDate] = true;
    rows.push(row);
  });

  const detailResult = parseFreightDetails(wb.Sheets[DETAIL_SHEET]);
  errors.push(...detailResult.errors);

  if (errors.length) {
    document.getElementById('freight-status').textContent = '❌ 驗證失敗';
    document.getElementById('freight-preview').innerHTML = `
      <div class="import-alert import-alert-error">
        <b class="text-red">匯入已擋下</b><br>
        ${errors.slice(0, 8).map(e => `• ${e}`).join('<br>')}
        ${errors.length > 8 ? `<br>• 其餘 ${errors.length - 8} 項錯誤省略` : ''}
      </div>`;
    document.getElementById('freight-preview').style.display = 'block';
    document.getElementById('freight-btns').style.display = 'none';
    toast(`❌ 運費匯入驗證失敗：${errors.length} 項`);
    return;
  }

  if (!rows.length) { toast('❌ 找不到有效彙總資料'); return; }
  if (!detailResult.records.length) { toast('❌ 找不到有效明細資料'); return; }

  const totals = {
    daxi:     rows.reduce((s,r)=>s+r.daxi,0),
    dadu:     rows.reduce((s,r)=>s+r.dadu,0),
    gangshan: rows.reduce((s,r)=>s+r.gangshan,0),
  };
  parsedFreight = {
    rows,
    fileName,
    totals,
    detailRecords: detailResult.records,
    detailSummary: summarizeFreightDetails(detailResult.records),
    at: new Date(),
  };
  document.getElementById('freight-status').textContent = `✅ ${rows.length} 天 · ${detailResult.records.length} 筆`;
  showFreightPreview(parsedFreight);
  const btns = document.getElementById('freight-btns');
  btns.style.display = 'flex';
  toast(`✅ 解析完成：${rows.length} 天 / ${detailResult.records.length} 筆明細`);
}

function parseFreightDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const yyyy = value.getFullYear();
    const mm = String(value.getMonth() + 1).padStart(2, '0');
    const dd = String(value.getDate()).padStart(2, '0');
    return { full: `${yyyy}-${mm}-${dd}`, short: `${mm}/${dd}` };
  }
  const s = String(value || '').trim();
  const parts = s.split(/[/-]/);
  if (parts.length !== 3) return null;
  let yyyy = Number(parts[0]);
  const mm = Number(parts[1]);
  const dd = Number(parts[2]);
  if (!yyyy || !mm || !dd) return null;
  if (yyyy < 1911) yyyy += 1911;
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
  return {
    full: `${String(yyyy).padStart(4,'0')}-${String(mm).padStart(2,'0')}-${String(dd).padStart(2,'0')}`,
    short: `${String(mm).padStart(2,'0')}/${String(dd).padStart(2,'0')}`,
  };
}

function normalizedKey(s) {
  return String(s || '').replace(/\s+/g, '');
}

function valueByHeader(row, header) {
  const target = normalizedKey(header);
  for (const key of Object.keys(row)) {
    if (normalizedKey(key) === target) return row[key];
  }
  return undefined;
}

function parseMoney(value) {
  if (value === null || value === undefined || value === '') return 0;
  const n = Number(String(value).replace(/,/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

function parseFreightDetails(sheet) {
  const raw = XLSX.utils.sheet_to_json(sheet, { defval:'' });
  const errors = [];
  const records = [];
  if (!raw.length) return { records, errors:['明細分頁沒有資料'] };

  const required = ['倉別', '進貨日', '配送商', '預計計價結果', '到點計價結果', '計價結果'];
  const first = raw[0];
  required.forEach(name => {
    if (valueByHeader(first, name) === undefined) errors.push(`明細分頁缺少欄位：${name}`);
  });
  if (errors.length) return { records, errors };

  const duplicateRows = {};
  raw.forEach((r, i) => {
    const dateInfo = parseFreightDate(valueByHeader(r, '進貨日'));
    const warehouse = String(valueByHeader(r, '倉別') || '').trim();
    const vendor = String(valueByHeader(r, '配送商') || '').trim();
    const estimated = parseMoney(valueByHeader(r, '預計計價結果'));
    const point = parseMoney(valueByHeader(r, '到點計價結果'));
    const actual = parseMoney(valueByHeader(r, '計價結果'));
    if (!dateInfo || !warehouse || !vendor) return;
    if ([estimated, point, actual].some(v => Number.isNaN(v))) {
      errors.push(`明細分頁第 ${i + 2} 列金額不是數字`);
      return;
    }
    if ([estimated, point, actual].some(v => v < 0)) {
      errors.push(`明細分頁第 ${i + 2} 列出現負數，依規則不可匯入`);
      return;
    }

    const fullRowKey = Object.keys(r).sort().map(k => `${k}:${r[k]}`).join('|');
    if (duplicateRows[fullRowKey]) {
      errors.push(`明細分頁第 ${i + 2} 列與第 ${duplicateRows[fullRowKey]} 列完全重複`);
      return;
    }
    duplicateRows[fullRowKey] = i + 2;

    records.push({
      date: dateInfo.short,
      fullDate: dateInfo.full,
      warehouse,
      vendor,
      estimated,
      point,
      actual,
      diff: actual - estimated,
      rate: estimated ? actual / estimated * 100 : 0,
    });
  });

  return { records, errors };
}

function summarizeFreightDetails(records) {
  const vendors = {};
  records.forEach(r => {
    if (!vendors[r.vendor]) vendors[r.vendor] = { name:r.vendor, contract:0, point:0, actual:0, count:0, overCount:0, saveCount:0 };
    const v = vendors[r.vendor];
    v.contract += r.estimated;
    v.point += r.point;
    v.actual += r.actual;
    v.count += 1;
    if (r.rate > 90) v.overCount += 1;
    else v.saveCount += 1;
  });
  const vendorRows = Object.values(vendors).map(v => ({
    ...v,
    amount: v.actual - v.contract,
  }));
  const estimatedCost = records.reduce((s, r) => s + r.estimated, 0);
  const actualCost = records.reduce((s, r) => s + r.actual, 0);
  return {
    vendors: vendorRows,
    estimatedCost,
    actualCost,
    overCount: records.filter(r => r.rate > 90).length,
    saveCount: records.filter(r => r.rate <= 90).length,
    totalOrders: records.length,
  };
}

function showFreightPreview(parsed) {
  const { rows, totals, detailSummary } = parsed;
  const show = rows.length <= 8 ? rows :
    [...rows.slice(0,4), {date:'…',daxi:'…',dadu:'…',gangshan:'…'}, ...rows.slice(-3)];

  const trs = show.map(r => {
    const fmt = v => v === '…' ? '…' : '$' + Number(v).toLocaleString();
    return `<tr>
      <td class="preview-cell preview-label">${r.date}</td>
      <td class="preview-cell mono num-right text-daxi">${fmt(r.daxi)}</td>
      <td class="preview-cell mono num-right text-dadu">${fmt(r.dadu)}</td>
      <td class="preview-cell mono num-right text-gangshan">${fmt(r.gangshan)}</td>
    </tr>`;
  }).join('');

  document.getElementById('freight-preview').innerHTML = `
    <div class="import-preview-title">📋 資料預覽（${rows.length} 天 / ${detailSummary.totalOrders.toLocaleString()} 筆明細）</div>
    <div class="preview-stat-grid">
      <div class="preview-stat"><b>預計</b><br><span class="mono">${fmtMoney(detailSummary.estimatedCost)}</span></div>
      <div class="preview-stat"><b>實際</b><br><span class="mono">${fmtMoney(detailSummary.actualCost)}</span></div>
      <div class="preview-stat"><b>動支率&gt;90%</b><br><span class="mono">${detailSummary.overCount.toLocaleString()} 筆</span></div>
      <div class="preview-stat"><b>配送商</b><br><span class="mono">${detailSummary.vendors.length.toLocaleString()} 家</span></div>
    </div>
    <div class="preview-table-wrap preview-table-wrap-tall">
      <table class="preview-table">
        <thead class="preview-thead">
          <tr>
            <th class="preview-head">日期</th>
            <th class="preview-head preview-head-num text-daxi-light">大溪倉</th>
            <th class="preview-head preview-head-num text-dadu-light">大肚倉</th>
            <th class="preview-head preview-head-num text-gangshan-light">岡山倉</th>
          </tr>
        </thead>
        <tbody>${trs}</tbody>
        <tfoot class="preview-total">
          <tr>
            <td class="preview-cell preview-total-label">合計</td>
            <td class="preview-cell mono num-right preview-strong text-daxi">$${totals.daxi.toLocaleString()}</td>
            <td class="preview-cell mono num-right preview-strong text-dadu">$${totals.dadu.toLocaleString()}</td>
            <td class="preview-cell mono num-right preview-strong text-gangshan">$${totals.gangshan.toLocaleString()}</td>
          </tr>
        </tfoot>
      </table>
    </div>`;
  document.getElementById('freight-preview').style.display = 'block';
}

function applyFreight() {
  if (!parsedFreight) return;
  DATA.freight.dailyByWarehouse = parsedFreight.rows.map(r => [r.date, r.daxi, r.dadu, r.gangshan, r.fullDate]);
  DATA.freight.dailyTrend = parsedFreight.rows.map(r => [r.date, r.daxi + r.dadu + r.gangshan, r.fullDate]);
  DATA.freight.totalCost = parsedFreight.totals.daxi + parsedFreight.totals.dadu + parsedFreight.totals.gangshan;
  DATA.freight.estimatedCost = parsedFreight.detailSummary.estimatedCost;
  DATA.freight.actualCost = parsedFreight.detailSummary.actualCost;
  DATA.freight.totalOrders = parsedFreight.detailSummary.totalOrders;
  DATA.freight.overCount = parsedFreight.detailSummary.overCount;
  DATA.freight.saveCount = parsedFreight.detailSummary.saveCount;
  DATA.freight.diffThreshold = 90;
  DATA.freight.vendors = parsedFreight.detailSummary.vendors;
  DATA.freight.details = parsedFreight.detailRecords;
  const map = {};
  parsedFreight.rows.forEach(r => { map[r.fullDate] = r; });
  const existing = {};
  DATA.dispatch.daily.forEach(row => { existing[dispatchRowFullDate(row)] = row; });

  DATA.dispatch.daily = DATA.dispatch.daily.map(row => {
    const f = map[dispatchRowFullDate(row)];
    return f ? [row[0], row[1], f.daxi, row[3], f.dadu, row[5], f.gangshan, f.fullDate] : row;
  });

  parsedFreight.rows.forEach(r => {
    if (existing[r.fullDate]) return;
    DATA.dispatch.daily.push([r.date, 0, r.daxi, 0, r.dadu, 0, r.gangshan, r.fullDate]);
  });

  DATA.dispatch.daily.sort((a, b) => dispatchRowFullDate(a).localeCompare(dispatchRowFullDate(b)));
  updateDispatchLatestUploadDate(parsedFreight.rows.map(r => r.fullDate));
  if (currentPageId === 'annual') renderAnnualPage();
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
  const sheetNames = getLaborSheetNames(wb);
  if (!sheetNames.length) {
    toast('❌ 找不到有效工時工作表');
    document.getElementById('labor-status').textContent = '❌ 格式不符';
    return;
  }

  const raw = sheetNames.flatMap(sheetName =>
    XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' })
      .map((row, rowIndex) => ({ ...row, __sheetName: sheetName, __rowNumber: rowIndex + 2 }))
  );
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
  const warnings = [];
  raw.forEach((r, index) => {
    const hrs = Number(r['作業時數']) || 0;
    const cost = Number(r['實際費用']) || 0;
    const dateStr = excelSerialToDateString(r['日期']);
    const wh = normalizeWarehouseName(r['倉別']);
    const check = String(r['檢核'] || '').trim();
    const rowLabel = `${r.__sheetName} 第 ${r.__rowNumber || index + 2} 列`;
    if (!dateStr) warnings.push(`${rowLabel} 日期無法辨識`);
    if (!wh) warnings.push(`${rowLabel} 倉別無法辨識：${r['倉別']}`);
    if (check && check !== 'ok') warnings.push(`${rowLabel} 檢核=${check}`);
    if (hrs < 0 || cost < 0) warnings.push(`${rowLabel} 出現負數工時或費用`);
    if (!dateStr || !wh || hrs < 0 || cost < 0) return;

    const opArea = String(r['作業區域'] || '');
    records.push({
      wh,
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
      cost,
      check,
      sourceSheet: r.__sheetName,
    });
  });

  if (!records.length) { toast('❌ 找不到有效工時記錄'); return; }

  const sortedDates = [...new Set(records.map(r => r.date))].sort();
  if (sortedDates.length > 1) {
    const daySpan = Math.round((new Date(sortedDates[sortedDates.length - 1]) - new Date(sortedDates[0])) / 86400000) + 1;
    if (daySpan > 31) {
      toast(`❌ 資料涵蓋 ${daySpan} 天（${sortedDates[0]} ～ ${sortedDates[sortedDates.length - 1]}），超過 31 天上限，請分批上傳`);
      document.getElementById('labor-status').textContent = '❌ 超過 31 天';
      return;
    }
  }

  const dateRange = sortedDates.length
    ? { start: sortedDates[0], end: sortedDates[sortedDates.length - 1] }
    : null;

  const kpiRecords = getLaborKpiRecords(records);
  const totalHrs  = kpiRecords.reduce((s, r) => s + r.hours, 0);
  const totalCost = kpiRecords.reduce((s, r) => s + r.cost,  0);
  const personDays = new Set(kpiRecords.map(r => `${r.date}|${r.empId}`)).size;
  parsedLabor = {
    records,
    daily: summarizeLaborDaily(kpiRecords),
    warnings,
    fileName,
    sheetName: sheetNames.join('、'),
    sheetNames,
    dateRange,
    at: new Date(),
  };
  document.getElementById('labor-status').textContent = `✅ ${records.length.toLocaleString()} 筆 · ${totalHrs.toFixed(1)}h`;
  showLaborPreview(records, totalHrs, totalCost, personDays, warnings, { sheetNames, dateRange });
  document.getElementById('labor-btns').style.display = 'flex';
  toast(`✅ 工時解析完成：${records.length.toLocaleString()} 筆，警告 ${warnings.length.toLocaleString()} 項`);
}

function parseWorkbookPeriodFromFileName(fileName) {
  const match = String(fileName || '').match(/(20\d{4})(?:\s*[-_~～至]\s*(20\d{4}))?/);
  if (!match) return null;

  const start = match[1];
  const end = match[2] || start;
  return start <= end ? { start, end } : { start: end, end: start };
}

function getLaborSheetNames(wb) {
  const names = wb.SheetNames || [];
  const monthSheets = names.filter(name => /^20\d{4}$/.test(String(name))).sort();
  if (monthSheets.length) return monthSheets;
  const fallback = names.find(name => /月總?$/.test(String(name)));
  if (fallback) return [fallback];
  return names[0] ? [names[0]] : [];
}

function formatWorkbookPeriod(period) {
  if (!period) return '';
  const fmt = ym => `${ym.slice(0, 4)}/${ym.slice(4, 6)}`;
  return period.start === period.end ? fmt(period.start) : `${fmt(period.start)}-${fmt(period.end)}`;
}

function normalizeWarehouseName(value) {
  const s = String(value || '').trim();
  if (s === '大溪' || s === '大溪倉') return '大溪倉';
  if (s === '大肚' || s === '大肚倉') return '大肚倉';
  if (s === '岡山' || s === '岡山倉') return '岡山倉';
  return '';
}

function excelSerialToDateString(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const serial = Number(value);
  if (!Number.isFinite(serial) || serial <= 0) return '';
  const d = new Date(Math.round((serial - 25569) * 86400000));
  return d.toISOString().slice(0, 10);
}

function getLaborKpiRecords(records) {
  return records.filter(r => r.hours > 0 || r.cost > 0);
}

function summarizeLaborDaily(records) {
  const map = {};
  records.forEach(r => {
    if (!map[r.date]) map[r.date] = { date:r.date, '大溪倉':0, '大肚倉':0, '岡山倉':0 };
    map[r.date][r.wh] += r.cost;
  });
  return Object.values(map).sort((a, b) => a.date.localeCompare(b.date));
}

function showLaborPreview(records, totalHrs, totalCost, personDays, warnings, meta = {}) {
  const byOp = {};
  getLaborKpiRecords(records).forEach(r => {
    if (!byOp[r.opArea]) byOp[r.opArea] = { hrs: 0, cost: 0, personDays: new Set() };
    byOp[r.opArea].hrs   += r.hours;
    byOp[r.opArea].cost  += r.cost;
    byOp[r.opArea].personDays.add(`${r.date}|${r.empId}`);
  });
  const trs = Object.entries(byOp)
    .sort((a, b) => b[1].hrs - a[1].hrs)
    .map(([op, it]) => `<tr>
      <td class="preview-cell preview-label">${op}</td>
      <td class="preview-cell mono num-right">${it.hrs.toFixed(1)} h</td>
      <td class="preview-cell mono num-right">$${it.cost.toLocaleString()}</td>
      <td class="preview-cell num-right text-muted-strong">${it.personDays.size.toLocaleString()}</td>
    </tr>`).join('');

  const periodText = meta.dateRange ? ` / ${meta.dateRange.start} ～ ${meta.dateRange.end}` : '';
  const sheetText = meta.sheetNames?.length ? ` / ${meta.sheetNames.length} 個工作表` : '';
  document.getElementById('labor-preview').innerHTML = `
    <div class="import-preview-title">
      📋 作業區域摘要（${records.length.toLocaleString()} 筆 / ${personDays.toLocaleString()} 人日 / ${totalHrs.toFixed(1)}h / $${totalCost.toLocaleString()}${periodText}${sheetText}）
    </div>
    ${warnings.length ? `<div class="import-alert import-alert-warning">
      ⚠️ 警告但允許套用：${warnings.length.toLocaleString()} 項（前 5 項：${warnings.slice(0, 5).join('；')}）
    </div>` : ''}
    <div class="preview-table-wrap">
      <table class="preview-table">
        <thead class="preview-thead">
          <tr>
            <th class="preview-head">作業區域</th>
            <th class="preview-head preview-head-num">工時</th>
            <th class="preview-head preview-head-num">費用</th>
            <th class="preview-head preview-head-num">人日</th>
          </tr>
        </thead>
        <tbody>${trs}</tbody>
        <tfoot class="preview-total">
          <tr>
            <td class="preview-cell preview-total-label">合計</td>
            <td class="preview-cell mono num-right preview-total-value">${totalHrs.toFixed(1)} h</td>
            <td class="preview-cell mono num-right preview-total-value">$${totalCost.toLocaleString()}</td>
            <td class="preview-cell mono num-right preview-total-value">${personDays.toLocaleString()}</td>
          </tr>
        </tfoot>
      </table>
    </div>`;
  document.getElementById('labor-preview').style.display = 'block';
}

async function applyLabor() {
  if (!parsedLabor) return;
  LABOR_RAW = parsedLabor.records;
  applyLaborToDispatch(parsedLabor.daily);
  updateDispatchLatestUploadDate(parsedLabor.daily.map(r => r.date));
  const dates = parsedLabor.records.map(r => r.date).filter(Boolean).sort();
  DATA.dataLatest.labor = dates[dates.length - 1] || '';
  DATA.dataOldest = DATA.dataOldest || {};
  DATA.dataOldest.labor = dates[0] || '';
  if (currentPageId === 'labor') renderLaborPage();
  if (currentPageId === 'dispatch') renderDispatchPage();
  if (currentPageId === 'productivity') renderProductivityPage();
  if (currentPageId === 'annual') renderAnnualPage();
  updateStatus();
  toast('✅ 工時資料已套用！總費用動支率的人力欄位已更新');
  try {
    const result = await syncCloudLabor(parsedLabor);
    toast(`✅ 人力費用已同步雲端（${result.rows} 筆）`);
  } catch (err) {
    console.warn('Labor cloud sync failed:', err);
    toast(`⚠️ 人力費用已套用，但雲端同步失敗：${err.message}`);
  }
}

function applyLaborToDispatch(dailyRows) {
  const map = {};
  dailyRows.forEach(r => {
    map[r.date] = r;
  });
  const existing = {};
  DATA.dispatch.daily.forEach(row => { existing[dispatchRowFullDate(row)] = row; });

  DATA.dispatch.daily = DATA.dispatch.daily.map(row => {
    const labor = map[dispatchRowFullDate(row)];
    return labor
      ? [row[0], labor['大溪倉'], row[2], labor['大肚倉'], row[4], labor['岡山倉'], row[6], labor.date]
      : row;
  });

  dailyRows.forEach(r => {
    const mmdd = r.date.slice(5).replace('-', '/');
    if (existing[r.date]) return;
    DATA.dispatch.daily.push([mmdd, r['大溪倉'], 0, r['大肚倉'], 0, r['岡山倉'], 0, r.date]);
  });

  DATA.dispatch.daily.sort((a, b) => dispatchRowFullDate(a).localeCompare(dispatchRowFullDate(b)));
}

function syncDispatchLaborFromRaw() {
  if (typeof LABOR_RAW === 'undefined' || !LABOR_RAW.length) return;
  applyLaborToDispatch(summarizeLaborDaily(getLaborKpiRecords(LABOR_RAW)));
}

function updateDispatchLatestUploadDate(dates) {
  const validDates = (dates || []).filter(Boolean).sort();
  if (!validDates.length) return;
  DATA.dispatch.latestUploadDate = validDates[validDates.length - 1];
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
  const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
  if (!rawRows.length) { toast('❌ 找不到有效資料'); return; }

  // 正規化欄位名稱（去除前後空白）
  const raw = rawRows.map(r => {
    const n = {};
    Object.entries(r).forEach(([k, v]) => { n[k.trim()] = v; });
    return n;
  });

  const sample = raw[0];
  const required = ['倉別', '日期', '業務類別', '作業區', '工時區域', '揀次'];
  const missing = required.filter(c => !(c in sample));
  if (missing.length) {
    const found = Object.keys(sample).join('、');
    toast('❌ 缺少欄位：' + missing.join('、'));
    document.getElementById('picks-status').textContent = '❌ 格式不符';
    document.getElementById('picks-preview').style.display = 'block';
    document.getElementById('picks-preview').innerHTML = `
      <div style="background:#fdf0ec;border-left:3px solid #d9401b;padding:10px 14px;font-size:12px;line-height:1.8">
        <b>❌ 缺少欄位：${missing.join('、')}</b><br>
        檔案中找到的欄位：<span style="color:#1e5ca8">${found}</span>
      </div>`;
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
  document.getElementById('picks-status').textContent = `✅ ${records.length} 筆 · ${totalPicks.toLocaleString()} 次`;
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
      <td class="preview-cell preview-label">${op}</td>
      <td class="preview-cell mono num-right">${cnt.toLocaleString()}</td>
    </tr>`).join('');

  document.getElementById('picks-preview').innerHTML = `
    <div class="import-preview-title">📋 工時區域摘要（共 ${totalPicks.toLocaleString()} 揀次）</div>
    <div class="preview-table-wrap">
      <table class="preview-table">
        <thead class="preview-thead">
          <tr>
            <th class="preview-head">工時區域</th>
            <th class="preview-head preview-head-num">揀次</th>
          </tr>
        </thead>
        <tbody>${trs}</tbody>
        <tfoot class="preview-total">
          <tr>
            <td class="preview-cell preview-total-label">合計</td>
            <td class="preview-cell mono num-right preview-total-value">${totalPicks.toLocaleString()}</td>
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
  if (currentPageId === 'productivity') renderProductivityPage();
  updateStatus();
  toast('✅ 揀次資料已套用！切換至「揀次分析」頁查看');
}

function resetPicks() {
  parsedPicks = null;
  PICKS_RAW = [];
  document.getElementById('picks-status').textContent = '尚未上傳';
  document.getElementById('picks-preview').style.display = 'none';
  document.getElementById('picks-btns').style.display = 'none';
  document.getElementById('picks-file').value = '';
  if (currentPageId === 'picks') renderPicksPage();
  if (currentPageId === 'productivity') renderProductivityPage();
  updateStatus();
}

function updateStatus() {
  const latestFreight = (() => {
    if (DATA.freight.details?.length)
      return DATA.freight.details.map(r => r.fullDate).filter(Boolean).sort().pop() || '';
    return DATA.freight.dailyTrend.map(r => r[2] || shortToFreightFullDate(r[0])).filter(Boolean).sort().pop() || '';
  })();
  const oldestFreight = (() => {
    if (DATA.freight.details?.length)
      return DATA.freight.details.map(r => r.fullDate).filter(Boolean).sort()[0] || '';
    return DATA.freight.dailyTrend.map(r => r[2] || shortToFreightFullDate(r[0])).filter(Boolean).sort()[0] || '';
  })();
  const latestLabor = (() => {
    const raw = (typeof LABOR_RAW !== 'undefined') ? LABOR_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort().pop() || '';
  })();
  const oldestLabor = (() => {
    const raw = (typeof LABOR_RAW !== 'undefined') ? LABOR_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort()[0] || '';
  })();
  const latestPicks = (() => {
    const raw = (typeof PICKS_RAW !== 'undefined') ? PICKS_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort().pop() || '';
  })();
  const oldestPicks = (() => {
    const raw = (typeof PICKS_RAW !== 'undefined') ? PICKS_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort()[0] || '';
  })();

  const rows = [
    { type:'💰 年度預算', real:hasDispatchBudget(), latest: DATA.dataLatest?.budget || '', oldest: DATA.dataOldest?.budget || '' },
    { type:'🚚 運務費用', real:!!parsedFreight,      latest: latestFreight, oldest: oldestFreight },
    { type:'💵 人力費用', real:!!parsedLabor || !!((typeof LABOR_RAW !== 'undefined') && LABOR_RAW.length), latest: latestLabor, oldest: oldestLabor },
    { type:'⚡ 揀次資料', real:!!parsedPicks,        latest: latestPicks,   oldest: oldestPicks  },
  ];

  const dateCell = v => v
    ? `<span style="font-family:monospace;font-size:13px">${v}</span>`
    : `<span style="color:#bbb;font-size:12px">—</span>`;

  document.getElementById('status-tbody').innerHTML = rows.map(r => {
    const c = r.real ? '#1b7c33' : '#e07855';
    const s = r.real ? '✅ 已套用' : '⚠️ 尚未上傳';
    return `<tr>
      <td style="font-weight:700">${r.type}</td>
      <td><span style="color:${c};font-weight:700">${s}</span></td>
      <td>${dateCell(r.latest)}</td>
      <td>${dateCell(r.oldest)}</td>
    </tr>`;
  }).join('');
}

// ════════════════════════════════════════════
// Picks Page 揀次分析
// ════════════════════════════════════════════
function initPicksPage() { renderPicksPage(); }

function renderPicksPage() {
  syncDashboardDateInputs('picks');
  const wh = document.getElementById('picks-wh')?.value || '';
  const op = document.getElementById('picks-op')?.value || '';

  let data = (typeof PICKS_RAW !== 'undefined') ? PICKS_RAW : [];
  data = data.filter(r => dateInSelectedRange(r.date));
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

  const picksPeriod = dates.length
    ? (() => {
        const y0 = dates[0].slice(0,7), y1 = dates[dates.length-1].slice(0,7);
        if (y0 === y1) { const d = new Date(dates[0]); return `${d.getFullYear()}年${d.getMonth()+1}月`; }
        return `${y0.replace('-','/')} ~ ${y1.replace('-','/')}`;
      })()
    : '—';

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
    <div class="gold-band">P001 · TOTAL</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>月總揀次</div></div>
    <div style="padding:20px 16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-blue);line-height:1;margin-bottom:6px">${total.toLocaleString()}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted)">三倉合計 · ${picksPeriod}</div>
    </div>
  </div>
  <div class="w s4">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">P002 · DAILY AVG</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>日均揀次</div></div>
    <div style="padding:20px 16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-ink);line-height:1;margin-bottom:6px">${avgDaily.toLocaleString()}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted)">每日平均（${dates.length} 天）</div>
    </div>
  </div>
  <div class="w s4">
    <div class="gold-band" style="background:var(--ry-red);color:white">P003 · PEAK DAY</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-red)"></div>峰值日</div></div>
    <div style="padding:20px 16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-red);line-height:1;margin-bottom:6px">${peakDate.slice(5) || '—'}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted)">${peakVal.toLocaleString()} 揀次</div>
    </div>
  </div>
  <div class="w s12">
    <div class="gold-band">P004 · 📅 每日揀次趨勢</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>每日三倉合計揀次（${picksPeriod}）</div><span class="wmeta">金色=峰值</span></div>
    ${trendHtml}
  </div>
  <div class="w s12">
    <div class="gold-band">P005 · 📊 作業區域分析</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各作業區域揀次量 × 三倉</div><span class="wmeta">單位：次</span></div>
    <table class="tbl">
      <thead><tr><th>作業區域</th><th style="text-align:right">大肚倉</th><th style="text-align:right">大溪倉</th><th style="text-align:right">岡山倉</th><th style="text-align:right">合計</th><th>佔比</th></tr></thead>
      <tbody>${opRows}</tbody>
    </table>
  </div>`;

  const meta = document.getElementById('picks-meta');
  if (meta) meta.textContent = `資料：${DATA.dateFrom} ~ ${DATA.dateTo} · ${picksPeriod} · 三倉 · ${data.length.toLocaleString()} 筆作業記錄`;
}

// ════════════════════════════════════════════
// Labor Page 人力工時結構
// ════════════════════════════════════════════
function initLaborPage() {
  syncLaborDeptOptions();
  renderLaborPage();
}

function syncLaborDeptOptions() {
  const select = document.getElementById('labor-vendor');
  if (!select) return;

  const current = select.value;
  const rawData = (typeof LABOR_RAW !== 'undefined') ? LABOR_RAW : [];
  const depts = [...new Set(rawData.map(r => r.dept).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, 'zh-Hant'));

  select.innerHTML = '<option value="">全部課別</option>' +
    depts.map(dept => `<option value="${dept}">${deptDisplayName(dept) || dept}</option>`).join('');
  if (current && depts.includes(current)) select.value = current;
}

function renderLaborPage() {
  syncDashboardDateInputs('labor');
  syncLaborDeptOptions();
  const shiftFilter  = document.getElementById('labor-shift')?.value  || '';
  const deptFilter = document.getElementById('labor-vendor')?.value || '';

  const rawData = (typeof LABOR_RAW !== 'undefined') ? LABOR_RAW : [];
  if (!rawData.length) {
    const fmeta = document.getElementById('labor-filter-meta');
    if (fmeta) fmeta.textContent = '0 筆工時記錄';
    const meta = document.getElementById('labor-meta');
    if (meta) meta.textContent = '資料：尚未匯入 · 0 人日 · 0 筆明細';
    document.getElementById('labor-grid').innerHTML = `
      <div class="w s12">
        <div class="wh"><div class="wl"><div class="wdot"></div>尚未匯入工時資料</div></div>
        <div style="padding:32px;text-align:center;color:var(--ry-muted);font-size:var(--fs-sm);line-height:1.8">
          請到「資料匯入」上傳全區工時 Excel，套用後此頁會自動產生人力工時結構。
        </div>
      </div>`;
    return;
  }

  let data = rawData.filter(r => dateInSelectedRange(r.date));
  data = data.filter(r => r.opArea !== '午休時間' && r.hours > 0);
  if (shiftFilter)  data = data.filter(r => r.shift  === shiftFilter);
  if (deptFilter) data = data.filter(r => r.dept === deptFilter);

  const totalHrs  = data.reduce((s, r) => s + r.hours, 0);
  const totalCost = data.reduce((s, r) => s + r.cost,  0);
  const avgRate   = totalHrs > 0 ? Math.round(totalCost / totalHrs) : 0;
  const empCount  = new Set(data.map(r => r.empId)).size;
  const personDays = new Set(data.map(r => `${r.date}|${r.empId}`)).size;
  const laborDates = [...new Set(data.map(r => r.date))].sort();
  const laborPeriod = laborDates.length
    ? (() => {
        const y0 = laborDates[0].slice(0,7), y1 = laborDates[laborDates.length-1].slice(0,7);
        if (y0 === y1) { const d = new Date(laborDates[0]); return `${d.getFullYear()}年${d.getMonth()+1}月`; }
        return `${y0.replace('-','/')} ~ ${y1.replace('-','/')}`;
      })()
    : '—';

  const byOp = {};
  data.forEach(r => {
    if (!byOp[r.opArea]) byOp[r.opArea] = { hrs: 0, cost: 0 };
    byOp[r.opArea].hrs  += r.hours;
    byOp[r.opArea].cost += r.cost;
  });
  const ops = Object.keys(byOp).sort((a, b) => byOp[b].hrs - byOp[a].hrs);
  const COLORS = ['#1e5ca8', '#f5c400', '#d9401b', '#2ea85a', '#e07855', '#5a6478'];

  const opStats = ops.map((o, i) => {
    const pct  = totalHrs > 0 ? (byOp[o].hrs / totalHrs * 100) : 0;
    const rate = byOp[o].hrs > 0 ? Math.round(byOp[o].cost / byOp[o].hrs) : 0;
    return { name: o, hrs: byOp[o].hrs, pct, rate, color: COLORS[i % COLORS.length] };
  });

  function pieSlicePath(cx, cy, r, start, end) {
    const startAngle = start * Math.PI * 2 - Math.PI / 2;
    const endAngle = end * Math.PI * 2 - Math.PI / 2;
    const x1 = cx + r * Math.cos(startAngle);
    const y1 = cy + r * Math.sin(startAngle);
    const x2 = cx + r * Math.cos(endAngle);
    const y2 = cy + r * Math.sin(endAngle);
    const largeArc = end - start > 0.5 ? 1 : 0;
    return `M ${cx} ${cy} L ${x1.toFixed(3)} ${y1.toFixed(3)} A ${r} ${r} 0 ${largeArc} 1 ${x2.toFixed(3)} ${y2.toFixed(3)} Z`;
  }

  let cursor = 0;
  const pieSlices = opStats.map(s => {
    const start = cursor;
    const share = totalHrs > 0 ? s.hrs / totalHrs : 0;
    cursor += share;
    if (share >= 0.9999) return `<circle cx="120" cy="120" r="94" fill="${s.color}"></circle>`;
    return `<path d="${pieSlicePath(120, 120, 94, start, cursor)}" fill="${s.color}"></path>`;
  }).join('');

  const structHtml = opStats.length ? `
    <div class="labor-pie-layout">
      <svg class="labor-pie-chart" viewBox="0 0 240 240" role="img" aria-label="作業區域工時占比圓餅圖">
        ${pieSlices}
        <circle cx="120" cy="120" r="94" fill="none" stroke="var(--ry-paper)" stroke-width="2"></circle>
      </svg>
      <div class="labor-top-list">
        <div class="labor-top-title">前五名</div>
        ${opStats.slice(0, 5).map((s, i) => `
          <div class="labor-top-item">
            <span class="labor-top-rank">${String(i + 1).padStart(2, '0')}</span>
            <span class="labor-top-swatch" style="background:${s.color}"></span>
            <span class="labor-top-name">${s.name}</span>
            <span class="labor-top-value">${s.hrs.toFixed(1)}h · ${s.pct.toFixed(1)}%</span>
          </div>
        `).join('')}
      </div>
    </div>` : '<div style="color:var(--ry-muted);padding:16px;text-align:center">無資料</div>';

  const byShift = {};
  data.forEach(r => {
    if (!byShift[r.shift]) byShift[r.shift] = { hrs: 0, cost: 0 };
    byShift[r.shift].hrs  += r.hours;
    byShift[r.shift].cost += r.cost;
  });
  const shiftRows = ['日', '中', '夜'].filter(s => byShift[s]).map(s => {
    const it   = byShift[s];
    const rate = it.hrs > 0 ? Math.round(it.cost / it.hrs) : 0;
    return `<tr>
      <td><b>${s}班</b></td>
      <td style="text-align:right;font-family:var(--f-mono)">${it.hrs.toFixed(1)}</td>
      <td style="text-align:right;font-family:var(--f-mono)">$${it.cost.toLocaleString()}</td>
      <td style="text-align:right;font-family:var(--f-mono)">$${rate}</td>
    </tr>`;
  }).join('');

  const byDept = {};
  data.forEach(r => {
    const dept = deptDisplayName(r.dept) || r.dept || '未分類';
    if (!byDept[dept]) byDept[dept] = { hrs: 0, cost: 0, emps: new Set() };
    byDept[dept].hrs  += r.hours;
    byDept[dept].cost += r.cost;
    byDept[dept].emps.add(r.empId);
  });
  const deptRows = Object.entries(byDept)
    .sort((a, b) => b[1].hrs - a[1].hrs)
    .map(([v, it]) => {
      const rate = it.hrs > 0 ? Math.round(it.cost / it.hrs) : 0;
      return `<tr>
        <td class="labor-dept-name">${v}</td>
        <td class="mono num-right">${it.emps.size}</td>
        <td class="mono num-right">${it.hrs.toFixed(1)}</td>
        <td class="mono num-right">$${it.cost.toLocaleString()}</td>
        <td class="mono num-right">$${rate}</td>
      </tr>`;
    }).join('');

  const fmeta = document.getElementById('labor-filter-meta');
  if (fmeta) fmeta.textContent = `${data.length} 筆工時記錄`;

  document.getElementById('labor-grid').innerHTML = `
  <div class="w s3 metric-card">
    <div class="gold-band">L001 · HOURS</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>總工時</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-blue);line-height:1">${totalHrs.toFixed(1)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">小時</div>
    </div>
  </div>
  <div class="w s3 metric-card">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">L002 · COST</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>總費用</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">$${totalCost.toLocaleString()}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">元</div>
    </div>
  </div>
  <div class="w s3 metric-card">
    <div class="gold-band" style="background:#2ea85a;color:white">L003 · RATE</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:#2ea85a"></div>平均時薪</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:#2ea85a;line-height:1">$${avgRate}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">元/小時</div>
    </div>
  </div>
  <div class="w s3 metric-card">
    <div class="gold-band" style="background:var(--ry-muted);color:white">L004 · PEOPLE</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-muted)"></div>出勤人日</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">${personDays.toLocaleString()}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">${empCount.toLocaleString()} 位員工</div>
    </div>
  </div>
  <div class="w s6">
    <div class="gold-band">L005 · ⚡ 工時結構 · 作業區域</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各作業區域工時佔比</div><span class="wmeta">總 ${totalHrs.toFixed(1)} h</span></div>
    <div class="labor-pie-wrap">${structHtml}</div>
  </div>
  <div class="w s6 table-card labor-shift-card">
    <div class="gold-band">L006 · 🌙 班別工時分析</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>日班 / 中班 / 夜班</div></div>
    <table class="tbl">
      <thead><tr><th>班別</th><th style="text-align:right">工時(h)</th><th style="text-align:right">費用</th><th style="text-align:right">時薪</th></tr></thead>
      <tbody>${shiftRows || '<tr><td colspan="4" style="text-align:center;color:var(--ry-muted)">無資料</td></tr>'}</tbody>
    </table>
  </div>
  <div class="w s12 table-card labor-dept-card">
    <div class="gold-band">L007 · 🏢 課別工時彙總</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各作業課別工時與費用</div></div>
    <div class="table-edge labor-dept-edge">
      <div class="scroll-x">
        <table class="tbl labor-dept-table">
          <thead>
            <tr>
              <th>課別</th>
              <th class="num-right">人次</th>
              <th class="num-right">工時(H)</th>
              <th class="num-right">費用</th>
              <th class="num-right">時薪</th>
            </tr>
          </thead>
          <tbody>${deptRows || '<tr><td colspan="5" class="labor-dept-empty">無資料</td></tr>'}</tbody>
        </table>
      </div>
      <div class="table-note labor-formula-note">
        📌 課別：依 Excel「作業課別」欄位分組<br>
        📌 人次：該課別不重複員編數 · 工時(H)：該課別作業時數加總<br>
        📌 費用：該課別實際費用加總 · 時薪：費用 ÷ 工時
      </div>
    </div>
  </div>`;

  const meta = document.getElementById('labor-meta');
  if (meta) meta.textContent = `資料：${DATA.dateFrom} ~ ${DATA.dateTo} · ${laborPeriod} · 全區各課 · ${personDays.toLocaleString()} 人日 · ${data.length.toLocaleString()} 筆明細`;
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
          <div class="org-row-name">${deptDisplayName(d.name) || d.name} <span class="org-type-tag" style="background:${tc}20;color:${tc}">● ${d.type}</span></div>
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

// ════════════════════════════════════════════
// Productivity Page 人效監控
// ════════════════════════════════════════════
function initProductivityPage() { renderProductivityPage(); }

function renderProductivityPage() {
  syncDashboardDateInputs('productivity');

  const labor = (typeof LABOR_RAW !== 'undefined' ? LABOR_RAW : [])
    .filter(r => dateInSelectedRange(r.date) && r.hours > 0 && r.opArea !== '午休時間');
  const picks = (typeof PICKS_RAW !== 'undefined' ? PICKS_RAW : [])
    .filter(r => dateInSelectedRange(r.date));

  if (!labor.length && !picks.length) {
    document.getElementById('productivity-grid').innerHTML = `
    <div class="w s12">
      <div class="wh"><div class="wl"><div class="wdot"></div>尚未匯入資料</div></div>
      <div style="padding:32px;text-align:center;color:var(--ry-muted);font-size:var(--fs-sm);line-height:1.8">
        請先到「資料匯入」上傳工時資料與揀次資料，套用後此頁會自動產生人效指標。
      </div>
    </div>`;
    return;
  }

  const t = productivityTotals(labor, picks);
  const picksTotal = picks.reduce((s, r) => s + r.picks, 0);
  const byDate = {};
  picks.forEach(r => { byDate[r.date] = (byDate[r.date] || 0) + r.picks; });
  const pickDates = Object.keys(byDate).sort();
  const dailyVals = pickDates.map(d => byDate[d]);
  const avgDaily = pickDates.length ? Math.round(picksTotal / pickDates.length) : 0;
  const maxDay = Math.max(...dailyVals, 1);
  const peakVal = dailyVals.length ? Math.max(...dailyVals) : 0;
  const volumePeriod = pickDates.length
    ? (() => {
        const y0 = pickDates[0].slice(0, 7);
        const y1 = pickDates[pickDates.length - 1].slice(0, 7);
        if (y0 === y1) {
          const d = new Date(`${pickDates[0]}T00:00:00`);
          return `${d.getFullYear()}年${d.getMonth() + 1}月`;
        }
        return `${y0.replace('-', '/')} ~ ${y1.replace('-', '/')}`;
      })()
    : '—';

  let trendHtml = '<div class="prod-trend-bars">';
  pickDates.forEach((d, i) => {
    const h = Math.round(dailyVals[i] / maxDay * 64);
    const day = d.slice(8);
    const isPeak = dailyVals[i] === peakVal && peakVal > 0;
    trendHtml += `<div class="prod-trend-day" title="${d}: ${dailyVals[i].toLocaleString()}">
      <div class="prod-trend-bar" style="height:${Math.max(h, 2)}px;background:${isPeak ? 'var(--ry-gold)' : 'var(--ry-blue)'}"></div>
      <div class="prod-trend-label">${day}</div>
    </div>`;
  });
  trendHtml += '</div>';

  const volumeCards = `
  <div class="w s3 metric-card">
    <div class="gold-band">E001 · TOTAL</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>總揀次</div></div>
    <div class="prod-kpi-body">
      <div class="prod-kpi-value text-blue">${picksTotal.toLocaleString()}</div>
      <div class="prod-kpi-note">三倉合計 · ${volumePeriod}</div>
    </div>
  </div>
  <div class="w s3 metric-card">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">E002 · DAILY AVG</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>日均揀次</div></div>
    <div class="prod-kpi-body">
      <div class="prod-kpi-value">${avgDaily.toLocaleString()}</div>
      <div class="prod-kpi-note">每日平均 · ${pickDates.length} 天</div>
    </div>
  </div>`;

  document.getElementById('productivity-grid').innerHTML = [
    volumeCards,
    renderM019(t, 'E003', 's3'),
    renderM020(t, 'E004', 's3'),
    renderM023(labor, picks, 'E005'),
    `<div class="w s12 chart-card">
      <div class="gold-band">E006 · 每日揀次趨勢</div>
      <div class="wh"><div class="wl"><div class="wdot"></div>每日三倉合計揀次</div><span class="wmeta">金色=峰值</span></div>
      ${trendHtml}
    </div>`,
    renderM025(labor, picks, 'E007'),
  ].join('');

  const dates = [...new Set([...labor.map(r => r.date), ...picks.map(r => r.date)])].sort();
  const period = dates.length >= 2
    ? `${dates[0]} ~ ${dates[dates.length - 1]}`
    : dates[0] || DATA.dateFrom;
  const meta = document.getElementById('productivity-meta');
  if (meta) {
    meta.textContent = `資料區間：${period} · ${t.picks.toLocaleString()} 揀次 · ${t.hrs.toFixed(1)} h 工時 · PPH ${t.pph.toFixed(1)}`;
  }
}

// ════════════════════════════════════════════
// Monthly Page 月度結算
// ════════════════════════════════════════════
function initMonthlyPage() { renderMonthlyPage(); }

function renderMonthlyPage() {
  syncDashboardDateInputs('monthly');

  const labor = (typeof LABOR_RAW !== 'undefined' ? LABOR_RAW : [])
    .filter(r => dateInSelectedRange(r.date) && r.hours > 0 && r.opArea !== '午休時間');
  const picks = (typeof PICKS_RAW !== 'undefined' ? PICKS_RAW : [])
    .filter(r => dateInSelectedRange(r.date));

  const laborByWh   = sumLaborByWh(labor);
  const freightByWh = sumFreightByWh();
  const budget      = getMonthlyBudgetByWh();
  const freightTotal = Object.values(freightByWh).reduce((s, v) => s + v, 0);
  const hasAny = labor.length || picks.length || freightTotal;

  if (!hasAny) {
    document.getElementById('monthly-grid').innerHTML = `
    <div class="w s12">
      <div class="wh"><div class="wl"><div class="wdot"></div>尚未匯入資料</div></div>
      <div style="padding:32px;text-align:center;color:var(--ry-muted);font-size:var(--fs-sm);line-height:1.8">
        請先到「資料匯入」上傳工時資料與運費資料，套用後此頁會自動產生月度結算報表。
      </div>
    </div>`;
    return;
  }

  document.getElementById('monthly-grid').innerHTML = [
    labor.length  ? renderM001(labor)                    : '',
    labor.length  ? renderM002(labor)                    : '',
    freightTotal  ? renderMFreight()                     : '',
    renderMSummaryKpi(labor, picks, freightTotal, budget),
    renderM011(laborByWh, freightByWh, budget),
    labor.length  ? renderECBreakdown(labor)             : '',
    labor.length && picks.length ? renderMEfficiency(labor, picks) : '',
  ].join('');

  const allDates = [...labor.map(r => r.date), ...picks.map(r => r.date)].filter(Boolean).sort();
  const period = allDates.length >= 2
    ? `${allDates[0]} ~ ${allDates[allDates.length - 1]}`
    : allDates[0] || DATA.dateFrom;
  const meta = document.getElementById('monthly-meta');
  if (meta) meta.textContent = `月度結算 · ${period}`;
}

// ════════════════════════════════════════════
// Annual Page 年度規劃分析
// ════════════════════════════════════════════
function renderAnnualPage() {
  const grid = document.getElementById('annual-grid');
  if (!grid) return;
  const isLabor = annualViewMode === 'labor';
  grid.innerHTML = `
    <div class="w s12 annual-mode-panel">
      <div>
        <div class="annual-mode-title">年度分析項目</div>
        <div class="annual-mode-subtitle">切換檢視人力費用或運務費用，圖表與明細會同步更新。</div>
      </div>
      <div class="annual-segment" role="tablist" aria-label="年度分析項目切換">
        <button class="annual-segment-btn ${isLabor ? 'active' : ''}" type="button" onclick="setAnnualViewMode('labor')" aria-selected="${isLabor}">人力</button>
        <button class="annual-segment-btn ${!isLabor ? 'active' : ''}" type="button" onclick="setAnnualViewMode('freight')" aria-selected="${!isLabor}">運務</button>
      </div>
    </div>
    ${renderAnnualSection(annualViewMode)}
  `;
}

function setAnnualViewMode(mode) {
  annualViewMode = mode === 'freight' ? 'freight' : 'labor';
  renderAnnualPage();
}
