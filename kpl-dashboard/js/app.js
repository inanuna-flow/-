// ═══════════════════════════════════════════════════════
// app.js · 頁面設定 / 導覽 / 初始化
// ═══════════════════════════════════════════════════════

// ── 超級管理員設定 ──
const ADMIN_USER_ID = 'inari';
let currentUserId = '';
let pagePermissions = {}; // 從伺服器載入

function isAdmin() {
  return (currentUserId || sessionStorage.getItem('kpl_user') || '').toLowerCase() === ADMIN_USER_ID;
}

function handleAuthExpired(res) {
  if (res.status !== 401) return false;
  sessionStorage.removeItem('kpl_auth');
  sessionStorage.removeItem('kpl_user');
  location.href = 'login.html';
  return true;
}

async function loadSession() {
  const res = await fetch('/api/session', { cache: 'no-store' });
  if (handleAuthExpired(res)) return false;
  if (!res.ok) {
    location.href = 'login.html';
    return false;
  }
  const data = await res.json();
  currentUserId = String(data.userId || '');
  sessionStorage.setItem('kpl_user', currentUserId);
  sessionStorage.setItem('kpl_auth', '1');
  return true;
}

async function loadPagePermissions() {
  try {
    const res = await fetch('/api/page-permissions');
    if (handleAuthExpired(res)) return;
    if (res.ok) pagePermissions = await res.json();
  } catch { /* 若載入失敗，管理員帳號仍可看全部 */ }
}

function isPageVisible(pageId) {
  if (isAdmin()) return true; // 管理員永遠看得到所有頁面
  return pagePermissions[pageId] !== false;
}

const PAGES = [
  {
    group: '📊 儀表板總覽',
    items: [
      { id:'daily',        icon:'📅', label:'每日動支監控', status:'ready' },
      { id:'dispatch',     icon:'💼', label:'總費用動支率', status:'ready' },
    ]
  },
  {
    group: '💰 成本分析',
    items: [
      { id:'freight',      icon:'🚚', label:'運費損益分析', status:'ready' },
      { id:'labor',        icon:'⏱', label:'人力工時結構', status:'ready' },
    ]
  },
  {
    group: '⚡ 效率分析',
    items: [
      { id:'picks',        icon:'⚡', label:'揀次分析', status:'ready' },
      { id:'productivity', icon:'📊', label:'揀次人效分析', status:'ready' },
    ]
  },
  {
    group: '📋 預算規劃',
    items: [
      { id:'annual',       icon:'📋', label:'年度規劃分析', status:'ready' },
    ]
  },
  {
    group: '📁 資料管理',
    items: [
      { id:'import',       icon:'📤', label:'資料匯入',     status:'ready' },
      { id:'versions',     icon:'🕘', label:'版本資訊',     status:'ready' },
    ]
  },
  {
    group: '🔐 帳號權限管理',
    items: [
      { id:'accountPermissions', icon:'👥', label:'帳號權限',         status:'ready', adminOnly: true },
      { id:'accountIpRules',     icon:'🛡️', label:'後臺限制 IP 管理', status:'ready', adminOnly: true },
      { id:'accountLoginAudit',  icon:'📜', label:'系統登入日誌',     status:'ready', adminOnly: true },
      { id:'accountErrorAudit',  icon:'🚨', label:'全域錯誤例外日誌', status:'ready', adminOnly: true },
    ]
  },
  {
    group: '⚙️ 系統設定',
    items: [
      { id:'org',          icon:'🏢', label:'組織設定',     status:'ready' },
      { id:'typography',   icon:'🔤', label:'文字樣式設定', status:'ready' },
    ]
  },
];

// 索引必須與 PAGES 順序一一對應；label 用 Map 去重，若重複會把後面的群組併進前面
const SIDEBAR_GROUP_META = [
  { label: '儀表板總覽', icon: '📊' },
  { label: '成本分析',   icon: '💼' },
  { label: '效率分析',   icon: '⚡' },
  { label: '預算規劃',   icon: '📈' },
  { label: '資料管理',   icon: '📁' },
  { label: '帳號權限管理', icon: '🔐' },
  { label: '系統設定',   icon: '⚙️' },
];
const THEME_STORAGE_KEY = 'kplThemeMode';
const SIDEBAR_PINNED_KEY = 'kplSidebarPinned';
const THEME_MODES = ['original', 'day', 'night'];
const RELEASE_READ_KEY = 'kplReleaseRead';
const FALLBACK_RELEASES = [
  {
    version: 'v0.4.0',
    date: '2026-06-06',
    title: '儀表板閱讀體驗更新',
    summary: '統一元件標題、表格線條與數字字重。',
    changes: [
      '元件標題移至卡片外，並調整標題間距與字重。',
      '區間動支彙總改為中性色表格，統一細線與分組樣式。',
      '移除摘要卡與動支百分比進度線，重新整理資訊對齊。',
      '元件數字統一使用正常字重，提升長時間閱讀舒適度。',
      '運費損益分析由展示版介面改為資料驅動的 KPI、圖表與決策表元件。',
    ],
  },
  {
    version: 'v0.3.0',
    date: '2026-05-26',
    title: '版面與主題系統',
    summary: '新增可釘選側邊欄與三種主題。',
    changes: [
      '新增原味、白天、黑夜主題並保存使用者選擇。',
      '新增可釘選、收合與滑鼠靠近浮出的側邊欄。',
      '全站套用 Inter 與 Geist Mono 字體。',
    ],
  },
];
let RELEASES = [...FALLBACK_RELEASES];

let currentPageId = 'daily';
let annualViewMode = 'labor';

const TYPOGRAPHY_STORAGE_KEY = 'kplTypographySettings';
const TYPOGRAPHY_VERSION_KEY = 'kplTypographyVersion';
const TYPOGRAPHY_VERSION = '4';
const TYPOGRAPHY_FONT_MAP = {
  ui: 'var(--font-ui)',
  number: 'var(--font-number)',
};
const TYPOGRAPHY_CONTROLS = [
  { key:'page-title', label:'頁面主標題', hint:'頁面 H1，例如「運費損益分析」', sample:'運費損益分析', font:'ui', size:26, weight:600, line:1.3 },
  { key:'page-kicker', label:'標題上方小標籤', hint:'breadcrumb 或頁面分類', sample:'成本分析 > 運費損益分析', font:'ui', size:13, weight:500, line:1.5 },
  { key:'nav', label:'導覽 / 側邊選單', hint:'側邊選單群組與頁面名稱', sample:'系統設定 / 文字樣式設定', font:'ui', size:14, weight:500, line:1.45 },
  { key:'control', label:'按鈕與篩選區', hint:'select、input、button、篩選 label', sample:'倉別 三倉總覽 月份 2026年03月', font:'ui', size:13, weight:500, line:1.45 },
  { key:'form-hint', label:'表單輔助文字', hint:'篩選列提示與表單說明', sample:'日期區間已鎖定為整個月份', font:'ui', size:12, weight:400, line:1.55 },
  { key:'widget-title', label:'元件標題', hint:'卡片、圖表與表格區塊標題', sample:'核心 KPI 總覽', font:'ui', size:16, weight:500, line:1.45 },
  { key:'widget-note', label:'元件說明小文字', hint:'meta、note、補充說明', sample:'依資料庫預算計算，資料每月更新', font:'ui', size:12, weight:400, line:1.55 },
  { key:'metric', label:'元件數字', hint:'KPI 大數字、金額、百分比', sample:'NT$14.53M', font:'number', size:32, weight:400, line:1.1 },
  { key:'table', label:'表格文字', hint:'表頭、儲存格與密集數值', sample:'費用項目  預算金額  實際動支', font:'ui', size:12, weight:400, line:1.5 },
  { key:'badge', label:'狀態標籤 / 圖表文字', hint:'badge、pill、圖表小標、空狀態標題', sample:'低於預算 · 使用健康', font:'ui', size:11, weight:500, line:1.4 },
];


function toggleAccordion(categoryEl) {
  const isExpanded = categoryEl.classList.contains('expanded');
  document.querySelectorAll('.sb-category.expanded').forEach(el => el.classList.remove('expanded'));
  if (!isExpanded) categoryEl.classList.add('expanded');
}

function getSidebarGroups() {
  const grouped = [];
  const byLabel = new Map();
  PAGES.forEach((group, index) => {
    const meta = SIDEBAR_GROUP_META[index] || { label: group.group, icon: '•' };
    if (!byLabel.has(meta.label)) {
      const next = { label: meta.label, icon: meta.icon, group: group.group, items: [] };
      grouped.push(next);
      byLabel.set(meta.label, next);
    }
    byLabel.get(meta.label).items.push(...group.items);
  });
  return grouped;
}

function renderSidebar() {
  const sb = document.getElementById('sidebar');
  if (!sb) return;
  const userId = sessionStorage.getItem('kpl_user') || 'User';
  const adminBadge = isAdmin() ? '<span class="sb-user-badge">ADMIN</span>' : '';
  const sidebarGroups = getSidebarGroups();

  let html = `
  <div class="sb-header">
    <div class="sb-header-brand">
      <div class="sb-brand-badge">RY</div>
      <div>
        <div class="sb-header-title">KPL 營運儀表板</div>
        <div class="sb-header-subtitle">Re-Yi Distribution</div>
      </div>
    </div>
    <button class="sb-pin" onclick="toggleSidebarPinned()" type="button" aria-label="釘選側欄" title="釘選/隱藏側欄">
      <svg class="sb-pin-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M14 4l6 6-3.5 1.2-2.9 4.7 2.1 2.1-1.7 1.7-4.1-4.1-4.2 4.2-1.5-1.5 4.2-4.2-4.1-4.1L6 8.3l2.1 2.1 4.7-2.9L14 4z" fill="currentColor"/>
      </svg>
    </button>
    <button class="sb-close" onclick="closeSidebar()" type="button" aria-label="關閉側欄">×</button>
  </div>
  <nav class="sb-nav">`;

  sidebarGroups.forEach(group => {
    const visibleItems = group.items.filter(item => {
      if (item.adminOnly) return isAdmin();
      return isPageVisible(item.id);
    });
    if (visibleItems.length === 0) return;

    const isActiveGroup = visibleItems.some(item => item.id === currentPageId);
    const expandedClass = isActiveGroup ? ' expanded' : '';
    const activeCatClass = isActiveGroup ? ' active' : '';

    html += `
    <div class="sb-category${expandedClass}">
      <button class="sb-category-btn${activeCatClass}" onclick="toggleAccordion(this.parentElement)" type="button">
        <span class="sb-category-main">
          <span class="sb-category-icon">${group.icon}</span>
          <span class="sb-category-label">${group.label}</span>
        </span>
      </button>
      <div class="sb-pages">`;

    visibleItems.forEach(item => {
      const active = item.id === currentPageId ? ' active' : '';
      let badge = '';
      if (item.status === 'wip') badge = '<span class="sb-item-badge wip">WIP</span>';
      else if (item.status === 'placeholder') badge = '<span class="sb-item-badge soon">TBD</span>';
      const hiddenMark = (isAdmin() && pagePermissions[item.id] === false && !item.adminOnly)
        ? '<span class="sb-item-badge muted">HIDDEN</span>' : '';
      const itemStyle = item.adminOnly ? ' data-admin-item="true"' : '';
      html += `
        <a href="#${item.id}" class="sb-item${active}" onclick="navigate(event,'${item.id}')"${itemStyle}>
          <span class="sb-item-icon" aria-hidden="true">${item.icon}</span>
          <span class="sb-item-label">${item.label}</span>
          ${badge}${hiddenMark}
        </a>`;
    });

    html += `
      </div>
    </div>`;
  });

  html += `
  </nav>
  <div class="sb-user">
    <div class="sb-user-avatar">${isAdmin() ? 'A' : userId.slice(0, 1).toUpperCase()}</div>
    <div class="sb-user-info">
      <div class="sb-user-name">${userId}${adminBadge}</div>
      <div class="sb-user-role">Re-Yi Distribution</div>
    </div>
    <button class="sb-logout" onclick="logout()" title="登出">⏻</button>
  </div>`;

  sb.innerHTML = html;
}

// ── 登出 ──
function logout() {
  fetch('/api/logout', { method: 'POST' }).catch(() => {});
  sessionStorage.removeItem('kpl_user');
  sessionStorage.removeItem('kpl_auth');
  location.href = 'login.html';
}

// ── 切換頁面 ──
function navigate(event, pageId) {
  if (event) event.preventDefault();
  currentPageId = pageId;
  renderSidebar();
  renderTopbarTabs();
  loadPage(pageId);
  if (isMobileLayout() || document.body.classList.contains('sidebar-collapsed')) {
    closeSidebar();
  }
}

function escapeReleaseText(value) {
  return String(value ?? '').replace(/[&<>"']/g, char => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  }[char]));
}

// 釋出版本連結僅允許 https GitHub URL，避免 javascript: 之類的偽協議被塞入 href
function safeReleaseUrl(value) {
  const s = String(value ?? '').trim();
  if (!/^https:\/\//i.test(s)) return '';
  return escapeReleaseText(s);
}

async function loadGithubReleases() {
  try {
    const res = await fetch('/api/releases', { cache: 'no-store' });
    if (handleAuthExpired(res)) return;
    const data = await res.json();
    if (res.ok && data.ok && Array.isArray(data.releases) && data.releases.length) {
      RELEASES = data.releases;
    }
  } catch {
    RELEASES = [...FALLBACK_RELEASES];
  }
}

function renderReleaseNotice() {
  const panel = document.getElementById('release-notice');
  const dot = document.getElementById('release-notif-dot');
  const latest = RELEASES[0];
  if (!panel || !latest) return;
  panel.innerHTML = `
    <div class="release-notice-kicker">最新版本 ${escapeReleaseText(latest.version)}</div>
    <div class="release-notice-title">${escapeReleaseText(latest.title)}</div>
    <div class="release-notice-summary">${escapeReleaseText(latest.summary)}</div>
    <button type="button" class="release-notice-link" onclick="openVersionInfo(event)">查看詳細版本資訊</button>`;
  if (dot) dot.hidden = localStorage.getItem(RELEASE_READ_KEY) === latest.version;
}

function toggleReleaseNotice() {
  const panel = document.getElementById('release-notice');
  if (!panel) return;
  panel.hidden = !panel.hidden;
  if (!panel.hidden && RELEASES[0]) {
    localStorage.setItem(RELEASE_READ_KEY, RELEASES[0].version);
    const dot = document.getElementById('release-notif-dot');
    if (dot) dot.hidden = true;
  }
}

function openVersionInfo(event) {
  if (event) event.stopPropagation();
  const panel = document.getElementById('release-notice');
  if (panel) panel.hidden = true;
  navigate(null, 'versions');
}

function renderVersionsPage() {
  const main = document.getElementById('main');
  main.innerHTML = `
    <div class="page-head">
      <div class="page-eyebrow">資料管理 &gt; 版本資訊</div>
      <h1 class="page-h">版本資訊</h1>
    </div>
    <div class="release-history">
      ${RELEASES.map((release, index) => `
        <article class="release-entry">
          <div class="release-entry-meta">
            <span class="release-version">${escapeReleaseText(release.version)}</span>
            <time datetime="${escapeReleaseText(release.date)}">${escapeReleaseText(release.date)}</time>
            ${index === 0 ? '<span class="release-latest">最新版本</span>' : ''}
          </div>
          <h2>${escapeReleaseText(release.title)}</h2>
          <p>${escapeReleaseText(release.summary)}</p>
          <ul>${release.changes.map(change => `<li>${escapeReleaseText(change)}</li>`).join('')}</ul>
          ${(() => { const u = safeReleaseUrl(release.url); return u ? `<a class="release-github-link" href="${u}" target="_blank" rel="noopener noreferrer">在 GitHub 查看 Release</a>` : ''; })()}
        </article>`).join('')}
    </div>`;
}

// ── 更新 Topbar 頁面名稱 ──
function updateTopbarPageName(pageId) {
  const el = document.getElementById('topbar-page-name');
  const group = PAGES.find(g => g.items.some(i => i.id === pageId));
  const page = group?.items.find(i => i.id === pageId);
  if (page) {
    if (el) el.textContent = `${page.icon} ${page.label}`;
    const groupIcon = document.getElementById('topbar-breadcrumb-icon');
    const groupEl = document.getElementById('topbar-breadcrumb-group');
    const pageEl = document.getElementById('topbar-breadcrumb-page');
    if (groupIcon) groupIcon.textContent = group?.group.match(/^\S+/)?.[0] || page.icon || '📊';
    if (groupEl) groupEl.textContent = (group?.group || '').replace(/^\S+\s*/, '') || '儀表板總覽';
    if (pageEl) pageEl.textContent = page.label;
    document.title = `${page.label} · KPL 儀表板`;
  } else {
    if (el) el.textContent = '';
    const groupEl = document.getElementById('topbar-breadcrumb-group');
    const pageEl = document.getElementById('topbar-breadcrumb-page');
    if (groupEl) groupEl.textContent = 'KPL';
    if (pageEl) pageEl.textContent = '營運儀表板';
    document.title = 'KPL 儀表板';
  }
}

function updatePageBreadcrumb(pageId) {
  const group = getCurrentGroup(pageId);
  const page = group?.items.find(item => item.id === pageId);
  const el = document.querySelector('#main .page-eyebrow');
  if (!el || !group || !page) return;
  el.textContent = `${group.group} > ${page.label}`;
}

// ── 載入頁面 ──
function loadPage(pageId) {
  updateTopbarPageName(pageId);

  // 舊 admin 路由轉址到新頁面
  if (pageId === 'admin') {
    navigate(null, 'accountPermissions');
    return;
  }

  if (pageId.startsWith('account')) {
    if (!isAdmin()) { navigate(null, 'daily'); return; }
    renderAccountManagementPage(pageId);
    return;
  }

  // 非管理員：檢查頁面是否有權限
  if (!isAdmin() && pagePermissions[pageId] === false) {
    navigate(null, 'daily');
    return;
  }

  if (pageId === 'versions') {
    renderVersionsPage();
    return;
  }

  const main = document.getElementById('main');
  const tpl = PAGE_TEMPLATES[pageId];
  if (!tpl) {
    main.innerHTML = '<div class="wip-page"><div class="wip-icon">⚠️</div><div class="wip-title">找不到頁面</div></div>';
    return;
  }
  main.innerHTML = tpl;
  updatePageBreadcrumb(pageId);

  if (pageId === 'daily')        initDailyPage();
  else if (pageId === 'dispatch') initDispatchPage();
  else if (pageId === 'freight') initFreightPage();
  else if (pageId === 'picks')   initPicksPage();
  else if (pageId === 'labor')   initLaborPage();
  else if (pageId === 'import')  initImportPage();
  else if (pageId === 'org')     initOrgPage();
  else if (pageId === 'typography') initTypographyPage();
  else if (pageId === 'productivity') initProductivityPage();
  else if (pageId === 'monthly') initMonthlyPage();
  else if (pageId === 'annual')  renderAnnualPage();

  normalizeDateFilterBars();
}

// ── 各頁面初始化 ──
async function initDailyPage() {
  renderDailyPage();
  DATA.dailySummary.dateFrom = DATA.dateFrom;
  DATA.dailySummary.dateTo = DATA.dateTo;
  DATA.dailySummary.laborRows = [];
  DATA.dailySummary.freightRows = [];
  await Promise.all([
    loadCloudBudgetData(),
    loadCloudLaborData({ summary: true }),
    loadCloudFreightData({ summary: true }),
  ]);
  if (currentPageId === 'daily') renderDailyPage();
}

async function initDispatchPage() {
  renderDispatchPage();
  await Promise.all([
    loadCloudBudgetData(),
    loadCloudLaborData({ summary: true }),
    loadCloudFreightData({ summary: true }),
  ]);
  if (currentPageId === 'dispatch') renderDispatchPage();
}

async function initFreightPage() {
  renderFreightPage();
  await Promise.all([loadCloudBudgetData(), loadCloudFreightData()]);
  if (currentPageId === 'freight') renderFreightPage();
}
async function initImportPage() {
  await loadCloudDataRange();
  updateStatus();
}

const DASHBOARD_DATE_FILTERS = {
  daily:        { from:'filter-from',        to:'filter-to',        meta:'filter-meta' },
  dispatch:     { from:'dispatch-from',     to:'dispatch-to',     meta:null },
  freight:      { from:'freight-from',      to:'freight-to',      meta:null },
  picks:        { from:'picks-from',        to:'picks-to',        meta:'picks-date-meta' },
  labor:        { from:'labor-from',        to:'labor-to',        meta:'labor-date-meta' },
  productivity: { from:'productivity-from', to:'productivity-to', meta:'productivity-date-meta' },
  monthly:      { from:'monthly-from',      to:'monthly-to',      meta:'monthly-date-meta' },
};

function normalizeDateFilterBars() {
  document.querySelectorAll('#main .filter-bar').forEach(bar => {
    if (!bar.querySelector('input[type="date"]')) {
      bar.classList.remove('date-range-filter');
      return;
    }
    bar.classList.add('date-range-filter');
    const separator = bar.querySelector('.range-arrow');
    if (separator) separator.textContent = '→';
    const meta = bar.querySelector('.filter-meta');
    if (meta) meta.textContent = '日期區間已鎖定為整個月份';
  });
}

function syncDashboardDateInputs(pageId = currentPageId) {
  const cfg = DASHBOARD_DATE_FILTERS[pageId];
  if (!cfg) return;
  const from = document.getElementById(cfg.from);
  const to = document.getElementById(cfg.to);
  if (from) from.value = DATA.dateFrom;
  if (to) to.value = DATA.dateTo;
  if (cfg.meta) {
    const meta = document.getElementById(cfg.meta);
    if (meta) meta.textContent = '日期區間已鎖定為整個月份';
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

const ACCOUNT_MANAGEMENT_PAGES = {
  accountIpRules: {
    icon: '🛡️',
    title: '後臺限制 IP 管理',
  },
  accountLoginAudit: {
    icon: '📜',
    title: '系統登入日誌',
    desc: '查詢登入成功、登入失敗、頻率限制與 Session 到期事件。',
    columns: ['時間', '帳號', '來源 IP', '結果', '訊息'],
    rows: [],
    next: ['建立 login_audit_logs 資料表', '把 /api/check-user 成功與失敗寫入日誌', '提供日期、帳號、IP 篩選'],
  },
  accountErrorAudit: {
    icon: '🚨',
    title: '全域錯誤例外日誌',
    desc: '集中查看後端 API 例外、前端錯誤與資料匯入失敗。',
    columns: ['時間', '來源', '嚴重度', '錯誤摘要', '處理狀態'],
    rows: [],
    next: ['建立 error_audit_logs 資料表', '捕捉 server.js API 例外', '前端加入 window error / unhandledrejection 回報'],
  },
};

function renderAccountManagementPage(pageId) {
  if (!isAdmin()) return;
  const main = document.getElementById('main');
  if (!main) return;

  if (pageId === 'accountPermissions') {
    const allItems = PAGES.flatMap(g => g.items).filter(item => !item.adminOnly);
    const rows = allItems.map(item => {
      const isOn = pagePermissions[item.id] !== false;
      return `<tr>
        <td>${item.icon} ${item.label}</td>
        <td class="mono text-muted-strong">${item.id}</td>
        <td>
          <label class="admin-toggle">
            <input type="checkbox" id="perm-${item.id}" ${isOn ? 'checked' : ''}>
            <span class="admin-toggle-slider"></span>
          </label>
        </td>
      </tr>`;
    }).join('');

    main.innerHTML = `
      <div class="page-head">
        <div class="page-eyebrow">帳號權限管理 &gt; 帳號權限</div>
        <h1 class="page-h">👥 頁面權限設定</h1>
      </div>
      <div class="grid">
        <div class="w s12">
          <div class="wh">
            <div class="wl"><div class="wdot"></div>一般帳號可見頁面</div>
            <span class="wmeta">管理員設定</span>
          </div>
          <div style="padding:12px 18px 4px;font-size:13px;color:var(--app-muted)">
            控制一般帳號可以看到哪些頁面。管理員帳號（inari）永遠可以看到全部頁面。
          </div>
          <div class="admin-permission-table-frame">
            <table class="ops-compact-table admin-permission-table" style="width:100%;border-collapse:collapse">
              <thead><tr><th>頁面名稱</th><th>頁面 ID</th><th>一般帳號可見</th></tr></thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </div>

        <div class="w s12">
          <div class="wh">
            <div class="wl"><div class="wdot"></div>儲存設定</div>
          </div>
          <div id="admin-save-area" style="padding:16px 18px">
            <div style="margin-bottom:12px;padding:12px 14px;background:#fff8e6;border-radius:8px;border:1px solid #f5c400;font-size:13px;color:#7a6000">
              ⚠️ 儲存前需輸入您的密碼重新驗證身份，確保安全。
            </div>
            <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap">
              <div style="flex:1;min-width:200px">
                <label style="display:block;font-size:11px;font-weight:700;color:var(--app-muted);margin-bottom:6px;letter-spacing:.08em">確認密碼</label>
                <input type="password" id="admin-psw" placeholder="請輸入您的登入密碼"
                  style="width:100%;padding:10px 14px;border:1.5px solid var(--app-border);border-radius:4px;font-size:13px;background:var(--app-surface);color:var(--app-text)">
              </div>
              <button onclick="savePermissions()"
                style="padding:10px 28px;background:#1e5ca8;color:#fff;border:none;border-radius:4px;font-size:14px;font-weight:700;cursor:pointer;white-space:nowrap">
                💾 儲存設定
              </button>
            </div>
            <div id="admin-msg" style="margin-top:10px;font-size:13px;display:none"></div>
          </div>
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
    return;
  }

  if (pageId === 'accountIpRules') { renderAccountIpRulesPage(); return; }

  const config = ACCOUNT_MANAGEMENT_PAGES[pageId];
  if (!config) return;
  const rows = config.rows.length
    ? config.rows.map(row => `<tr>${row.map(cell => `<td>${escapeReleaseText(cell)}</td>`).join('')}</tr>`).join('')
    : `<tr><td colspan="${config.columns.length}" style="text-align:center;color:var(--app-muted);padding:28px 12px">尚未接 Cloud SQL，沒有可稽核資料。</td></tr>`;

  main.innerHTML = `
    <div class="page-head">
      <div class="page-eyebrow">帳號權限管理 &gt; ${escapeReleaseText(config.title)}</div>
      <h1 class="page-h">${config.icon} ${escapeReleaseText(config.title)}</h1>
    </div>
    <div class="grid">
      <div class="w s12">
        <div class="wh">
          <div class="wl"><div class="wdot"></div>管理範圍</div>
          <span class="wmeta">管理員設定</span>
        </div>
        <div style="padding:18px;color:var(--app-muted);font-size:13px;line-height:1.8">
          ${escapeReleaseText(config.desc)}
          <div style="margin-top:12px;padding:12px 14px;border:1px solid var(--app-border);border-left:3px solid var(--app-accent);border-radius:8px;background:var(--app-surface-soft);color:var(--app-text)">
            第一版先補齊管理入口與資訊架構。涉及帳號、IP、稽核日誌的正式資料保存需 Cloud SQL schema，需少佐另外同意 migration 後才會實作。
          </div>
        </div>
      </div>
      <div class="w s12">
        <div class="wh"><div class="wl"><div class="wdot"></div>目前資料</div></div>
        <div class="admin-permission-table-frame">
          <table class="ops-compact-table admin-permission-table" style="width:100%;border-collapse:collapse">
            <thead><tr>${config.columns.map(col => `<th>${escapeReleaseText(col)}</th>`).join('')}</tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>
      <div class="w s12">
        <div class="wh"><div class="wl"><div class="wdot"></div>後續需接上的正式能力</div></div>
        <div style="padding:16px 18px">
          <ul style="margin-left:18px;color:var(--app-muted);font-size:13px;line-height:1.9">
            ${config.next.map(item => `<li>${escapeReleaseText(item)}</li>`).join('')}
          </ul>
        </div>
      </div>
    </div>`;
}

// ════════════════════════════════════════════
// 🛡️ 後臺限制 IP 管理（黑名單 CRUD）
// ════════════════════════════════════════════

function renderAccountIpRulesPage() {
  const main = document.getElementById('main');
  if (!main) return;
  main.innerHTML = `
    <div class="page-head">
      <div class="page-eyebrow">帳號權限管理 &gt; 後臺限制 IP 管理</div>
      <h1 class="page-h">🛡️ 後臺限制 IP 管理</h1>
    </div>
    <div class="grid">
      <div class="w s12">
        <div style="padding:10px 18px 0;color:var(--app-muted);font-size:13px;line-height:1.8">
          黑名單邏輯：清單中的 IP 無法呼叫任何 API（靜態頁面不受影響）。支援 IPv4 單一位址或 CIDR 格式（例：10.0.0.0/8）。
        </div>
        <div id="ip-my-ip-banner" style="margin:10px 18px 0;padding:10px 14px;border-radius:8px;border:1px solid var(--app-warning);background:var(--app-warning-soft);font-size:13px;color:var(--app-text)">
          ⚠️ 您目前的來源 IP：<strong id="ip-my-ip" style="font-family:monospace">偵測中…</strong>
          &ensp;—&ensp;封鎖自己的 IP 或 CIDR 後，Session 到期前可進入管理頁解除；Session 到期後將無法登入，需透過 Cloud SQL 直接刪除記錄。
        </div>
      </div>

      <div class="w s12">
        <div class="wh">
          <div class="wl"><div class="wdot"></div>新增封鎖 IP</div>
        </div>
        <div style="padding:16px 18px">
          <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap">
            <div style="flex:0 0 210px">
              <label style="display:block;font-size:11px;font-weight:700;color:var(--app-muted);margin-bottom:6px;letter-spacing:.08em">IP 位址 / CIDR</label>
              <input type="text" id="ip-new-cidr" placeholder="192.168.1.1 或 10.0.0.0/8"
                style="width:100%;padding:10px 14px;border:1.5px solid var(--app-border);border-radius:4px;font-size:13px;background:var(--app-surface);color:var(--app-text);font-family:monospace">
            </div>
            <div style="flex:1;min-width:160px">
              <label style="display:block;font-size:11px;font-weight:700;color:var(--app-muted);margin-bottom:6px;letter-spacing:.08em">備註說明</label>
              <input type="text" id="ip-new-label" placeholder="例：惡意爬蟲、異常流量"
                style="width:100%;padding:10px 14px;border:1.5px solid var(--app-border);border-radius:4px;font-size:13px;background:var(--app-surface);color:var(--app-text)">
            </div>
            <button onclick="addIpRule()"
              style="padding:10px 22px;background:var(--app-danger);color:#fff;border:none;border-radius:4px;font-size:13px;font-weight:700;cursor:pointer;white-space:nowrap">
              🚫 封鎖此 IP
            </button>
          </div>
          <div id="ip-rules-msg" style="margin-top:10px;font-size:13px;display:none"></div>
        </div>
      </div>

      <div class="w s12">
        <div class="wh">
          <div class="wl"><div class="wdot"></div>封鎖清單</div>
          <span class="wmeta" id="ip-rules-count">載入中…</span>
        </div>
        <div class="admin-permission-table-frame" id="ip-rules-table-frame">
          <div style="padding:28px;text-align:center;color:var(--app-muted);font-size:13px">載入中…</div>
        </div>
      </div>
    </div>`;

  loadIpRules();
  fetch('/api/admin/my-ip').then(r => r.json()).then(d => {
    const el = document.getElementById('ip-my-ip');
    if (el && d.ok) el.textContent = d.ip;
  }).catch(() => {});
}

async function loadIpRules() {
  try {
    const res = await fetch('/api/admin/ip-rules');
    const data = await res.json();
    if (!data.ok) throw new Error(data.MSG);
    renderIpRulesTable(data.rules);
  } catch (err) {
    const frame = document.getElementById('ip-rules-table-frame');
    if (frame) frame.innerHTML = `<div style="padding:28px;text-align:center;color:var(--app-danger);font-size:13px">⚠️ 載入失敗：${escapeReleaseText(err.message)}</div>`;
  }
}

function renderIpRulesTable(rules) {
  const countEl = document.getElementById('ip-rules-count');
  const frame = document.getElementById('ip-rules-table-frame');
  if (!frame) return;
  const active = rules.filter(r => r.is_active).length;
  if (countEl) countEl.textContent = `共 ${rules.length} 條・封鎖中 ${active} 條`;

  if (rules.length === 0) {
    frame.innerHTML = `<div style="padding:28px;text-align:center;color:var(--app-muted);font-size:13px">目前無封鎖規則</div>`;
    return;
  }

  const rows = rules.map(r => {
    const ts = new Date(r.created_at).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei', hour12: false });
    return `<tr>
      <td class="mono" style="font-size:13px">${escapeReleaseText(r.ip_cidr)}</td>
      <td style="color:var(--app-muted)">${escapeReleaseText(r.label || '—')}</td>
      <td>
        <label class="admin-toggle" title="${r.is_active ? '點擊停用' : '點擊啟用'}">
          <input type="checkbox" ${r.is_active ? 'checked' : ''} onchange="toggleIpRule(${r.id}, this.checked)">
          <span class="admin-toggle-slider"></span>
        </label>
        <span style="margin-left:8px;font-size:12px;color:${r.is_active ? 'var(--app-danger)' : 'var(--app-muted)'}">${r.is_active ? '封鎖中' : '已停用'}</span>
      </td>
      <td style="color:var(--app-muted);font-size:12px">${escapeReleaseText(r.created_by)}</td>
      <td style="color:var(--app-muted);font-size:12px;white-space:nowrap">${ts}</td>
      <td>
        <button onclick="deleteIpRule(${r.id})"
          style="padding:4px 12px;background:var(--app-danger-soft);color:var(--app-danger);border:1px solid var(--app-danger);border-radius:4px;font-size:12px;cursor:pointer">
          刪除
        </button>
      </td>
    </tr>`;
  }).join('');

  frame.innerHTML = `
    <table class="ops-compact-table admin-permission-table" style="width:100%;border-collapse:collapse">
      <thead>
        <tr>
          <th>IP / CIDR</th><th>備註</th><th>狀態</th>
          <th>建立者</th><th>建立時間</th><th>操作</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;
}

async function addIpRule() {
  const cidrEl  = document.getElementById('ip-new-cidr');
  const labelEl = document.getElementById('ip-new-label');
  const ipCidr  = cidrEl?.value.trim() || '';
  const label   = labelEl?.value.trim() || '';
  if (!ipCidr) { showIpRulesMsg('❌ 請輸入 IP 或 CIDR', 'error'); cidrEl?.focus(); return; }

  try {
    const res = await fetch('/api/admin/ip-rules', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ip_cidr: ipCidr, label }),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.MSG);
    const msg = data.cacheStale ? '✅ 已新增（DB 快取暫時無法更新，請稍後重整）' : '✅ 已新增封鎖規則';
    showIpRulesMsg(msg, 'ok');
    if (cidrEl)  cidrEl.value  = '';
    if (labelEl) labelEl.value = '';
    await loadIpRules();
  } catch (err) {
    showIpRulesMsg(`❌ 新增失敗：${err.message}`, 'error');
  }
}

async function deleteIpRule(id) {
  if (!confirm('確定刪除此封鎖規則？刪除後立即生效，該 IP 將可重新存取。')) return;
  try {
    const res = await fetch(`/api/admin/ip-rules/${id}`, { method: 'DELETE' });
    const data = await res.json();
    if (!data.ok) throw new Error(data.MSG);
    const msg = data.cacheStale ? '✅ 已刪除（DB 快取暫時無法更新，請稍後重整）' : '✅ 已刪除';
    showIpRulesMsg(msg, 'ok');
    await loadIpRules();
  } catch (err) {
    showIpRulesMsg(`❌ 刪除失敗：${err.message}`, 'error');
  }
}

async function toggleIpRule(id, isActive) {
  try {
    const res = await fetch(`/api/admin/ip-rules/${id}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ is_active: isActive }),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.MSG);
    const base = isActive ? '✅ 已啟用封鎖' : '✅ 已停用（該 IP 可正常存取）';
    showIpRulesMsg(data.cacheStale ? `${base}（快取暫時無法更新）` : base, 'ok');
    await loadIpRules();
  } catch (err) {
    showIpRulesMsg(`❌ 操作失敗：${err.message}`, 'error');
    await loadIpRules();
  }
}

function showIpRulesMsg(text, type) {
  const el = document.getElementById('ip-rules-msg');
  if (!el) return;
  el.style.display = 'block';
  el.style.color = type === 'error' ? 'var(--app-danger)' : 'var(--app-success)';
  el.textContent = text;
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.style.display = 'none'; }, 4000);
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
    if (handleAuthExpired(res)) return;
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

function defaultTypographySettings() {
  return TYPOGRAPHY_CONTROLS.reduce((map, item) => {
    map[item.key] = {
      font: item.font,
      size: item.size,
      weight: item.weight,
      line: item.line,
    };
    return map;
  }, {});
}

function loadTypographySettings() {
  const defaults = defaultTypographySettings();
  try {
    if (localStorage.getItem(TYPOGRAPHY_VERSION_KEY) !== TYPOGRAPHY_VERSION) {
      localStorage.setItem(TYPOGRAPHY_STORAGE_KEY, JSON.stringify(defaults));
      localStorage.setItem(TYPOGRAPHY_VERSION_KEY, TYPOGRAPHY_VERSION);
      return defaults;
    }
    const saved = JSON.parse(localStorage.getItem(TYPOGRAPHY_STORAGE_KEY) || '{}');
    TYPOGRAPHY_CONTROLS.forEach(item => {
      const row = saved[item.key] || {};
      defaults[item.key] = {
        font: TYPOGRAPHY_FONT_MAP[row.font] ? row.font : defaults[item.key].font,
        size: Number(row.size) || defaults[item.key].size,
        weight: Number(row.weight) || defaults[item.key].weight,
        line: Number(row.line) || defaults[item.key].line,
      };
    });
  } catch {
    return defaults;
  }
  return defaults;
}

function applyTypographySettings(settings = loadTypographySettings()) {
  const root = document.documentElement;
  TYPOGRAPHY_CONTROLS.forEach(item => {
    const row = settings[item.key] || item;
    root.style.setProperty(`--typo-${item.key}-font`, TYPOGRAPHY_FONT_MAP[row.font] || TYPOGRAPHY_FONT_MAP.ui);
    root.style.setProperty(`--typo-${item.key}-size`, `${Number(row.size) || item.size}px`);
    root.style.setProperty(`--typo-${item.key}-weight`, `${Number(row.weight) || item.weight}`);
    root.style.setProperty(`--typo-${item.key}-line`, `${Number(row.line) || item.line}`);
  });
}

let typographyDraft = null;

function initTypographyPage() {
  typographyDraft = loadTypographySettings();
  applyTypographySettings(typographyDraft);
  renderTypographyPage();
}

function renderTypographyPage() {
  const grid = document.getElementById('typography-grid');
  if (!grid) return;
  const rows = TYPOGRAPHY_CONTROLS.map(item => {
    const row = typographyDraft[item.key];
    return `
      <div class="typo-row">
        <div class="typo-row-info">
          <div class="typo-row-title">${item.label}</div>
          <div class="typo-row-hint">${item.hint}</div>
          <div class="typo-row-sample" id="typo-sample-${item.key}" style="font-family:${TYPOGRAPHY_FONT_MAP[row.font]};font-size:${row.size}px;font-weight:${row.weight};line-height:${row.line}">${item.sample}</div>
        </div>
        <label class="typo-field">
          <span>字體</span>
          <select class="filter-input" onchange="updateTypographyDraft('${item.key}','font',this.value)">
            <option value="ui" ${row.font === 'ui' ? 'selected' : ''}>中文介面</option>
            <option value="number" ${row.font === 'number' ? 'selected' : ''}>數字字體</option>
          </select>
        </label>
        <label class="typo-field">
          <span>字級</span>
          <input class="filter-input" type="number" min="9" max="56" step="1" value="${row.size}" oninput="updateTypographyDraft('${item.key}','size',this.value)">
        </label>
        <label class="typo-field">
          <span>字重</span>
          <select class="filter-input" onchange="updateTypographyDraft('${item.key}','weight',this.value)">
            ${[400,500,600,700,800,900].map(w => `<option value="${w}" ${Number(row.weight) === w ? 'selected' : ''}>${w}</option>`).join('')}
          </select>
        </label>
        <label class="typo-field">
          <span>行高</span>
          <input class="filter-input" type="number" min="1" max="2" step="0.05" value="${row.line}" oninput="updateTypographyDraft('${item.key}','line',this.value)">
        </label>
      </div>`;
  }).join('');

  grid.innerHTML = `
    <section class="w s7 typography-panel">
      <div class="wh">
        <div class="wl"><div class="wdot"></div>文字層級設定</div>
        <span class="wmeta">localStorage 保存</span>
      </div>
      <div class="typo-list">${rows}</div>
    </section>
    <section class="w s5 typography-preview-panel">
      <div class="wh">
        <div class="wl"><div class="wdot dot-picks"></div>即時預覽</div>
        <span class="wmeta">全站 token</span>
      </div>
      <div id="typography-preview">${renderTypographyPreview()}</div>
    </section>`;
}

function renderTypographyPreview() {
  return `
    <div class="typo-preview-page-kicker">成本分析 &gt; 運費損益分析</div>
    <div class="typo-preview-page-title">運費損益分析</div>
    <div class="typo-preview-nav">系統設定 / 文字樣式設定</div>
    <div class="typo-preview-filter">
      <span>倉別</span>
      <button type="button">三倉總覽</button>
      <span>月份</span>
      <button type="button">2026年03月</button>
      <button type="button" class="primary">套用</button>
    </div>
    <div class="typo-preview-hint">日期區間已鎖定為整個月份</div>
    <div class="typo-preview-card">
      <div class="typo-preview-widget-title">核心 KPI 總覽</div>
      <div class="typo-preview-widget-note">依資料庫預算計算，資料每月更新</div>
      <div class="typo-preview-metric">NT$14.53M</div>
      <span class="typo-preview-badge">低於預算 · 使用健康</span>
    </div>
    <table class="typo-preview-table">
      <thead><tr><th>費用項目</th><th>預算金額</th><th>實際動支</th></tr></thead>
      <tbody><tr><td>主線</td><td>33.20M</td><td>11.56M</td></tr></tbody>
    </table>
    <div class="typo-preview-empty">
      <div class="typo-preview-widget-title">尚未匯入資料</div>
      <div class="typo-preview-widget-note">請先到資料匯入上傳資料，套用後此頁會自動產生指標。</div>
    </div>`;
}

function updateTypographyDraft(key, prop, value) {
  if (!typographyDraft?.[key]) return;
  typographyDraft[key][prop] = prop === 'font' ? value : Number(value);
  applyTypographySettings(typographyDraft);
  const sample = document.getElementById(`typo-sample-${key}`);
  if (sample) {
    const row = typographyDraft[key];
    sample.style.fontFamily = TYPOGRAPHY_FONT_MAP[row.font] || TYPOGRAPHY_FONT_MAP.ui;
    sample.style.fontSize = `${row.size}px`;
    sample.style.fontWeight = row.weight;
    sample.style.lineHeight = row.line;
  }
  const preview = document.getElementById('typography-preview');
  if (preview) preview.innerHTML = renderTypographyPreview();
}

function saveTypographySettings() {
  const settings = typographyDraft || loadTypographySettings();
  localStorage.setItem(TYPOGRAPHY_STORAGE_KEY, JSON.stringify(settings));
  localStorage.setItem(TYPOGRAPHY_VERSION_KEY, TYPOGRAPHY_VERSION);
  applyTypographySettings(settings);
  toast('✅ 文字樣式設定已儲存');
}

function resetTypographySettings() {
  typographyDraft = defaultTypographySettings();
  localStorage.removeItem(TYPOGRAPHY_STORAGE_KEY);
  localStorage.setItem(TYPOGRAPHY_VERSION_KEY, TYPOGRAPHY_VERSION);
  applyTypographySettings(typographyDraft);
  renderTypographyPage();
  toast('↺ 已恢復文字樣式預設值');
}

function rerenderDashboardPage(pageId = currentPageId) {
  if (pageId === 'daily') renderDailyPage();
  else if (pageId === 'dispatch') renderDispatchPage();
  else if (pageId === 'freight') renderFreightPage();
  else if (pageId === 'picks') renderPicksPage();
  else if (pageId === 'labor') renderLaborPage();
  else if (pageId === 'productivity') renderProductivityPage();
  else if (pageId === 'monthly') renderMonthlyPage();
  else if (pageId === 'annual') renderAnnualPage();
  normalizeDateFilterBars();
}

async function applyDashboardDateFilter(pageId = currentPageId) {
  if (!setSharedDateRangeFromInputs(pageId)) return;
  if (pageId === 'daily' || pageId === 'dispatch') {
    if (pageId === 'daily') {
      DATA.dailySummary.dateFrom = DATA.dateFrom;
      DATA.dailySummary.dateTo = DATA.dateTo;
      DATA.dailySummary.laborRows = [];
      DATA.dailySummary.freightRows = [];
    }
    await Promise.all([
      loadCloudBudgetData(),
      loadCloudLaborData({ summary: true }),
      loadCloudFreightData({ summary: true }),
    ]);
  } else {
    await Promise.all([
      loadCloudBudgetData(),
      loadCloudLaborData(),
      loadCloudPicksData(),
      loadCloudFreightData(),
    ]);
  }
  if (pageId === 'dispatch') syncDispatchBudgetForCurrentMonth();
  rerenderDashboardPage(pageId);
  toast('🔄 日期區間已更新');
}

// ── Drawer 開關（桌機與手機統一行為）──
function toggleSidebar() {
  if (isMobileLayout()) {
    const isOpen = document.getElementById('sidebar').classList.toggle('drawer-open');
    document.body.classList.toggle('sidebar-mobile-open', isOpen);
    return;
  }
  setSidebarPinned(!document.body.classList.contains('sidebar-pinned'));
}

function closeSidebar() {
  document.getElementById('sidebar').classList.remove('drawer-open');
  document.body.classList.remove('sidebar-mobile-open', 'sidebar-peek');
}

function isMobileLayout() {
  return window.matchMedia('(max-width: 768px)').matches;
}

function setSidebarPinned(isPinned) {
  document.body.classList.toggle('sidebar-pinned', isPinned);
  document.body.classList.toggle('sidebar-collapsed', !isPinned);
  document.body.classList.remove('sidebar-peek');
  localStorage.setItem(SIDEBAR_PINNED_KEY, isPinned ? '1' : '0');
}

function toggleSidebarPinned() {
  setSidebarPinned(!document.body.classList.contains('sidebar-pinned'));
}

// ── 側邊欄寬度：可拖拉調整、記憶於 localStorage ──
const SIDEBAR_WIDTH_KEY = 'kpl_sidebar_width';
const SIDEBAR_WIDTH_MIN = 200;
const SIDEBAR_WIDTH_MAX = 440;
const SIDEBAR_WIDTH_DEFAULT = 260;

function applySidebarWidth(px) {
  const w = Math.min(SIDEBAR_WIDTH_MAX, Math.max(SIDEBAR_WIDTH_MIN, Math.round(px)));
  document.documentElement.style.setProperty('--drawer-width', `${w}px`);
  return w;
}

function initSidebarResize() {
  const handle = document.getElementById('sidebar-resize-handle');
  if (!handle) return;
  const saved = Number(localStorage.getItem(SIDEBAR_WIDTH_KEY));
  if (saved >= SIDEBAR_WIDTH_MIN && saved <= SIDEBAR_WIDTH_MAX) applySidebarWidth(saved);

  let dragging = false;
  const onMove = (e) => {
    if (!dragging) return;
    const x = e.touches ? e.touches[0].clientX : e.clientX;
    const sidebar = document.getElementById('sidebar');
    const sidebarLeft = sidebar ? sidebar.getBoundingClientRect().left : 0;
    applySidebarWidth(x - sidebarLeft);
    if (e.cancelable) e.preventDefault();
  };
  const onUp = () => {
    if (!dragging) return;
    dragging = false;
    document.body.classList.remove('sidebar-resizing');
    const w = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--drawer-width'), 10);
    if (w) localStorage.setItem(SIDEBAR_WIDTH_KEY, String(w));
  };
  const onDown = (e) => {
    dragging = true;
    document.body.classList.add('sidebar-resizing');
    if (e.cancelable) e.preventDefault();
  };
  handle.addEventListener('mousedown', onDown);
  handle.addEventListener('touchstart', onDown, { passive: false });
  window.addEventListener('mousemove', onMove);
  window.addEventListener('touchmove', onMove, { passive: false });
  window.addEventListener('mouseup', onUp);
  window.addEventListener('touchend', onUp);
  // 雙擊把手還原預設寬度
  handle.addEventListener('dblclick', () => {
    applySidebarWidth(SIDEBAR_WIDTH_DEFAULT);
    localStorage.setItem(SIDEBAR_WIDTH_KEY, String(SIDEBAR_WIDTH_DEFAULT));
  });
}

function openSidebarPeek() {
  if (!isMobileLayout() && document.body.classList.contains('sidebar-collapsed')) {
    document.body.classList.add('sidebar-peek');
  }
}

function closeSidebarPeek() {
  if (!isMobileLayout()) document.body.classList.remove('sidebar-peek');
}

function initSidebarState() {
  const saved = localStorage.getItem(SIDEBAR_PINNED_KEY);
  setSidebarPinned(saved !== '0');
  const zone = document.getElementById('sidebar-hover-zone');
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebar-overlay');
  if (zone) zone.addEventListener('mouseenter', openSidebarPeek);
  if (sidebar) {
    sidebar.addEventListener('mouseenter', openSidebarPeek);
    sidebar.addEventListener('mouseleave', closeSidebarPeek);
  }
  if (overlay) overlay.addEventListener('click', closeSidebar);
}

function setThemeMode(mode) {
  const nextMode = THEME_MODES.includes(mode) ? mode : 'original';
  document.documentElement.dataset.theme = nextMode;
  localStorage.setItem(THEME_STORAGE_KEY, nextMode);
  document.querySelectorAll('[data-theme-option]').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.themeOption === nextMode);
  });
}

function initThemeMode() {
  setThemeMode(localStorage.getItem(THEME_STORAGE_KEY) || 'original');
}

// ── Topbar：分類 Tab 導覽 ──
function getCurrentGroup(pageId) {
  return PAGES.find(g => g.items.some(i => i.id === pageId)) || null;
}

function navigateGroup(groupIdx) {
  const group = PAGES[groupIdx];
  if (!group) return;
  const firstVisible = group.items.find(item => {
    if (item.adminOnly) return isAdmin();
    return isPageVisible(item.id);
  });
  if (firstVisible) navigate(null, firstVisible.id);
}

function renderTopbarTabs() {
  const container = document.getElementById('topbar-tabs');
  if (!container) return;
  const currentGroup = getCurrentGroup(currentPageId);
  let html = '';
  PAGES.forEach((group, idx) => {
    const visibleItems = group.items.filter(item => {
      if (item.adminOnly) return isAdmin();
      return isPageVisible(item.id);
    });
    if (visibleItems.length === 0) return;
    const active = currentGroup && currentGroup.group === group.group ? ' active' : '';
    let items = '';
    visibleItems.forEach(item => {
      const ia = item.id === currentPageId ? ' active' : '';
      items += `<button class="topbar-dropdown-item${ia}" onclick="navigateTopbarItem(event,'${item.id}')" type="button">${item.icon} ${item.label}</button>`;
    });
    html += `<div class="topbar-tab-wrap" onmouseenter="ddShow(this)" onmouseleave="ddHide()"><button class="topbar-tab${active}" onclick="onTabClick(event,${idx})" type="button">${group.group}</button><div class="topbar-dropdown" onmouseenter="ddCancel()" onmouseleave="ddHide()">${items}</div></div>`;
  });
  container.innerHTML = html;
}

// ── Topbar Dropdown：hover 展開子選單 ──
let _ddT = null;

function ddShow(wrapEl) {
  clearTimeout(_ddT);
  document.querySelectorAll('.topbar-dropdown.dd-open').forEach(d => d.classList.remove('dd-open'));
  const dd = wrapEl.querySelector('.topbar-dropdown');
  if (!dd) return;
  dd.classList.add('dd-open');
}

function ddHide() {
  _ddT = setTimeout(() => {
    document.querySelectorAll('.topbar-dropdown.dd-open').forEach(d => d.classList.remove('dd-open'));
  }, 120);
}

function ddCancel() { clearTimeout(_ddT); }

function closeTopbarDropdowns() {
  clearTimeout(_ddT);
  document.querySelectorAll('.topbar-dropdown.dd-open').forEach(d => d.classList.remove('dd-open'));
}

function navigateTopbarItem(event, pageId) {
  navigate(event, pageId);
  closeTopbarDropdowns();
}

function onTabClick(event, groupIdx) {
  if (window.matchMedia('(hover: hover)').matches) {
    navigateGroup(groupIdx);
  } else {
    event.stopPropagation();
    const wrapEl = event.currentTarget.closest('.topbar-tab-wrap');
    const dd = wrapEl?.querySelector('.topbar-dropdown');
    if (!dd) return;
    if (dd.classList.contains('dd-open')) { dd.classList.remove('dd-open'); }
    else { ddShow(wrapEl); }
  }
}

// ── Topbar：初始化使用者資訊 ──
function initTopbar() {
  const userId = sessionStorage.getItem('kpl_user') || '使用者';
  const nameEl   = document.getElementById('topbar-user-name');
  const avatarEl = document.getElementById('topbar-user-avatar');
  const umName   = document.getElementById('um-name');
  const umRole   = document.getElementById('um-role');
  if (nameEl)   nameEl.textContent   = userId;
  if (avatarEl) avatarEl.textContent = isAdmin() ? '👑' : '👤';
  if (umName)   umName.textContent   = userId;
  if (umRole)   umRole.textContent   = isAdmin() ? '管理員 · 日翊文化行銷' : '日翊文化行銷';
}

// ── Topbar：使用者選單 ──
function toggleUserMenu() {
  const menu = document.getElementById('user-menu');
  if (!menu) return;
  menu.hidden = !menu.hidden;
}

// ── 點擊外部自動關閉 User Menu 與 Topbar Dropdown ──
document.addEventListener('click', function handleClickOutside(e) {
  const menu    = document.getElementById('user-menu');
  const userBtn = document.getElementById('user-btn');
  const releaseNotice = document.getElementById('release-notice');
  if (menu && !menu.hidden && userBtn &&
      !menu.contains(e.target) && !userBtn.contains(e.target)) {
    menu.hidden = true;
  }
  if (releaseNotice && !releaseNotice.hidden && !e.target.closest('.topbar-notif-wrap')) {
    releaseNotice.hidden = true;
  }
  if (!e.target.closest('.topbar-tab-wrap')) {
    document.querySelectorAll('.topbar-dropdown.dd-open').forEach(d => d.classList.remove('dd-open'));
  }
});

// ── 頂部時間（時鐘，保留向下相容） ──
function updateTime() {
  const el = document.getElementById('nav-time');
  if (el) el.textContent = new Date().toLocaleString('zh-TW', { hour12:false });
}

function getBudgetYear() {
  const year = Number(String(DATA.dateFrom || '').slice(0, 4));
  return Number.isInteger(year) && year >= 2000 && year <= 2100 ? year : new Date().getFullYear();
}

function checkMobile() {
  if (isMobileLayout()) {
    document.body.classList.remove('sidebar-peek');
  } else {
    document.getElementById('sidebar')?.classList.remove('drawer-open');
    document.body.classList.remove('sidebar-mobile-open');
  }
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
  DATA.budgetSource = 'cloud';
  DATA.dispatch.budget = buildDispatchBudget(labor, freight, getCurrentMonthIndex());
  DATA.dataLatest.budget = latest || `${getCurrentMonthIndex() + 1}月`;
  DATA.dataOldest = DATA.dataOldest || {};
  DATA.dataOldest.budget = oldest;
  return true;
}

async function loadCloudBudgetData() {
  try {
    const res = await fetch(`/api/data/budget?year=${getBudgetYear()}`);
    if (handleAuthExpired(res)) return false;
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
  if (handleAuthExpired(res)) throw new Error('登入已逾時，請重新登入');
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

function applyFreightToDispatch(dailyRows) {
  const freightByDate = {};
  dailyRows.forEach(row => {
    const fullDate = row[4];
    if (fullDate) freightByDate[fullDate] = row;
  });

  const existing = {};
  DATA.dispatch.daily = DATA.dispatch.daily.map(row => {
    const fullDate = dispatchRowFullDate(row);
    existing[fullDate] = true;
    if (!fullDate || fullDate < DATA.dateFrom || fullDate > DATA.dateTo) return row;
    const freight = freightByDate[fullDate];
    return [
      row[0],
      row[1],
      freight?.[1] || 0,
      row[3],
      freight?.[2] || 0,
      row[5],
      freight?.[3] || 0,
      fullDate,
    ];
  });

  dailyRows.forEach(row => {
    const fullDate = row[4];
    if (!fullDate || existing[fullDate]) return;
    DATA.dispatch.daily.push([row[0], 0, row[1] || 0, 0, row[2] || 0, 0, row[3] || 0, fullDate]);
  });

  DATA.dispatch.daily.sort((a, b) => dispatchRowFullDate(a).localeCompare(dispatchRowFullDate(b)));
}

function applyLaborSummaryToDispatch(rows) {
  const dailyByDate = {};
  rows.forEach(row => {
    if (!dailyByDate[row.date]) {
      dailyByDate[row.date] = { date: row.date, '大溪倉': 0, '大肚倉': 0, '岡山倉': 0 };
    }
    if (dailyByDate[row.date][row.wh] !== undefined) {
      dailyByDate[row.date][row.wh] += Number(row.cost) || 0;
    }
  });
  resetDispatchLaborForRange();
  const dailyRows = Object.values(dailyByDate);
  applyLaborToDispatch(dailyRows);
  updateDispatchLatestUploadDate(dailyRows.map(row => row.date));
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

function applyCloudPicksRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    PICKS_RAW = [];
    DATA.dataLatest.picks = '';
    DATA.dataOldest = DATA.dataOldest || {};
    DATA.dataOldest.picks = '';
    return true;
  }

  PICKS_RAW = rows.map(row => ({
    wh: row.wh || '',
    date: normalizeCloudDate(row.date),
    biz: row.biz || '',
    area: row.area || '',
    op: row.op || '',
    picks: Number(row.picks) || 0,
    sourceSheet: 'Cloud SQL',
  })).filter(row => row.date && row.wh && row.picks > 0);

  const dates = PICKS_RAW.map(row => row.date).filter(Boolean).sort();
  DATA.dataLatest.picks = dates[dates.length - 1] || '';
  DATA.dataOldest = DATA.dataOldest || {};
  DATA.dataOldest.picks = dates[0] || '';
  return true;
}

async function loadCloudDataRange() {
  try {
    const res = await fetch('/api/data/range');
    if (handleAuthExpired(res)) return;
    if (!res.ok) return;
    const data = await res.json();
    if (data.ok) {
      DATA.cloudRange = DATA.cloudRange || {};
      DATA.cloudRange.labor              = data.labor              || { min: '', max: '' };
      DATA.cloudRange.freight            = data.freight            || { min: '', max: '' };
      DATA.cloudRange.freightMainline    = data.freightMainline    || { min: '', max: '' };
      DATA.cloudRange.freightNonMainline = data.freightNonMainline || { min: '', max: '' };
      DATA.cloudRange.picks              = data.picks              || { min: '', max: '' };
    }
  } catch {}
}

async function loadCloudLaborData({ summary = false } = {}) {
  try {
    const params = new URLSearchParams({
      date_from: DATA.dateFrom,
      date_to: DATA.dateTo,
    });
    if (summary) params.set('summary', '1');
    const res = await fetch(`/api/data/labor?${params.toString()}`);
    if (handleAuthExpired(res)) return false;
    if (!res.ok) return false;
    const data = await res.json();
    if (!data.ok) return false;
    if (summary) {
      DATA.dailySummary.laborRows = (data.rows || []).map(row => ({
        wh: row.wh || '',
        date: normalizeCloudDate(row.date),
        opArea: 'daily-summary',
        hours: Number(row.hours) || 0,
        cost: Number(row.cost) || 0,
      })).filter(row => row.date && row.wh);
      applyLaborSummaryToDispatch(DATA.dailySummary.laborRows);
      return true;
    }
    return applyCloudLaborRows(data.rows);
  } catch {
    return false;
  }
}

async function loadCloudPicksData() {
  try {
    const params = new URLSearchParams({
      date_from: DATA.dateFrom,
      date_to: DATA.dateTo,
    });
    const res = await fetch(`/api/data/picks?${params.toString()}`);
    if (handleAuthExpired(res)) return false;
    if (!res.ok) return false;
    const data = await res.json();
    return data.ok ? applyCloudPicksRows(data.rows) : false;
  } catch {
    return false;
  }
}

async function loadCloudFreightData({ summary = false } = {}) {
  try {
    const params = new URLSearchParams({ date_from: DATA.dateFrom, date_to: DATA.dateTo });
    if (summary) params.set('summary', '1');
    const res = await fetch(`/api/data/freight?${params.toString()}`);
    if (handleAuthExpired(res)) return false;
    if (!res.ok) return false;
    const data = await res.json();
    if (!data.ok) return false;
    if (summary) {
      DATA.dailySummary.freightRows = (data.dailyCosts || []).map(r => {
        const mmdd = `${r.date.slice(5, 7)}/${r.date.slice(8, 10)}`;
        return [mmdd, r.daxi || 0, r.dadu || 0, r.gangshan || 0, r.date];
      });
      applyFreightToDispatch(DATA.dailySummary.freightRows);
    } else {
      applyCloudFreightData(data);
    }
    return true;
  } catch {
    return false;
  }
}

function applyCloudFreightData(data) {
  const dailyCosts = data.dailyCosts || [];
  const details    = data.details    || [];
  const threshold  = DATA.freight.diffThreshold || 90;

  // dailyByWarehouse: [mm/dd, 大溪, 大肚, 岡山, fullDate]
  DATA.freight.dailyByWarehouse = dailyCosts.map(r => {
    const mmdd = `${r.date.slice(5, 7)}/${r.date.slice(8, 10)}`;
    return [mmdd, r.daxi || 0, r.dadu || 0, r.gangshan || 0, r.date];
  });
  applyFreightToDispatch(DATA.freight.dailyByWarehouse);

  // dailyTrend: [mm/dd, 合計, fullDate]
  DATA.freight.dailyTrend = dailyCosts.map(r => {
    const mmdd = `${r.date.slice(5, 7)}/${r.date.slice(8, 10)}`;
    return [mmdd, (r.daxi || 0) + (r.dadu || 0) + (r.gangshan || 0), r.date];
  });

  // 主線明細（用於 F002/F003）
  DATA.freight.details = details.map(r => ({
    fullDate:  r.date,
    vendor:    r.vendor,
    estimated: r.estimated || 0,
    actual:    r.actual    || 0,
    sourceType: r.source_type || '',
    reason:    r.reason || '',
    route:     r.route || '',
    categoryL1: r.category_l1 || '',
    categoryL2: r.category_l2 || '',
    budgetWarehouse: r.budget_warehouse || '',
    point:     0,
    rate:      r.source_type === 'nonmainline' ? 0 : r.estimated > 0 ? (r.actual / r.estimated) * 100 : 100,
  }));

  const allActual = dailyCosts.reduce((s, r) => s + (r.daxi || 0) + (r.dadu || 0) + (r.gangshan || 0), 0);
  DATA.freight.actualCost    = allActual;
  DATA.freight.totalCost     = allActual;
  DATA.freight.estimatedCost = details.reduce((s, r) => s + (r.estimated || 0), 0);
  DATA.freight.totalOrders   = details.length;
  DATA.freight.overCount     = DATA.freight.details.filter(r => r.rate > threshold).length;
  DATA.freight.saveCount     = DATA.freight.details.filter(r => r.rate <= threshold).length;
  DATA.freight.lastMonthCost = data.lastMonthTotal || 0;
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
  if (handleAuthExpired(res)) throw new Error('登入已逾時，請重新登入');
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.ok) throw new Error(data.MSG || `HTTP ${res.status}`);
  return data;
}

async function syncCloudPicks(parsed) {
  if (!parsed?.records?.length) return false;
  const payload = {
    records: parsed.records,
    fileName: parsed.fileName,
    importedBy: sessionStorage.getItem('kpl_user') || '',
  };
  const res = await fetch('/api/import/picks', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (handleAuthExpired(res)) throw new Error('登入已逾時，請重新登入');
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.ok) throw new Error(data.MSG || `HTTP ${res.status}`);
  return data;
}

async function uploadFreightWithConfirm(endpoint, records, fileName, label) {
  const importedBy = sessionStorage.getItem('kpl_user') || '';
  const dryRes = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ records, fileName, importedBy, dryRun: true }),
  });
  if (handleAuthExpired(dryRes)) throw new Error('登入已逾時，請重新登入');
  const dryData = await dryRes.json().catch(() => ({}));
  if (!dryRes.ok || !dryData.ok) throw new Error(dryData.MSG || `HTTP ${dryRes.status}`);

  const { rowsToImport, existingInRange, dateFrom, dateTo } = dryData;
  let msg = `${label}：${rowsToImport} 筆（${dateFrom} ～ ${dateTo}）`;
  if (existingInRange > 0) msg += `\n\n⚠️ 雲端已有此區間 ${existingInRange} 筆，上傳後將被覆蓋`;
  msg += '\n\n確認上傳？';
  if (!confirm(msg)) return null;

  const res = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ records, fileName, importedBy, dryRun: false }),
  });
  if (handleAuthExpired(res)) throw new Error('登入已逾時，請重新登入');
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.ok) throw new Error(data.MSG || `HTTP ${res.status}`);
  return data;
}

function parseFreightMainline(wb, fileName) {
  const sheetName = wb.SheetNames.find(n => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '', header: 1 });
    return rows.length && rows[0].some(c => String(c).replace(/\s+/g,'').includes('進貨日'));
  });
  if (!sheetName) {
    toast('❌ 找不到含「進貨日」欄位的工作表');
    document.getElementById('freight-mainline-status').textContent = '❌ 找不到工作表';
    return;
  }
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
  if (!raw.length) { toast('❌ 主線運費工作表沒有資料'); return; }
  const first = raw[0];
  const missing = ['進貨日', '倉別'].filter(col => valueByHeader(first, col) === undefined);
  if (missing.length) {
    toast('❌ 主線運費缺少欄位：' + missing.join('、'));
    document.getElementById('freight-mainline-status').textContent = '❌ 欄位不足';
    return;
  }
  const validRows = raw.filter(r => valueByHeader(r, '進貨日') !== '');
  const validCount = validRows.length;
  if (!validCount) { toast('❌ 主線運費沒有有效資料列'); return; }

  parsedFreightMainline = { records: validRows, fileName, rowCount: validCount, sheetName };
  document.getElementById('freight-mainline-status').textContent = `✅ ${validCount} 筆 · ${sheetName}`;
  const prev = document.getElementById('freight-mainline-preview');
  prev.innerHTML = `<div class="import-alert import-alert-info"><b>工作表：</b>${escapeReleaseText(sheetName)}　<b>有效列：</b>${validCount} 筆<br><small>確認後直接上傳至雲端資料庫</small></div>`;
  prev.style.display = 'block';
  document.getElementById('freight-mainline-btns').style.display = 'flex';
  toast(`✅ 主線運費解析完成：${validCount} 筆`);
}

function parseFreightNonMainline(wb, fileName) {
  const sheetName = wb.SheetNames.find(n => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '', header: 1 });
    return rows.length && rows[0].some(c => String(c).replace(/\s+/g,'').includes('派車原因'));
  });
  if (!sheetName) {
    toast('❌ 找不到含「派車原因」欄位的工作表');
    document.getElementById('freight-non-mainline-status').textContent = '❌ 找不到工作表';
    return;
  }
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
  if (!raw.length) { toast('❌ 非主線運費工作表沒有資料'); return; }
  const first = raw[0];
  const missing = ['進貨日', '廠商'].filter(col => valueByHeader(first, col) === undefined);
  if (missing.length) {
    toast('❌ 非主線運費缺少欄位：' + missing.join('、'));
    document.getElementById('freight-non-mainline-status').textContent = '❌ 欄位不足';
    return;
  }
  const validRows = raw.filter(r => valueByHeader(r, '進貨日') !== '');
  const validCount = validRows.length;
  if (!validCount) { toast('❌ 非主線運費沒有有效資料列'); return; }

  parsedFreightNonMainline = { records: validRows, fileName, rowCount: validCount, sheetName };
  document.getElementById('freight-non-mainline-status').textContent = `✅ ${validCount} 筆 · ${sheetName}`;
  const prev = document.getElementById('freight-non-mainline-preview');
  prev.innerHTML = `<div class="import-alert import-alert-info"><b>工作表：</b>${escapeReleaseText(sheetName)}　<b>有效列：</b>${validCount} 筆<br><small>伺服器將自動分類後寫入雲端</small></div>`;
  prev.style.display = 'block';
  document.getElementById('freight-non-mainline-btns').style.display = 'flex';
  toast(`✅ 非主線運費解析完成：${validCount} 筆`);
}

function resetFreightMainline() {
  parsedFreightMainline = null;
  setImportResultVisible('freight-mainline', false);
  document.getElementById('freight-mainline-status').textContent = '尚未上傳';
  document.getElementById('freight-mainline-preview').style.display = 'none';
  document.getElementById('freight-mainline-btns').style.display = 'none';
  document.getElementById('freight-mainline-file').value = '';
}

function resetFreightNonMainline() {
  parsedFreightNonMainline = null;
  setImportResultVisible('freight-non-mainline', false);
  document.getElementById('freight-non-mainline-status').textContent = '尚未上傳';
  document.getElementById('freight-non-mainline-preview').style.display = 'none';
  document.getElementById('freight-non-mainline-btns').style.display = 'none';
  document.getElementById('freight-non-mainline-file').value = '';
}

async function applyFreightMainline() {
  if (!parsedFreightMainline) return;
  const btns = document.getElementById('freight-mainline-btns').querySelectorAll('button');
  btns.forEach(b => b.disabled = true);
  try {
    const result = await uploadFreightWithConfirm(
      '/api/import/freight-mainline',
      parsedFreightMainline.records,
      parsedFreightMainline.fileName,
      '主線運費'
    );
    if (!result) return;
    document.getElementById('freight-mainline-status').textContent = `✅ 已上傳 ${result.rows} 筆`;
    toast(`✅ 主線運費已上傳雲端（${result.rows} 筆）`);
    await loadCloudDataRange();
    updateStatus();
  } catch (err) {
    toast('❌ ' + err.message);
    document.getElementById('freight-mainline-status').textContent = '❌ 上傳失敗';
  } finally {
    btns.forEach(b => b.disabled = false);
  }
}

async function applyFreightNonMainline() {
  if (!parsedFreightNonMainline) return;
  const btnsDiv = document.getElementById('freight-non-mainline-btns');
  const btns = btnsDiv.querySelectorAll('button');
  btns.forEach(b => b.disabled = true);
  const importedBy = sessionStorage.getItem('kpl_user') || '';
  try {
    // 步驟 1：Dry-run → 取得分類摘要
    const dryRes = await fetch('/api/import/freight-non-mainline', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        records: parsedFreightNonMainline.records,
        fileName: parsedFreightNonMainline.fileName,
        importedBy,
        dryRun: true,
      }),
    });
    if (handleAuthExpired(dryRes)) throw new Error('登入已逾時，請重新登入');
    const dryData = await dryRes.json().catch(() => ({}));
    if (!dryRes.ok || !dryData.ok) throw new Error(dryData.MSG || `HTTP ${dryRes.status}`);

    const { rowsToImport, existingInRange, dateFrom, dateTo, classification = [] } = dryData;
    const unclassified = classification.filter(c => c.l1 === '無法判斷').reduce((s, c) => s + c.count, 0);

    // 顯示分類明細表
    const prev = document.getElementById('freight-non-mainline-preview');
    prev.innerHTML = `
      <div class="import-alert import-alert-info">
        <b>${dateFrom} ～ ${dateTo}　共 ${rowsToImport} 筆</b>
        ${existingInRange ? `<br><span style="color:var(--app-warning)">⚠ 雲端已有此區間 ${existingInRange} 筆（將被覆蓋）</span>` : ''}
        ${unclassified ? `<br><span style="color:var(--app-danger)">⚠ ${unclassified} 筆無法判斷，上傳後需人工確認</span>` : ''}
      </div>
      <table class="tbl ops-compact-table" style="margin-top:8px;width:100%">
        <thead><tr><th>大分類</th><th>細分類</th><th style="text-align:right">筆數</th></tr></thead>
        <tbody>${classification.map(c => `
          <tr>
            <td>${escapeReleaseText(c.l1)}</td>
            <td>${c.l2 ? escapeReleaseText(c.l2) : '—'}</td>
            <td class="mono" style="text-align:right">${Number(c.count)}</td>
          </tr>`).join('')}
        </tbody>
      </table>`;
    prev.style.display = 'block';
    btns.forEach(b => b.disabled = false);

    // 步驟 2：使用者確認
    let confirmMsg = `非主線運費：${rowsToImport} 筆（${dateFrom} ～ ${dateTo}）`;
    if (unclassified) confirmMsg += `\n⚠️ ${unclassified} 筆無法判斷，上傳後須人工核對`;
    if (existingInRange) confirmMsg += `\n⚠️ 雲端已有此區間 ${existingInRange} 筆，將被覆蓋`;
    confirmMsg += '\n\n確認上傳至雲端？';
    if (!confirm(confirmMsg)) return;

    // 步驟 3：正式上傳
    btns.forEach(b => b.disabled = true);
    const res = await fetch('/api/import/freight-non-mainline', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        records: parsedFreightNonMainline.records,
        fileName: parsedFreightNonMainline.fileName,
        importedBy,
        dryRun: false,
      }),
    });
    if (handleAuthExpired(res)) throw new Error('登入已逾時，請重新登入');
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) throw new Error(data.MSG || `HTTP ${res.status}`);

    document.getElementById('freight-non-mainline-status').textContent = `✅ 已上傳 ${data.rows} 筆`;
    toast(`✅ 非主線運費已上傳雲端（${data.rows} 筆）`);
    await loadCloudDataRange();
    updateStatus();
  } catch (err) {
    toast('❌ ' + err.message);
    document.getElementById('freight-non-mainline-status').textContent = '❌ 上傳失敗';
  } finally {
    btns.forEach(b => b.disabled = false);
  }
}

// ── 初始化 ──
document.addEventListener('DOMContentLoaded', async () => {
  // 驗證登入
  if (!await loadSession()) return;

  // 載入頁面權限（管理員也載入，用於顯示哪些頁面被隱藏）
  await loadPagePermissions();

  // 預設日期必須在 cloud data load 之前設定，否則 API 收到空日期參數
  const _now = new Date();
  const _y = _now.getFullYear();
  const _m = String(_now.getMonth() + 1).padStart(2, '0');
  const _d = String(_now.getDate()).padStart(2, '0');
  DATA.dateFrom = `${_y}-${_m}-01`;
  DATA.dateTo   = `${_y}-${_m}-${_d}`;

  await Promise.all([
    loadCloudBudgetData(),
    loadCloudLaborData(),
    loadCloudPicksData(),
    loadCloudFreightData(),
  ]);

  initThemeMode();
  initSidebarState();
  initSidebarResize();
  applyTypographySettings();

  renderSidebar();
  initTopbar();
  renderReleaseNotice();
  void loadGithubReleases().then(() => {
    renderReleaseNotice();
    if (currentPageId === 'versions') renderVersionsPage();
  });

  const hash = location.hash.slice(1);
  if (hash) {
    const allItems = PAGES.flatMap(g => g.items);
    if (allItems.find(i => i.id === hash)) {
      currentPageId = hash;
    }
  }
  renderTopbarTabs();
  loadPage(currentPageId);

  updateTime();
  setInterval(updateTime, 60000);
  checkMobile();
  window.addEventListener('resize', checkMobile);

  // 初始化完成，淡出開機載入畫面
  if (typeof window.hideBootLoader === 'function') {
    window.hideBootLoader();
  }
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
function getFreightMonthValue() {
  const monthInput = document.getElementById('freight-month')?.value;
  if (/^\d{4}-\d{2}$/.test(monthInput || '')) return monthInput;
  const base = DATA.dateFrom || new Date().toISOString().slice(0, 10);
  return String(base).slice(0, 7);
}

function lockFreightToMonth(monthValue) {
  const safeValue = /^\d{4}-\d{2}$/.test(monthValue || '')
    ? monthValue
    : new Date().toISOString().slice(0, 7);
  const [year, month] = safeValue.split('-').map(Number);
  const lastDate = new Date(year, month, 0).getDate();
  DATA.dateFrom = `${safeValue}-01`;
  DATA.dateTo = `${safeValue}-${String(lastDate).padStart(2, '0')}`;
  const monthEl = document.getElementById('freight-month');
  const fromEl = document.getElementById('freight-from');
  const toEl = document.getElementById('freight-to');
  if (monthEl) monthEl.value = safeValue;
  if (fromEl) fromEl.value = DATA.dateFrom;
  if (toEl) toEl.value = DATA.dateTo;
}

function setupFreightMonthShell() {
  const pageHead = document.querySelector('#main .page-head');
  const filterBar = document.querySelector('#main .filter-bar');
  document.getElementById('freight-options-filter')?.remove();
  if (pageHead) {
    pageHead.style.display = '';
    const eyebrow = pageHead.querySelector('.page-eyebrow');
    const title = pageHead.querySelector('.page-h');
    if (eyebrow) eyebrow.textContent = '💰 成本分析 > 運費損益分析';
    if (title) title.textContent = '運費損益分析';
  }
  if (filterBar) {
    filterBar.style.display = '';
    filterBar.className = 'filter-bar freight-date-filter';
    filterBar.innerHTML = `
      <span class="filter-label">月份</span>
      <input type="month" class="filter-input" id="freight-month" onchange="applyFreightFilter()">
      <div class="filter-divider"></div>
      <span class="filter-label">倉別</span>
      <select class="filter-input" id="freight-warehouse" onchange="applyFreightFilter()">
        <option value="all">三倉總覽</option>
        <option value="大溪倉">大溪倉</option>
        <option value="大肚倉">大肚倉</option>
        <option value="岡山倉">岡山倉</option>
      </select>
      <div class="filter-divider"></div>
      <button class="btn btn-primary" onclick="applyFreightFilter()">套用</button>
      <span class="filter-meta">以月度檢視運費損益</span>
      <button class="btn btn-ghost export-btn" onclick="downloadPageReport('freight')" title="下載 Excel">↓ 匯出</button>
    `;
  }
  lockFreightToMonth(getFreightMonthValue());
  const warehouseEl = document.getElementById('freight-warehouse');
  if (warehouseEl) warehouseEl.value = DATA.freightSelectedWarehouse || 'all';
}

function renderFreightPage() {
  const grid = document.getElementById('freight-grid');
  grid.innerHTML = renderFreightReferenceDashboard();
  const pageHead = document.querySelector('#main .page-head');
  const filterBar = document.querySelector('#main .filter-bar');
  if (pageHead) pageHead.style.display = 'none';
  if (filterBar) filterBar.style.display = 'none';
  setupFreightMonthShell();

  const ffrom = document.getElementById('freight-from'); if (ffrom) ffrom.value = DATA.dateFrom;
  const fto   = document.getElementById('freight-to');   if (fto)   fto.value   = DATA.dateTo;
  const summary = typeof getFreightSummaryForPage === 'function'
    ? getFreightSummaryForPage()
    : { totalOrders: DATA.freight.totalOrders };
  const days = typeof getFreightTrendRowsForPage === 'function'
    ? getFreightTrendRowsForPage().length
    : DATA.freight.dailyTrend.length;
  document.getElementById('freight-meta').textContent =
    `資料區間：${DATA.dateFrom} ~ ${DATA.dateTo} · ${days} 天 · 共 ${summary.totalOrders.toLocaleString()} 筆配送`;
  normalizeDateFilterBars();
}

async function applyFreightFilter() {
  DATA.freightSelectedWarehouse = document.getElementById('freight-warehouse')?.value || 'all';
  lockFreightToMonth(getFreightMonthValue());
  // 月份改變後必須向後端重新抓該月資料，否則只會在記憶體舊資料上過濾，
  // 找不到該月就掉回展示資料（費用總計、動支率、配送筆數變成 demo 值）。
  await Promise.all([loadCloudBudgetData(), loadCloudFreightData()]);
  if (currentPageId === 'freight') renderFreightPage();
}

// ════════════════════════════════════════════
// Import Page 邏輯
// ════════════════════════════════════════════
let parsedFreightMainline    = null;
let parsedFreightNonMainline = null;
let parsedLabor   = null;
let parsedPicks   = null;
let parsedBudget  = null;
let activeImportRequest = 0;
const IMPORT_TYPES = ['budget', 'freight-mainline', 'freight-non-mainline', 'labor', 'picks'];

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
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = '';
  parseExcel(file, type);
}

function setImportResultVisible(type, visible = true) {
  const row = document.querySelector(`.import-row[data-import-result="${type}"]`);
  if (row) row.classList.toggle('is-active', visible);

  const title = document.getElementById('import-results-title');
  if (title) {
    title.hidden = !document.querySelector('.import-row[data-import-result].is-active');
  }
}

function clearImportPreview(type) {
  if (type === 'budget') parsedBudget = null;
  else if (type === 'freight-mainline') parsedFreightMainline = null;
  else if (type === 'freight-non-mainline') parsedFreightNonMainline = null;
  else if (type === 'labor') parsedLabor = null;
  else if (type === 'picks') parsedPicks = null;

  const preview = document.getElementById(type + '-preview');
  if (preview) {
    preview.innerHTML = '';
    preview.style.display = 'none';
  }

  const buttons = document.getElementById(type + '-btns');
  if (buttons) buttons.style.display = 'none';

  const fileInput = document.getElementById(type + '-file');
  if (fileInput) fileInput.value = '';

  setImportResultVisible(type, false);
}

function beginImportPreview(file, type) {
  IMPORT_TYPES.forEach(clearImportPreview);
  setImportResultVisible(type, true);

  const status = document.getElementById(type + '-status');
  if (status) status.textContent = `讀取中：${file.name}`;
}

function parseExcel(file, type) {
  const requestId = ++activeImportRequest;
  beginImportPreview(file, type);
  setImportResultVisible(type, true);
  document.getElementById(type + '-status').textContent = '解析中…';
  const reader = new FileReader();
  reader.onload = e => {
    if (requestId !== activeImportRequest) return;
    try {
      const wb = XLSX.read(e.target.result, { type:'array' });
      const detectedType = detectWorkbookType(wb);
      if (detectedType && detectedType !== type) {
        document.getElementById(type + '-status').textContent = `↪ 已改由${importTypeLabel(detectedType)}解析`;
        setImportResultVisible(type, false);
        type = detectedType;
        setImportResultVisible(type, true);
        document.getElementById(type + '-status').textContent = '解析中…';
      }
      if (type === 'freight-mainline') parseFreightMainline(wb, file.name);
      else if (type === 'freight-non-mainline') parseFreightNonMainline(wb, file.name);
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
    'freight-mainline': '主線運費',
    'freight-non-mainline': '非主線運費',
    labor: '工時資料',
    picks: '揀次資料',
  }[type] || '正確類型';
}

function downloadImportTemplate(type) {
  const wb = XLSX.utils.book_new();
  const addSheet = (name, headers, example = []) => {
    const ws = XLSX.utils.aoa_to_sheet([headers, example]);
    ws['!cols'] = headers.map(header => ({ wch: Math.max(12, String(header).length * 2 + 2) }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  };

  if (type === 'budget') {
    const months = Array.from({ length:12 }, (_, index) => `${index + 1}月`);
    addSheet('人力', ['區別', '單位', '類別', '作業項目', ...months, '合計'], ['北區', '理貨一課', '人力', '夜配', ...Array(12).fill(0), 0]);
    addSheet('運務', ['區別', '倉別', '類別', '作業項目', ...months, '合計'], ['北區', '大溪倉', '主線', '夜配', ...Array(12).fill(0), 0]);
  } else if (type === 'freight-mainline') {
    addSheet('主線運費', ['倉別', '進貨日', '路線來源', '路線', '配別', '配送商', '預計計價結果', '到點計價結果', '計價結果'], ['大溪倉', '2026/03/01', '主線', '大溪－北區', '夜配', '台灣宅配通', 0, 0, 0]);
  } else if (type === 'freight-non-mainline') {
    addSheet('非主線加派費用', ['進貨日', '廠商', '配別', '配送商', '派車原因', '計價方式', '單價', '趟數', '費用小計', '備註'], ['2026/03/01', '大溪倉', '夜配', '台灣宅配通', '爆量', '趟', 0, 1, 0, '']);
  } else if (type === 'labor') {
    addSheet('工時資料', ['倉別', '日期', '廠商', '班別', '員編', '作業課別', '姓名', '作業區域', '作業時數', '實際費用', '裝箱時數', '夜間時數', '正常時數'], ['大溪倉', 46082, '日翊', '日', 'A0001', '理貨一課', '範例人員', '夜配理貨', 8, 0, 0, 0, 8]);
  } else if (type === 'picks') {
    addSheet('揀次資料', ['倉別', '日期', '業務類別', '作業區', '工時區域', '揀次'], ['大溪倉', 46082, '夜配', '理貨區', '夜配理貨', 0]);
  } else {
    return;
  }

  XLSX.writeFile(wb, `${importTypeLabel(type)}_匯入範本.xlsx`);
}

function downloadPageReport(page) {
  const wb = XLSX.utils.book_new();
  const addSheet = (name, rows) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = rows[0].map(h => ({ wch: Math.max(12, String(h).length * 2 + 2) }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  };

  const from = DATA.dateFrom || '';
  const to   = DATA.dateTo   || '';
  const labels = { dispatch: '總費用動支率', freight: '運費損益', labor: '人力工時', picks: '揀次分析' };
  const filename = `KPL_${labels[page]}_${from}~${to}.xlsx`;

  if (page === 'freight') {
    if (!DATA.freight?.details?.length) {
      toast('⚠️ 運費資料尚未載入，請稍候再匯出');
      return;
    }
    const wh = DATA.freightSelectedWarehouse || document.getElementById('freight-warehouse')?.value || 'all';
    let data = DATA.freight.details;
    if (wh !== 'all') data = data.filter(r => r.budgetWarehouse === wh);
    addSheet('明細', [
      ['日期', '倉別', '廠商', '路線', '派車原因', '預計費用', '計費結果', '到點費用', '大分類', '小分類', '來源類型'],
      ...data.map(r => [r.fullDate, r.budgetWarehouse, r.vendor, r.route, r.reason,
                        r.estimated, r.actual, r.point, r.categoryL1, r.categoryL2, r.sourceType])
    ]);

  } else if (page === 'labor') {
    const shiftFilter = document.getElementById('labor-shift')?.value || '';
    const deptFilter  = document.getElementById('labor-vendor')?.value || '';
    let data = (LABOR_RAW || []).filter(r =>
      dateInSelectedRange(r.date) && r.opArea !== '午休時間' && r.hours > 0
    );
    if (shiftFilter) data = data.filter(r => r.shift === shiftFilter);
    if (deptFilter)  data = data.filter(r => r.dept === deptFilter);

    addSheet('明細', [
      ['倉別', '日期', '廠商', '班別', '員編', '作業課別', '姓名', '作業區域', '作業時數', '裝箱時數', '夜間時數', '正常時數', '費用'],
      ...data.map(r => [r.wh, r.date, r.vendor, r.shift, r.empId, r.dept,
                        r.name, r.opArea, r.hours, r.boxHours, r.nightHrs, r.normHrs, r.cost])
    ]);

    const byDept = {};
    data.forEach(r => {
      const k = r.dept || '未分類';
      if (!byDept[k]) byDept[k] = { hrs: 0, cost: 0, emps: new Set() };
      byDept[k].hrs  += r.hours;
      byDept[k].cost += r.cost;
      byDept[k].emps.add(r.empId);
    });
    addSheet('彙總', [
      ['課別', '人次', '工時(H)', '費用', '時薪'],
      ...Object.entries(byDept)
        .sort((a, b) => b[1].hrs - a[1].hrs)
        .map(([dept, v]) => [dept, v.emps.size, v.hrs, v.cost,
                             v.hrs > 0 ? Math.round(v.cost / v.hrs) : 0])
    ]);

  } else if (page === 'dispatch') {
    const rows = (DATA.dispatch.daily || []).filter(r => dateInSelectedRange(r[7]));
    addSheet('每日明細', [
      ['日期', '大溪_人力', '大溪_運務', '大肚_人力', '大肚_運務', '岡山_人力', '岡山_運務', '日合計'],
      ...rows.map(r => {
        const total = r[1] + r[2] + r[3] + r[4] + r[5] + r[6];
        return [r[7], r[1], r[2], r[3], r[4], r[5], r[6], total];
      })
    ]);

  } else if (page === 'picks') {
    const wh = document.getElementById('picks-wh')?.value || '';
    const op = document.getElementById('picks-op')?.value || '';
    let data = (PICKS_RAW || []).filter(r => dateInSelectedRange(r.date));
    if (wh) data = data.filter(r => r.wh === wh);
    if (op) data = data.filter(r => r.op === op);

    addSheet('明細', [
      ['日期', '倉別', '業務類別', '作業區', '工時區域', '揀次'],
      ...data.map(r => [r.date, r.wh, r.biz, r.area, r.op, r.picks])
    ]);

    const allWhs = [...new Set(data.map(r => r.wh))].sort();
    const byOp = {};
    data.forEach(r => {
      if (!byOp[r.op]) byOp[r.op] = { total: 0 };
      byOp[r.op][r.wh] = (byOp[r.op][r.wh] || 0) + r.picks;
      byOp[r.op].total += r.picks;
    });
    addSheet('彙總', [
      ['作業區域', ...allWhs, '合計'],
      ...Object.entries(byOp)
        .sort((a, b) => b[1].total - a[1].total)
        .map(([op, v]) => [op, ...allWhs.map(w => v[w] || 0), v.total])
    ]);

  } else {
    return;
  }

  XLSX.writeFile(wb, filename);
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
      <table class="preview-table ops-compact-table">
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
  DATA.budgetSource = 'excel-pending';
  DATA.dispatch.budget = parsedBudget.dispatchBudget;
  DATA.dataLatest.budget = `${parsedBudget.monthIndex + 1}月`;
  if (currentPageId === 'dispatch') renderDispatchPage();
  if (currentPageId === 'daily') renderDailyPage();
  if (currentPageId === 'annual') renderAnnualPage();
  updateStatus();
  toast('✅ 年度預算已套用！總費用動支率預算已更新');
  try {
    const result = await syncCloudBudget(parsedBudget);
    await loadCloudBudgetData();
    rerenderDashboardPage(currentPageId);
    toast(`✅ 年度預算已同步雲端（${result.rows} 筆）`);
  } catch (err) {
    console.warn('Budget cloud sync failed:', err);
    toast(`⚠️ 年度預算已套用，但雲端同步失敗：${err.message}`);
  }
}

function resetBudget() {
  parsedBudget = null;
  setImportResultVisible('budget', false);
  document.getElementById('budget-status').textContent = '尚未上傳';
  document.getElementById('budget-preview').style.display = 'none';
  document.getElementById('budget-btns').style.display = 'none';
  document.getElementById('budget-file').value = '';
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
    const reason = String(
      valueByHeader(r, '派車原因')
      || valueByHeader(r, '非主線派車原因')
      || valueByHeader(r, '原因')
      || ''
    ).trim();
    const route = String(
      valueByHeader(r, '主線派車路線')
      || valueByHeader(r, '派車路線')
      || valueByHeader(r, '路線')
      || ''
    ).trim();
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
      reason,
      route,
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
      <table class="preview-table ops-compact-table">
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
      <table class="preview-table ops-compact-table">
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
    await loadCloudDataRange();
    updateStatus();
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
  setImportResultVisible('labor', false);
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
      <table class="preview-table ops-compact-table">
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

async function applyPicks() {
  if (!parsedPicks) return;
  PICKS_RAW = parsedPicks.records;
  if (currentPageId === 'picks') renderPicksPage();
  if (currentPageId === 'productivity') renderProductivityPage();
  updateStatus();
  toast('✅ 揀次資料已套用！切換至「揀次分析」頁查看');
  try {
    const result = await syncCloudPicks(parsedPicks);
    toast(`✅ 揀次資料已同步雲端（${result.rows} 筆）`);
    await loadCloudDataRange();
    updateStatus();
  } catch (err) {
    console.warn('Picks cloud sync failed:', err);
    toast(`⚠️ 揀次已套用，但雲端同步失敗：${err.message}`);
  }
}

function resetPicks() {
  parsedPicks = null;
  PICKS_RAW = [];
  setImportResultVisible('picks', false);
  document.getElementById('picks-status').textContent = '尚未上傳';
  document.getElementById('picks-preview').style.display = 'none';
  document.getElementById('picks-btns').style.display = 'none';
  document.getElementById('picks-file').value = '';
  if (currentPageId === 'picks') renderPicksPage();
  if (currentPageId === 'productivity') renderProductivityPage();
  updateStatus();
}

function updateStatus() {
  const latestLabor = DATA.cloudRange?.labor?.max || (() => {
    const raw = (typeof LABOR_RAW !== 'undefined') ? LABOR_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort().pop() || '';
  })();
  const oldestLabor = DATA.cloudRange?.labor?.min || (() => {
    const raw = (typeof LABOR_RAW !== 'undefined') ? LABOR_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort()[0] || '';
  })();
  const latestPicks = DATA.cloudRange?.picks?.max || (() => {
    const raw = (typeof PICKS_RAW !== 'undefined') ? PICKS_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort().pop() || '';
  })();
  const oldestPicks = DATA.cloudRange?.picks?.min || (() => {
    const raw = (typeof PICKS_RAW !== 'undefined') ? PICKS_RAW : [];
    return raw.map(r => r.date).filter(Boolean).sort()[0] || '';
  })();

  const rows = [
    { key:'budget', type:'💰 年度預算', real:hasDispatchBudget(), latest: DATA.dataLatest?.budget || '', oldest: DATA.dataOldest?.budget || '' },
    { key:'freight-mainline', type:'🚛 主線運費', real: !!(DATA.cloudRange?.freightMainline?.max), latest: DATA.cloudRange?.freightMainline?.max || '', oldest: DATA.cloudRange?.freightMainline?.min || '' },
    { key:'freight-non-mainline', type:'🚐 非主線運費', real: !!(DATA.cloudRange?.freightNonMainline?.max), latest: DATA.cloudRange?.freightNonMainline?.max || '', oldest: DATA.cloudRange?.freightNonMainline?.min || '' },
    { key:'labor', type:'💵 人力費用', real:!!parsedLabor || !!((typeof LABOR_RAW !== 'undefined') && LABOR_RAW.length), latest: latestLabor, oldest: oldestLabor },
    { key:'picks', type:'⚡ 揀次資料', real:!!parsedPicks || !!((typeof PICKS_RAW !== 'undefined') && PICKS_RAW.length) || !!latestPicks, latest: latestPicks, oldest: oldestPicks },
  ];

  const dateCell = v => v
    ? `<span style="font-family:var(--f-mono);font-size:13px">${v}</span>`
    : `<span style="color:#bbb;font-size:12px">—</span>`;

  document.getElementById('status-tbody').innerHTML = rows.map(r => {
    const c = r.real ? '#1b7c33' : '#e07855';
    const s = r.real ? '✅ 已套用' : '⚠️ 尚未上傳';
    return `<tr>
      <td style="font-weight:700">${r.type}</td>
      <td><span style="color:${c};font-weight:700">${s}</span></td>
      <td>${dateCell(r.latest)}</td>
      <td>${dateCell(r.oldest)}</td>
      <td><button class="btn btn-primary import-table-action" onclick="document.getElementById('${r.key}-file').click()">選擇檔案</button></td>
      <td><button class="btn btn-ghost import-table-action" onclick="downloadImportTemplate('${r.key}')">下載範本</button></td>
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
    <div class="ops-table-frame">
      <table class="tbl ops-compact-table">
        <thead><tr><th>作業區域</th><th style="text-align:right">大肚倉</th><th style="text-align:right">大溪倉</th><th style="text-align:right">岡山倉</th><th style="text-align:right">合計</th><th>佔比</th></tr></thead>
        <tbody>${opRows}</tbody>
      </table>
    </div>
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

  // 數字格式：≥ 萬以「萬」顯示（最多 2 位小數），未達萬正常顯示
  const wanNum = (v) => {
    v = Number(v) || 0;
    return v >= 10000
      ? `${(v / 10000).toLocaleString('zh-TW', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}萬`
      : v.toLocaleString('zh-TW', { maximumFractionDigits: 1 });
  };
  const wanMoney = (v) => `$${wanNum(v)}`;
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
            <span class="labor-top-value">${wanNum(s.hrs)}h · ${s.pct.toFixed(1)}%</span>
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
      <td style="text-align:right;font-family:var(--f-mono)">${wanNum(it.hrs)}</td>
      <td style="text-align:right;font-family:var(--f-mono)">${wanMoney(it.cost)}</td>
      <td style="text-align:right;font-family:var(--f-mono)">${wanMoney(rate)}</td>
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
        <td class="mono num-right">${wanNum(it.hrs)}</td>
        <td class="mono num-right">${wanMoney(it.cost)}</td>
        <td class="mono num-right">${wanMoney(rate)}</td>
      </tr>`;
    }).join('');

  const fmeta = document.getElementById('labor-filter-meta');
  if (fmeta) fmeta.textContent = `${data.length} 筆工時記錄`;

  document.getElementById('labor-grid').innerHTML = `
  <div class="w s3 metric-card">
    <div class="gold-band">L001 · HOURS</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>總工時</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-blue);line-height:1">${wanNum(totalHrs)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">小時</div>
    </div>
  </div>
  <div class="w s3 metric-card">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">L002 · COST</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>總費用</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">${wanMoney(totalCost)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">元</div>
    </div>
  </div>
  <div class="w s3 metric-card">
    <div class="gold-band" style="background:#2ea85a;color:white">L003 · RATE</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:#2ea85a"></div>平均時薪</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:#2ea85a;line-height:1">${wanMoney(avgRate)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">元/小時</div>
    </div>
  </div>
  <div class="w s3 metric-card">
    <div class="gold-band" style="background:var(--ry-muted);color:white">L004 · PEOPLE</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:var(--ry-muted)"></div>出勤人日</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">${wanNum(personDays)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">${empCount.toLocaleString()} 位員工</div>
    </div>
  </div>
  <div class="w s6 labor-struct-card">
    <div class="gold-band">L005 · ⚡ 工時結構 · 作業區域</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各作業區域工時佔比</div><span class="wmeta">總 ${wanNum(totalHrs)} h</span></div>
    <div class="labor-pie-wrap">${structHtml}</div>
  </div>
  <div class="w s6 table-card labor-shift-card">
    <div class="gold-band">L006 · 🌙 班別工時分析</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>班別成本</div></div>
    <table class="tbl">
      <thead><tr><th>班別</th><th style="text-align:right">工時(h)</th><th style="text-align:right">費用</th><th style="text-align:right">時薪</th></tr></thead>
      <tbody>${shiftRows || '<tr><td colspan="4" style="text-align:center;color:var(--ry-muted)">無資料</td></tr>'}</tbody>
    </table>
  </div>
  <div class="w s12 table-card labor-dept-card">
    <div class="gold-band">L007 · 🏢 課別工時彙總</div>
    <div class="wh"><div class="wl"><div class="wdot"></div>各作業課別工時與費用</div></div>
    <div class="ops-table-frame labor-dept-edge">
      <div class="scroll-x">
        <table class="tbl labor-dept-table ops-compact-table">
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
