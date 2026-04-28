// ═══════════════════════════════════════════════════════
// utils.js · 共用工具函式
// ═══════════════════════════════════════════════════════

// 三色門檻（< 75% 綠 · 75-90% 橘 · >= 90% 紅）
function colorFor(v) {
  if (v >= 90) return '#d9401b';
  if (v >= 75) return '#e07855';
  return '#1b7c33';
}

function labelFor(v) {
  if (v >= 90) return '🔴 危險';
  if (v >= 75) return '🟡 注意';
  return '🟢 安全';
}

function bgFor(v) {
  if (v >= 90) return '#fdf0ec';
  if (v >= 75) return '#fff4e8';
  return '#eef9f0';
}

// 金額格式化（含千分位、$符號）
function fmtMoney(n) {
  return '$' + Number(n).toLocaleString();
}

// 右下角 Toast 通知
function toast(msg) {
  let t = document.getElementById('toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'toast';
    t.className = 'toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.classList.add('on');
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => t.classList.remove('on'), 2200);
}

// 計算常用數值
function getActualPct() {
  if (!DATA.budget) return 0;
  return DATA.actual / DATA.budget * 100;
}

function getMonthProgress() {
  if (!DATA.totalDays) return 0;
  return DATA.dayOfMonth / DATA.totalDays;
}

function getProjectedPct() {
  const progress = getMonthProgress();
  if (!progress) return 0;
  return getActualPct() / getMonthProgress();
}
