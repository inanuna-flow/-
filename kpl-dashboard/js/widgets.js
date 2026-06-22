// ═══════════════════════════════════════════════════════
// widgets.js · 儀表板 Widget 渲染函式
// ═══════════════════════════════════════════════════════
// 元件樣式原則：固定外觀放在 widget.css 的 CSS class；
// JS 只保留資料結構、條件 class，以及必要的動態 style。

// M012 預算達成率（從已匯入的人力 + 運費資料自動計算）
function renderM012() {
  const summaryMatchesRange = DATA.dailySummary?.dateFrom === DATA.dateFrom &&
    DATA.dailySummary?.dateTo === DATA.dateTo;
  const laborSource = summaryMatchesRange
    ? DATA.dailySummary.laborRows
    : (typeof LABOR_RAW !== 'undefined' ? LABOR_RAW : []);
  const freightSource = summaryMatchesRange
    ? DATA.dailySummary.freightRows
    : getFreightDailyRowsFiltered();
  const laborRaw  = laborSource
    .filter(r => dateInSelectedRange(r.date) && r.hours > 0 && r.opArea !== '午休時間');
  const laborCost   = laborRaw.reduce((s, r) => s + r.cost, 0);
  const freightCost = freightSource
    .reduce((s, r) => s + (r[1]||0) + (r[2]||0) + (r[3]||0), 0);
  const actual = laborCost + freightCost;

  const idx    = monthIndexFromDate(DATA.dateFrom);
  const WHS    = ['大溪倉', '大肚倉', '岡山倉'];
  const budget = WHS.reduce((s, wh) =>
    s + (DATA.annualBudget?.labor?.[wh]?.[idx]   || 0)
      + (DATA.annualBudget?.freight?.[wh]?.[idx] || 0), 0);

  if (!budget && !actual) {
    return `
  <div class="w s12 table-card">
    <div class="wh"><div class="wl"><div class="wdot"></div>${DATA.widgetLabels?.m012 || 'M012 預算達成率'}</div></div>
    <div class="empty-state">請先匯入年度預算與費用資料</div>
  </div>`;
  }

  const dayOfMonth = Number((DATA.dateTo  || '').slice(8, 10)) || 0;
  const totalDays  = daysInMonthFor(DATA.dateFrom);
  const progress   = totalDays  ? dayOfMonth / totalDays : 0;
  const pct        = budget     ? actual / budget * 100  : 0;
  const projected  = progress   ? pct / progress         : 0;
  const curColor   = colorFor(pct);
  const projColor  = colorFor(projected);

  return `
  <div class="w s12">
    <div class="gold-band gold-band-dynamic" style="background:${projColor}">
      ${DATA.widgetLabels?.m012 || 'M012 預算達成率'} · 基準 = ${idx+1}月預算 ${fmtMoney(Math.round(budget))}（來源：年度預算）
    </div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${projColor}"></div>動支率監控</div>
      <div class="wmeta">${dayOfMonth}/${totalDays} 天 · 月進度 ${(progress*100).toFixed(1)}%</div>
    </div>
    <div class="m012-grid">
      <div class="m012-panel m012-panel-bordered">
        <div class="metric-label">累計動支率</div>
        <div class="metric-row">
          <div class="metric-value" style="color:${curColor}">${pct.toFixed(1)}</div>
          <div class="metric-unit" style="color:${curColor}">%</div>
          <div class="metric-status">${labelFor(pct)}</div>
        </div>
        <div class="metric-track metric-track-lg">
          <div class="metric-fill" style="width:${Math.min(pct,100)}%;background:${curColor}"></div>
          <div class="metric-limit metric-limit-90"></div>
        </div>
        <div class="metric-sub">${fmtMoney(Math.round(actual))} / ${fmtMoney(Math.round(budget))}</div>
        <div class="m012-cost-row">
          <div class="m012-cost-item">
            <div class="m012-cost-label">人力</div>
            <div class="m012-cost-val">${fmtWan(laborCost)}</div>
          </div>
          <div class="m012-cost-sep"></div>
          <div class="m012-cost-item">
            <div class="m012-cost-label">運費</div>
            <div class="m012-cost-val">${fmtWan(freightCost)}</div>
          </div>
        </div>
      </div>
      <div class="m012-panel">
        <div class="metric-label">預估月底動支率</div>
        <div class="metric-row">
          <div class="metric-value" style="color:${projColor}">${projected.toFixed(1)}</div>
          <div class="metric-unit" style="color:${projColor}">%</div>
          <div class="metric-status">${labelFor(projected)}</div>
        </div>
        <div class="metric-track metric-track-lg">
          <div class="metric-fill" style="width:${Math.min(projected,120)*0.8}%;background:${projColor}"></div>
          <div class="metric-limit metric-limit-72"></div>
        </div>
        <div class="metric-sub">依目前速度線性外推</div>
        <div class="m012-note" style="background:${bgFor(projected)};border-left-color:${projColor}">
          ${projected>=90 ? `<b class="text-danger">⚠️ 月底預估將超過 90% 目標上限</b>，建議立即盤點未執行項目、延後可遞延支出。`
            : projected>=75 ? `<b class="text-warning">🟡 月底預估接近警戒線</b>，請留意未來兩週的費用節奏。`
                            : `<b class="text-safe">🟢 月底預估控管良好</b>，可維持目前節奏。`}
        </div>
      </div>
    </div>
  </div>`;
}

// 每月動支率時序列
function renderM015(selectedYear) {
  // 用 DATA.m015（全年彙總，不受日期篩選器限制）
  const laborSource   = DATA.m015?.laborRows   || [];
  const freightSource = DATA.m015?.freightRows || [];
  const hasData = laborSource.length > 0 || freightSource.length > 0;

  if (!hasData) {
    return `
  <div class="w s12 table-card m015-card">
    <div class="wh"><div class="wl"><div class="wdot dot-freight"></div>${DATA.widgetLabels?.m015 || '每月動支率時序列'}</div></div>
    <div class="empty-state">請先匯入人力費用資料</div>
  </div>`;
  }

  // 依月份彙總人力費用
  const laborByMonth = {};
  laborSource.forEach(r => {
    if (!r.date) return;
    const ym = r.date.slice(0, 7);
    if (!laborByMonth[ym]) laborByMonth[ym] = 0;
    laborByMonth[ym] += r.cost;
  });

  // 依月份彙總運務費用
  const freightByMonth = {};
  freightSource.forEach(r => {
    const fullDate = r[4] || shortToFreightFullDate(r[0]);
    if (!fullDate) return;
    const ym = fullDate.slice(0, 7);
    if (!freightByMonth[ym]) freightByMonth[ym] = 0;
    freightByMonth[ym] += (r[1] || 0) + (r[2] || 0) + (r[3] || 0);
  });

  const allMonths = [...new Set([...Object.keys(laborByMonth), ...Object.keys(freightByMonth)])].sort();

  // 可用年份清單（降序）；預設取日期篩選器的年份，找不到則取最新
  const years = [...new Set(allMonths.map(ym => ym.slice(0, 4)))].sort().reverse();
  const filterYear = (DATA.dateFrom || '').slice(0, 4);
  const activeYear = selectedYear || DATA.m015?.year || (years.includes(filterYear) ? filterYear : years[0]) || '';

  // 篩選選定年份的月份（倒序：12月→1月）
  const months = allMonths.filter(ym => ym.startsWith(activeYear)).reverse();

  const yearOpts = years.map(y =>
    `<option value="${y}" ${y === activeYear ? 'selected' : ''}>${y}年</option>`
  ).join('');

  const rows = months.map(ym => {
    const labor   = laborByMonth[ym]   || 0;
    const freight = freightByMonth[ym] || 0;
    const total   = labor + freight;
    const mIdx    = Number(ym.slice(5, 7)) - 1;
    const WHS     = ['大溪倉', '大肚倉', '岡山倉'];
    const budget  = WHS.reduce((s, wh) =>
      s + (DATA.annualBudget?.labor?.[wh]?.[mIdx]   || 0)
        + (DATA.annualBudget?.freight?.[wh]?.[mIdx] || 0), 0);
    const pct     = budget ? total / budget * 100 : 0;
    const color   = budget ? colorFor(pct) : 'var(--ry-muted)';
    const [y, m]  = ym.split('-');
    return `<tr>
      <td style="font-weight:700">${y}年${Number(m)}月</td>
      <td class="mono num-right">${fmtMoney(Math.round(labor))}</td>
      <td class="mono num-right">${fmtMoney(Math.round(freight))}</td>
      <td class="mono num-right actual-strong">${fmtMoney(Math.round(total))}</td>
      <td class="mono num-right">${budget ? fmtMoney(Math.round(budget)) : '—'}</td>
      <td class="num-right">
        ${budget
          ? `<span class="solid-pct-badge" style="background:${color}">${pct.toFixed(1)}%</span>`
          : '—'}
      </td>
    </tr>`;
  }).join('');

  return `
  <div class="w s12 table-card m015-card">
    <div class="freight-ref-matrix-heading" style="padding:0 0 10px 0">
      <div class="wl"><div class="wdot dot-freight"></div>${DATA.widgetLabels?.m015 || '每月動支率時序列'}</div>
      <label class="freight-ref-matrix-view">
        <span>年份</span>
        <select class="filter-input" id="m015-year-select">${yearOpts}</select>
      </label>
    </div>
    <div class="ops-table-frame">
      <table class="tbl ops-compact-table">
        <thead><tr>
          <th>月份</th>
          <th class="num-right">人力費用</th>
          <th class="num-right">運務費用</th>
          <th class="num-right">合計實際</th>
          <th class="num-right">月預算</th>
          <th class="num-right">動支率</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="table-note">
        📌 動支率 = 合計實際 ÷ 月預算（需載入年度預算）· 三色門檻：&lt; 75% 🟢 安全 · 75–90% 🟡 注意 · &gt; 90% 🔴 危險
      </div>
    </div>
  </div>`;
}

function shortToFreightFullDate(mmdd) {
  const year = (DATA.dateFrom || '').slice(0, 4) || String(new Date().getFullYear());
  const parts = String(mmdd || '').split('/');
  if (parts.length !== 2) return '';
  return `${year}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`;
}

function freightDateInRange(fullDate) {
  if (!fullDate) return false;
  if (DATA.dateFrom && fullDate < DATA.dateFrom) return false;
  if (DATA.dateTo && fullDate > DATA.dateTo) return false;
  return true;
}

function getFreightDailyRowsFiltered() {
  return DATA.freight.dailyByWarehouse.filter(row => freightDateInRange(row[4] || shortToFreightFullDate(row[0])));
}

function getFreightTrendFiltered() {
  const selectedWarehouse = getSelectedFreightWarehouse();
  if (selectedWarehouse !== 'all') {
    const amountIndex = FREIGHT_WAREHOUSE_ORDER.indexOf(selectedWarehouse) + 1;
    return getFreightDailyRowsFiltered().map(row => [
      row[0],
      Number(row[amountIndex] || 0),
      row[4] || shortToFreightFullDate(row[0]),
    ]);
  }
  return DATA.freight.dailyTrend.filter(row => freightDateInRange(row[2] || shortToFreightFullDate(row[0])));
}

function getFreightDetailsFiltered() {
  if (!DATA.freight.details) return null;
  const selectedWarehouse = getSelectedFreightWarehouse();
  return DATA.freight.details.filter(r =>
    freightDateInRange(r.fullDate) &&
    (selectedWarehouse === 'all' || r.budgetWarehouse === selectedWarehouse)
  );
}

function summarizeFreightVendorsForRange(records) {
  const vendors = {};
  records.forEach(r => {
    if (!vendors[r.vendor]) vendors[r.vendor] = { name:r.vendor, contract:0, point:0, actual:0, count:0 };
    vendors[r.vendor].contract += r.estimated;
    vendors[r.vendor].point += r.point;
    vendors[r.vendor].actual += r.actual;
    vendors[r.vendor].count += 1;
  });
  return Object.values(vendors).map(v => ({ ...v, amount: v.actual - v.contract }));
}

function getFreightFilteredSummary() {
  const details = getFreightDetailsFiltered();
  if (details) {
    const threshold = DATA.freight.diffThreshold || 90;
    return {
      estimatedCost: details.reduce((s, r) => s + r.estimated, 0),
      actualCost: details.reduce((s, r) => s + r.actual, 0),
      totalOrders: details.length,
      overCount: details.filter(r => r.rate > threshold).length,
      saveCount: details.filter(r => r.rate <= threshold).length,
      vendors: summarizeFreightVendorsForRange(details),
    };
  }
  return {
    estimatedCost: DATA.freight.estimatedCost,
    actualCost: DATA.freight.actualCost,
    totalOrders: DATA.freight.totalOrders,
    overCount: DATA.freight.overCount,
    saveCount: DATA.freight.saveCount,
    vendors: DATA.freight.vendors,
  };
}

// F001 月總運費
const FREIGHT_WAREHOUSE_ORDER = ['大溪倉', '大肚倉', '岡山倉'];
const FREIGHT_DEMO_BUDGET = {
  '大溪倉': 33200602,
  '大肚倉': 17067880,
  '岡山倉': 21497185,
};

function getFreightWarehouseOrder() {
  const annual = DATA.annualBudget?.freight || {};
  const legacy = DATA.freight?.warehouseBudget || {};
  const names = Object.keys(annual).length ? Object.keys(annual) : Object.keys(legacy);
  const preferred = FREIGHT_WAREHOUSE_ORDER.filter(name => names.includes(name));
  const order = preferred.length === FREIGHT_WAREHOUSE_ORDER.length
    ? preferred
    : names.length ? names.slice(0, 3) : FREIGHT_WAREHOUSE_ORDER;
  const selectedWarehouse = getSelectedFreightWarehouse();
  return selectedWarehouse === 'all' ? order : order.filter(name => name === selectedWarehouse);
}

function getSelectedFreightWarehouse() {
  return FREIGHT_WAREHOUSE_ORDER.includes(DATA.freightSelectedWarehouse)
    ? DATA.freightSelectedWarehouse
    : 'all';
}

function hasRealFreightRows() {
  return getFreightDailyRowsFiltered().length > 0 || getFreightTrendFiltered().length > 0;
}

function freightDemoFullDate(day) {
  const base = DATA.dateFrom || '2026-03-01';
  const year = String(base).slice(0, 4) || '2026';
  const month = String(base).slice(5, 7) || '03';
  return `${year}-${month}-${String(day).padStart(2, '0')}`;
}

function getFreightDemoDailyRows() {
  const days = 18;
  return Array.from({ length: days }, (_, index) => {
    const day = index + 1;
    const weekWave = [0.92, 1.08, 1.02, 1.18, 0.88, 0.76, 0.82][index % 7];
    const campaignLift = index >= 8 && index <= 12 ? 1.28 : 1;
    const daxi = Math.round(980000 * weekWave * campaignLift + index * 18500);
    const dadu = Math.round(520000 * (1.06 - (index % 4) * 0.03) * campaignLift);
    const gangshan = Math.round(690000 * (0.9 + (index % 5) * 0.06));
    return [`${Number(String(DATA.dateFrom || '2026-03-01').slice(5, 7)) || 3}/${day}`, daxi, dadu, gangshan, freightDemoFullDate(day)];
  });
}

function getFreightDailyRowsForPage() {
  // 沒有該月實際資料就回空陣列（顯示 0），不再掉回展示資料。
  return getFreightDailyRowsFiltered();
}

function getFreightTrendRowsForPage() {
  // 沒有該月實際資料就回空陣列（顯示 0），不再掉回展示資料。
  return getFreightTrendFiltered();
}

function getFreightAnalysisMode() {
  return hasRealFreightRows() ? 'real' : 'demo';
}

function getFreightSummaryForPage() {
  // 一律使用實際過濾結果；沒資料時各欄位自然為 0 / 空。
  return getFreightFilteredSummary();
}

function getFreightBudgetMonthIndex() {
  if (typeof monthIndexFromDate === 'function') return monthIndexFromDate(DATA.dateFrom);
  const month = Number(String(DATA.dateFrom || '').slice(5, 7));
  return month >= 1 && month <= 12 ? month - 1 : 0;
}

function getFreightBudgetByWarehouse() {
  const order = getFreightWarehouseOrder();
  const monthIndex = getFreightBudgetMonthIndex();
  const source = DATA.annualBudget?.freight || {};
  const allowDbBudget = DATA.budgetSource === 'cloud';
  return order.reduce((map, warehouse) => {
    // 預算只讀資料庫年度預算；沒有就 0，不再使用展示預算。
    map[warehouse] = allowDbBudget ? Number(source[warehouse]?.[monthIndex] || 0) : 0;
    return map;
  }, {});
}

function getFreightActualByWarehouse() {
  const order = getFreightWarehouseOrder();
  const totals = order.reduce((map, warehouse) => {
    map[warehouse] = 0;
    return map;
  }, {});
  getFreightDailyRowsForPage().forEach(row => {
    order.forEach(warehouse => {
      const amountIndex = FREIGHT_WAREHOUSE_ORDER.indexOf(warehouse) + 1;
      totals[warehouse] += Number(row[amountIndex] || 0);
    });
  });
  return totals;
}

function getFreightAnalysisData() {
  const order = getFreightWarehouseOrder();
  const budget = getFreightBudgetByWarehouse();
  const actual = getFreightActualByWarehouse();
  const totalBudget = order.reduce((sum, warehouse) => sum + (budget[warehouse] || 0), 0);
  const totalActual = order.reduce((sum, warehouse) => sum + (actual[warehouse] || 0), 0);
  const rate = totalBudget ? totalActual / totalBudget * 100 : 0;
  const variance = totalBudget ? totalActual - totalBudget : null;
  return {
    order,
    budget,
    actual,
    totalBudget,
    totalActual,
    rate,
    variance,
    hasBudget: totalBudget > 0,
  };
}

function freightRiskLabel(rate, hasBudget) {
  if (!hasBudget) return '等待資料庫預算';
  if (rate >= 90) return '高風險';
  if (rate >= 75) return '需關注';
  return '健康';
}

function renderFreightKpiOverview() {
  const data = getFreightAnalysisData();
  const summary = getFreightSummaryForPage();
  const diff = data.variance;
  const riskColor = data.hasBudget ? colorFor(data.rate) : 'var(--ry-muted)';
  const isDemo = getFreightAnalysisMode() === 'demo';
  return `
  <div class="w s12 freight-kpi-panel">
    <div class="freight-kpi-card freight-kpi-primary">
      <div class="freight-kpi-label">月總運費 ${isDemo ? '<span class="freight-demo-badge">展示資料</span>' : ''}</div>
      <div class="freight-kpi-value">${fmtMoney(Math.round(data.totalActual))}</div>
      <div class="freight-kpi-note">${summary.totalOrders.toLocaleString()} 筆運單 · 依日期區間彙總</div>
    </div>
    <div class="freight-kpi-card">
      <div class="freight-kpi-label">資料庫預算</div>
      <div class="freight-kpi-value">${data.hasBudget ? fmtMoney(Math.round(data.totalBudget)) : '未接入'}</div>
      <div class="freight-kpi-note">${isDemo ? '目前用展示預算預覽版型' : '運費預算只讀取資料庫年度預算'}</div>
    </div>
    <div class="freight-kpi-card">
      <div class="freight-kpi-label">預算差異</div>
      <div class="freight-kpi-value" style="color:${diff === null ? 'var(--ry-muted)' : diff > 0 ? 'var(--ry-red)' : 'var(--ry-green)'}">
        ${diff === null ? '-' : `${diff > 0 ? '+' : ''}${fmtMoney(Math.round(diff))}`}
      </div>
      <div class="freight-kpi-note">${diff === null ? '需先有資料庫預算' : diff > 0 ? '實際高於預算' : '實際低於預算'}</div>
    </div>
    <div class="freight-kpi-card">
      <div class="freight-kpi-label">動支率</div>
      <div class="freight-kpi-value" style="color:${riskColor}">${data.hasBudget ? `${data.rate.toFixed(1)}%` : '-'}</div>
      <div class="freight-kpi-note">${freightRiskLabel(data.rate, data.hasBudget)}</div>
    </div>
  </div>`;
}

function renderFreightWarehouseBullets() {
  const data = getFreightAnalysisData();
  const maxRate = Math.max(100, ...data.order.map(warehouse => {
    const budget = data.budget[warehouse] || 0;
    return budget ? data.actual[warehouse] / budget * 100 : 0;
  }));
  const rows = data.order.map(warehouse => {
    const budget = data.budget[warehouse] || 0;
    const actual = data.actual[warehouse] || 0;
    const rate = budget ? actual / budget * 100 : 0;
    const width = Math.min(100, rate / maxRate * 100);
    const color = budget ? colorFor(rate) : 'var(--ry-muted)';
    return `
      <div class="freight-bullet-row">
        <div class="freight-bullet-name">${warehouse}</div>
        <div class="freight-bullet-track">
          <span class="freight-bullet-limit freight-bullet-limit-75"></span>
          <span class="freight-bullet-limit freight-bullet-limit-90"></span>
          <span class="freight-bullet-fill" style="width:${width}%;background:${color}"></span>
        </div>
        <div class="freight-bullet-value">${budget ? `${rate.toFixed(1)}%` : '-'}</div>
      </div>
      <div class="freight-bullet-sub">${fmtMoney(Math.round(actual))} / ${budget ? fmtMoney(Math.round(budget)) : '資料庫預算未接入'}</div>`;
  }).join('');

  return `
  <div class="w s6 freight-analysis-card">
    <div class="wh">
      <div class="wl"><div class="wdot dot-freight"></div>三倉動支率子彈圖</div>
      <span class="wmeta">75% / 90% 警戒線</span>
    </div>
    <div class="freight-bullet-list">${rows}</div>
  </div>`;
}

function renderFreightWaterfall() {
  const data = getFreightAnalysisData();
  const maxAbs = Math.max(
    1,
    data.totalBudget,
    data.totalActual,
    ...data.order.map(warehouse => Math.abs((data.actual[warehouse] || 0) - (data.budget[warehouse] || 0)))
  );
  const bars = [
    { label:'預算', value:data.totalBudget, kind:'base' },
    ...data.order.map(warehouse => ({
      label: warehouse,
      value: (data.actual[warehouse] || 0) - (data.budget[warehouse] || 0),
      kind: 'variance',
    })),
    { label:'實際', value:data.totalActual, kind:'actual' },
  ];
  const body = bars.map(bar => {
    const isVariance = bar.kind === 'variance';
    const isPositive = bar.value > 0;
    const width = Math.max(2, Math.abs(bar.value) / maxAbs * 100);
    const color = bar.kind === 'base'
      ? 'var(--ry-blue)'
      : bar.kind === 'actual'
        ? 'var(--ry-blue-dark)'
        : isPositive ? 'var(--ry-red)' : 'var(--ry-green)';
    return `
      <div class="freight-waterfall-item">
        <div class="freight-waterfall-label">${bar.label}</div>
        <div class="freight-waterfall-track">
          <span class="freight-waterfall-bar ${isVariance && !isPositive ? 'is-saving' : ''}" style="width:${width}%;background:${color}"></span>
        </div>
        <div class="freight-waterfall-value">${bar.value === 0 ? '-' : `${isVariance && isPositive ? '+' : ''}${fmtMoney(Math.round(bar.value))}`}</div>
      </div>`;
  }).join('');

  return `
  <div class="w s6 freight-analysis-card">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>預算差異瀑布圖</div>
      <span class="wmeta">${data.hasBudget ? '資料庫預算 vs 實際' : '等待資料庫預算'}</span>
    </div>
    <div class="freight-waterfall">${body}</div>
  </div>`;
}

function renderFreightDailyHeatmap() {
  const rows = getFreightTrendRowsForPage();
  const max = Math.max(1, ...rows.map(row => Number(row[1] || 0)));
  const cells = rows.map(row => {
    const value = Number(row[1] || 0);
    const level = Math.ceil(value / max * 5);
    return `
      <div class="freight-heat-cell level-${level}" title="${row[0]} ${fmtMoney(Math.round(value))}">
        <span>${String(row[0]).split('/').pop()}</span>
      </div>`;
  }).join('');

  return `
  <div class="w s12 freight-analysis-card">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>每日運費熱度圖</div>
      <span class="wmeta">${rows.length} 天 · 顏色越深代表當日運費越高</span>
    </div>
    ${rows.length ? `<div class="freight-heatmap">${cells}</div>` : '<div class="empty-state">此日期區間沒有運費資料</div>'}
  </div>`;
}

function renderFreightDecisionTable() {
  const data = getFreightAnalysisData();
  const rows = data.order.map(warehouse => {
    const budget = data.budget[warehouse] || 0;
    const actual = data.actual[warehouse] || 0;
    const diff = budget ? actual - budget : null;
    const rate = budget ? actual / budget * 100 : 0;
    const share = data.totalActual ? actual / data.totalActual * 100 : 0;
    const color = budget ? colorFor(rate) : 'var(--ry-muted)';
    return `
      <tr>
        <td class="date-strong">${warehouse}</td>
        <td class="mono num-right">${budget ? fmtMoney(Math.round(budget)) : '-'}</td>
        <td class="mono num-right actual-strong">${fmtMoney(Math.round(actual))}</td>
        <td class="mono num-right" style="color:${diff === null ? 'var(--ry-muted)' : diff > 0 ? 'var(--ry-red)' : 'var(--ry-green)'}">${diff === null ? '-' : `${diff > 0 ? '+' : ''}${fmtMoney(Math.round(diff))}`}</td>
        <td class="num-right"><span class="soft-pct-badge" style="background:${budget ? bgFor(rate) : '#f3f4f6'};color:${color}">${budget ? `${rate.toFixed(1)}%` : '-'}</span></td>
        <td>
          <div class="freight-share-bar"><span style="width:${Math.min(100, share)}%"></span></div>
          <div class="freight-share-text">${share.toFixed(1)}%</div>
        </td>
      </tr>`;
  }).join('');

  return `
  <div class="w s12 table-card freight-decision-card">
    <div class="wh">
      <div class="wl"><div class="wdot dot-freight"></div>倉別損益決策表</div>
      <span class="wmeta">排序固定為營運三倉</span>
    </div>
    <div class="table-edge">
      <table class="tbl freight-decision-table">
        <thead>
          <tr>
            <th>倉別</th>
            <th class="num-right">資料庫預算</th>
            <th class="num-right">實際運費</th>
            <th class="num-right">差異</th>
            <th class="num-right">動支率</th>
            <th>占總運費</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="table-note">運費預算由資料庫年度預算載入；若預算為空，頁面只呈現實際運費與占比。</div>
    </div>
  </div>`;
}

function renderFreightBudgetActualCombo() {
  const data = getFreightAnalysisData();
  const maxValue = Math.max(1, ...data.order.flatMap(warehouse => [
    data.budget[warehouse] || 0,
    data.actual[warehouse] || 0,
  ]));
  const W = 760, H = 280, padL = 52, padR = 44, padT = 28, padB = 46;
  const chartW = W - padL - padR;
  const chartH = H - padT - padB;
  const groupW = chartW / data.order.length;
  const barW = Math.min(42, groupW * 0.2);
  const yFor = value => padT + chartH - (Number(value || 0) / maxValue * chartH);
  const rateFor = warehouse => {
    const budget = data.budget[warehouse] || 0;
    return budget ? data.actual[warehouse] / budget * 100 : 0;
  };
  const maxRate = Math.max(100, ...data.order.map(rateFor));
  const yRate = rate => padT + chartH - (rate / maxRate * chartH);

  const bars = data.order.map((warehouse, index) => {
    const cx = padL + groupW * index + groupW / 2;
    const budget = data.budget[warehouse] || 0;
    const actual = data.actual[warehouse] || 0;
    const yBudget = yFor(budget);
    const yActual = yFor(actual);
    const rate = rateFor(warehouse);
    return `
      <rect x="${cx - barW - 4}" y="${yBudget}" width="${barW}" height="${padT + chartH - yBudget}" rx="4" fill="#3b82f6"></rect>
      <rect x="${cx + 4}" y="${yActual}" width="${barW}" height="${padT + chartH - yActual}" rx="4" fill="#fb6f92"></rect>
      <text x="${cx - barW / 2 - 4}" y="${yBudget - 7}" text-anchor="middle" font-size="10" fill="#123d74" font-family="Noto Sans TC">${(budget / 1000000).toFixed(1)}</text>
      <text x="${cx + barW / 2 + 4}" y="${yActual - 7}" text-anchor="middle" font-size="10" fill="#9f1239" font-family="Noto Sans TC">${(actual / 1000000).toFixed(1)}</text>
      <text x="${cx}" y="${H - 18}" text-anchor="middle" font-size="12" fill="#1a1d24" font-weight="700">${warehouse.replace('倉', '')}</text>
      <circle cx="${cx}" cy="${yRate(rate)}" r="4" fill="#1d4ed8" stroke="white" stroke-width="2"></circle>
      <text x="${cx}" y="${yRate(rate) - 10}" text-anchor="middle" font-size="11" fill="#1d4ed8" font-family="Noto Sans TC" font-weight="700">${rate.toFixed(1)}%</text>`;
  }).join('');

  const ratePath = data.order.map((warehouse, index) => {
    const cx = padL + groupW * index + groupW / 2;
    return `${index === 0 ? 'M' : 'L'}${cx},${yRate(rateFor(warehouse))}`;
  }).join(' ');

  return `
  <div class="w s6 freight-visual-card">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>倉別預算與實際動支</div>
      <span class="wmeta">藍：預算 · 紅：實際 · 線：動支率</span>
    </div>
    <svg class="freight-combo-chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="倉別預算與實際動支">
      <line x1="${padL}" y1="${padT + chartH}" x2="${padL + chartW}" y2="${padT + chartH}" stroke="#dde2ec"></line>
      <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${padT + chartH}" stroke="#dde2ec"></line>
      <line x1="${padL + chartW}" y1="${padT}" x2="${padL + chartW}" y2="${padT + chartH}" stroke="#dde2ec" stroke-dasharray="3 3"></line>
      <text x="${padL - 8}" y="${padT + 4}" text-anchor="end" font-size="10" fill="#5a6478" font-family="Noto Sans TC">${(maxValue / 1000000).toFixed(0)}M</text>
      <text x="${padL - 8}" y="${padT + chartH + 4}" text-anchor="end" font-size="10" fill="#5a6478" font-family="Noto Sans TC">0</text>
      <path d="${ratePath}" fill="none" stroke="#1d4ed8" stroke-width="2.5"></path>
      ${bars}
    </svg>
  </div>`;
}

function renderFreightStructureBars() {
  const data = getFreightAnalysisData();
  const maxValue = Math.max(1, ...data.order.flatMap(warehouse => [
    data.budget[warehouse] || 0,
    data.actual[warehouse] || 0,
  ]));
  const rows = data.order.map(warehouse => {
    const budget = data.budget[warehouse] || 0;
    const actual = data.actual[warehouse] || 0;
    const share = data.totalActual ? actual / data.totalActual * 100 : 0;
    return `
      <div class="freight-structure-row">
        <div class="freight-structure-name">${warehouse}</div>
        <div class="freight-structure-bars">
          <div class="freight-structure-bar budget" style="width:${budget / maxValue * 100}%"><span>${(budget / 1000000).toFixed(2)}M</span></div>
          <div class="freight-structure-bar actual" style="width:${actual / maxValue * 100}%"><span>${(actual / 1000000).toFixed(2)}M</span></div>
        </div>
        <div class="freight-share-ring" style="--share:${Math.min(100, share)}"><span>${share.toFixed(1)}%</span></div>
      </div>`;
  }).join('');
  return `
  <div class="w s6 freight-visual-card">
    <div class="wh">
      <div class="wl"><div class="wdot dot-freight"></div>費用結構與占比</div>
      <span class="wmeta">實際動支占總運費</span>
    </div>
    <div class="freight-structure">${rows}</div>
    <div class="freight-legend-row">
      <span><i class="legend-budget"></i>預算</span>
      <span><i class="legend-actual"></i>實際</span>
    </div>
  </div>`;
}

function renderFreightMonthlyBudgetMatrix() {
  const data = getFreightAnalysisData();
  const months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  const demoRatios = [0.29, 0.26, 0.27, 0.25, 0.28, 0.24, 0.31, 0.26, 0.29, 0.27];
  const currentMonth = (getFreightBudgetMonthIndex() || 0) + 1;

  function budgetFor(warehouse, month) {
    const source = DATA.annualBudget?.freight?.[warehouse]?.[month - 1] || 0;
    if (DATA.budgetSource === 'cloud' && source) return Number(source);
    const demoBase = FREIGHT_DEMO_BUDGET[warehouse] || 0;
    const drift = 1 + ((month % 3) - 1) * 0.035;
    return Math.round(demoBase * drift);
  }

  function actualFor(warehouse, month) {
    if (getFreightAnalysisMode() === 'real' && month === currentMonth) {
      return data.actual[warehouse] || 0;
    }
    const budget = budgetFor(warehouse, month);
    const warehouseOffset = data.order.indexOf(warehouse) * 0.018;
    return Math.round(budget * (demoRatios[month - 1] + warehouseOffset));
  }

  function valueCell(value, isRate = false, isDiff = false) {
    const color = isDiff && value > 0 ? 'var(--ry-red)' : isDiff && value < 0 ? 'var(--ry-green)' : '';
    return `<td class="mono num-right" style="${color ? `color:${color};font-weight:900` : ''}">${isRate ? `${value.toFixed(1)}%` : fmtMoney(Math.round(value))}</td>`;
  }

  const warehouseRows = data.order.map(warehouse => {
    const metricRows = [
      {
        label: '預算金額',
        values: months.map(month => budgetFor(warehouse, month)),
        type: 'money',
      },
      {
        label: '實際動支',
        values: months.map(month => actualFor(warehouse, month)),
        type: 'money',
      },
      {
        label: '預算差異',
        values: months.map(month => actualFor(warehouse, month) - budgetFor(warehouse, month)),
        type: 'diff',
      },
      {
        label: '動支率',
        values: months.map(month => {
          const budget = budgetFor(warehouse, month);
          return budget ? actualFor(warehouse, month) / budget * 100 : 0;
        }),
        type: 'rate',
      },
    ];
    return metricRows.map((row, index) => `
      <tr>
        ${index === 0 ? `<td class="freight-matrix-group" rowspan="${metricRows.length}">${warehouse}</td>` : ''}
        <td class="freight-matrix-metric">${row.label}</td>
        ${row.values.map(value => valueCell(value, row.type === 'rate', row.type === 'diff')).join('')}
      </tr>`).join('');
  }).join('');

  return `
  <div class="w s12 table-card freight-matrix-card">
    <div class="freight-matrix-title">月別預算動支明細</div>
    <div class="freight-matrix-scroll">
      <table class="tbl freight-matrix-table">
        <thead>
          <tr>
            <th>倉別</th>
            <th>資料指標</th>
            ${months.map(month => `<th class="num-right">${month}月</th>`).join('')}
          </tr>
        </thead>
        <tbody>${warehouseRows}</tbody>
      </table>
    </div>
  </div>`;
}

const FREIGHT_BUDGET_WEIGHT_DEFAULTS = { mon:1, tue:1.4, wed:1.4, thu:1.4, fri:1, sat:0.6, sun:0.5 };
let freightBudgetWeights = { ...FREIGHT_BUDGET_WEIGHT_DEFAULTS };
let freightBudgetHolidays = [];

function freightBudgetDayWeight(date) {
  if (freightBudgetHolidays.includes(date)) return freightBudgetWeights.sun;
  const day = new Date(`${date}T00:00:00`).getDay();
  return [
    freightBudgetWeights.sun,
    freightBudgetWeights.mon,
    freightBudgetWeights.tue,
    freightBudgetWeights.wed,
    freightBudgetWeights.thu,
    freightBudgetWeights.fri,
    freightBudgetWeights.sat,
  ][day];
}

function apportionFreightBudget(total, weights) {
  const weightTotal = weights.reduce((sum, value) => sum + value, 0);
  if (!weightTotal) return weights.map(() => 0);
  const raw = weights.map(value => total * value / weightTotal);
  const result = raw.map(Math.floor);
  let remainder = Math.round(total - result.reduce((sum, value) => sum + value, 0));
  raw.map((value, index) => ({ index, fraction:value - result[index] }))
    .sort((a, b) => b.fraction - a.fraction)
    .slice(0, remainder)
    .forEach(({ index }) => { result[index] += 1; });
  return result;
}

function toggleFreightBudgetPanel() {
  const body = document.getElementById('freight-budget-panel-body');
  const arrow = document.getElementById('freight-budget-panel-arrow');
  if (!body || !arrow) return;
  const opening = body.hidden;
  body.hidden = !opening;
  arrow.textContent = opening ? '▲ 收合' : '▼ 展開';
}

function applyFreightBudgetWeights() {
  const keys = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];
  keys.forEach(key => {
    const value = Number(document.getElementById(`freight-weight-${key}`)?.value);
    if (Number.isFinite(value) && value >= 0) freightBudgetWeights[key] = value;
  });
  freightBudgetHolidays = String(document.getElementById('freight-budget-holidays')?.value || '')
    .split(/[,\s]+/)
    .filter(value => /^\d{4}-\d{2}-\d{2}$/.test(value));
  renderFreightPage();
  toast('✅ 已更新運費每日預算分攤');
}

function resetFreightBudgetWeights() {
  freightBudgetWeights = { ...FREIGHT_BUDGET_WEIGHT_DEFAULTS };
  freightBudgetHolidays = [];
  renderFreightPage();
  toast('↺ 已還原預算分攤預設值');
}

function renderFreightReferenceDashboard() {
  const overview = getFreightAnalysisData();
  const summary = getFreightSummaryForPage();
  const actualTotal = overview.totalActual;
  const budgetTotal = overview.totalBudget;
  const variance = actualTotal - budgetTotal;
  // 以 dateTo（範圍結尾月）為基準，前月 = 結尾月-1。server 已在 lastMonthCost 算好整月金額；
  // 若 server 沒回值（舊 API），再 fallback 用本地 dailyByWarehouse 自行加總。
  const rangeEndStr = DATA.dateTo || DATA.dateFrom || new Date().toISOString().slice(0, 10);
  const rangeEnd = new Date(`${rangeEndStr}T00:00:00`);
  const previousMonthDate = new Date(rangeEnd.getFullYear(), rangeEnd.getMonth() - 1, 1);
  const previousMonthKey = `${previousMonthDate.getFullYear()}-${String(previousMonthDate.getMonth() + 1).padStart(2, '0')}`;
  const previousMonthDays = new Date(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1, 0).getDate();
  const selectedWarehouse = getSelectedFreightWarehouse();
  const serverPrevMonth = selectedWarehouse === 'all' ? Number(DATA.freight?.lastMonthCost || 0) : 0;
  const previousMonthTotal = serverPrevMonth > 0
    ? serverPrevMonth
    : (DATA.freight?.dailyByWarehouse || []).reduce((sum, row) => {
        const fullDate = row[4] || shortToFreightFullDate(row[0]);
        if (!fullDate?.startsWith(previousMonthKey)) return sum;
        if (selectedWarehouse === 'all') {
          return sum + Number(row[1] || 0) + Number(row[2] || 0) + Number(row[3] || 0);
        }
        const amountIndex = FREIGHT_WAREHOUSE_ORDER.indexOf(selectedWarehouse) + 1;
        return sum + Number(row[amountIndex] || 0);
      }, 0);
  const fullMonthDifference = actualTotal - previousMonthTotal;
  const previousMonthOrders = (DATA.freight?.details || []).filter(row =>
    row.fullDate?.startsWith(previousMonthKey) &&
    (selectedWarehouse === 'all' || row.budgetWarehouse === selectedWarehouse)
  ).length;
  const orderDifference = summary.totalOrders - previousMonthOrders;
  const selectedDetails = getFreightDetailsFiltered() || [];
  const nonMainlineDetails = selectedDetails.filter(row => row.sourceType === 'nonmainline');
  const mainlineDetails = selectedDetails.filter(row => !nonMainlineDetails.includes(row));
  const specialTruckCount = nonMainlineDetails.filter(row => row.categoryL2 === '專車').length;
  const nonMainlineRatio = selectedDetails.length ? nonMainlineDetails.length / selectedDetails.length * 100 : 0;
  const spendingRate = budgetTotal ? actualTotal / budgetTotal * 100 : 0;
  const spendingState = !budgetTotal ? 'muted' : spendingRate < 75 ? 'safe' : spendingRate <= 90 ? 'warning' : 'danger';
  const actualReasonRows = Object.values(nonMainlineDetails.reduce((map, row) => {
    const reason = String(row.reason || '').trim();
    if (!reason) return map;
    if (!map[reason]) map[reason] = { reason, cost:0, count:0 };
    map[reason].cost += Number(row.actual || 0);
    map[reason].count += 1;
    return map;
  }, {}));
  const reasonRows = actualReasonRows
    .map(row => ({ ...row, average: row.count ? row.cost / row.count : 0 }))
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 5);
  const actualRouteRows = Object.values(mainlineDetails.reduce((map, row) => {
    const route = String(row.route || row.vendor || '').trim();
    if (!route) return map;
    if (!map[route]) map[route] = { reason:route, cost:0, count:0 };
    map[route].cost += Number(row.actual || 0);
    map[route].count += 1;
    return map;
  }, {}));
  const routeRows = actualRouteRows
    .map(row => ({ ...row, average: row.count ? row.cost / row.count : 0 }))
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 5);
  function buildFreightRankChart(rows) {
    if (!rows.length) {
      return {
        bars: '<text x="350" y="120" text-anchor="middle" class="freight-ref-axis">尚無資料</text>',
        path: '',
      };
    }
    const maxCost = Math.max(1, ...rows.map(row => row.cost));
    const maxAverage = Math.max(1, ...rows.map(row => row.average));
    const bars = rows.map((row, index) => {
      const y = 32 + index * 43;
      const width = row.cost / maxCost * 420;
      const pointX = 176 + row.average / maxAverage * 420;
      const reason = escapeReleaseText(row.reason);
      return `
      <text x="160" y="${y + 12}" text-anchor="end" class="freight-ref-axis freight-ref-reason-name">${reason}</text>
      <rect x="176" y="${y}" width="${width}" height="16" rx="2" fill="var(--freight-kpi-blue)" class="freight-ref-chart-hit">
        <title>${reason}｜費用 $${(row.cost / 10000).toFixed(1)}W</title>
      </rect>
      <circle cx="${pointX}" cy="${y + 8}" r="12" fill="transparent" class="freight-ref-chart-hit">
        <title>${reason}｜平均單價 $${Math.round(row.average).toLocaleString()}</title>
      </circle>
      <circle cx="${pointX}" cy="${y + 8}" r="3.5" fill="var(--freight-kpi-pink)" pointer-events="none"></circle>`;
    }).join('');
    const path = rows.map((row, index) => {
      const x = 176 + row.average / maxAverage * 420;
      const y = 40 + index * 43;
      return `${index ? 'L' : 'M'}${x},${y}`;
    }).join(' ');
    return { bars, path };
  }
  const reasonChart = buildFreightRankChart(reasonRows);
  const routeChart = buildFreightRankChart(routeRows);
  const topReasonCost = reasonRows.reduce((sum, row) => sum + row.cost, 0);
  const topReasonCount = reasonRows.reduce((sum, row) => sum + row.count, 0);
  const topReasonAverage = topReasonCount ? topReasonCost / topReasonCount : 0;
  const topRouteCost = routeRows.reduce((sum, row) => sum + row.cost, 0);
  const topRouteCount = routeRows.reduce((sum, row) => sum + row.count, 0);
  const topRouteAverage = topRouteCount ? topRouteCost / topRouteCount : 0;

  const workGroups = [
    { category:'主線', items:['夜配'] },
    { category:'加派', items:['正物流', '逆物流'] },
    { category:'其他', items:['專車', '違規罰款', '離島海陸空運(馬祖)', '離島運費(澎湖、金門)', '全台共配費'] },
    { category:'轉運', items:['花蓮轉運費', '跨區轉運費'] },
  ];
  const workItems = workGroups.flatMap(group => group.items.map(item => ({ category:group.category, item })));
  const matrixDates = [];
  const matrixStart = new Date(`${DATA.dateFrom}T00:00:00`);
  const matrixEnd = new Date(`${DATA.dateTo}T00:00:00`);
  for (let date = new Date(matrixStart); date <= matrixEnd; date.setDate(date.getDate() + 1)) {
    matrixDates.push(`${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`);
  }
  function classifyFreightWorkItem(row) {
    // 非主線：採用匯入時 classifyNonMainline 寫入的權威分類；無法判斷暫歸專車
    if (row.sourceType === 'nonmainline') return row.categoryL2 || '專車';
    // 主線：一律夜配（業務確認主線目前只有夜配）
    return '夜配';
  }
  const workActualByDate = Object.fromEntries(workItems.map(({ item }) => [item, Object.fromEntries(matrixDates.map(date => [date, 0]))]));
  selectedDetails.forEach(row => {
    const item = classifyFreightWorkItem(row);
    if (workActualByDate[item] && workActualByDate[item][row.fullDate] !== undefined) {
      workActualByDate[item][row.fullDate] += Number(row.actual || 0);
    }
  });
  const workActualTotals = Object.fromEntries(workItems.map(({ item }) => [item, matrixDates.reduce((sum, date) => sum + workActualByDate[item][date], 0)]));
  const classifiedActualTotal = Object.values(workActualTotals).reduce((sum, value) => sum + value, 0);
  const workShares = Object.fromEntries(workItems.map(({ item }) => [item, classifiedActualTotal ? workActualTotals[item] / classifiedActualTotal : 1 / workItems.length]));
  const formatFreightWorkItemLabel = item => item === '離島海陸空運(馬祖)'
    ? '<span class="freight-ref-matrix-item-label">離島海陸空運<span>(馬祖)</span></span>'
    : item;
  const apportionedDailyBudget = apportionFreightBudget(budgetTotal, matrixDates.map(freightBudgetDayWeight));
  const dailyBudgetTotal = Object.fromEntries(matrixDates.map((date, index) => [date, apportionedDailyBudget[index]]));
  const workBudgetByDate = Object.fromEntries(workItems.map(({ item }) => [
    item,
    Object.fromEntries(matrixDates.map(date => [date, dailyBudgetTotal[date] * workShares[item]])),
  ]));
  const matrixValue = (value, type) => {
    if (type === 'rate') {
      const state = value > 90 ? 'danger' : value >= 75 ? 'warning' : 'safe';
      return `<td class="freight-ref-matrix-number rate ${state}">${value ? `${value.toFixed(1)}%` : '—'}</td>`;
    }
    return `<td class="freight-ref-matrix-number">${value ? Math.round(value).toLocaleString('zh-TW') : '—'}</td>`;
  };
  const matrixRows = workGroups.map(group => group.items.map((item, itemIndex) => {
    const budgetValues = matrixDates.map(date => workBudgetByDate[item][date]);
    const actualValues = matrixDates.map(date => workActualByDate[item][date]);
    const rateValues = actualValues.map((actual, index) => budgetValues[index] ? actual / budgetValues[index] * 100 : 0);
    const budgetTotalForItem = budgetValues.reduce((sum, value) => sum + value, 0);
    const actualTotalForItem = actualValues.reduce((sum, value) => sum + value, 0);
    const lines = [
      { label:'預算', type:'budget', values:budgetValues, total:budgetTotalForItem },
      { label:'實際', type:'actual', values:actualValues, total:actualTotalForItem },
      { label:'動支', type:'rate', values:rateValues, total:budgetTotalForItem ? actualTotalForItem / budgetTotalForItem * 100 : 0 },
    ];
    return lines.map((line, lineIndex) => `
      <tr class="matrix-${line.type} ${lineIndex === 0 ? 'work-start' : ''} ${itemIndex === 0 && lineIndex === 0 ? 'group-start' : ''}" data-work-category="${group.category}">
        ${itemIndex === 0 && lineIndex === 0 ? `<td rowspan="${group.items.length * 3}" class="freight-ref-matrix-category">${group.category}</td>` : ''}
        ${lineIndex === 0 ? `<td rowspan="3" class="freight-ref-matrix-item">${formatFreightWorkItemLabel(item)}</td>` : ''}
        <td class="freight-ref-matrix-metric">${line.label}</td>
        ${line.values.map(value => matrixValue(value, line.type)).join('')}
        ${matrixValue(line.total, line.type)}
      </tr>`).join('');
  }).join('')).join('');
  const totalBudgetValues = matrixDates.map(date => dailyBudgetTotal[date]);
  const totalActualValues = matrixDates.map(date => workItems.reduce((sum, { item }) => sum + workActualByDate[item][date], 0));
  const totalRateValues = totalActualValues.map((actual, index) => totalBudgetValues[index] ? actual / totalBudgetValues[index] * 100 : 0);
  const totalBudgetAmount = totalBudgetValues.reduce((sum, value) => sum + value, 0);
  const totalActualAmount = totalActualValues.reduce((sum, value) => sum + value, 0);
  const mainlineActualValues = matrixDates.map(date => mainlineDetails
    .filter(row => row.fullDate === date)
    .reduce((sum, row) => sum + Number(row.actual || 0), 0));
  const nonMainlineActualValues = matrixDates.map(date => nonMainlineDetails
    .filter(row => row.fullDate === date)
    .reduce((sum, row) => sum + Number(row.actual || 0), 0));
  const trendMax = Math.max(1, ...totalBudgetValues, ...totalActualValues);
  const trendW = 1240;
  const trendH = 250;
  const trendPadL = 58;
  const trendPadR = 14;
  const trendPadT = 12;
  const trendPadB = 28;
  const trendPlotH = trendH - trendPadT - trendPadB;
  const trendBarW = (trendW - trendPadL - trendPadR) / Math.max(1, matrixDates.length);
  const trendY = value => trendPadT + trendPlotH - value / trendMax * trendPlotH;
  const trendGrid = Array.from({ length:5 }, (_, index) => {
    const value = trendMax / 4 * index;
    const y = trendY(value);
    return `<line x1="${trendPadL}" y1="${y}" x2="${trendW - trendPadR}" y2="${y}" stroke="var(--app-border)" stroke-width="0.5"></line>
      <text x="${trendPadL - 7}" y="${y + 4}" text-anchor="end" class="freight-ref-axis">${Math.round(value / 10000)}萬</text>`;
  }).join('');
  const trendBars = matrixDates.map((date, index) => {
    const x = trendPadL + index * trendBarW + 2;
    const width = Math.max(2, trendBarW - 4);
    const main = mainlineActualValues[index];
    const nonmain = nonMainlineActualValues[index];
    const total = main + nonmain;
    const label = Number(date.slice(8, 10));
    return `<rect x="${x}" y="${trendY(main)}" width="${width}" height="${trendH - trendPadB - trendY(main)}" fill="var(--freight-kpi-blue)">
        <title>${date}｜主線 $${Math.round(main).toLocaleString()}｜非主線 $${Math.round(nonmain).toLocaleString()}｜預算 $${Math.round(totalBudgetValues[index]).toLocaleString()}</title>
      </rect>
      <rect x="${x}" y="${trendY(total)}" width="${width}" height="${trendY(main) - trendY(total)}" fill="var(--freight-kpi-blue)" opacity="0.4"></rect>
      ${label % 2 ? `<text x="${x + width / 2}" y="${trendH - 9}" text-anchor="middle" class="freight-ref-axis">${label}</text>` : ''}`;
  }).join('');
  const budgetTrendPath = totalBudgetValues.map((value, index) => {
    const x = trendPadL + index * trendBarW + trendBarW / 2;
    return `${index ? 'L' : 'M'}${x},${trendY(value)}`;
  }).join(' ');
  const weightLabels = [
    ['mon', '週一'], ['tue', '週二'], ['wed', '週三'], ['thu', '週四'],
    ['fri', '週五'], ['sat', '週六'], ['sun', '週日／假日'],
  ];
  const weightInputs = weightLabels.map(([key, label]) => `
    <label class="freight-budget-weight">
      <span>${label}</span>
      <input type="number" min="0" step="0.1" id="freight-weight-${key}" value="${freightBudgetWeights[key]}">
    </label>`).join('');
  const totalRows = [
    { label:'預算', type:'budget', values:totalBudgetValues, total:totalBudgetAmount },
    { label:'實際', type:'actual', values:totalActualValues, total:totalActualAmount },
    { label:'動支', type:'rate', values:totalRateValues, total:totalBudgetAmount ? totalActualAmount / totalBudgetAmount * 100 : 0 },
  ].map((line, index) => `
    <tr class="total matrix-${line.type} ${index === 0 ? 'work-start group-start' : ''}">
      ${index === 0 ? '<td rowspan="3" colspan="2" class="freight-ref-matrix-item">總計</td>' : ''}
      <td class="freight-ref-matrix-metric">${line.label}</td>
      ${line.values.map(value => matrixValue(value, line.type)).join('')}
      ${matrixValue(line.total, line.type)}
    </tr>`).join('');

  return `
  <section class="freight-ref-page">
    <div class="freight-ref-section">
      <div class="freight-ref-kpis">
        <article class="freight-ref-kpi">
          <div class="freight-ref-kpi-icon" aria-hidden="true">💰</div>
          <div class="freight-ref-kpi-label">費用總計</div>
          <div class="freight-ref-kpi-main">$${(actualTotal / 10000).toFixed(2)}W</div>
          <div class="freight-ref-kpi-ctx">
            前月（${previousMonthDays}天）${previousMonthTotal ? `$${(previousMonthTotal / 10000).toFixed(2)}W` : '—'}${previousMonthTotal ? `<span class="freight-ref-delta ${fullMonthDifference <= 0 ? 'good' : 'bad'}">${fullMonthDifference < 0 ? '▼' : '▲'} $${Math.abs(fullMonthDifference / 10000).toFixed(2)}W</span>` : ''}
          </div>
        </article>
        <article class="freight-ref-kpi">
          <div class="freight-ref-kpi-icon" aria-hidden="true">📊</div>
          <div class="freight-ref-kpi-label">動支率</div>
          <div class="freight-ref-kpi-main">${spendingRate.toFixed(1)}%</div>
          <div class="freight-ref-progress state-${spendingState}" data-tooltip="綠色：低於 75% 安全｜黃色：75–90% 注意｜紅色：高於 90% 危險"><i style="width:${Math.min(100, spendingRate)}%"></i></div>
          <div class="freight-ref-kpi-ctx freight-ref-ctx-state-${spendingState}">${!budgetTotal ? '未設定預算' : spendingRate < 75 ? `安全 · < 75%` : spendingRate <= 90 ? `注意 · 75–90%` : `超支 · > 90%`}</div>
        </article>
        <article class="freight-ref-kpi">
          <div class="freight-ref-kpi-icon" aria-hidden="true">📦</div>
          <div class="freight-ref-kpi-label">配送筆數</div>
          <div class="freight-ref-kpi-main">${summary.totalOrders.toLocaleString()}</div>
          <div class="freight-ref-kpi-ctx">
            前月（${previousMonthDays}天）${previousMonthOrders ? `${previousMonthOrders.toLocaleString()} 筆` : '—'}${previousMonthOrders ? `<span class="freight-ref-delta neutral">${orderDifference < 0 ? '▼' : '▲'} ${Math.abs(orderDifference).toLocaleString()} 筆</span>` : ''}
          </div>
        </article>
        <article class="freight-ref-kpi gauge">
          <div class="freight-ref-kpi-icon" aria-hidden="true">🚚</div>
          <div class="freight-ref-kpi-label">非主線占比</div>
          <div class="freight-ref-gauge" style="--gauge-deg:${Math.min(180, nonMainlineRatio * 1.8)}deg">
            <div class="freight-ref-gauge-value">${nonMainlineRatio.toFixed(1)}%</div>
          </div>
          <div class="freight-ref-kpi-ctx">非主線 ${nonMainlineDetails.length.toLocaleString()} 筆 · 其中專車 ${specialTruckCount.toLocaleString()} 筆</div>
        </article>
      </div>
    </div>

    <section class="freight-budget-trend-card">
      <div class="freight-ref-card-title">每日費用趨勢</div>
      <div class="freight-trend-legend">
        <span><i class="mainline"></i>主線</span>
        <span><i class="nonmainline"></i>非主線</span>
        <span><i class="budget"></i>每日預算分攤</span>
      </div>
      <svg class="freight-budget-trend-chart" viewBox="0 0 ${trendW} ${trendH}">
        ${trendGrid}
        ${trendBars}
        <path d="${budgetTrendPath}" fill="none" stroke="var(--app-warning)" stroke-width="2" stroke-dasharray="6 4"></path>
      </svg>
      <button type="button" class="freight-budget-panel-toggle" onclick="toggleFreightBudgetPanel()">
        <span><strong>⚙️ 預算分攤控制台</strong>　物流量集中在非國定假日的週二～四，權重可調</span>
        <span id="freight-budget-panel-arrow">▼ 展開</span>
      </button>
      <div class="freight-budget-panel-body" id="freight-budget-panel-body" hidden>
        <p>每日預算 ＝ 月預算 × 當日權重 ÷ 全月權重總和。國定假日採週日權重；僅影響本頁預算分攤顯示，不修改資料庫預算。</p>
        <div class="freight-budget-weights">${weightInputs}</div>
        <div class="freight-budget-actions">
          <label>
            <span>國定假日（YYYY-MM-DD，逗號或換行分隔）</span>
            <textarea id="freight-budget-holidays">${freightBudgetHolidays.join('\n')}</textarea>
          </label>
          <div>
            <button class="btn btn-primary" onclick="applyFreightBudgetWeights()">套用</button>
            <button class="btn btn-ghost" onclick="resetFreightBudgetWeights()">還原預設</button>
          </div>
        </div>
      </div>
    </section>

    <section class="freight-ref-matrix-section">
      <div class="freight-ref-matrix-heading">
        <div class="freight-ref-card-title freight-ref-matrix-title">每日作業項目預算動支明細</div>
        <label class="freight-ref-matrix-view">
          <span>查看</span>
          <select class="filter-input" id="freight-work-category-filter" onchange="filterFreightWorkMatrix(this.value)">
            <option value="all">全部</option>
            <option value="主線">主線</option>
            <option value="加派">加派</option>
            <option value="其他">其他</option>
            <option value="轉運">轉運</option>
          </select>
        </label>
      </div>
      <article class="freight-ref-matrix-card">
      <div class="freight-ref-matrix-scroll">
        <table class="freight-ref-matrix">
          <thead>
            <tr>
              <th>類別</th>
              <th>作業項目</th>
              <th>項目</th>
              ${matrixDates.map(date => `<th>${Number(date.slice(5, 7))}月${Number(date.slice(8, 10))}日</th>`).join('')}
              <th>合計</th>
            </tr>
          </thead>
          <tbody>${matrixRows}${totalRows}</tbody>
        </table>
      </div>
      </article>
    </section>
    <section class="freight-exception-panel">
      <div>
        <div class="freight-ref-card-title">⊘ 不列入預算明細</div>
        <p>上收、誤 key 與尚未完成分類的資料，不計入動支率與費用總計。</p>
      </div>
      <span>目前後端 API 尚未提供不列入明細，待資料介面完成後即可在此檢視與分類。</span>
    </section>
  </section>`;
}

function filterFreightWorkMatrix(category = 'all') {
  document.querySelectorAll('.freight-ref-matrix tbody tr[data-work-category]').forEach(row => {
    row.hidden = category !== 'all' && row.dataset.workCategory !== category;
  });
}

// T003 每日動支明細：依領域（全部/人力/運務）過濾
function filterT003Domain(value = 'all') {
  document.querySelectorAll('.freight-ref-matrix tr[data-t003-domain]').forEach(row => {
    row.hidden = value !== 'all' && row.dataset.t003Domain !== value;
  });
}

function renderF001() {
  const f = DATA.freight;
  const daily = getFreightTrendFiltered();
  const summary = getFreightFilteredSummary();
  const totalCost = daily.reduce((s, d) => s + d[1], 0);
  const trend = f.lastMonthCost ? (totalCost - f.lastMonthCost) / f.lastMonthCost * 100 : 0;
  const up = trend >= 0;
  const trendColor = up ? '#d9401b' : '#1b7c33';
  return `
  <div class="w s4 widget-border-blue">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>F001 月總運費</div>
      <span class="wmeta">本期結算</span>
    </div>
    <div class="kv kv-lg text-blue">${fmtMoney(totalCost)}</div>
    <div class="kd">
      <span class="trend-value" style="color:${trendColor}">${up ? '↑' : '↓'} ${Math.abs(trend).toFixed(1)}%</span>
      較上月（${fmtMoney(f.lastMonthCost)}）· ${summary.totalOrders.toLocaleString()} 筆配送
    </div>
  </div>`;
}

// F002 預計 vs 實際差異率
function renderF002() {
  const summary = getFreightFilteredSummary();
  const diff = summary.actualCost - summary.estimatedCost;
  const diffPct = summary.estimatedCost ? (diff / summary.estimatedCost * 100) : 0;
  const color = diff > 0 ? '#d9401b' : '#1b7c33';
  const label = diff > 0 ? '🔴 超支' : '🟢 節省';
  return `
  <div class="w s4">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${color}"></div>F002 預計 vs 實際差異率</div>
      <span class="wmeta">${label}</span>
    </div>
    <div class="kv kv-lg" style="color:${color}">${diffPct > 0 ? '+' : ''}${diffPct.toFixed(2)}%</div>
    <div class="kd">
      實際 ${fmtMoney(summary.actualCost)}<br>
      預計 ${fmtMoney(summary.estimatedCost)}（差異 ${diff > 0 ? '+' : ''}${fmtMoney(diff)}）
    </div>
  </div>`;
}

// F003 超支/節省筆數
function renderF003() {
  const f = DATA.freight;
  const summary = getFreightFilteredSummary();
  const total = summary.overCount + summary.saveCount;
  const overPct = total ? (summary.overCount / total * 100).toFixed(1) : '0.0';
  return `
  <div class="w s4 widget-border-red">
    <div class="wh">
      <div class="wl"><div class="wdot dot-red"></div>F003 超支／節省筆數</div>
      <span class="wmeta">動支率 &gt; ${f.diffThreshold || 90}%</span>
    </div>
    <div class="f003-count-row">
      <div>
        <div class="f003-label text-red">🔴 超支</div>
        <div class="f003-value text-red">${summary.overCount}</div>
      </div>
      <div class="f003-divider">／</div>
      <div>
        <div class="f003-label text-safe">🟢 節省</div>
        <div class="f003-value text-safe">${summary.saveCount}</div>
      </div>
    </div>
    <div class="kd">共 ${total.toLocaleString()} 筆明細 · 動支率 &gt; ${f.diffThreshold || 90}% 佔 ${overPct}%</div>
  </div>`;
}

// F009 日費用趨勢（折線圖）
function renderF009() {
  const data = getFreightTrendRowsForPage();
  if (!data.length) {
    return `
  <div class="w s12 table-card">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>F009 日費用趨勢</div>
      <span class="wmeta">0 天</span>
    </div>
    <div class="empty-state">此日期區間沒有運費資料</div>
  </div>`;
  }
  const maxCost = Math.max(...data.map(d => d[1]));
  const minCost = Math.min(...data.map(d => d[1]));
  const avgCost = data.reduce((s, d) => s + d[1], 0) / data.length;

  const W = 1000, H = 240, padL = 56, padR = 20, padT = 20, padB = 36;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const xStep = data.length > 1 ? plotW / (data.length - 1) : plotW;
  const range = maxCost - minCost || 1;
  const yFor = v => padT + plotH - ((v - minCost) / range * plotH);

  const pathD = data.map((d, i) => {
    const x = padL + i * xStep;
    const y = yFor(d[1]);
    return `${i === 0 ? 'M' : 'L'}${x.toFixed(1)},${y.toFixed(1)}`;
  }).join(' ');

  const areaD = pathD + ` L${padL + plotW},${padT + plotH} L${padL},${padT + plotH} Z`;

  const xLabels = data.map((d, i) => i % 4 === 0 || i === data.length - 1
    ? `<text x="${padL + i * xStep}" y="${H - 10}" text-anchor="middle" font-size="10" fill="#5a6478" font-family="Noto Sans TC">${d[0]}</text>`
    : '').join('');

  const dots = data.map((d, i) =>
    `<circle cx="${padL + i * xStep}" cy="${yFor(d[1])}" r="3" fill="#1e5ca8" stroke="white" stroke-width="1.5"><title>${d[0]}：${fmtMoney(d[1])}</title></circle>`
  ).join('');

  const avgY = yFor(avgCost);

  const yTicks = [minCost, (minCost + maxCost) / 2, maxCost].map(v =>
    `<line x1="${padL}" y1="${yFor(v)}" x2="${padL + plotW}" y2="${yFor(v)}" stroke="#dde2ec" stroke-width="0.5" stroke-dasharray="3,3"/>
     <text x="${padL - 8}" y="${yFor(v) + 3}" text-anchor="end" font-size="9" fill="#5a6478" font-family="Noto Sans TC">${(v/1000).toFixed(0)}K</text>`
  ).join('');

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>F009 日費用趨勢</div>
      <span class="wmeta">${data.length} 天 · 平均 ${fmtMoney(Math.round(avgCost))}</span>
    </div>
    <svg class="chart-svg" viewBox="0 0 ${W} ${H}">
      ${yTicks}
      <line x1="${padL}" y1="${avgY}" x2="${padL + plotW}" y2="${avgY}" stroke="#e07855" stroke-width="1" stroke-dasharray="5,4"/>
      <text x="${padL + plotW - 4}" y="${avgY - 4}" text-anchor="end" font-size="10" fill="#e07855" font-family="Noto Sans TC">平均 ${fmtMoney(Math.round(avgCost))}</text>
      <path d="${areaD}" fill="#1e5ca8" fill-opacity="0.08"/>
      <path d="${pathD}" fill="none" stroke="#1e5ca8" stroke-width="2" stroke-linejoin="round"/>
      ${dots}
      ${xLabels}
    </svg>
    <div class="chart-note">
      📌 最高 ${fmtMoney(maxCost)} · 最低 ${fmtMoney(minCost)} · 橘色虛線為日均 · 滑鼠移到點上可看詳細數字
    </div>
  </div>`;
}

// F010 三倉每日運費動支
function renderF010() {
  const budget = getFreightBudgetByWarehouse();
  const daily  = getFreightDailyRowsForPage();
  const totalDays = daily.length;

  if (!totalDays) {
    return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot dot-freight"></div>F010 三倉每日運費動支</div>
      <span class="wmeta">0 天</span>
    </div>
    <div class="empty-state">此日期區間沒有運費資料</div>
  </div>`;
  }

  const dailyBudget = {
    '大溪倉': budget['大溪倉'] / totalDays,
    '大肚倉': budget['大肚倉'] / totalDays,
    '岡山倉': budget['岡山倉'] / totalDays,
  };

  const accumActual = { '大溪倉':0, '岡山倉':0, '大肚倉':0 };
  daily.forEach(row => {
    accumActual['大溪倉'] += row[1];
    accumActual['大肚倉'] += row[2];
    accumActual['岡山倉'] += row[3];
  });

  function cellsFor(actual, budgetVal) {
    const pct = budgetVal ? actual / budgetVal * 100 : 0;
    const color = colorFor(pct);
    const bg = bgFor(pct);
    return `
      <td class="mono num-right">${fmtMoney(Math.round(budgetVal))}</td>
      <td class="mono num-right actual-strong" style="color:${color}">${fmtMoney(actual)}</td>
      <td class="mono num-right">
        <span class="soft-pct-badge" style="background:${bg};color:${color}">${pct.toFixed(1)}%</span>
      </td>`;
  }

  const rows = daily.map(row => {
    const [date, d1, d2, d3] = row;
    return `<tr>
      <td class="date-strong">${date}</td>
      ${cellsFor(d1, dailyBudget['大溪倉'])}
      ${cellsFor(d2, dailyBudget['大肚倉'])}
      ${cellsFor(d3, dailyBudget['岡山倉'])}
    </tr>`;
  }).join('');

  function summaryCell(actual, budgetVal) {
    const pct = budgetVal ? actual / budgetVal * 100 : 0;
    const color = colorFor(pct);
    return `
      <td class="mono num-right">${fmtMoney(budgetVal)}</td>
      <td class="mono num-right actual-heavy" style="color:${color}">${fmtMoney(actual)}</td>
      <td class="mono num-right">
        <span class="solid-pct-badge" style="background:${color}">${pct.toFixed(1)}%</span>
      </td>`;
  }

  const summary = `<tr class="f010-summary-row">
    <td class="f010-summary-label">月累計</td>
    ${summaryCell(accumActual['大溪倉'], budget['大溪倉'])}
    ${summaryCell(accumActual['大肚倉'], budget['大肚倉'])}
    ${summaryCell(accumActual['岡山倉'], budget['岡山倉'])}
  </tr>`;

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot dot-freight"></div>F010 三倉每日運費動支</div>
      <span class="wmeta">${totalDays} 天 · 縱向捲動 · 含月累計彙總</span>
    </div>
    <div class="table-edge">
      <div class="f010-scroll">
        <table class="tbl f010-table">
          <thead class="sticky-thead">
            <tr class="f010-head-row">
              <th rowspan="2" class="f010-date-head">日期</th>
              <th colspan="3" class="f010-wh-head f010-daxi">🏭 大溪倉</th>
              <th colspan="3" class="f010-wh-head f010-dadu">🏭 大肚倉</th>
              <th colspan="3" class="f010-wh-head f010-gangshan">🏭 岡山倉</th>
            </tr>
            <tr class="f010-subhead-row">
              <th class="num-right">預算</th><th class="num-right">實際</th><th class="num-right">動支率</th>
              <th class="num-right">預算</th><th class="num-right">實際</th><th class="num-right">動支率</th>
              <th class="num-right">預算</th><th class="num-right">實際</th><th class="num-right">動支率</th>
            </tr>
          </thead>
          <tbody>${rows}${summary}</tbody>
        </table>
      </div>
      <div class="table-note">
        📌 單日預算 = 月預算 ÷ ${totalDays} 天 · 動支率 = 實際 / 單日預算<br>
        📌 三色門檻：&lt; 75% 🟢 安全 · 75–90% 🟡 注意 · &gt; 90% 🔴 危險
      </div>
    </div>
  </div>`;
}

// ── 總費用動支共用工具 ──

function getDispatchDailyFiltered() {
  return DATA.dispatch.daily.filter(row => {
    const fullDate = dispatchRowFullDate(row);
    return fullDate >= DATA.dateFrom && fullDate <= DATA.dateTo;
  });
}

function dispatchRowFullDate(row) {
  if (row[7]) return row[7];
  const year = (DATA.dateFrom || '').slice(0, 4) || String(new Date().getFullYear());
  const parts = String(row[0] || '').split('/');
  if (parts.length !== 2) return '';
  return `${year}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`;
}

function daysInDispatchMonth() {
  const y = Number((DATA.dateFrom || '').slice(0, 4));
  const m = Number((DATA.dateFrom || '').slice(5, 7));
  if (!y || !m) return 30;
  return new Date(y, m, 0).getDate();
}

function dispatchBudgetMonthLabel() {
  const m = Number((DATA.dateFrom || '').slice(5, 7));
  return m >= 1 && m <= 12 ? `${m}月` : '目前月份';
}

function hasAnnualDispatchBudget() {
  const labor = DATA.annualBudget?.labor || {};
  const freight = DATA.annualBudget?.freight || {};
  return Object.keys(DATA.dispatch.budget || {}).some(wh =>
    (labor[wh] || []).some(v => v > 0) || (freight[wh] || []).some(v => v > 0)
  );
}

function monthIndexFromDate(dateStr) {
  const m = Number(String(dateStr || '').slice(5, 7));
  return m >= 1 && m <= 12 ? m - 1 : 0;
}

function daysInMonthFor(dateStr) {
  const y = Number(String(dateStr || '').slice(0, 4));
  const m = Number(String(dateStr || '').slice(5, 7));
  if (!y || !m) return 30;
  return new Date(y, m, 0).getDate();
}

function monthlyDispatchBudget(warehouse, monthIndex) {
  if (hasAnnualDispatchBudget()) {
    return {
      labor: DATA.annualBudget.labor?.[warehouse]?.[monthIndex] || 0,
      freight: DATA.annualBudget.freight?.[warehouse]?.[monthIndex] || 0,
    };
  }

  return DATA.dispatch.budget[warehouse] || { labor: 0, freight: 0 };
}

function dailyDispatchBudget(warehouse, dateStr) {
  const monthly = monthlyDispatchBudget(warehouse, monthIndexFromDate(dateStr));
  const days = daysInMonthFor(dateStr);
  return {
    labor: monthly.labor / days,
    freight: monthly.freight / days,
  };
}

function selectedDispatchBudget(warehouse) {
  const dates = selectedDispatchDateRange();
  return dates.reduce((sum, dateStr) => {
    const daily = dailyDispatchBudget(warehouse, dateStr);
    sum.labor += daily.labor;
    sum.freight += daily.freight;
    return sum;
  }, { labor: 0, freight: 0 });
}

function selectedDispatchDayCount() {
  return selectedDispatchDateRange().length;
}

function selectedDispatchDateRange() {
  const from = new Date(`${DATA.dateFrom}T00:00:00`);
  const to = new Date(`${DATA.dateTo}T00:00:00`);
  if (Number.isNaN(from.getTime()) || Number.isNaN(to.getTime()) || to < from) return [];
  const dates = [];
  for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    dates.push(`${yyyy}-${mm}-${dd}`);
  }
  return dates;
}

function getDispatchLatestUploadDate() {
  if (DATA.dispatch.latestUploadDate) return DATA.dispatch.latestUploadDate;
  const dates = DATA.dispatch.daily.map(dispatchRowFullDate).filter(Boolean).sort();
  return dates.length ? dates[dates.length - 1] : '';
}

function sumWarehouse(rows, warehouseIdx) {
  const names = ['大溪倉', '大肚倉', '岡山倉'];
  const laborCol = 1 + warehouseIdx * 2;
  const freightCol = 2 + warehouseIdx * 2;
  let labor = 0, freight = 0;
  rows.forEach(r => { labor += r[laborCol]; freight += r[freightCol]; });
  const total = labor + freight;
  const b = selectedDispatchBudget(names[warehouseIdx]);
  const laborBudget   = b.labor;
  const freightBudget = b.freight;
  const periodBudget  = laborBudget + freightBudget;
  return {
    name: names[warehouseIdx],
    labor, freight, total,
    laborBudget, freightBudget, budget: periodBudget,
    laborPct:   laborBudget   ? labor   / laborBudget   * 100 : 0,
    freightPct: freightBudget ? freight / freightBudget * 100 : 0,
    pct:        periodBudget  ? total   / periodBudget  * 100 : 0,
  };
}

function sumAll(rows) {
  const w1 = sumWarehouse(rows, 0);
  const w2 = sumWarehouse(rows, 1);
  const w3 = sumWarehouse(rows, 2);
  const labor   = w1.labor   + w2.labor   + w3.labor;
  const freight = w1.freight + w2.freight + w3.freight;
  const laborBudget   = w1.laborBudget   + w2.laborBudget   + w3.laborBudget;
  const freightBudget = w1.freightBudget + w2.freightBudget + w3.freightBudget;
  const total = labor + freight;
  const budget = laborBudget + freightBudget;
  return {
    name: '全區',
    labor, freight, total,
    laborBudget, freightBudget, budget,
    laborPct:   laborBudget   ? labor   / laborBudget   * 100 : 0,
    freightPct: freightBudget ? freight / freightBudget * 100 : 0,
    pct:        budget        ? total   / budget        * 100 : 0,
  };
}

// T001 總動支關鍵數據卡片
function renderT001() {
  const rows = getDispatchDailyFiltered();
  const all      = sumAll(rows);
  const daxi     = sumWarehouse(rows, 0);
  const dadu     = sumWarehouse(rows, 1);
  const gangshan = sumWarehouse(rows, 2);

  const colors = {
    '全區':   'var(--ry-blue-dark)',
    '大溪倉': '#0E7BAD',
    '大肚倉': '#2DA870',
    '岡山倉': '#E07855',
  };

  const toWan = (v) => `$${(Number(v || 0) / 10000).toLocaleString('zh-TW', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}萬`;

  function card(stats, isAll) {
    const color = colors[stats.name];
    const rateColor = colorFor(stats.pct);
    return `
    <div class="w s3 t001-card">
      <div class="wh">
        <div class="wl"><div class="wdot" style="background:${color}"></div>${isAll ? '全區' : stats.name}</div>
      </div>
      <span class="t001-status">${labelFor(stats.pct)}</span>
      <div class="t001-rate-row">
        <div class="kv kv-xl" style="color:${rateColor}">${stats.pct.toFixed(1)}</div>
        <div class="t001-rate-unit" style="color:${rateColor}">%</div>
      </div>
      <div class="t001-cost-grid">
        <div class="t001-cost-col">
          <div class="mini-label">人力費用</div>
          <div class="mono mini-value">${toWan(stats.labor)}</div>
        </div>
        <div class="t001-cost-col">
          <div class="mini-label">運務費用</div>
          <div class="mono mini-value">${toWan(stats.freight)}</div>
        </div>
      </div>
      <div class="t001-hover-card">
        <div class="t001-hover-row"><span>總實際</span><span>${fmtMoney(stats.total)}</span></div>
        <div class="t001-hover-row"><span>總預算</span><span>${fmtMoney(Math.round(stats.budget))}</span></div>
      </div>
    </div>`;
  }

  return [card(all, true), card(daxi), card(dadu), card(gangshan)].join('');
}

// T002 期間動支彙總表
function renderT002() {
  const rows = getDispatchDailyFiltered();
  const daxi     = sumWarehouse(rows, 0);
  const dadu     = sumWarehouse(rows, 1);
  const gangshan = sumWarehouse(rows, 2);
  const all      = sumAll(rows);
  const warehouses = [daxi, dadu, gangshan, all];

  // 數值/動支率儲存格：沿用 freight-ref-matrix 樣式
  const numCell = (v) => `<td class="freight-ref-matrix-number">${fmtMoney(Math.round(v || 0))}</td>`;
  const rateCell = (pct) => {
    const state = pct > 90 ? 'danger' : pct >= 75 ? 'warning' : 'safe';
    return `<td class="freight-ref-matrix-number rate ${state}">${(pct || 0).toFixed(1)}%</td>`;
  };

  const domains = [
    { label:'人力', icon:'💰', key:{ b:'laborBudget',   a:'labor',   p:'laborPct'   } },
    { label:'運務', icon:'🚚', key:{ b:'freightBudget', a:'freight', p:'freightPct' } },
    { label:'合計', icon:'📊', key:{ b:'budget',        a:'total',   p:'pct'        } },
  ];

  const matrixRows = domains.map(d => {
    const lines = [
      { label:'預算', type:'budget', cells: warehouses.map(w => numCell(w[d.key.b])).join('') },
      { label:'實際', type:'actual', cells: warehouses.map(w => numCell(w[d.key.a])).join('') },
      { label:'動支', type:'rate',   cells: warehouses.map(w => rateCell(w[d.key.p])).join('') },
    ];
    return lines.map((line, i) => `
      <tr class="matrix-${line.type} ${i === 0 ? 'work-start group-start' : ''}">
        ${i === 0 ? `<td rowspan="3" class="freight-ref-matrix-category">${d.icon} ${d.label}</td>` : ''}
        <td class="freight-ref-matrix-metric">${line.label}</td>
        ${line.cells}
      </tr>`).join('');
  }).join('');

  const matrixHeader = warehouses
    .map(w => `<th>${w.name}</th>`)
    .join('');

  return `
  <section class="t002-section">
    <div class="t002-section-title">區間動支彙總</div>
    <article class="freight-ref-matrix-card t002-matrix-card">
      <div class="freight-ref-matrix-scroll">
        <table class="freight-ref-matrix t002-matrix">
          <thead>
            <tr>
              <th>領域</th>
              <th>項目</th>
              ${matrixHeader}
            </tr>
          </thead>
          <tbody>${matrixRows}</tbody>
        </table>
      </div>
      <details class="table-note t002-note">
        <summary>說明</summary>
        <div class="t002-note-content">
          <div>📌 動支% = 實際 ÷ 期間預算 · 期間預算 = 每月預算依篩選日期比例加總</div>
          <div>📌 預算來源：${hasAnnualDispatchBudget() ? '年度預算（支援跨月加總）' : `${dispatchBudgetMonthLabel()}預算`}</div>
          <div>📌 三色門檻：&lt; 75% 🟢 安全 · 75–90% 🟡 注意 · &gt; 90% 🔴 危險</div>
        </div>
      </details>
    </article>
  </section>`;
}

// ════════════════════════════════════════════
// 人效監控共用工具
// ════════════════════════════════════════════

function productivityTotals(labor, picks) {
  const hrs      = labor.reduce((s, r) => s + r.hours,   0);
  const cost     = labor.reduce((s, r) => s + r.cost,    0);
  const normHrs  = labor.reduce((s, r) => s + r.normHrs, 0);
  const totalPicks = picks.reduce((s, r) => s + r.picks, 0);
  const otHrs    = Math.max(0, hrs - normHrs);
  return {
    hrs, cost, normHrs, otHrs,
    picks: totalPicks,
    pph:   hrs        > 0 ? totalPicks / hrs  : 0,
    cpp:   totalPicks > 0 ? cost / totalPicks : 0,
    otPct: hrs        > 0 ? otHrs / hrs * 100 : 0,
  };
}

// M019 全區 PPH
function renderM019(t, code = 'M019', span = 's4') {
  return `
  <div class="w ${span} metric-card">
    <div class="gold-band">${code} · PPH</div>
    <div class="wh">
      <div class="wl"><div class="wdot"></div>每小時揀次 PPH</div>
      <span class="wmeta">全倉期間均值</span>
    </div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-blue);line-height:1">${t.pph.toFixed(1)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">次 / h</div>
    </div>
    <div class="kd">${t.picks.toLocaleString()} 揀次 ÷ ${t.hrs.toFixed(1)} h</div>
  </div>`;
}

// M020 單次揀貨成本
function renderM020(t, code = 'M020', span = 's4') {
  return `
  <div class="w ${span} metric-card">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">${code} · CPP</div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>單次揀貨成本</div>
      <span class="wmeta">元 / 揀次</span>
    </div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:var(--ry-ink);line-height:1">${t.cpp.toFixed(2)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">元 / 揀次</div>
    </div>
    <div class="kd">${fmtMoney(Math.round(t.cost))} ÷ ${t.picks.toLocaleString()} 次</div>
  </div>`;
}

// M021 加班佔比
function renderM021(t) {
  const pct   = t.otPct;
  const color = pct > 20 ? '#d9401b' : pct > 10 ? '#e07855' : '#1b7c33';
  const label = pct > 20 ? '🔴 偏高' : pct > 10 ? '🟡 注意' : '🟢 正常';
  return `
  <div class="w s4">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${color}"></div>M021 加班佔比</div>
      <span class="wmeta">${label}</span>
    </div>
    <div class="kv kv-lg" style="color:${color}">${pct.toFixed(1)}%</div>
    <div class="metric-track metric-track-lg">
      <div class="metric-fill" style="width:${Math.min(pct, 100)}%;background:${color}"></div>
    </div>
    <div class="kd">加班 ${t.otHrs.toFixed(1)} h ÷ 總工時 ${t.hrs.toFixed(1)} h</div>
  </div>`;
}

// M023 三倉 PPH 比較（返回 3 個 s4 widget）
function renderM023(labor, picks, code = 'M023') {
  const WHS    = ['大溪倉', '大肚倉', '岡山倉'];
  const colors = { '大溪倉': '#0E7BAD', '大肚倉': '#2DA870', '岡山倉': '#E07855' };
  const stats  = WHS.map(wh => ({
    wh,
    ...productivityTotals(labor.filter(r => r.wh === wh), picks.filter(r => r.wh === wh)),
  }));
  const maxPph = Math.max(...stats.map(s => s.pph), 1);

  return stats.map(s => {
    const c    = colors[s.wh];
    const barW = (s.pph / maxPph * 100).toFixed(1);
    const has  = s.hrs > 0 || s.picks > 0;
    return `
  <div class="w s4 metric-card">
    <div class="gold-band" style="background:${c};color:white">${code} · ${s.wh}</div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${c}"></div>${s.wh} PPH</div>
    </div>
    <div style="padding:12px 16px">
      <div style="font-size:1.8rem;font-weight:900;color:${c};line-height:1">${has ? s.pph.toFixed(1) : '—'}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">次 / h</div>
      ${has ? `<div class="metric-track metric-track-sm" style="margin-top:8px">
        <div class="metric-fill" style="width:${barW}%;background:${c}"></div>
      </div>` : ''}
    </div>
    <div class="kd">${has ? `${s.picks.toLocaleString()} 次 / ${s.hrs.toFixed(1)} h` : '尚未有資料'}</div>
  </div>`;
  }).join('');
}

// M025 工時區域人效矩陣
function renderM025(labor, picks, code = 'M025') {
  const laborByArea = {};
  labor.forEach(r => {
    if (!laborByArea[r.opArea]) laborByArea[r.opArea] = { hrs: 0, cost: 0, normHrs: 0 };
    laborByArea[r.opArea].hrs     += r.hours;
    laborByArea[r.opArea].cost    += r.cost;
    laborByArea[r.opArea].normHrs += r.normHrs;
  });

  const picksByArea = {};
  picks.forEach(r => { picksByArea[r.op] = (picksByArea[r.op] || 0) + r.picks; });

  const allAreas = new Set([...Object.keys(laborByArea), ...Object.keys(picksByArea)]);
  const rows = [...allAreas].map(area => {
    const l   = laborByArea[area] || { hrs: 0, cost: 0, normHrs: 0 };
    const p   = picksByArea[area] || 0;
    const otHrs = Math.max(0, l.hrs - l.normHrs);
    return {
      area,
      hrs:   l.hrs,
      cost:  l.cost,
      picks: p,
      pph:   l.hrs > 0 ? p / l.hrs        : 0,
      cpp:   p     > 0 ? l.cost / p        : 0,
      otPct: l.hrs > 0 ? otHrs / l.hrs * 100 : 0,
    };
  }).sort((a, b) => b.picks - a.picks || b.hrs - a.hrs);

  if (!rows.length) {
    return `
  <div class="w s12 table-card">
    <div class="wh"><div class="wl"><div class="wdot"></div>${code} 工時區域人效矩陣</div></div>
    <div class="empty-state">此日期區間沒有資料</div>
  </div>`;
  }

  const maxPph = Math.max(...rows.map(r => r.pph), 1);
  const trs = rows.map(r => {
    const barW    = r.pph > 0 ? (r.pph / maxPph * 100).toFixed(1) : '0';
    const otColor = r.otPct > 20 ? '#d9401b' : r.otPct > 10 ? '#e07855' : '#5a6478';
    return `<tr>
      <td>${r.area}</td>
      <td class="mono num-right">${r.picks.toLocaleString()}</td>
      <td class="mono num-right">${r.hrs.toFixed(1)}</td>
      <td style="min-width:120px">
        <div style="display:flex;align-items:center;gap:6px">
          <div style="flex:1;height:8px;background:var(--ry-line);border-radius:2px">
            <div style="height:8px;border-radius:2px;background:var(--ry-blue);width:${barW}%"></div>
          </div>
          <span class="mono" style="min-width:36px;text-align:right">${r.pph > 0 ? r.pph.toFixed(1) : '—'}</span>
        </div>
      </td>
      <td class="mono num-right">${r.cpp > 0 ? r.cpp.toFixed(2) : '—'}</td>
      <td class="mono num-right" style="color:${otColor}">${r.otPct > 0 ? r.otPct.toFixed(1) + '%' : '—'}</td>
    </tr>`;
  }).join('');

  return `
  <div class="w s12 table-card">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>${code} 工時區域人效矩陣</div>
      <span class="wmeta">${rows.length} 個作業區域</span>
    </div>
    <div class="ops-table-frame">
      <table class="tbl ops-compact-table">
        <thead><tr>
          <th>作業區域</th>
          <th class="num-right">揀次</th>
          <th class="num-right">工時(h)</th>
          <th>PPH</th>
          <th class="num-right">元/揀次</th>
          <th class="num-right">加班%</th>
        </tr></thead>
        <tbody>${trs}</tbody>
      </table>
      <div class="table-note">
        📌 PPH = 揀次 ÷ 工時 · 元/揀次 = 人力費用 ÷ 揀次 · 加班% = 加班工時 ÷ 總工時<br>
        📌 加班工時 = 總工時 − 正常工時（需 Excel 含「正常時數」欄位）<br>
        📌 揀次資料「工時區域」欄位需與人力資料「作業區域」欄位名稱對應
      </div>
    </div>
  </div>`;
}

// ════════════════════════════════════════════
// 月度結算共用工具
// ════════════════════════════════════════════

// ════════════════════════════════════════════
// 年度規劃分析
// ════════════════════════════════════════════

function getAnnualLaborActual() {
  const result = { '大溪倉': Array(12).fill(0), '大肚倉': Array(12).fill(0), '岡山倉': Array(12).fill(0) };
  const raw = typeof LABOR_RAW !== 'undefined' ? LABOR_RAW : [];
  raw.filter(r => r.hours > 0 && r.opArea !== '午休時間').forEach(r => {
    const mi = Number(r.date.slice(5, 7)) - 1;
    if (mi >= 0 && mi < 12 && result[r.wh]) result[r.wh][mi] += r.cost;
  });
  return result;
}

function getAnnualFreightActual() {
  const result = { '大溪倉': Array(12).fill(0), '大肚倉': Array(12).fill(0), '岡山倉': Array(12).fill(0) };
  DATA.freight.dailyByWarehouse.forEach(r => {
    const fullDate = r[4] || shortToFreightFullDate(r[0]);
    if (!fullDate) return;
    const mi = Number(fullDate.slice(5, 7)) - 1;
    if (mi < 0 || mi >= 12) return;
    result['大溪倉'][mi] += r[1] || 0;
    result['大肚倉'][mi] += r[2] || 0;
    result['岡山倉'][mi] += r[3] || 0;
  });
  return result;
}

function renderAnnualSection(type) {
  const label  = type === 'labor' ? '人力費用' : '運務費用';
  const dotCls = type === 'labor' ? ''         : 'dot-freight';
  const budget = DATA.annualBudget[type];
  const actual = type === 'labor' ? getAnnualLaborActual() : getAnnualFreightActual();
  const WHS    = ['大溪倉', '大肚倉', '岡山倉'];

  const totalBudget = Array(12).fill(0).map((_, mi) => WHS.reduce((s, wh) => s + (budget[wh]?.[mi] || 0), 0));
  const totalActual = Array(12).fill(0).map((_, mi) => WHS.reduce((s, wh) => s + (actual[wh]?.[mi] || 0), 0));
  const hasBudget   = totalBudget.some(v => v > 0);
  const hasActual   = totalActual.some(v => v > 0);

  if (!hasBudget && !hasActual) {
    return `
  <div class="w s12 table-card">
    <div class="gold-band">${label}區</div>
    <div class="wh"><div class="wl"><div class="wdot ${dotCls}"></div>${label}區</div></div>
    <div class="empty-state">請先匯入${type === 'labor' ? '工時資料與' : '運費資料與'}年度預算</div>
  </div>`;
  }

  return [
    renderAnnualCharts(label, dotCls, totalBudget, totalActual, hasBudget, hasActual),
    renderAnnualTable(label, dotCls, budget, actual, WHS, totalBudget, totalActual),
  ].join('');
}

function renderAnnualCharts(label, dotCls, totalBudget, totalActual, hasBudget, hasActual) {
  const W = 600, H = 200, padL = 52, padR = 16, padT = 24, padB = 36;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const xFor  = mi => padL + mi * (plotW / 11);
  const xLabels = Array(12).fill(0).map((_, mi) =>
    `<text x="${xFor(mi).toFixed(1)}" y="${H - 8}" text-anchor="middle" font-size="10" fill="#5a6478" font-family="Noto Sans TC">${mi + 1}月</text>`
  ).join('');

  // Chart 1: 達成率折線
  const pcts = Array(12).fill(0).map((_, mi) =>
    totalBudget[mi] > 0 ? totalActual[mi] / totalBudget[mi] * 100 : null
  );
  const hasPct = pcts.some(v => v !== null);
  const maxPct = 120;
  const yPct   = v => padT + plotH - (v / maxPct * plotH);

  const pctPath = (() => {
    const segs = []; let penDown = false;
    pcts.forEach((p, mi) => {
      if (p === null) { penDown = false; return; }
      segs.push(`${penDown ? 'L' : 'M'}${xFor(mi).toFixed(1)},${yPct(p).toFixed(1)}`);
      penDown = true;
    });
    return segs.join(' ');
  })();

  const pctDots = pcts.map((p, mi) => {
    if (p === null) return '';
    const c = colorFor(p);
    return `<circle cx="${xFor(mi).toFixed(1)}" cy="${yPct(p).toFixed(1)}" r="4" fill="${c}" stroke="white" stroke-width="1.5"><title>${mi + 1}月：${p.toFixed(1)}%</title></circle>`;
  }).join('');

  const pctYTicks = [0, 25, 50, 75, 90, 100].map(v => {
    const y = yPct(v).toFixed(1);
    const isRef = v === 75 || v === 90;
    return `<line x1="${padL}" y1="${y}" x2="${padL + plotW}" y2="${y}" stroke="${v === 75 ? '#f59e0b' : v === 90 ? '#d9401b' : '#dde2ec'}" stroke-width="${isRef ? 1 : 0.5}" stroke-dasharray="${isRef ? '4,3' : ''}"/>
    <text x="${padL - 4}" y="${(Number(y) + 3).toFixed(1)}" text-anchor="end" font-size="9" fill="#5a6478">${v}%</text>`;
  }).join('');

  const chart1 = `
  <div class="w s6 chart-card">
    <div class="wh"><div class="wl"><div class="wdot ${dotCls}"></div>達成率趨勢（全區合計）</div><span class="wmeta">${label}</span></div>
    ${!hasPct ? '<div class="empty-state">需同時匯入預算與費用資料</div>' : `
    <svg class="chart-svg" viewBox="0 0 ${W} ${H}">
      ${pctYTicks}
      ${pctPath ? `<path d="${pctPath}" fill="none" stroke="var(--ry-blue)" stroke-width="2" stroke-linejoin="round"/>` : ''}
      ${pctDots}
      ${xLabels}
    </svg>
    <div class="chart-note">📌 黃虛線 75% · 紅虛線 90% · 懸停查看月份數值</div>`}
  </div>`;

  // Chart 2: 預算 vs 實際金額折線
  const maxAmt = Math.max(...totalBudget, ...totalActual, 1);
  const yAmt   = v => padT + plotH - (v / maxAmt * plotH);

  const buildAmtPath = vals => {
    const segs = []; let penDown = false;
    vals.forEach((v, mi) => {
      if (v === 0 && totalBudget[mi] === 0 && totalActual[mi] === 0) { penDown = false; return; }
      segs.push(`${penDown ? 'L' : 'M'}${xFor(mi).toFixed(1)},${yAmt(v).toFixed(1)}`);
      penDown = true;
    });
    return segs.join(' ');
  };

  const amtYTicks = [0, maxAmt / 2, maxAmt].map(v =>
    `<line x1="${padL}" y1="${yAmt(v).toFixed(1)}" x2="${padL + plotW}" y2="${yAmt(v).toFixed(1)}" stroke="#dde2ec" stroke-width="0.5" stroke-dasharray="3,3"/>
     <text x="${padL - 4}" y="${(yAmt(v) + 3).toFixed(1)}" text-anchor="end" font-size="9" fill="#5a6478">${(v / 1000).toFixed(0)}K</text>`
  ).join('');

  const chart2 = `
  <div class="w s6 chart-card">
    <div class="wh"><div class="wl"><div class="wdot ${dotCls}"></div>預算 vs 實際（全區合計）</div><span class="wmeta">${label}</span></div>
    <svg class="chart-svg" viewBox="0 0 ${W} ${H}">
      ${amtYTicks}
      <line x1="${padL + 4}" y1="${padT - 12}" x2="${padL + 20}" y2="${padT - 12}" stroke="#1e5ca8" stroke-width="2" stroke-dasharray="5,3"/>
      <text x="${padL + 24}" y="${padT - 8}" font-size="10" fill="#1e5ca8">預算</text>
      <line x1="${padL + 60}" y1="${padT - 12}" x2="${padL + 76}" y2="${padT - 12}" stroke="#2ea85a" stroke-width="2"/>
      <text x="${padL + 80}" y="${padT - 8}" font-size="10" fill="#2ea85a">實際</text>
      ${hasBudget ? `<path d="${buildAmtPath(totalBudget)}" fill="none" stroke="#1e5ca8" stroke-width="2" stroke-dasharray="6,4" stroke-linejoin="round"/>` : ''}
      ${hasActual ? `<path d="${buildAmtPath(totalActual)}" fill="none" stroke="#2ea85a" stroke-width="2" stroke-linejoin="round"/>` : ''}
      ${xLabels}
    </svg>
    <div class="chart-note">📌 藍虛線 = 預算 · 綠實線 = 實際 · Y 軸單位：千元</div>
  </div>`;

  return chart1 + chart2;
}

function renderAnnualTable(label, dotCls, budget, actual, WHS, totalBudget, totalActual) {
  const grandBudget = totalBudget.reduce((s, v) => s + v, 0);
  const grandActual = totalActual.reduce((s, v) => s + v, 0);

  const pctBadge = (b, a) => {
    if (!b && !a) return '<td class="annual-empty" colspan="3">—</td>';
    const pct   = b > 0 ? a / b * 100 : null;
    const color = pct !== null ? colorFor(pct) : 'var(--ry-muted)';
    const badge = pct !== null
      ? `<span class="annual-pct-badge" style="background:${color}">${pct.toFixed(1)}%</span>`
      : '—';
    return `<td class="mono num-right">${b > 0 ? fmtMoney(Math.round(b)) : '—'}</td>` +
           `<td class="mono num-right">${a > 0 ? fmtMoney(Math.round(a)) : '—'}</td>` +
           `<td class="num-right">${badge}</td>`;
  };

  const thMonths = Array(12).fill(0).map((_, mi) =>
    `<th colspan="3" class="annual-th-month">${mi + 1}月</th>`
  ).join('');
  const thSubs = Array(12).fill(0).map(() =>
    '<th class="num-right annual-th-sub">預算</th><th class="num-right annual-th-sub">實際</th><th class="num-right annual-th-sub">達成率</th>'
  ).join('');

  const whRows = WHS.map(wh => {
    const yearB = budget[wh].reduce((s, v) => s + v, 0);
    const yearA = actual[wh].reduce((s, v) => s + v, 0);
    const cells = Array(12).fill(0).map((_, mi) => pctBadge(budget[wh][mi], actual[wh][mi])).join('');
    return `<tr>
      <td class="annual-td-wh">${wh}</td>
      ${cells}
      ${pctBadge(yearB, yearA)}
    </tr>`;
  }).join('');

  const totalCells = Array(12).fill(0).map((_, mi) => pctBadge(totalBudget[mi], totalActual[mi])).join('');

  return `
  <div class="w s12 table-card">
    <div class="wh">
      <div class="wl"><div class="wdot ${dotCls}"></div>月度明細表</div>
      <span class="wmeta">全年合計：預算 ${fmtMoney(Math.round(grandBudget))} · 實際 ${fmtMoney(Math.round(grandActual))}</span>
    </div>
    <div class="table-edge annual-table-scroll">
      <table class="tbl annual-tbl">
        <thead>
          <tr>
            <th rowspan="2" class="annual-td-wh">倉別</th>
            ${thMonths}
            <th colspan="3" class="annual-th-month annual-th-year">全年合計</th>
          </tr>
          <tr>
            ${thSubs}
            <th class="num-right annual-th-sub">預算</th>
            <th class="num-right annual-th-sub">實際</th>
            <th class="num-right annual-th-sub">達成率</th>
          </tr>
        </thead>
        <tbody>
          ${whRows}
          <tr class="annual-row-total">
            <td class="annual-td-wh">三倉合計</td>
            ${totalCells}
            ${pctBadge(grandBudget, grandActual)}
          </tr>
        </tbody>
      </table>
      <div class="table-note">
        📌 達成率 = 實際 ÷ 預算 · 顏色：🟢 &lt; 75% · 🟡 75–90% · 🔴 &gt; 90%<br>
        📌 人力費用來源：工時 Excel · 運務費用來源：運務 Excel · 預算來源：年度預算V2
      </div>
    </div>
  </div>`;
}

function getMonthlyBudgetByWh() {
  const idx = monthIndexFromDate(DATA.dateFrom);
  return ['大溪倉', '大肚倉', '岡山倉'].reduce((acc, wh) => {
    acc[wh] = {
      labor:   DATA.annualBudget?.labor?.[wh]?.[idx]   || 0,
      freight: DATA.annualBudget?.freight?.[wh]?.[idx] || 0,
    };
    return acc;
  }, {});
}

function getDeptTypeMap() {
  const map = {};
  (DATA.org?.depts || []).forEach(d => { if (d.name) map[d.name] = d.type; });
  return map;
}

function sumLaborByWh(labor) {
  const r = { '大溪倉': { hrs: 0, cost: 0 }, '大肚倉': { hrs: 0, cost: 0 }, '岡山倉': { hrs: 0, cost: 0 } };
  labor.forEach(x => { if (r[x.wh]) { r[x.wh].hrs += x.hours; r[x.wh].cost += x.cost; } });
  return r;
}

function sumFreightByWh() {
  const r = { '大溪倉': 0, '大肚倉': 0, '岡山倉': 0 };
  getFreightDailyRowsFiltered().forEach(row => {
    r['大溪倉'] += row[1];
    r['大肚倉'] += row[2];
    r['岡山倉'] += row[3];
  });
  return r;
}

// M001 總人力成本
function renderM001(labor) {
  const cost = labor.reduce((s, r) => s + r.cost,  0);
  const hrs  = labor.reduce((s, r) => s + r.hours, 0);
  const emp  = new Set(labor.map(r => r.empId)).size;
  return `
  <div class="w s3">
    <div class="gold-band">M001 · 人力成本</div>
    <div class="wh">
      <div class="wl"><div class="wdot"></div>總人力成本</div>
      <span class="wmeta">${emp.toLocaleString()} 人 · ${hrs.toFixed(1)} h</span>
    </div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-blue);line-height:1">${fmtMoney(Math.round(cost))}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">人力費用合計</div>
    </div>
  </div>`;
}

// M002 總工時
function renderM002(labor) {
  const hrs   = labor.reduce((s, r) => s + r.hours,   0);
  const norm  = labor.reduce((s, r) => s + r.normHrs, 0);
  const ot    = Math.max(0, hrs - norm);
  const otPct = hrs > 0 ? ot / hrs * 100 : 0;
  const otColor = otPct > 20 ? '#d9401b' : otPct > 10 ? '#e07855' : '#5a6478';
  return `
  <div class="w s3">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">M002 · 總工時</div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:var(--ry-gold)"></div>總工時</div>
      <span class="wmeta" style="color:${otColor}">加班 ${otPct.toFixed(1)}%</span>
    </div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">${hrs.toFixed(1)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">正常 ${norm.toFixed(1)} h + 加班 ${ot.toFixed(1)} h</div>
    </div>
  </div>`;
}

// F001 月總運費（月度結算版 s3 卡片）
function renderMFreight() {
  const rows    = getFreightTrendFiltered();
  const total   = rows.reduce((s, d) => s + d[1], 0);
  const summary = getFreightFilteredSummary();
  return `
  <div class="w s3">
    <div class="gold-band" style="background:#E07855;color:white">F001 · 月總運費</div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:#E07855"></div>月總運費</div>
      <span class="wmeta">${summary.totalOrders.toLocaleString()} 筆配送</span>
    </div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:#E07855;line-height:1">${fmtMoney(Math.round(total))}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">${rows.length} 天結算</div>
    </div>
  </div>`;
}

// 第 4 KPI：優先 PPH > 全區動支率 > 總費用
function renderMSummaryKpi(labor, picks, freightTotal, budget) {
  const laborCost   = labor.reduce((s, r) => s + r.cost, 0);
  const totalActual = laborCost + freightTotal;
  const totBudget   = Object.values(budget).reduce((s, b) => s + b.labor + b.freight, 0);

  if (labor.length > 0 && picks.length > 0) {
    const t = productivityTotals(labor, picks);
    return `
  <div class="w s3">
    <div class="gold-band" style="background:#2ea85a;color:white">M019 · PPH</div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:#2ea85a"></div>每小時揀次</div>
      <span class="wmeta">${t.picks.toLocaleString()} 揀次</span>
    </div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:#2ea85a;line-height:1">${t.pph.toFixed(1)}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">次 / h</div>
    </div>
    <div class="kd">${t.picks.toLocaleString()} 次 ÷ ${t.hrs.toFixed(1)} h</div>
  </div>`;
  }

  if (totBudget > 0) {
    const pct   = totalActual / totBudget * 100;
    const color = colorFor(pct);
    return `
  <div class="w s3">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${color}"></div>全區動支率</div>
      <span class="wmeta">${labelFor(pct)}</span>
    </div>
    <div class="kv kv-lg" style="color:${color}">${pct.toFixed(1)}%</div>
    <div class="metric-track metric-track-lg">
      <div class="metric-fill" style="width:${Math.min(pct, 100)}%;background:${color}"></div>
    </div>
    <div class="kd">${fmtMoney(Math.round(totalActual))} ÷ ${fmtMoney(Math.round(totBudget))}</div>
  </div>`;
  }

  return `
  <div class="w s3">
    <div class="gold-band" style="background:#5a6478;color:white">總費用</div>
    <div class="wh"><div class="wl"><div class="wdot" style="background:#5a6478"></div>人力 + 運費合計</div></div>
    <div style="padding:16px;text-align:center">
      <div style="font-size:1.8rem;font-weight:900;color:var(--ry-ink);line-height:1">${fmtMoney(Math.round(totalActual))}</div>
      <div style="font-size:var(--fs-xs);color:var(--ry-muted);margin-top:4px">合計費用</div>
    </div>
  </div>`;
}

// M011 本部費用彙總（三倉 × 人力/運費 × 預算/實際）
function renderM011(laborByWh, freightByWh, budget) {
  const WHS      = ['大溪倉', '大肚倉', '岡山倉'];
  const whColors = { '大溪倉': 'var(--tbl-daxi)', '大肚倉': 'var(--tbl-dadu)', '岡山倉': 'var(--tbl-gangshan)' };
  const hasBudget = Object.values(budget).some(b => b.labor > 0 || b.freight > 0);

  function pctBadge(pct) {
    if (!hasBudget) return '';
    return `<td class="num-right"><span class="solid-pct-badge" style="background:${colorFor(pct)}">${pct.toFixed(1)}%</span></td>`;
  }

  const dataRows = WHS.map(wh => {
    const l   = laborByWh[wh]   || { cost: 0 };
    const f   = freightByWh[wh] || 0;
    const b   = budget[wh]      || { labor: 0, freight: 0 };
    const tot = l.cost + f;
    const totB = b.labor + b.freight;
    const lPct = b.labor   ? l.cost / b.labor   * 100 : 0;
    const fPct = b.freight ? f      / b.freight * 100 : 0;
    const tPct = totB      ? tot    / totB       * 100 : 0;
    const lColor = hasBudget ? colorFor(lPct) : 'var(--ry-ink)';
    const fColor = hasBudget ? colorFor(fPct) : 'var(--ry-ink)';
    const tColor = hasBudget ? colorFor(tPct) : 'var(--ry-ink)';
    return `<tr>
      <td style="font-weight:700;background:${whColors[wh]};color:white;padding:4px 10px">${wh}</td>
      ${hasBudget ? `<td class="mono num-right">${fmtMoney(b.labor)}</td>` : ''}
      <td class="mono num-right actual-strong" style="color:${lColor}">${fmtMoney(Math.round(l.cost))}</td>
      ${pctBadge(lPct)}
      ${hasBudget ? `<td class="mono num-right">${fmtMoney(b.freight)}</td>` : ''}
      <td class="mono num-right actual-strong" style="color:${fColor}">${fmtMoney(Math.round(f))}</td>
      ${pctBadge(fPct)}
      ${hasBudget ? `<td class="mono num-right">${fmtMoney(totB)}</td>` : ''}
      <td class="mono num-right actual-strong" style="color:${tColor}">${fmtMoney(Math.round(tot))}</td>
      ${pctBadge(tPct)}
    </tr>`;
  });

  const tL  = WHS.reduce((s, wh) => s + (laborByWh[wh]?.cost || 0), 0);
  const tF  = WHS.reduce((s, wh) => s + (freightByWh[wh] || 0), 0);
  const tBL = WHS.reduce((s, wh) => s + (budget[wh]?.labor || 0), 0);
  const tBF = WHS.reduce((s, wh) => s + (budget[wh]?.freight || 0), 0);
  const tB  = tBL + tBF;
  const tA  = tL + tF;
  const tLP = tBL ? tL / tBL * 100 : 0;
  const tFP = tBF ? tF / tBF * 100 : 0;
  const tTP = tB  ? tA / tB  * 100 : 0;
  const tlC = hasBudget ? colorFor(tLP) : 'var(--ry-ink)';
  const tfC = hasBudget ? colorFor(tFP) : 'var(--ry-ink)';
  const ttC = hasBudget ? colorFor(tTP) : 'var(--ry-ink)';

  const totalRow = `<tr style="font-weight:800;border-top:2px solid var(--ry-line)">
    <td>全區合計</td>
    ${hasBudget ? `<td class="mono num-right">${fmtMoney(tBL)}</td>` : ''}
    <td class="mono num-right actual-heavy" style="color:${tlC}">${fmtMoney(Math.round(tL))}</td>
    ${pctBadge(tLP)}
    ${hasBudget ? `<td class="mono num-right">${fmtMoney(tBF)}</td>` : ''}
    <td class="mono num-right actual-heavy" style="color:${tfC}">${fmtMoney(Math.round(tF))}</td>
    ${pctBadge(tFP)}
    ${hasBudget ? `<td class="mono num-right">${fmtMoney(tB)}</td>` : ''}
    <td class="mono num-right actual-heavy" style="color:${ttC}">${fmtMoney(Math.round(tA))}</td>
    ${pctBadge(tTP)}
  </tr>`;

  const lHead = hasBudget
    ? `<th class="num-right">人力預算</th><th class="num-right">人力實際</th><th class="num-right">動支</th>`
    : `<th class="num-right">人力實際</th>`;
  const fHead = hasBudget
    ? `<th class="num-right">運費預算</th><th class="num-right">運費實際</th><th class="num-right">動支</th>`
    : `<th class="num-right">運費實際</th>`;
  const tHead = hasBudget
    ? `<th class="num-right">合計預算</th><th class="num-right">合計實際</th><th class="num-right">動支</th>`
    : `<th class="num-right">合計實際</th>`;
  const monthLabel = monthIndexFromDate(DATA.dateFrom) + 1;

  return `
  <div class="w s12 table-card">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>M011 本部費用彙總</div>
      <span class="wmeta">${hasBudget ? `對照 ${monthLabel} 月預算` : '⚠️ 尚未載入年度預算'}</span>
    </div>
    <div class="table-edge">
      <div class="scroll-x">
        <table class="tbl">
          <thead><tr><th>倉別</th>${lHead}${fHead}${tHead}</tr></thead>
          <tbody>${dataRows.join('')}${totalRow}</tbody>
        </table>
      </div>
      ${hasBudget ? `<div class="table-note">📌 三色門檻：&lt; 75% 🟢 安全 · 75–90% 🟡 注意 · &gt; 90% 🔴 危險 · 預算來源：年度預算 ${monthLabel} 月</div>` : ''}
    </div>
  </div>`;
}

// M013/M014 EC 課別費用分析（依 org 設定分類）
function renderECBreakdown(labor) {
  const typeMap = getDeptTypeMap();
  const EC_DEFS = [
    { type: '服務EC',   id: 'M013', color: '#1e5ca8' },
    { type: '營收EC',   id: 'M014', color: '#e07855' },
    { type: '後勤支援', id: 'M-SUP', color: '#2ea85a' },
  ];

  const byType = {};
  EC_DEFS.forEach(d => { byType[d.type] = {}; });

  labor.forEach(r => {
    const displayName = deptDisplayName(r.dept);
    if (displayName === null) return;
    const type = typeMap[r.dept] || '未分類';
    if (!byType[type]) byType[type] = {};
    if (!byType[type][displayName]) byType[type][displayName] = { hrs: 0, cost: 0 };
    byType[type][displayName].hrs  += r.hours;
    byType[type][displayName].cost += r.cost;
  });

  const active = EC_DEFS.filter(d => Object.keys(byType[d.type] || {}).length > 0);
  if (!active.length) return '';

  const colClass = active.length >= 3 ? 's4' : 's6';

  return active.map(({ type, id, color }) => {
    const depts   = Object.entries(byType[type]).sort((a, b) => b[1].cost - a[1].cost);
    const totCost = depts.reduce((s, [, v]) => s + v.cost, 0);
    const totHrs  = depts.reduce((s, [, v]) => s + v.hrs, 0);

    const rows = depts.map(([dept, v]) => {
      const pct  = totCost > 0 ? v.cost / totCost * 100 : 0;
      const rate = v.hrs > 0 ? Math.round(v.cost / v.hrs) : 0;
      return `<tr>
        <td>${dept}</td>
        <td class="mono num-right">${v.hrs.toFixed(1)}</td>
        <td class="mono num-right">${fmtMoney(Math.round(v.cost))}</td>
        <td class="mono num-right">$${rate.toLocaleString()}</td>
        <td style="min-width:80px">
          <div style="background:var(--ry-line);border-radius:2px;height:6px;margin-bottom:2px">
            <div style="background:${color};height:6px;border-radius:2px;width:${pct.toFixed(1)}%"></div>
          </div>
          <span style="font-size:9px;color:var(--ry-muted)">${pct.toFixed(1)}%</span>
        </td>
      </tr>`;
    }).join('');

    return `
  <div class="w ${colClass}">
    <div class="gold-band" style="background:${color};color:white">${id} · ${type}</div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${color}"></div>${type}課別費用</div>
      <span class="wmeta">${fmtMoney(Math.round(totCost))} · ${totHrs.toFixed(1)} h</span>
    </div>
    <table class="tbl">
      <thead><tr>
        <th>課別</th>
        <th class="num-right">工時(h)</th>
        <th class="num-right">費用</th>
        <th class="num-right">時薪</th>
        <th>佔比</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
  }).join('');
}

// 本月揀次效率摘要（有揀次資料時顯示）
function renderMEfficiency(labor, picks) {
  const WHS    = ['大溪倉', '大肚倉', '岡山倉'];
  const colors = { '大溪倉': '#0E7BAD', '大肚倉': '#2DA870', '岡山倉': '#E07855' };
  const all    = productivityTotals(labor, picks);

  const rows = WHS.map(wh => {
    const s = productivityTotals(labor.filter(r => r.wh === wh), picks.filter(r => r.wh === wh));
    const c = colors[wh];
    return `<tr>
      <td style="color:${c};font-weight:700">${wh}</td>
      <td class="mono num-right">${s.picks.toLocaleString()}</td>
      <td class="mono num-right">${s.hrs.toFixed(1)}</td>
      <td class="mono num-right">${s.pph > 0 ? s.pph.toFixed(1) : '—'}</td>
      <td class="mono num-right">${s.cpp > 0 ? s.cpp.toFixed(2) : '—'}</td>
      <td class="mono num-right">${s.otPct > 0 ? s.otPct.toFixed(1) + '%' : '—'}</td>
    </tr>`;
  }).join('');

  const totRow = `<tr style="font-weight:800;border-top:2px solid var(--ry-line)">
    <td>全區合計</td>
    <td class="mono num-right">${all.picks.toLocaleString()}</td>
    <td class="mono num-right">${all.hrs.toFixed(1)}</td>
    <td class="mono num-right">${all.pph > 0 ? all.pph.toFixed(1) : '—'}</td>
    <td class="mono num-right">${all.cpp > 0 ? all.cpp.toFixed(2) : '—'}</td>
    <td class="mono num-right">${all.otPct > 0 ? all.otPct.toFixed(1) + '%' : '—'}</td>
  </tr>`;

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>本月揀次效率摘要</div>
      <span class="wmeta">${all.picks.toLocaleString()} 揀次 · ${all.hrs.toFixed(1)} h</span>
    </div>
    <div class="table-edge">
      <table class="tbl">
        <thead><tr>
          <th>倉別</th>
          <th class="num-right">揀次</th>
          <th class="num-right">工時(h)</th>
          <th class="num-right">PPH</th>
          <th class="num-right">元/揀次</th>
          <th class="num-right">加班%</th>
        </tr></thead>
        <tbody>${rows}${totRow}</tbody>
      </table>
    </div>
    <div class="table-note">📌 PPH = 揀次 ÷ 工時 · 元/揀次 = 人力費用 ÷ 揀次 · 加班% = 加班工時 ÷ 總工時</div>
  </div>`;
}

// T003 每日動支明細表
function renderT003() {
  const rows = getDispatchDailyFiltered();
  function pctBadge(pct) {
    const c = colorFor(pct);
    return `<span class="t003-pct-badge" style="background:${c}">${pct.toFixed(1)}%</span>`;
  }

  function dateLabel(row) {
    const full = dispatchRowFullDate(row);
    const parts = full.split('-').map(Number);
    if (parts.length !== 3 || parts.some(n => !n)) return row[0];
    return `${parts[0]}/${parts[1]}/${parts[2]}`;
  }

  const whDefs = [
    { name:'全區', laborCol:null, freightCol:null, laborBudget:0, freightBudget:0 },
    { name:'大溪倉', laborCol:1, freightCol:2 },
    { name:'大肚倉', laborCol:3, freightCol:4 },
    { name:'岡山倉', laborCol:5, freightCol:6 },
  ];

  const domains = [
    { key:'labor', label:'人力', icon:'💰', color:'#0E7BAD', budgetOf:(w,row)=>dailyBudgetForTable(w, row).labor, actualOf:(row,w)=>w.name === '全區' ? row[1]+row[3]+row[5] : row[w.laborCol] },
    { key:'freight', label:'運務', icon:'🚚', color:'#E07855', budgetOf:(w,row)=>dailyBudgetForTable(w, row).freight, actualOf:(row,w)=>w.name === '全區' ? row[2]+row[4]+row[6] : row[w.freightCol] },
    { key:'total', label:'合計', icon:'📊', color:'var(--ry-blue-dark)', budgetOf:(w,row)=> {
      const b = dailyBudgetForTable(w, row);
      return b.labor + b.freight;
    }, actualOf:(row,w)=> {
      if (w.name === '全區') return row[1]+row[2]+row[3]+row[4]+row[5]+row[6];
      return row[w.laborCol] + row[w.freightCol];
    }},
  ];

  function dailyBudgetForTable(w, row) {
    const date = dispatchRowFullDate(row);
    if (w.name !== '全區') return dailyDispatchBudget(w.name, date);
    return ['大溪倉', '大肚倉', '岡山倉'].reduce((sum, wh) => {
      const b = dailyDispatchBudget(wh, date);
      sum.labor += b.labor;
      sum.freight += b.freight;
      return sum;
    }, { labor: 0, freight: 0 });
  }

  const numCell = (v, latest) => `<td class="freight-ref-matrix-number${latest ? ' t003-latest' : ''}">${fmtMoney(Math.round(v || 0))}</td>`;
  const rateMatrixCell = (pct, latest) => {
    const state = pct > 90 ? 'danger' : pct >= 75 ? 'warning' : 'safe';
    return `<td class="freight-ref-matrix-number rate ${state}${latest ? ' t003-latest' : ''}">${(pct || 0).toFixed(1)}%</td>`;
  };

  if (!rows.length) {
    return `
  <section class="t002-section">
    <div class="t002-section-title">每日動支明細</div>
    <article class="freight-ref-matrix-card t003-matrix-card">
      <div class="t003-empty">此日期區間沒有總費用資料</div>
    </article>
  </section>`;
  }

  const dateHeaders = rows.map(row => {
    const isLatest = dispatchRowFullDate(row) === getDispatchLatestUploadDate();
    return `<th class="${isLatest ? 't003-latest-head' : ''}">${dateLabel(row)}</th>`;
  }).join('');

  const matrixRows = domains.map(domain => {
    const domainRowspan = whDefs.length * 3;
    return whDefs.map((w, whIndex) => {
      return ['預算', '實際', '動支'].map((item, itemIndex) => {
        const isFirstDomainRow = whIndex === 0 && itemIndex === 0;
        const type = item === '預算' ? 'budget' : item === '實際' ? 'actual' : 'rate';
        const cells = rows.map(row => {
          const latest = dispatchRowFullDate(row) === getDispatchLatestUploadDate();
          const budget = domain.budgetOf(w, row);
          const actual = domain.actualOf(row, w);
          if (item === '預算') return numCell(budget, latest);
          if (item === '實際') return numCell(actual, latest);
          return rateMatrixCell(budget ? actual / budget * 100 : 0, latest);
        }).join('');
        return `<tr class="matrix-${type} ${itemIndex === 0 ? 'work-start' : ''} ${isFirstDomainRow ? 'group-start' : ''}" data-t003-domain="${domain.label}">
          ${isFirstDomainRow ? `<td rowspan="${domainRowspan}" class="freight-ref-matrix-category">${domain.icon} ${domain.label}</td>` : ''}
          ${itemIndex === 0 ? `<td rowspan="3" class="freight-ref-matrix-item">${w.name}</td>` : ''}
          <td class="freight-ref-matrix-metric">${item}</td>
          ${cells}
        </tr>`;
      }).join('');
    }).join('');
  }).join('');

  return `
  <section class="t002-section">
    <div class="freight-ref-matrix-heading">
      <div class="t002-section-title">每日動支明細</div>
      <label class="freight-ref-matrix-view">
        <span>查看</span>
        <select class="filter-input" id="t003-domain-filter" onchange="filterT003Domain(this.value)">
          <option value="all">全部</option>
          <option value="人力">人力</option>
          <option value="運務">運務</option>
        </select>
      </label>
    </div>
    <article class="freight-ref-matrix-card t003-matrix-card">
      <div class="freight-ref-matrix-scroll">
        <table class="freight-ref-matrix">
          <thead>
            <tr>
              <th>領域</th>
              <th>倉別</th>
              <th>項目</th>
              ${dateHeaders}
            </tr>
          </thead>
          <tbody>${matrixRows}</tbody>
        </table>
      </div>
      <details class="table-note t002-note">
        <summary>說明</summary>
        <div class="t002-note-content">
          <div>📌 日期欄依篩選區間橫向展開 · 淺藍欄為已匯入資料最新日期 ${getDispatchLatestUploadDate() || '尚未匯入'}</div>
          <div>📌 動支率 = 當日實際 ÷ 單日預算 · 單日預算 = 月預算 ÷ ${daysInDispatchMonth()} 天</div>
        </div>
      </details>
    </article>
  </section>`;
}
