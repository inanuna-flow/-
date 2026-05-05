// ═══════════════════════════════════════════════════════
// widgets.js · 儀表板 Widget 渲染函式
// ═══════════════════════════════════════════════════════
// 元件樣式原則：固定外觀放在 widget.css 的 CSS class；
// JS 只保留資料結構、條件 class，以及必要的動態 style。

// M012 預算達成率
function renderM012() {
  const pct       = getActualPct();
  const progress  = getMonthProgress();
  const projected = getProjectedPct();
  const curColor  = colorFor(pct);
  const projColor = colorFor(projected);

  return `
  <div class="w s12" style="border-top:3px solid ${projColor}">
    <div class="gold-band gold-band-dynamic" style="background:${projColor}">
      M012 預算達成率 · 基準 = 月預算 ${fmtMoney(DATA.budget)}（來源：預算 Excel）
    </div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${projColor}"></div>動支率監控</div>
      <div class="wmeta">${DATA.dayOfMonth}/${DATA.totalDays} 天 · 月進度 ${(progress*100).toFixed(1)}%</div>
    </div>
    <div class="m012-grid">
      <div class="m012-panel m012-panel-bordered">
        <div class="metric-label">今日動支率</div>
        <div class="metric-row">
          <div class="metric-value" style="color:${curColor}">${pct.toFixed(1)}</div>
          <div class="metric-unit" style="color:${curColor}">%</div>
          <div class="metric-status">${labelFor(pct)}</div>
        </div>
        <div class="metric-track metric-track-lg">
          <div class="metric-fill" style="width:${Math.min(pct,100)}%;background:${curColor}"></div>
          <div class="metric-limit metric-limit-90"></div>
        </div>
        <div class="metric-sub">
          ${fmtMoney(DATA.actual)} / ${fmtMoney(DATA.budget)}
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
      </div>
    </div>
    <div class="m012-note" style="background:${bgFor(projected)};border-left-color:${projColor}">
      ${projected>=90 ? `<b class="text-danger">⚠️ 月底預估將超過 90% 目標上限</b>，建議立即盤點未執行項目、延後可遞延支出。`
        : projected>=75 ? `<b class="text-warning">🟡 月底預估接近警戒線</b>，請留意未來兩週的費用節奏。`
                       : `<b class="text-safe">🟢 月底預估控管良好</b>，可維持目前節奏。`}
    </div>
  </div>`;
}

// M015 Business Units 表格
function renderM015() {
  const normalUnits = DATA.units.filter(u => u.type !== 'sup');
  const totalFee = normalUnits.reduce((s, u) => s + u.fee, 0);

  function statusTag(u) {
    if (u.type === 'sup') return { label:'SUP', className:'tag-sup' };
    const pct = totalFee ? u.fee / totalFee * 100 : 0;
    if (pct > DATA.thresholdPeak)   return { label:'PEAK',   className:'tag-peak' };
    if (pct >= DATA.thresholdStable) return { label:'STABLE', className:'tag-stable' };
    return { label:'IDLE', className:'tag-idle' };
  }

  const sorted = [...DATA.units].sort((a, b) => {
    if (a.type === 'sup') return 1;
    if (b.type === 'sup') return -1;
    return b.fee - a.fee;
  });

  const rows = sorted.map(u => {
    const pct = u.type === 'sup' ? null : (totalFee ? u.fee / totalFee * 100 : 0);
    const tag = statusTag(u);
    return `
    <tr>
      <td>${u.name}</td>
      <td class="mono">${fmtMoney(u.fee)}</td>
      <td class="mono">${u.hr}h</td>
      <td class="mono">${pct === null ? '—' : pct.toFixed(1) + '%'}</td>
      <td><span class="tag ${tag.className}">${tag.label}</span></td>
    </tr>`;
  }).join('');

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot dot-freight"></div>M015 Business Units 表格</div>
      <span class="wmeta">⚠️ 門檻值待確認</span>
    </div>
    <div class="table-edge">
      <table class="tbl">
        <thead><tr><th>作業課別</th><th>費用</th><th>工時</th><th>佔比</th><th>狀態</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="table-note">
        📌 門檻：&gt; ${DATA.thresholdPeak}% = <b class="text-warning">PEAK</b> ｜
        ${DATA.thresholdStable}–${DATA.thresholdPeak}% = <b class="text-safe">STABLE</b> ｜
        &lt; ${DATA.thresholdStable}% = <b class="text-muted-strong">IDLE</b><br>
        📌 佔比計算排除 SUP 後勤支援課
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
  return DATA.freight.dailyTrend.filter(row => freightDateInRange(row[2] || shortToFreightFullDate(row[0])));
}

function getFreightDetailsFiltered() {
  if (!DATA.freight.details) return null;
  return DATA.freight.details.filter(r => freightDateInRange(r.fullDate));
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
  <div class="w s4" style="border-top:3px solid ${color}">
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
  const data = getFreightTrendFiltered();
  if (!data.length) {
    return `
  <div class="w s12">
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
    ? `<text x="${padL + i * xStep}" y="${H - 10}" text-anchor="middle" font-size="10" fill="#5a6478" font-family="Courier New, monospace">${d[0]}</text>`
    : '').join('');

  const dots = data.map((d, i) =>
    `<circle cx="${padL + i * xStep}" cy="${yFor(d[1])}" r="3" fill="#1e5ca8" stroke="white" stroke-width="1.5"><title>${d[0]}：${fmtMoney(d[1])}</title></circle>`
  ).join('');

  const avgY = yFor(avgCost);

  const yTicks = [minCost, (minCost + maxCost) / 2, maxCost].map(v =>
    `<line x1="${padL}" y1="${yFor(v)}" x2="${padL + plotW}" y2="${yFor(v)}" stroke="#dde2ec" stroke-width="0.5" stroke-dasharray="3,3"/>
     <text x="${padL - 8}" y="${yFor(v) + 3}" text-anchor="end" font-size="9" fill="#5a6478" font-family="Courier New, monospace">${(v/1000).toFixed(0)}K</text>`
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
      <text x="${padL + plotW - 4}" y="${avgY - 4}" text-anchor="end" font-size="10" fill="#e07855" font-family="Courier New, monospace">平均 ${fmtMoney(Math.round(avgCost))}</text>
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
  const budget = DATA.freight.warehouseBudget;
  const daily  = getFreightDailyRowsFiltered();
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
  const rows = getDispatchDailyFiltered();
  return rows.reduce((sum, row) => {
    const daily = dailyDispatchBudget(warehouse, dispatchRowFullDate(row));
    sum.labor += daily.labor;
    sum.freight += daily.freight;
    return sum;
  }, { labor: 0, freight: 0 });
}

function selectedDispatchDayCount() {
  const from = new Date(`${DATA.dateFrom}T00:00:00`);
  const to = new Date(`${DATA.dateTo}T00:00:00`);
  if (Number.isNaN(from.getTime()) || Number.isNaN(to.getTime()) || to < from) return 0;
  return Math.round((to - from) / 86400000) + 1;
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

  function card(stats, isAll) {
    const color = colors[stats.name];
    const rateColor = colorFor(stats.pct);
    return `
    <div class="w s3">
      <div class="wh">
        <div class="wl"><div class="wdot" style="background:${color}"></div>${isAll ? '🌐 全區' : stats.name}</div>
        <span class="wmeta">${labelFor(stats.pct)}</span>
      </div>
      <div class="t001-rate-row">
        <div class="kv kv-xl" style="color:${rateColor}">${stats.pct.toFixed(1)}</div>
        <div class="t001-rate-unit" style="color:${rateColor}">%</div>
      </div>
      <div class="metric-track metric-track-sm">
        <div class="metric-fill" style="width:${Math.min(stats.pct, 100)}%;background:${rateColor}"></div>
        <div class="metric-limit metric-limit-90 subtle"></div>
      </div>
      <div class="t001-cost-grid">
        <div>
          <div class="mini-label">人力費用</div>
          <div class="mono mini-value">${fmtMoney(stats.labor)}</div>
        </div>
        <div>
          <div class="mini-label">運務費用</div>
          <div class="mono mini-value">${fmtMoney(stats.freight)}</div>
        </div>
      </div>
      <div class="t001-card-note">
        <div>本期總計 ----------</div>
        <div>總實際 ${fmtMoney(stats.total)}</div>
        <div>總預算 ${fmtMoney(Math.round(stats.budget))}</div>
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

  const whColor = {
    '大溪倉': 'var(--tbl-daxi)',
    '大肚倉': 'var(--tbl-dadu)',
    '岡山倉': 'var(--tbl-gangshan)',
    '全區':   'var(--tbl-all)',
  };

  const warehouseHeader = warehouses.map(w => {
    const isAll = w.name === '全區';
    return `<th class="t002-wh-head" style="background:${whColor[w.name]}">${isAll ? '🌐 ' : ''}${w.name}</th>`;
  }).join('');

  function domainRows(domainLabel, domainIcon, domainBg, keyMap) {
    const budgetCells = warehouses.map(w => {
      return `<td class="mono num-right ${w.name === '全區' ? 'total-cell' : ''}">${fmtMoney(Math.round(w[keyMap.b]))}</td>`;
    }).join('');

    const actualCells = warehouses.map(w => {
      const color = colorFor(w[keyMap.p]);
      return `<td class="mono num-right actual-strong ${w.name === '全區' ? 'total-cell' : ''}" style="color:${color}">${fmtMoney(w[keyMap.a])}</td>`;
    }).join('');

    const pctCells = warehouses.map(w => {
      const color = colorFor(w[keyMap.p]);
      return `<td class="num-right ${w.name === '全區' ? 'total-cell' : ''}">
        <span class="solid-pct-badge solid-pct-badge-wide" style="background:${color}">${w[keyMap.p].toFixed(1)}%</span>
      </td>`;
    }).join('');

    const domainCell = `<td rowspan="3" class="t002-domain-cell" style="background:${domainBg}">
      <div class="domain-icon">${domainIcon}</div>
      ${domainLabel}
    </td>`;

    return `
      <tr class="t002-domain-start">
        ${domainCell}
        <td class="t002-item-cell">預算</td>
        ${budgetCells}
      </tr>
      <tr>
        <td class="t002-item-cell">實際</td>
        ${actualCells}
      </tr>
      <tr>
        <td class="t002-item-cell t002-item-strong">動支</td>
        ${pctCells}
      </tr>`;
  }

  const keyLabor   = { b:'laborBudget',   a:'labor',   p:'laborPct'   };
  const keyFreight = { b:'freightBudget', a:'freight', p:'freightPct' };
  const keyTotal   = { b:'budget',        a:'total',   p:'pct'        };

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>T002 期間動支彙總</div>
      <span class="wmeta">${DATA.dateFrom} ~ ${DATA.dateTo} · 共 ${rows.length} 天</span>
    </div>
    <div class="table-edge">
      <div class="scroll-x">
        <table class="tbl t002-table">
          <thead>
            <tr class="t002-head-row">
              <th class="t002-head-domain">領域</th>
              <th class="t002-head-item">項目</th>
              ${warehouseHeader}
            </tr>
          </thead>
          <tbody>
            ${domainRows('人力', '💰', '#0E7BAD', keyLabor)}
            ${domainRows('運務', '🚚', '#E07855', keyFreight)}
            ${domainRows('合計', '📊', 'var(--ry-blue-dark)', keyTotal)}
          </tbody>
        </table>
      </div>
      <div class="table-note">
        📌 動支% = 實際 ÷ 期間預算 · 期間預算 = 每月預算依篩選日期比例加總<br>
        📌 預算來源：${hasAnnualDispatchBudget() ? '年度預算（支援跨月加總）' : `${dispatchBudgetMonthLabel()}預算`}<br>
        📌 三色門檻：&lt; 75% 🟢 安全 · 75–90% 🟡 注意 · &gt; 90% 🔴 危險
      </div>
    </div>
  </div>`;
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
function renderM019(t) {
  return `
  <div class="w s4">
    <div class="gold-band">M019 · PPH</div>
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
function renderM020(t) {
  return `
  <div class="w s4">
    <div class="gold-band" style="background:var(--ry-gold);color:var(--ry-blue-dark)">M020 · CPP</div>
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
  <div class="w s4" style="border-top:3px solid ${color}">
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
function renderM023(labor, picks) {
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
  <div class="w s4">
    <div class="gold-band" style="background:${c};color:white">M023 · ${s.wh}</div>
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
function renderM025(labor, picks) {
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
  <div class="w s12">
    <div class="wh"><div class="wl"><div class="wdot"></div>M025 工時區域人效矩陣</div></div>
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
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>M025 工時區域人效矩陣</div>
      <span class="wmeta">${rows.length} 個作業區域</span>
    </div>
    <div class="table-edge">
      <table class="tbl">
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
  <div class="w s3" style="border-top:3px solid ${color}">
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
  <div class="w s12">
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

  function valueFor(row, w, domain, item) {
    const budget = domain.budgetOf(w, row);
    const actual = domain.actualOf(row, w);
    if (item === '預算') return `<span class="mono">${fmtMoney(Math.round(budget))}</span>`;
    if (item === '實際') return `<span class="mono t003-actual">${fmtMoney(actual)}</span>`;
    const pct = budget ? actual / budget * 100 : 0;
    return pctBadge(pct);
  }

  function matrixRows() {
    return domains.map(domain => {
      const domainRowspan = whDefs.length * 3;
      return whDefs.map((w, whIndex) => {
        return ['預算', '實際', '動支'].map((item, itemIndex) => {
          const firstDomainCell = whIndex === 0 && itemIndex === 0
            ? `<td rowspan="${domainRowspan}" class="t003-domain-cell" style="background:${domain.color}">
                <div class="t003-domain-icon">${domain.icon}</div>${domain.label}
              </td>`
            : '';
          const whCell = itemIndex === 0
            ? `<td rowspan="3" class="t003-wh-cell ${w.name === '全區' ? 't003-wh-total' : ''}">${w.name}</td>`
            : '';
          const dateCells = rows.map(row => {
            const full = dispatchRowFullDate(row);
            const isLatest = full === getDispatchLatestUploadDate();
            return `<td class="t003-date-cell ${isLatest ? 't003-latest' : ''}">${valueFor(row, w, domain, item)}</td>`;
          }).join('');
          return `<tr>
            ${firstDomainCell}
            ${whCell}
            <td class="t003-item-cell">${item}</td>
            ${dateCells}
          </tr>`;
        }).join('');
      }).join('');
    }).join('');
  }

  if (!rows.length) {
    return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot t003-title-dot"></div>T003 每日動支明細</div>
      <span class="wmeta">0 天 · 資料最新日期 ${getDispatchLatestUploadDate() || '尚未匯入'}</span>
    </div>
    <div class="t003-empty">此日期區間沒有總費用資料</div>
  </div>`;
  }

  const dateHeaders = rows.map(row => {
    const full = dispatchRowFullDate(row);
    const isLatest = full === getDispatchLatestUploadDate();
    return `<th class="t003-date-head ${isLatest ? 't003-latest-head' : ''}">${dateLabel(row)}</th>`;
  }).join('');

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot t003-title-dot"></div>T003 每日動支明細</div>
      <span class="wmeta">${rows.length} 天 · 橫向日期 · 資料最新日期 ${getDispatchLatestUploadDate()}</span>
    </div>
    <div class="t003-shell">
      <div class="t003-scroll">
        <table class="tbl t003-table" style="min-width:${234 + rows.length * 136}px">
          <thead class="t003-thead">
            <tr class="t003-head-row">
              <th class="t003-head-sticky t003-head-domain">領域</th>
              <th class="t003-head-sticky t003-head-wh">倉別</th>
              <th class="t003-head-sticky t003-head-item">項目</th>
              ${dateHeaders}
            </tr>
          </thead>
          <tbody>${matrixRows()}</tbody>
        </table>
      </div>
      <div class="t003-note">
        📌 日期欄依篩選區間橫向展開 · 淺藍欄為目前已匯入資料最新日期 ${getDispatchLatestUploadDate() || '尚未匯入'}<br>
        📌 動支率 = 當日實際 ÷ 單日預算 · 單日預算 = 月預算 ÷ ${daysInDispatchMonth()} 天
      </div>
    </div>
  </div>`;
}
