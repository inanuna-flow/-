// ═══════════════════════════════════════════════════════
// widgets.js · 儀表板 Widget 渲染函式
// ═══════════════════════════════════════════════════════

// M012 預算達成率
function renderM012() {
  const pct       = getActualPct();
  const progress  = getMonthProgress();
  const projected = getProjectedPct();
  const curColor  = colorFor(pct);
  const projColor = colorFor(projected);

  return `
  <div class="w s12" style="border-top:3px solid ${projColor}">
    <div class="gold-band" style="background:${projColor};color:white">
      M012 預算達成率 · 基準 = 月預算 ${fmtMoney(DATA.budget)}（來源：預算 Excel）
    </div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${projColor}"></div>動支率監控</div>
      <div class="wmeta">${DATA.dayOfMonth}/${DATA.totalDays} 天 · 月進度 ${(progress*100).toFixed(1)}%</div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
      <div style="border-right:1px solid var(--ry-line);padding-right:20px">
        <div style="font-size:var(--fs-xs);font-weight:700;letter-spacing:.06em;color:var(--ry-muted);margin-bottom:6px">今日動支率</div>
        <div style="display:flex;align-items:baseline;gap:6px">
          <div style="font-size:40px;font-weight:900;font-family:var(--f-mono);color:${curColor};line-height:1">${pct.toFixed(1)}</div>
          <div style="font-size:18px;color:${curColor};font-weight:700">%</div>
          <div style="margin-left:auto;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono)">${labelFor(pct)}</div>
        </div>
        <div style="height:10px;background:var(--ry-bg);border-radius:99px;overflow:hidden;margin:10px 0 6px;position:relative">
          <div style="width:${Math.min(pct,100)}%;height:100%;background:${curColor};transition:width .4s"></div>
          <div style="position:absolute;left:90%;top:-2px;bottom:-2px;width:2px;background:var(--ry-ink);opacity:.6"></div>
        </div>
        <div style="font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono)">
          ${fmtMoney(DATA.actual)} / ${fmtMoney(DATA.budget)}
        </div>
      </div>
      <div>
        <div style="font-size:var(--fs-xs);font-weight:700;letter-spacing:.06em;color:var(--ry-muted);margin-bottom:6px">預估月底動支率</div>
        <div style="display:flex;align-items:baseline;gap:6px">
          <div style="font-size:40px;font-weight:900;font-family:var(--f-mono);color:${projColor};line-height:1">${projected.toFixed(1)}</div>
          <div style="font-size:18px;color:${projColor};font-weight:700">%</div>
          <div style="margin-left:auto;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono)">${labelFor(projected)}</div>
        </div>
        <div style="height:10px;background:var(--ry-bg);border-radius:99px;overflow:hidden;margin:10px 0 6px;position:relative">
          <div style="width:${Math.min(projected,120)*0.8}%;height:100%;background:${projColor};transition:width .4s"></div>
          <div style="position:absolute;left:72%;top:-2px;bottom:-2px;width:2px;background:var(--ry-ink);opacity:.6"></div>
        </div>
        <div style="font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono)">依目前速度線性外推</div>
      </div>
    </div>
    <div style="margin-top:14px;padding:10px 14px;background:${bgFor(projected)};border-left:3px solid ${projColor};border-radius:2px;font-size:var(--fs-sm);line-height:1.7">
      ${projected>=90 ? `<b style="color:#d9401b">⚠️ 月底預估將超過 90% 目標上限</b>，建議立即盤點未執行項目、延後可遞延支出。`
        : projected>=75 ? `<b style="color:#e07855">🟡 月底預估接近警戒線</b>，請留意未來兩週的費用節奏。`
                       : `<b style="color:#1b7c33">🟢 月底預估控管良好</b>，可維持目前節奏。`}
    </div>
  </div>`;
}

// M015 Business Units 表格
function renderM015() {
  const normalUnits = DATA.units.filter(u => u.type !== 'sup');
  const totalFee = normalUnits.reduce((s, u) => s + u.fee, 0);

  function statusTag(u) {
    if (u.type === 'sup') return { label:'SUP', bg:'#eff6ff', fg:'#1e5ca8', border:'#bfdbfe' };
    const pct = u.fee / totalFee * 100;
    if (pct > DATA.thresholdPeak)   return { label:'PEAK',   bg:'#fff3e0', fg:'#e65100', border:'#ffcc80' };
    if (pct >= DATA.thresholdStable) return { label:'STABLE', bg:'#e8f5e9', fg:'#1b7c33', border:'#a5d6a7' };
    return { label:'IDLE', bg:'#f3f4f6', fg:'#6b7280', border:'#d1d5db' };
  }

  const sorted = [...DATA.units].sort((a, b) => {
    if (a.type === 'sup') return 1;
    if (b.type === 'sup') return -1;
    return b.fee - a.fee;
  });

  const rows = sorted.map(u => {
    const pct = u.type === 'sup' ? null : (u.fee / totalFee * 100);
    const tag = statusTag(u);
    return `
    <tr>
      <td>${u.name}</td>
      <td class="mono">${fmtMoney(u.fee)}</td>
      <td class="mono">${u.hr}h</td>
      <td class="mono">${pct === null ? '—' : pct.toFixed(1) + '%'}</td>
      <td><span class="tag" style="background:${tag.bg};color:${tag.fg};border-color:${tag.border}">${tag.label}</span></td>
    </tr>`;
  }).join('');

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:#E07855"></div>M015 Business Units 表格</div>
      <span class="wmeta">⚠️ 門檻值待確認</span>
    </div>
    <div style="padding:0;margin:0 -18px -16px">
      <table class="tbl">
        <thead><tr><th>作業課別</th><th>費用</th><th>工時</th><th>佔比</th><th>狀態</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div style="padding:10px 18px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono);background:var(--ry-bg);border-top:1px solid var(--ry-line);line-height:1.7">
        📌 門檻：&gt; ${DATA.thresholdPeak}% = <b style="color:#e65100">PEAK</b> ｜
        ${DATA.thresholdStable}–${DATA.thresholdPeak}% = <b style="color:#1b7c33">STABLE</b> ｜
        &lt; ${DATA.thresholdStable}% = <b style="color:#6b7280">IDLE</b><br>
        📌 佔比計算排除 SUP 後勤支援課
      </div>
    </div>
  </div>`;
}

// F001 月總運費
function renderF001() {
  const f = DATA.freight;
  const trend = (f.totalCost - f.lastMonthCost) / f.lastMonthCost * 100;
  const up = trend >= 0;
  const trendColor = up ? '#d9401b' : '#1b7c33';
  return `
  <div class="w s4" style="border-top:3px solid var(--ry-blue)">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>F001 月總運費</div>
      <span class="wmeta">本期結算</span>
    </div>
    <div class="kv" style="font-size:28px;color:var(--ry-blue)">${fmtMoney(f.totalCost)}</div>
    <div class="kd">
      <span style="color:${trendColor};font-weight:700">${up ? '↑' : '↓'} ${Math.abs(trend).toFixed(1)}%</span>
      較上月（${fmtMoney(f.lastMonthCost)}）· ${f.totalOrders.toLocaleString()} 筆配送
    </div>
  </div>`;
}

// F002 預計 vs 實際差異率
function renderF002() {
  const f = DATA.freight;
  const diff = f.actualCost - f.estimatedCost;
  const diffPct = (diff / f.estimatedCost * 100);
  const color = diff > 0 ? '#d9401b' : '#1b7c33';
  const label = diff > 0 ? '🔴 超支' : '🟢 節省';
  return `
  <div class="w s4" style="border-top:3px solid ${color}">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:${color}"></div>F002 預計 vs 實際差異率</div>
      <span class="wmeta">${label}</span>
    </div>
    <div class="kv" style="font-size:28px;color:${color}">${diffPct > 0 ? '+' : ''}${diffPct.toFixed(2)}%</div>
    <div class="kd">
      實際 ${fmtMoney(f.actualCost)}<br>
      預計 ${fmtMoney(f.estimatedCost)}（差異 ${diff > 0 ? '+' : ''}${fmtMoney(diff)}）
    </div>
  </div>`;
}

// F003 超支/節省筆數
function renderF003() {
  const f = DATA.freight;
  const total = f.overCount + f.saveCount;
  const overPct = (f.overCount / total * 100).toFixed(1);
  return `
  <div class="w s4" style="border-top:3px solid var(--ry-red)">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:var(--ry-red)"></div>F003 超支／節省筆數</div>
      <span class="wmeta">差異 &gt; ${fmtMoney(f.diffThreshold)}</span>
    </div>
    <div style="display:flex;align-items:baseline;gap:16px;margin:4px 0">
      <div>
        <div style="font-size:9px;color:var(--ry-red);font-weight:700;letter-spacing:.08em">🔴 超支</div>
        <div style="font-family:var(--f-mono);font-size:26px;font-weight:800;color:var(--ry-red)">${f.overCount}</div>
      </div>
      <div style="font-size:14px;color:var(--ry-muted)">／</div>
      <div>
        <div style="font-size:9px;color:#1b7c33;font-weight:700;letter-spacing:.08em">🟢 節省</div>
        <div style="font-family:var(--f-mono);font-size:26px;font-weight:800;color:#1b7c33">${f.saveCount}</div>
      </div>
    </div>
    <div class="kd">共 ${total.toLocaleString()} 筆有顯著差異 · 超支佔 ${overPct}%</div>
  </div>`;
}

// F009 日費用趨勢（折線圖）
function renderF009() {
  const data = DATA.freight.dailyTrend;
  const maxCost = Math.max(...data.map(d => d[1]));
  const minCost = Math.min(...data.map(d => d[1]));
  const avgCost = data.reduce((s, d) => s + d[1], 0) / data.length;

  const W = 1000, H = 240, padL = 56, padR = 20, padT = 20, padB = 36;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const xStep = plotW / (data.length - 1);
  const range = maxCost - minCost;
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
    <svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block">
      ${yTicks}
      <line x1="${padL}" y1="${avgY}" x2="${padL + plotW}" y2="${avgY}" stroke="#e07855" stroke-width="1" stroke-dasharray="5,4"/>
      <text x="${padL + plotW - 4}" y="${avgY - 4}" text-anchor="end" font-size="10" fill="#e07855" font-family="Courier New, monospace">平均 ${fmtMoney(Math.round(avgCost))}</text>
      <path d="${areaD}" fill="#1e5ca8" fill-opacity="0.08"/>
      <path d="${pathD}" fill="none" stroke="#1e5ca8" stroke-width="2" stroke-linejoin="round"/>
      ${dots}
      ${xLabels}
    </svg>
    <div style="padding:8px 14px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono);border-top:1px solid var(--ry-line);background:var(--ry-bg);margin:12px -18px -16px">
      📌 最高 ${fmtMoney(maxCost)} · 最低 ${fmtMoney(minCost)} · 橘色虛線為日均 · 滑鼠移到點上可看詳細數字
    </div>
  </div>`;
}

// F005 配送商損益紅綠榜
function renderF005() {
  const vendors = [...DATA.freight.vendors].sort((a, b) => b.amount - a.amount);
  const maxAbs = Math.max(...vendors.map(v => Math.abs(v.amount)));
  const scale = maxAbs * 1.1;

  const html = vendors.map(r => {
    const pct = (Math.abs(r.amount) / scale * 100).toFixed(1);
    const isOver = r.amount > 0;
    const tip = `合約預計：${fmtMoney(r.contract)}\n到點計價：${fmtMoney(r.actual)}\n差　　異：${r.amount > 0 ? '+' : ''}${r.amount.toLocaleString()}`;
    const bar = isOver
      ? `<div style="width:${pct}%;height:14px;background:var(--ry-red);margin-left:auto;border-radius:0 2px 2px 0" title="${tip}"></div>`
      : `<div style="width:${pct}%;height:14px;background:#1b7c33;margin-right:auto;border-radius:2px 0 0 2px" title="${tip}"></div>`;
    return `
    <div style="display:grid;grid-template-columns:64px 1fr 1fr 82px;gap:4px;padding:7px 14px;align-items:center;border-bottom:1px solid var(--ry-line)" title="${tip}">
      <span style="font-size:var(--fs-xs);font-weight:700">${r.name}</span>
      <div style="display:flex;height:14px">${isOver ? '' : bar}</div>
      <div style="display:flex;height:14px">${isOver ? bar : ''}</div>
      <span style="font-family:var(--f-mono);font-weight:700;font-size:var(--fs-xs);text-align:right;color:${isOver ? 'var(--ry-red)' : '#1b7c33'}">
        ${isOver ? '+' : ''}${r.amount.toLocaleString()}
      </span>
    </div>`;
  }).join('');

  return `
  <div class="w s6">
    <div class="gold-band" style="font-size:var(--fs-xs);line-height:1.6">
      F005 配送商損益紅綠榜 · 基準線（0）= 合約預計金額
    </div>
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:#E07855"></div>到點計價 vs 合約預計差異</div>
      <span class="wmeta">滑鼠移到長條可看明細</span>
    </div>
    <div style="margin:0 -18px -16px">
      <div style="display:grid;grid-template-columns:64px 1fr 1fr 82px;gap:4px;padding:8px 14px;font-family:var(--f-mono);font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--ry-muted);border-bottom:1.5px solid var(--ry-blue-dark);background:var(--ry-bg)">
        <span>廠商</span>
        <span style="text-align:right;padding-right:6px;color:var(--ry-red)">🔴 超支</span>
        <span style="padding-left:6px;color:#1b7c33">🟢 節省</span>
        <span style="text-align:right">差異金額</span>
      </div>
      ${html}
      <div style="padding:10px 14px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono);background:var(--ry-bg);line-height:1.7">
        📌 中心軸 0 = 合約預計金額 · 🔴 = 到點 &gt; 合約 · 🟢 = 到點 &lt; 合約
      </div>
    </div>
  </div>`;
}

// F004 配送商費用排行
function renderF004() {
  const vendors = [...DATA.freight.vendors].sort((a, b) => b.actual - a.actual);
  const maxActual = vendors[0].actual;

  const rows = vendors.map((v, i) => {
    const pct = (v.actual / maxActual * 100).toFixed(1);
    const rankColor = i === 0 ? '#f5c400' : i === 1 ? '#9ca3af' : i === 2 ? '#cd7f32' : '#dde2ec';
    return `
    <div style="padding:8px 14px;border-bottom:1px solid var(--ry-line)">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
        <span style="display:inline-block;width:20px;height:20px;background:${rankColor};color:white;border-radius:50%;font-size:var(--fs-xs);font-weight:900;font-family:var(--f-mono);text-align:center;line-height:20px">${i + 1}</span>
        <span style="font-size:var(--fs-sm);font-weight:700;flex:1">${v.name}</span>
        <span style="font-family:var(--f-mono);font-size:var(--fs-sm);font-weight:700;color:var(--ry-blue)">${fmtMoney(v.actual)}</span>
      </div>
      <div style="height:6px;background:var(--ry-bg);border-radius:99px;overflow:hidden">
        <div style="width:${pct}%;height:100%;background:var(--ry-blue);border-radius:99px;transition:width .4s"></div>
      </div>
    </div>`;
  }).join('');

  const total = vendors.reduce((s, v) => s + v.actual, 0);

  return `
  <div class="w s6">
    <div class="wh">
      <div class="wl"><div class="wdot"></div>F004 配送商費用排行</div>
      <span class="wmeta">依到點實際費用</span>
    </div>
    <div style="margin:0 -18px -16px">
      ${rows}
      <div style="padding:10px 14px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono);background:var(--ry-bg);border-top:1px solid var(--ry-line);line-height:1.7">
        📌 共 ${vendors.length} 家 · 總運費 ${fmtMoney(total)} · TOP3：${vendors.slice(0,3).map(v=>v.name).join('、')}
      </div>
    </div>
  </div>`;
}

// F010 三倉每日運費動支
function renderF010() {
  const budget = DATA.freight.warehouseBudget;
  const daily  = DATA.freight.dailyByWarehouse;
  const totalDays = daily.length;

  const dailyBudget = {
    '大溪倉': budget['大溪倉'] / totalDays,
    '岡山倉': budget['岡山倉'] / totalDays,
    '大肚倉': budget['大肚倉'] / totalDays,
  };

  const accumActual = { '大溪倉':0, '岡山倉':0, '大肚倉':0 };
  daily.forEach(row => {
    accumActual['大溪倉'] += row[1];
    accumActual['岡山倉'] += row[2];
    accumActual['大肚倉'] += row[3];
  });

  function cellsFor(actual, budgetVal) {
    const pct = actual / budgetVal * 100;
    const color = colorFor(pct);
    const bg = bgFor(pct);
    return `
      <td class="mono" style="text-align:right">${fmtMoney(Math.round(budgetVal))}</td>
      <td class="mono" style="text-align:right;color:${color};font-weight:700">${fmtMoney(actual)}</td>
      <td class="mono" style="text-align:right">
        <span style="display:inline-block;padding:2px 8px;background:${bg};color:${color};border-radius:99px;font-weight:700;font-size:var(--fs-xs)">${pct.toFixed(1)}%</span>
      </td>`;
  }

  const rows = daily.map(row => {
    const [date, d1, d2, d3] = row;
    return `<tr>
      <td style="font-weight:700;color:var(--ry-ink)">${date}</td>
      ${cellsFor(d1, dailyBudget['大溪倉'])}
      ${cellsFor(d2, dailyBudget['岡山倉'])}
      ${cellsFor(d3, dailyBudget['大肚倉'])}
    </tr>`;
  }).join('');

  function summaryCell(actual, budgetVal) {
    const pct = actual / budgetVal * 100;
    const color = colorFor(pct);
    return `
      <td class="mono" style="text-align:right">${fmtMoney(budgetVal)}</td>
      <td class="mono" style="text-align:right;color:${color};font-weight:800">${fmtMoney(actual)}</td>
      <td class="mono" style="text-align:right">
        <span style="display:inline-block;padding:3px 10px;background:${color};color:white;border-radius:99px;font-weight:800;font-size:var(--fs-xs)">${pct.toFixed(1)}%</span>
      </td>`;
  }

  const summary = `<tr style="background:var(--ry-blue-pale);border-top:2px solid var(--ry-blue-dark);position:sticky;bottom:0;font-weight:800">
    <td style="font-weight:900;color:var(--ry-blue-dark)">月累計</td>
    ${summaryCell(accumActual['大溪倉'], budget['大溪倉'])}
    ${summaryCell(accumActual['岡山倉'], budget['岡山倉'])}
    ${summaryCell(accumActual['大肚倉'], budget['大肚倉'])}
  </tr>`;

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:#E07855"></div>F010 三倉每日運費動支</div>
      <span class="wmeta">${totalDays} 天 · 縱向捲動 · 含月累計彙總</span>
    </div>
    <div style="margin:0 -18px -16px">
      <div style="max-height:420px;overflow-y:auto;overflow-x:auto">
        <table class="tbl" style="min-width:880px">
          <thead style="position:sticky;top:0;z-index:10">
            <tr style="background:var(--ry-blue-dark)">
              <th rowspan="2" style="color:white;border-right:2px solid var(--ry-blue);vertical-align:middle">日期</th>
              <th colspan="3" style="color:white;text-align:center;border-right:2px solid var(--ry-blue);background:#0E7BAD">🏭 大溪倉</th>
              <th colspan="3" style="color:white;text-align:center;border-right:2px solid var(--ry-blue);background:#E07855">🏭 岡山倉</th>
              <th colspan="3" style="color:white;text-align:center;background:#2DA870">🏭 大肚倉</th>
            </tr>
            <tr style="background:var(--ry-bg);border-bottom:2px solid var(--ry-blue-dark)">
              <th style="text-align:right">預算</th><th style="text-align:right">實際</th><th style="text-align:right">動支率</th>
              <th style="text-align:right">預算</th><th style="text-align:right">實際</th><th style="text-align:right">動支率</th>
              <th style="text-align:right">預算</th><th style="text-align:right">實際</th><th style="text-align:right">動支率</th>
            </tr>
          </thead>
          <tbody>${rows}${summary}</tbody>
        </table>
      </div>
      <div style="padding:10px 14px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono);background:var(--ry-bg);border-top:1px solid var(--ry-line);line-height:1.7">
        📌 單日預算 = 月預算 ÷ ${totalDays} 天 · 動支率 = 實際 / 單日預算<br>
        📌 三色門檻：&lt; 75% 🟢 安全 · 75–90% 🟡 注意 · &gt; 90% 🔴 危險
      </div>
    </div>
  </div>`;
}

// ── 總費用動支共用工具 ──

function getDispatchDailyFiltered() {
  const from = DATA.dateFrom.slice(5).replace('-', '/');
  const to   = DATA.dateTo.slice(5).replace('-', '/');
  return DATA.dispatch.daily.filter(row => row[0] >= from && row[0] <= to);
}

function sumWarehouse(rows, warehouseIdx) {
  const names = ['大溪倉', '大肚倉', '岡山倉'];
  const laborCol = 1 + warehouseIdx * 2;
  const freightCol = 2 + warehouseIdx * 2;
  let labor = 0, freight = 0;
  rows.forEach(r => { labor += r[laborCol]; freight += r[freightCol]; });
  const total = labor + freight;
  const b = DATA.dispatch.budget[names[warehouseIdx]];
  const totalDays = DATA.dispatch.daily.length;
  const periodRatio = rows.length / totalDays;
  const laborBudget   = b.labor   * periodRatio;
  const freightBudget = b.freight * periodRatio;
  const periodBudget  = laborBudget + freightBudget;
  return {
    name: names[warehouseIdx],
    labor, freight, total,
    laborBudget, freightBudget, budget: periodBudget,
    laborPct:   labor   / laborBudget   * 100,
    freightPct: freight / freightBudget * 100,
    pct:        total   / periodBudget  * 100,
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
    laborPct:   labor   / laborBudget   * 100,
    freightPct: freight / freightBudget * 100,
    pct:        total   / budget        * 100,
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
    <div class="w s3" style="border-top:3px solid ${isAll ? 'var(--ry-gold)' : color}">
      <div class="wh">
        <div class="wl"><div class="wdot" style="background:${color}"></div>${isAll ? '🌐 全區' : stats.name}</div>
        <span class="wmeta">${labelFor(stats.pct)}</span>
      </div>
      <div style="display:flex;align-items:baseline;gap:4px;margin-bottom:10px">
        <div class="kv" style="font-size:32px;color:${rateColor};line-height:1">${stats.pct.toFixed(1)}</div>
        <div style="font-size:16px;color:${rateColor};font-weight:700">%</div>
      </div>
      <div style="height:6px;background:var(--ry-bg);border-radius:99px;overflow:hidden;margin-bottom:10px;position:relative">
        <div style="width:${Math.min(stats.pct, 100)}%;height:100%;background:${rateColor};transition:width .4s"></div>
        <div style="position:absolute;left:90%;top:-2px;bottom:-2px;width:2px;background:var(--ry-ink);opacity:.5"></div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;padding-top:8px;border-top:1px solid var(--ry-line)">
        <div>
          <div style="font-size:9px;font-weight:700;color:var(--ry-muted);letter-spacing:.06em">人力費用</div>
          <div class="mono" style="font-size:var(--fs-sm);font-weight:700;color:var(--ry-ink);margin-top:2px">${fmtMoney(stats.labor)}</div>
        </div>
        <div>
          <div style="font-size:9px;font-weight:700;color:var(--ry-muted);letter-spacing:.06em">運務費用</div>
          <div class="mono" style="font-size:var(--fs-sm);font-weight:700;color:var(--ry-ink);margin-top:2px">${fmtMoney(stats.freight)}</div>
        </div>
      </div>
      <div style="margin-top:6px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono)">
        本期總計 ${fmtMoney(stats.total)} / 預算 ${fmtMoney(Math.round(stats.budget))}
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
    return `<th style="text-align:right;color:white;padding:12px 16px;font-size:14px;font-weight:900;letter-spacing:.04em;background:${whColor[w.name]}">${isAll ? '🌐 ' : ''}${w.name}</th>`;
  }).join('');

  function domainRows(domainLabel, domainIcon, domainBg, keyMap) {
    const budgetCells = warehouses.map(w => {
      const cellBg = w.name === '全區' ? 'background:var(--ry-blue-pale);' : '';
      return `<td class="mono" style="text-align:right;padding:10px 14px;${cellBg}">${fmtMoney(Math.round(w[keyMap.b]))}</td>`;
    }).join('');

    const actualCells = warehouses.map(w => {
      const cellBg = w.name === '全區' ? 'background:var(--ry-blue-pale);' : '';
      const color = colorFor(w[keyMap.p]);
      return `<td class="mono" style="text-align:right;padding:10px 14px;font-weight:700;color:${color};${cellBg}">${fmtMoney(w[keyMap.a])}</td>`;
    }).join('');

    const pctCells = warehouses.map(w => {
      const cellBg = w.name === '全區' ? 'background:var(--ry-blue-pale);' : '';
      const color = colorFor(w[keyMap.p]);
      return `<td style="text-align:right;padding:10px 14px;${cellBg}">
        <span style="display:inline-block;padding:3px 12px;background:${color};color:white;border-radius:99px;font-weight:800;font-size:var(--fs-xs);font-family:var(--f-mono)">${w[keyMap.p].toFixed(1)}%</span>
      </td>`;
    }).join('');

    const domainCell = `<td rowspan="3" style="vertical-align:middle;text-align:center;background:${domainBg};color:white;font-weight:800;font-size:var(--fs-sm);padding:14px 10px;border-right:2px solid var(--ry-white);letter-spacing:.04em;line-height:1.5">
      <div style="font-size:20px;margin-bottom:4px">${domainIcon}</div>
      ${domainLabel}
    </td>`;

    return `
      <tr style="border-top:2px solid var(--ry-line)">
        ${domainCell}
        <td style="padding:11px 16px;font-size:var(--fs-sm);font-weight:700;color:var(--ry-ink);background:var(--ry-paper);text-align:center;border-right:2px solid var(--ry-line)">預算</td>
        ${budgetCells}
      </tr>
      <tr>
        <td style="padding:11px 16px;font-size:var(--fs-sm);font-weight:700;color:var(--ry-ink);background:var(--ry-paper);text-align:center;border-right:2px solid var(--ry-line)">實際</td>
        ${actualCells}
      </tr>
      <tr>
        <td style="padding:11px 16px;font-size:var(--fs-sm);font-weight:800;text-align:center;color:var(--ry-ink);background:var(--ry-paper);border-right:2px solid var(--ry-line)">動支</td>
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
    <div style="margin:0 -18px -16px">
      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:var(--fs-sm);min-width:780px">
          <thead>
            <tr style="background:var(--ry-blue-dark);border-bottom:3px solid var(--ry-gold)">
              <th style="width:90px;color:white;text-align:center;padding:12px 14px;font-size:var(--fs-lg);font-weight:900;letter-spacing:.06em">領域</th>
              <th style="width:80px;color:white;text-align:center;padding:12px 14px;font-size:var(--fs-lg);font-weight:900;letter-spacing:.06em">項目</th>
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
      <div style="padding:10px 14px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono);background:var(--ry-bg);border-top:1px solid var(--ry-line);line-height:1.7">
        📌 動支% = 實際 ÷ 期間預算 · 期間預算 = 月預算 × (篩選天數 ÷ 當月天數)<br>
        📌 三色門檻：&lt; 75% 🟢 安全 · 75–90% 🟡 注意 · &gt; 90% 🔴 危險
      </div>
    </div>
  </div>`;
}

// T003 每日動支明細表
function renderT003() {
  const rows = getDispatchDailyFiltered();
  const totalDays = DATA.dispatch.daily.length;

  const dailyBudget = {
    daxi:     (DATA.dispatch.budget['大溪倉'].labor + DATA.dispatch.budget['大溪倉'].freight) / totalDays,
    dadu:     (DATA.dispatch.budget['大肚倉'].labor + DATA.dispatch.budget['大肚倉'].freight) / totalDays,
    gangshan: (DATA.dispatch.budget['岡山倉'].labor + DATA.dispatch.budget['岡山倉'].freight) / totalDays,
  };
  const dailyBudgetAll = dailyBudget.daxi + dailyBudget.dadu + dailyBudget.gangshan;

  function dayRows(row) {
    const [date, dxL, dxF, ddL, ddF, gsL, gsF] = row;
    const dxTotal = dxL + dxF;
    const ddTotal = ddL + ddF;
    const gsTotal = gsL + gsF;
    const allTotal = dxTotal + ddTotal + gsTotal;
    const dayPct = allTotal / dailyBudgetAll * 100;
    const dayColor = colorFor(dayPct);

    const laborRow = `<tr>
      <td rowspan="2" style="vertical-align:middle;font-weight:700;color:var(--ry-ink);background:var(--ry-paper);border-right:2px solid var(--ry-line)">${date}</td>
      <td style="font-size:var(--fs-xs);color:var(--ry-muted)">💰 人力</td>
      <td class="mono" style="text-align:right">${fmtMoney(dxL)}</td>
      <td class="mono" style="text-align:right">${fmtMoney(ddL)}</td>
      <td class="mono" style="text-align:right">${fmtMoney(gsL)}</td>
      <td class="mono" style="text-align:right;background:var(--ry-blue-pale);font-weight:700">${fmtMoney(dxL + ddL + gsL)}</td>
      <td rowspan="2" style="vertical-align:middle;text-align:center;background:${bgFor(dayPct)};border-left:2px solid var(--ry-line)">
        <span style="display:inline-block;padding:4px 12px;background:${dayColor};color:white;border-radius:99px;font-weight:800;font-size:var(--fs-xs);font-family:var(--f-mono)">${dayPct.toFixed(1)}%</span>
      </td>
    </tr>`;
    const freightRow = `<tr style="border-bottom:2px solid var(--ry-line)">
      <td style="font-size:var(--fs-xs);color:var(--ry-muted)">🚚 運務</td>
      <td class="mono" style="text-align:right">${fmtMoney(dxF)}</td>
      <td class="mono" style="text-align:right">${fmtMoney(ddF)}</td>
      <td class="mono" style="text-align:right">${fmtMoney(gsF)}</td>
      <td class="mono" style="text-align:right;background:var(--ry-blue-pale);font-weight:700">${fmtMoney(dxF + ddF + gsF)}</td>
    </tr>`;
    return laborRow + freightRow;
  }

  return `
  <div class="w s12">
    <div class="wh">
      <div class="wl"><div class="wdot" style="background:var(--ry-gold);"></div>T003 每日動支明細</div>
      <span class="wmeta">${rows.length} 天 · 縱向捲動</span>
    </div>
    <div style="margin:0 -18px -16px">
      <div style="max-height:520px;overflow-y:auto;overflow-x:auto">
        <table class="tbl" style="min-width:860px">
          <thead style="position:sticky;top:0;z-index:10">
            <tr style="background:var(--ry-blue-dark);border-bottom:3px solid var(--ry-gold)">
              <th style="width:90px;color:white;font-size:var(--fs-lg);font-weight:900;padding:12px 14px">日期</th>
              <th style="width:110px;color:white;font-size:var(--fs-lg);font-weight:900;padding:12px 14px">費用類別</th>
              <th style="color:white;font-size:var(--fs-lg);font-weight:900;padding:12px 14px;text-align:right;background:var(--tbl-daxi)">大溪倉</th>
              <th style="color:white;font-size:var(--fs-lg);font-weight:900;padding:12px 14px;text-align:right;background:var(--tbl-dadu)">大肚倉</th>
              <th style="color:white;font-size:var(--fs-lg);font-weight:900;padding:12px 14px;text-align:right;background:var(--tbl-gangshan)">岡山倉</th>
              <th style="color:white;font-size:var(--fs-lg);font-weight:900;padding:12px 14px;text-align:right;background:var(--tbl-all)">🌐 全區</th>
              <th style="color:white;font-size:var(--fs-lg);font-weight:900;padding:12px 14px;text-align:center">當日動支率</th>
            </tr>
          </thead>
          <tbody>${rows.map(dayRows).join('')}</tbody>
        </table>
      </div>
      <div style="padding:10px 14px;font-size:var(--fs-xs);color:var(--ry-muted);font-family:var(--f-mono);background:var(--ry-bg);border-top:1px solid var(--ry-line);line-height:1.7">
        📌 當日動支率 = (當日三倉人力+運務合計) ÷ (單日預算 × 3) · 單日預算 = 各倉月預算 ÷ ${totalDays} 天<br>
        📌 每日兩列：💰 人力 + 🚚 運務；全區欄位為三倉加總
      </div>
    </div>
  </div>`;
}
