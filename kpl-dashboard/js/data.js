// ═══════════════════════════════════════════════════════
// data.js · 資料物件（未來改 API 只動這裡）
// ═══════════════════════════════════════════════════════
//
// 給資訊部的接手指引：
// 把 DATA 物件改成 async 函式即可：
//   let DATA = {};
//   async function loadData(filters) {
//     const res = await fetch('/api/kpl?' + new URLSearchParams(filters));
//     DATA = await res.json();
//   }
// ═══════════════════════════════════════════════════════

const DATA = {
  // 日期相關
  dateFrom:   '2026-03-01',
  dateTo:     '2026-03-31',
  dayOfMonth: 0,
  totalDays:  31,

  // 預算 & 累計費用（M012）
  budget:     0,
  actual:     0,

  // M015 Business Units
  units: [],
  thresholdPeak:   30,
  thresholdStable: 10,

  // 運費資料（F001-F009）
  freight: {
    totalCost:       0,
    lastMonthCost:   0,
    totalOrders:     0,

    estimatedCost:   0,
    actualCost:      0,

    overCount:       0,
    saveCount:       0,
    diffThreshold:   90,

    vendors: [],

    dailyTrend: [],

    warehouseBudget: {
      '大溪倉': 0,
      '岡山倉': 0,
      '大肚倉': 0,
    },

    // F010 每倉每日實際運費 [日期, 大溪, 岡山, 大肚]
    dailyByWarehouse: [],
  },

  // 年度預算（來源：年度預算.xlsx · 人力預算_轉換 / 運費預算_轉換）
  // 陣列 index 0~11 對應 1月~12月
  annualBudget: {
    labor: {
      '大溪倉': Array(12).fill(0),
      '大肚倉': Array(12).fill(0),
      '岡山倉': Array(12).fill(0),
    },
    freight: {
      '大溪倉': Array(12).fill(0),
      '大肚倉': Array(12).fill(0),
      '岡山倉': Array(12).fill(0),
    },
  },

  // 組織設定
  org: {
    warehouses: [
      { name: '大溪倉', region: '北區營運部',  color: '#1e5ca8' },
      { name: '岡山倉', region: '中南區營運部', color: '#d9401b' },
      { name: '大肚倉', region: '中南區營運部', color: '#2ea85a' },
    ],
    depts: [
      { name: 'EC驗收課',     type: '服務EC',  wh: '大溪倉', color: '#1e5ca8' },
      { name: 'EC理貨課',     type: '服務EC',  wh: '大溪倉', color: '#2ea85a' },
      { name: '岡山營運課',   type: '服務EC',  wh: '岡山倉', color: '#d9401b' },
      { name: '大肚服務EC課', type: '服務EC',  wh: '大肚倉', color: '#6366f1' },
      { name: '商品理貨課',   type: '營收EC',  wh: '大溪倉', color: '#f59e0b' },
      { name: '倉儲管理課',   type: '營收EC',  wh: '大溪倉', color: '#7c3aed' },
      { name: '大肚營收EC課', type: '營收EC',  wh: '大肚倉', color: '#0ea5e9' },
      { name: '其他單位',     type: '後勤支援', wh: '大溪倉', color: '#9ca3af' },
      { name: '大肚運務課',   type: '後勤支援', wh: '大肚倉', color: '#d1d5db' },
    ],
  },

  // 總費用動支率（T001 / T002 / T003）
  dispatch: {
    latestUploadDate: '',

    budget: {
      '大溪倉': { labor: 0, freight: 0 },
      '大肚倉': { labor: 0, freight: 0 },
      '岡山倉': { labor: 0, freight: 0 },
    },

    // [日期, 大溪人力, 大溪運務, 大肚人力, 大肚運務, 岡山人力, 岡山運務]
    daily: [],
  },
};
