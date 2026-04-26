# KPL 績效儀表板 · 開發防坑規格書

> 每次開發前請先讀這份文件，避免重複踩坑。

---

## 1. 專案架構

```
物流營運營表板/
├── kpl-dashboard/
│   ├── index.html          ← 主框架：Nav + Sidebar + 所有頁面模板
│   ├── login.html          ← 登入頁（sessionStorage 控制 auth）
│   ├── css/
│   │   ├── base.css        ← CSS 變數 Token（顏色、字型、間距）
│   │   ├── layout.css      ← Nav / Sidebar / Shell / Main
│   │   └── widget.css      ← 卡片、表格、進度條、Modal、Org
│   └── js/
│       ├── picks_data.js   ← 揀次靜態資料（let PICKS_RAW）
│       ├── labor_data.js   ← 工時靜態資料（let LABOR_RAW）
│       ├── data.js         ← 主資料物件（const DATA）
│       ├── utils.js        ← colorFor / fmtMoney 等工具函式
│       ├── widgets.js      ← renderXxx 函式（renderM012 等）
│       └── app.js          ← 頁面導航 + 所有頁面邏輯
└── DEVSPEC.md              ← 本文件
```

**載入順序（index.html 底部）：**
```
picks_data.js → labor_data.js → data.js → utils.js → widgets.js → app.js
```
順序錯誤會導致 `PICKS_RAW is not defined` 或 `DATA is not defined`。

---

## 2. 資料層設計

### 全域資料物件
| 變數 | 宣告方式 | 說明 |
|------|---------|------|
| `DATA` | `const` | 主資料（預算、運費、dispatch）。子欄位可被修改但物件本身不能重新賦值 |
| `PICKS_RAW` | **`let`** | 揀次陣列。資料匯入後 `PICKS_RAW = parsedPicks.records` 重新賦值 |
| `LABOR_RAW` | **`let`** | 工時陣列。資料匯入後 `LABOR_RAW = parsedLabor.records` 重新賦值 |

**坑：** 未來新增可上傳的資料陣列，一律用 `let` 而非 `const`，否則 `applyXxx()` 賦值時會報 TypeError。

### Excel 日期 Serial 轉換
```js
const d = new Date(Math.round((serial - 25569) * 86400000));
dateStr = d.toISOString().slice(0, 10); // → "2026-03-01"
```
Excel 日期是數字（如 46082），必須用此公式轉換。切勿用 `new Date(serial)`。

---

## 3. 頁面模板系統

每個頁面是 `index.html` 裡的 `<script id="tpl-xxx" type="text/html">` 標籤。

**新增頁面的完整 Checklist：**

1. 在 `index.html` 加 `<script id="tpl-新頁面" type="text/html">` 模板
2. 在 `index.html` 底部的 `PAGE_TEMPLATES` 物件加 `新頁面: document.getElementById('tpl-新頁面').innerHTML`
3. 在 `app.js` 的 `PAGES` 陣列加 `{ id:'新頁面', icon:'X', label:'...', status:'ready' }`
4. 在 `app.js` 的 `loadPage()` 加 `else if (pageId === '新頁面') init新頁面Page()`
5. 實作 `init新頁面Page()` 和 `render新頁面Page()`

**缺任何一步都會造成頁面空白或找不到頁面。**

---

## 4. 資料匯入頁面規範

每新增一種 Excel 匯入，需補齊以下 5 個地方：

| 步驟 | 位置 | 內容 |
|------|------|------|
| 1 | `index.html tpl-import` | 新增上傳 drop zone 卡片（`id="xxx-drop"`, `id="xxx-file"`, `id="xxx-import-status"`, `id="xxx-preview"`, `id="xxx-btns"`） |
| 2 | `app.js parseExcel()` | 加 `else if (type === 'xxx') parseXxx(wb, file.name)` |
| 3 | `app.js` | 實作 `parseXxx()`, `showXxxPreview()`, `applyXxx()`, `resetXxx()` |
| 4 | `app.js updateStatus()` | 在 `rows` 陣列加對應資料的狀態列 |
| 5 | `js/xxx_data.js` | 靜態資料用 `let` 宣告，讓 `applyXxx()` 能重新賦值 |

**坑：元素 ID 命名**
匯入頁面的 status badge 使用 `xxx-import-status`，不要與分析頁面的 `xxx-meta` 衝突。
例：揀次上傳區用 `picks-import-status`，揀次分析頁的描述欄用 `picks-meta`。

---

## 5. CSS 系統

### CSS 變數（base.css）
```css
--ry-blue      /* 主色 #1e5ca8 */
--ry-blue-dark /* 深藍 #123d74 */
--ry-gold      /* 金色 #f5c400 */
--ry-red       /* 紅色 #d9401b */
--ry-orange    /* 橘色 #e07855 */
--ry-green     /* 綠色 #2ea85a */
--ry-muted     /* 灰色 #5a6478 */
--ry-line      /* 線條 #e5e7eb */
--ry-bg        /* 背景 #f8f9fb */
--ry-ink       /* 文字 #1a1e2a */
```

### Grid 系統
`.w.s3` `.w.s4` `.w.s6` `.w.s8` `.w.s12`（12 格欄）
- `s3/s4` 在 1100px 以下變 span 6
- `s6` 在 1100px 以下變 span 12
- 匯入頁三卡並排使用 `s4`，讓 RWD 自然折行

---

## 6. HTML onclick 陷阱

**❌ 錯誤寫法（bind 在 HTML attribute 不支援）：**
```html
onclick="fn.bind(null,${i})('${c}')"
```

**✅ 正確寫法（直接傳參數）：**
```html
onclick="fn(${i},'${c}')"
```
或使用 helper function 產生 HTML 字串：
```js
const swatches = (i, sel) => COLORS.map(c =>
  `<div onclick="setColor(${i},'${c}')"></div>`
).join('');
```

---

## 7. Excel 欄位格式規範

### 運務費用（運務費用.xlsx）
- 分頁名：`進貨日與計價費用`
- 欄位：`列標籤`（民國日期 YYY/MM/DD）｜`大肚倉`｜`大溪倉`｜`岡山倉`｜`總計`

### 工時資料（全區人力工時.xlsx）
- 說明：涵蓋全倉全課，數萬筆；EC理貨課 0301 工時只是初版樣本資料
- 分頁：任意（取第一個分頁）
- 必要欄位：`倉別`｜`日期`（Excel serial）｜`廠商`｜`班別`｜`員編`｜`作業課別`｜`姓名`｜`作業區域`｜`作業時數`｜`實際費用`
- 選填欄位：`裝箱時數`｜`夜間時數`｜`正常時數`
- 過濾規則：排除 `作業區域 === '午休時間'` 和 `作業時數 <= 0`

### 揀次資料（貨量揀次.xlsx）
- 分頁：任意（取第一個分頁）
- 必要欄位：`倉別`｜`日期`（Excel serial）｜`業務類別`｜`作業區`｜`工時區域`｜`揀次`
- 過濾規則：排除 `揀次 <= 0`

### 待整合（資料尚未提供）
- `EC出貨預估與時機.xlsx` — EC出貨預估 vs 實際對比頁面
- `年度預算表.xlsx` — 年度預算整合到各儀表板 KPI

---

## 8. 狀態管理

- **登入狀態**：`sessionStorage.getItem('kpl_auth')` / `sessionStorage.getItem('kpl_user')`
- **當前頁面**：`currentPageId`（全域 let）
- **Org 編輯狀態**：`orgEditWh`、`orgEditDept`（全域 let，-1 表示未編輯）
- **已上傳資料**：`parsedFreight`、`parsedLabor`、`parsedPicks`（全域 let，null 表示未上傳）

---

## 9. 待辦 / 已知 WIP 頁面

| 頁面 | 狀態 | 待確認事項 |
|------|------|-----------|
| 人效監控 | WIP | 揀次 + 工時合併邏輯（PPH）確認後開發 |
| 月度結算 | WIP | 月底結算流程確認後開發 |
| EC出貨分析 | 未建 | 待「EC出貨預估與時機.xlsx」提供 |
| 年度預算 | 未建 | 待「年度預算表.xlsx」提供 |

---

## 10. Git / 部署

- 倉庫：`inanuna-flow/-`，GitHub Pages 自動部署
- 推送指令：`git add -A && git commit -m "..." && git push`
- 登入頁保留 Dev Bypass 按鈕，待 IT 部門提供 CheckUserId API 後移除
