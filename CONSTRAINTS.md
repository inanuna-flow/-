# 🚫 KPL Dashboard 重做的技術約束

> **這些是「鐵則」。AI 必須嚴格遵守。**
> 違反任一條都會破壞功能或破壞部署。

---

## 1️⃣ 絕對不能改的檔案

| 檔案 | 角色 | 為什麼不能改 |
|------|------|------------|
| `server.js` | 後端 API + 資料庫整合 | 改了會破壞雲端 DB 上傳、Session 機制 |
| `kpl-dashboard/js/data.js` | 預算資料處理 | 改了會破壞數字計算 |
| `kpl-dashboard/js/picks_data.js` | 揀次資料處理 | 改了會破壞揀次分析 |
| `kpl-dashboard/js/labor_data.js` | 人力資料處理 | 改了會破壞人力分析 |
| `kpl-dashboard/js/utils.js` | 共用工具函數 | 改了會連鎖破壞 |
| `cloudbuild.yaml` | GCP 部署設定 | 改了會破壞自動部署 |
| `Dockerfile` | 容器建構設定 | 改了會破壞 image build |
| `package.json` / `package-lock.json` | 依賴清單 | 改了會破壞建構 |
| `page_permissions.json` | 權限資料 | 這是執行期資料，不要動 |

---

## 2️⃣ 慎改檔案（可改但要小心）

| 檔案 | 可改範圍 | 不可改範圍 |
|------|---------|----------|
| `kpl-dashboard/js/widgets.js` | HTML 字串內的 class 名、結構 | 函數名、資料計算、資料屬性傳遞 |
| `kpl-dashboard/js/app.js` | 渲染相關的 HTML 結構 | 路由邏輯、fetch 呼叫、事件綁定 |
| `kpl-dashboard/login.html` | 視覺、版型 | 表單欄位 name、fetch URL |

---

## 3️⃣ 可自由改的檔案

| 檔案 | 說明 |
|------|------|
| `kpl-dashboard/index.html` | 11 個 template + 主結構（注意保留 onclick/id）|
| `kpl-dashboard/css/base.css` | 設計 token，可整個重寫 |
| `kpl-dashboard/css/layout.css` | 排版結構，可整個重寫 |
| `kpl-dashboard/css/widget.css` | UI 元件，可整個重寫 |

---

## 4️⃣ API 必須相容（不能改 URL、不能改參數、不能改 method）

```
GET    /api/session                                  ← 檢查登入狀態
POST   /api/check-user                               ← 登入（送 USER_ID, PSW）
POST   /api/logout                                   ← 登出
GET    /api/page-permissions                         ← 讀取頁面權限
POST   /api/page-permissions                         ← 儲存頁面權限（管理員用）
GET    /api/data/range                               ← 取得資料日期範圍
GET    /api/data/budget?year=YYYY                    ← 取得年度預算
GET    /api/data/labor?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD
GET    /api/data/picks?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD
GET    /api/data/freight-mainline?date_from=...&date_to=...      ← 主線運費（新）
GET    /api/data/freight-non-mainline?date_from=...&date_to=...  ← 非主線運費（新）
POST   /api/import/budget                            ← 上傳預算
POST   /api/import/labor                             ← 上傳人力
POST   /api/import/picks                             ← 上傳揀次
POST   /api/import/freight-mainline                  ← 主線運費（新；支援 dryRun: true）
POST   /api/import/freight-non-mainline              ← 非主線運費（新；支援 dryRun: true）
```

> ⚠️ **已淘汰：**`/api/import/freight` 與 `/api/data/freight` 回 410 Gone，請改用上方主線 / 非主線版本。
> 💡 **dryRun 機制：**運費 import 端點接受 `{dryRun: true}`，回傳 `{rowsToImport, existingInRange, dateFrom, dateTo}`；前端據此跳 confirm 對話框提示「將覆蓋 N 筆」。

**回應格式：** 全部都是 JSON，成功回 `{ok: true, ...}`，失敗回 `{ok: false, MSG: "999 錯誤訊息"}`

---

## 5️⃣ 必須保留的 DOM ID（JS 還在用）

### 全站
- `#sidebar`, `#main`, `#topbar-tabs`
- `#menu-toggle`, `#user-btn`, `#user-menu`
- `#topbar-user-name`, `#topbar-user-avatar`
- `#um-name`, `#um-role`
- `#sidebar-overlay`
- `#toast`
- `#notif-btn`
- `#boot-loader`

### 各頁篩選器
- `#filter-from`, `#filter-to`, `#filter-meta`
- `#dispatch-from`, `#dispatch-to`
- `#freight-from`, `#freight-to`
- `#picks-from`, `#picks-to`, `#picks-wh`, `#picks-op`, `#picks-date-meta`
- `#labor-from`, `#labor-to`, `#labor-shift`, `#labor-vendor`, `#labor-date-meta`, `#labor-filter-meta`
- `#attendance-from`, `#attendance-to`, `#attendance-dept`, `#attendance-shift`, `#attendance-emp-id`, `#attendance-emp-list`, `#attendance-date-meta`, `#attendance-meta`
- `#productivity-from`, `#productivity-to`, `#productivity-date-meta`
- `#monthly-from`, `#monthly-to`, `#monthly-date-meta`

### 各頁網格容器
- `#daily-grid`, `#dispatch-grid`, `#freight-grid`, `#picks-grid`, `#labor-grid`, `#attendance-grid`
- `#productivity-grid`, `#monthly-grid`, `#annual-grid`, `#org-grid`, `#typography-grid`

### 各頁 meta 文字
- `#page-meta`, `#dispatch-meta`, `#freight-meta`, `#picks-meta`, `#labor-meta`
- `#monthly-meta`, `#productivity-meta`

### 資料匯入頁（特別重要，元件多）
- `#status-tbody`
- `#budget-drop`, `#budget-file`, `#budget-status`, `#budget-preview`, `#budget-btns`
- `#freight-mainline-drop`, `#freight-mainline-file`, `#freight-mainline-status`, `#freight-mainline-preview`, `#freight-mainline-btns`
- `#freight-non-mainline-drop`, `#freight-non-mainline-file`, `#freight-non-mainline-status`, `#freight-non-mainline-preview`, `#freight-non-mainline-btns`
- `#labor-drop`, `#labor-file`, `#labor-status`, `#labor-preview`, `#labor-btns`
- `#picks-drop`, `#picks-file`, `#picks-status`, `#picks-preview`, `#picks-btns`
- ~~`#freight-drop` 等舊版~~（已拆成主線/非主線）

### 登入頁
- `#userId`, `#password`, `#login-btn`, `#error-msg`

---

## 6️⃣ 必須保留的 inline onclick / onchange 函數名

> AI 重做時可以改 class 名、改 HTML 結構，但 **onclick="..."** 裡的函數名必須保留。

### 全站
- `toggleSidebar()`, `toggleUserMenu()`, `logout()`

### 各頁套用按鈕
- `applyFilter()` — daily
- `applyDispatchFilter()` — dispatch
- `applyFreightFilter()` — freight
- `applyDashboardDateFilter('picks' | 'labor' | 'attendance' | 'productivity' | 'monthly')`
- `renderPicksPage()`, `renderLaborPage()` — 給 select 的 onchange
- `renderAttendancePage()` — 個人出勤頁課別/班別 onchange、員編 oninput（即時查詢；走既有 `/api/data/labor`，無新端點）
- `downloadAttendance()` — 個人出勤頁「↓ 下載明細」按鈕（XLSX 匯出當前篩選明細）

### 資料匯入頁
- `onDragOver(event, dropId)`, `onDragLeave(dropId)`, `onDrop(event, type)`
- `onFileSelect(event, type)` — type 支援：`budget` / `labor` / `picks` / `freight-mainline` / `freight-non-mainline` / `freight`（舊，僅本地預覽）
- `applyBudget()`, `applyLabor()`, `applyPicks()`
- `applyFreightMainline()`, `applyFreightNonMainline()` — 新（上傳至 Cloud SQL，含 dry-run 警告）
- `resetBudget()`, `resetLabor()`, `resetPicks()`
- `resetFreightMainline()`, `resetFreightNonMainline()` — 新
- `applyFreight()`, `resetFreight()`, `downloadFreight()` — 舊（僅本地預覽用，未來會刪）

### 組織設定頁
- `saveOrgSettings()`

### 文字設定頁
- `saveTypographySettings()`, `resetTypographySettings()`

---

## 7️⃣ 外部依賴（不要換、不要升級）

| 套件 | 來源 | 用途 |
|------|------|------|
| XLSX.js v0.18.5 | `https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js` | Excel 讀寫 |

**不要引入：**
- ❌ React / Vue / Svelte / Angular
- ❌ jQuery / Lodash（沒在用，純 vanilla）
- ❌ Tailwind CSS（會破壞既有 CSS）
- ❌ Bootstrap（同上）
- ❌ Chart.js / D3 / ECharts（圖表都是純 HTML/CSS/SVG 寫的）
- ❌ Material UI / Ant Design 等元件庫

> 💡 如果真的需要圖表函式庫，**事先問使用者**。

---

## 8️⃣ 部署環境

| 項目 | 規格 |
|------|------|
| 平台 | GCP Cloud Run |
| 區域 | asia-east1 |
| 容器埠號 | 8080 |
| Node 版本 | 18+ |
| 資料庫 | Cloud SQL PostgreSQL |
| Schema | `kpl_dashboard` |

部署流程：
1. `git push origin main`
2. Cloud Build 自動觸發（讀 `cloudbuild.yaml`）
3. 約 2-3 分鐘後 image 推到 Artifact Registry
4. 使用者手動到 Cloud Run Console 部署最新 image

---

## 9️⃣ 安全性要求（不要破壞）

- ✅ 保留 `httpOnly` + `SameSite=Strict` cookie 機制
- ✅ 保留 Rate Limit（15 分鐘 10 次登入）
- ✅ 保留 Session 重新驗證機制
- ✅ 保留 CSP / X-Frame-Options 等 security headers
- ✅ 資料匯入 / 變更端點（`/api/import/*`）後端須經 `requireDataManager`（B 級以上），C 級一般使用者禁止（前端隱藏頁面不足以把關）
- ✅ 後台管理端點（帳號 / IP / 權限儲存）須經 `sessionIsAdmin` + 密碼二次驗證；資料讀取（圖表）端點維持 `requireSession`（人人可讀）
- ❌ 不要在前端寫入任何密碼 / API Key
- ❌ 不要把資料寫到 localStorage / sessionStorage（除了臨時 UI 狀態）

---

## 9️⃣b 環境變數與機密管理

本專案使用 `.env` 管理本機開發環境所需的環境變數。

### 規則

- `.env` **不得提交**到 GitHub
- `.env` 必須加入 `.gitignore`（已加好）
- 專案只提供 `.env.example` 作為範例
- `.env.example` **不得包含**真實 API Key、密碼、Token 或任何敏感資訊
- 新增環境變數時，必須**同步更新** `.env.example`
- 若 `.env` 或任何密鑰不小心被提交，必須**立即撤銷**並重新產生相關金鑰

### 檔案說明

| 檔案 | 是否提交 | 說明 |
|------|---------|------|
| `.env` | ❌ 不提交 | 本機實際使用的環境變數，可能包含敏感資訊 |
| `.env.local` | ❌ 不提交 | 本機覆蓋設定 |
| `.env.example` | ✅ 可提交 | 給開發者參考的環境變數範本，不放真實值 |

### 範例

請複製 `.env.example` 建立自己的 `.env`：

```bash
cp .env.example .env
```

### 正式環境（Cloud Run）

正式環境的 `DB_PASSWORD` 等密鑰**不放 .env**，而是直接設在 Cloud Run 的「變數與密鑰」分頁，由 Google 託管。

---

## 🔟 驗證清單（每階段完成後必跑）

### 視覺
- [ ] 11 個頁面都能切換、不破版
- [ ] topbar 在所有頁面正常顯示
- [ ] sidebar 在所有頁面正常顯示
- [ ] 響應式：1280px 桌機正常

### 功能
- [ ] 登入：用 EIP 帳號可成功登入
- [ ] 登出：點選後跳回登入頁
- [ ] 上傳：上傳一份小 Excel，顯示「000 Imported」
- [ ] 查詢：換瀏覽器登入後看到一樣的雲端資料
- [ ] 篩選：日期區間 + 下拉選擇都能即時更新內容
- [ ] 管理員：`inari` 帳號可進管理頁，可控制其他人權限

### 部署
- [ ] `git push` 後 Cloud Build 成功
- [ ] 手動部署最新 image 後，訪問網址正常
- [ ] 上面的功能驗證全部再跑一次

---

## ❓ 不確定時怎麼辦

**絕對不要硬幹。** 使用者明確說過：「失敗 3 次必停，盤點後給 3 個方向選」。

如果你（AI）遇到不確定，**先問使用者**：
- 「這個元件你希望保留還是換成 _____ ？」
- 「我看到 X 跟 Y 都可以實現，你偏好哪個？」
- 「我注意到 Z 檔案沒在 CONSTRAINTS.md，可以動嗎？」

---

## 📌 一句話摘要

> **改 CSS 和 HTML 結構 OK；改 JS 邏輯、API、後端、部署設定 → 停下來問使用者。**
