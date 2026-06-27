# 📋 KPL Dashboard 功能保留清單

> 重設視覺時，**以下功能必須 100% 保留**。
> 視覺可以改、HTML 結構可以調，但「能做什麼」不能少。

---

## 🏠 通用結構（所有頁面共用）

### Topbar（頂部列）
- 左側：漢堡選單按鈕（手機收合 sidebar 用）+ 主分類 tabs
- 右側：通知鈴鐺（紅點代表有未讀）+ 使用者選單（含登出）
- 高度約 56px，固定置頂

### Sidebar（左側選單）
- 顯示頁面分類與項目清單
- 管理員（`inari`）會多出「管理員設定」項目（金色標示）
- 手機尺寸自動收合，點漢堡開啟

### Page Header（每頁頂部）
- Eyebrow 小字（英文分類，如 `DAILY BUDGET MONITOR`）
- 大標 h1（中文頁名）
- 副標（一行說明 + 載入狀態）

### Filter Bar（篩選列）
- 日期區間：from + to 日期選擇器
- 「套用」按鈕（藍色主按鈕）
- 篩選結果提示文字

### Toast（提示訊息）
- 右下角彈出，3 秒自動消失
- 用於：上傳成功、儲存完成、錯誤提示

### Boot Loader（開機載入畫面）
- 登入後立即顯示
- 內容：品牌名 + spinner + 載入中...
- 資料載入完淡出

---

## 📄 11 個頁面詳細功能

### 1. `tpl-daily` 每日動支監控
- **路由 id**：`daily`
- **篩選器**：日期區間 from/to
- **內容**：
  - M012 - 預算達成率 KPI 卡片
  - M015 - 月度費用時間軸圖表
- **互動**：套用按鈕 `applyFilter()`

### 2. `tpl-dispatch` 總費用動支率
- **路由 id**：`dispatch`
- **篩選器**：日期區間 from/to（與其他頁共用）
- **內容**：
  - 預算狀態提示
  - 動支率 KPI
- **互動**：套用按鈕 `applyDispatchFilter()`

### 3. `tpl-freight` 運費損益分析
- **路由 id**：`freight`
- **篩選器**：日期區間 from/to
- **內容**（多個 widget 組合，請保留全部）：
  - `renderFreightKpiOverview` - 總覽 KPI
  - `renderFreightWarehouseBullets` - 各倉表現條圖
  - `renderFreightWaterfall` - 瀑布圖
  - `renderFreightDailyHeatmap` - 每日熱力圖
  - `renderFreightDecisionTable` - 決策表格
  - `renderFreightBudgetActualCombo` - 預算/實際對照
  - `renderFreightStructureBars` - 結構長條圖
  - `renderFreightMonthlyBudgetMatrix` - 月度預算矩陣
  - `renderFreightReferenceDashboard` - 參考儀表板
- **互動**：套用按鈕 `applyFreightFilter()`

### 4. `tpl-picks` 揀次分析
- **路由 id**：`picks`
- **篩選器**：日期區間 + 倉庫下拉 + 作業區域下拉
- **內容**：揀次明細分析、人效、區域比較
- **互動**：
  - 套用按鈕 `applyDashboardDateFilter('picks')`
  - 倉庫切換 `renderPicksPage()`
  - 作業區域切換 `renderPicksPage()`

### 5. `tpl-labor` 人力工時結構
- **路由 id**：`labor`
- **篩選器**：日期區間 + 班別下拉（日/中/夜）+ 課別下拉
- **內容**：工時分布、課別比較、班別比較
- **互動**：
  - 套用按鈕 `applyDashboardDateFilter('labor')`
  - 班別切換 `renderLaborPage()`
  - 課別切換 `renderLaborPage()`

### 5b. `tpl-attendance` 個人出勤查詢 ⭐ 新增
- **路由 id**：`attendance`（側邊欄「成本分析」群組）
- **篩選器**：日期區間 + 課別下拉 + 班別下拉 + 員工編號（選填，input + datalist 自動補全）
- **兩種模式**：
  - **全部員工（員編留空）**：彙總卡（員工數/出勤人日/總工時/總人力成本）+ 員工出勤彙總表（員編/課別/出勤天數/總工時/人力成本），**點任一列鑽進該員明細**
  - **單一員工（填員編）**：摘要卡（出勤天數/總工時/平均每日工時/人力成本）+ 逐日出勤明細表，含「← 回全部員工」連結
- **資料來源**：複用 `LABOR_RAW`（既有 `/api/data/labor`，**無新端點**），前端依 課別/班別/員編 client-side 過濾
- **互動**：
  - 套用按鈕 `applyDashboardDateFilter('attendance')`
  - 課別/班別下拉 `onchange="renderAttendancePage()"`；員編 `oninput="renderAttendancePage()"`（即時）
  - 下載明細 `downloadAttendance()` — 匯出當前篩選的逐筆明細 Excel（欄位：日期/倉別/廠商/班別/員編/作業課別/作業時數）
- **備註**：雲端只存員編不存姓名（姓名與打卡起訖時間未納入下載，資料源/DB 未存）；排除「午休時間」與 0 工時列

### 6. `tpl-productivity` 揀次人效分析
- **路由 id**：`productivity`
- **篩選器**：日期區間
- **內容**：整合揀次、工時、成本與區域效率
- **互動**：套用按鈕 `applyDashboardDateFilter('productivity')`

### 7. `tpl-monthly` 月度結算
- **路由 id**：`monthly`
- **篩選器**：日期區間
- **內容**：本月口袋總結（給月底結算用）
- **互動**：套用按鈕 `applyDashboardDateFilter('monthly')`

### 8. `tpl-annual` 年度規劃分析
- **路由 id**：`annual`
- **無篩選器**（自動顯示當年度）
- **內容**：12 個月預算 vs 實際（人力 + 運務）
- **互動**：無（純檢視）

### 9. `tpl-import` 資料匯入 ⭐ **第一階段重做這頁**
- **路由 id**：`import`
- **內容區塊**：
  - **儀表板資料狀態表**（id: `status-tbody`）
    - 欄位：資料類型 / 狀態 / 最新資料日期 / 最舊資料日期
  - **5 個上傳卡片**（每個都有拖曳區 + 預覽 + 按鈕）：
    - 💰 年度預算（budget）
    - 🚚 主線運費（freight-mainline）★ 新；上傳至 Cloud SQL `freight_mainline_daily`
    - 🚛 非主線運費（freight-non-mainline）★ 新；上傳至 Cloud SQL `freight_non_mainline_daily`
    - ⏱ 工時資料（labor）
    - ⚡ 揀次資料（picks）
  - 舊「🚚 運務費用」單卡已淘汰（拆成上面兩張）
- **互動函數**（必須保留名稱）：
  - `onDragOver(event, id)`, `onDragLeave(id)`, `onDrop(event, type)`
  - `onFileSelect(event, type)` — type 支援：`budget` / `labor` / `picks` / `freight-mainline` / `freight-non-mainline`
  - `applyBudget()`, `applyLabor()`, `applyPicks()`
  - `applyFreightMainline()`, `applyFreightNonMainline()` — 含 dry-run 覆蓋警告
  - `resetBudget()`, `resetLabor()`, `resetPicks()`
  - `resetFreightMainline()`, `resetFreightNonMainline()`
- **必須保留的 DOM id**：
  - `status-tbody`
  - `budget-drop`, `budget-file`, `budget-status`, `budget-preview`, `budget-btns`
  - `freight-mainline-drop`, `freight-mainline-file`, `freight-mainline-status`, `freight-mainline-preview`, `freight-mainline-btns`
  - `freight-non-mainline-drop`, `freight-non-mainline-file`, `freight-non-mainline-status`, `freight-non-mainline-preview`, `freight-non-mainline-btns`
  - 同 pattern：`labor-*`, `picks-*`

### 10. `tpl-org` 組織設定
- **路由 id**：`org`
- **內容**：可編輯的組織架構表格
- **互動**：點任一列進入編輯 + 儲存按鈕 `saveOrgSettings()`

### 11. `tpl-typography` 文字樣式設定
- **路由 id**：`typography`
- **內容**：全站文字層級調整 + 即時預覽
- **互動**：
  - 即時預覽 `renderTypographyPreview()`
  - 重設按鈕 `resetTypographySettings()`
  - 儲存按鈕 `saveTypographySettings()`

---

## 🔐 登入頁（login.html）

- **路徑**：根目錄即顯示 / 或未登入時自動跳轉
- **欄位**：
  - 帳號（USER_ID）
  - 密碼（PSW）
- **互動**：
  - Enter 觸發登入
  - 「忘記密碼」連結（如有）
  - 顯示錯誤訊息區
- **驗證機制**：
  - 已登入會自動跳 `index.html`（透過 `/api/session` 檢查）
  - 登入打 `/api/check-user` POST

---

## 👑 管理員專屬功能

- **觸發條件**：使用者 ID = `inari`
- **額外頁面**：「管理員設定」可控制其他帳號能看到哪些頁面
- **權限儲存**：寫入 `page_permissions.json`（後端 `/api/page-permissions`）
- **二次驗證**：儲存權限時要再次輸入密碼

---

## 🎯 視覺重做時要保留的核心使用體驗

1. **使用者一登入就立刻看到 boot loader**（不要空白頁）
2. **頂部 topbar 永遠在**（切頁不會閃爍）
3. **左側 sidebar 永遠在**（管理員會看到金色「管理員」標示）
4. **每頁都有 Page Header**（eyebrow + h1 + 副標）
5. **每頁都有 Filter Bar**（除了 annual / org / typography）
6. **資料載入時要有狀態指示**（不要靜默的空畫面）
7. **操作完成要有 Toast 提示**（成功 / 失敗）

---

## 📊 資料來源（給 AI 理解）

所有資料來自 GCP Cloud SQL（PostgreSQL），透過後端 API 取得：

| 資料類型 | 來源 API |
|---------|---------|
| 人力工時 | `GET /api/data/labor?date_from=&date_to=` |
| 揀次資料 | `GET /api/data/picks?date_from=&date_to=` |
| 年度預算 | `GET /api/data/budget?year=` |
| 資料日期範圍 | `GET /api/data/range` |
| 上傳人力 | `POST /api/import/labor` |
| 上傳揀次 | `POST /api/import/picks` |
| 上傳預算 | `POST /api/import/budget` |
| 登入 | `POST /api/check-user` |
| Session 檢查 | `GET /api/session` |
| 登出 | `POST /api/logout` |
| 讀取頁面權限 | `GET /api/page-permissions` |
| 儲存頁面權限 | `POST /api/page-permissions` |

**重做視覺時這些 API 不能改、URL 不能改、參數不能改。**
