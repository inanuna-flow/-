-- ════════════════════════════════════════════════════════════════
-- Migration 002: 拆分 freight_daily → freight_mainline_daily + freight_non_mainline_daily
-- ════════════════════════════════════════════════════════════════
--
-- 背景：原 freight_daily 僅一個 freight_amount 欄位，無法承載：
--   1. 主線運費 47 欄完整資訊（預計/到點/實際三組計價 + 行政欄位）
--   2. 非主線運費 11 種子分類（加派/其他/轉運下的細項）
--
-- 變更：
--   - DROP freight_daily（及其 trigger / index）
--   - 新建 freight_mainline_daily（47 欄）
--   - 新建 freight_non_mainline_daily（14 欄 + category_l1/l2 預留）
--   - 更新 import_batches.source_type CHECK
--
-- 注意：使用者要求砍掉舊資料重灌（Q5），所以直接 DROP，不做資料遷移。
-- ════════════════════════════════════════════════════════════════

BEGIN;

-- ─── 1. DROP 舊 freight_daily ───────────────────────────────────
DROP TRIGGER IF EXISTS set_updated_at_freight_daily ON kpl_dashboard.freight_daily;
DROP INDEX IF EXISTS kpl_dashboard.idx_freight_daily_date;
DROP INDEX IF EXISTS kpl_dashboard.idx_freight_daily_warehouse;
DROP TABLE IF EXISTS kpl_dashboard.freight_daily;

-- ─── 2. 更新 import_batches CHECK 允許新 source_type ────────────
ALTER TABLE kpl_dashboard.import_batches
  DROP CONSTRAINT IF EXISTS import_batches_source_type_check;

ALTER TABLE kpl_dashboard.import_batches
  ADD CONSTRAINT import_batches_source_type_check
  CHECK (source_type IN (
    'budget',
    'freight',                 -- 保留向下相容（舊 batch 紀錄不刪）
    'freight_mainline',
    'freight_non_mainline',
    'labor',
    'picks'
  ));

-- ─── 3. 主線運費表（freight_mainline_daily）──────────────────────
-- 對應 Excel：貨運費用明細總表匯出（47 欄全保留）
CREATE TABLE IF NOT EXISTS kpl_dashboard.freight_mainline_daily (
  id                       BIGSERIAL PRIMARY KEY,

  -- 基本資訊（9 欄）
  warehouse_name           TEXT NOT NULL,                  -- 倉別
  status_locked            TEXT,                           -- 明細狀態（未鎖定/鎖定）
  carry_month              TEXT,                           -- 結轉年月
  work_date                DATE NOT NULL,                  -- 進貨日（民國年→西元年）
  month                    DATE NOT NULL CHECK (date_trunc('month', month)::date = month),
  shipper                  TEXT,                           -- 配別（日翊）
  route_source             TEXT,                           -- 路線來源（主線）
  -- route_code / carrier 設為 NOT NULL DEFAULT ''，避免 NULL 在 UNIQUE 比對時被 PostgreSQL 視為「不衝突」而插入重複列
  route_code               TEXT NOT NULL DEFAULT '',       -- 路線（如 1A/1B/1C；空字串代表未填）
  task_category            TEXT,                           -- 任務分類（主線計價）
  carrier                  TEXT NOT NULL DEFAULT '',       -- 配送商（創新/...；空字串代表未填）

  -- 路線狀態（2 欄）
  route_status             TEXT,                           -- 路線完成狀態（已返廠/...）
  task_note                TEXT,                           -- 任務備註

  -- 預計計價（11 欄）
  est_pricing_type         TEXT,                           -- 預計計價類型
  est_price_list           TEXT,                           -- 預計價目表
  est_trip_type            TEXT,                           -- 預計全/半趟
  est_vehicle_tonnage      TEXT,                           -- 合約車輛噸型
  est_pricing_qty          NUMERIC(14,4),                  -- 預計計價數
  est_pricing_tier         TEXT,                           -- 預計計價級距
  est_pricing_unit         TEXT,                           -- 預計計價單位
  est_tier_rate            NUMERIC(14,4),                  -- 預計級距費率
  est_dispatch_price       NUMERIC(14,4),                  -- 預計出車價
  est_pricing_result       NUMERIC(14,4),                  -- 預計計價結果
  est_with_tax             TEXT,                           -- 預計是否含稅

  -- 到點計價（7 欄）
  arr_vehicle_tonnage      TEXT,                           -- 到點車輛噸型
  arr_pricing_qty          NUMERIC(14,4),                  -- 到點計價數
  arr_pricing_tier         TEXT,                           -- 到點計價級距
  arr_pricing_unit         TEXT,                           -- 到點計價單位
  arr_tier_rate            NUMERIC(14,4),                  -- 到點級距費率
  arr_dispatch_price       NUMERIC(14,4),                  -- 到點車輛噸型出車價
  arr_pricing_result       NUMERIC(14,4),                  -- 到點計價結果

  -- 實際計價（11 欄，pricing_result 是 ★ 最終金額）
  pricing_type             TEXT,                           -- 計價類型
  price_list               TEXT,                           -- 計價價目表
  trip_type                TEXT,                           -- 全/半趟
  vehicle_tonnage          TEXT,                           -- 車輛噸型
  pricing_qty              NUMERIC(14,4),                  -- 計價數
  pricing_tier             TEXT,                           -- 計價級距
  pricing_unit             TEXT,                           -- 計價單位
  tier_rate                NUMERIC(14,4),                  -- 級距費率
  dispatch_price           NUMERIC(14,4),                  -- 出車價
  pricing_result           NUMERIC(14,2) NOT NULL DEFAULT 0,  -- ★ 計價結果（實際金額）
  with_tax                 TEXT,                           -- 是否含稅

  -- 行政欄位（7 欄，追溯用）
  adjust_reason            TEXT,                           -- 調整原因
  system_updated_at_raw    TEXT,                           -- 系統最後更新時間（原字串）
  est_created_by           TEXT,                           -- 預計計價建立人員
  est_created_at_raw       TEXT,                           -- 預計計價建立時間
  actual_updated_by        TEXT,                           -- 實際/調整後計價更新人員
  actual_updated_at_raw    TEXT,                           -- 實際/調整後計價更新時間
  status_updated_by        TEXT,                           -- 狀態更新人員
  status_updated_at_raw    TEXT,                           -- 狀態更新時間
  carry_by                 TEXT,                           -- 結轉人員
  carry_at_raw             TEXT,                           -- 結轉時間

  -- 追溯
  source_file              TEXT NOT NULL,
  import_batch_id          BIGINT REFERENCES kpl_dashboard.import_batches(id),
  created_at               TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at               TIMESTAMPTZ NOT NULL DEFAULT now(),

  -- 去重（策略 B：同 day + 倉 + 路線 + 配送商 視為同一筆）
  UNIQUE (work_date, warehouse_name, route_code, carrier)
);

CREATE INDEX IF NOT EXISTS idx_freight_mainline_date
  ON kpl_dashboard.freight_mainline_daily(work_date);

CREATE INDEX IF NOT EXISTS idx_freight_mainline_warehouse
  ON kpl_dashboard.freight_mainline_daily(warehouse_name);

CREATE INDEX IF NOT EXISTS idx_freight_mainline_month
  ON kpl_dashboard.freight_mainline_daily(month);

-- 防呆：若先前已建表（含 NULLable route_code/carrier），把 NULL 補成空字串再強制 NOT NULL
-- 這樣即使表已存在、CREATE TABLE IF NOT EXISTS 跳過，UNIQUE 仍然能正確比對
UPDATE kpl_dashboard.freight_mainline_daily
  SET route_code = COALESCE(route_code, ''),
      carrier    = COALESCE(carrier, '');

ALTER TABLE kpl_dashboard.freight_mainline_daily
  ALTER COLUMN route_code SET DEFAULT '',
  ALTER COLUMN route_code SET NOT NULL,
  ALTER COLUMN carrier    SET DEFAULT '',
  ALTER COLUMN carrier    SET NOT NULL;

DROP TRIGGER IF EXISTS set_updated_at_freight_mainline_daily ON kpl_dashboard.freight_mainline_daily;
CREATE TRIGGER set_updated_at_freight_mainline_daily
BEFORE UPDATE ON kpl_dashboard.freight_mainline_daily
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

COMMENT ON TABLE kpl_dashboard.freight_mainline_daily IS
  '主線運費明細（每車一筆，47 欄完整保留）';

-- ─── 4. 非主線運費表（freight_non_mainline_daily）───────────────
-- 對應 Excel：加派車維護基本檔（14 欄 + 衍生分類 2 欄）
CREATE TABLE IF NOT EXISTS kpl_dashboard.freight_non_mainline_daily (
  id                       BIGSERIAL PRIMARY KEY,

  -- Excel 原始欄位（14 欄，項次不存）
  work_date                DATE NOT NULL,                  -- 進貨日
  month                    DATE NOT NULL CHECK (date_trunc('month', month)::date = month),
  shipper                  TEXT,                           -- 配別（日翊）
  region                   TEXT,                           -- 區域（北/中/南）
  vehicle_type             TEXT,                           -- 車型（3.5噸14呎;160箱）
  delivery_item            TEXT,                           -- 配送項目（區配）
  warehouse_name           TEXT NOT NULL,                  -- 廠商欄位 = 倉別
  carrier                  TEXT,                           -- 配送商（華碩/銓程/...）
  dispatch_reason          TEXT,                           -- 派車原因 ★ 分類依據
  pricing_method           TEXT,                           -- 計價方式 ★ 分類依據
  unit_price               NUMERIC(14,2) NOT NULL DEFAULT 0,
  trip_count               INTEGER NOT NULL DEFAULT 0,
  amount                   NUMERIC(14,2) NOT NULL DEFAULT 0,  -- 費用小計
  note                     TEXT,                           -- 備註 ★ 分類依據

  -- 衍生分類（規則由 server.js classifyNonMainline() 計算寫入）
  -- 規則由業務同仁確認，2026-06-08 啟用；序 1~12（命中即停）
  category_l1              TEXT,                           -- 加派/其他/轉運/不列入/無法判斷
  category_l2              TEXT,                           -- 正物流/逆物流/專車/違規罰款/誤key/上收/...
  budget_warehouse         TEXT,                           -- 預算歸屬倉（跨區轉運費統一歸「大溪倉」；其餘 = warehouse_name）

  -- 追溯
  source_file              TEXT NOT NULL,
  import_batch_id          BIGINT REFERENCES kpl_dashboard.import_batches(id),
  created_at               TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at               TIMESTAMPTZ NOT NULL DEFAULT now()
  -- 非主線無天然唯一鍵（同 key 可能合法多筆），不設 UNIQUE
);

CREATE INDEX IF NOT EXISTS idx_freight_nonmain_date
  ON kpl_dashboard.freight_non_mainline_daily(work_date);

CREATE INDEX IF NOT EXISTS idx_freight_nonmain_warehouse
  ON kpl_dashboard.freight_non_mainline_daily(warehouse_name);

CREATE INDEX IF NOT EXISTS idx_freight_nonmain_month
  ON kpl_dashboard.freight_non_mainline_daily(month);

CREATE INDEX IF NOT EXISTS idx_freight_nonmain_category
  ON kpl_dashboard.freight_non_mainline_daily(category_l1, category_l2);

DROP TRIGGER IF EXISTS set_updated_at_freight_non_mainline_daily ON kpl_dashboard.freight_non_mainline_daily;
CREATE TRIGGER set_updated_at_freight_non_mainline_daily
BEFORE UPDATE ON kpl_dashboard.freight_non_mainline_daily
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

COMMENT ON TABLE kpl_dashboard.freight_non_mainline_daily IS
  '非主線運費明細（每車一筆，含派車原因/計價方式/備註用於分類；分類規則待確認）';

COMMIT;
