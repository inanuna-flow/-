BEGIN;

CREATE SCHEMA IF NOT EXISTS kpl_dashboard;

CREATE OR REPLACE FUNCTION kpl_dashboard.set_updated_at()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$;

CREATE TABLE IF NOT EXISTS kpl_dashboard.import_batches (
  id BIGSERIAL PRIMARY KEY,
  source_type TEXT NOT NULL CHECK (source_type IN ('budget', 'freight', 'labor', 'picks')),
  source_file TEXT NOT NULL,
  period_month DATE CHECK (period_month IS NULL OR date_trunc('month', period_month)::date = period_month),
  imported_by TEXT,
  imported_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  note TEXT
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.warehouses (
  id BIGSERIAL PRIMARY KEY,
  warehouse_code TEXT NOT NULL UNIQUE,
  warehouse_name TEXT NOT NULL UNIQUE,
  region TEXT,
  is_active BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.budget_monthly (
  id BIGSERIAL PRIMARY KEY,
  month DATE NOT NULL CHECK (date_trunc('month', month)::date = month),
  warehouse_name TEXT NOT NULL,
  cost_type TEXT NOT NULL CHECK (cost_type IN ('labor', 'freight')),
  budget_amount NUMERIC(14,2) NOT NULL DEFAULT 0,
  source_file TEXT NOT NULL,
  import_batch_id BIGINT REFERENCES kpl_dashboard.import_batches(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (month, warehouse_name, cost_type)
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.freight_daily (
  id BIGSERIAL PRIMARY KEY,
  work_date DATE NOT NULL,
  month DATE NOT NULL CHECK (date_trunc('month', month)::date = month),
  warehouse_name TEXT NOT NULL,
  freight_amount NUMERIC(14,2) NOT NULL DEFAULT 0,
  source_file TEXT NOT NULL,
  import_batch_id BIGINT REFERENCES kpl_dashboard.import_batches(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (work_date, warehouse_name)
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.labor_daily (
  id BIGSERIAL PRIMARY KEY,
  work_date DATE NOT NULL,
  month DATE NOT NULL CHECK (date_trunc('month', month)::date = month),
  warehouse_name TEXT NOT NULL,
  vendor_name TEXT,
  shift_name TEXT NOT NULL DEFAULT '',
  employee_id TEXT NOT NULL DEFAULT '',
  department_name TEXT,
  operation_area TEXT NOT NULL,
  hours NUMERIC(10,2) NOT NULL DEFAULT 0,
  cost NUMERIC(14,2) NOT NULL DEFAULT 0,
  source_file TEXT NOT NULL,
  import_batch_id BIGINT REFERENCES kpl_dashboard.import_batches(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (work_date, warehouse_name, employee_id, operation_area, shift_name)
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.picks_daily (
  id BIGSERIAL PRIMARY KEY,
  work_date DATE NOT NULL,
  month DATE NOT NULL CHECK (date_trunc('month', month)::date = month),
  warehouse_name TEXT NOT NULL,
  business_type TEXT NOT NULL DEFAULT '',
  area_name TEXT NOT NULL DEFAULT '',
  operation_area TEXT NOT NULL,
  picks_count INTEGER NOT NULL DEFAULT 0,
  source_file TEXT NOT NULL,
  import_batch_id BIGINT REFERENCES kpl_dashboard.import_batches(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (work_date, warehouse_name, business_type, area_name, operation_area)
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.operation_area_map (
  id BIGSERIAL PRIMARY KEY,
  source_name TEXT NOT NULL UNIQUE,
  standard_name TEXT NOT NULL,
  warehouse_name TEXT,
  is_active BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.vendor_map (
  id BIGSERIAL PRIMARY KEY,
  source_name TEXT NOT NULL UNIQUE,
  standard_name TEXT NOT NULL,
  vendor_type TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS kpl_dashboard.shift_map (
  id BIGSERIAL PRIMARY KEY,
  source_name TEXT NOT NULL UNIQUE,
  standard_name TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_budget_monthly_month
  ON kpl_dashboard.budget_monthly(month);

CREATE INDEX IF NOT EXISTS idx_budget_monthly_warehouse
  ON kpl_dashboard.budget_monthly(warehouse_name);

CREATE INDEX IF NOT EXISTS idx_freight_daily_date
  ON kpl_dashboard.freight_daily(work_date);

CREATE INDEX IF NOT EXISTS idx_freight_daily_warehouse
  ON kpl_dashboard.freight_daily(warehouse_name);

CREATE INDEX IF NOT EXISTS idx_labor_daily_date
  ON kpl_dashboard.labor_daily(work_date);

CREATE INDEX IF NOT EXISTS idx_labor_daily_warehouse_area
  ON kpl_dashboard.labor_daily(warehouse_name, operation_area);

CREATE INDEX IF NOT EXISTS idx_picks_daily_date
  ON kpl_dashboard.picks_daily(work_date);

CREATE INDEX IF NOT EXISTS idx_picks_daily_warehouse_area
  ON kpl_dashboard.picks_daily(warehouse_name, operation_area);

DROP TRIGGER IF EXISTS set_updated_at_warehouses ON kpl_dashboard.warehouses;
CREATE TRIGGER set_updated_at_warehouses
BEFORE UPDATE ON kpl_dashboard.warehouses
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_budget_monthly ON kpl_dashboard.budget_monthly;
CREATE TRIGGER set_updated_at_budget_monthly
BEFORE UPDATE ON kpl_dashboard.budget_monthly
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_freight_daily ON kpl_dashboard.freight_daily;
CREATE TRIGGER set_updated_at_freight_daily
BEFORE UPDATE ON kpl_dashboard.freight_daily
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_labor_daily ON kpl_dashboard.labor_daily;
CREATE TRIGGER set_updated_at_labor_daily
BEFORE UPDATE ON kpl_dashboard.labor_daily
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_picks_daily ON kpl_dashboard.picks_daily;
CREATE TRIGGER set_updated_at_picks_daily
BEFORE UPDATE ON kpl_dashboard.picks_daily
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_operation_area_map ON kpl_dashboard.operation_area_map;
CREATE TRIGGER set_updated_at_operation_area_map
BEFORE UPDATE ON kpl_dashboard.operation_area_map
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_vendor_map ON kpl_dashboard.vendor_map;
CREATE TRIGGER set_updated_at_vendor_map
BEFORE UPDATE ON kpl_dashboard.vendor_map
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_shift_map ON kpl_dashboard.shift_map;
CREATE TRIGGER set_updated_at_shift_map
BEFORE UPDATE ON kpl_dashboard.shift_map
FOR EACH ROW EXECUTE FUNCTION kpl_dashboard.set_updated_at();

INSERT INTO kpl_dashboard.warehouses (warehouse_code, warehouse_name, region)
VALUES
  ('DX', U&'\5927\6EAA\5009', U&'\5317\5340'),
  ('DD', U&'\5927\809A\5009', U&'\4E2D\5340'),
  ('GS', U&'\5CA1\5C71\5009', U&'\5357\5340')
ON CONFLICT (warehouse_code) DO UPDATE
SET warehouse_name = EXCLUDED.warehouse_name,
    region = EXCLUDED.region,
    is_active = true;

INSERT INTO kpl_dashboard.shift_map (source_name, standard_name)
VALUES
  (U&'\65E5', U&'\65E5'),
  (U&'\65E5\73ED', U&'\65E5'),
  (U&'\65E9\73ED', U&'\65E5'),
  (U&'\4E2D', U&'\4E2D'),
  (U&'\4E2D\73ED', U&'\4E2D'),
  (U&'\591C', U&'\591C'),
  (U&'\591C\73ED', U&'\591C')
ON CONFLICT (source_name) DO UPDATE
SET standard_name = EXCLUDED.standard_name;

COMMIT;
