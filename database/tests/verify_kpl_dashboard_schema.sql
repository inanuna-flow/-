BEGIN;

INSERT INTO kpl_dashboard.import_batches (source_type, source_file, period_month, imported_by, note)
VALUES ('budget', 'annual_budget_v2.xlsx', '2026-03-01', 'schema-test', 'verify budget upsert')
RETURNING id;

INSERT INTO kpl_dashboard.budget_monthly
  (month, warehouse_name, cost_type, budget_amount, source_file)
VALUES
  ('2026-03-01', U&'\5927\6EAA\5009', 'labor', 1000000, 'annual_budget_v2.xlsx'),
  ('2026-03-01', U&'\5927\6EAA\5009', 'freight', 300000, 'annual_budget_v2.xlsx'),
  ('2026-03-01', U&'\5927\809A\5009', 'labor', 800000, 'annual_budget_v2.xlsx'),
  ('2026-03-01', U&'\5CA1\5C71\5009', 'freight', 250000, 'annual_budget_v2.xlsx')
ON CONFLICT (month, warehouse_name, cost_type) DO UPDATE
SET budget_amount = EXCLUDED.budget_amount,
    source_file = EXCLUDED.source_file;

INSERT INTO kpl_dashboard.freight_daily
  (work_date, month, warehouse_name, freight_amount, source_file)
VALUES
  ('2026-03-01', '2026-03-01', U&'\5927\6EAA\5009', 12000, 'freight.xlsx'),
  ('2026-03-01', '2026-03-01', U&'\5927\6EAA\5009', 12500, 'freight_revised.xlsx')
ON CONFLICT (work_date, warehouse_name) DO UPDATE
SET freight_amount = EXCLUDED.freight_amount,
    source_file = EXCLUDED.source_file;

INSERT INTO kpl_dashboard.labor_daily
  (work_date, month, warehouse_name, vendor_name, shift_name, employee_id, department_name, operation_area, hours, cost, source_file)
VALUES
  ('2026-03-01', '2026-03-01', U&'\5927\6EAA\5009', 'vendor_a', U&'\65E5', 'E001', 'EC', 'EC', 8, 1600, 'labor.xlsx'),
  ('2026-03-01', '2026-03-01', U&'\5927\6EAA\5009', 'vendor_a', U&'\65E5', 'E001', 'EC', 'EC', 8, 1680, 'labor_revised.xlsx')
ON CONFLICT (work_date, warehouse_name, employee_id, operation_area, shift_name) DO UPDATE
SET vendor_name = EXCLUDED.vendor_name,
    department_name = EXCLUDED.department_name,
    hours = EXCLUDED.hours,
    cost = EXCLUDED.cost,
    source_file = EXCLUDED.source_file;

INSERT INTO kpl_dashboard.picks_daily
  (work_date, month, warehouse_name, business_type, area_name, operation_area, picks_count, source_file)
VALUES
  ('2026-03-01', '2026-03-01', U&'\5927\6EAA\5009', 'EC', 'A', 'EC', 5000, 'picks.xlsx'),
  ('2026-03-01', '2026-03-01', U&'\5927\6EAA\5009', 'EC', 'A', 'EC', 5200, 'picks_revised.xlsx')
ON CONFLICT (work_date, warehouse_name, business_type, area_name, operation_area) DO UPDATE
SET picks_count = EXCLUDED.picks_count,
    source_file = EXCLUDED.source_file;

SELECT 'schema_count' AS check_name, count(*) AS value
FROM information_schema.schemata
WHERE schema_name = 'kpl_dashboard';

SELECT 'public_kpl_table_count' AS check_name, count(*) AS value
FROM information_schema.tables
WHERE table_schema = 'public'
  AND table_name IN (
    'import_batches',
    'warehouses',
    'budget_monthly',
    'freight_daily',
    'labor_daily',
    'picks_daily',
    'operation_area_map',
    'vendor_map',
    'shift_map'
  );

SELECT 'freight_upsert_amount' AS check_name, freight_amount AS value
FROM kpl_dashboard.freight_daily
WHERE work_date = '2026-03-01'
  AND warehouse_name = U&'\5927\6EAA\5009';

SELECT 'labor_upsert_cost' AS check_name, cost AS value
FROM kpl_dashboard.labor_daily
WHERE work_date = '2026-03-01'
  AND warehouse_name = U&'\5927\6EAA\5009'
  AND employee_id = 'E001';

SELECT 'picks_upsert_count' AS check_name, picks_count AS value
FROM kpl_dashboard.picks_daily
WHERE work_date = '2026-03-01'
  AND warehouse_name = U&'\5927\6EAA\5009'
  AND operation_area = 'EC';

ROLLBACK;
