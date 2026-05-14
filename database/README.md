# KPL Dashboard Cloud SQL PostgreSQL

This folder contains the first-phase database implementation for the KPL dashboard.
It creates the PostgreSQL schema only. It does not change the Cloud Run app, frontend,
Excel import flow, or authentication flow.

## Target

- Database: GCP Cloud SQL for PostgreSQL
- Schema: `kpl_dashboard`
- Do not create KPL tables in `public`
- Re-import policy: use `INSERT ... ON CONFLICT DO UPDATE` for month/date keyed data

## Files

- `migrations/001_create_kpl_dashboard_schema.sql`
  - Creates `kpl_dashboard`
  - Creates all first-phase tables
  - Creates indexes, unique keys, update timestamp triggers, and default warehouse/shift mappings
- `tests/verify_kpl_dashboard_schema.sql`
  - Inserts sample 2026-03 data in a transaction
  - Verifies upsert behavior
  - Rolls back at the end, leaving no test data behind

## Apply Migration

Run this against the target Cloud SQL PostgreSQL database:

```powershell
psql "$env:DATABASE_URL" -f database/migrations/001_create_kpl_dashboard_schema.sql
```

Or with explicit connection parameters:

```powershell
psql `
  --host "$env:PGHOST" `
  --port "$env:PGPORT" `
  --username "$env:PGUSER" `
  --dbname "$env:PGDATABASE" `
  -f database/migrations/001_create_kpl_dashboard_schema.sql
```

## Verify

```powershell
psql "$env:DATABASE_URL" -f database/tests/verify_kpl_dashboard_schema.sql
```

Expected verification signals:

- `schema_count = 1`
- `public_kpl_table_count = 0`
- `freight_upsert_amount = 12500.00`
- `labor_upsert_cost = 1680.00`
- `picks_upsert_count = 5200`

## Table Model

Core fact tables:

- `budget_monthly`: monthly labor/freight budget by warehouse
- `freight_daily`: daily freight amount by warehouse
- `labor_daily`: daily labor detail by warehouse, employee, area, and shift
- `picks_daily`: daily pick count by warehouse and operation area

Reference tables:

- `warehouses`
- `operation_area_map`
- `vendor_map`
- `shift_map`

Import tracking:

- `import_batches`

## Notes For The Next Phase

The frontend currently reads in-memory `DATA`, `LABOR_RAW`, and `PICKS_RAW`.
The next implementation phase should add API endpoints that query this schema and
return shapes compatible with the existing dashboard render functions before moving
Excel parsing server-side.
