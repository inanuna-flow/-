# CLAUDE.md

## Project Context

This repository is the KPL logistics performance dashboard.

- Local root: this repository folder on the user's desktop/OneDrive
- Cloud Run project: `l000-493901`
- Cloud Run service: `kpl-dashboard`
- Cloud Run region: `asia-east1`
- Public URL: `https://kpl-dashboard-146561234228.asia-east1.run.app/login.html`
- Git remote: `https://github.com/inanuna-flow/-.git`
- PostgreSQL schema name: `kpl_dashboard`
- Do not use the PostgreSQL `public` schema for KPL business tables.

## Coordination Rules

The user may use both Claude Code Desktop and Codex/VS Code against the same folder.

Before editing, committing, pushing, or deploying:

```powershell
git status --short
git fetch origin
git log --oneline --decorate -5
```

If another agent just worked on the repo, synchronize first:

```powershell
git pull --ff-only
```

Only one agent should be the active writer/deployer at a time. If another agent is deploying Cloud Run, wait until it finishes before making new deploys. The last Cloud Run deploy wins, even if GitHub contains a different version.

Never add unrelated untracked files automatically. In this repo, `.claude/`, `.vscode/`, and root `README.md` may be local/user files unless explicitly requested.

## GCP Notes

PowerShell may block `gcloud.ps1` because script execution is disabled. Use `gcloud.cmd` instead:

```powershell
gcloud.cmd --version
gcloud.cmd auth list
gcloud.cmd config list
```

The active account should usually be:

```text
inari@fme.com.tw
```

The project should be:

```powershell
gcloud.cmd config set project l000-493901
```

The company network may cause certificate errors such as:

```text
CERTIFICATE_VERIFY_FAILED
self-signed certificate in certificate chain
Missing Authority Key Identifier
```

The more correct fix is to use the company CA file:

```text
C:/Users/inari/fme_ca_certs.pem
```

If urgent and the user explicitly approves, temporarily bypass SSL validation:

```powershell
gcloud.cmd config set auth/disable_ssl_validation true
```

Always restore it immediately after the GCP operation:

```powershell
gcloud.cmd config unset auth/disable_ssl_validation
gcloud.cmd config get-value auth/disable_ssl_validation
```

Expected restored value:

```text
(unset)
```

Do not leave SSL validation disabled.

## Cloud Run Deployment

This project now includes:

- `Dockerfile`
- `.dockerignore`
- `.gcloudignore`
- `server.js`

Deploy command:

```powershell
gcloud.cmd run deploy kpl-dashboard --source . --project l000-493901 --region asia-east1 --allow-unauthenticated --quiet
```

Verify deployed revision:

```powershell
gcloud.cmd run services describe kpl-dashboard --project l000-493901 --region asia-east1 --format "value(status.latestReadyRevisionName,status.url)"
```

Known service account:

```text
146561234228-compute@developer.gserviceaccount.com
```

If the deploy succeeds but IAM policy setting warns, the service can still be serving traffic. Verify by opening:

```text
https://kpl-dashboard-146561234228.asia-east1.run.app/login.html
```

## Cloud SQL Notes

Known Cloud SQL information from IT/user:

```text
Connection Name: l000-493901:asia-east1:l000-db
Instance Name: l000-db
Database Name: L000-DB
DB User: L000LDB
Database Version: POSTGRES_18
Region: asia-east1
State: RUNNABLE
```

The DB password was once shared in chat/screenshot. Treat it as compromised and recommend asking IT to rotate it after setup is confirmed. Do not write DB passwords into files, commits, or logs.

The user account may be able to describe the instance but not list databases/users:

```text
gcloud sql databases list -> 403
gcloud sql users list -> 403
```

This can mean the user has Cloud SQL Client but not Cloud SQL Viewer/Admin.

Local machine previously did not have `psql` in PATH:

```text
Psql client not found.
```

If `psql` is still missing, either install PostgreSQL client tools, use Cloud Shell, or ask IT to run:

```text
database/init_database.sql
```

The database initialization script is:

```text
database/init_database.sql
```

It creates the `kpl_dashboard` schema and first-phase tables. It intentionally does not place KPL tables in `public`.

First-phase confirmed data domains:

- annual budget: `budget_monthly`
- labor cost: `labor_daily`
- picks: `picks_daily`
- import tracking: `import_batches`
- references: `warehouses`, `operation_area_map`, `vendor_map`, `shift_map`

Freight data is not fully confirmed yet. `freight_daily` exists in the schema, but formal import should wait until field definitions are confirmed or use staging first.

## Git History To Remember

Important recent commits:

```text
0297df0 Add IT database initialization script
0b3b266 Add Cloud Run deployment config
0c1ac58 Add Cloud SQL PostgreSQL schema plan
```

If local state differs from GitHub, inspect before changing:

```powershell
git status --short
git log --oneline --decorate origin/main..HEAD
git log --oneline --decorate HEAD..origin/main
```
稱呼我為【少佐】
CLAUDE CODE 自稱[タチコマ]