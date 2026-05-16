const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { Pool } = require('pg');

const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const STATIC_ROOT = path.join(ROOT, 'kpl-dashboard');
const EIP_CHECK_USER_URL = 'https://eip.fme.com.tw/FMEIP/AasApi/CheckUserId';
const PERMISSIONS_FILE = path.join(ROOT, 'page_permissions.json');
const ADMIN_USER_ID = (process.env.ADMIN_USER_ID || 'inari').toLowerCase();
const DB_SCHEMA = process.env.DB_SCHEMA || 'kpl_dashboard';
const DB_INSTANCE_CONNECTION_NAME = process.env.DB_INSTANCE_CONNECTION_NAME || '';
const DB_NAME = process.env.DB_NAME || process.env.PGDATABASE || '';
const DB_USER = process.env.DB_USER || process.env.PGUSER || '';
const DB_PASSWORD = process.env.DB_PASSWORD || process.env.PGPASSWORD || '';
const DB_HOST = process.env.DB_HOST || (DB_INSTANCE_CONNECTION_NAME ? `/cloudsql/${DB_INSTANCE_CONNECTION_NAME}` : '');

const VALID_PAGE_IDS = new Set(['daily','dispatch','freight','labor','productivity','monthly','annual','import','org']);
const BUDGET_TYPES = new Set(['labor', 'freight']);

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
};

const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
};

function sendJson(res, status, payload) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    ...SECURITY_HEADERS,
  });
  res.end(JSON.stringify(payload));
}

function readRequestBody(req, maxBytes = 4096) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.setEncoding('utf8');
    req.on('data', chunk => {
      body += chunk;
      if (body.length > maxBytes) {
        reject(new Error('REQUEST_TOO_LARGE'));
        req.destroy();
      }
    });
    req.on('end', () => resolve(body));
    req.on('error', reject);
  });
}

function extractSupportId(text) {
  return text.match(/support ID is:\s*([^<\s]+)/i)?.[1] || '';
}

let dbPool = null;

function isDbConfigured() {
  return Boolean(DB_HOST && DB_NAME && DB_USER && DB_PASSWORD);
}

function getDbPool() {
  if (!isDbConfigured()) return null;
  if (!dbPool) {
    dbPool = new Pool({
      host: DB_HOST,
      database: DB_NAME,
      user: DB_USER,
      password: DB_PASSWORD,
      max: 5,
      idleTimeoutMillis: 30000,
      connectionTimeoutMillis: 10000,
    });
  }
  return dbPool;
}

function sqlIdent(name) {
  return `"${String(name).replace(/"/g, '""')}"`;
}

function schemaTable(table) {
  return `${sqlIdent(DB_SCHEMA)}.${sqlIdent(table)}`;
}

function monthDate(year, monthIndex) {
  const y = Number(year);
  const m = Number(monthIndex) + 1;
  if (!Number.isInteger(y) || y < 2000 || y > 2100) return '';
  if (!Number.isInteger(m) || m < 1 || m > 12) return '';
  return `${y}-${String(m).padStart(2, '0')}-01`;
}

function budgetRowsFromPayload(payload) {
  const year = Number(payload.year);
  const rows = [];
  for (const costType of ['labor', 'freight']) {
    const byWarehouse = payload[costType];
    if (!byWarehouse || typeof byWarehouse !== 'object') continue;
    for (const [warehouseName, months] of Object.entries(byWarehouse)) {
      if (!warehouseName || !Array.isArray(months)) continue;
      months.forEach((amount, monthIndex) => {
        const month = monthDate(year, monthIndex);
        const budgetAmount = Number(amount) || 0;
        if (month) {
          rows.push({
            month,
            warehouseName,
            costType,
            budgetAmount,
          });
        }
      });
    }
  }
  return rows;
}

function isIsoDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ''));
}

function monthStartFromDate(value) {
  return isIsoDate(value) ? `${String(value).slice(0, 7)}-01` : '';
}

function laborRowsFromPayload(payload) {
  const records = Array.isArray(payload.records) ? payload.records : [];
  const rows = records
    .map(record => ({
      workDate: String(record.date || ''),
      month: monthStartFromDate(record.date),
      warehouseName: String(record.wh || '').trim(),
      vendorName: String(record.vendor || '').trim(),
      shiftName: String(record.shift || '').trim(),
      employeeId: String(record.empId || '').trim(),
      departmentName: String(record.dept || '').trim(),
      operationArea: String(record.opArea || '').trim(),
      hours: Number(record.hours) || 0,
      cost: Number(record.cost) || 0,
    }))
    .filter(row =>
      isIsoDate(row.workDate) &&
      row.month &&
      row.warehouseName &&
      row.operationArea &&
      (row.hours !== 0 || row.cost !== 0)
    );
  const grouped = new Map();
  rows.forEach(row => {
    const key = [
      row.workDate,
      row.warehouseName,
      row.employeeId,
      row.operationArea,
      row.shiftName,
    ].join('\u001f');
    const existing = grouped.get(key);
    if (existing) {
      existing.hours += row.hours;
      existing.cost += row.cost;
      if (!existing.vendorName && row.vendorName) existing.vendorName = row.vendorName;
      if (!existing.departmentName && row.departmentName) existing.departmentName = row.departmentName;
    } else {
      grouped.set(key, { ...row });
    }
  });
  return Array.from(grouped.values()).map(row => ({
    ...row,
    hours: Math.round(row.hours * 100) / 100,
    cost: Math.round(row.cost * 100) / 100,
  }));
}

async function handleCheckUser(req, res) {
  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  let credentials;
  try {
    const body = await readRequestBody(req);
    credentials = JSON.parse(body || '{}');
  } catch (err) {
    sendJson(res, 400, { ok: false, MSG: '999 登入資料格式錯誤' });
    return;
  }

  const userId = String(credentials.USER_ID || '').trim();
  const password = String(credentials.PSW || '');
  if (!userId || !password) {
    sendJson(res, 400, { ok: false, MSG: '999 請輸入帳號與密碼' });
    return;
  }

  try {
    const upstream = await fetch(EIP_CHECK_USER_URL, {
      method: 'POST',
      headers: {
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json',
        'Origin': 'https://eip.fme.com.tw',
        'Referer': 'https://eip.fme.com.tw/FMEIP/',
        'User-Agent': 'Mozilla/5.0 KPL-Dashboard-Auth-Proxy',
      },
      body: JSON.stringify({ USER_ID: userId, PSW: password }),
    });

    const text = await upstream.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      const supportId = extractSupportId(text);
      sendJson(res, 502, {
        ok: false,
        MSG: supportId
          ? `999 日翊 EIP 拒絕驗證請求，Support ID: ${supportId}`
          : `999 日翊 EIP 回傳非 JSON 內容，HTTP ${upstream.status}`,
      });
      return;
    }

    sendJson(res, upstream.ok ? 200 : 502, data);
  } catch (err) {
    sendJson(res, 502, {
      ok: false,
      MSG: `999 無法連線至日翊 EIP：${err.message}`,
    });
  }
}

function handleStatic(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const pathname = decodeURIComponent(url.pathname);
  let relativePath = pathname === '/' ? 'login.html' : pathname.slice(1);
  if (relativePath.startsWith('kpl-dashboard/')) {
    relativePath = relativePath.slice('kpl-dashboard/'.length);
  }
  const filePath = path.resolve(STATIC_ROOT, relativePath);

  if (!filePath.startsWith(STATIC_ROOT + path.sep) && filePath !== STATIC_ROOT) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
      'Cache-Control': 'no-store',
      ...SECURITY_HEADERS,
    });
    res.end(data);
  });
}

// ── 讀取頁面權限 ──
function handleGetPermissions(req, res) {
  fs.readFile(PERMISSIONS_FILE, 'utf8', (err, data) => {
    if (err) {
      // 若檔案不存在，回傳全開預設值
      const defaults = {};
      VALID_PAGE_IDS.forEach(id => { defaults[id] = true; });
      sendJson(res, 200, defaults);
      return;
    }
    try {
      sendJson(res, 200, JSON.parse(data));
    } catch {
      sendJson(res, 500, { ok: false, MSG: '999 權限檔案格式錯誤' });
    }
  });
}

// ── 儲存頁面權限（需 inari 重新驗證）──
async function handleSavePermissions(req, res) {
  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  let body;
  try {
    const raw = await readRequestBody(req);
    body = JSON.parse(raw || '{}');
  } catch {
    sendJson(res, 400, { ok: false, MSG: '999 資料格式錯誤' });
    return;
  }

  const userId   = String(body.USER_ID || '').trim().toLowerCase();
  const password = String(body.PSW || '');
  const perms    = body.permissions;

  if (!userId || !password) {
    sendJson(res, 400, { ok: false, MSG: '999 請提供帳號與密碼' });
    return;
  }
  if (userId !== ADMIN_USER_ID) {
    sendJson(res, 403, { ok: false, MSG: '403 僅限管理員操作' });
    return;
  }
  if (!perms || typeof perms !== 'object') {
    sendJson(res, 400, { ok: false, MSG: '999 權限資料格式錯誤' });
    return;
  }

  // 重新驗證身份
  try {
    const upstream = await fetch(EIP_CHECK_USER_URL, {
      method: 'POST',
      headers: {
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json',
        'Origin': 'https://eip.fme.com.tw',
        'Referer': 'https://eip.fme.com.tw/FMEIP/',
        'User-Agent': 'Mozilla/5.0 KPL-Dashboard-Auth-Proxy',
      },
      body: JSON.stringify({ USER_ID: body.USER_ID, PSW: password }),
    });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); } catch { data = {}; }
    if (!String(data.MSG || '').startsWith('000')) {
      sendJson(res, 401, { ok: false, MSG: '401 密碼驗證失敗，權限未儲存' });
      return;
    }
  } catch (err) {
    sendJson(res, 502, { ok: false, MSG: `999 無法連線至 EIP：${err.message}` });
    return;
  }

  // 過濾只保留合法的 page id，值只允許 boolean
  const sanitized = {};
  VALID_PAGE_IDS.forEach(id => {
    sanitized[id] = perms[id] !== false;
  });

  fs.writeFile(PERMISSIONS_FILE, JSON.stringify(sanitized, null, 2), 'utf8', err => {
    if (err) {
      sendJson(res, 500, { ok: false, MSG: '999 儲存失敗' });
      return;
    }
    sendJson(res, 200, { ok: true, MSG: '000 權限已儲存', permissions: sanitized });
  });
}

async function handleBudgetImport(req, res) {
  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  const pool = getDbPool();
  if (!pool) {
    sendJson(res, 503, {
      ok: false,
      MSG: '503 Database is not configured. Set DB_INSTANCE_CONNECTION_NAME, DB_NAME, DB_USER, and DB_PASSWORD.',
    });
    return;
  }

  let payload;
  try {
    const raw = await readRequestBody(req, 1024 * 1024);
    payload = JSON.parse(raw || '{}');
  } catch (err) {
    sendJson(res, err.message === 'REQUEST_TOO_LARGE' ? 413 : 400, {
      ok: false,
      MSG: '999 Invalid budget import payload',
    });
    return;
  }

  const rows = budgetRowsFromPayload(payload);
  if (!rows.length) {
    sendJson(res, 400, { ok: false, MSG: '999 No valid budget rows to import' });
    return;
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const batchResult = await client.query(
      `INSERT INTO ${schemaTable('import_batches')}
        (source_type, source_file, period_month, imported_by, note)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING id`,
      [
        'budget',
        String(payload.fileName || 'budget.xlsx'),
        monthDate(Number(payload.year), Number(payload.monthIndex || 0)),
        String(payload.importedBy || ''),
        String(payload.version || ''),
      ],
    );
    const batchId = batchResult.rows[0].id;

    const upsertSql = `
      INSERT INTO ${schemaTable('budget_monthly')}
        (month, warehouse_name, cost_type, budget_amount, source_file, import_batch_id)
      VALUES ($1, $2, $3, $4, $5, $6)
      ON CONFLICT (month, warehouse_name, cost_type) DO UPDATE
      SET budget_amount = EXCLUDED.budget_amount,
          source_file = EXCLUDED.source_file,
          import_batch_id = EXCLUDED.import_batch_id
    `;

    for (const row of rows) {
      if (!BUDGET_TYPES.has(row.costType)) continue;
      await client.query(upsertSql, [
        row.month,
        row.warehouseName,
        row.costType,
        row.budgetAmount,
        String(payload.fileName || 'budget.xlsx'),
        batchId,
      ]);
    }

    await client.query('COMMIT');
    sendJson(res, 200, {
      ok: true,
      MSG: '000 Budget imported',
      batchId,
      rows: rows.length,
    });
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    sendJson(res, 500, { ok: false, MSG: `999 Budget import failed: ${err.message}` });
  } finally {
    client.release();
  }
}

async function handleBudgetData(req, res) {
  if (req.method !== 'GET') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  const pool = getDbPool();
  if (!pool) {
    sendJson(res, 503, {
      ok: false,
      MSG: '503 Database is not configured. Set DB_INSTANCE_CONNECTION_NAME, DB_NAME, DB_USER, and DB_PASSWORD.',
    });
    return;
  }

  const url = new URL(req.url, `http://${req.headers.host}`);
  const year = Number(url.searchParams.get('year') || new Date().getFullYear());
  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    sendJson(res, 400, { ok: false, MSG: '999 Invalid year' });
    return;
  }

  try {
    const result = await pool.query(
      `SELECT month, warehouse_name, cost_type, budget_amount, source_file, import_batch_id, updated_at
       FROM ${schemaTable('budget_monthly')}
       WHERE month >= $1::date AND month < ($1::date + interval '1 year')
       ORDER BY month, warehouse_name, cost_type`,
      [`${year}-01-01`],
    );

    sendJson(res, 200, {
      ok: true,
      year,
      rows: result.rows.map(row => ({
        month: row.month,
        warehouseName: row.warehouse_name,
        costType: row.cost_type,
        budgetAmount: Number(row.budget_amount) || 0,
        sourceFile: row.source_file,
        importBatchId: row.import_batch_id,
        updatedAt: row.updated_at,
      })),
    });
  } catch (err) {
    sendJson(res, 500, { ok: false, MSG: `999 Budget query failed: ${err.message}` });
  }
}

async function handleLaborImport(req, res) {
  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  const pool = getDbPool();
  if (!pool) {
    sendJson(res, 503, {
      ok: false,
      MSG: '503 Database is not configured. Set DB_INSTANCE_CONNECTION_NAME, DB_NAME, DB_USER, and DB_PASSWORD.',
    });
    return;
  }

  let payload;
  try {
    const raw = await readRequestBody(req, 25 * 1024 * 1024);
    payload = JSON.parse(raw || '{}');
  } catch (err) {
    sendJson(res, err.message === 'REQUEST_TOO_LARGE' ? 413 : 400, {
      ok: false,
      MSG: '999 Invalid labor import payload',
    });
    return;
  }

  const rows = laborRowsFromPayload(payload);
  if (!rows.length) {
    sendJson(res, 400, { ok: false, MSG: '999 No valid labor rows to import' });
    return;
  }

  const dates = rows.map(row => row.workDate).sort();
  const periodMonth = monthStartFromDate(dates[0]);
  const warehouses = Array.from(new Set(rows.map(row => row.warehouseName).filter(Boolean)));
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const batchResult = await client.query(
      `INSERT INTO ${schemaTable('import_batches')}
        (source_type, source_file, period_month, imported_by, note)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING id`,
      [
        'labor',
        String(payload.fileName || 'labor.xlsx'),
        periodMonth || null,
        String(payload.importedBy || ''),
        `date_range=${dates[0]}..${dates[dates.length - 1]}; rows=${rows.length}`,
      ],
    );
    const batchId = batchResult.rows[0].id;

    await client.query(
      `DELETE FROM ${schemaTable('labor_daily')}
       WHERE work_date >= $1::date
         AND work_date <= $2::date
         AND warehouse_name = ANY($3::text[])`,
      [dates[0], dates[dates.length - 1], warehouses],
    );

    const upsertSql = `
      INSERT INTO ${schemaTable('labor_daily')}
        (work_date, month, warehouse_name, vendor_name, shift_name, employee_id,
         department_name, operation_area, hours, cost, source_file, import_batch_id)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
      ON CONFLICT (work_date, warehouse_name, employee_id, operation_area, shift_name) DO UPDATE
      SET vendor_name = EXCLUDED.vendor_name,
          department_name = EXCLUDED.department_name,
          hours = EXCLUDED.hours,
          cost = EXCLUDED.cost,
          source_file = EXCLUDED.source_file,
          import_batch_id = EXCLUDED.import_batch_id
    `;

    for (const row of rows) {
      await client.query(upsertSql, [
        row.workDate,
        row.month,
        row.warehouseName,
        row.vendorName,
        row.shiftName,
        row.employeeId,
        row.departmentName,
        row.operationArea,
        row.hours,
        row.cost,
        String(payload.fileName || 'labor.xlsx'),
        batchId,
      ]);
    }

    await client.query('COMMIT');
    sendJson(res, 200, {
      ok: true,
      MSG: '000 Labor imported',
      batchId,
      rows: rows.length,
      dateFrom: dates[0],
      dateTo: dates[dates.length - 1],
    });
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    sendJson(res, 500, { ok: false, MSG: `999 Labor import failed: ${err.message}` });
  } finally {
    client.release();
  }
}

async function handleLaborData(req, res) {
  if (req.method !== 'GET') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  const pool = getDbPool();
  if (!pool) {
    sendJson(res, 503, {
      ok: false,
      MSG: '503 Database is not configured. Set DB_INSTANCE_CONNECTION_NAME, DB_NAME, DB_USER, and DB_PASSWORD.',
    });
    return;
  }

  const url = new URL(req.url, `http://${req.headers.host}`);
  const dateFrom = String(url.searchParams.get('date_from') || '').trim();
  const dateTo = String(url.searchParams.get('date_to') || '').trim();
  if (!isIsoDate(dateFrom) || !isIsoDate(dateTo) || dateFrom > dateTo) {
    sendJson(res, 400, { ok: false, MSG: '999 Invalid labor date range' });
    return;
  }

  try {
    const result = await pool.query(
      `SELECT work_date, warehouse_name, vendor_name, shift_name, employee_id,
              department_name, operation_area, hours, cost, source_file,
              import_batch_id, updated_at
       FROM ${schemaTable('labor_daily')}
       WHERE work_date >= $1::date AND work_date <= $2::date
       ORDER BY work_date, warehouse_name, employee_id, operation_area, shift_name`,
      [dateFrom, dateTo],
    );

    sendJson(res, 200, {
      ok: true,
      dateFrom,
      dateTo,
      rows: result.rows.map(row => ({
        date: row.work_date,
        wh: row.warehouse_name,
        vendor: row.vendor_name || '',
        shift: row.shift_name || '',
        empId: row.employee_id || '',
        dept: row.department_name || '',
        opArea: row.operation_area,
        hours: Number(row.hours) || 0,
        cost: Number(row.cost) || 0,
        sourceFile: row.source_file,
        importBatchId: row.import_batch_id,
        updatedAt: row.updated_at,
      })),
    });
  } catch (err) {
    sendJson(res, 500, { ok: false, MSG: `999 Labor query failed: ${err.message}` });
  }
}

const server = http.createServer((req, res) => {
  if (req.url.startsWith('/api/check-user')) {
    handleCheckUser(req, res);
    return;
  }
  if (req.url.startsWith('/api/page-permissions')) {
    if (req.method === 'GET')  { handleGetPermissions(req, res);  return; }
    if (req.method === 'POST') { handleSavePermissions(req, res); return; }
  }
  if (req.url.startsWith('/api/import/budget')) {
    handleBudgetImport(req, res);
    return;
  }
  if (req.url.startsWith('/api/data/budget')) {
    handleBudgetData(req, res);
    return;
  }
  if (req.url.startsWith('/api/import/labor')) {
    handleLaborImport(req, res);
    return;
  }
  if (req.url.startsWith('/api/data/labor')) {
    handleLaborData(req, res);
    return;
  }

  handleStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`KPL dashboard running at http://localhost:${PORT}/`);
});
