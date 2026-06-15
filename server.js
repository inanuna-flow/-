const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
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
// K_SERVICE 是 Cloud Run 自動設定的環境變數，用來判斷是否在正式環境
const IS_PROD = Boolean(process.env.K_SERVICE || process.env.NODE_ENV === 'production');

const VALID_PAGE_IDS = new Set(['daily','dispatch','freight','picks','labor','productivity','monthly','annual','import','org','typography']);
const BUDGET_TYPES = new Set(['labor', 'freight']);

// ── Session 設定 ──
const SESSION_TTL_MS = 8 * 60 * 60 * 1000; // 8 小時
const sessions = new Map(); // token -> { userId, expiresAt }

// ── 登入頻率限制設定 ──
const RATE_LIMIT_WINDOW_MS = 15 * 60 * 1000; // 15 分鐘
const RATE_LIMIT_MAX = 10; // 同一 IP 在 15 分鐘內最多嘗試 10 次
const loginAttempts = new Map(); // ip -> { count, windowStart }

// 定時清理過期的 session 和登入紀錄，避免記憶體無限增長
setInterval(() => {
  const now = Date.now();
  for (const [token, s] of sessions) {
    if (s.expiresAt <= now) sessions.delete(token);
  }
  for (const [ip, entry] of loginAttempts) {
    if (now - entry.windowStart > RATE_LIMIT_WINDOW_MS) loginAttempts.delete(ip);
  }
}, 30 * 60 * 1000);

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ttf': 'font/ttf',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

// 寬鬆的 CSP：允許 inline onclick、CDN（SheetJS 等），仍擋掉 object/embed
const CSP = [
  "default-src 'self'",
  "script-src 'self' 'unsafe-inline' 'unsafe-eval' https:",
  "style-src 'self' 'unsafe-inline' https:",
  "img-src 'self' data: https:",
  "connect-src 'self' https:",
  "font-src 'self' data: https:",
  "object-src 'none'",
  "base-uri 'self'",
  "form-action 'self'",
].join('; ');

const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
  'Content-Security-Policy': CSP,
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

// ── Cookie 解析 ──
function parseCookies(req) {
  const header = req.headers.cookie || '';
  const cookies = {};
  for (const part of header.split(';')) {
    const eqIdx = part.indexOf('=');
    if (eqIdx < 0) continue;
    const name = part.slice(0, eqIdx).trim();
    const value = part.slice(eqIdx + 1).trim();
    if (name) {
      try { cookies[name] = decodeURIComponent(value); }
      catch { cookies[name] = value; }
    }
  }
  return cookies;
}

// 取得客戶端真實 IP（Cloud Run 前面有 load balancer，真實 IP 在 x-forwarded-for）
function getClientIp(req) {
  return String(req.headers['x-forwarded-for'] || req.socket.remoteAddress || 'unknown')
    .split(',')[0].trim();
}

// ── 登入頻率限制 ──
function checkRateLimit(ip) {
  const now = Date.now();
  const entry = loginAttempts.get(ip);
  if (!entry || now - entry.windowStart > RATE_LIMIT_WINDOW_MS) {
    loginAttempts.set(ip, { count: 1, windowStart: now });
    return true;
  }
  if (entry.count >= RATE_LIMIT_MAX) return false;
  entry.count++;
  return true;
}

function resetRateLimit(ip) {
  loginAttempts.delete(ip);
}

// ── Session 管理 ──
function createSession(userId) {
  const token = crypto.randomBytes(32).toString('hex');
  sessions.set(token, { userId, expiresAt: Date.now() + SESSION_TTL_MS });
  return token;
}

function getSession(req) {
  const token = parseCookies(req)['kpl_session'];
  if (!token) return null;
  const session = sessions.get(token);
  if (!session || session.expiresAt <= Date.now()) {
    if (token) sessions.delete(token);
    return null;
  }
  return session;
}

function sessionCookieHeader(token) {
  const maxAge = Math.floor(SESSION_TTL_MS / 1000);
  let cookie = `kpl_session=${token}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${maxAge}`;
  if (IS_PROD) cookie += '; Secure';
  return cookie;
}

function clearSessionCookieHeader() {
  return 'kpl_session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0';
}

// 需要登入才能使用的 API 呼叫此函式；未登入時回傳 401 並返回 null
function requireSession(req, res) {
  const session = getSession(req);
  if (!session) {
    sendJson(res, 401, { ok: false, MSG: '401 請先登入' });
    return null;
  }
  return session;
}

function handleSession(req, res) {
  if (req.method !== 'GET') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }
  const session = requireSession(req, res);
  if (!session) return;
  sendJson(res, 200, {
    ok: true,
    userId: session.userId,
    isAdmin: session.userId === ADMIN_USER_ID,
  });
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
    dbPool.on('error', err => {
      console.error('[DB] pool 發生錯誤:', err.message);
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

function picksRowsFromPayload(payload) {
  const records = Array.isArray(payload.records) ? payload.records : [];
  const rows = records
    .map(record => ({
      workDate:      String(record.date  || '').trim(),
      month:         monthStartFromDate(record.date),
      warehouseName: String(record.wh    || '').trim(),
      businessType:  String(record.biz   || '').trim(),
      areaName:      String(record.area  || '').trim(),
      operationArea: String(record.op    || '').trim(),
      picksCount:    Number(record.picks) || 0,
    }))
    .filter(row =>
      isIsoDate(row.workDate) &&
      row.month &&
      row.warehouseName &&
      row.picksCount > 0
    );
  const grouped = new Map();
  rows.forEach(row => {
    const key = [
      row.workDate,
      row.warehouseName,
      row.businessType,
      row.areaName,
      row.operationArea,
    ].join('');
    const existing = grouped.get(key);
    if (existing) {
      existing.picksCount += row.picksCount;
    } else {
      grouped.set(key, { ...row });
    }
  });
  return Array.from(grouped.values());
}

async function handlePicksImport(req, res) {
  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  if (!requireSession(req, res)) return;

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
      MSG: '999 Invalid picks import payload',
    });
    return;
  }

  const rows = picksRowsFromPayload(payload);
  if (!rows.length) {
    sendJson(res, 400, { ok: false, MSG: '999 No valid picks rows to import' });
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
        'picks',
        String(payload.fileName || 'picks.xlsx'),
        periodMonth || null,
        String(payload.importedBy || ''),
        `date_range=${dates[0]}..${dates[dates.length - 1]}; rows=${rows.length}`,
      ],
    );
    const batchId = batchResult.rows[0].id;

    await client.query(
      `DELETE FROM ${schemaTable('picks_daily')}
       WHERE work_date >= $1::date
         AND work_date <= $2::date
         AND warehouse_name = ANY($3::text[])`,
      [dates[0], dates[dates.length - 1], warehouses],
    );

    const upsertSql = `
      INSERT INTO ${schemaTable('picks_daily')}
        (work_date, month, warehouse_name, business_type, area_name, operation_area,
         picks_count, source_file, import_batch_id)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
      ON CONFLICT (work_date, warehouse_name, business_type, area_name, operation_area) DO UPDATE
      SET picks_count     = EXCLUDED.picks_count,
          source_file     = EXCLUDED.source_file,
          import_batch_id = EXCLUDED.import_batch_id
    `;

    for (const row of rows) {
      await client.query(upsertSql, [
        row.workDate,
        row.month,
        row.warehouseName,
        row.businessType,
        row.areaName,
        row.operationArea,
        row.picksCount,
        String(payload.fileName || 'picks.xlsx'),
        batchId,
      ]);
    }

    await client.query('COMMIT');
    sendJson(res, 200, {
      ok: true,
      MSG: '000 Picks imported',
      batchId,
      rows: rows.length,
      dateFrom: dates[0],
      dateTo: dates[dates.length - 1],
    });
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    console.error('[picks] 匯入失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 揀次資料匯入失敗，請稍後再試' });
  } finally {
    client.release();
  }
}

async function handleCheckUser(req, res) {
  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  // 頻率限制：同一 IP 15 分鐘內超過 10 次就擋掉
  const ip = getClientIp(req);
  if (!checkRateLimit(ip)) {
    sendJson(res, 429, { ok: false, MSG: '999 登入嘗試次數過多，請 15 分鐘後再試' });
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

    const success = upstream.ok && String(data.MSG || '').startsWith('000');
    if (success) {
      // 登入成功：建立 session，寫入 HttpOnly cookie
      resetRateLimit(ip);
      const token = createSession(userId.toLowerCase());
      res.writeHead(200, {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store',
        'Set-Cookie': sessionCookieHeader(token),
        ...SECURITY_HEADERS,
      });
      res.end(JSON.stringify(data));
    } else {
      sendJson(res, upstream.ok ? 200 : 502, data);
    }
  } catch (err) {
    console.error('[auth] EIP 連線錯誤:', err.message);
    sendJson(res, 502, {
      ok: false,
      MSG: '999 無法連線至日翊 EIP，請稍後再試',
    });
  }
}

function handleLogout(req, res) {
  const cookies = parseCookies(req);
  const token = cookies['kpl_session'];
  if (token) sessions.delete(token);
  res.writeHead(200, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    'Set-Cookie': clearSessionCookieHeader(),
    ...SECURITY_HEADERS,
  });
  res.end(JSON.stringify({ ok: true, MSG: '000 已登出' }));
}

function handleStatic(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const pathname = decodeURIComponent(url.pathname);
  let relativePath = pathname === '/' ? 'login.html' : pathname.slice(1);
  if (relativePath.startsWith('kpl-dashboard/')) {
    relativePath = relativePath.slice('kpl-dashboard/'.length);
  }

  if (relativePath === 'index.html' && !getSession(req)) {
    res.writeHead(302, {
      'Location': 'login.html',
      'Cache-Control': 'no-store',
      ...SECURITY_HEADERS,
    });
    res.end();
    return;
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
  if (!requireSession(req, res)) return;
  fs.readFile(PERMISSIONS_FILE, 'utf8', (err, data) => {
    if (err) {
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

  if (!requireSession(req, res)) return;

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
    console.error('[permissions] EIP 連線錯誤:', err.message);
    sendJson(res, 502, { ok: false, MSG: '999 無法連線至 EIP，請稍後再試' });
    return;
  }

  const sanitized = {};
  VALID_PAGE_IDS.forEach(id => {
    sanitized[id] = perms[id] !== false;
  });

  fs.writeFile(PERMISSIONS_FILE, JSON.stringify(sanitized, null, 2), 'utf8', err => {
    if (err) {
      console.error('[permissions] 寫入失敗:', err.message);
      sendJson(res, 500, { ok: false, MSG: '999 儲存失敗，請稍後再試' });
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

  if (!requireSession(req, res)) return;

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
    console.error('[budget] 匯入失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 預算匯入失敗，請稍後再試' });
  } finally {
    client.release();
  }
}

async function handleBudgetData(req, res) {
  if (req.method !== 'GET') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  if (!requireSession(req, res)) return;

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
    console.error('[budget] 查詢失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 預算查詢失敗，請稍後再試' });
  }
}

async function handleLaborImport(req, res) {
  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  if (!requireSession(req, res)) return;

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
    console.error('[labor] 匯入失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 人力資料匯入失敗，請稍後再試' });
  } finally {
    client.release();
  }
}

async function handleLaborData(req, res) {
  if (req.method !== 'GET') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  if (!requireSession(req, res)) return;

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
  const summaryOnly = url.searchParams.get('summary') === '1';
  if (!isIsoDate(dateFrom) || !isIsoDate(dateTo) || dateFrom > dateTo) {
    sendJson(res, 400, { ok: false, MSG: '999 Invalid labor date range' });
    return;
  }

  try {
    if (summaryOnly) {
      const result = await pool.query(
        `SELECT work_date, warehouse_name,
                SUM(hours)::float AS hours,
                SUM(cost)::float AS cost
         FROM ${schemaTable('labor_daily')}
         WHERE work_date >= $1::date AND work_date <= $2::date
           AND operation_area <> '午休時間'
         GROUP BY work_date, warehouse_name
         ORDER BY work_date, warehouse_name`,
        [dateFrom, dateTo],
      );

      sendJson(res, 200, {
        ok: true,
        dateFrom,
        dateTo,
        rows: result.rows.map(row => ({
          date: row.work_date,
          wh: row.warehouse_name,
          hours: Number(row.hours) || 0,
          cost: Number(row.cost) || 0,
        })),
      });
      return;
    }

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
    console.error('[labor] 查詢失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 人力資料查詢失敗，請稍後再試' });
  }
}

async function handlePicksData(req, res) {
  if (req.method !== 'GET') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  if (!requireSession(req, res)) return;

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
    sendJson(res, 400, { ok: false, MSG: '999 Invalid picks date range' });
    return;
  }

  try {
    const result = await pool.query(
      `SELECT work_date, warehouse_name, business_type, area_name, operation_area,
              picks_count, source_file, import_batch_id, updated_at
       FROM ${schemaTable('picks_daily')}
       WHERE work_date >= $1::date AND work_date <= $2::date
       ORDER BY work_date, warehouse_name, business_type, area_name, operation_area`,
      [dateFrom, dateTo],
    );

    sendJson(res, 200, {
      ok: true,
      dateFrom,
      dateTo,
      rows: result.rows.map(row => ({
        date: row.work_date,
        wh: row.warehouse_name,
        biz: row.business_type || '',
        area: row.area_name || '',
        op: row.operation_area || '',
        picks: Number(row.picks_count) || 0,
        sourceFile: row.source_file,
        importBatchId: row.import_batch_id,
        updatedAt: row.updated_at,
      })),
    });
  } catch (err) {
    console.error('[picks] 查詢失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 揀次資料查詢失敗，請稍後再試' });
  }
}

async function handleFreightData(req, res) {
  if (req.method !== 'GET') { sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' }); return; }
  if (!requireSession(req, res)) return;
  const pool = getDbPool();
  if (!pool) { sendJson(res, 503, { ok: false, MSG: '503 Database not configured' }); return; }

  const url = new URL(req.url, `http://${req.headers.host}`);
  const dateFrom = String(url.searchParams.get('date_from') || '').trim();
  const dateTo   = String(url.searchParams.get('date_to')   || '').trim();
  const summaryOnly = url.searchParams.get('summary') === '1';
  if (!isIsoDate(dateFrom) || !isIsoDate(dateTo) || dateFrom > dateTo) {
    sendJson(res, 400, { ok: false, MSG: '400 date_from / date_to 格式錯誤' }); return;
  }

  try {
    // 三個查詢都是 read-only 且互不相依 → 並行送往 DB，縮短整體回應時間
    const dailyQuery = pool.query(
      `WITH mainline AS (
         SELECT work_date,
                warehouse_name,
                COALESCE(pricing_result, 0) AS cost
         FROM ${schemaTable('freight_mainline_daily')}
         WHERE work_date BETWEEN $1 AND $2
       ),
       nonmain AS (
         SELECT work_date,
                COALESCE(budget_warehouse, '大溪倉') AS warehouse_name,
                COALESCE(amount, 0) AS cost
         FROM ${schemaTable('freight_non_mainline_daily')}
         WHERE work_date BETWEEN $1 AND $2
           AND category_l1 IS DISTINCT FROM '不列入'
       ),
       combined AS (
         SELECT work_date, warehouse_name, cost FROM mainline
         UNION ALL
         SELECT work_date, warehouse_name, cost FROM nonmain
       )
       SELECT
         work_date::text AS date,
         SUM(CASE WHEN warehouse_name = '大溪倉' THEN cost ELSE 0 END)::float AS daxi,
         SUM(CASE WHEN warehouse_name = '大肚倉' THEN cost ELSE 0 END)::float AS dadu,
         SUM(CASE WHEN warehouse_name = '岡山倉' THEN cost ELSE 0 END)::float AS gangshan
       FROM combined
       GROUP BY work_date
       ORDER BY work_date`,
      [dateFrom, dateTo],
    );

    // 主線 + 非主線明細（供運費損益明細元件）
    const detailQuery = summaryOnly ? Promise.resolve({ rows: [] }) : pool.query(
      `SELECT * FROM (
         SELECT
           work_date::text AS date,
           'mainline'::text AS source_type,
           COALESCE(carrier, '未知') AS vendor,
           COALESCE(est_pricing_result, 0)::float AS estimated,
           COALESCE(pricing_result, 0)::float AS actual,
           NULL::text AS reason,
           route_code::text AS route,
           '主線'::text AS category_l1,
           task_category::text AS category_l2,
           warehouse_name::text AS budget_warehouse
         FROM ${schemaTable('freight_mainline_daily')}
         WHERE work_date BETWEEN $1 AND $2
         UNION ALL
         SELECT
           work_date::text AS date,
           'nonmainline'::text AS source_type,
           COALESCE(carrier, '非主線') AS vendor,
           0::float AS estimated,
           COALESCE(amount, 0)::float AS actual,
           dispatch_reason::text AS reason,
           delivery_item::text AS route,
           category_l1::text,
           category_l2::text,
           budget_warehouse::text
         FROM ${schemaTable('freight_non_mainline_daily')}
         WHERE work_date BETWEEN $1 AND $2
           AND category_l1 IS DISTINCT FROM '不列入'
       ) details
       ORDER BY date, source_type`,
      [dateFrom, dateTo],
    );

    // 上月總費用（供 F001 月趨勢比較）
    // 基準改為 dateTo（範圍結尾月），更符合「最近月的上一個月」直覺
    const lastMonthQuery = pool.query(
      `SELECT (
         COALESCE((
           SELECT SUM(COALESCE(pricing_result, 0))
           FROM ${schemaTable('freight_mainline_daily')}
           WHERE date_trunc('month', work_date) = date_trunc('month', $1::date) - INTERVAL '1 month'
         ), 0) +
         COALESCE((
           SELECT SUM(COALESCE(amount, 0))
           FROM ${schemaTable('freight_non_mainline_daily')}
           WHERE date_trunc('month', work_date) = date_trunc('month', $1::date) - INTERVAL '1 month'
             AND category_l1 IS DISTINCT FROM '不列入'
         ), 0)
       )::float AS total`,
      [dateTo],
    );

    const [dailyResult, detailResult, lastMonthResult] =
      await Promise.all([dailyQuery, detailQuery, lastMonthQuery]);

    sendJson(res, 200, {
      ok: true,
      dailyCosts:     dailyResult.rows,
      details:        detailResult.rows,
      lastMonthTotal: lastMonthResult.rows[0]?.total || 0,
    });
  } catch (err) {
    console.error('[freight-data] 查詢失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 運費資料查詢失敗' });
  }
}

async function handleDataRange(req, res) {
  if (req.method !== 'GET') {
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }
  if (!requireSession(req, res)) return;
  const pool = getDbPool();
  if (!pool) {
    sendJson(res, 503, { ok: false, MSG: '503 Database not configured' });
    return;
  }
  try {
    const baseResult = await pool.query(
      `SELECT
         (SELECT MIN(work_date)::text FROM ${schemaTable('labor_daily')}) AS labor_min,
         (SELECT MAX(work_date)::text FROM ${schemaTable('labor_daily')}) AS labor_max,
         (SELECT MIN(work_date)::text FROM ${schemaTable('picks_daily')}) AS picks_min,
         (SELECT MAX(work_date)::text FROM ${schemaTable('picks_daily')}) AS picks_max`,
    );
    const base = baseResult.rows[0] || {};

    // migration 002 可能尚未執行，freight 表不存在時不應拖累整個 range API
    let freightRow = { freight_main_min: null, freight_main_max: null, freight_nonmain_min: null, freight_nonmain_max: null };
    try {
      const frResult = await pool.query(
        `SELECT
           (SELECT MIN(work_date)::text FROM ${schemaTable('freight_mainline_daily')})     AS freight_main_min,
           (SELECT MAX(work_date)::text FROM ${schemaTable('freight_mainline_daily')})     AS freight_main_max,
           (SELECT MIN(work_date)::text FROM ${schemaTable('freight_non_mainline_daily')}) AS freight_nonmain_min,
           (SELECT MAX(work_date)::text FROM ${schemaTable('freight_non_mainline_daily')}) AS freight_nonmain_max`,
      );
      freightRow = frResult.rows[0] || freightRow;
    } catch (frErr) {
      console.warn('[range] freight 表尚未建立（migration 002 未執行）:', frErr.message);
    }

    const mins = [freightRow.freight_main_min, freightRow.freight_nonmain_min].filter(Boolean).sort();
    const maxs = [freightRow.freight_main_max, freightRow.freight_nonmain_max].filter(Boolean).sort();
    sendJson(res, 200, {
      ok: true,
      labor:   { min: base.labor_min || '', max: base.labor_max || '' },
      // 為了向下相容，保留 freight 欄位（取兩表的聯集範圍）
      freight: { min: mins[0] || '', max: maxs[maxs.length - 1] || '' },
      freightMainline:    { min: freightRow.freight_main_min    || '', max: freightRow.freight_main_max    || '' },
      freightNonMainline: { min: freightRow.freight_nonmain_min || '', max: freightRow.freight_nonmain_max || '' },
      picks:   { min: base.picks_min || '', max: base.picks_max || '' },
    });
  } catch (err) {
    console.error('[range] 查詢失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 查詢失敗' });
  }
}

// ════════════════════════════════════════════════════════════════
// 運費（主線 / 非主線）匯入與查詢
// ════════════════════════════════════════════════════════════════

// 民國年日期字串轉西元 ISO：'115/05/01' → '2026-05-01'
function minguoToIso(value) {
  const s = String(value || '').trim();
  const m = s.match(/^(\d{2,3})[/.\-](\d{1,2})[/.\-](\d{1,2})/);
  if (!m) return '';
  const y = Number(m[1]) + 1911;
  const mm = String(Number(m[2])).padStart(2, '0');
  const dd = String(Number(m[3])).padStart(2, '0');
  return `${y}-${mm}-${dd}`;
}

// Excel serial date → ISO：46143 → '2026-05-01'
// Excel epoch 1899-12-30（避開 1900 leap-year bug）
function excelSerialToIso(value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 1 || n > 80000) return '';
  const epoch = Date.UTC(1899, 11, 30);
  const ms = epoch + Math.round(n) * 86400000;
  const d = new Date(ms);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

// 任意日期輸入轉 ISO（支援：ISO 字串、民國年字串、Excel 序號、Date 物件）
function normalizeWorkDate(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date && !isNaN(value)) {
    const yyyy = value.getFullYear();
    const mm = String(value.getMonth() + 1).padStart(2, '0');
    const dd = String(value.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
  if (typeof value === 'number') return excelSerialToIso(value);
  const s = String(value).trim();
  if (isIsoDate(s)) return s;
  const minguo = minguoToIso(s);
  if (minguo) return minguo;
  // 也接受 2026/05/01
  const m = s.match(/^(\d{4})[/.\-](\d{1,2})[/.\-](\d{1,2})/);
  if (m) {
    return `${m[1]}-${String(Number(m[2])).padStart(2, '0')}-${String(Number(m[3])).padStart(2, '0')}`;
  }
  // 8 位純數字優先當 YYYYMMDD 解析（避免被誤判為極大 Excel 序號）
  const compactMatch = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (compactMatch) {
    const mm = Number(compactMatch[2]);
    const dd = Number(compactMatch[3]);
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      return `${compactMatch[1]}-${compactMatch[2]}-${compactMatch[3]}`;
    }
  }
  // 純數字字串當 Excel 序號
  if (/^\d+(\.\d+)?$/.test(s)) return excelSerialToIso(Number(s));
  return '';
}

function toNumOrNull(value) {
  if (value === '' || value == null) return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function toNumOrZero(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function toIntOrZero(value) {
  const n = parseInt(value, 10);
  return Number.isFinite(n) ? n : 0;
}

function toTextOrNull(value) {
  if (value == null) return null;
  const s = String(value).trim();
  return s === '' ? null : s;
}

// 非主線運費分類規則（業務同仁 2026-06-08 確認版）
// 規則順序「由上往下、命中即停」，總共 12 條。
// 回傳 [category_l1, category_l2, budgetWarehouse]
//   budgetWarehouse: 跨區轉運費統一歸「大溪倉」；其餘回 null 代表沿用 row.warehouse_name
function classifyNonMainline(row) {
  const I = String(row.dispatch_reason || '');
  const J = String(row.pricing_method  || '');
  const F = String(row.delivery_item   || '');
  const N = String(row.note            || '');
  const wh = row.warehouse_name || null;

  // 序 1：誤 key（資料錯誤）→ 不列入
  if (/誤\s*key/i.test(N)) return ['不列入', '誤key', wh];

  // 序 2：上收（不在預算範圍）→ 不列入
  if (/上收/.test(I) || /上收/.test(N)) return ['不列入', '上收', wh];

  // 序 3：違規罰款（關鍵字優先，避免被後續轉運/共配誤判）
  if (/違規|罰款|警察|驅趕|封路|施工|無法停靠|火災/.test(N)) return ['其他', '違規罰款', wh];

  // 序 4：花蓮轉運費（轉運 + 花蓮）
  if ((/轉運/.test(I) || /轉運/.test(F)) && /花蓮/.test(`${J} ${N}`)) {
    return ['轉運', '花蓮轉運費', wh];
  }

  // 序 5：跨區轉運費（其餘轉運）→ 統一歸大溪倉
  if (/轉運/.test(I) || /轉運/.test(F)) {
    return ['轉運', '跨區轉運費', '大溪倉'];
  }

  // 序 6：離島海陸空運（馬祖）
  if (/離島/.test(I) && /馬祖/.test(N)) return ['其他', '離島海陸空運(馬祖)', wh];

  // 序 7：離島運費（澎湖、金門）
  if (/離島/.test(I) || /澎湖|金門/.test(N)) return ['其他', '離島運費(澎湖、金門)', wh];

  // 序 8：全台共配費
  if (/共配/.test(N)) return ['其他', '全台共配費', wh];

  // 序 9：正物流（爆量）
  if (/爆量/.test(I)) return ['加派', '正物流', wh];

  // 序 10：逆物流（回頭收/店回/固定店回專車）
  if (/回頭收|店回|固定店回專車/.test(I)) return ['加派', '逆物流', wh];

  // 序 11：專車（特殊點專車 OR 新開店 OR J=其他專車系列 OR 派車原因=其他 之 fallback）
  // 實測資料中 J 為「其他專車」「其他專車2」「其他專車3」皆屬此類；先級規則先攔截後，
  // 落到此處的 /其他專車/ 都應歸專車（已驗證不會誤捕跨區/離島/上收等）。
  // 業務確認：派車原因 = '其他' 沒被前面規則攔截的個案（如 POC 棧板回收）也歸專車。
  if (/特殊點專車|新開店/.test(I) || /其他專車/.test(J) || /^其他$/.test(I.trim())) {
    return ['其他', '專車', wh];
  }

  // 序 12：無法判斷 → 人工確認
  return ['無法判斷', null, wh];
}

function freightMainlineRowsFromPayload(payload) {
  const records = Array.isArray(payload.records) ? payload.records : [];
  return records.map(r => {
    const workDate = normalizeWorkDate(r['進貨日'] ?? r.workDate ?? r.work_date);
    const month = monthStartFromDate(workDate);
    return {
      warehouse_name:        toTextOrNull(r['倉別']            ?? r.warehouse_name),
      status_locked:         toTextOrNull(r['明細狀態']        ?? r.status_locked),
      carry_month:           toTextOrNull(r['結轉年月']        ?? r.carry_month),
      work_date:             workDate,
      month,
      shipper:               toTextOrNull(r['配別']            ?? r.shipper),
      route_source:          toTextOrNull(r['路線來源']        ?? r.route_source),
      // route_code / carrier 必須為非 NULL，否則 ON CONFLICT 比對會把 NULL 當成「不衝突」→ 重複列
      route_code:            toTextOrNull(r['路線']            ?? r.route_code) || '',
      task_category:         toTextOrNull(r['任務分類']        ?? r.task_category),
      carrier:               toTextOrNull(r['配送商']          ?? r.carrier) || '',
      route_status:          toTextOrNull(r['路線完成\r\n狀態'] ?? r['路線完成狀態'] ?? r.route_status),
      task_note:             toTextOrNull(r['任務備註']        ?? r.task_note),

      est_pricing_type:      toTextOrNull(r['預計\r\n計價類型'] ?? r['預計計價類型'] ?? r.est_pricing_type),
      est_price_list:        toTextOrNull(r['預計\r\n價目表']   ?? r['預計價目表']   ?? r.est_price_list),
      est_trip_type:         toTextOrNull(r['預計\r\n全/半趟']  ?? r['預計全/半趟']  ?? r.est_trip_type),
      est_vehicle_tonnage:   toTextOrNull(r['合約\r\n車輛噸型'] ?? r['合約車輛噸型'] ?? r.est_vehicle_tonnage),
      est_pricing_qty:       toNumOrNull (r['預計\r\n計價數']   ?? r['預計計價數']   ?? r.est_pricing_qty),
      est_pricing_tier:      toTextOrNull(r['預計\r\n計價級距'] ?? r['預計計價級距'] ?? r.est_pricing_tier),
      est_pricing_unit:      toTextOrNull(r['預計\r\n計價單位'] ?? r['預計計價單位'] ?? r.est_pricing_unit),
      est_tier_rate:         toNumOrNull (r['預計\r\n級距費率'] ?? r['預計級距費率'] ?? r.est_tier_rate),
      est_dispatch_price:    toNumOrNull (r['預計\r\n出車價']   ?? r['預計出車價']   ?? r.est_dispatch_price),
      est_pricing_result:    toNumOrNull (r['預計\r\n計價結果'] ?? r['預計計價結果'] ?? r.est_pricing_result),
      est_with_tax:          toTextOrNull(r['預計\r\n是否含稅'] ?? r['預計是否含稅'] ?? r.est_with_tax),

      arr_vehicle_tonnage:   toTextOrNull(r['到點\r\n車輛噸型'] ?? r['到點車輛噸型'] ?? r.arr_vehicle_tonnage),
      arr_pricing_qty:       toNumOrNull (r['到點\r\n計價數']   ?? r['到點計價數']   ?? r.arr_pricing_qty),
      arr_pricing_tier:      toTextOrNull(r['到點\r\n計價級距'] ?? r['到點計價級距'] ?? r.arr_pricing_tier),
      arr_pricing_unit:      toTextOrNull(r['到點\r\n計價單位'] ?? r['到點計價單位'] ?? r.arr_pricing_unit),
      arr_tier_rate:         toNumOrNull (r['到點\r\n級距費率'] ?? r['到點級距費率'] ?? r.arr_tier_rate),
      arr_dispatch_price:    toNumOrNull (r['到點\r\n車輛噸型出車價'] ?? r['到點車輛噸型出車價'] ?? r.arr_dispatch_price),
      arr_pricing_result:    toNumOrNull (r['到點\r\n計價結果'] ?? r['到點計價結果'] ?? r.arr_pricing_result),

      pricing_type:          toTextOrNull(r['計價類型']         ?? r.pricing_type),
      price_list:            toTextOrNull(r['計價\r\n價目表']   ?? r['計價價目表']   ?? r.price_list),
      trip_type:             toTextOrNull(r['計價\r\n全/半趟']  ?? r['計價全/半趟']  ?? r.trip_type),
      vehicle_tonnage:       toTextOrNull(r['計價\r\n車輛噸型'] ?? r['計價車輛噸型'] ?? r.vehicle_tonnage),
      pricing_qty:           toNumOrNull (r['計價數']           ?? r.pricing_qty),
      pricing_tier:          toTextOrNull(r['計價級距']         ?? r.pricing_tier),
      pricing_unit:          toTextOrNull(r['計價單位']         ?? r.pricing_unit),
      tier_rate:             toNumOrNull (r['計價\r\n級距費率'] ?? r['計價級距費率'] ?? r.tier_rate),
      dispatch_price:        toNumOrNull (r['出車價']           ?? r.dispatch_price),
      pricing_result:        toNumOrNull (r['計價結果']         ?? r.pricing_result),
      with_tax:              toTextOrNull(r['計價\r\n是否含稅'] ?? r['計價是否含稅'] ?? r.with_tax),

      adjust_reason:         toTextOrNull(r['調整原因']         ?? r.adjust_reason),
      system_updated_at_raw: toTextOrNull(r['系統最後\r\n更新時間'] ?? r['系統最後更新時間'] ?? r.system_updated_at_raw),
      est_created_by:        toTextOrNull(r['預計計價\r\n建立人員'] ?? r['預計計價建立人員'] ?? r.est_created_by),
      est_created_at_raw:    toTextOrNull(r['預計計價\r\n建立時間'] ?? r['預計計價建立時間'] ?? r.est_created_at_raw),
      actual_updated_by:     toTextOrNull(r['實際/調整後\r\n計價更新人員'] ?? r['實際/調整後計價更新人員'] ?? r.actual_updated_by),
      actual_updated_at_raw: toTextOrNull(r['實際/調整後\r\n計價更新時間'] ?? r['實際/調整後計價更新時間'] ?? r.actual_updated_at_raw),
      status_updated_by:     toTextOrNull(r['狀態\r\n更新人員'] ?? r['狀態更新人員'] ?? r.status_updated_by),
      status_updated_at_raw: toTextOrNull(r['狀態\r\n更新時間'] ?? r['狀態更新時間'] ?? r.status_updated_at_raw),
      carry_by:              toTextOrNull(r['結轉人員']         ?? r.carry_by),
      carry_at_raw:          toTextOrNull(r['結轉時間']         ?? r.carry_at_raw),
    };
  }).filter(row => isIsoDate(row.work_date) && row.warehouse_name);
}

function freightNonMainlineRowsFromPayload(payload) {
  const records = Array.isArray(payload.records) ? payload.records : [];
  return records.map(r => {
    const workDate = normalizeWorkDate(r['進貨日'] ?? r.work_date);
    const month = monthStartFromDate(workDate);
    const unit    = toNumOrZero(r['單價'] ?? r.unit_price);
    const trip    = toIntOrZero(r['趟數'] ?? r.trip_count);
    const rawAmt  = r['費用小計'] ?? r.amount;
    const amt     = toNumOrZero(rawAmt);
    return {
      work_date:       workDate,
      month,
      shipper:         toTextOrNull(r['配別']     ?? r.shipper),
      region:          toTextOrNull(r['區域']     ?? r.region),
      vehicle_type:    toTextOrNull(r['車型']     ?? r.vehicle_type),
      delivery_item:   toTextOrNull(r['配送項目'] ?? r.delivery_item),
      warehouse_name:  toTextOrNull(r['廠商']     ?? r.warehouse_name),
      carrier:         toTextOrNull(r['配送商']   ?? r.carrier),
      dispatch_reason: toTextOrNull(r['派車原因'] ?? r.dispatch_reason),
      pricing_method:  toTextOrNull(r['計價方式'] ?? r.pricing_method),
      unit_price:      unit,
      trip_count:      trip,
      // rawAmt 為空/undefined → 回退計算值；rawAmt 明確填 0 → 保留 0
      amount:          (rawAmt !== '' && rawAmt != null) ? amt : (unit * trip),
      note:            toTextOrNull(r['備註']     ?? r.note),
      category_l1:     null,   // 由 classifyNonMainline() 於寫入前填寫
      category_l2:     null,
      budget_warehouse: null,
    };
  }).filter(row => isIsoDate(row.work_date) && row.warehouse_name);
}

async function handleFreightMainlineImport(req, res) {
  if (req.method !== 'POST') { sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' }); return; }
  if (!requireSession(req, res)) return;

  const pool = getDbPool();
  if (!pool) { sendJson(res, 503, { ok: false, MSG: '503 Database is not configured.' }); return; }

  let payload;
  try {
    const raw = await readRequestBody(req, 50 * 1024 * 1024);
    payload = JSON.parse(raw || '{}');
  } catch (err) {
    sendJson(res, err.message === 'REQUEST_TOO_LARGE' ? 413 : 400, { ok: false, MSG: '999 Invalid freight-mainline import payload' });
    return;
  }

  const rows = freightMainlineRowsFromPayload(payload);
  if (!rows.length) { sendJson(res, 400, { ok: false, MSG: '999 No valid freight-mainline rows to import' }); return; }

  const dryRun = Boolean(payload.dryRun);
  const fileName = String(payload.fileName || 'freight_mainline.xlsx');

  // 計算「將被覆蓋」的筆數（同 work_date + warehouse + route_code + carrier 已存在者）
  const dates = rows.map(r => r.work_date).sort();
  const dateFrom = dates[0];
  const dateTo = dates[dates.length - 1];

  try {
    const dupResult = await pool.query(
      `SELECT COUNT(*)::int AS n FROM ${schemaTable('freight_mainline_daily')} t
       WHERE work_date >= $1::date AND work_date <= $2::date`,
      [dateFrom, dateTo],
    );
    const existingInRange = dupResult.rows[0]?.n || 0;

    if (dryRun) {
      sendJson(res, 200, {
        ok: true,
        dryRun: true,
        rowsToImport: rows.length,
        existingInRange,
        dateFrom, dateTo,
        MSG: '000 Dry run complete',
      });
      return;
    }
  } catch (err) {
    console.error('[freight-mainline] dry-run 失敗:', err.message);
    const msg = err.code === '42P01'
      ? '503 資料表尚未建立，請先執行 migration 002（002_split_freight_tables.sql）'
      : '999 主線運費檢查失敗，請稍後再試';
    sendJson(res, err.code === '42P01' ? 503 : 500, { ok: false, MSG: msg });
    return;
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const batchResult = await client.query(
      `INSERT INTO ${schemaTable('import_batches')} (source_type, source_file, period_month, imported_by, note)
       VALUES ($1, $2, $3, $4, $5) RETURNING id`,
      ['freight_mainline', fileName, monthStartFromDate(dateFrom) || null, String(payload.importedBy || ''),
       `date_range=${dateFrom}..${dateTo}; rows=${rows.length}`],
    );
    const batchId = batchResult.rows[0].id;

    const upsertCols = [
      'warehouse_name','status_locked','carry_month','work_date','month','shipper',
      'route_source','route_code','task_category','carrier','route_status','task_note',
      'est_pricing_type','est_price_list','est_trip_type','est_vehicle_tonnage',
      'est_pricing_qty','est_pricing_tier','est_pricing_unit','est_tier_rate',
      'est_dispatch_price','est_pricing_result','est_with_tax',
      'arr_vehicle_tonnage','arr_pricing_qty','arr_pricing_tier','arr_pricing_unit',
      'arr_tier_rate','arr_dispatch_price','arr_pricing_result',
      'pricing_type','price_list','trip_type','vehicle_tonnage','pricing_qty',
      'pricing_tier','pricing_unit','tier_rate','dispatch_price','pricing_result','with_tax',
      'adjust_reason','system_updated_at_raw','est_created_by','est_created_at_raw',
      'actual_updated_by','actual_updated_at_raw','status_updated_by','status_updated_at_raw',
      'carry_by','carry_at_raw','source_file','import_batch_id',
    ];
    const updateClause = upsertCols
      .filter(c => !['warehouse_name','work_date','route_code','carrier'].includes(c))
      .map(c => `${c} = EXCLUDED.${c}`).join(', ');

    const BATCH = 500;
    for (let i = 0; i < rows.length; i += BATCH) {
      const batch = rows.slice(i, i + BATCH);
      const valueGroups = batch.map((_, bi) =>
        `(${upsertCols.map((__, ci) => `$${bi * upsertCols.length + ci + 1}`).join(', ')})`
      ).join(', ');
      const values = batch.flatMap(row => upsertCols.map(c => {
        if (c === 'source_file')     return fileName;
        if (c === 'import_batch_id') return batchId;
        return row[c];
      }));
      await client.query(
        `INSERT INTO ${schemaTable('freight_mainline_daily')} (${upsertCols.join(', ')})
         VALUES ${valueGroups}
         ON CONFLICT (work_date, warehouse_name, route_code, carrier) DO UPDATE
         SET ${updateClause}`,
        values,
      );
    }

    await client.query('COMMIT');
    sendJson(res, 200, {
      ok: true, MSG: '000 Freight mainline imported',
      batchId, rows: rows.length, dateFrom, dateTo,
    });
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    console.error('[freight-mainline] 匯入失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 主線運費匯入失敗，請稍後再試' });
  } finally {
    client.release();
  }
}

async function handleFreightNonMainlineImport(req, res) {
  if (req.method !== 'POST') { sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' }); return; }
  if (!requireSession(req, res)) return;

  const pool = getDbPool();
  if (!pool) { sendJson(res, 503, { ok: false, MSG: '503 Database is not configured.' }); return; }

  let payload;
  try {
    const raw = await readRequestBody(req, 50 * 1024 * 1024);
    payload = JSON.parse(raw || '{}');
  } catch (err) {
    sendJson(res, err.message === 'REQUEST_TOO_LARGE' ? 413 : 400, { ok: false, MSG: '999 Invalid freight-non-mainline import payload' });
    return;
  }

  const rows = freightNonMainlineRowsFromPayload(payload);
  if (!rows.length) { sendJson(res, 400, { ok: false, MSG: '999 No valid freight-non-mainline rows to import' }); return; }

  const dryRun = Boolean(payload.dryRun);
  const fileName = String(payload.fileName || 'freight_non_mainline.xlsx');
  const dates = rows.map(r => r.work_date).sort();
  const dateFrom = dates[0];
  const dateTo = dates[dates.length - 1];

  // 非主線無天然唯一鍵，重複上傳會新增 → 採「刪除日期區間內所有資料再寫入」策略（與 labor/picks 一致）
  try {
    const dupResult = await pool.query(
      `SELECT COUNT(*)::int AS n FROM ${schemaTable('freight_non_mainline_daily')}
       WHERE work_date >= $1::date AND work_date <= $2::date`,
      [dateFrom, dateTo],
    );
    const existingInRange = dupResult.rows[0]?.n || 0;

    if (dryRun) {
      const classMap = {};
      rows.forEach(row => {
        const [l1, l2] = classifyNonMainline(row);
        const key = `${l1}||${l2 || ''}`;
        if (!classMap[key]) classMap[key] = { l1, l2: l2 || null, count: 0 };
        classMap[key].count++;
      });
      const classification = Object.values(classMap).sort((a, b) => b.count - a.count);
      sendJson(res, 200, {
        ok: true,
        dryRun: true,
        rowsToImport: rows.length,
        existingInRange,
        dateFrom, dateTo,
        classification,
        MSG: '000 Dry run complete',
      });
      return;
    }
  } catch (err) {
    console.error('[freight-non-mainline] dry-run 失敗:', err.message);
    const msg = err.code === '42P01'
      ? '503 資料表尚未建立，請先執行 migration 002（002_split_freight_tables.sql）'
      : '999 非主線運費檢查失敗，請稍後再試';
    sendJson(res, err.code === '42P01' ? 503 : 500, { ok: false, MSG: msg });
    return;
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const batchResult = await client.query(
      `INSERT INTO ${schemaTable('import_batches')} (source_type, source_file, period_month, imported_by, note)
       VALUES ($1, $2, $3, $4, $5) RETURNING id`,
      ['freight_non_mainline', fileName, monthStartFromDate(dateFrom) || null, String(payload.importedBy || ''),
       `date_range=${dateFrom}..${dateTo}; rows=${rows.length}`],
    );
    const batchId = batchResult.rows[0].id;

    await client.query(
      `DELETE FROM ${schemaTable('freight_non_mainline_daily')}
       WHERE work_date >= $1::date AND work_date <= $2::date`,
      [dateFrom, dateTo],
    );

    const NM_COLS = 19;
    const BATCH = 500;

    for (const row of rows) {
      const [l1, l2, budgetWh] = classifyNonMainline(row);
      row.category_l1 = l1;
      row.category_l2 = l2;
      row.budget_warehouse = budgetWh || row.warehouse_name;
    }

    for (let i = 0; i < rows.length; i += BATCH) {
      const batch = rows.slice(i, i + BATCH);
      const valueGroups = batch.map((_, bi) =>
        `(${Array.from({length: NM_COLS}, (__, ci) => `$${bi * NM_COLS + ci + 1}`).join(', ')})`
      ).join(', ');
      const values = batch.flatMap(row => [
        row.work_date, row.month, row.shipper, row.region, row.vehicle_type, row.delivery_item,
        row.warehouse_name, row.carrier, row.dispatch_reason, row.pricing_method,
        row.unit_price, row.trip_count, row.amount, row.note,
        row.category_l1, row.category_l2, row.budget_warehouse, fileName, batchId,
      ]);
      await client.query(
        `INSERT INTO ${schemaTable('freight_non_mainline_daily')}
          (work_date, month, shipper, region, vehicle_type, delivery_item, warehouse_name,
           carrier, dispatch_reason, pricing_method, unit_price, trip_count, amount, note,
           category_l1, category_l2, budget_warehouse, source_file, import_batch_id)
         VALUES ${valueGroups}`,
        values,
      );
    }

    await client.query('COMMIT');
    sendJson(res, 200, {
      ok: true, MSG: '000 Freight non-mainline imported',
      batchId, rows: rows.length, dateFrom, dateTo,
    });
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    console.error('[freight-non-mainline] 匯入失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 非主線運費匯入失敗，請稍後再試' });
  } finally {
    client.release();
  }
}

async function handleFreightMainlineData(req, res) {
  if (req.method !== 'GET') { sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' }); return; }
  if (!requireSession(req, res)) return;
  const pool = getDbPool();
  if (!pool) { sendJson(res, 503, { ok: false, MSG: '503 Database is not configured.' }); return; }

  const url = new URL(req.url, `http://${req.headers.host}`);
  const dateFrom = String(url.searchParams.get('date_from') || '').trim();
  const dateTo   = String(url.searchParams.get('date_to')   || '').trim();
  if (!isIsoDate(dateFrom) || !isIsoDate(dateTo) || dateFrom > dateTo) {
    sendJson(res, 400, { ok: false, MSG: '999 Invalid freight-mainline date range' });
    return;
  }

  try {
    const result = await pool.query(
      `SELECT work_date, month, warehouse_name, route_code, task_category, carrier,
              route_status, task_note, vehicle_tonnage, trip_type, pricing_qty,
              pricing_tier, pricing_unit, tier_rate, dispatch_price, pricing_result,
              with_tax, source_file, import_batch_id, updated_at
       FROM ${schemaTable('freight_mainline_daily')}
       WHERE work_date >= $1::date AND work_date <= $2::date
       ORDER BY work_date, warehouse_name, route_code, carrier`,
      [dateFrom, dateTo],
    );
    sendJson(res, 200, {
      ok: true, dateFrom, dateTo,
      rows: result.rows.map(row => ({
        date: row.work_date, month: row.month,
        wh: row.warehouse_name, route: row.route_code,
        taskCategory: row.task_category, carrier: row.carrier,
        routeStatus: row.route_status, note: row.task_note,
        vehicleTonnage: row.vehicle_tonnage, tripType: row.trip_type,
        pricingQty: Number(row.pricing_qty) || 0,
        pricingTier: row.pricing_tier, pricingUnit: row.pricing_unit,
        tierRate: Number(row.tier_rate) || 0,
        dispatchPrice: Number(row.dispatch_price) || 0,
        amount: Number(row.pricing_result) || 0,
        withTax: row.with_tax,
        sourceFile: row.source_file,
        importBatchId: row.import_batch_id,
        updatedAt: row.updated_at,
      })),
    });
  } catch (err) {
    console.error('[freight-mainline] 查詢失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 主線運費查詢失敗' });
  }
}

async function handleFreightNonMainlineData(req, res) {
  if (req.method !== 'GET') { sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' }); return; }
  if (!requireSession(req, res)) return;
  const pool = getDbPool();
  if (!pool) { sendJson(res, 503, { ok: false, MSG: '503 Database is not configured.' }); return; }

  const url = new URL(req.url, `http://${req.headers.host}`);
  const dateFrom = String(url.searchParams.get('date_from') || '').trim();
  const dateTo   = String(url.searchParams.get('date_to')   || '').trim();
  if (!isIsoDate(dateFrom) || !isIsoDate(dateTo) || dateFrom > dateTo) {
    sendJson(res, 400, { ok: false, MSG: '999 Invalid freight-non-mainline date range' });
    return;
  }

  try {
    const result = await pool.query(
      `SELECT work_date, month, warehouse_name, region, vehicle_type, delivery_item,
              carrier, dispatch_reason, pricing_method, unit_price, trip_count, amount,
              note, category_l1, category_l2, budget_warehouse,
              source_file, import_batch_id, updated_at
       FROM ${schemaTable('freight_non_mainline_daily')}
       WHERE work_date >= $1::date AND work_date <= $2::date
       ORDER BY work_date, warehouse_name, carrier`,
      [dateFrom, dateTo],
    );
    sendJson(res, 200, {
      ok: true, dateFrom, dateTo,
      rows: result.rows.map(row => ({
        date: row.work_date, month: row.month,
        wh: row.warehouse_name, region: row.region,
        vehicleType: row.vehicle_type, deliveryItem: row.delivery_item,
        carrier: row.carrier, dispatchReason: row.dispatch_reason,
        pricingMethod: row.pricing_method,
        unitPrice: Number(row.unit_price) || 0,
        tripCount: Number(row.trip_count) || 0,
        amount: Number(row.amount) || 0,
        note: row.note,
        categoryL1: row.category_l1,
        categoryL2: row.category_l2,
        budgetWarehouse: row.budget_warehouse,
        sourceFile: row.source_file,
        importBatchId: row.import_batch_id,
        updatedAt: row.updated_at,
      })),
    });
  } catch (err) {
    console.error('[freight-non-mainline] 查詢失敗:', err.message);
    sendJson(res, 500, { ok: false, MSG: '999 非主線運費查詢失敗' });
  }
}

// 舊端點：已拆成主線/非主線，回 410 Gone 提醒前端切換
function handleFreightGone(req, res) {
  sendJson(res, 410, {
    ok: false,
    MSG: '410 此端點已淘汰；請改用 /api/import/freight-mainline 或 /api/import/freight-non-mainline（查詢端同理）',
  });
}

// 改用精確 pathname 比對，避免 startsWith 因新增端點而被誤路由
// 例如：/api/data/freight 與 /api/data/freight-mainline 若用 startsWith 順序就敏感；用 === 則安全
const ROUTES = {
  '/api/check-user':         handleCheckUser,
  '/api/session':            handleSession,
  '/api/logout':             handleLogout,
  '/api/import/budget':      handleBudgetImport,
  '/api/data/budget':        handleBudgetData,
  '/api/import/labor':       handleLaborImport,
  '/api/import/picks':       handlePicksImport,
  '/api/import/freight-mainline':     handleFreightMainlineImport,
  '/api/import/freight-non-mainline': handleFreightNonMainlineImport,
  '/api/data/freight-mainline':       handleFreightMainlineData,
  '/api/data/freight-non-mainline':   handleFreightNonMainlineData,
  '/api/data/freight':       handleFreightData,
  '/api/import/freight':     handleFreightGone,   // 舊端點已淘汰
  '/api/data/range':         handleDataRange,
  '/api/data/labor':         handleLaborData,
  '/api/data/picks':         handlePicksData,
};

const server = http.createServer((req, res) => {
  // 切掉 query string 取純路徑做比對
  const pathname = req.url.split('?')[0];

  // page-permissions 需依 method 分流
  if (pathname === '/api/page-permissions') {
    if (req.method === 'GET')  { handleGetPermissions(req, res);  return; }
    if (req.method === 'POST') { handleSavePermissions(req, res); return; }
    sendJson(res, 405, { ok: false, MSG: '999 Method Not Allowed' });
    return;
  }

  const handler = ROUTES[pathname];
  if (handler) { handler(req, res); return; }

  handleStatic(req, res);
});

// ── 全域錯誤保護，避免未預期的錯誤讓伺服器當掉 ──
process.on('unhandledRejection', (reason) => {
  console.error('[server] 未處理的 Promise 拒絕:', reason);
});

process.on('uncaughtException', (err) => {
  console.error('[server] 未捕捉的例外:', err.message);
});

server.listen(PORT, () => {
  console.log(`KPL dashboard running at http://localhost:${PORT}/`);
});
