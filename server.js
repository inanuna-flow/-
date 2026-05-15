const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const EIP_CHECK_USER_URL = 'https://eip.fme.com.tw/FMEIP/AasApi/CheckUserId';
const PERMISSIONS_FILE = path.join(ROOT, 'page_permissions.json');
const ADMIN_USER_ID = (process.env.ADMIN_USER_ID || 'inari').toLowerCase();

const VALID_PAGE_IDS = new Set(['daily','dispatch','freight','labor','productivity','monthly','annual','import','org']);

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

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.setEncoding('utf8');
    req.on('data', chunk => {
      body += chunk;
      if (body.length > 4096) {
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
  const relativePath = pathname === '/' ? 'index.html' : pathname.slice(1);
  const filePath = path.resolve(ROOT, relativePath);

  if (!filePath.startsWith(ROOT)) {
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

const server = http.createServer((req, res) => {
  if (req.url.startsWith('/api/check-user')) {
    handleCheckUser(req, res);
    return;
  }
  if (req.url.startsWith('/api/page-permissions')) {
    if (req.method === 'GET')  { handleGetPermissions(req, res);  return; }
    if (req.method === 'POST') { handleSavePermissions(req, res); return; }
  }

  handleStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`KPL dashboard running at http://localhost:${PORT}/`);
});
