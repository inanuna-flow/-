const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const EIP_CHECK_USER_URL = 'https://eip.fme.com.tw/FMEIP/AasApi/CheckUserId';

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

function sendJson(res, status, payload) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
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
    });
    res.end(data);
  });
}

const server = http.createServer((req, res) => {
  if (req.url.startsWith('/api/check-user')) {
    handleCheckUser(req, res);
    return;
  }

  handleStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`KPL dashboard running at http://localhost:${PORT}/`);
});
