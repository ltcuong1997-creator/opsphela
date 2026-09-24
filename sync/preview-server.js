/**
 * Server xem trước chạy trên máy (chỉ nghe 127.0.0.1): phục vụ thư mục ../public
 * và cho trang web đọc Firestore bằng service account trong sync/.env - để xem
 * thử giao diện khi chưa tạo tài khoản đăng nhập.
 *   /api/days?from&to  tóm tắt ngày (d05Days) - trang web dùng cái này
 *   /api/d05?from&to   chi tiết từng cửa hàng-ngày (d05Daily, nặng)
 * Bản deploy thật trên Firebase Hosting KHÔNG dùng file này: ở đó trang web
 * đăng nhập Firebase Auth rồi đọc thẳng Firestore.
 *
 *   node preview-server.js      -> http://localhost:5180
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { loadEnv, initFirebase } = require('./lib');

const PORT = Number(process.env.PORT) || 5180;
const PUBLIC_DIR = path.join(__dirname, '..', 'public');
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon' };
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

loadEnv();
const db = initFirebase();

async function d05(from, to) {
  const snap = await db.collection('d05Daily').where('date', '>=', from).where('date', '<=', to).get();
  return snap.docs.map((d) => {
    const { updatedAt, hash, ...rest } = d.data();
    return { id: d.id, ...rest, updatedAt: updatedAt ? updatedAt.toDate().toISOString() : null };
  });
}

// Doc tóm tắt ngày (d05Days) - trang web dùng cái này, nhẹ hơn d05Daily nhiều
async function days(from, to) {
  const snap = await db.collection('d05Days').where('date', '>=', from).where('date', '<=', to).get();
  return snap.docs.map((d) => {
    const { updatedAt, ...rest } = d.data();
    return { ...rest, updatedAt: updatedAt ? updatedAt.toDate().toISOString() : null };
  });
}

// Số bán từng món × từng cửa hàng (d05ItemStores) - bảng "Món theo chi nhánh"
async function itemStores(from, to) {
  const snap = await db.collection('d05ItemStores').where('date', '>=', from).where('date', '<=', to).get();
  return snap.docs.map((d) => {
    const { updatedAt, ...rest } = d.data();
    return rest;
  });
}

// Đọc toàn bộ, xem trước trang "Chi nhánh" / "Target" - không ghi được ở chế độ local
// (không có Firebase Auth thật để qua firestore.rules), chỉ để xem trước giao diện.
async function readAll(collection) {
  const snap = await db.collection(collection).get();
  return snap.docs.map((d) => {
    const { updatedAt, ...rest } = d.data();
    return { id: d.id, ...rest, updatedAt: updatedAt ? updatedAt.toDate().toISOString() : null };
  });
}

function send(res, code, body, type) {
  res.writeHead(code, { 'Content-Type': type || 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(body);
}

http
  .createServer(async (req, res) => {
    const url = new URL(req.url, 'http://localhost');
    try {
      if (url.pathname === '/api/days' || url.pathname === '/api/d05') {
        const from = url.searchParams.get('from');
        const to = url.searchParams.get('to');
        if (!DATE_RE.test(from) || !DATE_RE.test(to)) return send(res, 400, '{"error":"from/to phải dạng YYYY-MM-DD"}');
        return send(res, 200, JSON.stringify(url.pathname === '/api/days' ? await days(from, to) : await d05(from, to)));
      }
      if (url.pathname === '/api/itemstores') {
        const from = url.searchParams.get('from');
        const to = url.searchParams.get('to');
        if (!DATE_RE.test(from) || !DATE_RE.test(to)) return send(res, 400, '{"error":"from/to phải dạng YYYY-MM-DD"}');
        return send(res, 200, JSON.stringify(await itemStores(from, to)));
      }
      if (url.pathname === '/api/cstickets') {
        const from = url.searchParams.get('from');
        const to = url.searchParams.get('to');
        if (!DATE_RE.test(from) || !DATE_RE.test(to)) return send(res, 400, '{"error":"from/to phải dạng YYYY-MM-DD"}');
        const snap = await db.collection('csTickets').where('date', '>=', from).where('date', '<=', to).get();
        return send(res, 200, JSON.stringify(snap.docs.map((d) => { const { importedAt, ...x } = d.data(); return x; })));
      }
      if (url.pathname === '/api/config') return send(res, 200, JSON.stringify(await readAll('config')));
      if (url.pathname === '/api/stores') return send(res, 200, JSON.stringify(await readAll('stores')));
      if (url.pathname === '/api/targets') return send(res, 200, JSON.stringify(await readAll('targets')));
      const file = path.normalize(path.join(PUBLIC_DIR, url.pathname === '/' ? 'index.html' : url.pathname));
      if (!file.startsWith(PUBLIC_DIR)) return send(res, 403, 'forbidden', 'text/plain');
      fs.readFile(file, (err, buf) => {
        // SPA: đường dẫn lạ thì trả index.html, giống rewrite trong firebase.json
        if (err) return fs.readFile(path.join(PUBLIC_DIR, 'index.html'), (e2, idx) => send(res, e2 ? 404 : 200, idx || 'not found', TYPES['.html']));
        send(res, 200, buf, TYPES[path.extname(file)] || 'application/octet-stream');
      });
    } catch (e) {
      console.error(e);
      send(res, 500, JSON.stringify({ error: e.message }));
    }
  })
  .listen(PORT, '127.0.0.1', () => console.log(`Xem trước: http://localhost:${PORT}`));
