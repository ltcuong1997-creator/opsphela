/**
 * Phần dùng chung cho sync-report.js (tự động, hôm nay) và import-file.js
 * (nạp tay file Excel D05 cũ): đọc file, gộp ngày × cửa hàng × món, và chỉ
 * ghi lên Firestore những document thật sự thay đổi.
 *
 * Cấu trúc Firestore:
 *   d05Daily/{date}_{storeId}    1 doc = 1 cửa hàng trong 1 ngày (storeId = mã Fabi)
 *     { date, storeId, storeName, revenue, quantity, bills, hash, updatedAt,
 *       items: { [itemKey]: { code, name, group, quantity, revenue } } }
 *   syncState/d05_{date}         { hashes: { [docId]: hash } }
 *   d05ItemStores/{date}         số bán từng món × từng cửa hàng trong ngày, cho bảng
 *     "Món theo chi nhánh" trên web (1 doc/ngày thay vì ~90 doc d05Daily)
 *   d05Days/{date}               tóm tắt cả ngày cho trang web: tổng ngày, tổng từng
 *     cửa hàng (stores{}), tổng từng món cả chuỗi (items{}) - web đọc 1 doc/ngày
 *     thay vì ~90 doc nặng, xem 1 tháng chỉ tốn ~30 lượt đọc.
 *
 * Mốc EXCEL_UNTIL: các ngày tới hết mốc lấy số từ báo cáo nội bộ Excel (hist-to-days.js
 * chép histDays -> d05Days); Fabi chỉ tự nạp các ngày SAU mốc. Để cùng cách tính với
 * Excel và target, d05Days của ngày sau mốc ghi doanh thu TRƯỚC VAT:
 *   revenue = (doanh thu Fabi + hoa hồng app) / 1,08     (số gốc giữ ở revenueGross)
 * Riêng doanh thu từng món (items) vẫn là số gốc Fabi.
 *
 * Mỗi lần ghi, đọc 1 doc syncState của ngày đó (1 lượt đọc), so hash từng
 * cửa hàng, chỉ ghi các cửa hàng có số liệu khác đi + cập nhật lại syncState.
 * Không có gì đổi thì không ghi gì cả.
 */

const XLSX = require('xlsx');
const admin = require('firebase-admin');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const COLLECTION = 'd05Daily';
const STATE_COLLECTION = 'syncState';
const DAYS_COLLECTION = 'd05Days';
const ITEM_STORES_COLLECTION = 'd05ItemStores';
const STORES_COLLECTION = 'stores';

const {
  EXCEL_UNTIL, COL, ALIASES, ET_KEYS, rowsFromWorkbook, keyOf, aggregate, hashInput, toNet, itemStoresDoc, daySummary,
} = require('../public/d05-core.js');


function loadEnv() {
  const envFile = path.join(__dirname, '.env');
  if (fs.existsSync(envFile)) process.loadEnvFile(envFile);
}

function initFirebase() {
  let json = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (!json && process.env.FIREBASE_SERVICE_ACCOUNT_FILE) {
    json = fs.readFileSync(process.env.FIREBASE_SERVICE_ACCOUNT_FILE, 'utf8');
  }
  if (!json) throw new Error('Thiếu FIREBASE_SERVICE_ACCOUNT hoặc FIREBASE_SERVICE_ACCOUNT_FILE');
  admin.initializeApp({ credential: admin.credential.cert(JSON.parse(json)) });
  return admin.firestore();
}

// Đọc file Excel D05 -> các dòng (xem rowsFromWorkbook trong public/d05-core.js)
function readRows(filePath) {
  return rowsFromWorkbook(XLSX.readFile(filePath, { cellDates: true }), XLSX);
}

function hashOf(doc) {
  return crypto.createHash('sha1').update(hashInput(doc)).digest('hex').slice(0, 16);
}

// Mạng chậm hay ném DEADLINE_EXCEEDED giữa chừng: thử lại vài lần trước khi bỏ ngày đó.
// Firestore set() là ghi đè nguyên doc nên chạy lại an toàn, không nhân đôi dữ liệu.
async function withRetry(label, fn, times = 4) {
  for (let i = 1; ; i++) {
    try { return await fn(); } catch (e) {
      if (i >= times) throw e;
      console.log(`  thử lại ${label} (lần ${i}): ${String(e.message).split('\n')[0].slice(0, 60)}`);
      await new Promise((r) => setTimeout(r, 2000 * i));
    }
  }
}

// Chỉ ghi các doc có hash khác với lần ghi trước. Trả về số doc đã ghi.
async function writeChanged(db, docs) {
  const byDate = {};
  for (const [id, doc] of Object.entries(docs)) (byDate[doc.date] ||= {})[id] = doc;

  let written = 0;
  let skipped = 0;
  const changedDocs = [];
  const locked = Object.keys(byDate).filter((d) => d <= EXCEL_UNTIL);
  if (locked.length && !process.argv.includes('--force')) {
    throw new Error(`Ngày ${locked.join(', ')} nằm trong khoảng lấy số từ Excel (tới ${EXCEL_UNTIL}) - không nạp Fabi đè lên. `
      + 'Thêm --force nếu chỉ muốn lưu món/PTTT gốc vào d05Daily (d05Days vẫn giữ số Excel).');
  }
  for (const [date, dayDocs] of Object.entries(byDate)) {
    const stateRef = db.collection(STATE_COLLECTION).doc(`d05_${date}`);
    const oldHashes = (await stateRef.get()).data()?.hashes || {};
    const newHashes = { ...oldHashes };
    const changed = [];
    for (const [id, doc] of Object.entries(dayDocs)) {
      const h = hashOf(doc);
      newHashes[id] = h;
      if (oldHashes[id] === h) skipped++;
      else changed.push([id, { ...doc, hash: h }]);
    }
    if (!changed.length) continue;

    // Tóm tắt ngày cần đủ mọi cửa hàng của ngày đó. File xuất "tất cả cửa hàng" thì
    // dayDocs đã đủ; lỡ chỉ nạp 1 phần (file vài cửa hàng) thì đọc bù phần còn lại.
    let allDay = dayDocs;
    const missing = Object.keys(oldHashes).filter((id) => !dayDocs[id]);
    if (missing.length) {
      const snap = await db.collection(COLLECTION).where('date', '==', date).get();
      allDay = {};
      snap.forEach((d) => (allDay[d.id] = d.data()));
      Object.assign(allDay, dayDocs);
    }
    const summary = daySummary(date, allDay);

    // Chia batch 40 doc - mỗi doc cửa hàng kèm danh sách món nên batch to hay bị
    // DEADLINE_EXCEEDED khi mạng chậm. Hai doc nặng nhất (tóm tắt ngày ~300 KB và
    // món × cửa hàng ~130 KB) tách ra ghi riêng, KHÔNG nhét chung batch cuối.
    for (let i = 0; i < changed.length; i += 40) {
      const part = changed.slice(i, i + 40);
      await withRetry(`${date}: ghi ${part.length} cửa hàng`, async () => {
        const batch = db.batch();
        for (const [id, doc] of part) {
          batch.set(db.collection(COLLECTION).doc(id), {
            ...doc,
            updatedAt: admin.firestore.FieldValue.serverTimestamp(),
          });
        }
        await batch.commit();
      });
    }
    if (date > EXCEL_UNTIL) {
      await withRetry(`${date}: tóm tắt ngày`, () => db.collection(DAYS_COLLECTION).doc(date).set({
        ...summary,
        updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      }));
    }
    // Món × cửa hàng: chỉ Fabi mới có, không dính mốc Excel nên ngày nào cũng ghi
    await withRetry(`${date}: món × cửa hàng`, () => db.collection(ITEM_STORES_COLLECTION).doc(date).set({
      ...itemStoresDoc(date, allDay),
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    }));
    // syncState ghi SAU CÙNG: có nó nghĩa là cả ngày đã vào đủ, nên --skip-done /
    // --skip-complete mới tin được.
    await withRetry(`${date}: syncState`, () => stateRef.set({ hashes: newHashes }));
    written += changed.length;
    for (const [, doc] of changed) changedDocs.push(doc);
  }

  // Danh mục cửa hàng dùng cho trang "Chi nhánh" (gán Region/City/AM). Merge để không
  // đụng field admin đã gán trên web - script chỉ giữ storeId/storeName luôn đủ và đúng.
  // Chỉ các cửa hàng vừa có doc thay đổi: cửa hàng mới / đổi tên luôn làm hash đổi nên
  // vẫn được cập nhật, còn lần chạy không có gì mới thì không tốn thêm lượt ghi nào.
  const stores = new Map();
  for (const doc of changedDocs) stores.set(doc.storeId, doc.storeName);
  if (stores.size) {
    const catalogBatch = db.batch();
    for (const [storeId, storeName] of stores) {
      catalogBatch.set(db.collection(STORES_COLLECTION).doc(storeId), {
        storeId, storeName, updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      }, { merge: true });
    }
    await catalogBatch.commit();
  }

  return { written, skipped };
}

// Dựng lại d05Days cho các ngày đã có trong d05Daily (dùng 1 lần khi mới thêm d05Days,
// hoặc khi nghi ngờ tóm tắt lệch). Mỗi ngày: đọc ~90 doc, ghi 1 doc.
async function rebuildDaySummaries(db, dates) {
  for (const date of dates) {
    if (date <= EXCEL_UNTIL) { console.log(`${date}: bỏ qua - ngày lấy số từ Excel`); continue; }
    const snap = await db.collection(COLLECTION).where('date', '==', date).get();
    if (snap.empty) continue;
    const dayDocs = {};
    snap.forEach((d) => (dayDocs[d.id] = d.data()));
    await db.collection(DAYS_COLLECTION).doc(date).set({
      ...daySummary(date, dayDocs),
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    console.log(`${date}: ${snap.size} cửa hàng`);
  }
}

function summarize(docs) {
  const list = Object.values(docs);
  const dates = [...new Set(list.map((d) => d.date))].sort();
  const revenue = list.reduce((s, d) => s + d.revenue, 0);
  return `${list.length} cửa hàng-ngày, ngày ${dates[0]} → ${dates[dates.length - 1]}, tổng tiền ${revenue.toLocaleString('vi-VN')} ₫`;
}

// { [tên cửa hàng không dấu]: mã Fabi } từ danh mục stores - dùng khi nạp file không có mã.
async function loadStoreIds(db) {
  const snap = await db.collection(STORES_COLLECTION).get();
  const map = {};
  snap.forEach((d) => { if (d.data().storeName) map[keyOf(d.data().storeName)] = d.id; });
  return map;
}

module.exports = { EXCEL_UNTIL, ET_KEYS, toNet, daySummary, itemStoresDoc, loadStoreIds, ALIASES, loadEnv, initFirebase, readRows, aggregate, writeChanged, summarize, rebuildDaySummaries, COL };
