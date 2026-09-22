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

// Tên cột trong file Excel D05 (khác bảng trên web: "Ngày" thay cho "Thời
// gian", "Hoá đơn" thay cho "Mã hoá đơn", và có thêm cột "Cửa hàng").
const COL = {
  time: 'Ngày',
  bill: 'Hoá đơn',
  itemCode: 'Mã hàng',
  itemName: 'Tên hàng',
  group: 'Nhóm món',
  quantity: 'Số lượng',
  revenue: 'Tổng tiền',
  store: 'Cửa hàng',
};

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

// File D05 "tất cả cửa hàng": mỗi cửa hàng có bán là 1 sheet, tên sheet dạng
// "108654-Phê La – 145 Trích..." (mã cửa hàng trên Fabi + tên bị cắt). Dòng
// cuối mỗi sheet là dòng tổng cộng (không có Ngày) - aggregate() tự bỏ qua.
// Ngoài ra có 1 sheet "Tất cả cửa hàng" chứa LẠI toàn bộ dòng của các sheet
// trên - phải bỏ, không thì doanh thu bị nhân đôi. Nếu file chỉ có đúng sheet
// đó (xuất 1 cửa hàng / kiểu khác) thì mới dùng nó.
function readRows(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });
  const storeSheets = workbook.SheetNames.filter((n) => /^\d+-/.test(n));
  const sheets = storeSheets.length ? storeSheets : workbook.SheetNames;
  return sheets.flatMap((name) => {
    const storeId = name.match(/^(\d+)-/)?.[1] || null;
    return XLSX.utils
      .sheet_to_json(workbook.Sheets[name], { defval: null })
      .map((row) => ({ ...row, __storeId: storeId }));
  });
}

function toNumber(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  // VND không có phần thập phân; dấu "," / "." chỉ là phân cách hàng nghìn.
  return Number(String(v).replace(/[^0-9-]/g, '')) || 0;
}

function normalizeDate(v) {
  // "22/09/2026 06:05" (dd/mm/yyyy) - new Date() sẽ hiểu nhầm thành mm/dd.
  if (!v) return null;
  if (v instanceof Date) {
    if (isNaN(v)) return null;
    const p = (n) => String(n).padStart(2, '0');
    return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
  }
  const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const iso = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return iso ? iso[0] : null;
}

function keyOf(s) {
  // Firestore doc id / field name an toàn: bỏ dấu, chỉ giữ chữ số và "-".
  return String(s)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/gi, 'd')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '') || 'khong-ro';
}

// rows (từng dòng món trong hoá đơn) -> { [docId]: doc } gộp theo ngày × cửa hàng.
function aggregate(rows) {
  const docs = {};
  const bills = {};
  for (const row of rows) {
    const date = normalizeDate(row[COL.time]);
    if (!date) continue;
    const storeName = String(row[COL.store] || 'Không rõ').trim();
    const storeId = row.__storeId || keyOf(storeName);
    const id = `${date}_${storeId}`;
    const doc = (docs[id] ||= { date, storeId, storeName, revenue: 0, quantity: 0, bills: 0, items: {} });
    const qty = toNumber(row[COL.quantity]);
    const rev = toNumber(row[COL.revenue]);
    doc.revenue += rev;
    doc.quantity += qty;
    (bills[id] ||= new Set()).add(row[COL.bill]);

    const code = row[COL.itemCode] || row[COL.itemName];
    const item = (doc.items[keyOf(code)] ||= {
      code: row[COL.itemCode] || null,
      name: row[COL.itemName] || null,
      group: row[COL.group] || null,
      quantity: 0,
      revenue: 0,
    });
    item.quantity += qty;
    item.revenue += rev;
  }
  for (const id in docs) docs[id].bills = bills[id].size;
  return docs;
}

function hashOf(doc) {
  const items = Object.keys(doc.items).sort().map((k) => [k, doc.items[k]]);
  const stable = JSON.stringify([doc.date, doc.storeId, doc.storeName, doc.revenue, doc.quantity, doc.bills, items]);
  return crypto.createHash('sha1').update(stable).digest('hex').slice(0, 16);
}

// Chỉ ghi các doc có hash khác với lần ghi trước. Trả về số doc đã ghi.
async function writeChanged(db, docs) {
  const byDate = {};
  for (const [id, doc] of Object.entries(docs)) (byDate[doc.date] ||= {})[id] = doc;

  let written = 0;
  let skipped = 0;
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

    // Batch tối đa 500 thao tác; chừa 1 chỗ cho syncState.
    for (let i = 0; i < changed.length; i += 450) {
      const batch = db.batch();
      for (const [id, doc] of changed.slice(i, i + 450)) {
        batch.set(db.collection(COLLECTION).doc(id), {
          ...doc,
          updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        });
      }
      if (i + 450 >= changed.length) batch.set(stateRef, { hashes: newHashes });
      await batch.commit();
    }
    written += changed.length;
  }
  return { written, skipped };
}

function summarize(docs) {
  const list = Object.values(docs);
  const dates = [...new Set(list.map((d) => d.date))].sort();
  const revenue = list.reduce((s, d) => s + d.revenue, 0);
  return `${list.length} cửa hàng-ngày, ngày ${dates[0]} → ${dates[dates.length - 1]}, tổng tiền ${revenue.toLocaleString('vi-VN')} ₫`;
}

module.exports = { loadEnv, initFirebase, readRows, aggregate, writeChanged, summarize, COL };
