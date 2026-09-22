/**
 * Dựng d05ItemStores/{ngày} (số bán từng món × từng cửa hàng) từ d05Daily đã có.
 * Chạy 1 lần cho dữ liệu cũ; từ nay sync-report.js/import-file.js tự ghi kèm (lib.js).
 * Web dùng doc này cho bảng "Món theo chi nhánh" - 1 lượt đọc mỗi ngày thay vì ~90.
 *
 *   node build-item-stores.js --dry-run
 *   node build-item-stores.js [--from 2026-01-01] [--to 2026-09-22]
 */

const admin = require('firebase-admin');
const { loadEnv, initFirebase, itemStoresDoc } = require('./lib');

const arg = (n) => { const i = process.argv.indexOf(n); return i > 0 ? process.argv[i + 1] : null; };

async function main() {
  loadEnv();
  const dryRun = process.argv.includes('--dry-run');
  const from = arg('--from') || '2000-01-01';
  const to = arg('--to') || '2999-12-31';
  const db = initFirebase();

  const dates = (await db.collection('syncState').get()).docs
    .map((d) => d.id.replace(/^d05_/, '')).filter((d) => d >= from && d <= to).sort();
  console.log(`${dates.length} ngày có dữ liệu Fabi: ${dates[0]} → ${dates[dates.length - 1]}`);

  let written = 0;
  for (const date of dates) {
    const snap = await db.collection('d05Daily').where('date', '==', date).get();
    if (snap.empty) continue;
    const dayDocs = {};
    snap.forEach((d) => (dayDocs[d.id] = d.data()));
    const doc = itemStoresDoc(date, dayDocs);
    const size = doc.payload.length;
    const nItems = doc.items;
    console.log(`  ${date}: ${snap.size} cửa hàng, ${nItems} món, ~${Math.round(size / 1024)} KB`
      + (size > 900000 ? '  ⚠ gần giới hạn 1 MB của Firestore' : ''));
    if (dryRun) continue;
    await db.collection('d05ItemStores').doc(date).set({ ...doc, updatedAt: admin.firestore.FieldValue.serverTimestamp() });
    written++;
  }
  console.log(dryRun ? '(chạy thử, chưa ghi gì)' : `✅ đã ghi ${written} ngày vào d05ItemStores`);
}

main().then(() => process.exit(0)).catch((e) => { console.error('❌ Lỗi:', e.message); process.exit(1); });
