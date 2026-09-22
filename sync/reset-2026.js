/**
 * Chuyển năm 2026 sang dùng 100% dữ liệu Fabi:
 *  1. Xoá các doc d05Days của 2026 đang mang số Excel (source 'excel' / 'excel+fabi')
 *  2. Xoá histDays của 2026 (2024-2025 giữ nguyên - vẫn là nguồn cho d05Days trước mốc)
 *  3. Dựng lại d05Days từ d05Daily cho những ngày Fabi đã nạp
 * Chạy lại được nhiều lần. Ngày Fabi chưa có sẽ trống cho tới khi nạp bù xong.
 *
 *   node reset-2026.js --dry-run
 *   node reset-2026.js
 */

const admin = require('firebase-admin');
const { loadEnv, initFirebase, daySummary } = require('./lib');

const FROM = '2026-01-01';

// Mạng chậm hay ném DEADLINE_EXCEEDED giữa chừng -> thử lại từng mẻ nhỏ.
async function retry(label, fn, times = 5) {
  for (let i = 1; ; i++) {
    try { return await fn(); } catch (e) {
      if (i >= times) throw e;
      console.log(`  thử lại ${label} (lần ${i}): ${String(e.message).split('\n')[0].slice(0, 60)}`);
      await new Promise((r) => setTimeout(r, 1500 * i));
    }
  }
}

async function deleteDocs(db, docs, label, dryRun, chunk = 50) {
  console.log(`${label}: ${docs.length} doc`);
  if (dryRun || !docs.length) return;
  for (let i = 0; i < docs.length; i += chunk) {
    const part = docs.slice(i, i + chunk);
    await retry(`xoá ${part[0].id}`, async () => {
      const batch = db.batch();
      part.forEach((d) => batch.delete(d.ref));
      await batch.commit();
    });
  }
  console.log(`  ✅ đã xoá ${docs.length}`);
}

async function main() {
  loadEnv();
  const dryRun = process.argv.includes('--dry-run');
  const db = initFirebase();

  const days = await db.collection('d05Days').where('date', '>=', FROM).select('source').get();
  await deleteDocs(db, days.docs.filter((d) => d.data().source), 'd05Days 2026 mang số Excel', dryRun);

  const hist = await db.collection('histDays').where('date', '>=', FROM).select().get();
  // doc histDays nặng (~60 KB) -> mẻ nhỏ, không thì Firestore báo "Transaction too big"
  await deleteDocs(db, hist.docs, 'histDays 2026 (giữ lại 2024-2025)', dryRun, 8);

  const dates = (await db.collection('syncState').select().get()).docs
    .map((d) => d.id.replace(/^d05_/, '')).filter((d) => d >= FROM).sort();
  console.log(`Dựng lại d05Days từ d05Daily: ${dates.length} ngày`);
  if (!dryRun) {
    for (const date of dates) {
      const snap = await retry(`đọc ${date}`, () => db.collection('d05Daily').where('date', '==', date).get());
      if (snap.empty) continue;
      const dayDocs = {};
      snap.forEach((d) => (dayDocs[d.id] = d.data()));
      await retry(`ghi ${date}`, () => db.collection('d05Days').doc(date).set({
        ...daySummary(date, dayDocs), updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      }));
      process.stdout.write(`  ${date} (${snap.size} CH)\n`);
    }
  }

  const left = await db.collection('d05Days').where('date', '>=', FROM).select('source').get();
  const src = {};
  left.forEach((d) => { const k = d.data().source || 'fabi'; src[k] = (src[k] || 0) + 1; });
  console.log(`Còn lại trên d05Days trong 2026: ${left.size} ngày ${JSON.stringify(src)}`);
}

main().then(() => process.exit(0)).catch((e) => { console.error('❌ Lỗi:', e.message); process.exit(1); });
