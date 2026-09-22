/**
 * Dựng lại doc tóm tắt d05Days/{ngày} từ d05Daily cho mọi ngày đã nạp (hoặc các
 * ngày chỉ định). Chạy 1 lần sau khi thêm d05Days; sync-report / import-file tự
 * cập nhật d05Days từ đó về sau.
 *
 *   node rebuild-days.js                     tất cả ngày có trong syncState
 *   node rebuild-days.js 2026-09-01 2026-09-02
 */
const { loadEnv, initFirebase, rebuildDaySummaries } = require('./lib');

(async () => {
  loadEnv();
  const db = initFirebase();
  let dates = process.argv.slice(2);
  if (!dates.length) {
    const snap = await db.collection('syncState').get();
    dates = snap.docs.map((d) => d.id.replace(/^d05_/, '')).sort();
  }
  await rebuildDaySummaries(db, dates);
})().catch((e) => {
  console.error('❌ Lỗi:', e.message);
  process.exit(1);
});
