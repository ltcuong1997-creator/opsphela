// Đọc + gộp file D05 trong worker riêng: XLSX đọc đồng bộ, file lỗi có thể làm treo cả tiến
// trình (21/10/2026: file ngày 20/10 treo cả tiếng, chặn luôn tab khác đang ghi Firestore).
// Chạy ở đây thì tiến trình chính vẫn hẹn giờ và huỷ được.
const { parentPort, workerData } = require('worker_threads');
const XLSX = require('xlsx');
const { rowsFromWorkbook } = require('../public/d05-core.js');
const { aggregate } = require('./lib');

const t0 = Date.now();
const wb = XLSX.readFile(workerData.filePath, { cellDates: true });
const refs = wb.SheetNames.filter((n) => /^\d+-/.test(n)).map((n) => wb.Sheets[n]['!ref']);
console.log(`  mở file ${((Date.now() - t0) / 1000).toFixed(1)}s, ${wb.SheetNames.length} sheet, vùng lớn nhất ${refs.sort((a, b) => b.length - a.length)[0] || '-'}`);
const rows = rowsFromWorkbook(wb, XLSX);
console.log(`  đọc dòng ${((Date.now() - t0) / 1000).toFixed(1)}s`);
parentPort.postMessage({ docs: aggregate(rows), rows: rows.length });
