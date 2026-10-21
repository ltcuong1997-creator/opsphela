// Đọc + gộp file D05 trong worker riêng: XLSX đọc đồng bộ, file lỗi có thể làm treo cả tiến
// trình (21/10/2026: file ngày 20/10 treo cả tiếng, chặn luôn tab khác đang ghi Firestore).
// Chạy ở đây thì tiến trình chính vẫn hẹn giờ và huỷ được.
const { parentPort, workerData } = require('worker_threads');
const XLSX = require('xlsx');
const { rowsFromWorkbook } = require('../public/d05-core.js');
const { aggregate } = require('./lib');

// Log gửi qua parentPort (console.log của worker có thể chưa kịp in nếu worker bị huỷ).
const log = (msg) => parentPort.postMessage({ log: msg });
const t0 = Date.now();
log('  bắt đầu mở file…');
// cellText/cellHTML: tắt dựng chữ định dạng cho TỪNG ô (không dùng tới - sheet_to_json lấy
// giá trị gốc) - định dạng số lạ có thể làm bước này chạy rất lâu với file lớn.
const wb = XLSX.readFile(workerData.filePath, { cellDates: true, cellText: false, cellHTML: false, cellStyles: false });
const refs = wb.SheetNames.filter((n) => /^\d+-/.test(n)).map((n) => wb.Sheets[n]['!ref']);
log(`  mở file ${((Date.now() - t0) / 1000).toFixed(1)}s, ${wb.SheetNames.length} sheet, vùng lớn nhất ${refs.sort((a, b) => b.length - a.length)[0] || '-'}`);
const rows = rowsFromWorkbook(wb, XLSX);
log(`  đọc dòng ${((Date.now() - t0) / 1000).toFixed(1)}s`);
parentPort.postMessage({ docs: aggregate(rows), rows: rows.length });
