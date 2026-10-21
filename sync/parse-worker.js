// Đọc + gộp file D05 trong worker riêng: XLSX đọc đồng bộ, file lỗi có thể làm treo cả tiến
// trình (21/10/2026: file ngày 20/10 treo cả tiếng, chặn luôn tab khác đang ghi Firestore).
// Chạy ở đây thì tiến trình chính vẫn hẹn giờ và huỷ được.
const { parentPort, workerData } = require('worker_threads');
const { readRows, aggregate } = require('./lib');

const rows = readRows(workerData.filePath);
parentPort.postMessage({ docs: aggregate(rows), rows: rows.length });
