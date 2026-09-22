/**
 * Chạy 1 LẦN (hoặc mỗi khi thêm/sửa 1 loại báo cáo) để tạo/cập nhật các
 * document trong collection "reportTypes" - đây là thứ webapp đọc để
 * vẽ sidebar và biết cách vẽ biểu đồ cho từng loại báo cáo.
 *
 * Chạy: FIREBASE_SERVICE_ACCOUNT='...' node seed-report-types.js
 */
const admin = require('firebase-admin');

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

// Thêm 1 object vào mảng này mỗi khi có 1 báo cáo mới muốn hiện trên dashboard.
const REPORT_TYPES = [
  {
    id: 'doanh-thu-tong-quan',
    order: 1,
    label: 'Doanh thu tổng quan',
    chartType: 'line', // vẽ theo thời gian
    dimensionKey: 'date',
    metricKeys: ['revenue', 'orders', 'customers'],
    metricLabels: { revenue: 'Doanh thu', orders: 'Số hoá đơn', customers: 'Số khách' },
    units: { revenue: 'đ' },
  },
  {
    id: 'ban-hang-theo-cua-hang',
    order: 2,
    label: 'Bán hàng theo cửa hàng',
    chartType: 'bar', // so sánh giữa các cửa hàng
    dimensionKey: 'storeName',
    metricKeys: ['revenue', 'quantity'],
    metricLabels: { revenue: 'Doanh thu', quantity: 'Số lượng bán' },
    units: { revenue: 'đ' },
  },
];

(async () => {
  const batch = db.batch();
  REPORT_TYPES.forEach((rt) => {
    const { id, ...data } = rt;
    batch.set(db.collection('reportTypes').doc(id), data, { merge: true });
  });
  await batch.commit();
  console.log(`✅ Đã tạo/cập nhật ${REPORT_TYPES.length} reportTypes.`);
})();
