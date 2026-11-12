const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const orderDir = path.join(__dirname, 'Order');
console.log('🔄 เริ่มต้นประมวลผลข้อมูลจากโฟลเดอร์ Order...');

if (!fs.existsSync(orderDir)) {
  console.error('❌ ไม่พบโฟลเดอร์ Order กรุณาตรวจสอบว่าโฟลเดอร์วางอยู่ระดับเดียวกับไฟล์นี้');
  process.exit(1);
}

const files = fs.readdirSync(orderDir).filter(f => f.endsWith('.xlsx') || f.endsWith('.csv') || f.endsWith('.xls'));

function cleanSellerSku(rawSku) {
  if (!rawSku || rawSku === '-' || String(rawSku).trim() === '') return 'Uncategorized';
  const str = String(rawSku).trim();
  if (str.includes('_')) {
    const parts = str.split('_');
    parts.shift();
    return parts.join('_').trim() || str;
  }
  return str;
}

function parseDateParts(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    const dObj = xlsx.SSF.parse_date_code(val);
    if (dObj) return { year: String(dObj.y), month: String(dObj.m).padStart(2, '0'), day: String(dObj.d).padStart(2, '0') };
  }
  const str = String(val).trim();
  const dmyMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmyMatch) return { year: dmyMatch[3], month: dmyMatch[2].padStart(2, '0'), day: dmyMatch[1].padStart(2, '0') };
  const d = new Date(str);
  if (!isNaN(d.getTime())) return { year: String(d.getFullYear()), month: String(d.getMonth() + 1).padStart(2, '0'), day: String(d.getDate()).padStart(2, '0') };
  return null;
}

const timeline = {};
const topProductsMap = {};
const auctionMap = {};
const gwpMap = {};
const allYears = new Set();

for (const file of files) {
  console.log(`📂 กำลังอ่านไฟล์: ${file}`);
  const wb = xlsx.readFile(path.join(orderDir, file), { dense: true, cellStyles: false, cellHTML: false, cellFormula: false });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });

  for (const row of rows) {
    let orderId = '', buyer = '', productName = '', sku = '-', sellerSku = '-', quantity = 1;
    let dateRaw = '', amount = 0, itemPaid = null, statusRaw = '';

    for (const k of Object.keys(row)) {
      const kl = k.trim().toLowerCase().replace(/[\s_\-]/g, '');
      const v = row[k];
      if (!productName && (kl.includes('productname') || kl.includes('ชื่อสินค้า'))) productName = String(v).trim();
      if (!sellerSku && (kl.includes('sellersku') || kl.includes('skuผู้ขาย'))) sellerSku = String(v).trim();
      if (!sku && (kl.includes('skuid') || kl.includes('รหัสsku') || kl === 'sku')) sku = String(v).trim();
      if (!orderId && (kl.includes('orderid') || kl.includes('หมายเลขคำสั่งซื้อ') || kl === 'id')) orderId = String(v).trim();
      if (!buyer && (kl.includes('buyer') || kl.includes('ผู้ซื้อ') || kl.includes('username'))) buyer = String(v).trim();
      if (kl.includes('quantity') || kl.includes('จำนวน') || kl === 'qty') { const q = parseInt(v); if (!isNaN(q) && q > 0) quantity = q; }
      if (!dateRaw && (kl.includes('time') || kl.includes('date') || kl.includes('วันที่'))) dateRaw = v;
      if (itemPaid === null && (kl.includes('skusubtotalafterdiscount') || kl.includes('ราคาหลังหักส่วนลด'))) itemPaid = parseFloat(String(v).replace(/[^0-9.-]+/g, '')) || 0;
      if (!amount && (kl.includes('orderamount') || kl.includes('ยอดรวมของคำสั่งซื้อ') || kl === 'gmv')) amount = parseFloat(String(v).replace(/[^0-9.-]+/g, '')) || 0;
      if (!statusRaw && (kl.includes('orderstatus') || kl.includes('สถานะ'))) statusRaw = String(v).trim();
    }

    const dParts = parseDateParts(dateRaw) || { year: '2026', month: '01', day: '01' };
    const cleanAmount = amount;
    const paid = itemPaid !== null ? itemPaid : cleanAmount;
    const isCanceled = statusRaw.toLowerCase().includes('cancel') || statusRaw.includes('ยกเลิก');
    const isCompleted = !isCanceled;

    const pLower = productName.toLowerCase();
    const sLowerSku = String(sellerSku).toLowerCase();
    const cleanSku = cleanSellerSku(sellerSku);

    // ตรวจสอบกลุ่มของแถม (GWP)
    const isExplicitGwp = (
      cleanSku.toUpperCase() === '1D' || cleanSku.toUpperCase() === '1DAY' || cleanSku.toUpperCase() === 'WTP' ||
      cleanSku.toUpperCase().includes('INF1') || cleanSku.toUpperCase().includes('GWP') ||
      sLowerSku.includes('wtp') || sLowerSku.includes('inf1') || sLowerSku.includes('1day') ||
      pLower.includes('1day') || /\b1d\b/.test(pLower)
    );

    const isGWP = (
      paid === 0 || cleanAmount === 0 || isExplicitGwp ||
      sLowerSku.includes('gwp') || pLower.includes('ซองกันน้ำ') || pLower.includes('ของแถม')
    );

    // ตรวจสอบกลุ่มประมูล (นาทีทอง / ประมูล / action)
    const isAuction = !isGWP && (
      pLower.includes('นาทีทอง') || 
      pLower.includes('ประมูล') || 
      pLower.includes('action')
    );

    const yr = dParts.year, mo = dParts.month, dy = dParts.day;
    allYears.add(yr);

    if (!timeline[yr]) timeline[yr] = { compOrders: new Set(), cancOrders: new Set(), compGmv: 0, cancGmv: 0, buyers: new Set(), months: {} };
    if (!timeline[yr].months[mo]) timeline[yr].months[mo] = { compOrders: new Set(), cancOrders: new Set(), compGmv: 0, cancGmv: 0, buyers: new Set(), days: {} };
    if (!timeline[yr].months[mo].days[dy]) timeline[yr].months[mo].days[dy] = { compOrders: new Set(), cancOrders: new Set(), compGmv: 0, cancGmv: 0, buyers: new Set() };

    const oId = orderId || `${productName}_${Math.random()}`;
    const bId = buyer || oId;

    [timeline[yr], timeline[yr].months[mo], timeline[yr].months[mo].days[dy]].forEach(node => {
      node.buyers.add(bId);
      if (isCompleted) {
        if (!node.compOrders.has(oId)) {
          node.compOrders.add(oId);
          if (!isGWP) node.compGmv += cleanAmount;
        }
      } else {
        if (!node.cancOrders.has(oId)) {
          node.cancOrders.add(oId);
          if (!isGWP) node.cancGmv += cleanAmount;
        }
      }
    });

    // แยกเก็บลงตารางตามหมวดหมู่
    const targetMap = isGWP ? gwpMap : (isAuction ? auctionMap : topProductsMap);
    const itemKey = `${productName}___${sku}`;
    if (!targetMap[itemKey]) {
      targetMap[itemKey] = { name: productName || 'สินค้าไม่ระบุชื่อ', sku, quantity: 0, completedCount: 0, canceledCount: 0, totalAmount: 0, _seen: new Set() };
    }
    targetMap[itemKey].quantity += quantity;
    if (!targetMap[itemKey]._seen.has(oId)) {
      targetMap[itemKey]._seen.add(oId);
      if (isCompleted) {
        targetMap[itemKey].completedCount++;
        if (!isGWP) targetMap[itemKey].totalAmount += cleanAmount;
      } else {
        targetMap[itemKey].canceledCount++;
      }
    }
  }
}

const cleanTimeline = {};
let grandTotalOrders = 0;

for (const yr of Object.keys(timeline)) {
  const y = timeline[yr];
  grandTotalOrders += y.compOrders.size + y.cancOrders.size;
  cleanTimeline[yr] = {
    compCount: y.compOrders.size, cancCount: y.cancOrders.size, compGmv: Math.round(y.compGmv), cancGmv: Math.round(y.cancGmv), buyersCount: y.buyers.size, months: {}
  };
  for (const mo of Object.keys(y.months)) {
    const m = y.months[mo];
    cleanTimeline[yr].months[mo] = {
      compCount: m.compOrders.size, cancCount: m.cancOrders.size, compGmv: Math.round(m.compGmv), cancGmv: Math.round(m.cancGmv), buyersCount: m.buyers.size, days: {}
    };
    for (const dy of Object.keys(m.days)) {
      const d = m.days[dy];
      cleanTimeline[yr].months[mo].days[dy] = {
        compCount: d.compOrders.size, cancCount: d.cancOrders.size, compGmv: Math.round(d.compGmv), cancGmv: Math.round(d.cancGmv), buyersCount: d.buyers.size
      };
    }
  }
}

const sanitizeList = (map, sortBy) => Object.values(map).map(item => { delete item._seen; return item; }).sort((a, b) => b[sortBy] - a[sortBy]).slice(0, 50);

const finalOutput = {
  totalCount: grandTotalOrders,
  years: Array.from(allYears).sort().reverse(),
  timeline: cleanTimeline,
  topProducts: sanitizeList(topProductsMap, 'totalAmount'),
  auctionProducts: sanitizeList(auctionMap, 'totalAmount'),
  gwpItems: sanitizeList(gwpMap, 'quantity')
};

const outPath = path.join(__dirname, 'summary.json');
fs.writeFileSync(outPath, JSON.stringify(finalOutput));
const kb = (fs.statSync(outPath).size / 1024).toFixed(1);
console.log(`\n🎉 สรุปยอดเสร็จสิ้น! บันทึกเป็น summary.json ขนาด: ${kb} KB (พบสินค้าประมูล ${finalOutput.auctionProducts.length} รายการ)`);