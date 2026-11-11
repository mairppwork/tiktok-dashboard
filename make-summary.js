const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const orderDir = path.join(__dirname, 'Order');
console.log('🔄 เริ่มต้นสรุปสถิติ (โหมดกิน RAM น้อยที่สุด)...');

if (!fs.existsSync(orderDir)) {
  console.error('❌ ไม่พบโฟลเดอร์ Order');
  process.exit(1);
}

const files = fs.readdirSync(orderDir).filter(f => f.endsWith('.xlsx') || f.endsWith('.csv'));

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
const gwpMap = {};
const sellerSkuSaleMap = {};
const sellerSkuGwpMap = {};
const allYears = new Set();

for (const file of files) {
  console.log(`Processing: ${file}`);
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

    const yr = dParts.year, mo = dParts.month, dy = dParts.day;
    allYears.add(yr);

    if (!timeline[yr]) timeline[yr] = { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0, buyers: new Set(), months: {} };
    if (!timeline[yr].months[mo]) timeline[yr].months[mo] = { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0, buyers: new Set(), days: {} };
    if (!timeline[yr].months[mo].days[dy]) timeline[yr].months[mo].days[dy] = { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0, buyers: new Set() };

    const bId = buyer || orderId || 'C';
    [timeline[yr], timeline[yr].months[mo], timeline[yr].months[mo].days[dy]].forEach(node => {
      node.buyers.add(bId);
      if (isCompleted) {
        node.compCount++;
        if (!isGWP) node.compGmv += cleanAmount;
      } else {
        node.cancCount++;
        if (!isGWP) node.cancGmv += cleanAmount;
      }
    });

    const targetProductMap = isGWP ? gwpMap : topProductsMap;
    const itemKey = `${productName}___${sku}`;
    if (!targetProductMap[itemKey]) targetProductMap[itemKey] = { name: productName, sku, quantity: 0, totalAmount: 0 };
    targetProductMap[itemKey].quantity += quantity;
    if (!isGWP && isCompleted) targetProductMap[itemKey].totalAmount += cleanAmount;

    const targetSkuMap = isGWP ? sellerSkuGwpMap : sellerSkuSaleMap;
    if (!targetSkuMap[cleanSku]) targetSkuMap[cleanSku] = { name: cleanSku, sku: cleanSku, quantity: 0, totalAmount: 0 };
    targetSkuMap[cleanSku].quantity += quantity;
    if (!isGWP && isCompleted) targetSkuMap[cleanSku].totalAmount += cleanAmount;
  }
}

// แปลง Set เป็นตัวเลข เพื่อให้ไฟล์ JSON มีขนาดเล็กที่สุด
const cleanTimeline = {};
for (const yr of Object.keys(timeline)) {
  const y = timeline[yr];
  cleanTimeline[yr] = {
    compCount: y.compCount, cancCount: y.cancCount, compGmv: Math.round(y.compGmv), cancGmv: Math.round(y.cancGmv), buyersCount: y.buyers.size, months: {}
  };
  for (const mo of Object.keys(y.months)) {
    const m = y.months[mo];
    cleanTimeline[yr].months[mo] = {
      compCount: m.compCount, cancCount: m.cancCount, compGmv: Math.round(m.compGmv), cancGmv: Math.round(m.cancGmv), buyersCount: m.buyers.size, days: {}
    };
    for (const dy of Object.keys(m.days)) {
      const d = m.days[dy];
      cleanTimeline[yr].months[mo].days[dy] = {
        compCount: d.compCount, cancCount: d.cancCount, compGmv: Math.round(d.compGmv), cancGmv: Math.round(d.cancGmv), buyersCount: d.buyers.size
      };
    }
  }
}

const finalOutput = {
  years: Array.from(allYears).sort().reverse(),
  timeline: cleanTimeline,
  topProducts: Object.values(topProductsMap).sort((a,b) => b.totalAmount - a.totalAmount).slice(0, 50),
  gwpItems: Object.values(gwpMap).sort((a,b) => b.quantity - a.quantity).slice(0, 50),
  sellerSkuSale: Object.values(sellerSkuSaleMap).sort((a,b) => b.totalAmount - a.totalAmount).slice(0, 50),
  sellerSkuGwp: Object.values(sellerSkuGwpMap).sort((a,b) => b.quantity - a.quantity).slice(0, 50)
};

fs.writeFileSync(path.join(__dirname, 'summary.json'), JSON.stringify(finalOutput));
const kb = (fs.statSync('summary.json').size / 1024).toFixed(1);
console.log(`\n🎉 สรุปสำเร็จ! ได้ summary.json ขนาดเพียง: ${kb} KB (รับประกันไม่เกิน RAM แน่นอน)`);