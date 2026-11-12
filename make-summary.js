const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const orderDir = path.join(__dirname, 'Order');
console.log('🔄 เริ่มต้นประมวลผลข้อมูลฉบับสมบูรณ์ (สัดส่วน + Forecast + เทียบวัน)...');

if (!fs.existsSync(orderDir)) {
  console.error('❌ ไม่พบโฟลเดอร์ Order');
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
const sellerSkuMap = {};
const statusBreakdown = {}; // เก็บสัดส่วนทุกสถานะในไฟล์
const allYears = new Set();
const activeMonthsSet = new Set(); // เก็บจำนวนเดือนที่มีข้อมูลจริงเพื่อหารค่าเฉลี่ย Forecast

for (const file of files) {
  console.log(`📂 กำลังอ่านไฟล์: ${file}`);
  const wb = xlsx.readFile(path.join(orderDir, file), { dense: true, cellStyles: false, cellHTML: false, cellFormula: false });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });

  for (const row of rows) {
    let orderId = '', buyer = '', productName = '', sku = '-', sellerSku = '-', quantity = 1;
    let dateRaw = '', amount = 0, itemPaid = null, statusRaw = 'Completed', province = '';

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
      if (kl.includes('orderstatus') || kl.includes('สถานะ')) statusRaw = String(v).trim() || 'Completed';
      if (!province && (kl.includes('province') || kl.includes('จังหวัด'))) province = String(v).trim();
    }

    const dParts = parseDateParts(dateRaw) || { year: '2026', month: '01', day: '01' };
    const yr = dParts.year, mo = dParts.month, dy = dParts.day;
    const ymKey = `${yr}-${mo}`;
    allYears.add(yr);
    activeMonthsSet.add(ymKey);

    const cleanAmount = amount;
    const paid = itemPaid !== null ? itemPaid : cleanAmount;
    const isCanceled = statusRaw.toLowerCase().includes('cancel') || statusRaw.includes('ยกเลิก');
    const isCompleted = !isCanceled;

    // เก็บสัดส่วนทุกสถานะดิบจากไฟล์
    statusBreakdown[statusRaw] = (statusBreakdown[statusRaw] || 0) + 1;

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

    const isAuction = !isGWP && (
      pLower.includes('นาทีทอง') || 
      pLower.includes('ประมูล') || 
      pLower.includes('action')
    );

    // โครงสร้าง Timeline ละเอียดถึงระดับวัน (รองรับเปรียบเทียบวันข้ามเดือน/ข้ามปี)
    if (!timeline[yr]) timeline[yr] = { compOrders: new Set(), cancOrders: new Set(), compGmv: 0, cancGmv: 0, buyers: new Set(), months: {} };
    if (!timeline[yr].months[mo]) timeline[yr].months[mo] = { compOrders: new Set(), cancOrders: new Set(), compGmv: 0, cancGmv: 0, buyers: new Set(), days: {} };
    if (!timeline[yr].months[mo].days[dy]) timeline[yr].months[mo].days[dy] = { compOrders: new Set(), cancOrders: new Set(), compGmv: 0, cancGmv: 0, buyers: new Set(), gwpQty: 0, auctionGmv: 0 };

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

    if (isGWP) timeline[yr].months[mo].days[dy].gwpQty += quantity;
    if (isAuction && isCompleted) timeline[yr].months[mo].days[dy].auctionGmv += cleanAmount;

    // เก็บข้อมูล Seller SKU รายเดือน เพื่อใช้คำนวณ Forecast แม่นยำ
    const skuKey = `${cleanSku}___${isGWP ? 'GWP' : 'SALE'}`;
    if (!sellerSkuMap[skuKey]) {
      sellerSkuMap[skuKey] = {
        sellerSku: cleanSku,
        sampleName: productName,
        type: isGWP ? 'GWP' : (isAuction ? 'Auction' : 'Regular'),
        totalQty: 0,
        completedQty: 0,
        totalGmv: 0,
        monthlyQty: {} // เก็บยอดขายแต่ละเดือน เช่น { '2026-01': 120, '2026-02': 150 }
      };
    }
    sellerSkuMap[skuKey].totalQty += quantity;
    if (isCompleted) {
      sellerSkuMap[skuKey].completedQty += quantity;
      sellerSkuMap[skuKey].monthlyQty[ymKey] = (sellerSkuMap[skuKey].monthlyQty[ymKey] || 0) + quantity;
      if (!isGWP) sellerSkuMap[skuKey].totalGmv += cleanAmount;
    }

    // เก็บตารางสินค้า
    const targetMap = isGWP ? gwpMap : (isAuction ? auctionMap : topProductsMap);
    const itemKey = `${productName}___${sku}`;
    if (!targetMap[itemKey]) {
      targetMap[itemKey] = { name: productName || 'สินค้าไม่ระบุชื่อ', sku, sellerSku: cleanSku, quantity: 0, completedCount: 0, canceledCount: 0, totalAmount: 0, _seen: new Set() };
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

// คำนวณ Forecast ราย SKU สำหรับเดือนถัดไป
const sortedMonths = Array.from(activeMonthsSet).sort();
const numMonths = Math.max(sortedMonths.length, 1);
const latestMonth = sortedMonths[sortedMonths.length - 1];
const prevMonth = sortedMonths.length > 1 ? sortedMonths[sortedMonths.length - 2] : latestMonth;

const forecastList = Object.values(sellerSkuMap).map(item => {
  const avgMonthly = item.completedQty / numMonths;
  const lastM = item.monthlyQty[latestMonth] || 0;
  const prevM = item.monthlyQty[prevMonth] || lastM;

  // คำนวณอัตราการเติบโต (Growth Rate) แบบจำกัดเพดานไม่ให้เพี้ยน (-30% ถึง +50%)
  let growthRate = 0;
  if (prevM > 0 && sortedMonths.length > 1) {
    growthRate = (lastM - prevM) / prevM;
    growthRate = Math.max(-0.3, Math.min(0.5, growthRate));
  }

  // สูตรพยากรณ์: (ค่าเฉลี่ยรายเดือน 60% + เดือนล่าสุด 40%) * (1 + แนวโน้มเติบโต)
  const baseForecast = sortedMonths.length > 1 ? (avgMonthly * 0.6) + (lastM * 0.4) : lastM;
  const predictedSales = Math.ceil(baseForecast * (1 + growthRate));
  const recommendedOrder = Math.ceil(predictedSales * 1.15); // บวก Safety Stock 15%

  return {
    sellerSku: item.sellerSku,
    sampleName: item.sampleName,
    type: item.type,
    totalCompletedQty: item.completedQty,
    avgMonthlyQty: Math.round(avgMonthly),
    lastMonthQty: lastM,
    growthPct: Math.round(growthRate * 100),
    predictedNextMonth: predictedSales,
    recommendedRestock: recommendedOrder
  };
}).sort((a, b) => b.recommendedRestock - a.recommendedRestock);

// แปลง Timeline
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
        compCount: d.compOrders.size, cancCount: d.cancOrders.size, compGmv: Math.round(d.compGmv), cancGmv: Math.round(d.cancGmv), buyersCount: d.buyers.size, gwpQty: d.gwpQty, auctionGmv: Math.round(d.auctionGmv)
      };
    }
  }
}

const sanitizeList = (map, sortBy) => Object.values(map).map(item => { delete item._seen; return item; }).sort((a, b) => b[sortBy] - a[sortBy]).slice(0, 60);

const finalOutput = {
  totalCount: grandTotalOrders,
  activeMonths: sortedMonths,
  statusBreakdown,
  years: Array.from(allYears).sort().reverse(),
  timeline: cleanTimeline,
  topProducts: sanitizeList(topProductsMap, 'totalAmount'),
  auctionProducts: sanitizeList(auctionMap, 'totalAmount'),
  gwpItems: sanitizeList(gwpMap, 'quantity'),
  forecast: forecastList
};

const outPath = path.join(__dirname, 'summary.json');
fs.writeFileSync(outPath, JSON.stringify(finalOutput));
const kb = (fs.statSync(outPath).size / 1024).toFixed(1);
console.log(`\n🎉 สรุปยอดเสร็จสิ้น! บันทึกเป็น summary.json ขนาด: ${kb} KB (คำนวณ Forecast ${forecastList.length} SKUs)`);