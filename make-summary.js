const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const orderDir = path.join(__dirname, 'Order');
console.log('🔄 เริ่มต้นประมวลผลข้อมูลแบบแยกตามวัน/เดือน/ปี และแยกรายสินค้าครบถ้วน...');

if (!fs.existsSync(orderDir)) {
  console.error('❌ ไม่พบโฟลเดอร์ Order');
  process.exit(1);
}

const files = fs.readdirSync(orderDir).filter(f => f.endsWith('.xlsx') || f.endsWith('.csv') || f.endsWith('.xls'));

function cleanSellerSku(rawSku) {
  if (!rawSku || rawSku === '-' || String(rawSku).trim() === '') return '';
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

// ตัวช่วยสร้างกล่องเก็บข้อมูลของแต่ละช่วงเวลา (ปี / เดือน / วัน)
function createPeriodNode() {
  return {
    compOrders: new Set(),
    cancOrders: new Set(),
    compGmv: 0,
    cancGmv: 0,
    buyers: new Set(),
    regGmv: 0,
    aucGmv: 0,
    gwpQty: 0,
    statusMap: {},
    products: {}, // เก็บ topProducts ของช่วงเวลานั้น
    auctions: {}, // เก็บ auctionProducts ของช่วงเวลานั้น
    gwps: {}      // เก็บ gwpItems ของช่วงเวลานั้น
  };
}

const timeline = {};
const forecastSkuMap = {};
const allYears = new Set();
const activeMonthsSet = new Set();

for (const file of files) {
  console.log(`📂 กำลังอ่านไฟล์: ${file}`);
  const wb = xlsx.readFile(path.join(orderDir, file), { dense: true, cellStyles: false, cellHTML: false, cellFormula: false });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });

  for (const row of rows) {
    let orderId = '', buyer = '', productName = '', sku = '-', sellerSku = '-', quantity = 1;
    let dateRaw = '', amount = 0, itemPaid = null, statusRaw = 'Completed';

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

    // สร้างโหนด ปี -> เดือน -> วัน
    if (!timeline[yr]) timeline[yr] = { ...createPeriodNode(), months: {} };
    if (!timeline[yr].months[mo]) timeline[yr].months[mo] = { ...createPeriodNode(), days: {} };
    if (!timeline[yr].months[mo].days[dy]) timeline[yr].months[mo].days[dy] = createPeriodNode();

    const oId = orderId || `${productName}_${Math.random()}`;
    const bId = buyer || oId;
    const itemKey = `${productName || 'สินค้าไม่ระบุชื่อ'}___${sku}`;

    // บันทึกข้อมูลลงทั้ง 3 ระดับ (ปี, เดือน, วัน)
    [timeline[yr], timeline[yr].months[mo], timeline[yr].months[mo].days[dy]].forEach(node => {
      node.buyers.add(bId);
      node.statusMap[statusRaw] = (node.statusMap[statusRaw] || 0) + 1;

      if (isCompleted) {
        if (!node.compOrders.has(oId)) {
          node.compOrders.add(oId);
          if (!isGWP) {
            node.compGmv += cleanAmount;
            if (isAuction) node.aucGmv += cleanAmount;
            else node.regGmv += cleanAmount;
          }
        }
      } else {
        if (!node.cancOrders.has(oId)) {
          node.cancOrders.add(oId);
          if (!isGWP) node.cancGmv += cleanAmount;
        }
      }

      if (isGWP) node.gwpQty += quantity;

      // เก็บรายการสินค้าแยกตามช่วงเวลา
      const targetBucket = isGWP ? node.gwps : (isAuction ? node.auctions : node.products);
      if (!targetBucket[itemKey]) {
        targetBucket[itemKey] = { name: productName || 'สินค้าไม่ระบุชื่อ', sku, quantity: 0, completedCount: 0, canceledCount: 0, totalAmount: 0, _seen: new Set() };
      }
      targetBucket[itemKey].quantity += quantity;
      if (!targetBucket[itemKey]._seen.has(oId)) {
        targetBucket[itemKey]._seen.add(oId);
        if (isCompleted) {
          targetBucket[itemKey].completedCount++;
          if (!isGWP) targetBucket[itemKey].totalAmount += cleanAmount;
        } else {
          targetBucket[itemKey].canceledCount++;
        }
      }
    });

    // แก้ไขปัญหา Forecast รวมเหลือ 2 ตัว: ใช้รหัส SKU + ชื่อสินค้าเป็น Key เพื่อแยกสินค้าทุกตัวชัดเจน
    const displaySku = cleanSku || (sku !== '-' ? sku : 'NO-SKU');
    const forecastKey = `${displaySku}___${productName || 'ไม่ระบุชื่อ'}___${isGWP ? 'GWP' : 'SALE'}`;

    if (!forecastSkuMap[forecastKey]) {
      forecastSkuMap[forecastKey] = {
        sellerSku: displaySku,
        sampleName: productName || 'สินค้าไม่ระบุชื่อ',
        type: isGWP ? 'GWP' : (isAuction ? 'Auction' : 'Regular'),
        totalQty: 0,
        completedQty: 0,
        monthlyQty: {}
      };
    }
    forecastSkuMap[forecastKey].totalQty += quantity;
    if (isCompleted) {
      forecastSkuMap[forecastKey].completedQty += quantity;
      forecastSkuMap[forecastKey].monthlyQty[ymKey] = (forecastSkuMap[forecastKey].monthlyQty[ymKey] || 0) + quantity;
    }
  }
}

// แปลงข้อมูลแต่ละ Period ให้เป็น JSON ขนาดกะทัดรัด (จำกัด Top 25 ต่อวัน เพื่อไม่ให้ไฟล์ใหญ่เกินไป)
function serializeNode(node, limit = 30) {
  const cleanBucket = (bucket, sortKey) => Object.values(bucket)
    .map(item => { const copy = { ...item }; delete copy._seen; return copy; })
    .sort((a, b) => b[sortKey] - a[sortKey])
    .slice(0, limit);

  return {
    compCount: node.compOrders.size,
    cancCount: node.cancOrders.size,
    compGmv: Math.round(node.compGmv),
    cancGmv: Math.round(node.cancGmv),
    buyersCount: node.buyers.size,
    regGmv: Math.round(node.regGmv),
    aucGmv: Math.round(node.aucGmv),
    gwpQty: node.gwpQty,
    statusBreakdown: node.statusMap,
    topProducts: cleanBucket(node.products, 'totalAmount'),
    auctionProducts: cleanBucket(node.auctions, 'totalAmount'),
    gwpItems: cleanBucket(node.gwps, 'quantity')
  };
}

const cleanTimeline = {};
let grandTotalOrders = 0;

for (const yr of Object.keys(timeline)) {
  const y = timeline[yr];
  grandTotalOrders += y.compOrders.size + y.cancOrders.size;
  cleanTimeline[yr] = { ...serializeNode(y, 50), months: {} };

  for (const mo of Object.keys(y.months)) {
    const m = y.months[mo];
    cleanTimeline[yr].months[mo] = { ...serializeNode(m, 40), days: {} };

    for (const dy of Object.keys(m.days)) {
      const d = m.days[dy];
      cleanTimeline[yr].months[mo].days[dy] = serializeNode(d, 20);
    }
  }
}

// คำนวณ Forecast แยกรายสินค้าทุกตัว
const sortedMonths = Array.from(activeMonthsSet).sort();
const numMonths = Math.max(sortedMonths.length, 1);
const latestMonth = sortedMonths[sortedMonths.length - 1];
const prevMonth = sortedMonths.length > 1 ? sortedMonths[sortedMonths.length - 2] : latestMonth;

const forecastList = Object.values(forecastSkuMap).map(item => {
  const avgMonthly = item.completedQty / numMonths;
  const lastM = item.monthlyQty[latestMonth] || 0;
  const prevM = item.monthlyQty[prevMonth] || lastM;

  let growthRate = 0;
  if (prevM > 0 && sortedMonths.length > 1) {
    growthRate = Math.max(-0.3, Math.min(0.5, (lastM - prevM) / prevM));
  }

  const baseForecast = sortedMonths.length > 1 ? (avgMonthly * 0.6) + (lastM * 0.4) : lastM;
  const predictedSales = Math.max(1, Math.ceil(baseForecast * (1 + growthRate)));
  const recommendedOrder = Math.ceil(predictedSales * 1.15);

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

const finalOutput = {
  totalCount: grandTotalOrders,
  activeMonths: sortedMonths,
  years: Array.from(allYears).sort().reverse(),
  timeline: cleanTimeline,
  forecast: forecastList
};

const outPath = path.join(__dirname, 'summary.json');
fs.writeFileSync(outPath, JSON.stringify(finalOutput));
const kb = (fs.statSync(outPath).size / 1024).toFixed(1);
console.log(`\n🎉 สรุปยอดสำเร็จ! ขนาดไฟล์: ${kb} KB (พบสินค้าในหน้า Forecast ทั้งหมด ${forecastList.length} รายการ)`);