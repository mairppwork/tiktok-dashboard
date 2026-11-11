const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const csv = require('csv-parser');

const app = express();
app.use(cors());
app.use(express.static(__dirname));

const ORDER_DIR = path.join(__dirname, 'Order');

let cachedOrders = [];
let isProcessing = false;

function parseDateParts(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    const dateObj = xlsx.SSF.parse_date_code(val);
    if (dateObj) {
      return {
        year: String(dateObj.y),
        month: String(dateObj.m).padStart(2, '0'),
        day: String(dateObj.d).padStart(2, '0')
      };
    }
  }

  const str = String(val).trim();
  const dmyMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmyMatch) {
    return {
      year: dmyMatch[3],
      month: dmyMatch[2].padStart(2, '0'),
      day: dmyMatch[1].padStart(2, '0')
    };
  }

  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    return {
      year: String(d.getFullYear()),
      month: String(d.getMonth() + 1).padStart(2, '0'),
      day: String(d.getDate()).padStart(2, '0')
    };
  }
  return null;
}

function extractRowData(row) {
  let orderId = null;
  let buyer = null;
  let productName = null;
  let sku = null;
  let sellerSku = null;
  let quantity = 1;
  let dateRaw = null;
  let orderAmountRaw = null;
  let actualPaidItemPriceRaw = null;
  let statusRaw = null;

  for (const k of Object.keys(row)) {
    const keyTrim = k.trim();
    const keyLower = keyTrim.toLowerCase().replace(/[\s_\-]/g, '');
    const val = row[k];

    // 1. Product Name
    if (!productName) {
      if (
        keyTrim.toLowerCase() === 'product name' ||
        keyLower === 'productname' ||
        keyTrim === 'ชื่อสินค้า' ||
        keyTrim.includes('ชื่อสินค้า') ||
        keyLower.includes('productname')
      ) {
        if (val && String(val).trim() !== '') productName = String(val).trim();
      }
    }

    // 2. Seller SKU
    if (!sellerSku) {
      if (
        keyTrim.toLowerCase() === 'seller sku' ||
        keyLower === 'sellersku' ||
        keyTrim.includes('Seller SKU') ||
        keyTrim.includes('SKU ผู้ขาย') ||
        keyLower.includes('sellersku')
      ) {
        if (val && String(val).trim() !== '') sellerSku = String(val).trim();
      }
    }

    // 3. SKU ID
    if (!sku) {
      if (
        keyTrim.toLowerCase() === 'sku id' ||
        keyLower === 'skuid' ||
        keyTrim.includes('SKU ID') ||
        keyTrim.includes('รหัส SKU') ||
        keyLower.includes('skuid') ||
        keyLower === 'sku'
      ) {
        if (val && String(val).trim() !== '') sku = String(val).trim();
      }
    }

    // 4. Order ID
    if (!orderId && (
      keyLower.includes('orderid') ||
      keyLower.includes('หมายเลขคำสั่งซื้อ') ||
      keyLower.includes('เลขที่คำสั่งซื้อ') ||
      keyLower === 'id'
    )) {
      orderId = String(val).trim();
    }

    // 5. Buyer
    if (!buyer && (
      keyLower.includes('buyer') ||
      keyLower.includes('ผู้ซื้อ') ||
      keyLower.includes('customer') ||
      keyLower.includes('recipient') ||
      keyLower.includes('username')
    )) {
      buyer = String(val).trim();
    }

    // 6. Quantity
    if (keyLower.includes('quantity') || keyLower.includes('จำนวน') || keyLower === 'qty') {
      const q = parseInt(val);
      if (!isNaN(q) && q > 0) quantity = q;
    }

    // 7. Date
    if (!dateRaw && (
      keyLower.includes('createdtime') ||
      keyLower.includes('ordertime') ||
      keyLower.includes('paidtime') ||
      keyLower.includes('วันที่สร้าง') ||
      keyLower.includes('เวลาสร้าง') ||
      keyLower.includes('date')
    )) {
      dateRaw = val;
    }

    // 8. ยอดเฉพาะชิ้นนี้หลังหักส่วนลด
    if (actualPaidItemPriceRaw === null && (
      keyLower.includes('skusubtotalafterdiscount') ||
      keyLower.includes('skuunitdiscountedprice') ||
      keyLower.includes('itemamount') ||
      keyLower.includes('productsubtotal') ||
      keyLower.includes('ยอดรวมย่อยของสินค้า') ||
      keyLower.includes('ราคาหลังหักส่วนลด')
    )) {
      actualPaidItemPriceRaw = val;
    }

    // 9. Order Amount
    if (orderAmountRaw === null && (
      keyLower.includes('orderamount') ||
      keyLower.includes('orderamoud') ||
      keyLower.includes('ยอดรวมของคำสั่งซื้อ') ||
      keyLower.includes('ยอดคำสั่งซื้อ') ||
      keyLower.includes('totalamount') ||
      keyLower === 'gmv'
    )) {
      orderAmountRaw = val;
    }

    // 10. Status
    if (!statusRaw && (
      keyLower.includes('orderstatus') ||
      keyLower.includes('ordersubstatus') ||
      keyLower.includes('สถานะคำสั่งซื้อ') ||
      keyLower.includes('สถานะ') ||
      keyLower.includes('orderstate') ||
      keyLower === 'status'
    )) {
      statusRaw = String(val).trim();
    }
  }

  const dateParts = parseDateParts(dateRaw);
  if (!dateParts) return null;

  function parseNumber(raw) {
    if (typeof raw === 'number') return raw;
    if (typeof raw === 'string') return parseFloat(raw.replace(/[^0-9.-]+/g, '')) || 0;
    return 0;
  }

  const cleanOrderAmount = parseNumber(orderAmountRaw);
  const cleanItemPaidPrice = actualPaidItemPriceRaw !== null ? parseNumber(actualPaidItemPriceRaw) : cleanOrderAmount;

  const sLower = (statusRaw || '').toLowerCase();
  const isCanceled = (
    sLower.includes('cancel') ||
    sLower.includes('ยกเลิก') ||
    sLower.includes('refund') ||
    sLower.includes('คืนเงิน') ||
    sLower.includes('closed') ||
    sLower.includes('unpaid')
  );

  const pNameLower = (productName || '').toLowerCase();
  const rawSkuLower = String(sku || '').toLowerCase();
  const rawSellerSkuLower = String(sellerSku || '').toLowerCase();

  // ตรวจจับ 1D, 1Day, WTP, INF1
  const isExplicitGwp = (
    rawSellerSkuLower.includes('wtp') ||
    rawSkuLower.includes('wtp') ||
    pNameLower.includes('wtp') ||
    rawSellerSkuLower.includes('inf1') ||
    rawSkuLower.includes('inf1') ||
    pNameLower.includes('inf1') ||
    rawSellerSkuLower.includes('1day') ||
    rawSellerSkuLower.includes('_1d') ||
    rawSellerSkuLower.endsWith('1d') ||
    rawSkuLower.includes('1day') ||
    rawSkuLower.includes('_1d') ||
    rawSkuLower.endsWith('1d') ||
    pNameLower.includes('1day') ||
    pNameLower.includes('1-day') ||
    /\b1d\b/.test(pNameLower)
  );

  const hasGwpInSku = (
    rawSkuLower.includes('gwp') ||
    rawSellerSkuLower.includes('gwp')
  );

  const isGWP = (
    cleanItemPaidPrice === 0 ||
    cleanOrderAmount === 0 ||
    isExplicitGwp ||
    hasGwpInSku ||
    pNameLower.includes('ซองกันน้ำ') ||
    pNameLower.includes('waterproof') ||
    pNameLower.includes('pouch') ||
    pNameLower.includes('gwp') ||
    pNameLower.includes('ของแถม') ||
    pNameLower.includes('ห้ามจำหน่าย') ||
    pNameLower.includes('not for sale') ||
    pNameLower.includes('free gift') ||
    pNameLower.includes('tester') ||
    pNameLower.includes('sample')
  );

  return {
    orderId: orderId,
    buyer: buyer || orderId || 'Customer',
    productName: productName || 'ไม่ระบุชื่อสินค้า',
    sku: sku || '-',
    sellerSku: sellerSku || sku || '-',
    quantity: quantity,
    year: dateParts.year,
    month: dateParts.month,
    day: dateParts.day,
    amount: isGWP ? 0 : cleanOrderAmount,
    itemPaidPrice: cleanItemPaidPrice,
    status: isCanceled ? 'Canceled' : 'Completed',
    isGWP: isGWP
  };
}

async function preloadOrderData() {
  if (isProcessing) return;
  isProcessing = true;
  console.time('⚡ ประมวลผลข้อมูลเสร็จสิ้นใน');
  console.log('🔄 กำลังอ่านไฟล์คำสั่งซื้อจากโฟลเดอร์ Order...');

  try {
    if (!fs.existsSync(ORDER_DIR)) {
      console.log(`❌ ไม่พบโฟลเดอร์: ${ORDER_DIR} -> กำลังสร้างโฟลเดอร์อัตโนมัติ...`);
      fs.mkdirSync(ORDER_DIR, { recursive: true });
      isProcessing = false;
      return;
    }

    const files = fs.readdirSync(ORDER_DIR).filter(f => f.endsWith('.xlsx') || f.endsWith('.csv') || f.endsWith('.xls'));
    const allRows = [];

    for (const file of files) {
      const filePath = path.join(ORDER_DIR, file);
      console.log(`📂 กำลังอ่านไฟล์: ${file}`);

      if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
        const workbook = xlsx.readFile(filePath, { 
          dense: true,
          cellStyles: false,
          cellHTML: false,
          cellFormula: false
        });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });

        for (let i = 0; i < rows.length; i++) {
          const item = extractRowData(rows[i]);
          if (item) allRows.push(item);
        }

        delete workbook.Sheets[sheetName];
        delete workbook.Sheets;
      } else if (file.endsWith('.csv')) {
        await new Promise((resolve) => {
          fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (row) => {
              const item = extractRowData(row);
              if (item) allRows.push(item);
            })
            .on('end', resolve);
        });
      }
    }

    cachedOrders = allRows;

    const uniqueOrders = new Set(cachedOrders.map(o => o.orderId).filter(Boolean)).size;
    const uniqueProducts = new Set(cachedOrders.map(o => o.productName)).size;
    const uniqueSKUs = new Set(cachedOrders.map(o => o.sellerSku || o.sku).filter(s => s !== '-')).size;

    console.log(`✅ โหลดข้อมูลสำเร็จ:`);
    console.log(`   - จำนวนแถวสินค้าทั้งหมด: ${cachedOrders.length.toLocaleString()} แถว`);
    console.log(`   - จำนวนคำสั่งซื้อ (Orders): ${uniqueOrders.toLocaleString()} บิล`);
    console.log(`   - จำนวนชื่อสินค้า (Products): ${uniqueProducts.toLocaleString()} รายการ`);
    console.log(`   - จำนวน Seller SKU ที่พบ: ${uniqueSKUs.toLocaleString()} รหัส`);
  } catch (err) {
    console.error('❌ ข้อผิดพลาด:', err);
  } finally {
    console.timeEnd('⚡ ประมวลผลข้อมูลเสร็จสิ้นใน');
    isProcessing = false;
  }
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/api/orders', (req, res) => {
  if (isProcessing) {
    return res.json({ success: false, message: 'กำลังประมวลผลข้อมูล โปรดรอสักครู่...' });
  }
  res.json({
    success: true,
    totalOrders: cachedOrders.length,
    orders: cachedOrders
  });
});

app.get('/api/reload', async (req, res) => {
  await preloadOrderData();
  res.json({ success: true, count: cachedOrders.length });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`🚀 Server พร้อมทำงานแล้วที่ Port: ${PORT}`);
  await preloadOrderData();
});