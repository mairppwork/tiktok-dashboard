const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const ORDER_DIR = path.join(__dirname, 'Order');
const OUTPUT_FILE = path.join(__dirname, 'orders.json');

console.log('🔄 กำลังอ่านและบีบอัดข้อมูลเป็น orders.json (โหมดลดขนาดไฟล์)...');

if (!fs.existsSync(ORDER_DIR)) {
  console.log('❌ ไม่พบโฟลเดอร์ Order');
  process.exit(1);
}

const files = fs.readdirSync(ORDER_DIR).filter(f => f.endsWith('.xlsx') || f.endsWith('.csv') || f.endsWith('.xls'));
const allRows = [];

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

for (const file of files) {
  console.log(`📂 อ่านไฟล์: ${file}`);
  const wb = xlsx.readFile(path.join(ORDER_DIR, file), { 
    dense: true, 
    cellStyles: false, 
    cellHTML: false, 
    cellFormula: false 
  });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });

  for (const row of rows) {
    let orderId = '';
    let buyer = '';
    let productName = '';
    let sku = '-';
    let sellerSku = '-';
    let quantity = 1;
    let dateRaw = '';
    let orderAmountRaw = null;
    let itemPriceRaw = null;
    let statusRaw = '';

    for (const k of Object.keys(row)) {
      const keyTrim = k.trim();
      const keyLower = keyTrim.toLowerCase().replace(/[\s_\-]/g, '');
      const val = row[k];

      if (!productName && (keyTrim.toLowerCase() === 'product name' || keyLower === 'productname' || keyTrim.includes('ชื่อสินค้า'))) {
        if (val) productName = String(val).trim();
      }
      if (!sellerSku && (keyTrim.toLowerCase() === 'seller sku' || keyLower === 'sellersku' || keyTrim.includes('SKU ผู้ขาย'))) {
        if (val) sellerSku = String(val).trim();
      }
      if (!sku && (keyTrim.toLowerCase() === 'sku id' || keyLower === 'skuid' || keyTrim.includes('รหัส SKU') || keyLower === 'sku')) {
        if (val) sku = String(val).trim();
      }
      if (!orderId && (keyLower.includes('orderid') || keyLower.includes('หมายเลขคำสั่งซื้อ') || keyLower === 'id')) {
        orderId = String(val).trim();
      }
      if (!buyer && (keyLower.includes('buyer') || keyLower.includes('ผู้ซื้อ') || keyLower.includes('username'))) {
        buyer = String(val).trim();
      }
      if (keyLower.includes('quantity') || keyLower.includes('จำนวน') || keyLower === 'qty') {
        const q = parseInt(val);
        if (!isNaN(q) && q > 0) quantity = q;
      }
      if (!dateRaw && (keyLower.includes('createdtime') || keyLower.includes('ordertime') || keyLower.includes('วันที่สร้าง') || keyLower.includes('date'))) {
        dateRaw = val;
      }
      if (itemPriceRaw === null && (keyLower.includes('skusubtotalafterdiscount') || keyLower.includes('itemamount') || keyLower.includes('ราคาหลังหักส่วนลด'))) {
        itemPriceRaw = val;
      }
      if (orderAmountRaw === null && (keyLower.includes('orderamount') || keyLower.includes('ยอดรวมของคำสั่งซื้อ') || keyLower === 'gmv')) {
        orderAmountRaw = val;
      }
      if (!statusRaw && (keyLower.includes('orderstatus') || keyLower.includes('สถานะคำสั่งซื้อ') || keyLower === 'status')) {
        statusRaw = String(val).trim();
      }
    }

    const dateParts = parseDateParts(dateRaw) || { year: '2026', month: '01', day: '01' };
    const amount = parseFloat(String(orderAmountRaw || 0).replace(/[^0-9.-]+/g, '')) || 0;
    const itemPaid = itemPriceRaw !== null ? parseFloat(String(itemPriceRaw).replace(/[^0-9.-]+/g, '')) || 0 : amount;
    const sLower = statusRaw.toLowerCase();
    const isCanceled = sLower.includes('cancel') || sLower.includes('ยกเลิก') || sLower.includes('refund');

    const pLower = (productName || '').toLowerCase();
    const sLowerSku = String(sellerSku || '').toLowerCase();
    const skuLower = String(sku || '').toLowerCase();
    const cleanSku = cleanSellerSku(sellerSku);

    const isExplicitGwp = (
      cleanSku.toUpperCase() === '1D' ||
      cleanSku.toUpperCase() === '1DAY' ||
      cleanSku.toUpperCase() === 'WTP' ||
      cleanSku.toUpperCase().includes('INF1') ||
      cleanSku.toUpperCase().includes('GWP') ||
      sLowerSku.includes('wtp') || sLowerSku.includes('inf1') || sLowerSku.includes('1day') ||
      skuLower.includes('wtp') || skuLower.includes('inf1') ||
      pLower.includes('1day') || /\b1d\b/.test(pLower)
    );

    const isGWP = (
      itemPaid === 0 || amount === 0 || isExplicitGwp ||
      skuLower.includes('gwp') || sLowerSku.includes('gwp') ||
      pLower.includes('ซองกันน้ำ') || pLower.includes('ของแถม') || pLower.includes('not for sale')
    );

    // เก็บเฉพาะข้อมูลที่จำเป็นและย่อความยาวตัวอักษร
    allRows.push({
      orderId: orderId,
      buyer: buyer ? buyer.substring(0, 15) : 'Customer',
      productName: productName.substring(0, 80),
      sku: sku,
      sellerSku: sellerSku,
      quantity,
      year: dateParts.year,
      month: dateParts.month,
      day: dateParts.day,
      amount: isGWP ? 0 : amount,
      status: isCanceled ? 'Canceled' : 'Completed',
      isGWP
    });
  }
}

// เขียนไฟล์แบบ Compact (ไม่เว้นวรรค)
fs.writeFileSync(OUTPUT_FILE, JSON.stringify(allRows));
const stats = fs.statSync(OUTPUT_FILE);
const mb = (stats.size / (1024 * 1024)).toFixed(2);
console.log(`\n🎉 แปลงไฟล์สำเร็จ! ขนาดไฟล์ orders.json ใหม่: ${mb} MB (เป้าหมาย < 50 MB)`);