const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.static(__dirname));

const SUMMARY_FILE = path.join(__dirname, 'summary.json');
let cachedData = null;

function loadData() {
  if (fs.existsSync(SUMMARY_FILE)) {
    try {
      console.log('🚀 พบไฟล์ summary.json กำลังโหลดข้อมูลเข้าสู่ระบบ...');
      const raw = fs.readFileSync(SUMMARY_FILE, 'utf8');
      cachedData = JSON.parse(raw);
      console.log('✅ โหลดข้อมูลจาก summary.json สำเร็จเรียบร้อย');
    } catch (e) {
      console.error('❌ อ่านไฟล์ summary.json ไม่สำเร็จ:', e.message);
    }
  } else {
    console.warn('⚠️ ไม่พบ summary.json ในไดเรกทอรี');
  }
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/api/orders', (req, res) => {
  if (!cachedData) {
    loadData();
  }
  if (!cachedData) {
    return res.json({ success: false, message: 'ไม่พบข้อมูล summary.json' });
  }
  res.json({
    success: true,
    data: cachedData
  });
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`🚀 Server พร้อมทำงานแล้วที่ Port: ${PORT}`);
  loadData();
});