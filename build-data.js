let summaryData = null;

    async function fetchOrders() {
      const badge = document.getElementById('syncStatusBadge');
      badge.innerText = '● Syncing...';
      badge.className = 'text-[11px] bg-amber-50 text-amber-700 font-bold px-2.5 py-1 rounded-full border border-amber-200';

      try {
        const res = await fetch('/api/orders');
        const result = await res.json();

        if (result.success && result.data) {
          summaryData = result.data;
          badge.innerText = `● Synced Ready`;
          badge.className = 'text-[11px] bg-[#faeeef] text-[#ea5a7e] font-bold px-2.5 py-1 rounded-full border border-[#f5d5db]';

          const yearSelect = document.getElementById('yearFilter');
          yearSelect.innerHTML = '';
          (summaryData.years || ['2026']).forEach(y => {
            const opt = document.createElement('option');
            opt.value = y;
            opt.innerText = y;
            yearSelect.appendChild(opt);
          });

          currentTopProductsList = summaryData.topProducts || [];
          currentAuctionList = summaryData.auctionProducts || [];
          currentGwpList = summaryData.gwpItems || [];
          currentSellerSkuSaleList = summaryData.sellerSkuSale || [];
          currentSellerSkuGwpList = summaryData.sellerSkuGwp || [];

          populateDayDropdown();
          renderDashboard();
        }
      } catch (err) {
        console.error(err);
        badge.innerText = '● Server Offline';
      }
    }

    function renderDashboard() {
      if (!summaryData || !summaryData.timeline) return;

      const selectedYear = document.getElementById('yearFilter').value;
      const selectedMonth = document.getElementById('monthFilter').value;
      const selectedDay = document.getElementById('dayFilter').value;

      const yNode = summaryData.timeline[selectedYear] || { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0, buyersCount: 0, months: {} };

      let labels = [];
      let compCountList = [];
      let cancCountList = [];
      let compGmvList = [];
      let cancGmvList = [];
      let totalComp = 0, totalCanc = 0, totalCompGmv = 0, totalCancGmv = 0, totalBuyers = 0;

      if (currentMode === 'year') {
        document.getElementById('periodHeading').innerText = `สรุปภาพรวมคำสั่งซื้อ: ประจำปี ${selectedYear} (ทั้งปี)`;
        document.getElementById('chartSubText').innerText = `สถิติคำสั่งซื้อแยกตาม 12 เดือน ประจำปี ${selectedYear}`;

        labels = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

        for (let i = 1; i <= 12; i++) {
          const mStr = String(i).padStart(2, '0');
          const mNode = yNode.months[mStr] || { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0 };
          compCountList.push(mNode.compCount);
          cancCountList.push(mNode.cancCount);
          compGmvList.push(mNode.compGmv);
          cancGmvList.push(mNode.cancGmv);
        }

        totalComp = yNode.compCount;
        totalCanc = yNode.cancCount;
        totalCompGmv = yNode.compGmv;
        totalCancGmv = yNode.cancGmv;
        totalBuyers = yNode.buyersCount;
      } else {
        const mNode = yNode.months[selectedMonth] || { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0, buyersCount: 0, days: {} };
        const daysInMonth = new Date(selectedYear, parseInt(selectedMonth), 0).getDate();

        if (selectedDay === 'all') {
          document.getElementById('periodHeading').innerText = `สรุปภาพรวมคำสั่งซื้อ: เดือน ${selectedMonth}/${selectedYear} (ทั้งเดือน)`;
          document.getElementById('chartSubText').innerText = `สถิติคำสั่งซื้อรายวัน วันที่ 1 - ${daysInMonth}`;

          for (let d = 1; d <= daysInMonth; d++) {
            const dStr = String(d).padStart(2, '0');
            labels.push(`วันที่ ${d}`);
            const dNode = mNode.days[dStr] || { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0 };
            compCountList.push(dNode.compCount);
            cancCountList.push(dNode.cancCount);
            compGmvList.push(dNode.compGmv);
            cancGmvList.push(dNode.cancGmv);
          }

          totalComp = mNode.compCount;
          totalCanc = mNode.cancCount;
          totalCompGmv = mNode.compGmv;
          totalCancGmv = mNode.cancGmv;
          totalBuyers = mNode.buyersCount;
        } else {
          document.getElementById('periodHeading').innerText = `สรุปภาพรวมคำสั่งซื้อ: วันที่ ${parseInt(selectedDay)} เดือน ${selectedMonth}/${selectedYear}`;
          labels.push(`วันที่ ${parseInt(selectedDay)}`);
          const dNode = mNode.days[selectedDay] || { compCount: 0, cancCount: 0, compGmv: 0, cancGmv: 0, buyersCount: 0 };
          compCountList.push(dNode.compCount);
          cancCountList.push(dNode.cancCount);
          compGmvList.push(dNode.compGmv);
          cancGmvList.push(dNode.cancGmv);

          totalComp = dNode.compCount;
          totalCanc = dNode.cancCount;
          totalCompGmv = dNode.compGmv;
          totalCancGmv = dNode.cancGmv;
          totalBuyers = dNode.buyersCount;
        }
      }

      const totalOrders = totalComp + totalCanc;
      const totalGmv = totalCompGmv + totalCancGmv;
      currentTotalGmv = totalGmv;

      const compRate = totalOrders > 0 ? ((totalComp / totalOrders) * 100).toFixed(1) : '0.0';
      const cancRate = totalOrders > 0 ? ((totalCanc / totalOrders) * 100).toFixed(1) : '0.0';
      const aov = totalComp > 0 ? (totalCompGmv / totalComp) : 0;

      document.getElementById('cardTotalGmv').innerText = '฿' + totalGmv.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
      document.getElementById('cardTotalOrders').innerText = totalOrders.toLocaleString();
      document.getElementById('cardTotalBuyers').innerText = totalBuyers.toLocaleString();
      document.getElementById('cardAov').innerText = '฿' + aov.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});

      document.getElementById('txtCompCount').innerText = `${totalComp.toLocaleString()} ออเดอร์`;
      document.getElementById('badgeCompRate').innerText = `${compRate}%`;
      document.getElementById('txtCompGmv').innerText = '฿' + totalCompGmv.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});

      document.getElementById('txtCancCount').innerText = `${totalCanc.toLocaleString()} ออเดอร์`;
      document.getElementById('badgeCancRate').innerText = `${cancRate}%`;
      document.getElementById('txtCancGmv').innerText = '฿' + totalCancGmv.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});

      renderBarChart(labels, compCountList, cancCountList, compGmvList, cancGmvList);
      renderAuctionList();
      filterProductDisplay();
      filterSellerSkuDisplay();
    }