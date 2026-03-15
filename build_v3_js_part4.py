content = open('C:/Users/jamic/庫存分析/v3_parts/part_js_p4.txt', 'w', encoding='utf-8')

js = """
// ══ 21. RENDER: REPORT ══
function renderReport() {
    generateTextReport();
    renderSnapshotsList();
}

function generateTextReport() {
    const el = document.getElementById('report-text');
    if (!el || !aggregatedData.products) return;

    const arr = aggregatedData.products;
    const totalInv = aggregatedData.totalVal || 0;
    const dsi = aggregatedData._dsi || 0;
    const health = aggregatedData._health || 0;
    const gmroi = aggregatedData._gmroi || 0;
    const turnover = aggregatedData._turnover || 0;
    const stockouts = aggregatedData._stockouts || 0;
    const aging90Pct = aggregatedData._aging90Pct || 0;
    const mKeys = aggregatedData.mKeys || [];
    const ms = aggregatedData.monthly || {};

    const today = new Date().toISOString().slice(0,10);
    const abcCounts = {A:0,B:0,C:0};
    arr.forEach(p => abcCounts[p.abc]++);
    const totalLoss = arr.reduce((s,p)=>s+p.estimatedLoss, 0);
    const catTotals = {};
    arr.forEach(p => { catTotals[p.category] = (catTotals[p.category]||0) + p.totalCurrentValue; });
    const topCat = Object.entries(catTotals).sort((a,b)=>b[1]-a[1])[0];
    const highRisk = arr.filter(p=>p.avgAging>90).length;
    const lastMoGrowth = mKeys.length > 1 ? (ms[mKeys[mKeys.length-1]]?.growth || 0) : 0;

    const report =
'═══════════════════════════════════════════\n' +
'       庫存分析智能儀表板 V3.0 週報\n' +
'       報告日期：' + today + '\n' +
'═══════════════════════════════════════════\n\n' +
'【一、庫存概覽】\n' +
'  庫存總金額：' + (totalInv/10000).toFixed(1) + ' 萬元\n' +
'  庫存天數(DSI)：' + dsi.toFixed(1) + ' 天 ' + (dsi > 90 ? '⚠ 超標' : '✓ 正常') + '\n' +
'  存貨周轉率：' + turnover.toFixed(2) + ' 次/年\n' +
'  庫存健康分：' + health.toFixed(0) + ' / 100 分\n' +
'  GMROI：' + gmroi.toFixed(2) + ' 倍\n\n' +
'【二、ABC分類】\n' +
'  A類品項：' + abcCounts.A + ' 件（高價值，重點管控）\n' +
'  B類品項：' + abcCounts.B + ' 件（中價值，定期審查）\n' +
'  C類品項：' + abcCounts.C + ' 件（低價值，批量管理）\n\n' +
'【三、庫齡風險】\n' +
'  > 90天庫存佔比：' + aging90Pct.toFixed(1) + '%\n' +
'  高庫齡品項數：' + highRisk + ' 件\n' +
'  預估總損失：' + (totalLoss/10000).toFixed(1) + ' 萬元\n\n' +
'【四、供應鏈狀態】\n' +
'  缺貨品項數：' + stockouts + ' 件 ' + (stockouts > 5 ? '⚠ 需關注' : '✓ 正常') + '\n' +
'  最大庫存品類：' + (topCat ? topCat[0] + ' (' + (topCat[1]/10000).toFixed(1) + '萬元)' : '—') + '\n' +
'  上月COGS環比：' + (lastMoGrowth >= 0 ? '+' : '') + (lastMoGrowth*100).toFixed(1) + '%\n\n' +
'【五、行動建議】\n' +
(dsi > 90 ? '  ⚠ DSI超標，建議加強促銷或削減採購\n' : '') +
(aging90Pct > 15 ? '  ⚠ 高庫齡比例偏高，建議清倉促銷\n' : '') +
(stockouts > 5 ? '  ⚠ 多個品項缺貨，建議緊急補貨\n' : '') +
(gmroi < 1 ? '  ⚠ GMROI偏低，建議優化品項結構\n' : '') +
(health >= 80 ? '  ✓ 整體庫存健康狀況良好\n' : '') +
'\n' +
'═══════════════════════════════════════════\n' +
'   Inventory Strategic OS V3.0 自動生成\n' +
'═══════════════════════════════════════════';

    el.textContent = report;
}

function renderSnapshotsList() {
    const container = document.getElementById('snapshots-list');
    if (!container) return;
    snapshots = JSON.parse(localStorage.getItem('v3Snapshots') || '[]');
    if (snapshots.length === 0) {
        container.innerHTML = '<div class="text-center py-3 text-slate-500">尚無歷史快照</div>';
        return;
    }
    container.innerHTML = snapshots.slice().reverse().map((s, ri) => {
        const i = snapshots.length - 1 - ri;
        return '<div class="flex items-center justify-between p-2 rounded-lg bg-white/5">' +
            '<div><div class="text-slate-300 text-xs font-medium">' + s.date + '</div>' +
            '<div class="text-slate-500 text-[0.65rem]">庫存:' + s.inv + '萬 / DSI:' + s.dsi + '天 / 健康:' + s.health + '</div></div>' +
            '<button onclick="deleteSnapshot(' + i + ')" class="text-red-400 text-xs hover:text-red-300 px-2">刪除</button>' +
            '</div>';
    }).join('');
}

function saveSnapshot() {
    if (!aggregatedData.products) return;
    const snap = {
        date: new Date().toLocaleString('zh-TW'),
        inv: (aggregatedData.totalVal/10000).toFixed(1),
        dsi: (aggregatedData._dsi||0).toFixed(1),
        health: (aggregatedData._health||0).toFixed(0),
        gmroi: (aggregatedData._gmroi||0).toFixed(2),
        stockouts: aggregatedData._stockouts || 0,
        products: aggregatedData.products.length
    };
    snapshots.push(snap);
    localStorage.setItem('v3Snapshots', JSON.stringify(snapshots));
    renderSnapshotsList();
    alert('快照已儲存：' + snap.date);
}

function deleteSnapshot(i) {
    snapshots.splice(i, 1);
    localStorage.setItem('v3Snapshots', JSON.stringify(snapshots));
    renderSnapshotsList();
}

// ══ 22. RENDER: SETTINGS (static) ══
function renderSettings() {
    // Mostly static HTML with accordion, nothing dynamic needed
}

// ══ 23. EXPORT: PDF ══
function exportPDF() {
    const body = document.body;
    body.classList.add('export-mode');
    const activeTab = document.querySelector('.tab-content.active');
    const title = document.querySelector('.glass-btn.active[data-tab]')?.textContent?.trim() || '庫存分析';

    html2canvas(activeTab || document.querySelector('main'), {
        scale: 1.5,
        useCORS: true,
        backgroundColor: document.body.classList.contains('light-mode') ? '#f1f5f9' : '#0f172a',
        onclone: (cloned) => {
            cloned.querySelectorAll('.glass-panel').forEach(el => {
                el.style.backdropFilter = 'none';
                el.style.webkitBackdropFilter = 'none';
                el.style.background = document.body.classList.contains('light-mode') ? '#ffffff' : '#1e293b';
            });
        }
    }).then(canvas => {
        body.classList.remove('export-mode');
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
        const imgW = 277, imgH = canvas.height / canvas.width * imgW;
        let y = 0;
        while (y < imgH) {
            if (y > 0) pdf.addPage();
            pdf.addImage(canvas.toDataURL('image/jpeg', 0.85), 'JPEG', 10, 10, imgW, Math.min(190, imgH - y));
            y += 190;
        }
        pdf.save('庫存分析報告_' + new Date().toISOString().slice(0,10) + '.pdf');
    }).catch(() => {
        body.classList.remove('export-mode');
        alert('PDF匯出失敗，請重試');
    });
}

// ══ 24. EXPORT: JSON ══
function exportJSON() {
    if (!aggregatedData.products) return;
    const data = {
        exportDate: new Date().toISOString(),
        summary: {
            totalInventoryValue: aggregatedData.totalVal,
            dsi: aggregatedData._dsi,
            health: aggregatedData._health,
            gmroi: aggregatedData._gmroi,
            stockouts: aggregatedData._stockouts
        },
        products: aggregatedData.products,
        monthly: aggregatedData.monthly,
        seasonalIndex: aggregatedData.seasonalIndex
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = '庫存分析_' + new Date().toISOString().slice(0,10) + '.json';
    a.click(); URL.revokeObjectURL(url);
}

// ══ 25. EXPORT: CSV ══
function exportCSV() {
    if (!aggregatedData.products) return;
    const headers = ['ProductID','名稱','品類','ABC','XYZ','ABCXYZ','現有庫存','庫存金額','庫齡(天)','DSI','GMROI','售磬率','ROP','安全庫存','EOQ','前置時間','單位成本','單位售價'];
    const rows = aggregatedData.products.map(p => [
        p.id, p.name, p.category, p.abc, p.xyz, p.abcxyz,
        p.totalCurrent, p.totalCurrentValue.toFixed(0), p.avgAging.toFixed(1),
        (p.totalCurrentValue>0 ? (p.totalCurrentValue/aggregatedData.totalCOGS*CFG.DAYS).toFixed(1) : 0),
        p.gmroi.toFixed(3), (p.str*100).toFixed(1),
        p.rop.toFixed(0), p.safetyStock.toFixed(0), p.eoq.toFixed(0),
        p.leadTime, p.unitCost, p.unitPrice
    ]);
    const csv = [headers, ...rows].map(r => r.map(v => '"' + String(v).replace(/"/g,'""') + '"').join(',')).join('\n');
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = '庫存品項分析_' + new Date().toISOString().slice(0,10) + '.csv';
    a.click(); URL.revokeObjectURL(url);
}

// ══ 26. EXPORT: EXCEL ══
function exportExcel() {
    if (!aggregatedData.products || typeof XLSX === 'undefined') {
        alert('SheetJS 未載入，請確認網路連線');
        return;
    }
    const wb = XLSX.utils.book_new();
    const arr = aggregatedData.products;

    // Sheet 1: KPIs
    const kpiData = [
        ['指標', '數值', '說明'],
        ['庫存總金額(元)', aggregatedData.totalVal?.toFixed(0), '所有品項現有庫存價值'],
        ['DSI(天)', aggregatedData._dsi?.toFixed(1), '庫存天數'],
        ['存貨周轉率(次/年)', aggregatedData._turnover?.toFixed(2), '年周轉次數'],
        ['庫存健康分', aggregatedData._health?.toFixed(0), '0-100分'],
        ['GMROI(倍)', aggregatedData._gmroi?.toFixed(2), '毛利庫存投資報酬率'],
        ['缺貨品項數', aggregatedData._stockouts, '低於ROP的品項'],
        ['>90天庫齡佔比(%)', aggregatedData._aging90Pct?.toFixed(1), '高風險庫存比例'],
        ['Fisher季節性指數', aggregatedData.fisherRatio?.toFixed(2), '>6.5表示具季節性']
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(kpiData), 'KPI總覽');

    // Sheet 2: Products
    const prodHeaders = ['ProductID','名稱','品類','ABC','XYZ','ABCXYZ','現有庫存','庫存金額(元)','庫齡(天)','GMROI','售磬率','前置時間(天)','安全庫存','ROP','EOQ','年需求量','單位成本','單位售價'];
    const prodRows = arr.map(p => [
        p.id, p.name, p.category, p.abc, p.xyz, p.abcxyz,
        p.totalCurrent, Math.round(p.totalCurrentValue), Math.round(p.avgAging),
        +p.gmroi.toFixed(3), +(p.str*100).toFixed(1),
        p.leadTime, Math.round(p.safetyStock), Math.round(p.rop), Math.round(p.eoq),
        Math.round(p.annualDemand), p.unitCost, p.unitPrice
    ]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([prodHeaders, ...prodRows]), '品項分析');

    // Sheet 3: ROP
    const ropHeaders = ['ProductID','名稱','日均需求','前置時間(天)','安全庫存','ROP','現有庫存','狀態'];
    const ropRows = arr.map(p => [
        p.id, p.name, +p.dailySalesAvg.toFixed(2), p.leadTime,
        Math.round(p.safetyStock), Math.round(p.rop), p.totalCurrent,
        p.totalCurrent < p.rop ? '⚠低於ROP' : '✓正常'
    ]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([ropHeaders, ...ropRows]), 'ROP補貨');

    // Sheet 4: Actions
    const actions = [];
    arr.forEach(p => {
        if (p.totalCurrent < p.rop) actions.push(['缺貨', p.id, p.name, '庫存' + p.totalCurrent + '低於ROP ' + Math.round(p.rop)]);
        if (p.avgAging > 90) actions.push(['高庫齡', p.id, p.name, '平均庫齡' + p.avgAging.toFixed(0) + '天']);
        if (p.gmroi < 0.5 && p.totalCurrentValue > 10000) actions.push(['低GMROI', p.id, p.name, 'GMROI=' + p.gmroi.toFixed(2)]);
    });
    const actHeaders = ['類型','ProductID','名稱','說明'];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([actHeaders, ...actions]), '策略行動');

    XLSX.writeFile(wb, '庫存分析_' + new Date().toISOString().slice(0,10) + '.xlsx');
}

// ══ 27. DARK MODE ══
function toggleDarkMode() {
    darkMode = !darkMode;
    document.body.classList.toggle('light-mode', !darkMode);
    localStorage.setItem('v3DarkMode', darkMode ? 'true' : 'false');
    const icon = document.querySelector('#btn-darkmode i[data-lucide]');
    if (icon) { icon.setAttribute('data-lucide', darkMode ? 'moon' : 'sun'); lucide.createIcons(); }
    renderedPages.clear();
    Object.values(myCharts).forEach(c => c.destroy());
    myCharts = {};
    renderCurrentTab();
}

// ══ 28. FILE UPLOAD ══
function handleFileUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            if (typeof XLSX === 'undefined') { alert('SheetJS未載入'); return; }
            const wb = XLSX.read(evt.target.result, { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { defval: 0 });
            if (!rows || rows.length === 0) { alert('讀取失敗：工作表無資料'); return; }

            // Normalize column names
            const normalized = rows.map(r => {
                const norm = {};
                Object.keys(r).forEach(k => {
                    const lk = k.trim().toLowerCase().replace(/\s+/g,'_');
                    norm[lk] = r[k];
                });
                return {
                    ProductID: norm['productid'] || norm['product_id'] || norm['pid'] || norm['品項id'] || norm['品號'] || '—',
                    ProductName: norm['productname'] || norm['product_name'] || norm['名稱'] || norm['品名'] || norm['productid'] || '未知',
                    Category: norm['category'] || norm['品類'] || norm['分類'] || '未分類',
                    Date: norm['date'] ? (norm['date'] instanceof Date ? norm['date'].toISOString().slice(0,10) : String(norm['date']).slice(0,10)) : '—',
                    in_qty: parseFloat(norm['in_qty'] || norm['入庫'] || norm['in'] || 0),
                    out_qty: parseFloat(norm['out_qty'] || norm['出庫'] || norm['out'] || 0),
                    current_inventory: parseFloat(norm['current_inventory'] || norm['庫存'] || norm['inventory'] || 0),
                    aging: parseFloat(norm['aging'] || norm['庫齡'] || 0),
                    UnitCost: parseFloat(norm['unitcost'] || norm['unit_cost'] || norm['成本'] || 0),
                    UnitPrice: parseFloat(norm['unitprice'] || norm['unit_price'] || norm['售價'] || 0),
                    LeadTime: parseFloat(norm['leadtime'] || norm['lead_time'] || norm['前置時間'] || 14)
                };
            });

            const valid = normalized.filter(r => r.ProductID && r.ProductID !== '—' && r.Date && r.Date !== '—');
            if (valid.length === 0) { alert('格式不符：無法識別必要欄位(ProductID, Date)'); return; }
            processData(valid);
            alert('成功載入 ' + valid.length + ' 筆資料，' + [...new Set(valid.map(r=>r.ProductID))].length + ' 個品項');
        } catch(err) {
            alert('檔案解析失敗：' + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
}

// ══ 29. ACCORDION ══
function toggleAccordion(header) {
    const body = header.nextElementSibling;
    const icon = header.querySelector('[data-lucide]');
    const isOpen = body.classList.contains('open');
    body.classList.toggle('open', !isOpen);
    body.style.maxHeight = isOpen ? '0' : body.scrollHeight + 'px';
    if (icon) { icon.setAttribute('data-lucide', isOpen ? 'chevron-down' : 'chevron-up'); lucide.createIcons(); }
}

// ══ 30. CLOCK ══
function updateClock() {
    const el = document.getElementById('header-clock');
    if (el) el.textContent = new Date().toLocaleString('zh-TW', { hour12:false, year:'numeric', month:'2-digit', day:'2-digit', hour:'2-digit', minute:'2-digit', second:'2-digit' });
}

// ══ 31. EVENT LISTENERS ══
document.addEventListener('DOMContentLoaded', () => {

    // Apply dark mode on load
    if (!darkMode) document.body.classList.add('light-mode');

    // Clock
    updateClock();
    setInterval(updateClock, 1000);

    // Tab navigation
    document.querySelectorAll('.glass-btn[data-tab]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.glass-btn[data-tab]').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
            btn.classList.add('active');
            const tabEl = document.getElementById(btn.dataset.tab);
            if (tabEl) tabEl.classList.add('active');
            renderCurrentTab();
        });
    });

    // Dark mode buttons
    document.getElementById('btn-darkmode')?.addEventListener('click', toggleDarkMode);
    document.getElementById('btn-darkmode-settings')?.addEventListener('click', toggleDarkMode);

    // File upload
    document.getElementById('file-upload')?.addEventListener('change', handleFileUpload);

    // Drop zone
    const dropZone = document.getElementById('drop-zone');
    if (dropZone) {
        dropZone.addEventListener('click', () => document.getElementById('file-upload')?.click());
        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.style.borderColor = '#6366f1'; });
        dropZone.addEventListener('dragleave', () => { dropZone.style.borderColor = ''; });
        dropZone.addEventListener('drop', e => {
            e.preventDefault();
            dropZone.style.borderColor = '';
            const file = e.dataTransfer.files[0];
            if (file) {
                const fakeEvent = { target: { files: [file], value: '' } };
                handleFileUpload(fakeEvent);
            }
        });
    }

    // SSR toggle
    document.getElementById('btn-ssr-long')?.addEventListener('click', () => {
        ssrMode = 'long';
        document.getElementById('btn-ssr-long').classList.add('active');
        document.getElementById('btn-ssr-short').classList.remove('active');
        renderedPages.delete('tab-trend');
        renderSSRChart();
    });
    document.getElementById('btn-ssr-short')?.addEventListener('click', () => {
        ssrMode = 'short';
        document.getElementById('btn-ssr-short').classList.add('active');
        document.getElementById('btn-ssr-long').classList.remove('active');
        renderedPages.delete('tab-trend');
        renderSSRChart();
    });

    // ABC mode toggle
    document.getElementById('btn-abc-rel')?.addEventListener('click', () => {
        abcMode = 'relative';
        document.getElementById('btn-abc-rel').classList.add('active');
        document.getElementById('btn-abc-abs').classList.remove('active');
        renderedPages.delete('tab-structure');
        renderStructure();
    });
    document.getElementById('btn-abc-abs')?.addEventListener('click', () => {
        abcMode = 'absolute';
        document.getElementById('btn-abc-abs').classList.add('active');
        document.getElementById('btn-abc-rel').classList.remove('active');
        renderedPages.delete('tab-structure');
        renderStructure();
    });

    // Service level buttons
    const svcMap = { 'btn-svc-90': 1.282, 'btn-svc-95': 1.645, 'btn-svc-99': 2.326 };
    Object.entries(svcMap).forEach(([id, z]) => {
        document.getElementById(id)?.addEventListener('click', () => {
            svcLevel = z;
            Object.keys(svcMap).forEach(bid => document.getElementById(bid)?.classList.remove('active'));
            document.getElementById(id)?.classList.add('active');
            if (rawData.length > 0) processData(rawData);
        });
    });

    // Loss rate update
    document.getElementById('btn-update-loss')?.addEventListener('click', () => {
        if (rawData.length > 0) processData(rawData);
    });

    // Items filters
    ['item-search','item-cat-filter','item-abc-filter','item-xyz-filter','item-sort'].forEach(id => {
        document.getElementById(id)?.addEventListener('change', () => { currentPage = 1; applyItemFilters(); });
    });
    document.getElementById('item-search')?.addEventListener('input', () => { currentPage = 1; applyItemFilters(); });

    // Pagination
    document.getElementById('btn-prev-page')?.addEventListener('click', () => {
        if (currentPage > 1) { currentPage--; renderItemsPage(); }
    });
    document.getElementById('btn-next-page')?.addEventListener('click', () => {
        const total = Math.ceil(filteredItems.length / rowsPerPage);
        if (currentPage < total) { currentPage++; renderItemsPage(); }
    });
    document.getElementById('items-per-page')?.addEventListener('change', () => {
        rowsPerPage = parseInt(document.getElementById('items-per-page').value);
        currentPage = 1;
        renderItemsPage();
    });

    // CSV export from items tab
    document.getElementById('btn-export-items-csv')?.addEventListener('click', exportCSV);

    // Threshold update
    document.getElementById('btn-update-thresholds')?.addEventListener('click', () => {
        alertCfg.dsiMax = parseFloat(document.getElementById('threshold-dsi')?.value || 90);
        alertCfg.agingMax = parseFloat(document.getElementById('threshold-aging')?.value || 20);
        alertCfg.gmroiMin = parseFloat(document.getElementById('threshold-gmroi')?.value || 1.2);
        alertCfg.imbMax = parseFloat(document.getElementById('threshold-imbalance')?.value || 0.3);
        renderedPages.delete('tab-actions');
        renderActions();
    });

    // What-if sliders
    const sliderConfig = [
        { id:'slider-growth', valId:'val-growth', param:'growth', fmt: v => (v>=0?'+':'') + v + '%' },
        { id:'slider-svc',    valId:'val-svc',    param:'svc',    fmt: v => v + '%' },
        { id:'slider-hold',   valId:'val-hold',   param:'hold',   fmt: v => v + '%' },
        { id:'slider-lead',   valId:'val-lead',   param:'lead',   fmt: v => (v>=0?'+':'') + v + '天' }
    ];
    sliderConfig.forEach(cfg => {
        document.getElementById(cfg.id)?.addEventListener('input', e => {
            const v = parseInt(e.target.value);
            whatIfParams[cfg.param] = v;
            const el = document.getElementById(cfg.valId);
            if (el) el.textContent = cfg.fmt(v);
            updateWhatIf();
        });
    });

    // EOQ calc button
    document.getElementById('btn-calc-eoq')?.addEventListener('click', () => {
        CFG.ORDER_COST = parseFloat(document.getElementById('eoq-order-cost')?.value || 500);
        CFG.HOLDING_RATE = parseFloat(document.getElementById('eoq-holding-rate')?.value || 25) / 100;
        renderedPages.delete('tab-eoq');
        renderEOQ();
    });

    // FIFO product selector
    document.getElementById('fifo-product-select')?.addEventListener('change', e => {
        renderFIFOProduct(e.target.value);
    });

    // Report buttons
    document.getElementById('btn-pdf')?.addEventListener('click', exportPDF);
    document.getElementById('btn-excel')?.addEventListener('click', exportExcel);
    document.getElementById('btn-copy-report')?.addEventListener('click', () => {
        const txt = document.getElementById('report-text')?.textContent || '';
        navigator.clipboard?.writeText(txt).then(() => alert('報告已複製到剪貼板')).catch(() => alert('複製失敗，請手動選取'));
    });
    document.getElementById('btn-save-snapshot')?.addEventListener('click', saveSnapshot);

    // Settings buttons
    document.getElementById('btn-load-demo')?.addEventListener('click', generateDemoData);
    document.getElementById('btn-export-json')?.addEventListener('click', exportJSON);
    document.getElementById('btn-export-csv')?.addEventListener('click', exportCSV);
    document.getElementById('btn-reset')?.addEventListener('click', () => {
        if (confirm('確定要重置所有資料嗎？')) {
            rawData = []; aggregatedData = {};
            Object.values(myCharts).forEach(c=>c.destroy()); myCharts = {};
            renderedPages.clear();
            document.querySelectorAll('[id^="kpi-"]').forEach(el => el.textContent = '—');
            generateDemoData();
        }
    });

    // Header download
    document.getElementById('btn-header-download')?.addEventListener('click', exportExcel);

    // Init lucide icons
    if (typeof lucide !== 'undefined') lucide.createIcons();

    // Auto-load demo data
    generateDemoData();
});
"""

content.write(js)
content.close()
print("Part 4 written:", len(js), "chars")
