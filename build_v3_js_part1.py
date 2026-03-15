content = open('C:/Users/jamic/庫存分析/v3_parts/part_js_p1.txt', 'w', encoding='utf-8')

js = """
// ══════════════════════════════════════════════════════
// 庫存分析智能儀表板 V3.0 — JavaScript
// ══════════════════════════════════════════════════════

// ══ 1. CONSTANTS & STATE ══
const CFG = {
    DAYS: 180,
    ORDER_COST: 500,
    HOLDING_RATE: 0.25,
    SVC_LEVEL_Z: 1.645,
    ABC_A: 0.5, ABC_B: 0.7,
    XYZ_X: 0.5, XYZ_Y: 1.0,
    LOSS_RATES: [0, 0, 0.10, 0.30, 0.60],
    ALERT: { DSI_MAX: 90, AGING_MAX: 60, GMROI_MIN: 0.5, IMB_MAX: 0.10, STOCKOUT_MAX: 5, EOQ_DEV_MAX: 0.30 }
};
let rawData = [], aggregatedData = {}, myCharts = {}, renderedPages = new Set();
let darkMode = localStorage.getItem('v3DarkMode') !== 'false';
let currentPage = 1, rowsPerPage = 20, filteredItems = [];
let svcLevel = 1.645;
let ssrMode = 'long';
let abcMode = 'relative';
let snapshots = JSON.parse(localStorage.getItem('v3Snapshots') || '[]');
let alertCfg = { dsiMax: 90, agingMax: 60, gmroiMin: 0.5, imbMax: 0.10 };
let whatIfParams = { growth: 10, svc: 95, hold: 25, lead: 0 };

// ══ 2. DEMO DATA GENERATOR ══
function generateDemoData() {
    const categories = ['電子產品','服飾配件','家用電器','家具寢具','玩具文具'];
    const productNames = {
        '電子產品': ['智慧型手機','藍牙耳機','平板電腦','智慧手錶','無線充電器','筆記型電腦','電競鍵盤','4K顯示器','USB集線器','網路攝影機'],
        '服飾配件': ['男士西裝','女士洋裝','牛仔褲','棉質T恤','皮革皮帶','太陽眼鏡','運動球鞋','羊毛圍巾','真皮手包','帆布背包'],
        '家用電器': ['空氣清淨機','智慧冰箱','滾筒洗衣機','微波爐','電磁爐','掃地機器人','電熱水瓶','電動牙刷','除濕機','烤箱'],
        '家具寢具': ['雙人記憶床墊','木質書架','L型沙發','辦公椅','衣帽架','棉麻床單組','遮光窗簾','儲物抽屜櫃','折疊餐桌','浴室置物架'],
        '玩具文具': ['積木套組','遙控飛機','繪圖水彩組','兒童益智拼圖','鋼筆禮盒','筆記本套裝','桌遊卡牌','手工黏土','彩色鉛筆','文具收納盒']
    };
    const unitCostBase = { '電子產品':800, '服飾配件':200, '家用電器':600, '家具寢具':400, '玩具文具':80 };
    const leadTimeBase = { '電子產品':21, '服飾配件':14, '家用電器':30, '家具寢具':45, '玩具文具':7 };

    const startDate = new Date('2025-09-16');
    const rows = [];
    let pid = 1001;

    // Simple seeded random
    let seed = 42;
    function rnd() { seed = (seed * 1664525 + 1013904223) & 0xffffffff; return (seed >>> 0) / 0xffffffff; }

    categories.forEach(cat => {
        const names = productNames[cat];
        for (let pi = 0; pi < 10; pi++) {
            const productId = 'P' + pid++;
            const baseCost = unitCostBase[cat] * (0.7 + rnd() * 0.6);
            const unitCost = Math.round(baseCost * 10) / 10;
            const unitPrice = Math.round(unitCost * (1.3 + rnd() * 1.2) * 10) / 10;
            const baseDemand = 2 + rnd() * 8;
            const leadTime = leadTimeBase[cat] + Math.floor((rnd() - 0.5) * 10);
            let inventory = Math.floor(50 + rnd() * 150);
            let aging = Math.floor(rnd() * 30);

            for (let d = 0; d < CFG.DAYS; d++) {
                const date = new Date(startDate);
                date.setDate(startDate.getDate() + d);
                const dateStr = date.toISOString().slice(0,10);
                const month = date.getMonth() + 1;
                const seasonFactor = 1 + 0.3 * Math.sin((month - 3) * Math.PI / 6);
                const dailyDemand = baseDemand * seasonFactor;

                let in_qty = 0;
                if (rnd() < 0.05 || inventory < 10) {
                    in_qty = Math.floor(50 + rnd() * 150);
                    inventory += in_qty;
                    aging = Math.max(0, aging - Math.floor(aging * 0.4));
                }

                let rawDemand = dailyDemand * (0.5 + rnd());
                let out_qty = Math.min(inventory, Math.max(0, Math.floor(rawDemand)));
                inventory = Math.max(0, inventory - out_qty);
                aging = in_qty > 0 ? aging : Math.min(aging + 1, 365);

                rows.push({
                    ProductID: productId,
                    ProductName: names[pi],
                    Category: cat,
                    Date: dateStr,
                    in_qty: in_qty,
                    out_qty: out_qty,
                    current_inventory: inventory,
                    aging: aging,
                    UnitCost: unitCost,
                    UnitPrice: unitPrice,
                    LeadTime: leadTime
                });
            }
        }
    });

    processData(rows);
}

// ══ 3. PROCESS DATA ══
function processData(raw) {
    rawData = raw;
    const ps = {}, ms = {};

    raw.forEach(r => {
        const pid = r.ProductID;
        if (!ps[pid]) {
            ps[pid] = {
                id: pid, name: r.ProductName || pid, category: r.Category || '未分類',
                totalSales: 0, totalCOGS: 0, totalIn: 0, totalRev: 0,
                totalCurrent: 0, totalCurrentValue: 0,
                agingSum: 0, agingCount: 0,
                leadTime: r.LeadTime || 14,
                unitCost: r.UnitCost || 100,
                unitPrice: r.UnitPrice || 150,
                monthlyOut: {}, dailySalesArr: [], batches: []
            };
        }
        const p = ps[pid];
        p.totalSales += r.out_qty || 0;
        p.totalCOGS += (r.out_qty || 0) * (r.UnitCost || 0);
        p.totalIn += r.in_qty || 0;
        p.totalRev += (r.out_qty || 0) * (r.UnitPrice || 0);
        p.agingSum += (r.aging || 0);
        p.agingCount++;
        p.dailySalesArr.push(r.out_qty || 0);
        if ((r.in_qty || 0) > 0) {
            p.batches.push({ date: r.Date, qty: r.in_qty, cost: r.UnitCost || 0 });
        }
        const mo = r.Date ? r.Date.slice(0,7) : '2025-09';
        p.monthlyOut[mo] = (p.monthlyOut[mo] || 0) + (r.out_qty || 0);

        if (!ms[mo]) ms[mo] = { invVal:0, cogs:0, rev:0, count:0, totalIn:0, totalOut:0 };
        ms[mo].invVal += (r.current_inventory || 0) * (r.UnitCost || 0);
        ms[mo].cogs += (r.out_qty || 0) * (r.UnitCost || 0);
        ms[mo].rev += (r.out_qty || 0) * (r.UnitPrice || 0);
        ms[mo].count++;
        ms[mo].totalIn += r.in_qty || 0;
        ms[mo].totalOut += r.out_qty || 0;
    });

    const latestByPid = {};
    raw.forEach(r => {
        if (!latestByPid[r.ProductID] || r.Date > latestByPid[r.ProductID].Date) {
            latestByPid[r.ProductID] = r;
        }
    });
    Object.keys(ps).forEach(pid => {
        const lr = latestByPid[pid];
        ps[pid].totalCurrent = lr ? (lr.current_inventory || 0) : 0;
        ps[pid].totalCurrentValue = ps[pid].totalCurrent * (ps[pid].unitCost || 0);
        ps[pid].avgAging = ps[pid].agingCount > 0 ? ps[pid].agingSum / ps[pid].agingCount : 0;
    });

    let arr = Object.values(ps);
    const totalVal = arr.reduce((s,p) => s + p.totalCurrentValue, 0);
    const totalCOGS = arr.reduce((s,p) => s + p.totalCOGS, 0);

    arr.sort((a,b) => b.totalCurrentValue - a.totalCurrentValue);
    let cumVal = 0;
    const lossRates = readLossRates();
    arr.forEach(p => {
        cumVal += p.totalCurrentValue;
        p.abcPct = totalVal > 0 ? cumVal / totalVal : 0;
        p.abc = p.abcPct <= CFG.ABC_A ? 'A' : (p.abcPct <= CFG.ABC_B ? 'B' : 'C');
        p.valPct = totalVal > 0 ? p.totalCurrentValue / totalVal : 0;
        p.cogsPct = totalCOGS > 0 ? p.totalCOGS / totalCOGS : 0;
        p.imbalance = p.valPct - p.cogsPct;

        const mo = Object.values(p.monthlyOut || {});
        const avg = mo.length > 0 ? mo.reduce((s,v)=>s+v,0)/mo.length : 0;
        const std = Math.sqrt(mo.reduce((s,v)=>s+(v-avg)**2,0)/(mo.length||1));
        p.coV = avg > 0 ? std/avg : 0;
        p.xyz = p.coV < CFG.XYZ_X ? 'X' : (p.coV < CFG.XYZ_Y ? 'Y' : 'Z');
        p.abcxyz = p.abc + p.xyz;

        p.strFast = p.totalIn > 0 ? p.totalSales/p.totalIn : 0;
        p.strSlow = (p.totalSales + p.totalCurrent) > 0 ? p.totalSales/(p.totalSales+p.totalCurrent) : 0;
        p.strType = p.leadTime > 21 ? 'slow' : 'fast';
        p.str = p.strType === 'slow' ? p.strSlow : p.strFast;

        p.dailySalesAvg = p.totalSales / CFG.DAYS;
        const dailyStd = p.dailySalesAvg * 0.3;
        p.safetyStock = svcLevel * Math.sqrt(p.leadTime) * dailyStd;
        p.rop = p.dailySalesAvg * p.leadTime + p.safetyStock;

        p.annualDemand = p.dailySalesAvg * 365;
        const hcRate = CFG.HOLDING_RATE;
        const hc = p.unitCost * hcRate;
        p.eoq = hc > 0 ? Math.sqrt(2 * p.annualDemand * CFG.ORDER_COST / hc) : 0;
        const avgBatch = p.batches.length > 0
            ? p.batches.reduce((s,b)=>s+b.qty,0)/p.batches.length
            : (p.eoq || 100);
        p.eoqDeviation = p.eoq > 0 ? Math.abs(avgBatch - p.eoq) / p.eoq : 0;

        p.grossMargin = p.unitPrice > 0 ? (p.unitPrice - p.unitCost)/p.unitPrice : 0;
        p.gmroi = p.totalCurrentValue > 0 ? (p.totalRev * p.grossMargin)/p.totalCurrentValue : 0;

        const a = p.avgAging;
        p.agingBucket = a<=30?0:(a<=60?1:(a<=90?2:(a<=120?3:4)));
        p.lossRate = lossRates[p.agingBucket];
        p.estimatedLoss = p.totalCurrentValue * p.lossRate;
    });

    const mKeys = Object.keys(ms).sort();
    mKeys.forEach((mo, i) => {
        const m = ms[mo];
        m.avgInv = m.count > 0 ? m.invVal / m.count : 0;
        m.turnover = m.avgInv > 0 ? m.cogs / m.avgInv : 0;
        m.dsi = m.turnover > 0 ? 30/m.turnover : 30;
        m.ssrLong = m.rev > 0 ? m.avgInv / m.rev : 0;
        const prevAvg = i > 0 ? ms[mKeys[i-1]].avgInv : m.avgInv;
        m.ssrShort = m.totalOut > 0 ? (prevAvg + m.totalIn) / m.totalOut : 0;
        m.growth = i > 0 && ms[mKeys[i-1]].cogs > 0 ? (m.cogs - ms[mKeys[i-1]].cogs)/ms[mKeys[i-1]].cogs : 0;
        if (i >= 2) m.ma3 = (ms[mKeys[i]].totalOut + ms[mKeys[i-1]].totalOut + ms[mKeys[i-2]].totalOut) / 3;
        if (i >= 5) m.ma6 = mKeys.slice(i-5, i+1).reduce((s,k) => s + ms[k].totalOut, 0) / 6;
    });

    const monthlyGroups = {};
    mKeys.forEach(mo => {
        const mn = parseInt(mo.split('-')[1]);
        if (!monthlyGroups[mn]) monthlyGroups[mn] = [];
        monthlyGroups[mn].push(ms[mo].totalOut);
    });
    const grandAvgOut = mKeys.length > 0 ? mKeys.reduce((s,k) => s + ms[k].totalOut, 0) / mKeys.length : 1;
    const seasonalIndex = {};
    for (let m = 1; m <= 12; m++) {
        const vals = monthlyGroups[m] || [grandAvgOut];
        seasonalIndex[m] = (vals.reduce((s,v)=>s+v,0)/vals.length) / (grandAvgOut||1);
    }

    const monthAvgs = Object.values(monthlyGroups).map(v => v.reduce((s,x)=>s+x,0)/v.length);
    const mGrand = monthAvgs.reduce((s,v)=>s+v,0) / (monthAvgs.length||1);
    const monthStd = Math.sqrt(monthAvgs.reduce((s,v)=>s+(v-mGrand)**2,0)/(monthAvgs.length||1));
    const yearVals = mKeys.map(k => ms[k].totalOut);
    const yGrand = yearVals.reduce((s,v)=>s+v,0)/(yearVals.length||1);
    const yearStd = Math.sqrt(yearVals.reduce((s,v)=>s+(v-yGrand)**2,0)/(yearVals.length||1));
    const fisherRatio = yearStd > 0 ? monthStd/yearStd : 0;

    const qualityIssues = checkDataQuality(raw);

    aggregatedData = { products: arr, monthly: ms, mKeys, totalVal, totalCOGS, seasonalIndex, fisherRatio, qualityIssues };
    updateDashboard();
}

function readLossRates() {
    const get = (id, def) => (parseFloat(document.getElementById(id)?.value) || def) / 100;
    return [get('loss-rate-0',0), get('loss-rate-30',5), get('loss-rate-60',15), get('loss-rate-90',30), get('loss-rate-180',60)];
}

// ══ 4. UPDATE DASHBOARD ══
function updateDashboard() {
    if (!aggregatedData.products) return;
    const arr = aggregatedData.products;
    const totalInv = aggregatedData.totalVal || 0;
    const totalCOGS = aggregatedData.totalCOGS || 0;

    const aging90Val = arr.filter(p=>p.avgAging>90).reduce((s,p)=>s+p.totalCurrentValue,0);
    const aging90Pct = totalInv > 0 ? aging90Val/totalInv*100 : 0;
    const dsi = totalCOGS > 0 ? (totalInv / totalCOGS) * CFG.DAYS : 0;
    const turnover = totalInv > 0 && totalCOGS > 0 ? totalCOGS / totalInv * (365/CFG.DAYS) : 0;
    const stockouts = arr.filter(p => p.totalCurrent < p.rop).length;
    const totalGMROI = totalInv > 0 ? arr.reduce((s,p) => s + p.gmroi * p.totalCurrentValue, 0) / totalInv : 0;
    const avgSTR = arr.reduce((s,p) => s + p.str * p.valPct, 0);

    let health = 100;
    if (dsi > 90) health -= 20; else if (dsi > 60) health -= 10;
    if (aging90Pct > 30) health -= 15; else if (aging90Pct > 15) health -= 7;
    if (stockouts > 5) health -= 10; else if (stockouts > 2) health -= 5;
    if (totalGMROI < 0.5) health -= 15; else if (totalGMROI < 1) health -= 10;
    health = Math.max(0, Math.min(100, health));

    const setText = (id, val) => { const el=document.getElementById(id); if(el) el.textContent=val; };
    const setWidth = (id, pct) => { const el=document.getElementById(id); if(el) el.style.width=Math.min(100,Math.max(0,pct))+'%'; };

    setText('kpi-total-inv', (totalInv/10000).toFixed(1));
    setWidth('pb-total-inv', Math.min(100, (totalInv/10000)/500*100));
    setText('trend-total-inv', totalInv > 0 ? '庫存在管' : '無資料');

    setText('kpi-dsi', dsi.toFixed(1));
    setWidth('pb-dsi', Math.min(100, dsi/120*100));
    setText('trend-dsi', dsi > 90 ? '⚠ 偏高' : dsi < 30 ? '✓ 良好' : '— 正常');

    setText('kpi-aging', aging90Pct.toFixed(1));
    setWidth('pb-aging', aging90Pct);
    setText('trend-aging', aging90Pct > 20 ? '⚠ 注意' : '✓ 正常');

    setText('kpi-health', health.toFixed(0));
    setWidth('pb-health', health);
    setText('trend-health', health >= 80 ? '✓ 健康' : health >= 60 ? '— 待改善' : '⚠ 警示');

    setText('kpi-turnover', turnover.toFixed(2));
    setWidth('pb-turnover', Math.min(100, turnover/10*100));
    setText('trend-turnover', turnover > 6 ? '✓ 高效' : '— 普通');

    setText('kpi-str', avgSTR.toFixed(2));
    setWidth('pb-str', Math.min(100, avgSTR*100));
    setText('trend-str', avgSTR > 0.7 ? '✓ 快銷' : '— 滯銷');

    setText('kpi-gmroi', totalGMROI.toFixed(2));
    setWidth('pb-gmroi', Math.min(100, totalGMROI/3*100));
    setText('trend-gmroi', totalGMROI > 1.5 ? '✓ 優良' : totalGMROI > 0.5 ? '— 正常' : '⚠ 偏低');

    setText('kpi-stockouts', stockouts);
    setWidth('pb-stockouts', Math.min(100, stockouts/Math.max(arr.length,1)*100));
    setText('trend-stockouts', stockouts > 5 ? '⚠ 過多' : '✓ 正常');

    setText('supply-gmroi', totalGMROI.toFixed(2));
    const stockoutLoss = arr.filter(p=>p.totalCurrent<p.rop).reduce((s,p)=>s+p.dailySalesAvg*p.unitPrice*p.leadTime,0);
    setText('supply-stockout-loss', (stockoutLoss/10000).toFixed(1));
    const avgLead = arr.reduce((s,p)=>s+p.leadTime,0)/(arr.length||1);
    setText('supply-avg-leadtime', avgLead.toFixed(1));

    aggregatedData._health = health;
    aggregatedData._dsi = dsi;
    aggregatedData._turnover = turnover;
    aggregatedData._stockouts = stockouts;
    aggregatedData._gmroi = totalGMROI;
    aggregatedData._str = avgSTR;
    aggregatedData._aging90Pct = aging90Pct;

    renderedPages.clear();
    renderCurrentTab();
    generateTextReport();
    renderSnapshotsList();
}

// ══ 5. TAB NAVIGATION ══
function renderCurrentTab() {
    const active = document.querySelector('.tab-content.active');
    if (!active) return;
    const id = active.id;
    if (renderedPages.has(id)) return;
    renderedPages.add(id);
    const map = {
        'tab-overview': renderOverview,
        'tab-trend': renderTrend,
        'tab-structure': renderStructure,
        'tab-aging': renderAging,
        'tab-forecast': renderForecast,
        'tab-supply': renderSupply,
        'tab-items': renderItems,
        'tab-actions': renderActions,
        'tab-whatif': renderWhatIf,
        'tab-eoq': renderEOQ,
        'tab-fifo': renderFIFO,
        'tab-quality': renderQuality,
        'tab-report': renderReport,
        'tab-settings': renderSettings
    };
    if (map[id]) map[id]();
}

// ══ 6. CHART HELPER ══
function cc(id, type, data, opts) {
    if (myCharts[id]) { myCharts[id].destroy(); delete myCharts[id]; }
    const canvas = document.getElementById(id);
    if (!canvas) return null;
    const isDark = !document.body.classList.contains('light-mode');
    const gridColor = isDark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
    const tickColor = isDark ? '#94a3b8' : '#64748b';
    const isCircular = type === 'pie' || type === 'doughnut' || type === 'polarArea';
    const defaults = {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: tickColor, font:{size:10}, boxWidth:12 } }, tooltip: { mode:'index', intersect:false } },
        scales: type === 'radar' ? {
            r: { ticks:{color:tickColor,font:{size:9},backdropColor:'transparent'}, pointLabels:{color:tickColor,font:{size:10}}, grid:{color:gridColor}, angleLines:{color:gridColor} }
        } : isCircular ? {} : {
            x: { ticks:{color:tickColor,font:{size:9}}, grid:{color:gridColor} },
            y: { ticks:{color:tickColor,font:{size:9}}, grid:{color:gridColor} }
        }
    };
    const mergedOpts = deepMerge(defaults, opts || {});
    myCharts[id] = new Chart(canvas, { type, data, options: mergedOpts });
    return myCharts[id];
}

function deepMerge(target, source) {
    const out = Object.assign({}, target);
    for (const key in source) {
        if (source[key] !== null && typeof source[key] === 'object' && !Array.isArray(source[key])) {
            out[key] = deepMerge(target[key] || {}, source[key]);
        } else {
            out[key] = source[key];
        }
    }
    return out;
}

const PALETTE = ['#6366f1','#22d3ee','#10b981','#f59e0b','#f43f5e','#8b5cf6','#06b6d4','#84cc16','#fb923c','#ec4899'];
const CAT_COLORS = {'電子產品':'#6366f1','服飾配件':'#22d3ee','家用電器':'#10b981','家具寢具':'#f59e0b','玩具文具':'#f43f5e'};
"""

content.write(js)
content.close()
print("Part 1 written:", len(js), "chars")
