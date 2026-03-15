content = open('C:/Users/jamic/庫存分析/v3_parts/part_js_p3.txt', 'w', encoding='utf-8')

js = """
// ══ 13. RENDER: ITEMS ══
function renderItems() {
    if (!aggregatedData.products) return;
    applyItemFilters();
}

function applyItemFilters() {
    if (!aggregatedData.products) return;
    const search = (document.getElementById('item-search')?.value || '').toLowerCase();
    const cat = document.getElementById('item-cat-filter')?.value || '';
    const abc = document.getElementById('item-abc-filter')?.value || '';
    const xyz = document.getElementById('item-xyz-filter')?.value || '';
    const sort = document.getElementById('item-sort')?.value || 'value_desc';

    filteredItems = aggregatedData.products.filter(p => {
        if (search && !p.name.toLowerCase().includes(search) && !p.id.toLowerCase().includes(search)) return false;
        if (cat && p.category !== cat) return false;
        if (abc && p.abc !== abc) return false;
        if (xyz && p.xyz !== xyz) return false;
        return true;
    });

    const sortMap = {
        'value_desc': (a,b) => b.totalCurrentValue - a.totalCurrentValue,
        'value_asc': (a,b) => a.totalCurrentValue - b.totalCurrentValue,
        'aging_desc': (a,b) => b.avgAging - a.avgAging,
        'aging_asc': (a,b) => a.avgAging - b.avgAging,
        'gmroi_desc': (a,b) => b.gmroi - a.gmroi,
        'str_desc': (a,b) => b.str - a.str
    };
    filteredItems.sort(sortMap[sort] || sortMap['value_desc']);

    rowsPerPage = parseInt(document.getElementById('items-per-page')?.value || '20');
    renderItemsPage();
}

function renderItemsPage() {
    const total = filteredItems.length;
    const totalPages = Math.max(1, Math.ceil(total / rowsPerPage));
    currentPage = Math.min(currentPage, totalPages);
    const start = (currentPage - 1) * rowsPerPage;
    const pageItems = filteredItems.slice(start, start + rowsPerPage);

    const pageInfo = document.getElementById('page-info');
    if (pageInfo) pageInfo.textContent = '第 ' + currentPage + ' / ' + totalPages + ' 頁（共 ' + total + ' 筆）';

    const tbody = document.getElementById('items-table-body');
    if (!tbody) return;

    const abcColor = { A:'text-blue-400', B:'text-amber-400', C:'text-slate-400' };
    const xyzColor = { X:'text-emerald-400', Y:'text-yellow-400', Z:'text-red-400' };

    tbody.innerHTML = pageItems.map(p => {
        const ropStatus = p.totalCurrent < p.rop ? '<span class="text-red-400">⚠ 低</span>' : '<span class="text-emerald-400">✓ 正常</span>';
        return '<tr>' +
            '<td>' + p.category + '</td>' +
            '<td class="text-slate-400">' + p.id + '</td>' +
            '<td>' + p.name + '</td>' +
            '<td class="text-center ' + abcColor[p.abc] + ' font-bold">' + p.abc + '</td>' +
            '<td class="text-center ' + xyzColor[p.xyz] + ' font-bold">' + p.xyz + '</td>' +
            '<td class="text-center font-bold">' + p.abcxyz + '</td>' +
            '<td class="text-right">' + p.totalCurrent.toLocaleString() + '</td>' +
            '<td class="text-right">' + p.totalCurrentValue.toLocaleString() + '</td>' +
            '<td class="text-right ' + (p.avgAging > 90 ? 'text-red-400' : p.avgAging > 60 ? 'text-amber-400' : '') + '">' + p.avgAging.toFixed(0) + '</td>' +
            '<td>' + (p.strType === 'fast' ? '快銷' : '慢銷') + '</td>' +
            '<td class="text-right">' + (p.str*100).toFixed(1) + '%</td>' +
            '<td class="text-right ' + (p.gmroi < 0.5 ? 'text-red-400' : p.gmroi > 1.5 ? 'text-emerald-400' : '') + '">' + p.gmroi.toFixed(2) + '</td>' +
            '<td>' + ropStatus + '</td>' +
            '</tr>';
    }).join('');

    // Update category filter options
    const catSel = document.getElementById('item-cat-filter');
    if (catSel && catSel.options.length <= 1) {
        const cats = [...new Set(aggregatedData.products.map(p => p.category))].sort();
        cats.forEach(cat => {
            const opt = document.createElement('option');
            opt.value = cat; opt.textContent = cat;
            catSel.appendChild(opt);
        });
    }
}

// ══ 14. RENDER: ACTIONS ══
function renderActions() {
    if (!aggregatedData.products) return;
    const arr = aggregatedData.products;
    const actions = [];
    const totalInv = aggregatedData.totalVal || 1;

    // 1. High DSI
    const dsi = aggregatedData._dsi || 0;
    if (dsi > alertCfg.dsiMax) {
        actions.push({ type:'error', icon:'clock', title:'DSI 超標', desc:'庫存天數 ' + dsi.toFixed(0) + ' 天，超過上限 ' + alertCfg.dsiMax + ' 天', action:'建議促銷或調撥過剩庫存' });
    }

    // 2. Aging risk
    const agingRisk = arr.filter(p => p.avgAging > 90);
    if (agingRisk.length > 0) {
        const val = agingRisk.reduce((s,p)=>s+p.totalCurrentValue,0);
        actions.push({ type:'warning', icon:'alert-triangle', title:'高庫齡品項', desc: agingRisk.length + ' 個品項平均庫齡超過90天，合計 ' + (val/10000).toFixed(1) + ' 萬元', action:'建議降價清倉或轉調庫存' });
    }

    // 3. Low GMROI
    const lowGMROI = arr.filter(p => p.gmroi < alertCfg.gmroiMin && p.totalCurrentValue > 10000);
    if (lowGMROI.length > 0) {
        actions.push({ type:'warning', icon:'trending-down', title:'GMROI 偏低', desc: lowGMROI.length + ' 個品項 GMROI < ' + alertCfg.gmroiMin, action:'建議提升毛利或降低庫存持有' });
    }

    // 4. Stockouts
    const stockouts = arr.filter(p => p.totalCurrent < p.rop);
    if (stockouts.length > 0) {
        actions.push({ type:'error', icon:'package', title:'庫存低於ROP', desc: stockouts.length + ' 個品項庫存低於再訂購點', action:'建議立即啟動補貨流程' });
    }

    // 5. Supply-demand imbalance
    const imbalanced = arr.filter(p => Math.abs(p.imbalance) > alertCfg.imbMax);
    if (imbalanced.length > 0) {
        actions.push({ type:'info', icon:'bar-chart-2', title:'供需失衡品項', desc: imbalanced.length + ' 個品項供需失衡指數 > ' + (alertCfg.imbMax*100).toFixed(0) + '%', action:'建議重新評估採購計畫' });
    }

    // 6. EOQ deviation
    const eoqDev = arr.filter(p => p.eoqDeviation > CFG.ALERT.EOQ_DEV_MAX);
    if (eoqDev.length > 0) {
        actions.push({ type:'info', icon:'calculator', title:'EOQ偏離過大', desc: eoqDev.length + ' 個品項批量偏離EOQ超過30%', action:'建議調整訂購批量以降低總成本' });
    }

    // 7. Data quality
    const qualIssues = aggregatedData.qualityIssues || [];
    const critIssues = qualIssues.filter(q => q.severity === 'error');
    if (critIssues.length > 0) {
        actions.push({ type:'error', icon:'shield-alert', title:'資料品質問題', desc: critIssues.length + ' 筆嚴重資料異常需要修正', action:'請至資料品質頁面查看詳情' });
    }

    const actionList = document.getElementById('action-list');
    if (actionList) {
        if (actions.length === 0) {
            actionList.innerHTML = '<div class="text-center text-emerald-400 py-8">✓ 目前無需緊急處理的策略行動</div>';
        } else {
            const typeColors = { error:'border-red-500 bg-red-500/10', warning:'border-amber-500 bg-amber-500/10', info:'border-blue-500 bg-blue-500/10' };
            const typeBadge = { error:'<span class="text-[0.6rem] bg-red-500 text-white px-1.5 py-0.5 rounded">緊急</span>',
                warning:'<span class="text-[0.6rem] bg-amber-500 text-white px-1.5 py-0.5 rounded">警告</span>',
                info:'<span class="text-[0.6rem] bg-blue-500 text-white px-1.5 py-0.5 rounded">建議</span>' };
            actionList.innerHTML = actions.map(a =>
                '<div class="border-l-2 rounded-r-lg p-3 ' + (typeColors[a.type]||'border-slate-500') + '">' +
                '<div class="flex items-center gap-2 mb-1">' + typeBadge[a.type] + '<span class="text-sm font-medium">' + a.title + '</span></div>' +
                '<div class="text-xs text-slate-300 mb-1">' + a.desc + '</div>' +
                '<div class="text-xs text-slate-400">→ ' + a.action + '</div>' +
                '</div>'
            ).join('');
        }
    }

    // Scorecard
    const health = aggregatedData._health || 0;
    const turnover = aggregatedData._turnover || 0;
    const aging90Pct = aggregatedData._aging90Pct || 0;
    const gmroi = aggregatedData._gmroi || 0;
    const str = aggregatedData._str || 0;
    const qualScore = Math.max(0, 100 - (qualIssues.filter(q=>q.severity==='error').length * 10) - (qualIssues.filter(q=>q.severity==='warning').length * 3));
    const supplyScore = Math.max(0, 100 - (stockouts.length / Math.max(arr.length,1)) * 200);

    const setBar = (barId, valId, pct, val) => {
        const b = document.getElementById(barId);
        const v = document.getElementById(valId);
        if (b) b.style.width = Math.min(100, Math.max(0, pct)) + '%';
        if (v) v.textContent = val;
    };
    setBar('sc-health', 'sc-health-val', health, health.toFixed(0) + '分');
    setBar('sc-turn', 'sc-turn-val', Math.min(100, turnover/10*100), turnover.toFixed(2) + '次');
    setBar('sc-aging', 'sc-aging-val', Math.max(0, 100 - aging90Pct*2), (100-aging90Pct*2).toFixed(0) + '分');
    setBar('sc-forecast', 'sc-forecast-val', qualScore, qualScore.toFixed(0) + '分');
    setBar('sc-supply', 'sc-supply-val', supplyScore, supplyScore.toFixed(0) + '分');
    setBar('sc-profit', 'sc-profit-val', Math.min(100, gmroi/3*100), gmroi.toFixed(2) + '倍');
}

// ══ 15. RENDER: WHAT-IF ══
function renderWhatIf() {
    if (!aggregatedData.products) return;
    updateWhatIf();
}

function updateWhatIf() {
    const arr = aggregatedData.products;
    if (!arr) return;
    const totalInv = aggregatedData.totalVal || 0;
    const totalCOGS = aggregatedData.totalCOGS || 0;
    const holdRate = whatIfParams.hold / 100;
    const growthRate = whatIfParams.growth / 100;
    const leadAdj = whatIfParams.lead;
    const svcZ = whatIfParams.svc <= 90 ? 1.282 : whatIfParams.svc <= 95 ? 1.645 : 2.326;

    function calcScenario(gMult, ssZ, holdMult, ltAdj) {
        const projCOGS = totalCOGS * (1 + gMult);
        const projInv = totalInv * (1 + gMult * 0.6);
        const dsi = projCOGS > 0 ? (projInv / projCOGS) * CFG.DAYS : 0;
        const totalSS = arr.reduce((s,p) => {
            const lt = Math.max(1, p.leadTime + ltAdj);
            const dailyStd = p.dailySalesAvg * (1+gMult) * 0.3;
            return s + ssZ * Math.sqrt(lt) * dailyStd * p.unitCost;
        }, 0);
        const holdingCost = projInv * holdMult;
        const gmroi = projInv > 0 ? arr.reduce((s,p)=>s+(p.totalRev*(1+gMult)*p.grossMargin),0) / projInv : 0;
        return { inv: projInv/10000, dsi, ss: totalSS/10000, hc: holdingCost/10000, gmroi };
    }

    const base = calcScenario(growthRate, svcZ, holdRate, leadAdj);
    const pes = calcScenario(growthRate - 0.1, Math.max(svcZ-0.3, 1.0), holdRate + 0.03, leadAdj + 3);
    const opt = calcScenario(growthRate + 0.3, Math.min(svcZ+0.3, 2.5), holdRate - 0.03, Math.max(0, leadAdj - 3));

    const setText = (id, val) => { const el=document.getElementById(id); if(el) el.textContent=val; };
    setText('wi-inv-pes', pes.inv.toFixed(1));
    setText('wi-inv-base', base.inv.toFixed(1));
    setText('wi-inv-opt', opt.inv.toFixed(1));
    setText('wi-dsi-pes', pes.dsi.toFixed(1));
    setText('wi-dsi-base', base.dsi.toFixed(1));
    setText('wi-dsi-opt', opt.dsi.toFixed(1));
    setText('wi-ss-pes', pes.ss.toFixed(1));
    setText('wi-ss-base', base.ss.toFixed(1));
    setText('wi-ss-opt', opt.ss.toFixed(1));
    setText('wi-hc-pes', pes.hc.toFixed(1));
    setText('wi-hc-base', base.hc.toFixed(1));
    setText('wi-hc-opt', opt.hc.toFixed(1));
    setText('wi-gm-pes', pes.gmroi.toFixed(2));
    setText('wi-gm-base', base.gmroi.toFixed(2));
    setText('wi-gm-opt', opt.gmroi.toFixed(2));

    const metrics = ['庫存金額(萬)', 'DSI(天)', '安全庫存(萬)', '持有成本(萬)', 'GMROI'];
    const pesVals = [pes.inv, pes.dsi, pes.ss, pes.hc, pes.gmroi];
    const baseVals = [base.inv, base.dsi, base.ss, base.hc, base.gmroi];
    const optVals = [opt.inv, opt.dsi, opt.ss, opt.hc, opt.gmroi];

    cc('chartWhatif', 'bar', {
        labels: metrics,
        datasets: [{
            label: '悲觀',
            data: pesVals.map(v => v.toFixed(2)),
            backgroundColor: 'rgba(244,63,94,0.7)',
            borderRadius: 3
        }, {
            label: '基準',
            data: baseVals.map(v => v.toFixed(2)),
            backgroundColor: 'rgba(99,102,241,0.7)',
            borderRadius: 3
        }, {
            label: '樂觀',
            data: optVals.map(v => v.toFixed(2)),
            backgroundColor: 'rgba(16,185,129,0.7)',
            borderRadius: 3
        }]
    }, {
        plugins: { legend:{ position:'top' } }
    });
}

// ══ 16. RENDER: EOQ ══
function renderEOQ() {
    if (!aggregatedData.products) return;
    const arr = aggregatedData.products;
    const orderCost = parseFloat(document.getElementById('eoq-order-cost')?.value || '500');
    const holdingRatePct = parseFloat(document.getElementById('eoq-holding-rate')?.value || '25') / 100;

    // Aggregate EOQ curve for "average" product
    const avgAnnualD = arr.reduce((s,p)=>s+p.annualDemand,0)/arr.length;
    const avgUnitCost = arr.reduce((s,p)=>s+p.unitCost,0)/arr.length;
    const H = avgUnitCost * holdingRatePct;
    const eoqOptimal = H > 0 ? Math.sqrt(2 * avgAnnualD * orderCost / H) : 100;
    const qRange = [];
    for (let q = Math.max(10, eoqOptimal*0.2); q <= eoqOptimal*3; q += eoqOptimal*0.15) qRange.push(Math.round(q));
    const holdCosts = qRange.map(q => (q/2 * H).toFixed(0));
    const orderCosts = qRange.map(q => (avgAnnualD/q * orderCost).toFixed(0));
    const totalCosts = qRange.map((q,i) => (parseFloat(holdCosts[i]) + parseFloat(orderCosts[i])).toFixed(0));

    cc('chartEOQ', 'line', {
        labels: qRange.map(q => Math.round(q)),
        datasets: [{
            label: '持有成本',
            data: holdCosts,
            borderColor: '#f59e0b',
            backgroundColor: 'rgba(245,158,11,0.1)',
            fill: false,
            tension: 0.4,
            pointRadius: 2
        }, {
            label: '訂購成本',
            data: orderCosts,
            borderColor: '#6366f1',
            backgroundColor: 'rgba(99,102,241,0.1)',
            fill: false,
            tension: 0.4,
            pointRadius: 2
        }, {
            label: '總成本',
            data: totalCosts,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16,185,129,0.1)',
            fill: false,
            tension: 0.4,
            pointRadius: 3,
            borderWidth: 2
        }]
    }, {
        plugins: {
            legend:{ position:'top' },
            tooltip:{ callbacks:{ afterLabel: ctx => ctx.datasetIndex===2 && Math.abs(qRange[ctx.dataIndex]-eoqOptimal)<eoqOptimal*0.15 ? '← EOQ最優點' : '' } }
        },
        scales: { y:{ ticks:{ callback: v => v.toLocaleString() + '元' } } }
    });

    // EOQ table
    const tbody = document.getElementById('eoq-table-body');
    if (tbody) {
        const top15 = [...arr].sort((a,b)=>b.annualDemand-a.annualDemand).slice(0,15);
        tbody.innerHTML = top15.map(p => {
            const h = p.unitCost * holdingRatePct;
            const eoq = h > 0 ? Math.round(Math.sqrt(2*p.annualDemand*orderCost/h)) : 0;
            const annualHC = (eoq/2 * h).toFixed(0);
            const annualOC = eoq > 0 ? (p.annualDemand/eoq * orderCost).toFixed(0) : 0;
            const devPct = p.eoqDeviation > 0.3 ? 'text-red-400' : 'text-emerald-400';
            return '<tr>' +
                '<td>' + p.name + '</td>' +
                '<td class="text-right">' + Math.round(p.annualDemand).toLocaleString() + '</td>' +
                '<td class="text-right">' + eoq.toLocaleString() + '</td>' +
                '<td class="text-right ' + devPct + '">' + (p.eoqDeviation*100).toFixed(1) + '%</td>' +
                '<td class="text-right">' + parseFloat(annualHC).toLocaleString() + '</td>' +
                '<td class="text-right">' + parseFloat(annualOC).toLocaleString() + '</td>' +
                '</tr>';
        }).join('');
    }
}

// ══ 17. RENDER: FIFO ══
function renderFIFO() {
    if (!aggregatedData.products) return;
    const sel = document.getElementById('fifo-product-select');
    if (sel && sel.options.length <= 1) {
        aggregatedData.products.forEach(p => {
            const opt = document.createElement('option');
            opt.value = p.id; opt.textContent = p.id + ' ' + p.name;
            sel.appendChild(opt);
        });
    }
    const selectedId = sel?.value;
    if (!selectedId) { renderFIFOProduct(aggregatedData.products[0]?.id); return; }
    renderFIFOProduct(selectedId);
}

function renderFIFOProduct(productId) {
    if (!productId || !aggregatedData.products) return;
    const fifoResult = computeFIFO(productId);
    if (!fifoResult) return;

    const { batches, stats } = fifoResult;
    const setText = (id, val) => { const el=document.getElementById(id); if(el) el.textContent=val; };
    setText('fifo-batch-count', batches.filter(b=>b.remaining>0).length);
    setText('fifo-avg-aging', stats.avgAging.toFixed(1) + ' 天');
    setText('fifo-total-value', stats.totalValue.toLocaleString() + ' 元');

    // FIFO waterfall chart (horizontal stacked)
    const batchLabels = batches.map((b,i) => '批次' + (i+1) + ' (' + b.date.slice(5) + ')');
    cc('chartFIFO', 'bar', {
        labels: batchLabels,
        datasets: [{
            label: '已耗用',
            data: batches.map(b => b.consumed),
            backgroundColor: 'rgba(148,163,184,0.4)',
            borderRadius: 2
        }, {
            label: '剩餘庫存',
            data: batches.map(b => b.remaining),
            backgroundColor: batches.map(b => b.aging > 90 ? 'rgba(239,68,68,0.7)' : b.aging > 60 ? 'rgba(245,158,11,0.7)' : 'rgba(16,185,129,0.7)'),
            borderRadius: 2
        }]
    }, {
        indexAxis: 'y',
        scales: { x:{ stacked:true }, y:{ stacked:true } },
        plugins:{ legend:{ position:'top' } }
    });

    // Batch detail table
    const tbody = document.getElementById('fifo-table-body');
    if (tbody) {
        const today = new Date();
        tbody.innerHTML = batches.map((b, i) => {
            const statusClass = b.remaining === 0 ? 'text-slate-500' : b.aging > 90 ? 'text-red-400' : b.aging > 60 ? 'text-amber-400' : 'text-emerald-400';
            const status = b.remaining === 0 ? '已耗盡' : b.aging > 90 ? '⚠ 高齡' : b.aging > 60 ? '注意' : '✓ 正常';
            return '<tr>' +
                '<td>批次' + (i+1) + '</td>' +
                '<td>' + b.date + '</td>' +
                '<td class="text-right">' + b.qty + '</td>' +
                '<td class="text-right">' + b.consumed + '</td>' +
                '<td class="text-right ' + (b.remaining>0?'text-blue-400':'text-slate-500') + '">' + b.remaining + '</td>' +
                '<td class="text-right">' + b.cost.toFixed(2) + '</td>' +
                '<td class="text-right">' + (b.remaining * b.cost).toFixed(0) + '</td>' +
                '<td class="text-right ' + (b.aging > 90 ? 'text-red-400' : b.aging > 60 ? 'text-amber-400' : '') + '">' + b.aging + '</td>' +
                '<td class="' + statusClass + '">' + status + '</td>' +
                '</tr>';
        }).join('');
    }
}

// ══ 18. COMPUTE FIFO ══
function computeFIFO(productId) {
    const product = aggregatedData.products?.find(p => p.id === productId);
    if (!product) return null;

    const productRaw = rawData.filter(r => r.ProductID === productId).sort((a,b) => a.Date.localeCompare(b.Date));
    const batches = [];
    let batchNum = 0;

    productRaw.forEach(r => {
        if ((r.in_qty || 0) > 0) {
            batches.push({ date: r.Date, qty: r.in_qty, remaining: r.in_qty, consumed: 0, cost: r.UnitCost || product.unitCost });
            batchNum++;
        }
        let toConsume = r.out_qty || 0;
        for (let i = 0; i < batches.length && toConsume > 0; i++) {
            const take = Math.min(batches[i].remaining, toConsume);
            batches[i].remaining -= take;
            batches[i].consumed += take;
            toConsume -= take;
        }
    });

    const today = new Date();
    batches.forEach(b => {
        const d = new Date(b.date);
        b.aging = Math.floor((today - d) / 86400000);
    });

    const activeBatches = batches.filter(b => b.qty > 0);
    const totalRemaining = activeBatches.reduce((s,b)=>s+b.remaining, 0);
    const totalValue = activeBatches.reduce((s,b)=>s+b.remaining*b.cost, 0);
    const avgAging = totalRemaining > 0
        ? activeBatches.reduce((s,b)=>s+b.remaining*b.aging,0)/totalRemaining
        : 0;

    return { batches: activeBatches, stats: { totalRemaining, totalValue, avgAging } };
}

// ══ 19. RENDER: QUALITY ══
function renderQuality() {
    const issues = aggregatedData.qualityIssues || [];
    const errors = issues.filter(q => q.severity === 'error').length;
    const warnings = issues.filter(q => q.severity === 'warning').length;
    const notices = issues.filter(q => q.severity === 'info').length;
    const totalChecks = errors * 10 + warnings * 3 + notices;
    const qualityScore = Math.max(0, Math.min(100, 100 - Math.min(totalChecks, 100)));

    const setText = (id, val) => { const el=document.getElementById(id); if(el) el.textContent=val; };
    setText('quality-score', qualityScore.toFixed(0));
    setText('quality-errors', errors);
    setText('quality-warnings', warnings);
    setText('quality-notices', notices);

    // Gauge arc
    const arc = document.getElementById('gauge-arc');
    if (arc) {
        const totalLen = 188.5;
        const filled = totalLen * (qualityScore / 100);
        arc.setAttribute('stroke-dashoffset', (totalLen - filled).toFixed(1));
        arc.setAttribute('stroke', qualityScore >= 80 ? '#10b981' : qualityScore >= 60 ? '#f59e0b' : '#ef4444');
    }

    // Checklist
    const checkMark = (ok) => ok ? '<span class="text-emerald-400">✓</span>' : '<span class="text-red-400">✗</span>';
    const hasDateIssue = issues.some(q => q.message.includes('日期'));
    const hasQtyIssue = issues.some(q => q.message.includes('負值') || q.message.includes('超過'));
    const hasCostIssue = issues.some(q => q.message.includes('成本') || q.message.includes('單價'));
    const hasSKUIssue = issues.some(q => q.message.includes('ProductID'));
    const hasBalIssue = issues.some(q => q.message.includes('平衡') || q.message.includes('庫存'));
    const hasCatIssue = issues.some(q => q.message.includes('品類'));

    [['qc-dates', !hasDateIssue], ['qc-qty', !hasQtyIssue], ['qc-cost', !hasCostIssue],
     ['qc-sku', !hasSKUIssue], ['qc-balance', !hasBalIssue], ['qc-cat', !hasCatIssue]
    ].forEach(([id, ok]) => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = checkMark(ok);
    });

    // Issues table
    const tbody = document.getElementById('quality-table-body');
    if (tbody) {
        if (issues.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center text-emerald-400 py-6">✓ 資料品質良好，無異常</td></tr>';
        } else {
            const sevClass = { error:'text-red-400', warning:'text-amber-400', info:'text-blue-400' };
            const sevLabel = { error:'🔴 嚴重', warning:'🟡 警告', info:'🔵 注意' };
            tbody.innerHTML = issues.slice(0, 50).map(q =>
                '<tr>' +
                '<td class="' + (sevClass[q.severity]||'text-slate-400') + '">' + (sevLabel[q.severity]||q.severity) + '</td>' +
                '<td>' + (q.product||'—') + '</td>' +
                '<td>' + (q.date||'—') + '</td>' +
                '<td>' + q.message + '</td>' +
                '</tr>'
            ).join('');
        }
    }
}

// ══ 20. CHECK DATA QUALITY ══
function checkDataQuality(raw) {
    const issues = [];
    const pidSeen = {};
    const datesByPid = {};

    raw.forEach(r => {
        const pid = r.ProductID;
        if (!pid) { issues.push({ severity:'error', product:'—', date:r.Date||'—', message:'ProductID 欄位缺失' }); return; }
        if (!r.Date) { issues.push({ severity:'error', product:pid, date:'—', message:'日期欄位缺失' }); return; }

        if (!datesByPid[pid]) datesByPid[pid] = new Set();
        const dateKey = r.Date;
        if (datesByPid[pid].has(dateKey)) {
            issues.push({ severity:'warning', product:pid, date:r.Date, message:'重複日期記錄' });
        }
        datesByPid[pid].add(dateKey);

        if ((r.in_qty || 0) < 0) issues.push({ severity:'error', product:pid, date:r.Date, message:'入庫數量為負值 (' + r.in_qty + ')' });
        if ((r.out_qty || 0) < 0) issues.push({ severity:'error', product:pid, date:r.Date, message:'出庫數量為負值 (' + r.out_qty + ')' });
        if ((r.current_inventory || 0) < 0) issues.push({ severity:'warning', product:pid, date:r.Date, message:'當日庫存為負值 (' + r.current_inventory + ')' });
        if ((r.UnitCost || 0) <= 0) issues.push({ severity:'warning', product:pid, date:r.Date, message:'單位成本異常 (' + r.UnitCost + ')' });
        if ((r.UnitPrice || 0) < (r.UnitCost || 0)) issues.push({ severity:'info', product:pid, date:r.Date, message:'售價低於成本' });
        if ((r.out_qty || 0) > (r.current_inventory || 0) + (r.in_qty || 0)) {
            issues.push({ severity:'warning', product:pid, date:r.Date, message:'出庫量超過可用庫存' });
        }
        // Extreme outlier: out_qty > 10x average
        pidSeen[pid] = pidSeen[pid] || [];
        pidSeen[pid].push(r.out_qty || 0);
    });

    // Outlier check
    Object.entries(pidSeen).forEach(([pid, outs]) => {
        const avg = outs.reduce((s,v)=>s+v,0)/outs.length;
        const maxOut = Math.max(...outs);
        if (maxOut > avg * 10 && avg > 0) {
            issues.push({ severity:'info', product:pid, date:'—', message:'出庫量存在極端離群值 (最大=' + maxOut + ', 均值=' + avg.toFixed(1) + ')' });
        }
    });

    return issues;
}
"""

content.write(js)
content.close()
print("Part 3 written:", len(js), "chars")
