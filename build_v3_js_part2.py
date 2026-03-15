content = open('C:/Users/jamic/庫存分析/v3_parts/part_js_p2.txt', 'w', encoding='utf-8')

js = """
// ══ 7. RENDER: OVERVIEW ══
function renderOverview() {
    if (!aggregatedData.products) return;
    const arr = aggregatedData.products;
    const ms = aggregatedData.monthly || {};
    const mKeys = aggregatedData.mKeys || [];

    // Trend chart - monthly inventory value
    const trendLabels = mKeys.map(k => k.slice(5));
    const trendVals = mKeys.map(k => (ms[k].avgInv || 0) / 10000);
    cc('chartTrendOverview', 'line', {
        labels: trendLabels,
        datasets: [{
            label: '庫存金額(萬元)',
            data: trendVals,
            borderColor: '#6366f1',
            backgroundColor: 'rgba(99,102,241,0.12)',
            fill: true,
            tension: 0.4,
            pointRadius: 4,
            pointBackgroundColor: '#6366f1'
        }, {
            label: 'COGS(萬元)',
            data: mKeys.map(k => (ms[k].cogs || 0) / 10000),
            borderColor: '#22d3ee',
            backgroundColor: 'rgba(34,211,238,0.08)',
            fill: false,
            tension: 0.4,
            pointRadius: 3,
            pointBackgroundColor: '#22d3ee'
        }]
    }, {
        plugins: { legend: { position: 'top' } },
        scales: {
            y: { ticks: { callback: v => v.toFixed(0) + '萬' } }
        }
    });

    // Category donut
    const catTotals = {};
    arr.forEach(p => { catTotals[p.category] = (catTotals[p.category] || 0) + p.totalCurrentValue; });
    const cats = Object.keys(catTotals);
    cc('chartCategoryOverview', 'doughnut', {
        labels: cats,
        datasets: [{
            data: cats.map(c => catTotals[c]),
            backgroundColor: cats.map(c => CAT_COLORS[c] || '#6366f1'),
            borderWidth: 2,
            borderColor: 'rgba(0,0,0,0.2)'
        }]
    }, {
        plugins: {
            legend: { position: 'bottom' },
            tooltip: { callbacks: { label: ctx => ctx.label + ': ' + (ctx.raw/10000).toFixed(1) + '萬元' } }
        },
        cutout: '60%'
    });
}

// ══ 8. RENDER: TREND ══
function renderTrend() {
    if (!aggregatedData.monthly) return;
    const ms = aggregatedData.monthly;
    const mKeys = aggregatedData.mKeys || [];
    const labels = mKeys.map(k => k.slice(5));

    // DSI / Turnover dual axis
    cc('chartTurnover', 'bar', {
        labels,
        datasets: [{
            label: 'DSI(天)',
            data: mKeys.map(k => (ms[k].dsi || 0).toFixed(1)),
            backgroundColor: 'rgba(99,102,241,0.6)',
            yAxisID: 'y'
        }, {
            label: '周轉率(次/月)',
            data: mKeys.map(k => (ms[k].turnover || 0).toFixed(2)),
            type: 'line',
            borderColor: '#22d3ee',
            backgroundColor: 'transparent',
            pointRadius: 4,
            yAxisID: 'y1'
        }]
    }, {
        scales: {
            y: { position:'left', ticks:{ callback: v => v+'天' } },
            y1: { position:'right', grid:{ drawOnChartArea:false }, ticks:{ callback: v => v+'次' } }
        },
        plugins: { legend:{ position:'top' } }
    });

    // SSR chart
    renderSSRChart();

    // Growth chart
    const growthData = mKeys.map(k => ((ms[k].growth || 0) * 100).toFixed(1));
    cc('chartGrowth', 'bar', {
        labels,
        datasets: [{
            label: 'MoM成長率(%)',
            data: growthData,
            backgroundColor: growthData.map(v => parseFloat(v) >= 0 ? 'rgba(16,185,129,0.7)' : 'rgba(244,63,94,0.7)'),
            borderRadius: 3
        }]
    }, {
        scales: { y: { ticks:{ callback: v => v+'%' } } },
        plugins: { legend:{ display:false } }
    });

    // Moving averages
    const actualData = mKeys.map(k => ms[k].totalOut || 0);
    const ma3Data = mKeys.map((k,i) => ms[k].ma3 || null);
    const ma6Data = mKeys.map((k,i) => ms[k].ma6 || null);
    cc('chartMA', 'line', {
        labels,
        datasets: [{
            label: '實際出庫量',
            data: actualData,
            borderColor: '#94a3b8',
            backgroundColor: 'rgba(148,163,184,0.1)',
            fill: false,
            pointRadius: 2
        }, {
            label: '3M MA',
            data: ma3Data,
            borderColor: '#6366f1',
            backgroundColor: 'transparent',
            fill: false,
            pointRadius: 0,
            borderWidth: 2,
            tension: 0.4,
            spanGaps: true
        }, {
            label: '6M MA',
            data: ma6Data,
            borderColor: '#f59e0b',
            backgroundColor: 'transparent',
            fill: false,
            pointRadius: 0,
            borderWidth: 2,
            tension: 0.4,
            spanGaps: true
        }]
    }, {
        plugins: { legend:{ position:'top' } }
    });
}

function renderSSRChart() {
    const ms = aggregatedData.monthly;
    const mKeys = aggregatedData.mKeys || [];
    const labels = mKeys.map(k => k.slice(5));
    const isLong = ssrMode === 'long';
    const ssrData = mKeys.map(k => (isLong ? (ms[k].ssrLong||0) : (ms[k].ssrShort||0)).toFixed(3));
    cc('chartSSR', 'line', {
        labels,
        datasets: [{
            label: isLong ? '存銷比(長週期)' : '存銷比(短週期)',
            data: ssrData,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16,185,129,0.1)',
            fill: true,
            tension: 0.4,
            pointRadius: 4,
            pointBackgroundColor: '#10b981'
        }]
    }, {
        plugins: { legend:{ display:false } },
        scales: { y: { ticks:{ callback: v => v.toFixed(2) } } }
    });
}

// ══ 9. RENDER: STRUCTURE ══
function renderStructure() {
    if (!aggregatedData.products) return;
    const arr = aggregatedData.products;

    // ABC pie
    const abcCounts = {A:0, B:0, C:0};
    const abcVals = {A:0, B:0, C:0};
    arr.forEach(p => { abcCounts[p.abc]++; abcVals[p.abc] += p.totalCurrentValue; });
    cc('chartABC', 'pie', {
        labels: ['A類', 'B類', 'C類'],
        datasets: [{
            data: [abcVals.A, abcVals.B, abcVals.C],
            backgroundColor: ['#6366f1','#f59e0b','#94a3b8'],
            borderWidth: 2,
            borderColor: 'rgba(0,0,0,0.2)'
        }]
    }, {
        plugins: {
            legend:{ position:'bottom' },
            tooltip:{ callbacks:{ label: ctx => ctx.label + ': ' + (ctx.raw/10000).toFixed(1) + '萬 (' + abcCounts[['A','B','C'][ctx.dataIndex]] + '品)' } }
        }
    });

    // Imbalance scatter
    const scatterData = arr.map(p => ({ x: (p.cogsPct*100).toFixed(2), y: (p.valPct*100).toFixed(2), r: 5 }));
    cc('chartImbalance', 'scatter', {
        datasets: [{
            label: '品項',
            data: scatterData,
            backgroundColor: arr.map(p => p.imbalance > 0.05 ? 'rgba(244,63,94,0.7)' : p.imbalance < -0.05 ? 'rgba(34,211,238,0.7)' : 'rgba(99,102,241,0.7)'),
            pointRadius: 5
        }, {
            label: '均衡線',
            data: [{x:0,y:0},{x:10,y:10}],
            type: 'line',
            borderColor: 'rgba(255,255,255,0.3)',
            borderDash: [4,4],
            pointRadius: 0,
            backgroundColor: 'transparent'
        }]
    }, {
        scales: {
            x: { title:{ display:true, text:'COGS佔比(%)', color:'#94a3b8' } },
            y: { title:{ display:true, text:'庫存值佔比(%)', color:'#94a3b8' } }
        },
        plugins:{ legend:{ position:'top' } }
    });

    // Pareto
    const sorted = [...arr].sort((a,b) => b.totalCurrentValue - a.totalCurrentValue);
    const totalVal = aggregatedData.totalVal || 1;
    let cum = 0;
    const paretoLabels = sorted.slice(0,20).map(p => p.id);
    const paretoVals = sorted.slice(0,20).map(p => (p.totalCurrentValue/10000).toFixed(1));
    const paretoCum = sorted.slice(0,20).map(p => { cum += p.totalCurrentValue; return (cum/totalVal*100).toFixed(1); });
    cc('chartPareto', 'bar', {
        labels: paretoLabels,
        datasets: [{
            label: '庫存值(萬)',
            data: paretoVals,
            backgroundColor: 'rgba(99,102,241,0.7)',
            yAxisID: 'y'
        }, {
            label: '累積%',
            data: paretoCum,
            type: 'line',
            borderColor: '#f59e0b',
            backgroundColor: 'transparent',
            pointRadius: 3,
            yAxisID: 'y1'
        }]
    }, {
        scales: {
            y: { position:'left' },
            y1: { position:'right', min:0, max:100, ticks:{ callback: v=>v+'%' }, grid:{ drawOnChartArea:false } }
        },
        plugins:{ legend:{ position:'top' } }
    });

    // ABC-XYZ matrix
    const cells = {};
    ['AX','AY','AZ','BX','BY','BZ','CX','CY','CZ'].forEach(k => cells[k] = {count:0, val:0});
    arr.forEach(p => {
        const k = p.abcxyz;
        if (cells[k]) { cells[k].count++; cells[k].val += p.totalCurrentValue; }
    });
    ['AX','AY','AZ','BX','BY','BZ','CX','CY','CZ'].forEach(k => {
        const el = document.getElementById('cnt-' + k.toLowerCase());
        if (el) el.textContent = cells[k].count + '品 / ' + (cells[k].val/10000).toFixed(0) + '萬';
    });
}

// ══ 10. RENDER: AGING ══
function renderAging() {
    if (!aggregatedData.products) return;
    const arr = aggregatedData.products;
    const cats = [...new Set(arr.map(p => p.category))];
    const buckets = ['0-30天','30-60天','60-90天','90-120天','>120天'];
    const bucketColors = ['rgba(16,185,129,0.7)','rgba(234,179,8,0.7)','rgba(245,158,11,0.7)','rgba(249,115,22,0.7)','rgba(239,68,68,0.7)'];
    const bucketData = cats.map(cat => {
        const catItems = arr.filter(p => p.category === cat);
        return buckets.map((_, bi) => catItems.filter(p => p.agingBucket === bi).reduce((s,p) => s+p.totalCurrentValue, 0) / 10000);
    });

    cc('chartAging', 'bar', {
        labels: cats,
        datasets: buckets.map((label, bi) => ({
            label,
            data: cats.map((cat, ci) => bucketData[ci][bi].toFixed(1)),
            backgroundColor: bucketColors[bi],
            borderRadius: 2
        }))
    }, {
        scales: { x: { stacked:true }, y: { stacked:true, ticks:{ callback: v=>v+'萬' } } },
        plugins: { legend:{ position:'top' } }
    });

    // Loss estimate
    const lossRates = readLossRates();
    const lossByBucket = [0,1,2,3,4].map(bi =>
        arr.filter(p=>p.agingBucket===bi).reduce((s,p)=>s+p.totalCurrentValue*lossRates[bi],0)
    );
    const lossIds = ['loss-30','loss-60','loss-90','loss-180'];
    const fmt = v => (v/10000).toFixed(1) + '萬元';
    [1,2,3,4].forEach((bi,i) => {
        const el = document.getElementById(lossIds[i]);
        if (el) el.textContent = fmt(lossByBucket[bi]);
    });
    const totalLoss = lossByBucket.reduce((s,v)=>s+v,0);
    const elTot = document.getElementById('loss-total');
    if (elTot) elTot.textContent = fmt(totalLoss);

    // Divergence signal
    const mKeys = aggregatedData.mKeys || [];
    const ms = aggregatedData.monthly;
    const divPanel = document.getElementById('divergence-panel');
    if (divPanel && mKeys.length >= 2) {
        const lastMo = mKeys[mKeys.length-1];
        const prevMo = mKeys[mKeys.length-2];
        const dsiNow = ms[lastMo].dsi || 0;
        const dsiPrev = ms[prevMo].dsi || 0;
        const outNow = ms[lastMo].totalOut || 0;
        const outPrev = ms[prevMo].totalOut || 0;
        let signal = '— 正常範圍';
        let signalColor = 'text-slate-400';
        if (dsiNow > dsiPrev*1.2 && outNow < outPrev*0.9) {
            signal = '⚠ 高庫存低銷量 背離';
            signalColor = 'text-red-400';
        } else if (dsiNow < dsiPrev*0.8 && outNow > outPrev*1.1) {
            signal = '✓ 低庫存高銷量 健康';
            signalColor = 'text-emerald-400';
        }
        divPanel.innerHTML = '<div class="' + signalColor + ' font-medium">' + signal + '</div>' +
            '<div class="mt-2 space-y-1">' +
            '<div>本月DSI: <span class="text-white">' + dsiNow.toFixed(1) + '天</span></div>' +
            '<div>上月DSI: <span class="text-white">' + dsiPrev.toFixed(1) + '天</span></div>' +
            '<div>本月出庫: <span class="text-white">' + outNow.toFixed(0) + '</span></div>' +
            '<div>上月出庫: <span class="text-white">' + outPrev.toFixed(0) + '</span></div>' +
            '</div>';
    }

    // Health radar
    const health = aggregatedData._health || 0;
    const dsi = aggregatedData._dsi || 0;
    const turnover = aggregatedData._turnover || 0;
    const gmroi = aggregatedData._gmroi || 0;
    const str = aggregatedData._str || 0;
    const aging90 = aggregatedData._aging90Pct || 0;
    cc('chartHealth', 'radar', {
        labels: ['庫存健康','周轉效率','庫齡控制','GMROI','售磬率'],
        datasets: [{
            label: '當前狀態',
            data: [
                health,
                Math.min(100, turnover/10*100),
                Math.max(0, 100 - aging90*2),
                Math.min(100, gmroi/3*100),
                Math.min(100, str*100)
            ],
            backgroundColor: 'rgba(99,102,241,0.2)',
            borderColor: '#6366f1',
            pointBackgroundColor: '#6366f1',
            pointRadius: 4
        }]
    }, {
        scales: { r: { min:0, max:100, ticks:{ stepSize:25 } } },
        plugins: { legend:{ display:false } }
    });
}

// ══ 11. RENDER: FORECAST ══
function renderForecast() {
    if (!aggregatedData.products) return;
    const ms = aggregatedData.monthly;
    const mKeys = aggregatedData.mKeys || [];
    const seasonalIndex = aggregatedData.seasonalIndex || {};
    const fisherRatio = aggregatedData.fisherRatio || 0;

    // Fisher panel
    const fisherPanel = document.getElementById('fisher-panel');
    if (fisherPanel) {
        const threshold = 6.5;
        const hasSeason = fisherRatio > threshold;
        fisherPanel.innerHTML =
            '<div class="flex justify-between mb-2"><span class="text-slate-400">Fisher比值</span><span class="' + (hasSeason?'text-emerald-400':'text-slate-300') + ' font-bold">' + fisherRatio.toFixed(2) + '</span></div>' +
            '<div class="flex justify-between mb-2"><span class="text-slate-400">臨界值(n=12)</span><span class="text-slate-300">' + threshold + '</span></div>' +
            '<div class="flex justify-between mb-3"><span class="text-slate-400">判定結果</span><span class="' + (hasSeason?'text-emerald-400':'text-amber-400') + ' font-semibold">' + (hasSeason?'✓ 具季節性':'— 無顯著季節性') + '</span></div>' +
            '<div class="text-[0.65rem] text-slate-500">' + (hasSeason?'建議使用季節調整預測模型':'建議使用趨勢/移動均值模型') + '</div>' +
            '<div class="mt-2 text-[0.65rem] text-slate-500">H₀: 無週期性 | 拒絕域: F > ' + threshold + '</div>';
    }

    // Seasonal radar
    const monthLabels = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
    const siData = monthLabels.map((_, i) => (seasonalIndex[i+1] || 1).toFixed(3));
    cc('chartSeasonalRadar', 'radar', {
        labels: monthLabels,
        datasets: [{
            label: '季節指數',
            data: siData,
            backgroundColor: 'rgba(34,211,238,0.15)',
            borderColor: '#22d3ee',
            pointBackgroundColor: '#22d3ee',
            pointRadius: 4
        }, {
            label: '基準線(1.0)',
            data: Array(12).fill(1),
            backgroundColor: 'transparent',
            borderColor: 'rgba(255,255,255,0.2)',
            borderDash: [4,4],
            pointRadius: 0
        }]
    }, {
        plugins: { legend:{ position:'top' } },
        scales: { r: { min:0.5, max:1.5, ticks:{ stepSize:0.25 } } }
    });

    // 3-method forecast (extend 3 months)
    const histLabels = mKeys.map(k => k.slice(5));
    const histOut = mKeys.map(k => ms[k].totalOut || 0);
    const lastMo = mKeys[mKeys.length-1] || '2026-02';
    const futureLabels = [];
    for (let i = 1; i <= 3; i++) {
        const d = new Date(lastMo + '-01');
        d.setMonth(d.getMonth() + i);
        futureLabels.push(d.toISOString().slice(0,7).slice(5));
    }
    const allLabels = [...histLabels, ...futureLabels];

    // MA3 forecast
    const ma3base = histOut.slice(-3).reduce((s,v)=>s+v,0)/3;
    // Trend forecast
    const n = histOut.length;
    const xBar = (n-1)/2;
    const yBar = histOut.reduce((s,v)=>s+v,0)/n;
    const slope = histOut.reduce((s,v,i)=>s+(i-xBar)*(v-yBar),0) / histOut.reduce((s,_,i)=>s+(i-xBar)**2,0);
    const intercept = yBar - slope*xBar;
    // Composite forecast with seasonal
    const nextMonths = [1,2,3].map(i => {
        const d = new Date(lastMo + '-01');
        d.setMonth(d.getMonth() + i);
        return d.getMonth() + 1;
    });

    const histNulls = Array(histLabels.length).fill(null);
    const ma3Forecast = [...histNulls, ...nextMonths.map(() => Math.round(ma3base))];
    const trendForecast = [...histNulls, ...nextMonths.map((_, i) => Math.max(0, Math.round(intercept + slope*(n+i))))];
    const compositeForecast = [...histNulls, ...nextMonths.map((mn, i) => {
        const si = seasonalIndex[mn] || 1;
        return Math.max(0, Math.round(((ma3base*0.6 + (intercept+slope*(n+i))*0.4)) * si));
    })];

    // Upper/lower bands for composite
    const bandUpper = [...histNulls, ...compositeForecast.slice(n).map(v => Math.round(v * 1.15))];
    const bandLower = [...histNulls, ...compositeForecast.slice(n).map(v => Math.round(v * 0.85))];

    cc('chartForecast', 'line', {
        labels: allLabels,
        datasets: [{
            label: '歷史出庫',
            data: [...histOut, ...Array(3).fill(null)],
            borderColor: '#94a3b8',
            backgroundColor: 'rgba(148,163,184,0.08)',
            fill: false,
            pointRadius: 2
        }, {
            label: 'MA3預測',
            data: ma3Forecast,
            borderColor: '#6366f1',
            borderDash: [6,3],
            backgroundColor: 'transparent',
            pointRadius: 4,
            spanGaps: false
        }, {
            label: '趨勢預測',
            data: trendForecast,
            borderColor: '#f59e0b',
            borderDash: [6,3],
            backgroundColor: 'transparent',
            pointRadius: 4,
            spanGaps: false
        }, {
            label: '複合預測',
            data: compositeForecast,
            borderColor: '#10b981',
            borderDash: [6,3],
            backgroundColor: 'rgba(16,185,129,0.1)',
            fill: false,
            pointRadius: 5,
            spanGaps: false
        }, {
            label: '上限(+15%)',
            data: bandUpper,
            borderColor: 'rgba(16,185,129,0.3)',
            backgroundColor: 'rgba(16,185,129,0.05)',
            fill: '+1',
            pointRadius: 0,
            borderDash: [2,4],
            spanGaps: false
        }, {
            label: '下限(-15%)',
            data: bandLower,
            borderColor: 'rgba(16,185,129,0.3)',
            backgroundColor: 'rgba(16,185,129,0.05)',
            fill: false,
            pointRadius: 0,
            borderDash: [2,4],
            spanGaps: false
        }]
    }, {
        plugins: { legend:{ position:'top' } }
    });

    // ROP table
    const ropBody = document.getElementById('rop-table-body');
    if (ropBody) {
        const top20 = [...aggregatedData.products].sort((a,b) => b.totalCurrentValue - a.totalCurrentValue).slice(0,20);
        ropBody.innerHTML = top20.map(p => {
            const status = p.totalCurrent < p.rop ? '⚠ 低於ROP' : '✓ 充足';
            const statusClass = p.totalCurrent < p.rop ? 'text-red-400' : 'text-emerald-400';
            return '<tr>' +
                '<td>' + p.name + '</td>' +
                '<td class="text-right">' + p.dailySalesAvg.toFixed(2) + '</td>' +
                '<td class="text-right">' + p.leadTime + '</td>' +
                '<td class="text-right">' + Math.round(p.safetyStock) + '</td>' +
                '<td class="text-right">' + Math.round(p.rop) + '</td>' +
                '<td class="text-right">' + p.totalCurrent + '</td>' +
                '<td class="' + statusClass + '">' + status + '</td>' +
                '</tr>';
        }).join('');
    }
}

// ══ 12. RENDER: SUPPLY ══
function renderSupply() {
    if (!aggregatedData.products) return;
    const arr = aggregatedData.products;
    const ms = aggregatedData.monthly;
    const mKeys = aggregatedData.mKeys || [];

    // GMROI by category
    const catGMROI = {};
    const catInv = {};
    arr.forEach(p => {
        catGMROI[p.category] = (catGMROI[p.category] || 0) + p.gmroi * p.totalCurrentValue;
        catInv[p.category] = (catInv[p.category] || 0) + p.totalCurrentValue;
    });
    const cats = Object.keys(catGMROI);
    const gmroiVals = cats.map(c => catInv[c] > 0 ? (catGMROI[c]/catInv[c]).toFixed(2) : 0);
    cc('chartGMROI', 'bar', {
        labels: cats,
        datasets: [{
            label: 'GMROI',
            data: gmroiVals,
            backgroundColor: cats.map(c => CAT_COLORS[c] || '#6366f1'),
            borderRadius: 4
        }]
    }, {
        indexAxis: 'y',
        plugins: { legend:{ display:false } },
        scales: {
            x: { ticks:{ callback: v => v + '倍' } },
            y: {}
        }
    });

    // Heatmap using Canvas 2D
    const heatmapContainer = document.getElementById('chartHeatmap');
    if (heatmapContainer) {
        heatmapContainer.innerHTML = '';
        const canvas = document.createElement('canvas');
        const cellW = 50, cellH = 28, padL = 70, padT = 40;
        const months = mKeys.slice(-6);
        canvas.width = padL + months.length * cellW + 10;
        canvas.height = padT + cats.length * cellH + 10;
        heatmapContainer.appendChild(canvas);
        const ctx = canvas.getContext('2d');
        const isDark = !document.body.classList.contains('light-mode');
        ctx.fillStyle = isDark ? '#1e293b' : '#f1f5f9';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.font = '10px sans-serif';
        ctx.fillStyle = isDark ? '#94a3b8' : '#64748b';

        // Category labels (Y axis)
        cats.forEach((cat, ci) => {
            ctx.fillStyle = isDark ? '#94a3b8' : '#64748b';
            ctx.textAlign = 'right';
            ctx.fillText(cat.slice(0,4), padL - 4, padT + ci * cellH + cellH/2 + 3);
        });

        // Month labels (X axis)
        months.forEach((mo, mi) => {
            ctx.fillStyle = isDark ? '#94a3b8' : '#64748b';
            ctx.textAlign = 'center';
            ctx.fillText(mo.slice(5), padL + mi * cellW + cellW/2, padT - 10);
        });

        // Build out_qty by cat+month
        const catMoData = {};
        cats.forEach(cat => {
            catMoData[cat] = {};
            months.forEach(mo => {
                catMoData[cat][mo] = arr.filter(p => p.category === cat)
                    .reduce((s,p) => s + (p.monthlyOut[mo] || 0), 0);
            });
        });
        const maxVal = Math.max(...cats.flatMap(cat => months.map(mo => catMoData[cat][mo])));

        cats.forEach((cat, ci) => {
            months.forEach((mo, mi) => {
                const val = catMoData[cat][mo];
                const ratio = maxVal > 0 ? val/maxVal : 0;
                const r = Math.round(99 + (239-99)*ratio);
                const g = Math.round(102 + (68-102)*ratio);
                const b = Math.round(241 + (94-241)*ratio);
                ctx.fillStyle = 'rgba(' + r + ',' + g + ',' + b + ',0.85)';
                ctx.fillRect(padL + mi*cellW + 1, padT + ci*cellH + 1, cellW-2, cellH-2);
                ctx.fillStyle = '#fff';
                ctx.textAlign = 'center';
                ctx.fillText(val.toFixed(0), padL + mi*cellW + cellW/2, padT + ci*cellH + cellH/2 + 3);
            });
        });
    }

    // Replenishment table
    const repBody = document.getElementById('replenish-table-body');
    if (repBody) {
        const toReplenish = arr.filter(p => p.totalCurrent <= p.rop * 1.1)
            .sort((a,b) => (a.totalCurrent/a.rop) - (b.totalCurrent/b.rop))
            .slice(0, 15);
        if (toReplenish.length === 0) {
            repBody.innerHTML = '<tr><td colspan="6" class="text-center text-emerald-400 py-4">✓ 目前無緊急補貨需求</td></tr>';
        } else {
            repBody.innerHTML = toReplenish.map(p => {
                const repQty = Math.max(0, Math.round(p.eoq || (p.dailySalesAvg * p.leadTime * 1.5)));
                const urgency = p.totalCurrent < p.safetyStock ? '🔴 緊急' : p.totalCurrent < p.rop ? '🟡 建議' : '🟢 觀察';
                const daysToReplenish = p.dailySalesAvg > 0 ? Math.floor((p.totalCurrent - p.safetyStock) / p.dailySalesAvg) : 99;
                const repDate = new Date();
                repDate.setDate(repDate.getDate() + Math.max(0, daysToReplenish - p.leadTime));
                return '<tr>' +
                    '<td>' + p.category + '</td>' +
                    '<td>' + p.name + '</td>' +
                    '<td class="text-right">' + repQty + ' 件</td>' +
                    '<td class="text-right">' + repDate.toISOString().slice(0,10) + '</td>' +
                    '<td class="text-right">' + (repQty * p.unitCost).toLocaleString() + ' 元</td>' +
                    '<td>' + urgency + '</td>' +
                    '</tr>';
            }).join('');
        }
    }
}
"""

content.write(js)
content.close()
print("Part 2 written:", len(js), "chars")
