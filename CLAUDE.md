# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## 專案概述

此目錄是「庫存分析智能儀表板」專案，包含 V1 / V2 / V3 三個版本的零依賴單一 HTML 儀表板。**主力開發版本為 V3**（`Inventory_Strategic_OS_V3.html`）。

---

## 常用指令

### 開啟儀表板

```bash
# 直接用系統預設瀏覽器開啟（Windows）
start Inventory_Strategic_OS_V3.html

# 若需要 HTTP 伺服器（避免 file:// CORS 問題）
python -m http.server 8080
# 然後瀏覽 http://localhost:8080/Inventory_Strategic_OS_V3.html
```

### 重新組裝 V3（修改 part 檔案後必須執行）

```bash
cd C:\Users\jamic\庫存分析

# Step 1：合併 4 個 JS 部件 → v3_parts/part_js.txt
python assemble_js.py

# Step 2：組裝最終 HTML（含字串修復器）→ Inventory_Strategic_OS_V3.html
python reassemble_v3.py

# Step 3：語法驗證
node --check v3_debug/full_js.js
```

### 生成範例資料

```bash
# 產生 Sample_Data_V1.xlsx 和 Sample_Data_V2.xlsx
python generate_sample_data_v1_v2.py
```

### 驗證功能完整性

```bash
python -c "
import re, sys
sys.stdout.reconfigure(encoding='utf-8')
with open('Inventory_Strategic_OS_V3.html', 'r', encoding='utf-8') as f: c = f.read()
checks = ['generateDemoData','DOMContentLoaded','applyGlobalFilter','renderFIFO','renderEOQ','renderWhatIf']
[print('PASS' if t in c else 'FAIL', t) for t in checks]
"
```

---

## V3 建置架構（重要）

V3 **不直接編輯** `Inventory_Strategic_OS_V3.html`，而是修改部件檔案後重新組裝。

```
v3_parts/
├── part_html.txt        ← HTML 結構 + CSS（由 build_v3_html.py 生成）
├── part_js_p1.txt       ← JS：常數、狀態變數、processData()、全局篩選器
├── part_js_p2.txt       ← JS：renderOverview/Trend/Structure/Aging/Supply/Items
├── part_js_p3.txt       ← JS：renderForecast/EOQ/FIFO/WhatIf/Quality/Actions
├── part_js_p4.txt       ← JS：renderReport/export functions/DOMContentLoaded
└── part_js.txt          ← 由 assemble_js.py 合併 p1~p4（勿直接編輯）

reassemble_v3.py         ← 最終組裝腳本（含正則感知字串修復器）
Inventory_Strategic_OS_V3.html  ← 最終輸出（勿直接編輯）
```

**修改 CSS 或 HTML 結構** → 編輯 `v3_parts/part_html.txt`
**修改 JS 邏輯** → 編輯對應的 `v3_parts/part_js_p*.txt`
**組裝** → 執行 `python assemble_js.py && python reassemble_v3.py`

---

## 資料流與核心架構

```
rawData（原始日誌）
    │  processData(raw)
    ├─→ _rawDataFull（完整資料備份，全局篩選用）
    │
    ▼
aggregatedData {
    products[],      # 品項陣列（ABC/XYZ/KPI 已計算）
    monthly{},       # 月度聚合 { "YYYY-MM": {invVal, cogs, ...} }
    mKeys[],         # 月份鍵列表
    totalVal, totalCOGS, seasonalIndex, fisherRatio, qualityIssues
}
    │  updateDashboard() → updateFilterUI()
    │
    ▼
renderCurrentTab()
    └─→ renderMap['tab-xxx']()   # 懶載入，切換頁籤時才呼叫
```

**全局篩選器流程**：
```
globalFilter.{years, categories, search}
    │  applyGlobalFilter()
    ├─→ 過濾 _rawDataFull → filtered[]
    ├─→ _skipRawUpdate = true
    ├─→ processData(filtered)    # 重新聚合，不覆蓋 _rawDataFull
    └─→ _skipRawUpdate = false
```

---

## 關鍵設計限制（避免踩坑）

### CDN 版本鎖定

| 函式庫 | 正確 CDN | 常見錯誤 |
|--------|---------|---------|
| Chart.js | `cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js` | 全域名是 `window.Chart`，**不是** `ChartJS` |
| SheetJS | `cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js` | `cdnjs` 上 0.18.5 以後不更新，0.20.x 在 cdnjs 會 404 |
| chartjs-plugin-datalabels | `cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0` | 需用 `Chart.register(ChartDataLabels)`，不是 `ChartJS.register` |

### reassemble_v3.py 字串修復器

修復器必須辨識正則字面量（`/pattern/flags`）。**不要**用簡單的引號配對來解析 JS——`/"/g` 中的 `"` 不是字串開始。`is_regex_start()` 函式以前一個非空白字符是否為運算子來判斷。

### 頁面渲染懶載入

`renderedPages` Set 追蹤已渲染的頁籤 ID。更新資料後呼叫 `renderedPages.clear()` 強制所有頁籤在下次切換時重繪。若要強制重繪特定頁籤：`renderedPages.delete('tab-xxx')`。

### Canvas 熱力圖（供應鏈頁）

使用 Canvas 2D 手繪，需讀取容器 `clientWidth` 動態計算 `cellW`，並乘以 `window.devicePixelRatio` 設定實體像素、CSS 邏輯像素分開設定，否則 Retina 螢幕模糊。

### html2canvas 匯出限制

不支援 `backdrop-filter: blur()`。PDF 匯出前需啟用 `.export-mode` CSS class，在 `onclone` 回調中強制 inline style 為不透明背景色。

---

## 全局狀態變數（part_js_p1.txt）

| 變數 | 說明 |
|------|------|
| `rawData` | 當前工作資料集（篩選後可能是子集） |
| `_rawDataFull` | 永久保留的完整原始資料 |
| `aggregatedData` | processData() 計算的所有聚合結果 |
| `globalFilter` | `{ years: Set, categories: Set, search: string }` |
| `_skipRawUpdate` | `true` 時 processData 不覆蓋 rawData/_rawDataFull |
| `myCharts` | Chart.js 實例快取，切換模式前須逐一 `destroy()` |
| `renderedPages` | 已渲染頁籤的 Set，清空可強制全頁重繪 |
| `CFG` | 全域閾值常數（DAYS=180、ABC_A=0.5 等） |

---

## 文件參照

| 文件 | 用途 |
|------|------|
| `V3_使用說明書.md` | 完整使用者操作說明（14 頁籤、篩選器、Excel 格式） |
| `V3_修改紀錄.md` | 逐版次修改記錄（v3.0.0 ~ v3.0.3） |
| `V3_設計藍圖.md` | V3 原始設計規格（KPI 定義、演算法說明） |
| `Sample_Data_V1.xlsx` | V1/V2 相容的範例資料（50品項×180天） |
| `Sample_Data_V2.xlsx` | 同上，欄位稍有差異 |

---

## 溝通語言

**繁體中文**（對話、commit message、文件、程式碼註解皆使用繁體中文）
