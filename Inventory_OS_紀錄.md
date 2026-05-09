# 庫存分析智能儀表板 (Inventory Strategic OS) 開發紀錄

## 專案概述
本專案依據《PowerBI_庫存分析完整技術報告.docx》與《庫存計算 1-68.pdf》中的核心思維，以純網頁前端技術（HTML + TailwindCSS + Chart.js + SheetJS）完整復刻並強化了一系列高互動性的庫存分析儀表板。專案無需架設後端資料庫，透過 Excel (或 CSV) 上傳即可在瀏覽器端動態計算並產生視覺化圖表。

## 核心功能開發重點

### 1. 介面與使用者體驗 (UI/UX)
*   **設計風格**：採用深色毛玻璃 (Dark Glassmorphism) 設計，帶來高度沈浸、專業的商業智慧視覺感受。
*   **導覽架構**：規劃了 8 大功能分頁，包含總覽、趨勢、結構與ABC、庫齡風險、預測ROP、品項深度剖析、策略行動與資料設定。
*   **狀態即時性**：在頁首放置即時時鐘與「系統運作中」的呼吸燈動畫，增強軟體動態生命力。

### 2. 資料處理引擎 (SheetJS)
*   系統利用 `FileReader` 讀取使用者匯入的 Excel 檔案。
*   **資料聚合 (Aggregation)**：在 `processData` 函數中將逐日的異動紀錄 (Transaction-level) 轉換為依據品項聚合 (Product-level) 與按月聚合 (Monthly-level) 兩種維度。
    這取代了在 PowerBI 中撰寫大量的 DAX 量值，改為使用 JavaScript 的 Array `map`, `reduce`, `filter` 來達成相同的計算目的。

### 3. 分析模型實作 (還原 DAX 邏輯)
*   **存貨周轉率與 DSI**：使用聚合的「月均庫存」與「當月總銷貨成本」計算月化與年化的存貨周轉率，進而計算出 DSI（Days Sales of Inventory）。
*   **ABC 分析 (累計比例法)**：針對聚合後的現有庫存總額進行降冪排序，累計總額前 50% 標註為 A 類，50~70% 為 B 類，並提供視覺化的供需失衡矩陣圖（散佈圖）讓管理者找尋問題點。
*   **庫齡分布與減損**：利用資料中的 `Average_Aging_Days` 將庫存依據 30, 60, 90, 120 天切分成五個 Bucket，並對大於 30 天以上的各個級距乘上動態減損比例（例如 >120 天提列 60% 呆帳風險）。
*   **預測與 ROP (安全庫存)**：利用歷史平均日銷量與前置時間 (Lead time)，加上模擬的安全庫存係數，推導出各產品的 Reorder Point (ROP)。如果實際庫存 < ROP 則觸發警示列表。

### 4. 自動化生成範例資料 (Python 腳本)
為了解決純粹空白看板無示範效果的問題，編寫了 `generate_sample_data.py` 利用 `numpy` 與 `pandas` 產生了一套涵蓋 5 個品類、50 個品項、連續 180 天的模擬日常庫存動態紀錄 Excel。此腳本實作了常態分佈、季節性波動及隨機補貨邏輯，使得儀表板的視覺變化更具真實商業數據的美感。

## 待優化與未來展望
*   目前資料量限制在瀏覽器記憶體可以負載的範圍（實測 1~5萬筆交易量通常流暢）。未來若需要更大資料量整合，可考慮拆解成 Web Worker 或引入 SQLite WASM 進行本地端 SQL 分析。
*   預測邏輯目前使用單純的均值加成模型，未來可引入時間序列演算法 (如 Exponential Smoothing 或 ARIMA) 以 JavaScript 實作更精準的未來趨勢預測。

---

## 離線Standalone版本轉換紀錄 (2026-03-19)

### 問題背景
使用者在無網路連線的環境下需要使用本儀表板，但原始版本依賴多個外部 CDN：
- Tailwind CSS CDN
- Chart.js CDN
- SheetJS CDN
- Lucide Icons CDN

當 CDN 無法訪問時，圖表無法正常顯示。

### 解決方案
編寫 Python 轉換腳本 `create_standalone_v1.py`，將所有外部依賴內嵌到單一 HTML 檔案中。

### 轉換程序

1. **內嵌 Chart.js 與 SheetJS**
   - 從 `chartjs.txt` 與 `sheetjs.txt` 讀取已下載的庫檔案
   - 移除 CDN 引用標籤
   - 在 `</head>` 前插入內嵌的 `<script>` 標籤

2. **Lucide 圖示替代方案**
   - 建立 Lucide icon 到 emoji 的對照表
   - 將所有 `data-lucide="icon-name"` 屬性替換為 `title="icon-name"`
   - 將 `<i data-lucide="icon-name"></i>` 替換為 `<span>emoji</span>`

3. **Tailwind CSS 等效樣式**
   - 將所有使用的 Tailwind 實用類別轉換為純 CSS
   - 建立完整的 CSS 樣式表，包含：
     - 顯示 (display)、彈性盒 (flex)、網格 (grid)
     - 間距 (padding/margin)
     - 文字樣式 (字體大小、顏色，粗細)
     - 背景與邊框樣式
     - 陰影、圓角、動畫等

4. **Lucide Stub 問題修復**
   - 移除 Lucide CDN 後，程式碼仍呼叫 `lucide.createIcons()`
   - 拋出 `ReferenceError: lucide is not defined`
   - 在第一次呼叫前加入 Stub：
     ```javascript
     if (typeof lucide === 'undefined') { var lucide = { createIcons: function() {} }; }
     ```

### 產出檔案
- `Inventory_Strategic_OS_Standalone.html` (約 963KB)
- 包含內嵌的 Chart.js (約 205KB) 與 SheetJS (約 709KB)

### 轉換腳本
```python
# create_standalone_v1.py 主要功能
- 讀取 chartjs.txt, sheetjs.txt, Inventory_Strategic_OS.html
- 移除 CDN 引用
- 內嵌 JavaScript 庫
- 替換 Lucide 圖示為 emoji
- 添加 Tailwind 等效 CSS
- 輸出 _Standalone.html
```

### 技術要點
- `@media` 規則必須放在 CSS 根層級，不能巢狀在選擇器內
- `lucide.createIcons()` 呼叫需使用 `typeof lucide !== 'undefined'` 條件檢查
- Emoji 替代方案兼顧美觀與效能
