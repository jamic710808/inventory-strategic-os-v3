#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_v3_html.py
生成庫存分析智能儀表板 V3.0 的 HTML 結構與 CSS 部分（不含 JavaScript）
輸出至 v3_parts/part_html.txt
"""

import os

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "v3_parts")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "part_html.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)

html = r"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>庫存分析智能儀表板 V3.0</title>

  <!-- CDN Scripts -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
  <script src="https://unpkg.com/lucide@latest"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>

  <style>
    /* ============================================================
       基礎 / 深色模式
    ============================================================ */
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Segoe UI', 'PingFang TC', 'Microsoft JhengHei', sans-serif;
      background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #020617 100%);
      background-attachment: fixed;
      color: #e2e8f0;
      min-height: 100vh;
    }

    /* ============================================================
       淺色模式
    ============================================================ */
    body.light-mode {
      background: linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 50%, #f8fafc 100%);
      background-attachment: fixed;
      color: #1e293b;
    }
    body.light-mode .glass-panel {
      background: rgba(255, 255, 255, 0.85);
      border: 1px solid rgba(100, 116, 139, 0.2);
    }
    body.light-mode .glass-header {
      background: rgba(241, 245, 249, 0.85);
      border-bottom: 1px solid rgba(100, 116, 139, 0.2);
    }
    body.light-mode .glass-btn {
      background: rgba(59, 130, 246, 0.12);
      color: #1e40af;
      border-color: rgba(59, 130, 246, 0.25);
    }
    body.light-mode .glass-btn.active {
      background: linear-gradient(135deg, #2563eb, #1d4ed8);
      color: #ffffff;
    }
    body.light-mode .kpi-card {
      background: rgba(255, 255, 255, 0.9);
      border: 1px solid rgba(100, 116, 139, 0.15);
    }
    body.light-mode .data-table th {
      background: rgba(241, 245, 249, 0.9);
      color: #475569;
    }
    body.light-mode .data-table td {
      border-color: rgba(100, 116, 139, 0.15);
    }
    body.light-mode .tab-sub-btn {
      background: rgba(59, 130, 246, 0.1);
      color: #1e40af;
    }
    body.light-mode .tab-sub-btn.active {
      background: linear-gradient(135deg, #3b82f6, #2563eb);
      color: #ffffff;
    }
    body.light-mode input[type="range"] { accent-color: #3b82f6; }
    body.light-mode input[type="number"],
    body.light-mode input[type="text"],
    body.light-mode select {
      background: rgba(255, 255, 255, 0.9);
      color: #1e293b;
      border-color: rgba(100, 116, 139, 0.3);
    }
    body.light-mode .accordion-header {
      background: rgba(241, 245, 249, 0.8);
      color: #1e293b;
    }
    body.light-mode .accordion-body {
      background: rgba(255, 255, 255, 0.7);
      color: #334155;
    }
    body.light-mode .heatmap-cell { color: #0f172a; }
    body.light-mode ::-webkit-scrollbar-track { background: #e2e8f0; }
    body.light-mode ::-webkit-scrollbar-thumb { background: #94a3b8; }

    /* ============================================================
       Glass UI
    ============================================================ */
    .glass-panel {
      background: rgba(30, 41, 59, 0.4);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border: 1px solid rgba(255, 255, 255, 0.1);
      border-radius: 1rem;
    }
    .glass-header {
      position: sticky;
      top: 0;
      z-index: 50;
      background: rgba(15, 23, 42, 0.75);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border-bottom: 1px solid rgba(255, 255, 255, 0.08);
    }
    .glass-btn {
      display: inline-flex;
      align-items: center;
      gap: 0.35rem;
      padding: 0.35rem 0.75rem;
      border-radius: 0.5rem;
      font-size: 0.75rem;
      font-weight: 500;
      cursor: pointer;
      border: 1px solid rgba(255, 255, 255, 0.12);
      background: rgba(59, 130, 246, 0.15);
      color: #93c5fd;
      transition: background 0.2s, color 0.2s, transform 0.15s;
      white-space: nowrap;
    }
    .glass-btn:hover { background: rgba(59, 130, 246, 0.28); color: #bfdbfe; }
    .glass-btn.active {
      background: linear-gradient(135deg, #3b82f6, #2563eb);
      color: #ffffff;
      border-color: transparent;
    }

    /* ============================================================
       Tab 系統
    ============================================================ */
    .tab-content { display: none; }
    .tab-content.active {
      display: block;
      animation: fadeIn 0.3s ease;
    }
    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(6px); }
      to   { opacity: 1; transform: translateY(0); }
    }

    /* 子分頁按鈕 */
    .tab-sub-btn {
      padding: 0.3rem 0.9rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 500;
      cursor: pointer;
      border: 1px solid rgba(255,255,255,0.12);
      background: rgba(255,255,255,0.08);
      color: #94a3b8;
      transition: background 0.2s, color 0.2s;
    }
    .tab-sub-btn.active {
      background: linear-gradient(135deg, #3b82f6, #2563eb);
      color: #ffffff;
      border-color: transparent;
    }

    /* ============================================================
       KPI Cards
    ============================================================ */
    .kpi-card {
      background: rgba(30, 41, 59, 0.5);
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 0.875rem;
      padding: 1rem 1.25rem;
      display: flex;
      flex-direction: column;
      gap: 0.4rem;
      transition: transform 0.2s, box-shadow 0.2s;
    }
    .kpi-card:hover {
      transform: translateY(-3px);
      box-shadow: 0 8px 32px rgba(59, 130, 246, 0.18);
    }
    .kpi-label { font-size: 0.7rem; letter-spacing: 0.06em; color: #94a3b8; text-transform: uppercase; }
    .kpi-value { font-size: 1.75rem; font-weight: 700; line-height: 1.1; }
    .kpi-unit  { font-size: 0.7rem; color: #64748b; }
    .kpi-trend { font-size: 0.72rem; font-weight: 500; }

    /* ============================================================
       Charts
    ============================================================ */
    .chart-container {
      position: relative;
      height: 320px;
      width: 100%;
    }
    .chart-container-sm { position: relative; height: 220px; width: 100%; }
    .chart-container-lg { position: relative; height: 420px; width: 100%; }

    /* ============================================================
       Tables
    ============================================================ */
    .data-table { width: 100%; border-collapse: collapse; font-size: 0.8rem; }
    .data-table th {
      background: rgba(15, 23, 42, 0.6);
      color: #94a3b8;
      font-weight: 600;
      letter-spacing: 0.04em;
      padding: 0.6rem 0.8rem;
      text-align: left;
      border-bottom: 1px solid rgba(255, 255, 255, 0.08);
      white-space: nowrap;
    }
    .data-table td {
      padding: 0.55rem 0.8rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.05);
      vertical-align: middle;
    }
    .data-table tbody tr:hover { background: rgba(59, 130, 246, 0.08); }
    .table-wrapper { overflow-x: auto; border-radius: 0.75rem; }

    /* ============================================================
       Badges
    ============================================================ */
    .badge {
      display: inline-flex; align-items: center; justify-content: center;
      padding: 0.15rem 0.5rem;
      border-radius: 9999px;
      font-size: 0.65rem;
      font-weight: 700;
      letter-spacing: 0.06em;
    }
    .badge-a  { background: rgba(59,130,246,0.25); color: #60a5fa; border: 1px solid rgba(59,130,246,0.4); }
    .badge-b  { background: rgba(245,158,11,0.25); color: #fbbf24; border: 1px solid rgba(245,158,11,0.4); }
    .badge-c  { background: rgba(100,116,139,0.25); color: #94a3b8; border: 1px solid rgba(100,116,139,0.4); }
    .badge-x  { background: rgba(16,185,129,0.25); color: #34d399; border: 1px solid rgba(16,185,129,0.4); }
    .badge-y  { background: rgba(234,179,8,0.25);  color: #facc15; border: 1px solid rgba(234,179,8,0.4); }
    .badge-z  { background: rgba(239,68,68,0.25);  color: #f87171; border: 1px solid rgba(239,68,68,0.4); }
    .badge-ax { background: rgba(59,130,246,0.35); color: #93c5fd; border: 1px solid rgba(59,130,246,0.5); }

    /* ============================================================
       Progress Bars
    ============================================================ */
    .progress-bar {
      width: 100%;
      height: 4px;
      background: rgba(255, 255, 255, 0.08);
      border-radius: 9999px;
      overflow: hidden;
      margin-top: 0.35rem;
    }
    .progress-fill {
      height: 100%;
      border-radius: 9999px;
      background: linear-gradient(90deg, #3b82f6, #06b6d4);
      transition: width 0.6s ease;
    }
    .progress-fill.warning { background: linear-gradient(90deg, #f59e0b, #ef4444); }
    .progress-fill.success { background: linear-gradient(90deg, #10b981, #06b6d4); }

    /* ============================================================
       Alert Colours
    ============================================================ */
    .alert-green  { color: #34d399; }
    .alert-yellow { color: #facc15; }
    .alert-red    { color: #f87171; }

    .alert-card {
      border-radius: 0.75rem;
      padding: 0.85rem 1rem;
      border-left: 4px solid;
      margin-bottom: 0.6rem;
    }
    .alert-card.sev-high   { border-color: #ef4444; background: rgba(239,68,68,0.08); }
    .alert-card.sev-medium { border-color: #f59e0b; background: rgba(245,158,11,0.08); }
    .alert-card.sev-low    { border-color: #3b82f6; background: rgba(59,130,246,0.08); }

    /* ============================================================
       Heatmap
    ============================================================ */
    .heatmap-cell {
      text-align: center;
      font-size: 0.7rem;
      font-weight: 600;
      padding: 0.45rem 0.3rem;
      border-radius: 0.25rem;
      color: #fff;
      transition: opacity 0.2s;
    }
    .heatmap-cell:hover { opacity: 0.8; }

    /* ============================================================
       ABC-XYZ Matrix
    ============================================================ */
    .abcxyz-cell {
      border-radius: 0.5rem;
      padding: 0.6rem;
      text-align: center;
      cursor: pointer;
      transition: transform 0.15s, box-shadow 0.15s;
      min-height: 5rem;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 0.2rem;
    }
    .abcxyz-cell:hover { transform: scale(1.04); box-shadow: 0 4px 16px rgba(0,0,0,0.3); }
    .abcxyz-cell .cell-label { font-size: 1rem; font-weight: 800; }
    .abcxyz-cell .cell-count { font-size: 0.65rem; color: rgba(255,255,255,0.75); }

    /* Colour map for 9 cells */
    .cell-ax { background: linear-gradient(135deg, #1d4ed8, #2563eb); }
    .cell-ay { background: linear-gradient(135deg, #1e40af, #3b82f6); }
    .cell-az { background: linear-gradient(135deg, #7c3aed, #8b5cf6); }
    .cell-bx { background: linear-gradient(135deg, #065f46, #059669); }
    .cell-by { background: linear-gradient(135deg, #d97706, #f59e0b); }
    .cell-bz { background: linear-gradient(135deg, #9a3412, #ea580c); }
    .cell-cx { background: linear-gradient(135deg, #0e7490, #06b6d4); }
    .cell-cy { background: linear-gradient(135deg, #374151, #6b7280); }
    .cell-cz { background: linear-gradient(135deg, #7f1d1d, #ef4444); }

    /* ============================================================
       Drop Zone
    ============================================================ */
    #drop-zone {
      border: 2px dashed rgba(99, 102, 241, 0.45);
      border-radius: 1rem;
      padding: 2.5rem;
      text-align: center;
      cursor: pointer;
      transition: border-color 0.2s, background 0.2s;
    }
    #drop-zone.drag-over {
      border-color: #6366f1;
      background: rgba(99, 102, 241, 0.08);
    }
    #drop-zone:hover { border-color: rgba(99, 102, 241, 0.7); }

    /* ============================================================
       Accordion (DAX)
    ============================================================ */
    .accordion-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 0.75rem 1rem;
      background: rgba(30, 41, 59, 0.6);
      border-radius: 0.5rem;
      cursor: pointer;
      font-size: 0.82rem;
      font-weight: 600;
      color: #93c5fd;
      user-select: none;
      transition: background 0.2s;
    }
    .accordion-header:hover { background: rgba(59, 130, 246, 0.15); }
    .accordion-body {
      display: none;
      padding: 0.85rem 1rem;
      background: rgba(15, 23, 42, 0.4);
      border-radius: 0 0 0.5rem 0.5rem;
      font-size: 0.78rem;
    }
    .accordion-body.open { display: block; animation: fadeIn 0.2s ease; }
    .accordion-body pre {
      background: rgba(0, 0, 0, 0.3);
      border-radius: 0.4rem;
      padding: 0.75rem;
      overflow-x: auto;
      font-family: 'Cascadia Code', 'Consolas', monospace;
      font-size: 0.72rem;
      line-height: 1.6;
      color: #a5f3fc;
      white-space: pre-wrap;
      word-break: break-word;
    }

    /* ============================================================
       Form Controls
    ============================================================ */
    input[type="range"] {
      width: 100%;
      accent-color: #6366f1;
      cursor: pointer;
    }
    input[type="number"],
    input[type="text"],
    select {
      background: rgba(30, 41, 59, 0.6);
      color: #e2e8f0;
      border: 1px solid rgba(255, 255, 255, 0.1);
      border-radius: 0.5rem;
      padding: 0.4rem 0.7rem;
      font-size: 0.8rem;
      outline: none;
      transition: border-color 0.2s;
    }
    input[type="number"]:focus,
    input[type="text"]:focus,
    select:focus { border-color: #6366f1; }

    /* ============================================================
       Gauge (Quality Score)
    ============================================================ */
    .gauge-wrap { position: relative; width: 140px; height: 80px; margin: 0 auto; }
    .gauge-svg { width: 140px; height: 80px; }
    .gauge-text {
      position: absolute;
      bottom: 0; left: 0; right: 0;
      text-align: center;
      font-size: 1.4rem;
      font-weight: 700;
    }

    /* ============================================================
       Scrollbar
    ============================================================ */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: rgba(15, 23, 42, 0.5); border-radius: 3px; }
    ::-webkit-scrollbar-thumb { background: rgba(99, 102, 241, 0.5); border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: rgba(99, 102, 241, 0.8); }

    /* ============================================================
       Export Mode (PDF — no backdrop-filter)
    ============================================================ */
    .export-mode .glass-panel {
      background: #1e293b !important;
      backdrop-filter: none !important;
      -webkit-backdrop-filter: none !important;
    }
    .export-mode.light-mode .glass-panel { background: #ffffff !important; }

    /* ============================================================
       Utility helpers
    ============================================================ */
    .text-muted  { color: #64748b; }
    .text-accent { color: #6366f1; }
    .divider { border-top: 1px solid rgba(255,255,255,0.07); margin: 1rem 0; }
    .section-title {
      font-size: 0.7rem;
      letter-spacing: 0.1em;
      text-transform: uppercase;
      color: #64748b;
      margin-bottom: 0.75rem;
    }

    /* ============================================================
       Responsive
    ============================================================ */
    @media (max-width: 1024px) {
      .nav-scroll { overflow-x: auto; }
      .kpi-grid { grid-template-columns: repeat(2, 1fr) !important; }
    }
    @media (max-width: 640px) {
      .kpi-grid { grid-template-columns: 1fr !important; }
      .chart-container { height: 240px; }
    }
  </style>
</head>
<body>

<!-- ================================================================
     HEADER
================================================================ -->
<header class="glass-header px-4 py-2">
  <div class="max-w-[1600px] mx-auto flex items-center gap-3">

    <!-- Left: Logo + Title -->
    <div class="flex items-center gap-2 shrink-0">
      <div class="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-500 to-indigo-600 flex items-center justify-center">
        <i data-lucide="boxes" class="w-4 h-4 text-white"></i>
      </div>
      <div class="leading-tight">
        <div class="text-sm font-bold text-white">庫存分析智能儀表板</div>
        <div class="text-[0.6rem] text-slate-400">Inventory Strategic OS v3.0</div>
      </div>
    </div>

    <!-- Center: Nav -->
    <nav class="flex-1 nav-scroll">
      <div class="flex items-center gap-1 min-w-max">
        <button class="glass-btn active" data-tab="tab-overview">
          <i data-lucide="layout-dashboard" class="w-3 h-3"></i>總覽
        </button>
        <button class="glass-btn" data-tab="tab-trend">
          <i data-lucide="trending-up" class="w-3 h-3"></i>趨勢周轉
        </button>
        <button class="glass-btn" data-tab="tab-structure">
          <i data-lucide="pie-chart" class="w-3 h-3"></i>結構ABC
        </button>
        <button class="glass-btn" data-tab="tab-aging">
          <i data-lucide="clock" class="w-3 h-3"></i>庫齡風險
        </button>
        <button class="glass-btn" data-tab="tab-forecast">
          <i data-lucide="line-chart" class="w-3 h-3"></i>預測補貨
        </button>
        <button class="glass-btn" data-tab="tab-supply">
          <i data-lucide="link" class="w-3 h-3"></i>供應鏈
        </button>
        <button class="glass-btn" data-tab="tab-items">
          <i data-lucide="list" class="w-3 h-3"></i>品項剖析
        </button>
        <button class="glass-btn" data-tab="tab-actions">
          <i data-lucide="zap" class="w-3 h-3"></i>策略行動
        </button>
        <button class="glass-btn" data-tab="tab-whatif">
          <i data-lucide="sliders" class="w-3 h-3"></i>情境模擬
        </button>
        <button class="glass-btn" data-tab="tab-eoq">
          <i data-lucide="calculator" class="w-3 h-3"></i>EOQ成本
        </button>
        <button class="glass-btn" data-tab="tab-fifo">
          <i data-lucide="layers" class="w-3 h-3"></i>FIFO批次
        </button>
        <button class="glass-btn" data-tab="tab-quality">
          <i data-lucide="shield-check" class="w-3 h-3"></i>資料品質
        </button>
        <button class="glass-btn" data-tab="tab-report">
          <i data-lucide="file-text" class="w-3 h-3"></i>報告中心
        </button>
        <button class="glass-btn" data-tab="tab-settings">
          <i data-lucide="settings" class="w-3 h-3"></i>設定DAX
        </button>
      </div>
    </nav>

    <!-- Right: Controls -->
    <div class="flex items-center gap-2 shrink-0">
      <!-- Clock -->
      <span id="header-clock" class="text-xs text-slate-400 tabular-nums hidden md:block"></span>
      <!-- System status -->
      <span class="flex items-center gap-1 text-[0.65rem] text-emerald-400 hidden md:flex">
        <span class="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>系統運作中
      </span>
      <!-- Dark / Light toggle -->
      <button id="btn-darkmode" class="glass-btn" title="切換深淺色模式">
        <i data-lucide="moon" class="w-3 h-3"></i>
      </button>
      <!-- Upload -->
      <button class="glass-btn text-blue-400" onclick="document.getElementById('file-upload').click()">
        <i data-lucide="upload-cloud" class="w-3 h-3"></i>上傳資料
      </button>
      <input type="file" id="file-upload" accept=".xlsx,.xls,.csv" class="hidden" />
      <!-- Download -->
      <button id="btn-header-download" class="glass-btn text-emerald-400">
        <i data-lucide="download" class="w-3 h-3"></i>下載資料
      </button>
    </div>

  </div>
</header>

<!-- ================================================================
     MAIN CONTENT
================================================================ -->
<main class="p-4 max-w-[1600px] mx-auto">

  <!-- ==============================================================
       TAB 1: 總覽 Overview
  ============================================================== -->
  <div id="tab-overview" class="tab-content active">

    <!-- KPI Row 1 -->
    <div class="grid grid-cols-4 gap-3 mb-3 kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">庫存總金額</div>
        <div class="flex items-end gap-1">
          <span id="kpi-total-inv" class="kpi-value text-blue-400">—</span>
          <span class="kpi-unit mb-1">萬元</span>
        </div>
        <div class="progress-bar"><div id="pb-total-inv" class="progress-fill" style="width:0%"></div></div>
        <div id="trend-total-inv" class="kpi-trend alert-green">—</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">庫存天數 DSI</div>
        <div class="flex items-end gap-1">
          <span id="kpi-dsi" class="kpi-value text-amber-400">—</span>
          <span class="kpi-unit mb-1">天</span>
        </div>
        <div class="progress-bar"><div id="pb-dsi" class="progress-fill warning" style="width:0%"></div></div>
        <div id="trend-dsi" class="kpi-trend alert-yellow">—</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">庫齡 &gt;90天比例</div>
        <div class="flex items-end gap-1">
          <span id="kpi-aging" class="kpi-value text-red-400">—</span>
          <span class="kpi-unit mb-1">%</span>
        </div>
        <div class="progress-bar"><div id="pb-aging" class="progress-fill warning" style="width:0%"></div></div>
        <div id="trend-aging" class="kpi-trend alert-red">—</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">庫存健康分</div>
        <div class="flex items-end gap-1">
          <span id="kpi-health" class="kpi-value text-emerald-400">—</span>
          <span class="kpi-unit mb-1">/ 100</span>
        </div>
        <div class="progress-bar"><div id="pb-health" class="progress-fill success" style="width:0%"></div></div>
        <div id="trend-health" class="kpi-trend alert-green">—</div>
      </div>
    </div>

    <!-- KPI Row 2 -->
    <div class="grid grid-cols-4 gap-3 mb-3 kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">存貨周轉率</div>
        <div class="flex items-end gap-1">
          <span id="kpi-turnover" class="kpi-value text-indigo-400">—</span>
          <span class="kpi-unit mb-1">次/年</span>
        </div>
        <div class="progress-bar"><div id="pb-turnover" class="progress-fill" style="width:0%"></div></div>
        <div id="trend-turnover" class="kpi-trend alert-green">—</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">存銷比 STR</div>
        <div class="flex items-end gap-1">
          <span id="kpi-str" class="kpi-value text-cyan-400">—</span>
          <span class="kpi-unit mb-1">快銷</span>
        </div>
        <div class="progress-bar"><div id="pb-str" class="progress-fill success" style="width:0%"></div></div>
        <div id="trend-str" class="kpi-trend alert-green">—</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">GMROI</div>
        <div class="flex items-end gap-1">
          <span id="kpi-gmroi" class="kpi-value text-purple-400">—</span>
          <span class="kpi-unit mb-1">倍</span>
        </div>
        <div class="progress-bar"><div id="pb-gmroi" class="progress-fill" style="width:0%"></div></div>
        <div id="trend-gmroi" class="kpi-trend alert-green">—</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">缺貨次數</div>
        <div class="flex items-end gap-1">
          <span id="kpi-stockouts" class="kpi-value text-rose-400">—</span>
          <span class="kpi-unit mb-1">次</span>
        </div>
        <div class="progress-bar"><div id="pb-stockouts" class="progress-fill warning" style="width:0%"></div></div>
        <div id="trend-stockouts" class="kpi-trend alert-red">—</div>
      </div>
    </div>

    <!-- Charts Row -->
    <div class="grid grid-cols-3 gap-3">
      <div class="glass-panel p-4 col-span-2">
        <div class="section-title">庫存金額趨勢 (12個月)</div>
        <div class="chart-container">
          <canvas id="chartTrendOverview"></canvas>
        </div>
      </div>
      <div class="glass-panel p-4">
        <div class="section-title">品類庫存佔比</div>
        <div class="chart-container">
          <canvas id="chartCategoryOverview"></canvas>
        </div>
      </div>
    </div>

  </div><!-- /tab-overview -->

  <!-- ==============================================================
       TAB 2: 趨勢周轉 Trend
  ============================================================== -->
  <div id="tab-trend" class="tab-content">
    <div class="grid grid-cols-2 gap-3 mb-3">
      <div class="glass-panel p-4">
        <div class="section-title">DSI / 存貨周轉率 (雙軸)</div>
        <div class="chart-container">
          <canvas id="chartTurnover"></canvas>
        </div>
      </div>
      <div class="glass-panel p-4">
        <div class="flex items-center justify-between mb-2">
          <div class="section-title">存銷比 SSR</div>
          <div class="flex gap-1">
            <button id="btn-ssr-long"  class="tab-sub-btn active">長週期SSR</button>
            <button id="btn-ssr-short" class="tab-sub-btn">短週期SSR</button>
          </div>
        </div>
        <div class="chart-container">
          <canvas id="chartSSR"></canvas>
        </div>
      </div>
    </div>
    <div class="grid grid-cols-2 gap-3">
      <div class="glass-panel p-4">
        <div class="section-title">月環比成長率 MoM Growth</div>
        <div class="chart-container">
          <canvas id="chartGrowth"></canvas>
        </div>
      </div>
      <div class="glass-panel p-4">
        <div class="section-title">移動平均線 (Actual / 3M MA / 6M MA)</div>
        <div class="chart-container">
          <canvas id="chartMA"></canvas>
        </div>
      </div>
    </div>
  </div><!-- /tab-trend -->

  <!-- ==============================================================
       TAB 3: 結構ABC Structure
  ============================================================== -->
  <div id="tab-structure" class="tab-content">
    <!-- ABC Toggle -->
    <div class="flex items-center gap-2 mb-3">
      <span class="text-xs text-slate-400">ABC 分析方法：</span>
      <button id="btn-abc-rel" class="tab-sub-btn active">相對法</button>
      <button id="btn-abc-abs" class="tab-sub-btn">絕對法</button>
    </div>

    <div class="grid grid-cols-3 gap-3 mb-3">
      <div class="glass-panel p-4">
        <div class="section-title">ABC 佔比</div>
        <div class="chart-container">
          <canvas id="chartABC"></canvas>
        </div>
      </div>
      <div class="glass-panel p-4">
        <div class="section-title">供需失衡散佈圖</div>
        <div class="chart-container">
          <canvas id="chartImbalance"></canvas>
        </div>
      </div>
      <div class="glass-panel p-4">
        <div class="section-title">Pareto 累積曲線</div>
        <div class="chart-container">
          <canvas id="chartPareto"></canvas>
        </div>
      </div>
    </div>

    <!-- ABC-XYZ 9-cell Matrix -->
    <div class="glass-panel p-4">
      <div class="section-title mb-3">ABC-XYZ 九宮格矩陣</div>
      <div class="overflow-x-auto">
        <table id="abcxyz-matrix" class="w-full" style="border-collapse:separate; border-spacing:6px;">
          <thead>
            <tr>
              <th class="text-xs text-slate-500 p-2 text-right w-16">ABC ↓ / XYZ →</th>
              <th class="text-xs text-emerald-400 p-2 text-center">X (高穩定)</th>
              <th class="text-xs text-yellow-400 p-2 text-center">Y (中波動)</th>
              <th class="text-xs text-red-400 p-2 text-center">Z (高波動)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td class="text-xs text-blue-400 font-bold p-2 text-right">A (高值)</td>
              <td><div class="abcxyz-cell cell-ax" data-cell="AX"><span class="cell-label">AX</span><span class="cell-count" id="cnt-ax">—品</span></div></td>
              <td><div class="abcxyz-cell cell-ay" data-cell="AY"><span class="cell-label">AY</span><span class="cell-count" id="cnt-ay">—品</span></div></td>
              <td><div class="abcxyz-cell cell-az" data-cell="AZ"><span class="cell-label">AZ</span><span class="cell-count" id="cnt-az">—品</span></div></td>
            </tr>
            <tr>
              <td class="text-xs text-amber-400 font-bold p-2 text-right">B (中值)</td>
              <td><div class="abcxyz-cell cell-bx" data-cell="BX"><span class="cell-label">BX</span><span class="cell-count" id="cnt-bx">—品</span></div></td>
              <td><div class="abcxyz-cell cell-by" data-cell="BY"><span class="cell-label">BY</span><span class="cell-count" id="cnt-by">—品</span></div></td>
              <td><div class="abcxyz-cell cell-bz" data-cell="BZ"><span class="cell-label">BZ</span><span class="cell-count" id="cnt-bz">—品</span></div></td>
            </tr>
            <tr>
              <td class="text-xs text-slate-400 font-bold p-2 text-right">C (低值)</td>
              <td><div class="abcxyz-cell cell-cx" data-cell="CX"><span class="cell-label">CX</span><span class="cell-count" id="cnt-cx">—品</span></div></td>
              <td><div class="abcxyz-cell cell-cy" data-cell="CY"><span class="cell-label">CY</span><span class="cell-count" id="cnt-cy">—品</span></div></td>
              <td><div class="abcxyz-cell cell-cz" data-cell="CZ"><span class="cell-label">CZ</span><span class="cell-count" id="cnt-cz">—品</span></div></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div><!-- /tab-structure -->

  <!-- ==============================================================
       TAB 4: 庫齡風險 Aging
  ============================================================== -->
  <div id="tab-aging" class="tab-content">
    <div class="grid grid-cols-3 gap-3 mb-3">
      <div class="glass-panel p-4 col-span-2">
        <div class="section-title">庫齡分層分析 (5級)</div>
        <div class="chart-container">
          <canvas id="chartAging"></canvas>
        </div>
      </div>
      <div class="flex flex-col gap-3">
        <!-- Loss Estimate -->
        <div class="glass-panel p-4 flex-1">
          <div class="section-title">預估損失金額</div>
          <div id="loss-estimate" class="space-y-2 text-sm">
            <div class="flex justify-between"><span class="text-slate-400">30-60天</span><span id="loss-30" class="text-yellow-400">—</span></div>
            <div class="flex justify-between"><span class="text-slate-400">60-90天</span><span id="loss-60" class="text-amber-400">—</span></div>
            <div class="flex justify-between"><span class="text-slate-400">90-180天</span><span id="loss-90" class="text-orange-400">—</span></div>
            <div class="flex justify-between"><span class="text-slate-400">&gt;180天</span><span id="loss-180" class="text-red-400">—</span></div>
            <div class="divider"></div>
            <div class="flex justify-between font-bold"><span>合計</span><span id="loss-total" class="text-red-400">—</span></div>
          </div>
        </div>
        <!-- Divergence Signal -->
        <div class="glass-panel p-4 flex-1">
          <div class="section-title">背離信號偵測</div>
          <div id="divergence-panel" class="space-y-2 text-xs text-slate-400">
            <div class="text-center py-2">載入資料後顯示</div>
          </div>
        </div>
      </div>
    </div>

    <div class="grid grid-cols-2 gap-3">
      <div class="glass-panel p-4">
        <div class="section-title">庫存健康雷達</div>
        <div class="chart-container">
          <canvas id="chartHealth"></canvas>
        </div>
      </div>
      <!-- Configurable Loss Rates -->
      <div class="glass-panel p-4">
        <div class="section-title">損失率參數設定</div>
        <div class="space-y-3 text-sm">
          <div class="flex items-center justify-between gap-3">
            <label class="text-slate-400 whitespace-nowrap w-28">0-30天損失率</label>
            <input type="number" id="loss-rate-0" class="w-24" value="0" min="0" max="100" step="1" />
            <span class="text-slate-500">%</span>
          </div>
          <div class="flex items-center justify-between gap-3">
            <label class="text-slate-400 whitespace-nowrap w-28">30-60天損失率</label>
            <input type="number" id="loss-rate-30" class="w-24" value="5" min="0" max="100" step="1" />
            <span class="text-slate-500">%</span>
          </div>
          <div class="flex items-center justify-between gap-3">
            <label class="text-slate-400 whitespace-nowrap w-28">60-90天損失率</label>
            <input type="number" id="loss-rate-60" class="w-24" value="15" min="0" max="100" step="1" />
            <span class="text-slate-500">%</span>
          </div>
          <div class="flex items-center justify-between gap-3">
            <label class="text-slate-400 whitespace-nowrap w-28">90-180天損失率</label>
            <input type="number" id="loss-rate-90" class="w-24" value="30" min="0" max="100" step="1" />
            <span class="text-slate-500">%</span>
          </div>
          <div class="flex items-center justify-between gap-3">
            <label class="text-slate-400 whitespace-nowrap w-28">&gt;180天損失率</label>
            <input type="number" id="loss-rate-180" class="w-24" value="60" min="0" max="100" step="1" />
            <span class="text-slate-500">%</span>
          </div>
          <button id="btn-update-loss" class="glass-btn mt-2">
            <i data-lucide="refresh-cw" class="w-3 h-3"></i>更新計算
          </button>
        </div>
      </div>
    </div>
  </div><!-- /tab-aging -->

  <!-- ==============================================================
       TAB 5: 預測補貨 Forecast
  ============================================================== -->
  <div id="tab-forecast" class="tab-content">
    <div class="grid grid-cols-3 gap-3 mb-3">
      <!-- Fisher Test Panel -->
      <div class="glass-panel p-4">
        <div class="section-title">Fisher 週期性檢定</div>
        <div id="fisher-panel" class="space-y-2 text-xs">
          <div class="text-slate-400 text-center py-3">載入資料後顯示</div>
        </div>
      </div>
      <!-- Seasonal Radar -->
      <div class="glass-panel p-4 col-span-2">
        <div class="section-title">季節性指數 (12個月)</div>
        <div class="chart-container">
          <canvas id="chartSeasonalRadar"></canvas>
        </div>
      </div>
    </div>

    <div class="glass-panel p-4 mb-3">
      <div class="flex items-center justify-between mb-2">
        <div class="section-title">複合預測模型 (移動均值 / 趨勢 / 複合)</div>
        <div class="flex items-center gap-2">
          <span class="text-xs text-slate-400">服務水準：</span>
          <button id="btn-svc-90" class="tab-sub-btn">90%</button>
          <button id="btn-svc-95" class="tab-sub-btn active">95%</button>
          <button id="btn-svc-99" class="tab-sub-btn">99%</button>
        </div>
      </div>
      <div class="chart-container-lg">
        <canvas id="chartForecast"></canvas>
      </div>
    </div>

    <!-- ROP Table -->
    <div class="glass-panel p-4">
      <div class="section-title">再訂購點 (ROP) 建議清單</div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>
              <th>品名</th><th>日均需求</th><th>前置時間(天)</th>
              <th>安全庫存</th><th>ROP</th><th>現有庫存</th><th>狀態</th>
            </tr>
          </thead>
          <tbody id="rop-table-body">
            <tr><td colspan="7" class="text-center text-slate-500 py-4">載入資料後顯示</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div><!-- /tab-forecast -->

  <!-- ==============================================================
       TAB 6: 供應鏈 Supply
  ============================================================== -->
  <div id="tab-supply" class="tab-content">
    <!-- Supply KPIs -->
    <div class="grid grid-cols-3 gap-3 mb-3">
      <div class="kpi-card">
        <div class="kpi-label">整體 GMROI</div>
        <div class="flex items-end gap-1">
          <span id="supply-gmroi" class="kpi-value text-purple-400">—</span>
          <span class="kpi-unit mb-1">倍</span>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">缺貨損失估算</div>
        <div class="flex items-end gap-1">
          <span id="supply-stockout-loss" class="kpi-value text-red-400">—</span>
          <span class="kpi-unit mb-1">萬元</span>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">平均前置時間</div>
        <div class="flex items-end gap-1">
          <span id="supply-avg-leadtime" class="kpi-value text-cyan-400">—</span>
          <span class="kpi-unit mb-1">天</span>
        </div>
      </div>
    </div>

    <div class="grid grid-cols-2 gap-3 mb-3">
      <div class="glass-panel p-4">
        <div class="section-title">各品類 GMROI 比較</div>
        <div class="chart-container">
          <canvas id="chartGMROI"></canvas>
        </div>
      </div>
      <div class="glass-panel p-4">
        <div class="section-title">供需熱力圖 (品類 × 月份)</div>
        <div id="chartHeatmap" class="overflow-auto"></div>
      </div>
    </div>

    <!-- Replenishment Table -->
    <div class="glass-panel p-4">
      <div class="section-title">補貨排程建議</div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>
              <th>品類</th><th>品名</th><th>建議補貨量</th>
              <th>建議補貨時間</th><th>預估成本</th><th>優先度</th>
            </tr>
          </thead>
          <tbody id="replenish-table-body">
            <tr><td colspan="6" class="text-center text-slate-500 py-4">載入資料後顯示</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div><!-- /tab-supply -->

  <!-- ==============================================================
       TAB 7: 品項剖析 Items
  ============================================================== -->
  <div id="tab-items" class="tab-content">
    <!-- Filters -->
    <div class="glass-panel p-3 mb-3">
      <div class="flex flex-wrap items-center gap-2">
        <div class="flex items-center gap-1">
          <i data-lucide="search" class="w-3.5 h-3.5 text-slate-400"></i>
          <input type="text" id="item-search" placeholder="搜尋品項名稱/ID…" class="w-44" />
        </div>
        <select id="item-cat-filter">
          <option value="">全部品類</option>
        </select>
        <select id="item-abc-filter">
          <option value="">全部ABC</option>
          <option value="A">A 類</option>
          <option value="B">B 類</option>
          <option value="C">C 類</option>
        </select>
        <select id="item-xyz-filter">
          <option value="">全部XYZ</option>
          <option value="X">X (穩定)</option>
          <option value="Y">Y (波動)</option>
          <option value="Z">Z (高波動)</option>
        </select>
        <select id="item-sort">
          <option value="value_desc">庫存金額 ↓</option>
          <option value="value_asc">庫存金額 ↑</option>
          <option value="aging_desc">庫齡 ↓</option>
          <option value="aging_asc">庫齡 ↑</option>
          <option value="gmroi_desc">GMROI ↓</option>
          <option value="str_desc">售磬率 ↓</option>
        </select>
        <button id="btn-export-items-csv" class="glass-btn text-emerald-400 ml-auto">
          <i data-lucide="file-down" class="w-3 h-3"></i>匯出CSV
        </button>
      </div>
    </div>

    <!-- Table -->
    <div class="glass-panel p-0 overflow-hidden mb-3">
      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>
              <th>品類</th><th>產品ID</th><th>名稱</th>
              <th>ABC</th><th>XYZ</th><th>ABCXYZ</th>
              <th class="text-right">庫存量</th>
              <th class="text-right">庫存金額(元)</th>
              <th class="text-right">庫齡(天)</th>
              <th>STR類型</th>
              <th class="text-right">售磬率</th>
              <th class="text-right">GMROI</th>
              <th>ROP狀態</th>
            </tr>
          </thead>
          <tbody id="items-table-body">
            <tr><td colspan="13" class="text-center text-slate-500 py-6">載入資料後顯示</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Pagination -->
    <div class="flex items-center justify-between text-xs text-slate-400">
      <div class="flex items-center gap-2">
        <span>每頁顯示</span>
        <select id="items-per-page" class="w-16">
          <option value="20">20</option>
          <option value="50" selected>50</option>
          <option value="100">100</option>
        </select>
        <span>筆</span>
      </div>
      <div id="page-info" class="text-slate-400">第 — / — 頁（共 — 筆）</div>
      <div class="flex items-center gap-1">
        <button id="btn-prev-page" class="glass-btn">
          <i data-lucide="chevron-left" class="w-3 h-3"></i>上一頁
        </button>
        <button id="btn-next-page" class="glass-btn">
          下一頁<i data-lucide="chevron-right" class="w-3 h-3"></i>
        </button>
      </div>
    </div>
  </div><!-- /tab-items -->

  <!-- ==============================================================
       TAB 8: 策略行動 Actions
  ============================================================== -->
  <div id="tab-actions" class="tab-content">
    <div class="grid grid-cols-3 gap-3">
      <!-- Alert List -->
      <div class="col-span-2 glass-panel p-4">
        <div class="section-title">策略行動清單</div>
        <div id="action-list" class="space-y-2">
          <div class="text-center text-slate-500 py-6">載入資料後顯示行動建議</div>
        </div>
      </div>

      <!-- Right Panel -->
      <div class="flex flex-col gap-3">
        <!-- Thresholds -->
        <div class="glass-panel p-4">
          <div class="section-title">閾值設定</div>
          <div class="space-y-3 text-sm">
            <div>
              <label class="text-slate-400 text-xs">DSI 上限 (天)</label>
              <input type="number" id="threshold-dsi" class="w-full mt-1" value="90" min="1" />
            </div>
            <div>
              <label class="text-slate-400 text-xs">庫齡占比上限 (%)</label>
              <input type="number" id="threshold-aging" class="w-full mt-1" value="20" min="0" max="100" />
            </div>
            <div>
              <label class="text-slate-400 text-xs">GMROI 下限</label>
              <input type="number" id="threshold-gmroi" class="w-full mt-1" value="1.2" step="0.1" min="0" />
            </div>
            <div>
              <label class="text-slate-400 text-xs">供需失衡指數上限</label>
              <input type="number" id="threshold-imbalance" class="w-full mt-1" value="0.3" step="0.05" min="0" max="1" />
            </div>
            <button id="btn-update-thresholds" class="glass-btn w-full justify-center">
              <i data-lucide="save" class="w-3 h-3"></i>更新閾值
            </button>
          </div>
        </div>

        <!-- KPI Scorecard -->
        <div class="glass-panel p-4 flex-1">
          <div class="section-title">KPI 記分卡</div>
          <div id="scorecard" class="space-y-3">
            <div>
              <div class="flex justify-between text-xs mb-1">
                <span class="text-slate-400">庫存健康</span>
                <span id="sc-health-val" class="text-white">—</span>
              </div>
              <div class="progress-bar"><div id="sc-health" class="progress-fill success" style="width:0%"></div></div>
            </div>
            <div>
              <div class="flex justify-between text-xs mb-1">
                <span class="text-slate-400">周轉效率</span>
                <span id="sc-turn-val" class="text-white">—</span>
              </div>
              <div class="progress-bar"><div id="sc-turn" class="progress-fill" style="width:0%"></div></div>
            </div>
            <div>
              <div class="flex justify-between text-xs mb-1">
                <span class="text-slate-400">庫齡控制</span>
                <span id="sc-aging-val" class="text-white">—</span>
              </div>
              <div class="progress-bar"><div id="sc-aging" class="progress-fill warning" style="width:0%"></div></div>
            </div>
            <div>
              <div class="flex justify-between text-xs mb-1">
                <span class="text-slate-400">預測準確度</span>
                <span id="sc-forecast-val" class="text-white">—</span>
              </div>
              <div class="progress-bar"><div id="sc-forecast" class="progress-fill success" style="width:0%"></div></div>
            </div>
            <div>
              <div class="flex justify-between text-xs mb-1">
                <span class="text-slate-400">供應鏈韌性</span>
                <span id="sc-supply-val" class="text-white">—</span>
              </div>
              <div class="progress-bar"><div id="sc-supply" class="progress-fill" style="width:0%"></div></div>
            </div>
            <div>
              <div class="flex justify-between text-xs mb-1">
                <span class="text-slate-400">獲利貢獻</span>
                <span id="sc-profit-val" class="text-white">—</span>
              </div>
              <div class="progress-bar"><div id="sc-profit" class="progress-fill success" style="width:0%"></div></div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div><!-- /tab-actions -->

  <!-- ==============================================================
       TAB 9: 情境模擬 What-If
  ============================================================== -->
  <div id="tab-whatif" class="tab-content">
    <div class="grid grid-cols-3 gap-3">
      <!-- Sliders -->
      <div class="glass-panel p-4">
        <div class="section-title">情境參數</div>
        <div class="space-y-5">
          <!-- Growth Rate -->
          <div>
            <div class="flex justify-between text-xs mb-1">
              <label class="text-slate-300">需求成長率</label>
              <span id="val-growth" class="text-blue-400 font-bold">+10%</span>
            </div>
            <input type="range" id="slider-growth" min="-30" max="50" value="10" step="1" />
            <div class="flex justify-between text-[0.6rem] text-slate-500 mt-0.5"><span>-30%</span><span>+50%</span></div>
          </div>
          <!-- Service Level -->
          <div>
            <div class="flex justify-between text-xs mb-1">
              <label class="text-slate-300">服務水準</label>
              <span id="val-svc" class="text-emerald-400 font-bold">95%</span>
            </div>
            <input type="range" id="slider-svc" min="85" max="99" value="95" step="1" />
            <div class="flex justify-between text-[0.6rem] text-slate-500 mt-0.5"><span>85%</span><span>99%</span></div>
          </div>
          <!-- Holding Cost -->
          <div>
            <div class="flex justify-between text-xs mb-1">
              <label class="text-slate-300">持有成本率</label>
              <span id="val-hold" class="text-amber-400 font-bold">25%</span>
            </div>
            <input type="range" id="slider-hold" min="15" max="35" value="25" step="1" />
            <div class="flex justify-between text-[0.6rem] text-slate-500 mt-0.5"><span>15%</span><span>35%</span></div>
          </div>
          <!-- Lead Time -->
          <div>
            <div class="flex justify-between text-xs mb-1">
              <label class="text-slate-300">前置時間調整</label>
              <span id="val-lead" class="text-purple-400 font-bold">±0天</span>
            </div>
            <input type="range" id="slider-lead" min="-7" max="14" value="0" step="1" />
            <div class="flex justify-between text-[0.6rem] text-slate-500 mt-0.5"><span>-7天</span><span>+14天</span></div>
          </div>
        </div>
      </div>

      <!-- What-If Table + Chart -->
      <div class="col-span-2 flex flex-col gap-3">
        <div class="glass-panel p-4">
          <div class="section-title">三情境比較</div>
          <div class="table-wrapper">
            <table id="whatif-table" class="data-table">
              <thead>
                <tr>
                  <th>指標</th>
                  <th class="text-center">悲觀情境</th>
                  <th class="text-center">基準情境</th>
                  <th class="text-center">樂觀情境</th>
                </tr>
              </thead>
              <tbody>
                <tr><td>庫存金額(萬)</td><td id="wi-inv-pes" class="text-center text-red-400">—</td><td id="wi-inv-base" class="text-center text-blue-400">—</td><td id="wi-inv-opt" class="text-center text-emerald-400">—</td></tr>
                <tr><td>DSI(天)</td><td id="wi-dsi-pes" class="text-center text-red-400">—</td><td id="wi-dsi-base" class="text-center text-blue-400">—</td><td id="wi-dsi-opt" class="text-center text-emerald-400">—</td></tr>
                <tr><td>安全庫存(萬)</td><td id="wi-ss-pes" class="text-center text-red-400">—</td><td id="wi-ss-base" class="text-center text-blue-400">—</td><td id="wi-ss-opt" class="text-center text-emerald-400">—</td></tr>
                <tr><td>持有成本(萬)</td><td id="wi-hc-pes" class="text-center text-red-400">—</td><td id="wi-hc-base" class="text-center text-blue-400">—</td><td id="wi-hc-opt" class="text-center text-emerald-400">—</td></tr>
                <tr><td>GMROI</td><td id="wi-gm-pes" class="text-center text-red-400">—</td><td id="wi-gm-base" class="text-center text-blue-400">—</td><td id="wi-gm-opt" class="text-center text-emerald-400">—</td></tr>
              </tbody>
            </table>
          </div>
        </div>
        <div class="glass-panel p-4 flex-1">
          <div class="section-title">情境影響圖</div>
          <div class="chart-container">
            <canvas id="chartWhatif"></canvas>
          </div>
        </div>
      </div>
    </div>
  </div><!-- /tab-whatif -->

  <!-- ==============================================================
       TAB 10: EOQ 成本 EOQ
  ============================================================== -->
  <div id="tab-eoq" class="tab-content">
    <div class="grid grid-cols-3 gap-3">
      <!-- EOQ Params -->
      <div class="glass-panel p-4">
        <div class="section-title">EOQ 參數</div>
        <div class="space-y-4 text-sm">
          <div>
            <label class="text-slate-400 text-xs">每次訂購成本 (ORDER_COST, 元)</label>
            <input type="number" id="eoq-order-cost" class="w-full mt-1" value="500" min="1" />
          </div>
          <div>
            <label class="text-slate-400 text-xs">年持有成本率 (HOLDING_RATE, %)</label>
            <input type="number" id="eoq-holding-rate" class="w-full mt-1" value="25" min="1" max="100" step="0.5" />
          </div>
          <button id="btn-calc-eoq" class="glass-btn w-full justify-center">
            <i data-lucide="calculator" class="w-3 h-3"></i>計算 EOQ
          </button>
          <div class="divider"></div>
          <div class="text-xs text-slate-500">
            <p>EOQ = √(2DS / H)</p>
            <p class="mt-1">D=年需求量 S=訂購成本</p>
            <p>H=單位持有成本/年</p>
          </div>
        </div>
      </div>

      <!-- EOQ Chart -->
      <div class="glass-panel p-4 col-span-2">
        <div class="section-title">總成本曲線 (EOQ 最優點)</div>
        <div class="chart-container-lg">
          <canvas id="chartEOQ"></canvas>
        </div>
      </div>
    </div>

    <!-- EOQ Table -->
    <div class="glass-panel p-4 mt-3">
      <div class="section-title">EOQ 建議清單</div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>
              <th>品名</th>
              <th class="text-right">年需求量</th>
              <th class="text-right">EOQ建議量</th>
              <th class="text-right">目前批量偏離</th>
              <th class="text-right">年持有成本(元)</th>
              <th class="text-right">年訂購成本(元)</th>
            </tr>
          </thead>
          <tbody id="eoq-table-body">
            <tr><td colspan="6" class="text-center text-slate-500 py-4">計算後顯示</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div><!-- /tab-eoq -->

  <!-- ==============================================================
       TAB 11: FIFO 批次 FIFO
  ============================================================== -->
  <div id="tab-fifo" class="tab-content">
    <div class="grid grid-cols-3 gap-3 mb-3">
      <!-- Product Selector + Stats -->
      <div class="glass-panel p-4">
        <div class="section-title">批次選擇</div>
        <div class="mb-3">
          <label class="text-xs text-slate-400">選擇品項</label>
          <select id="fifo-product-select" class="w-full mt-1">
            <option value="">載入資料後選擇…</option>
          </select>
        </div>
        <div class="divider"></div>
        <div class="section-title mt-2">FIFO 統計</div>
        <div class="space-y-2 text-sm">
          <div class="flex justify-between">
            <span class="text-slate-400">剩餘批次數</span>
            <span id="fifo-batch-count" class="text-blue-400">—</span>
          </div>
          <div class="flex justify-between">
            <span class="text-slate-400">加權平均庫齡</span>
            <span id="fifo-avg-aging" class="text-amber-400">— 天</span>
          </div>
          <div class="flex justify-between">
            <span class="text-slate-400">剩餘庫存總值</span>
            <span id="fifo-total-value" class="text-emerald-400">— 元</span>
          </div>
        </div>
      </div>

      <!-- FIFO Chart -->
      <div class="glass-panel p-4 col-span-2">
        <div class="section-title">FIFO 批次瀑布圖</div>
        <div class="chart-container">
          <canvas id="chartFIFO"></canvas>
        </div>
      </div>
    </div>

    <!-- Batch Detail Table -->
    <div class="glass-panel p-4">
      <div class="section-title">批次明細</div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>
              <th>批次號</th><th>入庫日期</th><th>入庫數量</th>
              <th class="text-right">已耗用</th><th class="text-right">剩餘數量</th>
              <th class="text-right">單位成本(元)</th><th class="text-right">批次金額(元)</th>
              <th class="text-right">庫齡(天)</th><th>狀態</th>
            </tr>
          </thead>
          <tbody id="fifo-table-body">
            <tr><td colspan="9" class="text-center text-slate-500 py-4">選擇品項後顯示</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div><!-- /tab-fifo -->

  <!-- ==============================================================
       TAB 12: 資料品質 Quality
  ============================================================== -->
  <div id="tab-quality" class="tab-content">
    <div class="grid grid-cols-3 gap-3 mb-3">
      <!-- Gauge -->
      <div class="glass-panel p-4 flex flex-col items-center justify-center">
        <div class="section-title w-full">資料品質總分</div>
        <div class="gauge-wrap my-4">
          <svg class="gauge-svg" viewBox="0 0 140 80">
            <path d="M 10 75 A 60 60 0 0 1 130 75" fill="none" stroke="rgba(255,255,255,0.08)" stroke-width="12" stroke-linecap="round"/>
            <path id="gauge-arc" d="M 10 75 A 60 60 0 0 1 130 75" fill="none" stroke="#10b981" stroke-width="12" stroke-linecap="round"
                  stroke-dasharray="188.5" stroke-dashoffset="188.5" style="transition:stroke-dashoffset 1s ease;"/>
          </svg>
          <div class="gauge-text text-emerald-400" id="quality-score">—</div>
        </div>
        <div class="text-xs text-slate-400">資料完整度與一致性</div>
      </div>

      <!-- Issue Summary -->
      <div class="glass-panel p-4">
        <div class="section-title">問題摘要</div>
        <div class="space-y-3 mt-2">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span class="w-2 h-2 rounded-full bg-red-500"></span>
              <span class="text-sm text-slate-300">嚴重錯誤</span>
            </div>
            <span id="quality-errors" class="text-red-400 font-bold text-lg">—</span>
          </div>
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span class="w-2 h-2 rounded-full bg-amber-400"></span>
              <span class="text-sm text-slate-300">警告</span>
            </div>
            <span id="quality-warnings" class="text-amber-400 font-bold text-lg">—</span>
          </div>
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span class="w-2 h-2 rounded-full bg-blue-400"></span>
              <span class="text-sm text-slate-300">注意事項</span>
            </div>
            <span id="quality-notices" class="text-blue-400 font-bold text-lg">—</span>
          </div>
        </div>
      </div>

      <!-- Quality Checklist -->
      <div class="glass-panel p-4">
        <div class="section-title">品質檢核項目</div>
        <div class="space-y-2 text-xs">
          <div class="flex items-center gap-2"><span id="qc-dates" class="text-slate-500">○</span><span>日期欄位完整性</span></div>
          <div class="flex items-center gap-2"><span id="qc-qty" class="text-slate-500">○</span><span>數量欄位非負值</span></div>
          <div class="flex items-center gap-2"><span id="qc-cost" class="text-slate-500">○</span><span>成本欄位合理性</span></div>
          <div class="flex items-center gap-2"><span id="qc-sku" class="text-slate-500">○</span><span>SKU唯一性</span></div>
          <div class="flex items-center gap-2"><span id="qc-balance" class="text-slate-500">○</span><span>進出庫平衡</span></div>
          <div class="flex items-center gap-2"><span id="qc-cat" class="text-slate-500">○</span><span>品類代碼一致性</span></div>
        </div>
      </div>
    </div>

    <!-- Issues Table -->
    <div class="glass-panel p-4">
      <div class="section-title">問題清單</div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>
              <th>嚴重度</th><th>品項</th><th>日期</th><th>問題說明</th>
            </tr>
          </thead>
          <tbody id="quality-table-body">
            <tr><td colspan="4" class="text-center text-slate-500 py-4">載入資料後顯示</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div><!-- /tab-quality -->

  <!-- ==============================================================
       TAB 13: 報告中心 Report
  ============================================================== -->
  <div id="tab-report" class="tab-content">
    <div class="grid grid-cols-2 gap-3 mb-3">
      <!-- PDF Export -->
      <div class="glass-panel p-4">
        <div class="section-title">PDF 報告匯出</div>
        <p class="text-xs text-slate-400 mb-3">將目前儀表板完整匯出為 PDF 報告，包含所有圖表與 KPI 數據。</p>
        <button id="btn-pdf" class="glass-btn text-red-400">
          <i data-lucide="file-type" class="w-3 h-3"></i>匯出 PDF 報告
        </button>
      </div>

      <!-- Excel Export -->
      <div class="glass-panel p-4">
        <div class="section-title">Excel 資料匯出</div>
        <p class="text-xs text-slate-400 mb-3">將所有分析結果匯出為多工作表 Excel 檔案，包含 ABC、庫齡、ROP 等。</p>
        <button id="btn-excel" class="glass-btn text-emerald-400">
          <i data-lucide="table-2" class="w-3 h-3"></i>匯出 Excel
        </button>
      </div>
    </div>

    <!-- Text Summary -->
    <div class="glass-panel p-4 mb-3">
      <div class="flex items-center justify-between mb-2">
        <div class="section-title">文字摘要報告</div>
        <button id="btn-copy-report" class="glass-btn">
          <i data-lucide="copy" class="w-3 h-3"></i>複製
        </button>
      </div>
      <div id="report-text" class="text-xs text-slate-300 leading-relaxed bg-black/20 rounded-lg p-3 min-h-[6rem] whitespace-pre-wrap">
        載入資料後自動生成摘要報告…
      </div>
    </div>

    <!-- Snapshots -->
    <div class="glass-panel p-4">
      <div class="flex items-center justify-between mb-3">
        <div class="section-title">歷史快照</div>
        <button id="btn-save-snapshot" class="glass-btn text-indigo-400">
          <i data-lucide="camera" class="w-3 h-3"></i>儲存快照
        </button>
      </div>
      <div id="snapshots-list" class="space-y-2 text-xs text-slate-400">
        <div class="text-center py-3">尚無歷史快照</div>
      </div>
    </div>
  </div><!-- /tab-report -->

  <!-- ==============================================================
       TAB 14: 設定DAX Settings
  ============================================================== -->
  <div id="tab-settings" class="tab-content">
    <div class="grid grid-cols-2 gap-3 mb-3">
      <!-- Upload Zone -->
      <div class="glass-panel p-4">
        <div class="section-title">資料上傳</div>
        <div id="drop-zone" class="mb-3">
          <i data-lucide="upload-cloud" class="w-8 h-8 text-indigo-400 mx-auto mb-2"></i>
          <p class="text-sm text-slate-300">拖曳 Excel / CSV 至此，或點擊選擇檔案</p>
          <p class="text-xs text-slate-500 mt-1">支援 .xlsx、.xls、.csv</p>
        </div>
        <div class="flex flex-wrap gap-2">
          <button id="btn-load-demo" class="glass-btn text-blue-400">
            <i data-lucide="database" class="w-3 h-3"></i>載入範例資料
          </button>
          <button id="btn-export-json" class="glass-btn text-amber-400">
            <i data-lucide="braces" class="w-3 h-3"></i>匯出JSON
          </button>
          <button id="btn-export-csv" class="glass-btn text-emerald-400">
            <i data-lucide="file-down" class="w-3 h-3"></i>匯出CSV
          </button>
          <button id="btn-reset" class="glass-btn text-red-400">
            <i data-lucide="trash-2" class="w-3 h-3"></i>重置
          </button>
        </div>
      </div>

      <!-- Display Settings -->
      <div class="glass-panel p-4">
        <div class="section-title">顯示設定</div>
        <div class="space-y-3">
          <div class="flex items-center justify-between">
            <span class="text-sm text-slate-300">深色 / 淺色模式</span>
            <button id="btn-darkmode-settings" class="glass-btn">
              <i data-lucide="moon" class="w-3 h-3"></i>切換模式
            </button>
          </div>
          <div class="divider"></div>
          <p class="text-xs text-slate-500">
            系統版本：Inventory Strategic OS v3.0<br/>
            建置日期：2026-03-14<br/>
            支援格式：xlsx / xls / csv<br/>
            圖表引擎：Chart.js 4.4.0
          </p>
        </div>
      </div>
    </div>

    <!-- DAX Formula Reference -->
    <div class="glass-panel p-4">
      <div class="section-title mb-3">DAX 公式庫（Power BI 參考）</div>
      <div class="space-y-1" id="dax-accordion">

        <!-- 1 日期表 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>1. 日期表 (Date Table)</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>DateTable =
ADDCOLUMNS(
    CALENDAR(DATE(2020,1,1), DATE(2026,12,31)),
    "Year",       YEAR([Date]),
    "Month",      MONTH([Date]),
    "MonthName",  FORMAT([Date], "MMM"),
    "Quarter",    "Q" & ROUNDUP(MONTH([Date])/3, 0),
    "YearMonth",  FORMAT([Date], "YYYY-MM"),
    "WeekNo",     WEEKNUM([Date])
)</pre>
          </div>
        </div>

        <!-- 2 累計入庫 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>2. 累計入庫數量</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>累計入庫數量 =
CALCULATE(
    SUM(Inventory[InQty]),
    FILTER(
        ALL(DateTable),
        DateTable[Date] <= MAX(DateTable[Date])
    )
)</pre>
          </div>
        </div>

        <!-- 3 累計出庫 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>3. 累計出庫數量</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>累計出庫數量 =
CALCULATE(
    SUM(Inventory[OutQty]),
    FILTER(
        ALL(DateTable),
        DateTable[Date] <= MAX(DateTable[Date])
    )
)</pre>
          </div>
        </div>

        <!-- 4 庫存數量 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>4. 庫存數量</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>庫存數量 = [累計入庫數量] - [累計出庫數量]</pre>
          </div>
        </div>

        <!-- 5 庫存金額 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>5. 庫存金額</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>庫存金額 =
SUMX(
    Inventory,
    Inventory[StockQty] * Inventory[UnitCost]
)</pre>
          </div>
        </div>

        <!-- 6 期初庫存 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>6. 期初庫存金額</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>期初庫存金額 =
CALCULATE(
    [庫存金額],
    STARTOFMONTH(DateTable[Date])
)</pre>
          </div>
        </div>

        <!-- 7 平均庫存 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>7. 平均庫存金額</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>平均庫存金額 =
AVERAGEX(
    VALUES(DateTable[YearMonth]),
    CALCULATE([庫存金額])
)</pre>
          </div>
        </div>

        <!-- 8 加權平均成本 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>8. 加權平均成本 (WAC)</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>加權平均成本 =
DIVIDE(
    SUMX(Inventory, Inventory[InQty] * Inventory[InUnitCost]),
    SUM(Inventory[InQty]),
    0
)</pre>
          </div>
        </div>

        <!-- 9 存貨周轉率 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>9. 存貨周轉率</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>存貨周轉率 =
DIVIDE(
    [年度銷售成本],
    [平均庫存金額],
    0
)</pre>
          </div>
        </div>

        <!-- 10 DSI -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>10. 庫存天數 DSI</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>DSI =
DIVIDE(
    [平均庫存金額],
    DIVIDE([年度銷售成本], 365, 0),
    0
)
-- 或簡化：DSI = 365 / [存貨周轉率]</pre>
          </div>
        </div>

        <!-- 11 售磬率 快銷 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>11. 售磬率（快銷 &lt;30天）</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>售磬率_快銷 =
DIVIDE(
    CALCULATE(SUM(Sales[Qty]), Sales[LeadDays] < 30),
    SUM(Inventory[AvgStockQty]),
    0
)</pre>
          </div>
        </div>

        <!-- 12 售磬率 慢銷 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>12. 售磬率（慢銷 ≥30天）</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>售磬率_慢銷 =
DIVIDE(
    CALCULATE(SUM(Sales[Qty]), Sales[LeadDays] >= 30),
    SUM(Inventory[AvgStockQty]),
    0
)</pre>
          </div>
        </div>

        <!-- 13 存銷比 長週期 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>13. 存銷比 STR（長週期，12週）</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>存銷比_長週期 =
DIVIDE(
    [庫存數量],
    CALCULATE(
        AVERAGE(Sales[WeeklySales]),
        DATESINPERIOD(DateTable[Date], MAX(DateTable[Date]), -12, WEEK)
    ),
    0
)</pre>
          </div>
        </div>

        <!-- 14 存銷比 短週期 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>14. 存銷比 STR（短週期，4週）</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>存銷比_短週期 =
DIVIDE(
    [庫存數量],
    CALCULATE(
        AVERAGE(Sales[WeeklySales]),
        DATESINPERIOD(DateTable[Date], MAX(DateTable[Date]), -4, WEEK)
    ),
    0
)</pre>
          </div>
        </div>

        <!-- 15 ABC 累計比例 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>15. ABC 累計金額比例</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>ABC_累計比例 =
VAR TotalValue = CALCULATE(SUM(Inventory[StockValue]), ALL(Inventory))
VAR CumValue =
    CALCULATE(
        SUM(Inventory[StockValue]),
        FILTER(
            ALL(Inventory),
            Inventory[StockValue] >= MIN(Inventory[StockValue])
        )
    )
RETURN DIVIDE(CumValue, TotalValue, 0)</pre>
          </div>
        </div>

        <!-- 16 ABC 分類 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>16. ABC 分類</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>ABC分類 =
SWITCH(
    TRUE(),
    [ABC_累計比例] <= 0.80, "A",
    [ABC_累計比例] <= 0.95, "B",
    "C"
)</pre>
          </div>
        </div>

        <!-- 17 供需失衡指數 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>17. 供需失衡指數</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>供需失衡指數 =
VAR InQty  = SUM(Inventory[InQty])
VAR OutQty = SUM(Inventory[OutQty])
VAR Total  = InQty + OutQty
RETURN
    IF(Total = 0, 0,
        ABS(DIVIDE(InQty - OutQty, Total, 0))
    )</pre>
          </div>
        </div>

        <!-- 18 FIFO 加權平均庫齡 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>18. FIFO 加權平均庫齡</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>FIFO加權平均庫齡 =
DIVIDE(
    SUMX(
        FIFOBatch,
        FIFOBatch[RemainQty] *
        DATEDIFF(FIFOBatch[InDate], TODAY(), DAY)
    ),
    SUM(FIFOBatch[RemainQty]),
    0
)</pre>
          </div>
        </div>

        <!-- 19 背離信號 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>19. 背離信號</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>背離信號 =
VAR DSI_now  = [DSI]
VAR DSI_3m   = CALCULATE([DSI], DATESINPERIOD(DateTable[Date], MAX(DateTable[Date]), -90, DAY))
VAR Sales_now = SUM(Sales[Qty])
VAR Sales_3m  = CALCULATE(SUM(Sales[Qty]), DATESINPERIOD(DateTable[Date], MAX(DateTable[Date]), -90, DAY))
RETURN
    IF(DSI_now > DSI_3m * 1.2 && Sales_now < Sales_3m * 0.9,
        "⚠ 高庫存低銷量 背離",
    IF(DSI_now < DSI_3m * 0.8 && Sales_now > Sales_3m * 1.1,
        "✓ 低庫存高銷量 健康",
        "— 正常範圍"
    ))</pre>
          </div>
        </div>

        <!-- 20 安全庫存 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>20. 安全庫存 (Safety Stock)</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>-- Z 因子：90%=1.28，95%=1.645，99%=2.326
安全庫存 =
VAR Z = 1.645  -- 95% 服務水準
VAR σ_demand  = STDEV.P(Sales[DailySales])
VAR σ_lead    = STDEV.P(Supplier[LeadDays])
VAR avg_lead  = AVERAGE(Supplier[LeadDays])
VAR avg_daily = AVERAGE(Sales[DailySales])
RETURN
    Z * SQRT(avg_lead * σ_demand^2 + avg_daily^2 * σ_lead^2)</pre>
          </div>
        </div>

        <!-- 21 ROP -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>21. 再訂購點 ROP</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>ROP =
    AVERAGE(Sales[DailySales]) * AVERAGE(Supplier[LeadDays])
    + [安全庫存]</pre>
          </div>
        </div>

        <!-- 22 EOQ -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>22. 經濟訂購量 EOQ</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>EOQ =
VAR D = SUM(Sales[AnnualQty])      -- 年需求量
VAR S = 500                         -- 每次訂購成本（元）
VAR H = [加權平均成本] * 0.25       -- 年單位持有成本
RETURN SQRT(DIVIDE(2 * D * S, H, 0))</pre>
          </div>
        </div>

        <!-- 23 Fisher 比值 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>23. Fisher 週期性比值</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>-- Fisher 週期性檢定：判斷銷售是否具顯著季節性
-- 計算需在 Python/Excel 中完成 DFT，再匯入 Power BI
Fisher比值 =
DIVIDE(
    MAX(Fourier[MaxAmplitude]),
    AVERAGE(Fourier[AllAmplitudes]),
    0
)
-- 臨界值（n=12月）約 6.5，>6.5 表示具顯著季節性</pre>
          </div>
        </div>

        <!-- 24 季節指數 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>24. 季節指數 (Seasonal Index)</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>季節指數 =
VAR AvgMonthly = AVERAGEX(VALUES(DateTable[YearMonth]), CALCULATE(SUM(Sales[Qty])))
VAR ThisMonth  = CALCULATE(SUM(Sales[Qty]))
RETURN DIVIDE(ThisMonth, AvgMonthly, 1)</pre>
          </div>
        </div>

        <!-- 25 複合預測模型 -->
        <div class="accordion-item">
          <div class="accordion-header" onclick="toggleAccordion(this)">
            <span>25. 複合預測模型 (3M MA + 趨勢 + 季節)</span><i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>
          </div>
          <div class="accordion-body">
            <pre>移動平均3M =
AVERAGEX(
    DATESINPERIOD(DateTable[Date], MAX(DateTable[Date]), -3, MONTH),
    CALCULATE(SUM(Sales[Qty]))
)

移動平均6M =
AVERAGEX(
    DATESINPERIOD(DateTable[Date], MAX(DateTable[Date]), -6, MONTH),
    CALCULATE(SUM(Sales[Qty]))
)

複合預測 =
VAR ma3  = [移動平均3M]
VAR ma6  = [移動平均6M]
VAR si   = [季節指數]
VAR base = (ma3 * 0.6 + ma6 * 0.4)  -- 加權組合
RETURN base * si                      -- 乘以季節指數調整</pre>
          </div>
        </div>

      </div><!-- /dax-accordion -->
    </div>

  </div><!-- /tab-settings -->

</main>

<script>
"""

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(html)

print(f"✓ 已輸出至 {OUTPUT_FILE}")
print(f"  檔案大小：{len(html.encode('utf-8')) / 1024:.1f} KB")
