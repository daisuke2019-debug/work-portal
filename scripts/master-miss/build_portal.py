# -*- coding: utf-8 -*-
import json, re, os

# 読み書きの場所はこのスクリプトの位置から決める（どのフォルダで実行しても同じ結果になる）
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_JSON = os.path.join(ROOT, 'data', 'portal_json_data.json')
OUT_HTML = os.path.join(ROOT, 'site', 'master-miss', 'index.html')

with open(DATA_JSON, 'r', encoding='utf-8') as f:
    json_data = json.load(f)

staff_data = json_data['staff_data']
col_ranking = json_data['col_ranking']

js_staff_data = json.dumps(staff_data, ensure_ascii=False)
js_col_ranking = json.dumps(col_ranking, ensure_ascii=False)

# Build html by replacing placeholders
with open(os.path.abspath(__file__), 'r', encoding='utf-8') as f:
    orig_code = f.read()

# Generate index.html safely
html_content = """<!DOCTYPE html>
<html lang="ja" class="scroll-smooth">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>店舗マスタチェック指摘カウントWebポータル | AI Company Executive</title>
    <!-- Tailwind CSS v3 Play CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&family=Noto+Sans+JP:wght@400;500;700;900&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
    <style>
        body {
            font-family: 'Inter', 'Noto Sans JP', -apple-system, BlinkMacSystemFont, sans-serif;
            background-color: #f8fafc;
            color: #0f172a;
            -webkit-font-smoothing: antialiased;
        }
        
        .header-gradient {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
            border-bottom: 2px solid #334155;
        }
        
        .btn-action-dark {
            background-color: #0f172a !important;
            color: #ffffff !important;
            border: 1px solid #1e293b !important;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2) !important;
        }
        .btn-action-dark:hover {
            background-color: #000000 !important;
            color: #ffffff !important;
        }

        .btn-action-red {
            background-color: #dc2626 !important;
            color: #ffffff !important;
            border: 1px solid #b91c1c !important;
            box-shadow: 0 2px 5px rgba(220,38,38,0.3) !important;
        }
        .btn-action-red:hover {
            background-color: #b91c1c !important;
            color: #ffffff !important;
        }

        .btn-action-green {
            background-color: #059669 !important;
            color: #ffffff !important;
            border: 1px solid #047857 !important;
            box-shadow: 0 2px 5px rgba(5,150,105,0.3) !important;
        }
        .btn-action-green:hover {
            background-color: #047857 !important;
            color: #ffffff !important;
        }

        .glass-card {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            box-shadow: 0 4px 20px -2px rgba(15, 23, 42, 0.06);
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .kpi-card {
            background: #ffffff;
            border: 2px solid #cbd5e1;
            border-radius: 1rem;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .kpi-card:hover {
            transform: translateY(-2px);
        }

        .drop-zone {
            border: 2px dashed #ef4444;
            background: linear-gradient(180deg, #fef2f2 0%, #ffffff 100%);
            transition: all 0.2s ease;
        }
        .drop-zone:hover, .drop-zone.dragover {
            background-color: #fee2e2;
            border-color: #dc2626;
        }

        @media print {
            body {
                background-color: #ffffff !important;
                color: #000000 !important;
                padding-bottom: 0 !important;
            }
            header, #sidebar, #sidebarOverlay, #dropZone, #searchInput, .print-hide {
                display: none !important;
            }
            .glass-card, .kpi-card {
                box-shadow: none !important;
                border: 1px solid #94a3b8 !important;
                break-inside: avoid;
            }
            #detailModal {
                position: static !important;
                display: block !important;
                background: transparent !important;
            }
        }

        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
            background: #94a3b8;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #64748b;
        }
    </style>
</head>
<body class="min-h-screen pb-20 selection:bg-red-500 selection:text-white">

    <!-- サイドバー (ハンバーガーメニュー) -->
    <div id="sidebarOverlay" onclick="toggleSidebar()" class="fixed inset-0 bg-slate-900/70 backdrop-blur-sm z-50 hidden transition-opacity duration-300"></div>
    <aside id="sidebar" class="fixed top-0 left-0 w-80 h-full overflow-y-auto overscroll-contain bg-white z-50 transform -translate-x-full shadow-2xl flex flex-col justify-between p-6 border-r-2 border-slate-300 transition-transform duration-300 ease-out">
        <div>
            <div class="flex items-center justify-between pb-5 border-b-2 border-slate-200">
                <div class="flex items-center gap-2.5">
                    <div class="w-9 h-9 rounded-xl bg-red-600 flex items-center justify-center text-white font-black text-lg shadow-md">
                        AI
                    </div>
                    <div>
                        <h2 class="font-black text-base text-slate-900 leading-tight">AI Company</h2>
                        <p class="text-xs text-slate-800 font-bold">Executive Portal</p>
                    </div>
                </div>
                <button onclick="toggleSidebar()" class="w-8 h-8 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 flex items-center justify-center text-xl font-black">
                    ✕
                </button>
            </div>

            <nav class="mt-6 space-y-2">
                <div class="px-3 py-2 text-xs font-black text-slate-700 uppercase tracking-wider">ビュー切替・日別データ選択</div>
                
                <button onclick="setDateFilter('all'); toggleSidebar();" class="w-full text-left p-3 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-900 font-black text-sm flex items-center justify-between transition-all group">
                    <div class="flex items-center gap-2.5">
                        <span class="text-lg">🌐</span>
                        <span>全期間 累計データ</span>
                    </div>
                </button>

                <div id="sidebarDateList" class="space-y-1.5 pt-2">
                    <!-- JavaScriptで動的生成 -->
                </div>
            </nav>
        </div>

        <div class="pt-4 border-t-2 border-slate-200">
            <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <p class="text-xs text-slate-800 font-black">🏢 BOARD役員統括ポータル</p>
                <p class="text-xs text-slate-700 font-bold mt-0.5">Dai CEO爆速意思決定システム v3.5</p>
            </div>
        </div>
    </aside>

    <!-- ヘッダー -->
    <header class="header-gradient text-white sticky top-0 z-40 shadow-xl">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 md:px-8 py-3.5 flex items-center justify-between gap-2">
            
            <div class="flex items-center gap-3">
                <button onclick="toggleSidebar()" class="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-white flex items-center justify-center text-xl font-black shadow-md border border-slate-600 transition-all active:scale-95">
                    ☰
                </button>
                <div class="flex items-center gap-2.5">
                    <div class="w-9 h-9 sm:w-10 sm:h-10 rounded-xl bg-red-600 flex items-center justify-center text-white font-black text-base sm:text-lg shadow-md shrink-0 border border-red-500">
                        AI
                    </div>
                    <div>
                        <div class="flex items-center gap-2">
                            <h1 class="text-base sm:text-xl font-black tracking-tight text-white">店舗マスタチェック指摘カウント</h1>
                            <span class="bg-red-600 text-white text-xs font-black px-2 py-0.5 rounded border border-red-400 shrink-0">10月度</span>
                        </div>
                        <p class="text-xs text-slate-300 hidden sm:block font-bold">全スタッフの店舗巡回・マスタミス指摘結果の一覧・日別比較分析</p>
                    </div>
                </div>
            </div>

            <div class="flex items-center gap-2 shrink-0">
                <button onclick="window.print()" class="btn-action-dark text-xs sm:text-sm font-black px-3.5 py-2 rounded-xl flex items-center gap-1.5 shadow-md active:scale-95 transition-all print-hide">
                    <span>🖨️</span> <span class="hidden sm:inline">印刷 / PDF保存</span>
                </button>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-3 sm:px-6 md:px-8 mt-6 space-y-6">

        
        <!-- 🚨 Feature 2: 複数日連続赤指摘店舗の重点警告アラート -->
        
            <div id="consecutiveShopChips" class="flex flex-wrap items-center gap-1.5 w-full md:w-auto">
                <!-- JS描画 -->
            </div>
        </div>


        <!-- 📅 日付別切り替えタブバー (最重要・各日付独立保持＆ワンタップ切替) -->
        <div class="glass-card p-4 sm:p-5 rounded-2xl border-2 border-slate-300 bg-white shadow-sm flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div class="flex items-center gap-3">
                <div class="w-10 h-10 rounded-xl bg-red-100 text-red-600 flex items-center justify-center text-xl font-black shadow-inner border border-red-200">
                    📅
                </div>
                <div>
                    <h3 class="text-base sm:text-lg font-black text-slate-900 flex items-center gap-2">
                        <span>対象日付の切替</span>
                        <span id="currentDateDisplayBadge" class="bg-slate-900 text-white text-xs px-2.5 py-0.5 rounded-full font-black">全期間 (累計)</span>
                    </h3>
                    <p class="text-xs sm:text-sm text-slate-700 font-bold mt-0.5">タップすると全KPI・指摘ランキング・テーブル・詳細がその日付のデータに即座に切り替わります</p>
                </div>
            </div>
            
            <div id="dateFilterButtons" class="flex flex-wrap items-center gap-2 w-full md:w-auto">
                <!-- JavaScriptで動的生成される日付ボタン群 -->
            </div>
        </div>

        <!-- ドラッグ＆ドロップエリア (透明オーバーレイ式・全環境100%確実起動) -->
        <section id="dropZone" class="drop-zone p-6 sm:p-8 rounded-2xl text-center relative shadow-sm border-2 border-dashed border-red-400 hover:border-red-600 transition-all overflow-hidden">
            <input type="file" id="fileInput" accept=".xlsx, .xls" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-30" onchange="handleFileSelect(event)">
            <div class="flex flex-col items-center justify-center gap-3 relative z-10 pointer-events-none">
                <div class="w-14 h-14 rounded-2xl bg-red-100 text-red-600 flex items-center justify-center text-2xl font-black shadow-inner border border-red-200">
                    📥
                </div>
                <div>
                    <h2 class="text-base sm:text-xl font-black text-slate-900">毎朝のマスタチェックExcelファイル（.xlsx）を流し込んで自動統合</h2>
                    <p id="currentStatusText" class="text-xs sm:text-sm text-slate-800 mt-1 font-bold">ここをタップ または ファイルをドラッグ＆ドロップで即座にデータ追加</p>
                </div>
                <div class="btn-action-red font-black text-sm sm:text-base px-8 py-3 rounded-xl shadow-md mt-1">
                    📁 Excelファイルを選択してデータ追加
                </div>
            </div>
        </section>

        <!-- KPIカード -->
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div class="kpi-card p-4 sm:p-5 shadow-sm">
                <div class="text-slate-800 text-xs sm:text-sm font-black">対象チェック店舗数</div>
                <div id="statTotalShops" class="text-3xl sm:text-4xl font-black text-slate-900 mt-1 tracking-tight">0 <span class="text-sm font-bold text-slate-700">店舗</span></div>
                <div id="statStaffCountSub" class="text-xs text-slate-800 mt-1 font-bold">13名のスタッフ対象</div>
            </div>

            <div class="kpi-card p-4 sm:p-5 border-2 border-red-400 bg-red-50/60 shadow-sm">
                <div class="text-red-950 text-xs sm:text-sm font-black">A列赤指摘店舗数 (マスタミス)</div>
                <div id="statARedCount" class="text-3xl sm:text-4xl font-black text-red-600 mt-1 tracking-tight">0 <span class="text-sm font-bold text-slate-800">店舗</span></div>
                <div id="statTopMissStaff" class="text-xs text-red-950 font-black mt-1 truncate">最多: --</div>
            </div>

            <div class="kpi-card p-4 sm:p-5 border-2 border-amber-400 bg-amber-50/60 shadow-sm">
                <div class="text-amber-950 text-xs sm:text-sm font-black">A列外 赤指摘項目</div>
                <div id="statNonARedCount" class="text-3xl sm:text-4xl font-black text-amber-800 mt-1 tracking-tight">0 <span class="text-sm font-bold text-slate-800">箇所</span></div>
                <div class="text-xs text-amber-950 mt-1 font-black">赤セル項目総数</div>
            </div>

            <div class="kpi-card p-4 sm:p-5 border-2 border-emerald-400 bg-emerald-50/60 shadow-sm">
                <div class="text-emerald-950 text-xs sm:text-sm font-black">ミス0件の良好スタッフ</div>
                <div id="statPerfectCount" class="text-3xl sm:text-4xl font-black text-emerald-800 mt-1 tracking-tight">0 <span class="text-sm font-bold text-slate-800">名</span></div>
                <div id="statPerfectRate" class="text-xs text-emerald-950 font-black mt-1">0% がミスなし</div>
            </div>
        </div>

        <!-- メインテーブル -->
        <section id="staffTableSection" class="glass-card rounded-2xl overflow-hidden border-2 border-slate-300">
            <div class="p-4 sm:p-6 border-b-2 border-slate-300 space-y-4 bg-slate-100">
                <div class="flex flex-col md:flex-row justify-between md:items-center gap-4">
                    <div>
                        <h2 id="tableHeading" class="text-lg sm:text-xl font-black text-slate-900">
                            👔 最終更新スタッフ別 マスタミス指摘数 (一番ミスが多い順)
                        </h2>
                        <p id="tableSubHeading" class="text-xs sm:text-sm text-slate-800 mt-0.5 font-bold">A列の赤店舗マーカーをカウント。ミス率が高い順に整列中</p>
                    </div>

                    <div class="flex items-center gap-2.5 w-full md:w-auto">
                        <input type="text" id="searchInput" oninput="filterStaffList()" placeholder="スタッフ名で絞り込み..." class="w-full md:w-60 px-4 py-2.5 text-sm bg-white border-2 border-slate-400 rounded-xl font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-red-600 shadow-sm">
                        <button onclick="exportToCSV()" class="btn-action-dark text-xs sm:text-sm font-black px-4 py-2.5 rounded-xl shadow-md shrink-0 cursor-pointer active:scale-95 transition-all print-hide">
                            📥 CSV出力
                        </button>
                    </div>
                </div>

                <div class="flex flex-wrap items-center gap-2 pt-2 border-t border-slate-300 print-hide">
                    <span class="text-xs font-black text-slate-700 mr-1">状態絞り込み:</span>
                    <button id="filterBtnAll" onclick="setQuickFilter('all')" class="px-3.5 py-1.5 rounded-xl bg-slate-900 text-white font-black text-xs shadow-sm transition-all">
                        すべて表示
                    </button>
                    <button id="filterBtnMiss" onclick="setQuickFilter('miss')" class="px-3.5 py-1.5 rounded-xl bg-white border-2 border-red-500 text-red-700 hover:bg-red-50 font-black text-xs shadow-sm transition-all">
                        🔴 ミスありスタッフのみ
                    </button>
                    <button id="filterBtnPerfect" onclick="setQuickFilter('perfect')" class="px-3.5 py-1.5 rounded-xl bg-white border-2 border-emerald-500 text-emerald-800 hover:bg-emerald-50 font-black text-xs shadow-sm transition-all">
                        ✨ ミス0件(完璧)のみ
                    </button>
                    <button id="filterBtnNonA" onclick="setQuickFilter('nona')" class="px-3.5 py-1.5 rounded-xl bg-white border-2 border-amber-500 text-amber-900 hover:bg-amber-50 font-black text-xs shadow-sm transition-all">
                        ⚠️ A列外赤セルありのみ
                    </button>
                </div>
            </div>

            <!-- PC用テーブルビュー -->
            <div class="hidden md:block overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-200/80 text-slate-900 text-xs sm:text-sm font-black border-b-2 border-slate-300">
                            <th class="py-4 px-6">順位 / スタッフ名</th>
                            <th class="py-4 px-4 text-center">巡回・チェック店舗数</th>
                            <th class="py-4 px-4 text-center bg-red-100/60 text-red-950">A列 赤指摘店舗 (ミス)</th>
                            <th class="py-4 px-4 text-center">マスタミス率</th>
                            <th class="py-4 px-4 text-center bg-amber-100/60 text-amber-950">A列外 赤セル箇所</th>
                            <th class="py-4 px-6 text-center">アクション / 指導文面</th>
                        </tr>
                    </thead>
                    <tbody id="staffTableBody" class="divide-y divide-slate-200 text-sm font-medium">
                        <!-- JS描画 -->
                    </tbody>
                </table>
            </div>

            <!-- スマホ用カードビュー -->
            <div id="staffCardList" class="md:hidden divide-y divide-slate-200">
                <!-- JS描画 -->
            </div>
        </section>

        <!-- ワースト指摘項目ランキング -->
        <section class="glass-card p-5 sm:p-6 rounded-2xl border-2 border-slate-300 space-y-4">
            <div class="flex items-center justify-between">
                <div>
                    <h2 class="text-base sm:text-lg font-black text-slate-900 flex items-center gap-2">
                        <span>⚠️ A列外 赤セル指摘項目 ワーストランキング</span>
                    </h2>
                    <p class="text-xs sm:text-sm text-slate-800 font-bold mt-0.5">どの項目で指摘が多く発生しているか一覧可視化</p>
                </div>
            </div>
            <div id="columnRankingContainer" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                <!-- JS描画 -->
            </div>
        </section>
    </main>

    <!-- モーダルダイアログ (詳細店舗一覧 & 指導チャットコピー) -->
    <div id="detailModal" class="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-50 hidden flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
        <div class="bg-white rounded-3xl max-w-3xl w-full border-2 border-slate-400 shadow-2xl overflow-hidden my-8 transform transition-all">
            <div class="header-gradient text-white p-5 sm:p-6 flex items-center justify-between border-b-2 border-slate-700">
                <div class="flex items-center gap-3">
                    <div class="w-11 h-11 rounded-2xl bg-red-600 flex items-center justify-center text-white font-black text-xl shadow-md border border-red-400">
                        👤
                    </div>
                    <div>
                        <h3 id="modalStaffName" class="text-lg sm:text-2xl font-black tracking-tight">スタッフ名</h3>
                        <p id="modalStaffSubText" class="text-xs sm:text-sm text-slate-300 font-bold mt-0.5">巡回店舗の指摘詳細一覧</p>
                    </div>
                </div>
                <button onclick="closeModal()" class="w-9 h-9 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center text-2xl font-black transition-all">
                    ✕
                </button>
            </div>

            <div class="p-5 sm:p-6 space-y-6 max-h-[75vh] overflow-y-auto">
                <div class="grid grid-cols-3 gap-3">
                    <div class="p-3.5 rounded-2xl bg-slate-100 border border-slate-300 text-center">
                        <div class="text-xs text-slate-800 font-black">巡回チェック店舗</div>
                        <div id="modalShopCount" class="text-xl sm:text-2xl font-black text-slate-900 mt-1">0 <span class="text-xs text-slate-700">店舗</span></div>
                    </div>
                    <div class="p-3.5 rounded-2xl bg-red-100/80 border border-red-300 text-center">
                        <div class="text-xs text-red-950 font-black">A列 赤指摘 (ミス)</div>
                        <div id="modalARedCount" class="text-xl sm:text-2xl font-black text-red-700 mt-1">0 <span class="text-xs text-red-900">店舗</span></div>
                    </div>
                    <div class="p-3.5 rounded-2xl bg-amber-100/80 border border-amber-300 text-center">
                        <div class="text-xs text-amber-950 font-black">A列外 赤セル箇所</div>
                        <div id="modalNonARedCount" class="text-xl sm:text-2xl font-black text-amber-900 mt-1">0 <span class="text-xs text-amber-950">箇所</span></div>
                    </div>
                </div>

                <div class="flex flex-col sm:flex-row items-center justify-between gap-3 p-4 bg-emerald-50 rounded-2xl border-2 border-emerald-300">
                    <div>
                        <div class="text-xs font-black text-emerald-950">💬 指導・確認用チャット文面自動生成</div>
                        <div class="text-xs text-emerald-900 font-bold mt-0.5">ワンタップでLINEやチャットツールにそのまま貼れる指示文をコピーできます</div>
                    </div>
                    <button onclick="copyStaffChatModal()" class="btn-action-green font-black text-xs sm:text-sm px-5 py-2.5 rounded-xl shadow-md shrink-0 active:scale-95 transition-all">
                        📋 指導用チャット文面をコピー
                    </button>
                </div>

                <div>
                    <h4 class="text-sm font-black text-slate-900 mb-3 flex items-center justify-between">
                        <span>🏬 巡回店舗一覧 & 各セル指摘詳細</span>
                        <span id="modalShopBadge" class="text-xs bg-slate-200 text-slate-800 px-2.5 py-0.5 rounded-full font-bold">0店舗</span>
                    </h4>
                    <div id="modalShopList" class="space-y-2.5">
                        <!-- JS描画 -->
                    </div>
                </div>
            </div>

            <div class="p-4 bg-slate-100 border-t-2 border-slate-300 flex justify-end">
                <button onclick="closeModal()" class="px-6 py-2.5 bg-slate-800 hover:bg-slate-900 text-white font-black text-sm rounded-xl shadow-md transition-all">
                    閉じる
                </button>
            </div>
        </div>
    </div>

    <script>
        let currentStaffData = """ + js_staff_data + """;
        let currentColRanking = """ + js_col_ranking + """;
        let activeDateFilter = 'all';
        let activeQuickFilter = 'all';
        let activeModalStaffName = null;

        function toggleSidebar() {
            const sidebar = document.getElementById('sidebar');
            const overlay = document.getElementById('sidebarOverlay');
            if (sidebar.classList.contains('-translate-x-full')) {
                sidebar.classList.remove('-translate-x-full');
                overlay.classList.remove('hidden');
            } else {
                sidebar.classList.add('-translate-x-full');
                overlay.classList.add('hidden');
            }
        }

        function setDateFilter(dateVal) {
            activeDateFilter = dateVal;
            renderPortal();
        }

        function setQuickFilter(type) {
            activeQuickFilter = type;
            renderPortal();
        }

        function extractDateFromFilename(fileName) {
            if (!fileName) return '10/02';
            const m = fileName.match(/_result(\\d{2})(\\d{2})/i);
            if (m) {
                return parseInt(m[1]) + '/' + m[2];
            }
            const mYmd = fileName.match(/202\\d(\\d{2})(\\d{2})/);
            if (mYmd) {
                return parseInt(mYmd[1]) + '/' + mYmd[2];
            }
            return '10/02';
        }

        function handleFileSelect(event) {
            const file = event.target.files[0];
            if (!file) return;

            const detectedDate = extractDateFromFilename(file.name);
            const statusEl = document.getElementById('currentStatusText');
            if (statusEl) statusEl.textContent = `⏳ ${file.name} (${detectedDate}分) 解析中...`;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellStyles: true });

                    let targetSheetName = workbook.SheetNames.includes('店舗マスタ') ? '店舗マスタ' : workbook.SheetNames[0];
                    const sheet = workbook.Sheets[targetSheetName];
                    if (!sheet || !sheet['!ref']) throw new Error('データシートが見つかりません');

                    const range = XLSX.utils.decode_range(sheet['!ref']);
                    const headers = [];
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: C })];
                        headers.push(cell && cell.v ? String(cell.v).trim() : `Column_${C+1}`);
                    }

                    const parsedShops = [];

                    for (let R = 1; R <= range.e.r; ++R) {
                        const shopCell = sheet[XLSX.utils.encode_cell({ r: R, c: 0 })];
                        const staffCell = sheet[XLSX.utils.encode_cell({ r: R, c: 3 })];
                        const svCell = sheet[XLSX.utils.encode_cell({ r: R, c: 4 })];

                        if (!shopCell || !shopCell.v) continue;
                        const shopName = String(shopCell.v).trim();
                        if (shopName === '店舗' || shopName === '店舗名' || shopName.includes('合計')) continue;

                        const staffName = staffCell && staffCell.v ? String(staffCell.v).trim() : '未設定';
                        const svName = svCell && svCell.v ? String(svCell.v).trim() : '';

                        let isARed = false;
                        const nonARedItems = [];

                        for (let C = 0; C <= range.e.c; ++C) {
                            const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
                            let isRed = false;

                            if (cell && cell.s && cell.s.fill) {
                                const fill = cell.s.fill;
                                const fg = fill.fgColor || fill.bgColor;
                                if (fg) {
                                    const rgb = String(fg.rgb || '').toUpperCase();
                                    if (rgb === 'FFFF0000' || rgb === 'FF0000' || rgb === 'RED' || rgb.endsWith('FF0000')) {
                                        isRed = true;
                                    }
                                }
                            }

                            if (isRed) {
                                const colName = headers[C] || `Col${C+1}`;
                                if (C === 0) {
                                    isARed = true;
                                } else {
                                    nonARedItems.push(colName);
                                }
                            }
                        }

                        parsedShops.push({
                            name: shopName,
                            staff: staffName,
                            sv: svName,
                            date: detectedDate,
                            is_a_red: isARed,
                            non_a_count: nonARedItems.length,
                            non_a_items: nonARedItems
                        });
                    }

                    if (parsedShops.length === 0) {
                        throw new Error('有効な店舗データが見つかりませんでした');
                    }

                    const updatedStaffData = {};

                    for (let s in currentStaffData) {
                        const existingShops = currentStaffData[s].shops.filter(sh => sh.date !== detectedDate);
                        updatedStaffData[s] = {
                            a_count: 0,
                            non_a_red_total: 0,
                            shops: existingShops
                        };
                    }

                    for (let item of parsedShops) {
                        const s = item.staff;
                        if (!updatedStaffData[s]) {
                            updatedStaffData[s] = { a_count: 0, non_a_red_total: 0, shops: [] };
                        }
                        updatedStaffData[s].shops.push({
                            name: item.name,
                            sv: item.sv,
                            date: item.date,
                            is_a_red: item.is_a_red,
                            non_a_count: item.non_a_count,
                            non_a_items: item.non_a_items
                        });
                    }

                    currentStaffData = updatedStaffData;
                    try {
                        localStorage.setItem('master_miss_portal_staff_data', JSON.stringify(currentStaffData));
                    } catch(e) { console.error('LocalStorage save error:', e); }
                    activeDateFilter = detectedDate;

                    if (statusEl) statusEl.textContent = `✅ ${file.name} (${detectedDate}分) スマート統合完了`;
                    renderPortal();

                } catch (err) {
                    console.error('File Read Error:', err);
                    if (statusEl) statusEl.textContent = `⚠️ 解析エラー: ${err.message || '読み込み失敗'}`;
                    alert(`【エラー】\n${err.message || 'ファイルの読み込みに失敗しました'}`);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        
        function getFilteredData() {
            const dateSet = new Set();
            for (let s in currentStaffData) {
                for (let sh of currentStaffData[s].shops) {
                    if (sh.date) dateSet.add(sh.date);
                }
            }
            const dates = Array.from(dateSet).sort();

            const filteredStaffMap = {};
            const filteredColRanking = {};
            const consecutiveMissShops = []; // Feature 2: 複数日連続指摘店舗

            let totalShopsCount = 0;
            let totalARedCount = 0;
            let totalNonARedCount = 0;

            // 店舗ごとの日付別指摘履歴マップ作成 (Feature 2)
            const shopHistoryMap = {};
            for (let s in currentStaffData) {
                for (let sh of currentStaffData[s].shops) {
                    if (!shopHistoryMap[sh.name]) {
                        shopHistoryMap[sh.name] = { name: sh.name, staff: s, dates: {}, datesWithRed: [] };
                    }
                    const isRed = sh.is_a_red || (sh.non_a_count && sh.non_a_count > 0);
                    if (isRed) {
                        shopHistoryMap[sh.name].datesWithRed.push(sh.date);
                    }
                    shopHistoryMap[sh.name].dates[sh.date] = isRed;
                }
            }

            for (let shopName in shopHistoryMap) {
                const info = shopHistoryMap[shopName];
                if (info.datesWithRed.length >= 2) {
                    consecutiveMissShops.push(info);
                }
            }

            for (let s in currentStaffData) {
                const targetShops = currentStaffData[s].shops.filter(sh => {
                    if (activeDateFilter === 'all') return true;
                    return sh.date === activeDateFilter;
                });

                if (targetShops.length === 0) continue;

                let aCount = 0;
                let nonACount = 0;

                for (let sh of targetShops) {
                    if (sh.is_a_red) aCount += 1;
                    nonACount += sh.non_a_count;
                    for (let colItem of sh.non_a_items) {
                        filteredColRanking[colItem] = (filteredColRanking[colItem] || 0) + 1;
                    }
                }

                totalShopsCount += targetShops.length;
                totalARedCount += aCount;
                totalNonARedCount += nonACount;

                // Feature 1: トレンド計算 (直前日付 vs 最新日付)
                let trendVal = 0;
                let trendType = 'same'; // 'up' (悪化), 'down' (改善), 'same'
                if (dates.length >= 2) {
                    const latestDate = dates[dates.length - 1];
                    const prevDate = dates[dates.length - 2];

                    const latestShops = currentStaffData[s].shops.filter(sh => sh.date === latestDate);
                    const prevShops = currentStaffData[s].shops.filter(sh => sh.date === prevDate);

                    const latestMisses = latestShops.reduce((acc, sh) => acc + (sh.is_a_red ? 1 : 0) + sh.non_a_count, 0);
                    const prevMisses = prevShops.reduce((acc, sh) => acc + (sh.is_a_red ? 1 : 0) + sh.non_a_count, 0);

                    const diff = latestMisses - prevMisses;
                    if (diff > 0) {
                        trendType = 'up'; // 悪化
                        trendVal = diff;
                    } else if (diff < 0) {
                        trendType = 'down'; // 改善
                        trendVal = Math.abs(diff);
                    }
                }

                filteredStaffMap[s] = {
                    a_count: aCount,
                    non_a_red_total: nonACount,
                    shops: targetShops,
                    miss_rate: targetShops.length > 0 ? Math.round((aCount / targetShops.length) * 1000) / 10 : 0,
                    trendType: trendType,
                    trendVal: trendVal
                };
            }

            return {
                dates: dates,
                staffMap: filteredStaffMap,
                colRanking: filteredColRanking,
                totalShops: totalShopsCount,
                totalARed: totalARedCount,
                totalNonARed: totalNonARedCount,
                consecutiveMissShops: consecutiveMissShops
            };
        }


        function renderPortal() {
            const dataInfo = getFilteredData();
            const staffMap = dataInfo.staffMap;
            const dates = dataInfo.dates;

            
            
            // 1. 日付フィルターUIのシームレス描画 (直近タブ + 連動ドロップダウン)
            const dateBtnContainer = document.getElementById('dateFilterButtons');
            if (dateBtnContainer) {
                let btnsHtml = '';

                // 🌐 全期間 (累計) ボタン
                const isAllActive = activeDateFilter === 'all';
                btnsHtml += `
                    <button onclick="setDateFilter('all')" class="px-3.5 py-2 rounded-xl text-xs sm:text-sm font-black transition-all shadow-sm flex items-center gap-1.5 ${isAllActive ? 'bg-slate-900 text-white border-2 border-slate-900 shadow-md ring-2 ring-slate-900' : 'bg-slate-100 hover:bg-slate-200 text-slate-800 border-2 border-slate-300'}">
                        <span>🌐</span>
                        <span>全期間 (累計)</span>
                    </button>
                `;

                if (dates.length > 0) {
                    const latestDate = dates[dates.length - 1];
                    const isLatestActive = activeDateFilter === latestDate;
                    const latestCount = Object.values(currentStaffData).reduce((acc, st) => acc + st.shops.filter(s => s.date === latestDate).length, 0);

                    // ⚡ 最新日ボタン
                    btnsHtml += `
                        <button onclick="setDateFilter('${latestDate}')" class="px-3.5 py-2 rounded-xl text-xs sm:text-sm font-black transition-all shadow-sm flex items-center gap-1.5 ${isLatestActive ? 'bg-red-600 text-white border-2 border-red-600 shadow-md ring-2 ring-red-600' : 'bg-red-50 hover:bg-red-100 text-red-700 border-2 border-red-300'}">
                            <span>⚡</span>
                            <span>最新 ${latestDate}分 (${latestCount}店舗)</span>
                        </button>
                    `;

                    // 📅 前回日ボタン (2日以上ある場合)
                    if (dates.length >= 2) {
                        const prevDate = dates[dates.length - 2];
                        const isPrevActive = activeDateFilter === prevDate;
                        const prevCount = Object.values(currentStaffData).reduce((acc, st) => acc + st.shops.filter(s => s.date === prevDate).length, 0);

                        btnsHtml += `
                            <button onclick="setDateFilter('${prevDate}')" class="px-3.5 py-2 rounded-xl text-xs sm:text-sm font-black transition-all shadow-sm flex items-center gap-1.5 ${isPrevActive ? 'bg-red-600 text-white border-2 border-red-600 shadow-md ring-2 ring-red-600' : 'bg-slate-100 hover:bg-slate-200 text-slate-800 border-2 border-slate-300'}">
                                <span>📅</span>
                                <span>前回 ${prevDate}分 (${prevCount}店舗)</span>
                            </button>
                        `;
                    }
                }

                // 📅 シームレス日付ドロップダウンセレクト
                btnsHtml += `
                    <div class="relative inline-block">
                        <select id="mainDateSelectDropdown" onchange="setDateFilter(this.value)" class="px-3.5 py-2 rounded-xl text-xs sm:text-sm font-black text-slate-900 border-2 border-slate-400 focus:outline-none focus:ring-2 focus:ring-red-600 shadow-sm cursor-pointer ${activeDateFilter !== 'all' ? 'bg-red-50 border-red-500 text-red-950 font-black' : 'bg-white'}">
                            <option value="all" ${activeDateFilter === 'all' ? 'selected' : ''}>🌐 全期間 (累計) 表示</option>
                            ${dates.map(d => {
                                const cnt = Object.values(currentStaffData).reduce((acc, st) => acc + st.shops.filter(s => s.date === d).length, 0);
                                return `<option value="${d}" ${activeDateFilter === d ? 'selected' : ''}>📅 ${d}分データ (${cnt}店舗)</option>`;
                            }).join('')}
                        </select>
                    </div>
                `;

                dateBtnContainer.innerHTML = btnsHtml;
            }



            // 2. サイドバーの日付切り替え表示
            const sidebarDateList = document.getElementById('sidebarDateList');
            if (sidebarDateList) {
                let sideHtml = '';
                for (let d of dates) {
                    const aCount = Object.values(currentStaffData).reduce((acc, st) => acc + st.shops.filter(s => s.date === d && s.is_a_red).length, 0);
                    const nonACount = Object.values(currentStaffData).reduce((acc, st) => acc + st.shops.filter(s => s.date === d).reduce((a, s) => a + s.non_a_count, 0), 0);
                    sideHtml += `
                        <button onclick="setDateFilter('${d}'); toggleSidebar();" class="w-full text-left p-3 rounded-xl hover:bg-slate-100 text-slate-900 font-black text-sm flex items-center justify-between transition-all ${activeDateFilter === d ? 'bg-red-50 border-2 border-red-400' : 'border border-slate-200'}">
                            <div class="flex items-center gap-2">
                                <span>📅</span>
                                <span>${d}分 データ</span>
                            </div>
                            <div class="flex items-center gap-1 text-xs">
                                <span class="bg-red-600 text-white px-2 py-0.5 rounded font-black">${aCount}店舗</span>
                                <span class="bg-amber-500 text-white px-2 py-0.5 rounded font-black">${nonACount}箇所</span>
                            </div>
                        </button>
                    `;
                }
                sidebarDateList.innerHTML = sideHtml;
            }

            // 3. バッジ
            const badgeEl = document.getElementById('currentDateDisplayBadge');
            if (badgeEl) {
                badgeEl.textContent = activeDateFilter === 'all' ? '全期間 (累計表示)' : `${activeDateFilter}分データ表示中`;
                badgeEl.className = activeDateFilter === 'all' ? 'bg-slate-900 text-white text-xs px-2.5 py-0.5 rounded-full font-black' : 'bg-red-600 text-white text-xs px-2.5 py-0.5 rounded-full font-black';
            }

            // 4. KPI
            document.getElementById('statTotalShops').innerHTML = `${dataInfo.totalShops} <span class="text-sm font-bold text-slate-700">店舗</span>`;
            document.getElementById('statARedCount').innerHTML = `${dataInfo.totalARed} <span class="text-sm font-bold text-slate-800">店舗</span>`;
            document.getElementById('statNonARedCount').innerHTML = `${dataInfo.totalNonARed} <span class="text-sm font-bold text-slate-800">箇所</span>`;

            let maxStaff = '--';
            let maxVal = 0;
            let perfectCount = 0;
            const staffList = Object.keys(staffMap);

            if (activeDateFilter === 'all') {
                // 【全期間 累計時】: A列外の赤セル累積箇所数が最も多いスタッフを探す
                for (let s of staffList) {
                    if (staffMap[s].non_a_red_total > maxVal) {
                        maxVal = staffMap[s].non_a_red_total;
                        maxStaff = s;
                    }
                    if (staffMap[s].a_count === 0 && staffMap[s].non_a_red_total === 0) {
                        perfectCount++;
                    }
                }
                document.getElementById('statTopMissStaff').textContent = maxVal > 0 ? `最多(A外赤): ${maxStaff} (${maxVal}箇所)` : '指摘なし';
            } else {
                // 【日別時】: A列赤指摘店舗が最も多いスタッフ
                for (let s of staffList) {
                    if (staffMap[s].a_count > maxVal) {
                        maxVal = staffMap[s].a_count;
                        maxStaff = s;
                    }
                    if (staffMap[s].a_count === 0) {
                        perfectCount++;
                    }
                }
                document.getElementById('statTopMissStaff').textContent = maxVal > 0 ? `最多: ${maxStaff} (${maxVal}店舗)` : 'ミス者なし';
            }
            document.getElementById('statPerfectCount').innerHTML = `${perfectCount} <span class="text-sm font-bold text-slate-800">名</span>`;
            const pRate = staffList.length > 0 ? Math.round((perfectCount / staffList.length) * 1000) / 10 : 0;
            document.getElementById('statPerfectRate').textContent = `${pRate}% がミスなし`;

            renderTableAndCards(staffMap);
            renderColumnRanking(dataInfo.colRanking);
        }

        function renderTableAndCards(staffMap) {
            const searchVal = (document.getElementById('searchInput')?.value || '').trim().toLowerCase();
            const tbody = document.getElementById('staffTableBody');
            const cardContainer = document.getElementById('staffCardList');
            if (tbody) tbody.innerHTML = '';
            if (cardContainer) cardContainer.innerHTML = '';

            let sortedStaff = Object.keys(staffMap).sort((a, b) => {
                if (staffMap[b].a_count !== staffMap[a].a_count) {
                    return staffMap[b].a_count - staffMap[a].a_count;
                }
                return staffMap[b].non_a_red_total - staffMap[a].non_a_red_total;
            });

            sortedStaff = sortedStaff.filter(s => {
                const info = staffMap[s];
                if (searchVal && !s.toLowerCase().includes(searchVal)) return false;
                if (activeQuickFilter === 'miss' && info.a_count === 0) return false;
                if (activeQuickFilter === 'perfect' && info.a_count > 0) return false;
                if (activeQuickFilter === 'nona' && info.non_a_red_total === 0) return false;
                return true;
            });

            let rank = 1;
            sortedStaff.forEach(s => {
                const info = staffMap[s];
                const totalShops = info.shops.length;

                if (tbody) {
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-50 transition-colors border-b border-slate-200';
                    tr.innerHTML = `

                                                <td class="py-4 px-6 font-black text-slate-900">
                            <div class="flex items-center gap-3">
                                <span class="w-7 h-7 rounded-lg ${rank <= 3 && info.a_count > 0 ? 'bg-red-600 text-white' : 'bg-slate-200 text-slate-700'} flex items-center justify-center text-xs font-black shrink-0">
                                    ${rank}
                                </span>
                                <div>
                                    <div class="text-base font-black">${s}</div>
                                    <div class="text-xs text-slate-500 font-bold">${totalShops}店舗巡回</div>
                                </div>
                            </div>
                        </td>

                        <td class="py-4 px-4 text-center font-black text-slate-800 text-base">${totalShops}</td>
                        <td class="py-4 px-4 text-center font-black text-base ${info.a_count > 0 ? 'text-red-600 bg-red-50/50' : 'text-slate-400'}">
                            ${info.a_count} 店舗
                        </td>
                        <td class="py-4 px-4 text-center font-black text-base">
                            <span class="px-2.5 py-1 rounded-lg ${info.miss_rate > 20 ? 'bg-red-100 text-red-700' : (info.miss_rate > 0 ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800')}">
                                ${info.miss_rate}%
                            </span>
                        </td>
                        <td class="py-4 px-4 text-center font-black text-base ${info.non_a_red_total > 0 ? 'text-amber-800 bg-amber-50/50' : 'text-slate-400'}">
                            ${info.non_a_red_total} 箇所
                        </td>
                        <td class="py-4 px-6 text-center">
                            <div class="flex items-center justify-center gap-2">
                                <button onclick="openModal('${s}')" class="px-3.5 py-1.5 rounded-xl bg-slate-900 hover:bg-black text-white text-xs font-black shadow-sm active:scale-95 transition-all">
                                    🔍 詳細一覧
                                </button>
                                <button onclick="copyChatText('${s}')" class="px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black shadow-sm active:scale-95 transition-all">
                                    📋 チャット作成
                                </button>
                            </div>
                        </td>
                    `;
                    tbody.appendChild(tr);
                }

                if (cardContainer) {
                    const card = document.createElement('div');
                    card.className = 'p-4 bg-white hover:bg-slate-50 space-y-3';
                    card.innerHTML = `
                        <div class="flex items-center justify-between">
                            <div class="flex items-center gap-2.5">
                                <span class="w-6 h-6 rounded-lg ${rank <= 3 && info.a_count > 0 ? 'bg-red-600 text-white' : 'bg-slate-200 text-slate-700'} flex items-center justify-center text-xs font-black shrink-0">
                                    ${rank}
                                </span>
                                <span class="font-black text-base text-slate-900">${s}</span>
                            </div>
                            <span class="text-xs font-black ${info.miss_rate > 0 ? 'text-red-600' : 'text-emerald-600'}">
                                ミス率 ${info.miss_rate}%
                            </span>
                        </div>
                        <div class="grid grid-cols-3 gap-2 text-center text-xs font-black">
                            <div class="p-2 rounded-lg bg-slate-100">
                                <div class="text-slate-500">巡回</div>
                                <div class="text-sm mt-0.5">${totalShops}店舗</div>
                            </div>
                            <div class="p-2 rounded-lg ${info.a_count > 0 ? 'bg-red-100 text-red-700' : 'bg-slate-100'}">
                                <div>A列赤</div>
                                <div class="text-sm mt-0.5">${info.a_count}店舗</div>
                            </div>
                            <div class="p-2 rounded-lg ${info.non_a_red_total > 0 ? 'bg-amber-100 text-amber-900' : 'bg-slate-100'}">
                                <div>A列外赤</div>
                                <div class="text-sm mt-0.5">${info.non_a_red_total}箇所</div>
                            </div>
                        </div>
                        <div class="flex items-center gap-2 pt-1">
                            <button onclick="openModal('${s}')" class="flex-1 py-2 rounded-xl bg-slate-900 text-white text-xs font-black text-center shadow-sm">
                                🔍 詳細
                            </button>
                            <button onclick="copyChatText('${s}')" class="flex-1 py-2 rounded-xl bg-emerald-600 text-white text-xs font-black text-center shadow-sm">
                                📋 チャットコピー
                            </button>
                        </div>
                    `;
                    cardContainer.appendChild(card);
                }
                rank++;
            });
        }

        function renderColumnRanking(colRankingMap) {
            const container = document.getElementById('columnRankingContainer');
            if (!container) return;

            const sortedCols = Object.keys(colRankingMap).sort((a, b) => colRankingMap[b] - colRankingMap[a]);

            if (sortedCols.length === 0) {
                container.innerHTML = `<div class="col-span-full p-4 text-center text-xs font-bold text-slate-500">指摘箇所はありませんでした</div>`;
                return;
            }

            let html = '';
            let idx = 1;
            for (let col of sortedCols.slice(0, 6)) {
                const count = colRankingMap[col];
                html += `
                    <div class="p-3.5 rounded-xl bg-slate-50 border border-slate-200 flex items-center justify-between">
                        <div class="flex items-center gap-2.5 truncate">
                            <span class="w-6 h-6 rounded-lg ${idx <= 3 ? 'bg-amber-500 text-white' : 'bg-slate-200 text-slate-700'} flex items-center justify-center text-xs font-black shrink-0">
                                ${idx}
                            </span>
                            <span class="text-xs sm:text-sm font-black text-slate-900 truncate">${col}</span>
                        </div>
                        <span class="text-xs font-black bg-amber-100 text-amber-900 px-2.5 py-1 rounded-lg shrink-0">${count} 箇所</span>
                    </div>
                `;
                idx++;
            }
            container.innerHTML = html;
        }

        function openModal(staffName) {
            activeModalStaffName = staffName;
            const dataInfo = getFilteredData();
            const info = dataInfo.staffMap[staffName];
            if (!info) return;

            document.getElementById('modalStaffName').textContent = staffName;
            document.getElementById('modalStaffSubText').textContent = activeDateFilter === 'all' ? `${staffName} さんの全期間指摘一覧` : `${staffName} さんの${activeDateFilter}分 指摘一覧`;

            document.getElementById('modalShopCount').innerHTML = `${info.shops.length} <span class="text-xs text-slate-700">店舗</span>`;
            document.getElementById('modalARedCount').innerHTML = `${info.a_count} <span class="text-xs text-red-900">店舗</span>`;
            document.getElementById('modalNonARedCount').innerHTML = `${info.non_a_red_total} <span class="text-xs text-amber-950">箇所</span>`;
            document.getElementById('modalShopBadge').textContent = `${info.shops.length}店舗`;

            const shopListEl = document.getElementById('modalShopList');
            shopListEl.innerHTML = '';

            info.shops.forEach(s => {
                const item = document.createElement('div');
                item.className = `p-3.5 rounded-2xl border-2 transition-all ${s.is_a_red ? 'bg-red-50/70 border-red-300' : 'bg-slate-50 border-slate-200'}`;
                
                let nonAItemsHtml = '';
                if (s.non_a_items && s.non_a_items.length > 0) {
                    nonAItemsHtml = `
                        <div class="mt-2 flex flex-wrap gap-1.5">
                            ${s.non_a_items.map(it => `<span class="text-xs bg-amber-100 text-amber-900 font-black px-2 py-0.5 rounded border border-amber-300">⚠️ ${it}</span>`).join('')}
                        </div>
                    `;
                }

                item.innerHTML = `
                    <div class="flex items-center justify-between">
                        <div class="flex items-center gap-2">
                            <span class="text-xs font-black bg-slate-200 text-slate-800 px-2 py-0.5 rounded">${s.date || '10/01'}</span>
                            <span class="font-black text-sm text-slate-900">${s.name}</span>
                        </div>
                        ${s.is_a_red ? '<span class="text-xs bg-red-600 text-white font-black px-2.5 py-0.5 rounded-full">🚨 A列 赤指摘</span>' : '<span class="text-xs bg-emerald-600 text-white font-black px-2.5 py-0.5 rounded-full">OK</span>'}
                    </div>
                    ${nonAItemsHtml}
                `;
                shopListEl.appendChild(item);
            });

            document.getElementById('detailModal').classList.remove('hidden');
        }

        function closeModal() {
            document.getElementById('detailModal').classList.add('hidden');
        }

        function copyChatText(staffName) {
            const dataInfo = getFilteredData();
            const info = dataInfo.staffMap[staffName];
            if (!info) return;

            const text = `【店舗マスタチェック指摘・修正のお願い】
${staffName} さん

お疲れ様です。マスタチェックにて巡回店舗の指摘箇所が検出されました。早急に確認と修正をお願いいたします。

■ 巡回スタッフ: ${staffName}
■ 対象期間: ${activeDateFilter === 'all' ? '全期間 (累計)' : activeDateFilter + '分'}
■ 巡回店舗数: ${info.shops.length} 店舗
■ A列赤指摘 (マスタミス): ${info.a_count} 回
■ A列外赤セル内訳: ${info.non_a_red_total} 箇所

何卒よろしくお願いいたします。`;

            navigator.clipboard.writeText(text).then(() => {
                alert(`✅ 【${staffName}さん宛ての指導チャット文面】をコピーしました！\\nLINEやチャットツールにそのまま貼り付けて送信できます。`);
            }).catch(err => {
                console.error(err);
                alert(`文面: \\n${text}`);
            });
        }

        function copyStaffChatModal() {
            if (activeModalStaffName) {
                copyChatText(activeModalStaffName);
            }
        }

        function exportToCSV() {
            const dataInfo = getFilteredData();
            const staffMap = dataInfo.staffMap;

            let csv = '\\uFEFF順位,スタッフ名,巡回店舗数,A列赤指摘店舗数,ミス率(%),A列外赤セル箇所\\n';
            let rank = 1;
            for (let s in staffMap) {
                const st = staffMap[s];
                csv += `${rank},"${s}",${st.shops.length},${st.a_count},${st.miss_rate},${st.non_a_red_total}\\n`;
                rank++;
            }

            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `店舗マスタミス集計_${activeDateFilter === 'all' ? '全期間' : activeDateFilter}.csv`;
            link.click();
        }

        
        // DropZone Event Listeners
        document.addEventListener('DOMContentLoaded', () => {
            const dropZone = document.getElementById('dropZone');
            const fileInput = document.getElementById('fileInput');

            if (dropZone && fileInput) {
                dropZone.addEventListener('click', (e) => {
                    if (e.target.tagName !== 'BUTTON' && e.target !== fileInput) {
                        fileInput.click();
                    }
                });

                dropZone.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    dropZone.classList.add('dragover');
                });

                dropZone.addEventListener('dragleave', () => {
                    dropZone.classList.remove('dragover');
                });

                dropZone.addEventListener('drop', (e) => {
                    e.preventDefault();
                    dropZone.classList.remove('dragover');
                    if (e.dataTransfer && e.dataTransfer.files.length > 0) {
                        fileInput.files = e.dataTransfer.files;
                        handleFileSelect({ target: fileInput });
                    }
                });
            }
            renderPortal();
        });

    </script>
</body>
</html>"""

with open(OUT_HTML, 'w', encoding='utf-8') as out:
    out.write(html_content)

print('Generated clean index.html successfully with non-f-string template!')
