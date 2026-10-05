import os

base_dir = r"G:\我的雲端硬碟\3_大衛工作\2_進金生能源\8.0 專案機會\8.2 儲能工程\8.2.3 錶後儲能：用電大戶\高雄_榮總醫院\vghks-bess-consulting-platform"
html_path = os.path.join(base_dir, "index.html")

html_content = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>高雄榮民總醫院 表後儲能系統建置及節能技術服務 ｜ 無雜訊顧問決策簡報</title>
  <meta name="description" content="高雄榮民總醫院 ENG-1150058 4.38億元表後儲能巨額採購案決策分析、招標重點摘要、六大隱藏風險疑慮與投標必勝五維戰略">
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Noto+Sans+TC:wght@300;400;500;700;900&display=swap" rel="stylesheet">
  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          colors: {
            brand: {
              navy: '#051C2C',
              royal: '#1E3A8A',
              gold: '#D97706',
              forest: '#004F32',
              lime: '#29BA74',
              purple: '#A100FF',
              cyan: '#06B6D4'
            }
          },
          fontFamily: {
            sans: ['Inter', 'Noto Sans TC', 'sans-serif'],
          }
        }
      }
    }
  </script>
  <style>
    body {
      font-family: 'Inter', 'Noto Sans TC', sans-serif;
      background-color: #0b1120;
      color: #f1f5f9;
    }
    .slide-card {
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .custom-scrollbar::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
      background: rgba(15, 23, 42, 0.6);
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
      background: rgba(100, 116, 139, 0.5);
      border-radius: 4px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {
      background: rgba(148, 163, 184, 0.8);
    }
  </style>
</head>
<body class="min-h-screen flex flex-col bg-slate-950 text-slate-100 antialiased selection:bg-amber-500 selection:text-slate-950">

  <!-- Top Navigation Bar -->
  <header class="sticky top-0 z-50 bg-slate-900/90 backdrop-blur-md border-b border-slate-800 px-4 lg:px-8 py-3 flex items-center justify-between">
    <div class="flex items-center gap-3">
      <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-amber-500 to-blue-600 flex items-center justify-center font-black text-slate-950 shadow-lg shadow-blue-500/20">
        6692
      </div>
      <div>
        <div class="flex items-center gap-2">
          <span class="text-xs font-bold uppercase tracking-wider text-amber-400 bg-amber-950/60 border border-amber-500/30 px-2 py-0.5 rounded-full">
            巨額採購案 ｜ ENG-1150058
          </span>
          <span class="text-xs text-slate-400 hidden sm:inline">進金生能源 投標決策專用</span>
        </div>
        <h1 class="text-sm sm:text-base font-bold text-slate-100 tracking-tight">
          高雄榮總 表後儲能系統建置及節能技術服務 ｜ 無雜訊顧問簡報
        </h1>
      </div>
    </div>

    <!-- Controls -->
    <div class="flex items-center gap-2 sm:gap-3">
      <!-- Calculator Modal Trigger -->
      <button onclick="toggleModal('calc-modal')" class="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs sm:text-sm font-semibold flex items-center gap-1.5 transition border border-slate-700">
        <i data-lucide="calculator" class="w-4 h-4 text-emerald-400"></i>
        <span class="hidden md:inline">收益拆分試算</span>
      </button>

      <!-- Toggle Speaker Notes -->
      <button id="toggle-notes-btn" onclick="toggleNotes()" class="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs sm:text-sm font-semibold flex items-center gap-1.5 transition border border-slate-700">
        <i data-lucide="message-square" class="w-4 h-4 text-blue-400"></i>
        <span class="hidden md:inline">備忘講稿</span>
      </button>

      <!-- Download PPTX -->
      <a href="vghks_bess_zero_noise_consulting.pptx" download class="px-3.5 py-1.5 rounded-lg bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-400 hover:to-amber-500 text-slate-950 font-bold text-xs sm:text-sm flex items-center gap-1.5 shadow-md shadow-amber-500/20 transition">
        <i data-lucide="download" class="w-4 h-4"></i>
        <span>下載 PPTX</span>
      </a>

      <!-- Fullscreen Toggle -->
      <button onclick="toggleFullScreen()" class="p-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition border border-slate-700">
        <i data-lucide="maximize" class="w-4 h-4"></i>
      </button>
    </div>
  </header>

  <!-- Main Slide Viewport -->
  <main class="flex-1 flex flex-col items-center justify-center p-3 sm:p-6 lg:p-8 max-w-7xl mx-auto w-full">

    <!-- Slide Container (16:9 ratio card) -->
    <div id="slide-viewport" class="w-full bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl overflow-hidden relative flex flex-col slide-card">
      
      <!-- Slide Top Banner -->
      <div class="px-6 py-4 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
        <div class="flex items-center gap-2">
          <span id="slide-archetype-badge" class="px-2.5 py-1 rounded-md text-xs font-bold uppercase tracking-wider bg-blue-950/80 text-blue-400 border border-blue-500/30">
            McKinsey Deep Navy ｜ 封面頁
          </span>
          <span id="slide-counter-badge" class="text-xs font-mono text-slate-400 bg-slate-800 px-2 py-0.5 rounded">
            01 / 12
          </span>
        </div>
        <div class="flex items-center gap-4 text-xs text-slate-400 font-mono">
          <span>Source: 高雄榮民總醫院 ENG-1150058</span>
          <span class="hidden sm:inline">16:9 Zero-Noise Architecture</span>
        </div>
      </div>

      <!-- Slide Main Body -->
      <div class="p-6 sm:p-8 flex-1 flex flex-col justify-between">
        
        <!-- Header Zone: Action Title & Sub -->
        <div class="mb-6">
          <h2 id="slide-action-title" class="text-xl sm:text-2xl lg:text-3xl font-black text-slate-100 tracking-tight leading-snug">
            高雄榮民總醫院「表後儲能系統建置及節能技術服務」
          </h2>
          <p id="slide-sub-title" class="mt-2 text-sm sm:text-base text-slate-300 font-normal">
            採購案號：ENG-1150058 ｜ 4.38 億元巨額採購案決策分析與投標必勝攻略
          </p>
        </div>

        <!-- Evidence Zone: Left Cards & Right Image -->
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch flex-1">
          
          <!-- Left: Evidence & Analysis Card (7 cols) -->
          <div id="slide-evidence-container" class="lg:col-span-7 bg-slate-950/60 border border-slate-800 rounded-xl p-5 sm:p-6 flex flex-col justify-center space-y-3.5">
            <!-- Dynamic Bullets -->
          </div>

          <!-- Right: Visual Diagram / Image (5 cols) -->
          <div class="lg:col-span-5 bg-slate-950/80 border border-slate-800 rounded-xl p-2.5 flex flex-col items-center justify-center relative group overflow-hidden">
            <img id="slide-img" src="assets/images/slide_01_cover_campus.jpg" alt="簡報圖像" class="w-full h-full object-cover rounded-lg shadow-md cursor-zoom-in transition duration-300 group-hover:scale-[1.02]" onclick="openLightbox(this.src)">
            <div class="absolute bottom-4 right-4 bg-slate-950/80 backdrop-blur-md border border-slate-700 text-[11px] font-mono px-2 py-1 rounded text-slate-300 pointer-events-none flex items-center gap-1">
              <i data-lucide="zoom-in" class="w-3 h-3 text-amber-400"></i>
              點擊放大檢視
            </div>
          </div>

        </div>

      </div>

      <!-- Slide Bottom Control Bar -->
      <div class="px-6 py-3 bg-slate-950/80 border-t border-slate-800 flex items-center justify-between">
        <button onclick="prevSlide()" class="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs sm:text-sm font-semibold flex items-center gap-1.5 transition border border-slate-700">
          <i data-lucide="chevron-left" class="w-4 h-4"></i>
          <span>上一頁 (Prev)</span>
        </button>

        <!-- Slide Indicators Dots -->
        <div id="slide-dots" class="hidden sm:flex items-center gap-1.5">
          <!-- Populated by JS -->
        </div>

        <button onclick="nextSlide()" class="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs sm:text-sm font-semibold flex items-center gap-1.5 transition border border-slate-700">
          <span>下一頁 (Next)</span>
          <i data-lucide="chevron-right" class="w-4 h-4"></i>
        </button>
      </div>

    </div>

    <!-- Speaker Notes Drawer (Toggled) -->
    <div id="notes-drawer" class="w-full mt-4 bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-xl hidden transition-all">
      <div class="flex items-center justify-between pb-3 border-b border-slate-800 mb-3">
        <div class="flex items-center gap-2">
          <i data-lucide="mic" class="w-4 h-4 text-amber-400"></i>
          <span class="text-sm font-bold text-slate-200">顧問報告備忘錄 (Speaker Notes)</span>
        </div>
        <span class="text-xs text-slate-400">按三段式講稿：結論 ➔ 證據細節 ➔ 下頁轉場</span>
      </div>
      <div id="notes-content" class="text-sm text-slate-300 whitespace-pre-line leading-relaxed font-sans bg-slate-950 p-4 rounded-lg border border-slate-800/80">
        <!-- Notes here -->
      </div>
    </div>

    <!-- Thumbnails Strip -->
    <div class="w-full mt-6">
      <div class="flex items-center justify-between mb-2">
        <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
          <i data-lucide="layers" class="w-3.5 h-3.5 text-blue-400"></i>
          全頁目錄速覽 (12 Slides Index)
        </h3>
        <span class="text-xs text-slate-400 font-mono">支援鍵盤 ◀ ▶ 箭頭翻頁</span>
      </div>
      <div id="thumbnails-container" class="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-12 gap-2 overflow-x-auto pb-2 custom-scrollbar">
        <!-- Thumbnails populated by JS -->
      </div>
    </div>

  </main>

  <!-- Interactive Calculator Modal -->
  <div id="calc-modal" class="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm hidden flex items-center justify-center p-4">
    <div class="bg-slate-900 border border-slate-800 rounded-2xl max-w-lg w-full p-6 shadow-2xl relative">
      <button onclick="toggleModal('calc-modal')" class="absolute top-4 right-4 text-slate-400 hover:text-white">
        <i data-lucide="x" class="w-5 h-5"></i>
      </button>

      <div class="flex items-center gap-2 mb-4">
        <div class="p-2 rounded-lg bg-emerald-950 text-emerald-400 border border-emerald-500/30">
          <i data-lucide="trending-up" class="w-5 h-5"></i>
        </div>
        <div>
          <h3 class="text-lg font-bold text-slate-100">收益分享比率評分與營收試算器</h3>
          <p class="text-xs text-slate-400">招標上限 85.7%，每降 0.5% 加 1 分，最高 20 分</p>
        </div>
      </div>

      <div class="space-y-4 text-sm">
        <div>
          <div class="flex justify-between items-center mb-1">
            <label class="text-slate-300 font-semibold">廠商投標收益分享比率 (%)</label>
            <span id="calc-pct-label" class="text-amber-400 font-bold text-base font-mono">81.70%</span>
          </div>
          <input id="calc-slider" type="range" min="75" max="85.7" step="0.1" value="81.7" oninput="updateCalculator(this.value)" class="w-full accent-amber-500 h-2 bg-slate-700 rounded-lg cursor-pointer">
          <div class="flex justify-between text-[11px] text-slate-400 mt-1">
            <span>75.0% (過度讓利)</span>
            <span class="text-emerald-400 font-bold">81.7% (最佳滿分點)</span>
            <span>85.7% (基準上限)</span>
          </div>
        </div>

        <div class="grid grid-cols-2 gap-3 pt-2">
          <div class="bg-slate-950 p-3 rounded-xl border border-slate-800">
            <span class="text-xs text-slate-400">評選項目四得分</span>
            <div id="calc-score" class="text-2xl font-black text-emerald-400 mt-1">20 / 20 分</div>
            <span id="calc-score-sub" class="text-[11px] text-emerald-500">★ 獲得滿分最高評價</span>
          </div>
          <div class="bg-slate-950 p-3 rounded-xl border border-slate-800">
            <span class="text-xs text-slate-400">預估每月可得收益</span>
            <div id="calc-monthly" class="text-2xl font-black text-amber-400 mt-1">290.0 萬元</div>
            <span class="text-[11px] text-slate-400">以月效益 355 萬計算</span>
          </div>
        </div>

        <div class="bg-slate-950 p-4 rounded-xl border border-slate-800 space-y-2">
          <div class="flex justify-between">
            <span class="text-slate-400 text-xs">12 年 (144 個月) 累計總收益：</span>
            <span id="calc-total" class="font-bold text-slate-200 font-mono">4.17 億元</span>
          </div>
          <div class="flex justify-between">
            <span class="text-slate-400 text-xs">20 年全生命週期總收益：</span>
            <span id="calc-life" class="font-bold text-slate-200 font-mono">6.96 億元</span>
          </div>
          <div class="flex justify-between border-t border-slate-800 pt-2">
            <span class="text-slate-400 text-xs">相較 85.7% 減少拆分比例：</span>
            <span id="calc-diff" class="text-rose-400 font-semibold text-xs font-mono">-4.00% (-1,704萬元)</span>
          </div>
        </div>

        <div class="p-3 bg-amber-950/40 border border-amber-500/30 rounded-lg text-xs text-amber-200 leading-relaxed">
          <strong>顧問洞察：</strong> 報價 81.70% 恰好降 4.0%，取得 8 點級距加分穩拿滿分 20 分。低於 81.70% 分數不再增加，白白損失合約收益；高於 81.70% 則每 0.5% 失 1 分，極易在序位法遭到淘汰！
        </div>
      </div>
    </div>
  </div>

  <!-- Image Lightbox Modal -->
  <div id="lightbox-modal" class="fixed inset-0 z-50 bg-black/90 backdrop-blur-md hidden flex items-center justify-center p-4 cursor-zoom-out" onclick="closeLightbox()">
    <div class="max-w-5xl max-h-[90vh] w-full flex flex-col items-center justify-center relative">
      <img id="lightbox-img" src="" alt="放大圖片" class="max-w-full max-h-[85vh] object-contain rounded-xl shadow-2xl border border-slate-700">
      <p class="text-xs text-slate-400 mt-2">點擊任意處關閉放大檢視</p>
    </div>
  </div>

  <!-- Footer -->
  <footer class="bg-slate-950 border-t border-slate-800/80 px-6 py-4 text-center text-xs text-slate-400">
    <div class="flex flex-wrap items-center justify-center gap-4">
      <span>進金生能源服務股份有限公司 (6692)</span>
      <span>•</span>
      <span>高雄榮民總醫院 表後儲能系統建置及節能技術服務 (ENG-1150058)</span>
      <span>•</span>
      <span>8 種無雜訊顧問版型 × 三大顧問視覺系統</span>
    </div>
  </footer>

  <!-- Interactive Presentation Script -->
  <script>
    const slidesData = [
      {
        num: 1,
        tag: "PROJECT OVERVIEW ｜ 決策旗艦報告",
        archetype: "McKinsey Deep Navy ｜ 封面頁",
        badgeColor: "bg-blue-950/80 text-blue-400 border-blue-500/30",
        actionTitle: "高雄榮民總醫院「表後儲能系統建置及節能技術服務」",
        subTitle: "採購案號：ENG-1150058 ｜ 4.38 億元巨額採購案決策分析與投標必勝攻略",
        img: "assets/images/slide_01_cover_campus.jpg",
        bullets: [
          { bold: "標案性質：", text: "巨額勞務採購（限制性招標，公開評選準用最有利標，序位法決標）" },
          { bold: "金額規模：", text: "本期預算 4.38 億元（12年），含兩次擴充總採購金額上看 7.30 億元" },
          { bold: "商業模式：", text: "純 ESCO / BOO（廠商全額投資建置與維運，按月實際效益分成）" },
          { bold: "核心要求：", text: "中斷供電供應普通電源≧2小時、充放電效率≧85%、嚴格防逆送" }
        ],
        notes: "【採用版型】：封面頁 ｜ 視覺風格：McKinsey Deep Navy\\n① 本頁結論：高雄榮總表後儲能案為 4.38 億之巨額純 ESCO 標案，投標決勝在於控制工期與醫療供電安全。\\n② 補充說明：本報告提供招標重點摘要、6大隱藏風險疑慮，以及確保奪標的五維準備方案。\\n③ 下一頁轉場：首先呈現全案的核心決策總覽。"
      },
      {
        num: 2,
        tag: "EXECUTIVE SUMMARY ｜ 決策架構總覽",
        archetype: "McKinsey Deep Navy ｜ 版型 3：結合論法 (Combined Tree)",
        badgeColor: "bg-blue-950/80 text-blue-400 border-blue-500/30",
        actionTitle: "「純 ESCO 全額出資」結合「醫療級 2 小時備援」與「81.7% 定價博弈」，鎖定 4.38 億案場",
        subTitle: "三大核心前提同時成立，推導出進金生能源必勝奪標之頂層策略",
        img: "assets/images/slide_02_bess_overview.jpg",
        bullets: [
          { bold: "商業支柱：", text: "純 ESCO 效益分成模式，廠商承擔全部 CAPEX，依每月實際收益支付" },
          { bold: "技術支柱：", text: "污水場 660m² 防腐防汛、普通電源 2 小時孤島供電、全非陸製資通訊" },
          { bold: "博弈支柱：", text: "收益分享比率鎖定 81.70% 奪滿分 20 分，兼顧評選第 1 與營收最大化" },
          { bold: "實績門檻：", text: "特定電力供應業電業執照 ＋ 5年內單次 1.75 億或累計 4.38 億實績" }
        ],
        notes: "【採用版型】：版型 3：結合論法 ｜ 視覺風格：McKinsey Deep Navy\\n① 本頁結論：三大支柱必須同時到位才能確保拿下此巨額案場。\\n② 補充說明：商業模式需耐受無保底風險，工程需克服污水場腐蝕，定價須精準卡位滿分。\\n③ 下一頁轉場：接著檢視招標文件所載的合約採購規模與擴充機制。"
      },
      {
        num: 3,
        tag: "SCALE & EXPANSION ｜ 採購規模與擴充",
        archetype: "Accenture Core Purple ｜ 版型 1：一對一論法 (One-to-One)",
        badgeColor: "bg-purple-950/80 text-purple-400 border-purple-500/30",
        actionTitle: "本期預算 4.38 億元加上兩次續約擴充，全生命週期採購總規模上看 7.30 億元",
        subTitle: "單一強力核心指標：本案為橫跨 20 年之長生命週期巨額能源技術服務",
        img: "assets/images/slide_03_contract_scale.jpg",
        bullets: [
          { bold: "本期 12 年 (144 個月)：", text: "預算上限 4億3,809萬8,400 元（每月預估 355 萬 × 85.7%）" },
          { bold: "第 1 次續約 (第 13~16 年)：", text: "增購上限 1億4,603萬2,800 元（續約門檻：RTE ≧ 85%）" },
          { bold: "第 2 次續約 (第 17~20 年)：", text: "增購上限 1億4,603萬2,800 元（續約門檻：RTE ≧ 80%）" },
          { bold: "全案總額：", text: "本期加擴充合計達 7億3,016萬4,000 元；履約保證金 400 萬元" }
        ],
        notes: "【採用版型】：版型 1：一對一論法 ｜ 視覺風格：Accenture Core Purple\\n① 本頁結論：這是一檔長達 20 年、總額高達 7.3 億元的巨額技術服務合約。\\n② 補充說明：第 13 年仍要求 85% 效率，投標團隊必須在中期編列電芯置換 CAPEX。\\n③ 下一頁轉場：接下來剖析合約明訂的商業模式與計價公式。"
      },
      {
        num: 4,
        tag: "BUSINESS MODEL ｜ 商業模式與計價",
        archetype: "McKinsey Deep Navy ｜ 版型 5：麥肯錫論法 (McKinsey Matrix)",
        badgeColor: "bg-blue-950/80 text-blue-400 border-blue-500/30",
        actionTitle: "「零期初工程款 ＋ 每月實際收益拆分」，廠商承擔全額 CAPEX 與回收風險",
        subTitle: "沿著生命週期流程，診斷契約計價機制與廠商財務護欄",
        img: "assets/images/slide_04_revenue_ems.jpg",
        bullets: [
          { bold: "建置階段：", text: "廠商全額自費出資設備、土建、消防與併聯引接，院方無期初工程款" },
          { bold: "運轉調度：", text: "EMS 實施尖離峰削峰填谷，嚴格落實防逆送保護，維持 ≧ 85% RTE" },
          { bold: "收益公式：", text: "（尖峰電價×供電量）－（離峰電價×充電量）＋ 輔助服務／需量反應收益" },
          { bold: "風險對策：", text: "無保底機制（No Take-or-Pay），需藉由高榮白天大基載確保放電全額吸收" }
        ],
        notes: "【採用版型】：版型 5：麥肯錫論法 ｜ 視覺風格：McKinsey Deep Navy\\n① 本頁結論：合約是純 ESCO 模式，廠商承擔所有投資與運轉責任，按實際效益分成。\\n② 補充說明：每月款項需經書面驗收，計量電錶必須具備經濟部標檢局檢定合格。\\n③ 下一頁轉場：要支撐此效益，必須深入高榮院區的真實用電結構。"
      },
      {
        num: 5,
        tag: "LOAD & CAPACITY ｜ 用電分析與容量",
        archetype: "BCG Forest Green ｜ 版型 1：一對一論法 (One-to-One)",
        badgeColor: "bg-emerald-950/80 text-emerald-400 border-emerald-500/30",
        actionTitle: "契約容量 8,300 kW 搭配夏月超約特性，精算 3~4 MW / 8~10 MWh 為最佳化投標容量",
        subTitle: "以高榮 19 個月實測用電大數據，反推最佳經濟效益與用地平衡點",
        img: "assets/images/slide_05_load_capacity.jpg",
        bullets: [
          { bold: "用電總量：", text: "114/01~115/07 用電 7,435 萬度，電費 3.06 億元，平均單價 4.12 元/度" },
          { bold: "負載分佈：", text: "空調設備佔 54.0%、動力設備佔 17.6%、照明佔 15.8%，夏季下午極度吃重" },
          { bold: "超約痛點：", text: "114 年 5~7 月連續產生超約附加費（累計逾 31 萬元），削峰急迫性極高" },
          { bold: "最佳配置：", text: "3~4 MW / 8~10 MWh 系統可平抑夏月尖峰，日放電 1.2 萬度，直逼效益目標" }
        ],
        notes: "【採用版型】：版型 1：一對一論法 ｜ 視覺風格：BCG Forest Green\\n① 本頁結論：高榮用電基載大且夏月空調尖峰明顯，3~4 MW 儲能最具經濟削峰效益。\\n② 補充說明：污水處理場 660 m² 土地面積亦剛好可容納約 2~3 座 40 呎集裝箱儲能櫃。\\n③ 下一頁轉場：然而，污水場現場環境十分嚴苛，工程防護必須做足。"
      },
      {
        num: 6,
        tag: "SITE CONDITIONS ｜ 場址與工程防護",
        archetype: "McKinsey Deep Navy ｜ 版型 6：對立論法 (Opposition)",
        badgeColor: "bg-blue-950/80 text-blue-400 border-blue-500/30",
        actionTitle: "污水處理場 660 m² 硫化氫腐蝕環境，全面導入 C5-M 重防腐與基礎防汛抬高工程",
        subTitle: "針對污水廠嚴酷環境挑戰，逐項提出對應之重工業級工程防護對策",
        img: "assets/images/slide_06_sewage_civil.jpg",
        bullets: [
          { bold: "化學腐蝕 vs. C5-M 塗裝：", text: "污水場高濃度 H2S 與 NH3 氣體 ➔ 貨櫃外殼 C5-M 海洋重防腐，空調加裝化學吸附濾網" },
          { bold: "地下管溝 vs. 筏式微型樁：", text: "場址地下密布大型集水槽與管線 ➔ 採用淺層筏式基礎搭配微型抗拔樁，避開管溝" },
          { bold: "防汛積水 vs. 基座墊高：", text: "高雄夏季常有強降雨積淹水疑慮 ➔ 混凝土設備基座墊高 800~1200mm，配置雙迴路排水" },
          { bold: "動線維持 vs. 模組化佈局：", text: "契約規定維持污水廠維修通行 ➔ 預留 4 米法定應急救援通道，滿足消防指引要求" }
        ],
        notes: "【採用版型】：版型 6：對立論法 ｜ 視覺風格：McKinsey Deep Navy\\n① 本頁結論：污水場是高腐蝕與受限空間，非一般空地，防腐與防汛設計是審標硬指標。\\n② 補充說明：若未在企劃書提出 C5-M 防腐與基座墊高，將被工務室委員重扣分數。\\n③ 下一頁轉場：除土建外，醫院最重視的核心功能是市電斷電時的電力救援切換。"
      },
      {
        num: 7,
        tag: "RESILIENCE & MICROGRID ｜ 醫療供電韌性",
        archetype: "Accenture Core Purple ｜ 版型 4：連鎖論法 (Chain / Flow)",
        badgeColor: "bg-purple-950/80 text-purple-400 border-purple-500/30",
        actionTitle: "市電中斷瞬時切換孤島，確保普通電源 2 小時續航並與柴油發電機防衝擊協同",
        subTitle: "端到端展示台電斷電時之毫秒級自動切換與發電機互鎖時序流程",
        img: "assets/images/slide_07_resilience_switch.jpg",
        bullets: [
          { bold: "市電中斷 ➔ 瞬時解聯：", text: "台電電網斷電時，高壓斷路器 200ms 快速脫扣，形成孤島隔離" },
          { bold: "電網形成 ➔ 普通電源 2 小時：", text: "PCS 轉為 VF 電壓源模式，維持全院照明、電梯、空調等普通電源" },
          { bold: "平時調度 ➔ 最低 SOC 護欄：", text: "EMS 平時調度鎖定最低備援荷電量 ≧ 30%，保證隨時有 2 小時電量" },
          { bold: "發電機協同 ➔ 防反充保護：", text: "儲能系統與既設柴發並聯時增設逆功率防護電驛，防止發電機反充跳脫" }
        ],
        notes: "【採用版型】：版型 4：連鎖論法 ｜ 視覺風格：Accenture Core Purple\\n① 本頁結論：微電網孤島切換與防發電機逆功率反充是醫療韌性技術關鍵。\\n② 補充說明：重症 ICU 與手術室有專屬 UPS 支撐，儲能專責普通電源備援至少 2 小時。\\n③ 下一頁轉場：看似完整的方案背後，招標文件實則存在 6 大潛在疑慮。"
      },
      {
        num: 8,
        tag: "RISK DIAGNOSIS ｜ 招標文件風險診斷",
        archetype: "BCG Forest Green ｜ 版型 8：BCG 2×2 矩陣 (BCG 2x2 Matrix)",
        badgeColor: "bg-emerald-950/80 text-emerald-400 border-emerald-500/30",
        actionTitle: "影響度 × 發生機率象限診斷：防逆送限制、工期衝突與無保底收益為致命痛點",
        subTitle: "以四象限策略矩陣，深度診斷招標文件中的 6 大潛在法律、營運與工程風險",
        img: "assets/images/slide_08_risk_fire_matrix.jpg",
        bullets: [
          { bold: "第一象限（致命衝擊 × 高機率）：", text: "① 防逆送限制阻礙電力交易大額收益；② 240/300 天工期衝突面臨逾期違約金" },
          { bold: "第二象限（重大衝擊 × 中機率）：", text: "③ 純按實際收益分成，無保底付款；④ 附記條款嚴格排陸，供應鏈穿透查核風險" },
          { bold: "第三象限（長期必然 × 中等衝擊）：", text: "⑤ 第 13 年需達 85% RTE，第 8~10 年需投入數千萬電芯置換 CAPEX" },
          { bold: "第四象限（合約退場 × 中等衝擊）：", text: "⑥ 契約期滿未續約時，廠商須全自費將儲能系統拆除清運、回復原狀" }
        ],
        notes: "【採用版型】：版型 8：BCG 2×2 矩陣 ｜ 視覺風格：BCG Forest Green\\n① 本頁結論：工期矛盾、防逆送限制與無保底機制是投標團隊的三大致命風險。\\n② 補充說明：必須透過投標前正式釋疑與建議書避險設計，化解這 6 大疑慮。\\n③ 下一頁轉場：其中最急迫的法律衝突，是招標文件中寫死的工期矛盾。"
      },
      {
        num: 9,
        tag: "TIMELINE ANALYSIS ｜ 工期矛盾與關鍵路徑",
        archetype: "Accenture Core Purple ｜ 版型 4：連鎖論法 (Chain / Flow)",
        badgeColor: "bg-purple-950/80 text-purple-400 border-purple-500/30",
        actionTitle: "台電審查與標檢局案場確證需 6~9 個月，240 天工期無容錯空間需立即啟動釋疑",
        subTitle: "對比契約文件條款矛盾，展開關鍵路徑 (CPM) 揭露潛在違約風險",
        img: "assets/images/slide_09_timeline_conflict.jpg",
        bullets: [
          { bold: "文件條款直接矛盾：", text: "契約首頁載明 240 日曆天完工，但契約第 7 條卻載明決標起 300 日內設置" },
          { bold: "關鍵路徑 CPM 耗時 280 天：", text: "台電內線審查 (60-90天) ＋ 消防審查 (45天) ＋ 標檢局確證 (60天) ＋ 設備交期與試車 (90天)" },
          { bold: "行政審查不可控：", text: "台電與標檢局審查屬於外部公權力，240 天工期若遇審查延宕，每日遭罰千分之一" },
          { bold: "法律行動建議：", text: "於等標期內依採購法正式行文請求機關「釋疑」，爭取以 300 天為準且扣除外審時間" }
        ],
        notes: "【採用版型】：版型 4：連鎖論法 ｜ 視覺風格：Accenture Core Purple\\n① 本頁結論：240 天與 300 天的矛盾關係到每日數十萬的違約金，必須於投標前發文釋疑。\\n② 補充說明：建議釋疑要求「台電外審及標檢局文件審查期間，非可歸責廠商者，應不計入工期」。\\n③ 下一頁轉場：在價格評選上，招標文件藏有一個精確的博弈滿分密碼。"
      },
      {
        num: 10,
        tag: "PRICING STRATEGY ｜ 收益分享比率博弈",
        archetype: "BCG Forest Green ｜ 版型 8：BCG 決策矩陣",
        badgeColor: "bg-emerald-950/80 text-emerald-400 border-emerald-500/30",
        actionTitle: "評選公式鎖定降至 81.70% 奪下 20 分滿分，兼顧評審高分與長期 12 年營收最大化",
        subTitle: "剖析評選須知評分公式，推導出奪標機率與財務收益的最佳平衡點",
        img: "assets/images/slide_10_pricing_gametheory.jpg",
        bullets: [
          { bold: "評分規則：", text: "廠商承諾比率上限 85.7%（得基本分 12 分），每降 0.5% 加 1 分，最高 20 分" },
          { bold: "報價 85.70%：", text: "僅得 12 分 ➔ 痛失 8 分關鍵評選分，排名必定大幅落後對手" },
          { bold: "報價 81.70% (黃金出價)：", text: "降 4.0% 穩拿 20 分滿分！保留 81.7% 高收益，獲利最大化" },
          { bold: "報價 80.00% 或更低：", text: "仍得 20 分（分數封頂不累加），白白損失未來 12 年數百萬拆分利潤" }
        ],
        notes: "【採用版型】：版型 8：BCG 決策矩陣 ｜ 視覺風格：BCG Forest Green\\n① 本頁結論：投標標價清單上收益分享比率精確填寫 81.70%，是唯一能拿滿分又保留最多利潤的解。\\n② 補充說明：多降無益（上限 20 分），少降失分，這是純數學博弈。\\n③ 下一頁轉場：結合技術、價格與法規，我們制定了五維必勝投標藍圖。"
      },
      {
        num: 11,
        tag: "EXECUTION BLUEPRINT ｜ 投標必勝五維藍圖",
        archetype: "McKinsey Deep Navy ｜ 版型 5：麥肯錫論法 (McKinsey Matrix)",
        badgeColor: "bg-blue-950/80 text-blue-400 border-blue-500/30",
        actionTitle: "資格組隊、技術客製、資安脫陸、合約釋疑與評選攻防之端到端落地方案",
        subTitle: "五大維度環環相扣，構築無可挑剔的投標服務建議書與評選實力",
        img: "assets/images/slide_11_five_pillars.jpg",
        bullets: [
          { bold: "維度一（資格與團隊）：", text: "能源技術代表領軍，結盟特定電力電業執照＋甲級承裝，協議書完成公證" },
          { bold: "維度二（定價與財務）：", text: "精準鎖定 81.70% 收益分享，建立 144 個月現金流模型與電芯更換折現" },
          { bold: "維度三（技術與工程）：", text: "污水場 C5-M 防腐、基座抬高 1 米、總 1 變電站無縫孤島控制與發電機互鎖" },
          { bold: "維度四（資安與安全）：", text: "非陸製供應鏈（台達/士林/西門子）穿透切結、NFPA 855 消防、標檢局 MOU" },
          { bold: "維度五（簡報與詢答）：", text: "聚焦院方「醫療零跳電」與專家「保護協調與防逆送」，進行 3D 兵棋推演" }
        ],
        notes: "【採用版型】：版型 5：麥肯錫論法 ｜ 視覺風格：McKinsey Deep Navy\\n① 本頁結論：五大維度無死角覆蓋，直接對應評選委員評分表各項子項指標。\\n② 補充說明：特別檢附電機技師與高級能管員證照，直取履約能力 15 分高分。\\n③ 下一頁轉場：最後，轉化為具體的下一步時程行動清單。"
      },
      {
        num: 12,
        tag: "ACTION ROADMAP ｜ 關鍵行動與時程管制",
        archetype: "Accenture Core Purple ｜ 版型 2：並列論法 (Parallel)",
        badgeColor: "bg-purple-950/80 text-purple-400 border-purple-500/30",
        actionTitle: "立即啟動「釋疑公文、公證協議、標檢局 MOU」三大作戰支柱，確保 100% 奪標",
        subTitle: "倒數計時推進，將全案投標準備轉化為立即可落地執行的實體行動",
        img: "assets/images/slide_12_partnership_nextsteps.jpg",
        bullets: [
          { bold: "作戰支柱 1（合約法律防禦）：", text: "Day 1~3 正式行文機關辦理招標文件釋疑，釐清 240/300 天工期衝突與防逆送電力交易認列" },
          { bold: "作戰支柱 2（團隊執照鎖定）：", text: "Day 4~7 簽署共同投標協議書並赴公證處認證，備妥特定電力執照、技師證照與 1.75 億元實績證明" },
          { bold: "作戰支柱 3（技術藍圖與評選演練）：", text: "Day 8~14 完成 C5-M 污水場配置圖、簽署標檢局驗證機構 MOU、編印 15 份建議書並實兵推演" }
        ],
        notes: "【採用版型】：版型 2：並列論法 ｜ 視覺風格：Accenture Core Purple\\n① 本頁結論：分秒必爭，三大作戰支柱並行推進，奪取高雄榮總指標性醫療級儲能案。\\n② 補充說明：進金生能源具備完整電力工程實績與綠能統包實力，有信心打造全台標竿。\\n③ 簡報結束：感謝聆聽，敬請評審長官指教！"
      }
    ];

    let currentSlide = 0;

    function renderSlide(index) {
      if (index < 0) index = 0;
      if (index >= slidesData.length) index = slidesData.length - 1;
      currentSlide = index;

      const slide = slidesData[currentSlide];

      // Archetype Badge
      const badge = document.getElementById("slide-archetype-badge");
      badge.textContent = `${slide.tag} ｜ ${slide.archetype}`;
      badge.className = `px-2.5 py-1 rounded-md text-xs font-bold uppercase tracking-wider ${slide.badgeColor}`;

      // Counter Badge
      document.getElementById("slide-counter-badge").textContent = `${String(slide.num).padStart(2, '0')} / 12`;

      // Titles
      document.getElementById("slide-action-title").textContent = slide.actionTitle;
      document.getElementById("slide-sub-title").textContent = slide.subTitle;

      // Image
      document.getElementById("slide-img").src = slide.img;

      // Evidence Bullets
      const container = document.getElementById("slide-evidence-container");
      container.innerHTML = "";
      slide.bullets.forEach(item => {
        const div = document.createElement("div");
        div.className = "p-3.5 bg-slate-900/90 rounded-lg border border-slate-800/80 hover:border-slate-700 transition flex items-start gap-2.5";
        div.innerHTML = `
          <div class="w-2 h-2 rounded-full bg-amber-400 mt-2 shrink-0"></div>
          <div class="text-xs sm:text-sm text-slate-200 leading-relaxed">
            <span class="font-bold text-amber-300">${item.bold}</span>
            <span class="text-slate-300">${item.text}</span>
          </div>
        `;
        container.appendChild(div);
      });

      // Notes
      document.getElementById("notes-content").textContent = slide.notes;

      // Update Dots
      renderDots();
      // Update Thumbnails
      renderThumbnails();
    }

    function renderDots() {
      const container = document.getElementById("slide-dots");
      container.innerHTML = "";
      slidesData.forEach((_, i) => {
        const dot = document.createElement("button");
        dot.className = `w-2 h-2 rounded-full transition-all ${i === currentSlide ? 'bg-amber-400 w-6' : 'bg-slate-700 hover:bg-slate-500'}`;
        dot.onclick = () => renderSlide(i);
        container.appendChild(dot);
      });
    }

    function renderThumbnails() {
      const container = document.getElementById("thumbnails-container");
      container.innerHTML = "";
      slidesData.forEach((s, i) => {
        const thumb = document.createElement("div");
        thumb.className = `cursor-pointer rounded-lg border p-1 text-center transition flex flex-col items-center justify-center ${i === currentSlide ? 'border-amber-400 bg-amber-950/40 shadow-md shadow-amber-500/10' : 'border-slate-800 bg-slate-900/60 hover:border-slate-600'}`;
        thumb.onclick = () => renderSlide(i);
        thumb.innerHTML = `
          <div class="text-[10px] font-mono font-bold ${i === currentSlide ? 'text-amber-400' : 'text-slate-400'}">P.${String(s.num).padStart(2, '0')}</div>
          <div class="text-[9px] truncate max-w-[80px] text-slate-300 mt-0.5">${s.tag.split(' ｜ ')[0]}</div>
        `;
        container.appendChild(thumb);
      });
    }

    function prevSlide() {
      if (currentSlide > 0) renderSlide(currentSlide - 1);
    }

    function nextSlide() {
      if (currentSlide < slidesData.length - 1) renderSlide(currentSlide + 1);
    }

    function toggleNotes() {
      const drawer = document.getElementById("notes-drawer");
      drawer.classList.toggle("hidden");
    }

    function toggleModal(id) {
      const modal = document.getElementById(id);
      modal.classList.toggle("hidden");
    }

    function openLightbox(src) {
      document.getElementById("lightbox-img").src = src;
      document.getElementById("lightbox-modal").classList.remove("hidden");
    }

    function closeLightbox() {
      document.getElementById("lightbox-modal").classList.add("hidden");
    }

    function toggleFullScreen() {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
      } else {
        if (document.exitFullscreen) {
          document.exitFullscreen();
        }
      }
    }

    // Keyboard navigation
    document.addEventListener("keydown", (e) => {
      if (e.key === "ArrowRight" || e.key === " " || e.key === "PageDown") {
        nextSlide();
      } else if (e.key === "ArrowLeft" || e.key === "PageUp") {
        prevSlide();
      } else if (e.key === "Escape") {
        closeLightbox();
        document.getElementById("calc-modal").classList.add("hidden");
      }
    });

    // Calculator function
    function updateCalculator(val) {
      val = parseFloat(val);
      document.getElementById("calc-pct-label").textContent = val.toFixed(2) + "%";

      // Score calculation:
      // Base: 85.7% -> 12 points
      // Each 0.5% lower -> +1 point, max 20 points
      const diff = 85.7 - val;
      const steps = Math.floor((diff + 0.0001) / 0.5);
      let score = 12 + steps;
      if (score > 20) score = 20;
      if (score < 12) score = 12;

      const scoreEl = document.getElementById("calc-score");
      const scoreSubEl = document.getElementById("calc-score-sub");
      scoreEl.textContent = `${score} / 20 分`;

      if (score === 20) {
        scoreSubEl.textContent = "★ 獲得滿分最高評價 (極限最佳解)";
        scoreSubEl.className = "text-[11px] text-emerald-400 font-semibold";
      } else if (score >= 16) {
        scoreSubEl.textContent = "✔ 良好競爭力，但仍落後滿分 " + (20 - score) + " 分";
        scoreSubEl.className = "text-[11px] text-amber-400";
      } else {
        scoreSubEl.textContent = "⚠ 嚴重落後滿分 " + (20 - score) + " 分，極高淘汰風險";
        scoreSubEl.className = "text-[11px] text-rose-400";
      }

      // Money calculations (Monthly estimated revenue: 3.55M NTD)
      const monthlyRev = 3550000 * (val / 100);
      document.getElementById("calc-monthly").textContent = (monthlyRev / 10000).toFixed(1) + " 萬元";

      const total12yr = monthlyRev * 144;
      document.getElementById("calc-total").textContent = (total12yr / 100000000).toFixed(2) + " 億元";

      const total20yr = monthlyRev * 240;
      document.getElementById("calc-life").textContent = (total20yr / 100000000).toFixed(2) + " 億元";

      const baseMonthly = 3550000 * (85.7 / 100);
      const totalBase = baseMonthly * 144;
      const lost12yr = total12yr - totalBase;
      const pctDiff = val - 85.7;

      document.getElementById("calc-diff").textContent = `${pctDiff.toFixed(2)}% (${(lost12yr / 10000).toFixed(0)}萬元)`;
    }

    // Initialize
    window.addEventListener("DOMContentLoaded", () => {
      lucide.createIcons();
      renderSlide(0);
      updateCalculator(81.7);
    });
  </script>
</body>
</html>"""

with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Created index.html: {html_path}")
