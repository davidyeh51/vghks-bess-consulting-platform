import os
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

base_dir = r"G:\我的雲端硬碟\3_大衛工作\2_進金生能源\8.0 專案機會\8.2 儲能工程\8.2.3 錶後儲能：用電大戶\高雄_榮總醫院\vghks-bess-consulting-platform"
img_dir = os.path.join(base_dir, "assets", "images")
output_pptx = os.path.join(base_dir, "vghks_bess_zero_noise_consulting.pptx")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Consulting Color Palette
NAVY = RGBColor(5, 28, 44)
ROYAL = RGBColor(30, 58, 138)
GOLD = RGBColor(217, 119, 6)
WHITE = RGBColor(255, 255, 255)
DARK_GRAY = RGBColor(31, 41, 55)
LIGHT_GRAY = RGBColor(243, 244, 246)
BORDER_GRAY = RGBColor(203, 213, 225)
GREEN = RGBColor(0, 79, 50)
LIME = RGBColor(41, 186, 116)
PURPLE = RGBColor(161, 0, 255)
CYAN = RGBColor(6, 182, 212)
CARD_BG = RGBColor(248, 250, 252)

slides_data = [
    {
        "num": 1,
        "tag": "PROJECT OVERVIEW ｜ 決策旗艦報告",
        "archetype": "封面頁 ｜ McKinsey Deep Navy",
        "action_title": "高雄榮民總醫院「表後儲能系統建置及節能技術服務」",
        "sub": "採購案號：ENG-1150058 ｜ 4.38 億元巨額採購案決策分析與投標必勝攻略",
        "img": "slide_01_cover_campus.jpg",
        "bullets": [
            "■ 標案性質：巨額勞務採購（限制性招標，公開評選準用最有利標，序位法決標）",
            "■ 預算規模：本期預算 4.38 億元（12年），含兩次擴充總採購金額上看 7.30 億元",
            "■ 商業模式：純 ESCO / BOO（廠商全額投資建置與維運，按月實際效益分成）",
            "■ 核心要求：中斷供電供應普通電源≧2小時、充放電效率≧85%、嚴格防逆送"
        ],
        "notes": "【採用版型】：封面頁 ｜ 視覺風格：McKinsey Deep Navy\n① 本頁結論：高雄榮總表後儲能案為 4.38 億之巨額純 ESCO 標案，投標決勝在於控制工期與醫療供電安全。\n② 補充說明：本報告提供招標重點摘要、6大隱藏風險疑慮，以及確保奪標的五維準備方案。\n③ 下一頁轉場：首先呈現全案的核心決策總覽。"
    },
    {
        "num": 2,
        "tag": "EXECUTIVE SUMMARY ｜ 決策架構總覽",
        "archetype": "版型 3：結合論法 (Combined Tree) ｜ McKinsey Deep Navy",
        "action_title": "「純 ESCO 全額出資」結合「醫療級 2 小時備援」與「81.7% 定價博弈」，鎖定 4.38 億案場",
        "sub": "三大核心前提同時成立，推導出進金生能源必勝奪標之頂層策略",
        "img": "slide_02_bess_overview.jpg",
        "bullets": [
            "【商業支柱】：純 ESCO 效益分成模式，廠商承擔全部 CAPEX，依每月實際收益支付",
            "【技術支柱】：污水場 660m² 防腐防汛、普通電源 2 小時孤島供電、全非陸製資通訊",
            "【博弈支柱】：收益分享比率鎖定 81.70% 奪滿分 20 分，兼顧評選第 1 與營收最大化",
            "【實績門檻】：特定電力供應業電業執照 ＋ 5年內單次 1.75 億或累計 4.38 億實績"
        ],
        "notes": "【採用版型】：版型 3：結合論法 ｜ 視覺風格：McKinsey Deep Navy\n① 本頁結論：三大支柱必須同時到位才能確保拿下此巨額案場。\n② 補充說明：商業模式需耐受無保底風險，工程需克服污水場腐蝕，定價須精準卡位滿分。\n③ 下一頁轉場：接著檢視招標文件所載的合約採購規模與擴充機制。"
    },
    {
        "num": 3,
        "tag": "SCALE & EXPANSION ｜ 採購規模與擴充",
        "archetype": "版型 1：一對一論法 (One-to-One) ｜ Accenture Core Purple",
        "action_title": "本期預算 4.38 億元加上兩次續約擴充，全生命週期採購總規模上看 7.30 億元",
        "sub": "單一強力核心指標：本案為橫跨 20 年之長生命週期巨額能源技術服務",
        "img": "slide_03_contract_scale.jpg",
        "bullets": [
            "【本期 12 年 (144 個月)】：預算上限 4億3,809萬8,400 元（每月預估 355 萬 × 85.7%）",
            "【第 1 次續約 (第 13~16 年)】：增購上限 1億4,603萬2,800 元（續約門檻：RTE ≧ 85%）",
            "【第 2 次續約 (第 17~20 年)】：增購上限 1億4,603萬2,800 元（續約門檻：RTE ≧ 80%）",
            "【全案總額】：本期加擴充合計達 7億3,016萬4,000 元；履約保證金 400 萬元"
        ],
        "notes": "【採用版型】：版型 1：一對一論法 ｜ 視覺風格：Accenture Core Purple\n① 本頁結論：這是一檔長達 20 年、總額高達 7.3 億元的巨額技術服務合約。\n② 補充說明：第 13 年仍要求 85% 效率，投標團隊必須在中期編列電芯置換 CAPEX。\n③ 下一頁轉場：接下來剖析合約明訂的商業模式與計價公式。"
    },
    {
        "num": 4,
        "tag": "BUSINESS MODEL ｜ 商業模式與計價",
        "archetype": "版型 5：麥肯錫論法 (McKinsey Matrix) ｜ McKinsey Deep Navy",
        "action_title": "「零期初工程款 ＋ 每月實際收益拆分」，廠商承擔全額 CAPEX 與回收風險",
        "sub": "沿著生命週期流程，診斷契約計價機制與廠商財務護欄",
        "img": "slide_04_revenue_ems.jpg",
        "bullets": [
            "【建置階段】：廠商全額自費出資設備、土建、消防與併聯引接，院方無期初工程款",
            "【運轉調度】：EMS 實施尖離峰削峰填谷，嚴格落實防逆送保護，維持 ≧ 85% RTE",
            "【收益公式】：（尖峰電價×供電量）－（離峰電價×充電量）＋ 輔助服務／需量反應收益",
            "【風險對策】：無保底機制（No Take-or-Pay），需藉由高榮白天大基載確保放電全額吸收"
        ],
        "notes": "【採用版型】：版型 5：麥肯錫論法 ｜ 視覺風格：McKinsey Deep Navy\n① 本頁結論：合約是純 ESCO 模式，廠商承擔所有投資與運轉責任，按實際效益分成。\n② 補充說明：每月款項需經書面驗收，計量電錶必須具備經濟部標檢局檢定合格。\n③ 下一頁轉場：要支撐此效益，必須深入高榮院區的真實用電結構。"
    },
    {
        "num": 5,
        "tag": "LOAD & CAPACITY ｜ 用電分析與容量",
        "archetype": "版型 1：一對一論法 (One-to-One) ｜ BCG Forest Green",
        "action_title": "契約容量 8,300 kW 搭配夏月超約特性，精算 3~4 MW / 8~10 MWh 為最佳化投標容量",
        "sub": "以高榮 19 個月實測用電大數據，反推最佳經濟效益與用地平衡點",
        "img": "slide_05_load_capacity.jpg",
        "bullets": [
            "【用電總量】：114/01~115/07 用電 7,435 萬度，電費 3.06 億元，平均單價 4.12 元/度",
            "【負載分佈】：空調設備佔 54.0%、動力設備佔 17.6%、照明佔 15.8%，夏季下午極度吃重",
            "【超約痛點】：114 年 5~7 月連續產生超約附加費（累計逾 31 萬元），削峰急迫性極高",
            "【最佳配置】：3~4 MW / 8~10 MWh 系統可平抑夏月尖峰，日放電 1.2 萬度，直逼效益目標"
        ],
        "notes": "【採用版型】：版型 1：一對一論法 ｜ 視覺風格：BCG Forest Green\n① 本頁結論：高榮用電基載大且夏月空調尖峰明顯，3~4 MW 儲能最具經濟削峰效益。\n② 補充說明：污水處理場 660 m² 土地面積亦剛好可容納約 2~3 座 40 呎集裝箱儲能櫃。\n③ 下一頁轉場：然而，污水場現場環境十分嚴苛，工程防護必須做足。"
    },
    {
        "num": 6,
        "tag": "SITE CONDITIONS ｜ 場址與工程防護",
        "archetype": "版型 6：對立論法 (Opposition) ｜ McKinsey Deep Navy",
        "action_title": "污水處理場 660 m² 硫化氫腐蝕環境，全面導入 C5-M 重防腐與基礎防汛抬高工程",
        "sub": "針對污水廠嚴酷環境挑戰，逐項提出對應之重工業級工程防護對策",
        "img": "slide_06_sewage_civil.jpg",
        "bullets": [
            "【化學腐蝕 vs. C5-M 塗裝】：污水場高濃度 H2S 與 NH3 氣體 ➔ 貨櫃外殼 C5-M 海洋重防腐，空調加裝化學吸附濾網",
            "【地下管溝 vs. 筏式微型樁】：場址地下密布大型集水槽與管線 ➔ 採用淺層筏式基礎搭配微型抗拔樁，避開管溝",
            "【防汛積水 vs. 基座墊高】：高雄夏季常有強降雨積淹水疑慮 ➔ 混凝土設備基座墊高 800~1200mm，配置雙迴路排水",
            "【動線維持 vs. 模組化佈局】：契約規定維持污水廠維修通行 ➔ 預留 4 米法定應急救援通道，滿足消防指引要求"
        ],
        "notes": "【採用版型】：版型 6：對立論法 ｜ 視覺風格：McKinsey Deep Navy\n① 本頁結論：污水場是高腐蝕與受限空間，非一般空地，防腐與防汛設計是審標硬指標。\n② 補充說明：若未在企劃書提出 C5-M 防腐與基座墊高，將被工務室委員重扣分數。\n③ 下一頁轉場：除土建外，醫院最重視的核心功能是市電斷電時的電力救援切換。"
    },
    {
        "num": 7,
        "tag": "RESILIENCE & MICROGRID ｜ 醫療供電韌性",
        "archetype": "版型 4：連鎖論法 (Chain / Flow) ｜ Accenture Core Purple",
        "action_title": "市電中斷瞬時切換孤島，確保普通電源 2 小時續航並與柴油發電機防衝擊協同",
        "sub": "端到端展示台電斷電時之毫秒級自動切換與發電機互鎖時序流程",
        "img": "slide_07_resilience_switch.jpg",
        "bullets": [
            "【市電中斷 ➔ 瞬時解聯】：台電電網斷電時，高壓斷路器 200ms 快速脫扣，形成孤島隔離",
            "【電網形成 ➔ 普通電源 2 小時】：PCS 轉為 VF 電壓源模式，維持全院照明、電梯、空調等普通電源",
            "【平時調度 ➔ 最低 SOC 護欄】：EMS 平時調度鎖定最低備援荷電量 ≧ 30%，保證隨時有 2 小時電量",
            "【發電機協同 ➔ 防反充保護】：儲能系統與既設柴發並聯時增設逆功率防護電驛，防止發電機反充跳脫"
        ],
        "notes": "【採用版型】：版型 4：連鎖論法 ｜ 視覺風格：Accenture Core Purple\n① 本頁結論：微電網孤島切換與防發電機逆功率反充是醫療韌性技術關鍵。\n② 補充說明：重症 ICU 與手術室有專屬 UPS 支撐，儲能專責普通電源備援至少 2 小時。\n③ 下一頁轉場：看似完整的方案背後，招標文件實則存在 6 大潛在疑慮。"
    },
    {
        "num": 8,
        "tag": "RISK DIAGNOSIS ｜ 招標文件風險診斷",
        "archetype": "版型 8：BCG 2×2 矩陣 (BCG 2x2 Matrix) ｜ BCG Forest Green",
        "action_title": "影響度 × 發生機率象限診斷：防逆送限制、工期衝突與無保底收益為致命痛點",
        "sub": "以四象限策略矩陣，深度診斷招標文件中的 6 大潛在法律、營運與工程風險",
        "img": "slide_08_risk_fire_matrix.jpg",
        "bullets": [
            "【第一象限：致命衝擊 × 高機率 (高亮焦點)】：① 防逆送限制阻礙電力交易大額收益；② 240/300 天工期衝突面臨逾期違約金",
            "【第二象限：重大衝擊 × 中機率】：③ 純按實際收益分成，無保底付款；④ 附記條款嚴格排陸，供應鏈穿透查核風險",
            "【第三象限：長期必然 × 中等衝擊】：⑤ 第 13 年需達 85% RTE，第 8~10 年需投入數千萬電芯置換 CAPEX",
            "【第四象限：合約退場 × 中等衝擊】：⑥ 契約期滿未續約時，廠商須全自費將儲能系統拆除清運、回復原狀"
        ],
        "notes": "【採用版型】：版型 8：BCG 2×2 矩陣 ｜ 視覺風格：BCG Forest Green\n① 本頁結論：工期矛盾、防逆送限制與無保底機制是投標團隊的三大致命風險。\n② 補充說明：必須透過投標前正式釋疑與建議書避險設計，化解這 6 大疑慮。\n③ 下一頁轉場：其中最急迫的法律衝突，是招標文件中寫死的工期矛盾。"
    },
    {
        "num": 9,
        "tag": "TIMELINE ANALYSIS ｜ 工期矛盾與關鍵路徑",
        "archetype": "版型 4：連鎖論法 (Chain / Flow) ｜ Accenture Core Purple",
        "action_title": "台電審查與標檢局案場確證需 6~9 個月，240 天工期無容錯空間需立即啟動釋疑",
        "sub": "對比契約文件條款矛盾，展開關鍵路徑 (CPM) 揭露潛在違約風險",
        "img": "slide_09_timeline_conflict.jpg",
        "bullets": [
            "【文件條款直接矛盾】：契約首頁載明 240 日曆天完工，但契約第 7 條卻載明決標起 300 日內設置",
            "【關鍵路徑 CPM 耗時 280 天】：台電內線審查 (60-90天) ＋ 消防審查 (45天) ＋ 標檢局確證 (60天) ＋ 設備交期與試車 (90天)",
            "【行政審查不可控】：台電與標檢局審查屬於外部公權力，240 天工期若遇審查延宕，每日遭罰千分之一",
            "【法律行動建議】：於等標期內依採購法正式行文請求機關「釋疑」，爭取以 300 天為準且扣除外審時間"
        ],
        "notes": "【採用版型】：版型 4：連鎖論法 ｜ 視覺風格：Accenture Core Purple\n① 本頁結論：240 天與 300 天的矛盾關係到每日數十萬的違約金，必須於投標前發文釋疑。\n② 補充說明：建議釋疑要求「台電外審及標檢局文件審查期間，非可歸責廠商者，應不計入工期」。\n③ 下一頁轉場：在價格評選上，招標文件藏有一個精確的博弈滿分密碼。"
    },
    {
        "num": 10,
        "tag": "PRICING STRATEGY ｜ 收益分享比率博弈",
        "archetype": "版型 8：BCG 決策矩陣 ｜ BCG Forest Green",
        "action_title": "評選公式鎖定降至 81.70% 奪下 20 分滿分，兼顧評審高分與長期 12 年營收最大化",
        "sub": "剖析評選須知評分公式，推導出奪標機率與財務收益的最佳平衡點",
        "img": "slide_10_pricing_gametheory.jpg",
        "bullets": [
            "【評分規則】：廠商承諾比率上限 85.7%（得基本分 12 分），每降 0.5% 加 1 分，最高 20 分",
            "【報價 85.70%】：僅得 12 分 ➔ 痛失 8 分關鍵評選分，排名必定大幅落後對手",
            "【報價 81.70% (黃金出價)】：降 4.0% 穩拿 20 分滿分！保留 81.7% 高收益，獲利最大化",
            "【報價 80.00% 或更低】：仍得 20 分（分數封頂不累加），白白損失未來 12 年數百萬拆分利潤"
        ],
        "notes": "【採用版型】：版型 8：BCG 決策矩陣 ｜ 視覺風格：BCG Forest Green\n① 本頁結論：投標標價清單上收益分享比率精確填寫 81.70%，是唯一能拿滿分又保留最多利潤的解。\n② 補充說明：多降無益（上限 20 分），少降失分，這是純數學博弈。\n③ 下一頁轉場：結合技術、價格與法規，我們制定了五維必勝投標藍圖。"
    },
    {
        "num": 11,
        "tag": "EXECUTION BLUEPRINT ｜ 投標必勝五維藍圖",
        "archetype": "版型 5：麥肯錫論法 (McKinsey Matrix) ｜ McKinsey Deep Navy",
        "action_title": "資格組隊、技術客製、資安脫陸、合約釋疑與評選攻防之端到端落地方案",
        "sub": "五大維度環環相扣，構築無可挑剔的投標服務建議書與評選實力",
        "img": "slide_11_five_pillars.jpg",
        "bullets": [
            "【維度一：資格與團隊】：能源技術代表領軍，結盟特定電力電業執照＋甲級承裝，協議書完成公證",
            "【維度二：定價與財務】：精準鎖定 81.70% 收益分享，建立 144 個月現金流模型與電芯更換折現",
            "【維度三：技術與工程】：污水場 C5-M 防腐、基座抬高 1 米、總 1 變電站無縫孤島控制與發電機互鎖",
            "【維度四：資安與安全】：非陸製供應鏈（台達/士林/西門子）穿透切結、NFPA 855 消防、標檢局 MOU",
            "【維度五：簡報與詢答】：聚焦院方「醫療零跳電」與專家「保護協調與防逆送」，進行 3D 兵棋推演"
        ],
        "notes": "【採用版型】：版型 5：麥肯錫論法 ｜ 視覺風格：McKinsey Deep Navy\n① 本頁結論：五大維度無死角覆蓋，直接對應評選委員評分表各項子項指標。\n② 補充說明：特別檢附電機技師與高級能管員證照，直取履約能力 15 分高分。\n③ 下一頁轉場：最後，轉化為具體的下一步時程行動清單。"
    },
    {
        "num": 12,
        "tag": "ACTION ROADMAP ｜ 關鍵行動與時程管制",
        "archetype": "版型 2：並列論法 (Parallel) ｜ Accenture Core Purple",
        "action_title": "立即啟動「釋疑公文、公證協議、標檢局 MOU」三大作戰支柱，確保 100% 奪標",
        "sub": "倒數計時推進，將全案投標準備轉化為立即可落地執行的實體行動",
        "img": "slide_12_partnership_nextsteps.jpg",
        "bullets": [
            "【作戰支柱 1：合約法律防禦 (Day 1~3)】：正式行文機關辦理招標文件釋疑，釐清 240/300 天工期衝突與防逆送電力交易認列",
            "【作戰支柱 2：團隊執照鎖定 (Day 4~7)】：簽署共同投標協議書並赴公證處認證，備妥特定電力執照、技師證照與 1.75 億元實績證明",
            "【作戰支柱 3：技術藍圖與評選演練 (Day 8~14)】：完成 C5-M 污水場配置圖、簽署標檢局驗證機構 MOU、編印 15 份建議書並實兵推演"
        ],
        "notes": "【採用版型】：版型 2：並列論法 ｜ 視覺風格：Accenture Core Purple\n① 本頁結論：分秒必爭，三大作戰支柱並行推進，奪取高雄榮總指標性醫療級儲能案。\n② 補充說明：進金生能源具備完整電力工程實績與綠能統包實力，有信心打造全台標竿。\n③ 簡報結束：感謝聆聽，敬請評審長官指教！"
    }
]

for s in slides_data:
    slide = prs.slides.add_slide(blank_layout)
    
    # 1. Background fill
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = CARD_BG
    bg.line.fill.background()
    
    # Top Category Tag
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8), Inches(0.35))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = f"{s['tag']} ｜ {s['archetype']}"
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = ROYAL if "McKinsey" in s["archetype"] else (GREEN if "BCG" in s["archetype"] else PURPLE)
    
    # Action Title (Top-down conclusion)
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.85))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = s["action_title"]
    p_title.font.size = Pt(19)
    p_title.font.bold = True
    p_title.font.color.rgb = NAVY
    
    # Subtitle / Context
    p_sub = tf_title.add_paragraph()
    p_sub.text = s["sub"]
    p_sub.font.size = Pt(12)
    p_sub.font.color.rgb = DARK_GRAY
    
    # Divider line
    div_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.68), Inches(11.733), Inches(0.02))
    div_line.fill.solid()
    div_line.fill.fore_color.rgb = BORDER_GRAY
    div_line.line.fill.background()
    
    # Left Box / Evidence Content Cards
    card_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.85), Inches(6.8), Inches(5.1))
    card_shape.fill.solid()
    card_shape.fill.fore_color.rgb = WHITE
    card_shape.line.color.rgb = BORDER_GRAY
    card_shape.line.width = Pt(1)
    
    # Card Text
    text_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(6.4), Inches(4.9))
    tf_card = text_box.text_frame
    tf_card.word_wrap = True
    
    for idx, b in enumerate(s["bullets"]):
        p_b = tf_card.paragraphs[0] if idx == 0 else tf_card.add_paragraph()
        p_b.text = b
        p_b.font.size = Pt(12)
        p_b.font.color.rgb = DARK_GRAY
        p_b.space_before = Pt(8)
        p_b.space_after = Pt(8)
    
    # Right Box: Dedicated Image for Every Slide
    img_filename = s["img"]
    img_path = os.path.join(img_dir, img_filename)
    if os.path.exists(img_path):
        img_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(1.85), Inches(4.733), Inches(5.1))
        img_card.fill.solid()
        img_card.fill.fore_color.rgb = WHITE
        img_card.line.color.rgb = BORDER_GRAY
        img_card.line.width = Pt(1)
        
        slide.shapes.add_picture(img_path, Inches(7.9), Inches(1.95), width=Inches(4.533), height=Inches(4.9))
    
    # Bottom Source / Confidentiality Footnote
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.3))
    tf_foot = foot_box.text_frame
    p_foot = tf_foot.paragraphs[0]
    p_foot.text = f"Source: 高雄榮民總醫院 ENG-1150058 招標文件 ｜ 投標決策專用 ｜ Page {s['num']} of 12"
    p_foot.font.size = Pt(9)
    p_foot.font.color.rgb = BORDER_GRAY
    
    # Speaker Notes
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = s["notes"]

prs.save(output_pptx)
print(f"Successfully generated PPTX: {output_pptx}")

parent_pptx = os.path.join(r"G:\我的雲端硬碟\3_大衛工作\2_進金生能源\8.0 專案機會\8.2 儲能工程\8.2.3 錶後儲能：用電大戶\高雄_榮總醫院", "高雄榮總_表後儲能建置及節能技術服務_無雜訊顧問簡報.pptx")
prs.save(parent_pptx)
print(f"Successfully saved copy to: {parent_pptx}")
