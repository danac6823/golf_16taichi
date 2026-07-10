# CLAUDE.md — 意象太極高爾夫球隊 LIFF 系統

## 專案概觀
LINE LIFF 高爾夫球隊管理系統(第16屆)。單一 Google Sheet 為資料庫,GAS 為後端(JSONP),前端單檔 vanilla JS 部署 GitHub Pages。使用者:總幹事張凌華(管理員/開發者)、財務長張曼華、53 位會員。介面語言:繁體中文(台灣)。

- 分享連結:https://liff.line.me/2010396373-bdFy7e3d(LIFF_ID 寫在兩檔內)
- 前端:danac6823.github.io/fengsheng(index.html,重新上傳後必換 `?v=`)
- 後端:GAS Web App(Code.gs,**存檔≠部署,必須「部署新版本」**)

## 檔案
| 檔案 | 說明 |
|---|---|
| `Code.gs` | 後端全部。`APP_VERSION` 每次改動要 bump(格式 vYYYY.MM.DD-tag);`?action=ver` 可驗證部署 |
| `index.html` | 前端全部(含 CSS/JS/43KB base64 logo)。狀態機:`state` 物件 + `renderMain()` 全量重繪 + `wireMain()` 綁事件 |
| `test_harness.js` | Node 測試(28 條)。前 63 行是 mock sheets 沙箱,可 `vm` 載入後直呼後端函式 |
| `build_preview.py` | 產 `index_preview.html`:離線假後端(跑真 Code.gs)+ 三身分切換列(會員郭土水/管理員張凌華/財務長張曼華 TREASURER_UID) |
| `部署清單.md` | 部署步驟、測試清單、待辦 |

## 驗證儀式(每次改動必跑)
```bash
node -e "new Function(require('fs').readFileSync('Code.gs','utf8'));console.log('OK')"   # GAS 語法
node -e '…extract non-src <script> blocks… new Function(code)'                            # 前端語法
node test_harness.js        # 28/28
python3 build_preview.py    # 重建預覽
# PDF 改動時:wkhtmltopdf --page-size A4 → pypdf 數頁數,必須 1 頁
```
沙箱測試模板:`vm` 載入 harness 前 63 行 → `sandbox`;寫入型測試需注入
`sb.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})}` 並 patch
`G.sheet(名).deleteRows=function(st,c){this._rows.splice(st-1,c);}`。

## 資料分頁(自動建立)
Roster/Members(userId|name|role|gender|birthYear|phone|lineName)/會員名單(姓名|性別|年初差點|出生年|會員身份|狀態←自動建)/Scores/技術獎(日期|獎項|得獎者|條數)/Payments/行事曆/Config/報名紀錄(封存)/公告/採購紀錄(場次|類別|品項|單價(棄用)|數量|小計=金額|請款|已請款|ts)/帳務(月份|本月收入|餐費|其他|備註|更新時間|其他說明←自動補欄)/加碼獎(場次|獎項|得獎者|金額|出資者|計入球隊v|建立時間)

## 關鍵語意規則(勿破壞)
1. **採購**:列「請款」勾=向球隊請款 → 才算球隊支出(monthPurchaseTotal_ 只加勾選列);沒勾=總幹事自付招待。「已請款」(整場)=請款完成回填,控制首頁「⚠ 待請款」提醒。
2. **加碼獎**:逐筆「計入球隊」勾選決定入帳(monthAddonTeamTotal_);出資者是誰不決定入帳(會長可計入也可自付,併存)。日期**跟隨成績頁「比賽日期」欄(smDate)**。
3. **帳務**:模式 B(每月可填收入);起始月/起始餘額在 Config(管理員設);獎金 prizeFor()×Scores 自動;餘額逐月結轉。isTreasurer_ = Members 綁定姓名 === cfg.treasurer。
4. **會員狀態**:getMasterList() 只回有效會員(排除 停權/退隊)→ 下游(尚未登記/會費/代報名)自動排除;masterAll_() 回全部(會員管理用)。改名同步 Members+本場名單,歷史 Scores 保留舊名。
5. **送球歸送球、金錢歸金錢**:送球(技術獎條數)只在管理員送球庫存統計,不進財務。
6. **PDF 單頁**:有加碼 → LOW=false 密排 + FS≤13;洞洞有獎同(金額|出資|計入)合併一列;技術獎同獎項合併。改版面必跑 8/20/30/40 人 × 有無加碼回歸。
7. 成績重送=覆蓋(清該場 Scores+技術獎,還原差點);加碼獎獨立不清。

## 待辦(見部署清單.md 第五節)
C8 通訊錄匯出/請款拋轉 Phase 2/會員繳費移交財務長。

## 工作習慣
計畫先行、小步交付、每步驗證;回覆繁中、先講結論;每次改 Code.gs 要 bump APP_VERSION 並提醒「兩檔部署 + 換 ?v=」。
