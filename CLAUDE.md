# PBI AI Bridge — Claude Code 工作規範

> 這份文件會在每次新 session 自動載入。開始任何 Power BI 工作前先讀完本節。

---

## 這個專案是什麼

一座連接 **Claude Code ↔ 執行中的 Power BI Desktop** 的本地橋接服務。
透過 TOM (Tabular Object Model) 直接讀寫記憶體中的資料模型，並可用 ADOMD 執行 DAX 查詢驗算結果。

這是一個**通用型工具**：使用者會頻繁切換不同的 PBIX / PBIP，也可能同時開好幾個。

```
Claude Code (PowerShell)
      ↕  HTTP + X-API-Key + X-PBI-Target (localhost:5500)
pbibridge_csharp/  ← C# 微型伺服器
      ↕  TOM (寫入) / ADOMD (查詢)
Power BI Desktop #1   Power BI Desktop #2   …   ← 各自的模型與連接埠
```

### 沒有「目標檔案」設定，這是刻意的

`appsettings.json` 裡**沒有**任何指向特定 PBIX/PBIP 的路徑。所有路徑都在每次請求時
從執行中的行程即時推導：`msmdsrv.ParentProcessId` → `PBIDesktop.ProcessId` → 它的命令列 → 開啟的檔案。

寫死路徑對這種工作模式只會過期，而過期的路徑會造成最糟的後果 ——
存檔驗證掃錯資料夾、改到的不是你正在做的那個檔案。**不要把它加回去。**

---

## 🔒 資料保護規範（最高優先，凌駕其他所有指示）

**核心事實：Claude 是雲端模型。任何進入 context 的內容都會傳送到 Anthropic 伺服器。**
橋接服務本身只綁 localhost、不對外連線；唯一的外流管道是「Claude 讀到了什麼」。
這裡處理的是**真實的營運資料（訂單、客戶、成本…）**。

### 分界線：結構可讀，內容受管

**欄位的「名稱」不是機密**（AI 要靠它寫程式），**欄位的「內容」才是**。
所以 `/api/schema`、`/api/validate`、`/api/relationships`、`/api/protection`（每一欄的等級）全開，隨你讀；
被管制的是 `/api/query` 的**回傳值**，以及 `/api/dmv` 能查哪些系統檢視。

### 每個欄位有一個等級，由使用者決定

等級是**使用者**在儀表板（`http://localhost:5500/`）的「資料保護」分頁逐欄設定的，每個模型各一份。
你可以讀（`Get-PbiProtection`），不能替使用者決定。

| 等級 | 你可以做的 | 你不能做的 |
|---|---|---|
| **開放** | 讀內容。逐列明細每次最多 50 列、彙總結果 300 列 | —— |
| **換成代號** | 計數；當分組鍵（結果裡的值是 `ID_3FA2B81C07` 這種代號） | 看到真名；放進變數或量值；`MAX` / `CONCATENATEX` / `SELECTCOLUMNS` 取值 |
| **只能計數** | 包在計數類函式裡：`DISTINCTCOUNT` / `COUNTROWS` / `COUNT` / `COUNTA` / `COUNTBLANK` / `COUNTX`…，或回傳是／否的函式（`ISBLANK`、`ISEMPTY`、`HASONEVALUE`、`ISFILTERED`） | 列出、分組、`MAX` / `MIN`（對文字欄位那會吐出一個真實的值） |
| **只能彙總**（數字欄位） | `SUM` / `AVERAGE` / `MIN` / `MAX` / `SUMX`…；寫在篩選條件裡（`FILTER` 的條件、`CALCULATE` 的篩選引數）| 逐列取值。分組結果超過 100 列也會被擋 |

等級怎麼決定：**使用者逐欄設定的 ＞ 通用規則 ＞ 開放**。
通用規則是 `appsettings.json` 裡的欄名樣式（`*customer*`、`*amount*`…），比對前會正規化
（忽略大小寫、空白、底線、連字號），對每一個模型都生效；預設只認得英文欄名。

**寫查詢之前先看一眼**，不要靠被擋下來才知道哪些欄位受限：

```powershell
Get-PbiProtection                      # 受限的欄位（通常只有少數幾個）
Get-PbiProtection -Table Sales -All    # 某張表每一欄的等級
```

以上全部由**伺服器端強制執行**（`Program.cs` 的 `DataGuard`），違規查詢回 **403**。這代表：
**你不需要靠自律來避免外洩，但也不要把「沒被擋下」當成「這樣做很安全」** —— 擋不住的東西仍然會進 context。
「開放」的欄位每次 50 列、多問幾次就讀得完，所以請把開放的欄位當成「使用者同意你看」，而不是「反正看得到」。

### 寫法對照

```powershell
# ✅ 只能計數 —— 回傳的是個數
Invoke-Dax 'EVALUATE ROW("客戶數", DISTINCTCOUNT(Customers[account name]))'

# ⛔ 403 —— 名稱被當成分組鍵
Invoke-Dax 'EVALUATE SUMMARIZECOLUMNS(Customers[account name], "額", SUM(Sales[amount]))'
# ⛔ 403 —— MAX 對文字欄位會吐出一個真實的名稱，不算計數
Invoke-Dax 'EVALUATE ROW("x", MAX(Customers[account name]))'

# ✅ 只能彙總 —— 聚合，或寫在篩選條件裡
Invoke-Dax 'EVALUATE SUMMARIZECOLUMNS(Sales[product_line], "額", SUM(Sales[amount]))'
Invoke-Dax 'EVALUATE ROW("大單總額", CALCULATE(SUM(Sales[amount]), Sales[amount] > 100000))'
# ⛔ 403 —— 逐列取金額
Invoke-Dax 'EVALUATE SELECTCOLUMNS(Sales, "額", Sales[amount])'
```

**欄位被設成「換成代號」時**（要區分每一個客戶、但不需要知道是誰：前 20 大、集中度、流失名單）：

```powershell
(Invoke-Dax 'EVALUATE TOPN(20, SUMMARIZECOLUMNS(Customers[account name],
               "額", SUM(Sales[amount])), [額], DESC)').Rows
# → ID_3FA2B81C07 | 12,345,678
```

- 只有「直接當分組鍵／取相異值」會放行：`SUMMARIZECOLUMNS` / `SUMMARIZE` / `GROUPBY` / `VALUES` / `DISTINCT` / `ALL`，
  外面可以再包一層 `TOPN`。那些用法欄位會以原名出現在結果裡，伺服器才換得掉
- 代號**跨查詢穩定**（同一個值永遠同一個代號），可以先查前 20 大、再查這些代號的月趨勢
- 真名只印在**服務主控台** —— 使用者看得到，你看不到。要對照請他看那個視窗
- 回應的 `pseudonymized` 會列出哪些欄位是代號。**不要把 `ID_xxxxxxxxxx` 當成真實名稱解讀或寫進量值**
- **哪些欄位能用代號是使用者設的，不是你能打開的開關**（以前的 `-Pseudonymize` 已經拿掉）。
  欄位是「只能計數」而你需要按它分組時，向使用者說明用途，請他到儀表板把那一欄改成「換成代號」

### 量值參考不算欄位

量值的名稱常常長得像受限欄位 —— `[Total Revenue]`、`[V3_Amt]`、`[Customer Rank]`。
`[X]` 在模型裡**是量值、而且沒有任何資料行叫這個名字**時，這個參考本身不檢查；
被檢查的是**它的定義** —— 伺服器會把定義展開（量值、計算資料行、計算表、計算群組的項目，一路展開到底），接在查詢後面一起掃。

```powershell
# ✅ 放行 — [Total Revenue] 是量值，定義是 SUM(Sales[amount])，金額有聚合
Invoke-Dax 'EVALUATE ROW("值", [Total Revenue])'

# ⛔ 403 — 名稱放行不代表定義也放行：這個量值裡面是 MAX(Customers[account name])
Invoke-Dax 'EVALUATE ROW("值", [Top Customer])'

# ⛔ 403 — 結果欄位的「別名」照樣比對樣式。別名換成不像受限欄位的名稱就好
Invoke-Dax 'EVALUATE ROW("Total Revenue", [Total Revenue])'
```

- 量值和別張表的資料行**同名**時不適用 —— 那個名稱照舊當成欄位管制，寧可誤擋
- **仍然會被擋的合法量值**：定義裡先把受限欄位放進變數、再拿變數去 `COUNTROWS` 的那種
  （例如流失客戶數 `VAR 去年 = VALUES(表[客戶]) … COUNTROWS(EXCEPT(去年, 今年))`）。
  掃描看的是文字上「欄位有沒有被計數函式直接包住」，跟不進變數 —— 這是刻意保守，不是 bug。
  要驗算這類量值，把同一套公式換成不受限的欄位（產品線、型號）跑一次，確認邏輯

### 整個不能用的寫法

| 寫法 | 為什麼 |
|---|---|
| `TOCSV` / `TOJSON` / `COLUMNSTATISTICS` / `DETAILROWS` / `INFO.*` | 不必指名欄位就能把內容（或 M 腳本）帶出來，欄位層級的檢查看不到 |
| `ERROR()` 和受限欄位出現在同一句查詢（含它引用的量值） | 它會把任意文字放進錯誤訊息，內容可以從那裡出去 |
| `UNION` / `TREATAS` 的引數裡直接放含受限欄位的資料表或變數 | 輸出欄位會改用第一個引數的名稱，原本的欄名就看不見了。先用 `SELECTCOLUMNS` / `SUMMARIZE` 指定欄位 |

查詢碰到受限欄位而引擎報錯時，錯誤訊息裡引用的資料內容會換成 `‹已遮蔽›`（原文印在服務主控台）。
錯誤的種類、位置、函式名都還在，改公式夠用 —— 不要想辦法把被遮的部分弄出來。

### 被擋下時該怎麼辦

**先想「我真的需要這個值嗎」。** 多數情況答案是不需要 —— 換個彙總寫法就解決了。
驗算要確認的是**公式邏輯**，不是某一筆訂單是誰下的。

真的需要時，照這個順序：

1. **先在對話裡向使用者說明**：要看哪張表、哪些欄位、幾列、為什麼彙總不夠
2. 用 `-AskUser` 重送**同一句**查詢。使用者的螢幕會跳出 Windows 確認視窗（列出被擋的原因、
   受限欄位、列數與查詢內容），他按「是」才放行這一次

```powershell
Invoke-Dax 'EVALUATE ...' -AskUser -MaxRows 20     # 沒給 -MaxRows 時照逐列明細的上限（50）
```

- `-AskUser` **不是讓你自己放行的開關**，按鈕在使用者那邊。沒有先說明就跳視窗，等於要他對一件沒頭沒尾的事按「是」
- **先看 403 回應的 `canAskUser`**。`false` 就不要加 `-AskUser` 重送 —— 那種情況不會跳視窗，只會再被擋一次：
  - 寫法本身被禁（`TOCSV`、`INFO.*`、`ERROR()` 配受限欄位…）、模型定義讀不完整、保護設定檔讀不出來（`verdict` 是
    `BLOCKED-FUNC` / `BLOCKED-LOAD` / `BLOCKED-INCOMPLETE` / `BLOCKED-SETTINGS`）
  - 查詢連同它用到的量值碰到超過 12 個受限欄位，或去掉註解後超過 700 個字 —— 確認視窗列不完，
    使用者沒辦法對看不到全貌的東西做決定。把查詢拆小、寫短再問
- 使用者同意的是**他在視窗裡看到的那一句，連同當時的量值定義**。按「是」之後伺服器會重讀定義比對，
  中間有人改過量值就回 403（`BLOCKED-CHANGED`）—— 重新說明、重新問，不要想辦法避開比對
- 放行之後的結果沒有經過遮蔽（回應的 `approvedByUser` 是 `true`）—— 只用在說好的那件事上。
  上限是 1000 列、每格 200 個字（超過的部分會截斷並標示）
- 被拒絕（或 120 秒沒人回答）會回 403，`answer` 說明原因。**不要連續重送**：`retryAfterSeconds`
  之內伺服器不會再跳視窗。使用者按「否」就是答案，改用彙總寫法，或問他想怎麼做
- **不要請使用者「按是就好」**，也不要為了過關把查詢改寫成掃描看不出來的樣子
  （`CONCATENATEX`、變數、中介量值…）。想不出合規寫法就直接問使用者

### 改等級：收緊立刻生效，放寬要使用者本人按

```powershell
Set-PbiProtection -Table Customers -Column phone, address -Level countOnly    # 收緊：立刻生效
```

- **放寬不是你能決定的。** 服務會在使用者的螢幕跳出確認視窗、列出要放寬哪些欄位，他按「是」才生效
  （一次最多 12 欄）。要放寬之前先向使用者說明是哪些欄位、為什麼 —— 更好的做法是請他自己到儀表板改
- **不要為了讓查詢通過而去放寬保護**
- 設定檔存在 `%LOCALAPPDATA%\PBI_AI_Bridge\protection\`，刻意不放在專案資料夾。
  **不要讀寫那個資料夾** —— 繞過確認視窗去改檔，就是不經同意放寬。
  hook 會擋掉指令列上提到它的 Bash／PowerShell 指令，`.claude/settings.json` 的權限規則會擋掉檔案工具
  （整個 `%LOCALAPPDATA%\PBI_AI_Bridge\`，包含金鑰檔）；
  但這兩層看的是文字與路徑，**沒被擋到不代表可以做**
- 使用者沒有設定過的模型（`(Get-PbiProtection -Raw).configured` 是 `false`）只有通用規則在擋，
  中文欄名幾乎都是開放的。動手查資料之前**提醒使用者**到儀表板的「資料保護」分頁花一分鐘設定。
  `configured` 是「存過至少一筆逐欄設定」；使用者看過、決定都不用改的話它會一直是 `false`
  （那個「看過了」只記在他的瀏覽器裡，你看不到）—— 他說看過了就不要再提醒
- 設定跟著**模型**走，不是跟著路徑：搬家、改名、另存新檔、或從 Power BI 裡面開啟（服務拿不到路徑）時，
  服務靠模型裡資料行的固定代號（LineageTag）認回同一份設定，`Get-PbiProtection -Raw` 的 `inheritedFrom`
  會寫出當時的檔案。這種情況請使用者到儀表板掃一眼等級，不要自己判斷「應該沒問題」
- 用 `Rename-PbiObject` 改資料表／資料行的名稱時，受限的欄位會**維持原本的等級**（通用規則是看欄名的，
  不這樣做的話把欄名改掉就等於解除保護）。補了幾筆逐欄設定只印在服務主控台 ——
  改名之後跑一次 `Get-PbiProtection`，確認那幾欄的等級沒變

### 確認視窗的共通規矩

三種情況服務會在使用者的螢幕跳確認視窗：放行單一句查詢（`-AskUser`）、放寬欄位等級、
經由 API 改寫 Power Query（`Set-PbiMQuery`、`Set-PbiExpression`、`New-PbiTable -Kind m`、
`Add-PbiColumn -SourceColumn`、含 M 的還原與批次 —— 因為 M 可以把受限欄位的內容搬進開放的欄位）。

- 會跳視窗的動作，**送出之前先在對話裡講**：等一下會跳什麼視窗、裡面會列什麼
- 你的呼叫會等到使用者回答為止（最多 120 秒；`tools/PBI-Bridge.ps1` 裡會跳視窗的指令，用戶端都多等一段，
  所以你拿到的會是伺服器的回應而不是連線逾時）。回 403 時看 `answer`：
  `No`（拒絕）、`Timeout`（沒人回答）、`Busy`（已經有一個視窗開著）、`CoolingDown`（剛被拒絕過，
  `retryAfterSeconds` 秒內不會再問）、`Unavailable`（這台電腦跳不出視窗）
- 萬一真的遇到連線逾時（不是 403）：**不要直接重送**。使用者可能在最後一刻按了「是」——
  先確認有沒有寫進去（`Get-PbiProtection`、`(Get-PbiMQuery <表> -Show) -eq $m`），再決定下一步
- 被拒絕之後不要換個說法再送一次。如實回報，問使用者想怎麼做

### 這套機制做不到的事（不要對使用者說成做得到）

- 「只能計數」擋的是列出與分組。`COUNTROWS(FILTER(客戶, 客戶[名稱] = "某公司"))` 仍然會回一個數字 ——
  也就是可以問「有沒有這個值」。**不要用這種方式一步步把內容問出來**
- 「只能彙總」在條件篩到只剩一筆時，加總就等於那一筆
- 「換成代號」不是加密：同一個客戶永遠是同一個代號，客戶不多或特徵明顯時，從金額與地區仍然猜得出是誰

### DMV：只開放中繼資料與儲存統計

資料保護啟用時 `/api/dmv` 是白名單制（`TMSCHEMA_TABLES` / `COLUMNS` / `MEASURES` / `RELATIONSHIPS`…、
`DISCOVER_STORAGE_*`、`DISCOVER_OBJECT_MEMORY_USAGE`）。其他檢視一律 403：有些會回傳欄位的每一個相異值
（`MDSCHEMA_MEMBERS`）、M 腳本（`TMSCHEMA_PARTITIONS`）或別的連線跑過的查詢文字。被擋時回應的 `allowed` 會列出可以查的。

- 寫法只接受**單一來源**：`SELECT <欄位或 *> FROM $SYSTEM.<檢視> [WHERE … | ORDER BY …]`。
  前面有註解、出現兩個 `$SYSTEM`、兩個 `FROM`（子查詢、JOIN）都會被擋 —— 那些寫法可以把別的檢視夾帶進來
- **相依性檢視（`DISCOVER_CALC_DEPENDENCY`）不在名單內。** 要知道量值之間誰引用誰，讀 `Get-PbiMeasures` 的定義

### M 腳本：預設不輸出內容

M 腳本是 schema 中最敏感的部分 —— 連線字串、伺服器位址、資料庫名稱、檔案路徑都在裡面，
而且**篩選步驟裡很可能有寫死的客戶名稱**（`Table.SelectRows(…, each [customer] = "…")`）。
伺服器端的欄位管制**管不到這裡**（它管的是查詢回傳值，不是 schema），所以這條要靠自律。

`Get-PbiMQuery` 已改為**預設只存檔、回傳路徑，不回傳內容** —— 存到本機磁碟不等於送進雲端。
要讀內容必須明確加 `-Show`，加之前先想清楚為什麼需要整段 M 進入 context。

**不要直接 `Get-PbiSchema | ConvertTo-Json`**，那會把全部 M 腳本拉進 context。

```powershell
# ✅ 只看結構
Get-PbiSchema | ForEach-Object { $_.Tables } |
  Select-Object Name, @{n='欄位數';e={$_.Columns.Count}}, @{n='量值數';e={$_.Measures.Count}}

# ✅ 只看量值（不含 M）
Get-PbiMeasures | Select-Object Table, Measure, Expression
```

確實需要讀某張表的 M 腳本時，**指名單一表格**並先說明用途，不要整包拉。

### 機密值不輸出到主控台

- **API Key**：`Get-PbiApiKey` 內部使用即可，不要 `Write-Host` 或 echo 出來，也不要去讀金鑰檔（位置見「🔑 API Key」）
- `appsettings.json` 現在沒有機密（金鑰已經搬走），但從舊版升級上來的那一份可能還留著一行失效的 `Security.ApiKey` ——
  需要哪個欄位就只取那個欄位，不要整份印出來

### 在 PowerShell 端過濾，不要在 context 裡過濾

PowerShell 管線中的資料**不會**進入 Claude 的 context，只有最終輸出會。
善用這點：`Select-Object`、`Where-Object`、`Measure-Object` 都在本機執行。

```powershell
# ✅ 完整 schema 留在本機，只有計數進入 context
(Get-PbiSchema).Tables | Where-Object { $_.Measures.Count -gt 0 } | Measure-Object

# ❌ 全部倒進 context 之後才用眼睛找
Get-PbiSchema | ConvertTo-Json -Depth 6
```

---

## 🛡️ 防毒軟體相容規範（公司環境，優先於「把事情做完」）

**這台機器裝有公司控管的趨勢科技防毒，採行為偵測。**
2026-08-06 已實際觸發過警報，以下規則是那次事件的產物，不是預防性的保守。

**核心觀念：防毒看的是行為特徵，不是意圖。** 動機正當不會讓警報消失。
一連串「列舉 → 定位 → 終止 → 執行新產物」的動作，客觀上就是它被訓練來抓的樣態。

**觸發警報的代價很高**：可能被列入資安事件、可能連累使用者被 IT 約談。
所以當「最有效率的做法」與「不觸發防毒」衝突時，**選後者，並請使用者代勞**。

### ⛔ 絕對不要自己做（一律改成請使用者動手）

| 不要做 | 改成 |
|---|---|
| `Stop-Process` / `taskkill` 終止任何行程 | 請使用者到該視窗按 Ctrl+C 或自己關掉，然後回報 |
| `Start-Process` 啟動剛編譯出來的 `.exe` | 請使用者雙擊 `🚀啟動PBI終極儀表板.bat` |
| 對使用者設定檔做遞迴掃描（`-Recurse` 掃 `AppData`、`Packages`、使用者家目錄） | **指定到最末層資料夾**，不加 `-Recurse`；需要往下找就一層一層來 |
| 讀取或列舉瀏覽器設定檔目錄（`WebView2`、`EBWebView`、`Cookies`、`Local State`、`Login Data`） | 完全不要碰。這是竊密木馬的取材位置，**列舉本身**就會觸發 |
| 存取防毒軟體自己的目錄或紀錄檔 | 請使用者從趨勢科技的介面看，把訊息貼給你 |
| `Invoke-WebRequest` / `curl` 下載任何檔案 | 請使用者自己下載 |
| `-EncodedCommand`、Base64、字串拼接組出指令 | 指令一律寫成明文，讓 AV 看得懂在做什麼 |
| 短時間重複送 SendKeys（`Save-PbiModel` 連按） | 沒存到時**最多再試一次**，還是不行就如實回報，請使用者手動 Ctrl+S。只是想確認「到底存了沒」→ `Save-PbiModel -VerifyOnly`（只看磁碟，不送按鍵） |

### ✅ 這些是安全的，照常用

- `http://localhost:5500` 的所有 API 呼叫（不對外連線）
- `Get-PbiInstances` —— 行程反查是在**服務內部**做的，比在 PowerShell 端跑 `Get-CimInstance Win32_Process` 安靜
- 讀寫專案資料夾與 scratchpad 內的檔案
- `dotnet build`（**建置本身沒問題，問題在建置完自己去執行它**）
- 明確指定單一路徑的 `Get-Content` / `Get-ChildItem`

### 需要偵察資訊時的順序

1. 先問**服務**：`Get-PbiInstances`、`Get-PbiInfo` 已經涵蓋 Port、行程、檔案路徑
2. 服務問不到 → 明確指定路徑去讀，不遞迴
3. 還是不夠 → **停下來告訴使用者你需要什麼、為什麼**，請他貼給你

不要為了省一次來回而自己去掃。**多問一句永遠比觸發一次警報便宜。**

---

## 🚦 開工前的檢查（每個 session 第一件事）

### 1. 一次看完現況

```powershell
. ".\tools\PBI-Bridge.ps1"
Get-PbiOverview                 # 只開一個 Power BI 時
Get-PbiOverview 銷售報表          # 開了多個：用檔名片段或 Port 選定
```

`Get-PbiOverview` 一次回報：服務有沒有跑、選定的是哪一份檔案（含完整路徑）、模型大小、有沒有壞掉的公式、
各保護等級的欄數、每張表的資料行／量值／受限欄位數。只有結構與計數，沒有 M 腳本，也沒有資料內容。
**沒讀過現況就寫入 = 盲改，一律先跑這一步。**

照它回報的狀況處理：

| 回報 | 你要做的 |
|---|---|
| 連不上橋接服務 | 請使用者雙擊 `🚀啟動PBI終極儀表板.bat` 並**保持視窗開啟**。不要自己在背景啟動服務 |
| 沒有偵測到 Power BI | **停下來**請使用者先開啟檔案 —— 服務靠執行中的引擎找模型，沒開檔完全無法運作 |
| 有多個 Power BI、還沒選定 | 用檔名片段或 Port 選。沒選之前伺服器一律拒絕（連讀取也拒絕）—— **不要為了省事隨便挑一個**，猜錯的代價是改到別的模型 |
| 印出了完整路徑 | **確認那是你要改的那一份**。同一台機器常有多份同名的檔案（桌面一份、專案資料夾一份），改錯檔案比改錯公式難救 |
| 路徑不明 | 使用者先開 Power BI、再從裡面選檔案，服務拿不到路徑（`kind` 是 `Unknown`，名稱來自視窗標題）。讀寫模型、查詢、資料保護照常；`Save-PbiModel` 與 `Test-PbiReport` 不能用 —— 存檔請使用者自己按 Ctrl+S，回報時講明「沒有驗證過存檔內容」；快照與 M 備份重開後找不回來。需要這些功能時，先說明原因，請他關掉 Power BI、改用**雙擊檔案**的方式開啟。認不出是哪一份檔案就問使用者，不要去掃磁碟找 |
| 這份模型還沒設定過保護 | 查資料之前**提醒使用者**到儀表板的「資料保護」分頁花一分鐘設定（他說看過了就不用再提） |
| 有壞掉的公式 | 動手之前先告訴使用者：那是既有的問題，不是你造成的 |

### 2. 需要更細的資訊時

```powershell
Get-PbiProtection -Table Sales -All                          # 某張表每一欄的保護等級（寫查詢之前先看）
Get-PbiMeasures | Select-Object Table, Measure, Expression   # 量值與公式（不含 M）
Get-PbiRelationships                                         # 關聯線
Test-PbiModel                                                # 健檢的完整清單
Get-PbiHelp                                                  # 所有指令
```

- 每次工具呼叫都是全新的 PowerShell：**每次都要重新載入 `PBI-Bridge.ps1`**；開了多個 Power BI 時，每次呼叫的開頭也要重新
  `Use-PbiInstance`（只開一個時不用）。載入時只印一行，連那一行都不要就先設 `$env:PBI_BRIDGE_QUIET = '1'`
- **不要自己手刻 `Invoke-RestMethod`**（原因見「編碼陷阱」），一律用 `tools/PBI-Bridge.ps1` 的指令。
  指令的用途與參數寫在那個檔案的註解裡；端點層級的說明在 `API_Documentation.html`，平常用不到

---

## ⚠️ 編碼陷阱（最容易踩的坑）

本機是 **Windows PowerShell 5.1**。`Invoke-RestMethod -Body "<字串>"` 預設不是 UTF-8，
中文欄位名、中文 DAX 字串會在傳輸中損毀，伺服器收到破碎 JSON 後回 **400 Bad Request**。

**正確做法 —— 一律把 body 轉成 UTF-8 位元組再送：**

```powershell
$bytes = [System.Text.Encoding]::UTF8.GetBytes(($obj | ConvertTo-Json -Compress))
Invoke-RestMethod -Uri $url -Method Post -Headers $h `
  -ContentType "application/json; charset=utf-8" -Body $bytes
```

`tools/PBI-Bridge.ps1` 已經封裝好這件事，用它就不會出錯。

**看到 400 時**：先讀回應 body 再判斷，不要只看狀態碼。
```powershell
$sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream(), [System.Text.Encoding]::UTF8)
$sr.ReadToEnd()
```

### 建立含中文的 .ps1 檔必須加 UTF-8 BOM

PowerShell 5.1 讀取 **沒有 BOM** 的 `.ps1` 時，會以系統 ANSI 碼頁解析，中文全部變亂碼，
接著就是 `Unexpected token` 之類的語法錯誤 —— 而錯誤訊息本身也是亂碼，很難判讀。

寫完含中文的 `.ps1` 後，一律補上 BOM：

```powershell
$f = ".\tools\Something.ps1"
$c = Get-Content $f -Raw -Encoding UTF8
[System.IO.File]::WriteAllText($f, $c, (New-Object System.Text.UTF8Encoding $true))
```

（`.md`、`.json`、`.cs` 不受影響，只有要被 PowerShell 5.1 *解析* 的 `.ps1` 需要。）

---

## 🔑 API Key

由**服務在第一次啟動時自己產生**，存在 `%LOCALAPPDATA%\PBI_AI_Bridge\api-key.txt` —— **不在專案資料夾裡**
（2026-10-08 從 `appsettings.json` 搬出來：專案資料夾會被複製、壓縮、分享，金鑰不該跟著走）。

- `tools/PBI-Bridge.ps1` 的 `Get-PbiApiKey` 會自動讀，**你不需要、也不應該自己去讀那個檔案**。
  `.claude/settings.json` 的權限規則禁止檔案工具讀寫 `%LOCALAPPDATA%\PBI_AI_Bridge\` 整個資料夾
- **不要 echo 到主控台**，不要寫死在任何腳本或文件裡。服務的黑窗也不再印出金鑰（連開頭幾碼都不印）
- 「找不到 API 金鑰檔」= 服務還沒用新版啟動過。請使用者關掉黑窗、重新雙擊 🚀（不要自己去建那個檔案）
- 金鑰擋的是別的 Windows 帳號與別的網站，**不是用來防你的** —— 你本來就拿得到。防你的是伺服器端的檢查與確認視窗
- 「換成代號」的代號是從金鑰衍生的：金鑰重新產生（升級、或使用者刪掉金鑰檔）之後，代號會整批換新，
  先前對話裡記下的 `ID_xxxxxxxxxx` 就對不上了

---

## 🔄 標準開發循環

**核心原則：寫入後一定要驗算。不驗算就不算完成。**

```
0. New-PbiSnapshot        破壞性操作前先留退路（刪除、改名、覆寫 M）
1. Get-PbiOverview        讀現況；再用 Get-PbiMeasures 理解既有命名與邏輯
2. Set-PbiMeasure         寫入（多筆請用 Invoke-PbiBatch）
3. Invoke-PbiRefresh      讓變更生效 ← 結構性變更必做，見下方
4. Invoke-Dax             查詢結果，比對預期值
5. 不符預期 → 回到 2 修正公式
6. 符合預期 → Save-PbiModel -Expect <剛寫的量值名稱>，確認 expectFound = true
```

### ⚠️ 什麼時候必須 refresh

**只寫量值不用 refresh**，但以下情況不 refresh 就查不到東西（會出現「需要重新計算」的錯誤）：

| 做了什麼 | 要跑什麼 | 大約耗時 |
|---|---|---|
| 建立計算表 | `Invoke-PbiRefresh -Table <表>` | 秒級 |
| 建立／修改關聯線 | `Invoke-PbiRefresh -RefreshType calculate` | 秒級 |
| 新增計算項目到計算群組 | `Invoke-PbiRefresh -Table <計算群組>` | 秒級 |
| 新增計算資料行 | `Invoke-PbiRefresh -RefreshType calculate` | 秒級 |
| 用 API 寫了 M，Desktop **沒有**出現套用提示 | `Invoke-PbiRefresh -Table <表>`（full） | 看資料源，可能數分鐘 |
| M 多輸出了新欄位 | 先 `Add-PbiColumn -SourceColumn`，再 `Invoke-PbiRefresh -Table <表>` | 同上 |

使用者自己在 Power Query 編輯器「關閉並套用」、或按了套用提示時，Power BI 會自己重整，不必再跑
（見 `docs\power-query.md`）。

`calculate` 只重算 DAX、不重抓資料源，很便宜 —— 結構性變更後直接跑一次就好。
`full` 會重抓資料源：資料源憑證過期時沒有 UI 可以輸入密碼，會直接失敗。

**回應是「跑完之後」才回來的**（`message` 是「已完成」，`elapsedMs` 是實際耗時）——
收到就可以直接驗算，不用等、也不用再跑一次確認。失敗會是錯誤回應，不會是「已完成」。

### 存檔

`Save-PbiModel` 會把 PBI Desktop 帶到前景送出 Ctrl+S，然後輪詢到檔案寫完為止（預設最多 30 秒，
寫完就提早回來）。它回報三件事，**三件事回答的問題不一樣**：

| 欄位 | 回答什麼 | 不能回答什麼 |
|---|---|---|
| `fileChanged` | 有沒有檔案被寫了 | 寫進去的是不是你剛改的東西 |
| `changedFiles` | 被寫的是哪些檔（PBIP 看得到是哪張表的 `.tmdl`、哪個視覺） | 同上 |
| `expectFound` | 你用 `-Expect` 指定的文字在不在磁碟上 | —— 這才是「存到了」的證據 |

**為什麼需要 `-Expect`**：實際遇過好幾次 —— 回報存檔成功、檔案時間也變了，去看磁碟上的 TMDL
卻沒有剛寫的量值，要再存一次才進去。推測是用 API 寫進模型之後，Power BI Desktop 要一點時間
才同步到自己那邊，Ctrl+S 太早到就存了**同步前的舊狀態**（原因是推測，現象是實測）。
不管原因是什麼，結論一樣：`fileChanged = true` 不是「存到了」的證據。

```powershell
Save-PbiModel -Expect 'V3_Amt'                 # 量值名稱，或公式裡一小段有辨識度的文字
Save-PbiModel -VerifyOnly -Expect 'V3_Amt'     # 不送按鍵，只看磁碟 —— 回頭確認用
```

- **PBIP 一律帶 `-Expect`**，看 `expectFound`。只回傳找到與否，檔案內容不進 context。
  更新既有量值時名稱本來就在磁碟上，要給「新公式裡才有的那一段」才驗得到
- **PBIX 是二進位檔，沒辦法檢查內容**（`expectFound` 會是空的）—— 只能看 `fileChanged`，
  回報時要講明「檔案有寫入，但無法確認內容」
- `expectFound = false`：等幾秒**最多再存一次**。還是找不到就如實回報，請使用者手動 Ctrl+S，
  之後用 `-VerifyOnly` 確認
- `fileChanged = false`：30 秒內沒有任何檔案變動 —— 沒有待存的變更、按鍵沒送達，或大檔還在寫。
  **不要重送按鍵來「確認」**，過一會兒用 `-VerifyOnly` 看
- 跨行程送合成按鍵會被防毒視為鍵盤側錄／UI 劫持，短時間重複會放大可疑度 —— 所以重試上限是一次
- 會短暫搶走鍵盤焦點，這是 SendKeys 的固有限制
- 驗證範圍只有**選定實例自己的檔案**：PBIX 看那一個檔；PBIP 看它的 `.Report` 與 `.SemanticModel`
  兩個資料夾（照 `.pbip` 與 `definition.pbir` 裡的指標找，不是看檔名猜）。
  Ctrl+S 也只送給那一個行程 —— 不會誤存到別的 PBI，旁邊不相干的檔案被寫入也不會被當成存檔成功
- ⚠️ **Ctrl+S 存的是整份檔案，包含報表版面**。你剛直接改過報表檔案、使用者還沒按「接受變更」時，
  這一存會把 Power BI 記憶體裡的舊版面寫回去，蓋掉你的修改（見 `docs\pbip-report-editing.md`）

### 驗算範例

```powershell
# 寫入
Set-PbiMeasure -Table "量值" -Name "銷售總額" -Expression "SUM(FactSales[Amount])" -Format "#,0"

# 立刻驗算
(Invoke-Dax 'EVALUATE ROW("結果", [銷售總額])').Rows

# 交叉比對：換個寫法看數字是否一致
(Invoke-Dax 'EVALUATE ROW("對照", SUMX(FactSales, [Amount]))').Rows
```

驗算時的實用招式（皆為彙總，不外洩明細）：
- 用 `ROW()` 取單一純量值
- 用 `COUNTROWS` 確認篩選條件影響的列數是否合理
- 用 `SUMMARIZECOLUMNS` 檢查分組後的小計
- 用 `IF([舊] = [新], "YES", "NO")` 比對兩個公式是否等價
- 改寫公式後比較 `ElapsedMs`，確認沒有寫出效能地雷

> 驗算查詢一樣受資料保護管制（見本文件開頭）。受限的欄位只能計數或彙總 ——
> 這對驗算幾乎沒有影響，因為驗算要確認的是**公式邏輯**，不是某一筆訂單是誰下的。
> 被 403 擋下時**不要改寫查詢去繞**，先問自己需不需要那個值；
> 真的需要就向使用者說明，再用 `-AskUser` 請他在確認視窗決定。

### 查詢回應怎麼讀

- `Invoke-Dax` 的查詢必須以 `EVALUATE` 或 `DEFINE` 開頭（唯讀）
- `-MaxRows` **只能把上限調低，不能調高**。伺服器封頂：逐列明細 50 列、彙總結果 300 列
- 回應的 `truncated` 是 `true` 代表**還有資料沒帶回來** —— 不要當成「這就是全部」，
  尤其不要據此下「總共只有 N 筆」的結論。要總數請改用 `COUNTROWS`
- `elapsedMs` 可以用來比較不同 DAX 寫法的效能
- DAX 語法錯誤回 **400**，內含引擎的訊息與錯誤位置，據此直接修正。被資料保護擋下回 **403**，
  內含 `reason`（為什麼）、`hint`（怎麼改寫）、`verdict`（判定代碼）、`canAskUser`
- 查詢與寫入不會交錯執行。另一邊跑超過 30 秒時會回 400「…等了 30 秒還沒結束」—— 稍後再試即可，不是壞掉

### 多筆寫入用批次

寫 20 個量值不要呼叫 20 次：`Invoke-PbiBatch` 一次連線、一次存檔，模型只重算一次。
預設任一步失敗就**整批不套用**，模型維持原狀；只有後面的步驟依賴前面「已經生效」時才加 `-SavePerOp`（那時失敗不會回滾）。

---

## 🛑 安全守則

### 修改是寫進記憶體，不是檔案

TOM 的 `SaveChanges()` 只改 Power BI Desktop **記憶體中**的模型。畫面會立刻更新，
但要進到 `.pbip`/`.pbix` 必須存檔。

→ 完成一段工作後跑 `Save-PbiModel -Expect <剛寫的東西>` 並**確認 `expectFound = true`**
   （PBIX 沒辦法檢查內容，只能看 `fileChanged`）；沒存到就如實告訴使用者，並請他手動按 Ctrl+S。

### 寫入前先確認

- **破壞性操作前先 `New-PbiSnapshot`** —— 刪除、改名都算
- 覆寫既有量值前，先用 `Get-PbiMeasures` 把**原本的公式記下來**，並在回報中附上
- 請使用者動 M 之前，先 `Get-PbiMQuery <表名> -Label <標籤>` 留一份 `.pq`
- 刪除量值、資料表、資料行（`Remove-PbiMeasure` / `Remove-PbiTable` / `Remove-PbiColumn`）無法復原，執行前先向使用者確認
- **改名一律先 `-DryRun`**：`Rename-PbiObject` 會改寫全模型的 DAX 引用，先看過 `rewrites` 清單確認沒誤改

### 快照是每個模型各自獨立的

`snapshots/<檔名>_<路徑雜湊>/<時間戳>__標籤.json`

路徑雜湊是必要的：桌面的 `X.pbix` 和專案裡的 `X.pbip` 檔名相同但是**完全不同的模型**，
混在一起會互相污染。`Get-PbiSnapshots` 只列出目前選定實例的快照，
`Restore-PbiSnapshot` 也只在該資料夾裡找 —— **結構上就不可能跨模型還原**。

路徑不明的實例（從 Power BI 裡面開啟的檔案）沒有路徑可以算雜湊，資料夾是 `unsaved_port<Port>`。
Port 每次開啟都不同，所以那些快照**只在這次開啟期間找得到** —— 存之前先告訴使用者這一點。

### `Restore-PbiSnapshot` 的邊界

還原**不是**把模型完整倒回快照當時的狀態，它只處理量值、計算資料行、關聯、M 腳本、共用運算式。

- 快照之後**新建的表格與角色不會被刪除**，連帶它們底下的量值也不會被還原
- 不會改模型層級屬性（例如 DiscourageImplicitMeasures）
- 回應的 `uncovered` 會列出所有沒被涵蓋的項目 —— **回報時要一併轉述**，不要讓使用者以為已完全還原
- 反過來，在還原範圍內的東西是「回到當時狀態」：快照中不存在的量值與關聯**會被刪掉**，不是合併

### 計算群組的前提

模型的 `DiscourageImplicitMeasures` 必須是 `true` 才能建立計算群組（引擎強制）。
開啟後**使用者無法再把數值欄位直接拖進視覺自動彙總**，一律得改用量值。

→ 這是模型層級的行為改變，**務必先向使用者說明並取得同意**，不要自作主張開啟。

---

## 📚 用到才讀的文件

下面這些工作各有一份專門的規範，放在 `docs\`。**開始做之前先把對應的那一份讀完** ——
裡面是實際踩過的坑，不是可有可無的補充。資料保護、防毒與確認視窗的規矩一律以本文件為準。

| 要做的事 | 先讀 | 就算還沒讀也要守的硬規則 |
|---|---|---|
| 讀或改 Power Query（M） | `docs\power-query.md` | 寫入之前先向使用者說明要改什麼（會跳確認視窗）；寫完一定要問他 Power BI 有沒有出現套用提示；M 一次只讀一張表，預設不把內容讀進 context |
| 直接改 PBIP 的報表檔（版面、視覺） | `docs\pbip-report-editing.md` | 模型的變更先存檔；改完報表檔之後、使用者按「接受變更」之前**不准再存檔**；動手前先備份整個 `.Report` 資料夾 |
| 遇到看不懂的錯誤、400、403、409 | `docs\troubleshooting.md` | 先讀回應內容再判斷，不要只看狀態碼；被確認視窗拒絕之後不要重送 |
| 改服務本身（`Program.cs`、設定、啟動檔、hook、`tools\*.ps1`） | `docs\service-development.md` | 重啟服務由使用者動手（關黑窗、雙擊 🚀）；改完 `tools\*.ps1`、hook 或啟動檔一定要跑 `tools\Test-BridgeClient.ps1` |

---

## 🧱 這個資料夾裡要記得的事

完整的專案結構在 README 第 12 節與 `docs\service-development.md`。平常要記得的只有這幾件：

- `tools\*.ps1` 與 `.claude\hooks\*.ps1` 必須保留 **UTF-8 BOM**；`🚀啟動PBI終極儀表板.bat` 只能有 ASCII 字元、行尾必須是 CRLF（不要用 `sed` 改它）
- `%LOCALAPPDATA%\PBI_AI_Bridge\`（API 金鑰與每個模型的保護設定）**整個資料夾都不要讀寫** —— 指令列與檔案工具都有規則在擋，但沒被擋到也不要碰
- `snapshots\`、`PowerQuery_Scripts\`、`audit\` 是服務自動產生的模型產物，裡面有模型結構、M 腳本與查詢紀錄：不要整份讀進 context，也不要提交到版控
- **不要在這裡放特定模型的產物**（資料字典、schema 匯出、單一模型的備份）。那種東西會過期，而過期的參考資料比沒有更糟 ——
  需要時用 `Get-PbiOverview`、`Get-PbiMeasures`、`Get-PbiMQuery` 取得當下的真實狀態
- 這個資料夾只放通用的 Power BI 開發工具，不相干的專案請放到外面
