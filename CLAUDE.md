# PBI AI Bridge — Claude Code 工作規範

> 這份文件會在每次新 session 自動載入。開始任何 Power BI 工作前先讀完本節。

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

## 🚦 開工前的檢查順序（每個 session 第一件事）

**務必按順序做完，不要跳過。**

### 1. 載入輔助函式並檢查環境

```powershell
. ".\tools\PBI-Bridge.ps1"
Test-PbiBridge
```

`Test-PbiBridge` 一次告訴你三件事：服務有沒有跑、有哪些 PBI 可以操作、目前選了哪一個。

- 服務沒起來 → 請使用者雙擊 `🚀啟動PBI終極儀表板.bat`，**並保持該視窗開啟**。
  不要自己在背景偷偷啟動服務而不告知使用者。
- 沒有任何實例 → **停下來請使用者先開啟 PBI 檔案**。橋接服務靠 `msmdsrv` 行程反查 Port，沒開檔完全無法運作。

**不要自己手刻 Invoke-RestMethod**，原因見下方「編碼陷阱」。

載入時只會印一行（非互動環境不重印指令清單 —— 每次工具呼叫都要重新載入，十幾行清單白佔 context）。
要看有哪些指令：`Get-PbiHelp`。連那一行都不要：載入前設 `$env:PBI_BRIDGE_QUIET = '1'`。

### 2. 選定要操作哪一個 Power BI ← 這是通用工具，這步不能跳

這台機器上可以同時開好幾個 PBI，每一個都有自己的 msmdsrv 與連接埠。

```powershell
Get-PbiInstances                # 列出全部
Use-PbiInstance 銷售報表             # 用檔名片段選
Use-PbiInstance 55820           # 或用 Port 精確指定
```

規則：

- **只有一個實例** → 不用選，自動採用
- **有多個實例而沒選** → 伺服器一律**拒絕執行**（連讀取也拒絕），並列出清單。
  這是刻意的：猜錯的代價是改到別的模型，寧可回錯誤也不要預設挑一個
- 選定後會用 Port 鎖定；PBI 關掉重開後 Port 會變，`Invoke-PbiApi` 會用記住的完整路徑**自動重新解析**

**同一台機器常有多份同名的 `.pbip` / `.pbix`**（例如桌面一份、專案資料夾一份）。
`Use-PbiInstance` 的回應會印出完整路徑 —— **確認那是你要改的那一份**，改錯檔案比改錯公式難救。

**`kind` 是 `Unknown`、路徑是空的** = 使用者先開 Power BI、再從裡面選檔案。服務只能從 Power BI 的啟動參數
得知路徑，這種開法拿不到，`fileName` 是從視窗標題來的（不含副檔名）。這時：

- 讀寫模型、查詢、資料保護（設定照樣存得起來、認得回來）都照常
- **`Save-PbiModel` 與 `Test-PbiReport` 不能用**（沒有檔案可以看；`Save-PbiModel` 會直接回 400、不送 Ctrl+S）——
  存檔請使用者自己按 Ctrl+S，並在回報裡講明「沒有驗證過存檔內容」。快照與 M 備份的資料夾是用 Port 命名的，
  Power BI 重開就找不回來
- 用視窗標題認不出是哪一份檔案時，**問使用者**，不要去掃磁碟找同名檔案
- 要做存檔驗證、改報表檔、或需要留得住的快照 → 先向使用者說明，請他關掉 Power BI、改用**雙擊檔案**的方式開啟

### 3. 先讀模型再動手

```powershell
(Get-PbiSchema).Tables | Select-Object Name, @{n='欄位數';e={$_.Columns.Count}}, @{n='量值數';e={$_.Measures.Count}}
Get-PbiMeasures | Select-Object Table, Measure, Expression   # 量值（不含 M，M 見「M 腳本：預設不輸出內容」）
Get-PbiRelationships                       # 關聯線
Test-PbiModel                              # 健檢：有沒有既存的壞公式
Get-PbiProtection                          # 哪些欄位的內容受限（寫查詢之前先知道）
```

沒讀過現況就寫入 = 盲改。**一律先掃描。**

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

## 📡 API 端點總表

Base: `http://localhost:5500`
Headers: `X-API-Key: <key>`　＋　`X-PBI-Target: <Port 或檔名片段>`（多實例時必帶）

> `X-PBI-Target` 只能放 ASCII（HTTP 標頭限制），所以 `Use-PbiInstance` 內部一律換算成 **Port** 再送出。
> `tools/PBI-Bridge.ps1` 會自動帶上，不需要手動處理。

### 讀取（安全，隨時可用）

| 端點 | 方法 | 用途 |
|---|---|---|
| `/ping` | GET | 健康檢查（免 Key） |
| `/` | GET | 網頁儀表板（給人用，免 Key —— 伺服器把金鑰帶進頁面）。所有請求只接受 Host 為 localhost / 127.0.0.1 |
| `/api/schema` | GET | 全部資料表、欄位、DAX 量值原始碼、Power Query M 腳本 |
| `/api/relationships` | GET | 關聯線：基數、雙向篩選、是否啟用 |
| `/api/roles` | GET | RLS 角色與資料表篩選規則 |
| `/api/expressions` | GET | Power Query 共用運算式（參數、函式） |
| `/api/model-props` | GET | 模型層級屬性、相容性層級、計算群組數 |
| `/api/instances` | GET | **列出所有執行中的 PBI**（Port、檔案、PBIX/PBIP）— 切換檔案的第一站 |
| `/api/pbi-info` | GET | 目前解析到的目標實例是哪一個 |
| `/api/validate` | GET | 模型健檢：壞掉的公式、雙向關聯、孤島表、疑似沒人用的欄位 |
| `/api/validate-report` | GET | **報表健檢（僅 PBIP）**：報表檔案的 JSON 能否解析、引用的欄位在模型裡存不存在、視覺重疊／超出頁面。只回結構，不回篩選條件裡的值 |
| `/api/snapshots` | GET | 列出所有快照 |
| `/api/protection` | GET | 每個欄位的保護等級、是誰決定的（逐欄設定／通用規則／預設）。等級是結構資訊，可以放心讀 |
| `/api/query` | POST | **執行唯讀 DAX 查詢** — 驗算用。`AskUser` 見下方 |
| `/api/dmv` | POST | `$SYSTEM` DMV 查詢（中繼資料、儲存統計、記憶體佔用）。只接受單一來源的 `SELECT ... FROM $SYSTEM.xxx`；資料保護啟用時是白名單制，見「DMV：只開放中繼資料與儲存統計」 |

### 安全網（動手前先做）

| 端點 | 方法 | 用途 |
|---|---|---|
| `/api/snapshot` | POST | 把整個模型序列化成 TMSL 存檔 |
| `/api/restore` | POST | 從快照還原量值／計算資料行／關聯／M／共用運算式（支援 `DryRun`） |

### 讓變更生效

| 端點 | 方法 | 用途 |
|---|---|---|
| `/api/refresh` | POST | **重新整理資料**。建完計算表／計算項目指定該表；建／改關聯、計算資料行用 `calculate`；用 API 寫了 M 而 Desktop 沒出現套用提示時，對該表跑 `full`。**跑完才回應**（`message` 是「已完成」，`elapsedMs` 是實際耗時） |
| `/api/save` | POST | **模擬 Ctrl+S 存檔並驗證**：回報哪些檔案變了（`changedFiles`），`Expect` 可指定「必須出現在磁碟上的文字」（`expectFound`）；`VerifyOnly` 不送按鍵、只看磁碟 |

### 寫入（會改動模型，動手前先確認）

| 端點 | 方法 | 用途 |
|---|---|---|
| `/api/upsert-measure` | POST | 新增／覆寫量值（公式、格式、說明、DisplayFolder、IsHidden） |
| `/api/delete-measure` | POST | 刪除量值 |
| `/api/move-measure` | POST | 量值搬移到其他表（目標表不存在會自動建立） |
| `/api/add-column` | POST | 新增資料行：給 `Expression` 是 DAX 計算資料行；給 `SourceColumn` 是對應 M 輸出欄位的來源資料行（M 多輸出一欄時用） |
| `/api/delete-column` | POST | 刪除資料行（仍被關聯使用時會擋下） |
| `/api/set-column-props` | POST | 格式、DisplayFolder、隱藏、SortByColumn、SummarizeBy、資料類別 |
| `/api/create-table` | POST | 建表：`calculated`（DAX）／`m`（M，需給 Columns）／`measureHolder` |
| `/api/update-m` | POST | **覆寫資料表的 M 腳本**。只改模型那一份、不會重抓資料。寫完要讓它生效（Desktop 有提示就請使用者按「套用」，沒有就 refresh 該表），再確認留下的版本 |
| `/api/delete-table` | POST | 刪表（仍有關聯線時會擋下） |
| `/api/rename` | POST | **改名並同步改寫所有 DAX 引用**，支援 `DryRun` |
| `/api/upsert-relationship` | POST | 建立／更新關聯（基數、篩選方向、是否啟用） |
| `/api/delete-relationship` | POST | 刪除關聯 |
| `/api/upsert-expression` · `/api/delete-expression` | POST | Power Query 參數與共用函式 |
| `/api/upsert-calc-group` | POST | 建立計算群組（需先開 DiscourageImplicitMeasures） |
| `/api/upsert-calc-item` · `/api/delete-calc-item` | POST | 計算項目（用 `SELECTEDMEASURE()`） |
| `/api/upsert-role` · `/api/delete-role` | POST | RLS 角色（TablePermissions 為**整組取代**） |
| `/api/set-model-props` | POST | 模型層級屬性（DiscourageImplicitMeasures 等） |
| `/api/batch` | POST | **一次連線、一次 SaveChanges 套用多個操作** |
| `/api/protection` | POST | 變更欄位的保護等級。收緊立刻生效；**放寬會在使用者的螢幕跳確認視窗**，他按「是」才套用（一次最多 12 欄） |

資料保護啟用時，會動到 Power Query 的寫入（`update-m`、`upsert-expression`、`create-table` 的 `Kind=m`、
`add-column` 的 `SourceColumn`、含 `mquery`／`expressions` 範圍的 `restore`、含上述操作的 `batch`）
同樣會先跳確認視窗 —— 見「資料保護規範」的「確認視窗的共通規矩」。

查詢與寫入不會交錯執行：查詢從讀模型定義到讀完結果全程擋住寫入，反之亦然。
另一邊跑超過 30 秒時會回 400「…等了 30 秒還沒結束」—— 稍後再試即可，不是壞掉。

### `/api/batch` 為什麼重要

寫 20 個量值 = 20 次 HTTP + 20 次 `SaveChanges()` + 20 次模型重算。批次只重算一次。

更關鍵的是**失敗語意**：預設 `StopOnError=true` 且最後才存檔 —— 任一步失敗，整批變更全部丟棄，模型維持原狀。
只有在後續操作依賴前面結果「已生效」時才需要 `SavePerOp=true`（此時失敗不會回滾）。

### `/api/query` 參數

```json
{ "Query": "EVALUATE ...", "MaxRows": 1000, "TimeoutSeconds": 60, "AskUser": false }
```

- `Query` **必須以 `EVALUATE` 或 `DEFINE` 開頭**（唯讀防護，擋 XMLA 命令）
- `MaxRows` **只能把上限調低，不能調高**。伺服器封頂：逐列明細 `MaxDetailRows`（預設 50）、
  彙總結果 `MaxAggregateRows`（預設 300）。收緊時主控台會印出「列數上限收緊為 N 列」
- 回應的 `truncated` 為 `true` 代表**還有資料沒帶回來** —— 不要當成「這就是全部」，
  尤其不要據此下「總共只有 N 筆」的結論。要總數請改用 `COUNTROWS`
- `TimeoutSeconds` 預設 60，上限 600
- 回應含 `elapsedMs`，可用來比較不同 DAX 寫法的效能
- 被資料保護擋下回 **403**，body 有 `reason`（為什麼）、`hint`（怎麼改寫）、`verdict`（判定代碼）、
  `canAskUser`（能不能用 `-AskUser` 請使用者放行這一次）
- 回應的欄位名稱都是 camelCase（請求的參數名稱不分大小寫）；PowerShell 取屬性不分大小寫，
  所以 `$r.Rows` 和 `$r.rows` 都可以
- DAX 語法錯誤回 **400**，body 含引擎的訊息與錯誤位置 → 據此直接修正。
  查詢碰到受限欄位時，訊息裡引用的資料內容會換成 `‹已遮蔽›`
- `AskUser`：這句查詢被資料保護擋下時，在使用者的螢幕跳確認視窗，由他決定要不要放行這一次。
  沒被擋就照常回傳、不會跳視窗。用法與規矩見「被擋下時該怎麼辦」

---

## 🔄 標準開發循環

**核心原則：寫入後一定要驗算。不驗算就不算完成。**

```
0. New-PbiSnapshot        破壞性操作前先留退路（刪除、改名、覆寫 M）
1. Get-PbiSchema          讀現況，理解既有命名與邏輯
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
（見下方「Power Query M 的開發互動模式」）。

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
  這一存會把 Power BI 記憶體裡的舊版面寫回去，蓋掉你的修改（見「直接編輯 PBIP 報表檔」）

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

---

## 🧪 Power Query M 的開發互動模式

使用者會**自己開著 Power Query 編輯器**（那裡有預覽，是 AI 看不到的東西）。
分工原則：**使用者只做判斷，M 的碼一律由 AI 寫。**

| | 使用者 | Claude |
|---|---|---|
| 看 | 預覽長怎樣、哪一欄怪 | 列數、空值數、相異值數、總和有沒有跑掉 |
| 決定 | 這個轉換對不對 | — |
| 寫 | — | 全部的 M |

### M 可以寫入，但寫入不等於生效

`/api/update-m` 與 `Set-PbiMQuery` 曾於 2026-08-06 移除，**2026-09-16 恢復** —— 當初移除的理由
（「按套用也解不開」）是錯的。

**Power BI Desktop 的 Power Query 文件與 TOM 模型是兩份獨立的東西。**
用 TOM 改 M 只動到模型那份，而且**寫入本身不會重抓資料** —— 寫完那一刻，資料還是舊的。
寫入之後 Desktop 的反應有兩種，**兩種都實際遇過，事先無法預期是哪一種**：

| Desktop 的反應 | 怎麼讓 M 生效 | 什麼時候看到的 |
|---|---|---|
| 顯示「查詢中有暫止的變更尚未套用」 | 使用者按「套用」，Power BI 自己重整 | 2026-09-16 使用者回報 |
| **完全沒有提示** | `Invoke-PbiRefresh -Table <表>`（full）—— 由你跑 | 2026-10-01，PBIP 專案 |

所以寫完 M **一定要問使用者一句**：「Power BI 有沒有出現『查詢中有暫止的變更尚未套用』的提示？」
不要假設會出現，也不要假設不會。兩條硬規則：

- **寫入之前先講要改什麼。** 資料保護啟用時，經由 API 寫 M 會先在使用者的螢幕跳確認視窗
  （M 可以把受限欄位的內容搬進開放的欄位，掃描看不懂 M，所以由人把關）。
  他按「是」才寫入；沒先說明就跳視窗，他只能對一件不知道內容的事做決定
- **有提示時，套用由使用者按。** 不要想辦法用按鍵或 UI 自動化去按它
- **沒提示時，refresh 之前先講一聲**：full refresh 會重抓資料源，大表可能要好幾分鐘，
  而且資料源憑證過期時沒有 UI 可以輸入密碼、會直接失敗

**生效之後要確認留下的是哪一版**：兩邊內容不同時，Desktop 那份可能把 API 寫進去的 M 蓋掉。

```powershell
$before = Get-PbiTableProfile <表名> -Columns <欄位…>   # 基準線
Get-PbiMQuery <表名> -Label before                      # 留一份 .pq 備份
Set-PbiMQuery <表名> -Expression $m                      # 寫入模型那一份
# → 問使用者有沒有出現套用提示。有：請他按套用。沒有：Invoke-PbiRefresh -Table <表名>
(Get-PbiMQuery <表名> -Show) -eq $m                      # True＝留下的是這次寫的版本
Compare-PbiTableProfile $before (Get-PbiTableProfile <表名> -Columns <欄位…>)
```

最後那個比對是在 PowerShell 端做的，進 context 的只有 `True` / `False`，M 的內容不會進來
（`$m` 要在同一次呼叫裡；跨呼叫的話從你存的 `.pq` 檔讀回來）。

**還沒驗證過的事，照保守的方式處理：**

- 走「沒提示 → refresh」那條路之後，**Power Query 編輯器裡那一份有沒有跟著更新，沒有驗證過**。
  請使用者在存檔並重開檔案之前，先不要在編輯器裡「關閉並套用」—— 編輯器那份如果還是舊的，
  套用會把模型蓋回去。重開之後再比對一次留下的版本
- PBIP 還有另一條路：M 以檔案形式存在專案資料夾（TMDL），改檔案再讓 Desktop 重載 ——
  **本工具尚未實作也未驗證**，不要自己臨時發明做法

**M 多輸出了一個新欄位**：模型不會自己長出資料行（結構偵測是 Desktop 套用查詢時才做的事）。
走 API 這條路時要自己補：

```powershell
Set-PbiMQuery <表名> -Expression $m                                   # M 已經輸出新欄位
Add-PbiColumn -Table <表名> -Name <欄名> -SourceColumn <M 裡的欄名> -DataType text   # 型別要和 M 一致
Invoke-PbiRefresh -Table <表名>                                       # 失敗的話 Remove-PbiColumn 拿掉再查原因
```

**共用查詢**（多張表引用的那種，不屬於任何一張表）用 `Set-PbiExpression` 寫，
之後對**用到它的表**重新整理 —— 共用查詢本身沒有資料可以重整。

使用者偏好自己貼也可以 —— 給他完整的 `let...in`，貼完「關閉並套用」時 Power BI 會自己重整
（新欄位也會自己偵測），那條路不需要 `Invoke-PbiRefresh`。

> ⚠️ 還有兩條路也會寫 M，是刻意保留的，用之前要想清楚後果（同樣只改模型那一份，同樣要讓它生效）：
> `/api/restore` 的 `mquery` 範圍（**預設就包含**，`Restore-PbiSnapshot` 不指定 `-Scope` 就會改到 M —— 
> 只想還原量值就明確給 `-Scope measures`）、以及 `/api/create-table` 的 `Kind=m`。

### 循環

**① 開工前（不需要使用者動手）**

```powershell
Get-PbiMQuery DimProduct -Label baseline     # 讀單張表的 M 並留備份
$before = Get-PbiTableProfile DimProduct -Columns Category, ProductName   # 記下基準線
New-PbiSnapshot -Label "before-pq-work"
```

**② 提案時先講人話，不要先丟 code**

> 我打算把 `[TransactionDate]` 從文字轉日期（看起來是 YYYYMMDD）。
> 轉失敗的列傾向設成 null 而不是整批擋掉 —— 這樣壞資料不會讓整張表消失。可以嗎？

使用者回「可以 / 不行，要怎樣」。**不讓他讀 code 做判斷。**

**③ 給一份完整、可直接全選取代的 `let...in`**

**不要給片段** —— 片段會逼使用者自己判斷貼哪裡，那是工作不是判斷。

**④ 使用者只需回三種其中一種**

- `OK` → 進下一步
- `錯` + 直接貼錯誤訊息（不必整理措辭）
- `怪` + 截圖（截圖比打字快，看得懂就好）

**⑤ 套用後做量化驗證 —— 這是預覽抓不到的**

```powershell
$after = Get-PbiTableProfile DimProduct -Columns Category, ProductName
Compare-PbiTableProfile $before $after
```

預覽只給前 1000 列。轉換把後面 30% 的資料弄掉了，預覽看起來完全正常 ——
`Compare-PbiTableProfile` 會直接告訴你「列數 -30.0%」。

**沒有變動時它會安靜**，只列出有差異的指標，不佔版面。

### 注意事項

- `Get-PbiMQuery` **一次只讀一張表**。M 腳本含連線字串、伺服器位址、資料庫名稱，
  整包拉會把所有連線資訊送進 context（見資料保護規範的「M 腳本：預設不輸出內容」）
- 使用者自己「關閉並套用」或按了套用提示時，Power BI 會自己重整，**不需要**再跑 `Invoke-PbiRefresh`；
  用 API 寫 M 而沒有出現提示時才需要（見上方表格）
- 資料源需要認證而 PBI 快取憑證失效時，TOM 觸發的 refresh **沒有 UI 可以輸入密碼**，可能失敗或卡住
- 用 TOM 建立全新的 M 表格時**必須明確指定 Columns**，引擎不會自動推導結構；
  既有的表多輸出欄位時同理，用 `Add-PbiColumn -SourceColumn` 補

---

## 🛑 安全守則

### 修改是寫進記憶體，不是檔案

TOM 的 `SaveChanges()` 只改 Power BI Desktop **記憶體中**的模型。畫面會立刻更新，
但要進到 `.pbip`/`.pbix` 必須存檔。

→ 完成一段工作後跑 `Save-PbiModel -Expect <剛寫的東西>` 並**確認 `expectFound = true`**
   （PBIX 沒辦法檢查內容，只能看 `fileChanged`）；沒存到就如實告訴使用者，並請他手動按 Ctrl+S。

### 寫入前先確認

- **破壞性操作前先 `New-PbiSnapshot`** —— 刪除、改名都算
- 覆寫既有量值前，先用 `/api/schema` 把**原內容記下來**，並在回報中附上
- 請使用者動 M 之前，先 `Get-PbiMQuery <表名> -Label <標籤>` 留一份 `.pq`
- `delete-measure` / `delete-table` / `delete-column` 無法復原，執行前先向使用者確認
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

## 🎨 直接編輯 PBIP 報表檔（版面與視覺）

PBIP 的報表是一堆文字檔：`<名稱>.Report\definition\pages\<頁>\visuals\<視覺>\visual.json`。
**Power BI 開著的時候可以直接改這些檔案** —— Desktop 偵測到檔案變動會請使用者接受變更，
接受之後畫面就換成新的。不用關檔、不殺行程。（使用者 2026-09-24 起實際這樣用。）

這條路沒有 API，改檔案是你自己用腳本做的，所以**安全網也要自己做**：

```
1. 模型的變更先全部做完並存檔          Save-PbiModel -Expect …（之後到步驟 6 之前都不要再存）
2. 把整個 .Report 資料夾複製一份到別處   改壞了可以整包換回來
3. 用腳本改檔案                        一次改完；刪視覺用「搬走資料夾」，不要直接刪
4. Test-PbiReport                     errorCount 必須是 0；warnings 逐條看過
5. 請使用者到 Power BI 接受變更，並截圖給你看
6. 使用者確認畫面沒問題之後，才可以再存檔
```

**踩過的坑（都是實際發生過的）：**

- **順序錯了會把自己的修改蓋掉。** Ctrl+S 存的是整份檔案，包含 Power BI 記憶體裡的版面。
  使用者還沒接受變更時送出 `Save-PbiModel`，舊版面就被寫回磁碟，你改的檔案無聲無息地不見。
  所以步驟 1 一定在前面，步驟 3 之後不准存檔
- **使用者說沒有跳出接受變更的提示** → 請他「**不要存檔**，直接關閉再重開」。
  這時候按 Ctrl+S 一樣會用舊版面蓋掉新檔案
- **布林欄位上的視覺層篩選會被 Power BI 存檔時丟掉**，沒有任何提示。
  要篩選請用文字欄位（例如另建一個「近 2 年／更早」的文字欄）。`Test-PbiReport` 會警告
- **換欄位要連 `queryRef` 一起換。** 條件式格式（底色、字色）是用 `queryRef` 對應欄位的，
  只改 `field` 不改 `queryRef`，或格式設定還指著已經拿掉的欄位，格式就默默失效。`Test-PbiReport` 會警告
- **腳本寫到一半失敗會留下半套檔案。** 先把所有要寫的內容在記憶體裡算好、檢查完，最後才一次寫出；
  寫出前先完成步驟 2 的備份
- **你看不到畫面。** `Test-PbiReport` 查得到結構（欄位存不存在、有沒有重疊、有沒有超出頁面），
  查不到文字被擠壓、字太小、配色難看。每次改完都要請使用者截圖，不要自己宣稱「排版沒問題」

**資料保護與防毒在這裡一樣適用：**

- `visual.json` 的篩選條件裡可能有寫死的真實值（客戶名稱）。**不要把整份檔案印出來讀**；
  用腳本改，只輸出結構資訊（視覺類型、位置、欄位名稱）。`Test-PbiReport` 就是這樣做的
- 報表資料夾常常在桌面或下載資料夾。**一層一層指定路徑去讀，不要對上層資料夾 `-Recurse`**
- 使用者沒說可以之前，不要去讀專案資料夾以外的 PBIP —— 先問

---

## 🧱 專案結構

```
PBI_AI_Bridge/
├── CLAUDE.md                      ← 本文件（AI 工作規範）
├── AGENTS.md                      ← 其他 AI 代理（Codex / Cursor …）的入口，指向本文件
├── README.md                      ← 給人看的完整說明（安裝、使用流程、資料保護清單、疑難排解）
├── 🚀啟動PBI終極儀表板.bat          ← 唯一的啟動檔（自動偵測階段：已在跑就開儀表板 → 缺 SDK／套件時詢問並安裝 → 程式有改才編譯 → 起服務）
│                                     ⚠️ 只能有 ASCII 字元、行尾必須 CRLF（cmd 以位元組定位，否則 goto 會錯位）
├── PowerBI_Visualizer.html        ← 網頁儀表板：「模型結構」與「資料保護」兩個檢視（使用者在後者逐欄設定 AI 能讀什麼）。
│                                     由服務在 http://localhost:5500/ 提供並帶入金鑰（直接雙擊會自動轉址）
├── API_Documentation.html         ← 端點總覽（網頁版）
├── tools/
│   ├── PBI-Bridge.ps1             ← PowerShell 輔助函式（一律用這個呼叫 API）⚠️ 必須保留 UTF-8 BOM
│   ├── Test-DataGuard.ps1         ← 資料保護紅隊測試（只回報擋下與否，不印出資料）。要開著服務與 Power BI
│   └── Test-BridgeClient.ps1      ← 用戶端自我檢查，不需要服務：BOM／控制字元／啟動檔編碼，
│                                     以及每個函式送出的請求。**改過 tools\*.ps1、hook 或啟動檔之後一定要跑**
├── pbibridge_csharp/              ← C# 橋接服務
│   ├── Program.cs                 ← 所有 API 端點與 DataGuard 都在這
│   ├── appsettings.template.json  ← 設定範本（可分享；啟動檔第一次執行時由它產生 appsettings.json）
│   └── appsettings.json           ← 資料保護的通用規則等設定（沒有機密 —— 金鑰不在這裡，見下方「專案資料夾之外」）
│                                     ⚠️ 這裡「沒有」目標檔案路徑 —— 見文件開頭說明
├── .claude/
│   ├── settings.json              ← PreToolUse hook 設定，以及禁止檔案工具讀寫保護設定資料夾的權限規則
│   └── hooks/guard-data-access.ps1 ← 檢查 Bash／PowerShell 指令的文字，擋掉繞過橋接服務的寫法：直接連 ADOMD、
│                                     手刻 HTTP 請求、指令裡出現保護設定資料夾。只看指令文字 —— 指令裡只是「提到」
│                                     那些字串（提交訊息、heredoc 寫文件）也會被擋，那種情況改用 Write／Edit／Grep 工具
│                                     ⚠️ stdin 必須以 UTF-8 讀取：用系統碼頁（cp950）讀，專案路徑的中文會讓
│                                        JSON 解析失敗，而腳本解析失敗時是放行 —— hook 會變成全部放行而且毫無徵兆
├── audit/                         ← 查詢稽核記錄（每月一檔，TSV，服務自動建立）
│                                     只記查詢文字與判定結果，不記回傳值
├── snapshots/                     ← 模型快照，每個模型一個子資料夾（服務自動建立）
│   └── <檔名>_<路徑雜湊>/
└── PowerQuery_Scripts/            ← M 腳本備份 (.pq)，Get-PbiMQuery 第一次執行時建立
    └── <檔名>_<路徑雜湊>/          ← 與 snapshots 相同的命名

（專案資料夾之外 —— 整個 %LOCALAPPDATA%\PBI_AI_Bridge\ 都不要用檔案工具或指令去讀寫）
%LOCALAPPDATA%\PBI_AI_Bridge\api-key.txt
                                   ← API 金鑰。服務第一次啟動時產生；Get-PbiApiKey 會自己讀
%LOCALAPPDATA%\PBI_AI_Bridge\protection\<檔名>_<雜湊>.json
                                   ← 每個模型的逐欄保護設定。刻意不放在專案資料夾：那是 AI 平常讀寫檔案的地方。
                                     只由服務讀寫 —— 不要碰（指令列與檔案工具都有規則在擋，但沒被擋到也不要碰）。
                                     設定跟著「模型」走：先比完整路徑，對不上就靠資料行的固定代號（LineageTag）認，
                                     所以搬家、改名、另存新檔、從 Power BI 裡面開啟都還在
```

**不要在這裡放特定模型的產物**（資料字典、schema 匯出、單一模型的備份）。
這種東西會過期，而過期的參考資料比沒有更糟 —— 需要時用 `Get-PbiSchema`、
`Get-PbiMeasures`、`Get-PbiMQuery` 取得當下的真實狀態。

這個資料夾**只放通用 PBI 開發工具**。不相干的專案請放到這個資料夾外面、各自獨立的地方，
不要混進來 —— 混在一起會讓 `.gitignore`、專案結構、與這份文件全部失真。

## 建置指令

```powershell
dotnet build .\pbibridge_csharp -c Release
```

改完 `Program.cs` 必須重新建置，**並重啟服務**才會生效。

**重啟服務的正確流程（不要自己動手，會觸發防毒）：**

1. 請使用者**自己關掉**主控台視窗（或在該視窗按 Ctrl+C）
2. 使用者回報關好了 → 你才跑 `dotnet build`
   （服務沒停時 DLL 被鎖住，會出現 `MSB3027 檔案鎖定者` —— 那不是程式碼有問題，
   編譯本身是過的，只是複製不進 `bin`）
3. 建置成功 → 請使用者**雙擊 `🚀啟動PBI終極儀表板.bat`**，並保持視窗開啟
4. 使用者回報啟動了 → 你再 `Test-PbiBridge` 驗證

**不要用 `Stop-Process` 停服務，也不要用 `Start-Process` 啟動剛編出來的 `.exe`。**
這兩個動作連在一起（終止行程 → 執行新編譯的未簽章執行檔）是防毒的高危特徵組合。

### ⚠️ 改 `appsettings.json` 不會立刻生效

服務啟動時會 `SetCurrentDirectory(AppContext.BaseDirectory)`，讀的是
**`bin\Release\net8.0\appsettings.json`** —— 那是建置時複製過去的副本，不是專案根目錄那份。

改完設定請使用者**關掉黑窗、重新雙擊 🚀** —— 啟動檔每次啟動都會把專案根目錄那份**直接複製**到 bin（不看時間戳）。
不經過啟動檔、只跑 `dotnet build` 則不保險，因為 MSBuild 是**依時間戳做增量複製**：
如果你是還原備份或用較舊的檔案覆蓋，時間戳沒變新，建置會直接跳過複製，設定依然是舊的。

不經過啟動檔時的保險做法：

```powershell
(Get-Item .\pbibridge_csharp\appsettings.json).LastWriteTime = Get-Date
Copy-Item .\pbibridge_csharp\appsettings.json .\pbibridge_csharp\bin\Release\net8.0\ -Force
```

**驗證方式**：比對兩份的 `Get-FileHash`，或看服務啟動時印出的設定摘要。
不要假設「我改了檔案所以生效了」。

---

## 常見問題對照

| 症狀 | 原因 | 處理 |
|---|---|---|
| `找不到正在執行的 Power BI 檔案` | 沒開 PBI 或沒開檔案 | 請使用者開啟 PBIP/PBIX |
| 400 且訊息是亂碼 | PowerShell 5.1 編碼問題 | 改用 UTF-8 位元組送出 |
| 401 Unauthorized | Key 錯誤或沒帶 Header（手刻請求、或金鑰在這個 PowerShell 工作階段載入之後被換過） | 重新載入 `tools\PBI-Bridge.ps1` 再試。不要自己去讀金鑰檔 |
| `找不到 API 金鑰檔` | 金鑰是服務啟動時產生的：服務還沒啟動過，或黑窗裡跑的是舊版 | 請使用者關掉黑窗、重新雙擊 🚀 |
| 404 Not Found | 執行中的是舊版建置 | 重新 build 並重啟服務（流程見「建置指令」，**由使用者關窗、由使用者開 `.bat`**） |
| `MSB3027 檔案鎖定者` | 服務還在跑，DLL 鎖住 | **編譯是過的**，只差複製。請使用者關掉主控台視窗，再重跑 `dotnet build` |
| `address already in use` / 啟動檔說 `Port 5500 is used by a different program` | 別的程式佔用 5500（常見是 VS Code Live Server） | **不要自己砍行程**（會觸發防毒）。請使用者關掉那個程式再雙擊 🚀。佔用者若是本服務，啟動檔會直接開儀表板、不會報錯 —— 要重啟服務就請使用者先關掉舊的黑窗 |
| `需要重新計算，因此未包含任何資料` | 建了計算表／關聯但沒重算 | `Invoke-PbiRefresh -RefreshType calculate` |
| PBI 顯示「查詢中有暫止的變更尚未套用」 | M 被 API 改過（`Set-PbiMQuery`、`restore` 的 mquery 範圍、`create-table` 的 `Kind=m`） | 請使用者按「套用」即可更新。套用後確認留下的是預期版本；不是的話再到進階編輯器貼上正確的 M 並「關閉並套用」 |
| 用 API 寫了 M，Desktop 沒有任何提示，資料也沒變 | 寫入只改模型裡的 M，不會重抓資料；而這次 Desktop 沒有跳套用提示（兩種情況都遇過） | 先跟使用者講一聲，再 `Invoke-PbiRefresh -Table <表>` |
| refresh 失敗，訊息提到找不到某個資料行 | 用 `Add-PbiColumn -SourceColumn` 加的資料行，M 其實沒有輸出那個欄位（名稱或大小寫不符） | `Remove-PbiColumn` 拿掉，核對 M 輸出的欄名後重加 |
| `Save-PbiModel` 回 `fileChanged = false` | 30 秒內沒有檔案變動：沒有待存變更、按鍵沒送達，或大檔還在寫 | 過一會兒用 `Save-PbiModel -VerifyOnly -Expect …` 看磁碟（不送按鍵）。確定沒存到就如實回報，請使用者手動 Ctrl+S。**不要連按 SendKeys** |
| `Save-PbiModel` 回 `expectFound = false` | 檔案有寫，但剛改的內容不在裡面（推測 Desktop 還沒同步到 API 的寫入就存了） | 等幾秒**最多再存一次**；還是沒有就請使用者手動 Ctrl+S，再用 `-VerifyOnly` 確認 |
| 查量值被 403，訊息說它是受限欄位 | 量值與某個資料行同名 | 這是刻意的，改用不同名稱的量值驗算 |
| 403，`answer` 是 `No` / `Timeout` | 使用者在確認視窗按了「否」，或 120 秒內沒有回答 | 不要重送。如實回報，問使用者想怎麼做。`Timeout` 多半是視窗被蓋住或在另一個螢幕 —— 請他按 Alt+Tab 找，等 `retryAfterSeconds` 過後再送 |
| 403，`answer` 是 `CoolingDown` / `Busy` | 剛被拒絕過（30 秒內不再詢問），或已經有一個確認視窗開著 | 等 `retryAfterSeconds` 秒。期間不要連續重送 |
| 403，`verdict` 是 `BLOCKED-SETTINGS`，訊息說保護設定檔讀不出來 | 這個模型的逐欄設定檔損毀。為了不讓「檔案壞掉＝解除保護」，查詢全部擋下 | 請使用者到儀表板的「資料保護」分頁按「重設」（會跳確認視窗），再重新設定。這個狀況不能用 `-AskUser` 繞過（`canAskUser` 是 `false`） |
| 403，`canAskUser` 是 `false` | 這種擋法不能請使用者放行：寫法本身被禁、模型定義讀不完整、設定檔讀不出來，或查詢太長／碰到太多受限欄位，確認視窗列不完 | 照 `hint` 改寫（拆小、寫短、換彙總寫法）。不要加 `-AskUser` 重送 |
| 403，`verdict` 是 `BLOCKED-CHANGED` | 使用者按「是」之後，這句查詢用到的量值定義和他看到的時候不一樣了 | 重新向使用者說明，再問一次 |
| 409，訊息說設定在等待確認的期間被改過 | 放寬的確認視窗開著的時候，另一個請求（多半是收緊）先存進去了 | `Get-PbiProtection` 重新讀一次，確認現況後再決定要不要重送 |
| `Get-PbiInstances` 的 `kind` 是 `Unknown`、路徑是空的 | 使用者先開 Power BI、再從裡面選檔案，服務拿不到路徑 | 讀寫模型與資料保護照常。`Save-PbiModel`、`Test-PbiReport` 不能用，快照重開就找不回來 —— 需要時請使用者改用雙擊檔案的方式開啟 |
| `Save-PbiModel` 回 400，訊息說不知道檔案路徑 | 同上。服務沒有送 Ctrl+S | 請使用者自己按 Ctrl+S，回報時講明沒有驗證過存檔內容 |
| `Get-PbiProtection -Raw` 的 `inheritedFrom` 有值 | 這份模型的設定是靠內容認回來的（檔案搬過家、改過名，或這次從 Power BI 裡面開） | 正常。請使用者到儀表板掃一眼等級；新增的欄位仍然是沒設定過的 |
| 400「…等了 30 秒還沒結束」 | 查詢與寫入不能交錯：另一邊（長時間的重新整理或查詢）還在跑 | 稍後再試 |
| 錯誤訊息裡有 `‹已遮蔽›` | 查詢碰到受限欄位，引擎訊息裡引用的資料內容被換掉了 | 訊息其餘部分照常可用。原文在服務主控台，需要時請使用者唸給你 |
| 查詢被 403，訊息提到「若 [X] 是你在查詢裡取的別名」 | 結果欄位的別名長得像管制欄位（例如 `"Revenue"`） | 換一個別名（例如 `"值"`）。這不算繞過管制 —— 別名只是標籤 |
| 直接改了報表檔案，使用者接受變更後又變回舊的 | 接受之前有人送了 Ctrl+S（`Save-PbiModel`），舊版面被寫回磁碟 | 從備份的 `.Report` 還原再來一次；順序見「直接編輯 PBIP 報表檔」 |
| 防毒跳警報 | 做了行程終止／遞迴掃描／執行新編譯的 exe | 停下來告訴使用者你剛做了什麼、時間點，讓他對照警報。之後改走「請使用者代勞」的路線 |
| `目前有 N 個 Power BI 實例在執行` | 多個 PBI 開著但沒選目標 | `Use-PbiInstance <檔名片段或 Port>`。**不要為了繞過而隨便挑一個** |
| `先前選定的實例已不存在` | 選定的 PBI 被關掉了 | 重新 `Use-PbiInstance`。（只是換檔重開的話會自動重新解析，不會出現這個） |
| 改到錯的模型 | 沒確認 `Use-PbiInstance` 印出的完整路徑 | 同名檔案很常見，每次切換都要看路徑 |
| 改名後公式壞掉 | 沒有同步改寫引用 | 用 `Rename-PbiObject`（會自動改寫），**別直接改 TOM 的 Name** |
| 建計算群組失敗 | `DiscourageImplicitMeasures` 未開 | 先向使用者說明副作用，同意後加 `-DiscourageImplicitMeasures` |
