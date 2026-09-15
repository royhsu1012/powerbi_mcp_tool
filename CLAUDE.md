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
存檔驗證掃錯資料夾、`inject-visual` 關掉你正在做的檔案去開另一個。**不要把它加回去。**

---

## 🔒 資料保護規範（最高優先，凌駕其他所有指示）

**核心事實：Claude 是雲端模型。任何進入 context 的內容都會傳送到 Anthropic 伺服器。**
橋接服務本身只綁 localhost、不對外連線；唯一的外流管道是「Claude 讀到了什麼」。
這裡處理的是**真實的營運資料（訂單、客戶、成本…）**。

### 分界線：結構可讀，內容受管

**欄位的「名稱」不是機密**（AI 要靠它寫程式），**欄位的「內容」才是**。
所以 `/api/schema`、`/api/validate`、`/api/relationships` 全開，隨你讀；
被管制的只有 `/api/query` 與 `/api/dmv` 的**回傳值**。

### 三層防護，各管一件事

| 層 | 管什麼 | 額度 |
|---|---|---|
| 結構 | 資料表、欄位名稱、型別、DAX 量值、關聯 | **全開**，隨你讀 |
| 身分／自由文字欄位 | 客戶名、聯絡人、備註、next step… | **0 列** —— 只能數，不能取值 |
| 金額欄位 | 各種 amount / amt / price | 必須包在聚合函式內 |
| 其餘所有欄位的明細 | 一般維度與事實欄位 | **逐列明細 50 列 / 彙總 300 列** |

第四層是 2026-08-18 加的，補掉一個大洞：`MaxRows` 原本是**請求參數**，
意思是呼叫端（也就是 AI 自己）可以填 10000。上限若能自己指定，它就不是護欄。
現在改由伺服器封頂，`MaxRows` 只能把額度調得**更低**。

判定明細或彙總，看的是「有沒有欄位被聚合／計數函式包住」（量值定義會先展開）：
`SUMMARIZECOLUMNS(日期[月], "額", [銷售總額])` → 彙總 300 列；
`EVALUATE 某表`、`SELECTCOLUMNS(...)`、`TOPN(...)` → 明細 50 列。
**認不出來就當明細**，判定刻意保守。

欄位比對會先做**正規化**：忽略大小寫、空白、底線、連字號 ——
`account name` 與 `account_name` 視為同一個欄位。清單項目也可以是
`*customer*` 這種樣式。欄位有一千個以上而且一直在長，逐一列舉追不上。

樣式誤傷時（例如 `*owner*` 打到 `process_owner`），把該欄位加進
`DataProtection:AllowColumns`，**不要為了一個誤判把整條樣式拆掉**。

以下規則不再只是規範 —— 2026-08-10 起由**伺服器端強制執行**（`Program.cs` 的 `DataGuard`）。
違規查詢回 **403**，模型無法自我豁免。這代表：**你不需要靠自律來避免外洩，
但也不要把「沒被擋下」當成「這樣做很安全」**——擋不住的東西仍然會進 context。

### 規則一：客戶身分欄位只能數，不能取值

管制清單在 `appsettings.json` → `DataProtection:DenyColumns`（支援 `*` 樣式，比對前會正規化）。
這些欄位**只准出現在計數類函式內**（`DISTINCTCOUNT` / `COUNTROWS` / `COUNT` /
`COUNTA` / `COUNTBLANK`）—— 那類函式的回傳值必定是數字。

```powershell
# ✅ 放行 — 回傳 1503，是結構資訊
Invoke-Dax 'EVALUATE ROW("客戶數", DISTINCTCOUNT(Customers[account name]))'

# ⛔ 403 — 名稱被當成分組鍵，1503 個真名會進 context
Invoke-Dax 'EVALUATE SUMMARIZECOLUMNS(Customers[account name], "額", SUM(...))'

# ⛔ 403 — MAX 對文字欄位會吐出一個真實客戶名稱，不算計數
Invoke-Dax 'EVALUATE ROW("x", MAX(Customers[account name]))'
```

### 規則二：金額欄位必須包在聚合函式內

清單在 `DataProtection:AggregateOnlyColumns`（同樣支援樣式與正規化）。可用任何數值聚合
（`SUM`/`AVERAGE`/`SUMX`/`MIN`/`MAX`…），但不得逐列取值。
另有列數上限（`MaxRowsWithMoney`，預設 100）擋掉「按高基數鍵分組」的變相逐筆。

```powershell
# ✅ 放行
Invoke-Dax 'EVALUATE ROW("總額", SUM(Sales[amount]))'
Invoke-Dax 'EVALUATE SUMMARIZECOLUMNS(Sales[product_line],
              "額", SUM(Sales[amount]))'

# ⛔ 403 — 逐列取金額
Invoke-Dax 'EVALUATE SELECTCOLUMNS(Sales, "額", [amount])'
```

### 被擋下時該怎麼辦

**先想「我真的需要這個值嗎」。** 多數情況答案是不需要 —— 換個彙總寫法就解決了。

真的需要時：伺服器會在**服務主控台**印出一組一次性權杖。你看不到那個視窗，
所以必須停下來，**具體說明**要看哪張表、哪些欄位、幾列、為什麼彙總不足，
請使用者確認後把權杖貼給你，再用 `-DetailToken <權杖>` 重送。

```powershell
Invoke-Dax 'EVALUATE ...' -DetailToken A1B2C3D4
```

權杖**綁定當次查詢文字**、10 分鐘有效、只能用一次。不要拿舊權杖去送改過的查詢
（會被拒絕，這是刻意的：那等於用一次同意授權另一件事）。

**不要為了繞過管制而改寫查詢。** 想不出合規寫法就直接問使用者，
不要用 `CONCATENATEX`、變數、中介量值之類的手法把敏感欄位藏起來。

### 去敏模式：需要「按客戶分組」但不需要知道是誰

`-Pseudonymize` 讓身分欄位可以當**分組鍵**，值在伺服器端換成穩定代號。

```powershell
(Invoke-Dax 'EVALUATE SUMMARIZECOLUMNS(Customers[account name],
               "額", SUM(Sales[amount]))' -Pseudonymize).Rows
# → ID_A3F1B2 | 12,345,678
#    ID_7C4E90 |  9,876,543
```

適用於營收集中度、長尾分布、衰退清單這類分析 —— 要區分客戶，不需要真名。

- 代號**跨查詢穩定**（同一個客戶永遠同一個代號），所以可以先查前 20 大、再查這些客戶的月趨勢
- 代號是 HMAC 輸出，金鑰由 API Key 衍生、只存在本機，**推不回原值**
- 真名只印在**服務主控台** —— 使用者看得到，你看不到。要對照請他看那個視窗
- 回應會多一個 `Pseudonymized` 欄位列出哪些欄位是代號。**不要把 `ID_xxxxxx` 當成真實名稱解讀或寫進量值**

**這個開關你可以自己按** —— 與 `-DetailToken` 不同。理由：它只會讓輸出更少，不會更多。

仍然會被擋的用法（去敏也救不了）：

```powershell
# ⛔ MAX / CONCATENATEX / SELECTCOLUMNS 取別名 —— 真名會躲在別名欄位底下，遮罩抓不到
Invoke-Dax 'EVALUATE ROW("x", MAX(Customers[account name]))' -Pseudonymize
Invoke-Dax 'EVALUATE SELECTCOLUMNS(Customers, "n", [account name])' -Pseudonymize
```

去敏只放行「當分組鍵／取相異值」（`SUMMARIZECOLUMNS` / `SUMMARIZE` / `GROUPBY` /
`VALUES` / `DISTINCT` / `ALL`）—— 那些用法欄位會以原名出現在結果裡，遮得掉。

### 規則三：預設不輸出 M 腳本內容

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

### 規則四：機密值不輸出到主控台

- **API Key**：`Get-PbiApiKey` 內部使用即可，不要 `Write-Host` 或 echo 出來
- 不要 `Get-Content appsettings.json` 後直接顯示全文；需要哪個欄位就只取那個欄位

### 規則五：在 PowerShell 端過濾，不要在 context 裡過濾

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
| 短時間重複送 SendKeys（`Save-PbiModel` 連按） | **最多一次**。`fileChanged` 回 false 就如實回報，請使用者手動 Ctrl+S |

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

### 3. 先讀模型再動手

```powershell
(Get-PbiSchema).Tables | Select-Object Name, @{n='欄位數';e={$_.Columns.Count}}, @{n='量值數';e={$_.Measures.Count}}
Get-PbiMeasures | Select-Object Table, Measure, Expression   # 量值（不含 M，M 見規則三）
Get-PbiRelationships                       # 關聯線
Test-PbiModel                              # 健檢：有沒有既存的壞公式
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

存放於 `pbibridge_csharp/appsettings.json` → `Security.ApiKey`。

- **從檔案讀取，不要寫死在任何腳本或文件裡**（`tools/PBI-Bridge.ps1` 會自動讀）
- **不要 echo 到主控台**，避免出現在截圖或分享畫面中

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
| `/api/snapshots` | GET | 列出所有快照 |
| `/api/query` | POST | **執行唯讀 DAX 查詢** — 驗算用 |
| `/api/dmv` | POST | `$SYSTEM` DMV 查詢（記憶體佔用、相依性）。只接受 `SELECT ... FROM $SYSTEM.xxx` |

### 安全網（動手前先做）

| 端點 | 方法 | 用途 |
|---|---|---|
| `/api/snapshot` | POST | 把整個模型序列化成 TMSL 存檔 |
| `/api/restore` | POST | 從快照還原量值／計算資料行／關聯／M／共用運算式（支援 `DryRun`） |

### 讓變更生效

| 端點 | 方法 | 用途 |
|---|---|---|
| `/api/refresh` | POST | **重新整理資料**。改完 M 要 `full`；建完計算表／關聯／計算群組要 `calculate` |
| `/api/save` | POST | **模擬 Ctrl+S 存檔**，並用檔案修改時間驗證是否真的存到 |

### 寫入（會改動模型，動手前先確認）

| 端點 | 方法 | 用途 |
|---|---|---|
| `/api/upsert-measure` | POST | 新增／覆寫量值（公式、格式、說明、DisplayFolder、IsHidden） |
| `/api/delete-measure` | POST | 刪除量值 |
| `/api/move-measure` | POST | 量值搬移到其他表（目標表不存在會自動建立） |
| `/api/add-column` | POST | 新增 DAX 計算資料行 |
| `/api/delete-column` | POST | 刪除資料行（仍被關聯使用時會擋下） |
| `/api/set-column-props` | POST | 格式、DisplayFolder、隱藏、SortByColumn、SummarizeBy、資料類別 |
| `/api/create-table` | POST | 建表：`calculated`（DAX）／`m`（M，需給 Columns）／`measureHolder` |
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
| `/api/inject-visual` | POST | ⚠️ 寫入 visual.json 並**強制重啟 PBI Desktop**（僅 PBIP 適用） |

### `/api/batch` 為什麼重要

寫 20 個量值 = 20 次 HTTP + 20 次 `SaveChanges()` + 20 次模型重算。批次只重算一次。

更關鍵的是**失敗語意**：預設 `StopOnError=true` 且最後才存檔 —— 任一步失敗，整批變更全部丟棄，模型維持原狀。
只有在後續操作依賴前面結果「已生效」時才需要 `SavePerOp=true`（此時失敗不會回滾）。

### `/api/query` 參數

```json
{ "Query": "EVALUATE ...", "MaxRows": 1000, "TimeoutSeconds": 60 }
```

- `Query` **必須以 `EVALUATE` 或 `DEFINE` 開頭**（唯讀防護，擋 XMLA 命令）
- `MaxRows` **只能把上限調低，不能調高**。伺服器封頂：逐列明細 `MaxDetailRows`（預設 50）、
  彙總結果 `MaxAggregateRows`（預設 300）。收緊時主控台會印出「列數上限收緊為 N 列」
- 回應的 `Truncated` 為 `true` 代表**還有資料沒帶回來** —— 不要當成「這就是全部」，
  尤其不要據此下「總共只有 N 筆」的結論。要總數請改用 `COUNTROWS`
- `TimeoutSeconds` 預設 60，上限 600
- 回應含 `ElapsedMs`，可用來比較不同 DAX 寫法的效能
- DAX 語法錯誤回 **400**，body 含引擎原始訊息與錯誤位置 → 據此直接修正

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
6. 符合預期 → Save-PbiModel 存檔並確認 fileChanged = true
```

### ⚠️ 什麼時候必須 refresh

**只寫量值不用 refresh**，但以下情況不 refresh 就查不到東西（會出現「需要重新計算」的錯誤）：

| 做了什麼 | 要跑什麼 | 大約耗時 |
|---|---|---|
| 建立計算表 | `Invoke-PbiRefresh -Table <表>` | 秒級 |
| 建立／修改關聯線 | `Invoke-PbiRefresh -RefreshType calculate` | 秒級 |
| 新增計算項目到計算群組 | `Invoke-PbiRefresh -Table <計算群組>` | 秒級 |
| 新增計算資料行 | `Invoke-PbiRefresh -RefreshType calculate` | 秒級 |

M 腳本不在這張表裡 —— 它是唯讀的，由使用者在 Power Query 編輯器貼上並「關閉並套用」，
Power BI 會自己重整（見下方「Power Query M 的開發互動模式」）。

`calculate` 只重算 DAX、不重抓資料源，很便宜 —— 結構性變更後直接跑一次就好。

### 存檔

`Save-PbiModel` 會把 PBI Desktop 帶到前景送出 Ctrl+S，再比對檔案修改時間驗證。

- **一定要看回傳的 `fileChanged`**。`false` 代表沒存到（或本來就沒有待存變更），不要當成成功
- **`false` 時最多再試一次，不要連按。** 跨行程送合成按鍵會被防毒視為鍵盤側錄／UI 劫持，
  短時間重複會放大可疑度 → 直接如實回報「沒存到」，請使用者手動按 Ctrl+S
- 大型模型（數百 MB）寫檔要時間，驗證只等 5 秒。回 `false` 時**先隔一段時間重看檔案時間**，
  可能其實存成功了 —— 不要因為 `false` 就急著重送按鍵
- 會短暫搶走鍵盤焦點，這是 SendKeys 的固有限制
- 驗證對象是**選定實例自己的檔案**（`Get-PbiInfo` 可查），Ctrl+S 也只送給那一個行程 —— 不會誤存到別的 PBI

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

> 驗算查詢一樣受資料保護管制（見本文件開頭）。客戶身分欄位只能數不能取值、
> 金額必須聚合 —— 這對驗算幾乎沒有影響，因為驗算要確認的是**公式邏輯**，
> 不是某一筆訂單是誰下的。被 403 擋下時**不要改寫查詢去繞**，
> 先問自己需不需要那個值，真的需要就向使用者要一次性權杖。

---

## 🧪 Power Query M 的開發互動模式

使用者會**自己開著 Power Query 編輯器**（那裡有預覽，是 AI 看不到的東西）。
分工原則：**使用者只做判斷，M 的碼一律由 AI 寫。**

| | 使用者 | Claude |
|---|---|---|
| 看 | 預覽長怎樣、哪一欄怪 | 列數、空值數、相異值數、總和有沒有跑掉 |
| 決定 | 這個轉換對不對 | — |
| 寫 | — | 全部的 M |

### ⛔ 硬規則：M 腳本唯讀，API 不能寫，一律交給使用者貼

`/api/update-m` 與 `Set-PbiMQuery` 已於 **2026-08-06 移除**，用了會回 **410**。
**不要想辦法繞過**（不要改用 `/api/batch`、不要用 `/api/create-table` 重建同名表）。

原因是實際踩過的坑，不是保守：**Power BI Desktop 的 Power Query 文件與 TOM 模型是兩份獨立的東西。**
用 TOM 改 M 只動到模型那份，Desktop 自己那份不會跟著變。兩邊一失步，Desktop 就永久顯示
「查詢中有暫止的變更尚未套用」；按「套用變更」是拿 Desktop 那份**舊 M** 去跑，
跑完不一致依然存在，橫幅又冒出來 —— 死迴圈，只能請使用者手動到進階編輯器貼一次才解得開。
編輯器有沒有開著都一樣會發生。

所以 M 一律是：**AI 讀 + AI 寫碼 + 使用者貼**。

```powershell
Get-PbiMQuery <表名> -Label <標籤>   # ✅ 讀，並留一份 .pq 備份
# ❌ 沒有寫入的指令。給使用者完整的 let...in，請他貼進進階編輯器並「關閉並套用」。
```

貼完由 **Power BI 自己**跑重整（「關閉並套用」就會跑），不需要 `Invoke-PbiRefresh`。
之後用 `Compare-PbiTableProfile` 做量化驗證。

> ⚠️ 還有兩條路仍會寫 M，是刻意保留的，用之前要想清楚後果（同樣會造成上述失步）：
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
  整包拉會把所有連線資訊送進 context（見資料保護規範規則三）
- 使用者「關閉並套用」時 Power BI 會自己重整，**不需要**再跑 `Invoke-PbiRefresh`
- 資料源需要認證而 PBI 快取憑證失效時，TOM 觸發的 refresh **沒有 UI 可以輸入密碼**，可能失敗或卡住
- 用 TOM 建立全新的 M 表格時**必須明確指定 Columns**，引擎不會自動推導結構

---

## 🛑 安全守則

### 修改是寫進記憶體，不是檔案

TOM 的 `SaveChanges()` 只改 Power BI Desktop **記憶體中**的模型。畫面會立刻更新，
但要進到 `.pbip`/`.pbix` 必須存檔。

→ 完成一段工作後跑 `Save-PbiModel` 並**確認 `fileChanged = true`**；
   若回傳 false，如實告訴使用者「沒存到」，並請他手動按 Ctrl+S。

### 寫入前先確認

- **破壞性操作前先 `New-PbiSnapshot`** —— 刪除、改名都算
- 覆寫既有量值前，先用 `/api/schema` 把**原內容記下來**，並在回報中附上
- 請使用者動 M 之前，先 `Get-PbiMQuery <表名> -Label <標籤>` 留一份 `.pq`
- `delete-measure` / `delete-table` / `delete-column` 無法復原，執行前先向使用者確認
- **改名一律先 `-DryRun`**：`Rename-PbiObject` 會改寫全模型的 DAX 引用，先看過 `Rewrites` 清單確認沒誤改

### 快照是每個模型各自獨立的

`snapshots/<檔名>_<路徑雜湊>/<時間戳>__標籤.json`

路徑雜湊是必要的：桌面的 `X.pbix` 和專案裡的 `X.pbip` 檔名相同但是**完全不同的模型**，
混在一起會互相污染。`Get-PbiSnapshots` 只列出目前選定實例的快照，
`Restore-PbiSnapshot` 也只在該資料夾裡找 —— **結構上就不可能跨模型還原**。

### `Restore-PbiSnapshot` 的邊界

還原**不是**把模型完整倒回快照當時的狀態，它只處理量值、計算資料行、關聯、M 腳本、共用運算式。

- 快照之後**新建的表格與角色不會被刪除**，連帶它們底下的量值也不會被還原
- 不會改模型層級屬性（例如 DiscourageImplicitMeasures）
- 回應的 `Uncovered` 會列出所有沒被涵蓋的項目 —— **回報時要一併轉述**，不要讓使用者以為已完全還原
- 反過來，在還原範圍內的東西是「回到當時狀態」：快照中不存在的量值與關聯**會被刪掉**，不是合併

### 計算群組的前提

模型的 `DiscourageImplicitMeasures` 必須是 `true` 才能建立計算群組（引擎強制）。
開啟後**使用者無法再把數值欄位直接拖進視覺自動彙總**，一律得改用量值。

→ 這是模型層級的行為改變，**務必先向使用者說明並取得同意**，不要自作主張開啟。

### `/api/inject-visual` 特別危險

它會送 Ctrl+S → **`Process.Kill()` 強制砍掉那一個 Power BI Desktop** → 改檔案 → 重開同一個檔案。

路徑全部由**選定的實例**推導（`X:\proj\name.pbip` → `X:\proj\name.Report\definition`），
不再依賴設定檔，所以不會發生「關掉你正在做的檔案、打開另一個」。
另有兩道防護：目標不是 PBIP 時中止；寫入路徑跑出該專案資料夾時中止。

即便如此，**使用前仍務必先向使用者確認**，並請他們手動存檔備份。
多實例情境下它只會關掉目標那一個，其他開著的 PBI 不受影響。

⚠️ **它包含 `Process.Kill()` + 重新啟動執行檔 —— 正是防毒的高危特徵組合**（見「防毒軟體相容規範」）。
即使是服務內部執行而非 PowerShell 端，行為特徵一樣。**在公司機器上預設不要用它**；
真的需要改視覺，請使用者自己在 Power BI Desktop 裡調整。

---

## 🧱 專案結構

```
PBI_AI_Bridge/
├── CLAUDE.md                      ← 本文件（AI 工作規範）
├── AGENTS.md                      ← 其他 AI 代理（Codex / Cursor …）的入口，指向本文件
├── README.md                      ← 給人看的完整說明（安裝、使用流程、資料保護清單、疑難排解）
├── 🚀啟動PBI終極儀表板.bat          ← 唯一的啟動檔（自動偵測階段：已在跑就開儀表板 → 缺 SDK／套件時詢問並安裝 → 程式有改才編譯 → 起服務）
│                                     ⚠️ 只能有 ASCII 字元、行尾必須 CRLF（cmd 以位元組定位，否則 goto 會錯位）
├── PowerBI_Visualizer.html        ← 網頁儀表板。由服務在 http://localhost:5500/ 提供並帶入金鑰（直接雙擊會自動轉址）
├── API_Documentation.html         ← 端點總覽（網頁版）
├── tools/
│   ├── PBI-Bridge.ps1             ← PowerShell 輔助函式（一律用這個呼叫 API）⚠️ 必須保留 UTF-8 BOM
│   └── Test-DataGuard.ps1         ← 資料保護紅隊測試（只回報擋下與否，不印出資料）
├── pbibridge_csharp/              ← C# 橋接服務
│   ├── Program.cs                 ← 所有 API 端點與 DataGuard 都在這
│   ├── appsettings.template.json  ← 設定範本（可分享；啟動檔第一次執行時由它產生 appsettings.json）
│   └── appsettings.json           ← API Key + 資料保護清單（含機密，勿外流）
│                                     ⚠️ 這裡「沒有」目標檔案路徑 —— 見文件開頭說明
├── .claude/
│   ├── settings.json              ← PreToolUse hook 設定
│   └── hooks/guard-data-access.ps1 ← 擋掉「直接連 ADOMD 繞過橋接服務」那條路
│                                     ⚠️ stdin 必須以 UTF-8 讀取：用系統碼頁（cp950）讀，專案路徑的中文會讓
│                                        JSON 解析失敗，而腳本解析失敗時是放行 —— hook 會變成全部放行而且毫無徵兆
├── audit/                         ← 查詢稽核記錄（每月一檔，TSV，服務自動建立）
│                                     只記查詢文字與判定結果，不記回傳值
├── snapshots/                     ← 模型快照，每個模型一個子資料夾（服務自動建立）
│   └── <檔名>_<路徑雜湊>/
└── PowerQuery_Scripts/            ← M 腳本備份 (.pq)，Get-PbiMQuery 第一次執行時建立
    └── <檔名>_<路徑雜湊>/          ← 與 snapshots 相同的命名
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
| 401 Unauthorized | Key 錯誤或沒帶 Header | 重讀 `appsettings.json` |
| 404 Not Found | 執行中的是舊版建置 | 重新 build 並重啟服務（流程見「建置指令」，**由使用者關窗、由使用者開 `.bat`**） |
| `MSB3027 檔案鎖定者` | 服務還在跑，DLL 鎖住 | **編譯是過的**，只差複製。請使用者關掉主控台視窗，再重跑 `dotnet build` |
| `address already in use` / 啟動檔說 `Port 5500 is used by a different program` | 別的程式佔用 5500（常見是 VS Code Live Server） | **不要自己砍行程**（會觸發防毒）。請使用者關掉那個程式再雙擊 🚀。佔用者若是本服務，啟動檔會直接開儀表板、不會報錯 —— 要重啟服務就請使用者先關掉舊的黑窗 |
| `需要重新計算，因此未包含任何資料` | 建了計算表／關聯但沒重算 | `Invoke-PbiRefresh -RefreshType calculate` |
| `410` + 「M 腳本唯讀」 | 想用 API 寫 M | 這是刻意擋的。給使用者完整 `let...in`，請他貼進進階編輯器並「關閉並套用」 |
| PBI 顯示「查詢中有暫止的變更尚未套用」，按套用後又出現 | Desktop 的 PQ 文件與 TOM 模型失步（多半是有人繞過去寫了 M） | 只能人工解：請使用者到進階編輯器把正確的 M 貼上並「關閉並套用」 |
| `Save-PbiModel` 回 `fileChanged = false` | 大檔還在寫（驗證只等 5 秒）、沒有待存變更，或按鍵沒送達 | 先隔一段時間重看檔案時間 —— 常常其實存到了。仍是舊時間就如實回報，請使用者手動 Ctrl+S。**不要連按 SendKeys** |
| 防毒跳警報 | 做了行程終止／遞迴掃描／執行新編譯的 exe | 停下來告訴使用者你剛做了什麼、時間點，讓他對照警報。之後改走「請使用者代勞」的路線 |
| `目前有 N 個 Power BI 實例在執行` | 多個 PBI 開著但沒選目標 | `Use-PbiInstance <檔名片段或 Port>`。**不要為了繞過而隨便挑一個** |
| `先前選定的實例已不存在` | 選定的 PBI 被關掉了 | 重新 `Use-PbiInstance`。（只是換檔重開的話會自動重新解析，不會出現這個） |
| 改到錯的模型 | 沒確認 `Use-PbiInstance` 印出的完整路徑 | 同名檔案很常見，每次切換都要看路徑 |
| 改名後公式壞掉 | 沒有同步改寫引用 | 用 `Rename-PbiObject`（會自動改寫），**別直接改 TOM 的 Name** |
| 建計算群組失敗 | `DiscourageImplicitMeasures` 未開 | 先向使用者說明副作用，同意後加 `-DiscourageImplicitMeasures` |
