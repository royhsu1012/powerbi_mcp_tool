# ⚡ PBI AI Bridge

讓 AI 代理（Claude Code 等）直接讀寫**執行中的 Power BI Desktop** 資料模型的本地橋接服務。

- 透過 TOM 讀寫**記憶體中**的模型：量值、計算資料行、關聯、計算群組、RLS 角色……
- 透過 ADOMD 執行 DAX 查詢**當場驗算** —— 寫完立刻查，不必存檔、不必重開
- 可同時開多個 PBIX / PBIP，隨時切換操作對象
- 伺服器端強制**資料保護**：AI 讀得到結構，拿不到客戶名稱與逐筆金額

```
AI 代理（PowerShell）            瀏覽器（儀表板）
          ↕                            ↕
   HTTP + X-API-Key + X-PBI-Target   （只綁 localhost:5500）
          ↕
   pbibridge_csharp/   C# 本地服務
          ↕  TOM（寫入）／ADOMD（查詢）
Power BI Desktop #1   Power BI Desktop #2   …
```

---

## 📖 文件導覽

| 你是… | 看這份 |
|---|---|
| **第一次拿到、要安裝使用的人** | **[使用說明.md](使用說明.md)** —— 從安裝到日常操作的逐步教學 |
| 想快速了解全貌的人 | 本文件 |
| AI 代理（Claude Code 自動載入） | [CLAUDE.md](CLAUDE.md) —— 完整工作規範 |
| 其他 AI 代理（Codex / Cursor / Antigravity…） | [AGENTS.md](AGENTS.md)，再讀 CLAUDE.md |
| 查 API 端點 | [API_Documentation.html](API_Documentation.html) |

---

## 🚀 快速開始

1. **解壓縮到本機資料夾**，例如 `C:\PBI_AI_Bridge`（不要放 OneDrive，不要在 zip 裡直接執行）
2. 雙擊 **`📦第一次使用請先點我.bat`**（只需一次）
3. **編輯 `pbibridge_csharp\appsettings.json` 的資料保護清單**，補上你模型裡的敏感欄位（見下方）
4. 開啟一份 Power BI 檔案 → 雙擊 **`🚀啟動PBI終極儀表板.bat`** → **保持黑窗開著**
5. 在這個資料夾開啟 Claude Code，用中文告訴它你要做什麼

每一步的畫面與錯誤處理見 [使用說明.md](使用說明.md)。

## 系統需求

| 項目 | 說明 |
|---|---|
| Windows 10 / 11 | 僅支援 Windows |
| **.NET 8 SDK** | ⚠️ 必須是 **SDK**，只有 Runtime 無法編譯。<https://dotnet.microsoft.com/download/dotnet/8.0> →「.NET SDK 8.x.x」Windows x64 |
| Power BI Desktop | 要**開著至少一份檔案**才有東西可以連 |
| Windows PowerShell 5.1 | 系統內建 |
| **網路（僅第一次編譯）** | 第一次編譯會從 **nuget.org** 下載 Analysis Services 元件，之後離線可用。公司 Proxy／防火牆擋 nuget.org 時請找 IT |
| AI 代理（選用） | Claude Code（建議），或其他能執行 PowerShell 的代理 |

---

## 安裝檔做了什麼

`📦第一次使用請先點我.bat` 只需要跑一次：

| 步驟 | 內容 |
|---|---|
| 1/5 | 確認收到的檔案完整 |
| 2/5 | 檢查 .NET SDK（分辨得出「只裝了 Runtime」，並幫你開下載頁） |
| 3/5 | 檢查資料夾位置（在 OneDrive 裡會警告） |
| 4/5 | 由範本產生 `appsettings.json`，**填入一把只屬於這台機器的隨機 API Key** |
| 5/5 | 編譯 Release 版（第一次約 1 分鐘，含下載 NuGet 套件） |

跑完會印出 API Key。網頁儀表板第一次會問；AI 使用的 `tools/PBI-Bridge.ps1` 會自己讀，不必手動輸入。
安裝檔不會碰你的 Power BI 檔案。

> **每台機器各自產生 Key，不要共用，也不要把 `appsettings.json` 傳給別人。**

## ⚙️ 安裝後第一件事：調整資料保護清單

`pbibridge_csharp/appsettings.json` → `DataProtection` 預設只有英文通用樣式（`*customer*`、`*amount*`…）。
**模型若用中文欄位名，預設清單幾乎擋不到東西**，請補上你自己的敏感欄位。

| 清單 | 放什麼 | 效果 |
|---|---|---|
| `DenyColumns` | 客戶名、聯絡人、電話、備註等身分／自由文字欄位 | 只能計數，不能取值 |
| `AggregateOnlyColumns` | 金額、單價、成本、薪資 | 必須包在 SUM／AVERAGE 等聚合內 |
| `AllowColumns` | 被樣式誤傷、其實不敏感的欄位 | 優先放行 |

比對時忽略大小寫、空白、底線、連字號，可用 `*` 樣式（例如 `*客戶*`）。
改完**關掉黑窗、重新雙擊 🚀** 就生效 —— 啟動器每次都會重新編譯並把設定複製到 `bin\`。

範例與驗證方法見 [使用說明.md 第 4 節](使用說明.md#4-設定資料保護清單重要)。

---

## 兩種使用方式

### A. 網頁儀表板（唯讀瀏覽）

啟動器會自動開啟 `PowerBI_Visualizer.html`：瀏覽資料表、欄位、DAX 量值、M 腳本。
右上角可切換要看哪一個 Power BI。第一次使用要輸入 API Key，瀏覽器會記住。

### B. AI 代理（開發）

在這個資料夾開啟 Claude Code，它會自動載入 `CLAUDE.md` 並自己呼叫 `tools/PBI-Bridge.ps1`。直接描述需求即可：

> 跑一次模型健檢，列出壞掉的公式
>
> 新增「毛利率」量值 = [毛利] / [營收]，放在「量值」表、百分比格式，做完用產品線驗算

也可以自己在 PowerShell 操作：

```powershell
. .\tools\PBI-Bridge.ps1
Test-PbiBridge                                  # 服務狀態 + 有哪些 PBI + 目前目標
Use-PbiInstance 銷售報表                          # 多個 PBI 時先選定（檔名片段或 Port）
New-PbiSnapshot -Label before                   # 動手前留退路
Set-PbiMeasure -Table 量值 -Name 銷售總額 -Expression 'SUM(FactSales[Amount])' -Format '#,0'
(Invoke-Dax 'EVALUATE ROW("結果", [銷售總額])').Rows   # 立刻驗算
Save-PbiModel                                   # 確認回傳 fileChanged = true
```

**不要自己手刻 `Invoke-RestMethod`**：PowerShell 5.1 預設不以 UTF-8 送出，中文欄位名與 DAX 會壞掉；`PBI-Bridge.ps1` 已處理好。

---

## 能做什麼

| 類別 | 指令 |
|---|---|
| 實例切換 | `Get-PbiInstances` / `Use-PbiInstance` / `Get-PbiInfo` |
| 讀取結構 | `Get-PbiSchema` / `Get-PbiMeasures` / `Get-PbiRelationships` / `Get-PbiRoles` / `Get-PbiExpressions` |
| 健檢 | `Test-PbiBridge` / `Test-PbiModel`（壞公式、雙向關聯、孤島表、疑似沒用的欄位） |
| 查詢驗算 | `Invoke-Dax` / `Invoke-PbiDmv` / `Get-PbiModelStats` |
| 量值 | `Set-PbiMeasure` / `Remove-PbiMeasure` / `Move-PbiMeasure` |
| 結構 | `Add-PbiColumn` / `Set-PbiColumn` / `Remove-PbiColumn` / `New-PbiTable` / `Remove-PbiTable` / `Rename-PbiObject`（同步改寫所有引用） |
| 關聯 | `Set-PbiRelationship` / `Remove-PbiRelationship` |
| 進階 | `New-PbiCalcGroup` / `Set-PbiCalcItem` / `Set-PbiRole` / `Set-PbiExpression` |
| 批次 | `Invoke-PbiBatch`（一次存檔；任一步失敗整批不套用） |
| Power Query | `Get-PbiMQuery`（唯讀）/ `Get-PbiTableProfile` / `Compare-PbiTableProfile` |
| 安全網 | `New-PbiSnapshot` / `Get-PbiSnapshots` / `Restore-PbiSnapshot` |
| 生效 | `Invoke-PbiRefresh` / `Save-PbiModel` |

各函式的用途與範例寫在 `tools/PBI-Bridge.ps1` 的註解裡；參數語法可用 `Get-Command Set-PbiMeasure -Syntax` 查。

### 標準開發循環

```
New-PbiSnapshot → Get-PbiSchema → 寫入 → Invoke-PbiRefresh（結構性變更才要）→ Invoke-Dax 驗算 → Save-PbiModel
```

| 做了什麼 | 需要 refresh 嗎 |
|---|---|
| 新增／修改量值 | 不用 |
| 建計算表、新增計算項目 | `Invoke-PbiRefresh -Table <表>` |
| 建／改關聯、新增計算資料行 | `Invoke-PbiRefresh -RefreshType calculate` |
| 改 M 腳本 | 不用 —— M 由使用者在進階編輯器貼上並「關閉並套用」，Power BI 會自己重整 |

**修改只進記憶體**：畫面立刻更新，但要寫進 `.pbix` / `.pbip` 必須 `Save-PbiModel`（或手動 Ctrl+S）。

---

## 🔒 資料保護

**AI 是雲端模型 —— 任何進入對話的內容都會傳到雲端。** 橋接服務只綁 localhost、不對外連線，唯一的外流管道是「AI 讀到了什麼」。

分界線：**欄位名稱可讀，欄位內容受管**。schema／關聯／健檢全開；被管制的只有 `/api/query`、`/api/dmv` 的回傳值。

由伺服器端強制執行（`Program.cs` 的 `DataGuard`），違規回 403，AI 無法自我豁免：

- **身分欄位**只能出現在計數函式內（`DISTINCTCOUNT` / `COUNTROWS` …）
- **金額欄位**必須聚合，且分組列數有上限
- **回傳列數封頂**：逐列明細 50 列、彙總 300 列（呼叫端只能調低）
- **真的需要看明細**：服務黑窗會印出一次性權杖（綁定該句查詢、10 分鐘、限用一次），由**你**決定要不要貼給 AI
- **去敏模式**（`-Pseudonymize`）：可以按客戶分組，但名稱換成 `ID_xxxxxx` 代號；真名只印在黑窗
- 每次查詢記錄到 `audit/`（只記查詢文字與判定，不記回傳值）
- Claude Code 另有 `.claude/hooks/guard-data-access.ps1`，擋掉「直接連 msmdsrv 繞過服務」的寫法（第二道防線；主防線是伺服器）

⚠️ 仍然擋不住的：

- 篩選條件夠窄時，彙總本身就是明細（例如某一家客戶的營收總額）
- M 腳本裡寫死的客戶名稱與連線字串 —— `Get-PbiMQuery` 預設只存檔、不回傳內容
- 錯誤訊息可能夾帶真實資料值 —— 貼錯誤給 AI 前先看一眼
- **清單沒涵蓋到的欄位** —— 所以請務必依自己的模型調整清單

驗證防護是否生效：`.\tools\Test-DataGuard.ps1`（用你模型裡的實際欄位模擬各種繞過手法，只回報擋下與否，不印任何資料）。

---

## 目前的限制

| 項目 | 狀態 |
|---|---|
| DAX 量值／計算資料行／關聯／計算群組／RLS | ✅ 完整 |
| 階層（Hierarchy） | ❌ 未實作 |
| 檢視方塊 / KPI / 多語系 / 增量重新整理原則 | ❌ 未實作 |
| Power Query M | 🔒 **唯讀**。用 TOM 改 M 會讓 Desktop 卡在「查詢中有暫止的變更尚未套用」且無法自行解開，所以 API 不提供寫入（`/api/update-m` 回 410）。AI 會寫好完整 `let...in` 請你貼進進階編輯器 |
| M 預覽 / 查詢摺疊分析 | ❌ 請在 Power Query 編輯器內進行 |
| `/api/inject-visual` | ⚠️ 僅 PBIP，會強制關閉並重開 Power BI。**有防毒控管的公司電腦上不建議使用** |

---

## 專案結構

```
PBI_AI_Bridge/
├── README.md                      本文件
├── 使用說明.md                    逐步使用教學
├── CLAUDE.md                      AI 工作規範（Claude Code 自動載入）
├── AGENTS.md                      其他 AI 代理的入口
├── 📦第一次使用請先點我.bat         安裝（只跑一次）
├── 🚀啟動PBI終極儀表板.bat          每次使用的啟動器
├── PowerBI_Visualizer.html        網頁儀表板
├── API_Documentation.html         API 端點總覽
├── tools/
│   ├── PBI-Bridge.ps1             PowerShell 輔助函式
│   └── Test-DataGuard.ps1         資料保護測試
├── pbibridge_csharp/
│   ├── Program.cs                 所有 API 端點與 DataGuard
│   ├── pbibridge_csharp.csproj
│   ├── appsettings.template.json  設定範本（可分享）
│   └── appsettings.json           安裝時產生，含你的 API Key（勿分享）
├── .claude/
│   ├── settings.json              Claude Code hook 設定
│   └── hooks/guard-data-access.ps1
├── snapshots/                     模型快照（自動產生，每個模型一個子資料夾）
├── PowerQuery_Scripts/            M 腳本備份（自動產生）
└── audit/                         查詢稽核記錄（自動產生）
```

## 📤 分享給別人時

直接複製資料夾或壓 zip 前，**先刪掉以下項目**（`.gitignore` 已排除，但複製檔案不會）：

| 刪掉 | 原因 |
|---|---|
| `pbibridge_csharp/appsettings.json` | 你的 API Key |
| `pbibridge_csharp/bin/`、`obj/` | `bin\Release\net8.0\` 裡有一份含同一把 Key 的設定副本 |
| `snapshots/` 裡的內容 | 你操作過的模型的完整結構與全部 DAX |
| `PowerQuery_Scripts/` | M 腳本：連線字串、伺服器、資料庫名稱 |
| `audit/` 裡的內容 | 你查過哪些表與欄位 |

想分享你調好的保護清單？把清單內容手動合進 `appsettings.template.json`（**不要**連 Key 一起帶過去）。

⚠️ 用文字編輯器修改 `tools/*.ps1` 後，**存檔時要保留「UTF-8 with BOM」**，否則同事的 PowerShell 5.1 會把中文讀成亂碼而無法載入。

---

## 常見問題

| 症狀 | 處理 |
|---|---|
| 安裝檔說找不到 SDK | 裝的是 Runtime，請下載「.NET SDK 8.x.x」 |
| 編譯失敗，訊息有 `NU1301` / 無法載入服務索引 | 連不到 nuget.org。確認網路或公司 Proxy 後重跑安裝檔 |
| 編譯失敗 `MSB3027` 檔案鎖定 | 服務還開著。關掉黑窗再重跑 |
| `Port 5500 is already in use` | 已有一個服務在跑（忘了關的舊黑窗；VS Code 的 Live Server 預設也用 5500）。關掉它再重開 |
| 雙擊 .bat 跳「Windows 已保護您的電腦」 | 從網路下載的檔案。按「其他資訊 → 仍要執行」，或解壓縮前在 zip 按右鍵 → 內容 → 解除封鎖 |
| PowerShell 說「已停用指令碼執行」或「未經數位簽署」 | 見 [使用說明.md 第 11 節](使用說明.md#powershell-無法執行腳本) |
| 載入 `PBI-Bridge.ps1` 出現一堆 `Unexpected token` | 檔案被存成沒有 BOM 的 UTF-8，請改存為 UTF-8 with BOM |
| `找不到正在執行的 Power BI 檔案` | 先開啟 PBIX / PBIP |
| `目前有 N 個 Power BI 實例在執行` | `Use-PbiInstance <檔名片段或 Port>` 選定目標 |
| 401 Unauthorized | Key 不對。網頁端重新輸入 `appsettings.json` 裡的 `Security.ApiKey` |
| 改了 `appsettings.json` 沒生效 | 關掉黑窗、重新雙擊 🚀（服務讀的是 `bin\` 下的副本） |
| `Save-PbiModel` 回 `fileChanged = false` | 大檔可能還在寫，稍等再看；確定沒存到就手動 Ctrl+S，**不要連按** |
| Power BI 一直顯示「查詢中有暫止的變更尚未套用」 | M 被 API 改過而失步。到進階編輯器貼上正確 M 並「關閉並套用」 |
| 防毒跳警報 | 停止操作、截圖、記下時間，對照 AI 剛做了什麼。詳見 CLAUDE.md「防毒軟體相容規範」 |
