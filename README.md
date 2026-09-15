# ⚡ PBI AI Bridge

讓 AI 代理（Claude Code 等）直接讀寫**執行中的 Power BI Desktop** 資料模型的本地橋接服務。

- 透過 TOM 讀寫**記憶體中**的模型：量值、計算資料行、關聯、計算群組、RLS 角色……
- 透過 ADOMD 執行 DAX 查詢**當場驗算** —— 寫完立刻查，不必存檔、不必重開
- 可同時開多個 PBIX / PBIP，隨時切換操作對象
- 伺服器端強制**資料保護**：AI 讀得到結構，拿不到客戶名稱與逐筆金額
- **只有一個啟動檔**：缺什麼會先問你、按 Y 直接裝好；儀表板自動打開，不用貼金鑰

```
AI 代理（PowerShell）            瀏覽器（儀表板 http://localhost:5500/）
          ↕                            ↕
   HTTP + X-API-Key + X-PBI-Target   （只接受 localhost）
          ↕
   pbibridge_csharp/   C# 本地服務
          ↕  TOM（寫入）／ADOMD（查詢）
Power BI Desktop #1   Power BI Desktop #2   …
```

---

## 📖 文件導覽

| 你是… | 看這份 |
|---|---|
| **第一次拿到、要安裝使用的人** | **[使用說明.md](使用說明.md)** —— 從第一次啟動到日常操作的逐步教學 |
| 想快速了解全貌的人 | 本文件 |
| AI 代理（Claude Code 自動載入） | [CLAUDE.md](CLAUDE.md) —— 完整工作規範 |
| 其他 AI 代理（Codex / Cursor / Antigravity…） | [AGENTS.md](AGENTS.md)，再讀 CLAUDE.md |
| 查 API 端點 | [API_Documentation.html](API_Documentation.html) |

---

## 🚀 快速開始

1. **解壓縮到本機資料夾**，例如 `C:\PBI_AI_Bridge`（不要放 OneDrive，不要在 zip 裡直接執行）
2. 開啟一份 Power BI 檔案（.pbix 或 .pbip）
3. 雙擊 **`🚀啟動PBI終極儀表板.bat`**，照畫面回答 Y / N
4. 儀表板自動打開後，**保持黑窗開著**（關掉＝停止服務）
5. 在這個資料夾開啟 Claude Code，用中文告訴它你要做什麼

第一次視需要安裝的東西而定，約 2～5 分鐘；之後每次幾秒鐘。每一步的畫面見 [使用說明.md](使用說明.md)。

## 系統需求

| 項目 | 說明 |
|---|---|
| Windows 10 / 11 | 僅支援 Windows |
| Power BI Desktop | 要**開著至少一份檔案**才有東西可以連 |
| .NET SDK 8 以上 | **沒有也沒關係**：啟動檔會問你要不要用 Windows 內建的 `winget` 直接安裝（約 250 MB，可能要系統管理員權限）。只裝 Runtime 不夠，必須是 SDK |
| 網路（僅第一次） | 從 **nuget.org**（`api.nuget.org`）下載 6 個相依套件，共約 17 MB，之後離線可用。公司 Proxy／防火牆擋 nuget.org 時請找 IT |
| Windows PowerShell 5.1 | 系統內建 |
| AI 代理（選用） | Claude Code（建議），或其他能執行 PowerShell 的代理 |

<details>
<summary>第一次啟動會下載的套件（需要請 IT 放行時用）</summary>

| 套件 | 版本 | 大小 |
|---|---|---|
| Microsoft.AnalysisServices.NetCore.retail.amd64 | 19.82.0 | 6.2 MB |
| Microsoft.AnalysisServices.AdomdClient.NetCore.retail.amd64 | 19.82.0 | 1.9 MB |
| System.Management | 8.0.0 | 0.8 MB |
| Microsoft.Identity.Client | 4.56.0 | 7.7 MB |
| System.CodeDom | 8.0.0 | 0.5 MB |
| Microsoft.IdentityModel.Abstractions | 6.22.0 | 0.1 MB |

前三個是專案直接引用，後三個是自動帶入的相依套件。來源：`https://api.nuget.org/v3/index.json`。
套件會存進使用者的 NuGet 快取（`%USERPROFILE%\.nuget\packages`），同一台機器之後不再下載。
只裝了 .NET 9 / 10 SDK 的電腦，還會多下載 .NET 8 的參考套件。

</details>

---

## 啟動檔做了什麼

只有一個啟動檔：`🚀啟動PBI終極儀表板.bat`。每次雙擊都依序檢查，**只做缺少的部分**：

| 檢查 | 第一次 | 之後 |
|---|---|---|
| 服務已經在跑？ | — | 是的話直接打開儀表板，不會重複啟動 |
| 1/5 檔案 | 確認資料夾完整；在 OneDrive 裡會警告 | 略過 |
| 2/5 .NET SDK | 沒有的話**問你要不要安裝**，按 Y 用 `winget` 直接裝；沒有 winget 時給連結與步驟，裝好按 R 重新檢查 | 略過 |
| 3/5 設定檔 | 產生 `appsettings.json` 與**只屬於這台機器的隨機 API Key**，並提醒設定資料保護清單（可直接開記事本） | 略過 |
| 4/5 套件 | **問你要不要下載**（約 17 MB）。失敗時自動判斷：沒註冊 nuget.org 就問要不要幫你加；網路問題就給檢查步驟，按 R 重試 | 略過 |
| 5/5 編譯 | 編譯（約 30 秒） | 程式有改才重新編譯 |
| 啟動 | 服務起來後**自動打開儀表板** | 同左 |

全程不會碰你的 Power BI 檔案。金鑰由儀表板與 AI 工具自動讀取，**不需要手動輸入**。

> **每台機器各自產生 Key，不要共用，也不要把 `appsettings.json` 傳給別人。**

## ⚙️ 第一次啟動時：設定資料保護清單

`pbibridge_csharp/appsettings.json` → `DataProtection` 預設只有英文通用樣式（`*customer*`、`*amount*`…）。
**模型若用中文欄位名，預設清單幾乎擋不到東西**，請補上你自己的敏感欄位。第一次啟動時，啟動檔會提醒並幫你用記事本打開。

| 清單 | 放什麼 | 效果 |
|---|---|---|
| `DenyColumns` | 客戶名、聯絡人、電話、備註等身分／自由文字欄位 | 只能計數，不能取值 |
| `AggregateOnlyColumns` | 金額、單價、成本、薪資 | 必須包在 SUM／AVERAGE 等聚合內 |
| `AllowColumns` | 被樣式誤傷、其實不敏感的欄位 | 優先放行 |

比對時忽略大小寫、空白、底線、連字號，可用 `*` 樣式（例如 `*客戶*`）。
之後再改：**關掉黑窗、重新雙擊 🚀** 就生效（啟動檔每次都會把設定複製到服務讀取的 `bin\`）。

範例與驗證方法見 [使用說明.md 第 4 節](使用說明.md#4-設定資料保護清單重要)。

---

## 兩種使用方式

### A. 網頁儀表板（唯讀瀏覽）

服務啟動後會自動打開 <http://localhost:5500/>：瀏覽資料表、欄位、DAX 量值、M 腳本，右上角可切換要看哪一個 Power BI。
頁面由服務提供並自動帶入金鑰，**不需要輸入任何東西**。直接雙擊 `PowerBI_Visualizer.html` 也會自動轉到這個網址。

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

服務本身的連線保護：

- 所有 `/api/*` 都要帶金鑰。儀表板頁面由服務提供並帶入金鑰，所以使用者不必輸入
- 只接受以 `localhost` / `127.0.0.1` 連線（擋 DNS rebinding），也不允許 `null` 來源（擋沙箱 iframe）—— 其他網站拿不到金鑰，也呼叫不了 API

⚠️ 仍然擋不住的：

- 篩選條件夠窄時，彙總本身就是明細（例如某一家客戶的營收總額）
- M 腳本裡寫死的客戶名稱與連線字串 —— `Get-PbiMQuery` 預設只存檔、不回傳內容
- 錯誤訊息可能夾帶真實資料值 —— 貼錯誤給 AI 前先看一眼
- **清單沒涵蓋到的欄位** —— 所以請務必依自己的模型調整清單

驗證防護是否生效：`.\tools\Test-DataGuard.ps1`（用你模型裡的實際欄位模擬各種繞過手法，只回報擋下與否，不印任何資料）。同時開多個 Power BI 時加 `-Target <Port>`。

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
├── 🚀啟動PBI終極儀表板.bat          唯一的啟動檔（自動偵測：安裝／編譯／啟動）
├── PowerBI_Visualizer.html        網頁儀表板（由服務在 localhost:5500 提供）
├── API_Documentation.html         API 端點總覽
├── tools/
│   ├── PBI-Bridge.ps1             PowerShell 輔助函式
│   └── Test-DataGuard.ps1         資料保護測試
├── pbibridge_csharp/
│   ├── Program.cs                 所有 API 端點與 DataGuard
│   ├── pbibridge_csharp.csproj
│   ├── appsettings.template.json  設定範本（可分享）
│   └── appsettings.json           第一次啟動時產生，含你的 API Key（勿分享）
├── .claude/
│   ├── settings.json              Claude Code hook 設定
│   └── hooks/guard-data-access.ps1
├── snapshots/                     模型快照（自動產生，每個模型一個子資料夾）
├── PowerQuery_Scripts/            M 腳本備份（自動產生）
└── audit/                         查詢稽核記錄（自動產生）
```

## 📤 分享給別人時

**最簡單**：請對方從 GitHub 下載（Code → Download ZIP），裡面不會有任何你的個人檔案。

直接複製你自己的資料夾或壓 zip 的話，**先刪掉以下項目**（`.gitignore` 已排除，但複製檔案不會）：

| 刪掉 | 原因 |
|---|---|
| `pbibridge_csharp/appsettings.json` | 你的 API Key |
| `pbibridge_csharp/bin/`、`obj/` | `bin\Release\net8.0\` 裡有一份含同一把 Key 的設定副本 |
| `snapshots/` 裡的內容 | 你操作過的模型的完整結構與全部 DAX |
| `PowerQuery_Scripts/` | M 腳本：連線字串、伺服器、資料庫名稱 |
| `audit/` 裡的內容 | 你查過哪些表與欄位 |

想分享你調好的保護清單？把清單內容手動合進 `appsettings.template.json`（**不要**連 Key 一起帶過去）。

⚠️ 修改檔案時的兩個編碼陷阱：

- `tools/*.ps1` 存檔要保留「**UTF-8 with BOM**」，否則 PowerShell 5.1 會把中文讀成亂碼而無法載入
- `🚀啟動PBI終極儀表板.bat` 只能有**英文字元**、行尾必須是 **CRLF**，否則 cmd 會讀錯位置、跳到錯的標籤

---

## 常見問題

| 症狀 | 處理 |
|---|---|
| 雙擊 .bat 跳「Windows 已保護您的電腦」 | 從網路下載的檔案。按「其他資訊 → 仍要執行」，或解壓縮前在 zip 按右鍵 → 內容 → 解除封鎖 |
| 啟動檔說找不到 .NET SDK | 按 Y 讓它用 winget 安裝。公司電腦沒有 winget 或被擋時，按 O 開下載頁手動安裝「.NET SDK 8.x.x」，裝好回到黑窗按 R |
| 啟動檔 4/5 下載失敗 | 照畫面處理：沒註冊 nuget.org 會問你要不要自動加上；網路問題請確認瀏覽器打得開 `https://api.nuget.org/v3/index.json`、公司 Proxy，再按 R 重試。詳見 [使用說明.md](使用說明.md#套件下載失敗) |
| 編譯失敗 `MSB3027` 或 `CS2012`，訊息有 `being used by another process` | 同時有兩個啟動檔在編譯。關掉多的黑窗再雙擊一次 |
| `Port 5500 is used by a different program` | 別的程式占用 5500（例如 VS Code Live Server）。關掉它再雙擊 🚀。本服務已經在跑的話不會出現這個，而是直接打開儀表板 |
| PowerShell 說「已停用指令碼執行」或「未經數位簽署」 | 見 [使用說明.md 第 11 節](使用說明.md#powershell-無法執行腳本) |
| 載入 `PBI-Bridge.ps1` 出現一堆 `Unexpected token` | 檔案被存成沒有 BOM 的 UTF-8，請改存為 UTF-8 with BOM |
| `找不到正在執行的 Power BI 檔案` | 先開啟 PBIX / PBIP |
| `目前有 N 個 Power BI 實例在執行` | `Use-PbiInstance <檔名片段或 Port>` 選定目標。同一份檔案開了兩次也會這樣 |
| 儀表板顯示「金鑰不符」 | 服務重新啟動過，按 F5 重新整理頁面 |
| 改了 `appsettings.json` 沒生效 | 關掉黑窗、重新雙擊 🚀 |
| `Save-PbiModel` 回 `fileChanged = false` | 大檔可能還在寫，稍等再看；確定沒存到就手動 Ctrl+S，**不要連按** |
| Power BI 一直顯示「查詢中有暫止的變更尚未套用」 | M 被 API 改過而失步。到進階編輯器貼上正確 M 並「關閉並套用」 |
| 防毒跳警報 | 停止操作、截圖、記下時間，對照 AI 剛做了什麼。詳見 CLAUDE.md「防毒軟體相容規範」 |
