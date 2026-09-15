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

> **其他文件**：[CLAUDE.md](CLAUDE.md) 是給 AI 的工作規範（Claude Code 自動載入）；[AGENTS.md](AGENTS.md) 是 Codex / Cursor 等其他代理的入口；API 端點說明在 [API_Documentation.html](API_Documentation.html)（服務開著時從儀表板右上角開啟）。

**目錄**

1. [快速開始](#1-快速開始)
2. [系統需求](#2-系統需求)
3. [第一次啟動](#3-第一次啟動)
4. [設定資料保護清單（重要）](#4-設定資料保護清單重要)
5. [每天的使用流程](#5-每天的使用流程)
6. [跟 AI 協作](#6-跟-ai-協作)
7. [網頁儀表板](#7-網頁儀表板)
8. [資料保護怎麼運作](#8-資料保護怎麼運作)
9. [快照與還原](#9-快照與還原)
10. [自己用 PowerShell 操作](#10-自己用-powershell-操作)
11. [目前的限制](#11-目前的限制)
12. [專案結構與分享](#12-專案結構與分享)
13. [疑難排解](#13-疑難排解)

---

## 1. 快速開始

1. **解壓縮到本機資料夾**，例如 `C:\PBI_AI_Bridge`（不要放 OneDrive，不要在 zip 裡直接執行）
2. 開啟一份 Power BI 檔案（.pbix 或 .pbip）
3. 雙擊 **`🚀啟動PBI終極儀表板.bat`**，照畫面回答 Y / N
4. 儀表板自動打開後，**保持黑窗開著**（關掉＝停止服務）
5. 在這個資料夾開啟 Claude Code，用中文告訴它你要做什麼

第一次視需要安裝的東西而定，約 2～5 分鐘；之後每次幾秒鐘。

> 💡 從 Email / Teams 下載的 zip：**解壓縮前**在 zip 上按右鍵 → 內容 → 勾選「解除封鎖」，可以避免之後 Windows 與 PowerShell 擋下檔案。

---

## 2. 系統需求

| 項目 | 說明 |
|---|---|
| Windows 10 / 11 | 僅支援 Windows |
| Power BI Desktop | 要**開著至少一份檔案**才有東西可以連 |
| .NET SDK 8 以上 | **沒有也沒關係**：啟動檔會問你要不要用 Windows 內建的 `winget` 直接安裝（約 250 MB，可能要系統管理員權限）。只裝 Runtime 不夠，必須是 SDK |
| 網路（僅第一次） | 從 **nuget.org** 下載 6 個相依套件，共約 17 MB，之後離線可用。瀏覽器打得開 <https://api.nuget.org/v3/index.json> 就沒問題 |
| Windows PowerShell 5.1 | 系統內建 |
| AI 代理（選用） | Claude Code（建議）—— Claude 桌面版的 **Code** 分頁，或終端機的 `claude` 指令 |

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

## 3. 第一次啟動

只有一個啟動檔：`🚀啟動PBI終極儀表板.bat`。每次雙擊都依序檢查，**只做缺少的部分**：

| 檢查 | 第一次 | 之後 |
|---|---|---|
| 服務已經在跑？ | — | 是的話直接打開儀表板，不會重複啟動 |
| 1/5 檔案 | 確認資料夾完整；在 OneDrive 裡會警告 | 略過 |
| 2/5 .NET SDK | 沒有的話**問你要不要安裝** | 略過 |
| 3/5 設定檔 | 產生 `appsettings.json` 與**只屬於這台機器的隨機 API Key**，並提醒設定資料保護清單（可直接開記事本） | 略過 |
| 4/5 套件 | **問你要不要下載**（約 17 MB） | 略過 |
| 5/5 編譯 | 編譯（約 30 秒） | 程式有改才重新編譯 |
| 啟動 | 服務起來後**自動打開儀表板** | 同左 |

第一次大概會看到：

```
[1/5] Files
      OK
[2/5] .NET SDK
      OK                                    ← 沒有的話會問你要不要安裝（見下方）
[3/5] Settings file
      Created, with an API Key made just for this computer.
      IMPORTANT - protect your data before AI queries it
      Open that file in Notepad now? [Y/N]  ← 按 Y，照第 4 節設定
[4/5] Libraries
      First run: 6 libraries must be downloaded from nuget.org, about 17 MB.
      Download them now? [Y/N]              ← 按 Y
      OK
[5/5] Compile
      Compiling, about 30 seconds...
      OK
----------------------------------------------
  Starting - the dashboard opens by itself.
  KEEP THIS WINDOW OPEN. Closing it stops the server.
----------------------------------------------
[保護] 資料保護啟用 — 身分欄位／樣式 23 條、金額欄位／樣式 11 條、安全白名單 1 條
[實例] 偵測到 1 個 Power BI 實例：
     · Port 51234  銷售報表.pbix  [PBIX]
[網頁] 網頁儀表板：http://localhost:5500/
```

瀏覽器會自動打開儀表板，看到你的資料表就完成了。**不需要抄金鑰、也不需要貼金鑰**，儀表板與 AI 工具會自己讀。
全程不會碰你的 Power BI 檔案。

> **每台機器各自產生 Key，不要共用，也不要把 `appsettings.json` 傳給別人。**

### 如果問你要不要安裝 .NET SDK

- 按 **Y**：用 `winget` 從微軟直接下載安裝。Windows 可能跳出「是否允許此 App 變更你的裝置」，按「是」。裝好會自動繼續
- 按 **N**，或安裝失敗（例如公司電腦停用了 winget）：畫面會列出手動步驟
  - 按 **O** 開啟下載頁 → 找「SDK 8.0.x」→ Windows **x64** Installer → 安裝（注意是 **SDK**，不是 Runtime）
  - 裝好回到黑窗按 **R** 重新檢查，不用重開；按 **Q** 離開

### 套件下載失敗

啟動檔會自己判斷原因：

- **這台電腦沒註冊 nuget.org**：會問要不要自動加上並重試，按 Y
- **網路連不到**：處理好之後按 **R** 重試
  1. 用瀏覽器開 <https://api.nuget.org/v3/index.json>，看得到一段文字就代表連得到
  2. 瀏覽器打得開、黑窗卻還是失敗，通常是公司 Proxy —— 請 IT 協助，或提供公司內部的 NuGet 來源
  3. 有開 VPN 的話，切換一下再試

---

## 4. 設定資料保護清單（重要）

預設清單只認得英文通用欄位名（`customer`、`amount`…）。
**模型如果用中文欄位名，預設清單幾乎擋不到任何東西。** 請在讓 AI 查詢資料之前完成這一步。
第一次啟動時按 Y 會直接用記事本打開；之後要改就自己開 `pbibridge_csharp\appsettings.json`。

在 `DataProtection` 段落的清單**後面加上**你模型的欄位（原本的英文樣式保留）。下面是示意，`…（原本的保留）…` 代表原有內容，**不要照打進去**：

```json
"DenyColumns": [
  "*customer*",
  …（原本的保留）…
  "*客戶*",
  "*經銷商*",
  "*聯絡人*",
  "*業務員*",
  "*電話*",
  "*備註*"
],
"AggregateOnlyColumns": [
  "*amount*",
  …（原本的保留）…
  "*金額*",
  "*單價*",
  "*成本*"
],
"AllowColumns": [
  "客戶等級"
],
```

| 清單 | 放什麼 | AI 可以 | AI 不行 |
|---|---|---|---|
| `DenyColumns` | 客戶名、公司名、人名、電話、Email、地址、備註等自由文字 | 數有幾個 | 看到任何一個值 |
| `AggregateOnlyColumns` | 金額、單價、成本、薪資 | 加總、平均、最大最小 | 逐筆列出 |
| `AllowColumns` | 被上面樣式誤傷、其實不敏感的欄位 | 一般使用 | — |

**寫法規則**

- `*` 代表任意文字：`*客戶*` 會比對到「客戶名稱」「終端客戶」「客戶代號」
- 比對時**忽略大小寫、空白、底線、連字號**：`account name`、`Account_Name`、`account-name` 視為同一個
- 誤傷時（例如 `*客戶*` 打到「客戶等級」），把**那個欄位名**加進 `AllowColumns`，不要刪掉整條樣式
- JSON 格式要正確：每一項用雙引號、項目之間用逗號、**每個清單最後一項後面不能有逗號**

> 不確定有哪些欄位？打開儀表板就看得到所有欄位名稱。也可以請 AI：「列出所有欄位名稱，挑出可能是客戶身分或金額的欄位，給我建議的清單」—— 這只會讀欄位名稱，不會讀資料。

**讓設定生效**：第一次啟動時，改完存檔、關掉記事本，回黑窗按任意鍵繼續。之後再改：關掉黑窗 → 重新雙擊 🚀。
確認黑窗「資料保護啟用」那一行的條數跟你清單的項目數一致。黑窗出現錯誤並顯示 `[EXIT] Server stopped` 的話，多半是 JSON 格式打錯。

**驗證防護真的有效**：服務與 Power BI 都開著時，請 AI 執行，或自己在工具資料夾的 PowerShell 跑：

```powershell
.\tools\Test-DataGuard.ps1                 # 只開一個 Power BI 時
.\tools\Test-DataGuard.ps1 -Target 61853   # 開了多個時指定 Port（黑窗「偵測到 N 個實例」那段看得到）
```

它拿你模型裡**實際存在**的管制欄位模擬各種「偷看」手法，**只回報有沒有被擋下，不會印出任何資料**。最後一行是 `N 通過 / 0 失敗` 就代表有效。

---

## 5. 每天的使用流程

```
① 開啟 Power BI 檔案
② 雙擊 🚀啟動PBI終極儀表板.bat（幾秒鐘，儀表板自動打開；黑窗保持開著）
③ 在工具資料夾開啟 Claude Code，開始對話
④ 做完 → 確認已存檔 → 關掉黑窗
```

- ①② 順序可以對調，服務啟動後才開的 Power BI 也偵測得到
- **黑窗關掉 = 服務停止**，AI 就連不到了
- 找不到儀表板分頁了？**再雙擊一次 🚀**，它發現服務已經在跑，會直接打開儀表板
- 可以同時開好幾個 Power BI 檔案，AI 會問你要改哪一個

---

## 6. 跟 AI 協作

**開啟方式**：Claude 桌面版的 Code 分頁 → 選擇資料夾 → 選這個工具的資料夾；或在終端機 `cd C:\PBI_AI_Bridge` 後輸入 `claude`。
**一定要在這個資料夾開啟**，AI 才會自動讀到 `CLAUDE.md` 的工作規範（資料保護、防毒相容、驗算習慣）。第一次開啟時如果詢問是否信任這個資料夾，選信任。

第一句話可以說：「幫我檢查橋接服務，然後告訴我目前這份模型的結構」。

| 你想做 | 可以這樣說 |
|---|---|
| 了解模型 | 「這份模型有哪些事實表跟維度表？關聯怎麼接？」 |
| 健檢 | 「跑一次模型健檢，列出壞掉的公式和可能沒用到的欄位」 |
| 新增量值 | 「新增『毛利率』= [毛利] / [營收]，放在『量值』表，百分比格式，做完用產品線驗算」 |
| 修公式 | 「『年增率』在一月份是空白，幫我找原因並修好」 |
| 改名 | 「把『銷售額』改成『銷售總額』，先給我看會改到哪些公式」 |
| Power Query | 「『訂單』表的日期欄是 YYYYMMDD 文字，幫我寫轉日期的 M」 |
| 效能 | 「哪幾張表最佔記憶體？」「這個量值很慢，有沒有更快的寫法？」 |
| 多個檔案 | 「我開了兩個 Power BI，這次改『銷售報表』那個」 |

**好習慣**：確認 AI 選定時印出的完整路徑是你要改的那份（同名檔案常有好幾份）；大改之前說「先存快照」；要求「做完驗算給我看」；結束前確認已存檔。

### AI 會請你幫忙的五件事

有些事 AI **刻意不自己做**（為了資料保護，或避免觸發公司防毒），會停下來請你動手。這是正常的。

| 情況 | 你要做的 |
|---|---|
| **① 一次性權杖**：查詢被資料保護擋下，AI 說明要看哪些欄位、幾列、為什麼 | 不同意就說「不行，用彙總就好」。同意的話，到黑窗找「若你確認要放行『這一句』查詢，把下列權杖貼給 AI」，把權杖貼給 AI。權杖只對那一句有效、10 分鐘過期、用一次就失效 |
| **② 貼 M 腳本**：AI 不能直接改 M（會讓 Power BI 卡在「查詢中有暫止的變更尚未套用」） | Power Query 編輯器 → 選那張查詢 → **進階編輯器** → 全選 → 貼上 AI 給的完整 `let … in` → **關閉並套用**。回覆 `OK`、`錯`＋錯誤訊息，或 `怪`＋截圖（貼之前看一眼有沒有真實資料） |
| **③ 手動存檔**：AI 回報 `fileChanged = false` | 切到 Power BI 按 Ctrl+S |
| **④ 重開服務**：AI 改了伺服器程式或設定 | 關掉黑窗 → 告訴它 → 重新雙擊 🚀（程式有改會自動重新編譯）→ 告訴它 |
| **⑤ 防毒警報** | 先截圖並告訴 AI，它會說明剛才做了什麼。AI 被規範不能自己讀防毒紀錄 |

---

## 7. 網頁儀表板

服務啟動後自動打開 <http://localhost:5500/>。分頁不見了就再雙擊一次 🚀，或直接在瀏覽器輸入這個網址。

- 瀏覽資料表、欄位、DAX 量值與 M 腳本；搜尋會標出符合的欄位
- 右上角切換 Power BI 檔案、深色／淺色，以及開啟 API 文件
- 滑過欄位或量值可以複製 DAX 參照
- **不需要輸入金鑰**：頁面由服務提供並自動帶入。直接雙擊 `PowerBI_Visualizer.html` 也會轉到上面的網址

儀表板**只能瀏覽，不會修改模型**。

---

## 8. 資料保護怎麼運作

**AI 是雲端模型 —— 任何進入對話的內容都會傳到雲端。** 橋接服務只綁 localhost、不對外連線，唯一的外流管道是「AI 讀到了什麼」。
分界線：**欄位名稱可讀，欄位內容受管**。schema／關聯／健檢全開；被管制的只有 `/api/query`、`/api/dmv` 的回傳值。

由伺服器端強制執行（`Program.cs` 的 `DataGuard`），違規回 403，AI 無法自我豁免：

- **身分欄位**只能出現在計數函式內（`DISTINCTCOUNT` / `COUNTROWS` …）
- **金額欄位**必須聚合，且分組列數有上限
- **回傳列數封頂**：逐列明細 50 列、彙總 300 列（呼叫端只能調低）
- **一次性權杖**：真的需要看明細時由**你**決定要不要放行（見第 6 節 ①）
- **去敏模式**：可以按客戶分組，但名稱換成 `ID_A3F1B2` 這種代號。代號穩定、推不回原值；真名只印在黑窗
- **稽核記錄**：`audit\query-audit-YYYYMM.tsv` 記錄每一句查詢與判定（不含回傳值）

其他防線：

- **服務連線**：所有 `/api/*` 都要帶金鑰；只接受以 `localhost` 連線（擋 DNS rebinding），不允許 `null` 來源（擋沙箱 iframe）—— 其他網站拿不到金鑰，也呼叫不了 API
- **Claude Code hook**：`.claude/hooks/guard-data-access.ps1` 擋掉「直接連 msmdsrv 繞過服務」的寫法

| 情況 | 結果 |
|---|---|
| AI 查「有幾個客戶」、「各產品線營收」 | ✅ 回傳數字／彙總 |
| AI 想列出客戶名稱、逐筆看金額 | ⛔ 擋下，AI 會改用計數、加總或去敏模式 |
| 查詢結果超過 50 列明細／300 列彙總 | 只回傳到上限，AI 會知道結果被截斷 |

⚠️ **仍然擋不住、你自己要注意的**：

- 條件很窄的彙總等於明細：「某某公司上個月營收」雖然是加總，其實就是那家公司的數字
- M 腳本裡寫死的客戶名稱與連線字串 —— AI 預設只存檔、不讀內容
- **錯誤訊息與截圖可能含真實資料**，貼給 AI 前先看一眼
- **清單沒列到的欄位不受保護**：模型新增了敏感欄位，記得回來補清單

---

## 9. 快照與還原

AI 在刪除、改名等破壞性操作前會自動存快照。你也可以隨時要求：「先存一個快照，標籤叫『改版前』」、「還原到『改版前』，只還原量值，先給我看會改什麼」。

快照存在 `snapshots\<檔名>_<雜湊>\`，每份模型各自獨立，不會還原到別的模型上。還原**不是**完整的時光機：

- 只還原量值、計算資料行、關聯、共用運算式、M 腳本
- 快照之後**新建的表格、角色不會被刪掉**
- 快照裡沒有的量值與關聯**會被刪掉**（回到當時狀態，不是合併）
- 還原 M 會造成「暫止的變更」問題，所以通常請 AI「只還原量值」

快照只是模型定義的備份，**不能取代存檔**。重要的報表請另外自己備份 .pbix / .pbip。

---

## 10. 自己用 PowerShell 操作

不透過 AI 也可以直接用。在工具資料夾開 PowerShell：

```powershell
. .\tools\PBI-Bridge.ps1                        # 載入（會列出所有可用指令）
Test-PbiBridge                                  # 服務狀態 + 有哪些 PBI + 目前目標
Use-PbiInstance 銷售報表                          # 多個 PBI 時先選定（檔名片段或 Port）
New-PbiSnapshot -Label before                   # 動手前留退路
Set-PbiMeasure -Table 量值 -Name 銷售總額 -Expression 'SUM(FactSales[Amount])' -Format '#,0'
(Invoke-Dax 'EVALUATE ROW("結果", [銷售總額])').Rows   # 立刻驗算
Save-PbiModel                                   # 確認回傳 fileChanged = true
```

**不要自己手刻 `Invoke-RestMethod`**：PowerShell 5.1 預設不以 UTF-8 送出，中文欄位名與 DAX 會壞掉；`PBI-Bridge.ps1` 已處理好。

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

各函式的用途與範例寫在 `tools/PBI-Bridge.ps1` 的註解裡；參數語法用 `Get-Command Set-PbiMeasure -Syntax` 查。

**修改只進記憶體**：畫面立刻更新，但要寫進 `.pbix` / `.pbip` 必須 `Save-PbiModel`（或手動 Ctrl+S）。結構性變更後要 refresh 才查得到：

| 做了什麼 | 需要 refresh 嗎 |
|---|---|
| 新增／修改量值 | 不用 |
| 建計算表、新增計算項目 | `Invoke-PbiRefresh -Table <表>` |
| 建／改關聯、新增計算資料行 | `Invoke-PbiRefresh -RefreshType calculate` |
| 改 M 腳本 | 不用 —— 使用者「關閉並套用」時 Power BI 會自己重整 |

---

## 11. 目前的限制

| 項目 | 狀態 |
|---|---|
| DAX 量值／計算資料行／關聯／計算群組／RLS | ✅ 完整 |
| 階層（Hierarchy） | ❌ 未實作 |
| 檢視方塊 / KPI / 多語系 / 增量重新整理原則 | ❌ 未實作 |
| Power Query M | 🔒 **唯讀**。用 TOM 改 M 會讓 Desktop 卡在「查詢中有暫止的變更尚未套用」且無法自行解開，所以 API 不提供寫入。AI 會寫好完整 `let...in` 請你貼進進階編輯器 |
| M 預覽 / 查詢摺疊分析 | ❌ 請在 Power Query 編輯器內進行 |
| `/api/inject-visual` | ⚠️ 僅 PBIP，會強制關閉並重開 Power BI。**有防毒控管的公司電腦上不建議使用** |

---

## 12. 專案結構與分享

```
PBI_AI_Bridge/
├── README.md                      本文件
├── CLAUDE.md                      AI 工作規範（Claude Code 自動載入）
├── AGENTS.md                      其他 AI 代理的入口
├── 🚀啟動PBI終極儀表板.bat          唯一的啟動檔（自動偵測：安裝／編譯／啟動）
├── PowerBI_Visualizer.html        網頁儀表板（由服務在 localhost:5500 提供）
├── API_Documentation.html         API 端點說明
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
├── snapshots/                     模型快照（自動產生）
├── PowerQuery_Scripts/            M 腳本備份（自動產生）
└── audit/                         查詢稽核記錄（自動產生）
```

**分享給別人**：最簡單是請對方從 GitHub 下載（Code → Download ZIP），裡面不會有任何你的個人檔案。
直接複製自己的資料夾或壓 zip 的話，**先刪掉**以下項目（`.gitignore` 已排除，但複製檔案不會）：

| 刪掉 | 原因 |
|---|---|
| `pbibridge_csharp/appsettings.json` | 你的 API Key |
| `pbibridge_csharp/bin/`、`obj/` | `bin\Release\net8.0\` 裡有一份含同一把 Key 的設定副本 |
| `snapshots/`、`PowerQuery_Scripts/`、`audit/` 的內容 | 你操作過的模型結構、DAX、M 腳本（連線字串）、查詢紀錄 |

想分享調好的保護清單？把清單內容手動合進 `appsettings.template.json`（**不要**連 Key 一起帶過去）。

⚠️ 修改檔案時的編碼陷阱：

- `tools/*.ps1` 與 `.claude/hooks/*.ps1` 存檔要保留「**UTF-8 with BOM**」，否則 PowerShell 5.1 會把中文讀成亂碼而無法載入
- hook 腳本讀取輸入時**必須用 UTF-8**；用系統碼頁讀，中文路徑會讓它解析失敗而全部放行
- `🚀啟動PBI終極儀表板.bat` 只能有**英文字元**、行尾必須是 **CRLF**，否則 cmd 會讀錯位置、跳到錯的標籤

---

## 13. 疑難排解

| 症狀 | 處理 |
|---|---|
| 雙擊 .bat 跳「Windows 已保護您的電腦」 | 從網路下載的檔案。按「其他資訊」→「仍要執行」，或解壓縮前在 zip 按右鍵 → 內容 → 解除封鎖 |
| 啟動檔找不到 .NET SDK | 按 Y 讓它用 winget 安裝；不行就按 O 手動安裝「.NET SDK 8.x.x」，裝好按 R。見[第 3 節](#如果問你要不要安裝-net-sdk) |
| 啟動檔 4/5 下載失敗 | 見[套件下載失敗](#套件下載失敗) |
| `[WARNING] This folder is inside OneDrive` | 建議把資料夾搬到本機，例如 `C:\PBI_AI_Bridge` |
| 編譯失敗，訊息有 `being used by another process`（`MSB3027` / `CS2012`） | 同時有兩個啟動檔在編譯。關掉多的黑窗再雙擊一次 |
| `Port 5500 is used by a different program` | 別的程式占用 5500（例如 VS Code Live Server）。關掉它再雙擊 🚀 |
| 黑窗顯示 `[EXIT] Server stopped` | 往上捲看錯誤。改過 `appsettings.json` 的話先檢查 JSON 格式（逗號、引號） |
| 儀表板顯示「無法讀取模型」 | 確認黑窗開著、Power BI 有開檔案；開了多個就在右上角選一個 |
| 儀表板顯示「金鑰不符」 | 服務重新啟動過，按 F5 重新整理頁面 |
| `找不到正在執行的 Power BI 檔案` | 先開啟 .pbix / .pbip |
| `目前有 N 個 Power BI 實例在執行` | 告訴 AI 要改哪一個（或 `Use-PbiInstance <檔名片段或 Port>`）。同一份檔案開了兩次也會這樣，關掉多的那個視窗 |
| 改了 `appsettings.json` 沒生效 | 關掉黑窗、重新雙擊 🚀 |
| `fileChanged = false` | 大檔可能還在寫，稍等再看；確定沒存到就手動 Ctrl+S，**不要連按** |
| Power BI 一直顯示「查詢中有暫止的變更尚未套用」 | M 被 API 改過而失步。到進階編輯器貼上正確 M 並「關閉並套用」 |
| 載入工具時出現一堆 `Unexpected token` 亂碼 | `.ps1` 被存成沒有 BOM 的檔案。VS Code 開啟 → 右下角編碼 →「以 UTF-8 with BOM 儲存」 |
| 防毒警報 | 停止操作、截圖、記下時間，告訴 AI 與 IT。詳見 CLAUDE.md「防毒軟體相容規範」 |

### PowerShell 無法執行腳本

錯誤訊息類似「因為這個系統上已停用指令碼執行」或「檔案未經數位簽署」。依序嘗試：

1. **解除下載封鎖**（最常見）—— 在工具資料夾開 PowerShell 執行：
   ```powershell
   Get-ChildItem .\tools, .\.claude\hooks -Filter *.ps1 | Unblock-File
   ```
2. **只對目前使用者放寬執行原則**：
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   ```
3. 仍被公司原則（群組原則）鎖住 → 請洽 IT，不要自行想辦法繞過
