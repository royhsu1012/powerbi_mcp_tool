# Power Query（M）的開發互動模式

> 這是 [CLAUDE.md](../CLAUDE.md) 的延伸，用到才讀：要讀或改 Power Query 之前，先把這一份讀完。
> 資料保護、防毒相容與確認視窗的規矩一律以 CLAUDE.md 為準。

使用者會**自己開著 Power Query 編輯器**（那裡有預覽，是 AI 看不到的東西）。
分工原則：**使用者只做判斷，M 的碼一律由 AI 寫。**

| | 使用者 | Claude |
|---|---|---|
| 看 | 預覽長怎樣、哪一欄怪 | 列數、空值數、相異值數、總和有沒有跑掉 |
| 決定 | 這個轉換對不對 | — |
| 寫 | — | 全部的 M |

## M 可以寫入，但寫入不等於生效

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

## 循環

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

## 注意事項

- `Get-PbiMQuery` **一次只讀一張表**。M 腳本含連線字串、伺服器位址、資料庫名稱，
  整包拉會把所有連線資訊送進 context（見 CLAUDE.md 資料保護規範的「M 腳本：預設不輸出內容」）
- 使用者自己「關閉並套用」或按了套用提示時，Power BI 會自己重整，**不需要**再跑 `Invoke-PbiRefresh`；
  用 API 寫 M 而沒有出現提示時才需要（見上方表格）
- 資料源需要認證而 PBI 快取憑證失效時，TOM 觸發的 refresh **沒有 UI 可以輸入密碼**，可能失敗或卡住
- 用 TOM 建立全新的 M 表格時**必須明確指定 Columns**，引擎不會自動推導結構；
  既有的表多輸出欄位時同理，用 `Add-PbiColumn -SourceColumn` 補
