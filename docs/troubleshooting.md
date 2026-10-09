# 常見問題對照（給 AI 代理）

> 這是 [CLAUDE.md](../CLAUDE.md) 的延伸，用到才讀：遇到看不懂的錯誤、400、403、409 時查這一份。給使用者看的疑難排解在 README 第 13 節。
> 資料保護、防毒相容與確認視窗的規矩一律以 CLAUDE.md 為準。

| 症狀 | 原因 | 處理 |
|---|---|---|
| `找不到正在執行的 Power BI 檔案` | 沒開 PBI 或沒開檔案 | 請使用者開啟 PBIP/PBIX |
| 400 且訊息是亂碼 | PowerShell 5.1 編碼問題 | 改用 UTF-8 位元組送出 |
| 401 Unauthorized | Key 錯誤或沒帶 Header（手刻請求、或金鑰在這個 PowerShell 工作階段載入之後被換過） | 重新載入 `tools\PBI-Bridge.ps1` 再試。不要自己去讀金鑰檔 |
| `找不到 API 金鑰檔` | 金鑰是服務啟動時產生的：服務還沒啟動過，或黑窗裡跑的是舊版 | 請使用者關掉黑窗、重新雙擊 🚀 |
| 404 Not Found | 執行中的是舊版建置 | 重新 build 並重啟服務（流程見 `docs\service-development.md`，**由使用者關窗、由使用者開 `.bat`**） |
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
| 直接改了報表檔案，使用者接受變更後又變回舊的 | 接受之前有人送了 Ctrl+S（`Save-PbiModel`），舊版面被寫回磁碟 | 從備份的 `.Report` 還原再來一次；順序見 `docs\pbip-report-editing.md` |
| 防毒跳警報 | 做了行程終止／遞迴掃描／執行新編譯的 exe | 停下來告訴使用者你剛做了什麼、時間點，讓他對照警報。之後改走「請使用者代勞」的路線 |
| `目前有 N 個 Power BI 實例在執行` | 多個 PBI 開著但沒選目標 | `Use-PbiInstance <檔名片段或 Port>`。**不要為了繞過而隨便挑一個** |
| `先前選定的實例已不存在` | 選定的 PBI 被關掉了 | 重新 `Use-PbiInstance`。（只是換檔重開的話會自動重新解析，不會出現這個） |
| 改到錯的模型 | 沒確認 `Use-PbiInstance` 印出的完整路徑 | 同名檔案很常見，每次切換都要看路徑 |
| 改名後公式壞掉 | 沒有同步改寫引用 | 用 `Rename-PbiObject`（會自動改寫），**別直接改 TOM 的 Name** |
| 建計算群組失敗 | `DiscourageImplicitMeasures` 未開 | 先向使用者說明副作用，同意後加 `-DiscourageImplicitMeasures` |
