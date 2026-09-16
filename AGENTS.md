# PBI AI Bridge — AI 代理工作規範（跨代理通用入口）

這個工具連接「**AI 代理 ↔ 執行中的 Power BI Desktop**」，透過本地 HTTP 服務（localhost:5500）用 TOM 讀寫記憶體中的資料模型、用 ADOMD 執行 DAX 驗算。

## 適用哪些 AI 代理

任何能在 Windows 上執行 **PowerShell** 的 AI 編碼代理都可用：Claude Code、OpenAI Codex、Cursor、Google Antigravity、Gemini CLI、Windsurf、GitHub Copilot 等。

- **Claude Code** 會自動載入 `CLAUDE.md`（完整工作規範）。
- **其他代理**會自動讀取本檔 `AGENTS.md`（跨工具的公開標準）。**請務必把 `CLAUDE.md` 也完整讀過** —— 那是最完整的工作規範，內容與「你是哪個代理」無關，一體適用。
- **Google Antigravity**：AGENTS.md 需要 v1.20.3 以上；較舊的版本請把本檔複製一份到 `.agents/rules/`。它自己的規則檔是 `~/.gemini/GEMINI.md`，且規則檔有 12,000 字元上限 —— 所以 `CLAUDE.md` 不要整份貼進規則檔，用讀檔的方式讀它。

## 五條最關鍵的規則（不論你是哪個代理）

1. **資料保護由伺服器端強制**，不是靠代理自律：客戶身分欄位只能數不能取值、金額必須包在聚合函式內，違規回 **403**。這是 `pbibridge_csharp/Program.cs` 的 `DataGuard` 做的，**任何代理都無法自我豁免**。管制清單在 `appsettings.json → DataProtection`（預設是通用樣式，請依你自己的資料模型增補）。詳見 `CLAUDE.md` 的「資料保護規範」。

2. **一律走 `tools/PBI-Bridge.ps1`**，不要自己直連 `msmdsrv` 或直接載入 ADOMD —— 那會繞過保護，也會踩到中文編碼問題。
   （Claude Code 另有 `.claude/hooks/guard-data-access.ps1` 擋這條繞路；**其他代理沒有這個 hook**，更要自律。但伺服器端的 DataGuard 仍會攔查詢回傳值，所以真正的防線一直都在。）

3. **標準開發流程**：
   讀（`Get-PbiSchema`）→ 動手前留退路（`New-PbiSnapshot`）→ 寫入（`Set-PbiMeasure` 等）→ 驗算（`Invoke-Dax`）→ 存檔（`Save-PbiModel`，確認 `fileChanged = true`）。

4. **防毒相容（公司電腦）**：不要用 `Stop-Process` 終止行程、不要執行剛編譯出來的 `.exe`、不要遞迴掃描使用者資料夾或瀏覽器設定檔、不要用 `-EncodedCommand`。服務要重開就請使用者自己雙擊 `🚀啟動PBI終極儀表板.bat`。完整清單見 `CLAUDE.md` 的「防毒軟體相容規範」。

5. **Windows 環境的兩個坑**：
   - **執行原則／下載封鎖**：從網路下載的 `tools/*.ps1` 會被標記，一般 PowerShell 會拒絕載入（`is not digitally signed`）。請使用者依 `README.md` 的〈PowerShell 無法執行腳本〉解除封鎖；或以 `powershell -NoProfile -ExecutionPolicy Bypass -Command ". .\tools\PBI-Bridge.ps1; ..."` 呼叫。
   - **中文編碼**：本機是 Windows PowerShell 5.1，預設不是 UTF-8。務必用 `tools/PBI-Bridge.ps1`（已處理好），不要自己手刻 `Invoke-RestMethod`；輸出若是亂碼，在該次工作階段先設 `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`。

## 完整文件

| 檔案 | 內容 |
|---|---|
| **`CLAUDE.md`** | 完整工作規範（Claude Code 自動載入；其他代理請主動讀取）。這是最權威的一份。 |
| **`README.md`** | 給使用者（人）的完整說明：安裝、啟動、資料保護清單、與 AI 協作時使用者要做的事、疑難排解。 |
| **`API_Documentation.html`** | 端點總覽。服務開著時也可從 <http://localhost:5500/API_Documentation.html> 開啟。 |

> 這是通用型工具，**不預設綁定任何特定報表或資料模型**。第一次使用請依 `README.md` 的「快速開始」（雙擊 `🚀啟動PBI終極儀表板.bat` 即可），並在 `appsettings.json` 補上你自己資料中的敏感欄位。
