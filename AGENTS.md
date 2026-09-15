# PBI AI Bridge — AI 代理工作規範（跨代理通用入口）

這個工具連接「**AI 代理 ↔ 執行中的 Power BI Desktop**」，透過本地 HTTP 服務（localhost:5500）用 TOM 讀寫記憶體中的資料模型、用 ADOMD 執行 DAX 驗算。

## 適用哪些 AI 代理

任何能執行 **PowerShell** 的 AI 編碼代理都可用：Claude Code、OpenAI Codex、Cursor、Google Antigravity 等。

- **Claude Code** 會自動載入 `CLAUDE.md`（完整工作規範）。
- **其他代理**（Codex / Antigravity / Cursor …）：多數會自動讀取本檔 `AGENTS.md`。**請務必把 `CLAUDE.md` 也完整讀過** —— 那是最完整的工作規範，內容與「你是哪個代理」無關，一體適用。

## 三條最關鍵的規則（不論你是哪個代理）

1. **資料保護由伺服器端強制**，不是靠代理自律：客戶身分欄位只能數不能取值、金額必須包在聚合函式內，違規回 **403**。這是 `pbibridge_csharp/Program.cs` 的 `DataGuard` 做的，**任何代理都無法自我豁免**。管制清單在 `appsettings.json → DataProtection`（預設是通用樣式，請依你自己的資料模型增補）。詳見 `CLAUDE.md` 的「資料保護規範」。

2. **一律走 `tools/PBI-Bridge.ps1`**，不要自己直連 `msmdsrv` 或直接載入 ADOMD —— 那會繞過保護，也會踩到中文編碼問題。
   （Claude Code 另有 `.claude/hooks/guard-data-access.ps1` 擋這條繞路；**其他代理沒有這個 hook**，更要自律。但伺服器端的 DataGuard 仍會攔查詢回傳值，所以真正的防線一直都在。）

3. **標準開發流程**：
   讀（`Get-PbiSchema`）→ 動手前留退路（`New-PbiSnapshot`）→ 寫入（`Set-PbiMeasure` 等）→ 驗算（`Invoke-Dax`）→ 存檔（`Save-PbiModel`，確認 `fileChanged = true`）。

## 完整文件

| 檔案 | 內容 |
|---|---|
| **`CLAUDE.md`** | 完整工作規範（Claude Code 自動載入；其他代理請主動讀取）。這是最權威的一份。 |
| **`README.md`** | 給使用者（人）的完整說明：安裝、啟動、資料保護清單、與 AI 協作時使用者要做的事、疑難排解。 |
| **`API_Documentation.html`** | 端點總覽。 |

> 這是通用型工具，**不預設綁定任何特定報表或資料模型**。第一次使用請依 `README.md` 的「快速開始」（雙擊 `🚀啟動PBI終極儀表板.bat` 即可），並在 `appsettings.json` 補上你自己資料中的敏感欄位。
