# 修改服務本身：專案結構、建置與重啟

> 這是 [CLAUDE.md](../CLAUDE.md) 的延伸，用到才讀：要改 `Program.cs`、設定檔、啟動檔、hook 或 `tools\*.ps1` 之前，先把這一份讀完。
> 資料保護、防毒相容與確認視窗的規矩一律以 CLAUDE.md 為準。

## 專案結構

```
PBI_AI_Bridge/
├── CLAUDE.md                      ← AI 工作規範：每次對話都會載入，只放常駐的規則 —— 加東西之前先想能不能放進 docs\
├── AGENTS.md                      ← 其他 AI 代理（Codex / Cursor …）的入口，指向 CLAUDE.md
├── README.md                      ← 給人看的完整說明（安裝、使用流程、資料保護清單、疑難排解）
├── docs/                          ← 給 AI 的延伸規範，用到才讀（CLAUDE.md 的「用到才讀的文件」列出什麼時候讀哪一份）
│   ├── power-query.md             ← 讀寫 Power Query（M）的流程
│   ├── pbip-report-editing.md     ← 直接改 PBIP 報表檔的步驟與踩過的坑
│   ├── troubleshooting.md         ← 錯誤訊息與狀態碼對照
│   └── service-development.md     ← 本文件
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
│                                     ⚠️ 這裡「沒有」目標檔案路徑 —— 見 CLAUDE.md 開頭的說明
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
