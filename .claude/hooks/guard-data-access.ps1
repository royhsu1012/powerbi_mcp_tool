# =============================================================================
# PreToolUse hook — 強制所有資料存取都走橋接服務
# =============================================================================
# 為什麼需要這一層：
#   C# 伺服器的欄位管制只能管到「經過它的請求」。PowerShell 可以直接載入 ADOMD
#   元件連上 localhost:<msmdsrv port>，完全不經過橋接服務 —— 那條路上伺服器再嚴
#   也管不到。這個 hook 就是為了封死那條路。
#
# 這裡刻意「不」比對敏感欄位名稱：
#   同一個欄位名，DISTINCTCOUNT(v[customer_name]) 是合法的統計量，
#   拿去當 SUMMARIZECOLUMNS 的分組鍵才是外洩。正則分不出這兩者，
#   但伺服器的括號配對掃描分得出來。在這裡擋名稱只會擋掉合法用法。
#
# 攔截時：exit 2 → 工具呼叫被擋下，stderr 內容會回饋給 Claude。
# =============================================================================

$raw = [Console]::In.ReadToEnd()
if (-not $raw) { exit 0 }

try { $payload = $raw | ConvertFrom-Json } catch { exit 0 }

# 取出指令文字（Bash 與 PowerShell 工具都用 command 欄位）
$cmd = ''
if ($payload.tool_input -and $payload.tool_input.command) {
    $cmd = [string]$payload.tool_input.command
}
if ([string]::IsNullOrWhiteSpace($cmd)) { exit 0 }

$rules = @(
    @{
        Pattern = 'AdomdConnection|Microsoft\.AnalysisServices|AnalysisServices\.AdomdClient|Invoke-ASCmd|Provider\s*=\s*MSOLAP|LoadWithPartialName'
        Reason  = '偵測到直接連線 Analysis Services / 載入 ADOMD 元件的寫法。'
        Why     = '這條路徑完全繞過橋接服務，伺服器端的客戶身分與金額欄位管制對它無效。'
    },
    @{
        Pattern = '(Invoke-RestMethod|Invoke-WebRequest|curl|wget)[^\r\n]{0,200}localhost:5500'
        Reason  = '偵測到手刻 HTTP 請求直接呼叫橋接服務。'
        Why     = '請改用 tools\PBI-Bridge.ps1 提供的函式（Invoke-Dax / Get-PbiSchema …）。手刻請求容易繞開用戶端的防護與 UTF-8 編碼處理。'
    },
    @{
        Pattern = 'DataProtection[^\r\n]{0,120}(Enabled[^\r\n]{0,40}(false|\$false)|DenyColumns[^\r\n]{0,40}@\(\s*\))'
        Reason  = '偵測到試圖關閉或清空資料保護設定。'
        Why     = '調整保護範圍應由使用者明確要求，並且要說明會放寬哪些欄位。'
    }
)

foreach ($rule in $rules) {
    if ($cmd -match $rule.Pattern) {
        $msg = @"
⛔ 已被 PreToolUse 資料保護 hook 攔截

原因：$($rule.Reason)
說明：$($rule.Why)

正確做法：所有模型查詢一律經由 tools\PBI-Bridge.ps1 → localhost:5500，
讓伺服器端的欄位管制生效。若這次確實有正當理由需要例外，
請停下來向使用者說明你想做什麼、為什麼，由使用者自行執行。
"@
        [Console]::Error.WriteLine($msg)
        exit 2
    }
}

exit 0
