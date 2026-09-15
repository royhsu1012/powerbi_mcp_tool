# =============================================================================
# Test-DataGuard.ps1 — 對資料保護機制做紅隊測試
# =============================================================================
# 用法：
#   . ".\tools\PBI-Bridge.ps1"
#   .\tools\Test-DataGuard.ps1                      # 唯讀測試
#   .\tools\Test-DataGuard.ps1 -IncludeWriteTests   # 另做「建衍生物件洗資料」測試
#
# 安全設計（重要）：
#   1. 每個攻擊查詢都設 MaxRows = 1。萬一真有漏洞，外洩的是一個值而不是整批。
#   2. 腳本**只回報有沒有被擋下，絕不印出回傳內容**。防護的測試本身不該變成外洩管道。
#   3. 管制清單從 appsettings.json 讀取，不寫死欄位名 —— 換模型也能用。
#
# 判讀：
#   ✅ 通過 = 行為與預期一致（該擋的擋了、該放行的放行了）
#   ❌ 失敗 = 有洞，或誤擋了合法查詢
# =============================================================================

param(
    [switch]$IncludeWriteTests,
    # 同時開多個 Power BI 時必填（Port 或檔名片段）。本腳本在自己的作用域載入橋接函式，
    # 呼叫端先前做過的 Use-PbiInstance 不會帶進來。
    [string]$Target
)

# ⚠️ 一定要在「本腳本自己的作用域」裡載入橋接函式，不能依賴呼叫端先 dot-source。
#    PowerShell 對非模組函式的 $script: 是動態解析到「當前執行中的腳本」，
#    不是函式定義處 —— 所以由呼叫端載入時，$script:PbiBaseUrl / PbiTarget
#    在這裡全會是 $null，症狀是 "Invalid URI: The hostname could not be parsed."
#    （長期解法是把 PBI-Bridge.ps1 改成 .psm1，模組作用域才是靜態的。）
. (Join-Path $PSScriptRoot "PBI-Bridge.ps1") | Out-Null
if ($Target) {
    Use-PbiInstance $Target
    if (-not $script:PbiTarget) { throw "找不到符合「$Target」的 Power BI 實例，請用 Get-PbiInstances 確認 Port" }
}

$root = Split-Path -Parent $PSScriptRoot
$cfg  = Get-Content (Join-Path $root "pbibridge_csharp\appsettings.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$dp   = $cfg.DataProtection
if (-not $dp -or -not $dp.Enabled) { throw "appsettings.json 的 DataProtection 未啟用" }

# 欄位比對必須與伺服器一致（Program.cs 的 DataGuard.Norm / Matches）：
# 忽略大小寫、空白、底線、連字號；清單項目可以是 * 樣式；AllowColumns 優先。
# 逐字 -contains 比對的話，預設清單全是 *customer* 這類樣式，永遠挑不到欄位。
function Get-NormColumnName([string]$s) {
    if (-not $s) { return '' }
    return ($s -replace '[ _-]', '').ToLowerInvariant()
}
function Test-ColumnListed([string]$Name, $List) {
    $n = Get-NormColumnName $Name
    if (-not $n) { return $false }
    foreach ($item in @($List)) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $p = Get-NormColumnName $item
        if ($item.Contains('*')) {
            $re = '^' + ((($p -split '\*') | ForEach-Object { [regex]::Escape($_) }) -join '.*') + '$'
            if ($n -cmatch $re) { return $true }
        } elseif ($n -ceq $p) { return $true }
    }
    return $false
}
function Test-DeniedColumn([string]$Name) {
    (Test-ColumnListed $Name $dp.DenyColumns) -and -not (Test-ColumnListed $Name $dp.AllowColumns)
}
function Test-MoneyColumn([string]$Name) {
    (Test-ColumnListed $Name $dp.AggregateOnlyColumns) -and
    -not (Test-ColumnListed $Name $dp.AllowColumns) -and -not (Test-DeniedColumn $Name)
}

# 從模型裡挑「實際存在」的管制欄位來測 —— 清單是通用的，模型不一定每個欄都有
$schema = Get-PbiSchema
$numericTypes = @('Int64', 'Double', 'Decimal')

# 身分欄位優先挑文字型：MAX / CONCATENATEX 取到的才會是真實名稱，最貼近實際風險
$pick = @(foreach ($t in $schema.Tables) { foreach ($c in $t.Columns) {
    if (Test-DeniedColumn $c.Name) { [PSCustomObject]@{ T = $t.Name; C = $c.Name; IsText = ($c.DataType -eq 'String') } }
} }) | Sort-Object { -not $_.IsText } | Select-Object -First 1
if (-not $pick) { throw "這個模型裡找不到任何符合 DenyColumns 清單（含 * 樣式）的欄位，無從測起。請確認 appsettings.json 的清單涵蓋你的欄位命名。" }

# 金額欄位只挑數值型：對文字欄 SUM 會是 DAX 錯誤，讓「合法：金額聚合」被誤判為失敗
$money = @(foreach ($t in $schema.Tables) { foreach ($c in $t.Columns) {
    if ((Test-MoneyColumn $c.Name) -and ($numericTypes -contains $c.DataType)) { [PSCustomObject]@{ T = $t.Name; C = $c.Name } }
} }) | Select-Object -First 1

$D = "'$($pick.T)'[$($pick.C)]"
Write-Host "🎯 測試對象" -ForegroundColor Cyan
Write-Host "   客戶身分欄位：$D"
if ($money) { $M = "'$($money.T)'[$($money.C)]"; Write-Host "   金額欄位：    $M" }
Write-Host ""

$cases = @(
    @{ N='直接把身分欄當分組鍵'; Block=$true
       Q="EVALUATE SUMMARIZECOLUMNS($D, `"c`", COUNTROWS('$($pick.T)'))" }
    @{ N='別名投影繞過欄名比對'; Block=$true
       Q="EVALUATE SELECTCOLUMNS('$($pick.T)', `"n`", $D)" }
    @{ N='整表 EVALUATE（文字裡無欄名）'; Block=$true
       Q="EVALUATE '$($pick.T)'" }
    @{ N='用 MAX 取單一真實值'; Block=$true
       Q="EVALUATE ROW(`"x`", MAX($D))" }
    @{ N='用 CONCATENATEX 串成一個字串'; Block=$true
       Q="EVALUATE ROW(`"x`", CONCATENATEX(TOPN(1, '$($pick.T)'), $D, `",`"))" }
    @{ N='DEFINE MEASURE 把欄位藏進量值'; Block=$true
       Q="DEFINE MEASURE '$($pick.T)'[__probe] = MAX($D) EVALUATE ROW(`"x`", [__probe])" }
    @{ N='TOPN 抽樣後投影'; Block=$true
       Q="EVALUATE SELECTCOLUMNS(TOPN(1, '$($pick.T)'), `"n`", $D)" }
    @{ N='✅ 合法：計數類函式取統計量'; Block=$false
       Q="EVALUATE ROW(`"n`", DISTINCTCOUNT($D))" }
    @{ N='✅ 合法：COUNTROWS'; Block=$false
       Q="EVALUATE ROW(`"n`", COUNTROWS('$($pick.T)'))" }
)

if ($money) {
    $cases += @{ N='金額逐列取值';        Block=$true;  Q="EVALUATE SELECTCOLUMNS('$($money.T)', `"a`", $M)" }
    $cases += @{ N='✅ 合法：金額聚合';   Block=$false; Q="EVALUATE ROW(`"s`", SUM($M))" }
}

$pass = 0; $fail = 0
foreach ($c in $cases) {
    $blocked = $false; $err = $null
    try {
        # 直接打 /api/query，刻意繞過 Invoke-Dax 的用戶端守門 ——
        # 用戶端守門不是安全邊界（模型可以選擇不用它），這裡要驗證的是伺服器。
        # MaxRows = 1：萬一沒被擋下，外洩的是一個值而不是整批。
        $null = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{ Query = $c.Q; MaxRows = 1 }
    } catch {
        $err = "$_"
        if ($err -match '403|資料保護') { $blocked = $true }
    }

    $ok = ($blocked -eq $c.Block)
    if ($ok) { $pass++ } else { $fail++ }

    $verdict = if ($ok) { '✅ 通過' } else { '❌ 失敗' }
    $actual  = if ($blocked) { '被擋下' } else { '通過' }
    Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
                $verdict, $c.N, $(if ($c.Block) { '擋下' } else { '放行' }), $actual) `
               -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    # 不是 403 的錯誤要顯示出來，否則 DAX 語法錯誤會被誤判成「防護生效」
    if ($err -and -not $blocked) { Write-Host "      （非管制錯誤：$($err -replace '\s+',' ' -replace '^(.{140}).*','$1…')）" -ForegroundColor DarkGray }
}

# ── 金額查詢的列數上限（後備防線）─────────────────────────────────────────
# 金額有被正確聚合（過得了 Layer A），但分組鍵基數太高 —— 那已接近逐筆金額。
# 注意：必須用比上限大的 MaxRows 送，否則結果會先被截斷到 MaxRows，
#       rows.Count 永遠不會超過上限，這道檢查就測不到。
if ($money) {
    $mt = $money.T
    # 動態找一個「非管制、且相異值 > 上限」的欄位當分組鍵，讓測試不綁死特定模型
    $groupCol = $null
    foreach ($c in ($schema.Tables | Where-Object { $_.Name -eq $mt }).Columns) {
        if (Test-DeniedColumn $c.Name) { continue }
        if (Test-ColumnListed $c.Name $dp.AggregateOnlyColumns) { continue }
        if ($c.DataType -ne 'String') { continue }
        try {
            $n = (Invoke-Dax "EVALUATE ROW(`"n`", DISTINCTCOUNT('$mt'[$($c.Name)]))").Rows[0].'[n]'
            if ($n -gt $dp.MaxRowsWithMoney) { $groupCol = $c.Name; break }
        } catch { }
    }
    if ($groupCol) {
        $blocked = $false
        try {
            $null = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{
                Query = "EVALUATE SUMMARIZECOLUMNS('$mt'[$groupCol], `"s`", SUM($M))"
                MaxRows = ($dp.MaxRowsWithMoney * 10) }
        } catch { if ("$_" -match '403|資料保護') { $blocked = $true } }
        if ($blocked) { $pass++ } else { $fail++ }
        Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
                    $(if ($blocked) { '✅ 通過' } else { '❌ 失敗' }),
                    "金額按高基數鍵分組($groupCol)", '擋下',
                    $(if ($blocked) { '被擋下' } else { '通過' })) `
                   -ForegroundColor $(if ($blocked) { 'Green' } else { 'Red' })
    } else {
        Write-Host "⚠️ 略過  金額列數上限測試 —— 這個模型找不到相異值 > $($dp.MaxRowsWithMoney) 的非管制文字欄位" -ForegroundColor Yellow
    }
}

# ── DMV 路徑 ───────────────────────────────────────────────────────────────
$blocked = $false
try { $null = Invoke-PbiDmv 'SELECT * FROM $SYSTEM.TMSCHEMA_PARTITIONS' } catch {
    if ("$_" -match '403|資料保護') { $blocked = $true }
}
if ($blocked) { $pass++ } else { $fail++ }
Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
            $(if ($blocked) { '✅ 通過' } else { '❌ 失敗' }), 'DMV 拉 M 腳本定義', '擋下',
            $(if ($blocked) { '被擋下' } else { '通過' })) `
           -ForegroundColor $(if ($blocked) { 'Green' } else { 'Red' })

# ── 寫入型測試：建衍生物件把資料洗到未管制的名稱下 ──────────────────────────
if ($IncludeWriteTests) {
    Write-Host "`n🧪 寫入型測試（會在模型中建立暫時量值，測完刪除）" -ForegroundColor Yellow
    New-PbiSnapshot -Label "before-dataguard-test" | Out-Null

    $tmp = "__guard_probe__"
    try {
        Set-PbiMeasure -Table $pick.T -Name $tmp -Expression "MAX($D)" | Out-Null
        $blocked = $false
        try {
            $null = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{
                Query = "EVALUATE ROW(`"x`", [$tmp])"; MaxRows = 1 }
        } catch {
            if ("$_" -match '403|資料保護') { $blocked = $true }
        }
        if ($blocked) { $pass++ } else { $fail++ }
        Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
                    $(if ($blocked) { '✅ 通過' } else { '❌ 失敗' }), '新建量值包裝敏感欄位', '擋下',
                    $(if ($blocked) { '被擋下' } else { '通過' })) `
                   -ForegroundColor $(if ($blocked) { 'Green' } else { 'Red' })
    } finally {
        try { Remove-PbiMeasure -Table $pick.T -Name $tmp | Out-Null; Write-Host "   已清除暫時量值 [$tmp]" -ForegroundColor Gray }
        catch { Write-Host "   ⚠️ 暫時量值 [$tmp] 清除失敗，請手動刪除" -ForegroundColor Yellow }
    }
}

Write-Host ""
Write-Host ("結果：$pass 通過 / $fail 失敗") -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { Write-Host "❗ 有項目未達預期 —— 上面標紅的那幾行就是漏洞或誤擋。" -ForegroundColor Red }
