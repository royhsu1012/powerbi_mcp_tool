# =============================================================================
# Test-BridgeClient.ps1 — 用戶端自我檢查（不需要服務，也不需要開 Power BI）
# =============================================================================
# 用法：
#   .\tools\Test-BridgeClient.ps1               # 只列出失敗的項目與總結
#   .\tools\Test-BridgeClient.ps1 -ShowPassed   # 連通過的也列出來
#
# 檢查的是兩類「壞了也不會有人發現」的問題：
#
#   ① 檔案層級的不變條件 —— 違反時症狀都很隱晦：
#        · .ps1 含中文卻沒有 UTF-8 BOM → PowerShell 5.1 把整份讀成亂碼
#        · .ps1 混進控制字元 → 實際發生過：正規表示式裡的「反斜線 b」被存成退格字元
#          (0x08)，Set-PbiMQuery 因此拒絕所有 M 腳本，直到真的要用才發現
#        · 啟動檔 .bat 出現非 ASCII 字元或 LF 行尾 → cmd 的 goto 會跳到錯的位置
#
#   ② 每個函式送出去的請求長什麼樣 —— 把 Invoke-PbiApi 換成記錄用的替身，逐一呼叫，
#      核對端點路徑與 body。伺服器不在場，所以這只能證明「用戶端這一半沒寫錯」；
#      伺服器那一半要開著服務跑 Test-DataGuard.ps1 與實際操作來驗證。
#
# 改過 tools\*.ps1、hook 或啟動檔之後跑一次。結束代碼：0 = 全過，1 = 有失敗。
# =============================================================================

param([switch]$ShowPassed)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$script:Pass     = 0
$script:Failures = New-Object System.Collections.ArrayList
$script:Covered  = @{}

function Check {
    param([string]$Name, [bool]$Ok, [string]$Detail)
    if ($Ok) {
        $script:Pass++
        if ($ShowPassed) { Write-Host "✅ $Name" -ForegroundColor Green }
    } else {
        [void]$script:Failures.Add($Name)
        Write-Host "❌ $Name" -ForegroundColor Red
        if ($Detail) { Write-Host "      $Detail" -ForegroundColor DarkGray }
    }
}

# =============================================================================
# ① 檔案層級的不變條件
# =============================================================================
Write-Host "① 檔案不變條件（BOM、控制字元、語法、啟動檔編碼）" -ForegroundColor Cyan

# 以 Latin-1 解讀位元組：每個位元組對到一個字元，可以直接用正規表示式找特定位元組
$latin1       = [System.Text.Encoding]::GetEncoding(28591)
$ctrlPattern  = '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]'
$bridgeBroken = $false

$scripts = @(Get-ChildItem -LiteralPath (Join-Path $root 'tools') -Filter *.ps1)
$hookDir = Join-Path $root '.claude\hooks'
if (Test-Path -LiteralPath $hookDir) { $scripts += @(Get-ChildItem -LiteralPath $hookDir -Filter *.ps1) }

foreach ($f in $scripts) {
    $rel    = $f.FullName.Substring($root.Length + 1)
    $bytes  = [System.IO.File]::ReadAllBytes($f.FullName)
    $raw    = $latin1.GetString($bytes)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $body   = if ($hasBom) { $raw.Substring(3) } else { $raw }

    Check "${rel}：含非 ASCII 字元時必須有 UTF-8 BOM" ($hasBom -or -not [regex]::IsMatch($body, '[\x80-\xFF]')) `
          "PowerShell 5.1 會用系統碼頁讀沒有 BOM 的檔案，中文全部變亂碼"

    $ctrl  = [regex]::Matches($body, $ctrlPattern)
    $where = @($ctrl | Select-Object -First 5 | ForEach-Object {
        $line = ([regex]::Matches($body.Substring(0, $_.Index), "`n")).Count + 1
        "第 {0} 行 0x{1:X2}" -f $line, [int][char]$_.Value
    }) -join '、'
    Check "${rel}：沒有控制字元" ($ctrl.Count -eq 0) `
          "找到 $($ctrl.Count) 個：$where —— 常見成因是跳脫序列在產生檔案時被當成真的控制字元寫進去"

    $tokens = $null; $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
    Check "${rel}：語法可解析" ($errors.Count -eq 0) `
          (($errors | Select-Object -First 3 | ForEach-Object { "第 $($_.Extent.StartLineNumber) 行：$($_.Message)" }) -join '；')
    if ($f.Name -eq 'PBI-Bridge.ps1' -and $errors.Count -gt 0) { $bridgeBroken = $true }
}

$bats = @(Get-ChildItem -LiteralPath $root -Filter *.bat)
Check "專案根目錄有啟動檔 .bat" ($bats.Count -ge 1) "找不到任何 .bat"
foreach ($b in $bats) {
    $raw = $latin1.GetString([System.IO.File]::ReadAllBytes($b.FullName))
    Check "啟動檔 .bat：只有 ASCII 字元" (-not [regex]::IsMatch($raw, '[\x80-\xFF]')) `
          "cmd 以位元組定位，多位元組字元會讓 goto 跳到錯的標籤（UTF-8 BOM 也算）"
    Check "啟動檔 .bat：行尾全部是 CRLF" (-not [regex]::IsMatch($raw, '(?<!\r)\n')) `
          "LF 行尾會讓 cmd 找不到標籤。Git Bash 的 sed -i 會把 CR 吃掉 —— 請用 PowerShell 編輯這個檔"
    Check "啟動檔 .bat：沒有控制字元" (-not [regex]::IsMatch($raw, $ctrlPattern)) ""
    Check "啟動檔 .bat：不產生、也不寫入 API 金鑰" (-not [regex]::IsMatch($raw, 'NewGuid|PUT_YOUR_OWN')) `
          "金鑰由服務第一次啟動時產生，存在使用者資料夾 —— 專案資料夾裡不該有任何一份"
}

# 金鑰不在專案資料夾裡：範本如果又長出金鑰欄位，啟動檔會把它原樣複製成 appsettings.json，
# 使用者就會以為要自己填一把（而服務根本不讀它）
$template = Join-Path $root 'pbibridge_csharp\appsettings.template.json'
if (Test-Path -LiteralPath $template) {
    $cfg = $null
    try { $cfg = Get-Content -LiteralPath $template -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
    Check "設定範本：是合法的 JSON" ($null -ne $cfg) "啟動檔會把它原樣複製成 appsettings.json，壞掉的話服務起不來"
    if ($cfg) {
        $names = @(); if ($cfg.Security) { $names = @($cfg.Security.PSObject.Properties.Name) }
        Check "設定範本：沒有 API 金鑰欄位" (-not ($names -contains 'ApiKey')) "Security.ApiKey 已經不使用"
    }
}

# =============================================================================
# ② 載入橋接函式
# =============================================================================
Write-Host "② 用戶端函式（請求路徑與 body）" -ForegroundColor Cyan

if ($bridgeBroken) {
    # 載入一份解析不了的檔案只會噴出整頁亂碼 —— 上面已經指出是哪一行，這裡直接收尾
    Check 'PBI-Bridge.ps1 可以載入' $false "語法解析失敗（見上），略過函式檢查"
    Write-Host ""
    Write-Host ("結果：{0} 通過 / {1} 失敗" -f $script:Pass, $script:Failures.Count) -ForegroundColor Red
    exit 1
}

$oldQuiet = $env:PBI_BRIDGE_QUIET
$env:PBI_BRIDGE_QUIET = '1'
try     { . (Join-Path $PSScriptRoot 'PBI-Bridge.ps1') }
finally { $env:PBI_BRIDGE_QUIET = $oldQuiet }

# ── ②-a 真的 Invoke-PbiApi：編碼與標頭 ─────────────────────────────────────
# 這是整個工具最容易踩的坑（PowerShell 5.1 送字串 body 不是 UTF-8），所以拿真的函式來測，
# 只把最底層的 Invoke-RestMethod 與金鑰讀取換掉 —— 不發出任何網路請求，也不讀真的金鑰。

# 換掉之前先看一眼真的那個：金鑰只該從使用者資料夾讀（只看原始碼，不執行它）
$keyReader = (Get-Command Get-PbiApiKey).Definition
Check 'Get-PbiApiKey 從使用者資料夾讀金鑰，不讀專案資料夾' `
      (($keyReader -match 'LocalApplicationData') -and ($keyReader -notmatch 'appsettings|Get-PbiRootPath')) `
      "金鑰由服務產生，存在 %LOCALAPPDATA%\PBI_AI_Bridge\api-key.txt；專案資料夾裡沒有金鑰"

$script:Rest = $null
function Get-PbiApiKey { 'TEST-KEY-NOT-REAL' }
function Invoke-RestMethod {
    param($Uri, $Method, $Headers, $ContentType, $Body, $TimeoutSec)
    $script:Rest = [PSCustomObject]@{
        Uri = "$Uri"; Method = "$Method"; Headers = $Headers; ContentType = "$ContentType"; Body = $Body
    }
    [PSCustomObject]@{ ok = $true }
}

$script:PbiTarget = $null
$null = Invoke-PbiApi -Path '/api/schema'
Check 'Invoke-PbiApi GET：網址正確、帶金鑰、沒選目標時不帶 X-PBI-Target' (
    $script:Rest.Uri -eq 'http://localhost:5500/api/schema' -and
    $script:Rest.Headers['X-API-Key'] -eq 'TEST-KEY-NOT-REAL' -and
    -not $script:Rest.Headers.ContainsKey('X-PBI-Target')) "實際送出的網址：$($script:Rest.Uri)"

$script:PbiTarget = '55820'
$null = Invoke-PbiApi -Path '/api/upsert-measure' -Method POST -Body @{
    TableName = '量值'; MeasureName = '銷售總額'; Expression = "SUM('訂單'[金額])"
}
$sent    = $script:Rest.Body
$decoded = if ($sent -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($sent) | ConvertFrom-Json } else { $null }
Check 'Invoke-PbiApi POST：body 以位元組送出（不是字串）' ($sent -is [byte[]]) `
      "送字串的話 PowerShell 5.1 會用系統碼頁編碼，中文在傳輸中損毀"
Check 'Invoke-PbiApi POST：中文經過 UTF-8 編碼後原樣還原' (
    $decoded -and $decoded.TableName -ceq '量值' -and $decoded.MeasureName -ceq '銷售總額' -and
    $decoded.Expression -ceq "SUM('訂單'[金額])") ""
Check 'Invoke-PbiApi POST：Content-Type 標明 charset=utf-8' ($script:Rest.ContentType -match 'charset=utf-8') ""
Check 'Invoke-PbiApi：選定目標後帶上 X-PBI-Target' ($script:Rest.Headers['X-PBI-Target'] -eq '55820') ""
$script:PbiTarget = $null
Remove-Item function:Invoke-RestMethod

# ── ②-b 把 Invoke-PbiApi 換成記錄用的替身，逐一核對每個函式送出的請求 ──────────
$script:Calls  = New-Object System.Collections.ArrayList
$script:Canned = @{}          # 端點路徑 → 要回傳的假資料

function Invoke-PbiApi {
    param([string]$Path, [string]$Method = 'GET', [hashtable]$Body,
          [int]$TimeoutSec = 120, [switch]$NoRetry)
    # 和真的一樣走一遍 JSON 序列化再解回來：核對的是「伺服器會收到的東西」，不是 PowerShell 物件
    $json = $null
    if ($Method -eq 'POST') { $json = ($Body | ConvertTo-Json -Compress -Depth 10) | ConvertFrom-Json }
    [void]$script:Calls.Add([PSCustomObject]@{
        Path = $Path; Method = $Method; Body = $Body; Json = $json
    })
    if ($script:Canned.ContainsKey($Path)) { return $script:Canned[$Path] }
    [PSCustomObject]@{ message = 'ok' }
}

function Test-Same($Actual, $Expected) {
    if ($Expected -is [array]) {
        return ($Actual -is [array]) -and ((@($Actual) -join "`n") -ceq (@($Expected) -join "`n"))
    }
    if ($Expected -is [bool]) { return ($Actual -is [bool]) -and ($Actual -eq $Expected) }
    return "$Actual" -ceq "$Expected"
}

function Expect-Call {
    <#  執行 $Act，核對它送出的最後一個請求：路徑、body 裡該有的值、不該出現的鍵  #>
    param([string]$Fn, [string]$Name, [scriptblock]$Act, [string]$Path,
          [hashtable]$Has = @{}, [string[]]$Lacks = @())
    $script:Covered[$Fn] = $true
    $n = $script:Calls.Count
    try { $null = & $Act 6>$null } catch { Check "${Fn}：$Name" $false "丟出例外：$_"; return }
    if ($script:Calls.Count -eq $n) { Check "${Fn}：$Name" $false "沒有呼叫任何端點"; return }
    $c = $script:Calls[$script:Calls.Count - 1]
    $problems = @()
    if ($c.Path -ne $Path) { $problems += "端點是 $($c.Path)，預期 $Path" }
    foreach ($k in $Has.Keys) {
        # 不能寫成 $actual = if (...) { $c.Json.$k }：if 當運算式用時輸出會經過管線，
        # 單一元素的陣列被攤平成純量 —— 剛好把「是不是陣列」這個要檢查的東西弄丟。
        $actual = $null
        if ($c.Json) { $actual = $c.Json.$k }
        if (-not (Test-Same $actual $Has[$k])) {
            $problems += "body.$k = [$(@($actual) -join ',')]（$(if ($null -eq $actual) { 'null' } else { $actual.GetType().Name })），預期 [$(@($Has[$k]) -join ',')]"
        }
    }
    foreach ($k in $Lacks) {
        if ($c.Json -and ($c.Json.PSObject.Properties.Name -contains $k)) { $problems += "body 不該有 $k（沒指定的參數不能送出去蓋掉原值）" }
    }
    Check "${Fn}：$Name" ($problems.Count -eq 0) ($problems -join '；')
}

function Expect-Throw {
    <#  $Act 必須在本機就丟出例外，而且不能有任何請求送出去  #>
    param([string]$Fn, [string]$Name, [scriptblock]$Act)
    $script:Covered[$Fn] = $true
    $n = $script:Calls.Count
    $threw = $false
    try { $null = & $Act 6>$null } catch { $threw = $true }
    Check "${Fn}：$Name" ($threw -and $script:Calls.Count -eq $n) `
          $(if (-not $threw) { '應該要被擋下，卻沒有丟出例外' } else { '丟了例外，但請求已經送出去了' })
}

# ── 讀取 ────────────────────────────────────────────────────────────────────
Expect-Call Get-PbiSchema        '端點' { Get-PbiSchema }        '/api/schema'
Expect-Call Get-PbiRelationships '端點' { Get-PbiRelationships } '/api/relationships'
Expect-Call Get-PbiRoles         '端點' { Get-PbiRoles }         '/api/roles'
Expect-Call Get-PbiExpressions   '端點' { Get-PbiExpressions }   '/api/expressions'
Expect-Call Test-PbiModel        '端點' { Test-PbiModel }        '/api/validate'
Expect-Call Test-PbiReport       '端點' { Test-PbiReport }       '/api/validate-report'
Expect-Call Get-PbiInfo          '端點' { Get-PbiInfo }          '/api/pbi-info'
Expect-Call Get-PbiModelProps    '端點' { Get-PbiModelProps }    '/api/model-props'

$script:Canned['/api/snapshots'] = [PSCustomObject]@{ Snapshots = @('a.json', 'b.json') }
Expect-Call Get-PbiSnapshots '端點' { Get-PbiSnapshots } '/api/snapshots'

$script:Canned['/api/schema'] = [PSCustomObject]@{ Tables = @(
    [PSCustomObject]@{
        Name     = 'Dim 產品'
        Columns  = @([PSCustomObject]@{ Name = '類別'; DataType = 'String' },
                     [PSCustomObject]@{ Name = '單價'; DataType = 'Double' })
        Measures = @([PSCustomObject]@{ Name = '銷售總額'; Expression = 'SUM(x)'; FormatString = '#,0'; DisplayFolder = ''; Description = '' })
    },
    [PSCustomObject]@{ Name = '空表'; Columns = @(); Measures = @() }
) }
$script:Covered['Get-PbiMeasures'] = $true
$m = @(Get-PbiMeasures)
Check 'Get-PbiMeasures：攤平成一列一個量值，格式欄位叫 Format（與 Set-PbiMeasure -Format 對應）' (
    $m.Count -eq 1 -and $m[0].Table -eq 'Dim 產品' -and $m[0].Measure -eq '銷售總額' -and $m[0].Format -eq '#,0') ""
Check 'Get-PbiMeasures：-TableFilter 會過濾資料表' (@(Get-PbiMeasures -TableFilter '空*').Count -eq 0) ""

# ── 實例切換 ────────────────────────────────────────────────────────────────
$script:Covered['Get-PbiInstances'] = $true
$script:Covered['Use-PbiInstance']  = $true
$one = [PSCustomObject]@{ port = 51111; fileName = '銷售報表.pbix'; filePath = 'C:\a\銷售報表.pbix'; windowTitle = '銷售報表'; kind = 'PBIX' }
$two = [PSCustomObject]@{ port = 52222; fileName = '銷售預測.pbip'; filePath = 'C:\b\銷售預測.pbip'; windowTitle = '銷售預測'; kind = 'PBIP' }
$script:Canned['/api/instances'] = [PSCustomObject]@{ Instances = @($one) }
$r = Get-PbiInstances
Check 'Get-PbiInstances：只有一個實例時仍是陣列（.Count = 1）' (($r -is [array]) -and $r.Count -eq 1) ""
$script:Canned['/api/instances'] = [PSCustomObject]@{ Instances = @($one, $two) }
$null = Use-PbiInstance 預測 6>$null
Check 'Use-PbiInstance：檔名片段唯一符合 → 記住 Port 與完整路徑' (
    $script:PbiTarget -eq '52222' -and $script:PbiTargetPath -eq 'C:\b\銷售預測.pbip') ""
$null = Use-PbiInstance 6>$null
Check 'Use-PbiInstance：不帶參數會清除選擇' ($null -eq $script:PbiTarget -and $null -eq $script:PbiTargetPath) ""
$null = Use-PbiInstance 銷售 6>$null
Check 'Use-PbiInstance：片段同時符合兩個實例 → 不選（不猜）' ($null -eq $script:PbiTarget) ""
$null = Use-PbiInstance 不存在的檔 6>$null
Check 'Use-PbiInstance：沒有符合的實例 → 不選' ($null -eq $script:PbiTarget) ""
$null = Use-PbiInstance 51111 6>$null
Check 'Use-PbiInstance：用 Port 精確指定' ($script:PbiTarget -eq '51111') ""
$null = Use-PbiInstance 6>$null

# ── 查詢 ────────────────────────────────────────────────────────────────────
# 呼叫端沒有任何「自己放行」的開關：欄位能不能讀、能不能用代號分組，由使用者替那一欄設的等級決定；
# 單一句查詢要例外放行，是伺服器在使用者的螢幕跳確認視窗（-AskUser 只是請伺服器去問）
Expect-Call Invoke-Dax '一般查詢（不帶任何放寬用的旗標）' { Invoke-Dax 'EVALUATE ROW("結果", [銷售總額])' } '/api/query' `
    -Has @{ Query = 'EVALUATE ROW("結果", [銷售總額])'; MaxRows = 1000; TimeoutSeconds = 60 } -Lacks Pseudonymize, AskUser
foreach ($gone in 'Pseudonymize', 'DetailToken') {
    Check "Invoke-Dax：沒有 -$gone 參數" (-not (Get-Command Invoke-Dax).Parameters.ContainsKey($gone)) `
          "這種開關或代碼等於讓呼叫端自己決定看不看得到；放行改由使用者在確認視窗決定"
}
$script:Covered['Invoke-Dax'] = $true
Expect-Call Invoke-Dax '整表查詢原樣送給伺服器判斷（本機不再另外擋一層）' { Invoke-Dax "EVALUATE 'Sales'" } '/api/query' `
    -Has @{ Query = "EVALUATE 'Sales'" } -Lacks AskUser
Expect-Call Invoke-Dax '-AskUser：請伺服器去問使用者；沒指定列數就不送 MaxRows' { Invoke-Dax "EVALUATE 'Sales'" -AskUser } '/api/query' `
    -Has @{ Query = "EVALUATE 'Sales'"; AskUser = $true } -Lacks MaxRows
Expect-Call Invoke-Dax '-AskUser 加上明確的 -MaxRows：照送' { Invoke-Dax "EVALUATE 'Sales'" -AskUser -MaxRows 20 } '/api/query' `
    -Has @{ AskUser = $true; MaxRows = 20 }
Expect-Call  Invoke-PbiDmv '端點與參數' { Invoke-PbiDmv 'SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLES' } '/api/dmv' `
    -Has @{ Query = 'SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLES'; MaxRows = 10000 }

$script:Covered['Get-PbiModelStats'] = $true
$script:Canned['/api/dmv'] = [PSCustomObject]@{ Truncated = $false; RowCount = 3; Rows = @(
    [PSCustomObject]@{ DIMENSION_NAME = 'A'; COLUMN_ID = 'c1'; USED_SIZE = 1MB },
    [PSCustomObject]@{ DIMENSION_NAME = 'A'; COLUMN_ID = 'c2'; USED_SIZE = 2MB },
    [PSCustomObject]@{ DIMENSION_NAME = 'B'; COLUMN_ID = 'c1'; USED_SIZE = 1MB }) }
$stats = @(Get-PbiModelStats)
Check 'Get-PbiModelStats：按資料表加總並由大到小排序' (
    $stats.Count -eq 2 -and $stats[0].Table -eq 'A' -and $stats[0].SizeMB -eq 3 -and $stats[0].Columns -eq 2) ""
$script:Canned.Remove('/api/dmv')

# ── 資料保護：每個欄位的等級 ─────────────────────────────────────────────────
$script:Covered['Get-PbiProtection'] = $true
$script:Canned['/api/protection'] = [PSCustomObject]@{
    enabled = $true; problem = $null
    summary = [PSCustomObject]@{ open = 2; countOnly = 1; pseudonym = 1; aggregateOnly = 0 }
    tables = @(
        [PSCustomObject]@{ name = '客戶'; columns = @(
            [PSCustomObject]@{ name = '客戶名稱'; dataType = 'String'; level = 'countOnly'; source = 'explicit'; rule = $null },
            [PSCustomObject]@{ name = '地區';     dataType = 'String'; level = 'open';      source = 'default';  rule = $null }) },
        [PSCustomObject]@{ name = '訂單'; columns = @(
            [PSCustomObject]@{ name = '業務員';   dataType = 'String'; level = 'pseudonym'; source = 'explicit'; rule = $null },
            [PSCustomObject]@{ name = '數量';     dataType = 'Int64';  level = 'open';      source = 'default';  rule = $null }) })
}
$p = @(Get-PbiProtection)
Check 'Get-PbiProtection：預設只列受限的欄位' (
    $p.Count -eq 2 -and $p[0].Table -eq '客戶' -and $p[0].Column -eq '客戶名稱' -and $p[0].Level -eq 'countOnly' -and $p[1].Level -eq 'pseudonym') ""
Check 'Get-PbiProtection：-All 列出全部，-Table 可以過濾' (
    @(Get-PbiProtection -All).Count -eq 4 -and @(Get-PbiProtection -All -Table '訂單').Count -eq 2) ""
Check 'Get-PbiProtection：-Raw 回傳伺服器的原始回應' ((Get-PbiProtection -Raw).summary.countOnly -eq 1) ""
$script:Canned.Remove('/api/protection')
$script:Covered['Set-PbiProtection'] = $true
$null = Set-PbiProtection -Table '客戶' -Column '電話' -Level countOnly
$sp = $script:Calls[$script:Calls.Count - 1]
Check 'Set-PbiProtection：只改一欄時 Changes 仍是陣列，表名、欄名、等級完整送出' (
    $sp.Path -eq '/api/protection' -and $sp.Method -eq 'POST' -and ($sp.Json.Changes -is [array]) -and $sp.Json.Changes.Count -eq 1 -and
    $sp.Json.Changes[0].Table -ceq '客戶' -and $sp.Json.Changes[0].Column -ceq '電話' -and $sp.Json.Changes[0].Level -ceq 'countOnly') ""
$null = Set-PbiProtection -Table '客戶' -Column '電話', '地址' -Level pseudonym
Check 'Set-PbiProtection：多個欄位各一筆變更' ($script:Calls[$script:Calls.Count - 1].Json.Changes.Count -eq 2) ""
Expect-Throw Set-PbiProtection '不認得的等級 → 本機先擋下' { Set-PbiProtection -Table T -Column c -Level readable }

# ── 資料表剖析（改 M 前後的量化比對）─────────────────────────────────────────
$script:Covered['Get-PbiTableProfile']     = $true
$script:Covered['Compare-PbiTableProfile'] = $true
$script:Canned['/api/query'] = [PSCustomObject]@{ Rows = @([PSCustomObject]@{
    '[n]' = 100; '[b0]' = 0; '[d0]' = 7; '[b1]' = 2; '[d1]' = 40; '[s1]' = 1234.5 }) }
$before = Get-PbiTableProfile 'Dim 產品' -Columns 類別, 單價
$q = $script:Calls[$script:Calls.Count - 1].Body.Query
Check 'Get-PbiTableProfile：產生的 DAX 把表名加上單引號、只對數值欄位加總' (
    $q.Contains("COUNTROWS('Dim 產品')") -and $q.Contains("COUNTBLANK('Dim 產品'[類別])") -and
    $q.Contains("SUM('Dim 產品'[單價])") -and -not $q.Contains("SUM('Dim 產品'[類別])")) "實際查詢：$q"
Check 'Get-PbiTableProfile：結果攤成屬性，別名對回欄名' (
    $before.列數 -eq 100 -and $before.'相異_類別' -eq 7 -and $before.'總和_單價' -eq 1234.5 -and $before.欄位數 -eq 2) ""
# 別名不能帶欄名：欄名符合保護樣式（amount、customer…）時，帶欄名的別名會被當成那個受限欄位擋下來
Check 'Get-PbiTableProfile：查詢裡的別名不含欄名' (-not ($q -match '"[^"]*(類別|單價)[^"]*"\s*,')) "實際查詢：$q"
# 這一項可以讀 schema（要靠它才知道欄位存不存在），但不能有查詢送出去
$queriesBefore = @($script:Calls | Where-Object { $_.Path -eq '/api/query' }).Count
$threw = $false
try { $null = Get-PbiTableProfile 'Dim 產品' -Columns 不存在 } catch { $threw = $true }
Check 'Get-PbiTableProfile：欄位不存在 → 丟出例外，而且沒有查詢送出去' (
    $threw -and @($script:Calls | Where-Object { $_.Path -eq '/api/query' }).Count -eq $queriesBefore) ""
$script:Canned['/api/query'] = [PSCustomObject]@{ Rows = @([PSCustomObject]@{
    '[n]' = 70; '[b0]' = 0; '[d0]' = 7; '[b1]' = 2; '[d1]' = 40; '[s1]' = 1234.5 }) }
$after = Get-PbiTableProfile 'Dim 產品' -Columns 類別, 單價
$diff  = @(Compare-PbiTableProfile $before $after 6>$null)
Check 'Compare-PbiTableProfile：只列出有變動的指標，並算出變化比例' (
    $diff.Count -eq 1 -and $diff[0].項目 -eq '列數' -and $diff[0].變化 -eq '-30.0%') "實際：$($diff | ForEach-Object { "$($_.項目) $($_.變化)" })"
Check 'Compare-PbiTableProfile：沒有變動時不輸出任何列' (@(Compare-PbiTableProfile $before $before 6>$null).Count -eq 0) ""
$script:Canned.Remove('/api/query')

# ── 存檔 / 重新整理 ─────────────────────────────────────────────────────────
Expect-Call Save-PbiModel '預設：最多等 30 秒，不帶 Expect / VerifyOnly' { Save-PbiModel } '/api/save' `
    -Has @{ WaitSeconds = 30 } -Lacks Expect, VerifyOnly
Expect-Call Save-PbiModel '-Expect 只給一個值時，送出去的仍是陣列' { Save-PbiModel -Expect '銷售總額' } '/api/save' `
    -Has @{ Expect = @('銷售總額') }
Expect-Call Save-PbiModel '-Expect 多個值' { Save-PbiModel -Expect '甲', '乙' -WaitSeconds 60 } '/api/save' `
    -Has @{ Expect = @('甲', '乙'); WaitSeconds = 60 }
Expect-Call Save-PbiModel '-VerifyOnly' { Save-PbiModel -VerifyOnly -Expect 'x' } '/api/save' -Has @{ VerifyOnly = $true }
Expect-Call Invoke-PbiRefresh '預設 full、不指定表' { Invoke-PbiRefresh } '/api/refresh' -Has @{ RefreshType = 'full' } -Lacks TableName
Expect-Call Invoke-PbiRefresh '指定表與類型' { Invoke-PbiRefresh -Table '日期表' -RefreshType calculate } '/api/refresh' `
    -Has @{ RefreshType = 'calculate'; TableName = '日期表' }

# ── 快照 ────────────────────────────────────────────────────────────────────
Expect-Call New-PbiSnapshot '標籤' { New-PbiSnapshot -Label '改版前' } '/api/snapshot' -Has @{ Label = '改版前' }
Expect-Call Restore-PbiSnapshot '預設不是 DryRun、不送 Scope' { Restore-PbiSnapshot 'a.json' } '/api/restore' `
    -Has @{ File = 'a.json'; DryRun = $false } -Lacks Scope
Expect-Call Restore-PbiSnapshot '-Scope 只給一個值時仍是陣列' { Restore-PbiSnapshot 'a.json' -DryRun -Scope measures } '/api/restore' `
    -Has @{ DryRun = $true; Scope = @('measures') }

# ── 量值 ────────────────────────────────────────────────────────────────────
Expect-Call Set-PbiMeasure '只送出有指定的屬性' {
    Set-PbiMeasure -Table '量值' -Name '銷售總額' -Expression 'SUM(F[Amount])' -Format '#,0'
} '/api/upsert-measure' -Has @{ TableName = '量值'; MeasureName = '銷售總額'; Expression = 'SUM(F[Amount])'; FormatString = '#,0' } `
  -Lacks Description, DisplayFolder, IsHidden
Expect-Call Set-PbiMeasure '-IsHidden $false 要送出去（明確取消隱藏）' {
    Set-PbiMeasure -Table '量值' -Name 'x' -IsHidden $false
} '/api/upsert-measure' -Has @{ IsHidden = $false } -Lacks Expression
Expect-Call Remove-PbiMeasure '端點' { Remove-PbiMeasure -Table '量值' -Name 'x' } '/api/delete-measure' -Has @{ TableName = '量值'; MeasureName = 'x' }
Expect-Call Move-PbiMeasure '端點' { Move-PbiMeasure -FromTable A -Name x -ToTable B } '/api/move-measure' -Has @{ FromTable = 'A'; MeasureName = 'x'; ToTable = 'B' }

# ── 資料行 ──────────────────────────────────────────────────────────────────
Expect-Call Add-PbiColumn 'DAX 計算資料行' { Add-PbiColumn -Table T -Name 毛利 -Expression 'T[a]-T[b]' -DataType decimal } '/api/add-column' `
    -Has @{ TableName = 'T'; ColumnName = '毛利'; Expression = 'T[a]-T[b]'; DataType = 'decimal' } -Lacks SourceColumn
Expect-Call Add-PbiColumn '來源資料行（M 輸出的欄位）' { Add-PbiColumn -Table T -Name region -SourceColumn region } '/api/add-column' `
    -Has @{ ColumnName = 'region'; SourceColumn = 'region'; DataType = 'text' } -Lacks Expression
Expect-Throw Add-PbiColumn '兩種都沒給 → 本機先擋下' { Add-PbiColumn -Table T -Name x }
Expect-Throw Add-PbiColumn '兩種都給 → 本機先擋下'   { Add-PbiColumn -Table T -Name x -Expression '1' -SourceColumn y }
Expect-Call Remove-PbiColumn '端點' { Remove-PbiColumn -Table T -Name x } '/api/delete-column' -Has @{ TableName = 'T'; ColumnName = 'x' }
Expect-Call Set-PbiColumn '只送出有指定的屬性' { Set-PbiColumn -Table T -Name x -Format '0%' -IsHidden $true } '/api/set-column-props' `
    -Has @{ FormatString = '0%'; IsHidden = $true } -Lacks SortByColumn, SummarizeBy, DataType, DisplayFolder
Expect-Call Set-PbiColumn '-SortByColumn 空字串要送出去（代表清除排序依據）' { Set-PbiColumn -Table T -Name x -SortByColumn '' } '/api/set-column-props' `
    -Has @{ SortByColumn = '' }

# ── Power Query M ───────────────────────────────────────────────────────────
$goodM = "let`r`n    Source = Table.FromRows({{1}}, {`"a`"}),`r`n    Kept = Table.SelectRows(Source, each [a] > 0)`r`nin`r`n    Kept"
Expect-Call Set-PbiMQuery '完整的 let … in 會送出去，內容原樣不變' { Set-PbiMQuery '訂單' -Expression $goodM } '/api/update-m' `
    -Has @{ TableName = '訂單'; Expression = $goodM }
Expect-Call Set-PbiMQuery '開頭有空白行與縮排也接受' { Set-PbiMQuery '訂單' -Expression ("`r`n  " + $goodM) } '/api/update-m'
Expect-Call Set-PbiMQuery '單行寫法也接受' { Set-PbiMQuery '訂單' -Expression 'let Source = 1 in Source' } '/api/update-m'
Expect-Throw Set-PbiMQuery '片段（沒有 let … in）→ 本機先擋下' { Set-PbiMQuery '訂單' -Expression 'Table.SelectRows(Source, each [a] > 0)' }
Expect-Throw Set-PbiMQuery '只有 let 沒有 in → 本機先擋下'      { Set-PbiMQuery '訂單' -Expression 'let Source = 1' }
Expect-Throw Set-PbiMQuery '只是字裡含有 let / in 的文字 → 本機先擋下' { Set-PbiMQuery '訂單' -Expression 'letter = "x" inside' }
Expect-Call Set-PbiExpression '共用運算式' { Set-PbiExpression -Name '起始日' -Expression '#date(2026,1,1)' } '/api/upsert-expression' `
    -Has @{ Name = '起始日'; Expression = '#date(2026,1,1)' } -Lacks Description
Expect-Call Remove-PbiExpression '端點' { Remove-PbiExpression -Name '起始日' } '/api/delete-expression' -Has @{ Name = '起始日' }

# ── 關聯 / 表格 / 改名 ──────────────────────────────────────────────────────
Expect-Call Set-PbiRelationship '預設 many → one、單向、啟用' {
    Set-PbiRelationship -FromTable Sales -FromColumn DateKey -ToTable '日期' -ToColumn DateKey
} '/api/upsert-relationship' -Has @{ FromCardinality = 'many'; ToCardinality = 'one'; CrossFilterDirection = 'single'; IsActive = $true; ToTable = '日期' }
Expect-Call Remove-PbiRelationship '端點' {
    Remove-PbiRelationship -FromTable Sales -FromColumn DateKey -ToTable '日期' -ToColumn DateKey
} '/api/delete-relationship' -Has @{ FromTable = 'Sales'; ToColumn = 'DateKey' }
Expect-Call New-PbiTable '計算表' { New-PbiTable -Name '日期表' -Expression 'CALENDARAUTO()' } '/api/create-table' `
    -Has @{ TableName = '日期表'; Kind = 'calculated'; Expression = 'CALENDARAUTO()'; IsHidden = $false } -Lacks Columns
$script:Covered['New-PbiTable'] = $true
$null = New-PbiTable -Name 'T' -Kind m -Expression 'let s = 1 in s' -Columns @(@{ Name = 'a'; DataType = 'int' })
$cols = $script:Calls[$script:Calls.Count - 1].Json.Columns
Check 'New-PbiTable：-Columns 只給一欄時，送出去的仍是陣列' (($cols -is [array]) -and $cols.Count -eq 1 -and $cols[0].Name -eq 'a') ""
Expect-Call Remove-PbiTable '端點' { Remove-PbiTable -Name 'T' } '/api/delete-table' -Has @{ TableName = 'T' }
Expect-Call Rename-PbiObject '-DryRun' { Rename-PbiObject -Type measure -Table '量值' -OldName a -NewName b -DryRun } '/api/rename' `
    -Has @{ ObjectType = 'measure'; TableName = '量值'; OldName = 'a'; NewName = 'b'; DryRun = $true }
Expect-Call Rename-PbiObject '預設不是 DryRun' { Rename-PbiObject -Type table -OldName a -NewName b } '/api/rename' -Has @{ DryRun = $false }
Expect-Throw Rename-PbiObject '改量值／資料行卻沒給 -Table → 本機先擋下' { Rename-PbiObject -Type measure -OldName a -NewName b }

# ── 模型屬性 / 計算群組 / 角色 ──────────────────────────────────────────────
Expect-Call  Set-PbiModelProps '端點' { Set-PbiModelProps -DiscourageImplicitMeasures $true } '/api/set-model-props' `
    -Has @{ DiscourageImplicitMeasures = $true } -Lacks Description
Expect-Throw Set-PbiModelProps '什麼都沒指定 → 本機先擋下' { Set-PbiModelProps }
Expect-Call New-PbiCalcGroup '預設不開 DiscourageImplicitMeasures' { New-PbiCalcGroup -Name '時間智慧' } '/api/upsert-calc-group' `
    -Has @{ TableName = '時間智慧'; ColumnName = '計算項目'; DiscourageImplicitMeasures = $false }
Expect-Call Set-PbiCalcItem '只送出有指定的屬性' { Set-PbiCalcItem -Group '時間智慧' -Name YTD -Expression 'CALCULATE(SELECTEDMEASURE(), DATESYTD(D[Date]))' } `
    '/api/upsert-calc-item' -Has @{ TableName = '時間智慧'; ItemName = 'YTD' } -Lacks Ordinal, FormatStringExpression
Expect-Call Remove-PbiCalcItem '端點' { Remove-PbiCalcItem -Group '時間智慧' -Name YTD } '/api/delete-calc-item' -Has @{ ItemName = 'YTD' }
$script:Covered['Set-PbiRole'] = $true
$null = Set-PbiRole -Name '業務員' -TablePermissions @(@{ TableName = 'F'; FilterExpression = '[Owner] = USERPRINCIPALNAME()' })
$role = $script:Calls[$script:Calls.Count - 1]
Check 'Set-PbiRole：端點、預設權限 read、TablePermissions 只給一條時仍是陣列' (
    $role.Path -eq '/api/upsert-role' -and $role.Json.ModelPermission -eq 'read' -and
    ($role.Json.TablePermissions -is [array]) -and $role.Json.TablePermissions[0].TableName -eq 'F') ""
Expect-Call Remove-PbiRole '端點' { Remove-PbiRole -Name '業務員' } '/api/delete-role' -Has @{ RoleName = '業務員' }

# ── 批次 ────────────────────────────────────────────────────────────────────
$script:Covered['Invoke-PbiBatch'] = $true
$null = Invoke-PbiBatch @(
    @{ Op = 'upsert-measure'; Args = @{ TableName = '量值'; MeasureName = '銷售總額'; Expression = "SUM('訂單'[金額])" } }
)
$batch = $script:Calls[$script:Calls.Count - 1]
Check 'Invoke-PbiBatch：只有一個操作時 Operations 仍是陣列，巢狀的中文參數完整保留' (
    $batch.Path -eq '/api/batch' -and ($batch.Json.Operations -is [array]) -and
    $batch.Json.Operations[0].Op -eq 'upsert-measure' -and
    $batch.Json.Operations[0].Args.MeasureName -ceq '銷售總額' -and
    $batch.Json.Operations[0].Args.Expression -ceq "SUM('訂單'[金額])") ""
Check 'Invoke-PbiBatch：預設任一步失敗就整批不套用（StopOnError = true、不逐步存檔、不是 DryRun）' (
    $batch.Json.StopOnError -eq $true -and $batch.Json.SavePerOp -eq $false -and $batch.Json.DryRun -eq $false) ""

# =============================================================================
# ③ 涵蓋率：PBI-Bridge.ps1 裡的每個函式都要有人管
# =============================================================================
# 新增函式卻沒在這裡加檢查的話，下面這項會失敗 —— 這份自我檢查才不會隨著時間變成擺設。
$notOffline = @{
    'Get-PbiRootPath' = '純路徑計算，載入本檔時已經用到'
    'Get-PbiApiKey'   = '讀真的金鑰檔；自我檢查不碰金鑰'
    'Invoke-PbiApi'   = '已在 ②-a 用真的函式測過'
    'Test-PbiBridge'  = '直接連服務的 /ping，需要服務在跑'
    'Get-PbiMQuery'   = '會把 M 腳本寫進 PowerQuery_Scripts\，自我檢查不產生檔案'
    'Get-PbiHelp'     = '只印指令清單'
}
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'PBI-Bridge.ps1'), [ref]$tokens, [ref]$errors)
$defined = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) |
             ForEach-Object { $_.Name })
$unchecked = @($defined | Where-Object { -not $script:Covered.ContainsKey($_) -and -not $notOffline.ContainsKey($_) })
Check "PBI-Bridge.ps1 的 $($defined.Count) 個函式都有對應的檢查" ($unchecked.Count -eq 0) "沒有檢查的函式：$($unchecked -join '、')"

# =============================================================================
Write-Host ""
$failed = $script:Failures.Count
Write-Host ("結果：{0} 通過 / {1} 失敗" -f $script:Pass, $failed) -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Red' })
if ($failed -gt 0) { exit 1 }
