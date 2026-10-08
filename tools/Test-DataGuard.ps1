# =============================================================================
# Test-DataGuard.ps1 — 對資料保護機制做紅隊測試
# =============================================================================
# 用法：
#   . ".\tools\PBI-Bridge.ps1"
#   .\tools\Test-DataGuard.ps1                      # 唯讀測試
#   .\tools\Test-DataGuard.ps1 -IncludeWriteTests   # 另做「建衍生物件洗資料」測試：
#                                                   # 會暫時建立幾個量值與一張計算表，測完刪除。
#                                                   # 不會存檔；跑之前會先存一份快照。
#
# 安全設計（重要）：
#   1. 每個攻擊查詢都設 MaxRows = 1。萬一真有漏洞，外洩的是一個值而不是整批。
#   2. 腳本**只回報有沒有被擋下，絕不印出回傳內容**。防護的測試本身不該變成外洩管道。
#   3. 受限的欄位直接問伺服器（Get-PbiProtection），不寫死欄位名 —— 換模型也能用。
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

# 只在寫入型測試用到：替暫時的量值取一個「符合通用規則樣式」的名字時，要照伺服器的方式比對
# （Program.cs 的 DataGuard.Norm / Matches：忽略大小寫、空白、底線、連字號；清單項目可以是 * 樣式）。
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

# 從模型裡挑「實際受限」的欄位來測。等級直接問伺服器（Get-PbiProtection）：
# 它已經把逐欄設定與通用規則合在一起算好了，在這裡重寫一遍比對邏輯只會和伺服器走岔。
$schema = Get-PbiSchema
$prot   = Get-PbiProtection -Raw
if (-not $prot.enabled) { throw "資料保護目前是關閉的，沒有東西可以測" }
if ($prot.problem)      { throw $prot.problem }
$levelOf = @{}
foreach ($t in $prot.tables) { foreach ($c in $t.columns) { $levelOf["$($t.name)`t$($c.name)"] = $c } }
# 上限與通用規則的樣式也問伺服器，不去讀 appsettings.json —— 服務實際用的是啟動時載入的那一份，問它才準
$maxMoneyRows  = [int]$prot.limits.maxRowsWithMoney
$denyPatterns  = @($prot.tables | ForEach-Object { $_.columns } | Where-Object { $_.source -eq 'pattern' -and $_.level -eq 'countOnly' }     | ForEach-Object { $_.rule } | Sort-Object -Unique)
$moneyPatterns = @($prot.tables | ForEach-Object { $_.columns } | Where-Object { $_.source -eq 'pattern' -and $_.level -eq 'aggregateOnly' } | ForEach-Object { $_.rule } | Sort-Object -Unique)

# 「只能計數」的欄位優先挑文字型：MAX / CONCATENATEX 取到的才會是真實名稱，最貼近實際風險
$pick = @(foreach ($t in $prot.tables) { foreach ($c in $t.columns) {
    if ($c.level -eq 'countOnly') { [PSCustomObject]@{ T = $t.name; C = $c.name; IsText = ($c.dataType -eq 'String') } }
} }) | Sort-Object { -not $_.IsText } | Select-Object -First 1
if (-not $pick) { throw "這個模型裡沒有任何欄位是「只能計數」，無從測起。請到儀表板的「資料保護」分頁至少設定一欄，或確認 appsettings.json 的清單涵蓋你的欄位命名。" }

# 「只能彙總」一定是數值欄位（伺服器不會把文字欄位歸到這一級）
$money = @(foreach ($t in $prot.tables) { foreach ($c in $t.columns) {
    if ($c.level -eq 'aggregateOnly') { [PSCustomObject]@{ T = $t.name; C = $c.name } }
} }) | Select-Object -First 1

$D = "'$($pick.T)'[$($pick.C)]"
Write-Host "🎯 測試對象" -ForegroundColor Cyan
Write-Host "   只能計數的欄位：  $D"
if ($money) { $M = "'$($money.T)'[$($money.C)]"; Write-Host "   只能彙總的欄位：$M" }
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

# 不必指名欄位就能把內容帶出來的函式 —— 查詢文字裡沒有任何中括號參考，結果欄位也只有別名，
# 所以欄位層級的兩道檢查都看不到，只能整個禁用
$cases += @{ N='TOCSV 把整張表變成一個字串';    Block=$true; Q="EVALUATE ROW(`"x`", TOCSV('$($pick.T)'))" }
$cases += @{ N='TOJSON 把整張表變成一個字串';   Block=$true; Q="EVALUATE ROW(`"x`", TOJSON('$($pick.T)'))" }
$cases += @{ N='COLUMNSTATISTICS 取每欄極值';  Block=$true; Q="EVALUATE COLUMNSTATISTICS()" }
$cases += @{ N='INFO 函式讀 M 腳本';           Block=$true; Q="EVALUATE INFO.PARTITIONS()" }

# UNION 的欄名取自第一個引數：拿一列假資料當第一個引數，整張表就掛在別名底下流出去
$colCount = @(($schema.Tables | Where-Object { $_.Name -eq $pick.T }).Columns).Count
$dummy    = (1..$colCount | ForEach-Object { "`"c$_`", 1" }) -join ', '
$cases += @{ N='UNION 把整張表掛在別名底下';   Block=$true;  Q="EVALUATE UNION(ROW($dummy), TOPN(1, '$($pick.T)'))" }
$cases += @{ N='變數裝著整張表再進 UNION';     Block=$true;  Q="EVALUATE VAR t = TOPN(1, '$($pick.T)') RETURN UNION(ROW($dummy), t)" }
$cases += @{ N='✅ 合法：UNION 幾個 ROW';      Block=$false; Q="EVALUATE UNION(ROW(`"k`", `"a`", `"v`", 1), ROW(`"k`", `"b`", `"v`", 2))" }

if ($money) {
    $cases += @{ N='金額逐列取值';                 Block=$true;  Q="EVALUATE SELECTCOLUMNS('$($money.T)', `"a`", $M)" }
    $cases += @{ N='DIVIDE 不算聚合（逐列取金額）'; Block=$true;  Q="EVALUATE SELECTCOLUMNS('$($money.T)', `"v`", DIVIDE($M, 1))" }
    $cases += @{ N='✅ 合法：金額聚合';            Block=$false; Q="EVALUATE ROW(`"s`", SUM($M))" }
    $cases += @{ N='✅ 合法：SUMX 逐列運算後加總';  Block=$false; Q="EVALUATE ROW(`"s`", SUMX('$($money.T)', $M))" }
    # 篩選條件只決定留下哪些列，本身不會變成輸出 —— 模型裡既有的量值常常這樣寫
    $cases += @{ N='✅ 合法：金額寫在篩選條件裡';   Block=$false; Q="EVALUATE ROW(`"s`", CALCULATE(SUM($M), $M > 0))" }
}

# ERROR() 會把任意文字放進引擎的錯誤訊息，所以只要查詢碰到受限欄位就整個不准用。
# 「該擋」的那一半只在寫入型測試裡、拿測試自己造的資料測（步驟 7）—— 拿真實欄位測的話，
# 萬一規則哪天失效，錯誤訊息就是一個真實的值，而這支腳本會把非管制的錯誤訊息印出來。
$cases += @{ N='✅ 合法：ERROR 沒碰受限欄位';    Block=$false; Q="EVALUATE ROW(`"x`", IF(1 = 2, ERROR(`"no`"), 1))" }

$pass = 0; $fail = 0
foreach ($c in $cases) {
    $blocked = $false; $err = $null
    try {
        # MaxRows = 1：萬一沒被擋下，外洩的是一個值而不是整批。
        $null = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{ Query = $c.Q; MaxRows = 1 }
    } catch {
        $err = "$_"
        if ($err -match '403|資料保護') { $blocked = $true }
    }

    # 「該放行」的案例必須真的查成功：被別的錯誤（DAX 語法、物件不存在）擋住不算放行
    $ok = if ($c.Block) { $blocked } else { -not $err }
    if ($ok) { $pass++ } else { $fail++ }

    $verdict = if ($ok) { '✅ 通過' } else { '❌ 失敗' }
    $actual  = if ($blocked) { '被擋下' } elseif ($err) { '查詢錯誤' } else { '通過' }
    Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
                $verdict, $c.N, $(if ($c.Block) { '擋下' } else { '放行' }), $actual) `
               -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    # 不是 403 的錯誤要顯示出來，否則 DAX 語法錯誤會被誤判成「防護生效」
    if ($err -and -not $blocked) { Write-Host "      （非管制錯誤：$($err -replace '\s+',' ' -replace '^(.{140}).*','$1…')）" -ForegroundColor DarkGray }
}

# ── 金額查詢的列數上限（後備防線）─────────────────────────────────────────
# 金額有被正確聚合（過得了查詢文字那一關），但分組鍵基數太高 —— 那已接近逐筆金額。
# 注意：必須用比上限大的 MaxRows 送，否則結果會先被截斷到 MaxRows，
#       rows.Count 永遠不會超過上限，這道檢查就測不到。
if ($money) {
    $mt = $money.T
    # 動態找一個「非管制、且相異值 > 上限」的欄位當分組鍵，讓測試不綁死特定模型
    $groupCol = $null
    foreach ($c in ($schema.Tables | Where-Object { $_.Name -eq $mt }).Columns) {
        if ($levelOf["$mt`t$($c.Name)"].level -ne 'open') { continue }      # 分組鍵要用開放的欄位
        if ($c.DataType -ne 'String') { continue }
        try {
            $n = (Invoke-Dax "EVALUATE ROW(`"n`", DISTINCTCOUNT('$mt'[$($c.Name)]))").Rows[0].'[n]'
            if ($n -gt $maxMoneyRows) { $groupCol = $c.Name; break }
        } catch { }
    }
    if ($groupCol) {
        $blocked = $false
        try {
            $null = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{
                Query = "EVALUATE SUMMARIZECOLUMNS('$mt'[$groupCol], `"s`", SUM($M))"
                MaxRows = ($maxMoneyRows * 10) }
        } catch { if ("$_" -match '403|資料保護') { $blocked = $true } }
        if ($blocked) { $pass++ } else { $fail++ }
        Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
                    $(if ($blocked) { '✅ 通過' } else { '❌ 失敗' }),
                    "金額按高基數鍵分組($groupCol)", '擋下',
                    $(if ($blocked) { '被擋下' } else { '通過' })) `
                   -ForegroundColor $(if ($blocked) { 'Green' } else { 'Red' })
    } else {
        Write-Host "⚠️ 略過  金額列數上限測試 —— 這個模型找不到相異值 > $($maxMoneyRows) 的非管制文字欄位" -ForegroundColor Yellow
    }
}

# ── DMV 路徑 ───────────────────────────────────────────────────────────────
# 系統檢視是白名單制：只開放模型的中繼資料與儲存統計。
# 該放行的那一項要「真的查成功」才算通過（被別的錯誤擋住不算）。
foreach ($d in @(
    @{ N = 'DMV 拉 M 腳本定義';         Block = $true;  Q = 'SELECT * FROM $SYSTEM.TMSCHEMA_PARTITIONS' },
    @{ N = 'DMV 列出欄位的所有相異值';   Block = $true;  Q = 'SELECT * FROM $SYSTEM.MDSCHEMA_MEMBERS' },
    @{ N = 'DMV 讀別的連線跑過的查詢';   Block = $true;  Q = 'SELECT * FROM $SYSTEM.DISCOVER_SESSIONS' },
    @{ N = '✅ 合法：DMV 讀量值清單';     Block = $false; Q = 'SELECT * FROM $SYSTEM.TMSCHEMA_MEASURES' },
    @{ N = '✅ 合法：DMV 讀儲存統計';     Block = $false; Q = 'SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLES' })) {
    $blocked = $false; $err = $null
    try { $null = Invoke-PbiApi -Path "/api/dmv" -Method POST -Body @{ Query = $d.Q; MaxRows = 1 } }
    catch { $err = "$_"; if ($err -match '403|資料保護') { $blocked = $true } }
    $ok = if ($d.Block) { $blocked } else { -not $err }
    if ($ok) { $pass++ } else { $fail++ }
    Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
                $(if ($ok) { '✅ 通過' } else { '❌ 失敗' }), $d.N, $(if ($d.Block) { '擋下' } else { '放行' }),
                $(if ($blocked) { '被擋下' } elseif ($err) { '查詢錯誤' } else { '通過' })) `
               -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    if ($err -and -not $blocked) { Write-Host "      （非管制錯誤：$($err -replace '\s+',' ' -replace '^(.{140}).*','$1…')）" -ForegroundColor DarkGray }
}

# ── 寫入型測試：建衍生物件把資料洗到未管制的名稱下 ──────────────────────────
if ($IncludeWriteTests) {
    Write-Host "`n🧪 寫入型測試（會在模型中建立暫時的量值與一張計算表，測完刪除；過程不會存檔）" -ForegroundColor Yellow
    New-PbiSnapshot -Label "before-dataguard-test" | Out-Null

    function Invoke-GuardCase([string]$Name, [bool]$ExpectBlock, [string]$Query) {
        $blocked = $false; $err = $null
        try { $null = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{ Query = $Query; MaxRows = 1 } }
        catch { $err = "$_"; if ($err -match '403|資料保護') { $blocked = $true } }
        # 「該放行」的案例必須真的查成功：被別的錯誤（DAX 語法、物件不存在）擋住不算放行
        $ok = if ($ExpectBlock) { $blocked } else { -not $err }
        if ($ok) { $script:pass++ } else { $script:fail++ }
        $actual = if ($blocked) { '被擋下' } elseif ($err) { '查詢錯誤' } else { '通過' }
        Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f
                    $(if ($ok) { '✅ 通過' } else { '❌ 失敗' }), $Name,
                    $(if ($ExpectBlock) { '擋下' } else { '放行' }), $actual) `
                   -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
        if ($err -and -not $blocked) {
            Write-Host "      （非管制錯誤：$($err -replace '\s+',' ' -replace '^(.{140}).*','$1…')）" -ForegroundColor DarkGray
        }
    }

    # 量值名稱若符合管制樣式（[Total Revenue]、[Customer Rank]…）以前會被當成欄位擋掉，
    # 連 ROW("v", [量值]) 都驗算不了。現在量值參考會放行、改由它的定義決定 —— 下面同時測
    # 「該放行的放行了」與「放行沒有變成漏洞」。
    function New-ProbeName($Patterns, [string]$Suffix) {
        foreach ($p in @($Patterns)) {
            if (-not $p -or -not $p.Contains('*')) { continue }
            $core = ($p -replace '\*', '') -replace '[^A-Za-z0-9一-鿿]', ''
            if ($core) { return "__guard_${core}_${Suffix}__" }
        }
        return $null
    }
    $columnNames = @{}; $measureNames = @{}
    foreach ($t in $schema.Tables) {
        foreach ($c in $t.Columns)  { $columnNames[(Get-NormColumnName $c.Name)] = $true }
        foreach ($m in $t.Measures) { $measureNames[$m.Name.ToLowerInvariant()]  = $true }
    }

    $tmp       = "__guard_probe__"
    $tmpTable  = "__guard_probe_table__"
    $moneyName = New-ProbeName $moneyPatterns 'probe'
    $denyName  = New-ProbeName $denyPatterns 'probe'
    $cleanup   = New-Object System.Collections.ArrayList      # 測完要刪的東西
    $settingsToClear = $false
    $countExpr = "COUNTROWS('$($pick.T)')"

    # 暫時物件一律用模型裡沒有的名字。Set-PbiMeasure 遇到同名量值是「覆寫」，測完又會刪掉 ——
    # 撞名的話等於把使用者原有的量值弄丟，所以寧可不測。
    foreach ($n in @($tmp, $moneyName, $denyName) | Where-Object { $_ }) {
        if ($measureNames.ContainsKey($n.ToLowerInvariant())) {
            throw "模型裡已經有量值 [$n]（可能是上次測試中斷留下的）。請先確認並刪除它，再重跑寫入型測試。"
        }
    }
    if ($schema.Tables | Where-Object { $_.Name -eq $tmpTable }) {
        throw "模型裡已經有資料表 '$tmpTable'（可能是上次測試中斷留下的）。請先刪除它，再重跑寫入型測試。"
    }

    try {
        # ── 1. 新建量值把敏感欄位包起來 ───────────────────────────────────────
        Set-PbiMeasure -Table $pick.T -Name $tmp -Expression "MAX($D)" | Out-Null
        [void]$cleanup.Add(@{ Kind = 'measure'; Table = $pick.T; Name = $tmp })
        Invoke-GuardCase '新建量值包裝敏感欄位' $true "EVALUATE ROW(`"x`", [$tmp])"

        # ── 2. 量值名稱像金額欄位 ─────────────────────────────────────────────
        if ($moneyName -and (Test-ColumnListed $moneyName $moneyPatterns) -and -not $columnNames.ContainsKey((Get-NormColumnName $moneyName))) {
            Set-PbiMeasure -Table $pick.T -Name $moneyName -Expression $countExpr | Out-Null
            [void]$cleanup.Add(@{ Kind = 'measure'; Table = $pick.T; Name = $moneyName })
            Invoke-GuardCase '✅ 合法：量值名稱像金額欄位' $false "EVALUATE ROW(`"v`", [$moneyName])"

            # 同一個名字，定義改成取出身分欄位 → 名稱被放行不代表定義也被放行
            Set-PbiMeasure -Table $pick.T -Name $moneyName -Expression "MAX($D)" | Out-Null
            Invoke-GuardCase '量值名稱像金額、定義取身分欄位' $true "EVALUATE ROW(`"v`", [$moneyName])"
        } else {
            Write-Host "⚠️ 略過  量值名稱像金額欄位 —— 這個模型沒有欄位是靠通用規則的樣式變成「只能彙總」的" -ForegroundColor Yellow
        }

        # ── 3. 量值名稱像身分欄位 ─────────────────────────────────────────────
        if ($denyName -and (Test-ColumnListed $denyName $denyPatterns) -and -not $columnNames.ContainsKey((Get-NormColumnName $denyName))) {
            Set-PbiMeasure -Table $pick.T -Name $denyName -Expression $countExpr | Out-Null
            [void]$cleanup.Add(@{ Kind = 'measure'; Table = $pick.T; Name = $denyName })
            Invoke-GuardCase '✅ 合法：量值名稱像身分欄位' $false "EVALUATE ROW(`"v`", [$denyName])"
        } else {
            Write-Host "⚠️ 略過  量值名稱像身分欄位 —— 這個模型沒有欄位是靠通用規則的樣式變成「只能計數」的" -ForegroundColor Yellow
        }

        # ── 4. 量值與金額「資料行」同名 ───────────────────────────────────────
        # 量值可以和別張表的資料行同名。放行量值參考時如果只看「有這個量值」，
        # 建一個同名量值就能讓那個金額欄位逐列流出去 —— 這裡確認資料行照樣受管制。
        if ($money) {
            $holder = $schema.Tables | Where-Object {
                $_.Name -ne $money.T -and -not $_.IsCalculationGroup -and
                -not ($_.Columns | Where-Object { $_.Name -eq $money.C })
            } | Select-Object -First 1
            $collided = $false
            # 模型裡已經有同名量值就不測（理由同上：不能覆寫再刪掉使用者的量值）
            if ($holder -and -not $measureNames.ContainsKey($money.C.ToLowerInvariant())) {
                try {
                    Set-PbiMeasure -Table $holder.Name -Name $money.C -Expression "1" | Out-Null
                    [void]$cleanup.Add(@{ Kind = 'measure'; Table = $holder.Name; Name = $money.C })
                    $collided = $true
                } catch { }
            }
            if ($collided) {
                Invoke-GuardCase '量值與金額欄位同名，欄位照樣受管制' $true "EVALUATE SELECTCOLUMNS('$($money.T)', `"a`", $M)"
            } else {
                Write-Host "⚠️ 略過  量值與金額欄位同名 —— 這個模型建不出同名量值（引擎不允許，或找不到可放的表）" -ForegroundColor Yellow
            }
        }

        # ── 5. 同名的計算資料行不能蓋掉量值的定義 ─────────────────────────────
        # 第 1 步的量值 [__guard_probe__] 還在。另一張表建一個同名、內容無害的計算資料行：
        # 展開定義時如果一個名稱只留一份，後讀到的「1」會蓋掉量值的 MAX(身分欄位)，掃描就看不到它。
        # 這張暫時的表內容是測試自己造的（alpha / beta / gamma），第 7 步拿它測四個等級。
        New-PbiTable -Name $tmpTable -Expression 'DATATABLE("probe_pid", STRING, "probe_grp", STRING, "probe_val", INTEGER, {{"alpha", "g1", 10}, {"beta", "g1", 20}, {"gamma", "g2", 30}})' | Out-Null
        [void]$cleanup.Add(@{ Kind = 'table'; Name = $tmpTable })
        Add-PbiColumn -Table $tmpTable -Name $tmp -Expression '1' -DataType int | Out-Null
        Invoke-GuardCase '同名計算資料行蓋不掉量值的定義' $true "EVALUATE ROW(`"x`", [$tmp])"

        # ── 6. 把取值的量值藏在一長串量值後面 ─────────────────────────────────
        # 以前定義只展開 6 層：[L0]=[L1]、[L1]=[L2] … 真正取值的放在更後面，掃描就看不到。
        $chain = @(0..8 | ForEach-Object { "__guard_chain_${_}__" })
        if (-not ($chain | Where-Object { $measureNames.ContainsKey($_.ToLowerInvariant()) })) {
            $ops = for ($i = 0; $i -lt $chain.Count; $i++) {
                $expr = if ($i -lt $chain.Count - 1) { "[$($chain[$i + 1])]" } else { "MAX($D)" }
                @{ Op = 'upsert-measure'; Args = @{ TableName = $pick.T; MeasureName = $chain[$i]; Expression = $expr } }
            }
            Invoke-PbiBatch $ops | Out-Null
            foreach ($n in $chain) { [void]$cleanup.Add(@{ Kind = 'measure'; Table = $pick.T; Name = $n }) }
            Invoke-GuardCase '量值串了九層才取值' $true "EVALUATE ROW(`"v`", [$($chain[0])])"
        }

        # ── 7. 四個等級：各自該擋的與該放行的 ─────────────────────────────────
        # 只做「收緊」（開放 → 受限），所以不會跳確認視窗。放寬要使用者按確認，那一段不在自動測試裡。
        $T = "'$tmpTable'"
        Invoke-PbiRefresh -Table $tmpTable | Out-Null             # 計算表要算過才查得到
        # 下面兩個設定都必須是「收緊」（開放 → 受限），才不會跳確認視窗。
        # 上次測試中斷留下的設定、或你自己的通用規則剛好符合 probe_ 這些欄名時，它們一開始就不是開放 —— 那就不測這一段
        $probeLevels = @(Get-PbiProtection -Table $tmpTable -All | Where-Object { $_.Level -ne 'open' })
        if ($probeLevels.Count -gt 0) { throw "暫時的表 '$tmpTable' 的欄位一開始就不是「開放」（$(($probeLevels | ForEach-Object { "$($_.Column)=$($_.Level)" }) -join '、')）。可能是上次測試留下的設定：請到儀表板的「資料保護」分頁按「清除這些設定」，再重跑。" }
        $settingsToClear = $true
        Set-PbiProtection -Table $tmpTable -Column probe_pid -Level pseudonym     | Out-Null
        Set-PbiProtection -Table $tmpTable -Column probe_val -Level aggregateOnly | Out-Null

        Invoke-GuardCase '✅ 換成代號：取相異值'               $false "EVALUATE VALUES($T[probe_pid])"
        Invoke-GuardCase '✅ 換成代號：分組，外面包 TOPN'       $false "EVALUATE TOPN(2, SUMMARIZECOLUMNS($T[probe_pid], `"s`", SUM($T[probe_val])), [s], DESC)"
        Invoke-GuardCase '✅ 換成代號：計數'                   $false "EVALUATE ROW(`"n`", DISTINCTCOUNT($T[probe_pid]))"
        Invoke-GuardCase '換成代號：用別名取值'                $true  "EVALUATE SELECTCOLUMNS($T, `"x`", $T[probe_pid])"
        Invoke-GuardCase '換成代號：用 MAX 取值'               $true  "EVALUATE ROW(`"x`", MAX($T[probe_pid]))"
        Invoke-GuardCase '換成代號：經由變數變成純量'           $true  "EVALUATE VAR v = TOPN(1, VALUES($T[probe_pid])) RETURN ROW(`"x`", v)"
        Invoke-GuardCase '換成代號：不帶表名的參考'             $true  "EVALUATE SELECTCOLUMNS($T, `"x`", [probe_pid])"
        Invoke-GuardCase '換成代號：在 TOPN 的排序引數裡比較'    $true  "EVALUATE TOPN(1, VALUES($T[probe_pid]), $T[probe_pid] = `"alpha`")"
        Invoke-GuardCase '只能彙總：逐列取值'              $true  "EVALUATE SELECTCOLUMNS($T, `"x`", $T[probe_val])"
        Invoke-GuardCase '✅ 只能彙總：加總'               $false "EVALUATE ROW(`"s`", SUM($T[probe_val]))"
        Invoke-GuardCase '整表取出（裡面有只能彙總的欄位）'  $true  "EVALUATE $T"
        Invoke-GuardCase '✅ 同一張表的開放欄位照常可讀'    $false "EVALUATE SELECTCOLUMNS($T, `"g`", $T[probe_grp])"

        # 值真的換成代號了嗎？這張表的值是測試自己造的，可以放心檢查內容
        $masked = $false
        try {
            $r = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{ Query = "EVALUATE VALUES($T[probe_pid])"; MaxRows = 10 }
            $vals = @($r.rows | ForEach-Object { $_.PSObject.Properties | ForEach-Object { "$($_.Value)" } })
            $masked = ($vals.Count -eq 3) -and (@($vals | Where-Object { $_ -notmatch '^ID_[0-9A-F]{10}$' }).Count -eq 0) -and
                      (@($r.pseudonymized).Count -eq 1)
        } catch { }
        if ($masked) { $script:pass++ } else { $script:fail++ }
        Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f $(if ($masked) { '✅ 通過' } else { '❌ 失敗' }),
                    '換成代號：回傳的值全部是代號', '代號', $(if ($masked) { '代號' } else { '不是' })) `
                   -ForegroundColor $(if ($masked) { 'Green' } else { 'Red' })

        Invoke-GuardCase '✅ 只能彙總：寫在篩選條件裡'       $false "EVALUATE SUMMARIZECOLUMNS($T[probe_grp], FILTER($T, $T[probe_val] > 15), `"s`", SUM($T[probe_val]))"
        Invoke-GuardCase 'ERROR 把內容放進錯誤訊息'         $true  "EVALUATE ROW(`"n`", COUNTROWS(FILTER($T, ERROR($T[probe_pid]))))"

        # 型別轉換失敗時，引擎的錯誤訊息裡就是那個轉不過去的值。外面有計數函式，逐欄檢查會放行 ——
        # 所以要確認回給呼叫端的錯誤訊息裡沒有值。這張表的值是測試自己造的，可以拿來比對。
        $leakErr = $null
        try { $null = Invoke-PbiApi -Path "/api/query" -Method POST -Body @{ Query = "EVALUATE ROW(`"n`", COUNTROWS(FILTER($T, VALUE($T[probe_pid]) > 0)))"; MaxRows = 1 } }
        catch { $leakErr = "$_" }
        if (-not $leakErr) {
            Write-Host "⚠️ 略過  錯誤訊息不夾帶欄位內容 —— 引擎沒有報錯，這個案例沒有測到" -ForegroundColor Yellow
        } else {
            $clean = $leakErr -notmatch 'alpha|beta|gamma'
            if ($clean) { $script:pass++ } else { $script:fail++ }
            Write-Host ("{0}  {1,-32} 預期={2,-6} 實際={3}" -f $(if ($clean) { '✅ 通過' } else { '❌ 失敗' }),
                        '錯誤訊息不夾帶欄位內容', '已遮蔽', $(if ($clean) { '已遮蔽' } else { '看得到值' })) `
                       -ForegroundColor $(if ($clean) { 'Green' } else { 'Red' })
        }

        Set-PbiProtection -Table $tmpTable -Column probe_pid -Level countOnly | Out-Null      # 再收緊一級
        Invoke-GuardCase '只能計數：不能分組'                $true  "EVALUATE VALUES($T[probe_pid])"
        Invoke-GuardCase '✅ 只能計數：計數'                 $false "EVALUATE ROW(`"n`", DISTINCTCOUNT($T[probe_pid]))"
    } finally {
        # 先刪表（連同它的計算資料行），再刪量值
        foreach ($item in @($cleanup | Sort-Object { $_.Kind -ne 'table' })) {
            $what = if ($item.Kind -eq 'table') { "計算表 '$($item.Name)'" } else { "量值 '$($item.Table)'[$($item.Name)]" }
            try {
                if ($item.Kind -eq 'table') { Remove-PbiTable -Name $item.Name | Out-Null }
                else { Remove-PbiMeasure -Table $item.Table -Name $item.Name | Out-Null }
                Write-Host "   已清除暫時的$what" -ForegroundColor Gray
            } catch {
                Write-Host "   ⚠️ 暫時的$what 清除失敗，請手動刪除（或關閉 Power BI 時不要存檔）" -ForegroundColor Yellow
            }
        }
        # 暫時的表刪掉之後，替它設的等級就對不到欄位了 —— 清掉。欄位已經不存在，所以這不算放寬、不會跳確認視窗。
        # 表沒刪成的話欄位還在，這時清設定就是真的放寬（會跳確認視窗、卡住最多兩分鐘）—— 那就不清，留給使用者處理
        $tableGone = -not ((Get-PbiSchema).Tables | Where-Object { $_.Name -eq $tmpTable })
        if ($settingsToClear -and -not $tableGone) {
            Write-Host "   ⚠️ 暫時的表還在，保護設定先不清。請手動刪掉 '$tmpTable'，再到儀表板「資料保護」分頁按「清除這些設定」" -ForegroundColor Yellow
        } elseif ($settingsToClear) {
            try { Set-PbiProtection -Table $tmpTable -Column probe_pid, probe_val -Level default | Out-Null; Write-Host "   已清除暫時的保護設定" -ForegroundColor Gray }
            catch { Write-Host "   ⚠️ 暫時的保護設定清除失敗：到儀表板「資料保護」分頁按「清除這些設定」即可" -ForegroundColor Yellow }
        }
    }
}

Write-Host ""
Write-Host ("結果：$pass 通過 / $fail 失敗") -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($fail -gt 0) { Write-Host "❗ 有項目未達預期 —— 上面標紅的那幾行就是漏洞或誤擋。" -ForegroundColor Red }
