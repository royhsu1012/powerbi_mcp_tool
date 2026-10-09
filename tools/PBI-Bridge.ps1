# =============================================================================
# PBI-Bridge.ps1 — Power BI 橋接服務輔助函式
#
# 用法：在專案根目錄執行  . ".\tools\PBI-Bridge.ps1"
#
# 為什麼要用這個而不是直接 Invoke-RestMethod：
#   Windows PowerShell 5.1 送出字串 body 時預設不是 UTF-8，中文欄位名與中文
#   DAX 字串會在傳輸中損毀，伺服器收到破碎 JSON 後回 400。本檔一律將 body
#   轉成 UTF-8 位元組後送出，並在出錯時自動讀回應內容。
# =============================================================================

$script:PbiRoot    = Split-Path -Parent $PSScriptRoot

# 專案根目錄。刻意用函式而不是直接讀 $script:PbiRoot ——
# 從「子腳本」呼叫本檔函式時（例如 tools\Test-DataGuard.ps1），$script: 作用域
# 不保證解析得到，會變成 $null，之後每一個用到專案路徑的函式都會報出不相干的錯誤。
# $PSScriptRoot 在函式內解析的是本檔所在位置，與呼叫端作用域無關。
function Get-PbiRootPath {
    if ($PSScriptRoot) { return (Split-Path -Parent $PSScriptRoot) }
    return $script:PbiRoot
}
$script:PbiBaseUrl = "http://localhost:5500"
$script:PbiApiKey  = $null
# 目前選定的目標實例。
#   PbiTarget     = Port（HTTP 標頭只能放 ASCII，中文檔名不能直接當標頭值）
#   PbiTargetPath = 完整路徑，供 PBI 重開、Port 變動後自動重新解析
$script:PbiTarget     = $null
$script:PbiTargetPath = $null

function Get-PbiApiKey {
    <#
      讀取 API Key（讀一次後快取，不寫死在檔案裡）。
      金鑰由橋接服務在第一次啟動時產生，存在這台電腦的使用者資料夾（%LOCALAPPDATA%\PBI_AI_Bridge\api-key.txt），
      不在專案資料夾裡 —— 複製或分享專案資料夾不會把它帶出去。內部使用即可，不要印出來。
    #>
    if ($script:PbiApiKey) { return $script:PbiApiKey }
    # 和服務用同一種方式找資料夾（不讀 $env:LOCALAPPDATA：那個變數可以被改掉，兩邊就對不上了）
    $keyPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PBI_AI_Bridge\api-key.txt'
    if (-not (Test-Path -LiteralPath $keyPath)) {
        throw ("找不到 API 金鑰檔。金鑰是橋接服務啟動時產生的 —— 請雙擊 🚀啟動PBI終極儀表板.bat。" +
               "黑窗已經開著的話，代表跑的是舊版：請先關掉它再雙擊一次。")
    }
    $key = ([System.IO.File]::ReadAllText($keyPath)).Trim()
    if (-not $key) { throw "API 金鑰檔是空的。請關掉服務的黑窗、重新雙擊 🚀啟動PBI終極儀表板.bat，服務會重新產生一把。" }
    $script:PbiApiKey = $key
    return $script:PbiApiKey
}

function Invoke-PbiApi {
    <#  所有 API 呼叫的統一入口：處理 UTF-8 編碼與錯誤內容讀取 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('GET','POST')][string]$Method = 'GET',
        [hashtable]$Body,
        [int]$TimeoutSec = 120,
        [switch]$NoRetry
    )
    $headers = @{ "X-API-Key" = (Get-PbiApiKey) }
    # 帶上選定的目標實例；沒選就交給伺服器判斷（單一實例自動採用，多實例會拒絕）
    if ($script:PbiTarget) { $headers["X-PBI-Target"] = $script:PbiTarget }
    $uri     = "$script:PbiBaseUrl$Path"
    try {
        if ($Method -eq 'GET') {
            return Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec $TimeoutSec
        }
        # ⚠️ 關鍵：轉成 UTF-8 位元組，否則中文會壞掉
        $json  = $Body | ConvertTo-Json -Compress -Depth 10
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        return Invoke-RestMethod -Uri $uri -Method Post -Headers $headers `
                                 -ContentType "application/json; charset=utf-8" `
                                 -Body $bytes -TimeoutSec $TimeoutSec
    } catch {
        $detail = $null
        if ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
                $detail = $reader.ReadToEnd()
            } catch { }
        }
        # PBI 關掉再開時 Port 會變，記住的目標就失效了。
        # 用記住的完整路徑重新解析一次再重試 —— 切換檔案是這個工具的日常，不該每次都要手動重選。
        if (-not $NoRetry -and $detail -match '找不到 Port' -and $script:PbiTargetPath) {
            $saved = $script:PbiTargetPath
            $script:PbiTarget = $null
            try {
                $again = @((Invoke-PbiApi -Path "/api/instances" -NoRetry).Instances |
                           Where-Object { $_.filePath -eq $saved })
            } catch { $again = @() }
            if ($again.Count -eq 1) {
                $script:PbiTarget = "$($again[0].port)"
                Write-Verbose "目標實例的 Port 已變更，自動更新為 $($again[0].port)"
                return Invoke-PbiApi -Path $Path -Method $Method -Body $Body -TimeoutSec $TimeoutSec -NoRetry
            }
            $script:PbiTargetPath = $null
            throw "API 錯誤 ($Path): 先前選定的實例已不存在（$saved）。請重新執行 Use-PbiInstance。"
        }

        if ($detail) { throw "API 錯誤 ($Path): $detail" }
        throw "API 錯誤 ($Path): $($_.Exception.Message)"
    }
}

function Test-PbiBridge {
    <#  開工前的環境檢查：服務有沒有跑、有哪些 PBI 可以操作、目前選了哪一個 #>
    try {
        $pong = Invoke-RestMethod "$script:PbiBaseUrl/ping" -TimeoutSec 3
        if ($pong -ne 'pong') { Write-Host "⚠️ 服務回應異常: $pong" -ForegroundColor Yellow; return }
    } catch {
        Write-Host "❌ 橋接服務未啟動 — 請雙擊 🚀啟動PBI終極儀表板.bat 並保持視窗開啟" -ForegroundColor Red
        return
    }
    Write-Host "✅ 橋接服務就緒 ($script:PbiBaseUrl)" -ForegroundColor Green

    $r = Invoke-PbiApi -Path "/api/instances"
    if ($r.Count -eq 0) {
        Write-Host "❌ 沒有偵測到執行中的 Power BI Desktop — 請先開啟 PBIX/PBIP 檔案" -ForegroundColor Red
        return
    }

    Write-Host "📂 偵測到 $($r.Count) 個 Power BI 實例：" -ForegroundColor Cyan
    $r.Instances | ForEach-Object {
        $mark = if ($script:PbiTarget -and
                    ("$($_.port)" -eq "$script:PbiTarget" -or $_.filePath -like "*$script:PbiTarget*")) { '►' } else { ' ' }
        "  {0} Port {1,-6} {2,-45} [{3}]" -f $mark, $_.port, $_.fileName, $_.kind
    }

    if ($script:PbiTarget) {
        Write-Host "🎯 目前目標：$script:PbiTarget" -ForegroundColor Green
    } elseif ($r.Count -gt 1) {
        Write-Host "⚠️ 有多個實例但尚未選定目標 — 所有請求（包含讀取）都會被拒絕。請先執行 Use-PbiInstance <檔名片段或 Port>" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# 實例切換（同時開多個 PBI 時使用）
# ---------------------------------------------------------------------------

function Get-PbiInstances {
    <#
      列出所有執行中的 Power BI 實例（各自的 Port、檔案、行程）。
      一律回傳陣列 —— PS 5.1 對單一元素會攤平成物件，讓 .Count 變成 $null。
      代價是不能直接接管線（整個陣列會被當成一個物件，Select-Object 會印出空白）：
      要接管線請加括號  (Get-PbiInstances) | Select-Object port, fileName
    #>
    ,@((Invoke-PbiApi -Path "/api/instances").Instances)
}

function Use-PbiInstance {
    <#
      選定接下來要操作的 Power BI 實例。之後所有指令都會自動帶上這個目標。

      -Target 可以是 Port 數字，或檔名／路徑的片段（不分大小寫）。
      不帶參數則清除選擇（回到「只有一個實例才自動採用」的模式）。

      範例：
        Use-PbiInstance 銷售報表
        Use-PbiInstance 55820
        Use-PbiInstance            # 清除
    #>
    param([Parameter(Position=0)][string]$Target)

    if (-not $Target) {
        $script:PbiTarget     = $null
        $script:PbiTargetPath = $null
        Write-Host "已清除目標選擇。" -ForegroundColor Gray
        return
    }

    $all = Get-PbiInstances
    $hits = @($all | Where-Object {
        "$($_.port)" -eq $Target -or $_.filePath -like "*$Target*" -or $_.windowTitle -like "*$Target*"
    })

    if ($hits.Count -eq 0) {
        Write-Host "❌ 沒有實例符合「$Target」。目前有：" -ForegroundColor Red
        $all | ForEach-Object { "     Port {0,-6} {1}" -f $_.port, $_.fileName }
        return
    }
    if ($hits.Count -gt 1) {
        Write-Host "❌「$Target」同時符合 $($hits.Count) 個實例，請改用 Port 精確指定：" -ForegroundColor Red
        $hits | ForEach-Object { "     Port {0,-6} {1}" -f $_.port, $_.fileName }
        return
    }

    # 用 Port 當標頭值（ASCII 安全），另外記住完整路徑供 PBI 重開後自動重新解析
    $script:PbiTarget     = "$($hits[0].port)"
    $script:PbiTargetPath = $hits[0].filePath
    Write-Host "🎯 目標已設定：$($hits[0].fileName)  (Port $($hits[0].port), $($hits[0].kind))" -ForegroundColor Green
    if ($hits[0].filePath) {
        Write-Host "   路徑：$($hits[0].filePath)" -ForegroundColor Gray
    } else {
        # 服務只能從 Power BI 的啟動參數得知路徑：先開 Power BI 再從裡面選檔案就拿不到
        Write-Host "   路徑：不明（這份檔案是從 Power BI 裡面開啟的）。讀寫模型與資料保護照常；" -ForegroundColor Yellow
        Write-Host "         存檔驗證、報表健檢不能用，快照與 M 備份下次開啟找不回來 —— 需要的話請使用者改用雙擊檔案的方式開啟" -ForegroundColor Yellow
    }
}

function Get-PbiOverview {
    <#
      開工用的一頁摘要：選定的 Power BI、模型大小、健檢、資料保護 —— 一次呼叫看完。
      取代開工時原本要分開跑的 Test-PbiBridge / Get-PbiInstances / Get-PbiSchema / Test-PbiModel / Get-PbiProtection。
      只輸出結構與計數：不含 M 腳本，也不含任何欄位的內容。

      -Target  同時開了多個 Power BI 時用它選定（檔名片段或 Port），等同先跑 Use-PbiInstance
      -Tables  資料表清單最多列幾張（預設 30；0＝不列）

      範例：
        Get-PbiOverview
        Get-PbiOverview 銷售報表
    #>
    param(
        [Parameter(Position=0)][string]$Target,
        [int]$Tables = 30
    )
    # 屬性不存在時 @($null).Count 會是 1 —— 計數一律經過這裡
    $count = { param($x) if ($null -eq $x) { 0 } else { @($x).Count } }
    $brief = { param($e) $s = "$e" -replace '\s+', ' '; if ($s.Length -gt 200) { $s.Substring(0, 200) + '…' } else { $s } }

    try { $all = @((Invoke-PbiApi -Path "/api/instances").Instances) }
    catch {
        "❌ 連不上橋接服務 —— 請使用者雙擊 🚀啟動PBI終極儀表板.bat 並保持視窗開啟。"
        "   （$(& $brief $_)）"
        return
    }
    if ($all.Count -eq 0) {
        "❌ 服務就緒，但沒有偵測到開著檔案的 Power BI Desktop —— 請使用者先開啟 .pbix / .pbip。"
        return
    }

    if ($Target) {
        Use-PbiInstance $Target
        if (-not $script:PbiTarget) { return }          # 選不到：Use-PbiInstance 已經說明原因
    }
    $cur = $null
    if ($script:PbiTarget) { $cur = $all | Where-Object { "$($_.port)" -eq "$script:PbiTarget" } | Select-Object -First 1 }
    if (-not $cur) {
        if ($all.Count -gt 1) {
            "⚠️ 有 $($all.Count) 個 Power BI 開著，還沒選定要操作哪一個（選定之前所有請求都會被拒絕）："
            $all | ForEach-Object { "   Port {0,-6} {1}  [{2}]" -f $_.port, $_.fileName, $_.kind }
            "   選定並看摘要：Get-PbiOverview <檔名片段或 Port>"
            return
        }
        $cur = $all[0]
    }

    "📂 $($cur.fileName)  (Port $($cur.port), $($cur.kind))"
    if ($cur.filePath) { "   路徑：$($cur.filePath)" }
    else { "   ⚠️ 路徑不明（從 Power BI 裡面開啟的）：Save-PbiModel、Test-PbiReport 不能用，快照與 M 備份重開後找不回來" }

    try {
        $schema = Get-PbiSchema
        $health = Test-PbiModel
        $prot   = Get-PbiProtection -Raw
    } catch {
        "❌ 讀取模型失敗：$(& $brief $_)"
        return
    }

    $tbl = @($schema.Tables)
    $nCols = 0; $nMeas = 0
    foreach ($t in $tbl) { $nCols += (& $count $t.Columns); $nMeas += (& $count $t.Measures) }
    "模型：$($tbl.Count) 張表／$nCols 個資料行／$nMeas 個量值／$([int]$health.totalRelationships) 條關聯"

    $f = $health.findings
    $broken = @(); if ($f -and $null -ne $f.brokenObjects) { $broken = @($f.brokenObjects) }
    if ($broken.Count -gt 0) {
        $names = @($broken | Select-Object -First 8 | ForEach-Object { $_.object }) -join '、'
        "健檢：⛔ 有 $($broken.Count) 個壞掉的公式：$names$(if ($broken.Count -gt 8) { ' …' })"
    } else {
        "健檢：沒有壞掉的公式"
    }
    if ($f) {
        $hints = [ordered]@{
            '雙向關聯'               = (& $count $f.biDirectionalRelationships)
            '多對多關聯'             = (& $count $f.manyToManyRelationships)
            '停用的關聯'             = (& $count $f.inactiveRelationships)
            '同名量值'               = (& $count $f.duplicateMeasureNames)
            '沒設格式的量值'         = (& $count $f.measuresWithoutFormat)
            '孤島表'                 = (& $count $f.islandTables)
            '可能沒用到的資料行'     = (& $count $f.possiblyUnusedColumns)
            '該改成量值的計算資料行' = (& $count $f.calcColumnsUsingAggregation)
            '自動日期表'             = [int]$f.autoDateTableCount
        }
        $shown = @($hints.Keys | Where-Object { $hints[$_] -gt 0 } | ForEach-Object { "$_ $($hints[$_])" })
        if ($shown.Count -gt 0) { "   其他建議：$($shown -join '、')（細節：Test-PbiModel）" }
    }

    $restricted = @{}
    $by = @{ open = 0; pseudonym = 0; countOnly = 0; aggregateOnly = 0 }
    foreach ($t in @($prot.tables)) {
        $n = 0
        foreach ($c in @($t.columns)) {
            $lv = "$($c.level)"
            if ($by.ContainsKey($lv)) { $by[$lv]++ }
            if ($lv -and $lv -ne 'open') { $n++ }
        }
        $restricted["$($t.name)"] = $n
    }
    if (-not $prot.enabled) {
        "資料保護：⚠️ 關閉中（appsettings.json 的 DataProtection:Enabled）—— 查詢可以取出任何欄位的內容"
    } elseif ($prot.problem) {
        "資料保護：⛔ 這份模型的設定檔讀不出來，查詢全部會被擋下 —— 請使用者到儀表板的「資料保護」分頁按「重設」"
    } else {
        "資料保護：開放 $($by.open)／換成代號 $($by.pseudonym)／只能計數 $($by.countOnly)／只能彙總 $($by.aggregateOnly)"
        if (-not $prot.configured) {
            "   ⚠️ 這份模型還沒設定過，只有通用規則在擋（中文欄名幾乎都是開放的）—— 查資料之前先提醒使用者到儀表板設定"
        }
        if ($prot.inheritedFrom) {
            "   設定是靠內容認回來的（當時的檔案：$($prot.inheritedFrom)）—— 請使用者到儀表板掃一眼等級"
        }
        $stale = (& $count $prot.stale)
        if ($stale -gt 0) { "   有 $stale 筆設定對不到現在的欄位（改名或刪除了）" }
    }

    if ($Tables -gt 0 -and $tbl.Count -gt 0) {
        "資料表（資料行／量值／受限的資料行）："
        foreach ($t in ($tbl | Select-Object -First $Tables)) {
            $r = 0; if ($restricted.ContainsKey("$($t.Name)")) { $r = $restricted["$($t.Name)"] }
            "   $($t.Name)：$(& $count $t.Columns)／$(& $count $t.Measures)／$r$(if ($t.IsHidden) { '  [隱藏]' })"
        }
        if ($tbl.Count -gt $Tables) { "   …還有 $($tbl.Count - $Tables) 張（Get-PbiOverview -Tables $($tbl.Count)）" }
    }
    "下一步：查某張表之前先看哪些欄位受限 → Get-PbiProtection -Table <表>；量值 → Get-PbiMeasures | Select-Object Table, Measure, Expression"
}

# ---------------------------------------------------------------------------
# 讀取
# ---------------------------------------------------------------------------

function Get-PbiSchema {
    <#  全部資料表、欄位、DAX 量值原始碼、Power Query M 腳本 #>
    Invoke-PbiApi -Path "/api/schema"
}

function Get-PbiRelationships {
    <#  關聯線：基數、雙向篩選、是否啟用 #>
    Invoke-PbiApi -Path "/api/relationships"
}

function Get-PbiMeasures {
    <#  攤平所有量值成一張表，方便 grep / 比對命名 #>
    param([string]$TableFilter)
    $schema = Get-PbiSchema
    $out = foreach ($t in $schema.Tables) {
        if ($TableFilter -and $t.Name -notlike $TableFilter) { continue }
        foreach ($m in $t.Measures) {
            [PSCustomObject]@{
                Table         = $t.Name
                Measure       = $m.Name
                Expression    = $m.Expression
                Format        = $m.FormatString
                DisplayFolder = $m.DisplayFolder
                Description   = $m.Description
            }
        }
    }
    $out
}

function Get-PbiRoles       { Invoke-PbiApi -Path "/api/roles" }
function Get-PbiExpressions { Invoke-PbiApi -Path "/api/expressions" }

function Test-PbiModel {
    <#  模型健檢：壞掉的公式、雙向關聯、孤島表、可能沒人用的欄位…… #>
    Invoke-PbiApi -Path "/api/validate" -TimeoutSec 180
}

function Test-PbiReport {
    <#
      報表健檢（僅 PBIP）：直接改過報表資料夾裡的檔案（visual.json、page.json…）之後，
      在請使用者按「接受變更」之前先跑一次。

      Errors   會讓視覺壞掉或變成空白：JSON 格式錯誤、引用了模型裡不存在的欄位或量值、
               頁面清單對不上、視覺名稱重複
      Warnings 值得看一眼：視覺超出頁面或互相重疊、互動設定指向不存在的視覺、
               格式設定指向已經不在視覺裡的欄位、布林欄位上的篩選

      比對的是「磁碟上的報表檔案」對「Power BI 記憶體中的模型」，所以要先選定那份 PBIP。
      只回傳結構資訊（頁面、視覺類型與標題、欄位名稱、座標），不含篩選條件裡的值。
      查不到的：文字有沒有被擠壓、配色、好不好讀 —— 那些要請使用者看畫面或截圖。

      範例：
        $r = Test-PbiReport
        "$($r.pages) 頁 / $($r.visuals) 個視覺：$($r.errorCount) 個錯誤、$($r.warningCount) 個警告"
        $r.errors   | Format-Table page, visual, kind, detail -Wrap
        $r.warnings | Group-Object kind | Select-Object Count, Name
    #>
    Invoke-PbiApi -Path "/api/validate-report" -TimeoutSec 180
}

function Get-PbiMQuery {
    <#
      讀取單一資料表的 Power Query M 腳本。

      ⚠️ M 腳本含連線字串、伺服器位址、資料庫名稱，而且篩選步驟裡很可能有
         寫死的客戶名稱（例如 Table.SelectRows(..., each [customer] = "…")）。
         所以**預設只存檔、回傳路徑，不回傳內容** —— 存到本機磁碟不等於送進雲端。

      要讀內容請明確加 -Show，並先想清楚為什麼需要整段 M 進入 context。
      多數情況你需要的是「這張表的 M 長什麼樣的結構」，那用 -Show 之前
      先問使用者能不能只描述問題段落。
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$Table,
        [switch]$Show,
        [string]$Label = "current"
    )
    $t = (Get-PbiSchema).Tables | Where-Object { $_.Name -eq $Table }
    if (-not $t)          { throw "找不到資料表：$Table" }
    if (-not $t.MQuery)   { throw "資料表 '$Table' 沒有 M 腳本（可能是計算表或計算群組）" }

    # 一律存檔。每個模型一個子資料夾，命名沿用 snapshots 的規則。
    # 直接跟伺服器要那個資料夾名，而不是在這裡重算雜湊 —— 演算法只該有一份實作。
    $key = Split-Path -Leaf (Invoke-PbiApi -Path "/api/snapshots").snapshotPath
    $dir = Join-Path (Join-Path (Get-PbiRootPath) "PowerQuery_Scripts") $key
    New-Item -ItemType Directory -Force $dir | Out-Null
    $safe = ($Table -replace '[\\/:*?"<>|]', '_')
    $path = Join-Path $dir ("{0}__{1:yyyyMMdd_HHmmss}__{2}.pq" -f $safe, (Get-Date), ($Label -replace '[^\w\-]','_'))
    [System.IO.File]::WriteAllText($path, $t.MQuery, (New-Object System.Text.UTF8Encoding $false))
    Write-Host "已存檔：$path" -ForegroundColor Gray

    if ($Show) { return $t.MQuery }

    # 預設回傳「這份 M 存在哪、有多大」，內容留在磁碟上不進 context
    [PSCustomObject]@{
        資料表 = $Table
        已存檔 = $path
        字元數 = $t.MQuery.Length
        提示   = "內容未載入 context。確實需要讀取時加 -Show，並先說明用途。"
    }
}

function Get-PbiTableProfile {
    <#
      資料表的量化剖析，用來比對「改 M 之前 / 之後」有沒有悄悄弄壞資料。
      預覽只給你前 1000 列，看不出後面掉了多少 —— 這個看得出來。

      全部是彙總查詢，不會有任何明細列進入 context。

      範例：
        $before = Get-PbiTableProfile DimProduct -Columns Category, ProductId
        # …改完 M、重整後…
        $after  = Get-PbiTableProfile DimProduct -Columns Category, ProductId
        Compare-PbiTableProfile $before $after
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$Table,
        [string[]]$Columns,
        [int]$TimeoutSeconds = 120
    )
    $t = (Get-PbiSchema).Tables | Where-Object { $_.Name -eq $Table }
    if (-not $t) { throw "找不到資料表：$Table" }

    $numeric = @('Int64','Double','Decimal')
    # 別名一律用不帶欄名的代碼（b0 / d0 / s0），結果回來之後再對回欄名。
    # 別名裡帶著欄名的話，欄名剛好符合保護規則的樣式（amount、customer…）時，
    # 結果欄位會被當成那個受限欄位擋下來 —— 而改 M 前後最常拿來剖析的正是那些欄位。
    $parts = @("`"n`", COUNTROWS('$Table')")
    $names = [ordered]@{ '[n]' = '列數' }
    $i = 0
    foreach ($c in $Columns) {
        $col = $t.Columns | Where-Object { $_.Name -eq $c }
        if (-not $col) { throw "資料表 '$Table' 沒有欄位 [$c]" }
        $ref = "'$Table'[$c]"
        # COUNTBLANK/SUM 在沒有結果時回傳 BLANK，會讓「0」和「查詢失敗」長得一樣 —— 一律轉成 0
        $parts += "`"b$i`", COALESCE(COUNTBLANK($ref), 0)";     $names["[b$i]"] = "空值_$c"
        $parts += "`"d$i`", COALESCE(DISTINCTCOUNT($ref), 0)";  $names["[d$i]"] = "相異_$c"
        if ($numeric -contains $col.DataType) { $parts += "`"s$i`", COALESCE(SUM($ref), 0)"; $names["[s$i]"] = "總和_$c" }
        $i++
    }

    $r = Invoke-Dax ("EVALUATE ROW(" + ($parts -join ", ") + ")") -TimeoutSeconds $TimeoutSeconds
    $out = [ordered]@{ 資料表 = $Table; 欄位數 = @($t.Columns).Count; 取樣時間 = (Get-Date).ToString('HH:mm:ss') }
    $row = $r.Rows[0]
    foreach ($k in $names.Keys) { $out[$names[$k]] = $row.$k }
    [PSCustomObject]$out
}


function Compare-PbiTableProfile {
    <#  比對兩份剖析結果，只列出有變動的項目。沒變動的不佔版面。 #>
    param(
        [Parameter(Mandatory, Position=0)]$Before,
        [Parameter(Mandatory, Position=1)]$After
    )
    $skip = '資料表','取樣時間'
    $rows = foreach ($p in $Before.PSObject.Properties) {
        if ($skip -contains $p.Name) { continue }
        $b = $p.Value; $a = $After.$($p.Name)
        if ("$b" -eq "$a") { continue }
        $delta = $null
        if ($b -is [ValueType] -and $a -is [ValueType] -and $b -ne 0) {
            $delta = "{0:+0.0#;-0.0#;0}%" -f ((($a - $b) / [double]$b) * 100)
        }
        [PSCustomObject]@{ 項目 = $p.Name; 改前 = $b; 改後 = $a; 變化 = $delta }
    }
    if (-not $rows) { Write-Host "✅ 所有指標都沒有變動" -ForegroundColor Green; return }
    Write-Host "⚠️ 以下指標有變動，請確認是否符合預期：" -ForegroundColor Yellow
    $rows
}

function Invoke-Dax {
    <#
      執行唯讀 DAX 查詢。查詢必須以 EVALUATE 或 DEFINE 開頭。
      範例：(Invoke-Dax 'EVALUATE ROW("結果", [銷售總額])').Rows

      ⚠️ 資料保護由伺服器把關。每個欄位有一個等級（使用者在儀表板的「資料保護」分頁設定，
         用 Get-PbiProtection 查）：
           開放       可以讀取內容（逐列明細有列數上限）
           換成代號   可以計數，或直接當分組鍵 —— 結果裡的值會是 ID_3FA2B81C07 這種代號
           只能計數   只能放在 DISTINCTCOUNT / COUNTROWS 這類計數函式裡
           只能彙總   只能放在 SUM / AVERAGE / MIN / MAX 這類聚合函式裡（寫在篩選條件裡也可以）

      被擋下（403）時：先想這個值是不是真的需要 —— 多半換成彙總寫法就夠了。
      真的需要時，先向使用者說明要看哪些欄位、幾列、為什麼，再加 -AskUser 重送「同一句」：
      使用者的螢幕會跳出確認視窗（列出原因、受限欄位、列數與查詢內容），他按「是」才放行這一次。
      -AskUser 不是讓你自己放行的開關，按鈕在使用者那邊。被拒絕之後不要連續重送，
      也不要改寫查詢去繞 —— 回應的 retryAfterSeconds 之內伺服器不會再跳視窗。

      「換成代號」欄位的寫法：直接當分組鍵，外面可以再包一層 TOPN。

           (Invoke-Dax 'EVALUATE TOPN(20, SUMMARIZECOLUMNS(Customers[account name],
                          "額", SUM(Sales[amount])), [額], DESC)').Rows
           → ID_3FA2B81C07 | 12,345,678

         代號穩定（同一個值跨查詢都是同一個代號，可以串接分析），但推不回原值。
         真名只印在**服務主控台**，使用者看得到、AI 看不到。
         回應的 pseudonymized 會列出哪些欄位是代號 —— 不要把代號當成真實名稱解讀或寫進量值。
         不能放進變數或量值，也不能用 MAX / CONCATENATEX / SELECTCOLUMNS 取值或拿來比較。
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$Query,
        [int]$MaxRows = 1000,
        [int]$TimeoutSeconds = 60,
        # 這句查詢被資料保護擋下時，請使用者在確認視窗決定要不要放行這一次。先向使用者說明過再用。
        [switch]$AskUser
    )
    $body    = @{ Query = $Query; MaxRows = $MaxRows; TimeoutSeconds = $TimeoutSeconds }
    $timeout = $TimeoutSeconds + 30
    if ($AskUser) {
        $body.AskUser = $true
        # 沒有明講要幾列就不送 MaxRows：放行之後帶回幾列，由伺服器照逐列明細的上限決定
        # （確認視窗會把列數寫給使用者看，預設的 1000 會讓他以為你要一千列）
        if (-not $PSBoundParameters.ContainsKey('MaxRows')) { $body.Remove('MaxRows') }
        # 伺服器會先跑一次（被擋下）、等使用者回答（最多 120 秒）、同意之後再跑一次
        $timeout = 2 * $TimeoutSeconds + 200
    }
    Invoke-PbiApi -Path "/api/query" -Method POST -Body $body -TimeoutSec $timeout
}

# ---------------------------------------------------------------------------
# 資料保護：每個欄位的等級
# ---------------------------------------------------------------------------

function Get-PbiProtection {
    <#
      每個欄位目前的保護等級。寫查詢之前先看一眼，就不用靠被擋下來才知道哪些欄位受限。
      等級是結構資訊（不是資料內容），可以放心讀。

        level   open（開放）/ pseudonym（換成代號）/ countOnly（只能計數）/ aggregateOnly（只能彙總）
        source  explicit＝使用者在儀表板設定的、pattern＝通用規則（appsettings.json 的樣式）、default＝預設開放

      預設只列出「受限」的欄位（通常只有少數幾個）；-All 列出全部。

      範例：
        Get-PbiProtection                         # 哪些欄位受限
        Get-PbiProtection -Table Sales -All       # Sales 每一欄的等級
        (Get-PbiProtection -Raw).summary          # 各等級有幾欄
    #>
    param(
        [string]$Table,
        [switch]$All,
        [switch]$Raw
    )
    $r = Invoke-PbiApi -Path "/api/protection"
    if ($Raw) { return $r }
    if (-not $r.enabled) { Write-Warning "資料保護目前是關閉的（appsettings.json 的 DataProtection:Enabled），下列等級不會被強制執行。" }
    if ($r.problem)      { Write-Warning $r.problem }
    foreach ($t in $r.tables) {
        if ($Table -and $t.name -notlike $Table) { continue }
        foreach ($c in $t.columns) {
            if (-not $All -and $c.level -eq 'open') { continue }
            [PSCustomObject]@{
                Table = $t.name; Column = $c.name; DataType = $c.dataType
                Level = $c.level; Source = $c.source; Rule = $c.rule
            }
        }
    }
}

function Set-PbiProtection {
    <#
      變更欄位的保護等級。

      ⚠️ 收緊（開放 → 受限、其他 → 只能計數）立刻生效。
         放寬不是你能決定的：服務會在使用者的電腦上跳出確認視窗，列出要放寬哪些欄位，
         使用者按「是」才生效；按「否」或兩分鐘沒回應，整批都不套用（回 403，
         回應的 answer 說明原因、retryAfterSeconds 之內不會再跳視窗）。一次最多放寬 12 欄。
         所以要放寬之前，先向使用者說明要放寬哪些欄位、為什麼 —— 更好的做法是請他
         自己到儀表板（http://localhost:5500/）的「資料保護」分頁調整。
         不要為了讓查詢通過而去放寬保護。

      -Level default：拿掉逐欄設定，回到通用規則。

      範例：
        Set-PbiProtection -Table Customers -Column phone, address -Level countOnly
        Set-PbiProtection -Table Sales -Column discount -Level aggregateOnly
    #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string[]]$Column,
        [Parameter(Mandatory)][ValidateSet('open','pseudonym','countOnly','aggregateOnly','default')][string]$Level
    )
    $changes = @($Column | ForEach-Object { @{ Table = $Table; Column = $_; Level = $Level } })
    # 放寬時伺服器會等使用者回應確認視窗（最多兩分鐘），逾時要比那個長
    Invoke-PbiApi -Path "/api/protection" -Method POST -Body @{ Changes = $changes } -TimeoutSec 180
}

function Invoke-PbiDmv {
    <#
      執行 $SYSTEM DMV 查詢（模型的中繼資料與儲存統計）。
      這些系統檢視只含中繼資料與統計，不會回傳事實資料列。
      範例：Invoke-PbiDmv 'SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLES'

      ⚠️ 資料保護啟用時只能查白名單上的檢視（模型的中繼資料與儲存統計：TMSCHEMA_TABLES /
         COLUMNS / MEASURES / RELATIONSHIPS …、DISCOVER_STORAGE_*、DISCOVER_OBJECT_MEMORY_USAGE）。
         其他檢視一律 403 —— 有些會回傳欄位的實際內容（MDSCHEMA_MEMBERS）、M 腳本
         （TMSCHEMA_PARTITIONS）或別的連線執行過的查詢文字。相依性檢視（DISCOVER_CALC_DEPENDENCY）也不在名單內 ——
         它會帶出運算式原文。被擋時回應的 allowed 會列出可以查的。
         寫法只接受 SELECT <欄位或 *> FROM $SYSTEM.<檢視> [WHERE … | ORDER BY …]，一句只查一個檢視。
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$Query,
        [int]$MaxRows = 10000,
        [int]$TimeoutSeconds = 120
    )
    Invoke-PbiApi -Path "/api/dmv" -Method POST -Body @{
        Query = $Query; MaxRows = $MaxRows; TimeoutSeconds = $TimeoutSeconds
    } -TimeoutSec ($TimeoutSeconds + 30)
}

function Get-PbiModelStats {
    <#
      每張表佔用的記憶體（由 VertiPaq 分段統計加總）。
      彙總在 PowerShell 端完成 —— 明細不會進入 context。
    #>
    param([int]$Top = 20)
    $r = Invoke-PbiDmv 'SELECT DIMENSION_NAME, COLUMN_ID, USED_SIZE FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS'
    if ($r.Truncated) {
        Write-Warning "分段數超過 $($r.RowCount) 列已被截斷，以下數字為低估值。請提高 -MaxRows 重跑。"
    }
    $r.Rows | Group-Object DIMENSION_NAME | ForEach-Object {
        [PSCustomObject]@{
            Table    = $_.Name
            SizeMB   = [math]::Round((($_.Group | Measure-Object USED_SIZE -Sum).Sum) / 1MB, 2)
            Columns  = ($_.Group | Select-Object -ExpandProperty COLUMN_ID -Unique).Count
        }
    } | Sort-Object SizeMB -Descending | Select-Object -First $Top
}

# ---------------------------------------------------------------------------
# 存檔 / 重新整理
# ---------------------------------------------------------------------------

function Save-PbiModel {
    <#
      模擬 Ctrl+S 存檔，並驗證是否真的存到。驗證範圍只有選定實例自己的檔案：
      PBIX 看那一個檔，PBIP 看它的 .Report 與 .SemanticModel 兩個資料夾。

      -WaitSeconds  最多等幾秒。偵測到檔案寫完就提早回傳，所以調高不會讓一般情況變慢。
      -Expect       存檔後必須出現在磁碟上的文字，例如剛寫入的量值名稱或公式裡一小段有辨識度的內容。
                    fileChanged = true 只代表「有檔案被寫了」，不代表剛寫的東西進去了 ——
                    實際遇過存下去的是舊狀態（推測是 Power BI 還沒同步到 API 的寫入）。給了 -Expect
                    就改看 expectFound。只回傳找到與否，檔案內容不會進 context。（僅 PBIP；PBIX 是二進位檔。）
      -VerifyOnly   不送 Ctrl+S，只檢查磁碟現況。存檔結果不確定時用這個回頭確認，不要重按。

      範例：
        Save-PbiModel -Expect '銷售總額'
        Save-PbiModel -VerifyOnly -Expect '銷售總額'      # 過一會兒再確認一次，不送按鍵

      ⚠️ 會把 PBI Desktop 帶到前景，短暫搶走鍵盤焦點（SendKeys 的固有限制）。
      ⚠️ 沒存到時最多再試一次，不要連按 —— 連續送合成按鍵會被防毒視為可疑行為。
    #>
    param(
        [int]$WaitSeconds = 30,
        [string[]]$Expect,
        [switch]$VerifyOnly
    )
    $body = @{ WaitSeconds = $WaitSeconds }
    if ($Expect)     { $body['Expect']     = @($Expect) }
    if ($VerifyOnly) { $body['VerifyOnly'] = $true }
    Invoke-PbiApi -Path "/api/save" -Method POST -Body $body -TimeoutSec ($WaitSeconds + 60)
}

function Invoke-PbiRefresh {
    <#
      重新整理資料，讓結構性變更生效：
        建計算表／新增計算項目        → Invoke-PbiRefresh -Table <表>
        建或改關聯線／新增計算資料行  → Invoke-PbiRefresh -RefreshType calculate
        用 API 寫了 M、Desktop 沒出現套用提示 → Invoke-PbiRefresh -Table <表>（full，會重抓資料源）
      只改量值不需要。使用者自己在 Power Query 編輯器「關閉並套用」時也不需要 —— Power BI 會自己重整。
      不指定 -Table 就是整個模型；full 會重抓資料源，大表可能很久。

      回應是在重新整理「做完之後」才回來的（message 是「已完成」，elapsedMs 是實際耗時）——
      收到回應就可以直接驗算，不必等。
    #>
    param(
        [string]$Table,
        [ValidateSet('full','calculate','dataOnly','automatic','add','clearValues','defragment')]
        [string]$RefreshType = 'full',
        [int]$TimeoutSec = 900
    )
    $body = @{ RefreshType = $RefreshType }
    if ($Table) { $body['TableName'] = $Table }
    Invoke-PbiApi -Path "/api/refresh" -Method POST -Body $body -TimeoutSec $TimeoutSec
}

# ---------------------------------------------------------------------------
# 快照 / 還原（所有破壞性操作前的安全網）
# ---------------------------------------------------------------------------

function New-PbiSnapshot {
    <#  把整個模型定義序列化成 TMSL 存檔。動大手術前先跑這個。 #>
    param([string]$Label)
    Invoke-PbiApi -Path "/api/snapshot" -Method POST -Body @{ Label = $Label } -TimeoutSec 180
}

function Get-PbiSnapshots {
    <#  列出所有快照，最新的在最前面 #>
    (Invoke-PbiApi -Path "/api/snapshots").Snapshots
}

function Restore-PbiSnapshot {
    <#
      從快照還原量值／計算資料行／關聯／M 腳本／共用運算式。
      ⚠️ 快照中不存在的量值與關聯會被「刪除」—— 還原是回到當時的狀態，不是合併。
         強烈建議先加 -DryRun 看差異清單，確認後再實際執行。
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$File,
        [switch]$DryRun,
        [ValidateSet('measures','columns','relationships','expressions','mquery')]
        [string[]]$Scope
    )
    $body = @{ File = $File; DryRun = [bool]$DryRun }
    if ($Scope) { $body['Scope'] = $Scope }
    Invoke-PbiApi -Path "/api/restore" -Method POST -Body $body -TimeoutSec 300
}

# ---------------------------------------------------------------------------
# 寫入：量值與資料行
# ---------------------------------------------------------------------------

function Set-PbiMeasure {
    <#  新增或覆寫量值。同名者直接覆蓋；未指定的屬性維持原值。 #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$Name,
        [string]$Expression,
        [string]$Format,
        [string]$Description,
        [string]$DisplayFolder,
        [bool]$IsHidden
    )
    $body = @{ TableName = $Table; MeasureName = $Name }
    if ($PSBoundParameters.ContainsKey('Expression'))    { $body['Expression']    = $Expression }
    if ($PSBoundParameters.ContainsKey('Format'))        { $body['FormatString']  = $Format }
    if ($PSBoundParameters.ContainsKey('Description'))   { $body['Description']   = $Description }
    if ($PSBoundParameters.ContainsKey('DisplayFolder')) { $body['DisplayFolder'] = $DisplayFolder }
    if ($PSBoundParameters.ContainsKey('IsHidden'))      { $body['IsHidden']      = $IsHidden }
    Invoke-PbiApi -Path "/api/upsert-measure" -Method POST -Body $body
}

function Remove-PbiMeasure {
    <#  刪除量值 — 無法復原，執行前請先 New-PbiSnapshot #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$Name
    )
    Invoke-PbiApi -Path "/api/delete-measure" -Method POST -Body @{
        TableName = $Table; MeasureName = $Name
    }
}

function Move-PbiMeasure {
    <#  搬移量值到其他表，目標表不存在會自動建立 #>
    param(
        [Parameter(Mandatory)][string]$FromTable,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ToTable
    )
    Invoke-PbiApi -Path "/api/move-measure" -Method POST -Body @{
        FromTable = $FromTable; MeasureName = $Name; ToTable = $ToTable
    }
}

function Add-PbiColumn {
    <#
      新增資料行。DataType: text/int/decimal/currency/bool/date

        -Expression    DAX 計算資料行（同名的計算資料行會被覆寫）
        -SourceColumn  來源資料行：對應 M 腳本輸出的欄位名稱。用 Set-PbiMQuery 讓查詢多輸出一欄時，
                       模型不會自己長出那個資料行，要用這個補上。DataType 要和 M 輸出的型別一致。

      兩者擇一。來源資料行加完要 Invoke-PbiRefresh -Table <表> 才有資料；
      M 其實沒有輸出那個欄位的話重新整理會失敗，屆時用 Remove-PbiColumn 拿掉。

      範例：
        Add-PbiColumn -Table Sales -Name 毛利 -Expression "Sales[Amount] - Sales[Cost]" -DataType decimal
        Add-PbiColumn -Table Sales -Name region -SourceColumn region -DataType text
    #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$Name,
        [string]$Expression,
        [string]$SourceColumn,
        [ValidateSet('text','int','decimal','currency','bool','date')][string]$DataType = 'text'
    )
    if ([bool]$Expression -eq [bool]$SourceColumn) {
        throw "Add-PbiColumn：-Expression（DAX 計算資料行）與 -SourceColumn（M 輸出的來源資料行）請擇一指定。"
    }
    $body = @{ TableName = $Table; ColumnName = $Name; DataType = $DataType }
    if ($Expression)   { $body['Expression']   = $Expression }
    if ($SourceColumn) { $body['SourceColumn'] = $SourceColumn }
    # 給了 -SourceColumn 時，資料保護啟用下會等使用者回答確認視窗（最多 120 秒）—— 逾時要比那個長
    Invoke-PbiApi -Path "/api/add-column" -Method POST -Body $body -TimeoutSec 180
}

function Remove-PbiColumn {
    <#  刪除資料行。仍被關聯線使用時會被擋下。 #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$Name
    )
    Invoke-PbiApi -Path "/api/delete-column" -Method POST -Body @{ TableName = $Table; ColumnName = $Name }
}

function Set-PbiColumn {
    <#
      設定資料行屬性：格式、顯示資料夾、隱藏、排序依據、摘要方式、資料類別。
      只送出有指定的參數，其餘維持原值。
    #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$Name,
        [string]$Format,
        [string]$DisplayFolder,
        [bool]$IsHidden,
        [string]$SortByColumn,
        [ValidateSet('none','default','sum','min','max','count','average','distinctCount')]
        [string]$SummarizeBy,
        [string]$DataCategory,
        [string]$Description,
        [ValidateSet('text','int','decimal','currency','bool','date')][string]$DataType
    )
    $body = @{ TableName = $Table; ColumnName = $Name }
    if ($PSBoundParameters.ContainsKey('Format'))        { $body['FormatString']  = $Format }
    if ($PSBoundParameters.ContainsKey('DisplayFolder')) { $body['DisplayFolder'] = $DisplayFolder }
    if ($PSBoundParameters.ContainsKey('IsHidden'))      { $body['IsHidden']      = $IsHidden }
    if ($PSBoundParameters.ContainsKey('SortByColumn'))  { $body['SortByColumn']  = $SortByColumn }
    if ($PSBoundParameters.ContainsKey('SummarizeBy'))   { $body['SummarizeBy']   = $SummarizeBy }
    if ($PSBoundParameters.ContainsKey('DataCategory'))  { $body['DataCategory']  = $DataCategory }
    if ($PSBoundParameters.ContainsKey('Description'))   { $body['Description']   = $Description }
    if ($PSBoundParameters.ContainsKey('DataType'))      { $body['DataType']      = $DataType }
    Invoke-PbiApi -Path "/api/set-column-props" -Method POST -Body $body
}

function Set-PbiMQuery {
    <#
      覆寫資料表的 M 腳本。2026-09-16 恢復（先前誤以為「按套用也解不開」而移除）。

      ⚠️ 這只改「模型」那一份，而且寫入本身不會重抓資料。Power BI Desktop 的 Power Query
         文件是另一份。寫完之後 Desktop 的反應有兩種，兩種都實際遇過：
           · 顯示「查詢中有暫止的變更尚未套用」→ 請使用者按「套用」
           · 完全沒有提示 → Invoke-PbiRefresh -Table <表名>，新的 M 才會套用到資料上
         兩份內容不同時，套用有可能是用 Desktop 那份覆蓋回來，所以生效之後一定要
         確認留下的是這次寫入的版本。

      建議流程：
        $before = Get-PbiTableProfile <表名> -Columns <欄位…>   # 基準線
        Get-PbiMQuery <表名> -Label before                      # 留一份 .pq 備份
        Set-PbiMQuery <表名> -Expression $m                      # 寫入
        # → 問使用者 Desktop 有沒有出現套用提示：有就請他按；沒有就 Invoke-PbiRefresh -Table <表名>
        (Get-PbiMQuery <表名> -Show) -eq $m                      # True＝留下的是這次寫入的版本（在本機比對，M 不進 context）
        Compare-PbiTableProfile $before (Get-PbiTableProfile <表名> -Columns <欄位…>)

      M 多輸出了新欄位：模型不會自己長出資料行，用 Add-PbiColumn -SourceColumn 補上再重新整理。
      改的是共用查詢（不是資料表自己的 M）：用 Set-PbiExpression，之後對用到它的表重新整理。
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$Table,
        [Parameter(Mandatory)][string]$Expression
    )
    # 片段會讓整張表的 M 變成不合法 —— 一律要求完整的 let ... in
    if ($Expression -notmatch '(?s)^\s*let\b.*\bin\b') {
        throw "M 腳本看起來不是完整的 let ... in，拒絕寫入（不要送片段）。"
    }
    # 資料保護啟用下會等使用者回答確認視窗（最多 120 秒）—— 逾時要比那個長
    Invoke-PbiApi -Path "/api/update-m" -Method POST -Body @{
        TableName = $Table; Expression = $Expression
    } -TimeoutSec 180
}

# ---------------------------------------------------------------------------
# 寫入：關聯線
# ---------------------------------------------------------------------------

function Set-PbiRelationship {
    <#
      建立或更新關聯線。預設 many → one、單向篩選、啟用。
      範例：Set-PbiRelationship -FromTable Sales -FromColumn DateKey -ToTable '日期' -ToColumn DateKey
    #>
    param(
        [Parameter(Mandatory)][string]$FromTable,
        [Parameter(Mandatory)][string]$FromColumn,
        [Parameter(Mandatory)][string]$ToTable,
        [Parameter(Mandatory)][string]$ToColumn,
        [ValidateSet('one','many')][string]$FromCardinality = 'many',
        [ValidateSet('one','many')][string]$ToCardinality   = 'one',
        [ValidateSet('single','both','auto')][string]$CrossFilter = 'single',
        [bool]$IsActive = $true
    )
    Invoke-PbiApi -Path "/api/upsert-relationship" -Method POST -Body @{
        FromTable = $FromTable; FromColumn = $FromColumn
        ToTable   = $ToTable;   ToColumn   = $ToColumn
        FromCardinality = $FromCardinality; ToCardinality = $ToCardinality
        CrossFilterDirection = $CrossFilter; IsActive = $IsActive
    }
}

function Remove-PbiRelationship {
    param(
        [Parameter(Mandatory)][string]$FromTable,
        [Parameter(Mandatory)][string]$FromColumn,
        [Parameter(Mandatory)][string]$ToTable,
        [Parameter(Mandatory)][string]$ToColumn
    )
    Invoke-PbiApi -Path "/api/delete-relationship" -Method POST -Body @{
        FromTable = $FromTable; FromColumn = $FromColumn; ToTable = $ToTable; ToColumn = $ToColumn
    }
}

# ---------------------------------------------------------------------------
# 寫入：表格結構
# ---------------------------------------------------------------------------

function New-PbiTable {
    <#
      建立新表格。
        -Kind calculated  用 DAX 建計算表（資料行由引擎自動推導）
        -Kind m           用 M 腳本建表，必須同時給 -Columns
        -Kind measureHolder  建立只放量值的空殼表

      -Columns 格式：@(@{ Name='Amount'; DataType='decimal' }, @{ Name='Cat'; DataType='text' })
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [ValidateSet('calculated','m','measureHolder')][string]$Kind = 'calculated',
        [string]$Expression,
        [hashtable[]]$Columns,
        [bool]$IsHidden = $false
    )
    $body = @{ TableName = $Name; Kind = $Kind; IsHidden = $IsHidden }
    if ($Expression) { $body['Expression'] = $Expression }
    if ($Columns)    { $body['Columns']    = $Columns }
    # -Kind m 在資料保護啟用下會等使用者回答確認視窗（最多 120 秒）—— 逾時要比那個長
    Invoke-PbiApi -Path "/api/create-table" -Method POST -Body $body -TimeoutSec 180
}

function Remove-PbiTable {
    <#  刪除表格。仍有關聯線指向它時會被擋下，請先刪關聯。 #>
    param([Parameter(Mandatory)][string]$Name)
    Invoke-PbiApi -Path "/api/delete-table" -Method POST -Body @{ TableName = $Name }
}

function Rename-PbiObject {
    <#
      改名並同步改寫所有 DAX 引用（TOM 本身不會做這件事，少做就會留下壞公式）。
      ⚠️ 強烈建議先跑 -DryRun 檢視 Rewrites 清單，確認沒有誤改再實際執行。

      範例：Rename-PbiObject -Type measure -Table 量值 -OldName 銷售額 -NewName 銷售總額 -DryRun
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('measure','column','table')][string]$Type,
        [string]$Table,
        [Parameter(Mandatory)][string]$OldName,
        [Parameter(Mandatory)][string]$NewName,
        [switch]$DryRun
    )
    if ($Type -ne 'table' -and -not $Table) { throw "改 measure/column 名稱時必須指定 -Table" }
    Invoke-PbiApi -Path "/api/rename" -Method POST -Body @{
        ObjectType = $Type; TableName = $Table; OldName = $OldName; NewName = $NewName
        DryRun = [bool]$DryRun
    } -TimeoutSec 180
}

# ---------------------------------------------------------------------------
# 寫入：計算群組（時間智慧的殺手鐧）
# ---------------------------------------------------------------------------

function Get-PbiInfo {
    <#  診斷用：目前選定的實例實際開的是哪個檔案（完整路徑、PBIX/PBIP、Port） #>
    Invoke-PbiApi -Path "/api/pbi-info"
}

function Get-PbiModelProps { Invoke-PbiApi -Path "/api/model-props" }

function Set-PbiModelProps {
    <#
      設定模型層級屬性。
      ⚠️ DiscourageImplicitMeasures = $true 是建立計算群組的前提，但會讓使用者
         無法再把數值欄位直接拖進視覺自動彙總（一律得改用量值）。
         要改回 $false，必須先刪掉模型中所有計算群組。
    #>
    param(
        [bool]$DiscourageImplicitMeasures,
        [string]$Description
    )
    $body = @{}
    if ($PSBoundParameters.ContainsKey('DiscourageImplicitMeasures')) { $body['DiscourageImplicitMeasures'] = $DiscourageImplicitMeasures }
    if ($PSBoundParameters.ContainsKey('Description'))                { $body['Description'] = $Description }
    if ($body.Count -eq 0) { throw "請至少指定一個要變更的屬性" }
    Invoke-PbiApi -Path "/api/set-model-props" -Method POST -Body $body
}

function New-PbiCalcGroup {
    <#
      建立計算群組。一組計算項目就能讓所有量值支援 YTD/MTD/YoY，
      不必為每個量值手寫衍生版本。

      ⚠️ -DiscourageImplicitMeasures 會讓使用者無法再把數值欄位直接拖進視覺
         自動彙總（必須改用量值）。這是模型層級的行為改變，預設不開。
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$ColumnName = '計算項目',
        [int]$Precedence = 0,
        [switch]$DiscourageImplicitMeasures
    )
    Invoke-PbiApi -Path "/api/upsert-calc-group" -Method POST -Body @{
        TableName = $Name; ColumnName = $ColumnName; Precedence = $Precedence
        DiscourageImplicitMeasures = [bool]$DiscourageImplicitMeasures
    }
}

function Set-PbiCalcItem {
    <#  新增或更新計算項目。用 SELECTEDMEASURE() 代表被套用的量值。 #>
    param(
        [Parameter(Mandatory)][string]$Group,
        [Parameter(Mandatory)][string]$Name,
        [string]$Expression,
        [string]$FormatStringExpression,
        [int]$Ordinal
    )
    $body = @{ TableName = $Group; ItemName = $Name }
    if ($PSBoundParameters.ContainsKey('Expression'))             { $body['Expression'] = $Expression }
    if ($PSBoundParameters.ContainsKey('FormatStringExpression')) { $body['FormatStringExpression'] = $FormatStringExpression }
    if ($PSBoundParameters.ContainsKey('Ordinal'))                { $body['Ordinal'] = $Ordinal }
    Invoke-PbiApi -Path "/api/upsert-calc-item" -Method POST -Body $body
}

function Remove-PbiCalcItem {
    param(
        [Parameter(Mandatory)][string]$Group,
        [Parameter(Mandatory)][string]$Name
    )
    Invoke-PbiApi -Path "/api/delete-calc-item" -Method POST -Body @{ TableName = $Group; ItemName = $Name }
}

# ---------------------------------------------------------------------------
# 寫入：RLS 角色與 Power Query 共用運算式
# ---------------------------------------------------------------------------

function Set-PbiRole {
    <#
      建立或更新 RLS 角色。
      ⚠️ -TablePermissions 是「整組取代」：送什麼就是最終狀態，未列出的規則會被清掉。

      範例：Set-PbiRole -Name 業務員 -TablePermissions @(
                @{ TableName='FactSales'; FilterExpression="[Owner] = USERPRINCIPALNAME()" })
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [ValidateSet('none','read','readRefresh','refresh','administrator')][string]$Permission = 'read',
        [hashtable[]]$TablePermissions
    )
    $body = @{ RoleName = $Name; ModelPermission = $Permission }
    if ($TablePermissions) { $body['TablePermissions'] = $TablePermissions }
    Invoke-PbiApi -Path "/api/upsert-role" -Method POST -Body $body
}

function Remove-PbiRole {
    param([Parameter(Mandatory)][string]$Name)
    Invoke-PbiApi -Path "/api/delete-role" -Method POST -Body @{ RoleName = $Name }
}

function Set-PbiExpression {
    <#  建立或更新 Power Query 共用運算式（參數、共用函式） #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$Expression,
        [string]$Description
    )
    $body = @{ Name = $Name }
    if ($PSBoundParameters.ContainsKey('Expression'))  { $body['Expression']  = $Expression }
    if ($PSBoundParameters.ContainsKey('Description')) { $body['Description'] = $Description }
    # 資料保護啟用下會等使用者回答確認視窗（最多 120 秒）—— 逾時要比那個長
    Invoke-PbiApi -Path "/api/upsert-expression" -Method POST -Body $body -TimeoutSec 180
}

function Remove-PbiExpression {
    param([Parameter(Mandatory)][string]$Name)
    Invoke-PbiApi -Path "/api/delete-expression" -Method POST -Body @{ Name = $Name }
}

# ---------------------------------------------------------------------------
# 批次
# ---------------------------------------------------------------------------

function Invoke-PbiBatch {
    <#
      一次連線、一次 SaveChanges 套用多個操作 —— 寫 20 個量值從 20 次模型重算變成 1 次。

      預設 StopOnError=true 且不逐步存檔：任一步失敗就整批丟棄，模型維持原狀。
      -SavePerOp 會逐步存檔（失敗時前面的不會回滾），只在後面的操作依賴前面的
      結果已生效時才需要。

      範例：
        Invoke-PbiBatch @(
          @{ Op='upsert-measure'; Args=@{ TableName='量值'; MeasureName='銷售總額'; Expression='SUM(f[Amt])'; FormatString='#,0' } }
          @{ Op='upsert-measure'; Args=@{ TableName='量值'; MeasureName='訂單數';   Expression='COUNTROWS(f)' } }
        )

      可用的 Op：upsert-measure / delete-measure / move-measure / add-column /
                upsert-relationship / delete-relationship / create-table / delete-table /
                delete-column / set-column-props / rename / upsert-calc-group /
                upsert-calc-item / delete-calc-item / upsert-role / delete-role /
                upsert-expression / delete-expression / update-m / set-model-props

      update-m 也可以放進批次，但寫入只改模型那一份、不會重抓資料 —— 之後仍要讓它生效
      （Desktop 有提示就請使用者按「套用」，沒有就 Invoke-PbiRefresh -Table），再確認留下的版本。
    #>
    param(
        [Parameter(Mandatory, Position=0)][hashtable[]]$Operations,
        [bool]$StopOnError = $true,
        [switch]$SavePerOp,
        [switch]$DryRun,
        [int]$TimeoutSec = 600
    )
    Invoke-PbiApi -Path "/api/batch" -Method POST -Body @{
        Operations  = $Operations
        StopOnError = $StopOnError
        SavePerOp   = [bool]$SavePerOp
        DryRun      = [bool]$DryRun
    } -TimeoutSec $TimeoutSec
}

function Get-PbiHelp {
    <#  列出可用指令 #>
    Write-Host "   開工  Get-PbiOverview   ← 一次看完：選定的 PBI、模型大小、健檢、資料保護" -ForegroundColor Yellow
    Write-Host "   實例  Get-PbiInstances / Use-PbiInstance / Get-PbiInfo   ← 同時開多個 PBI 時先用這個" -ForegroundColor Yellow
    Write-Host "   檢查  Test-PbiBridge / Test-PbiModel / Test-PbiReport" -ForegroundColor Gray
    Write-Host "   讀取  Get-PbiSchema / Get-PbiRelationships / Get-PbiMeasures / Get-PbiRoles / Get-PbiExpressions" -ForegroundColor Gray
    Write-Host "   查詢  Invoke-Dax / Invoke-PbiDmv / Get-PbiModelStats" -ForegroundColor Gray
    Write-Host "   保護  Get-PbiProtection / Set-PbiProtection   ← 每個欄位 AI 讀不讀得到（使用者在儀表板設定）" -ForegroundColor Gray
    Write-Host "   PQ    Get-PbiMQuery / Set-PbiMQuery / Get-PbiTableProfile / Compare-PbiTableProfile" -ForegroundColor Gray
    Write-Host "   安全  New-PbiSnapshot / Get-PbiSnapshots / Restore-PbiSnapshot" -ForegroundColor Gray
    Write-Host "   生效  Save-PbiModel / Invoke-PbiRefresh" -ForegroundColor Gray
    Write-Host "   量值  Set-PbiMeasure / Remove-PbiMeasure / Move-PbiMeasure" -ForegroundColor Gray
    Write-Host "   結構  Add-PbiColumn / Set-PbiColumn / Remove-PbiColumn / New-PbiTable / Remove-PbiTable / Rename-PbiObject" -ForegroundColor Gray
    Write-Host "   關聯  Set-PbiRelationship / Remove-PbiRelationship" -ForegroundColor Gray
    Write-Host "   進階  New-PbiCalcGroup / Set-PbiCalcItem / Remove-PbiCalcItem / Set-PbiRole / Remove-PbiRole" -ForegroundColor Gray
    Write-Host "         Set-PbiExpression / Remove-PbiExpression / Get-PbiModelProps / Set-PbiModelProps" -ForegroundColor Gray
    Write-Host "   批次  Invoke-PbiBatch" -ForegroundColor Gray
}

# 載入訊息。指令清單對人是提示；對 AI 代理則是每一次工具呼叫都要重付一次的十幾行 context
# （代理的 shell 狀態不跨呼叫，每次都得重新載入本檔）。所以：
#   · 互動式視窗        → 完整清單（和以前一樣）
#   · 非互動、或被別的腳本載入 → 只印一行
#   · $env:PBI_BRIDGE_QUIET = '1' → 什麼都不印
if ($env:PBI_BRIDGE_QUIET -ne '1') {
    $pbiBrief = [bool]$MyInvocation.ScriptName -or
                ([Environment]::GetCommandLineArgs() -contains '-NonInteractive')
    if ($pbiBrief) {
        Write-Host "✅ PBI-Bridge 已載入（指令清單：Get-PbiHelp）" -ForegroundColor Green
    } else {
        Write-Host "✅ PBI-Bridge 已載入。可用指令：" -ForegroundColor Green
        Get-PbiHelp
    }
    Remove-Variable pbiBrief
}
