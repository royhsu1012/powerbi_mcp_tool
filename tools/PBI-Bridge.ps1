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
# 不保證解析得到，會變成 $null，錯誤訊息還會誤導成「appsettings.json 缺少 ApiKey」。
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
    <#  從 appsettings.json 讀取 API Key（讀一次後快取，不寫死在檔案裡） #>
    if ($script:PbiApiKey) { return $script:PbiApiKey }
    $cfgPath = Join-Path (Get-PbiRootPath) "pbibridge_csharp\appsettings.json"
    if (-not (Test-Path $cfgPath)) {
        throw "找不到 appsettings.json：$cfgPath（請由 appsettings.template.json 複製建立）"
    }
    $cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $script:PbiApiKey = $cfg.Security.ApiKey
    if (-not $script:PbiApiKey) { throw "appsettings.json 中缺少 Security.ApiKey" }
    return $script:PbiApiKey
}

function Invoke-PbiApi {
    <#  所有 API 呼叫的統一入口：處理 UTF-8 編碼與錯誤內容讀取 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('GET','POST')][string]$Method = 'GET',
        [hashtable]$Body,
        [int]$TimeoutSec = 120,
        [switch]$NoRetry,
        [hashtable]$ExtraHeaders
    )
    $headers = @{ "X-API-Key" = (Get-PbiApiKey) }
    # 帶上選定的目標實例；沒選就交給伺服器判斷（單一實例自動採用，多實例會拒絕）
    if ($script:PbiTarget) { $headers["X-PBI-Target"] = $script:PbiTarget }
    if ($ExtraHeaders) { foreach ($k in $ExtraHeaders.Keys) { $headers[$k] = $ExtraHeaders[$k] } }
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
        Write-Host "⚠️ 有多個實例但尚未選定目標 — 寫入類操作會被拒絕。請先執行 Use-PbiInstance <檔名片段或 Port>" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# 實例切換（同時開多個 PBI 時使用）
# ---------------------------------------------------------------------------

function Get-PbiInstances {
    <#
      列出所有執行中的 Power BI 實例（各自的 Port、檔案、行程）。
      一律回傳陣列 —— PS 5.1 對單一元素會攤平成物件，讓 .Count 變成 $null。
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
    Write-Host "   路徑：$($hits[0].filePath)" -ForegroundColor Gray
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
    $parts   = @("`"列數`", COUNTROWS('$Table')")

    foreach ($c in $Columns) {
        $col = $t.Columns | Where-Object { $_.Name -eq $c }
        if (-not $col) { throw "資料表 '$Table' 沒有欄位 [$c]" }
        $ref = "'$Table'[$c]"
        # COUNTBLANK/SUM 在沒有結果時回傳 BLANK，會讓「0」和「查詢失敗」長得一樣 —— 一律轉成 0
        $parts += "`"空值_$c`", COALESCE(COUNTBLANK($ref), 0)"
        $parts += "`"相異_$c`", COALESCE(DISTINCTCOUNT($ref), 0)"
        if ($numeric -contains $col.DataType) { $parts += "`"總和_$c`", COALESCE(SUM($ref), 0)" }
    }

    $r = Invoke-Dax ("EVALUATE ROW(" + ($parts -join ", ") + ")") -TimeoutSeconds $TimeoutSeconds
    $out = [ordered]@{ 資料表 = $Table; 欄位數 = @($t.Columns).Count; 取樣時間 = (Get-Date).ToString('HH:mm:ss') }
    $r.Rows[0].PSObject.Properties | ForEach-Object { $out[($_.Name -replace '^\[|\]$','')] = $_.Value }
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

      ⚠️ 資料保護：真正的把關在伺服器端（欄位層級管制 + 結果欄位掃描），
         這裡的檢查只是提早失敗、給出比較好讀的訊息，不是安全邊界。

         客戶身分欄位只能用計數類函式（DISTINCTCOUNT / COUNTROWS …）取統計量；
         金額欄位必須包在聚合函式內。被擋時伺服器會在「服務主控台」印出一組
         一次性權杖 —— AI 看不到那個視窗，只能請使用者確認後貼過來，
         再用 -DetailToken <權杖> 重送「同一句」查詢。

      -Pseudonymize：讓身分欄位可以當分組鍵，但值會在伺服器端換成穩定代號。
         用途是「按客戶分組看營收分布」這類分析 —— 需要區分客戶，不需要知道是誰。

           (Invoke-Dax 'EVALUATE SUMMARIZECOLUMNS(Customers[account name],
                          "額", SUM(Sales[amount]))' -Pseudonymize).Rows
           → ID_A3F1B2 | 12,345,678

         代號穩定（同一個客戶跨查詢都是同一個代號，可以串接分析），但推不回原值。
         真名只印在**服務主控台**，你看得到、AI 看不到 —— 要對照請看那個視窗。

         仍然擋下的用法：MAX / CONCATENATEX / SELECTCOLUMNS 取別名 ——
         那些會讓真名躲在別名欄位底下，遮罩抓不到。去敏只放行「當分組鍵」。
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$Query,
        [int]$MaxRows = 1000,
        [int]$TimeoutSeconds = 60,
        # 使用者從服務主控台取得的一次性權杖。刻意不是 [switch] ——
        # 開關能被模型自己按下，權杖不能。
        [string]$DetailToken,
        # 去敏模式：身分欄位可以當分組鍵，但值會在伺服器端換成穩定代號（ID_xxxxxx）。
        # 這個「可以自己按」是安全的 —— 它只會讓輸出更少，不會讓輸出更多。
        [switch]$Pseudonymize
    )

    if (-not $DetailToken -and -not $Pseudonymize) {
        $q = $Query -replace '\s+', ' '
        $reason = $null

        # 樣式一：直接 EVALUATE 一張資料表（整表拉回）
        if ($q -match "(?i)^\s*EVALUATE\s+('[^']+'|\[?\w+\]?)\s*$") {
            $reason = "直接 EVALUATE 整張資料表"
        }
        # 樣式二：抽樣函式且未經欄位挑選／彙總包裝
        elseif ($q -match '(?i)\b(TOPN|SAMPLE)\s*\(' -and
                $q -notmatch '(?i)\b(SELECTCOLUMNS|SUMMARIZECOLUMNS|SUMMARIZE|ROW|CALCULATETABLE\s*\(\s*SUMMARIZE)\b') {
            $reason = "使用 TOPN/SAMPLE 抽取明細列，且未以 SELECTCOLUMNS/SUMMARIZE 挑選欄位"
        }

        if ($reason) {
            throw @"
⛔ 已擋下可能外洩明細資料的查詢
   原因：$reason

   此查詢會將真實資料列送入 Claude 的 context（即傳送至雲端）。

   建議改用彙總查詢：
     EVALUATE ROW("筆數", COUNTROWS(<表>), "總額", SUM(<表>[<欄>]))

   或去識別化後再取樣：
     EVALUATE TOPN(5, SELECTCOLUMNS(<表>, "分類", [<非敏感欄>]))

   確實必須查看明細時：先向使用者說明會看到哪些欄位、幾列、為何彙總不足。
   同意後請他把服務主控台印出的一次性權杖貼給你，再用 -DetailToken <權杖> 重送。
"@
        }
    }

    $extra = if ($DetailToken) { @{ "X-PBI-Allow-Detail" = $DetailToken } } else { $null }
    Invoke-PbiApi -Path "/api/query" -Method POST -Body @{
        Query          = $Query
        MaxRows        = $MaxRows
        TimeoutSeconds = $TimeoutSeconds
        Pseudonymize   = [bool]$Pseudonymize
    } -TimeoutSec ($TimeoutSeconds + 30) -ExtraHeaders $extra
}

function Invoke-PbiDmv {
    <#
      執行 $SYSTEM DMV 查詢（效能統計、相依性追蹤）。
      這些系統檢視只含中繼資料與統計，不會回傳事實資料列。
      範例：Invoke-PbiDmv 'SELECT * FROM $SYSTEM.DISCOVER_STORAGE_TABLES'

      ⚠️ $SYSTEM.TMSCHEMA_PARTITIONS 會回傳 M 腳本（含連線字串），
         依 CLAUDE.md 規則三，非必要不要查它。
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
      模擬 Ctrl+S 存檔，並用選定實例自己的檔案（PBIX 或 PBIP）修改時間驗證是否真的存到。
      ⚠️ 會把 PBI Desktop 帶到前景，短暫搶走鍵盤焦點（SendKeys 的固有限制）。
      ⚠️ fileChanged = false 時最多再試一次，不要連按 —— 連續送合成按鍵會被防毒視為可疑行為。
    #>
    param([int]$WaitSeconds = 5)
    Invoke-PbiApi -Path "/api/save" -Method POST -Body @{ WaitSeconds = $WaitSeconds } `
                  -TimeoutSec ($WaitSeconds + 60)
}

function Invoke-PbiRefresh {
    <#
      重新整理資料，讓結構性變更生效：
        建計算表／新增計算項目        → Invoke-PbiRefresh -Table <表>
        建或改關聯線／新增計算資料行  → Invoke-PbiRefresh -RefreshType calculate
      只改量值不需要。改 M 也不需要 —— 使用者在 Power Query 編輯器「關閉並套用」時
      Power BI 會自己重整。不指定 -Table 就是整個模型；full 會重抓資料源，大表可能很久。
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
    <#  新增 DAX 計算資料行。DataType: text/int/decimal/currency/bool/date #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Expression,
        [ValidateSet('text','int','decimal','currency','bool','date')][string]$DataType = 'text'
    )
    Invoke-PbiApi -Path "/api/add-column" -Method POST -Body @{
        TableName = $Table; ColumnName = $Name; Expression = $Expression; DataType = $DataType
    }
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

      ⚠️ 這只改「模型」那一份。Power BI Desktop 的 Power Query 文件是另一份 ——
         寫完 Desktop 會顯示「查詢中有暫止的變更尚未套用」，要請使用者按「套用」。
         兩份內容不同時，套用有可能是用 Desktop 那份覆蓋回來，所以套用後一定要
         用 Get-PbiMQuery 確認留下的是這次寫入的版本。

      建議流程：
        $before = Get-PbiTableProfile <表名> -Columns <欄位…>   # 基準線
        Get-PbiMQuery <表名> -Label before                      # 留一份 .pq 備份
        Set-PbiMQuery <表名> -Expression $m                      # 寫入
        # → 請使用者到 Power BI Desktop 按「套用」
        Get-PbiMQuery <表名> -Show                               # 確認留下的是哪一版
        Compare-PbiTableProfile $before (Get-PbiTableProfile <表名> -Columns <欄位…>)
    #>
    param(
        [Parameter(Mandatory, Position=0)][string]$Table,
        [Parameter(Mandatory)][string]$Expression
    )
    # 片段會讓整張表的 M 變成不合法 —— 一律要求完整的 let ... in
    if ($Expression -notmatch '(?s)^\s*let.*in') {
        throw "M 腳本看起來不是完整的 let ... in，拒絕寫入（不要送片段）。"
    }
    Invoke-PbiApi -Path "/api/update-m" -Method POST -Body @{
        TableName = $Table; Expression = $Expression
    }
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
    Invoke-PbiApi -Path "/api/create-table" -Method POST -Body $body
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
    Invoke-PbiApi -Path "/api/upsert-expression" -Method POST -Body $body
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
                upsert-expression / delete-expression / update-m

      update-m 也可以放進批次，但寫入只改模型那一份 —— 之後仍要請使用者在
      Power BI Desktop 按「套用」，並用 Get-PbiMQuery 確認留下的是這次寫入的版本。
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

Write-Host "✅ PBI-Bridge 已載入。可用指令：" -ForegroundColor Green
Write-Host "   實例  Get-PbiInstances / Use-PbiInstance / Get-PbiInfo   ← 同時開多個 PBI 時先用這個" -ForegroundColor Yellow
Write-Host "   檢查  Test-PbiBridge / Test-PbiModel" -ForegroundColor Gray
Write-Host "   讀取  Get-PbiSchema / Get-PbiRelationships / Get-PbiMeasures / Get-PbiRoles / Get-PbiExpressions" -ForegroundColor Gray
Write-Host "   查詢  Invoke-Dax / Invoke-PbiDmv / Get-PbiModelStats" -ForegroundColor Gray
Write-Host "   PQ    Get-PbiMQuery / Set-PbiMQuery / Get-PbiTableProfile / Compare-PbiTableProfile" -ForegroundColor Gray
Write-Host "   安全  New-PbiSnapshot / Get-PbiSnapshots / Restore-PbiSnapshot" -ForegroundColor Gray
Write-Host "   生效  Save-PbiModel / Invoke-PbiRefresh" -ForegroundColor Gray
Write-Host "   量值  Set-PbiMeasure / Remove-PbiMeasure / Move-PbiMeasure" -ForegroundColor Gray
Write-Host "   結構  Add-PbiColumn / Set-PbiColumn / Remove-PbiColumn / New-PbiTable / Remove-PbiTable / Rename-PbiObject" -ForegroundColor Gray
Write-Host "   關聯  Set-PbiRelationship / Remove-PbiRelationship" -ForegroundColor Gray
Write-Host "   進階  New-PbiCalcGroup / Set-PbiCalcItem / Set-PbiRole / Set-PbiExpression" -ForegroundColor Gray
Write-Host "   批次  Invoke-PbiBatch" -ForegroundColor Gray
