using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq;
using Microsoft.AnalysisServices.Tabular;
// ⚠️ 只引入需要的型別，不整個命名空間 using：
// AdomdClient 與 Tabular 都有 Measure 型別，全域引入會造成模稜兩可
using AdomdConnection = Microsoft.AnalysisServices.AdomdClient.AdomdConnection;
using AdomdErrorResponseException = Microsoft.AnalysisServices.AdomdClient.AdomdErrorResponseException;
// Tabular 的 JsonSerializer 會和 System.Text.Json.JsonSerializer 撞名，取別名區分
using TabularSerializer = Microsoft.AnalysisServices.Tabular.JsonSerializer;
// 同樣的理由：只取要用的 JSON 型別，不整個 using System.Text.Json
using JsonDocument  = System.Text.Json.JsonDocument;
using JsonElement   = System.Text.Json.JsonElement;
using JsonValueKind = System.Text.Json.JsonValueKind;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace PBIBridgeCSharp {

    /// <summary>一個執行中的 Power BI Desktop 實例（含它專屬的 msmdsrv 連接埠）。</summary>
    public record PbiInstance(
        int     Port,
        int     MsmdsrvPid,
        int     PbiPid,
        string? FilePath,
        string? FileName,
        string? WindowTitle,
        string  Kind          // PBIX / PBIP / Unknown
    );

    // ── 請求型別 ────────────────────────────────────────────────────────────
    public record UpsertMeasureRequest(string TableName, string MeasureName, string Expression,
                                       string FormatString, string Description,
                                       string DisplayFolder, bool? IsHidden);
    public record MoveMeasureRequest(string FromTable, string MeasureName, string ToTable);
    // Expression 與 SourceColumn 擇一：前者是 DAX 計算資料行，後者是 M 輸出的來源資料行
    public record AddColumnRequest(string TableName, string ColumnName, string Expression, string DataType,
                                   string? SourceColumn = null);
    public record DeleteMeasureRequest(string TableName, string MeasureName);
    // 唯讀 DAX 查詢：讓 AI 寫完量值後能立即驗算結果
    //   AskUser  這句查詢被資料保護擋下時，在使用者的螢幕跳出確認視窗，由他決定要不要放行這一次
    public record DaxQueryRequest(string Query, int? MaxRows, int? TimeoutSeconds, bool? AskUser = null);
    // DMV 查詢：$SYSTEM 系統檢視，用於查中繼資料、儲存統計與記憶體佔用
    public record DmvQueryRequest(string Query, int? MaxRows, int? TimeoutSeconds);

    // 存檔 / 重新整理
    //   WaitSeconds 最多等幾秒（偵測到檔案寫完就提早回傳）
    //   Expect      存檔後必須出現在磁碟上的文字（例如剛寫入的量值名稱），只回傳找到與否
    //   VerifyOnly  不送 Ctrl+S，只檢查磁碟上的現況
    public record SaveRequest(int? WaitSeconds, string[]? Expect = null, bool? VerifyOnly = null);
    public record RefreshRequest(string TableName, string RefreshType);

    // 逐欄資料保護設定。Level：open / pseudonym / countOnly / aggregateOnly，或 default（拿掉逐欄設定，回到通用規則）
    public record ProtectionChange(string Table, string Column, string Level);
    public record ProtectionRequest(ProtectionChange[]? Changes, bool? Reset = null);

    // 快照 / 還原
    public record SnapshotRequest(string Label);
    public record RestoreRequest(string File, bool? DryRun, string[] Scope);

    // 關聯線
    public record RelationshipRequest(string FromTable, string FromColumn, string ToTable, string ToColumn,
                                      string FromCardinality, string ToCardinality,
                                      string CrossFilterDirection, bool? IsActive);
    public record RelationshipRefRequest(string FromTable, string FromColumn, string ToTable, string ToColumn);

    // 表格 / 資料行結構
    public record ColumnDef(string Name, string DataType, string SourceColumn);
    public record CreateTableRequest(string TableName, string Kind, string Expression, ColumnDef[] Columns, bool? IsHidden);
    public record UpdateMRequest(string TableName, string Expression);
    public record TableRefRequest(string TableName);
    public record ColumnRefRequest(string TableName, string ColumnName);
    public record RenameRequest(string ObjectType, string TableName, string OldName, string NewName, bool? DryRun);
    public record ColumnPropsRequest(string TableName, string ColumnName, string FormatString, string DisplayFolder,
                                     bool? IsHidden, string SortByColumn, string SummarizeBy,
                                     string DataCategory, string Description, string DataType);

    // 模型層級屬性
    public record ModelPropsRequest(bool? DiscourageImplicitMeasures, string Description);

    // 計算群組
    public record CalcGroupRequest(string TableName, string ColumnName, int? Precedence, bool? DiscourageImplicitMeasures);
    public record CalcItemRequest(string TableName, string ItemName, string Expression, string FormatStringExpression, int? Ordinal);
    public record CalcItemRefRequest(string TableName, string ItemName);

    // 資料列層級安全性
    public record TablePermissionDef(string TableName, string FilterExpression);
    public record RoleRequest(string RoleName, string ModelPermission, TablePermissionDef[] TablePermissions);
    public record RoleRefRequest(string RoleName);

    // Power Query 共用運算式（參數與函式）
    public record ExpressionRequest(string Name, string Expression, string Description);
    public record ExpressionRefRequest(string Name);

    // 批次
    public record BatchOp(string Op, System.Text.Json.JsonElement Args);
    public record BatchRequest(BatchOp[] Operations, bool? StopOnError, bool? SavePerOp, bool? DryRun);

    /// <summary>可預期的操作失敗（找不到物件、參數不合法），對應 4xx 而非 500。</summary>
    public class OpException : Exception {
        public int StatusCode { get; }
        public OpException(string message, int statusCode = 400) : base(message) { StatusCode = statusCode; }
    }

    class Program {

        // =====================================================================
        // 基礎設施
        // =====================================================================

        // =====================================================================
        // 多實例探索
        //
        // 一台機器上可以同時開好幾個 Power BI Desktop，每一個都會生出自己的
        // msmdsrv 子行程、各自監聽不同的 Port。舊版直接取 msmdsrv[0]，開兩個檔
        // 就會隨機挑一個來改 —— 這裡改成完整列舉並要求明確指定目標。
        //
        // 配對方式：msmdsrv.ParentProcessId 就是它所屬的 PBIDesktop.ProcessId。
        // =====================================================================

        /// <summary>取得所有 LISTENING 的 127.0.0.1 連接埠，對應到擁有它的行程。</summary>
        static Dictionary<int, int> GetListeningPortsByPid() {
            var map = new Dictionary<int, int>();
            var psi = new ProcessStartInfo("netstat", "-ano") {
                RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true
            };
            using (var p = Process.Start(psi)!) {
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                foreach (var line in output.Split('\n')) {
                    // 格式：  TCP    127.0.0.1:55820    0.0.0.0:0    LISTENING    24752
                    var parts = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length < 5 || parts[3] != "LISTENING") continue;
                    var m = Regex.Match(parts[1], @"^127\.0\.0\.1:(\d+)$");
                    if (!m.Success) continue;
                    if (!int.TryParse(parts[^1], out int pid)) continue;
                    if (!map.ContainsKey(pid)) map[pid] = int.Parse(m.Groups[1].Value);
                }
            }
            return map;
        }

        /// <summary>從 PBIDesktop 的命令列取出它開啟的檔案路徑。</summary>
        static string? ExtractOpenFile(string? commandLine) {
            if (string.IsNullOrWhiteSpace(commandLine)) return null;
            var m = Regex.Match(commandLine, @"""([^""]+\.(?:pbix|pbip))""", RegexOptions.IgnoreCase);
            if (!m.Success) m = Regex.Match(commandLine, @"(\S+\.(?:pbix|pbip))", RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : null;
        }

        /// <summary>
        /// 命令列上沒有檔案路徑時（先開 Power BI、再從裡面選檔案開啟，就會這樣），至少從視窗標題認出這份報表叫什麼。
        /// 標題只有名字、沒有路徑，所以不能拿來找檔案 —— 只用來顯示、選定目標，以及替逐欄保護設定取檔名。
        /// </summary>
        static string? TitleName(string? title) {
            if (string.IsNullOrWhiteSpace(title)) return null;
            string t = Regex.Replace(title, @"\s*[-–—]\s*Power BI Desktop\s*$", "", RegexOptions.IgnoreCase).Trim();
            return t.Length == 0 || t.Equals("Power BI Desktop", StringComparison.OrdinalIgnoreCase) ? null : t;
        }

        static List<PbiInstance> DiscoverInstances() {
            var result = new List<PbiInstance>();

            // ① msmdsrv → 父行程
            var msmdToParent = new Dictionary<int, int>();
            using (var s = new System.Management.ManagementObjectSearcher(
                       "SELECT ProcessId, ParentProcessId FROM Win32_Process WHERE Name = 'msmdsrv.exe'")) {
                foreach (var o in s.Get())
                    msmdToParent[Convert.ToInt32(o["ProcessId"])] = Convert.ToInt32(o["ParentProcessId"]);
            }
            if (msmdToParent.Count == 0) return result;

            // ② PBIDesktop 命令列 → 開啟的檔案
            var pbiFiles = new Dictionary<int, string?>();
            using (var s = new System.Management.ManagementObjectSearcher(
                       "SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = 'PBIDesktop.exe'")) {
                foreach (var o in s.Get())
                    pbiFiles[Convert.ToInt32(o["ProcessId"])] = ExtractOpenFile(o["CommandLine"]?.ToString());
            }

            // ③ 一次 netstat 建 pid → port
            var ports = GetListeningPortsByPid();

            foreach (var kv in msmdToParent) {
                if (!ports.TryGetValue(kv.Key, out int port)) continue;   // 還沒起完，跳過
                pbiFiles.TryGetValue(kv.Value, out string? file);

                string? title = null;
                try { title = Process.GetProcessById(kv.Value).MainWindowTitle; } catch { }

                result.Add(new PbiInstance(
                    Port        : port,
                    MsmdsrvPid  : kv.Key,
                    PbiPid      : kv.Value,
                    FilePath    : file,
                    FileName    : file != null ? Path.GetFileName(file) : TitleName(title),
                    WindowTitle : string.IsNullOrWhiteSpace(title) ? null : title,
                    Kind        : file == null ? "Unknown"
                                  : file.EndsWith(".pbip", StringComparison.OrdinalIgnoreCase) ? "PBIP" : "PBIX"));
            }
            return result.OrderBy(i => i.FileName ?? "~").ThenBy(i => i.Port).ToList();
        }

        static string FormatInstanceList(IEnumerable<PbiInstance> list) =>
            string.Join("\n", list.Select(i =>
                $"   · Port {i.Port}  {i.FileName ?? "(未開啟檔案)"}  [{i.Kind}]"));

        /// <summary>
        /// 決定這次請求要操作哪一個 Power BI 實例。
        ///   · X-PBI-Target 標頭：Port 數字，或檔名/路徑/視窗標題的片段（不分大小寫）
        ///   · 沒指定時：只有一個實例就直接用；有多個一律拒絕，不猜。
        /// 猜錯的代價是改到別的模型，所以寧可回錯誤也不要預設挑一個。
        /// </summary>
        static PbiInstance ResolveInstance(HttpContext ctx) {
            var all = DiscoverInstances();
            if (all.Count == 0)
                throw new OpException("⚠️ 找不到正在執行的 Power BI 檔案。請先開啟一份 PBIX 或 PBIP。", 404);

            string? target = ctx.Request.Headers["X-PBI-Target"].FirstOrDefault();

            if (string.IsNullOrWhiteSpace(target)) {
                if (all.Count == 1) return all[0];
                throw new OpException(
                    $"❌ 目前有 {all.Count} 個 Power BI 實例在執行，請指定要操作哪一個（X-PBI-Target 標頭）：\n" +
                    FormatInstanceList(all) +
                    "\n   PowerShell 用法：Use-PbiInstance <檔名片段或 Port>");
            }

            if (int.TryParse(target, out int port)) {
                return all.FirstOrDefault(i => i.Port == port)
                       ?? throw new OpException($"❌ 找不到 Port {port} 的實例。目前有：\n{FormatInstanceList(all)}", 404);
            }

            var hits = all.Where(i =>
                (i.FilePath?.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (i.WindowTitle?.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();

            if (hits.Count == 1) return hits[0];
            if (hits.Count == 0)
                throw new OpException($"❌ 沒有實例符合「{target}」。目前有：\n{FormatInstanceList(all)}", 404);
            throw new OpException(
                $"❌「{target}」同時符合 {hits.Count} 個實例，請改用 Port 精確指定：\n{FormatInstanceList(hits)}");
        }

        // 查詢與寫入不能交錯。資料保護是先讀模型的定義（量值、計算欄位）來檢查，然後才執行查詢；
        // 兩步之間如果有另一個請求把量值換掉，被檢查的就不是真正執行的那一份。
        // 所以查詢從「讀定義」到「讀完結果」全程拿讀鎖，會改動模型的請求拿寫鎖。
        static readonly ReaderWriterLockSlim ModelGate = new();
        const int GateWaitMs = 30_000;

        sealed class GateLease : IDisposable {
            readonly bool _write;
            public GateLease(bool write) { _write = write; }
            public void Dispose() { if (_write) ModelGate.ExitWriteLock(); else ModelGate.ExitReadLock(); }
        }
        static IDisposable EnterGate(bool write) {
            bool ok = write ? ModelGate.TryEnterWriteLock(GateWaitMs) : ModelGate.TryEnterReadLock(GateWaitMs);
            if (!ok) throw new OpException(write
                ? "❌ 有查詢或另一個寫入正在進行，等了 30 秒還沒結束。請稍後再試。"
                : "❌ 模型正在被寫入（重新整理或套用變更），等了 30 秒還沒結束。請稍後再試。");
            return new GateLease(write);
        }

        /// <summary>
        /// 連上執行中的模型，執行 action，視需要 SaveChanges，並統一處理錯誤回應。
        /// 所有會動到模型的端點都走這裡，確保連線與錯誤處理只有一份實作。
        /// </summary>
        static IResult RunModel(HttpContext ctx, string label, Func<Model, PbiInstance, object> action, bool save,
                                Func<object, long, object>? afterSave = null, bool? exclusive = null) {
            try {
                var inst = ResolveInstance(ctx);
                Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] {label}  →  {inst.FileName ?? "(未命名)"} :{inst.Port}");
                // 會存檔的才拿寫鎖，擋開正在檢查中的查詢。唯讀與 DryRun 不拿：沒存檔的變更只在這條連線裡，別的查詢看不到。
                // exclusive 給「save 是 false、但 action 裡面自己會存檔」的呼叫端用（批次的逐步存檔）。
                using IDisposable? gate = (exclusive ?? save) ? EnterGate(write: true) : null;
                using (var server = new Server()) {
                    server.Connect($"Data Source=localhost:{inst.Port};");
                    var model = server.Databases[0].Model;
                    var result = action(model, inst);
                    // ⚠️ 未呼叫 SaveChanges 時，暫存的變更會隨 server 釋放而丟棄 —— DryRun 靠的就是這點
                    if (save) {
                        var sw = Stopwatch.StartNew();
                        model.SaveChanges();
                        // 重新整理是在 SaveChanges 裡才真的執行。afterSave 讓端點能回報
                        // 「做完了、花了多久」，而不是在動手之前就先把回應寫好。
                        if (afterSave != null) result = afterSave(result, sw.ElapsedMilliseconds);
                    }
                    Console.WriteLine($"✅ {label} — 完成{(save ? "" : "（DryRun，未寫入）")}");
                    return Results.Ok(result);
                }
            } catch (OpException ex) {
                Console.WriteLine($"⚠️ {label} — 已拒絕: {ex.Message}");
                return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
            } catch (Exception ex) {
                Console.WriteLine($"❌ {label} — 失敗: {ex.Message}");
                return Results.Problem(detail: ex.Message, statusCode: 500);
            }
        }

        static Table FindTable(Model m, string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new OpException("❌ TableName 不可為空");
            return m.Tables.Find(name) ?? throw new OpException($"找不到表格: {name}", 404);
        }

        static Column FindColumn(Table t, string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new OpException("❌ ColumnName 不可為空");
            return t.Columns.Find(name) ?? throw new OpException($"找不到資料行: '{t.Name}'[{name}]", 404);
        }

        static DataType ParseDataType(string s) => (s ?? "").ToLower() switch {
            "text" or "string"    => DataType.String,
            "int" or "integer" or "int64" => DataType.Int64,
            "decimal" or "double" => DataType.Double,
            "currency" or "money" => DataType.Decimal,
            "bool" or "boolean"   => DataType.Boolean,
            "date" or "datetime"  => DataType.DateTime,
            "binary"              => DataType.Binary,
            _                     => DataType.String
        };

        // 剝除查詢開頭的空白與註解（// 、-- 、/* */），只為了取得第一個有效關鍵字做安全檢查。
        // 回傳值僅供判斷用，實際送往 ADOMD 的仍是使用者的原始查詢字串。
        static string StripLeadingComments(string query) {
            string s = query ?? "";
            int i = 0;
            while (i < s.Length) {
                while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
                if (i + 1 < s.Length && s[i] == '/' && s[i + 1] == '/') {
                    while (i < s.Length && s[i] != '\n') i++;
                    continue;
                }
                if (i + 1 < s.Length && s[i] == '-' && s[i + 1] == '-') {
                    while (i < s.Length && s[i] != '\n') i++;
                    continue;
                }
                if (i + 1 < s.Length && s[i] == '/' && s[i + 1] == '*') {
                    int end = s.IndexOf("*/", i + 2);
                    if (end < 0) { i = s.Length; break; }
                    i = end + 2;
                    continue;
                }
                break;
            }
            return i >= s.Length ? "" : s.Substring(i);
        }

        static object ExtractSchema(PbiInstance inst) {
            using (var server = new Server()) {
                server.Connect($"Data Source=localhost:{inst.Port};");
                var model = server.Databases[0].Model;

                var visibleTables = model.Tables
                    .Where(t => !t.Name.StartsWith("LocalDateTable_") && !t.Name.StartsWith("DateTableTemplate_"))
                    .OrderBy(t => t.Name)
                    .Select(t => new {
                        Name = t.Name,
                        IsHidden = t.IsHidden,
                        IsCalculationGroup = t.CalculationGroup != null,
                        Columns = t.Columns.Where(c => c.Type != ColumnType.RowNumber).Select(c => new {
                            Name = c.Name,
                            DataType = c.DataType.ToString(),
                            IsHidden = c.IsHidden,
                            Kind = c.Type.ToString(),
                            DisplayFolder = c.DisplayFolder,
                            FormatString = c.FormatString,
                            SortByColumn = c.SortByColumn?.Name,
                            Expression = (c as CalculatedColumn)?.Expression
                        }).OrderBy(c => c.Name).ToList(),
                        Measures = t.Measures.Select(m => new {
                            Name = m.Name,
                            Expression = m.Expression,
                            FormatString = m.FormatString,
                            Description = m.Description,
                            DisplayFolder = m.DisplayFolder,
                            IsHidden = m.IsHidden
                        }).OrderBy(m => m.Name).ToList(),
                        CalculationItems = t.CalculationGroup == null
                            ? null
                            : t.CalculationGroup.CalculationItems
                                 .OrderBy(ci => ci.Ordinal)
                                 .Select(ci => new { ci.Name, ci.Expression, ci.Ordinal })
                                 .ToList<object>(),
                        MQuery = t.Partitions.OfType<Partition>().Select(p => p.Source as MPartitionSource).FirstOrDefault(m => m != null && !string.IsNullOrWhiteSpace(m.Expression))?.Expression
                    }).ToList();

                return new {
                    ExportTime  = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    SourceFile  = inst.FileName,
                    SourcePath  = inst.FilePath,
                    Port        = inst.Port,
                    TotalTables = visibleTables.Count,
                    Tables      = visibleTables
                };
            }
        }

        // =====================================================================
        // 寫入操作（Op*）—— 單一端點與 /api/batch 共用同一份實作
        //   · 一律只改動模型物件，不呼叫 SaveChanges（由呼叫端決定何時存）
        //   · 找不到物件或參數不合法時丟 OpException
        // =====================================================================

        static string OpUpsertMeasure(Model model, UpsertMeasureRequest req) {
            var table = FindTable(model, req.TableName);
            if (string.IsNullOrWhiteSpace(req.MeasureName)) throw new OpException("❌ MeasureName 不可為空");

            var existing = table.Measures.Find(req.MeasureName);
            string verb;
            if (existing != null) {
                if (!string.IsNullOrWhiteSpace(req.Expression)) existing.Expression = req.Expression;
                if (!string.IsNullOrEmpty(req.FormatString))  existing.FormatString  = req.FormatString;
                if (!string.IsNullOrEmpty(req.Description))   existing.Description   = req.Description;
                if (req.DisplayFolder != null)                existing.DisplayFolder = req.DisplayFolder;
                if (req.IsHidden.HasValue)                    existing.IsHidden      = req.IsHidden.Value;
                verb = "已覆蓋更新";
            } else {
                if (string.IsNullOrWhiteSpace(req.Expression)) throw new OpException("❌ 新增量值時 Expression 不可為空");
                table.Measures.Add(new Measure {
                    Name          = req.MeasureName,
                    Expression    = req.Expression,
                    FormatString  = req.FormatString  ?? "",
                    Description   = req.Description   ?? "",
                    DisplayFolder = req.DisplayFolder ?? "",
                    IsHidden      = req.IsHidden ?? false
                });
                verb = "已全新建立";
            }
            return $"量值 '{req.TableName}'[{req.MeasureName}] {verb}";
        }

        static string OpDeleteMeasure(Model model, DeleteMeasureRequest req) {
            var table   = FindTable(model, req.TableName);
            var measure = table.Measures.Find(req.MeasureName)
                          ?? throw new OpException($"找不到量值: {req.MeasureName}", 404);
            table.Measures.Remove(measure);
            return $"量值 '{req.TableName}'[{req.MeasureName}] 已刪除";
        }

        static string OpMoveMeasure(Model model, MoveMeasureRequest req) {
            var fromTable = model.Tables.Find(req.FromTable)
                            ?? throw new OpException($"找不到來源表格: {req.FromTable}", 404);
            var measure = fromTable.Measures.Find(req.MeasureName)
                          ?? throw new OpException($"找不到量值: {req.MeasureName}", 404);

            string expr = measure.Expression, fmt = measure.FormatString,
                   desc = measure.Description, folder = measure.DisplayFolder;
            bool hidden = measure.IsHidden;

            // 目標表格不存在時，自動建立一張「量值專用表」（單一隱藏欄位的空殼表）
            var toTable = model.Tables.Find(req.ToTable);
            if (toTable == null) {
                toTable = CreateMeasureHolderTable(req.ToTable);
                model.Tables.Add(toTable);
            }

            fromTable.Measures.Remove(measure);
            toTable.Measures.Add(new Measure {
                Name = req.MeasureName, Expression = expr, FormatString = fmt,
                Description = desc, DisplayFolder = folder, IsHidden = hidden
            });
            return $"量值 [{req.MeasureName}] 已搬移：'{req.FromTable}' → '{req.ToTable}'";
        }

        static Table CreateMeasureHolderTable(string name) {
            var t = new Table { Name = name };
            t.Columns.Add(new DataColumn { Name = "_dummy", DataType = DataType.Int64, IsHidden = true, SourceColumn = "_dummy" });
            t.Partitions.Add(new Partition {
                Name = name,
                Source = new MPartitionSource { Expression = "#table(type table [_dummy = Int64.Type], {{0}})" }
            });
            return t;
        }

        static string OpAddColumn(Model model, AddColumnRequest req) {
            var table = FindTable(model, req.TableName);
            if (string.IsNullOrWhiteSpace(req.ColumnName)) throw new OpException("❌ ColumnName 不可為空");

            // ── 來源資料行：M 腳本新輸出的欄位 ────────────────────────────────
            // 用 API 改 M 讓查詢多輸出一欄時，模型不會自己長出那個資料行 ——
            // 結構偵測是 Desktop「套用查詢」時才做的事。沒有這一步，新欄位就只能從 M 裡拿掉。
            if (!string.IsNullOrWhiteSpace(req.SourceColumn)) {
                if (!string.IsNullOrWhiteSpace(req.Expression))
                    throw new OpException("❌ SourceColumn 與 Expression 只能擇一：前者對應 M 輸出的欄位，後者是 DAX 計算資料行");
                if (!table.Partitions.Any(p => p.Source is MPartitionSource))
                    throw new OpException($"❌ 資料表 '{req.TableName}' 沒有 M 分割區（計算表／計算群組沒有來源資料行）");
                if (table.Columns.Find(req.ColumnName) != null)
                    throw new OpException($"❌ '{req.TableName}'[{req.ColumnName}] 已存在");
                if (string.IsNullOrWhiteSpace(req.DataType))
                    throw new OpException("❌ 來源資料行必須指定 DataType（text / int / decimal / currency / bool / date），要與 M 輸出的型別一致");
                table.Columns.Add(new DataColumn {
                    Name = req.ColumnName, DataType = ParseDataType(req.DataType), SourceColumn = req.SourceColumn
                });
                return $"來源資料行 '{req.TableName}'[{req.ColumnName}] 已新增（對應 M 輸出的 [{req.SourceColumn}]）。"
                     + "⚠️ 要重新整理這張表才會有資料；M 沒有輸出這個欄位的話重新整理會失敗，屆時請刪掉這個資料行";
            }
            if (string.IsNullOrWhiteSpace(req.Expression))
                throw new OpException("❌ 計算資料行需要 Expression（DAX）。要對應 M 輸出的新欄位請改給 SourceColumn");

            // 同名計算資料行已存在 → 覆蓋公式
            var existingCol = table.Columns.Find(req.ColumnName);
            if (existingCol is CalculatedColumn calcCol) {
                calcCol.Expression = req.Expression;
                if (!string.IsNullOrWhiteSpace(req.DataType)) calcCol.DataType = ParseDataType(req.DataType);
                return $"計算資料行 '{req.TableName}'[{req.ColumnName}] 已覆蓋更新";
            }
            if (existingCol != null)
                throw new OpException($"❌ '{req.TableName}'[{req.ColumnName}] 已存在且不是計算資料行（類型：{existingCol.Type}），無法覆寫");

            table.Columns.Add(new CalculatedColumn {
                Name = req.ColumnName, Expression = req.Expression, DataType = ParseDataType(req.DataType)
            });
            return $"計算資料行 '{req.TableName}'[{req.ColumnName}] 已新增";
        }

        // ── M 腳本 ────────────────────────────────────────────────────────────
        //
        // 歷程：OpUpdateM（透過 TOM 覆寫 MPartitionSource.Expression）在 2026-08-06 移除 ——
        // 當時誤以為寫入後 Desktop 會卡在待套用、解不開 —— 2026-09-16 恢復。
        //
        // Power BI Desktop 的 Power Query 文件與 TOM 模型是**兩份獨立的東西**，用 TOM 改
        // MPartitionSource 只動到模型那份。寫入之後 Desktop 的反應有兩種，兩種都實際遇過：
        //   · 顯示「查詢中有暫止的變更尚未套用」→ 使用者按「套用」就會更新（2026-09-16）
        //   · 完全沒有提示 → 資料不會自己重抓，要對那張表跑一次 full refresh，
        //     新的 M 才會套用到資料上（2026-10-01，PBIP 專案）
        // 兩份內容不同時，Desktop 那份有可能把 API 寫進去的 M 蓋掉 —— 所以呼叫端的責任是：
        // 寫完要讓 M 生效（套用或 refresh），再確認留下的是這次寫入的版本。
        //
        // M 多輸出一個欄位時，模型不會自己長出資料行 —— 見 OpAddColumn 的 SourceColumn。
        // PBIP 還有另一條路：M 以 TMDL 檔案存在專案資料夾，可以改檔案再讓 Desktop 重載。
        // 本服務尚未實作也未驗證，要做請先在測試檔上驗證再說。
        // 讀取請用 /api/schema 的 MQuery 欄位（Get-PbiMQuery）。

        /// <summary>
        /// 覆寫資料表的 M 腳本（MPartitionSource.Expression）。
        ///
        /// ⚠️ 這只改「模型」那一份，而且寫入本身不會重抓資料。讓它生效的方式見上方說明；
        /// 生效之後呼叫端要確認留下的 M 是這次寫入的版本。
        /// </summary>
        static string OpUpdateM(Model model, UpdateMRequest req) {
            if (string.IsNullOrWhiteSpace(req.TableName)) throw new OpException("❌ 缺少 TableName");
            if (string.IsNullOrWhiteSpace(req.Expression))
                throw new OpException("❌ 缺少 Expression —— 請給完整的 let...in，不要給片段");

            var t = model.Tables.Find(req.TableName)
                    ?? throw new OpException($"找不到表格: {req.TableName}", 404);
            var part = t.Partitions.FirstOrDefault(x => x.Source is MPartitionSource)
                    ?? throw new OpException($"❌ 資料表 '{req.TableName}' 沒有 M 分割區（計算表／計算群組沒有 M）");

            var src    = (MPartitionSource)part.Source;
            var before = src.Expression ?? "";
            if (before == req.Expression) return $"'{req.TableName}' 的 M 與現有內容相同，未變更";
            src.Expression = req.Expression;
            return $"已寫入 '{req.TableName}' 的 M（{before.Length} → {req.Expression.Length} 字元）";
        }

        // ── 關聯線 ────────────────────────────────────────────────────────────

        static SingleColumnRelationship? FindRelationship(Model model, string ft, string fc, string tt, string tc) =>
            model.Relationships.OfType<SingleColumnRelationship>().FirstOrDefault(r =>
                r.FromTable.Name  == ft && r.FromColumn.Name == fc &&
                r.ToTable.Name    == tt && r.ToColumn.Name   == tc);

        static string OpUpsertRelationship(Model model, RelationshipRequest req) {
            var fromTable = FindTable(model, req.FromTable);
            var toTable   = FindTable(model, req.ToTable);
            var fromCol   = FindColumn(fromTable, req.FromColumn);
            var toCol     = FindColumn(toTable,   req.ToColumn);

            var fromCard = (req.FromCardinality ?? "many").ToLower() switch {
                "one" or "1"    => RelationshipEndCardinality.One,
                "many" or "*"   => RelationshipEndCardinality.Many,
                _ => throw new OpException($"❌ 不支援的 FromCardinality: {req.FromCardinality}（可用 one / many）")
            };
            var toCard = (req.ToCardinality ?? "one").ToLower() switch {
                "one" or "1"    => RelationshipEndCardinality.One,
                "many" or "*"   => RelationshipEndCardinality.Many,
                _ => throw new OpException($"❌ 不支援的 ToCardinality: {req.ToCardinality}（可用 one / many）")
            };
            var crossFilter = (req.CrossFilterDirection ?? "single").ToLower() switch {
                "single" or "onedirection" => CrossFilteringBehavior.OneDirection,
                "both" or "bothdirections" => CrossFilteringBehavior.BothDirections,
                "auto" or "automatic"      => CrossFilteringBehavior.Automatic,
                _ => throw new OpException($"❌ 不支援的 CrossFilterDirection: {req.CrossFilterDirection}（可用 single / both / auto）")
            };

            var existing = FindRelationship(model, req.FromTable, req.FromColumn, req.ToTable, req.ToColumn);
            if (existing != null) {
                existing.FromCardinality        = fromCard;
                existing.ToCardinality          = toCard;
                existing.CrossFilteringBehavior = crossFilter;
                existing.IsActive               = req.IsActive ?? existing.IsActive;
                return $"關聯已更新：'{req.FromTable}'[{req.FromColumn}] → '{req.ToTable}'[{req.ToColumn}]";
            }

            model.Relationships.Add(new SingleColumnRelationship {
                Name                   = Guid.NewGuid().ToString(),
                FromColumn             = fromCol,
                ToColumn               = toCol,
                FromCardinality        = fromCard,
                ToCardinality          = toCard,
                CrossFilteringBehavior = crossFilter,
                IsActive               = req.IsActive ?? true
            });
            return $"關聯已建立：'{req.FromTable}'[{req.FromColumn}] → '{req.ToTable}'[{req.ToColumn}]（{fromCard}→{toCard}, {crossFilter}）";
        }

        static string OpDeleteRelationship(Model model, RelationshipRefRequest req) {
            var rel = FindRelationship(model, req.FromTable, req.FromColumn, req.ToTable, req.ToColumn)
                      ?? throw new OpException($"找不到關聯：'{req.FromTable}'[{req.FromColumn}] → '{req.ToTable}'[{req.ToColumn}]", 404);
            model.Relationships.Remove(rel);
            return $"關聯已刪除：'{req.FromTable}'[{req.FromColumn}] → '{req.ToTable}'[{req.ToColumn}]";
        }

        // ── 表格結構 ──────────────────────────────────────────────────────────

        static string OpCreateTable(Model model, CreateTableRequest req) {
            if (string.IsNullOrWhiteSpace(req.TableName)) throw new OpException("❌ TableName 不可為空");
            if (model.Tables.Find(req.TableName) != null) throw new OpException($"❌ 表格 '{req.TableName}' 已存在");

            string kind = (req.Kind ?? "calculated").ToLower();
            Table t;

            switch (kind) {
                case "calculated" or "dax":
                    if (string.IsNullOrWhiteSpace(req.Expression)) throw new OpException("❌ 計算表需要 Expression（DAX）");
                    t = new Table { Name = req.TableName };
                    // 計算表的資料行由引擎依 DAX 結果自動推導，不要手動指定
                    t.Partitions.Add(new Partition {
                        Name = req.TableName,
                        Source = new CalculatedPartitionSource { Expression = req.Expression }
                    });
                    break;

                case "m" or "powerquery":
                    if (string.IsNullOrWhiteSpace(req.Expression)) throw new OpException("❌ M 表格需要 Expression（M 腳本）");
                    if (req.Columns == null || req.Columns.Length == 0)
                        throw new OpException("❌ M 表格必須明確指定 Columns（引擎不會自動推導 TOM 建立的 M 表結構）");
                    t = new Table { Name = req.TableName };
                    foreach (var c in req.Columns) {
                        t.Columns.Add(new DataColumn {
                            Name         = c.Name,
                            DataType     = ParseDataType(c.DataType),
                            SourceColumn = string.IsNullOrWhiteSpace(c.SourceColumn) ? c.Name : c.SourceColumn
                        });
                    }
                    t.Partitions.Add(new Partition {
                        Name = req.TableName,
                        Source = new MPartitionSource { Expression = req.Expression }
                    });
                    break;

                case "measureholder" or "holder":
                    t = CreateMeasureHolderTable(req.TableName);
                    break;

                default:
                    throw new OpException($"❌ 不支援的 Kind: {req.Kind}（可用 calculated / m / measureHolder）");
            }

            t.IsHidden = req.IsHidden ?? false;
            model.Tables.Add(t);
            return $"表格 '{req.TableName}' 已建立（{kind}）";
        }

        static string OpDeleteTable(Model model, TableRefRequest req) {
            var table = FindTable(model, req.TableName);

            // 先擋掉還有關聯線指著它的情況 —— 直接刪會讓模型進入不一致狀態
            var refs = model.Relationships.OfType<SingleColumnRelationship>()
                            .Where(r => r.FromTable.Name == req.TableName || r.ToTable.Name == req.TableName)
                            .Select(r => $"'{r.FromTable.Name}'[{r.FromColumn.Name}] → '{r.ToTable.Name}'[{r.ToColumn.Name}]")
                            .ToList();
            if (refs.Count > 0)
                throw new OpException($"❌ 表格 '{req.TableName}' 仍有 {refs.Count} 條關聯線，請先刪除：{string.Join("; ", refs)}");

            int measureCount = table.Measures.Count;
            model.Tables.Remove(table);
            return $"表格 '{req.TableName}' 已刪除（含 {measureCount} 個量值）";
        }

        static string OpDeleteColumn(Model model, ColumnRefRequest req) {
            var table = FindTable(model, req.TableName);
            var col   = FindColumn(table, req.ColumnName);

            var refs = model.Relationships.OfType<SingleColumnRelationship>()
                            .Where(r => (r.FromTable.Name == req.TableName && r.FromColumn.Name == req.ColumnName) ||
                                        (r.ToTable.Name   == req.TableName && r.ToColumn.Name   == req.ColumnName))
                            .ToList();
            if (refs.Count > 0)
                throw new OpException($"❌ 資料行 '{req.TableName}'[{req.ColumnName}] 仍被 {refs.Count} 條關聯線使用，請先刪除關聯");

            string kind = col.Type.ToString();
            table.Columns.Remove(col);
            return $"資料行 '{req.TableName}'[{req.ColumnName}] 已刪除（類型：{kind}）" +
                   (col.Type == ColumnType.Data ? "。⚠️ 這是來源資料行，下次重新整理時可能會依 M 腳本重新出現" : "");
        }

        static string OpSetColumnProps(Model model, ColumnPropsRequest req) {
            var table = FindTable(model, req.TableName);
            var col   = FindColumn(table, req.ColumnName);
            var changed = new List<string>();

            if (req.FormatString  != null) { col.FormatString  = req.FormatString;  changed.Add("FormatString"); }
            if (req.DisplayFolder != null) { col.DisplayFolder = req.DisplayFolder; changed.Add("DisplayFolder"); }
            if (req.Description   != null) { col.Description   = req.Description;   changed.Add("Description"); }
            if (req.DataCategory  != null) { col.DataCategory  = req.DataCategory;  changed.Add("DataCategory"); }
            if (req.IsHidden.HasValue)     { col.IsHidden      = req.IsHidden.Value; changed.Add("IsHidden"); }

            if (req.SortByColumn != null) {
                col.SortByColumn = req.SortByColumn == ""
                    ? null
                    : FindColumn(table, req.SortByColumn);
                changed.Add("SortByColumn");
            }
            if (req.SummarizeBy != null) {
                col.SummarizeBy = req.SummarizeBy.ToLower() switch {
                    "none"          => AggregateFunction.None,
                    "default"       => AggregateFunction.Default,
                    "sum"           => AggregateFunction.Sum,
                    "min"           => AggregateFunction.Min,
                    "max"           => AggregateFunction.Max,
                    "count"         => AggregateFunction.Count,
                    "average"       => AggregateFunction.Average,
                    "distinctcount" => AggregateFunction.DistinctCount,
                    _ => throw new OpException($"❌ 不支援的 SummarizeBy: {req.SummarizeBy}")
                };
                changed.Add("SummarizeBy");
            }
            if (req.DataType != null) {
                if (col is not CalculatedColumn)
                    throw new OpException("❌ 只有計算資料行能改資料類型；來源資料行的類型由 Power Query 決定，請改 M 腳本");
                col.DataType = ParseDataType(req.DataType);
                changed.Add("DataType");
            }

            if (changed.Count == 0) throw new OpException("❌ 沒有指定任何要變更的屬性");
            return $"資料行 '{req.TableName}'[{req.ColumnName}] 已更新：{string.Join(", ", changed)}";
        }

        // ── 改名（含 DAX 引用改寫）────────────────────────────────────────────

        /// <summary>
        /// 走訪模型中所有 DAX 運算式，套用 rewrite 後回報變更。dryRun 時只回報不寫入。
        /// TOM 改名不會自動更新引用，所以這一步是必要的 —— 少做就會留下壞掉的公式。
        /// </summary>
        static List<object> RewriteAllDax(Model model, Func<string, string> rewrite, bool dryRun) {
            var changes = new List<object>();

            void Apply(string objectPath, string kind, string? before, Action<string> setter) {
                if (string.IsNullOrEmpty(before)) return;
                string after = rewrite(before);
                if (after == before) return;
                changes.Add(new { Object = objectPath, Kind = kind, Before = before, After = after });
                if (!dryRun) setter(after);
            }

            foreach (var t in model.Tables) {
                foreach (var m in t.Measures)
                    Apply($"'{t.Name}'[{m.Name}]", "Measure", m.Expression, v => m.Expression = v);

                foreach (var c in t.Columns.OfType<CalculatedColumn>())
                    Apply($"'{t.Name}'[{c.Name}]", "CalculatedColumn", c.Expression, v => c.Expression = v);

                foreach (var p in t.Partitions.Where(p => p.Source is CalculatedPartitionSource)) {
                    var src = (CalculatedPartitionSource)p.Source;
                    Apply($"'{t.Name}' (計算表)", "CalculatedTable", src.Expression, v => src.Expression = v);
                }

                if (t.CalculationGroup != null) {
                    foreach (var ci in t.CalculationGroup.CalculationItems) {
                        Apply($"'{t.Name}'::{ci.Name}", "CalculationItem", ci.Expression, v => ci.Expression = v);
                        Apply($"'{t.Name}'::{ci.Name} (格式)", "CalculationItemFormat",
                              ci.FormatStringDefinition?.Expression,
                              v => ci.FormatStringDefinition = new FormatStringDefinition { Expression = v });
                    }
                }
            }

            foreach (var role in model.Roles)
                foreach (var tp in role.TablePermissions)
                    Apply($"角色 [{role.Name}] → '{tp.Table.Name}'", "RLSFilter",
                          tp.FilterExpression, v => tp.FilterExpression = v);

            return changes;
        }

        static string RegexEscape(string s) => Regex.Escape(s);

        static object OpRename(Model model, RenameRequest req) {
            if (string.IsNullOrWhiteSpace(req.OldName) || string.IsNullOrWhiteSpace(req.NewName))
                throw new OpException("❌ OldName / NewName 不可為空");
            if (req.OldName == req.NewName) throw new OpException("❌ 新舊名稱相同");

            bool dryRun = req.DryRun ?? false;
            string type = (req.ObjectType ?? "").ToLower();
            string old  = req.OldName, neu = req.NewName;
            Func<string, string> rewrite;
            Action doRename;
            string target;

            switch (type) {
                case "measure": {
                    var table   = FindTable(model, req.TableName);
                    var measure = table.Measures.Find(old) ?? throw new OpException($"找不到量值: {old}", 404);
                    if (model.Tables.Any(t => t.Measures.Find(neu) != null))
                        throw new OpException($"❌ 模型中已存在量值 [{neu}]（量值名稱在整個模型必須唯一）");
                    // 量值引用一律不帶表名前綴，所以只改「前面不是 ] 、' 或文字」的 [Name]
                    var rx = new Regex($@"(?<![\]\w'])\[{RegexEscape(old)}\]");
                    rewrite  = s => rx.Replace(s, $"[{neu}]");
                    doRename = () => measure.Name = neu;
                    target   = $"量值 '{table.Name}'[{old}]";
                    break;
                }
                case "column": {
                    var table = FindTable(model, req.TableName);
                    var col   = FindColumn(table, old);
                    if (table.Columns.Find(neu) != null) throw new OpException($"❌ '{table.Name}' 已有資料行 [{neu}]");
                    string tn = RegexEscape(table.Name);
                    // 帶表名前綴的引用：'Table'[Old] 或 Table[Old]
                    var rxQualified = new Regex($@"(?<prefix>'{tn}'|(?<![\w'])({tn}))\[{RegexEscape(old)}\]");
                    // 不帶前綴的 [Old]（計算資料行內常見），但要避開其他表的同名欄位 —— 一併改寫並列在報告中供人工複核
                    var rxBare = new Regex($@"(?<![\]\w'])\[{RegexEscape(old)}\]");
                    rewrite  = s => rxBare.Replace(rxQualified.Replace(s, m => m.Groups["prefix"].Value + $"[{neu}]"), $"[{neu}]");
                    doRename = () => col.Name = neu;
                    target   = $"資料行 '{table.Name}'[{old}]";
                    break;
                }
                case "table": {
                    var table = model.Tables.Find(old) ?? throw new OpException($"找不到表格: {old}", 404);
                    if (model.Tables.Find(neu) != null) throw new OpException($"❌ 已存在表格 '{neu}'");
                    string oldEsc = RegexEscape(old);
                    var rxQuoted   = new Regex($@"'{oldEsc}'");
                    // 未加引號的表名：只在後面接 [ 或屬於獨立詞彙時才算引用
                    var rxUnquoted = new Regex($@"(?<![\w'\[]){oldEsc}(?![\w'\]])");
                    rewrite  = s => rxUnquoted.Replace(rxQuoted.Replace(s, $"'{neu}'"), $"'{neu}'");
                    doRename = () => table.Name = neu;
                    target   = $"表格 '{old}'";
                    break;
                }
                default:
                    throw new OpException($"❌ 不支援的 ObjectType: {req.ObjectType}（可用 measure / column / table）");
            }

            var changes = RewriteAllDax(model, rewrite, dryRun);
            if (!dryRun) doRename();

            return new {
                message       = $"{target} → [{neu}]{(dryRun ? "（DryRun：未實際套用）" : " 已改名")}",
                DryRun        = dryRun,
                RewriteCount  = changes.Count,
                Rewrites      = changes
            };
        }

        // ── 計算群組 ──────────────────────────────────────────────────────────

        static string OpUpsertCalcGroup(Model model, CalcGroupRequest req) {
            if (string.IsNullOrWhiteSpace(req.TableName)) throw new OpException("❌ TableName 不可為空");
            string colName = string.IsNullOrWhiteSpace(req.ColumnName) ? "計算項目" : req.ColumnName;

            var table = model.Tables.Find(req.TableName);
            if (table != null) {
                if (table.CalculationGroup == null)
                    throw new OpException($"❌ 表格 '{req.TableName}' 已存在但不是計算群組");
                if (req.Precedence.HasValue) table.CalculationGroup.Precedence = req.Precedence.Value;
                return $"計算群組 '{req.TableName}' 已更新（Precedence={table.CalculationGroup.Precedence}）";
            }

            // 引擎硬性要求：沒開 DiscourageImplicitMeasures 就不給建計算群組。
            // 這是模型層級的行為改變，不自動幫使用者開，但要講清楚為什麼失敗。
            if (!model.DiscourageImplicitMeasures && req.DiscourageImplicitMeasures != true)
                throw new OpException(
                    "❌ 建立計算群組前，模型的 DiscourageImplicitMeasures 必須設為 true（引擎強制要求）。\n" +
                    "   ⚠️ 開啟後使用者將無法再把數值欄位直接拖進視覺自動彙總，一律得改用量值。\n" +
                    "   確認可接受後，請加上 -DiscourageImplicitMeasures 重新執行；\n" +
                    "   事後可用 Set-PbiModelProps -DiscourageImplicitMeasures $false 改回（需先刪掉所有計算群組）。");

            var t = new Table {
                Name = req.TableName,
                CalculationGroup = new CalculationGroup { Precedence = req.Precedence ?? 0 }
            };
            // 計算群組的屬性資料行固定對應內建的 "Name" 來源
            t.Columns.Add(new DataColumn { Name = colName, DataType = DataType.String, SourceColumn = "Name" });
            t.Partitions.Add(new Partition { Name = req.TableName, Source = new CalculationGroupSource() });
            model.Tables.Add(t);

            string extra = "";
            if (req.DiscourageImplicitMeasures == true) {
                model.DiscourageImplicitMeasures = true;
                extra = "；⚠️ 已開啟 DiscourageImplicitMeasures：使用者將無法再把數值欄位直接拖進視覺自動彙總";
            }
            return $"計算群組 '{req.TableName}' 已建立（屬性資料行：[{colName}]，Precedence={req.Precedence ?? 0}）{extra}";
        }

        static string OpUpsertCalcItem(Model model, CalcItemRequest req) {
            var table = FindTable(model, req.TableName);
            var cg = table.CalculationGroup ?? throw new OpException($"❌ '{req.TableName}' 不是計算群組");
            if (string.IsNullOrWhiteSpace(req.ItemName)) throw new OpException("❌ ItemName 不可為空");

            var item = cg.CalculationItems.Find(req.ItemName);
            string verb;
            if (item != null) {
                if (!string.IsNullOrWhiteSpace(req.Expression)) item.Expression = req.Expression;
                if (req.Ordinal.HasValue) item.Ordinal = req.Ordinal.Value;
                verb = "已更新";
            } else {
                if (string.IsNullOrWhiteSpace(req.Expression)) throw new OpException("❌ 新增計算項目時 Expression 不可為空");
                item = new CalculationItem {
                    Name = req.ItemName,
                    Expression = req.Expression,
                    Ordinal = req.Ordinal ?? cg.CalculationItems.Count
                };
                cg.CalculationItems.Add(item);
                verb = "已建立";
            }
            if (!string.IsNullOrWhiteSpace(req.FormatStringExpression))
                item.FormatStringDefinition = new FormatStringDefinition { Expression = req.FormatStringExpression };

            return $"計算項目 '{req.TableName}'::[{req.ItemName}] {verb}";
        }

        static string OpDeleteCalcItem(Model model, CalcItemRefRequest req) {
            var table = FindTable(model, req.TableName);
            var cg = table.CalculationGroup ?? throw new OpException($"❌ '{req.TableName}' 不是計算群組");
            var item = cg.CalculationItems.Find(req.ItemName) ?? throw new OpException($"找不到計算項目: {req.ItemName}", 404);
            cg.CalculationItems.Remove(item);
            return $"計算項目 '{req.TableName}'::[{req.ItemName}] 已刪除";
        }

        // ── 模型層級屬性 ──────────────────────────────────────────────────────

        static string OpSetModelProps(Model model, ModelPropsRequest req) {
            var changed = new List<string>();
            if (req.DiscourageImplicitMeasures.HasValue) {
                model.DiscourageImplicitMeasures = req.DiscourageImplicitMeasures.Value;
                changed.Add($"DiscourageImplicitMeasures={req.DiscourageImplicitMeasures.Value}");
            }
            if (req.Description != null) { model.Description = req.Description; changed.Add("Description"); }
            if (changed.Count == 0) throw new OpException("❌ 沒有指定任何要變更的屬性");
            return $"模型屬性已更新：{string.Join(", ", changed)}";
        }

        // ── 資料列層級安全性 ──────────────────────────────────────────────────

        static string OpUpsertRole(Model model, RoleRequest req) {
            if (string.IsNullOrWhiteSpace(req.RoleName)) throw new OpException("❌ RoleName 不可為空");
            var perm = (req.ModelPermission ?? "read").ToLower() switch {
                "none"          => ModelPermission.None,
                "read"          => ModelPermission.Read,
                "readrefresh"   => ModelPermission.ReadRefresh,
                "refresh"       => ModelPermission.Refresh,
                "administrator" => ModelPermission.Administrator,
                _ => throw new OpException($"❌ 不支援的 ModelPermission: {req.ModelPermission}")
            };

            var role = model.Roles.Find(req.RoleName);
            string verb;
            if (role == null) {
                role = new ModelRole { Name = req.RoleName };
                model.Roles.Add(role);
                verb = "已建立";
            } else verb = "已更新";
            role.ModelPermission = perm;

            if (req.TablePermissions != null) {
                // 整組取代：呼叫端送什麼就是最終狀態，避免殘留舊規則
                foreach (var existing in role.TablePermissions.ToList()) role.TablePermissions.Remove(existing);
                foreach (var tp in req.TablePermissions) {
                    var t = FindTable(model, tp.TableName);
                    role.TablePermissions.Add(new TablePermission { Table = t, FilterExpression = tp.FilterExpression ?? "" });
                }
            }
            return $"角色 [{req.RoleName}] {verb}（權限={perm}，資料表規則={role.TablePermissions.Count} 條）";
        }

        static string OpDeleteRole(Model model, RoleRefRequest req) {
            var role = model.Roles.Find(req.RoleName) ?? throw new OpException($"找不到角色: {req.RoleName}", 404);
            model.Roles.Remove(role);
            return $"角色 [{req.RoleName}] 已刪除";
        }

        // ── Power Query 共用運算式（參數 / 函式）──────────────────────────────

        static string OpUpsertExpression(Model model, ExpressionRequest req) {
            if (string.IsNullOrWhiteSpace(req.Name)) throw new OpException("❌ Name 不可為空");
            var expr = model.Expressions.Find(req.Name);
            string verb;
            if (expr == null) {
                if (string.IsNullOrWhiteSpace(req.Expression)) throw new OpException("❌ 新增運算式時 Expression 不可為空");
                expr = new NamedExpression { Name = req.Name, Kind = ExpressionKind.M, Expression = req.Expression };
                model.Expressions.Add(expr);
                verb = "已建立";
            } else {
                if (!string.IsNullOrWhiteSpace(req.Expression)) expr.Expression = req.Expression;
                verb = "已更新";
            }
            if (req.Description != null) expr.Description = req.Description;
            return $"共用運算式 [{req.Name}] {verb}";
        }

        static string OpDeleteExpression(Model model, ExpressionRefRequest req) {
            var expr = model.Expressions.Find(req.Name) ?? throw new OpException($"找不到共用運算式: {req.Name}", 404);
            model.Expressions.Remove(expr);
            return $"共用運算式 [{req.Name}] 已刪除";
        }

        // =====================================================================
        // 批次派送
        // =====================================================================

        static readonly System.Text.Json.JsonSerializerOptions JsonOpts =
            new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };

        static T Args<T>(System.Text.Json.JsonElement el) {
            try {
                // 用完全限定呼叫，避免為了 JsonElement.Deserialize 擴充方法而 using System.Text.Json
                // ——那會讓 JsonSerializer 與 Tabular 的同名型別撞在一起
                return System.Text.Json.JsonSerializer.Deserialize<T>(el.GetRawText(), JsonOpts)
                       ?? throw new OpException("❌ Args 不可為 null");
            } catch (System.Text.Json.JsonException ex) {
                throw new OpException($"❌ Args 格式錯誤: {ex.Message}");
            }
        }

        /// <summary>把單一操作套用到模型上。/api/batch 與各單一端點共用這個派送表。</summary>
        static object DispatchOp(Model model, string op, System.Text.Json.JsonElement args) => (op ?? "").ToLower() switch {
            "upsert-measure"      => OpUpsertMeasure(model, Args<UpsertMeasureRequest>(args)),
            "delete-measure"      => OpDeleteMeasure(model, Args<DeleteMeasureRequest>(args)),
            "move-measure"        => OpMoveMeasure(model, Args<MoveMeasureRequest>(args)),
            "add-column"          => OpAddColumn(model, Args<AddColumnRequest>(args)),
            "update-m"            => OpUpdateM(model, Args<UpdateMRequest>(args)),
            "upsert-relationship" => OpUpsertRelationship(model, Args<RelationshipRequest>(args)),
            "delete-relationship" => OpDeleteRelationship(model, Args<RelationshipRefRequest>(args)),
            "create-table"        => OpCreateTable(model, Args<CreateTableRequest>(args)),
            "delete-table"        => OpDeleteTable(model, Args<TableRefRequest>(args)),
            "delete-column"       => OpDeleteColumn(model, Args<ColumnRefRequest>(args)),
            "set-column-props"    => OpSetColumnProps(model, Args<ColumnPropsRequest>(args)),
            "rename"              => OpRename(model, Args<RenameRequest>(args)),
            "set-model-props"     => OpSetModelProps(model, Args<ModelPropsRequest>(args)),
            "upsert-calc-group"   => OpUpsertCalcGroup(model, Args<CalcGroupRequest>(args)),
            "upsert-calc-item"    => OpUpsertCalcItem(model, Args<CalcItemRequest>(args)),
            "delete-calc-item"    => OpDeleteCalcItem(model, Args<CalcItemRefRequest>(args)),
            "upsert-role"         => OpUpsertRole(model, Args<RoleRequest>(args)),
            "delete-role"         => OpDeleteRole(model, Args<RoleRefRequest>(args)),
            "upsert-expression"   => OpUpsertExpression(model, Args<ExpressionRequest>(args)),
            "delete-expression"   => OpDeleteExpression(model, Args<ExpressionRefRequest>(args)),
            _ => throw new OpException($"❌ 不支援的操作: {op}")
        };

        // =====================================================================
        // 模型健檢
        // =====================================================================

        static object ValidateModel(Model model) {
            var userTables = model.Tables
                .Where(t => !t.Name.StartsWith("LocalDateTable_") && !t.Name.StartsWith("DateTableTemplate_"))
                .ToList();

            // 一次收集所有 DAX 文字，供「未被引用」的比對使用
            var allExpressions = new List<string>();
            foreach (var t in model.Tables) {
                allExpressions.AddRange(t.Measures.Select(m => m.Expression ?? ""));
                allExpressions.AddRange(t.Columns.OfType<CalculatedColumn>().Select(c => c.Expression ?? ""));
                allExpressions.AddRange(t.Partitions.Where(p => p.Source is CalculatedPartitionSource)
                                                    .Select(p => ((CalculatedPartitionSource)p.Source).Expression ?? ""));
                if (t.CalculationGroup != null)
                    allExpressions.AddRange(t.CalculationGroup.CalculationItems.Select(ci => ci.Expression ?? ""));
            }
            foreach (var r in model.Roles)
                allExpressions.AddRange(r.TablePermissions.Select(tp => tp.FilterExpression ?? ""));
            string daxBlob = string.Join("\n", allExpressions);

            // ① 引擎回報的錯誤 —— 這是「公式壞掉」最權威的訊號
            var brokenObjects = new List<object>();
            foreach (var t in model.Tables) {
                foreach (var m in t.Measures.Where(m => !string.IsNullOrEmpty(m.ErrorMessage)))
                    brokenObjects.Add(new { Object = $"'{t.Name}'[{m.Name}]", Kind = "Measure", Error = m.ErrorMessage });
                foreach (var c in t.Columns.OfType<CalculatedColumn>().Where(c => !string.IsNullOrEmpty(c.ErrorMessage)))
                    brokenObjects.Add(new { Object = $"'{t.Name}'[{c.Name}]", Kind = "CalculatedColumn", Error = c.ErrorMessage });
                if (t.CalculationGroup != null)
                    foreach (var ci in t.CalculationGroup.CalculationItems.Where(ci => !string.IsNullOrEmpty(ci.ErrorMessage)))
                        brokenObjects.Add(new { Object = $"'{t.Name}'::{ci.Name}", Kind = "CalculationItem", Error = ci.ErrorMessage });
            }

            var rels = model.Relationships.OfType<SingleColumnRelationship>().ToList();

            var bidirectional = rels
                .Where(r => r.CrossFilteringBehavior == CrossFilteringBehavior.BothDirections)
                .Select(r => new { Path = $"'{r.FromTable.Name}'[{r.FromColumn.Name}] ↔ '{r.ToTable.Name}'[{r.ToColumn.Name}]",
                                   Note = "雙向篩選可能造成模稜兩可的篩選路徑與效能問題" })
                .ToList<object>();

            var inactive = rels.Where(r => !r.IsActive)
                .Select(r => new { Path = $"'{r.FromTable.Name}'[{r.FromColumn.Name}] → '{r.ToTable.Name}'[{r.ToColumn.Name}]",
                                   Note = "停用中，需要 USERELATIONSHIP 才會生效" })
                .ToList<object>();

            var manyToMany = rels
                .Where(r => r.FromCardinality == RelationshipEndCardinality.Many && r.ToCardinality == RelationshipEndCardinality.Many)
                .Select(r => new { Path = $"'{r.FromTable.Name}'[{r.FromColumn.Name}] ↔ '{r.ToTable.Name}'[{r.ToColumn.Name}]",
                                   Note = "多對多關聯，請確認彙總結果符合預期" })
                .ToList<object>();

            var noFormat = userTables.SelectMany(t => t.Measures
                    .Where(m => string.IsNullOrWhiteSpace(m.FormatString) && !m.IsHidden)
                    .Select(m => $"'{t.Name}'[{m.Name}]"))
                .ToList();

            var duplicateMeasures = userTables
                .SelectMany(t => t.Measures.Select(m => new { Table = t.Name, m.Name }))
                .GroupBy(x => x.Name).Where(g => g.Count() > 1)
                .Select(g => new { Measure = g.Key, Tables = g.Select(x => x.Table).ToList() })
                .ToList<object>();

            var relatedTables = new HashSet<string>(rels.SelectMany(r => new[] { r.FromTable.Name, r.ToTable.Name }));
            var islandTables = userTables
                .Where(t => t.CalculationGroup == null
                            && !relatedTables.Contains(t.Name)
                            && t.Columns.Count(c => c.Type != ColumnType.RowNumber) > 1)  // 排除量值專用空殼表
                .Select(t => t.Name).ToList();

            int autoDateTables = model.Tables.Count(t => t.Name.StartsWith("LocalDateTable_") || t.Name.StartsWith("DateTableTemplate_"));

            // ② 可能沒人用的資料行（啟發式：名稱從未出現在任何 DAX 中，也沒被關聯或排序引用）
            var usedByRelationship = new HashSet<string>(rels.SelectMany(r => new[] {
                $"{r.FromTable.Name}|{r.FromColumn.Name}", $"{r.ToTable.Name}|{r.ToColumn.Name}" }));
            var usedBySort = new HashSet<string>(model.Tables.SelectMany(t =>
                t.Columns.Where(c => c.SortByColumn != null).Select(c => $"{t.Name}|{c.SortByColumn.Name}")));

            var possiblyUnused = new List<string>();
            foreach (var t in userTables) {
                if (t.CalculationGroup != null) continue;
                foreach (var c in t.Columns) {
                    if (c.Type == ColumnType.RowNumber || c.IsHidden) continue;
                    string key = $"{t.Name}|{c.Name}";
                    if (usedByRelationship.Contains(key) || usedBySort.Contains(key)) continue;
                    if (daxBlob.Contains($"[{c.Name}]")) continue;
                    possiblyUnused.Add($"'{t.Name}'[{c.Name}]");
                }
            }

            // ③ 用了彙總函式的計算資料行 —— 通常應該改寫成量值（省記憶體、隨篩選變動）
            var aggRx = new Regex(@"\b(SUM|SUMX|AVERAGE|AVERAGEX|COUNT|COUNTA|COUNTAX|COUNTX|COUNTROWS|DISTINCTCOUNT|MIN|MINX|MAX|MAXX)\s*\(",
                                  RegexOptions.IgnoreCase);
            var calcColsShouldBeMeasures = userTables.SelectMany(t => t.Columns.OfType<CalculatedColumn>()
                    .Where(c => aggRx.IsMatch(c.Expression ?? ""))
                    .Select(c => $"'{t.Name}'[{c.Name}]"))
                .ToList();

            var findings = new {
                BrokenObjects              = brokenObjects,
                BiDirectionalRelationships = bidirectional,
                ManyToManyRelationships    = manyToMany,
                InactiveRelationships      = inactive,
                DuplicateMeasureNames      = duplicateMeasures,
                MeasuresWithoutFormat      = noFormat,
                IslandTables               = islandTables,
                AutoDateTableCount         = autoDateTables,
                PossiblyUnusedColumns      = possiblyUnused,
                CalcColumnsUsingAggregation = calcColsShouldBeMeasures
            };

            int issueCount = brokenObjects.Count + bidirectional.Count + manyToMany.Count + inactive.Count
                           + duplicateMeasures.Count + noFormat.Count + islandTables.Count
                           + possiblyUnused.Count + calcColsShouldBeMeasures.Count
                           + (autoDateTables > 0 ? 1 : 0);

            return new {
                CheckedAt   = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                TotalTables = userTables.Count,
                TotalMeasures = userTables.Sum(t => t.Measures.Count),
                TotalRelationships = rels.Count,
                IssueCount  = issueCount,
                Findings    = findings,
                Notes = new[] {
                    "BrokenObjects 來自引擎回報，是唯一「確定壞掉」的項目，其餘皆為建議。",
                    "PossiblyUnusedColumns 為啟發式判斷：只比對 DAX 文字，未涵蓋報表視覺中的使用情形，刪除前務必人工確認。",
                    autoDateTables > 0 ? $"偵測到 {autoDateTables} 張自動日期表，會顯著膨脹模型大小；建議關閉「自動日期/時間」並改用自建日期表。" : "未偵測到自動日期表，很好。"
                }
            };
        }

        // =====================================================================
        // 報表健檢（PBIP 的報表資料夾）
        // =====================================================================

        /// <summary>
        /// 檢查 PBIP 報表資料夾裡的檔案：JSON 能不能解析、引用的欄位在模型裡存不存在、
        /// 互動設定有沒有指向不存在的視覺、視覺有沒有超出頁面或互相壓到。
        ///
        /// 用途：直接改 visual.json 之後、請使用者在 Power BI 按「接受變更」之前先跑一次。
        /// 改壞的檔案與打錯的欄位名 Power BI 不一定會報錯 —— 常常只是那個視覺變成空白，
        /// 要等人翻到那一頁才會發現。
        ///
        /// 回傳的只有結構資訊：頁面名稱、視覺類型與標題、欄位名稱、座標。
        /// 篩選條件裡的常值（可能是客戶名稱）不會被讀出來，錯誤訊息也只給行號、不引用檔案內容。
        /// </summary>
        static object ValidateReport(Model model, PbiInstance inst) {
            if (string.IsNullOrEmpty(inst.FilePath))
                throw new OpException("❌ 服務不知道這份報表的檔案路徑（先開 Power BI 再從裡面選檔案，或是還沒存過的新報表），找不到報表檔案可以檢查。"
                                    + "請使用者關掉 Power BI、改用雙擊 .pbip 的方式開啟。");
            if (inst.Kind != "PBIP")
                throw new OpException($"❌ 報表健檢只適用 PBIP（報表以文字檔存在專案資料夾裡）。目前開啟的是 {inst.Kind}。");

            var scope     = ResolveSaveScope(inst.FilePath);
            var reportDir = scope.Folders.FirstOrDefault(f =>
                                f.TrimEnd('\\', '/').EndsWith(".Report", StringComparison.OrdinalIgnoreCase))
                            ?? throw new OpException($"❌ 找不到這個 PBIP 的 .Report 資料夾（{scope.Problem ?? "預期在 .pbip 旁邊"}）", 404);
            var pagesDir  = Path.Combine(reportDir, "definition", "pages");
            if (!Directory.Exists(pagesDir))
                throw new OpException("❌ 這份報表不是分檔格式（沒有 definition\\pages 資料夾）。舊的單一 report.json 格式不在健檢範圍。");

            var errors   = new List<object>();
            var warnings = new List<object>();
            void Err(string page, string? visual, string kind, string detail) =>
                errors.Add(new { Page = page, Visual = visual, Kind = kind, Detail = detail });
            void Warn(string page, string? visual, string kind, string detail) =>
                warnings.Add(new { Page = page, Visual = visual, Kind = kind, Detail = detail });

            JsonDocument? Load(string file, string page, string? visual) {
                try {
                    return JsonDocument.Parse(File.ReadAllText(file));
                } catch (System.Text.Json.JsonException ex) {
                    // 只給位置，不帶 ex.Message —— 那裡面可能夾著檔案內容的片段
                    Err(page, visual, "JSON 無法解析",
                        $"{Path.GetFileName(file)} 第 {(ex.LineNumber ?? 0) + 1} 行附近格式錯誤（多了或少了逗號、括號、引號）");
                    return null;
                } catch (Exception ex) {
                    Err(page, visual, "檔案讀取失敗", $"{Path.GetFileName(file)}：{ex.GetType().Name}");
                    return null;
                }
            }

            string? CheckField(string kind, string entity, string prop) {
                var t = model.Tables.Find(entity);
                if (t == null) return $"資料表 '{entity}' 不存在（引用的是 [{prop}]）";
                bool isColumn = t.Columns.Find(prop) != null, isMeasure = t.Measures.Find(prop) != null;
                if (kind == "Column" && !isColumn)
                    return isMeasure ? $"'{entity}'[{prop}] 是量值，卻被當成資料行引用" : $"資料行 '{entity}'[{prop}] 不存在";
                if (kind == "Measure" && !isMeasure)
                    return isColumn ? $"'{entity}'[{prop}] 是資料行，卻被當成量值引用" : $"量值 '{entity}'[{prop}] 不存在";
                if (kind == "Hierarchy" && t.Hierarchies.Find(prop) == null)
                    return $"階層 '{entity}'[{prop}] 不存在";
                return null;
            }

            // 一個欄位引用長這樣：{ "Column": { "Expression": { "SourceRef": { "Entity": "表" } }, "Property": "欄" } }
            // 篩選條件裡 SourceRef 用的是別名（"Source": "d"），別名定義在同一層的 From 陣列。
            static (string Entity, string Prop)? ReadFieldRef(string kind, JsonElement obj, Dictionary<string, string>? aliases) {
                string key = kind == "Hierarchy" ? "Hierarchy" : "Property";
                if (obj.ValueKind != JsonValueKind.Object) return null;
                if (!obj.TryGetProperty("Expression", out var ex) || ex.ValueKind != JsonValueKind.Object) return null;
                if (!ex.TryGetProperty("SourceRef", out var sr) || sr.ValueKind != JsonValueKind.Object) return null;
                if (!obj.TryGetProperty(key, out var prop) || prop.ValueKind != JsonValueKind.String) return null;
                string? entity = null;
                if (sr.TryGetProperty("Entity", out var en) && en.ValueKind == JsonValueKind.String) entity = en.GetString();
                else if (aliases != null && sr.TryGetProperty("Source", out var so) && so.ValueKind == JsonValueKind.String)
                    aliases.TryGetValue(so.GetString()!, out entity);
                return entity == null ? null : (entity, prop.GetString()!);
            }

            void Walk(JsonElement el, Dictionary<string, string>? aliases, Action<string, string, string> onRef, int depth = 0) {
                if (depth > 64) return;
                if (el.ValueKind == JsonValueKind.Array) {
                    foreach (var x in el.EnumerateArray()) Walk(x, aliases, onRef, depth + 1);
                    return;
                }
                if (el.ValueKind != JsonValueKind.Object) return;

                if (el.TryGetProperty("From", out var from) && from.ValueKind == JsonValueKind.Array) {
                    var map = aliases == null ? new Dictionary<string, string>() : new Dictionary<string, string>(aliases);
                    foreach (var f in from.EnumerateArray())
                        if (f.ValueKind == JsonValueKind.Object
                            && f.TryGetProperty("Name", out var n)   && n.ValueKind == JsonValueKind.String
                            && f.TryGetProperty("Entity", out var e) && e.ValueKind == JsonValueKind.String)
                            map[n.GetString()!] = e.GetString()!;
                    aliases = map;
                }
                foreach (var p in el.EnumerateObject()) {
                    if (p.Name == "Column" || p.Name == "Measure" || p.Name == "Hierarchy") {
                        var r = ReadFieldRef(p.Name, p.Value, aliases);
                        if (r != null) onRef(p.Name, r.Value.Entity, r.Value.Prop);
                    }
                    Walk(p.Value, aliases, onRef, depth + 1);
                }
            }

            static double Num(JsonElement obj, string name) =>
                obj.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetDouble() : 0;

            static string? Str(JsonElement obj, string name) =>
                obj.ValueKind == JsonValueKind.Object && obj.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
                    ? v.GetString() : null;

            // 視覺的標題文字（讓人認得出是哪一個視覺；沒有標題就用類型 + 代號）
            static string? ReadTitle(JsonElement visual) {
                try {
                    if (!visual.TryGetProperty("visualContainerObjects", out var vco) || vco.ValueKind != JsonValueKind.Object) return null;
                    if (!vco.TryGetProperty("title", out var title) || title.ValueKind != JsonValueKind.Array) return null;
                    foreach (var item in title.EnumerateArray()) {
                        if (item.ValueKind == JsonValueKind.Object
                            && item.TryGetProperty("properties", out var props) && props.ValueKind == JsonValueKind.Object
                            && props.TryGetProperty("text", out var text)       && text.ValueKind == JsonValueKind.Object
                            && text.TryGetProperty("expr", out var expr)        && expr.ValueKind == JsonValueKind.Object
                            && expr.TryGetProperty("Literal", out var lit)      && lit.ValueKind == JsonValueKind.Object) {
                            string? v = Str(lit, "Value");
                            if (string.IsNullOrEmpty(v)) continue;
                            if (v!.Length >= 2 && v[0] == '\'' && v[^1] == '\'') v = v.Substring(1, v.Length - 2).Replace("''", "'");
                            return v.Length > 40 ? v.Substring(0, 40) + "…" : v;
                        }
                    }
                } catch { }
                return null;
            }

            // ── 頁面清單 ─────────────────────────────────────────────────────
            var pageDirs = Directory.GetDirectories(pagesDir)
                                    .ToDictionary(d => Path.GetFileName(d), d => d, StringComparer.OrdinalIgnoreCase);
            var order = new List<string>();
            var pagesJson = Path.Combine(pagesDir, "pages.json");
            if (!File.Exists(pagesJson)) {
                Err("(報表)", null, "缺少檔案", "pages\\pages.json 不存在");
            } else {
                using var doc = Load(pagesJson, "(報表)", null);
                if (doc != null) {
                    if (doc.RootElement.TryGetProperty("pageOrder", out var po) && po.ValueKind == JsonValueKind.Array)
                        order = po.EnumerateArray().Where(x => x.ValueKind == JsonValueKind.String).Select(x => x.GetString()!).ToList();
                    foreach (var id in order.Where(id => !pageDirs.ContainsKey(id)))
                        Err("(報表)", null, "頁面不存在", $"pages.json 的 pageOrder 列了 {id}，但沒有這個頁面資料夾");
                    foreach (var id in pageDirs.Keys.Where(id => !order.Contains(id, StringComparer.OrdinalIgnoreCase)))
                        Warn("(報表)", null, "頁面未列入順序", $"頁面資料夾 {id} 不在 pages.json 的 pageOrder 裡");
                    string? active = Str(doc.RootElement, "activePageName");
                    if (active != null && !pageDirs.ContainsKey(active))
                        Err("(報表)", null, "頁面不存在", $"activePageName 指向不存在的頁面 {active}");
                }
            }

            // ── 報表層篩選 ───────────────────────────────────────────────────
            var reportJson = Path.Combine(reportDir, "definition", "report.json");
            if (File.Exists(reportJson)) {
                using var doc = Load(reportJson, "(報表)", null);
                if (doc != null) {
                    var seen = new HashSet<string>();
                    Walk(doc.RootElement, null, (kind, entity, prop) => {
                        var why = CheckField(kind, entity, prop);
                        if (why != null && seen.Add(why)) Err("(報表)", null, "報表層引用不存在的欄位", why);
                    });
                }
            }

            // 裝飾用的視覺：底板、線條、圖片。它們本來就是拿來疊在別人底下的，不列入重疊檢查
            var decorative = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "shape", "basicShape", "image" };
            int visualCount = 0;

            foreach (var kv in pageDirs.OrderBy(k => {
                         int i = order.FindIndex(o => string.Equals(o, k.Key, StringComparison.OrdinalIgnoreCase));
                         return i < 0 ? int.MaxValue : i; })) {
                string pageName = kv.Key;
                double pageW = 1280, pageH = 720;
                var interactions = new List<(string Source, string Target)>();

                var pageFile = Path.Combine(kv.Value, "page.json");
                if (!File.Exists(pageFile)) {
                    Err(pageName, null, "缺少檔案", "page.json 不存在");
                } else {
                    using var doc = Load(pageFile, pageName, null);
                    if (doc != null) {
                        var r = doc.RootElement;
                        pageName = Str(r, "displayName") ?? pageName;
                        if (r.TryGetProperty("width", out var w)  && w.ValueKind == JsonValueKind.Number) pageW = w.GetDouble();
                        if (r.TryGetProperty("height", out var h) && h.ValueKind == JsonValueKind.Number) pageH = h.GetDouble();
                        if (r.TryGetProperty("visualInteractions", out var vi) && vi.ValueKind == JsonValueKind.Array)
                            foreach (var x in vi.EnumerateArray()) {
                                string? s = Str(x, "source"), t = Str(x, "target");
                                if (s != null && t != null) interactions.Add((s, t));
                            }
                        var seen = new HashSet<string>();
                        string pn = pageName;
                        Walk(r, null, (kind, entity, prop) => {
                            var why = CheckField(kind, entity, prop);
                            if (why != null && seen.Add(why)) Err(pn, null, "頁面層引用不存在的欄位", why);
                        });
                    }
                }

                var names = new HashSet<string>(StringComparer.Ordinal);
                var rects = new List<(string Label, double X, double Y, double W, double H)>();
                var visualsDir = Path.Combine(kv.Value, "visuals");
                if (!Directory.Exists(visualsDir)) continue;

                foreach (var vdir in Directory.GetDirectories(visualsDir)) {
                    string folder = Path.GetFileName(vdir);
                    var vf = Path.Combine(vdir, "visual.json");
                    if (!File.Exists(vf)) { Warn(pageName, folder, "缺少檔案", "視覺資料夾裡沒有 visual.json"); continue; }
                    using var doc = Load(vf, pageName, folder);
                    if (doc == null) continue;
                    visualCount++;

                    var r = doc.RootElement;
                    string name = Str(r, "name") ?? folder;
                    bool isGroup = r.TryGetProperty("visualGroup", out _);
                    string type = isGroup ? "群組" : "(未知類型)";
                    string? title = null;
                    bool hasVisual = r.TryGetProperty("visual", out var vis) && vis.ValueKind == JsonValueKind.Object;
                    if (hasVisual) {
                        type  = Str(vis, "visualType") ?? type;
                        title = ReadTitle(vis);
                    }
                    string label = title != null ? $"{type}「{title}」" : $"{type} {name}";
                    if (!names.Add(name)) Err(pageName, label, "視覺名稱重複", $"同一頁有兩個視覺的 name 都是 {name}");

                    // 欄位引用
                    var seen = new HashSet<string>();
                    Walk(r, null, (kind, entity, prop) => {
                        var why = CheckField(kind, entity, prop);
                        if (why != null && seen.Add(why)) Err(pageName, label, "引用不存在的欄位", why);
                    });

                    // queryRef：格式設定（條件式底色、欄寬…）是用 queryRef 對應欄位的。
                    // 換了欄位卻沒跟著改 queryRef，或格式設定還指著已經拿掉的欄位，格式就默默失效。
                    if (hasVisual) {
                        var queryRefs = new HashSet<string>(StringComparer.Ordinal);
                        if (vis.TryGetProperty("query", out var q) && q.ValueKind == JsonValueKind.Object
                            && q.TryGetProperty("queryState", out var qs) && qs.ValueKind == JsonValueKind.Object) {
                            foreach (var role in qs.EnumerateObject()) {
                                if (role.Value.ValueKind != JsonValueKind.Object
                                    || !role.Value.TryGetProperty("projections", out var projs) || projs.ValueKind != JsonValueKind.Array) continue;
                                foreach (var proj in projs.EnumerateArray()) {
                                    string? qr = Str(proj, "queryRef");
                                    if (qr == null) continue;
                                    queryRefs.Add(qr);
                                    if (!proj.TryGetProperty("field", out var field) || field.ValueKind != JsonValueKind.Object) continue;
                                    foreach (var kind in new[] { "Column", "Measure" }) {
                                        if (!field.TryGetProperty(kind, out var fo)) continue;
                                        var fr = ReadFieldRef(kind, fo, null);
                                        if (fr != null && qr != $"{fr.Value.Entity}.{fr.Value.Prop}")
                                            Warn(pageName, label, "queryRef 與欄位不一致",
                                                 $"欄位是 {fr.Value.Entity}.{fr.Value.Prop}，queryRef 卻是 {qr}");
                                    }
                                }
                            }
                        }
                        if (queryRefs.Count > 0 && vis.TryGetProperty("objects", out var objs)) {
                            var orphan = new HashSet<string>();
                            void FindSelectors(JsonElement el, int depth = 0) {
                                if (depth > 32) return;
                                if (el.ValueKind == JsonValueKind.Array) { foreach (var x in el.EnumerateArray()) FindSelectors(x, depth + 1); return; }
                                if (el.ValueKind != JsonValueKind.Object) return;
                                if (el.TryGetProperty("selector", out var sel) && sel.ValueKind == JsonValueKind.Object) {
                                    string? md = Str(sel, "metadata");
                                    if (md != null && !queryRefs.Contains(md)) orphan.Add(md);
                                }
                                foreach (var p in el.EnumerateObject()) FindSelectors(p.Value, depth + 1);
                            }
                            FindSelectors(objs);
                            foreach (var md in orphan)
                                Warn(pageName, label, "格式設定指向不在這個視覺裡的欄位", $"{md}（這個視覺的欄位裡沒有它，設定不會生效）");
                        }
                    }

                    // 布林欄位上的視覺層篩選：2026-09-24 遇過 Power BI 存檔時把它整個丟掉、而且沒有任何提示
                    if (r.TryGetProperty("filterConfig", out var fc) && fc.ValueKind == JsonValueKind.Object
                        && fc.TryGetProperty("filters", out var filters) && filters.ValueKind == JsonValueKind.Array) {
                        foreach (var f in filters.EnumerateArray()) {
                            if (f.ValueKind != JsonValueKind.Object || !f.TryGetProperty("filter", out _)) continue;
                            if (!f.TryGetProperty("field", out var ff) || ff.ValueKind != JsonValueKind.Object
                                || !ff.TryGetProperty("Column", out var fcol)) continue;
                            var fr = ReadFieldRef("Column", fcol, null);
                            if (fr == null) continue;
                            var col = model.Tables.Find(fr.Value.Entity)?.Columns.Find(fr.Value.Prop);
                            if (col != null && col.DataType == DataType.Boolean)
                                Warn(pageName, label, "篩選條件用了布林欄位",
                                     $"'{fr.Value.Entity}'[{fr.Value.Prop}] 是 True/False 欄位。曾遇過 Power BI 存檔時把這種視覺層篩選丟掉 —— 改用文字欄位比較保險");
                        }
                    }

                    // 位置
                    bool hidden  = r.TryGetProperty("isHidden", out var ih) && ih.ValueKind == JsonValueKind.True;
                    bool grouped = r.TryGetProperty("parentGroupName", out _);
                    if (!r.TryGetProperty("position", out var pos) || pos.ValueKind != JsonValueKind.Object) {
                        if (!isGroup) Warn(pageName, label, "缺少位置", "visual.json 沒有 position");
                        continue;
                    }
                    if (hidden || grouped || isGroup) continue;      // 群組內的座標是相對群組的，不在這裡判斷
                    double x = Num(pos, "x"), y = Num(pos, "y"), vw = Num(pos, "width"), vh = Num(pos, "height");
                    if (vw < 1 || vh < 1)
                        Warn(pageName, label, "大小為零", $"寬 {vw:0}、高 {vh:0}");
                    else if (x < -0.5 || y < -0.5 || x + vw > pageW + 0.5 || y + vh > pageH + 0.5)
                        Warn(pageName, label, "超出頁面", $"位置 ({x:0}, {y:0})、大小 {vw:0}×{vh:0}，頁面是 {pageW:0}×{pageH:0}");
                    if (!decorative.Contains(type) && vw >= 1 && vh >= 1) rects.Add((label, x, y, vw, vh));
                }

                foreach (var (s, t) in interactions) {
                    if (!names.Contains(s)) Warn(pageName, null, "互動設定指向不存在的視覺", $"source = {s}");
                    if (!names.Contains(t)) Warn(pageName, null, "互動設定指向不存在的視覺", $"target = {t}");
                }

                // 重疊：只抓「部分重疊」。一個完全落在另一個裡面，當成刻意的疊放（例如卡片裡再放一個小標籤）
                for (int i = 0; i < rects.Count; i++) {
                    for (int j = i + 1; j < rects.Count; j++) {
                        var a = rects[i]; var b = rects[j];
                        double ow = Math.Min(a.X + a.W, b.X + b.W) - Math.Max(a.X, b.X);
                        double oh = Math.Min(a.Y + a.H, b.Y + b.H) - Math.Max(a.Y, b.Y);
                        if (ow <= 1 || oh <= 1) continue;
                        bool aInB = a.X >= b.X - 1 && a.Y >= b.Y - 1 && a.X + a.W <= b.X + b.W + 1 && a.Y + a.H <= b.Y + b.H + 1;
                        bool bInA = b.X >= a.X - 1 && b.Y >= a.Y - 1 && b.X + b.W <= a.X + a.W + 1 && b.Y + b.H <= a.Y + a.H + 1;
                        if (aInB || bInA) continue;
                        Warn(pageName, a.Label, "視覺互相重疊", $"與 {b.Label} 重疊 {ow:0}×{oh:0}");
                    }
                }
            }

            return new {
                CheckedAt    = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                Report       = Path.GetFileName(reportDir.TrimEnd('\\', '/')),
                Pages        = pageDirs.Count,
                Visuals      = visualCount,
                ErrorCount   = errors.Count,
                WarningCount = warnings.Count,
                Errors       = errors.Take(200).ToList(),
                Warnings     = warnings.Take(200).ToList(),
                Notes = new[] {
                    "Errors 會讓視覺壞掉或變成空白；Warnings 是值得看一眼的版面與設定問題，不一定是錯。",
                    "檢查的是「磁碟上的報表檔案」對「Power BI 記憶體中的模型」。改完檔案、使用者還沒按接受變更之前，畫面上看到的仍是舊版。",
                    "查不到的：文字有沒有被擠壓、配色、圖表好不好讀 —— 那些要看畫面。書籤與行動版配置不在檢查範圍。"
                }
            };
        }

        // =====================================================================
        // 存檔（模擬 Ctrl+S）
        // =====================================================================

        /// <summary>
        /// 存檔驗證要看哪些檔案。
        ///   · .pbix 是單一檔案 → 只看它
        ///   · .pbip 存檔時寫的是它的 .Report 與 .SemanticModel 兩個資料夾 → 只看這兩個
        ///
        /// 以前是「掃 .pbip 所在的整個資料夾」，那有兩個問題：.pbip 放在桌面或下載資料夾時
        /// 等於遞迴掃過整個桌面；而且那幾秒內任何不相干的檔案被寫入（瀏覽器下載、雲端同步、
        /// 另一個專案存檔）都會被當成「存檔成功」。
        /// 拿不到範圍時一併回報原因 —— 靜靜當成「沒有變更」會讓呼叫端誤判。
        /// </summary>
        sealed class SaveScope {
            public string       BaseDir = "";      // .pbix / .pbip 所在資料夾，回報相對路徑用
            public List<string> Files   = new();   // 直接比對的單一檔案
            public List<string> Folders = new();   // 遞迴比對的資料夾 —— 只會是這個專案自己的
            public string?      Problem;           // 無法驗證的原因
        }

        // Power BI 專案資料夾的固定字尾（.Dataset 是語意模型資料夾的舊名）。
        // 指標指到不是這種名字的資料夾就不採用 —— 範圍永遠不會擴大到專案以外。
        static readonly string[] PbipFolderSuffixes = { ".Report", ".SemanticModel", ".Dataset" };

        static bool IsPbipFolder(string? path) =>
            !string.IsNullOrEmpty(path) && Directory.Exists(path) &&
            PbipFolderSuffixes.Any(s => path!.TrimEnd('\\', '/').EndsWith(s, StringComparison.OrdinalIgnoreCase));

        /// <summary>沿著鍵讀出 JSON 裡的一個字串；遇到陣列就取第一個有這個鍵的元素。讀不到回 null。</summary>
        static string? ReadJsonString(string file, params string[] keys) {
            try {
                if (!File.Exists(file)) return null;
                using var doc = System.Text.Json.JsonDocument.Parse(File.ReadAllText(file));
                var el = doc.RootElement;
                foreach (var k in keys) {
                    if (el.ValueKind == System.Text.Json.JsonValueKind.Array) {
                        bool hit = false;
                        foreach (var item in el.EnumerateArray()) {
                            if (item.ValueKind == System.Text.Json.JsonValueKind.Object && item.TryGetProperty(k, out var v)) {
                                el = v; hit = true; break;
                            }
                        }
                        if (!hit) return null;
                    } else if (el.ValueKind == System.Text.Json.JsonValueKind.Object && el.TryGetProperty(k, out var v)) {
                        el = v;
                    } else return null;
                }
                return el.ValueKind == System.Text.Json.JsonValueKind.String ? el.GetString() : null;
            } catch { return null; }
        }

        static SaveScope ResolveSaveScope(string? targetPath) {
            var scope = new SaveScope();
            try {
                if (string.IsNullOrWhiteSpace(targetPath)) { scope.Problem = "找不到可驗證的檔案路徑"; return scope; }
                var full = Path.GetFullPath(targetPath);
                scope.BaseDir = Path.GetDirectoryName(full) ?? "";
                if (!File.Exists(full)) { scope.Problem = $"檔案不存在：{full}"; return scope; }
                scope.Files.Add(full);
                if (!full.EndsWith(".pbip", StringComparison.OrdinalIgnoreCase)) return scope;

                // 照 Power BI 自己的指標走：.pbip 記著報表資料夾，報表的 definition.pbir 記著語意模型資料夾。
                // 讀不到指標才退回「同名 + .Report / .SemanticModel」的慣例（資料夾可能被改過名）。
                string stem = Path.GetFileNameWithoutExtension(full);

                string? report = null;
                var rel = ReadJsonString(full, "artifacts", "report", "path");
                if (!string.IsNullOrWhiteSpace(rel)) report = Path.GetFullPath(Path.Combine(scope.BaseDir, rel!));
                if (!IsPbipFolder(report)) report = Path.Combine(scope.BaseDir, stem + ".Report");

                string? model = null;
                if (IsPbipFolder(report)) {
                    scope.Folders.Add(report!);
                    rel = ReadJsonString(Path.Combine(report!, "definition.pbir"), "datasetReference", "byPath", "path");
                    if (!string.IsNullOrWhiteSpace(rel)) model = Path.GetFullPath(Path.Combine(report!, rel!));
                }
                if (!IsPbipFolder(model)) model = Path.Combine(scope.BaseDir, stem + ".SemanticModel");
                if (IsPbipFolder(model)) scope.Folders.Add(model!);

                if (scope.Folders.Count == 0)
                    scope.Problem = $"找不到這個 PBIP 的 .Report / .SemanticModel 資料夾（預期在 {scope.BaseDir}）";
            } catch (Exception ex) {
                scope.Problem = $"解析存檔範圍失敗：{ex.Message}";
            }
            return scope;
        }

        /// <summary>範圍內每個檔案的修改時間。只讀目錄項目，不開檔。</summary>
        static Dictionary<string, DateTime> ScanWriteTimes(SaveScope scope) {
            var map = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in scope.Files) {
                try { if (File.Exists(f)) map[f] = File.GetLastWriteTimeUtc(f); } catch { }
            }
            // 深度上限：報表最深是 definition\pages\<頁>\visuals\<視覺>\visual.json（5 層），留一點餘裕
            var opts = new EnumerationOptions {
                RecurseSubdirectories = true, MaxRecursionDepth = 8,
                IgnoreInaccessible = true, AttributesToSkip = 0
            };
            foreach (var d in scope.Folders) {
                try {
                    foreach (var fi in new DirectoryInfo(d).EnumerateFiles("*", opts))
                        map[fi.FullName] = fi.LastWriteTimeUtc;
                } catch { /* 存檔途中資料夾可能正在重建，下一輪輪詢再看 */ }
            }
            return map;
        }

        /// <summary>兩次掃描之間有變動的檔案：內容更新、新增、刪除。</summary>
        static List<(string Path, DateTime Time, string Kind)> DiffWriteTimes(
                Dictionary<string, DateTime> before, Dictionary<string, DateTime> now) {
            var list = new List<(string, DateTime, string)>();
            foreach (var kv in now) {
                if (!before.TryGetValue(kv.Key, out var t)) list.Add((kv.Key, kv.Value, "新增"));
                else if (kv.Value != t)                     list.Add((kv.Key, kv.Value, "更新"));
            }
            foreach (var kv in before)
                if (!now.ContainsKey(kv.Key)) list.Add((kv.Key, kv.Value, "刪除"));
            return list;
        }

        static string CollapseWhitespace(string? s) => Regex.Replace(s ?? "", @"\s+", " ").Trim();

        /// <summary>
        /// 要放進確認視窗的文字：壓成一行、去掉控制字元與改變顯示方向的字元、限制長度。
        /// 視窗是使用者做決定的唯一依據，而裡面的名稱（資料表、欄位、共用查詢、查詢文字）是呼叫端給的 ——
        /// 不處理的話，名稱裡塞幾個換行就能在視窗裡排出一段「例行更新，直接按是」之類的假說明。
        /// </summary>
        static string OneLine(string? s, int max = 160) {
            string t = CollapseWhitespace(Regex.Replace(s ?? "", @"[\p{Cc}\p{Cf}\p{Zl}\p{Zp}]", " "));
            return t.Length <= max ? t : t.Substring(0, max) + "…";
        }
        /// <summary>確認視窗裡用「」括起來的名稱：名稱自己帶的括號換掉，免得提早收尾、後面接上假的說明。</summary>
        static string Quoted(string? name) => OneLine(name, 60).Replace('「', '『').Replace('」', '』').Replace('[', '(').Replace(']', ')');

        static readonly HashSet<string> DefinitionExts = new(StringComparer.OrdinalIgnoreCase) {
            ".tmdl", ".bim", ".json", ".pbir", ".pbism"
        };

        /// <summary>模型／報表的定義文字檔。.pbi 資料夾是本機快取與個人設定，不算定義。</summary>
        static bool IsDefinitionFile(string path) =>
            DefinitionExts.Contains(Path.GetExtension(path)) &&
            path.IndexOf(Path.DirectorySeparatorChar + ".pbi" + Path.DirectorySeparatorChar,
                         StringComparison.OrdinalIgnoreCase) < 0;

        /// <summary>
        /// 在定義檔裡找預期文字 —— 回答的是「剛寫入的東西真的進到磁碟了嗎」。
        /// 檔案時間變了不代表內容到了：Power BI 還沒同步到 API 的寫入時，Ctrl+S 存下去的是舊狀態。
        ///
        /// 只記錄「哪一段在哪個檔案找到」，檔案內容不會出現在回應裡。空白一律壓成單一空格再比對，
        /// 因為 TMDL 會替多行運算式加縮排。checkedAt 讓沒變動的檔案不必重讀。
        /// </summary>
        static void FindExpected(IEnumerable<string> files, Dictionary<string, DateTime> stamps, string[] needles,
                                 Dictionary<string, string> found, Dictionary<string, DateTime> checkedAt) {
            foreach (var f in files) {
                if (found.Count >= needles.Length) return;
                if (!IsDefinitionFile(f) || !stamps.TryGetValue(f, out var stamp)) continue;
                if (checkedAt.TryGetValue(f, out var seen) && seen == stamp) continue;
                string text;
                try {
                    var fi = new FileInfo(f);
                    if (!fi.Exists || fi.Length > 16 * 1024 * 1024) continue;
                    using var fs = new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                    using var sr = new StreamReader(fs, System.Text.Encoding.UTF8);
                    text = CollapseWhitespace(sr.ReadToEnd());
                } catch { continue; }     // 正在寫入而讀不到：不記錄，下一輪再試
                checkedAt[f] = stamp;
                foreach (var n in needles)
                    if (!found.ContainsKey(n) && text.Contains(n, StringComparison.Ordinal)) found[n] = f;
            }
        }

        // Win32：把視窗搶到前景並送出按鍵。
        // ⚠️ WScript.Shell 的 AppActivate 在這裡行不通 —— 背景服務沒有前景權限，
        //    必須先 AttachThreadInput 到目前前景視窗的執行緒才能繞過 Windows 的前景鎖。
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        const int SW_RESTORE = 9;
        const byte VK_CONTROL = 0x11, VK_S = 0x53;
        const uint KEYEVENTF_KEYUP = 0x0002;

        static bool ForceForeground(IntPtr hWnd) {
            if (IsIconic(hWnd)) { ShowWindow(hWnd, SW_RESTORE); Thread.Sleep(300); }

            IntPtr fg = GetForegroundWindow();
            if (fg == hWnd) return true;

            uint fgThread  = GetWindowThreadProcessId(fg, out _);
            uint curThread = GetCurrentThreadId();
            bool attached = false;
            if (fgThread != 0 && fgThread != curThread)
                attached = AttachThreadInput(curThread, fgThread, true);
            try {
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
            } finally {
                if (attached) AttachThreadInput(curThread, fgThread, false);
            }
            Thread.Sleep(400);
            return GetForegroundWindow() == hWnd;
        }

        /// <summary>對指定的 PBI Desktop 行程送出 Ctrl+S。pid 必須是解析到的那個實例，不能隨便抓一個。</summary>
        static async Task<string> SendCtrlS(int pbiPid) {
            Process proc;
            try { proc = Process.GetProcessById(pbiPid); } catch { return "not_running"; }
            if (proc.MainWindowHandle == IntPtr.Zero) return "not_running";

            if (!ForceForeground(proc.MainWindowHandle)) return "activate_failed";

            keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
            keybd_event(VK_S,       0, 0, UIntPtr.Zero);
            await Task.Delay(60);
            keybd_event(VK_S,       0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            return "sent";
        }

        // =====================================================================
        // 主程式
        // =====================================================================

        /// <summary>
        /// 從執行檔位置往上找專案根 —— 判斷依據是「這一層底下有 pbibridge_csharp 資料夾」。
        /// 用途是讓快照落在專案根，而不是 bin\Release\net8.0（那裡會被 dotnet clean 清掉）。
        /// 找不到就回 null，呼叫端自行退回 BaseDirectory。
        /// </summary>
        static string? FindProjectRoot() {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            for (int i = 0; i < 6 && dir != null; i++, dir = dir.Parent) {
                if (Directory.Exists(Path.Combine(dir.FullName, "pbibridge_csharp")))
                    return dir.FullName;
            }
            return null;
        }

        // =====================================================================
        // 資料保護：欄位層級管制
        // =====================================================================

        /// <summary>
        /// 攔截會把敏感「值」送進 AI context 的查詢。
        ///
        /// 前提是一條分界線：欄位的「名稱」不是機密（AI 要靠它寫程式），欄位的「內容」才是。
        /// 所以 /api/schema 照常全開，被管的只有 /api/query 與 /api/dmv 的回傳值。
        ///
        /// 每個欄位是四個等級之一（見 Level）：
        ///   開放       — 只受回傳列數上限管制。
        ///   只能彙總   — 數值（金額）。可用任何數值聚合，但不得逐列取值。
        ///   換成代號   — 可以當分組鍵，值在回傳前換成穩定代號。
        ///   只能計數   — 只准出現在「計數類」函式內，因為那類函式的回傳值必定是數字。
        ///                刻意不含 MAX/MIN/CONCATENATEX —— 對文字欄位 MAX 會直接吐出一個真實的值。
        ///
        /// 為什麼要兩道掃描（CheckQuery + CheckResultColumns）：兩者的盲點互補。
        ///   · 查詢文字掃描抓得到改名／別名（SELECTCOLUMNS(x,"a",[customer_name])），
        ///     因為欄名一定會出現在文字裡；但抓不到整表 EVALUATE（文字裡沒有欄名）。
        ///   · 結果欄位掃描剛好相反：整表 EVALUATE 的結果欄名是 'Table'[Column]，一抓就中。
        /// 少任何一道都有洞。
        /// </summary>
        static class DataGuard {
            public static bool     Enabled          = false;
            public static string[] DenyColumns      = Array.Empty<string>();
            public static string[] MoneyColumns     = Array.Empty<string>();
            public static int      MaxRowsWithMoney = 100;
            public static string   AuditPath        = "";

            // 明確標記為安全的欄位：優先於 DenyColumns / MoneyColumns。
            // 樣式比對難免誤傷（*owner* 會打到 process_owner），這裡是解法 ——
            // 不必為了一個誤判就把整條樣式拿掉。
            public static string[] AllowColumns     = Array.Empty<string>();

            // 回傳列數的硬上限。由伺服器決定，呼叫端只能更保守。
            //   MaxDetailRows    逐列明細 —— 真實資料，額度要小
            //   MaxAggregateRows 彙總結果 —— 統計量不是資料細節，可以放寬
            public static int      MaxDetailRows    = 50;
            public static int      MaxAggregateRows = 300;

            // 回傳值必定是數字（或是／否）的函式 ——「只能計數」欄位唯一的合法容身處。
            // DISTINCTCOUNT(客戶[名稱]) 回傳的是個數，是結構資訊，該放行；
            // MAX(客戶[名稱]) 回傳一個真實的名稱，不在此列。
            static readonly HashSet<string> CountingFuncs = new(StringComparer.OrdinalIgnoreCase) {
                "COUNTROWS", "COUNT", "COUNTA", "COUNTAX", "COUNTX", "COUNTBLANK",
                "DISTINCTCOUNT", "DISTINCTCOUNTNOBLANK",
                "ISBLANK", "ISEMPTY", "HASONEVALUE", "ISFILTERED", "ISCROSSFILTERED"
            };

            // 數值聚合 ——「只能彙總」欄位的合法容身處：把一整欄收成一個數字的函式。
            // MAX / MIN 可以在這裡，因為數值的極值是統計量（文字欄位不會被歸到這一級，見 Policy）。
            // DIVIDE 與 RANKX 曾經在這份清單裡，那是錯的 —— 它們不是聚合：
            // SELECTCOLUMNS(Sales, "v", DIVIDE(Sales[amount], 1)) 就是逐列把金額原樣取出來。
            static readonly HashSet<string> NumericAggFuncs = new(StringComparer.OrdinalIgnoreCase) {
                "SUM", "SUMX", "AVERAGE", "AVERAGEX", "MIN", "MINX", "MAX", "MAXX",
                "MEDIAN", "MEDIANX", "PERCENTILE.INC", "PERCENTILE.EXC",
                "STDEV.P", "STDEV.S", "VAR.P", "VAR.S",
                "PRODUCT", "PRODUCTX", "GEOMEAN", "GEOMEANX"
            };

            /// <summary>
            /// 欄位名比對前的正規化：忽略大小寫、空白、底線、連字號。
            ///
            /// 為什麼非有不可 —— 同一個語意的欄位在不同資料表裡拼法不同：
            /// Customers[account name] 與 Opportunities[account_name]
            /// 是同一批客戶，逐字比對只擋得住其中一種，另一種整批漏出去。
            /// </summary>
            static string Norm(string s) {
                if (string.IsNullOrEmpty(s)) return "";
                var sb = new System.Text.StringBuilder(s.Length);
                foreach (char ch in s)
                    if (ch != ' ' && ch != '_' && ch != '-') sb.Append(char.ToLowerInvariant(ch));
                return sb.ToString();
            }

            static readonly Dictionary<string, Regex> _globCache = new();

            /// <summary>把 *foo*bar* 形式的樣式轉成 Regex。比對對象是正規化後的字串。</summary>
            static Regex GlobRegex(string glob) {
                lock (_globCache) {
                    if (_globCache.TryGetValue(glob, out var re)) return re;
                    re = new Regex("^" + string.Join(".*", glob.Split('*').Select(Regex.Escape)) + "$",
                                   RegexOptions.Compiled);
                    _globCache[glob] = re;
                    return re;
                }
            }

            /// <summary>
            /// 清單項目可以是完整欄名，也可以是含 * 的樣式。
            /// 樣式才是主力：大模型的欄位上千個，而且會一直新增 ——
            /// 逐一列舉永遠追不上，而漏掉的那一個不會有人發現。
            /// 第一個符合的樣式會從 rule 傳回去（給設定畫面顯示「是哪一條規則決定的」）。
            /// </summary>
            static bool Matches(string token, string[] cols, out string? rule) {
                rule = null;
                string t = Norm(token);
                if (t.Length == 0) return false;
                foreach (var c in cols) {
                    if (string.IsNullOrWhiteSpace(c)) continue;
                    bool hit = c.IndexOf('*') >= 0 ? GlobRegex(Norm(c)).IsMatch(t) : t == Norm(c);
                    if (hit) { rule = c; return true; }
                }
                return false;
            }

            // ── 保護等級 ──────────────────────────────────────────────────────
            // 四個等級不是一條由鬆到緊的直線：AggregateOnly 與 Pseudonym 不能比大小（一個不准分組、
            // 一個不准聚合以外的取值）。程式只做相等比較（見 IsLoosening、Combine），不依賴數值大小。
            public enum Level { Open = 0, AggregateOnly = 1, Pseudonym = 2, CountOnly = 3 }

            public static string LevelKey(Level l) => l switch {
                Level.Open => "open", Level.AggregateOnly => "aggregateOnly", Level.Pseudonym => "pseudonym", _ => "countOnly"
            };
            public static string LevelName(Level l) => l switch {
                Level.Open => "開放", Level.AggregateOnly => "只能彙總", Level.Pseudonym => "換成代號", _ => "只能計數"
            };
            public static bool TryParseLevel(string? s, out Level level) {
                switch ((s ?? "").Trim().ToLowerInvariant()) {
                    case "open":          level = Level.Open;          return true;
                    case "aggregateonly": level = Level.AggregateOnly; return true;
                    case "pseudonym":     level = Level.Pseudonym;     return true;
                    case "countonly":     level = Level.CountOnly;     return true;
                    default:              level = Level.CountOnly;     return false;
                }
            }

            /// <summary>
            /// 從 from 改成 to，AI 讀得到的東西會不會變多。會的話要使用者本人確認。
            /// 只有兩種情況確定不會變多：改成最嚴的「只能計數」，或原本就是開放。
            /// 「換成代號 ↔ 只能彙總」互相不能比較，一律當成放寬。
            /// </summary>
            public static bool IsLoosening(Level from, Level to) =>
                to != from && to != Level.CountOnly && from != Level.Open;

            /// <summary>同一個名稱可能指到好幾個資料行時，取最嚴的；受限的方式不一致就當成「只能計數」。</summary>
            static Level Combine(IEnumerable<Level> levels) {
                Level? acc = null;
                foreach (var l in levels) {
                    if (l == Level.Open) continue;
                    if (acc == null) acc = l;
                    else if (acc != l) return Level.CountOnly;
                }
                return acc ?? Level.Open;
            }

            // 「換成代號」欄位唯一被允許的取值方式：當分組鍵／取相異值。
            // 這些函式會讓欄位以 Table[Column] 的原名出現在結果欄位裡，結果欄位那一關認得出來，
            // 也就換得掉。反過來 MAX / CONCATENATEX / SELECTCOLUMNS 會讓真名躲在別名底下，
            // 欄名比對看不見 —— 那些用法一律擋。
            static readonly HashSet<string> GroupingShapes = new(StringComparer.OrdinalIgnoreCase) {
                "SUMMARIZECOLUMNS", "SUMMARIZE", "GROUPBY",
                "VALUES", "DISTINCT", "ALL", "ALLSELECTED", "ALLNOBLANKROW"
            };

            // TOPN 只有「第二個引數（資料表）」算分組位置：TOPN(20, SUMMARIZECOLUMNS(客戶[名稱], …), [額])。
            // 排序用的引數是逐列計算的，放在那裡等於可以拿欄位去跟任意值比較 —— 不算。
            // 沒有 TOPN 的話「前 N 大客戶」寫不出來，而金額查詢超過列數上限又會整個被擋，代號就形同虛設。
            static bool IsGroupingFrame(Frame f) =>
                GroupingShapes.Contains(f.Func) || (f.Arg == 1 && f.Func.Equals("TOPN", StringComparison.OrdinalIgnoreCase));

            // 不必指名欄位就能把整張表的內容帶出來的函式 —— 欄位層級的掃描看不到它們，只能整個禁用。
            //   TOCSV / TOJSON     把資料表整個序列化成一個字串
            //   COLUMNSTATISTICS   回傳模型裡每一個資料行的最小值與最大值（文字欄位的極值就是真實內容）
            // INFO.* 另外處理（見 CheckQuery）：那一族會回傳 M 腳本與運算式原文。
            //   DETAILROWS         回傳量值的「詳細資料列運算式」算出來的資料表，那份運算式不在量值的公式裡，展開看不到
            static readonly HashSet<string> BannedCalls = new(StringComparer.OrdinalIgnoreCase) {
                "TOCSV", "TOJSON", "COLUMNSTATISTICS", "DETAILROWS"
            };

            // 會「重新指定輸出欄位」的函式：經過它們之後，流出去的每一欄都有明寫的來源運算式可以檢查。
            // 用來判斷不帶中括號的資料表引用會不會原封不動流進 UNION（見 CheckQuery ③）。
            // 不在清單上的函式一律當成「原樣傳遞資料列」—— 認不出來就往嚴的方向判。
            static readonly HashSet<string> ProjectingFuncs = new(StringComparer.OrdinalIgnoreCase) {
                "SELECTCOLUMNS", "SUMMARIZE", "SUMMARIZECOLUMNS", "GROUPBY", "ROW", "CONCATENATEX"
            };

            static byte[] _pseudoKey = Array.Empty<byte>();

            /// <summary>代號金鑰由 API Key 衍生 —— 只存在這台機器上，永遠不會離開。</summary>
            public static void InitPseudonymKey(string? apiKey) {
                using var sha = System.Security.Cryptography.SHA256.Create();
                _pseudoKey = sha.ComputeHash(
                    System.Text.Encoding.UTF8.GetBytes((apiKey ?? "") + "|pseudonym|v1"));
            }

            /// <summary>
            /// 把一個真實值換成穩定代號。
            ///
            /// 用 HMAC 而不是流水號，是為了「同一個值永遠對到同一個代號」——
            /// 分組、去重、跨查詢串接才還能用（先查前 20 大客戶，再查這些客戶的月趨勢，
            /// 兩次的代號會對得起來）。代價是同一個代號在雲端會累積屬性，
            /// 但沒有金鑰就推不回原值，而金鑰不會離開這台機器。
            ///
            /// 取 10 個十六進位字元（40 bit）。以前只取 6 個：1,500 個客戶就有約 6.5% 的機率
            /// 出現兩個客戶撞同一個代號；10 個字元要到五十萬個相異值才會到那個機率。
            /// </summary>
            public static string? Pseudonym(object? value) {
                if (value == null) return null;
                string v = value.ToString() ?? "";
                if (v.Length == 0) return v;
                using var h = new System.Security.Cryptography.HMACSHA256(_pseudoKey);
                return "ID_" + Convert.ToHexString(
                    h.ComputeHash(System.Text.Encoding.UTF8.GetBytes(v)), 0, 5);
            }

            // =================================================================
            // 掃描：把 DAX 文字拆成「欄位參考 + 它所在的上下文」
            // =================================================================

            /// <summary>包住某個參考的一層函式，以及那個參考落在它的第幾個引數（0 起算）。</summary>
            public readonly record struct Frame(string Func, int Arg);

            /// <summary>一個中括號參考，或是一個不帶中括號的資料表／變數引用（Name 就是那個名稱、Table 為 null）。</summary>
            public sealed class Tok {
                public string  Name = "";
                public string? Table;                         // 緊鄰在 [ 之前的表名；沒有就是 null
                public Frame[] Enclosing = Array.Empty<Frame>();   // 由外到內
                public bool    InDefinition;                  // 位於 VAR…RETURN、DEFINE…EVALUATE，或展開進來的模型定義裡
            }

            public sealed class Scan {
                public List<Tok> Tokens = new();              // [欄位] 參考
                public List<Tok> Bare   = new();              // 不帶中括號的識別字／'引號名稱'（資料表或變數）
                public HashSet<string> Calls    = new(StringComparer.OrdinalIgnoreCase);   // 出現過的函式呼叫
                public HashSet<string> VarNames = new(StringComparer.OrdinalIgnoreCase);   // VAR 與 DEFINE TABLE 定義的名稱
            }

            static bool IsIdentChar(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.';

            /// <summary>取得緊接在左括號之前的識別字（函式名）。沒有就回空字串。</summary>
            static string IdentifierBefore(string s, int parenIndex, int start) {
                int j = parenIndex - 1;
                while (j >= start && char.IsWhiteSpace(s[j])) j--;
                int end = j;
                while (j >= start && IsIdentChar(s[j])) j--;
                return end > j ? s.Substring(j + 1, end - j) : "";
            }

            /// <summary>
            /// 掃過 DAX 文字。字串常值、註解都跳過 —— 裡頭的括號會打亂堆疊，而堆疊一亂，
            /// 「這個欄位有沒有被聚合包住」的判斷就跟著錯。
            ///
            /// 每個參考記三件事：
            ///   · 表名：只有「緊鄰」在 [ 前面才算（Sales[amount]、'Dim 產品'[類別]）。中間有空白就當成沒有表名，
            ///     之後會用同名資料行裡最嚴的等級 —— 寧可多擋，也不要把參考算到錯的表上
            ///   · 包住它的函式，以及它落在各層的第幾個引數
            ///   · 是不是在「定義」裡。變數會切斷文字上的包覆關係：
            ///       VAR v = VALUES(客戶[名稱]) RETURN ROW("x", v)
            ///     從 VALUES 看是合法的取相異值，實際上卻經由變數變成一個沒有名字的純量流出去。
            ///     所以定義裡的參考不適用「當分組鍵」那條放行規則。
            ///
            /// definitionsFrom：這個位置之後是 ExpandDerivations 接上來的模型定義，整段都算定義。
            /// boundaries：每一份接上來的定義的起點。查詢本體與每一份定義**各自獨立掃描** ——
            ///   某一份定義裡沒收尾的字串、註解或括號，不會把後面的定義一起吃掉。
            ///   （否則建一個結尾是「/*」的量值，排在它後面的定義就全部變成註解、不被檢查。）
            ///   分界用的是位置而不是分隔字元，因為字元可以被寫進查詢裡。
            /// </summary>
            public static Scan ScanDax(string? dax, int definitionsFrom = int.MaxValue, IEnumerable<int>? boundaries = null) {
                var scan = new Scan();
                string s = dax ?? "";
                var cuts = new List<int> { 0 };
                if (boundaries != null) cuts.AddRange(boundaries.Where(b => b > 0 && b < s.Length).OrderBy(b => b));
                cuts.Add(s.Length);
                for (int n = 0; n + 1 < cuts.Count; n++)
                    ScanSegment(s, cuts[n], cuts[n + 1], cuts[n] >= definitionsFrom, scan);
                return scan;
            }

            static void ScanSegment(string s, int start, int end, bool isDefinition, Scan scan) {
                var frames = new List<Frame>();
                var inVar  = new List<bool> { false };       // 每一層括號深度：是否位於 VAR 與它的 RETURN 之間
                bool inDefine = false;                       // DEFINE … EVALUATE 之間
                bool expectName = false;                     // 前一個字是 VAR（或 DEFINE 裡的 TABLE）：下一個名稱是被定義的名字
                string? pendingTable = null; int pendingEnd = -1;
                int i = start;

                Tok Make(string name, string? table) => new Tok {
                    Name = name, Table = table, Enclosing = frames.ToArray(),
                    InDefinition = isDefinition || inDefine || inVar.Contains(true)
                };

                while (i < end) {
                    char c = s[i];

                    if (c == '"') {                                   // 字串常值（"" 為跳脫）
                        i++;
                        while (i < end) {
                            if (s[i] == '"') {
                                if (i + 1 < end && s[i + 1] == '"') { i += 2; continue; }
                                i++; break;
                            }
                            i++;
                        }
                        continue;
                    }
                    if (c == '/' && i + 1 < end && s[i + 1] == '/') { while (i < end && s[i] != '\n') i++; continue; }
                    if (c == '-' && i + 1 < end && s[i + 1] == '-') { while (i < end && s[i] != '\n') i++; continue; }
                    if (c == '/' && i + 1 < end && s[i + 1] == '*') {
                        i += 2;
                        while (i + 1 < end && !(s[i] == '*' && s[i + 1] == '/')) i++;
                        i += 2;
                        continue;
                    }

                    if (c == '\'') {                                  // 'Table Name'（'' 為跳脫）
                        int j = i + 1;
                        var sb = new System.Text.StringBuilder();
                        while (j < end) {
                            if (s[j] == '\'') {
                                if (j + 1 < end && s[j + 1] == '\'') { sb.Append('\''); j += 2; continue; }
                                break;
                            }
                            sb.Append(s[j]); j++;
                        }
                        int after = Math.Min(j + 1, end);             // 結尾引號的下一格
                        string name = sb.ToString();
                        if (expectName) { scan.VarNames.Add(name); expectName = false; }
                        else if (after < end && s[after] == '[') { pendingTable = name; pendingEnd = after; }
                        else scan.Bare.Add(Make(name, null));
                        i = after;
                        continue;
                    }

                    if (c == '[') {                                   // [欄位或量值]（]] 為跳脫）
                        int j = i + 1;
                        var sb = new System.Text.StringBuilder();
                        bool closed = false;
                        while (j < end) {
                            if (s[j] == ']') {
                                if (j + 1 < end && s[j + 1] == ']') { sb.Append(']'); j += 2; continue; }
                                closed = true; break;
                            }
                            sb.Append(s[j]); j++;
                        }
                        if (!closed) { i++; continue; }
                        scan.Tokens.Add(Make(sb.ToString().Trim(), pendingEnd == i ? pendingTable : null));
                        pendingTable = null; pendingEnd = -1;
                        i = j + 1;
                        continue;
                    }

                    if (char.IsLetter(c) || c == '_') {               // 函式名、關鍵字、未加引號的表名、變數名
                        int j = i;
                        while (j < end && IsIdentChar(s[j])) j++;
                        string id = s.Substring(i, j - i);
                        int k = j;
                        while (k < end && char.IsWhiteSpace(s[k])) k++;
                        bool isCall = k < end && s[k] == '(';

                        if (expectName)                    { scan.VarNames.Add(id); expectName = false; }
                        else if (isCall)                   { scan.Calls.Add(id); }          // 堆疊在遇到 ( 的時候推
                        else if (j < end && s[j] == '[')   { pendingTable = id; pendingEnd = j; }
                        else switch (id.ToUpperInvariant()) {
                            case "VAR":      inVar[inVar.Count - 1] = true; expectName = true; break;
                            case "RETURN":   inVar[inVar.Count - 1] = false; break;
                            case "DEFINE":   inDefine = true; break;
                            case "EVALUATE": inDefine = false; if (frames.Count == 0) inVar[0] = false; break;
                            case "TABLE":    if (inDefine) expectName = true; else scan.Bare.Add(Make(id, null)); break;
                            case "MEASURE": case "COLUMN": case "FUNCTION":
                            case "ORDER": case "BY": case "ASC": case "DESC": case "START": case "AT":
                            case "IN": case "NOT": case "AND": case "OR": case "TRUE": case "FALSE":
                                break;
                            default:         scan.Bare.Add(Make(id, null)); break;
                        }
                        i = j;
                        continue;
                    }

                    if (c == '(' || c == '{') {
                        frames.Add(new Frame(c == '{' ? "{" : IdentifierBefore(s, i, start), 0));
                        inVar.Add(false);
                        i++; continue;
                    }
                    if (c == ')' || c == '}') {
                        if (frames.Count > 0) { frames.RemoveAt(frames.Count - 1); inVar.RemoveAt(inVar.Count - 1); }
                        i++; continue;
                    }
                    if (c == ',') {
                        if (frames.Count > 0) frames[^1] = frames[^1] with { Arg = frames[^1].Arg + 1 };
                        i++; continue;
                    }
                    i++;
                }
            }

            // 真正做聚合的函式。刻意不含 ISBLANK / ISEMPTY / HASONEVALUE 這類布林述詞 ——
            // 它們雖然回傳純量（所以在 CheckQuery 裡算合法的容身處），卻常出現在
            // FILTER 的條件裡，而 FILTER 回傳的是**原始資料列**。把它們當成
            // 「這是彙總查詢」，EVALUATE FILTER(表, ISBLANK(欄)) 就會拿到彙總額度，
            // 明細上限整個失效。這是實測抓到的，不是假想。
            static readonly Regex _aggCall = new Regex(
                @"\b(SUMX?|AVERAGEX?|MINX?|MAXX?|MEDIANX?|PRODUCTX?|GEOMEANX?|COUNTX?|COUNTAX?"
              + @"|COUNTROWS|COUNTA|COUNTBLANK|DISTINCTCOUNT|DISTINCTCOUNTNOBLANK"
              + @"|PERCENTILE\.(INC|EXC)|STDEV\.[PS]|VAR\.[PS])\s*\(",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            // 會產出「每列一組統計量」的表格函式。查詢最外層必須是這些之一，
            // 才有資格套用比較寬的彙總額度。FILTER / TOPN / SELECTCOLUMNS / VALUES
            // 一律不在此列 —— 它們吐的是資料列。
            // UNION 也不在：UNION(ROW("c", SUM(…)), SELECTCOLUMNS(大表, "c", 大表[欄])) 只要其中一支
            // 有聚合，整句就會被當成彙總，另一支的明細跟著拿到 300 列的額度。
            static readonly HashSet<string> AggregateShapes = new(StringComparer.OrdinalIgnoreCase) {
                "SUMMARIZECOLUMNS", "SUMMARIZE", "GROUPBY", "ROW", "DATATABLE"
            };

            /// <summary>取 EVALUATE 之後最外層的函式名。整表取出（沒有函式）回空字串。</summary>
            static string OuterFunction(string dax) {
                string s = dax ?? "";
                int e = s.IndexOf("EVALUATE", StringComparison.OrdinalIgnoreCase);
                if (e < 0) return "";
                int i = e + 8;
                while (i < s.Length && (char.IsWhiteSpace(s[i]) || s[i] == '(')) i++;
                int start = i;
                while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '_' || s[i] == '.')) i++;
                string name = s.Substring(start, i - start);
                while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
                return (i < s.Length && s[i] == '(') ? name : "";
            }

            /// <summary>
            /// 這句查詢是不是「彙總型」—— 只用來決定回傳列數上限，不決定放不放行。
            ///
            /// 兩個條件都要成立：
            ///   ① 最外層是會產出統計量的表格函式（SUMMARIZECOLUMNS / SUMMARIZE / GROUPBY / ROW）
            ///   ② 查詢裡真的有聚合呼叫（SUM( / COUNTROWS( …）
            ///
            /// 只看 ① 會讓 SUMMARIZECOLUMNS(表[欄]) —— 其實只是取相異值 —— 拿到寬額度；
            /// 只看 ② 會讓 FILTER(表, [額] > SUM(...)) 這種回傳原始列的查詢拿到寬額度。
            /// 兩個都要，判定才會偏保守：認不出來就當明細。
            ///
            /// 傳入的必須是「已展開量值定義」的文字，否則
            /// SUMMARIZECOLUMNS(日期[月], "額", [銷售總額]) 會因為看不到量值裡的 SUM
            /// 而被誤判成明細，正常的趨勢分析就會被卡在明細的列數上限。
            /// </summary>
            public static bool IsAggregateQuery(string dax) {
                if (!AggregateShapes.Contains(OuterFunction(dax))) return false;
                // 字串常值裡的 "SUM(" 不算數
                string bare = Regex.Replace(dax ?? "", "\"(?:[^\"]|\"\")*\"", "\"\"");
                return _aggCall.IsMatch(bare);
            }

            /// <summary>模型的中繼資料：DAX 衍生定義（量值、計算資料行、計算表）與欄位目錄。</summary>
            public sealed class Derivations {
                // 以中括號參考的（量值名、計算資料行名）。一個名稱可以對到不只一份定義：
                // 不同資料表的計算資料行可以同名，量值也可以和別張表的資料行同名。
                // 只留一份的話，後讀到的會蓋掉先讀到的 —— 被蓋掉的那份定義就不會被掃描。
                public Dictionary<string, List<string>> ByToken = new(StringComparer.OrdinalIgnoreCase);
                // 以表名參考的（計算表的定義、計算群組的各個計算項目）—— 表名不帶中括號，只能用字串包含比對
                public Dictionary<string, List<string>> ByTable = new(StringComparer.OrdinalIgnoreCase);
                public int Count => ByToken.Values.Sum(v => v.Count) + ByTable.Values.Sum(v => v.Count);

                // 模型裡全部的量值名與資料行名，用來分辨 [X] 指的是量值還是資料行。
                // ColumnNames 存的是正規化後的名稱（比對刻意放寬，寧可少豁免）。
                public HashSet<string> MeasureNames = new(StringComparer.OrdinalIgnoreCase);
                public HashSet<string> ColumnNames  = new(StringComparer.Ordinal);
                public bool NamesLoaded;      // 兩份名單都完整讀到才是 true

                // 欄位目錄：資料表 → (資料行 → 是不是數值型別)。不含系統的 RowNumber。
                // 逐欄設定是記在「哪張表的哪一欄」上的，要靠這份目錄才知道 Sales[name] 與 Customers[name] 是兩回事。
                public Dictionary<string, Dictionary<string, bool>> Catalog = new(StringComparer.OrdinalIgnoreCase);
                public bool CatalogLoaded;    // 目錄完整讀到才是 true；讀不到時改用比較嚴的判法（見 Policy.Resolve）

                // 有任何一份定義沒讀到（量值、資料行、計算表、計算群組）。沒讀到的定義就沒被檢查 ——
                // 呼叫端必須把查詢擋下來，不能當成「這個模型沒有那些東西」。
                public bool    LoadFailed;
                public string? LoadProblem;   // 第一個讀不到的系統檢視

                // 資料行的 LineageTag → 它現在叫什麼。LineageTag 是 Power BI 給每個資料行的固定代號，
                // 改名、搬到別張表都不會變 —— 逐欄設定靠它跟著欄位走，使用者把欄位改個名字保護不會掉。
                public Dictionary<string, (string Table, string Column)> TagToColumn = new(StringComparer.OrdinalIgnoreCase);

                /// <summary>從 TOM 模型建立欄位目錄（設定畫面用；查詢路徑走的是 DMV 版的 LoadDerivations）。</summary>
                public static Derivations FromModel(Model model) {
                    var d = new Derivations();
                    foreach (var t in model.Tables) {
                        var cols = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                        foreach (var c in t.Columns) {
                            if (c.Type == ColumnType.RowNumber) continue;
                            cols[c.Name] = c.DataType == DataType.Int64 || c.DataType == DataType.Double || c.DataType == DataType.Decimal;
                            d.ColumnNames.Add(Norm(c.Name));
                            if (!string.IsNullOrEmpty(c.LineageTag)) d.TagToColumn[c.LineageTag] = (t.Name, c.Name);
                        }
                        d.Catalog[t.Name] = cols;
                        foreach (var m in t.Measures) d.MeasureNames.Add(m.Name);
                    }
                    d.CatalogLoaded = true;
                    d.NamesLoaded   = true;
                    return d;
                }

                Dictionary<string, List<string>>? _owners;
                /// <summary>有這個資料行名稱的所有資料表。</summary>
                public List<string> TablesWithColumn(string column) {
                    if (_owners == null) {
                        _owners = new(StringComparer.OrdinalIgnoreCase);
                        foreach (var t in Catalog)
                            foreach (var c in t.Value.Keys) {
                                if (!_owners.TryGetValue(c, out var list)) _owners[c] = list = new List<string>();
                                list.Add(t.Key);
                            }
                    }
                    return _owners.TryGetValue(column, out var hit) ? hit : new List<string>();
                }

                /// <summary>
                /// [token] 在這個模型裡只可能是量值：有這個名字的量值，而且沒有任何資料行叫這個名字。
                ///
                /// 第二個條件不能省。量值可以和「別張表」的資料行同名 —— 只看「有這個量值」就豁免的話，
                /// 建一個叫 [amount] 的量值，就能讓 SELECTCOLUMNS(Sales, "x", Sales[amount]) 逐列取出金額。
                /// 名單沒讀完整時一律回 false：寧可像以前一樣誤擋，也不要因為少讀一份名單而放行。
                /// </summary>
                public bool IsMeasureOnly(string token) =>
                    NamesLoaded && MeasureNames.Contains(token) && !ColumnNames.Contains(Norm(token));
            }

            /// <summary>
            /// 讀取模型的衍生定義與欄位目錄。
            ///
            /// ⚠️ 這裡刻意「不做快取」。曾經放過 60 秒快取，那是個漏洞：
            ///    /api/upsert-measure 不受本管制（它是寫入端點），所以可以先建一個
            ///    [x] = MAX(Customers[account name])，再趁快取沒更新時查詢它 ——
            ///    展開查不到新量值，結果欄位又是別名，兩道掃描都會放行。
            ///    而且快取連正常流程都是錯的：Set-PbiMeasure 後馬上驗算是日常操作。
            ///    一次 DMV 往返只有幾毫秒，拿它換掉一個繞道通道非常划算。
            /// </summary>
            public static Derivations LoadDerivations(AdomdConnection conn) {
                var d = new Derivations();

                bool Slurp(string sql, Action<System.Data.IDataRecord> take) {
                    try {
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = sql;
                        using var rd = cmd.ExecuteReader();
                        while (rd.Read()) take(rd);
                        return true;
                    } catch {
                        return false;       // 缺了要不要緊由呼叫端決定（見 Fail）
                    }
                }
                // 少讀到一份定義不是「比較嚴」，而是比較鬆：量值沒讀到，[某量值] 就不會被展開，
                // 藏在裡面的 MAX(客戶[名稱]) 也就沒人檢查。所以記下來，由呼叫端把查詢擋掉。
                void Fail(string view) { d.LoadFailed = true; d.LoadProblem ??= view; }
                void AddDef(string? name, string? expr) {
                    if (string.IsNullOrEmpty(name) || string.IsNullOrWhiteSpace(expr)) return;
                    if (!d.ByToken.TryGetValue(name!, out var list)) d.ByToken[name!] = list = new List<string>();
                    list.Add(expr!);
                }

                bool okMeasures = Slurp("SELECT [Name], [Expression] FROM $SYSTEM.TMSCHEMA_MEASURES", r => {
                    string? n = r[0]?.ToString();
                    if (!string.IsNullOrWhiteSpace(n)) d.MeasureNames.Add(n!.Trim());
                    AddDef(n, r[1]?.ToString());
                });
                if (!okMeasures) Fail("TMSCHEMA_MEASURES");

                var tableNames  = new Dictionary<string, string>();         // 資料表 ID → 名稱
                var columnNames = new Dictionary<string, (string Table, string Column)>();   // 資料行 ID → 位置
                bool okTables = Slurp("SELECT [ID], [Name] FROM $SYSTEM.TMSCHEMA_TABLES", r => {
                    string? id = r[0]?.ToString(), n = r[1]?.ToString();
                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(n)) {
                        tableNames[id!] = n!;
                        if (!d.Catalog.ContainsKey(n!)) d.Catalog[n!] = new(StringComparer.OrdinalIgnoreCase);
                    }
                });
                if (!okTables) Fail("TMSCHEMA_TABLES");

                // 計算資料行：把客戶名稱複製到一個不在管制清單裡的新欄名，是最直覺的洗資料手法。
                // 一般資料行的 Expression 是空的，AddDef 會濾掉；但名稱全部都要記（見 IsMeasureOnly）。
                // 計算表的資料行名稱在 InferredName，不在 ExplicitName —— 兩個都收。
                // Type：1 一般、2 計算、3 RowNumber（系統用，不算）、4 計算表的資料行。
                // DataType：6 整數、8 小數、10 貨幣是數值；1（自動）與 19（未知）要改看 Inferred。
                static bool IsNumericType(string? code) => code == "6" || code == "8" || code == "10";
                bool okColumns = Slurp(
                    "SELECT [TableID], [ExplicitName], [InferredName], [Expression], [Type], [ExplicitDataType], [InferredDataType], [ID] "
                  + "FROM $SYSTEM.TMSCHEMA_COLUMNS", r => {
                    string? explicitName = r[1]?.ToString(), inferred = r[2]?.ToString();
                    string? name = string.IsNullOrEmpty(explicitName) ? inferred : explicitName;
                    if (!string.IsNullOrWhiteSpace(explicitName)) d.ColumnNames.Add(Norm(explicitName!));
                    if (!string.IsNullOrWhiteSpace(inferred))     d.ColumnNames.Add(Norm(inferred!));
                    AddDef(name, r[3]?.ToString());
                    if (r[4]?.ToString() == "3" || string.IsNullOrEmpty(name)) return;
                    if (tableNames.TryGetValue(r[0]?.ToString() ?? "", out var table)) {
                        string? type = r[5]?.ToString();
                        if (string.IsNullOrEmpty(type) || type == "1" || type == "19") type = r[6]?.ToString();
                        d.Catalog[table][name!] = IsNumericType(type);
                        columnNames[r[7]?.ToString() ?? ""] = (table, name!);
                    }
                });
                d.CatalogLoaded = okTables && okColumns;
                // LineageTag 另外查：舊的相容性層級沒有這個欄位，查不到就算了（逐欄設定改用名稱對應），
                // 不能因為它讓整份目錄讀取失敗。
                if (d.CatalogLoaded)
                    Slurp("SELECT [ID], [LineageTag] FROM $SYSTEM.TMSCHEMA_COLUMNS", r => {
                        string? tag = r[1]?.ToString();
                        if (!string.IsNullOrEmpty(tag) && columnNames.TryGetValue(r[0]?.ToString() ?? "", out var where))
                            d.TagToColumn[tag!] = where;
                    });
                if (!okColumns) {
                    // 退回舊的查詢：沒有目錄（判定改用比較嚴的那一套），但名單與計算資料行的定義照樣要有
                    d.Catalog.Clear();
                    okColumns = Slurp("SELECT [ExplicitName], [InferredName], [Expression] FROM $SYSTEM.TMSCHEMA_COLUMNS", r => {
                        string? explicitName = r[0]?.ToString(), inferred = r[1]?.ToString();
                        if (!string.IsNullOrWhiteSpace(explicitName)) d.ColumnNames.Add(Norm(explicitName!));
                        if (!string.IsNullOrWhiteSpace(inferred))     d.ColumnNames.Add(Norm(inferred!));
                        AddDef(string.IsNullOrEmpty(explicitName) ? inferred : explicitName, r[2]?.ToString());
                    });
                    if (!okColumns) Fail("TMSCHEMA_COLUMNS");
                }
                d.NamesLoaded = okMeasures && okColumns;

                // 計算表：同理，但參考時不帶中括號。
                //
                // ⚠️ 這裡必須只取 Type = 2（Calculated），**絕對不能連 M 分割區一起拿**。
                //    曾經全拿過，結果是把 M 腳本接進 DAX 掃描文字 —— M 同樣用 [欄位]
                //    中括號語法，卻沒有 SUM/DISTINCTCOUNT 這種 DAX 聚合脈絡，於是
                //    M 裡每個欄位參考都被判成「裸露引用」，任何碰到該表的合法查詢
                //    全部誤擋（實測：SUM([amount]) 被回報成 [opportunity id] 違規）。
                //    過度納入在這裡不是「比較嚴」，是直接壞掉。
                try {
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT [Name], [QueryDefinition], [Type] FROM $SYSTEM.TMSCHEMA_PARTITIONS";
                    using var rd = cmd.ExecuteReader();
                    while (rd.Read()) {
                        string? n = rd[0]?.ToString();
                        string? e = rd[1]?.ToString();
                        if (string.IsNullOrEmpty(n) || string.IsNullOrWhiteSpace(e)) continue;
                        if (!int.TryParse(rd[2]?.ToString(), out int type) || type != 2) continue; // 2 = Calculated
                        if (!d.ByTable.TryGetValue(n!, out var defs)) d.ByTable[n!] = defs = new List<string>();
                        defs.Add(e!);
                    }
                } catch {
                    // 計算表的定義沒讀到：EVALUATE 計算表 回來的是計算表自己的欄名，
                    // 結果欄位那一關認不出它是從哪個受限欄位複製來的
                    Fail("TMSCHEMA_PARTITIONS");
                }

                // 計算群組：計算項目的運算式會「取代」量值的計算，卻不屬於任何量值的定義。
                // 建一個內容是 MAX(客戶[名稱]) 的計算項目，再用 CALCULATE([任何量值], 計算群組[項目] = "那一項")
                // 套用它，取值的運算式就完全不在查詢文字裡，也不在任何量值的定義裡。
                // 所以查詢只要提到計算群組（表名，或它的資料行），就把那個群組每個項目的運算式接進來檢查。
                var groupTable = new Dictionary<string, string>();          // 計算群組 ID → 資料表名稱
                bool okGroups = Slurp("SELECT [ID], [TableID] FROM $SYSTEM.TMSCHEMA_CALCULATION_GROUPS", r => {
                    if (tableNames.TryGetValue(r[1]?.ToString() ?? "", out var t)) groupTable[r[0]?.ToString() ?? ""] = t;
                });
                if (!okGroups) {
                    // 相容性層級低於 1470 的模型不可能有計算群組，引擎也可能連這個檢視都不給查 —— 那種情況不算失敗。
                    int compat = 0;
                    bool okCompat = Slurp("SELECT [COMPATIBILITY_LEVEL] FROM $SYSTEM.DBSCHEMA_CATALOGS",
                                          r => { if (int.TryParse(r[0]?.ToString(), out int c) && c > compat) compat = c; });
                    if (!okCompat || compat == 0 || compat >= 1470) Fail("TMSCHEMA_CALCULATION_GROUPS");
                }
                if (groupTable.Count > 0 &&
                    !Slurp("SELECT [CalculationGroupID], [Expression] FROM $SYSTEM.TMSCHEMA_CALCULATION_ITEMS", r => {
                        string? expr = r[1]?.ToString();
                        if (string.IsNullOrWhiteSpace(expr) || !groupTable.TryGetValue(r[0]?.ToString() ?? "", out var t)) return;
                        if (!d.ByTable.TryGetValue(t, out var defs)) d.ByTable[t] = defs = new List<string>();
                        defs.Add(expr!);
                        if (d.Catalog.TryGetValue(t, out var cols))
                            foreach (var c in cols.Keys) AddDef(c, expr);
                    })) Fail("TMSCHEMA_CALCULATION_ITEMS");

                return d;
            }

            /// <summary>查詢文字加上它引用到的所有衍生定義。</summary>
            public sealed class Expanded {
                public string       Text = "";
                public int          DefinitionsFrom = int.MaxValue;   // 這個位置之後是接上來的定義
                public HashSet<int> Boundaries = new();               // 每一份定義的起點
                public bool         Incomplete;                       // 定義多到沒展開完 —— 呼叫端必須擋下查詢
            }

            /// <summary>
            /// 把查詢引用到的衍生定義接在文字後面，讓查詢掃描看得到藏在下游定義裡的欄位。
            /// 不展開的話，「建一個新量值／新資料行去複製敏感欄位」就能整個繞過管制。
            ///
            /// 一直展開到沒有新的定義為止。以前只展開 6 層：建一串 [L0]=[L1]、[L1]=[L2] … 的量值，
            /// 把真正取值的那一個放在第 7 層，掃描就看不到它。每個名稱只會被接一次，所以一定會停；
            /// 文字長到不合理時不再繼續，並回報「沒展開完」讓呼叫端擋下來 —— 沒看完的東西不能當成沒問題。
            /// </summary>
            public static Expanded ExpandDerivations(string? dax, Derivations d) {
                var sb   = new System.Text.StringBuilder(dax ?? "");
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var x    = new Expanded { DefinitionsFrom = sb.Length };
                const int MaxChars = 4_000_000;

                void Append(string definition) {
                    sb.Append('\n');
                    x.Boundaries.Add(sb.Length);
                    sb.Append(definition);
                }

                while (true) {
                    bool added = false;
                    string cur = sb.ToString();

                    foreach (var tok in ScanDax(cur, x.DefinitionsFrom, x.Boundaries).Tokens) {
                        if (!seen.Add(tok.Name)) continue;
                        if (d.ByToken.TryGetValue(tok.Name, out var exprs)) {
                            foreach (var expr in exprs) Append(expr);
                            added = true;
                        }
                    }
                    // 表名沒有中括號可循，只能看查詢文字有沒有提到它。
                    // 寧可過度納入 —— 多接一段定義只會讓掃描更嚴，不會放水。
                    foreach (var kv in d.ByTable) {
                        if (seen.Contains("\t" + kv.Key)) continue;
                        if (cur.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) < 0) continue;
                        seen.Add("\t" + kv.Key);
                        foreach (var def in kv.Value) Append(def);
                        added = true;
                    }
                    if (!added) break;
                    if (sb.Length > MaxChars) { x.Incomplete = true; break; }
                }
                x.Text = sb.ToString();
                return x;
            }

            /// <summary>
            /// 這次查詢適用的保護規則：這個模型的逐欄設定 ＞ 通用樣式（appsettings.json）＞ 開放。
            /// </summary>
            public sealed class Policy {
                readonly Dictionary<string, Level> _explicit;
                readonly Derivations? _model;

                public Policy(Dictionary<string, Level>? explicitLevels, Derivations? model) {
                    _explicit = explicitLevels ?? new Dictionary<string, Level>();
                    _model    = model;
                }

                public static string Key(string table, string column) =>
                    table.Trim().ToUpperInvariant() + "\u0001" + column.Trim().ToUpperInvariant();

                /// <summary>
                /// 通用樣式給這個欄名的等級。numeric：已知是數值欄位 true、已知不是 false、不知道（別名）null。
                /// 金額樣式只對數值欄位給「只能彙總」—— 文字欄位的 MAX / MIN 回傳的是一個真實的值，不是統計量，
                /// 所以名稱像金額的文字欄位（amount_note 之類）歸到「只能計數」。
                /// </summary>
                public static Level PatternLevel(string column, bool? numeric, out string? rule) {
                    if (Matches(column, AllowColumns, out rule)) return Level.Open;
                    if (Matches(column, DenyColumns,  out rule)) return Level.CountOnly;
                    if (Matches(column, MoneyColumns, out rule)) return numeric == false ? Level.CountOnly : Level.AggregateOnly;
                    rule = null;
                    return Level.Open;
                }

                bool? IsNumeric(string table, string column) =>
                    _model != null && _model.Catalog.TryGetValue(table, out var cols) && cols.TryGetValue(column, out var n) ? n : null;

                /// <summary>模型裡的某一欄的等級，以及這個等級是誰決定的。</summary>
                public Level OfColumn(string table, string column, out string source, out string? rule) {
                    bool? numeric = IsNumeric(table, column);
                    if (_explicit.TryGetValue(Key(table, column), out var set)) {
                        source = "explicit"; rule = null;
                        // 設定檔被手動改成「文字欄位只能彙總」的話，照樣當成只能計數（理由同 PatternLevel）
                        return set == Level.AggregateOnly && numeric == false ? Level.CountOnly : set;
                    }
                    var level = PatternLevel(column, numeric, out rule);
                    source = rule != null ? "pattern" : "default";
                    return level;
                }
                public Level OfColumn(string table, string column) => OfColumn(table, column, out _, out _);

                /// <summary>
                /// 查詢文字或結果欄位裡的一個參考，適用哪個等級。
                ///   · 表名明確、而且那張表真的有這一欄 → 那一欄的等級
                ///   · 沒有表名（或表名不是模型的資料表）→ 所有同名資料行裡最嚴的
                ///   · 模型裡沒有這個欄名（別名、查詢自己定義的欄位）→ 只看通用樣式
                /// </summary>
                public Level Resolve(string? table, string column) {
                    if (_model != null && _model.CatalogLoaded) {
                        if (table != null && _model.Catalog.TryGetValue(table, out var cols) && cols.ContainsKey(column))
                            return OfColumn(table, column);
                        var owners = _model.TablesWithColumn(column);
                        if (owners.Count > 0) return Combine(owners.Select(t => OfColumn(t, column)));
                        return PatternLevel(column, null, out _);
                    }
                    // 欄位目錄讀不到：分不出這個參考是哪張表的欄位 —— 同名的逐欄設定全部算進來取最嚴的。
                    // 逐欄設成「開放」的在這裡不會放寬任何東西（Combine 會略過開放）。
                    var levels = new List<Level> { PatternLevel(column, null, out _) };
                    string suffix = "\u0001" + column.Trim().ToUpperInvariant();
                    foreach (var kv in _explicit)
                        if (kv.Key.EndsWith(suffix, StringComparison.Ordinal)) levels.Add(kv.Value);
                    return Combine(levels);
                }

                /// <summary>這張表有沒有任何一欄不是開放的。</summary>
                public bool HasRestrictedColumn(string table) =>
                    _model != null && _model.Catalog.TryGetValue(table, out var cols)
                    && cols.Keys.Any(c => OfColumn(table, c) != Level.Open);
            }

            // ── 建議 ──────────────────────────────────────────────────────────
            // 給設定畫面的提示：目前是「開放」、但欄名看起來像敏感資料的欄位。
            // 只是提示 —— 不影響查詢的判定，使用者沒按套用就什麼都不會變。
            // 通用樣式（appsettings.json）預設只認得英文；這裡補上中文欄名，讓第一次設定不必從零開始。
            static readonly string[] IdentityWords = {
                "客戶", "顧客", "客人", "公司名", "廠商", "供應商", "經銷", "代理商", "聯絡", "姓名", "人名", "員工",
                "業務員", "業務代表", "業務姓名", "負責人", "窗口", "電話", "手機", "信箱", "郵件", "地址",
                "統編", "統一編號", "身分證", "護照", "帳號", "帳戶",
                "customer", "client", "account", "contact", "partner", "vendor", "supplier", "dealer", "distributor",
                "owner", "person", "employee", "staff", "salesrep", "salesperson", "email", "phone", "mobile", "address", "passport"
            };
            static readonly string[] FreeTextWords = {
                "備註", "註記", "說明", "描述", "意見", "留言", "memo", "comment", "remark", "note", "description"
            };
            static readonly string[] MoneyWords = {
                "金額", "單價", "價格", "售價", "成本", "毛利", "營收", "收入", "薪資", "薪水", "費用",
                "amount", "amt", "price", "revenue", "cost", "margin", "salary", "profit"
            };

            public static Level? Suggest(string column, bool numeric, out string? why) {
                string n = Norm(column);
                foreach (var w in IdentityWords)
                    if (n.Contains(w)) { why = $"欄名有「{w}」，看起來是可以辨識出特定人或公司的資料"; return Level.CountOnly; }
                foreach (var w in FreeTextWords)
                    if (n.Contains(w)) { why = $"欄名有「{w}」，自由填寫的文字裡常夾著姓名與細節"; return Level.CountOnly; }
                if (numeric)
                    foreach (var w in MoneyWords)
                        if (n.Contains(w)) { why = $"欄名有「{w}」，看起來是金額"; return Level.AggregateOnly; }
                why = null;
                return null;
            }

            static string Shown(Tok t) => t.Table != null ? $"'{t.Table}'[{t.Name}]" : $"[{t.Name}]";

            // 篩選條件的位置：FILTER 的條件、CALCULATE / CALCULATETABLE 的篩選引數。
            // 寫在那裡的運算式只決定留下哪些列，本身不會變成輸出。
            static readonly HashSet<string> FilterFuncs = new(StringComparer.OrdinalIgnoreCase) { "FILTER", "CALCULATE", "CALCULATETABLE" };
            static bool IsFilterSlot(Frame f) => f.Arg >= 1 && FilterFuncs.Contains(f.Func);

            /// <summary>
            /// 查詢連同它展開出來的所有定義的指紋。使用者在確認視窗同意的是「這些定義當下的內容」——
            /// 視窗開著的期間有人把量值換掉，指紋就對不上，那次同意不算數。
            /// 定義先排序再算：系統檢視回傳的順序不保證每次一樣。
            /// </summary>
            public static string Fingerprint(Expanded x) {
                var cuts  = x.Boundaries.Where(b => b <= x.Text.Length).OrderBy(b => b).ToList();
                var parts = new List<string>();
                int head  = cuts.Count > 0 ? cuts[0] : x.Text.Length;
                for (int i = 0; i < cuts.Count; i++) {
                    int end = i + 1 < cuts.Count ? cuts[i + 1] : x.Text.Length;
                    parts.Add(x.Text.Substring(cuts[i], end - cuts[i]));
                }
                parts.Sort(StringComparer.Ordinal);
                using var sha = System.Security.Cryptography.SHA256.Create();
                return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(
                    x.Text.Substring(0, head) + "\u0001" + string.Join("\u0001", parts))));
            }

            /// <summary>
            /// 引擎的錯誤訊息會引用資料：型別轉換失敗時，訊息裡就是那個轉不過去的值。
            /// 查詢碰到受限欄位時，訊息裡用引號括起來的片段只留「查詢自己寫的」與「模型裡的名稱」，其餘換掉 ——
            /// 錯誤的種類、位置、函式名都還在，改公式夠用；原文印在服務主控台。
            /// </summary>
            public static string RedactError(string? message, string? query, Derivations? model) {
                string q = query ?? "";
                static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_';
                // 必須是查詢裡「完整的一個字」：短的值（例如 TW）剛好是某個欄名的一部分不算
                bool InQuery(string s) {
                    for (int i = q.IndexOf(s, StringComparison.OrdinalIgnoreCase); i >= 0;
                         i = i + 1 < q.Length ? q.IndexOf(s, i + 1, StringComparison.OrdinalIgnoreCase) : -1) {
                        bool left  = i == 0 || !IsWordChar(q[i - 1]) || !IsWordChar(s[0]);
                        bool right = i + s.Length >= q.Length || !IsWordChar(q[i + s.Length]) || !IsWordChar(s[^1]);
                        if (left && right) return true;
                    }
                    return false;
                }
                bool Known(string s) =>
                    s.Length == 0 || !s.Any(char.IsLetterOrDigit) || InQuery(s)
                    || (model != null && (model.MeasureNames.Contains(s) || model.Catalog.ContainsKey(s)
                                          || model.ColumnNames.Contains(Norm(s))));
                string masked = Regex.Replace(message ?? "",
                    @"'([^'\r\n]*)'|""([^""\r\n]*)""|「([^」\r\n]*)」|‘([^’\r\n]*)’|“([^”\r\n]*)”", m => {
                        string inner = "";
                        for (int g = 1; g < m.Groups.Count; g++) if (m.Groups[g].Success) { inner = m.Groups[g].Value; break; }
                        return Known(inner) ? m.Value : $"{m.Value[0]}‹已遮蔽›{m.Value[^1]}";
                    });
                // 值本身帶引號（O'Brien）時，上面只配到前半段，後半段會露在引號外面 —— 遮罩後面緊接著的字一起遮掉
                return Regex.Replace(masked, @"(‹已遮蔽›['""」’”])[\p{L}\p{N}_][^\s'""」’”]*(['""」’”])?", "$1");
            }

            /// <summary>
            /// 確認視窗用：這句查詢（含展開的定義）提到了哪些受限欄位。
            /// 不帶中括號引用的資料表，底下每一個受限欄位都算 —— 整表 EVALUATE 的查詢文字裡沒有任何欄名。
            /// </summary>
            public static List<string> RestrictedMentions(Scan scan, Policy policy, Derivations? model) {
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var list = new List<string>();
                void Add(string? table, string name, Level level) {
                    string shown = table != null ? $"'{table}'[{name}]" : $"[{name}]";
                    if (seen.Add(shown)) list.Add($"{shown}（{LevelName(level)}）");
                }
                foreach (var t in scan.Tokens) {
                    if (model != null && model.IsMeasureOnly(t.Name)) continue;
                    var level = policy.Resolve(t.Table, t.Name);
                    if (level != Level.Open) Add(t.Table, t.Name, level);
                }
                if (model != null && model.CatalogLoaded)
                    foreach (var b in scan.Bare)
                        if (model.Catalog.TryGetValue(b.Name, out var cols))
                            foreach (var c in cols.Keys) {
                                var level = policy.OfColumn(b.Name, c);
                                if (level != Level.Open) Add(b.Name, c, level);
                            }
                return list;
            }

            /// <summary>
            /// 第一關：執行前掃描查詢文字（scan 必須是「已展開定義」的掃描結果）。null 表示放行。
            /// </summary>
            /// hard：被擋的原因是「這種寫法本身」（整個禁用的函式、ERROR），不是某個欄位的等級 —— 使用者的單次同意也不放行，
            /// 因為那些寫法帶出去的量不受「最多幾列」限制（TOCSV 一格就是整張表）。
            public static string? CheckQuery(Scan scan, Policy policy, Derivations? model, out bool touchesMoney, out bool hard) {
                touchesMoney = false;
                hard = false;
                if (!Enabled) return null;
                hard = true;        // ① 與 ERROR 那兩段回傳的都算；過了之後在 ② 之前改回 false

                // ── ① 整個禁用的函式 ─────────────────────────────────────────
                foreach (var call in scan.Calls) {
                    if (BannedCalls.Contains(call))
                        return $"{call.ToUpperInvariant()}() 在資料保護啟用時不能用：它不必指名欄位就能把資料表的內容帶出來，"
                             + "欄位層級的管制看不到它。要看結構請用 /api/schema，要看數字請用彙總查詢。";
                    if (call.StartsWith("INFO.", StringComparison.OrdinalIgnoreCase))
                        return $"{call.ToUpperInvariant()}() 在資料保護啟用時不能用：INFO 系列函式會回傳 M 腳本與運算式原文。"
                             + "模型結構請用 /api/schema、/api/validate。";
                }

                // ERROR() 會把任意文字放進引擎的錯誤訊息，而錯誤訊息會回給呼叫端 ——
                // COUNTROWS(FILTER(客戶, ERROR(客戶[名稱]))) 外面有計數函式，下面的逐欄檢查會放行，名稱卻從錯誤訊息出去了。
                // 用變數轉一手就看不出 ERROR 的引數是什麼，所以查詢（含展開的定義）只要碰到受限欄位，就整個不准用。
                if (scan.Calls.Contains("ERROR") && RestrictedMentions(scan, policy, model).Count > 0)
                    return "ERROR() 不能和受限欄位出現在同一句查詢裡（包含查詢引用到的量值）："
                         + "它會把任意文字放進錯誤訊息，受限欄位的內容可以從那裡出去。";

                hard = false;

                // ── ② 每一個欄位參考 ─────────────────────────────────────────
                var refs = new List<(Tok Tok, bool Implicit)>(scan.Tokens.Count);
                foreach (var t in scan.Tokens) refs.Add((t, false));
                // 只有一個資料行的資料表：不帶中括號引用它，等於引用那個資料行 ——
                // 1×1 的資料表會自動轉成純量，VAR v = TOPN(1, 單欄表) RETURN ROW("x", v) 一個中括號都不用寫。
                if (model != null && model.CatalogLoaded)
                    foreach (var b in scan.Bare)
                        if (model.Catalog.TryGetValue(b.Name, out var cols) && cols.Count == 1)
                            refs.Add((new Tok { Table = b.Name, Name = cols.Keys.First(),
                                                Enclosing = b.Enclosing, InDefinition = b.InDefinition }, true));

                foreach (var (t, isImplicit) in refs) {
                    // 量值參考不是欄位。[Total Revenue]、[Customer Rank] 這種名稱會符合管制樣式，
                    // 但量值回傳的是它的定義算出來的純量 —— 而那份定義已經由 ExpandDerivations
                    // 接在後面，會在這個迴圈裡逐一檢查，真正碰到敏感欄位的地方照樣受管制。
                    if (!isImplicit && model != null && model.IsMeasureOnly(t.Name)) continue;

                    var level = policy.Resolve(t.Table, t.Name);
                    if (level == Level.Open) continue;
                    bool counted = t.Enclosing.Any(f => CountingFuncs.Contains(f.Func));
                    string what  = isImplicit ? $"資料表 '{t.Table}'（它唯一的資料行 [{t.Name}]）" : Shown(t);
                    // 模型裡沒有這個資料行：多半是查詢自己取的別名，剛好長得像受限欄位。講清楚，否則訊息會誤導成「你沒聚合」
                    string aliasHint = !isImplicit && t.Table == null && model != null && model.CatalogLoaded
                                       && model.TablesWithColumn(t.Name).Count == 0
                        ? $"（模型裡沒有叫 [{t.Name}] 的資料行。如果它是你在查詢裡取的別名，這個名稱符合通用規則的樣式 —— "
                        + "換一個不像受限欄位的別名即可，例如 \"值\"。）" : "";

                    switch (level) {
                        case Level.CountOnly:
                            if (!counted)
                                return $"{what} 設為「只能計數」，不能當分組鍵或直接取值。"
                                     + "只能用計數類函式取統計量，例如 DISTINCTCOUNT / COUNTROWS / COUNTBLANK。"
                                     + "（MAX / MIN / CONCATENATEX 不算 —— 它們會回傳真實內容。）" + aliasHint;
                            break;

                        case Level.Pseudonym:
                            if (counted) break;
                            // 當分組鍵：必須直接寫在查詢本體裡（不能藏在變數或定義裡），而且一路往外都是分組函式 ——
                            // 這樣它才會以原名出現在結果欄位，值才換得成代號。沒有任何函式包住（例如 ORDER BY）也不算。
                            if (t.InDefinition || t.Enclosing.Length == 0 || !t.Enclosing.All(IsGroupingFrame))
                                return $"{what} 設為「換成代號」：可以計數，或直接當分組鍵"
                                     + "（SUMMARIZECOLUMNS / SUMMARIZE / GROUPBY / VALUES / DISTINCT，外面可以再包一層 TOPN），"
                                     + "結果會顯示成代號。不能放進變數、量值，也不能用 MAX / CONCATENATEX / SELECTCOLUMNS 取值或拿來比較。" + aliasHint;
                            break;

                        case Level.AggregateOnly:
                            touchesMoney = true;
                            // 寫在篩選條件裡不算取值。SUMX(FILTER(銷售, 銷售[金額] > 1000), …) 本來就放行，
                            // 同一個條件寫成 CALCULATE(SUM(…), 銷售[金額] > 1000) 沒有理由被擋 ——
                            // 模型裡既有的量值常常就是這樣寫的，而呼叫端沒辦法改寫別人的量值。
                            if (!counted && !t.Enclosing.Any(f => NumericAggFuncs.Contains(f.Func)) && !t.Enclosing.Any(IsFilterSlot))
                                return $"{what} 設為「只能彙總」，必須包在聚合函式內（SUM / AVERAGE / SUMX / MIN / MAX …），不能逐列取值。" + aliasHint;
                            break;
                    }
                }

                // ── ③ UNION / TREATAS：輸出欄位的名字不是來源欄位的名字 ──────────
                // UNION 的欄名取自第一個引數。UNION(ROW("a", 1, "b", 2, …), Customers) 會把 Customers 的
                // 每一列掛在 a、b… 這些無害的別名底下回傳，整句查詢沒有任何中括號參考，結果欄位也認不出來。
                // 所以：不帶中括號的資料表（或變數 —— 分不出裡面裝什麼）如果一路只經過「原樣傳遞資料列」的函式
                // 就到了 UNION，擋下。中間有 SELECTCOLUMNS / SUMMARIZE 這類重新指定欄位的函式就沒事。
                if (scan.Calls.Contains("UNION") || scan.Calls.Contains("TREATAS")) {
                    foreach (var b in scan.Bare) {
                        int reach = -1;
                        for (int k = b.Enclosing.Length - 1; k >= 0; k--) {
                            string f = b.Enclosing[k].Func;
                            if (f.Equals("UNION", StringComparison.OrdinalIgnoreCase) || f.Equals("TREATAS", StringComparison.OrdinalIgnoreCase)) { reach = k; break; }
                            if (ProjectingFuncs.Contains(f) || CountingFuncs.Contains(f) || NumericAggFuncs.Contains(f)) break;
                        }
                        if (reach < 0) continue;
                        if (b.Enclosing.Take(reach).Any(f => CountingFuncs.Contains(f.Func))) continue;   // 外面有計數函式：結果只是數字

                        bool isVar = scan.VarNames.Contains(b.Name);
                        bool risky = isVar
                                  || (model != null && model.CatalogLoaded ? policy.HasRestrictedColumn(b.Name)
                                                                           : true);    // 沒有目錄就分不出來，一律當成有
                        if (risky)
                            return $"{b.Enclosing[reach].Func.ToUpperInvariant()}() 的引數裡不能直接放"
                                 + (isVar ? $"變數 {b.Name}" : $"含受限欄位的資料表 '{b.Name}'")
                                 + "：輸出欄位會改用第一個引數的名稱，原本的欄位名就看不見了。"
                                 + "請用 SELECTCOLUMNS / SUMMARIZE 明確指定要哪些欄位，再放進去。";
                    }
                }
                return null;
            }

            /// <summary>
            /// 第二關：執行後、讀取任何一列之前，檢查結果欄位。null 表示放行；
            /// mask 是要把值換成代號的欄位索引。
            ///
            /// 模型的資料行在結果裡是 Table[Column]，查得到它是哪一欄；查詢自己取的別名只有 [別名]。
            /// </summary>
            public static string? CheckResultColumns(IReadOnlyList<string> columns, Policy policy, Derivations? model,
                                                     out HashSet<int> mask) {
                mask = new HashSet<int>();
                if (!Enabled) return null;
                for (int i = 0; i < columns.Count; i++) {
                    string s  = columns[i] ?? "";
                    int    lb = s.LastIndexOf('[');
                    string bare = (lb >= 0 ? s.Substring(lb + 1) : s).TrimEnd(']').Trim();
                    bool   isAlias = lb <= 0;

                    // 表名或欄名本身可以含有 [ ]，所以每一個 [ 都試著當成分界：
                    // 只要有一種切法對得上模型裡真的存在的 表[欄]，就照那一欄的等級（對上不只一種取最嚴的）。
                    var hits = new List<Level>();
                    if (!isAlias && s.EndsWith("]") && model != null && model.CatalogLoaded) {
                        for (int idx = s.IndexOf('['); idx > 0; idx = s.IndexOf('[', idx + 1)) {
                            string t = s.Substring(0, idx).Trim();
                            if (t.Length >= 2 && t[0] == '\'' && t[^1] == '\'') t = t.Substring(1, t.Length - 2).Replace("''", "'");
                            string c = s.Substring(idx + 1, s.Length - idx - 2);
                            foreach (var name in new[] { c, c.Replace("]]", "]") })
                                if (model.Catalog.TryGetValue(t, out var cols) && cols.ContainsKey(name))
                                    hits.Add(policy.OfColumn(t, name));
                        }
                    }

                    Level level;
                    if (hits.Count > 0) level = Combine(hits);
                    else if (isAlias && model != null && model.IsMeasureOnly(bare)) continue;   // 別名取成量值的名字：ROW("銷售總額", [銷售總額])
                    else level = policy.Resolve(null, bare);

                    // 別名被擋時要講清楚，否則訊息會誤導成「你沒聚合」
                    string aliasHint = isAlias
                        ? $"若 [{bare}] 是你在查詢裡取的別名，換一個不像受限欄位的名稱即可（例如 \"值\"）。" : "";

                    switch (level) {
                        case Level.CountOnly:
                            return $"查詢結果含「只能計數」的欄位 [{bare}]。整表 EVALUATE 會把原始欄位整批帶出來。{aliasHint}";
                        case Level.AggregateOnly:
                            return $"查詢結果含未經聚合的「只能彙總」欄位 [{bare}]。請改以 SUM / AVERAGE 等聚合後再取別名。{aliasHint}";
                        case Level.Pseudonym:
                            mask.Add(i);
                            break;
                    }
                }
                return null;
            }

            /// <summary>稽核記錄。寫的是查詢文字與判定結果，不寫任何查詢回傳值。</summary>
            public static void Audit(string verdict, string detail, string query, int rows, string? file) {
                if (string.IsNullOrEmpty(AuditPath)) return;
                try {
                    static string Flat(string? s) =>
                        (s ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
                    File.AppendAllText(AuditPath,
                        string.Join("\t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), verdict,
                                    Flat(file), rows.ToString(), Flat(detail), Flat(query))
                        + Environment.NewLine,
                        new System.Text.UTF8Encoding(true));
                } catch {
                    // 稽核寫入失敗不該讓查詢跟著失敗，但也不能無聲無息
                    Console.WriteLine("⚠️ 稽核記錄寫入失敗");
                }
            }
        }

        // 資料保護啟用時 /api/dmv 可以查的系統檢視：只有模型的中繼資料與儲存統計，沒有任何一個會回傳欄位內容。
        static readonly HashSet<string> AllowedDmvViews = new(StringComparer.OrdinalIgnoreCase) {
            "TMSCHEMA_MODEL", "TMSCHEMA_TABLES", "TMSCHEMA_COLUMNS", "TMSCHEMA_MEASURES", "TMSCHEMA_RELATIONSHIPS",
            "TMSCHEMA_HIERARCHIES", "TMSCHEMA_LEVELS", "TMSCHEMA_ROLES", "TMSCHEMA_TABLE_PERMISSIONS",
            "TMSCHEMA_CALCULATION_GROUPS", "TMSCHEMA_CALCULATION_ITEMS", "TMSCHEMA_PERSPECTIVES",
            "TMSCHEMA_TABLE_STORAGES", "TMSCHEMA_COLUMN_STORAGES", "TMSCHEMA_PARTITION_STORAGES",
            "TMSCHEMA_SEGMENT_STORAGES", "TMSCHEMA_SEGMENTS",
            "DISCOVER_STORAGE_TABLES", "DISCOVER_STORAGE_TABLE_COLUMNS", "DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS",
            "DISCOVER_OBJECT_MEMORY_USAGE", "DBSCHEMA_CATALOGS"
        };

        // =====================================================================
        // 逐欄保護設定（每個模型一份）
        // =====================================================================

        // 存逐欄設定要認得出這份模型：檔案路徑、報表名稱（視窗標題）、模型裡資料行的固定代號，有一樣就行。
        const string NoIdentityMessage =
            "❌ 服務認不出這份報表（沒有檔名，模型裡也沒有可以辨識的資料行），沒有地方記錄逐欄設定。"
          + "請先在 Power BI 存檔，再重新掃描。";

        /// <summary>模型的代號：檔名 + 完整路徑的雜湊。快照資料夾用它；保護設定檔優先用它（見 ProtectionStore.Find）。</summary>
        static string? ModelKey(PbiInstance inst) {
            if (string.IsNullOrEmpty(inst.FilePath)) return null;
            var name = Regex.Replace(Path.GetFileNameWithoutExtension(inst.FilePath), @"[\\/:*?""<>|]", "_");
            using var sha = System.Security.Cryptography.SHA256.Create();
            var hash = Convert.ToHexString(
                sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(inst.FilePath.ToLowerInvariant())))[..8];
            return $"{name}_{hash}";
        }

        /// <summary>
        /// API 金鑰：服務第一次啟動時自己產生，存在 %LOCALAPPDATA%\PBI_AI_Bridge\api-key.txt。
        ///
        /// 以前放在專案資料夾的 appsettings.json 裡，問題是那個資料夾會被複製、壓縮、分享 ——
        /// 金鑰跟著出去（bin\ 裡還有一份建置時複製的副本）。搬出來之後：
        ///   · 專案資料夾裡沒有任何機密，appsettings.json 可以放心打開、放心分享
        ///   · 使用者資料夾只有這個 Windows 帳號讀得到，正好是金鑰要擋的對象（別的帳號、別的網站）
        ///
        /// 金鑰不是用來防 AI 的 —— AI 的工具本來就要用它呼叫 API。防 AI 的是確認視窗（見 HumanConfirm）。
        ///
        /// 舊版留在 appsettings.json 裡的金鑰**不沿用**：它在專案資料夾裡待過，可能已經跟著資料夾出去了。
        /// 新版第一次啟動就換一把，舊的那把從此無效。代價是「換成代號」的代號會全部換一批（代號由金鑰衍生）。
        /// </summary>
        static class ApiKeyStore {
            public static string FilePath => Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PBI_AI_Bridge", "api-key.txt");

            static bool LooksValid(string k) => k.Length >= 32 && k.Length <= 128 && k.All(char.IsAsciiLetterOrDigit);

            /// <summary>讀出金鑰；沒有（或內容不像金鑰）就產生一把新的。created＝這次是不是新產生的。</summary>
            public static string LoadOrCreate(out bool created) {
                created = false;
                try {
                    if (File.Exists(FilePath)) {
                        string existing = File.ReadAllText(FilePath).Trim();
                        if (LooksValid(existing)) return existing;
                    }
                } catch (IOException) { /* 讀不到就往下重新產生；寫不進去的話下面會講清楚 */ }
                  catch (UnauthorizedAccessException) { }

                try {
                    string key = Convert.ToHexString(System.Security.Cryptography.RandomNumberGenerator.GetBytes(32));
                    Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
                    // 先寫到旁邊再換過去：寫到一半斷掉不會留下半把金鑰
                    string tmp = FilePath + ".tmp";
                    File.WriteAllText(tmp, key, new System.Text.UTF8Encoding(false));
                    File.Move(tmp, FilePath, overwrite: true);
                    created = true;
                    return key;
                } catch (Exception ex) {
                    throw new Exception($"❌ 無法建立 API 金鑰檔（{FilePath}）：{ex.Message}");
                }
            }
        }

        /// <summary>
        /// 使用者在儀表板「資料保護」分頁替某個模型逐欄設定的等級。
        ///
        /// 存在 %LOCALAPPDATA%\PBI_AI_Bridge\protection\，刻意不放在專案資料夾：
        /// 專案資料夾是 AI 代理平常讀寫檔案的地方，把「AI 能讀什麼」的設定放在那裡，等於把鑰匙掛在門上。
        ///
        /// 設定檔只在服務啟動後第一次用到時從磁碟讀一次，之後以記憶體為準、只由設定端點寫入。
        /// 所以直接改那些檔案不會立刻生效 —— 要等服務重開，和 appsettings.json 一樣需要人動手。
        ///
        /// 一份模型對到哪一份設定檔（見 Find）：
        ///   ① 完整路徑對得上的那一份
        ///   ② 沒有的話，靠內容認 —— 設定檔記著模型裡一部分資料行的固定代號（LineageTag），打開的模型裡有同樣的代號就是同一份。
        ///      服務只能從 Power BI 的啟動參數得知檔案路徑：先開 Power BI 再從裡面選檔案，路徑就拿不到；
        ///      檔案搬了家、另存新檔，路徑也會變。靠內容認，這幾種情況設定都還跟著模型。
        ///
        /// 檔案讀不出來（損毀、被改壞）時**不當成「沒有設定」**：那等於把檔案弄壞就能解除保護。
        /// 讀不出來 → 這個模型的查詢一律擋下，直到使用者在儀表板重設。
        /// </summary>
        static class ProtectionStore {
            public sealed class Entry {
                public string Table  { get; set; } = "";
                public string Column { get; set; } = "";
                public string Level  { get; set; } = "";
                public string? Tag   { get; set; }           // 資料行的 LineageTag：改名之後靠它找回來
            }
            sealed class FileShape {
                public int Version { get; set; } = 2;
                public string? File { get; set; }            // 上次儲存時這份模型的檔案（不知道路徑時是視窗標題）
                public string? SavedAt { get; set; }
                public List<string> Signature { get; set; } = new();     // 模型裡一部分資料行的固定代號，用來認模型
                public List<Entry> Columns { get; set; } = new();
            }
            /// <summary>讀進記憶體的一份設定檔。</summary>
            public sealed class Loaded {
                public string Key = "";
                public List<Entry> Entries = new();
                public HashSet<string> Tags = new(StringComparer.OrdinalIgnoreCase);    // Signature 加上每筆設定的固定代號
                public string? Problem, SavedAt, File;
            }
            /// <summary>這份模型對到的設定。</summary>
            public sealed class Match {
                public Loaded? File;            // null＝這份模型沒有設定過
                public string  Key = "";        // 要讀寫的檔名（不含副檔名）；空字串＝沒有任何東西可以認這份模型，存不了
                public bool    Exact;           // 照完整路徑（或在沒有固定代號可比時照名稱）對上的；false＝靠內容認出來的
                public string? Problem => File?.Problem;
                /// <summary>
                /// 靠內容認出來的設定，只照固定代號套用，不照「表名＋欄名」。
                /// 內容相符只代表兩份模型有共同的來源；名稱一樣的欄位未必是同一欄 ——
                /// 不這樣限制的話，別份模型上「把某欄開放」的決定，會套到這份模型同名的欄位上。
                /// </summary>
                public bool TagsOnly => File != null && !Exact;
                public List<Entry> Entries() =>
                    (File?.Entries ?? new List<Entry>()).Select(e => new Entry { Table = e.Table, Column = e.Column, Level = e.Level, Tag = e.Tag }).ToList();
            }

            static readonly object _lock = new();
            static readonly Dictionary<string, Loaded> _files = new(StringComparer.OrdinalIgnoreCase);
            static bool _scanned;
            static readonly System.Text.Json.JsonSerializerOptions _json = new() {
                PropertyNameCaseInsensitive = true, WriteIndented = true,
                PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            public static string Folder => Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PBI_AI_Bridge", "protection");

            static string PathFor(string key) => Path.Combine(Folder, key + ".json");

            static string Hash8(string s) {
                using var sha = System.Security.Cryptography.SHA256.Create();
                return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(s.ToLowerInvariant())))[..8];
            }

            /// <summary>這份報表的名字（檔名；沒有路徑時是視窗標題），整理成可以當檔名的樣子。</summary>
            static string NameOf(PbiInstance inst) {
                string raw = !string.IsNullOrEmpty(inst.FilePath) ? Path.GetFileNameWithoutExtension(inst.FilePath)
                           : inst.FileName ?? "";
                string name = Regex.Replace(raw, @"[\\/:*?""<>|\p{Cc}]", "_").Trim().TrimEnd('.');
                return name.Length <= 80 ? name : name.Substring(0, 80);
            }

            static Loaded ToLoaded(string key, FileShape shape) {
                var l = new Loaded { Key = key, Entries = shape.Columns, SavedAt = shape.SavedAt, File = shape.File };
                foreach (var t in shape.Signature) if (!string.IsNullOrWhiteSpace(t)) l.Tags.Add(t);
                foreach (var e in shape.Columns) if (!string.IsNullOrWhiteSpace(e.Tag)) l.Tags.Add(e.Tag!);
                return l;
            }

            static Loaded Read(string path) {
                string key = Path.GetFileNameWithoutExtension(path);
                try {
                    var shape = System.Text.Json.JsonSerializer.Deserialize<FileShape>(File.ReadAllText(path), _json)
                                ?? throw new InvalidDataException("檔案是空的");
                    shape.Columns   ??= new List<Entry>();
                    shape.Signature ??= new List<string>();
                    foreach (var e in shape.Columns) {
                        if (string.IsNullOrWhiteSpace(e.Table) || string.IsNullOrWhiteSpace(e.Column)
                            || !DataGuard.TryParseLevel(e.Level, out _))
                            throw new InvalidDataException($"有一筆設定看不懂（{e.Table}[{e.Column}] = {e.Level}）");
                    }
                    // 印出來讓使用者看得到：設定只該由確認過的儲存改動，筆數或時間對不上就是有人直接動過檔案
                    Console.WriteLine($"🛡️ 已載入逐欄保護設定「{key}」：{shape.Columns.Count} 筆（上次儲存 {shape.SavedAt ?? "不明"}）");
                    return ToLoaded(key, shape);
                } catch (Exception ex) {
                    Console.WriteLine($"⚠️ 保護設定檔讀取失敗：{path} —— {ex.Message}");
                    return new Loaded {
                        Key = key,
                        Problem = $"這個模型的資料保護設定檔讀不出來（{ex.Message}）。"
                                + "為了避免設定檔壞掉就等於解除保護，查詢先全部擋下。"
                                + "請到儀表板的「資料保護」分頁按「重設」，再重新設定。"
                    };
                }
            }

            static void EnsureLoaded() {
                if (_scanned) return;
                _scanned = true;
                if (!Directory.Exists(Folder)) return;
                foreach (var path in Directory.GetFiles(Folder, "*.json")) {
                    var l = Read(path);
                    _files[l.Key] = l;
                }
            }

            /// <summary>這份模型用哪一份設定。model 是它現在的欄位目錄（要有固定代號才能靠內容認）。</summary>
            public static Match Find(PbiInstance inst, DataGuard.Derivations? model) {
                lock (_lock) {
                    EnsureLoaded();
                    string? pathKey = ModelKey(inst);
                    string  name    = NameOf(inst);
                    var mine = new HashSet<string>(
                        model != null && model.CatalogLoaded ? model.TagToColumn.Keys : Enumerable.Empty<string>(),
                        StringComparer.OrdinalIgnoreCase);

                    // 沒設定過的話，新的設定檔要叫什麼
                    string newKey = pathKey
                        ?? (name.Length == 0 && mine.Count == 0 ? ""
                            : (name.Length > 0 ? name : "untitled") + "_"
                              + (mine.Count > 0 ? "m" + Hash8(mine.Min(StringComparer.OrdinalIgnoreCase)!) : "nopath"));

                    // ① 完整路徑對得上的那一份
                    if (pathKey != null && _files.TryGetValue(pathKey, out var byPath))
                        return new Match { File = byPath, Key = byPath.Key, Exact = true };

                    // ② 靠內容認
                    Loaded? best = null;
                    bool bestExact = false;
                    foreach (var f in _files.Values) {
                        bool sameName = name.Length > 0 && f.Key.StartsWith(name + "_", StringComparison.OrdinalIgnoreCase);
                        // 讀不出來的檔案看不到內容，只能照檔名認。寧可多擋：同名的報表在重設之前都查不了
                        if (f.Problem != null) {
                            if (sameName) return new Match { File = f, Key = f.Key };
                            continue;
                        }
                        bool hit, exact = false;
                        if (f.Tags.Count > 0 && mine.Count > 0) hit = f.Tags.Overlaps(mine);
                        // 這次讀不到模型的固定代號（或模型本來就沒有）：同名的設定照樣算數，不能因為少讀一份資料就當成沒設定過。
                        // 這種情況 TagsOnly 會讓受限的設定照欄名套用、開放的設定不套用 —— 只會比較嚴。
                        else if (f.Tags.Count > 0) hit = sameName;
                        // 設定檔裡沒有任何固定代號可比：只認同一個名字
                        else { hit = newKey.Length > 0 && f.Key.Equals(newKey, StringComparison.OrdinalIgnoreCase); exact = hit; }
                        if (!hit) continue;
                        if (best == null || string.CompareOrdinal(f.SavedAt ?? "", best.SavedAt ?? "") > 0) { best = f; bestExact = exact; }
                    }
                    return new Match { File = best, Key = best?.Key ?? newKey, Exact = bestExact };
                }
            }

            /// <summary>一筆設定現在指的是模型裡的哪一欄；找不到回 null。</summary>
            public static (string Table, string Column)? Locate(Entry e, DataGuard.Derivations model, bool tagsOnly) {
                if (!string.IsNullOrEmpty(e.Tag) && model.TagToColumn.TryGetValue(e.Tag!, out var byTag)) return byTag;
                if (tagsOnly && !string.IsNullOrEmpty(e.Tag)) return null;
                if (model.Catalog.TryGetValue(e.Table, out var cols) && cols.ContainsKey(e.Column)) return (e.Table, e.Column);
                return null;
            }

            /// <summary>
            /// 把存下來的設定對到模型「現在」的資料行，回傳 Policy 用的對照表。
            ///   · 有 LineageTag 而且模型裡找得到 → 跟著那個資料行走（改名、搬表都不會掉）
            ///   · 否則用存下來的 表＋欄 名稱（tagsOnly 時不用，見 Match.TagsOnly）
            ///   · 兩種都對不上（欄位被刪了，或在沒有 LineageTag 的情況下改了名）→ stale。
            ///     這種設定只要不是「開放」，仍然套用到所有同名的資料行 —— 資料表改名時保護才不會掉
            /// </summary>
            public static Dictionary<string, DataGuard.Level> LevelsOf(List<Entry> entries, DataGuard.Derivations? model,
                                                                        List<Entry>? stale = null, bool tagsOnly = false) {
                var map = new Dictionary<string, DataGuard.Level>();
                void Put(string table, string column, DataGuard.Level level) {
                    string k = DataGuard.Policy.Key(table, column);
                    if (!map.TryGetValue(k, out var old)) map[k] = level;
                    else if (old != level) map[k] = DataGuard.Level.CountOnly;       // 兩筆設定打架：取最嚴
                }
                var orphans = new List<(Entry E, DataGuard.Level L)>();
                foreach (var e in entries) {
                    DataGuard.TryParseLevel(e.Level, out var level);
                    if (model == null || !model.CatalogLoaded) {
                        if (!tagsOnly || level != DataGuard.Level.Open) Put(e.Table, e.Column, level);
                        continue;
                    }
                    var now = Locate(e, model, tagsOnly);
                    if (now != null) Put(now.Value.Table, now.Value.Column, level);
                    else orphans.Add((e, level));
                }
                foreach (var (e, level) in orphans) {
                    stale?.Add(e);
                    if (level == DataGuard.Level.Open) continue;
                    foreach (var t in model!.TablesWithColumn(e.Column)) {
                        string k = DataGuard.Policy.Key(t, e.Column);
                        if (!map.ContainsKey(k)) map[k] = level;
                    }
                }
                return map;
            }

            public static bool Same(List<Entry> a, List<Entry> b) {
                static IEnumerable<string> Lines(List<Entry> l) =>
                    l.Select(e => $"{e.Tag}\u0001{e.Table}\u0001{e.Column}\u0001{e.Level}".ToUpperInvariant()).OrderBy(s => s, StringComparer.Ordinal);
                return Lines(a).SequenceEqual(Lines(b));
            }

            public static Loaded Save(PbiInstance inst, DataGuard.Derivations? model, Match match, List<Entry> entries) {
                if (match.Key.Length == 0) throw new OpException(NoIdentityMessage);
                lock (_lock) {
                    Directory.CreateDirectory(Folder);
                    // 認模型用的固定代號：沿用原本記著的，再補上現在模型裡的一部分。欄位會增刪，留得多一點才不容易斷線
                    var signature = new List<string>();
                    if (match.File != null && match.File.Problem == null) signature.AddRange(match.File.Tags);
                    if (model != null && model.CatalogLoaded)
                        signature.AddRange(model.TagToColumn.Keys.OrderBy(t => t, StringComparer.OrdinalIgnoreCase).Take(64));
                    var shape = new FileShape {
                        File = inst.FilePath ?? inst.FileName,
                        SavedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        Signature = signature.Distinct(StringComparer.OrdinalIgnoreCase).Take(256).ToList(),
                        Columns = entries.OrderBy(e => e.Table, StringComparer.OrdinalIgnoreCase)
                                         .ThenBy(e => e.Column, StringComparer.OrdinalIgnoreCase).ToList()
                    };
                    // 先寫到旁邊再換過去：寫到一半斷掉不會留下半個檔案（那會被當成「讀不出來」而擋下所有查詢）
                    string path = PathFor(match.Key), tmp = path + ".tmp";
                    File.WriteAllText(tmp, System.Text.Json.JsonSerializer.Serialize(shape, _json), new System.Text.UTF8Encoding(false));
                    File.Move(tmp, path, overwrite: true);
                    var loaded = ToLoaded(match.Key, shape);
                    _files[match.Key] = loaded;
                    return loaded;
                }
            }

            // ── 改名 ──────────────────────────────────────────────────────────
            // 通用規則是看欄名的：把 [account name] 改名成 [c1]，樣式就對不上了，那一欄會從「只能計數」變回開放 ——
            // 而改名不需要任何人確認。所以改名前後各算一次每一欄的等級（靠固定代號認同一欄），
            // 變寬鬆的就補一筆逐欄設定，維持改名前的等級。

            public sealed class LevelSnapshot {
                public List<(string? Tag, string Table, string Column, DataGuard.Level Level)> Restricted = new();
            }

            /// <summary>改名之前：記下每一個受限欄位現在的等級。</summary>
            public static LevelSnapshot Capture(PbiInstance inst, DataGuard.Derivations model) {
                var snap  = new LevelSnapshot();
                var match = Find(inst, model);
                if (match.Problem != null) return snap;
                var policy = new DataGuard.Policy(LevelsOf(match.Entries(), model, null, match.TagsOnly), model);
                var tagOf  = new Dictionary<string, string>();
                foreach (var kv in model.TagToColumn) tagOf[DataGuard.Policy.Key(kv.Value.Table, kv.Value.Column)] = kv.Key;
                foreach (var t in model.Catalog)
                    foreach (var c in t.Value.Keys) {
                        var level = policy.OfColumn(t.Key, c);
                        if (level == DataGuard.Level.Open) continue;
                        snap.Restricted.Add((tagOf.TryGetValue(DataGuard.Policy.Key(t.Key, c), out var tag) ? tag : null, t.Key, c, level));
                    }
                return snap;
            }

            /// <summary>改名並存檔之後：把等級釘回去。回傳補了幾筆。</summary>
            public static int Pin(PbiInstance inst, DataGuard.Derivations after, LevelSnapshot before, IReadOnlyList<RenameRequest> renames) {
                var match = Find(inst, after);
                if (match.Problem != null || match.Key.Length == 0) return 0;
                var entries = match.Entries();
                bool changed = false;

                (string Table, string Column) Follow(string table, string column) {
                    foreach (var rn in renames) {
                        string type = (rn.ObjectType ?? "").ToLowerInvariant();
                        if (type == "table" && table.Equals(rn.OldName, StringComparison.OrdinalIgnoreCase)) table = rn.NewName;
                        else if (type == "column" && rn.TableName != null && table.Equals(rn.TableName, StringComparison.OrdinalIgnoreCase)
                                 && column.Equals(rn.OldName, StringComparison.OrdinalIgnoreCase)) column = rn.NewName;
                    }
                    return (table, column);
                }

                // 沒有固定代號可依靠的設定：照改名的內容把名字跟著改過去
                foreach (var e in entries) {
                    if (!string.IsNullOrEmpty(e.Tag) && after.TagToColumn.ContainsKey(e.Tag!)) continue;
                    var (t, c) = Follow(e.Table, e.Column);
                    if (t != e.Table || c != e.Column) { e.Table = t; e.Column = c; changed = true; }
                }

                var policy = new DataGuard.Policy(LevelsOf(entries, after, null, match.TagsOnly), after);
                int pinned = 0;
                foreach (var r in before.Restricted) {
                    (string Table, string Column) now = r.Tag != null && after.TagToColumn.TryGetValue(r.Tag, out var byTag)
                        ? byTag : Follow(r.Table, r.Column);
                    if (!after.Catalog.TryGetValue(now.Table, out var cols) || !cols.ContainsKey(now.Column)) continue;     // 欄位不在了
                    if (!DataGuard.IsLoosening(r.Level, policy.OfColumn(now.Table, now.Column))) continue;
                    string key = DataGuard.Policy.Key(now.Table, now.Column);
                    entries.RemoveAll(e => (r.Tag != null && string.Equals(e.Tag, r.Tag, StringComparison.OrdinalIgnoreCase))
                                           || DataGuard.Policy.Key(e.Table, e.Column) == key);
                    entries.Add(new Entry { Table = now.Table, Column = now.Column, Level = DataGuard.LevelKey(r.Level), Tag = r.Tag });
                    changed = true;
                    pinned++;
                }
                if (changed) Save(inst, after, match, entries);
                return pinned;
            }
        }

        // =====================================================================
        // 請使用者本人確認
        // =====================================================================

        /// <summary>
        /// 由服務自己跳出一個 Windows 確認視窗，使用者按「是」才繼續。
        ///
        /// 為什麼需要：設定端點與 AI 用的是同一把 API 金鑰 —— 只靠金鑰的話，AI 可以自己呼叫端點放寬保護。
        /// 確認視窗是服務行程自己畫的，內容由伺服器決定（列出到底要改什麼），按鈕只有坐在電腦前的人按得到。
        /// 不用「把一組代碼貼到某處」的做法，是因為代碼可以被轉手：使用者把它貼給 AI，同意的就不是他親眼看到的那件事。
        ///
        /// 三種情況會用到：放寬逐欄設定、放行單一句被擋下的查詢、經由 API 改寫 Power Query。
        ///
        /// 預設按鈕是「否」，一段時間沒回應也算否。被拒絕之後短時間內不再詢問，免得被連續跳出的視窗疲勞轟炸到按下去。
        /// 視窗蓋在最上層，但不主動搶鍵盤焦點（沒有 MB_SETFOREGROUND）：Yes/No 訊息方塊按 Y 就是「是」，
        /// 搶到焦點的話，使用者正在別處打的字會落進來，剛好打到 y 就等於同意。同意必須是看過內容之後動手點的。
        /// </summary>
        static class HumanConfirm {
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "MessageBoxTimeoutW")]
            static extern int MessageBoxTimeoutW(IntPtr hWnd, string text, string caption, uint type, ushort langId, uint ms);

            const uint MB_YESNO = 0x4, MB_ICONWARNING = 0x30, MB_DEFBUTTON2 = 0x100,
                       MB_SYSTEMMODAL = 0x1000, MB_TOPMOST = 0x40000;
            const int IDYES = 6, IDTIMEOUT = 32000;
            public const int TimeoutSeconds  = 120;
            public const int CooldownSeconds = 30;

            public enum Answer { Yes, No, Timeout, Busy, CoolingDown, Unavailable }

            static readonly SemaphoreSlim _one = new(1, 1);
            static DateTime _refuseUntil = DateTime.MinValue;

            public static async Task<Answer> AskAsync(string caption, string text) {
                if (DateTime.UtcNow < _refuseUntil) return Answer.CoolingDown;
                if (!await _one.WaitAsync(0)) return Answer.Busy;
                try {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] 🔐 等待使用者在確認視窗回答（{TimeoutSeconds} 秒內沒回應視為「否」）…");
                    int r = await Task.Run(() => {
                        uint flags = MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2 | MB_SYSTEMMODAL | MB_TOPMOST;
                        try { return MessageBoxTimeoutW(IntPtr.Zero, text, caption, flags, 0, (uint)TimeoutSeconds * 1000); }
                        // 沒有這個進入點的系統就不問（回 0 → Unavailable）。退回沒有計時的 MessageBoxW 的話，
                        // 沒人回答時視窗永遠不關，之後每一次確認都會得到「已經有一個視窗開著」
                        catch (EntryPointNotFoundException) { return 0; }
                    });
                    if (r == IDYES) { Console.WriteLine("   使用者按了「是」"); return Answer.Yes; }
                    _refuseUntil = DateTime.UtcNow.AddSeconds(CooldownSeconds);
                    var answer = r == IDTIMEOUT ? Answer.Timeout : r == 0 ? Answer.Unavailable : Answer.No;
                    Console.WriteLine($"   沒有同意（{answer}）");
                    return answer;
                } finally {
                    _one.Release();
                }
            }

            /// <summary>為什麼沒有獲得同意。只講原因；「所以什麼沒有發生」由呼叫端接在後面。</summary>
            public static string Explain(Answer a) => a switch {
                Answer.No          => "你在確認視窗選了「否」。",
                Answer.Timeout     => $"確認視窗 {TimeoutSeconds} 秒內沒有人回答，當成「否」。視窗可能在另一個螢幕，或沒有跳到最前面 —— 再試一次，並按 Alt+Tab 找它。",
                Answer.Busy        => "已經有另一個確認視窗開著，請先回答那一個。",
                Answer.CoolingDown => $"剛剛才按過「否」或逾時，{CooldownSeconds} 秒內不會再跳出確認視窗。",
                _                  => "這台電腦現在跳不出確認視窗（服務可能不是在你的桌面工作階段裡執行的）。"
            };

            /// <summary>隔多久再問才有意義（給畫面倒數用）。</summary>
            public static int RetryAfterSeconds(Answer a) => a switch {
                Answer.Yes or Answer.Unavailable => 0,
                Answer.Busy => 5,
                _ => Math.Clamp((int)Math.Ceiling((_refuseUntil - DateTime.UtcNow).TotalSeconds), 1, CooldownSeconds)
            };
        }

        /// <summary>查詢被資料保護擋下時的原因。/api/query 用它決定要不要、以及怎麼請使用者確認。</summary>
        sealed class GuardOutcome {
            public string? Verdict, Reason, File;
            public List<string> Mentioned = new();     // 查詢裡提到的受限欄位
            public string? Fingerprint;                // 查詢＋展開的定義（見 DataGuard.Fingerprint）
            public bool CanAskUser;                    // 這種擋法能不能請使用者單次放行
        }

        /// <summary>
        /// 繁中 Windows 的主控台預設是 cp950，裝不下 emoji —— 直接輸出會變成一排 "??"，看起來像亂碼。
        /// 包一層：常用的 emoji 換成文字標記，其他主控台編碼裝不下的字元直接拿掉。
        /// 主控台已經是 UTF-8 時不會用到這個類別，emoji 原樣輸出。
        /// </summary>
        sealed class ConsoleSymbolWriter : TextWriter {
            static readonly Dictionary<string, string> Tags = new() {
                ["🔐"] = "[安全]", ["💾"] = "[快照]", ["🛡"] = "[保護]", ["📋"] = "[稽核]",
                ["📂"] = "[實例]", ["🌐"] = "[網頁]", ["🔓"] = "[放行]",
                ["🎯"] = "[目標]", ["⚠"] = "[注意]", ["✅"] = "[OK]", ["❌"] = "[錯誤]", ["⛔"] = "[擋下]",
            };
            readonly TextWriter _inner;
            readonly System.Text.Encoding? _probe;   // 用來判斷「主控台裝不裝得下這個字」

            public ConsoleSymbolWriter(TextWriter inner, int codePage) {
                _inner = inner;
                try {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    _probe = System.Text.Encoding.GetEncoding(codePage,
                        System.Text.EncoderFallback.ExceptionFallback, System.Text.DecoderFallback.ExceptionFallback);
                } catch {
                    _probe = null;   // 拿不到碼頁就只做 emoji 對照，並拿掉代理字元組
                }
            }

            public override System.Text.Encoding Encoding => _inner.Encoding;

            string Clean(string? s) {
                if (string.IsNullOrEmpty(s)) return s ?? "";
                var sb = new System.Text.StringBuilder(s.Length);
                var e = System.Globalization.StringInfo.GetTextElementEnumerator(s);
                while (e.MoveNext()) {
                    var el = e.GetTextElement();
                    if (Tags.TryGetValue(el.Replace("️", ""), out var tag)) { sb.Append(tag); continue; }
                    if (_probe == null) {
                        if (!char.IsSurrogate(el[0])) sb.Append(el);
                        continue;
                    }
                    try { _probe.GetByteCount(el); sb.Append(el); }
                    catch (System.Text.EncoderFallbackException) { /* 主控台裝不下，拿掉 */ }
                }
                return sb.ToString();
            }

            public override void Write(char value)          => Write(value.ToString());
            public override void Write(string? value)       => _inner.Write(Clean(value));
            public override void WriteLine(string? value)   => _inner.WriteLine(Clean(value));
            public override void WriteLine()                => _inner.WriteLine();
            public override void Flush()                    => _inner.Flush();
        }

        static void Main(string[] args) {
            // 主控台不是 UTF-8 時（繁中 Windows 預設 cp950），emoji 會印成 "??" —— 改印文字標記
            var consoleCodePage = Console.OutputEncoding.CodePage;
            if (consoleCodePage != 65001)
                Console.SetOut(TextWriter.Synchronized(new ConsoleSymbolWriter(Console.Out, consoleCodePage)));

            // ✅ 強制將工作目錄設為 DLL 所在位置，防止捷徑啟動時找不到 appsettings.json
            System.IO.Directory.SetCurrentDirectory(System.AppContext.BaseDirectory);

            Console.WriteLine("======================================");
            Console.WriteLine("PBI AI Bridge - 服務啟動中...");
            Console.WriteLine("======================================");

            var builder = WebApplication.CreateBuilder(args);

            // ASP.NET 預設把每個請求都印成 info，會淹沒資料保護的警告 —— 只留警告以上
            builder.Logging.SetMinimumLevel(LogLevel.Warning);

            // 設定只來自 appsettings.json（資料保護的通用規則等）。金鑰不在裡面 —— 見 ApiKeyStore
            builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: false);
            var config          = builder.Configuration;
            var apiKey          = ApiKeyStore.LoadOrCreate(out bool apiKeyCreated);
            // 舊版的設定檔裡還留著一把金鑰（範本的佔位字串不算）。它已經不被接受，提醒使用者可以刪掉
            var legacyKey       = config["Security:ApiKey"];
            bool legacyKeyLeft  = !string.IsNullOrWhiteSpace(legacyKey) && !legacyKey!.StartsWith("__");

            // ⚠️ 這裡刻意不再有「目標 PBIP 檔案」的設定。
            //    這是通用工具，使用者會頻繁切換不同的 PBIX/PBIP —— 寫死一個路徑只會過期。
            //    所有路徑都從執行中的 PBIDesktop 行程推導（見 DiscoverInstances）。

            // 快照資料夾。留空是常態（複製給別人時 template 就是留空）——
            // 此時落在專案根的 snapshots\，不是 AppContext.BaseDirectory。
            // 後者是 bin\Release\net8.0，dotnet clean 或手動刪 bin 會把所有退路一起清掉。
            // 設定裡若給相對路徑，也一律相對專案根解析，理由相同。
            var projectRoot  = FindProjectRoot() ?? AppContext.BaseDirectory;
            var snapshotPath = config["PowerBI:SnapshotPath"];
            snapshotPath = string.IsNullOrWhiteSpace(snapshotPath)
                ? Path.Combine(projectRoot, "snapshots")
                : Path.GetFullPath(snapshotPath, projectRoot);
            Directory.CreateDirectory(snapshotPath);

            // ── 資料保護：欄位層級管制 ────────────────────────────────────────
            // 沒設定就預設關閉 —— 但關閉狀態會在開機訊息裡明講，不會靜悄悄地沒防護。
            DataGuard.Enabled          = config.GetValue<bool?>("DataProtection:Enabled") ?? false;
            DataGuard.DenyColumns      = config.GetSection("DataProtection:DenyColumns").Get<string[]>()
                                         ?? Array.Empty<string>();
            DataGuard.MoneyColumns     = config.GetSection("DataProtection:AggregateOnlyColumns").Get<string[]>()
                                         ?? Array.Empty<string>();
            DataGuard.MaxRowsWithMoney = config.GetValue<int?>("DataProtection:MaxRowsWithMoney") ?? 100;
            DataGuard.AllowColumns     = config.GetSection("DataProtection:AllowColumns").Get<string[]>()
                                         ?? Array.Empty<string>();
            DataGuard.MaxDetailRows    = config.GetValue<int?>("DataProtection:MaxDetailRows")    ?? 50;
            DataGuard.MaxAggregateRows = config.GetValue<int?>("DataProtection:MaxAggregateRows") ?? 300;
            DataGuard.InitPseudonymKey(apiKey);

            var auditDir = Path.Combine(projectRoot, "audit");
            Directory.CreateDirectory(auditDir);
            DataGuard.AuditPath = Path.Combine(auditDir, $"query-audit-{DateTime.Now:yyyyMM}.tsv");

            var allowedOrigins  = config.GetSection("Security:AllowedOrigins").Get<string[]>()
                                  ?? new[] { "http://localhost:5500", "http://127.0.0.1:5500" };
            // "null" 是舊版用 file:// 開儀表板時需要的來源。儀表板現在由本服務提供（同源），用不到了；
            // 留著反而等於允許任何網站的沙箱 iframe（來源同樣是 "null"）呼叫本服務 —— 舊設定檔裡有也一律剔除。
            allowedOrigins = allowedOrigins
                .Where(o => !string.Equals(o, "null", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            // 金鑰本身不印出來（連開頭幾碼也不印）：這個視窗常出現在截圖與分享畫面裡
            Console.WriteLine(apiKeyCreated
                ? "🔐 已產生這台電腦專用的 API 金鑰 —— 存在你的使用者資料夾，不在工具資料夾裡。儀表板與 AI 工具會自己讀"
                : "🔐 API 金鑰已載入（存在你的使用者資料夾，不在工具資料夾裡）");
            Console.WriteLine($"   位置：{ApiKeyStore.FilePath}");
            if (legacyKeyLeft)
                Console.WriteLine("   appsettings.json 裡還留著舊版的 Security.ApiKey —— 那一把已經失效、不再使用，可以把那一行刪掉");
            Console.WriteLine($"💾 快照資料夾: {snapshotPath}");
            if (DataGuard.Enabled) {
                Console.WriteLine($"🛡️ 資料保護啟用 — 通用規則（依欄名樣式）：只能計數 {DataGuard.DenyColumns.Length} 條、"
                                + $"只能彙總 {DataGuard.MoneyColumns.Length} 條、開放 {DataGuard.AllowColumns.Length} 條");
                Console.WriteLine("   逐欄設定：儀表板的「資料保護」分頁（每個模型各一份；放寬時會跳出確認視窗）");
                Console.WriteLine($"   回傳列數上限：逐列明細 {DataGuard.MaxDetailRows} 列、"
                                + $"彙總結果 {DataGuard.MaxAggregateRows} 列、"
                                + $"金額分組 {DataGuard.MaxRowsWithMoney} 列"
                                + "（伺服器強制，呼叫端只能更低）");
                Console.WriteLine($"📋 查詢稽核記錄: {DataGuard.AuditPath}");
            } else {
                Console.WriteLine("⚠️ 資料保護「未啟用」 — 查詢可取出任何欄位內容。"
                                + "若非本意，請檢查 appsettings.json 的 DataProtection:Enabled");
            }

            // 開機時列出偵測到的實例，讓使用者一眼看到「這台現在有哪些 PBI 可以操作」
            try {
                var found = DiscoverInstances();
                if (found.Count == 0) {
                    Console.WriteLine("📂 目前沒有偵測到執行中的 Power BI Desktop（開啟檔案後即可使用，不需重啟本服務）");
                } else {
                    Console.WriteLine($"📂 偵測到 {found.Count} 個 Power BI 實例：");
                    foreach (var i in found)
                        Console.WriteLine($"     · Port {i.Port}  {i.FileName ?? "(未開啟檔案)"}  [{i.Kind}]");
                    if (found.Count > 1)
                        Console.WriteLine("   ⚠️ 有多個實例 —— 請用 X-PBI-Target 標頭指定目標，否則所有請求（包含讀取）都會被拒絕");
                }
            } catch (Exception ex) {
                Console.WriteLine($"⚠️ 實例探索失敗: {ex.Message}");
            }

            // CORS 限制為白名單
            builder.Services.AddCors(options => {
                options.AddDefaultPolicy(policy => {
                    policy.WithOrigins(allowedOrigins)
                          .AllowAnyHeader()
                          .AllowAnyMethod();
                });
            });

            // 在 localhost 5500 啟動
            builder.WebHost.UseUrls("http://localhost:5500");

            var app = builder.Build();
            app.UseCors();

            // 只接受以 localhost / 127.0.0.1 呼叫。擋 DNS rebinding：惡意網站把自己的網域解析到
            // 127.0.0.1 之後，瀏覽器會把它當成同源 —— 沒有這道檢查，它就讀得到帶金鑰的儀表板頁面。
            app.Use(async (context, next) => {
                var host = context.Request.Host.Host;
                bool isLocal = host.Equals("localhost", StringComparison.OrdinalIgnoreCase)
                            || host == "127.0.0.1" || host == "[::1]" || host == "::1";
                if (!isLocal) {
                    context.Response.StatusCode = 400;
                    await context.Response.WriteAsync("❌ 400 只接受以 localhost 連線");
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 拒絕非 localhost 的連線（Host: {host}）");
                    return;
                }
                await next();
            });

            // API Key 驗證（所有 /api/* 都需帶 X-API-Key Header）
            app.Use(async (context, next) => {
                // 放行非 API 路徑（/ping、儀表板頁面）與瀏覽器的 CORS 預檢請求 (OPTIONS 不帶自訂標頭)
                if (!context.Request.Path.StartsWithSegments("/api") || context.Request.Method == "OPTIONS") {
                    await next();
                    return;
                }
                if (!context.Request.Headers.TryGetValue("X-API-Key", out var key) || key != apiKey) {
                    context.Response.StatusCode = 401;
                    context.Response.ContentType = "text/plain; charset=utf-8";
                    await context.Response.WriteAsync("❌ 401 Unauthorized：沒有帶 X-API-Key，或金鑰不對。"
                        + "金鑰由服務產生、存在這台電腦的使用者資料夾：tools\\PBI-Bridge.ps1 會自己讀，儀表板請按 F5 重新整理。");
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 未授權存取被攔截！");
                    return;
                }
                await next();
            });
            app.MapGet("/ping", () => "pong");

            // ── 網頁儀表板：由本服務提供，並把金鑰帶進頁面 ─────────────────────────
            // 以前用 file:// 直接開 HTML，只能叫使用者手動貼金鑰。改由服務提供後：
            //   · 頁面與 API 同源，使用者不需要知道金鑰的存在
            //   · 其他網站讀不到這個頁面：CORS 白名單不含它們，也不再允許 "null" 來源
            //   · DNS rebinding 由上方的 Host 檢查擋下
            // 同一個 Windows 帳號下的本機程式本來就能直接讀金鑰檔或連 msmdsrv，金鑰從來擋不住它們，
            // 所以把金鑰放進「只有本機拿得到」的頁面，不會降低安全性。
            // 每次請求都重新讀檔：改 HTML 不必重新編譯。
            IResult ServePage(HttpContext ctx, string fileName, bool injectKey) {
                var path = Path.Combine(projectRoot, fileName);
                if (!File.Exists(path)) return Results.NotFound($"找不到 {fileName}（應位於 {projectRoot}）");
                var html = File.ReadAllText(path, System.Text.Encoding.UTF8);
                if (injectKey) {
                    var script = $"<script>window.PBI_API_KEY = {System.Text.Json.JsonSerializer.Serialize(apiKey)};</script>";
                    html = html.Replace("</head>", script + "</head>");
                    ctx.Response.Headers.CacheControl = "no-store";
                }
                return Results.Content(html, "text/html; charset=utf-8");
            }
            app.MapGet("/", (HttpContext ctx) => ServePage(ctx, "PowerBI_Visualizer.html", injectKey: true));
            app.MapGet("/PowerBI_Visualizer.html", (HttpContext ctx) => ServePage(ctx, "PowerBI_Visualizer.html", injectKey: true));
            app.MapGet("/API_Documentation.html", (HttpContext ctx) => ServePage(ctx, "API_Documentation.html", injectKey: false));

            // =====================================================================
            // 讀取
            // =====================================================================

            // 列出所有執行中的 Power BI 實例 —— 切換檔案時的第一站
            app.MapGet("/api/instances", () => {
                try {
                    var list = DiscoverInstances();
                    return Results.Ok(new {
                        Count     = list.Count,
                        Instances = list,
                        Note = list.Count > 1
                            ? "有多個實例，請用 X-PBI-Target 標頭指定目標（Port 或檔名片段），否則所有請求（包含讀取）都會被拒絕。"
                            : list.Count == 1 ? "只有一個實例，不需指定目標。"
                            : "沒有偵測到執行中的 Power BI Desktop。"
                    });
                } catch (Exception ex) {
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            app.MapGet("/api/schema", (HttpContext ctx) =>
                RunModel(ctx, "掃描模型結構", (_, inst) => ExtractSchema(inst), save: false));

            app.MapGet("/api/relationships", (HttpContext ctx) => RunModel(ctx, "掃描資料表關聯結構", (model, _) => {
                var relationships = model.Relationships.OfType<SingleColumnRelationship>()
                    .Select(r => new {
                        FromTable  = r.FromTable.Name,
                        FromColumn = r.FromColumn.Name,
                        ToTable    = r.ToTable.Name,
                        ToColumn   = r.ToColumn.Name,
                        IsActive   = r.IsActive,
                        Cardinality = r.FromCardinality.ToString() + " -> " + r.ToCardinality.ToString(),
                        CrossFilterDirection = r.CrossFilteringBehavior.ToString()
                    }).ToList();
                return new { TotalRelationships = relationships.Count, Relationships = relationships };
            }, save: false));

            // 角色 / 共用運算式的清單
            app.MapGet("/api/roles", (HttpContext ctx) => RunModel(ctx, "列出 RLS 角色", (model, _) => new {
                Roles = model.Roles.Select(r => new {
                    r.Name,
                    ModelPermission  = r.ModelPermission.ToString(),
                    TablePermissions = r.TablePermissions.Select(tp => new { Table = tp.Table.Name, tp.FilterExpression }).ToList()
                }).ToList()
            }, save: false));

            app.MapGet("/api/expressions", (HttpContext ctx) => RunModel(ctx, "列出 Power Query 共用運算式", (model, _) => new {
                Expressions = model.Expressions.Select(e => new { e.Name, e.Kind, e.Expression, e.Description }).ToList()
            }, save: false));

            // 目前解析到的目標實例（診斷用：確認自己正在改哪個檔）
            app.MapGet("/api/pbi-info", (HttpContext ctx) => {
                try {
                    var inst = ResolveInstance(ctx);
                    return Results.Ok(new {
                        inst.Port, inst.PbiPid, inst.MsmdsrvPid,
                        inst.FilePath, inst.FileName, inst.WindowTitle, inst.Kind,
                        FileExists = inst.FilePath != null && File.Exists(inst.FilePath)
                    });
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            // 模型健檢
            app.MapGet("/api/validate", (HttpContext ctx) => RunModel(ctx, "模型健檢", (model, _) => ValidateModel(model), save: false));

            // 報表健檢（僅 PBIP）：直接改過報表檔案之後，確認沒有改壞
            app.MapGet("/api/validate-report", (HttpContext ctx) =>
                RunModel(ctx, "報表健檢", (model, inst) => ValidateReport(model, inst), save: false));

            // =====================================================================
            // 資料保護：逐欄設定
            // =====================================================================

            static bool IsSystemTable(Table t) =>
                t.Name.StartsWith("LocalDateTable_") || t.Name.StartsWith("DateTableTemplate_");

            const int MaxLoosenPerConfirm = 12;       // 確認視窗列得完的數量：使用者同意的必須是他看得到的

            // 每個欄位目前的等級、是誰決定的、有沒有建議。AI 也可以讀 —— 等級是結構資訊，
            // 知道哪些欄位受限，它才寫得出一次就過的查詢。
            app.MapGet("/api/protection", (HttpContext ctx) => RunModel(ctx, "讀取資料保護設定", (model, inst) => {
                var catalog = DataGuard.Derivations.FromModel(model);
                var stale   = new List<ProtectionStore.Entry>();
                var match   = ProtectionStore.Find(inst, catalog);
                var problem = match.Problem;
                var policy  = new DataGuard.Policy(
                    problem == null ? ProtectionStore.LevelsOf(match.Entries(), catalog, stale, match.TagsOnly)
                                    : new Dictionary<string, DataGuard.Level>(), catalog);

                var counts = new Dictionary<DataGuard.Level, int> {
                    [DataGuard.Level.Open] = 0, [DataGuard.Level.Pseudonym] = 0,
                    [DataGuard.Level.CountOnly] = 0, [DataGuard.Level.AggregateOnly] = 0 };
                int suggested = 0;
                var tables = new List<object>();
                foreach (var t in model.Tables.Where(t => !IsSystemTable(t)).OrderBy(t => t.Name)) {
                    var cols = new List<object>();
                    foreach (var c in t.Columns.Where(c => c.Type != ColumnType.RowNumber).OrderBy(c => c.Name)) {
                        bool numeric = catalog.Catalog[t.Name][c.Name];
                        var level = policy.OfColumn(t.Name, c.Name, out var source, out var rule);
                        counts[level]++;
                        string? why = null;
                        // 使用者親自決定過的欄位不再提示 —— 他把建議略過（明確設成開放）就是答案，
                        // 「有建議 0」才會等於「我都看過了」
                        DataGuard.Level? hint = level == DataGuard.Level.Open && source != "explicit"
                            ? DataGuard.Suggest(c.Name, numeric, out why) : null;
                        if (hint != null) suggested++;
                        cols.Add(new {
                            Name = c.Name, DataType = c.DataType.ToString(), Numeric = numeric, IsHidden = c.IsHidden,
                            Level = DataGuard.LevelKey(level), Source = source, Rule = rule,
                            Suggested = hint == null ? null : DataGuard.LevelKey(hint.Value), SuggestedWhy = why
                        });
                    }
                    tables.Add(new { Name = t.Name, IsHidden = t.IsHidden, Columns = cols });
                }
                return new {
                    Enabled    = DataGuard.Enabled,
                    File       = inst.FileName,
                    PathKnown  = inst.FilePath != null,           // false＝服務看不到檔案路徑（從 Power BI 裡面開啟的）
                    CanSave    = match.Key.Length > 0,
                    // 這個模型有沒有逐欄設定。看的是「有沒有任何一筆設定」，不是「設定檔在不在」——
                    // 重設之後檔案還在，但裡面是空的，那和沒設定過是同一回事
                    Configured = problem == null && match.File != null && match.File.Entries.Count > 0,
                    SavedAt    = problem == null ? match.File?.SavedAt : null,
                    // 設定是靠內容認出來的（檔案換過位置、改過名，或這次看不到路徑）：告訴使用者它原本是哪個檔案的
                    InheritedFrom = problem == null && match.File != null && !match.Exact ? match.File.File : null,
                    ModelId    = match.Key,                        // 儀表板用它記「這份模型我看過了」之類只屬於這個瀏覽器的狀態
                    MaxLoosenPerConfirm,
                    ConfirmTimeoutSeconds = HumanConfirm.TimeoutSeconds,
                    Problem    = problem,
                    Summary = new {
                        Open = counts[DataGuard.Level.Open], Pseudonym = counts[DataGuard.Level.Pseudonym],
                        CountOnly = counts[DataGuard.Level.CountOnly], AggregateOnly = counts[DataGuard.Level.AggregateOnly],
                        Suggested = suggested
                    },
                    Limits = new { DataGuard.MaxDetailRows, DataGuard.MaxAggregateRows, DataGuard.MaxRowsWithMoney },
                    // 對不到欄位的設定：欄位被刪了，或在沒有固定代號的情況下改了名
                    Stale  = stale.Select(e => new { e.Table, e.Column, e.Level }).ToList(),
                    Tables = tables
                };
            }, save: false));

            // 變更逐欄設定。收緊立刻生效；只要有任何一欄變寬鬆，就要使用者在確認視窗按「是」，否則整批不套用。
            // （儀表板因此把一次儲存拆成兩次送：先送收緊，再送放寬 —— 錯過視窗不會連收緊的也沒存到。）
            app.MapPost("/api/protection", async (ProtectionRequest req, HttpContext ctx) => {
                try {
                    var inst = ResolveInstance(ctx);
                    bool reset  = req.Reset == true;
                    var changes = req.Changes ?? Array.Empty<ProtectionChange>();
                    if (!reset && changes.Length == 0) return Results.BadRequest("❌ Changes 不可為空");
                    if (changes.Length > 5000) return Results.BadRequest("❌ 一次最多 5000 筆變更");

                    Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] 🛡️ 變更資料保護設定：{inst.FileName}（{(reset ? "重設" : changes.Length + " 筆")}）");

                    DataGuard.Derivations catalog;
                    using (var server = new Server()) {
                        server.Connect($"Data Source=localhost:{inst.Port};");
                        catalog = DataGuard.Derivations.FromModel(server.Databases[0].Model);
                    }
                    var tagOf = new Dictionary<string, string>();          // 表＋欄 → 固定代號
                    foreach (var kv in catalog.TagToColumn) tagOf[DataGuard.Policy.Key(kv.Value.Table, kv.Value.Column)] = kv.Key;

                    var match = ProtectionStore.Find(inst, catalog);
                    if (match.Key.Length == 0) return Results.BadRequest(NoIdentityMessage);
                    string? problem = match.Problem;
                    if (problem != null && !reset)
                        return Results.Json(new { Error = "⛔ 設定檔讀不出來，要先重設", Reason = problem }, statusCode: 409);

                    var oldEntries = problem == null ? match.Entries() : new List<ProtectionStore.Entry>();
                    var before = new DataGuard.Policy(ProtectionStore.LevelsOf(oldEntries, catalog, null, match.TagsOnly), catalog);

                    // 先把舊設定對到欄位現在的名字（改過名的在這裡跟上），對不上的原樣留著
                    var work = new Dictionary<string, ProtectionStore.Entry>();
                    var orphans = new List<ProtectionStore.Entry>();
                    if (!reset) foreach (var e in oldEntries) {
                        var now = ProtectionStore.Locate(e, catalog, match.TagsOnly);
                        if (now == null) { orphans.Add(e); continue; }
                        work[DataGuard.Policy.Key(now.Value.Table, now.Value.Column)] =
                            new ProtectionStore.Entry { Table = now.Value.Table, Column = now.Value.Column, Level = e.Level, Tag = e.Tag };
                    }

                    bool onlyStale = !reset;        // 這次只是在清「對不到欄位的殘留設定」
                    foreach (var ch in changes) {
                        if (string.IsNullOrWhiteSpace(ch.Table) || string.IsNullOrWhiteSpace(ch.Column))
                            return Results.BadRequest("❌ 每一筆變更都要有 Table 與 Column");
                        bool toDefault = (ch.Level ?? "").Trim().Equals("default", StringComparison.OrdinalIgnoreCase);
                        if (!toDefault && !DataGuard.TryParseLevel(ch.Level, out _))
                            return Results.BadRequest($"❌ 不認得的等級「{ch.Level}」（可用 open / pseudonym / countOnly / aggregateOnly / default）");

                        string key = DataGuard.Policy.Key(ch.Table, ch.Column);
                        bool exists = catalog.Catalog.TryGetValue(ch.Table, out var tcols) && tcols.ContainsKey(ch.Column);
                        if (!exists) {
                            // 欄位已經不在模型裡：只允許把殘留的設定清掉
                            if (!toDefault) return Results.BadRequest($"❌ 模型裡沒有 '{ch.Table}'[{ch.Column}]");
                            orphans.RemoveAll(o => DataGuard.Policy.Key(o.Table, o.Column) == key);
                            continue;
                        }
                        onlyStale = false;
                        if (toDefault) { work.Remove(key); continue; }

                        DataGuard.TryParseLevel(ch.Level, out var level);
                        if (level == DataGuard.Level.AggregateOnly && !tcols![ch.Column])
                            return Results.BadRequest($"❌ '{ch.Table}'[{ch.Column}] 不是數值欄位，不能設成「只能彙總」—— "
                                                    + "文字欄位的最大值／最小值就是一個真實的值。請改用「只能計數」或「換成代號」。");
                        // 用模型裡的正式名稱存（請求裡的大小寫可能不同）
                        string realTable  = catalog.Catalog.Keys.First(k => k.Equals(ch.Table, StringComparison.OrdinalIgnoreCase));
                        string realColumn = tcols!.Keys.First(k => k.Equals(ch.Column, StringComparison.OrdinalIgnoreCase));

                        // 本來就開放、也不是在蓋掉通用規則的欄位，「設成開放」不必記。
                        // 記了的話，這一欄會從「建議設成受限」的清單上消失、顯示成使用者親自決定過 ——
                        // 而這個端點分不出送請求的是使用者還是 AI。
                        if (level == DataGuard.Level.Open && !work.ContainsKey(key)
                            && before.OfColumn(realTable, realColumn) == DataGuard.Level.Open) continue;

                        work[key] = new ProtectionStore.Entry {
                            Table = realTable, Column = realColumn, Level = DataGuard.LevelKey(level),
                            Tag = tagOf.TryGetValue(key, out var tag) ? tag : null
                        };
                    }

                    var newEntries = work.Values.Concat(orphans).ToList();
                    if (!reset && problem == null && ProtectionStore.Same(oldEntries, newEntries))
                        return Results.Ok(new {
                            message = "沒有需要變更的項目", Tightened = 0, Loosened = 0, ConfirmedByUser = false,
                            SavedAt = match.File?.SavedAt
                        });
                    var after = new DataGuard.Policy(ProtectionStore.LevelsOf(newEntries, catalog, null, match.TagsOnly), catalog);

                    // 比的是「每一個欄位最後的等級」，不是請求裡寫了什麼 ——
                    // 清掉一筆殘留設定、或把某欄改回通用規則，都可能讓別的欄位跟著變寬鬆。
                    var loosened = new List<string>();
                    int tightened = 0;
                    foreach (var t in catalog.Catalog)
                        foreach (var c in t.Value.Keys) {
                            var was = before.OfColumn(t.Key, c);
                            var now = after.OfColumn(t.Key, c);
                            if (was == now) continue;
                            if (problem != null || DataGuard.IsLoosening(was, now))
                                loosened.Add(OneLine($"'{t.Key}'[{c}]：{DataGuard.LevelName(was)} → {DataGuard.LevelName(now)}"));
                            else tightened++;
                        }
                    loosened.Sort(StringComparer.Ordinal);      // 儀表板的等待畫面用同一個順序列，使用者才能逐行對照

                    bool needConfirm = loosened.Count > 0 || problem != null;
                    if (needConfirm) {
                        // 重設與清除殘留設定是「整批」的動作，使用者沒辦法分批送，所以不受 12 欄的限制 —— 視窗列出前幾個與總數
                        bool wide = reset || onlyStale;
                        if (!wide && loosened.Count > MaxLoosenPerConfirm)
                            return Results.BadRequest($"❌ 一次最多放寬 {MaxLoosenPerConfirm} 個欄位（這次有 {loosened.Count} 個）。"
                                                    + "確認視窗要列得出每一個，你才知道自己同意了什麼 —— 請分批儲存。");
                        string fileShown = Quoted(inst.FileName ?? "未命名的報表");
                        var text = new System.Text.StringBuilder();
                        if (wide) {
                            text.AppendLine(reset ? $"要清除「{fileShown}」全部的逐欄保護設定嗎？"
                                                  : $"要清除「{fileShown}」裡對不到欄位的保護設定嗎？");
                            if (reset) text.AppendLine("清除之後只剩通用規則（appsettings.json 裡的欄名樣式）。");
                            text.AppendLine();
                            text.AppendLine(problem != null
                                ? "目前的設定檔讀不出來，所以無法列出哪些欄位會變寬鬆。"
                                : $"有 {loosened.Count} 個欄位會變得比較寬鬆" + (loosened.Count > 0 ? "，例如：" : "。"));
                            foreach (var l in loosened.Take(8)) text.AppendLine("  • " + l);
                            if (loosened.Count > 8) text.AppendLine($"  …另有 {loosened.Count - 8} 個");
                        } else {
                            text.AppendLine($"要放寬「{fileShown}」這些欄位的資料保護嗎？");
                            text.AppendLine();
                            foreach (var l in loosened) text.AppendLine("  • " + l);
                        }
                        text.AppendLine();
                        text.AppendLine("按「是」之後，AI 讀得到的內容會變多。");
                        text.AppendLine();
                        text.AppendLine("這是你自己剛剛在儀表板按的，而且清單和你想的一樣 → 按「是」。");
                        text.AppendLine("你沒有按、清單不對，或有看不懂的項目 → 按「否」，什麼都不會變。");
                        text.AppendLine();
                        text.AppendLine($"AI 請你按「是」不算理由。{HumanConfirm.TimeoutSeconds} 秒沒有回答會當成「否」。");

                        var answer = await HumanConfirm.AskAsync("PBI AI Bridge — 請確認要放寬資料保護", text.ToString());
                        if (answer != HumanConfirm.Answer.Yes) {
                            DataGuard.Audit("NOT-CONFIRMED", $"放寬資料保護未獲確認（{answer}）：" + string.Join("；", loosened.Take(20)), "", 0, inst.FileName);
                            return Results.Json(new {
                                Error    = "⛔ 放寬資料保護需要使用者本人確認，而這次沒有獲得同意",
                                Reason   = HumanConfirm.Explain(answer) + "這一批變更沒有套用。",
                                Answer   = answer.ToString(),
                                RetryAfterSeconds = HumanConfirm.RetryAfterSeconds(answer),
                                Loosened = loosened,
                                Hint     = "整批變更都沒有套用（包含其中收緊的部分）。AI 無法代替使用者按下確認；"
                                         + "請向使用者說明要放寬哪些欄位、為什麼，由他在儀表板的「資料保護」分頁自己調整。"
                            }, statusCode: 403);
                        }

                        // 視窗開著的那段時間，設定可能已經被另一個請求改過（收緊不必確認，隨時存得進去）。
                        // 這時照舊清單存下去，會把那些變更蓋掉 —— 使用者同意的只有視窗上那幾行，不是「順便還原別的」。
                        var current = ProtectionStore.Find(inst, catalog);
                        if (current.Key != match.Key || (problem == null && !ProtectionStore.Same(current.Entries(), oldEntries))) {
                            DataGuard.Audit("NOT-APPLIED", "等待確認的期間設定被改過，這一批沒有套用", "", 0, inst.FileName);
                            return Results.Json(new {
                                Error  = "⛔ 設定在等待確認的期間被改過，這一批沒有套用",
                                Reason = "你回答確認視窗之前，這份模型的保護設定已經有別的變更存進去了。為了不把那些變更蓋掉，這一批沒有套用。"
                                       + "請重新整理後再試一次。"
                            }, statusCode: 409);
                        }
                    }

                    var saved = ProtectionStore.Save(inst, catalog, match, newEntries);
                    DataGuard.Audit("PROTECTION", $"逐欄設定已變更：收緊 {tightened}、放寬 {loosened.Count}"
                                  + (loosened.Count > 0 ? "（使用者已確認）：" + string.Join("；", loosened.Take(20)) : ""), "", 0, inst.FileName);
                    Console.WriteLine($"✅ 資料保護設定已儲存：收緊 {tightened} 欄、放寬 {loosened.Count} 欄" + (needConfirm ? "（使用者已確認）" : ""));
                    return Results.Ok(new {
                        message = reset ? "已清除這個模型的逐欄設定" : $"已儲存：收緊 {tightened} 欄、放寬 {loosened.Count} 欄",
                        Tightened = tightened, Loosened = loosened.Count, ConfirmedByUser = needConfirm,
                        SavedAt = saved.SavedAt
                    });
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    Console.WriteLine($"❌ 變更資料保護設定失敗: {ex.Message}");
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            // =====================================================================
            // 查詢
            // =====================================================================

            // 使用者同意放行單一句查詢時的額度。列數與每格的長度都要有上限 ——
            // 不然 CONCATENATEX 把整欄串成一格，「最多幾列」就成了空話。
            const int MaxApprovedRows      = 1000;
            const int MaxApprovedCellChars = 200;

            // 共用的 ADOMD 執行邏輯（/api/query 與 /api/dmv 都用它）
            //   approvedRows         使用者剛在確認視窗同意放行這一句：不擋、不換代號，最多回傳這麼多列
            //   approvedFingerprint  他同意的那一刻，查詢與它用到的定義的指紋；執行前對不上就不算數
            //   outcome              被擋下時把原因記在這裡（確認視窗要用）
            IResult RunAdomd(HttpContext ctx, string query, int? maxRowsIn, int? timeoutIn, string label,
                             bool isDax = false, int? approvedRows = null, string? approvedFingerprint = null,
                             GuardOutcome? outcome = null) {
                // 查詢碰到受限欄位時，引擎的錯誤訊息要先遮掉引用的資料內容才能回傳（見 DataGuard.RedactError）
                bool redactErrors = false;
                DataGuard.Derivations? errModel = null;
                try {
                    var inst = ResolveInstance(ctx);
                    int maxRows = approvedRows ?? Math.Clamp(maxRowsIn ?? 1000, 1, 10000);
                    int timeout = Math.Clamp(timeoutIn ?? 60, 1, 600);
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] 🔎 {label} → {inst.FileName}（上限 {maxRows} 列 / {timeout} 秒）...");
                    var sw = Stopwatch.StartNew();

                    // 攔截時的共用出口：記稽核，並告訴呼叫端怎麼請使用者決定。
                    var mentioned = new List<string>();      // 這句查詢提到的受限欄位（掃描完才有）
                    string? fingerprint = null;              // 查詢＋展開的定義（展開完才有）
                    IResult Blocked(string verdict, string reason, string hint, int rows) {
                        DataGuard.Audit(verdict, reason, query, rows, inst.FileName);
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⛔ 資料保護攔截：{reason}");
                        // 只有「某個欄位的等級不允許這樣取值」可以請使用者單次放行。
                        // 沒辦法檢查（設定檔壞了、定義讀不到或太多）、寫法本身被禁用、定義中途被換掉 —— 這些按一下「是」也不放行：
                        // 「我們看不到這句查詢會帶出什麼」不能變成「那就不檢查」。
                        bool canAsk = fingerprint != null && verdict is "BLOCKED-A" or "BLOCKED-B" or "BLOCKED-ROWS";
                        if (outcome != null) {
                            outcome.Verdict = verdict; outcome.Reason = reason;
                            outcome.File = inst.FileName; outcome.Mentioned = mentioned;
                            outcome.Fingerprint = fingerprint; outcome.CanAskUser = canAsk;
                        }
                        return Results.Json(new {
                            Error = "⛔ 已被資料保護機制攔截（伺服器端）", Reason = reason, Hint = hint,
                            Verdict = verdict, CanAskUser = canAsk,
                            HowToOverride = canAsk
                                ? "先想這個值是不是真的需要 —— 多半換成彙總寫法就夠了。真的需要時，先向使用者說明要看哪些欄位、為什麼，"
                                + "再以 Invoke-Dax -AskUser 重送同一句：使用者的螢幕會跳出確認視窗，由他決定要不要放行這一次。"
                                : verdict == "BLOCKED-SETTINGS"
                                ? "這個狀況不能用單次同意放行，要先請使用者到儀表板的「資料保護」分頁處理。"
                                : "這個狀況不能用單次同意放行（-AskUser 也一樣）。請換一種寫法。"
                        }, statusCode: 403);
                    }

                    // 從讀模型定義到讀完結果，中間不讓任何寫入插進來（見 ModelGate）
                    using var gate = EnterGate(write: false);
                    using (var conn = new AdomdConnection($"Data Source=localhost:{inst.Port};")) {
                        conn.Open();

                        // ── 第一關：執行前掃描查詢文字 ──────────────────────────────
                        // 放在 conn.Open() 之後，是因為展開量值定義、讀欄位目錄都需要連線。
                        // 敏感欄位可以藏在量值裡，不展開就等於沒檢查。
                        bool touchesMoney = false;
                        DataGuard.Policy?      policy  = null;
                        DataGuard.Derivations? derived = null;
                        // DMV 不走這兩關：它只能查白名單上的系統檢視（見 /api/dmv），那些檢視的欄名是固定的
                        // 中繼資料欄位（Name、Description、Expression…），拿模型的欄名樣式去比對沒有意義，只會誤擋。
                        bool enforce = DataGuard.Enabled && isDax;            // 這句查詢歸資料保護管
                        bool guarded = enforce && approvedRows == null;       // 沒有使用者的單次同意：照規則擋
                        if (enforce) {
                            derived = DataGuard.LoadDerivations(conn);
                            if (derived.LoadFailed)
                                return Blocked("BLOCKED-LOAD",
                                    $"讀不到模型的定義（$SYSTEM.{derived.LoadProblem}），沒辦法檢查這句查詢用到的量值與計算欄位裡有沒有受限欄位。"
                                  + "沒看到的東西不能當成沒問題。",
                                    "通常再送一次就好。持續發生的話，請把這段訊息回報給使用者。", 0);
                            int measureCount = derived.Count;
                            // 這份模型的逐欄設定。要先讀到模型的欄位目錄：服務看不到檔案路徑時，是靠欄位的固定代號認設定檔的
                            var settings = ProtectionStore.Find(inst, derived);
                            if (settings.Problem != null)
                                return Blocked("BLOCKED-SETTINGS", settings.Problem,
                                    "請使用者到儀表板（http://localhost:5500/）的「資料保護」分頁處理。", 0);
                            policy = new DataGuard.Policy(
                                ProtectionStore.LevelsOf(settings.Entries(), derived, null, settings.TagsOnly), derived);
                            var expanded = DataGuard.ExpandDerivations(query, derived);
                            if (expanded.Incomplete)
                                return Blocked("BLOCKED-INCOMPLETE", "這句查詢牽涉到的量值／計算資料行定義多到無法完整檢查，沒看完的東西不能當成沒問題。",
                                    "請把查詢拆小，或直接引用較底層的量值。", 0);
                            fingerprint = DataGuard.Fingerprint(expanded);
                            var scan = DataGuard.ScanDax(expanded.Text, expanded.DefinitionsFrom, expanded.Boundaries);
                            mentioned.AddRange(DataGuard.RestrictedMentions(scan, policy, derived));
                            redactErrors = mentioned.Count > 0;
                            errModel     = derived;

                            if (!guarded) {
                                // 使用者同意的是確認視窗列出來的那些內容。視窗開著的期間，量值或計算欄位的定義被換掉
                                // （寫入量值不需要確認），現在要執行的就不是他看到的那件事了。
                                if (fingerprint != approvedFingerprint)
                                    return Blocked("BLOCKED-CHANGED",
                                        "這句查詢用到的量值或計算欄位，在使用者回答確認視窗的期間被改過。剛才的同意是針對改之前的內容，不適用於現在的。",
                                        "不要在等待確認的期間修改模型。向使用者說明之後重新詢問一次。", 0);
                            } else {
                                var violation = DataGuard.CheckQuery(scan, policy, derived, out touchesMoney, out bool hard);
                                if (violation != null)
                                    return Blocked(hard ? "BLOCKED-FUNC" : "BLOCKED-A", $"{violation}（已展開 {measureCount} 筆量值定義）",
                                        "欄位名稱可以自由讀取，被擋的是欄位內容。改用彙總寫法即可，"
                                      + "例如 EVALUATE ROW(\"客戶數\", DISTINCTCOUNT(<表>[<欄>]))。"
                                      + "每個欄位目前的等級可以用 Get-PbiProtection 查。", 0);

                                // ── 列數上限：由伺服器決定，不由呼叫端決定 ─────────────────
                                // MaxRows 原本是請求參數 —— 那不是護欄，只是預設值：
                                // 呼叫端（也就是 AI）可以自己填 10000。這裡改成只能往下收。
                                //   逐列明細 → MaxDetailRows
                                //   彙總結果 → MaxAggregateRows（統計量不是資料細節）
                                bool isAgg = DataGuard.IsAggregateQuery(expanded.Text);
                                int  cap   = isAgg ? DataGuard.MaxAggregateRows : DataGuard.MaxDetailRows;
                                if (maxRows > cap) {
                                    Console.WriteLine($"   ↓ 列數上限收緊為 {cap} 列（{(isAgg ? "彙總查詢" : "逐列明細")}）");
                                    maxRows = cap;
                                }
                            }
                        }

                        using (var cmd = conn.CreateCommand()) {
                            cmd.CommandText = query;
                            cmd.CommandTimeout = timeout;
                            using (var reader = cmd.ExecuteReader()) {
                                // 欄位名稱去重，避免 DAX 回傳同名欄位時後者覆蓋前者
                                // rawNames 是引擎給的原名，資料保護用它判斷每一欄的等級；
                                // colNames 是去重之後的，只拿來當輸出的鍵（加了 _1 的名字對不回模型的欄位）
                                var colNames = new List<string>();
                                var rawNames = new List<string>();
                                var seen = new Dictionary<string, int>();
                                for (int i = 0; i < reader.FieldCount; i++) {
                                    string n = reader.GetName(i);
                                    if (string.IsNullOrEmpty(n)) n = $"Column{i + 1}";
                                    rawNames.Add(n);
                                    if (seen.TryGetValue(n, out int c)) { seen[n] = c + 1; n = $"{n}_{c + 1}"; }
                                    else seen[n] = 0;
                                    colNames.Add(n);
                                }

                                // ── 第二關：執行後、讀取任何一列之前檢查結果欄位 ──────────────
                                // 補第一關的盲點：整表 EVALUATE 的查詢文字裡沒有任何欄名，
                                // 但引擎回報的結果欄名是 Table[Column]，在這裡才抓得到。
                                // 注意順序 —— 這道檢查在 reader.Read() 之前，值連讀都沒讀進來。
                                // maskIdx：設為「換成代號」的欄位。它們的值要換成代號；對照表只印在主控台，不回傳。
                                var maskIdx = new HashSet<int>();
                                if (guarded) {
                                    var colViolation = DataGuard.CheckResultColumns(rawNames, policy!, derived, out maskIdx);
                                    if (colViolation != null)
                                        return Blocked("BLOCKED-B", colViolation,
                                            "請改以彙總方式取得所需資訊，不要整表取出。", 0);
                                }
                                var pseudoMap = new Dictionary<string, string>();

                                var rows = new List<Dictionary<string, object?>>();
                                bool truncated = false;
                                while (reader.Read()) {
                                    if (rows.Count >= maxRows) { truncated = true; break; }
                                    var row = new Dictionary<string, object?>();
                                    for (int i = 0; i < reader.FieldCount; i++) {
                                        var v = reader.GetValue(i);
                                        object? val = v == DBNull.Value ? null : v;
                                        if (maskIdx.Contains(i) && val != null) {
                                            string real = val.ToString() ?? "";
                                            string code = DataGuard.Pseudonym(real)!;
                                            if (real.Length > 0) pseudoMap[code] = real;
                                            val = code;
                                        } else if (approvedRows != null && val is string cell && cell.Length > MaxApprovedCellChars) {
                                            val = cell.Substring(0, MaxApprovedCellChars) + "…（已截斷）";
                                        }
                                        row[colNames[i]] = val;
                                    }
                                    rows.Add(row);
                                }
                                sw.Stop();

                                // ── 金額查詢的列數上限 ────────────────────────────────
                                // 後備防線：金額有被聚合（過了第一關），但分組鍵的基數太高。
                                // 例如按料號分組 → 幾千列的單一料號營收，
                                // 那已經接近逐筆金額，不是 KPI 彙總了。
                                if (guarded && touchesMoney && rows.Count > DataGuard.MaxRowsWithMoney)
                                    return Blocked("BLOCKED-ROWS",
                                        $"金額查詢回傳 {rows.Count} 列，超過上限 {DataGuard.MaxRowsWithMoney} 列。",
                                        "請改用基數較低的分組鍵（產品線、月份、廠別），或先篩選再彙總。", rows.Count);

                                DataGuard.Audit("OK", touchesMoney ? "含金額欄位" : "", query, rows.Count, inst.FileName);

                                Console.WriteLine($"✅ 查詢完成：{rows.Count} 列 / {sw.ElapsedMilliseconds} ms" +
                                                  (truncated ? $"（已截斷至 {maxRows} 列）" : ""));
                                if (maskIdx.Count > 0) {
                                    Console.WriteLine($"   🎭 {maskIdx.Count} 個「換成代號」欄位／{pseudoMap.Count} 個相異值已換成代號"
                                                    + "　對照表只印在這裡，不會回傳給 AI：");
                                    int shown = 0;
                                    foreach (var kv in pseudoMap) {
                                        if (shown++ >= 50) {
                                            Console.WriteLine($"      …另有 {pseudoMap.Count - 50} 筆未列出");
                                            break;
                                        }
                                        Console.WriteLine($"      {kv.Key}  =  {kv.Value}");
                                    }
                                }

                                return Results.Ok(new {
                                    Columns   = colNames,
                                    RowCount  = rows.Count,
                                    Truncated = truncated,
                                    ElapsedMs = sw.ElapsedMilliseconds,
                                    // 讓 AI 知道哪些欄位是代號，不要拿去當真名解讀
                                    Pseudonymized = maskIdx.Select(i => colNames[i]).ToList(),
                                    // 使用者剛在確認視窗同意放行：這份結果沒有經過遮蔽，裡面是真實內容
                                    ApprovedByUser = approvedRows != null,
                                    Rows      = rows
                                });
                            }
                        }
                    }
                } catch (AdomdErrorResponseException ex) {
                    // DAX 語法或執行期錯誤：回 400 並附上引擎的訊息，方便直接修正公式
                    Console.WriteLine($"❌ 查詢錯誤: {ex.Message}");
                    string shown = redactErrors ? DataGuard.RedactError(ex.Message, query, errModel) : ex.Message;
                    return Results.BadRequest($"❌ 查詢錯誤: {shown}"
                        + (shown != ex.Message ? "\n（這句查詢碰到受限欄位，訊息裡引用的資料內容已遮蔽；原文印在服務主控台。）" : ""));
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    Console.WriteLine($"❌ 查詢失敗: {ex.Message}");
                    return Results.Problem(detail: redactErrors ? DataGuard.RedactError(ex.Message, query, errModel) : ex.Message,
                                           statusCode: 500);
                }
            }

            // 拿掉 DAX 的註解（// …、-- …、/* … */）。字串、'引號名稱'、[中括號名稱] 裡面的不動。
            static string StripDaxComments(string s) {
                var sb = new System.Text.StringBuilder(s.Length);
                for (int i = 0; i < s.Length; ) {
                    char c = s[i];
                    if (c == '"' || c == '\'' || c == '[') {
                        char close = c == '[' ? ']' : c;
                        int j = i + 1;
                        while (j < s.Length) {
                            if (s[j] == close) { if (j + 1 < s.Length && s[j + 1] == close) { j += 2; continue; } break; }
                            j++;
                        }
                        j = Math.Min(j + 1, s.Length);
                        sb.Append(s, i, j - i);
                        i = j;
                        continue;
                    }
                    bool next(char x) => i + 1 < s.Length && s[i + 1] == x;
                    if (c == '/' && next('*')) { int e = s.IndexOf("*/", i + 2, StringComparison.Ordinal); i = e < 0 ? s.Length : e + 2; sb.Append(' '); continue; }
                    if ((c == '/' && next('/')) || (c == '-' && next('-'))) { int e = s.IndexOfAny(new[] { '\r', '\n' }, i); i = e < 0 ? s.Length : e; sb.Append(' '); continue; }
                    sb.Append(c);
                    i++;
                }
                return sb.ToString();
            }

            static string Shorten(string s, int max) => s.Length <= max ? s : s.Substring(0, max) + " …（後面還有）";
            static string FirstSentence(string? s, int max) {
                s ??= "";
                int dot = s.IndexOf('。');
                return Shorten(dot > 0 ? s.Substring(0, dot + 1) : s, max);
            }

            // 唯讀 DAX 查詢（驗算量值結果用）
            app.MapPost("/api/query", async (DaxQueryRequest req, HttpContext ctx) => {
                if (string.IsNullOrWhiteSpace(req.Query)) return Results.BadRequest("❌ Query 不可為空");
                // ✅ 唯讀防護：只接受 DAX 查詢語法，擋掉 XMLA 命令等會變更模型的內容
                var head = StripLeadingComments(req.Query);
                if (!head.StartsWith("EVALUATE", StringComparison.OrdinalIgnoreCase) &&
                    !head.StartsWith("DEFINE",   StringComparison.OrdinalIgnoreCase)) {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 非唯讀查詢被擋下");
                    return Results.BadRequest("❌ 只接受唯讀查詢：必須以 EVALUATE 或 DEFINE 開頭");
                }

                var outcome = new GuardOutcome();
                var result  = RunAdomd(ctx, req.Query, req.MaxRows, req.TimeoutSeconds, "執行 DAX 查詢", isDax: true, outcome: outcome);
                // 沒被擋，或呼叫端沒有要求詢問使用者 → 就是這個結果
                if (outcome.Verdict == null || req.AskUser != true) return result;
                // 不是每一種擋法都能請使用者放行（見 RunAdomd 的 Blocked）
                if (!outcome.CanAskUser) return result;

                // ── 請使用者決定要不要放行這一句 ──────────────────────────────
                // 以前是把一組代碼印在主控台，由使用者抄給 AI。代碼可以被轉手：貼出去的那一刻，
                // 同意的就不是他親眼看到的那件事，而且他得先在黑窗裡找到那行字。
                // 確認視窗的內容由伺服器決定（原因、受限欄位、列數、查詢），按鈕只有坐在電腦前的人按得到。
                int rows = Math.Clamp(req.MaxRows ?? DataGuard.MaxDetailRows, 1, MaxApprovedRows);

                // 使用者同意的必須是他看得到的：受限欄位要全部列得出來，查詢要整句放得進視窗。
                // 註解先拿掉 —— 不然開頭塞一段安撫人的說明，真正的查詢就被擠到看不見的地方去了。
                const int MaxDialogMentions = 12, MaxDialogQueryChars = 700;
                string shownQuery = OneLine(StripDaxComments(req.Query), int.MaxValue);
                if (outcome.Mentioned.Count > MaxDialogMentions || shownQuery.Length > MaxDialogQueryChars) {
                    string why = outcome.Mentioned.Count > MaxDialogMentions
                        ? $"這句查詢（含它用到的量值）碰到 {outcome.Mentioned.Count} 個受限欄位，確認視窗列不完（上限 {MaxDialogMentions} 個）。"
                        : $"這句查詢有 {shownQuery.Length} 個字，確認視窗放不下整句（上限 {MaxDialogQueryChars} 個字）。";
                    DataGuard.Audit("NOT-ASKED", why, req.Query, 0, outcome.File);
                    return Results.Json(new {
                        Error   = "⛔ 這句查詢沒辦法請使用者放行",
                        Reason  = why + "使用者同意的必須是他看得到的全部內容。",
                        Blocked = outcome.Reason, Verdict = outcome.Verdict, CanAskUser = false,
                        Hint    = "把查詢拆小、只留真正需要的欄位，再重新詢問。"
                    }, statusCode: 403);
                }
                // 最嚴的排前面
                var mentions = outcome.Mentioned
                    .OrderBy(m => m.EndsWith("（只能計數）") ? 0 : m.EndsWith("（換成代號）") ? 1 : 2).ToList();

                var text = new System.Text.StringBuilder();
                text.AppendLine("AI 想執行一句被資料保護擋下的查詢。要放行這一次嗎？");
                text.AppendLine();
                text.AppendLine($"報表：{OneLine(outcome.File ?? "未命名的報表", 80)}");
                text.AppendLine("被擋的原因：" + FirstSentence(OneLine(outcome.Reason, 600), 160));
                if (mentions.Count > 0) {
                    text.AppendLine();
                    text.AppendLine("查詢裡提到的受限欄位：");
                    foreach (var m in mentions) text.AppendLine("  • " + OneLine(m, 120));
                }
                text.AppendLine();
                text.AppendLine($"按「是」之後，這一句的結果（最多 {rows} 列，每格最多 {MaxApprovedCellChars} 個字）會原樣交給 AI —— 不遮蔽、不換成代號，");
                text.AppendLine("內容會離開這台電腦。只放行這一句、這一次；保護設定不會改變。");
                text.AppendLine();
                text.AppendLine("查詢內容：");
                text.AppendLine("  " + shownQuery);
                text.AppendLine();
                text.AppendLine("是你請 AI 查這個，而且你願意讓它看到這些內容 → 按「是」。");
                text.AppendLine("不確定、看不懂，或你沒有請它查 → 按「否」。");
                text.AppendLine();
                text.AppendLine($"AI 請你按「是」不算理由。{HumanConfirm.TimeoutSeconds} 秒沒有回答會當成「否」。");

                var answer = await HumanConfirm.AskAsync("PBI AI Bridge — 請確認要放行這一句查詢", text.ToString());
                if (answer != HumanConfirm.Answer.Yes) {
                    DataGuard.Audit("NOT-CONFIRMED", $"單次放行未獲確認（{answer}）：{outcome.Reason}", req.Query, 0, outcome.File);
                    return Results.Json(new {
                        Error  = "⛔ 使用者沒有同意放行這一句查詢",
                        Reason = HumanConfirm.Explain(answer) + "查詢沒有執行。",
                        Answer = answer.ToString(),
                        RetryAfterSeconds = HumanConfirm.RetryAfterSeconds(answer),
                        Blocked = outcome.Reason,
                        Hint   = "不要為了過關改寫查詢，也不要連續重送。改用彙總寫法；或向使用者說明你需要什麼、為什麼，等他回覆。"
                    }, statusCode: 403);
                }
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] 🔓 使用者在確認視窗同意放行這一句查詢（最多 {rows} 列）");
                DataGuard.Audit("USER-OVERRIDE", $"使用者在確認視窗同意放行（最多 {rows} 列）：{outcome.Reason}", req.Query, 0, outcome.File);
                return RunAdomd(ctx, req.Query, rows, req.TimeoutSeconds, "執行 DAX 查詢（使用者已同意放行）", isDax: true,
                                approvedRows: rows, approvedFingerprint: outcome.Fingerprint);
            });

            // DMV 查詢：$SYSTEM 系統檢視（中繼資料、VertiPaq 儲存統計、記憶體佔用）。資料保護啟用時是白名單制
            // 這些檢視只含中繼資料與統計，不會回傳事實資料列
            app.MapPost("/api/dmv", (DmvQueryRequest req, HttpContext ctx) => {
                if (string.IsNullOrWhiteSpace(req.Query)) return Results.BadRequest("❌ Query 不可為空");
                var head = StripLeadingComments(req.Query);
                if (!Regex.IsMatch(head, @"^SELECT\b[\s\S]*\bFROM\s+\$SYSTEM\.[A-Za-z0-9_]+",
                                   RegexOptions.IgnoreCase)) {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 非 $SYSTEM DMV 查詢被擋下");
                    return Results.BadRequest("❌ 只接受 $SYSTEM DMV 查詢，格式須為：SELECT ... FROM $SYSTEM.<檢視名稱>");
                }

                // 資料保護啟用時用「白名單」：只開放模型的中繼資料與儲存統計。
                //
                // 以前是黑名單，只擋會回傳 M／運算式原文的三個檢視。但 $SYSTEM 底下還有會回傳「資料內容」的檢視：
                // MDSCHEMA_MEMBERS 會列出每個欄位的每一個相異值；DISCOVER_SESSIONS / DISCOVER_COMMANDS 帶著別的連線
                // （包括 Power BI 自己的視覺）剛執行過的查詢文字，裡面有篩選用的真實值；TMSCHEMA_DATA_SOURCES 是連線字串。
                // 欄位層級的管制對它們無效 —— 那些內容躲在 MEMBER_CAPTION、COMMAND_TEXT 這種無害的欄名底下。
                // 會回傳內容的檢視列不完，所以反過來只列確定安全的。
                if (DataGuard.Enabled) {
                    // 只認一種寫法：SELECT <欄位清單> FROM $SYSTEM.<檢視> [WHERE … | ORDER BY …]，而且整句只有一個來源。
                    // 白名單是照「認得出來的檢視名稱」比對的 —— $SYSTEM.[名稱]、[$SYSTEM].[名稱]、第二個來源、子查詢
                    // 這些寫法只要有一種引擎接受而這裡沒認出來，就等於白名單外的檢視照樣查得到。認不出來的一律不收。
                    bool plain = Regex.IsMatch(head.Trim().TrimEnd(';').TrimEnd(),
                        @"^SELECT\s+(?:DISTINCT\s+)?(?:TOP\s+\d+\s+)?[\w\s,\*\[\]\.]+?\s+FROM\s+\$SYSTEM\.[A-Za-z0-9_]+\s*(?:$|(?:WHERE|ORDER\s+BY)\b[\s\S]*$)",
                        RegexOptions.IgnoreCase)
                        && Regex.Matches(head, @"\$\s*SYSTEM", RegexOptions.IgnoreCase).Count == 1
                        && Regex.Matches(head, @"\bFROM\b", RegexOptions.IgnoreCase).Count == 1
                        && head.IndexOf("$SYSTEM", StringComparison.OrdinalIgnoreCase) >= 0
                        && string.Equals(head.Trim(), req.Query.Trim(), StringComparison.Ordinal);      // 前面不能夾註解
                    if (!plain) {
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⛔ 資料保護攔截：DMV 查詢的寫法認不出來");
                        DataGuard.Audit("BLOCKED-DMV", "寫法不是單一來源的 SELECT … FROM $SYSTEM.<檢視>", req.Query, 0, null);
                        return Results.Json(new {
                            Error  = "⛔ 已被資料保護機制攔截（DMV 查詢）",
                            Reason = "資料保護啟用時，DMV 查詢只接受這一種寫法：SELECT <欄位或 *> FROM $SYSTEM.<檢視名稱> [WHERE … | ORDER BY …]，"
                                   + "一句只能查一個檢視，檢視名稱不加中括號，不能有子查詢或註解。",
                            Allowed = AllowedDmvViews.OrderBy(v => v).ToList()
                        }, statusCode: 403);
                    }
                    var views = Regex.Matches(head, @"\$SYSTEM\.([A-Za-z0-9_]+)", RegexOptions.IgnoreCase)
                                     .Select(m => m.Groups[1].Value.ToUpperInvariant()).Distinct().ToList();
                    var refused = views.Where(v => !AllowedDmvViews.Contains(v)).ToList();
                    if (refused.Count > 0) {
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⛔ 資料保護攔截：DMV 檢視 {string.Join(", ", refused)} 不在允許清單內");
                        DataGuard.Audit("BLOCKED-DMV", "檢視不在允許清單內：" + string.Join(", ", refused), req.Query, 0, null);
                        return Results.Json(new {
                            Error  = "⛔ 已被資料保護機制攔截（DMV 檢視）",
                            Reason = $"$SYSTEM.{refused[0]} 不在允許清單內。資料保護啟用時只開放模型的中繼資料與儲存統計 —— "
                                   + "有些系統檢視會回傳欄位的實際內容、M 腳本或別的連線執行過的查詢文字。",
                            Allowed = AllowedDmvViews.OrderBy(v => v).ToList(),
                            Hint   = "要看模型結構請用 /api/schema、/api/validate；"
                                   + "要看單張表的 M 請用 Get-PbiMQuery（預設只存檔不回傳內容）。"
                        }, statusCode: 403);
                    }
                }
                return RunAdomd(ctx, req.Query, req.MaxRows, req.TimeoutSeconds, "執行 DMV 查詢");
            });

            // =====================================================================
            // 存檔 / 重新整理
            // =====================================================================

            // 送出 Ctrl+S 給 PBI Desktop，並驗證是否真的存到：哪些檔案變了、預期的內容在不在磁碟上。
            // ⚠️ 會搶走前景視窗焦點，這是 SendKeys 的固有限制。
            app.MapPost("/api/save", async (SaveRequest? req, HttpContext ctx) => {
                try {
                    var inst = ResolveInstance(ctx);
                    int  wait       = Math.Clamp(req?.WaitSeconds ?? 30, 1, 180);
                    bool verifyOnly = req?.VerifyOnly == true;
                    var  needles    = (req?.Expect ?? Array.Empty<string>())
                                      .Select(CollapseWhitespace).Where(s => s.Length > 0).Distinct().Take(20).ToArray();
                    Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] 💾 " + (verifyOnly
                        ? $"檢查磁碟上的存檔內容（不送 Ctrl+S）：{inst.FileName}"
                        : $"存檔：{inst.FileName}（最多等 {wait} 秒，寫完就回報）..."));

                    var target = inst.FilePath;
                    if (string.IsNullOrEmpty(target))
                        return Results.BadRequest("❌ 服務不知道這份報表的檔案路徑（先開 Power BI 再從裡面選檔案，或是還沒存過的新報表），"
                            + "沒辦法驗證存檔結果，所以沒有送出 Ctrl+S。請使用者自己在 Power BI 按 Ctrl+S；"
                            + "要讓這個功能可以用，請他關掉 Power BI、改用雙擊檔案的方式開啟。");

                    var  scope     = ResolveSaveScope(target);
                    bool canVerify = scope.Problem == null;
                    bool isPbip    = scope.Folders.Count > 0;
                    var  before    = ScanWriteTimes(scope);
                    var  now       = before;
                    var  changed   = new List<(string Path, DateTime Time, string Kind)>();
                    var  found     = new Dictionary<string, string>();                 // 預期文字 → 找到它的檔案
                    var  checkedAt = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
                    string Rel(string p) { try { return Path.GetRelativePath(scope.BaseDir, p); } catch { return p; } }

                    var sw = Stopwatch.StartNew();
                    if (!verifyOnly) {
                        string result = await SendCtrlS(inst.PbiPid);

                        if (result == "not_running")
                            return Results.BadRequest("❌ 該 PBI Desktop 行程已結束或沒有可用視窗，無法存檔");
                        if (result == "activate_failed")
                            return Results.BadRequest("❌ 無法將 PBI Desktop 帶到前景，Ctrl+S 未送出。請確認視窗未最小化到系統匣。");

                        // 輪詢到「寫完」為止，而不是固定等幾秒：小檔一兩秒就回來，大檔寫得久也不會被誤判成沒存到。
                        // 「寫完」＝ 有檔案變動，而且連續三次輪詢（1.5 秒）沒有再出現新的變動。
                        // 有給 Expect 時還要等到預期內容出現 —— 檔案變了不代表剛寫的東西進去了。
                        int stable = 0;
                        (int Count, DateTime Latest) lastSig = (-1, DateTime.MinValue);
                        while (canVerify && sw.Elapsed < TimeSpan.FromSeconds(wait)) {
                            await Task.Delay(500);
                            now     = ScanWriteTimes(scope);
                            changed = DiffWriteTimes(before, now);
                            if (changed.Count == 0) continue;

                            var sig = (changed.Count, changed.Max(c => c.Time));
                            stable  = sig == lastSig ? stable + 1 : 0;
                            lastSig = sig;
                            if (stable < 3) continue;
                            // 寫完了。預期內容可能在這次沒被改寫的檔案裡（先前就存過），所以找的是整個範圍。
                            if (needles.Length > 0 && isPbip) FindExpected(now.Keys, now, needles, found, checkedAt);
                            if (found.Count >= needles.Length || !isPbip) break;
                        }
                    }
                    sw.Stop();

                    // 逾時或只檢查不存檔：預期內容還沒找齊的話，把整個範圍找一遍
                    if (needles.Length > 0 && isPbip && found.Count < needles.Length)
                        FindExpected(now.Keys, now, needles, found, checkedAt);

                    bool  fileChanged = canVerify && changed.Count > 0;
                    bool? expectFound = needles.Length == 0 || !isPbip ? null : found.Count >= needles.Length;
                    var   missing     = needles.Where(n => !found.ContainsKey(n)).ToList();

                    string msg; bool good;
                    if (!canVerify) {
                        msg  = (verifyOnly ? "無法檢查：" : "Ctrl+S 已送出，但無法驗證是否存檔成功：") + scope.Problem;
                        good = false;
                    } else if (verifyOnly) {
                        msg  = expectFound == true  ? "預期內容已在磁碟上"
                             : expectFound == false ? $"磁碟上找不到 {missing.Count} 段預期內容 —— 還沒存到"
                             : "只回報目前的檔案時間（沒有給 Expect，或目標是 PBIX）";
                        good = expectFound != false;
                    } else if (fileChanged && expectFound == false) {
                        msg  = $"檔案有變動（{changed.Count} 個），但找不到 {missing.Count} 段預期內容 —— "
                             + "存下去的可能是舊狀態：Power BI 還沒同步到剛才的寫入。等幾秒後最多再存一次";
                        good = false;
                    } else if (fileChanged) {
                        msg  = $"存檔成功（{changed.Count} 個檔案有變動" + (expectFound == true ? "，預期內容已在磁碟上）" : "）");
                        good = true;
                    } else if (expectFound == true) {
                        msg  = "檔案沒有變動，但預期內容已經在磁碟上 —— 先前就存過了";
                        good = true;
                    } else {
                        msg  = $"Ctrl+S 已送出，但 {wait} 秒內檔案沒有變動 —— 可能本來就沒有待存變更、按鍵沒送達，或大檔還在寫";
                        good = false;
                    }
                    if (needles.Length > 0 && !isPbip && canVerify)
                        msg += "（PBIX 是二進位檔，無法檢查內容）";
                    Console.WriteLine((good ? "✅ " : "⚠️ ") + msg);

                    return Results.Ok(new {
                        message               = msg,
                        FileChanged           = fileChanged,
                        ChangedCount          = changed.Count,
                        // 檔名是結構資訊（資料表名、頁面與視覺的代號），不含內容
                        ChangedFiles          = changed.OrderByDescending(c => c.Time).Take(30)
                                                       .Select(c => (c.Kind == "更新" ? "" : $"({c.Kind}) ") + Rel(c.Path)).ToList(),
                        ExpectFound           = expectFound,
                        Expect                = needles.Select(n => new {
                                                    Text  = n,
                                                    Found = found.ContainsKey(n),
                                                    File  = found.TryGetValue(n, out var f) ? Rel(f) : null
                                                }).ToList(),
                        TargetFile            = target,
                        VerifyOnly            = verifyOnly,
                        VerificationAvailable = canVerify,
                        VerificationProblem   = scope.Problem,
                        BeforeUtc             = before.Count > 0 ? before.Values.Max() : (DateTime?)null,
                        AfterUtc              = now.Count > 0 ? now.Values.Max() : (DateTime?)null,
                        WaitedSeconds         = Math.Round(sw.Elapsed.TotalSeconds, 1)
                    });
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    Console.WriteLine($"❌ 存檔失敗: {ex.Message}");
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            // 重新整理資料。回應在重新整理「做完之後」才送出 —— SaveChanges 會等引擎跑完。
            app.MapPost("/api/refresh", (RefreshRequest? req, HttpContext ctx) => {
                var rt = (req?.RefreshType ?? "full").ToLower() switch {
                    "full"        => RefreshType.Full,
                    "calculate"   => RefreshType.Calculate,
                    "dataonly"    => RefreshType.DataOnly,
                    "automatic"   => RefreshType.Automatic,
                    "add"         => RefreshType.Add,
                    "clearvalues" => RefreshType.ClearValues,
                    "defragment"  => RefreshType.Defragment,
                    _             => RefreshType.Full
                };
                string label = string.IsNullOrWhiteSpace(req?.TableName)
                    ? $"重新整理整個模型（{rt}）"
                    : $"重新整理表格 '{req!.TableName}'（{rt}）";

                return RunModel(ctx, label, (model, _) => {
                    if (string.IsNullOrWhiteSpace(req?.TableName)) {
                        model.RequestRefresh(rt);
                    } else {
                        var t = FindTable(model, req!.TableName);
                        t.RequestRefresh(rt);
                    }
                    return new { message = label + " 已排入" };   // 失敗時不會走到 afterSave，也不會回傳這個
                }, save: true,
                // 以前在 SaveChanges 之前就把回應寫成「已排入」，但呼叫端收到回應時其實已經跑完了 ——
                // 讀到「已排入」的人會以為還在跑，去等或重送一次。
                afterSave: (_, ms) => new { message = label + " 已完成", ElapsedMs = ms });
            });

            // =====================================================================
            // 快照 / 還原
            // =====================================================================

            static string SanitizeLabel(string? label) {
                if (string.IsNullOrWhiteSpace(label)) return "";
                var cleaned = Regex.Replace(label, @"[^\w\-]", "_");
                return "__" + (cleaned.Length > 40 ? cleaned.Substring(0, 40) : cleaned);
            }

            // 每個模型有自己的快照資料夾。名稱含完整路徑的雜湊 —— 不同資料夾下的同名
            // 檔案（例如桌面的 X.pbix 與專案裡的 X.pbip）必須分開，否則會互相污染。
            // create=false 用於「只是查看」的情境 —— 光是列出快照就建資料夾的話，
            // 每看過一個模型就會留下一個空資料夾。
            string SnapshotDirFor(PbiInstance inst, bool create = true) {
                string key = ModelKey(inst) ?? $"unsaved_port{inst.Port}";
                var dir = Path.Combine(snapshotPath, key);
                if (create) Directory.CreateDirectory(dir);
                return dir;
            }

            // 把整個模型定義序列化成 TMSL 存檔，作為所有破壞性操作前的安全網
            app.MapPost("/api/snapshot", (SnapshotRequest? req, HttpContext ctx) => {
                try {
                    var inst = ResolveInstance(ctx);
                    Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] 📸 建立快照：{inst.FileName}");
                    using (var server = new Server()) {
                        server.Connect($"Data Source=localhost:{inst.Port};");
                        var db = server.Databases[0];
                        var tmsl = TabularSerializer.SerializeDatabase(db, new SerializeOptions {
                            IgnoreInferredProperties = true,
                            IgnoreInferredObjects    = true,
                            IgnoreTimestamps         = true
                        });
                        string dir      = SnapshotDirFor(inst);
                        string fileName = $"{DateTime.Now:yyyyMMdd_HHmmss}{SanitizeLabel(req?.Label)}.json";
                        string full     = Path.Combine(dir, fileName);
                        File.WriteAllText(full, tmsl, new System.Text.UTF8Encoding(false));
                        var info = new FileInfo(full);
                        Console.WriteLine($"✅ 快照已建立：{fileName}（{info.Length / 1024} KB）");
                        return Results.Ok(new {
                            message  = "快照已建立",
                            File     = fileName,
                            Path     = full,
                            SourceFile = inst.FilePath,
                            SizeKB   = info.Length / 1024,
                            Tables   = db.Model.Tables.Count,
                            Measures = db.Model.Tables.Sum(t => t.Measures.Count)
                        });
                    }
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    Console.WriteLine($"❌ 快照失敗: {ex.Message}");
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            // 只列出「目標實例自己的」快照
            app.MapGet("/api/snapshots", (HttpContext ctx) => {
                try {
                    var inst = ResolveInstance(ctx);
                    var dir  = SnapshotDirFor(inst, create: false);
                    var files = (Directory.Exists(dir)
                            ? new DirectoryInfo(dir).GetFiles("*.json")
                            : Array.Empty<FileInfo>())
                        .OrderByDescending(f => f.LastWriteTime)
                        .Select(f => new { f.Name, SizeKB = f.Length / 1024, Created = f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") })
                        .ToList();
                    return Results.Ok(new {
                        SourceFile = inst.FilePath, SnapshotPath = dir,
                        Count = files.Count, Snapshots = files
                    });
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            // 從快照還原。只處理本橋接服務寫得到的物件類型，不做整庫覆蓋（那對執行中的
            // PBI Desktop 太危險）。預設 DryRun=false，但強烈建議先跑一次 DryRun 看差異。
            app.MapPost("/api/restore", async (RestoreRequest req, HttpContext ctx) => {
                if (string.IsNullOrWhiteSpace(req.File)) return Results.BadRequest("❌ File 不可為空");
                if (req.File.Contains("..") || req.File.Contains('/') || req.File.Contains('\\'))
                    return Results.BadRequest("❌ File 只接受快照資料夾內的檔名，不可含路徑");

                PbiInstance inst;
                try { inst = ResolveInstance(ctx); }
                catch (OpException ex) { return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message); }

                // 快照只在目標模型自己的資料夾裡找 —— 結構上就不可能還原到別的模型
                string full = Path.Combine(SnapshotDirFor(inst, create: false), req.File);
                if (!File.Exists(full))
                    return Results.NotFound($"在 {inst.FileName} 的快照資料夾中找不到 {req.File}（快照依模型分開存放，不能跨模型還原）");

                bool dryRun = req.DryRun ?? false;
                var scope = new HashSet<string>(
                    (req.Scope != null && req.Scope.Length > 0
                        ? req.Scope
                        : new[] { "measures", "columns", "relationships", "expressions", "mquery" })
                    .Select(s => s.ToLower()));

                Database snapDb;
                try {
                    snapDb = TabularSerializer.DeserializeDatabase(File.ReadAllText(full));
                } catch (Exception ex) {
                    return Results.BadRequest($"❌ 快照檔無法解析: {ex.Message}");
                }
                var snap = snapDb.Model;

                // 快照檔放在專案資料夾裡，是任何程式都改得到的文字檔 —— 還原 M 等於照著那個檔案改寫資料的載入方式
                if (!dryRun && (scope.Contains("mquery") || scope.Contains("expressions"))) {
                    var refused = await ConfirmDataShapeChange(inst, new[] {
                        $"從快照 {req.File} 還原 Power Query（{string.Join("、", scope.Where(s => s == "mquery" || s == "expressions").Select(s => s == "mquery" ? "各資料表的 M 腳本" : "共用查詢"))}）" });
                    if (refused != null) return refused;
                }

                return RunModel(ctx, $"從快照還原 {req.File}（範圍：{string.Join(",", scope)}）", (live, _) => {
                    var actions = new List<object>();
                    void Log(string action, string obj, string detail = "") =>
                        actions.Add(new { Action = action, Object = obj, Detail = detail });

                    foreach (var snapTable in snap.Tables) {
                        var liveTable = live.Tables.Find(snapTable.Name);
                        if (liveTable == null) { Log("跳過(表格不存在)", $"'{snapTable.Name}'", "還原不會重建整張表格"); continue; }

                        if (scope.Contains("measures")) {
                            foreach (var sm in snapTable.Measures) {
                                var lm = liveTable.Measures.Find(sm.Name);
                                if (lm == null) {
                                    if (!dryRun) liveTable.Measures.Add(new Measure {
                                        Name = sm.Name, Expression = sm.Expression, FormatString = sm.FormatString,
                                        Description = sm.Description, DisplayFolder = sm.DisplayFolder, IsHidden = sm.IsHidden });
                                    Log("新增量值", $"'{snapTable.Name}'[{sm.Name}]");
                                } else if (lm.Expression != sm.Expression || lm.FormatString != sm.FormatString) {
                                    if (!dryRun) { lm.Expression = sm.Expression; lm.FormatString = sm.FormatString;
                                                   lm.Description = sm.Description; lm.DisplayFolder = sm.DisplayFolder; }
                                    Log("還原量值公式", $"'{snapTable.Name}'[{sm.Name}]");
                                }
                            }
                            foreach (var lm in liveTable.Measures.ToList()) {
                                if (snapTable.Measures.Find(lm.Name) == null) {
                                    if (!dryRun) liveTable.Measures.Remove(lm);
                                    Log("刪除量值(快照中不存在)", $"'{snapTable.Name}'[{lm.Name}]");
                                }
                            }
                        }

                        if (scope.Contains("columns")) {
                            foreach (var sc in snapTable.Columns.OfType<CalculatedColumn>()) {
                                var lc = liveTable.Columns.Find(sc.Name);
                                if (lc == null) {
                                    if (!dryRun) liveTable.Columns.Add(new CalculatedColumn {
                                        Name = sc.Name, Expression = sc.Expression, DataType = sc.DataType });
                                    Log("新增計算資料行", $"'{snapTable.Name}'[{sc.Name}]");
                                } else if (lc is CalculatedColumn lcc && lcc.Expression != sc.Expression) {
                                    if (!dryRun) lcc.Expression = sc.Expression;
                                    Log("還原計算資料行公式", $"'{snapTable.Name}'[{sc.Name}]");
                                }
                            }
                        }

                        if (scope.Contains("mquery")) {
                            var sp = snapTable.Partitions.FirstOrDefault(p => p.Source is MPartitionSource);
                            var lp = liveTable.Partitions.FirstOrDefault(p => p.Source is MPartitionSource);
                            if (sp != null && lp != null) {
                                var sSrc = (MPartitionSource)sp.Source;
                                var lSrc = (MPartitionSource)lp.Source;
                                if (sSrc.Expression != lSrc.Expression) {
                                    if (!dryRun) lSrc.Expression = sSrc.Expression;
                                    Log("還原 M 腳本", $"'{snapTable.Name}'", "還原後需要 /api/refresh 才會生效");
                                }
                            }
                        }
                    }

                    if (scope.Contains("relationships")) {
                        string Key(SingleColumnRelationship r) =>
                            $"{r.FromTable.Name}|{r.FromColumn.Name}|{r.ToTable.Name}|{r.ToColumn.Name}";
                        var snapRels = snap.Relationships.OfType<SingleColumnRelationship>().ToList();
                        var liveRels = live.Relationships.OfType<SingleColumnRelationship>().ToList();
                        var snapKeys = new HashSet<string>(snapRels.Select(Key));
                        var liveKeys = new HashSet<string>(liveRels.Select(Key));

                        foreach (var sr in snapRels.Where(r => !liveKeys.Contains(Key(r)))) {
                            var ft = live.Tables.Find(sr.FromTable.Name);
                            var tt = live.Tables.Find(sr.ToTable.Name);
                            var fc = ft?.Columns.Find(sr.FromColumn.Name);
                            var tc = tt?.Columns.Find(sr.ToColumn.Name);
                            if (fc == null || tc == null) { Log("跳過關聯(欄位不存在)", Key(sr)); continue; }
                            if (!dryRun) live.Relationships.Add(new SingleColumnRelationship {
                                Name = Guid.NewGuid().ToString(), FromColumn = fc, ToColumn = tc,
                                FromCardinality = sr.FromCardinality, ToCardinality = sr.ToCardinality,
                                CrossFilteringBehavior = sr.CrossFilteringBehavior, IsActive = sr.IsActive });
                            Log("新增關聯", Key(sr));
                        }
                        foreach (var lr in liveRels.Where(r => !snapKeys.Contains(Key(r)))) {
                            if (!dryRun) live.Relationships.Remove(lr);
                            Log("刪除關聯(快照中不存在)", Key(lr));
                        }
                    }

                    if (scope.Contains("expressions")) {
                        foreach (var se in snap.Expressions) {
                            var le = live.Expressions.Find(se.Name);
                            if (le == null) {
                                if (!dryRun) live.Expressions.Add(new NamedExpression {
                                    Name = se.Name, Kind = se.Kind, Expression = se.Expression, Description = se.Description });
                                Log("新增共用運算式", se.Name);
                            } else if (le.Expression != se.Expression) {
                                if (!dryRun) le.Expression = se.Expression;
                                Log("還原共用運算式", se.Name);
                            }
                        }
                        foreach (var le in live.Expressions.ToList()) {
                            if (snap.Expressions.Find(le.Name) == null) {
                                if (!dryRun) live.Expressions.Remove(le);
                                Log("刪除共用運算式(快照中不存在)", le.Name);
                            }
                        }
                    }

                    // ⚠️ 還原刻意不刪除表格與角色。這裡把「沒被涵蓋到的東西」明確列出來，
                    //    否則呼叫端會誤以為模型已完全回到快照當時的狀態。
                    var uncoveredTables = live.Tables
                        .Where(lt => snap.Tables.Find(lt.Name) == null
                                     && !lt.Name.StartsWith("LocalDateTable_")
                                     && !lt.Name.StartsWith("DateTableTemplate_"))
                        .Select(lt => new {
                            Table = lt.Name,
                            Measures = lt.Measures.Count,
                            IsCalculationGroup = lt.CalculationGroup != null
                        }).ToList();

                    var uncoveredRoles = live.Roles
                        .Where(r => snap.Roles.Find(r.Name) == null)
                        .Select(r => r.Name).ToList();

                    bool implicitChanged = live.DiscourageImplicitMeasures != snap.DiscourageImplicitMeasures;

                    int uncovered = uncoveredTables.Count + uncoveredRoles.Count + (implicitChanged ? 1 : 0);

                    return new {
                        message     = (dryRun ? "DryRun 完成，未寫入任何變更" : $"還原完成，共套用 {actions.Count} 項變更")
                                      + (uncovered > 0 ? $"；⚠️ 另有 {uncovered} 項不在還原範圍內，請見 Uncovered" : ""),
                        Snapshot    = req.File,
                        DryRun      = dryRun,
                        Scope       = scope.ToList(),
                        ChangeCount = actions.Count,
                        Changes     = actions,
                        Uncovered   = new {
                            NewTablesNotRemoved = uncoveredTables,
                            NewRolesNotRemoved  = uncoveredRoles,
                            DiscourageImplicitMeasuresDiffers = implicitChanged
                                ? $"目前={live.DiscourageImplicitMeasures}，快照={snap.DiscourageImplicitMeasures}"
                                : null,
                            Note = "還原不會刪除表格、角色，也不會改模型層級屬性 —— 請用 Remove-PbiTable / Remove-PbiRole / Set-PbiModelProps 自行處理。"
                        }
                    };
                }, save: !dryRun);
            });

            // =====================================================================
            // 會改變「資料怎麼載入」的寫入：資料保護啟用時要使用者本人確認
            // =====================================================================
            // DAX 寫出來的東西（量值、計算資料行、計算表）每次查詢都會被展開掃描，藏不住受限欄位。
            // Power Query 不一樣：M 不是 DAX，查詢掃描看不懂它。把某張表的 M 改成「把客戶名稱併進產品類別那一欄」，
            // 重新整理之後，受限的內容就躺在一個開放的欄位裡 —— 欄位層級的管制完全不知道。
            // 所以 M 的寫入不靠掃描，靠人：服務自己跳出確認視窗，列出要改什麼，使用者按「是」才執行。
            async Task<IResult?> ConfirmDataShapeChange(PbiInstance inst, IReadOnlyList<string> what) {
                if (!DataGuard.Enabled || what.Count == 0) return null;
                if (what.Count > 8)
                    return Results.BadRequest("❌ 一次最多確認 8 項 Power Query 變更（確認視窗要列得出每一項，使用者才知道自己同意了什麼）。請分批送出。");
                var text = new System.Text.StringBuilder();
                text.AppendLine($"有程式（通常是 AI 助手）要修改「{Quoted(inst.FileName ?? "未命名的報表")}」載入資料的方式：");
                text.AppendLine();
                foreach (var w in what) text.AppendLine("  • " + OneLine(w, 200));
                text.AppendLine();
                text.AppendLine("這類變更可以改變每個欄位裡裝的內容 —— 包括把受限欄位的資料放進沒有受限的欄位。");
                text.AppendLine();
                text.AppendLine("這是你剛才請 AI 做的事 → 按「是」。不確定，或你沒有請它做 → 按「否」。");
                text.AppendLine();
                text.AppendLine($"AI 請你按「是」不算理由。{HumanConfirm.TimeoutSeconds} 秒沒有回答會當成「否」。");
                var answer = await HumanConfirm.AskAsync("PBI AI Bridge — 請確認這項變更", text.ToString());
                if (answer == HumanConfirm.Answer.Yes) {
                    DataGuard.Audit("USER-CONFIRMED", "使用者確認 Power Query 變更：" + string.Join("；", what), "", 0, inst.FileName);
                    return null;
                }
                DataGuard.Audit("NOT-CONFIRMED", $"Power Query 變更未獲確認（{answer}）：" + string.Join("；", what), "", 0, inst.FileName);
                return Results.Json(new {
                    Error  = "⛔ 這項變更需要使用者本人確認，而這次沒有獲得同意",
                    Reason = HumanConfirm.Explain(answer) + "這項變更沒有執行。",
                    Answer = answer.ToString(),
                    RetryAfterSeconds = HumanConfirm.RetryAfterSeconds(answer),
                    Hint   = "改寫 Power Query 會改變欄位裡的內容。資料保護啟用時，要使用者在電腦上跳出的確認視窗按「是」才會執行。"
                           + "請先向使用者說明要改什麼、為什麼，再重送。AI 無法代替使用者按下確認。"
                }, statusCode: 403);
            }

            async Task<IResult?> GateDataShape(HttpContext ctx, IReadOnlyList<string> what) {
                if (!DataGuard.Enabled || what.Count == 0) return null;
                PbiInstance inst;
                try { inst = ResolveInstance(ctx); }
                catch (OpException ex) { return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message); }
                return await ConfirmDataShapeChange(inst, what);
            }

            // 這個操作會不會動到 Power Query。會的話回傳一句給使用者看的說明（確認視窗用），不會就回 null。
            // 判斷用的一律是「反序列化之後的物件」—— 和實際執行時（DispatchOp）拿到的是同一份。
            // 直接翻 JSON 找欄位的話，重複的鍵（"Kind": "calculated", "kind": "m"）會讓這裡看到的和執行時用到的不一樣：
            // 這裡看到 calculated 而放行，執行時用的卻是 m。
            static string? NoteUpdateM(UpdateMRequest r) => $"覆寫資料表「{Quoted(r.TableName)}」的 Power Query（M 腳本）";
            static string? NoteExpression(ExpressionRequest r) => $"寫入共用查詢／參數「{Quoted(r.Name)}」";
            static string? NoteCreateTable(CreateTableRequest r) =>
                (r.Kind ?? "calculated").ToLower() is "m" or "powerquery" ? $"用 Power Query 新建資料表「{Quoted(r.TableName)}」" : null;
            static string? NoteAddColumn(AddColumnRequest r) => string.IsNullOrWhiteSpace(r.SourceColumn) ? null
                : $"在資料表「{Quoted(r.TableName)}」新增來源資料行「{Quoted(r.ColumnName)}」（對應 Power Query 輸出的「{Quoted(r.SourceColumn)}」）";
            static string? DataShapeNote(string? op, System.Text.Json.JsonElement args) => (op ?? "").ToLower() switch {
                "update-m"          => NoteUpdateM(Args<UpdateMRequest>(args)),
                "upsert-expression" => NoteExpression(Args<ExpressionRequest>(args)),
                "create-table"      => NoteCreateTable(Args<CreateTableRequest>(args)),
                "add-column"        => NoteAddColumn(Args<AddColumnRequest>(args)),
                _ => null
            };

            // =====================================================================
            // 寫入（單一操作）—— 全部走 DispatchOp，與 /api/batch 共用實作
            // =====================================================================

            // M 腳本寫入。只改模型那一份，寫入本身不會重抓資料。
            app.MapPost("/api/update-m", async (UpdateMRequest req, HttpContext ctx) => {
                var refused = await GateDataShape(ctx, new[] { NoteUpdateM(req)! });
                if (refused != null) return refused;
                return RunModel(ctx, $"寫入 M 腳本 '{req.TableName}'", (m, _) => new {
                    message = OpUpdateM(m, req),
                    note    = "這只改模型那一份，資料還沒變。Power BI Desktop 若顯示「查詢中有暫止的變更尚未套用」，請使用者按「套用」；" +
                              "沒有出現提示的話，對這張表跑一次 full refresh，新的 M 才會套用到資料上。" +
                              "之後用 Get-PbiMQuery 確認留下的是這次寫入的版本 —— 兩份不同時 Desktop 那份可能覆蓋回來。"
                }, save: true);
            });

            app.MapPost("/api/upsert-measure", (UpsertMeasureRequest req, HttpContext ctx) =>
                RunModel(ctx, $"寫入量值 '{req.TableName}'[{req.MeasureName}]", (m, _) => new { message = OpUpsertMeasure(m, req) }, save: true));

            app.MapPost("/api/delete-measure", (DeleteMeasureRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除量值 '{req.TableName}'[{req.MeasureName}]", (m, _) => new { message = OpDeleteMeasure(m, req) }, save: true));

            app.MapPost("/api/move-measure", (MoveMeasureRequest req, HttpContext ctx) =>
                RunModel(ctx, $"搬移量值 [{req.MeasureName}] → '{req.ToTable}'", (m, _) => new { message = OpMoveMeasure(m, req) }, save: true));

            app.MapPost("/api/add-column", async (AddColumnRequest req, HttpContext ctx) => {
                var note = NoteAddColumn(req);
                if (note != null) { var refused = await GateDataShape(ctx, new[] { note }); if (refused != null) return refused; }
                return RunModel(ctx, $"新增資料行 '{req.TableName}'[{req.ColumnName}]", (m, _) => new { message = OpAddColumn(m, req) }, save: true);
            });

            app.MapPost("/api/upsert-relationship", (RelationshipRequest req, HttpContext ctx) =>
                RunModel(ctx, "建立/更新關聯", (m, _) => new { message = OpUpsertRelationship(m, req) }, save: true));

            app.MapPost("/api/delete-relationship", (RelationshipRefRequest req, HttpContext ctx) =>
                RunModel(ctx, "刪除關聯", (m, _) => new { message = OpDeleteRelationship(m, req) }, save: true));

            app.MapPost("/api/create-table", async (CreateTableRequest req, HttpContext ctx) => {
                var note = NoteCreateTable(req);
                if (note != null) { var refused = await GateDataShape(ctx, new[] { note }); if (refused != null) return refused; }
                return RunModel(ctx, $"建立表格 '{req.TableName}'", (m, _) => new { message = OpCreateTable(m, req) }, save: true);
            });

            app.MapPost("/api/delete-table", (TableRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除表格 '{req.TableName}'", (m, _) => new { message = OpDeleteTable(m, req) }, save: true));

            app.MapPost("/api/delete-column", (ColumnRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除資料行 '{req.TableName}'[{req.ColumnName}]", (m, _) => new { message = OpDeleteColumn(m, req) }, save: true));

            app.MapPost("/api/set-column-props", (ColumnPropsRequest req, HttpContext ctx) =>
                RunModel(ctx, $"設定資料行屬性 '{req.TableName}'[{req.ColumnName}]", (m, _) => new { message = OpSetColumnProps(m, req) }, save: true));

            app.MapPost("/api/rename", (RenameRequest req, HttpContext ctx) => {
                PbiInstance? target = null;
                Model? live = null;
                ProtectionStore.LevelSnapshot? before = null;
                bool dry = req.DryRun ?? false;
                return RunModel(ctx, $"改名 {req.ObjectType} [{req.OldName}] → [{req.NewName}]",
                         (m, inst) => {
                             target = inst; live = m;
                             if (!dry) before = ProtectionStore.Capture(inst, DataGuard.Derivations.FromModel(m));
                             return OpRename(m, req);
                         }, save: !dry,
                         // 改名成功之後把保護等級釘回去 —— 否則改個名字，照欄名比對的通用規則就對不上，保護就掉了
                         afterSave: (r, _) => {
                             if (target != null && live != null && before != null) {
                                 int pinned = ProtectionStore.Pin(target, DataGuard.Derivations.FromModel(live), before, new[] { req });
                                 if (pinned > 0) Console.WriteLine($"🛡️ 改名之後有 {pinned} 個欄位的等級會變寬鬆，已補上逐欄設定維持原本的等級");
                             }
                             return r;
                         });
            });

            app.MapPost("/api/set-model-props", (ModelPropsRequest req, HttpContext ctx) =>
                RunModel(ctx, "設定模型層級屬性", (m, _) => new { message = OpSetModelProps(m, req) }, save: true));

            app.MapGet("/api/model-props", (HttpContext ctx) => RunModel(ctx, "讀取模型層級屬性", (m, _) => new {
                m.Name,
                m.DiscourageImplicitMeasures,
                m.Description,
                CompatibilityLevel = m.Database?.CompatibilityLevel,
                CalculationGroups  = m.Tables.Count(t => t.CalculationGroup != null)
            }, save: false));

            app.MapPost("/api/upsert-calc-group", (CalcGroupRequest req, HttpContext ctx) =>
                RunModel(ctx, $"建立/更新計算群組 '{req.TableName}'", (m, _) => new { message = OpUpsertCalcGroup(m, req) }, save: true));

            app.MapPost("/api/upsert-calc-item", (CalcItemRequest req, HttpContext ctx) =>
                RunModel(ctx, $"寫入計算項目 '{req.TableName}'::[{req.ItemName}]", (m, _) => new { message = OpUpsertCalcItem(m, req) }, save: true));

            app.MapPost("/api/delete-calc-item", (CalcItemRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除計算項目 '{req.TableName}'::[{req.ItemName}]", (m, _) => new { message = OpDeleteCalcItem(m, req) }, save: true));

            app.MapPost("/api/upsert-role", (RoleRequest req, HttpContext ctx) =>
                RunModel(ctx, $"建立/更新角色 [{req.RoleName}]", (m, _) => new { message = OpUpsertRole(m, req) }, save: true));

            app.MapPost("/api/delete-role", (RoleRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除角色 [{req.RoleName}]", (m, _) => new { message = OpDeleteRole(m, req) }, save: true));

            app.MapPost("/api/upsert-expression", async (ExpressionRequest req, HttpContext ctx) => {
                var refused = await GateDataShape(ctx, new[] { NoteExpression(req)! });
                if (refused != null) return refused;
                return RunModel(ctx, $"寫入共用運算式 [{req.Name}]", (m, _) => new { message = OpUpsertExpression(m, req) }, save: true);
            });

            app.MapPost("/api/delete-expression", (ExpressionRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除共用運算式 [{req.Name}]", (m, _) => new { message = OpDeleteExpression(m, req) }, save: true));

            // =====================================================================
            // 批次：一次連線、一次 SaveChanges，避免 N 次模型重算
            // =====================================================================
            app.MapPost("/api/batch", async (BatchRequest req, HttpContext ctx) => {
                if (req.Operations == null || req.Operations.Length == 0)
                    return Results.BadRequest("❌ Operations 不可為空");
                if (req.Operations.Length > 200)
                    return Results.BadRequest("❌ 單次批次上限 200 個操作");

                bool stopOnError = req.StopOnError ?? true;
                bool savePerOp   = req.SavePerOp   ?? false;
                bool dryRun      = req.DryRun      ?? false;

                // 批次裡夾著 Power Query 的寫入：一樣要使用者確認，不能因為包在批次裡就免了
                if (!dryRun) {
                    List<string> notes;
                    try { notes = req.Operations.Select(o => DataShapeNote(o.Op, o.Args)).Where(n => n != null).Select(n => n!).ToList(); }
                    catch (OpException ex) { return Results.BadRequest(ex.Message); }       // Args 格式不對：反正執行時也會失敗
                    var refused = await GateDataShape(ctx, notes);
                    if (refused != null) return refused;
                }

                PbiInstance? target = null;
                Model? live = null;
                ProtectionStore.LevelSnapshot? beforeRename = null;
                var renames = new List<RenameRequest>();       // 成功存檔之後，保護等級要跟著欄位走（見 ProtectionStore.Pin）
                void ApplyRenames() {
                    if (target == null || live == null || beforeRename == null || renames.Count == 0) return;
                    int pinned = ProtectionStore.Pin(target, DataGuard.Derivations.FromModel(live), beforeRename, renames.ToList());
                    if (pinned > 0) Console.WriteLine($"🛡️ 改名之後有 {pinned} 個欄位的等級會變寬鬆，已補上逐欄設定維持原本的等級");
                    renames.Clear();
                }

                return RunModel(ctx, $"批次執行 {req.Operations.Length} 個操作" +
                                (savePerOp ? "（逐步存檔）" : "（最後一次存檔）"), (model, inst) => {
                    target = inst; live = model;
                    if (!dryRun && req.Operations.Any(o => (o.Op ?? "").Equals("rename", StringComparison.OrdinalIgnoreCase)))
                        beforeRename = ProtectionStore.Capture(inst, DataGuard.Derivations.FromModel(model));
                    var results = new List<object>();
                    int ok = 0, failed = 0;

                    for (int i = 0; i < req.Operations.Length; i++) {
                        var opDef = req.Operations[i];
                        try {
                            var r = DispatchOp(model, opDef.Op, opDef.Args);
                            if (!dryRun && (opDef.Op ?? "").Equals("rename", StringComparison.OrdinalIgnoreCase))
                                renames.Add(Args<RenameRequest>(opDef.Args));
                            if (savePerOp && !dryRun) { model.SaveChanges(); ApplyRenames(); }
                            results.Add(new { Index = i, Op = opDef.Op, Status = "ok", Result = r });
                            ok++;
                        } catch (Exception ex) {
                            results.Add(new { Index = i, Op = opDef.Op, Status = "error", Result = (object)ex.Message });
                            failed++;
                            if (stopOnError) {
                                // 尚未存檔的變更會隨連線釋放而丟棄；已存檔的（savePerOp）則會留下
                                throw new OpException(
                                    $"❌ 第 {i} 個操作 [{opDef.Op}] 失敗：{ex.Message}\n" +
                                    (savePerOp
                                        ? $"⚠️ SavePerOp=true，前 {ok} 個操作已寫入模型，不會回滾。"
                                        : "✅ 本批次未存檔，所有變更已丟棄，模型維持原狀。"));
                            }
                        }
                    }
                    return new {
                        message = $"批次完成：{ok} 成功 / {failed} 失敗",
                        Total = req.Operations.Length, Succeeded = ok, Failed = failed,
                        DryRun = dryRun, Results = results
                    };
                }, save: !dryRun && !savePerOp, afterSave: (r, _) => { ApplyRenames(); return r; }, exclusive: !dryRun);
            });

            Console.WriteLine("✅ 伺服器已在背景運行 (等待網頁於 Port 5500 呼叫)！");
            // 服務真的開始接受連線後才開瀏覽器 —— 由啟動檔先開的話，頁面會比服務早到而顯示連線失敗。
            // 開的是網址（交給預設瀏覽器），不是執行檔。
            const string dashboardUrl = "http://localhost:5500/";
            app.Lifetime.ApplicationStarted.Register(() => {
                Console.WriteLine($"🌐 網頁儀表板：{dashboardUrl}");
                try { Process.Start(new ProcessStartInfo(dashboardUrl) { UseShellExecute = true }); }
                catch (Exception ex) { Console.WriteLine($"⚠️ 無法自動開啟瀏覽器（{ex.Message}），請手動開啟上面的網址"); }
            });

            app.Run();
        }
    }
}
