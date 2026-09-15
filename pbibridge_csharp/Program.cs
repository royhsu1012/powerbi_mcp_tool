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
    // 註：這裡刻意沒有 UpdateMQueryRequest —— M 腳本一律唯讀，寫入路徑已移除。
    //     原因見 OpUpdateM 移除處的說明。
    public record UpsertMeasureRequest(string TableName, string MeasureName, string Expression,
                                       string FormatString, string Description,
                                       string DisplayFolder, bool? IsHidden);
    public record MoveMeasureRequest(string FromTable, string MeasureName, string ToTable);
    public record AddColumnRequest(string TableName, string ColumnName, string Expression, string DataType);
    public record DeleteMeasureRequest(string TableName, string MeasureName);
    // 唯讀 DAX 查詢：讓 AI 寫完量值後能立即驗算結果
    public record DaxQueryRequest(string Query, int? MaxRows, int? TimeoutSeconds, bool? Pseudonymize = null);
    // DMV 查詢：$SYSTEM 系統檢視，用於效能分析與相依性追蹤
    public record DmvQueryRequest(string Query, int? MaxRows, int? TimeoutSeconds);

    // 存檔 / 重新整理
    public record SaveRequest(int? WaitSeconds);
    public record RefreshRequest(string TableName, string RefreshType);

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
    public record TableRefRequest(string TableName);
    public record ColumnRefRequest(string TableName, string ColumnName);
    public record RenameRequest(string ObjectType, string TableName, string OldName, string NewName, bool? DryRun);
    public record ColumnPropsRequest(string TableName, string ColumnName, string FormatString, string DisplayFolder,
                                     bool? IsHidden, string SortByColumn, string SummarizeBy,
                                     string DataCategory, string Description, string DataType);

    // 模型層級屬性
    public record ModelPropsRequest(bool? DiscourageImplicitMeasures, string DefaultMode, string Description);

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

    // 視覺層注入：寫入 visual.json 並自動重啟 PBI Desktop
    // ⛔ 安全修補：路徑改由 appsettings.json 讀取，不接受外部傳入
    public record InjectVisualRequest(
        string PageId,          // 目標頁面的 hash ID
        string PageDisplayName, // 頁面顯示名稱
        bool   IsNewPage,       // 是否為新頁面
        string VisualId,        // 視覺物件的 hash ID
        string VisualJson       // visual.json 的完整內容
    );

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
                    FileName    : file == null ? null : Path.GetFileName(file),
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

        /// <summary>
        /// 連上執行中的模型，執行 action，視需要 SaveChanges，並統一處理錯誤回應。
        /// 所有會動到模型的端點都走這裡，確保連線與錯誤處理只有一份實作。
        /// </summary>
        static IResult RunModel(HttpContext ctx, string label, Func<Model, PbiInstance, object> action, bool save) {
            try {
                var inst = ResolveInstance(ctx);
                Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] {label}  →  {inst.FileName ?? "(未命名)"} :{inst.Port}");
                using (var server = new Server()) {
                    server.Connect($"Data Source=localhost:{inst.Port};");
                    var model = server.Databases[0].Model;
                    var result = action(model, inst);
                    // ⚠️ 未呼叫 SaveChanges 時，暫存的變更會隨 server 釋放而丟棄 —— DryRun 靠的就是這點
                    if (save) model.SaveChanges();
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

        // ── M 腳本：唯讀 ──────────────────────────────────────────────────────
        //
        // 這裡曾經有 OpUpdateM（透過 TOM 覆寫 MPartitionSource.Expression），已於
        // 2026-08-06 移除，不要加回來。
        //
        // 原因：Power BI Desktop 的 Power Query 文件與 TOM 模型是**兩份獨立的東西**。
        // 用 TOM 改 MPartitionSource 只動到模型那份，Desktop 自己那份不會跟著變，
        // 兩邊一失步 Desktop 就會永久顯示「查詢中有暫止的變更尚未套用」——
        // 按「套用變更」是拿 Desktop 那份舊 M 去跑，跑完不一致依然存在，橫幅又回來，
        // 陷入死迴圈，只能請使用者手動到進階編輯器貼一次才解得開。
        //
        // 正確做法：M 只讀不寫。要改 M 就把完整的 let...in 交給使用者，
        // 由他在 Power Query 編輯器貼上並「關閉並套用」，走 Desktop 自己的路徑。
        // 讀取請用 /api/schema 的 MQuery 欄位（Get-PbiMQuery）。
        //
        // 註：/api/restore 的 mquery 範圍與 /api/create-table 的 Kind=m 仍會寫入 M，
        //     是刻意保留的；使用前請確認 Power BI Desktop 的後果。

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
            // "update-m" 已移除 —— M 腳本唯讀，原因見 OpUpdateM 移除處的說明
            "update-m"            => throw new OpException(
                "❌ update-m 已移除：M 腳本一律唯讀。用 TOM 改 M 會讓 Power BI Desktop 的 " +
                "Power Query 文件與模型失步，卡在「查詢中有暫止的變更尚未套用」。" +
                "請把完整的 let...in 交給使用者，由他在進階編輯器貼上並「關閉並套用」。", 410),
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
        // 存檔（模擬 Ctrl+S）
        // =====================================================================

        /// <summary>
        /// 取得目標檔案的最新修改時間，用來驗證存檔是否真的發生。
        ///   · .pbix 是單一檔案 → 直接看它的時間
        ///   · .pbip 會寫到旁邊的 .Report / .SemanticModel 資料夾 → 掃整個專案資料夾
        /// 拿不到時間時一併回報原因 —— 靜靜回傳 null 會讓呼叫端誤以為「沒有變更」。
        /// </summary>
        static (DateTime? Time, string? Problem) GetLatestWriteTime(string targetPath) {
            try {
                if (string.IsNullOrWhiteSpace(targetPath))
                    return (null, "找不到可驗證的檔案路徑");
                var full = Path.GetFullPath(targetPath);

                if (full.EndsWith(".pbix", StringComparison.OrdinalIgnoreCase)) {
                    if (!File.Exists(full)) return (null, $"檔案不存在：{full}");
                    return (File.GetLastWriteTimeUtc(full), null);
                }

                // .pbip：存檔會散落在同一層的多個資料夾，掃整個專案資料夾取最新
                var dir = Path.GetDirectoryName(full);
                if (dir == null || !Directory.Exists(dir))
                    return (null, $"資料夾不存在，無法驗證存檔結果：{dir}");

                DateTime latest = DateTime.MinValue;
                foreach (var f in Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories)) {
                    try { var w = File.GetLastWriteTimeUtc(f); if (w > latest) latest = w; } catch { }
                }
                return latest == DateTime.MinValue
                    ? (null, $"資料夾內沒有可讀取的檔案：{dir}")
                    : (latest, null);
            } catch (Exception ex) {
                return (null, $"掃描檔案時間失敗：{ex.Message}");
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
        /// 兩個等級：
        ///   Deny          — 客戶身分。只准出現在「計數類」函式內，因為那類函式的回傳值
        ///                   必定是數字。刻意不含 MAX/MIN/CONCATENATEX —— 對文字欄位
        ///                   MAX 會直接吐出一個真的客戶名稱。
        ///   AggregateOnly — 金額。可用任何數值聚合，但不得逐列取值。
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

            // 回傳值必定是數字的函式 —— 客戶身分欄位唯一的合法容身處。
            // DISTINCTCOUNT(v[customer_name]) 回傳 1503，是結構資訊，該放行；
            // MAX(v[account name]) 回傳一個真實客戶名稱，不在此列。
            static readonly HashSet<string> CountingFuncs = new(StringComparer.OrdinalIgnoreCase) {
                "COUNTROWS", "COUNT", "COUNTA", "COUNTAX", "COUNTX", "COUNTBLANK",
                "DISTINCTCOUNT", "DISTINCTCOUNTNOBLANK",
                "ISBLANK", "ISEMPTY", "HASONEVALUE", "ISFILTERED", "ISCROSSFILTERED"
            };

            // 數值聚合 —— 金額欄位的合法容身處。這裡可以有 MAX/MIN，
            // 因為金額的極值是統計量，不是身分。
            static readonly HashSet<string> NumericAggFuncs = new(StringComparer.OrdinalIgnoreCase) {
                "SUM", "SUMX", "AVERAGE", "AVERAGEX", "MIN", "MINX", "MAX", "MAXX",
                "MEDIAN", "MEDIANX", "PERCENTILE.INC", "PERCENTILE.EXC",
                "STDEV.P", "STDEV.S", "VAR.P", "VAR.S",
                "PRODUCT", "PRODUCTX", "GEOMEAN", "GEOMEANX", "DIVIDE", "RANKX"
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
                if (_globCache.TryGetValue(glob, out var re)) return re;
                re = new Regex("^" + string.Join(".*", glob.Split('*').Select(Regex.Escape)) + "$",
                               RegexOptions.Compiled);
                _globCache[glob] = re;
                return re;
            }

            /// <summary>
            /// 清單項目可以是完整欄名，也可以是含 * 的樣式。
            /// 樣式才是主力：欄位有一千五百個以上，而且會一直新增 ——
            /// 逐一列舉永遠追不上，而漏掉的那一個不會有人發現。
            /// </summary>
            // 去敏模式下，身分欄位唯一被允許的用法：當分組鍵／取相異值。
            // 這些函式會讓欄位以 'Table'[Column] 的原名出現在結果欄位裡，
            // Layer B 的欄名比對抓得到，也就遮得掉。
            // 反過來 MAX / CONCATENATEX / SELECTCOLUMNS 會讓真名躲在別名底下，
            // 欄名比對看不見 —— 那些用法即使開了去敏也一律擋。
            static readonly HashSet<string> GroupingShapes = new(StringComparer.OrdinalIgnoreCase) {
                "SUMMARIZECOLUMNS", "SUMMARIZE", "GROUPBY",
                "VALUES", "DISTINCT", "ALL", "ALLSELECTED", "ALLNOBLANKROW"
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
            /// </summary>
            public static string? Pseudonym(object? value) {
                if (value == null) return null;
                string v = value.ToString() ?? "";
                if (v.Length == 0) return v;
                using var h = new System.Security.Cryptography.HMACSHA256(_pseudoKey);
                return "ID_" + Convert.ToHexString(
                    h.ComputeHash(System.Text.Encoding.UTF8.GetBytes(v)), 0, 3);
            }

            /// <summary>結果欄位裡哪幾個索引需要換成代號。</summary>
            public static HashSet<int> MaskIndexes(IReadOnlyList<string> columns) {
                var idx = new HashSet<int>();
                for (int i = 0; i < columns.Count; i++) {
                    string bare = columns[i] ?? "";
                    int lb = bare.LastIndexOf('[');
                    if (lb >= 0) bare = bare.Substring(lb + 1);
                    bare = bare.TrimEnd(']').Trim();
                    if (Matches(bare, AllowColumns)) continue;
                    if (Matches(bare, DenyColumns)) idx.Add(i);
                }
                return idx;
            }

            static bool Matches(string token, string[] cols) {
                string t = Norm(token);
                if (t.Length == 0) return false;
                foreach (var c in cols) {
                    if (string.IsNullOrWhiteSpace(c)) continue;
                    if (c.IndexOf('*') >= 0) { if (GlobRegex(Norm(c)).IsMatch(t)) return true; }
                    else if (t == Norm(c)) return true;
                }
                return false;
            }

            /// <summary>取得緊接在左括號之前的識別字（函式名）。沒有就回空字串。</summary>
            static string IdentifierBefore(string s, int parenIndex) {
                int j = parenIndex - 1;
                while (j >= 0 && char.IsWhiteSpace(s[j])) j--;
                int end = j;
                while (j >= 0 && (char.IsLetterOrDigit(s[j]) || s[j] == '_' || s[j] == '.')) j--;
                return end > j ? s.Substring(j + 1, end - j) : "";
            }

            /// <summary>
            /// 掃過 DAX 文字，回傳每一個中括號參考以及包住它的函式名堆疊。
            /// 字串常值、單引號表名與註解都要跳過 —— 裡頭的括號會打亂堆疊，
            /// 而堆疊一亂，「這個欄位有沒有被聚合包住」的判斷就跟著錯。
            /// </summary>
            static List<(string Token, List<string> Enclosing)> ScanBracketTokens(string dax) {
                var result = new List<(string, List<string>)>();
                var stack  = new List<string>();
                string s = dax ?? "";
                int i = 0;
                while (i < s.Length) {
                    char c = s[i];

                    if (c == '"') {                                   // 字串常值（"" 為跳脫）
                        i++;
                        while (i < s.Length) {
                            if (s[i] == '"') {
                                if (i + 1 < s.Length && s[i + 1] == '"') { i += 2; continue; }
                                i++; break;
                            }
                            i++;
                        }
                        continue;
                    }
                    if (c == '\'') {                                  // 'Table Name' 形式的表名
                        i++;
                        while (i < s.Length && s[i] != '\'') i++;
                        i++;
                        continue;
                    }
                    if (c == '/' && i + 1 < s.Length && s[i + 1] == '/') { while (i < s.Length && s[i] != '\n') i++; continue; }
                    if (c == '-' && i + 1 < s.Length && s[i + 1] == '-') { while (i < s.Length && s[i] != '\n') i++; continue; }
                    if (c == '/' && i + 1 < s.Length && s[i + 1] == '*') {
                        i += 2;
                        while (i + 1 < s.Length && !(s[i] == '*' && s[i + 1] == '/')) i++;
                        i += 2;
                        continue;
                    }
                    if (c == '[') {
                        int close = s.IndexOf(']', i + 1);
                        if (close < 0) { i++; continue; }
                        result.Add((s.Substring(i + 1, close - i - 1).Trim(), new List<string>(stack)));
                        i = close + 1;
                        continue;
                    }
                    if (c == '(') { stack.Add(IdentifierBefore(s, i)); i++; continue; }
                    if (c == ')') { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); i++; continue; }
                    i++;
                }
                return result;
            }

            // 真正做聚合的函式。刻意不含 ISBLANK / ISEMPTY / HASONEVALUE 這類布林述詞 ——
            // 它們雖然回傳純量（所以在 CheckQuery 裡算合法的容身處），卻常出現在
            // FILTER 的條件裡，而 FILTER 回傳的是**原始資料列**。把它們當成
            // 「這是彙總查詢」，EVALUATE FILTER(表, ISBLANK(欄)) 就會拿到彙總額度，
            // 明細上限整個失效。這是實測抓到的，不是假想。
            static readonly Regex _aggCall = new Regex(
                @"\b(SUMX?|AVERAGEX?|MINX?|MAXX?|MEDIANX?|PRODUCTX?|GEOMEANX?|COUNTX?|COUNTAX?"
              + @"|COUNTROWS|COUNTA|COUNTBLANK|DISTINCTCOUNT|DISTINCTCOUNTNOBLANK"
              + @"|PERCENTILE\.(INC|EXC)|STDEV\.[PS]|VAR\.[PS]|DIVIDE|RANKX)\s*\(",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            // 會產出「每列一組統計量」的表格函式。查詢最外層必須是這些之一，
            // 才有資格套用比較寬的彙總額度。FILTER / TOPN / SELECTCOLUMNS / VALUES
            // 一律不在此列 —— 它們吐的是資料列。
            static readonly HashSet<string> AggregateShapes = new(StringComparer.OrdinalIgnoreCase) {
                "SUMMARIZECOLUMNS", "SUMMARIZE", "GROUPBY", "ROW", "UNION", "DATATABLE"
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
            /// 兩個都要，判定才會偏保守：認不出來就當明細，套 20 列。
            ///
            /// 傳入的必須是「已展開量值定義」的文字，否則
            /// SUMMARIZECOLUMNS(日期[月], "額", [銷售總額]) 會因為看不到量值裡的 SUM
            /// 而被誤判成明細，正常的趨勢分析就會被卡在 20 列。
            /// </summary>
            public static bool IsAggregateQuery(string dax) {
                if (!AggregateShapes.Contains(OuterFunction(dax))) return false;
                // 字串常值裡的 "SUM(" 不算數
                string bare = Regex.Replace(dax ?? "", "\"(?:[^\"]|\"\")*\"", "\"\"");
                return _aggCall.IsMatch(bare);
            }

            /// <summary>模型裡所有「DAX 衍生定義」：量值、計算資料行、計算表。</summary>
            public sealed class Derivations {
                // 以中括號參考的（量值名、計算資料行名）
                public Dictionary<string, string> ByToken = new(StringComparer.OrdinalIgnoreCase);
                // 以表名參考的（計算表）—— 表名不帶中括號，只能用字串包含比對
                public Dictionary<string, string> ByTable = new(StringComparer.OrdinalIgnoreCase);
                public int Count => ByToken.Count + ByTable.Count;
            }

            /// <summary>
            /// 讀取模型中所有 DAX 衍生定義。
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

                void Slurp(string sql, Action<string, string> take) {
                    try {
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = sql;
                        using var rd = cmd.ExecuteReader();
                        while (rd.Read()) {
                            string? n = rd[0]?.ToString();
                            string? e = rd[1]?.ToString();
                            if (!string.IsNullOrEmpty(n) && !string.IsNullOrWhiteSpace(e)) take(n!, e!);
                        }
                    } catch {
                        // 取不到就少展開一層。Layer B 仍在，但深度會下降 ——
                        // 這是降級不是靜默失敗，稽核記錄裡看得到展開筆數異常偏低。
                    }
                }

                Slurp("SELECT [Name], [Expression] FROM $SYSTEM.TMSCHEMA_MEASURES",
                      (n, e) => d.ByToken[n] = e);
                // 計算資料行：把客戶名稱複製到一個不在管制清單裡的新欄名，是最直覺的洗資料手法。
                // 一般資料行的 Expression 是空的，會被 Slurp 濾掉。
                Slurp("SELECT [ExplicitName], [Expression] FROM $SYSTEM.TMSCHEMA_COLUMNS",
                      (n, e) => d.ByToken[n] = e);

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
                        d.ByTable[n] = e;
                    }
                } catch { /* 同上：取不到就少展開一層，Layer B 仍在 */ }

                return d;
            }

            /// <summary>
            /// 把查詢引用到的衍生定義接在文字後面，讓 Layer A 掃得到藏在下游定義裡的欄位。
            /// 不展開的話，「建一個新量值／新資料行去複製敏感欄位」就能整個繞過管制。
            /// </summary>
            public static string ExpandDerivations(string dax, Derivations d) {
                var sb   = new System.Text.StringBuilder(dax ?? "");
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int depth = 0; depth < 6; depth++) {   // 衍生可以套衍生，但不無限展開
                    bool added = false;
                    string cur = sb.ToString();

                    foreach (var (token, _) in ScanBracketTokens(cur)) {
                        if (!seen.Add(token)) continue;
                        if (d.ByToken.TryGetValue(token, out var expr)) { sb.Append('\n').Append(expr); added = true; }
                    }
                    // 表名沒有中括號可循，只能看查詢文字有沒有提到它。
                    // 寧可過度納入 —— 多接一段定義只會讓掃描更嚴，不會放水。
                    foreach (var kv in d.ByTable) {
                        if (seen.Contains("\t" + kv.Key)) continue;
                        if (cur.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) < 0) continue;
                        seen.Add("\t" + kv.Key);
                        sb.Append('\n').Append(kv.Value);
                        added = true;
                    }
                    if (!added) break;
                }
                return sb.ToString();
            }

            /// <summary>Layer A：執行前掃描查詢文字。dax 應已展開量值。null 表示放行。</summary>
            public static string? CheckQuery(string dax, out bool touchesMoney, bool pseudonymize = false) {
                touchesMoney = false;
                if (!Enabled) return null;
                foreach (var (token, enclosing) in ScanBracketTokens(dax)) {
                    if (Matches(token, AllowColumns)) continue;      // 已明確標記為安全
                    if (Matches(token, DenyColumns) && !enclosing.Any(f => CountingFuncs.Contains(f))) {
                        // 去敏模式：只當分組鍵用時放行，值會在輸出前換成代號。
                        if (pseudonymize && enclosing.All(f => GroupingShapes.Contains(f))) continue;
                        return $"[{token}] 是客戶身分欄位，不能當分組鍵或直接取值。"
                             + "只能用計數類函式取統計量，例如 DISTINCTCOUNT / COUNTROWS / COUNTBLANK。"
                             + "（MAX / MIN / CONCATENATEX 不算 —— 它們會回傳真實名稱。）";
                    }
                    if (Matches(token, MoneyColumns)) {
                        touchesMoney = true;
                        if (!enclosing.Any(f => NumericAggFuncs.Contains(f) || CountingFuncs.Contains(f)))
                            return $"[{token}] 是金額欄位，必須包在聚合函式內（SUM / AVERAGE / SUMX …），不能逐列取值。";
                    }
                }
                return null;
            }

            /// <summary>Layer B：執行後、序列化前掃描結果欄位名。null 表示放行。</summary>
            public static string? CheckResultColumns(IEnumerable<string> columns, bool pseudonymize = false) {
                if (!Enabled) return null;
                foreach (var col in columns) {
                    string bare = col ?? "";
                    int lb = bare.LastIndexOf('[');
                    if (lb >= 0) bare = bare.Substring(lb + 1);
                    bare = bare.TrimEnd(']').Trim();

                    if (Matches(bare, AllowColumns)) continue;       // 已明確標記為安全
                    if (!pseudonymize && Matches(bare, DenyColumns))  // 去敏模式改成遮罩，不擋
                        return $"查詢結果含客戶身分欄位 [{bare}]。整表 EVALUATE 會把原始欄位整批帶出來。";
                    // 結果欄名若是 'Table'[money] 而非別名，代表該金額欄根本沒被聚合，是逐列輸出
                    if (Matches(bare, MoneyColumns))
                        return $"查詢結果含未經聚合的金額欄位 [{bare}]。請改以 SUM/AVERAGE 等聚合後再取別名。";
                }
                return null;
            }

            // ── 一次性放行權杖 ────────────────────────────────────────────────
            // 被攔截時才產生，並且只印在服務主控台。AI 看不到那個視窗，只能開口向
            // 使用者要 —— 「取得使用者同意」因此變成機制，而不是模型的自我承諾。
            // 權杖綁定當次查詢文字：不能拿一次同意去授權另一句更寬的查詢。
            static readonly Dictionary<string, (string QueryHash, DateTime Expires)> _tokens = new();

            static string HashQuery(string q) {
                using var sha = System.Security.Cryptography.SHA256.Create();
                string norm = Regex.Replace(q ?? "", @"\s+", " ").Trim();
                return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(norm)));
            }

            public static string IssueToken(string query) {
                string tok = Convert.ToHexString(
                    System.Security.Cryptography.RandomNumberGenerator.GetBytes(4));
                lock (_tokens) {
                    foreach (var k in _tokens.Where(kv => kv.Value.Expires < DateTime.UtcNow)
                                             .Select(kv => kv.Key).ToList())
                        _tokens.Remove(k);
                    _tokens[tok] = (HashQuery(query), DateTime.UtcNow.AddMinutes(10));
                }
                return tok;
            }

            /// <summary>驗證並「用掉」權杖。必須是同一句查詢，否則等於拿舊同意授權新查詢。</summary>
            public static bool RedeemToken(string? token, string query) {
                if (string.IsNullOrWhiteSpace(token)) return false;
                lock (_tokens) {
                    if (!_tokens.TryGetValue(token!, out var e)) return false;
                    _tokens.Remove(token!);              // 一次性：驗證成功與否都作廢
                    return e.Expires >= DateTime.UtcNow && e.QueryHash == HashQuery(query);
                }
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

        static void Main(string[] args) {
            // ✅ 強制將工作目錄設為 DLL 所在位置，防止捷徑啟動時找不到 appsettings.json
            System.IO.Directory.SetCurrentDirectory(System.AppContext.BaseDirectory);

            Console.WriteLine("======================================");
            Console.WriteLine("🚀 PBI Visual Explorer - 背景伺服器啟動中...");
            Console.WriteLine("======================================");

            var builder = WebApplication.CreateBuilder(args);

            // ✅ 修補①：讀取 appsettings.json，集中管理路徑與密鑰
            builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: false);
            var config          = builder.Configuration;
            var apiKey = config["Security:ApiKey"] ?? throw new Exception("❌ 缺少 Security:ApiKey");

            // ⚠️ 這裡刻意不再有「目標 PBIP 檔案」的設定。
            //    這是通用工具，使用者會頻繁切換不同的 PBIX/PBIP —— 寫死一個路徑只會過期。
            //    所有路徑都從執行中的 PBIDesktop 行程推導（見 DiscoverInstances）。
            //    AllowedBasePaths 是選用的額外白名單，設了就會在推導出的路徑之外再檢查一層。
            var extraAllowedPaths = config.GetSection("Security:AllowedBasePaths").Get<string[]>() ?? Array.Empty<string>();

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
            DataGuard.InitPseudonymKey(config["Security:ApiKey"]);

            var auditDir = Path.Combine(projectRoot, "audit");
            Directory.CreateDirectory(auditDir);
            DataGuard.AuditPath = Path.Combine(auditDir, $"query-audit-{DateTime.Now:yyyyMM}.tsv");

            var allowedOrigins  = config.GetSection("Security:AllowedOrigins").Get<string[]>()
                                  ?? new[] { "http://localhost:5500", "null" };

            Console.WriteLine($"🔐 安全模式啟用 — Key: {apiKey[..8]}...");
            Console.WriteLine($"💾 快照資料夾: {snapshotPath}");
            if (DataGuard.Enabled) {
                Console.WriteLine($"🛡️ 資料保護啟用 — 身分欄位／樣式 {DataGuard.DenyColumns.Length} 條、"
                                + $"金額欄位／樣式 {DataGuard.MoneyColumns.Length} 條、"
                                + $"安全白名單 {DataGuard.AllowColumns.Length} 條");
                Console.WriteLine($"   回傳列數上限：逐列明細 {DataGuard.MaxDetailRows} 列、"
                                + $"彙總結果 {DataGuard.MaxAggregateRows} 列、"
                                + $"金額分組 {DataGuard.MaxRowsWithMoney} 列"
                                + "（伺服器強制，呼叫端只能更低）");
                Console.WriteLine($"📋 查詢稽核記錄: {DataGuard.AuditPath}");
            } else {
                Console.WriteLine("⚠️ 資料保護「未啟用」 — 查詢可取出任何欄位內容。"
                                + "若非本意，請檢查 appsettings.json 的 DataProtection:Enabled");
            }
            if (extraAllowedPaths.Length > 0)
                Console.WriteLine($"🛡️ 額外路徑白名單: {string.Join(" | ", extraAllowedPaths)}");

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
                        Console.WriteLine("   ⚠️ 有多個實例 —— 請用 X-PBI-Target 標頭指定目標，否則寫入會被拒絕");
                }
            } catch (Exception ex) {
                Console.WriteLine($"⚠️ 實例探索失敗: {ex.Message}");
            }

            // ✅ 修補②：CORS 限制為白名單（不再 AllowAnyOrigin）
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

            // ✅ 修補③：全域 API Key 驗證（所有 /api/* 都需帶 X-API-Key Header）
            app.Use(async (context, next) => {
                // 放行 ping 測試與瀏覽器的 CORS 預檢請求 (OPTIONS 不帶自訂標頭)
                if (context.Request.Path == "/ping" || context.Request.Method == "OPTIONS") {
                    await next();
                    return;
                }
                if (!context.Request.Headers.TryGetValue("X-API-Key", out var key) || key != apiKey) {
                    context.Response.StatusCode = 401;
                    await context.Response.WriteAsync("❌ 401 Unauthorized");
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 未授權存取被攔截！");
                    return;
                }
                await next();
            });
            app.MapGet("/ping", () => "pong");

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
                            ? "有多個實例，請用 X-PBI-Target 標頭指定目標（Port 或檔名片段），否則寫入類端點會拒絕執行。"
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
                        FileExists = inst.FilePath != null && File.Exists(inst.FilePath),
                        Note = "/api/inject-visual 需要 PBIP 格式（展開成資料夾的專案）。Kind 是 PBIX 則不適用。"
                    });
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            // 模型健檢
            app.MapGet("/api/validate", (HttpContext ctx) => RunModel(ctx, "模型健檢", (model, _) => ValidateModel(model), save: false));

            // =====================================================================
            // 查詢
            // =====================================================================

            // 共用的 ADOMD 執行邏輯（/api/query 與 /api/dmv 都用它）
            IResult RunAdomd(HttpContext ctx, string query, int? maxRowsIn, int? timeoutIn, string label,
                             bool isDax = false, bool guardBypassed = false, bool pseudonymize = false) {
                try {
                    var inst = ResolveInstance(ctx);
                    int maxRows = Math.Clamp(maxRowsIn ?? 1000, 1, 10000);
                    int timeout = Math.Clamp(timeoutIn ?? 60, 1, 600);
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] 🔎 {label} → {inst.FileName}（上限 {maxRows} 列 / {timeout} 秒）...");
                    var sw = Stopwatch.StartNew();

                    // 攔截時的共用出口：記稽核、產生一次性權杖、把權杖印在主控台（只有使用者看得到）。
                    IResult Blocked(string verdict, string reason, string hint, int rows) {
                        DataGuard.Audit(verdict, reason, query, rows, inst.FileName);
                        string tok = DataGuard.IssueToken(query);
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⛔ 資料保護攔截：{reason}");
                        Console.WriteLine("   ┌─────────────────────────────────────────────────────────");
                        Console.WriteLine($"   │ 若你確認要放行「這一句」查詢，把下列權杖貼給 AI：");
                        Console.WriteLine($"   │     {tok}      （10 分鐘內有效，只能用一次）");
                        Console.WriteLine("   │ 它只授權上面這一句查詢，換一句就無效。不確定就別給。");
                        Console.WriteLine("   └─────────────────────────────────────────────────────────");
                        return Results.Json(new {
                            Error = "⛔ 已被資料保護機制攔截（伺服器端）", Reason = reason, Hint = hint,
                            HowToOverride = "服務主控台已印出一組一次性權杖。請使用者確認後把權杖貼給你，"
                                          + "再以 Invoke-Dax -DetailToken <權杖> 重送同一句查詢。AI 無法自行取得權杖。"
                        }, statusCode: 403);
                    }

                    using (var conn = new AdomdConnection($"Data Source=localhost:{inst.Port};")) {
                        conn.Open();

                        // ── Layer A：執行前掃描查詢文字 ──────────────────────────────
                        // 放在 conn.Open() 之後，是因為展開量值定義需要連線去讀 TMSCHEMA_MEASURES。
                        // 敏感欄位可以藏在量值裡，不展開就等於沒檢查。
                        bool touchesMoney = false;
                        if (DataGuard.Enabled && !guardBypassed) {
                            string scanText = query;
                            int measureCount = 0;
                            if (isDax) {
                                var derived = DataGuard.LoadDerivations(conn);
                                measureCount = derived.Count;
                                scanText = DataGuard.ExpandDerivations(query, derived);
                            }
                            var violation = DataGuard.CheckQuery(scanText, out touchesMoney, pseudonymize);
                            if (violation != null)
                                return Blocked("BLOCKED-A", $"{violation}（已展開 {measureCount} 筆量值定義）",
                                    "欄位名稱可以自由讀取，被擋的是欄位內容。改用彙總寫法即可，"
                                  + "例如 EVALUATE ROW(\"客戶數\", DISTINCTCOUNT(<表>[<欄>]))。", 0);

                            // ── 列數上限：由伺服器決定，不由呼叫端決定 ─────────────────
                            // MaxRows 原本是請求參數 —— 那不是護欄，只是預設值：
                            // 呼叫端（也就是 AI）可以自己填 10000。這裡改成只能往下收。
                            //   逐列明細 → MaxDetailRows
                            //   彙總結果 → MaxAggregateRows（統計量不是資料細節）
                            if (isDax) {
                                bool isAgg = DataGuard.IsAggregateQuery(scanText);
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
                                var colNames = new List<string>();
                                var seen = new Dictionary<string, int>();
                                for (int i = 0; i < reader.FieldCount; i++) {
                                    string n = reader.GetName(i);
                                    if (string.IsNullOrEmpty(n)) n = $"Column{i + 1}";
                                    if (seen.TryGetValue(n, out int c)) { seen[n] = c + 1; n = $"{n}_{c + 1}"; }
                                    else seen[n] = 0;
                                    colNames.Add(n);
                                }

                                // ── Layer B：執行後、序列化前掃描結果欄位名 ──────────────
                                // 補 Layer A 的盲點：整表 EVALUATE 的查詢文字裡沒有任何欄名，
                                // 但引擎回報的結果欄名是 'Table'[Column]，在這裡才抓得到。
                                // 注意順序 —— 這道檢查在 reader.Read() 之前，值連讀都沒讀進來。
                                if (DataGuard.Enabled && !guardBypassed) {
                                    var colViolation = DataGuard.CheckResultColumns(colNames, pseudonymize);
                                    if (colViolation != null)
                                        return Blocked("BLOCKED-B", colViolation,
                                            "請改以彙總方式取得所需資訊，不要整表取出。", 0);
                                }

                                // 去敏：哪幾個欄位的值要換成代號。對照表只印在主控台，不回傳。
                                var maskIdx   = (pseudonymize && DataGuard.Enabled && !guardBypassed)
                                              ? DataGuard.MaskIndexes(colNames) : new HashSet<int>();
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
                                        }
                                        row[colNames[i]] = val;
                                    }
                                    rows.Add(row);
                                }
                                sw.Stop();

                                // ── 金額查詢的列數上限 ────────────────────────────────
                                // 後備防線：金額有被聚合（過了 Layer A），但分組鍵的基數太高。
                                // 例如按 material_no 分組 → 7,927 列的單一料號營收，
                                // 那已經接近逐筆金額，不是 KPI 彙總了。
                                if (DataGuard.Enabled && !guardBypassed && touchesMoney &&
                                    rows.Count > DataGuard.MaxRowsWithMoney)
                                    return Blocked("BLOCKED-ROWS",
                                        $"金額查詢回傳 {rows.Count} 列，超過上限 {DataGuard.MaxRowsWithMoney} 列。",
                                        "請改用基數較低的分組鍵（產品線、月份、廠別），或先篩選再彙總。", rows.Count);

                                DataGuard.Audit("OK", touchesMoney ? "含金額欄位" : "", query, rows.Count, inst.FileName);

                                Console.WriteLine($"✅ 查詢完成：{rows.Count} 列 / {sw.ElapsedMilliseconds} ms" +
                                                  (truncated ? $"（已截斷至 {maxRows} 列）" : ""));
                                if (maskIdx.Count > 0) {
                                    Console.WriteLine($"   🎭 已去敏 {maskIdx.Count} 個欄位／{pseudoMap.Count} 個相異值"
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
                                    Rows      = rows
                                });
                            }
                        }
                    }
                } catch (AdomdErrorResponseException ex) {
                    // DAX 語法或執行期錯誤：回 400 並附上引擎原始訊息，方便直接修正公式
                    Console.WriteLine($"❌ 查詢錯誤: {ex.Message}");
                    return Results.BadRequest($"❌ 查詢錯誤: {ex.Message}");
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    Console.WriteLine($"❌ 查詢失敗: {ex.Message}");
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            }

            // 唯讀 DAX 查詢（驗算量值結果用）
            app.MapPost("/api/query", (DaxQueryRequest req, HttpContext ctx) => {
                if (string.IsNullOrWhiteSpace(req.Query)) return Results.BadRequest("❌ Query 不可為空");
                // ✅ 唯讀防護：只接受 DAX 查詢語法，擋掉 XMLA 命令等會變更模型的內容
                var head = StripLeadingComments(req.Query);
                if (!head.StartsWith("EVALUATE", StringComparison.OrdinalIgnoreCase) &&
                    !head.StartsWith("DEFINE",   StringComparison.OrdinalIgnoreCase)) {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 非唯讀查詢被擋下");
                    return Results.BadRequest("❌ 只接受唯讀查詢：必須以 EVALUATE 或 DEFINE 開頭");
                }

                // 一次性放行權杖。權杖只會出現在服務主控台，AI 拿不到 —— 這是刻意的：
                // 它讓「取得使用者同意」成為機制，而不是模型自己說了算的開關。
                string? tok = ctx.Request.Headers.TryGetValue("X-PBI-Allow-Detail", out var tv)
                              ? tv.ToString() : null;
                bool bypass = DataGuard.RedeemToken(tok, req.Query);
                if (!string.IsNullOrWhiteSpace(tok) && !bypass) {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 放行權杖無效（過期、已用過，或查詢已被改動）");
                    DataGuard.Audit("TOKEN-REJECTED", "權杖無效或與查詢不符", req.Query, 0, null);
                    return Results.Json(new {
                        Error = "⛔ 放行權杖無效",
                        Reason = "權杖可能已過期（10 分鐘）、已被使用過，或查詢文字與當初被攔截的那一句不同。",
                        Hint = "權杖綁定單一查詢。若要放行新的查詢，請重送一次讓伺服器產生新權杖。"
                    }, statusCode: 403);
                }
                if (bypass) {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] 🔓 使用者以一次性權杖放行本次查詢");
                    DataGuard.Audit("USER-OVERRIDE", "使用者以一次性權杖放行", req.Query, 0, null);
                }

                return RunAdomd(ctx, req.Query, req.MaxRows, req.TimeoutSeconds, "執行 DAX 查詢",
                                isDax: true, guardBypassed: bypass,
                                pseudonymize: req.Pseudonymize == true);
            });

            // DMV 查詢：$SYSTEM 系統檢視（VertiPaq 儲存統計、相依性追蹤等）
            // 這些檢視只含中繼資料與統計，不會回傳事實資料列
            app.MapPost("/api/dmv", (DmvQueryRequest req, HttpContext ctx) => {
                if (string.IsNullOrWhiteSpace(req.Query)) return Results.BadRequest("❌ Query 不可為空");
                var head = StripLeadingComments(req.Query);
                if (!Regex.IsMatch(head, @"^SELECT\b[\s\S]*\bFROM\s+\$SYSTEM\.[A-Za-z0-9_]+",
                                   RegexOptions.IgnoreCase)) {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ 非 $SYSTEM DMV 查詢被擋下");
                    return Results.BadRequest("❌ 只接受 $SYSTEM DMV 查詢，格式須為：SELECT ... FROM $SYSTEM.<檢視名稱>");
                }

                // 這幾個檢視會回傳「定義文字」（M 腳本、共用運算式、相依性運算式）。
                // 欄位層級管制對它們無效 —— 管制比對的是欄位名稱，而這裡敏感內容是
                // 藏在 QueryDefinition / Expression 這種無害欄名底下的字串。
                // M 腳本裡很可能有寫死的客戶名稱與連線字串，所以整個檢視擋掉。
                if (DataGuard.Enabled &&
                    Regex.IsMatch(head, @"\$SYSTEM\.(TMSCHEMA_PARTITIONS|TMSCHEMA_EXPRESSIONS|DISCOVER_CALC_DEPENDENCY)\b",
                                  RegexOptions.IgnoreCase)) {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⛔ 資料保護攔截：DMV 檢視會回傳 M／運算式定義文字");
                    DataGuard.Audit("BLOCKED-DMV", "檢視會回傳定義文字（M 腳本／運算式）", req.Query, 0, null);
                    return Results.Json(new {
                        Error  = "⛔ 已被資料保護機制攔截（DMV 檢視）",
                        Reason = "TMSCHEMA_PARTITIONS / TMSCHEMA_EXPRESSIONS / DISCOVER_CALC_DEPENDENCY "
                               + "會回傳 M 腳本與運算式原文，其中可能含寫死的客戶名稱與連線字串。",
                        Hint   = "要看模型結構請用 /api/schema、/api/validate；"
                               + "要看單張表的 M 請用 Get-PbiMQuery（預設只存檔不回傳內容）。"
                    }, statusCode: 403);
                }
                return RunAdomd(ctx, req.Query, req.MaxRows, req.TimeoutSeconds, "執行 DMV 查詢");
            });

            // =====================================================================
            // 存檔 / 重新整理
            // =====================================================================

            // 送出 Ctrl+S 給 PBI Desktop，並用檔案修改時間驗證是否真的存檔成功。
            // ⚠️ 會搶走前景視窗焦點，這是 SendKeys 的固有限制。
            app.MapPost("/api/save", async (SaveRequest? req, HttpContext ctx) => {
                try {
                    var inst = ResolveInstance(ctx);
                    int wait = Math.Clamp(req?.WaitSeconds ?? 5, 1, 120);
                    Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] 💾 存檔：{inst.FileName}（等待 {wait} 秒驗證）...");

                    var target = inst.FilePath;
                    if (string.IsNullOrEmpty(target))
                        return Results.BadRequest("❌ 這個實例沒有對應的檔案路徑（可能是未儲存過的新報表），無法驗證存檔結果。");

                    var (before, problem) = GetLatestWriteTime(target);
                    string result = await SendCtrlS(inst.PbiPid);

                    if (result == "not_running")
                        return Results.BadRequest("❌ 該 PBI Desktop 行程已結束或沒有可用視窗，無法存檔");
                    if (result == "activate_failed")
                        return Results.BadRequest("❌ 無法將 PBI Desktop 帶到前景，Ctrl+S 未送出。請確認視窗未最小化到系統匣。");

                    await Task.Delay(TimeSpan.FromSeconds(wait));
                    var (after, _) = GetLatestWriteTime(target);

                    bool canVerify = problem == null;
                    bool changed   = canVerify && before.HasValue && after.HasValue && after > before;

                    string msg;
                    if (!canVerify) {
                        msg = $"⚠️ Ctrl+S 已送出，但無法驗證是否存檔成功：{problem}";
                        Console.WriteLine(msg);
                    } else if (changed) {
                        msg = "存檔成功";
                        Console.WriteLine("✅ 已偵測到檔案更新，存檔成功");
                    } else {
                        msg = "Ctrl+S 已送出，但檔案時間未變動 —— 可能本來就沒有待存變更，或存檔尚未完成";
                        Console.WriteLine("⚠️ " + msg);
                    }

                    return Results.Ok(new {
                        message               = msg,
                        FileChanged           = changed,
                        TargetFile            = target,
                        VerificationAvailable = canVerify,
                        VerificationProblem   = problem,
                        BeforeUtc             = before,
                        AfterUtc              = after,
                        WaitedSeconds         = wait
                    });
                } catch (OpException ex) {
                    return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message);
                } catch (Exception ex) {
                    Console.WriteLine($"❌ 存檔失敗: {ex.Message}");
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            // 重新整理資料。改完 M 腳本後必須呼叫這個，資料才會真的更新。
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
                    var sw = Stopwatch.StartNew();
                    if (string.IsNullOrWhiteSpace(req?.TableName)) {
                        model.RequestRefresh(rt);
                    } else {
                        var t = FindTable(model, req!.TableName);
                        t.RequestRefresh(rt);
                    }
                    // SaveChanges 之後才會真的送出重新整理，這裡先記錄準備時間
                    return new { message = label + " 已排入", PreparedMs = sw.ElapsedMilliseconds };
                }, save: true);
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
                string key;
                if (string.IsNullOrEmpty(inst.FilePath)) {
                    key = $"unsaved_port{inst.Port}";
                } else {
                    var name = Regex.Replace(Path.GetFileNameWithoutExtension(inst.FilePath),
                                             @"[\\/:*?""<>|]", "_");
                    using var sha = System.Security.Cryptography.SHA256.Create();
                    var hash = Convert.ToHexString(
                        sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(inst.FilePath.ToLowerInvariant())))[..8];
                    key = $"{name}_{hash}";
                }
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
            app.MapPost("/api/restore", (RestoreRequest req, HttpContext ctx) => {
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
            // 寫入（單一操作）—— 全部走 DispatchOp，與 /api/batch 共用實作
            // =====================================================================

            // /api/update-m 已於 2026-08-06 移除 —— M 腳本唯讀。
            // 保留這個端點只為了回一句看得懂的話，而不是 404 讓人以為是版本不對。
            app.MapPost("/api/update-m", () => Results.Json(new {
                error = "M 腳本唯讀：/api/update-m 已移除",
                why   = "用 TOM 改 M 只動到模型，Power BI Desktop 的 Power Query 文件不會跟著變，" +
                        "兩邊失步後 Desktop 會永久顯示「查詢中有暫止的變更尚未套用」，按套用也解不開。",
                how   = "把完整的 let...in 交給使用者，請他在 Power Query 編輯器的進階編輯器貼上，再「關閉並套用」。",
                read  = "要讀 M 請用 GET /api/schema 的 MQuery 欄位（PowerShell：Get-PbiMQuery <表名>）"
            }, statusCode: 410));

            app.MapPost("/api/upsert-measure", (UpsertMeasureRequest req, HttpContext ctx) =>
                RunModel(ctx, $"寫入量值 '{req.TableName}'[{req.MeasureName}]", (m, _) => new { message = OpUpsertMeasure(m, req) }, save: true));

            app.MapPost("/api/delete-measure", (DeleteMeasureRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除量值 '{req.TableName}'[{req.MeasureName}]", (m, _) => new { message = OpDeleteMeasure(m, req) }, save: true));

            app.MapPost("/api/move-measure", (MoveMeasureRequest req, HttpContext ctx) =>
                RunModel(ctx, $"搬移量值 [{req.MeasureName}] → '{req.ToTable}'", (m, _) => new { message = OpMoveMeasure(m, req) }, save: true));

            app.MapPost("/api/add-column", (AddColumnRequest req, HttpContext ctx) =>
                RunModel(ctx, $"新增計算資料行 '{req.TableName}'[{req.ColumnName}]", (m, _) => new { message = OpAddColumn(m, req) }, save: true));

            app.MapPost("/api/upsert-relationship", (RelationshipRequest req, HttpContext ctx) =>
                RunModel(ctx, "建立/更新關聯", (m, _) => new { message = OpUpsertRelationship(m, req) }, save: true));

            app.MapPost("/api/delete-relationship", (RelationshipRefRequest req, HttpContext ctx) =>
                RunModel(ctx, "刪除關聯", (m, _) => new { message = OpDeleteRelationship(m, req) }, save: true));

            app.MapPost("/api/create-table", (CreateTableRequest req, HttpContext ctx) =>
                RunModel(ctx, $"建立表格 '{req.TableName}'", (m, _) => new { message = OpCreateTable(m, req) }, save: true));

            app.MapPost("/api/delete-table", (TableRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除表格 '{req.TableName}'", (m, _) => new { message = OpDeleteTable(m, req) }, save: true));

            app.MapPost("/api/delete-column", (ColumnRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除資料行 '{req.TableName}'[{req.ColumnName}]", (m, _) => new { message = OpDeleteColumn(m, req) }, save: true));

            app.MapPost("/api/set-column-props", (ColumnPropsRequest req, HttpContext ctx) =>
                RunModel(ctx, $"設定資料行屬性 '{req.TableName}'[{req.ColumnName}]", (m, _) => new { message = OpSetColumnProps(m, req) }, save: true));

            app.MapPost("/api/rename", (RenameRequest req, HttpContext ctx) =>
                RunModel(ctx, $"改名 {req.ObjectType} [{req.OldName}] → [{req.NewName}]",
                         (m, _) => OpRename(m, req), save: !(req.DryRun ?? false)));

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

            app.MapPost("/api/upsert-expression", (ExpressionRequest req, HttpContext ctx) =>
                RunModel(ctx, $"寫入共用運算式 [{req.Name}]", (m, _) => new { message = OpUpsertExpression(m, req) }, save: true));

            app.MapPost("/api/delete-expression", (ExpressionRefRequest req, HttpContext ctx) =>
                RunModel(ctx, $"刪除共用運算式 [{req.Name}]", (m, _) => new { message = OpDeleteExpression(m, req) }, save: true));

            // =====================================================================
            // 批次：一次連線、一次 SaveChanges，避免 N 次模型重算
            // =====================================================================
            app.MapPost("/api/batch", (BatchRequest req, HttpContext ctx) => {
                if (req.Operations == null || req.Operations.Length == 0)
                    return Results.BadRequest("❌ Operations 不可為空");
                if (req.Operations.Length > 200)
                    return Results.BadRequest("❌ 單次批次上限 200 個操作");

                bool stopOnError = req.StopOnError ?? true;
                bool savePerOp   = req.SavePerOp   ?? false;
                bool dryRun      = req.DryRun      ?? false;

                return RunModel(ctx, $"批次執行 {req.Operations.Length} 個操作" +
                                (savePerOp ? "（逐步存檔）" : "（最後一次存檔）"), (model, _) => {
                    var results = new List<object>();
                    int ok = 0, failed = 0;

                    for (int i = 0; i < req.Operations.Length; i++) {
                        var opDef = req.Operations[i];
                        try {
                            var r = DispatchOp(model, opDef.Op, opDef.Args);
                            if (savePerOp && !dryRun) model.SaveChanges();
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
                }, save: !dryRun && !savePerOp);
            });

            // =====================================================================
            // API: 視覺層注入 → 自動重啟 PBI Desktop
            // =====================================================================
            app.MapPost("/api/inject-visual", async (InjectVisualRequest req, HttpContext ctx) => {
                try {
                    // 目標實例決定一切：要改哪個報表資料夾、要砍哪個行程、要重開哪個檔案。
                    // 全部從執行中的行程推導，不再依賴 appsettings.json（設定會過期，行程不會）。
                    PbiInstance inst;
                    try { inst = ResolveInstance(ctx); }
                    catch (OpException ex) { return ex.StatusCode == 404 ? Results.NotFound(ex.Message) : Results.BadRequest(ex.Message); }

                    Console.WriteLine($"\n[{DateTime.Now:HH:mm:ss}] 視覺層注入：{inst.FileName} 頁面 [{req.PageDisplayName}] 視覺 [{req.VisualId}]");

                    // ✅ 修補④-A：PageId/VisualId 只允許 20 位 hex（防止路徑注入）
                    if (!Regex.IsMatch(req.PageId,   @"^[a-f0-9]{20}$") ||
                        !Regex.IsMatch(req.VisualId, @"^[a-f0-9]{20}$"))
                        return Results.BadRequest("❌ 非法的 PageId/VisualId（必須為 20 位 hex）");

                    // ✅ 修補④-B：驗證 VisualJson 必須是合法 JSON
                    try { System.Text.Json.JsonDocument.Parse(req.VisualJson); }
                    catch { return Results.BadRequest("❌ VisualJson 格式不合法"); }

                    // ✅ 修補④-C：視覺注入只適用 PBIP（展開成資料夾的專案）
                    if (inst.Kind != "PBIP" || string.IsNullOrEmpty(inst.FilePath))
                        return Results.BadRequest(
                            $"❌ 目標開啟的是 {inst.Kind}，視覺層注入需要 PBIP 格式（展開成資料夾的專案）。\n" +
                            "   請改用 PBIP 開啟，或改用 Power BI Desktop 介面手動調整視覺。");

                    // ✅ 修補④-D：報表定義資料夾由開啟中的檔案推導
                    //    X:\proj\name.pbip  →  X:\proj\name.Report\definition
                    var projectDir  = Path.GetDirectoryName(Path.GetFullPath(inst.FilePath))!;
                    var reportPath  = Path.Combine(projectDir,
                                          Path.GetFileNameWithoutExtension(inst.FilePath) + ".Report", "definition");
                    if (!Directory.Exists(reportPath))
                        return Results.BadRequest($"❌ 找不到報表定義資料夾：{reportPath}");

                    // ✅ 修補④-E：目標路徑必須落在專案資料夾內（防路徑穿越）
                    var resolvedVisDir = Path.GetFullPath(
                        Path.Combine(reportPath, "pages", req.PageId, "visuals", req.VisualId));
                    if (!resolvedVisDir.StartsWith(projectDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) {
                        Console.WriteLine($"⚠️ 路徑穿越攻擊被攔截！目標：{resolvedVisDir}");
                        return Results.BadRequest("❌ 存取路徑超出該 PBIP 專案資料夾範圍");
                    }
                    // 選用的額外白名單：設了就必須同時符合
                    if (extraAllowedPaths.Length > 0 &&
                        !extraAllowedPaths.Any(b => resolvedVisDir.StartsWith(Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase))) {
                        Console.WriteLine($"⚠️ 目標不在 Security:AllowedBasePaths 白名單內：{resolvedVisDir}");
                        return Results.BadRequest("❌ 存取路徑不在 appsettings.json 的 AllowedBasePaths 白名單內");
                    }

                    string pbipFilePath = inst.FilePath;
                    var pbipReportPath  = reportPath;

                    // ── 步驟 1：對「目標那一個」PBI Desktop 送 Ctrl+S ──────────────────
                    var pbiProcs = new[] { inst.PbiPid }
                        .Select(pid => { try { return Process.GetProcessById(pid); } catch { return null; } })
                        .Where(p => p != null).Select(p => p!).ToArray();
                    if (pbiProcs.Length > 0) {
                        Console.WriteLine("[1/5] 模擬 Ctrl+S 存檔中...");
                        var saveResult = await SendCtrlS(inst.PbiPid);
                        await Task.Delay(3000);
                        Console.WriteLine($"[1/5] ✅ 存檔指令已送出（{saveResult}）");

                        // ── 步驟 2：只關閉這一個實例，不動其他開著的 PBI ──────────────
                        Console.WriteLine($"[2/5] 關閉 PBI Desktop (PID {inst.PbiPid})...");
                        foreach (var p in pbiProcs) {
                            try { p.Kill(); } catch { }
                        }
                        await Task.Delay(2000); // 等待關閉完成
                        Console.WriteLine("[2/5] ✅ PBI Desktop 已關閉");
                    } else {
                        Console.WriteLine("[1/5] PBI Desktop 未開啟，跳過存檔步驟");
                    }

                    // ── 步驟 3：寫入 visual.json（路徑由 config 決定）────────────────────
                    Console.WriteLine("[3/5] 寫入視覺物件 JSON 檔案...");
                    Directory.CreateDirectory(resolvedVisDir);
                    await File.WriteAllTextAsync(Path.Combine(resolvedVisDir, "visual.json"), req.VisualJson);

                    // ── 步驟 3b：如果是新頁面，建立 page.json 並更新 pages.json ─────────
                    if (req.IsNewPage) {
                        var pageDir = Path.Combine(pbipReportPath, "pages", req.PageId);
                        Directory.CreateDirectory(pageDir);
                        var pageJson =
                            "{\r\n" +
                            "  \"$schema\": \"https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json\",\r\n" +
                            $"  \"name\": \"{req.PageId}\",\r\n" +
                            $"  \"displayName\": \"{req.PageDisplayName}\",\r\n" +
                            "  \"displayOption\": \"FitToPage\",\r\n" +
                            "  \"height\": 720,\r\n" +
                            "  \"width\": 1280\r\n" +
                            "}";
                        await File.WriteAllTextAsync(Path.Combine(pageDir, "page.json"), pageJson);

                        // 更新 pages.json，把新頁面加到最前面
                        var pagesJsonPath = Path.Combine(pbipReportPath, "pages", "pages.json");
                        var pagesContent  = await File.ReadAllTextAsync(pagesJsonPath);
                        var pagesObj      = System.Text.Json.JsonDocument.Parse(pagesContent);
                        var existingOrder = pagesObj.RootElement.GetProperty("pageOrder")
                                               .EnumerateArray()
                                               .Select(e => e.GetString()!)
                                               .ToList();
                        if (!existingOrder.Contains(req.PageId)) {
                            existingOrder.Insert(0, req.PageId);
                            var newPagesJson = System.Text.Json.JsonSerializer.Serialize(
                                new System.Collections.Generic.Dictionary<string, object> {
                                    ["$schema"]        = pagesObj.RootElement.GetProperty("$schema").GetString()!,
                                    ["pageOrder"]      = existingOrder,
                                    ["activePageName"] = req.PageId
                                },
                                new System.Text.Json.JsonSerializerOptions { WriteIndented = true }
                            );
                            await File.WriteAllTextAsync(pagesJsonPath, newPagesJson);
                            Console.WriteLine($"[3/5] ✅ pages.json 已更新，新頁面 [{req.PageDisplayName}] 已加入");
                        }
                    }
                    Console.WriteLine("[3/5] ✅ 視覺物件 JSON 寫入完畢");

                    // ── 步驟 4：重新開啟 PBIP（路徑由 config 決定，不接受外部傳入）────────
                    Console.WriteLine($"[4/5] 重新開啟 PBI Desktop：{pbipFilePath}");
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo {
                        FileName        = pbipFilePath,  // ✅ 使用 config 路徑
                        UseShellExecute = true
                    });
                    Console.WriteLine("[4/5] ✅ PBI Desktop 已重新啟動，等待載入...");

                    Console.WriteLine("[5/5] 🎉 視覺層注入完成！約 10~15 秒後 PBI 會顯示最新內容。");
                    return Results.Ok(new {
                        message = $"Visual [{req.VisualId}] injected to page [{req.PageDisplayName}]. PBI is restarting."
                    });

                } catch (Exception ex) {
                    Console.WriteLine($"❌ 視覺層注入失敗: {ex.Message}");
                    return Results.Problem(detail: ex.Message, statusCode: 500);
                }
            });

            Console.WriteLine("✅ 伺服器已在背景運行 (等待網頁於 Port 5500 呼叫)！");
            app.Run();
        }
    }
}
