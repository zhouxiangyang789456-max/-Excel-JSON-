using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Editor.Generator;
using ExcelToJsonPlugin.Editor.Mapping;
using ExcelToJsonPlugin.Editor.Validator;
using ExcelToJsonPlugin.Editor.Validator.Rules;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// 核心流水线调度器。
    /// 串联 Reader → Parser → Validator → Generator → Output。
    /// </summary>
    public class Pipeline
    {
        /// <summary>Set to true to cancel the current export operation.</summary>
        public static volatile bool CancelRequested;

        /// <summary>
        /// 流水线配置（每次运行时的参数）。
        /// </summary>
        public class Options
        {
            public int HeaderRow = 1;
            public int TypeRow = 2;
            public int CommentRow = 3;
            public int DataStartRow = 4;

            public bool SkipHiddenRows = true;
            public bool SkipHiddenColumns = true;
            public bool SkipEmptyRows = true;

            public string[] SkipSheetPrefixes = { "_", "#" };

            public string ExcelDir = "Assets/Excel";
            public string OutputDir = "Assets/Data";

            public CodeGenerator.Config CodeGenConfig = new CodeGenerator.Config();
            public AssetGenerator.Config AssetGenConfig = new AssetGenerator.Config();
            public JsonExportConfig JsonConfig = null;

            /// <summary>启用校验</summary>
            public bool EnableValidation = true;
            /// <summary>校验失败时阻止导出</summary>
            public bool BlockOnValidationError = true;

            /// <summary>同时导出 JSON</summary>
            public bool ExportJson = false;

            /// <summary>启用增量导出（仅处理变更的 Sheet）</summary>
            public bool UseIncrementalExport = false;
        }

        /// <summary>
        /// JSON 导出配置。
        /// </summary>
        public class JsonExportConfig
        {
            public string OutputDir = "Assets/Data";
            public bool PrettyPrint = true;
            public bool EnsureAscii = false;
            public string JsonMode = "array"; // array | id_keyed
        }

        /// <summary>
        /// 流水线执行结果。
        /// </summary>
        public class Result
        {
            public bool Success;
            public int FilesProcessed;
            public int SheetsProcessed;
            public int TotalRows;
            public int ErrorCount;
            public int WarningCount;
            public int FilesFailed;
            public int SheetsSkipped;
            public List<string> GeneratedAssets = new List<string>();
            public List<string> FailedFiles = new List<string>();
            public ValidationReport ValidationReport = new ValidationReport();
            public TimeSpan Elapsed;
        }

        // ============================================================
        // 单文件处理
        // ============================================================

        public static Result ProcessFile(string excelPath, Options options)
        {
            CancelRequested = false;
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var result = new Result();

            try
            {
                if (!File.Exists(excelPath))
                {
                    Debug.LogError($"[ExcelToJSON] 文件不存在: {excelPath}");
                    result.Success = false;
                    return result;
                }

                var fullPath = Path.GetFullPath(excelPath);

                var readResult = ExcelReader.Read(fullPath,
                    options.SkipHiddenRows,
                    options.SkipHiddenColumns,
                    options.SkipEmptyRows);

                result.FilesProcessed = 1;

                // 先解析 #Rules 表（如果存在）
                var customRules = new List<Validator.Rules.RuleConfig>();
                foreach (var sheetName in readResult.SheetNames)
                {
                    if (sheetName.StartsWith("#Rules") || sheetName.Equals("Rules", System.StringComparison.OrdinalIgnoreCase))
                    {
                        if (readResult.Sheets.TryGetValue(sheetName, out var rulesRows))
                        {
                            var parsed = Validator.Rules.RulesSheetParser.Parse(rulesRows, sheetName);
                            customRules.AddRange(parsed);
                            if (parsed.Count > 0)
                                Debug.Log($"[ExcelToJSON] 从 {sheetName} 加载了 {parsed.Count} 条自定义校验规则");
                        }
                    }
                }

                // --- 预解析所有 Sheet 结构，构建引用目标表的 ID 缓存（Stage 3） ---
                var refIdCache = BuildRefIdCache(readResult, options);

                // --- Mode B detection: scan all [ExcelTable] types ---
                var sheetTypeMap = Mapping.AttributeMapping.ScanTableSheetMap();

                // Schema migration: remember user choice across sheets
                int schemaSnapChoice = 0; // 0=未选择, 1=全部继续, 2=全部跳过

                // --- Incremental export: load hash cache ---
                Dictionary<string, string> hashCache = null;
                HashSet<string> changedSheets = null;
                if (options.UseIncrementalExport)
                {
                    hashCache = IncrementalCache.LoadCache(options.ExcelDir);
                    changedSheets = new HashSet<string>();
                }

                // Handle CSV files (already read, just process as single sheet)
                if (CsvReader.IsCsvFile(excelPath))
                {
                    ProcessCsvFile(fullPath, options, customRules, sheetTypeMap,
                        refIdCache, hashCache, changedSheets, result);
                    sw.Stop();
                    result.Elapsed = sw.Elapsed;
                    result.ErrorCount = result.ValidationReport.ErrorCount;
                    result.WarningCount = result.ValidationReport.WarningCount;
                    LogResult(result);
                    return result;
                }

                foreach (var sheetName in readResult.SheetNames)
                {
                    // 跳过隐藏 Sheet
                    if (readResult.SheetHidden.TryGetValue(sheetName, out var hidden) && hidden)
                        continue;

                    // 跳过 #Rules Sheet 本身
                    if (sheetName.StartsWith("#Rules") || sheetName.StartsWith("#"))
                        continue;

                    var rows = readResult.Sheets[sheetName];

                    // --- Incremental export: skip if unchanged ---
                    if (options.UseIncrementalExport && hashCache != null)
                    {
                        var hasChanged = IncrementalCache.HasChanged(options.ExcelDir,
                            readResult.FileName, sheetName, rows,
                            options.DataStartRow, hashCache);

                        // Also force re-export if this sheet depends on a changed sheet (ref chain)
                        if (!hasChanged && changedSheets != null && changedSheets.Count > 0)
                        {
                            var dependents = IncrementalCache.GetDependentSheets(
                                options.ExcelDir, sheetName);
                            foreach (var dep in dependents)
                            {
                                if (changedSheets.Contains(dep))
                                {
                                    hasChanged = true;
                                    Debug.Log($"[ExcelToJSON] {sheetName}: 关联表 \"{dep}\" 已变更，级联重新导出");
                                    break;
                                }
                            }
                        }

                        if (!hasChanged)
                        {
                            Debug.Log($"[ExcelToJSON] {sheetName}: 未变更，跳过（增量模式）");
                            result.SheetsProcessed++;
                            continue;
                        }
                        changedSheets.Add(sheetName);
                    }

                    // 解析结构
                    TableSchema schema;
                    try
                    {
                        schema = SchemaParser.Parse(rows, sheetName,
                            readResult.FileName,
                            options.HeaderRow,
                            options.TypeRow,
                            options.CommentRow,
                            options.DataStartRow,
                            options.SkipSheetPrefixes);
                    }
                    catch (FormatException ex)
                    {
                        result.ValidationReport.Add(
                            readResult.FileName, sheetName, 0, "",
                            "", "SchemaParse", ex.Message, ErrorLevel.Error);
                        continue;
                    }

                    // 跳过被标记的 Sheet
                    if (schema == null) continue;

                    // --- Schema 迁移检测（只弹一次，记住选择） ---
                    if (schemaSnapChoice == 0) // 0=未选择, 1=全部继续, 2=全部跳过
                    {
                        var snapshot = SchemaSnapshot.Load(schema.TableName, options.ExcelDir);
                        if (snapshot != null)
                        {
                            var changes = SchemaDiffer.Diff(schema, snapshot);
                            if (changes.Count > 0 && SchemaDiffer.HasDangerousChanges(changes))
                            {
                                var summary = SchemaDiffer.BuildSummary(changes, schema.TableName);
                                Debug.Log($"[ExcelToJSON] {summary}");

                                var choice = UnityEditor.EditorUtility.DisplayDialogComplex(
                                    $"Schema 变更: {schema.TableName}",
                                    $"检测到危险变更:\n{summary}\n\n如何继续？",
                                    "全部继续（不再提醒）",
                                    "跳过此表",
                                    "取消导出");

                                if (choice == 2) // 取消导出
                                {
                                    CancelRequested = true;
                                    Debug.LogWarning("[ExcelToJSON] 用户取消导出");
                                    break;
                                }
                                else if (choice == 1) // 跳过此表
                                {
                                    result.ValidationReport.Add(
                                        readResult.FileName, schema.TableName, 0, "", "",
                                        "SchemaMigration",
                                        $"Schema 变更被跳过: {changes.Count(c => c.IsDangerous)} 处危险变更",
                                        ErrorLevel.Warning);
                                    continue;
                                }
                                else // 全部继续
                                {
                                    schemaSnapChoice = 1;
                                    Debug.Log($"[ExcelToJSON] {schema.TableName}: 用户选择继续导出，后续不再提醒");
                                }
                            }
                        }
                    }
                    else if (schemaSnapChoice == 2)
                    {
                        continue; // Skip remaining sheets
                    }

                    // --- Cancel check ---
                    if (CancelRequested)
                    {
                        Debug.LogWarning("[ExcelToJSON] 导出被取消");
                        break;
                    }

                    // 解析数据
                    var errors = new List<ValidationError>();
                    var data = DataParser.ParseData(rows, schema, errors);
                    result.TotalRows += data.Count;

                    foreach (var err in errors)
                        result.ValidationReport.Errors.Add(err);

                    // 校验（含 Stage 1+2+3 引用完整性检查）
                    if (options.EnableValidation)
                    {
                        var validationReport = ValidationEngine.ValidateWithRefs(
                            rows, schema, readResult.FileName, refIdCache, customRules);
                        foreach (var err in validationReport.Errors)
                            result.ValidationReport.Errors.Add(err);
                    }

                    // 是否阻止导出
                    if (options.BlockOnValidationError
                        && result.ValidationReport.HasErrors)
                    {
                        Debug.LogError(
                            $"[ExcelToJSON] {sheetName}: 校验发现 {result.ValidationReport.ErrorCount} 个错误，导出已阻止");
                        continue;
                    }

                    // --- 模式检测：该 Sheet 是否为 Mode B（有 [ExcelTable] 标记的 C# 类） ---
                    var isModeB = sheetTypeMap.TryGetValue(schema.TableName, out var mappedType);

                    // Mode B: 执行属性映射兼容性校验
                    if (isModeB && mappedType != null)
                    {
                        var attrErrors = Mapping.AttributeMapping.ValidateCompatibility(
                            mappedType, schema, readResult.FileName);
                        foreach (var err in attrErrors)
                            result.ValidationReport.Errors.Add(err);
                    }

                    // 生成 C# 代码（Mode A 才生成，Mode B 已经有 C# 类）
                    if (!isModeB)
                    {
                        CodeGenerator.WriteToDisk(schema, options.CodeGenConfig);
                        AssetDatabase.Refresh();
                    }

                    // 生成 ScriptableObject（Mode B 传入映射类型）
                    var assetPath = AssetGenerator.Generate(data, schema, options.AssetGenConfig,
                        isModeB ? mappedType : null);
                    if (!string.IsNullOrEmpty(assetPath))
                        result.GeneratedAssets.Add(assetPath);

                    // 导出 JSON（可选）
                    if (options.ExportJson && options.JsonConfig != null)
                    {
                        ExportJson(data, schema, options.JsonConfig);
                    }

                    result.SheetsProcessed++;

                    // --- Save schema snapshot after successful export ---
                    SchemaSnapshot.Save(
                        SchemaSnapshot.FromSchema(schema, readResult.FileName),
                        options.ExcelDir);

                    // --- Incremental: update hash cache after successful export ---
                    if (options.UseIncrementalExport && hashCache != null)
                    {
                        IncrementalCache.UpdateCache(
                            readResult.FileName, sheetName, rows,
                            options.DataStartRow, hashCache, options.ExcelDir);
                    }
                }

                // --- Incremental: save reference graph for ref chain tracking ---
                if (options.UseIncrementalExport && refIdCache != null)
                {
                    var refGraph = new Dictionary<string, HashSet<string>>();
                    foreach (var sn in readResult.SheetNames)
                    {
                        if (!readResult.Sheets.TryGetValue(sn, out var sRows)) continue;
                        TableSchema sSchema;
                        try
                        {
                            sSchema = SchemaParser.Parse(sRows, sn, readResult.FileName,
                                options.HeaderRow, options.TypeRow, options.CommentRow,
                                options.DataStartRow, options.SkipSheetPrefixes);
                        }
                        catch { continue; }
                        if (sSchema == null) continue;

                        var targets = new HashSet<string>();
                        foreach (var f in sSchema.Fields)
                        {
                            if (f.IsCompositeType && !string.IsNullOrEmpty(f.CompositeParam)
                                && (f.NormalizedType?.StartsWith("ref:") == true
                                    || f.NormalizedType?.StartsWith("enum:") == true))
                            {
                                targets.Add(f.CompositeParam);
                            }
                        }
                        if (targets.Count > 0)
                            refGraph[sn] = targets;
                    }
                    IncrementalCache.SaveRefGraph(options.ExcelDir, refGraph);
                }

                AssetDatabase.SaveAssets();
                AssetDatabase.Refresh();

                result.Success = !result.ValidationReport.HasErrors;
            }
            catch (IOException ioEx)
            {
                result.FilesFailed++;
                result.FailedFiles.Add(excelPath);
                result.ValidationReport.Add(
                    Path.GetFileName(excelPath), "", 0, "", "",
                    "FileError", $"文件读取失败（可能被占用或损坏）: {ioEx.Message}", ErrorLevel.Error);
                Debug.LogError($"[ExcelToJSON] 文件 I/O 错误: {excelPath}: {ioEx.Message}");
            }
            catch (System.UnauthorizedAccessException uaEx)
            {
                result.FilesFailed++;
                result.FailedFiles.Add(excelPath);
                result.ValidationReport.Add(
                    Path.GetFileName(excelPath), "", 0, "", "",
                    "FileError", $"文件访问被拒绝: {uaEx.Message}", ErrorLevel.Error);
                Debug.LogError($"[ExcelToJSON] 文件访问错误: {excelPath}: {uaEx.Message}");
            }
            catch (Exception ex)
            {
                result.FilesFailed++;
                result.FailedFiles.Add(excelPath);
                result.ValidationReport.Add(
                    Path.GetFileName(excelPath), "", 0, "", "",
                    "FileError", $"处理异常: {ex.Message}", ErrorLevel.Error);
                Debug.LogError($"[ExcelToJSON] 处理 {excelPath} 时发生异常: {ex}");
            }

            sw.Stop();
            result.Elapsed = sw.Elapsed;
            result.ErrorCount = result.ValidationReport.ErrorCount;
            result.WarningCount = result.ValidationReport.WarningCount;

            LogResult(result);

            return result;
        }

        // ============================================================
        // 批量处理
        // ============================================================

        public static Result ProcessDirectory(string directory, Options options,
            System.Action<int, int, string> onProgress = null)
        {
            CancelRequested = false;
            var result = new Result();
            var sw = System.Diagnostics.Stopwatch.StartNew();

            if (!Directory.Exists(directory))
            {
                Debug.LogError($"[ExcelToJSON] 目录不存在: {directory}");
                result.Success = false;
                return result;
            }

            var excelFiles = Directory.GetFiles(directory, "*.xlsx", SearchOption.AllDirectories)
                .Concat(Directory.GetFiles(directory, "*.xls", SearchOption.AllDirectories))
                .Concat(Directory.GetFiles(directory, "*.csv", SearchOption.AllDirectories))
                .Concat(Directory.GetFiles(directory, "*.tsv", SearchOption.AllDirectories))
                .Where(f => !Path.GetFileName(f).StartsWith("~$")) // 跳过 Excel 临时文件
                .ToArray();

            for (int i = 0; i < excelFiles.Length; i++)
            {
                if (CancelRequested)
                {
                    Debug.LogWarning("[ExcelToJSON] 批量导出被取消");
                    break;
                }

                var file = excelFiles[i];
                var relativePath = GetRelativePath(file, options.ExcelDir);
                var fileName = Path.GetFileName(file);

                onProgress?.Invoke(i + 1, excelFiles.Length, fileName);

                var fileResult = ProcessFile(file, options);

                result.FilesProcessed += fileResult.FilesProcessed;
                result.SheetsProcessed += fileResult.SheetsProcessed;
                result.TotalRows += fileResult.TotalRows;
                result.FilesFailed += fileResult.FilesFailed;
                result.FailedFiles.AddRange(fileResult.FailedFiles);
                result.GeneratedAssets.AddRange(fileResult.GeneratedAssets);
                result.ValidationReport.Errors.AddRange(fileResult.ValidationReport.Errors);
            }

            sw.Stop();
            result.Elapsed = sw.Elapsed;
            result.ErrorCount = result.ValidationReport.ErrorCount;
            result.WarningCount = result.ValidationReport.WarningCount;
            result.Success = !result.ValidationReport.HasErrors;

            LogResult(result);

            return result;
        }

        // ============================================================
        // 辅助
        // ============================================================

        private static void LogResult(Result result)
        {
            if (result.SheetsProcessed == 0)
            {
                Debug.LogWarning("[ExcelToJSON] 没有处理任何 Sheet");
                return;
            }

            var status = result.ErrorCount > 0 ? "⚠ 有错误"
                : result.WarningCount > 0 ? "⚠ 有警告"
                : "✅ 全部通过";

            var failureInfo = result.FilesFailed > 0
                ? $" 失败: {result.FilesFailed} 文件" : "";

            Debug.Log(
                $"[ExcelToJSON] {status} | {result.FilesProcessed} 文件 → " +
                $"{result.SheetsProcessed} 表 → {result.TotalRows} 行 | " +
                $"错误: {result.ErrorCount} 警告: {result.WarningCount}{failureInfo} | " +
                $"耗时: {result.Elapsed.TotalSeconds:F1}s");
        }

        private static string GetRelativePath(string fullPath, string baseDir)
        {
            var full = Path.GetFullPath(fullPath).Replace("\\", "/");
            var base_ = Path.GetFullPath(baseDir).Replace("\\", "/");
            if (full.StartsWith(base_))
                return full.Substring(base_.Length).TrimStart('/');
            return Path.GetFileName(fullPath);
        }

        private static void ExportJson(
            List<Dictionary<string, object>> data,
            TableSchema schema,
            JsonExportConfig config)
        {
            var jsonPath = Path.Combine(config.OutputDir, $"{schema.TableName}.json");
            var json = SerializeToJson(data, schema, config);
            File.WriteAllText(jsonPath, json, System.Text.Encoding.UTF8);

            // Also write .hash for version comparison (hot-update)
            var hash = ComputeJsonHash(json);
            var hashPath = Path.Combine(config.OutputDir, $"{schema.TableName}.json.hash");
            File.WriteAllText(hashPath, hash, System.Text.Encoding.UTF8);
        }

        /// <summary>
        /// Process a CSV file (single sheet = file name).
        /// </summary>
        private static void ProcessCsvFile(string fullPath, Options options,
            List<Validator.Rules.RuleConfig> customRules,
            Dictionary<string, Type> sheetTypeMap,
            Dictionary<string, HashSet<int>> refIdCache,
            Dictionary<string, string> hashCache,
            HashSet<string> changedSheets,
            Result result)
        {
            var rows = CsvReader.Read(fullPath);
            var sheetName = Path.GetFileNameWithoutExtension(fullPath);
            var fileName = Path.GetFileName(fullPath);

            if (rows.Count == 0) return;

            TableSchema schema;
            try
            {
                schema = SchemaParser.Parse(rows, sheetName, fileName,
                    options.HeaderRow, options.TypeRow, options.CommentRow,
                    options.DataStartRow, options.SkipSheetPrefixes);
            }
            catch { return; }
            if (schema == null) return;

            var errors = new List<ValidationError>();
            var data = DataParser.ParseData(rows, schema, errors);
            result.TotalRows += data.Count;

            foreach (var err in errors)
                result.ValidationReport.Errors.Add(err);

            if (options.EnableValidation)
            {
                var vReport = ValidationEngine.ValidateWithRefs(
                    rows, schema, fileName, refIdCache, customRules);
                foreach (var err in vReport.Errors)
                    result.ValidationReport.Errors.Add(err);
            }

            if (!options.BlockOnValidationError || !result.ValidationReport.HasErrors)
            {
                if (!sheetTypeMap.TryGetValue(sheetName, out _))
                    CodeGenerator.WriteToDisk(schema, options.CodeGenConfig);

                var assetPath = AssetGenerator.Generate(data, schema, options.AssetGenConfig);
                if (!string.IsNullOrEmpty(assetPath))
                    result.GeneratedAssets.Add(assetPath);

                result.SheetsProcessed++;
                result.FilesProcessed = 1;

                SchemaSnapshot.Save(
                    SchemaSnapshot.FromSchema(schema, fileName), options.ExcelDir);
            }
        }

        private static string ComputeJsonHash(string json)
        {
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                var bytes = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(json));
                return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
            }
        }

        private static string SerializeToJson(
            List<Dictionary<string, object>> data,
            TableSchema schema,
            JsonExportConfig config)
        {
            var sb = new System.Text.StringBuilder();
            var indent = config.PrettyPrint ? "  " : "";
            var nl = config.PrettyPrint ? "\n" : "";

            sb.Append("[").Append(nl);

            for (int i = 0; i < data.Count; i++)
            {
                var row = data[i];
                sb.Append(indent).Append("{").Append(nl);

                var fieldList = row.ToList();
                for (int j = 0; j < fieldList.Count; j++)
                {
                    var kv = fieldList[j];
                    var comma = j < fieldList.Count - 1 ? "," : "";
                    var valueStr = SerializeJsonValue(kv.Value, config.EnsureAscii);
                    sb.Append(indent).Append(indent)
                        .Append($"\"{kv.Key}\": {valueStr}{comma}")
                        .Append(nl);
                }

                var rowComma = i < data.Count - 1 ? "," : "";
                sb.Append(indent).Append("}").Append(rowComma).Append(nl);
            }

            sb.Append("]");
            return sb.ToString();
        }

        private static string SerializeJsonValue(object value, bool ensureAscii)
        {
            if (value == null) return "null";
            if (value is string s)
            {
                var escaped = s.Replace("\\", "\\\\").Replace("\"", "\\\"");
                return $"\"{escaped}\"";
            }
            if (value is bool b) return b ? "true" : "false";
            if (value is float f) return f.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (value is int i) return i.ToString();
            if (value is Array arr)
            {
                var sb = new System.Text.StringBuilder("[");
                for (int idx = 0; idx < arr.Length; idx++)
                {
                    if (idx > 0) sb.Append(", ");
                    sb.Append(SerializeJsonValue(arr.GetValue(idx), ensureAscii));
                }
                sb.Append("]");
                return sb.ToString();
            }
            if (value is Vector2 v2) return $"[{v2.x}, {v2.y}]";
            if (value is Vector3 v3) return $"[{v3.x}, {v3.y}, {v3.z}]";
            if (value is Color c) return $"\"#{ColorUtility.ToHtmlStringRGBA(c)}\"";
            return value.ToString();
        }

        // ============================================================
        // 跨表引用缓存（Stage 3 校验用）
        // ============================================================

        /// <summary>
        /// Build a cache of all table→id-sets from this file and cross-file refs.
        /// Used by Stage 3 RefIntegrityRule.
        /// </summary>
        private static Dictionary<string, HashSet<int>> BuildRefIdCache(
            ExcelReader.ReadResult readResult, Options options)
        {
            var cache = new Dictionary<string, HashSet<int>>();
            var targetTables = new HashSet<string>();

            // Step 1: Collect all ref target table names from this file's sheets
            foreach (var sheetName in readResult.SheetNames)
            {
                if (readResult.SheetHidden.TryGetValue(sheetName, out var hidden) && hidden)
                    continue;
                if (sheetName.StartsWith("#Rules") || sheetName.StartsWith("#"))
                    continue;

                var rows = readResult.Sheets[sheetName];
                TableSchema schema;
                try
                {
                    schema = SchemaParser.Parse(rows, sheetName, readResult.FileName,
                        options.HeaderRow, options.TypeRow, options.CommentRow,
                        options.DataStartRow, options.SkipSheetPrefixes);
                }
                catch { continue; }
                if (schema == null) continue;

                foreach (var field in schema.Fields)
                {
                    if (field.IsCompositeType
                        && !string.IsNullOrEmpty(field.CompositeParam)
                        && (field.NormalizedType?.StartsWith("ref:") == true
                            || field.NormalizedType?.StartsWith("enum:") == true))
                    {
                        targetTables.Add(field.CompositeParam);
                    }
                }
            }

            // Step 2: Extract IDs from same-file sheets
            foreach (var sheetName in readResult.SheetNames)
            {
                if (readResult.SheetHidden.TryGetValue(sheetName, out var hidden) && hidden)
                    continue;
                if (sheetName.StartsWith("#Rules") || sheetName.StartsWith("#"))
                    continue;

                if (!targetTables.Contains(sheetName))
                    continue;

                var rows = readResult.Sheets[sheetName];
                TableSchema schema;
                try
                {
                    schema = SchemaParser.Parse(rows, sheetName, readResult.FileName,
                        options.HeaderRow, options.TypeRow, options.CommentRow,
                        options.DataStartRow, options.SkipSheetPrefixes);
                }
                catch { continue; }
                if (schema == null) continue;

                cache[sheetName] = RefIntegrityRule.ExtractIds(rows, schema);
            }

            // Step 3: For remaining ref targets not in this file, search other Excel files
            var unresolved = new HashSet<string>();
            foreach (var t in targetTables)
            {
                if (!cache.ContainsKey(t))
                    unresolved.Add(t);
            }

            if (unresolved.Count > 0 && Directory.Exists(options.ExcelDir))
            {
                var otherFiles = Directory.GetFiles(options.ExcelDir, "*.xlsx", SearchOption.AllDirectories)
                    .Concat(Directory.GetFiles(options.ExcelDir, "*.xls", SearchOption.AllDirectories))
                    .Where(f => !Path.GetFileName(f).StartsWith("~$")
                        && Path.GetFullPath(f) != Path.GetFullPath(
                            Path.Combine(options.ExcelDir, readResult.FileName)))
                    .ToArray();

                foreach (var filePath in otherFiles)
                {
                    if (unresolved.Count == 0) break;

                    try
                    {
                        var otherResult = ExcelReader.Read(filePath,
                            options.SkipHiddenRows,
                            options.SkipHiddenColumns,
                            options.SkipEmptyRows);

                        foreach (var sheetName in otherResult.SheetNames)
                        {
                            if (!unresolved.Contains(sheetName)) continue;

                            var rows = otherResult.Sheets[sheetName];
                            try
                            {
                                var schema = SchemaParser.Parse(rows, sheetName,
                                    otherResult.FileName,
                                    options.HeaderRow, options.TypeRow, options.CommentRow,
                                    options.DataStartRow, options.SkipSheetPrefixes);
                                cache[sheetName] = RefIntegrityRule.ExtractIds(rows, schema);
                                unresolved.Remove(sheetName);
                            }
                            catch { continue; }
                        }
                    }
                    catch { continue; }
                }
            }

            return cache;
        }
    }
}