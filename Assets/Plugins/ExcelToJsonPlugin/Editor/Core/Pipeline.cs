using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Editor.Generator;
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
            public List<string> GeneratedAssets = new List<string>();
            public ValidationReport ValidationReport = new ValidationReport();
            public TimeSpan Elapsed;
        }

        // ============================================================
        // 单文件处理
        // ============================================================

        public static Result ProcessFile(string excelPath, Options options)
        {
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

                foreach (var sheetName in readResult.SheetNames)
                {
                    // 跳过隐藏 Sheet
                    if (readResult.SheetHidden.TryGetValue(sheetName, out var hidden) && hidden)
                        continue;

                    var rows = readResult.Sheets[sheetName];

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

                    // 解析数据
                    var errors = new List<ValidationError>();
                    var data = DataParser.ParseData(rows, schema, errors);
                    result.TotalRows += data.Count;

                    foreach (var err in errors)
                        result.ValidationReport.Errors.Add(err);

                    // 基本校验（类型匹配已在 DataParser 中做）
                    if (options.EnableValidation)
                    {
                        // TODO Sprint 2: 完整校验引擎
                    }

                    // 是否阻止导出
                    if (options.BlockOnValidationError
                        && result.ValidationReport.HasErrors)
                    {
                        Debug.LogError(
                            $"[ExcelToJSON] {sheetName}: 校验发现 {result.ValidationReport.ErrorCount} 个错误，导出已阻止");
                        continue;
                    }

                    // 生成 C# 代码
                    var (rowPath, tablePath) = CodeGenerator.WriteToDisk(schema, options.CodeGenConfig);

                    // 刷新 AssetDatabase 使新生成的代码可被反射
                    AssetDatabase.Refresh();

                    // 生成 ScriptableObject
                    var assetPath = AssetGenerator.Generate(data, schema, options.AssetGenConfig);
                    if (!string.IsNullOrEmpty(assetPath))
                        result.GeneratedAssets.Add(assetPath);

                    // 导出 JSON（可选）
                    if (options.ExportJson && options.JsonConfig != null)
                    {
                        ExportJson(data, schema, options.JsonConfig);
                    }

                    result.SheetsProcessed++;
                }

                AssetDatabase.SaveAssets();
                AssetDatabase.Refresh();

                result.Success = !result.ValidationReport.HasErrors;
            }
            catch (Exception ex)
            {
                Debug.LogError($"[ExcelToJSON] 处理 {excelPath} 时发生异常: {ex}");
                result.Success = false;
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

        public static Result ProcessDirectory(string directory, Options options)
        {
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
                .Where(f => !Path.GetFileName(f).StartsWith("~$")) // 跳过 Excel 临时文件
                .ToArray();

            foreach (var file in excelFiles)
            {
                var relativePath = GetRelativePath(file, options.ExcelDir);
                var fileResult = ProcessFile(file, options);

                result.FilesProcessed += fileResult.FilesProcessed;
                result.SheetsProcessed += fileResult.SheetsProcessed;
                result.TotalRows += fileResult.TotalRows;
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

            Debug.Log(
                $"[ExcelToJSON] {status} | {result.FilesProcessed} 文件 → " +
                $"{result.SheetsProcessed} 表 → {result.TotalRows} 行 | " +
                $"错误: {result.ErrorCount} 警告: {result.WarningCount} | " +
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
            // TODO Sprint 5: 完整 JSON 导出
            var jsonDir = Path.Combine(config.OutputDir);
            Directory.CreateDirectory(jsonDir);

            var jsonPath = Path.Combine(jsonDir, $"{schema.TableName}.json");
            var json = SerializeToJson(data, schema, config);
            File.WriteAllText(jsonPath, json, System.Text.Encoding.UTF8);
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
    }
}
