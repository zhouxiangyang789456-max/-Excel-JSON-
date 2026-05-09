using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Editor.Generator;
using ExcelToJsonPlugin.Editor.Validator;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor
{
    /// <summary>
    /// 临时菜单入口（Sprint 1 MVP）。
    /// Sprint 2 将被完整的 Editor Window 取代。
    /// </summary>
    public static class ExcelToJsonMenu
    {
        private const string MenuRoot = "Tools/Excel To JSON/";

        [MenuItem(MenuRoot + "Export All (Assets/Excel → Assets/Data)", priority = 100)]
        public static void ExportAll()
        {
            var options = new Pipeline.Options
            {
                ExcelDir = "Assets/Excel",
                OutputDir = "Assets/Data",
                CodeGenConfig = new CodeGenerator.Config
                {
                    RowOutputDir = "Scripts/Generated/Data",
                    TableOutputDir = "Scripts/Generated/Tables",
                    Namespace = "Game.Data",
                    RowSuffix = "Row",
                    TableSuffix = "Table",
                    AssetsRoot = "Assets",
                },
                AssetGenConfig = new AssetGenerator.Config
                {
                    OutputDir = "Data",
                    AssetsRoot = "Assets",
                },
                EnableValidation = true,
                BlockOnValidationError = false,
                ExportJson = true,
                JsonConfig = new Pipeline.JsonExportConfig
                {
                    OutputDir = "Assets/Data",
                    PrettyPrint = true,
                    EnsureAscii = false,
                    JsonMode = "array",
                },
            };

            var result = Pipeline.ProcessDirectory(options.ExcelDir, options);

            if (result.Success)
            {
                if (!Application.isBatchMode)
                {
                    EditorUtility.DisplayDialog("导出完成",
                        $"处理了 {result.FilesProcessed} 个文件\n" +
                        $"{result.SheetsProcessed} 个 Sheet\n" +
                        $"{result.TotalRows} 行数据\n" +
                        $"耗时: {result.Elapsed.TotalSeconds:F1}s",
                        "确定");
                }
            }
            else
            {
                if (!Application.isBatchMode)
                {
                    EditorUtility.DisplayDialog("导出有错误",
                        $"{result.ErrorCount} 个错误, {result.WarningCount} 个警告\n" +
                        $"详情请查看 Console 窗口",
                        "确定");
                }
            }
        }

        [MenuItem(MenuRoot + "Validate Only", priority = 101)]
        public static void ValidateOnly()
        {
            var options = new Pipeline.Options
            {
                ExcelDir = "Assets/Excel",
                EnableValidation = true,
                BlockOnValidationError = false,
            };

            // Scan and validate all Excel files
            var excelDir = options.ExcelDir;
            if (!System.IO.Directory.Exists(excelDir))
            {
                Debug.LogWarning($"[ExcelToJSON] Excel directory not found: {excelDir}");
                return;
            }

            var files = System.IO.Directory.GetFiles(excelDir, "*.xlsx", System.IO.SearchOption.AllDirectories)
                .Concat(System.IO.Directory.GetFiles(excelDir, "*.xls", System.IO.SearchOption.AllDirectories))
                .Where(f => !System.IO.Path.GetFileName(f).StartsWith("~$"))
                .ToArray();

            var totalReport = new Editor.Core.Models.ValidationReport();

            foreach (var file in files)
            {
                var readResult = Core.ExcelReader.Read(file);
                foreach (var sheetName in readResult.SheetNames)
                {
                    if (!readResult.Sheets.TryGetValue(sheetName, out var rows)) continue;

                    try
                    {
                        var schema = Core.SchemaParser.Parse(rows, sheetName, readResult.FileName,
                            1, 2, 3, 4, new[] { "_", "#" });
                        if (schema == null) continue;

                        var report = Validator.ValidationEngine.ValidateAll(rows, schema, readResult.FileName);
                        totalReport.Errors.AddRange(report.Errors);
                    }
                    catch (System.FormatException ex)
                    {
                        totalReport.Add(readResult.FileName, sheetName, 0, "", "",
                            "SchemaError", ex.Message, ErrorLevel.Error);
                    }
                }
            }

            if (totalReport.Errors.Count == 0)
            {
                Debug.Log($"[ExcelToJSON] Validation passed: {files.Length} files OK");
                if (!Application.isBatchMode)
                    EditorUtility.DisplayDialog("Validation", $"All {files.Length} files passed.", "OK");
            }
            else
            {
                Debug.LogWarning($"[ExcelToJSON] Validation found {totalReport.ErrorCount} errors, {totalReport.WarningCount} warnings");
                foreach (var err in totalReport.Errors)
                {
                    if (err.Level == ErrorLevel.Error)
                        Debug.LogError(err.ToString());
                    else if (err.Level == ErrorLevel.Warning)
                        Debug.LogWarning(err.ToString());
                }
                if (!Application.isBatchMode)
                    EditorUtility.DisplayDialog("Validation Failed",
                        $"{totalReport.ErrorCount} errors, {totalReport.WarningCount} warnings\nSee Console for details.", "OK");
            }
        }

        [MenuItem(MenuRoot + "Open Excel Directory", priority = 200)]
        public static void OpenExcelDirectory()
        {
            var path = System.IO.Path.GetFullPath("Assets/Excel");
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            EditorUtility.RevealInFinder(path);
        }

        [MenuItem(MenuRoot + "Open Output Directory", priority = 201)]
        public static void OpenOutputDirectory()
        {
            var path = System.IO.Path.GetFullPath("Assets/Data");
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            EditorUtility.RevealInFinder(path);
        }
    }
}
