using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Generator;
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
            // TODO Sprint 2: 仅校验不导出的完整实现
            Debug.Log("[ExcelToJSON] 校验功能将在 Sprint 2 完善");
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
