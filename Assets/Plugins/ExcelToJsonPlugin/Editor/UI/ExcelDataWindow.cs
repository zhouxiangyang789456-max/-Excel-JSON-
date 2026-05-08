using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Runtime;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    /// <summary>
    /// 主编编辑器面板：文件树 + 5 个标签页 + 状态栏。
    /// 打开方式：Window > Excel Data Manager
    /// </summary>
    public class ExcelDataWindow : EditorWindow
    {
        // ===== 状态 =====
        private string excelDir = "Assets/Excel";
        private string outputDir = "Assets/Data";

        private List<ExcelFileEntry> fileEntries = new List<ExcelFileEntry>();
        private string selectedFilePath;
        private string selectedSheetName;
        private HashSet<string> multiSelectedFiles = new HashSet<string>();
        private string lastClickedFile; // Shift-click range anchor
        private int currentTab;

        private string statusMessage = "";
        private bool statusIsError;

        private Vector2 treeScroll, tabScroll;
        private string searchFilter = "";

        // Tab 内容
        private DataPreviewDrawer previewDrawer;
        private ValidationDrawer validationDrawer;
        private ExportDrawer exportDrawer;
        private RuntimeDrawer runtimeDrawer;
        private DashboardDrawer dashboardDrawer;

        // ===== 入口 =====
        [MenuItem("Window/Excel Data Manager", priority = 100)]
        public static void ShowWindow()
        {
            var window = GetWindow<ExcelDataWindow>("Excel Data Manager");
            window.minSize = new Vector2(900, 550);
            window.Show();
        }

        // ===== 生命周期 =====
        private void OnEnable()
        {
            previewDrawer = new DataPreviewDrawer(this);
            validationDrawer = new ValidationDrawer(this);
            exportDrawer = new ExportDrawer(this);
            runtimeDrawer = new RuntimeDrawer(this);
            dashboardDrawer = new DashboardDrawer(this);

            RefreshFileList();
        }

        private void OnFocus()
        {
            RefreshFileList();
        }

        // ===== 刷新文件列表 =====
        public void RefreshFileList()
        {
            fileEntries.Clear();
            if (!Directory.Exists(excelDir))
            {
                Directory.CreateDirectory(excelDir);
                return;
            }

            var files = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.AllDirectories)
                .Concat(Directory.GetFiles(excelDir, "*.xls", SearchOption.AllDirectories))
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .OrderBy(f => f);

            foreach (var fullPath in files)
            {
                var relativePath = GetRelativePath(fullPath);
                var entry = new ExcelFileEntry
                {
                    RelativePath = relativePath,
                    FullPath = fullPath,
                    FileName = Path.GetFileName(fullPath),
                };

                try
                {
                    // 快速读取 Sheet 名称（只读结构，不读数据）
                    var readResult = ExcelReader.Read(fullPath, false, false, false);
                    foreach (var sn in readResult.SheetNames)
                    {
                        if (!sn.StartsWith("_") && !sn.StartsWith("#"))
                            entry.Sheets.Add(new SheetEntry { Name = sn });
                    }
                }
                catch { /* 文件被占用或损坏时跳过 */ }

                fileEntries.Add(entry);
            }

            Repaint();
        }

        // ===== 绘制 =====
        private void OnGUI()
        {
            DrawToolbar();

            EditorGUILayout.BeginHorizontal();

            // 左侧：文件树
            EditorGUILayout.BeginVertical(GUILayout.Width(220));
            DrawFileTree();
            EditorGUILayout.EndVertical();

            // 分隔线
            EditorGUILayout.BeginVertical(GUILayout.Width(2));
            GUILayout.Box("", GUILayout.Width(2), GUILayout.ExpandHeight(true));
            EditorGUILayout.EndVertical();

            // 右侧：标签页内容
            EditorGUILayout.BeginVertical();
            DrawTabBar();
            DrawTabContent();
            EditorGUILayout.EndVertical();

            EditorGUILayout.EndHorizontal();

            // 底部状态栏
            DrawStatusBar();
        }

        // ===== 工具栏 =====
        private void DrawToolbar()
        {
            EditorGUILayout.BeginHorizontal(EditorStyles.toolbar);

            GUILayout.Label("Excel Data Manager", EditorStyles.boldLabel);

            GUILayout.FlexibleSpace();

            if (GUILayout.Button("Refresh", EditorStyles.toolbarButton, GUILayout.Width(60)))
                RefreshFileList();

            if (GUILayout.Button("Export All", EditorStyles.toolbarButton, GUILayout.Width(70)))
                RunExportAll();

            if (GUILayout.Button("Validate", EditorStyles.toolbarButton, GUILayout.Width(60)))
                RunValidate();

            GUILayout.Space(10);

            if (GUILayout.Button(EditorGUIUtility.IconContent("_Popup"), EditorStyles.toolbarButton, GUILayout.Width(30)))
                SettingsWindow.ShowWindow();

            EditorGUILayout.EndHorizontal();
        }

        // ===== 文件树 =====
        private void DrawFileTree()
        {
            EditorGUILayout.LabelField("Files", EditorStyles.boldLabel);

            searchFilter = EditorGUILayout.TextField(searchFilter, EditorStyles.toolbarSearchField);
            if (!string.IsNullOrEmpty(searchFilter))
                EditorGUILayout.LabelField($"  Filter: \"{searchFilter}\"", EditorStyles.miniLabel);

            // Multi-select controls
            if (multiSelectedFiles.Count > 0)
            {
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField($"{multiSelectedFiles.Count} selected", EditorStyles.miniLabel);
                if (GUILayout.Button("Export Selected", EditorStyles.miniButton, GUILayout.Width(90)))
                    RunExportSelected();
                if (GUILayout.Button("Validate Selected", EditorStyles.miniButton, GUILayout.Width(100)))
                    RunValidateSelected();
                if (GUILayout.Button("Clear", EditorStyles.miniButton, GUILayout.Width(40)))
                    multiSelectedFiles.Clear();
                EditorGUILayout.EndHorizontal();
                EditorGUILayout.Space(3);
            }

            treeScroll = EditorGUILayout.BeginScrollView(treeScroll);

            foreach (var entry in fileEntries)
            {
                if (!string.IsNullOrEmpty(searchFilter) &&
                    !entry.FileName.ToLower().Contains(searchFilter.ToLower()))
                    continue;

                bool isSelected = selectedFilePath == entry.RelativePath && selectedSheetName == null;
                bool isMultiSelected = multiSelectedFiles.Contains(entry.RelativePath);

                // 文件行
                var bgColor = isMultiSelected ? new Color(0.3f, 0.5f, 0.8f, 0.3f) : GUI.backgroundColor;
                var oldBg = GUI.backgroundColor;
                GUI.backgroundColor = bgColor;

                EditorGUILayout.BeginHorizontal();

                // Checkbox for multi-select
                var toggleRect = EditorGUILayout.GetControlRect(GUILayout.Width(18));
                var wasToggled = isMultiSelected;
                var nowToggled = EditorGUI.Toggle(toggleRect, wasToggled);
                if (nowToggled != wasToggled)
                    ToggleFileSelection(entry.RelativePath);

                // File label button
                var fileStyle = isSelected ? EditorStyles.boldLabel : EditorStyles.label;
                var fileIcon = entry.Sheets.Count > 0 ? "📄" : "📄";
                if (GUILayout.Button($"{fileIcon} {entry.FileName}", fileStyle, GUILayout.ExpandWidth(true)))
                {
                    HandleFileClick(entry.RelativePath);
                }

                EditorGUILayout.EndHorizontal();
                GUI.backgroundColor = oldBg;

                // Tooltip 显示完整路径
                var lastRect = GUILayoutUtility.GetLastRect();
                EditorGUI.LabelField(lastRect, new GUIContent("", entry.RelativePath));

                // 右键菜单（文件级）
                if (Event.current.type == EventType.ContextClick && lastRect.Contains(Event.current.mousePosition))
                {
                    var menu = new GenericMenu();
                    menu.AddItem(new GUIContent("Export This File"), false, () => RunExportSingle(entry.RelativePath, null));
                    menu.AddItem(new GUIContent("Validate This File"), false, () => RunValidateSingle(entry.RelativePath, null));
                    menu.AddItem(new GUIContent("Open Excel"), false, () => OpenExcel(entry.FullPath));
                    menu.AddItem(new GUIContent("Add to Selection"), false, () => multiSelectedFiles.Add(entry.RelativePath));
                    menu.ShowAsContext();
                    Event.current.Use();
                }

                // Sheet 列表
                if (isSelected || true) // 展开所有
                {
                    EditorGUI.indentLevel++;
                    foreach (var sheet in entry.Sheets)
                    {
                        var sheetStyle = isSelected && selectedSheetName == sheet.Name
                            ? EditorStyles.boldLabel : EditorStyles.label;

                        if (GUILayout.Button($"  📋 {sheet.Name}", sheetStyle, GUILayout.ExpandWidth(true)))
                        {
                            selectedFilePath = entry.RelativePath;
                            selectedSheetName = sheet.Name;
                            multiSelectedFiles.Clear();
                            lastClickedFile = entry.RelativePath;
                            currentTab = 1;
                        }

                        // 右键菜单（Sheet 级）
                        var sheetRect = GUILayoutUtility.GetLastRect();
                        if (Event.current.type == EventType.ContextClick && sheetRect.Contains(Event.current.mousePosition))
                        {
                            var menu = new GenericMenu();
                            menu.AddItem(new GUIContent("Export This Sheet"), false, () => RunExportSingle(entry.RelativePath, sheet.Name));
                            menu.AddItem(new GUIContent("Validate This Sheet"), false, () => RunValidateSingle(entry.RelativePath, sheet.Name));
                            menu.AddItem(new GUIContent("Open Excel"), false, () => OpenExcel(entry.FullPath));
                            menu.ShowAsContext();
                            Event.current.Use();
                        }
                    }
                    EditorGUI.indentLevel--;
                }
            }

            EditorGUILayout.EndScrollView();

            EditorGUILayout.Space(5);
            if (GUILayout.Button("+ Add Excel Directory"))
            {
                var newDir = EditorUtility.OpenFolderPanel("Select Excel Directory", "Assets", "");
                if (!string.IsNullOrEmpty(newDir))
                {
                    excelDir = GetRelativePath(newDir);
                    RefreshFileList();
                }
            }
        }

        private void HandleFileClick(string relPath)
        {
            bool ctrl = Event.current.control;
            bool shift = Event.current.shift;

            if (ctrl)
            {
                ToggleFileSelection(relPath);
                lastClickedFile = relPath;
            }
            else if (shift && lastClickedFile != null)
            {
                SelectFileRange(lastClickedFile, relPath);
            }
            else
            {
                selectedFilePath = relPath;
                selectedSheetName = null;
                multiSelectedFiles.Clear();
                lastClickedFile = relPath;
            }
        }

        private void ToggleFileSelection(string relPath)
        {
            if (multiSelectedFiles.Contains(relPath))
                multiSelectedFiles.Remove(relPath);
            else
                multiSelectedFiles.Add(relPath);
        }

        private void SelectFileRange(string from, string to)
        {
            multiSelectedFiles.Clear();
            bool inRange = false;
            foreach (var entry in fileEntries)
            {
                if (entry.RelativePath == from || entry.RelativePath == to)
                {
                    multiSelectedFiles.Add(entry.RelativePath);
                    inRange = !inRange;
                    if (!inRange) break; // Both ends found
                }
                else if (inRange)
                {
                    multiSelectedFiles.Add(entry.RelativePath);
                }
            }
        }

        // ===== 标签页 =====
        private static readonly string[] TabNames = { "Dashboard", "Data Preview", "Validation", "Export", "Runtime API" };

        private void DrawTabBar()
        {
            currentTab = GUILayout.Toolbar(currentTab, TabNames);
        }

        private void DrawTabContent()
        {
            tabScroll = EditorGUILayout.BeginScrollView(tabScroll);
            switch (currentTab)
            {
                case 0: dashboardDrawer?.Draw(); break;
                case 1: previewDrawer?.Draw(selectedFilePath, selectedSheetName, fileEntries); break;
                case 2: validationDrawer?.Draw(selectedFilePath, selectedSheetName, fileEntries); break;
                case 3: exportDrawer?.Draw(selectedFilePath, selectedSheetName, fileEntries); break;
                case 4: runtimeDrawer?.Draw(selectedFilePath, selectedSheetName, fileEntries); break;
            }
            EditorGUILayout.EndScrollView();
        }

        // ===== 状态栏 =====
        private void DrawStatusBar()
        {
            EditorGUILayout.BeginHorizontal(EditorStyles.toolbar);
            var style = statusIsError ? new GUIStyle(EditorStyles.label) { normal = { textColor = Color.red } }
                                      : EditorStyles.label;
            GUILayout.Label(string.IsNullOrEmpty(statusMessage) ? "Ready" : statusMessage, style);
            EditorGUILayout.EndHorizontal();
        }

        // ===== 操作 =====
        public void SetStatus(string msg, bool isError = false)
        {
            statusMessage = msg;
            statusIsError = isError;
            Repaint();
        }

        public void RunExportAll()
        {
            var options = PipelineOptions();

            try
            {
                EditorUtility.DisplayProgressBar("Excel To JSON", "Starting export...", 0f);

                var result = Pipeline.ProcessDirectory(excelDir, options,
                    (current, total, fileName) =>
                    {
                        var pct = (float)current / total;
                        EditorUtility.DisplayProgressBar(
                            "Excel To JSON — Exporting",
                            $"({current}/{total}) {fileName}",
                            pct);
                    });

                if (!Application.isBatchMode)
                    AssetDatabase.Refresh();

                if (result.Success)
                    SetStatus($"Export complete: {result.FilesProcessed} files, {result.SheetsProcessed} sheets, {result.TotalRows} rows ({result.Elapsed.TotalSeconds:F1}s)");
                else
                    SetStatus($"Export errors: {result.ErrorCount} errors, {result.WarningCount} warnings", true);

                // Show completion
                if (!Application.isBatchMode)
                {
                    Debug.Log($"[ExcelToJSON] Export complete: {result.FilesProcessed} files, {result.SheetsProcessed} sheets, {result.TotalRows} rows in {result.Elapsed.TotalSeconds:F1}s");
                }

                if (result.ErrorCount > 0)
                {
                    currentTab = 2; // Switch to Validation tab
                    validationDrawer?.SetReport(result.ValidationReport);
                }
            }
            finally
            {
                EditorUtility.ClearProgressBar();
            }
        }

        public void RunValidate()
        {
            var options = PipelineOptions();
            options.BlockOnValidationError = false; // 校验模式不阻止导出

            try
            {
                EditorUtility.DisplayProgressBar("Excel To JSON", "Validating...", 0f);

                var result = Pipeline.ProcessDirectory(excelDir, options,
                    (current, total, fileName) =>
                    {
                        EditorUtility.DisplayProgressBar(
                            "Excel To JSON — Validating",
                            $"({current}/{total}) {fileName}",
                            (float)current / total);
                    });

                if (result.ErrorCount > 0)
                {
                    SetStatus($"Validation: {result.ErrorCount} errors, {result.WarningCount} warnings", true);
                    currentTab = 2;
                    validationDrawer?.SetReport(result.ValidationReport);
                }
                else
                {
                    SetStatus($"Validation passed: {result.SheetsProcessed} sheets OK");
                }
            }
            finally
            {
                EditorUtility.ClearProgressBar();
            }
        }

        public void RunExportSingle(string filePath, string sheetName)
        {
            SetStatus($"Exporting {Path.GetFileName(filePath)}/{sheetName}...");
            var options = PipelineOptions();
            var result = Pipeline.ProcessFile(Path.Combine(excelDir, filePath), options);
            AssetDatabase.Refresh();

            SetStatus(result.Success
                ? $"Exported {Path.GetFileName(filePath)}: {result.SheetsProcessed} sheets, {result.TotalRows} rows"
                : $"Export error: {result.ErrorCount} errors",
                !result.Success);
        }

        public void RunValidateSingle(string filePath, string sheetName)
        {
            SetStatus($"Validating {Path.GetFileName(filePath)}/{sheetName}...");
            var options = PipelineOptions();
            var result = Pipeline.ProcessFile(Path.Combine(excelDir, filePath), options);

            if (result.ErrorCount > 0)
            {
                currentTab = 2;
                validationDrawer?.SetReport(result.ValidationReport);
            }
            SetStatus(result.ErrorCount > 0
                ? $"Validation: {result.ErrorCount} errors"
                : "Validation passed", result.ErrorCount > 0);
        }

        public void RunExportSelected()
        {
            if (multiSelectedFiles.Count == 0) return;

            var options = PipelineOptions();
            var total = multiSelectedFiles.Count;
            var completed = 0;
            var totalReport = new ValidationReport();

            try
            {
                EditorUtility.DisplayProgressBar("Excel To JSON", "Exporting selected files...", 0f);

                foreach (var relPath in multiSelectedFiles)
                {
                    var fullPath = Path.Combine(excelDir, relPath);
                    completed++;
                    EditorUtility.DisplayProgressBar("Excel To JSON — Export Selected",
                        $"({completed}/{total}) {Path.GetFileName(relPath)}",
                        (float)completed / total);

                    var fileResult = Pipeline.ProcessFile(fullPath, options);
                    totalReport.Errors.AddRange(fileResult.ValidationReport.Errors);
                }

                AssetDatabase.Refresh();
                SetStatus(totalReport.HasErrors
                    ? $"Export selected: {total} files, {totalReport.ErrorCount} errors"
                    : $"Export selected: {total} files OK");
            }
            finally
            {
                EditorUtility.ClearProgressBar();
            }
        }

        public void RunValidateSelected()
        {
            if (multiSelectedFiles.Count == 0) return;

            var options = PipelineOptions();
            options.BlockOnValidationError = false;
            var total = multiSelectedFiles.Count;
            var completed = 0;
            var totalReport = new ValidationReport();

            try
            {
                EditorUtility.DisplayProgressBar("Excel To JSON", "Validating selected files...", 0f);

                foreach (var relPath in multiSelectedFiles)
                {
                    var fullPath = Path.Combine(excelDir, relPath);
                    completed++;
                    EditorUtility.DisplayProgressBar("Excel To JSON — Validate Selected",
                        $"({completed}/{total}) {Path.GetFileName(relPath)}",
                        (float)completed / total);

                    var fileResult = Pipeline.ProcessFile(fullPath, options);
                    totalReport.Errors.AddRange(fileResult.ValidationReport.Errors);
                }

                if (totalReport.ErrorCount > 0)
                {
                    SetStatus($"Validation: {totalReport.ErrorCount} errors, {totalReport.WarningCount} warnings", true);
                    currentTab = 2;
                    validationDrawer?.SetReport(totalReport);
                }
                else
                {
                    SetStatus($"Validation: {total} files OK");
                }
            }
            finally
            {
                EditorUtility.ClearProgressBar();
            }
        }

        public void OpenExcel(string fullPath)
        {
            System.Diagnostics.Process.Start(fullPath);
        }

        private Pipeline.Options PipelineOptions()
        {
            var headerRow = EditorPrefs.GetInt("ExcelToJson.HeaderRow", 1);
            var typeRow = EditorPrefs.GetInt("ExcelToJson.TypeRow", 2);
            var commentRow = EditorPrefs.GetInt("ExcelToJson.CommentRow", 3);
            var dataStartRow = EditorPrefs.GetInt("ExcelToJson.DataStartRow", 4);
            var skipPrefixes = EditorPrefs.GetString("ExcelToJson.SkipPrefixes", "_,#");
            var strictMode = EditorPrefs.GetBool("ExcelToJson.StrictMode", false);
            var genPath = EditorPrefs.GetString("ExcelToJson.GenPath", "Scripts/Generated");
            var genNamespace = EditorPrefs.GetString("ExcelToJson.Namespace", "Game.Data");

            return new Pipeline.Options
            {
                ExcelDir = excelDir,
                OutputDir = outputDir,
                HeaderRow = headerRow,
                TypeRow = typeRow,
                CommentRow = commentRow,
                DataStartRow = dataStartRow,
                CodeGenConfig = new Generator.CodeGenerator.Config
                {
                    RowOutputDir = $"{genPath}/Data",
                    TableOutputDir = $"{genPath}/Tables",
                    Namespace = genNamespace,
                    RowSuffix = "Row",
                    TableSuffix = "Table",
                    AssetsRoot = "Assets",
                },
                AssetGenConfig = new Generator.AssetGenerator.Config
                {
                    OutputDir = "Data",
                    AssetsRoot = "Assets",
                },
                EnableValidation = true,
                BlockOnValidationError = strictMode,
                ExportJson = true,
                JsonConfig = new Pipeline.JsonExportConfig
                {
                    OutputDir = "Assets/Data",
                    PrettyPrint = true,
                    EnsureAscii = false,
                },
                SkipSheetPrefixes = skipPrefixes.Split(new[] { ',' }, System.StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToArray(),
            };
        }

        private string GetRelativePath(string fullPath)
        {
            var full = Path.GetFullPath(fullPath).Replace("\\", "/");
            var base_ = Path.GetFullPath(excelDir).Replace("\\", "/");
            if (full.StartsWith(base_))
                return full.Substring(base_.Length).TrimStart('/');
            return Path.GetFileName(fullPath);
        }

        // ===== 公开属性（供 Drawer 使用） =====
        public string ExcelDir { get => excelDir; set => excelDir = value; }
        public string OutputDir { get => outputDir; set => outputDir = value; }
        public string SelectedFilePath => selectedFilePath;
        public string SelectedSheetName => selectedSheetName;
        public List<ExcelFileEntry> FileEntries => fileEntries;

        public List<List<string>> GetSheetData(string relPath, string sheetName)
        {
            if (string.IsNullOrEmpty(relPath) || string.IsNullOrEmpty(sheetName)) return null;
            try
            {
                var fullPath = Path.Combine(excelDir, relPath);
                var read = ExcelReader.Read(fullPath);
                if (read.Sheets.TryGetValue(sheetName, out var rows))
                    return rows;
            }
            catch { }
            return null;
        }

        public TableSchema GetSchema(string relPath, string sheetName)
        {
            var rows = GetSheetData(relPath, sheetName);
            if (rows == null) return null;
            try
            {
                return SchemaParser.Parse(rows, sheetName,
                    Path.GetFileName(relPath ?? ""), 1, 2, 3, 4,
                    new[] { "_", "#" });
            }
            catch { return null; }
        }
    }

    // ===== 数据模型 =====
    [Serializable]
    public class ExcelFileEntry
    {
        public string RelativePath;
        public string FullPath;
        public string FileName;
        public List<SheetEntry> Sheets = new List<SheetEntry>();
    }

    [Serializable]
    public class SheetEntry
    {
        public string Name;
    }
}
