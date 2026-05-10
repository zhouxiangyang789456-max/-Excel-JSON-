using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Editor.Generator;
using ExcelToJsonPlugin.Editor.Mapping;
using ExcelToJsonPlugin.Editor.Watcher;
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

        // ===== 数据缓存 —— 避免每帧重新读取 Excel 文件 =====
        private readonly Dictionary<string, ExcelReader.ReadResult> _readResultCache = new Dictionary<string, ExcelReader.ReadResult>();
        private readonly Dictionary<string, DateTime> _readResultMtime = new Dictionary<string, DateTime>(); // 文件修改时间
        private readonly Dictionary<string, TableSchema> _schemaCache = new Dictionary<string, TableSchema>();
        private Dictionary<string, Type> _cachedSheetTypeMap; // ScanTableSheetMap 缓存

        // Tab 内容
        private DataPreviewDrawer previewDrawer;
        private ValidationDrawer validationDrawer;
        private ExportDrawer exportDrawer;
        private RuntimeDrawer runtimeDrawer;
        private DashboardDrawer dashboardDrawer;

        // File watcher (static to persist between window open/close)
        private static ExcelFileWatcher fileWatcher;
        private DateTime? lastExportTime;
        private bool isExporting;

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
            StartFileWatcherIfEnabled();
        }

        private void OnDisable()
        {
            // Do not stop watcher here — it persists via static reference
        }

        private void OnDestroy()
        {
            // Only stop on final destroy
        }

        private void OnFocus()
        {
            RefreshFileList();
            StartFileWatcherIfEnabled();
        }

        // ===== 刷新文件列表 =====
        public void RefreshFileList()
        {
            InvalidateCache();
            fileEntries.Clear();
            if (!Directory.Exists(excelDir))
            {
                // Only auto-create default relative dirs under the project (e.g. Assets/Excel).
                // Do not mkdir arbitrary names like "配置表" when an external folder pick was mis-resolved.
                if (Path.IsPathRooted(excelDir))
                {
                    UnityEngine.Debug.LogWarning($"[ExcelToJSON] Excel directory not found: {excelDir}");
                    return;
                }

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

        private void InvalidateCache()
        {
            _readResultCache.Clear();
            _readResultMtime.Clear();
            _schemaCache.Clear();
            _cachedSheetTypeMap = null;
        }

        // ===== 绘制 =====
        private void OnGUI()
        {
            HandleKeyboardShortcuts();
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
            DrawSheetInfoBar(); // Sheet info + mapping mode selector
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

            GUILayout.Label(Loc.Tr("window_title"), EditorStyles.boldLabel);

            GUILayout.FlexibleSpace();

            if (GUILayout.Button(Loc.Tr("toolbar_refresh"), EditorStyles.toolbarButton, GUILayout.Width(55)))
                RefreshFileList();

            if (GUILayout.Button(Loc.Tr("toolbar_export_all"), EditorStyles.toolbarButton, GUILayout.Width(65)))
                RunExportAll();

            if (GUILayout.Button(Loc.Tr("toolbar_validate"), EditorStyles.toolbarButton, GUILayout.Width(55)))
                RunValidate();

            if (GUILayout.Button(Loc.Tr("toolbar_template"), EditorStyles.toolbarButton, GUILayout.Width(60)))
            {
                var paths = Mapping.TemplateExporter.GenerateAllTemplates("Excel");
                var count = paths?.Count ?? 0;
                SetStatus(Loc.Tr("template_generated", count));
                RefreshFileList();
            }

            if (isExporting && GUILayout.Button("取消导出", EditorStyles.toolbarButton, GUILayout.Width(70)))
            {
                Core.Pipeline.CancelRequested = true;
                SetStatus("正在取消...");
            }

            GUILayout.Space(5);

            if (GUILayout.Button(EditorGUIUtility.IconContent("_Popup"), EditorStyles.toolbarButton, GUILayout.Width(30)))
                SettingsWindow.ShowWindow();

            EditorGUILayout.EndHorizontal();
        }

        // ===== 文件树 =====
        private void DrawFileTree()
        {
            EditorGUILayout.LabelField(Loc.Tr("files"), EditorStyles.boldLabel);

            searchFilter = EditorGUILayout.TextField(searchFilter, EditorStyles.toolbarSearchField);
            if (!string.IsNullOrEmpty(searchFilter))
                EditorGUILayout.LabelField($"  {Loc.Tr("filter")}: \"{searchFilter}\"", EditorStyles.miniLabel);

            // Multi-select controls
            if (multiSelectedFiles.Count > 0)
            {
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField(Loc.Tr("selected_count", multiSelectedFiles.Count), EditorStyles.miniLabel);
                if (GUILayout.Button(Loc.Tr("export_selected"), EditorStyles.miniButton, GUILayout.Width(90)))
                    RunExportSelected();
                if (GUILayout.Button(Loc.Tr("validate_selected"), EditorStyles.miniButton, GUILayout.Width(100)))
                    RunValidateSelected();
                if (GUILayout.Button(Loc.Tr("clear"), EditorStyles.miniButton, GUILayout.Width(40)))
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
                    menu.AddItem(new GUIContent(Loc.Tr("export_this_file")), false, () => RunExportSingle(entry.RelativePath, null));
                    menu.AddItem(new GUIContent(Loc.Tr("validate_this_file")), false, () => RunValidateSingle(entry.RelativePath, null));
                    menu.AddItem(new GUIContent(Loc.Tr("open_excel")), false, () => OpenExcel(entry.FullPath));
                    menu.AddItem(new GUIContent(Loc.Tr("add_to_selection")), false, () => multiSelectedFiles.Add(entry.RelativePath));
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
                            menu.AddItem(new GUIContent(Loc.Tr("export_this_sheet")), false, () => RunExportSingle(entry.RelativePath, sheet.Name));
                            menu.AddItem(new GUIContent(Loc.Tr("validate_this_sheet")), false, () => RunValidateSingle(entry.RelativePath, sheet.Name));
                            menu.AddItem(new GUIContent(Loc.Tr("open_excel")), false, () => OpenExcel(entry.FullPath));
                            menu.ShowAsContext();
                            Event.current.Use();
                        }
                    }
                    EditorGUI.indentLevel--;
                }
            }

            EditorGUILayout.EndScrollView();

            EditorGUILayout.Space(5);
            if (GUILayout.Button(Loc.Tr("add_excel_dir")))
            {
                var newDir = EditorUtility.OpenFolderPanel("Select Excel Directory", "Assets", "");
                if (!string.IsNullOrEmpty(newDir))
                {
                    // Store absolute normalized path. OpenFolderPanel returns a full OS path; using the old
                    // GetRelativePath(picker) collapsed external folders to Path.GetFileName only (e.g. "配置表"),
                    // which broke scanning Desktop/project-external Excel dirs.
                    excelDir = NormalizeDirectoryPath(newDir);
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
        private static string[] GetTabNames() => new[]
        {
            Loc.Tr("tab_dashboard"), Loc.Tr("tab_data_preview"), Loc.Tr("tab_validation"),
            Loc.Tr("tab_export"), Loc.Tr("tab_runtime"),
        };

        private void DrawTabBar()
        {
            currentTab = GUILayout.Toolbar(currentTab, GetTabNames());
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

        // ===== Sheet Info Bar =====
        private void DrawSheetInfoBar()
        {
            if (string.IsNullOrEmpty(selectedFilePath) || string.IsNullOrEmpty(selectedSheetName))
                return;

            var schema = GetSchema(selectedFilePath, selectedSheetName);
            var rows = GetSheetData(selectedFilePath, selectedSheetName);

            EditorGUILayout.BeginVertical(EditorStyles.helpBox);
            EditorGUILayout.BeginHorizontal();

            // File + Sheet info
            var fileInfo = $"{Path.GetFileName(selectedFilePath)} → {selectedSheetName}";
            var rowCount = rows?.Count ?? 0;
            var colCount = schema?.Fields?.Count ?? 0;
            EditorGUILayout.LabelField($"{fileInfo}  ({rowCount} rows, {colCount} cols)", EditorStyles.boldLabel);

            GUILayout.FlexibleSpace();

            // Mapping mode selector
            EditorGUILayout.LabelField(Loc.Tr("mapping_mode"), GUILayout.Width(60));
            var prefKey = $"ExcelToJson.SheetMode.{selectedSheetName}";
            var currentMode = EditorPrefs.GetInt(prefKey, 0);
            var modeNames = new[] { Loc.Tr("mode_auto"), Loc.Tr("mode_a"), Loc.Tr("mode_b") };
            var newMode = EditorGUILayout.Popup(currentMode, modeNames, GUILayout.Width(100));
            if (newMode != currentMode)
            {
                EditorPrefs.SetInt(prefKey, newMode);
                Debug.Log($"[ExcelToJSON] {selectedSheetName}: 映射模式切换为 {(newMode == 0 ? "Auto" : newMode == 1 ? "Mode A" : "Mode B")}");
            }

            // Target class display
            var targetClass = GetTargetClassDisplay(currentMode, selectedSheetName);
            if (!string.IsNullOrEmpty(targetClass))
            {
                EditorGUILayout.LabelField(Loc.Tr("target_class"), GUILayout.Width(45));
                EditorGUILayout.LabelField(targetClass, EditorStyles.miniLabel);
            }

            EditorGUILayout.EndHorizontal();
            EditorGUILayout.EndVertical();
        }

        private string GetTargetClassDisplay(int mode, string sheetName)
        {
            if (mode == 1) // Mode A forced
                return $"{CodeGenerator.ToPascalCaseStatic(sheetName)}Row (auto-generated)";

            // 缓存反射扫描结果，避免每帧遍历所有程序集
            if (_cachedSheetTypeMap == null)
                _cachedSheetTypeMap = Mapping.AttributeMapping.ScanTableSheetMap();

            if (_cachedSheetTypeMap.TryGetValue(sheetName, out var mappedType))
                return $"{mappedType.Name} (C# class)";

            if (mode == 2) // Mode B forced but no class found
                return "(no [ExcelTable] class found)";

            return $"{CodeGenerator.ToPascalCaseStatic(sheetName)}Row (auto-generated)";
        }

        // ===== 快捷键 =====
        private void HandleKeyboardShortcuts()
        {
            var e = Event.current;
            if (e.type != EventType.KeyDown) return;

            // Ctrl+E: Export selected file
            if (e.control && e.keyCode == KeyCode.E && !e.shift)
            {
                if (!string.IsNullOrEmpty(selectedFilePath))
                    RunExportSingle(selectedFilePath, selectedSheetName);
                e.Use();
            }
            // Ctrl+Shift+E: Export all
            else if (e.control && e.shift && e.keyCode == KeyCode.E)
            {
                RunExportAll();
                e.Use();
            }
            // Ctrl+V: Validate
            else if (e.control && e.keyCode == KeyCode.V)
            {
                RunValidate();
                e.Use();
            }
            // Ctrl+R or F5: Refresh
            else if ((e.control && e.keyCode == KeyCode.R) || e.keyCode == KeyCode.F5)
            {
                RefreshFileList();
                e.Use();
            }
        }

        // ===== 状态栏 =====
        private void DrawStatusBar()
        {
            EditorGUILayout.BeginHorizontal(EditorStyles.toolbar);
            var style = statusIsError ? new GUIStyle(EditorStyles.label) { normal = { textColor = Color.red } }
                                      : EditorStyles.label;
            GUILayout.Label(string.IsNullOrEmpty(statusMessage) ? Loc.Tr("ready") : statusMessage, style);

            GUILayout.FlexibleSpace();

            if (lastExportTime.HasValue)
            {
                EditorGUILayout.LabelField(
                    $"{Loc.Tr("last_export")} {lastExportTime.Value:HH:mm:ss}",
                    EditorStyles.miniLabel, GUILayout.Width(120));
            }

            // Keyboard shortcut hint
            EditorGUILayout.LabelField(
                Loc.Tr("shortcut_hint"),
                EditorStyles.miniLabel, GUILayout.Width(220));

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
            Core.Pipeline.CancelRequested = false;
            isExporting = true;
            var options = PipelineOptions();

            try
            {
                EditorUtility.DisplayProgressBar("Excel To JSON", "Starting export...", 0f);

                var result = Pipeline.ProcessDirectory(excelDir, options,
                    (current, total, fileName) =>
                    {
                        if (Core.Pipeline.CancelRequested)
                            EditorUtility.ClearProgressBar();
                        var pct = (float)current / total;
                        EditorUtility.DisplayProgressBar(
                            "Excel To JSON — Exporting",
                            $"({current}/{total}) {fileName}",
                            pct);
                    });

                if (!Application.isBatchMode)
                    AssetDatabase.Refresh();

                if (result.Success)
                {
                    lastExportTime = DateTime.Now;
                    SetStatus(Loc.Tr("status_export_complete", result.FilesProcessed, result.SheetsProcessed, result.TotalRows, result.Elapsed.TotalSeconds));
                }
                else
                    SetStatus(Loc.Tr("status_export_error", result.ErrorCount, result.WarningCount), true);

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
                isExporting = false;
                Repaint();
            }
        }

        public void RunValidate()
        {
            Core.Pipeline.CancelRequested = false;
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
                    SetStatus(Loc.Tr("status_validate_error", result.ErrorCount, result.WarningCount), true);
                    currentTab = 2;
                    validationDrawer?.SetReport(result.ValidationReport);
                }
                else
                {
                    SetStatus(Loc.Tr("status_validate_pass", result.SheetsProcessed));
                }
            }
            finally
            {
                EditorUtility.ClearProgressBar();
                isExporting = false;
                Repaint();
            }
        }

        public void RunExportSingle(string filePath, string sheetName)
        {
            SetStatus(Loc.Tr("status_exporting", Path.GetFileName(filePath)) + "...");
            var options = PipelineOptions();
            var result = Pipeline.ProcessFile(Path.Combine(excelDir, filePath), options);
            AssetDatabase.Refresh();

            if (result.Success)
                lastExportTime = DateTime.Now;

            SetStatus(result.Success
                ? Loc.Tr("status_export_complete", 1, result.SheetsProcessed, result.TotalRows, 0)
                : Loc.Tr("status_export_error", result.ErrorCount, result.WarningCount),
                !result.Success);
        }

        public void RunValidateSingle(string filePath, string sheetName)
        {
            SetStatus(Loc.Tr("status_exporting", Path.GetFileName(filePath)) + "...");
            var options = PipelineOptions();
            var result = Pipeline.ProcessFile(Path.Combine(excelDir, filePath), options);

            if (result.ErrorCount > 0)
            {
                currentTab = 2;
                validationDrawer?.SetReport(result.ValidationReport);
            }
            SetStatus(result.ErrorCount > 0
                ? Loc.Tr("status_validate_error", result.ErrorCount, result.WarningCount)
                : Loc.Tr("status_validate_pass", 0), result.ErrorCount > 0);
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
                isExporting = false;
                Repaint();
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
                isExporting = false;
                Repaint();
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

        /// <summary>Normalize a directory from OpenFolderPanel / user input for consistent scanning.</summary>
        private static string NormalizeDirectoryPath(string path)
        {
            var full = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return full.Replace("\\", "/");
        }

        private string GetRelativePath(string fullPath)
        {
            var full = Path.GetFullPath(fullPath).Replace("\\", "/");
            var base_ = Path.GetFullPath(excelDir).Replace("\\", "/");
            if (full.StartsWith(base_, StringComparison.OrdinalIgnoreCase))
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
                var lastWrite = File.GetLastWriteTimeUtc(fullPath);
                // 如果文件在缓存后被修改过，清除此文件的缓存
                if (_readResultCache.TryGetValue(fullPath, out var readResult))
                {
                    if (_readResultMtime.TryGetValue(fullPath, out var cachedMtime) && cachedMtime >= lastWrite)
                        return readResult.Sheets.TryGetValue(sheetName, out var r) ? r : null;
                    // 文件已更新，清除缓存
                    _readResultCache.Remove(fullPath);
                    _readResultMtime.Remove(fullPath);
                    CleanSchemaCacheForFile(fullPath);
                }
                // 重新读取
                readResult = ExcelReader.Read(fullPath);
                _readResultCache[fullPath] = readResult;
                _readResultMtime[fullPath] = lastWrite;
                if (readResult.Sheets.TryGetValue(sheetName, out var rows))
                    return rows;
            }
            catch { }
            return null;
        }

        private void CleanSchemaCacheForFile(string fullPath)
        {
            var prefix = fullPath + "|";
            var keys = new List<string>();
            foreach (var k in _schemaCache.Keys)
                if (k.StartsWith(prefix)) keys.Add(k);
            foreach (var k in keys)
                _schemaCache.Remove(k);
        }

        public TableSchema GetSchema(string relPath, string sheetName)
        {
            if (string.IsNullOrEmpty(relPath) || string.IsNullOrEmpty(sheetName)) return null;
            var fullPath = Path.Combine(excelDir, relPath);
            var cacheKey = fullPath + "|" + sheetName;
            // Schema 缓存
            if (_schemaCache.TryGetValue(cacheKey, out var cached))
                return cached;
            try
            {
                // 直接从缓存获取 ReadResult，避免重复读文件
                if (!_readResultCache.TryGetValue(fullPath, out var readResult))
                {
                    readResult = ExcelReader.Read(fullPath);
                    _readResultCache[fullPath] = readResult;
                }
                if (!readResult.Sheets.TryGetValue(sheetName, out var rows))
                    return null;
                var schema = SchemaParser.Parse(rows, sheetName,
                    Path.GetFileName(relPath ?? ""), 1, 2, 3, 4,
                    new[] { "_", "#" });
                _schemaCache[cacheKey] = schema;
                return schema;
            }
            catch { return null; }
        }

        // ===== File Watcher =====
        private void StartFileWatcherIfEnabled()
        {
            var enableAutoExport = EditorPrefs.GetBool("ExcelToJson.AutoExport", false);
            var debounceMs = EditorPrefs.GetInt("ExcelToJson.DebounceMs", 500);

            var fullPath = Path.GetFullPath(excelDir);

            if (enableAutoExport)
            {
                if (fileWatcher == null)
                {
                    fileWatcher = new ExcelFileWatcher();
                    fileWatcher.Start(fullPath, debounceMs);
                }
                else if (!fileWatcher.IsRunning)
                {
                    fileWatcher.Start(fullPath, debounceMs);
                }
            }
            else if (fileWatcher != null && fileWatcher.IsRunning)
            {
                fileWatcher.Stop();
            }
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
