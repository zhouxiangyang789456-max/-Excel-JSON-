using System;
using System.Collections.Generic;
using System.IO;
using ExcelToJsonPlugin.Editor.Core;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Watcher
{
    /// <summary>
    /// Monitors Excel directory for changes and auto-triggers Pipeline export.
    /// Uses FileSystemWatcher + debounce.
    /// </summary>
    public class ExcelFileWatcher : IDisposable
    {
        private FileSystemWatcher watcher;
        private readonly Dictionary<string, DateTime> pendingFiles = new Dictionary<string, DateTime>();
        private readonly object lockObj = new object();
        private float debounceMs = 500;
        private string watchDir;
        private bool isRunning;

        /// <summary>Is the watcher currently active?</summary>
        public bool IsRunning => isRunning;

        /// <summary>Current debounce delay in milliseconds</summary>
        public float DebounceMs
        {
            get => debounceMs;
            set => debounceMs = Mathf.Clamp(value, 100, 3000);
        }

        /// <summary>
        /// Start watching the given directory.
        /// </summary>
        public void Start(string directory, float debounceMs = 500)
        {
            if (isRunning) Stop();

            watchDir = Path.GetFullPath(directory);
            if (!Directory.Exists(watchDir))
            {
                Debug.LogWarning($"[ExcelToJSON Watcher] Directory not found: {watchDir}");
                return;
            }

            this.debounceMs = debounceMs;

            try
            {
                watcher = new FileSystemWatcher(watchDir)
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName,
                    Filter = "*.xlsx",
                    IncludeSubdirectories = true,
                    EnableRaisingEvents = true,
                };

                // Watch .xls too
                var xlsWatcher = new FileSystemWatcher(watchDir)
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName,
                    Filter = "*.xls",
                    EnableRaisingEvents = true,
                };

                watcher.Changed += OnFileChanged;
                watcher.Created += OnFileChanged;
                watcher.Renamed += OnFileRenamed;

                xlsWatcher.Changed += OnFileChanged;
                xlsWatcher.Created += OnFileChanged;

                // Register Editor update callback
                EditorApplication.update += OnEditorUpdate;

                isRunning = true;
                Debug.Log($"[ExcelToJSON Watcher] Watching: {watchDir} (debounce: {debounceMs}ms)");
            }
            catch (Exception ex)
            {
                Debug.LogError($"[ExcelToJSON Watcher] Failed to start: {ex.Message}");
            }
        }

        /// <summary>
        /// Stop watching and clean up.
        /// </summary>
        public void Stop()
        {
            if (watcher != null)
            {
                watcher.EnableRaisingEvents = false;
                watcher.Changed -= OnFileChanged;
                watcher.Created -= OnFileChanged;
                watcher.Renamed -= OnFileRenamed;
                watcher.Dispose();
                watcher = null;
            }

            EditorApplication.update -= OnEditorUpdate;
            isRunning = false;

            lock (lockObj)
                pendingFiles.Clear();

            Debug.Log("[ExcelToJSON Watcher] Stopped");
        }

        public void Dispose() => Stop();

        // ============================================================
        // File event handlers
        // ============================================================

        private void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            // Ignore Excel temp files
            var fileName = Path.GetFileName(e.FullPath);
            if (fileName.StartsWith("~$")) return;

            lock (lockObj)
            {
                pendingFiles[e.FullPath] = DateTime.UtcNow;
            }
        }

        private void OnFileRenamed(object sender, RenamedEventArgs e)
        {
            var fileName = Path.GetFileName(e.FullPath);
            if (fileName.StartsWith("~$")) return;

            lock (lockObj)
            {
                pendingFiles[e.FullPath] = DateTime.UtcNow;
            }
        }

        // ============================================================
        // Editor update (debounce check)
        // ============================================================

        private void OnEditorUpdate()
        {
            List<string> readyFiles = null;
            var now = DateTime.UtcNow;

            lock (lockObj)
            {
                foreach (var kv in pendingFiles)
                {
                    var elapsed = (now - kv.Value).TotalMilliseconds;
                    if (elapsed >= debounceMs)
                    {
                        readyFiles ??= new List<string>();
                        readyFiles.Add(kv.Key);
                    }
                }

                if (readyFiles != null)
                {
                    foreach (var file in readyFiles)
                        pendingFiles.Remove(file);
                }
            }

            if (readyFiles != null)
            {
                foreach (var filePath in readyFiles)
                {
                    // Wait a bit for file lock to be released by Excel
                    EditorApplication.delayCall += () =>
                    {
                        ProcessChangedFile(filePath);
                    };
                }
            }
        }

        // ============================================================
        // Process changed file
        // ============================================================

        private void ProcessChangedFile(string filePath)
        {
            if (!File.Exists(filePath)) return;

            var options = CreatePipelineOptions();
            Debug.Log($"[ExcelToJSON Watcher] Auto-export: {Path.GetFileName(filePath)}");

            try
            {
                var result = Pipeline.ProcessFile(filePath, options);
                AssetDatabase.Refresh();

                if (result.ErrorCount > 0)
                {
                    Debug.LogWarning(
                        $"[ExcelToJSON Watcher] {Path.GetFileName(filePath)}: " +
                        $"{result.ErrorCount} errors, {result.WarningCount} warnings");
                }
                else
                {
                    Debug.Log(
                        $"[ExcelToJSON Watcher] {Path.GetFileName(filePath)}: " +
                        $"{result.SheetsProcessed} sheets, {result.TotalRows} rows OK " +
                        $"({result.Elapsed.TotalSeconds:F1}s)");
                }
            }
            catch (Exception ex)
            {
                Debug.LogError($"[ExcelToJSON Watcher] Failed to process {Path.GetFileName(filePath)}: {ex.Message}");
            }
        }

        private Pipeline.Options CreatePipelineOptions()
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
                ExcelDir = (watchDir ?? "").Replace("\\", "/"),
                OutputDir = "Assets/Data",
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
                BlockOnValidationError = strictMode,
                ExportJson = false,
                SkipSheetPrefixes = skipPrefixes.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries),
            };
        }
    }
}
