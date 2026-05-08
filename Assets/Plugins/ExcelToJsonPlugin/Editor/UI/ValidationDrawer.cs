using System.Collections.Generic;
using System.Text;
using ExcelToJsonPlugin.Editor.Core.Models;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    /// <summary>Tab 2: 校验结果</summary>
    public class ValidationDrawer
    {
        private readonly ExcelDataWindow window;
        private ValidationReport currentReport;
        private Vector2 scroll;
        private bool showErrors = true;
        private bool showWarnings = true;
        private bool showInfo = false;

        public ValidationDrawer(ExcelDataWindow window) { this.window = window; }

        public void SetReport(ValidationReport report)
        {
            currentReport = report;
        }

        public void Draw(string selectedFile, string selectedSheet, List<ExcelFileEntry> entries)
        {
            if (currentReport == null || currentReport.Errors.Count == 0)
            {
                EditorGUILayout.HelpBox(
                    "No validation results yet.\n\n" +
                    "Click 'Validate' in the toolbar or right-click a sheet and select 'Validate This Sheet' to run validation.",
                    MessageType.Info);

                if (GUILayout.Button("Run Validation", GUILayout.Height(30)))
                    window.RunValidate();

                return;
            }

            // 摘要
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Validation Results", EditorStyles.boldLabel);
            EditorGUILayout.LabelField($"Error: {currentReport.ErrorCount}", GetCountStyle(Color.red));
            EditorGUILayout.LabelField($"Warn: {currentReport.WarningCount}", GetCountStyle(new Color(1f, 0.7f, 0f)));
            EditorGUILayout.LabelField($"Info: {currentReport.InfoCount}", GetCountStyle(Color.gray));
            EditorGUILayout.EndHorizontal();

            // 过滤器
            EditorGUILayout.BeginHorizontal();
            showErrors = EditorGUILayout.ToggleLeft("Errors", showErrors);
            showWarnings = EditorGUILayout.ToggleLeft("Warnings", showWarnings);
            showInfo = EditorGUILayout.ToggleLeft("Info", showInfo);
            GUILayout.FlexibleSpace();
            if (GUILayout.Button("Copy All", GUILayout.Width(70)))
                CopyToClipboard();
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(5);

            // 错误列表
            scroll = EditorGUILayout.BeginScrollView(scroll);

            foreach (var error in currentReport.Errors)
            {
                if (error.Level == ErrorLevel.Error && !showErrors) continue;
                if (error.Level == ErrorLevel.Warning && !showWarnings) continue;
                if (error.Level == ErrorLevel.Info && !showInfo) continue;

                DrawErrorRow(error);
            }

            EditorGUILayout.EndScrollView();
        }

        private void DrawErrorRow(ValidationError error)
        {
            var icon = error.Level == ErrorLevel.Error ? "❌"
                : error.Level == ErrorLevel.Warning ? "⚠️"
                : "ℹ️";

            EditorGUILayout.BeginHorizontal(EditorStyles.helpBox);

            EditorGUILayout.LabelField(icon, GUILayout.Width(25));

            EditorGUILayout.BeginVertical();
            EditorGUILayout.LabelField(error.Message, EditorStyles.wordWrappedLabel);

            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField($"File: {error.FileName}", EditorStyles.miniLabel);
            EditorGUILayout.LabelField($"Sheet: {error.SheetName}", EditorStyles.miniLabel);
            if (error.Row > 0)
                EditorGUILayout.LabelField($"Row: {error.Row}", EditorStyles.miniLabel);
            if (!string.IsNullOrEmpty(error.ColumnName))
                EditorGUILayout.LabelField($"Col: {error.ColumnName}", EditorStyles.miniLabel);
            EditorGUILayout.LabelField($"Rule: {error.RuleName}", EditorStyles.miniLabel, GUILayout.Width(120));
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.EndVertical();

            // 定位按钮
            if (GUILayout.Button("📍", GUILayout.Width(25), GUILayout.Height(35)))
            {
                var fullPath = System.IO.Path.Combine(window.ExcelDir, error.FileName);
                if (System.IO.File.Exists(fullPath))
                    System.Diagnostics.Process.Start(fullPath);
            }

            EditorGUILayout.EndHorizontal();
        }

        private void CopyToClipboard()
        {
            var sb = new StringBuilder();
            sb.AppendLine("File,Sheet,Row,Column,Value,Rule,Level,Message");
            foreach (var e in currentReport.Errors)
            {
                sb.AppendLine($"{e.FileName},{e.SheetName},{e.Row},{e.ColumnName}," +
                    $"{e.RawValue},{e.RuleName},{e.Level},{e.Message}");
            }
            GUIUtility.systemCopyBuffer = sb.ToString();
            Debug.Log("Validation results copied to clipboard (CSV format)");
        }

        private GUIStyle GetCountStyle(Color color)
        {
            return new GUIStyle(EditorStyles.boldLabel) { normal = { textColor = color } };
        }
    }
}
