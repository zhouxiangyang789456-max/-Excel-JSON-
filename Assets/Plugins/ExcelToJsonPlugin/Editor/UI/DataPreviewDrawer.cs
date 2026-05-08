using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    /// <summary>Tab 1: 数据预览</summary>
    public class DataPreviewDrawer
    {
        private readonly ExcelDataWindow window;
        private Vector2 scrollH, scrollV;
        private string searchQuery = "";
        private int searchResultRow = -1;
        private int pageSize = 100;
        private int currentPage;

        public DataPreviewDrawer(ExcelDataWindow window) { this.window = window; }

        public void Draw(string selectedFile, string selectedSheet, List<ExcelFileEntry> entries)
        {
            if (string.IsNullOrEmpty(selectedSheet))
            {
                EditorGUILayout.HelpBox("Select a sheet from the file tree to preview data.", MessageType.Info);
                return;
            }

            var rows = window.GetSheetData(selectedFile, selectedSheet);
            if (rows == null || rows.Count == 0)
            {
                EditorGUILayout.HelpBox("No data in this sheet.", MessageType.Warning);
                return;
            }

            var schema = window.GetSchema(selectedFile, selectedSheet);

            // 表信息
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField($"{Path.GetFileName(selectedFile)} → {selectedSheet}", EditorStyles.boldLabel);
            EditorGUILayout.LabelField($"({rows.Count} rows, {rows.FirstOrDefault()?.Count ?? 0} cols)", EditorStyles.miniLabel);
            EditorGUILayout.EndHorizontal();

            // 搜索
            EditorGUILayout.BeginHorizontal();
            searchQuery = EditorGUILayout.TextField("Search:", searchQuery);
            if (GUILayout.Button("Find", GUILayout.Width(60)))
                searchResultRow = FindRow(rows, searchQuery);
            if (searchResultRow >= 0)
                EditorGUILayout.LabelField($"Found at row {searchResultRow + 1}", EditorStyles.miniLabel);
            EditorGUILayout.EndHorizontal();

            // 分页
            int totalPages = Mathf.Max(1, Mathf.CeilToInt((float)rows.Count / pageSize));
            if (totalPages > 1)
            {
                EditorGUILayout.BeginHorizontal();
                if (GUILayout.Button("◀", GUILayout.Width(30))) currentPage = Mathf.Max(0, currentPage - 1);
                EditorGUILayout.LabelField($"Page {currentPage + 1} / {totalPages}", EditorStyles.centeredGreyMiniLabel);
                if (GUILayout.Button("▶", GUILayout.Width(30))) currentPage = Mathf.Min(totalPages - 1, currentPage + 1);
                EditorGUILayout.EndHorizontal();
            }
            else currentPage = 0;

            int startRow = currentPage * pageSize;
            int endRow = Mathf.Min(startRow + pageSize, rows.Count);

            // 数据表格
            scrollV = EditorGUILayout.BeginScrollView(scrollV);
            scrollH = EditorGUILayout.BeginScrollView(scrollH);

            int headerRows = schema?.DataStartRow - 1 ?? 3; // 表头行数

            for (int ri = 0; ri < rows.Count; ri++)
            {
                // 分页过滤
                if (ri < startRow || ri >= endRow) continue;

                var row = rows[ri];
                EditorGUILayout.BeginHorizontal();

                // 行号
                var rowLabelStyle = new GUIStyle(EditorStyles.label) { fixedWidth = 40 };
                if (ri < headerRows)
                {
                    var headerStyle = new GUIStyle(EditorStyles.boldLabel) { fixedWidth = 40, normal = { textColor = new Color(0.3f, 0.5f, 0.8f) } };
                    EditorGUILayout.LabelField($"#{ri + 1}", headerStyle);
                }
                else if (searchResultRow == ri)
                {
                    var hlStyle = new GUIStyle(EditorStyles.boldLabel) { fixedWidth = 40, normal = { textColor = Color.yellow } };
                    EditorGUILayout.LabelField($"#{ri + 1}", hlStyle);
                }
                else
                {
                    EditorGUILayout.LabelField($"#{ri + 1}", rowLabelStyle);
                }

                // 单元格
                for (int ci = 0; ci < row.Count; ci++)
                {
                    var cellValue = row[ci];
                    var displayValue = cellValue.Length > 30 ? cellValue.Substring(0, 27) + "..." : cellValue;

                    // 表头用粗体
                    var cellStyle = ri < headerRows
                        ? new GUIStyle(EditorStyles.label) { fontStyle = FontStyle.Bold, normal = { textColor = new Color(0.3f, 0.5f, 0.8f) } }
                        : EditorStyles.label;

                    // 类型行特殊颜色
                    if (ri == 1 && schema != null && ci < schema.Fields.Count)
                    {
                        cellStyle = new GUIStyle(EditorStyles.label)
                        {
                            fontStyle = FontStyle.Italic,
                            normal = { textColor = new Color(0.2f, 0.6f, 0.2f) }
                        };
                    }

                    // Tooltip 显示完整内容
                    var content = new GUIContent(displayValue, cellValue);
                    EditorGUILayout.LabelField(content, cellStyle, GUILayout.MinWidth(60));
                }

                EditorGUILayout.EndHorizontal();
            }

            EditorGUILayout.EndScrollView();
            EditorGUILayout.EndScrollView();

            EditorGUILayout.Space(10);
            EditorGUILayout.LabelField($"Showing {endRow - startRow} of {rows.Count} rows (page {currentPage + 1})", EditorStyles.miniLabel);
        }

        private int FindRow(List<List<string>> rows, string query)
        {
            if (string.IsNullOrEmpty(query)) return -1;
            var lower = query.ToLower();

            for (int ri = 0; ri < rows.Count; ri++)
            {
                foreach (var cell in rows[ri])
                {
                    if (cell.ToLower().Contains(lower))
                        return ri;
                }
            }
            return -1;
        }
    }
}
