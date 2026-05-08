using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    /// <summary>Tab 1: Data preview with inline validation coloring</summary>
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
                EditorGUILayout.HelpBox(Loc.Tr("select_sheet_hint"), MessageType.Info);
                return;
            }

            var rows = window.GetSheetData(selectedFile, selectedSheet);
            if (rows == null || rows.Count == 0)
            {
                EditorGUILayout.HelpBox(Loc.Tr("no_data"), MessageType.Warning);
                return;
            }

            var schema = window.GetSchema(selectedFile, selectedSheet);

            // 表信息 (handled by sheet info bar now)
            EditorGUILayout.Space(3);

            // 搜索
            EditorGUILayout.BeginHorizontal();
            searchQuery = EditorGUILayout.TextField(Loc.Tr("search"), searchQuery);
            if (GUILayout.Button(Loc.Tr("find"), GUILayout.Width(60)))
                searchResultRow = FindRow(rows, searchQuery);
            if (searchResultRow >= 0)
                EditorGUILayout.LabelField(Loc.Tr("found_at_row", searchResultRow + 1), EditorStyles.miniLabel);
            EditorGUILayout.EndHorizontal();

            // 分页
            int totalPages = Mathf.Max(1, Mathf.CeilToInt((float)rows.Count / pageSize));
            if (totalPages > 1)
            {
                EditorGUILayout.BeginHorizontal();
                if (GUILayout.Button("◀", GUILayout.Width(30))) currentPage = Mathf.Max(0, currentPage - 1);
                EditorGUILayout.LabelField($"{Loc.Tr("page")} {currentPage + 1} / {totalPages}", EditorStyles.centeredGreyMiniLabel);
                if (GUILayout.Button("▶", GUILayout.Width(30))) currentPage = Mathf.Min(totalPages - 1, currentPage + 1);
                EditorGUILayout.EndHorizontal();
            }
            else currentPage = 0;

            int startRow = currentPage * pageSize;
            int endRow = Mathf.Min(startRow + pageSize, rows.Count);

            // 数据表格
            scrollV = EditorGUILayout.BeginScrollView(scrollV);
            scrollH = EditorGUILayout.BeginScrollView(scrollH);

            int headerRows = schema?.DataStartRow - 1 ?? 3;
            var dataStartRow = headerRows; // data starts at the row after headers

            for (int ri = 0; ri < rows.Count; ri++)
            {
                if (ri < startRow || ri >= endRow) continue;

                var row = rows[ri];
                EditorGUILayout.BeginHorizontal();

                // 行号
                var rowLabelStyle = new GUIStyle(EditorStyles.label) { fixedWidth = 40 };
                if (ri < headerRows)
                {
                    var hdrStyle = new GUIStyle(EditorStyles.boldLabel) { fixedWidth = 40, normal = { textColor = new Color(0.3f, 0.5f, 0.8f) } };
                    EditorGUILayout.LabelField($"#{ri + 1}", hdrStyle);
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

                    GUIStyle cellStyle;
                    Color? cellColor = null;

                    if (ri < headerRows && schema != null && ci < schema.Fields.Count)
                    {
                        // Header rows: colored by type
                        var field = schema.Fields[ci];
                        var typeColor = GetTypeColor(field);
                        cellStyle = ri == 0
                            ? new GUIStyle(EditorStyles.label) { fontStyle = FontStyle.Bold, normal = { textColor = typeColor } }
                            : new GUIStyle(EditorStyles.label) { fontStyle = FontStyle.Italic, normal = { textColor = new Color(0.4f, 0.4f, 0.4f) } };
                    }
                    else if (ri >= dataStartRow && schema != null && ci < schema.Fields.Count)
                    {
                        // Data rows: inline validation coloring
                        var field = schema.Fields[ci];
                        var status = ValidateCell(cellValue, field);
                        cellColor = status.Color;
                        cellStyle = new GUIStyle(EditorStyles.label) { normal = { textColor = status.Color } };

                        // Add status icon prefix for resource types
                        if (status.Icon != null)
                            displayValue = status.Icon + " " + displayValue;
                    }
                    else if (ri < headerRows)
                    {
                        cellStyle = new GUIStyle(EditorStyles.label) { fontStyle = FontStyle.Bold, normal = { textColor = new Color(0.3f, 0.5f, 0.8f) } };
                    }
                    else
                    {
                        cellStyle = EditorStyles.label;
                    }

                    var tooltip = cellColor.HasValue && cellColor.Value != Color.white
                        ? $"⚠ Validation: {cellValue}" : cellValue;
                    var content = new GUIContent(displayValue, tooltip);
                    EditorGUILayout.LabelField(content, cellStyle, GUILayout.MinWidth(60));
                }

                EditorGUILayout.EndHorizontal();
            }

            EditorGUILayout.EndScrollView();
            EditorGUILayout.EndScrollView();

            EditorGUILayout.Space(5);
            EditorGUILayout.LabelField(Loc.Tr("showing_rows", endRow - startRow, rows.Count, currentPage + 1), EditorStyles.miniLabel);

            // Legend
            if (schema != null)
            {
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField("", GUILayout.Width(5));
                DrawColorLegend(Loc.Tr("legend_valid"), new Color(0.2f, 0.7f, 0.2f));
                DrawColorLegend(Loc.Tr("legend_warning"), new Color(0.9f, 0.7f, 0.1f));
                DrawColorLegend(Loc.Tr("legend_error"), new Color(0.9f, 0.2f, 0.2f));
                DrawColorLegend(Loc.Tr("legend_empty"), new Color(0.5f, 0.5f, 0.5f));
                EditorGUILayout.EndHorizontal();
            }
        }

        /// <summary>
        /// Quick inline validation result for a single cell.
        /// </summary>
        private struct CellStatus
        {
            public Color Color;
            public string Icon;
        }

        private static CellStatus ValidateCell(string rawValue, FieldDef field)
        {
            // Empty cell → neutral
            if (string.IsNullOrWhiteSpace(rawValue))
                return new CellStatus { Color = new Color(0.5f, 0.5f, 0.5f), Icon = null };

            var normType = (field.NormalizedType ?? "").Trim().ToLower();
            var value = rawValue.Trim();

            // Resource types: just show if has a value
            if (normType.StartsWith("res"))
            {
                var hasValidFormat = !value.Contains("\\") && !value.StartsWith("/") && !value.EndsWith("/") && !value.Contains("..");
                return hasValidFormat
                    ? new CellStatus { Color = new Color(0.1f, 0.6f, 0.6f), Icon = null }
                    : new CellStatus { Color = new Color(0.9f, 0.2f, 0.2f), Icon = null };
            }

            // int
            if (normType == "int" || (normType.StartsWith("ref:") || normType.StartsWith("enum:")))
            {
                return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)
                    ? new CellStatus { Color = Color.white, Icon = null }
                    : new CellStatus { Color = new Color(0.9f, 0.2f, 0.2f), Icon = null };
            }

            // float
            if (normType == "float" || normType == "double")
            {
                return float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
                    ? new CellStatus { Color = Color.white, Icon = null }
                    : new CellStatus { Color = new Color(0.9f, 0.7f, 0.1f), Icon = null };
            }

            // bool
            if (normType == "bool")
            {
                var v = value.ToLower();
                var valid = v == "true" || v == "false" || v == "0" || v == "1" || v == "是" || v == "否" || v == "yes" || v == "no";
                return valid
                    ? new CellStatus { Color = Color.white, Icon = null }
                    : new CellStatus { Color = new Color(0.9f, 0.2f, 0.2f), Icon = null };
            }

            // int array
            if (normType.StartsWith("int["))
            {
                var parts = value.Split(',', '|');
                foreach (var p in parts)
                {
                    var t = p.Trim().Trim('[', ']', '"');
                    if (!string.IsNullOrEmpty(t) && !int.TryParse(t, out _))
                        return new CellStatus { Color = new Color(0.9f, 0.2f, 0.2f), Icon = null };
                }
                return new CellStatus { Color = Color.white, Icon = null };
            }

            // Vector types
            if (normType == "vector2" || normType == "vec2" || normType == "vector3" || normType == "vec3")
            {
                var parts = value.Trim('[', ']').Split(',');
                int validCount = 0;
                foreach (var p in parts)
                {
                    if (float.TryParse(p.Trim(), out _)) validCount++;
                }
                return validCount >= 2
                    ? new CellStatus { Color = Color.white, Icon = null }
                    : new CellStatus { Color = new Color(0.9f, 0.2f, 0.2f), Icon = null };
            }

            // Default: text → valid
            return new CellStatus { Color = Color.white, Icon = null };
        }

        private static void DrawColorLegend(string text, Color color)
        {
            var s = new GUIStyle(EditorStyles.miniLabel) { normal = { textColor = color } };
            EditorGUILayout.LabelField(text, s, GUILayout.Width(80));
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

        private static Color GetTypeColor(FieldDef field)
        {
            var type = (field.NormalizedType ?? "").ToLower().Trim();
            if (type == "int" || type == "float" || type == "double")
                return new Color(0.2f, 0.4f, 0.9f);
            if (type == "string" || type == "loc" || type == "json")
                return new Color(0.2f, 0.7f, 0.2f);
            if (type == "bool")
                return new Color(0.9f, 0.5f, 0.1f);
            if (type.StartsWith("ref:"))
                return new Color(0.6f, 0.2f, 0.8f);
            if (type.StartsWith("enum:"))
                return new Color(0.7f, 0.3f, 0.7f);
            if (type.StartsWith("res"))
                return new Color(0.1f, 0.6f, 0.6f);
            if (type == "vector2" || type == "vec2" || type == "vector3" || type == "vec3")
                return new Color(0.7f, 0.4f, 0.1f);
            return new Color(0.3f, 0.3f, 0.3f);
        }
    }
}
