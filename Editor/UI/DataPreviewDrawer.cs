using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    /// <summary>Tab 1: Data preview — 预计算单元格样式以消除每帧重复校验</summary>
    public class DataPreviewDrawer
    {
        private readonly ExcelDataWindow window;
        private Vector2 scrollH, scrollV;
        private string searchQuery = "";
        private int searchResultRow = -1;
        private int pageSize = 100;
        private int currentPage;

        // Cached data — only reload when selection changes
        private string cachedSheet;
        private List<List<string>> cachedRows;
        private TableSchema cachedSchema;

        /// <summary>预计算的单元格样式矩阵 [row][col] → GUIStyle，只在数据加载时计算一次</summary>
        private List<List<GUIStyle>> cachedCellStyles;

        // Cached styles — 延迟创建，避免 OnEnable 阶段 EditorStyles 未初始化
        private bool stylesReady;
        private GUIStyle rowNumStyle, rowNumBoldStyle, rowNumHlStyle;
        private GUIStyle cellStyle, cellBoldStyle, cellItalicStyle;
        private GUIStyle cellValidStyle, cellWarnStyle, cellErrorStyle, cellEmptyStyle, cellResStyle;
        private Dictionary<string, GUIStyle> headerColorStyles = new Dictionary<string, GUIStyle>();

        /// <summary>支持 CJK 字符渲染的动态字体（延迟加载，全局缓存）</summary>
        private static Font _cjkFont;
        private static Font CjkFont
        {
            get
            {
                if (_cjkFont == null)
                {
                    // 按优先级尝试系统字体：Windows → macOS → 通用回退
                    _cjkFont = Font.CreateDynamicFontFromOSFont(
                        new[] { "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "SimHei", "STHeiti", "Arial Unicode MS" },
                        12);
                }
                return _cjkFont;
            }
        }

        public DataPreviewDrawer(ExcelDataWindow window)
        {
            this.window = window;
        }

        private void EnsureStyles()
        {
            if (stylesReady) return;
            stylesReady = true;
            BuildCachedStyles();
        }

        private void BuildCachedStyles()
        {
            var cjkFont = CjkFont;

            rowNumStyle = new GUIStyle(EditorStyles.label) { fixedWidth = 40 };
            rowNumBoldStyle = new GUIStyle(EditorStyles.boldLabel) { fixedWidth = 40, normal = { textColor = new Color(0.3f, 0.5f, 0.8f) } };
            rowNumHlStyle = new GUIStyle(EditorStyles.boldLabel) { fixedWidth = 40, normal = { textColor = Color.yellow } };

            cellStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, normal = { textColor = Color.white } };
            cellBoldStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, fontStyle = FontStyle.Bold };
            cellItalicStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, fontStyle = FontStyle.Italic, normal = { textColor = new Color(0.4f, 0.4f, 0.4f) } };
            cellValidStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, normal = { textColor = Color.white } };
            cellWarnStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, normal = { textColor = new Color(0.9f, 0.7f, 0.1f) } };
            cellErrorStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, normal = { textColor = new Color(0.9f, 0.2f, 0.2f) } };
            cellEmptyStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, normal = { textColor = new Color(0.5f, 0.5f, 0.5f) } };
            cellResStyle = new GUIStyle(EditorStyles.label) { font = cjkFont, normal = { textColor = new Color(0.1f, 0.6f, 0.6f) } };

            // Header color by type
            headerColorStyles["int"] = cellBoldStyle.WithColor(new Color(0.2f, 0.4f, 0.9f));
            headerColorStyles["float"] = headerColorStyles["int"];
            headerColorStyles["string"] = cellBoldStyle.WithColor(new Color(0.2f, 0.7f, 0.2f));
            headerColorStyles["bool"] = cellBoldStyle.WithColor(new Color(0.9f, 0.5f, 0.1f));
            headerColorStyles["ref"] = cellBoldStyle.WithColor(new Color(0.6f, 0.2f, 0.8f));
            headerColorStyles["enum"] = cellBoldStyle.WithColor(new Color(0.7f, 0.3f, 0.7f));
            headerColorStyles["res"] = cellBoldStyle.WithColor(new Color(0.1f, 0.6f, 0.6f));
        }

        private void ReloadData(string selectedFile, string selectedSheet)
        {
            var key = selectedFile + "|" + selectedSheet;
            if (key == cachedSheet) return; // data unchanged, skip reload
            cachedSheet = key;
            cachedRows = window.GetSheetData(selectedFile, selectedSheet);
            cachedSchema = window.GetSchema(selectedFile, selectedSheet);

            // 预计算所有单元格样式 — 只在数据加载时执行一次
            PrecomputeCellStyles();
        }

        /// <summary>
        /// 一次性预处理所有单元格的 GUIStyle，后续每帧渲染只做 O(1) 数组查表。
        /// </summary>
        private void PrecomputeCellStyles()
        {
            var rows = cachedRows;
            var schema = cachedSchema;
            cachedCellStyles = null;

            if (rows == null || rows.Count == 0) return;

            int rowCount = rows.Count;
            int headerRows = schema?.DataStartRow - 1 ?? 3;
            int dataStartRow = headerRows;
            int colCount = rows[0].Count;

            cachedCellStyles = new List<List<GUIStyle>>(rowCount);
            for (int ri = 0; ri < rowCount; ri++)
            {
                var row = rows[ri];
                var rowStyles = new List<GUIStyle>(colCount);
                for (int ci = 0; ci < colCount; ci++)
                {
                    GUIStyle st;
                    if (ri < headerRows && schema != null && ci < schema.Fields.Count)
                    {
                        // 表头行：按类型着色
                        var type = GetTypeKey(schema.Fields[ci]);
                        st = headerColorStyles.TryGetValue(type, out var s) ? s : cellBoldStyle;
                    }
                    else if (ri == 1 && schema != null)
                    {
                        // 类型行
                        st = cellItalicStyle;
                    }
                    else if (ri >= dataStartRow && schema != null && ci < schema.Fields.Count)
                    {
                        // 数据行：校验一次并缓存样式
                        var cellValue = ci < row.Count ? row[ci] : "";
                        st = EvaluateCellStyle(cellValue, schema.Fields[ci]);
                    }
                    else
                    {
                        st = cellStyle;
                    }
                    rowStyles.Add(st);
                }
                cachedCellStyles.Add(rowStyles);
            }
        }

        /// <summary>根据字段类型和值直接返回对应的 GUIStyle（无中间结构体）</summary>
        private GUIStyle EvaluateCellStyle(string rawValue, FieldDef field)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
                return cellEmptyStyle;

            var normType = (field.NormalizedType ?? "").Trim().ToLower();
            var value = rawValue.Trim();

            if (normType.StartsWith("res"))
            {
                bool ok = !value.Contains("\\") && !value.StartsWith("/") && !value.EndsWith("/") && !value.Contains("..");
                return ok ? cellResStyle : cellErrorStyle;
            }
            if (normType == "int" || normType.StartsWith("ref:") || normType.StartsWith("enum:"))
            {
                return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)
                    ? cellValidStyle : cellErrorStyle;
            }
            if (normType == "float" || normType == "double")
            {
                return float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
                    ? cellValidStyle : cellWarnStyle;
            }
            if (normType == "bool")
            {
                var v = value.ToLower();
                bool ok = v == "true" || v == "false" || v == "0" || v == "1" || v == "是" || v == "否" || v == "yes" || v == "no" || v == "y" || v == "n";
                return ok ? cellValidStyle : cellErrorStyle;
            }
            return cellValidStyle;
        }

        public void Draw(string selectedFile, string selectedSheet, List<ExcelFileEntry> entries)
        {
            EnsureStyles();
            if (string.IsNullOrEmpty(selectedSheet))
            {
                EditorGUILayout.HelpBox(Loc.Tr("select_sheet_hint"), MessageType.Info);
                return;
            }

            ReloadData(selectedFile, selectedSheet);
            var rows = cachedRows;
            var schema = cachedSchema;
            var cellStyles = cachedCellStyles;

            if (rows == null || rows.Count == 0)
            {
                EditorGUILayout.HelpBox(Loc.Tr("no_data"), MessageType.Warning);
                return;
            }

            EditorGUILayout.Space(3);

            // Search bar
            EditorGUILayout.BeginHorizontal();
            searchQuery = EditorGUILayout.TextField(Loc.Tr("search"), searchQuery);
            if (GUILayout.Button(Loc.Tr("find"), GUILayout.Width(60)))
                searchResultRow = FindRow(rows, searchQuery);
            if (searchResultRow >= 0)
                EditorGUILayout.LabelField(Loc.Tr("found_at_row", searchResultRow + 1), EditorStyles.miniLabel);
            EditorGUILayout.EndHorizontal();

            // Pagination
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

            int headerRows = schema?.DataStartRow - 1 ?? 3;
            int startRow = currentPage * pageSize;
            int endRow = Mathf.Min(startRow + pageSize, rows.Count);
            int colCount = rows.Count > 0 ? rows[0].Count : 0;
            float cellWidth = 80f;
            float totalWidth = 40f + colCount * cellWidth;

            // Data table — use GUI (not GUILayout) inside scrollview for performance
            EditorGUILayout.BeginVertical();
            scrollV = EditorGUILayout.BeginScrollView(scrollV, GUILayout.ExpandHeight(true));
            scrollH = EditorGUILayout.BeginScrollView(scrollH, GUILayout.ExpandWidth(true));

            // Reserve total content area
            int visibleRows = endRow - startRow;
            float rowHeight = EditorGUIUtility.singleLineHeight + 2f;
            GUILayoutUtility.GetRect(new GUIContent(""), GUIStyle.none, GUILayout.Width(totalWidth), GUILayout.Height(visibleRows * rowHeight));

            Rect tableRect = new Rect(0, 0, totalWidth, visibleRows * rowHeight);
            GUI.BeginGroup(tableRect);

            for (int vi = 0; vi < visibleRows; vi++)
            {
                int ri = startRow + vi;
                var row = rows[ri];
                var thisRowStyles = cellStyles?[ri]; // O(1) 数组查表
                float y = vi * rowHeight;
                Rect rowRect = new Rect(0, y, totalWidth, rowHeight);

                // Row background for header rows
                if (ri < headerRows)
                    EditorGUI.DrawRect(rowRect, new Color(0.15f, 0.15f, 0.18f));
                else if (searchResultRow == ri)
                    EditorGUI.DrawRect(rowRect, new Color(0.3f, 0.3f, 0.1f));
                else if (vi % 2 == 1)
                    EditorGUI.DrawRect(rowRect, new Color(0.05f, 0.05f, 0.07f));

                // Row number
                Rect numRect = new Rect(2, y + 1, 38, rowHeight);
                var numStyle = ri < headerRows ? rowNumBoldStyle :
                    (searchResultRow == ri ? rowNumHlStyle : rowNumStyle);
                GUI.Label(numRect, $"#{ri + 1}", numStyle);

                // Cells — 直接查表，零计算
                for (int ci = 0; ci < colCount; ci++)
                {
                    Rect cellRect = new Rect(40 + ci * cellWidth, y + 1, cellWidth - 2, rowHeight);
                    var cellValue = ci < row.Count ? row[ci] : "";
                    var displayValue = cellValue.Length > 25 ? cellValue.Substring(0, 22) + ".." : cellValue;

                    var st = thisRowStyles != null && ci < thisRowStyles.Count
                        ? thisRowStyles[ci] : cellStyle;

                    GUI.Label(cellRect, displayValue, st);

                    // Tooltip for long values
                    if (cellValue.Length > 25)
                    {
                        EditorGUI.LabelField(cellRect, new GUIContent("", cellValue));
                    }
                }
            }

            GUI.EndGroup();
            EditorGUILayout.EndScrollView();
            EditorGUILayout.EndScrollView();
            EditorGUILayout.EndVertical();

            EditorGUILayout.Space(5);
            EditorGUILayout.LabelField(Loc.Tr("showing_rows", endRow - startRow, rows.Count, currentPage + 1), EditorStyles.miniLabel);

            // Legend
            if (schema != null)
            {
                EditorGUILayout.BeginHorizontal();
                DrawColorLegend(Loc.Tr("legend_valid"), new Color(0.2f, 0.7f, 0.2f));
                DrawColorLegend(Loc.Tr("legend_warning"), new Color(0.9f, 0.7f, 0.1f));
                DrawColorLegend(Loc.Tr("legend_error"), new Color(0.9f, 0.2f, 0.2f));
                DrawColorLegend(Loc.Tr("legend_empty"), new Color(0.5f, 0.5f, 0.5f));
                EditorGUILayout.EndHorizontal();
            }
        }

        private static string GetTypeKey(FieldDef field)
        {
            var t = (field.NormalizedType ?? "").ToLower().Trim();
            if (t == "int" || t == "float" || t == "double") return "int";
            if (t == "string" || t == "loc" || t == "json") return "string";
            if (t == "bool") return "bool";
            if (t.StartsWith("ref:")) return "ref";
            if (t.StartsWith("enum:")) return "enum";
            if (t.StartsWith("res")) return "res";
            return "";
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
                    if (cell.ToLower().Contains(lower)) return ri;
            }
            return -1;
        }
    }

    /// <summary>Extension to avoid new GUIStyle for color changes</summary>
    internal static class GuiStyleEx
    {
        public static GUIStyle WithColor(this GUIStyle baseStyle, Color color)
        {
            var s = new GUIStyle(baseStyle);
            s.normal.textColor = color;
            return s;
        }
    }
}
