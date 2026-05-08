using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ExcelToJsonPlugin.Runtime;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Mapping
{
    /// <summary>
    /// 从带 [ExcelTable] / [ExcelColumn] 的 C# 类反向生成 Excel 模板。
    /// 策划拿到模板后填写数据即可。
    /// </summary>
    public static class TemplateExporter
    {
        /// <summary>
        /// Generate an empty Excel template from a C# class with [ExcelTable] attribute.
        /// </summary>
        /// <param name="type">The C# class with [ExcelTable] attribute</param>
        /// <param name="outputDir">Output directory (under Assets/)</param>
        /// <returns>Path to the generated .xlsx file, or null on failure</returns>
        public static string GenerateTemplate(Type type, string outputDir)
        {
            var tableAttr = type.GetCustomAttribute<ExcelTableAttribute>();
            if (tableAttr == null)
            {
                Debug.LogError($"[ExcelToJSON] 类型 {type.Name} 没有 [ExcelTable] 属性");
                return null;
            }

            var sheetName = tableAttr.SheetName;
            var fileName = string.IsNullOrEmpty(tableAttr.FileName)
                ? $"{sheetName}_template"
                : tableAttr.FileName;

            var fields = GetExportableFields(type);
            if (fields.Count == 0)
            {
                Debug.LogError($"[ExcelToJSON] 类型 {type.Name} 没有可导出的字段");
                return null;
            }

            try
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet(sheetName);

                // Create header styles
                var headerStyle = CreateHeaderStyle(workbook);
                var typeStyle = CreateTypeStyle(workbook);
                var commentStyle = CreateCommentStyle(workbook);

                // Row 1: Field names (from [ExcelColumn] or field name itself)
                var nameRow = sheet.CreateRow(0);
                // Row 2: Type declaration
                var typeRow = sheet.CreateRow(1);
                // Row 3: Comments (optional, empty by default)
                var commentRow = sheet.CreateRow(2);

                // Set column widths
                for (int i = 0; i < fields.Count; i++)
                    sheet.SetColumnWidth(i, 16 * 256);

                for (int col = 0; col < fields.Count; col++)
                {
                    var field = fields[col];
                    var colAttr = field.GetCustomAttribute<ExcelColumnAttribute>();
                    var colName = colAttr?.ColumnName ?? field.Name;
                    var excelType = InferExcelType(field.FieldType);

                    // Field name cell
                    var nameCell = nameRow.CreateCell(col);
                    nameCell.SetCellValue(colName);
                    nameCell.CellStyle = headerStyle;

                    // Type cell
                    var typeCell = typeRow.CreateCell(col);
                    typeCell.SetCellValue(excelType);
                    typeCell.CellStyle = typeStyle;

                    // Comment cell (empty)
                    var cmtCell = commentRow.CreateCell(col);
                    cmtCell.SetCellValue("");
                    cmtCell.CellStyle = commentStyle;
                }

                // Save to disk
                var fullOutputDir = Path.Combine("Assets", outputDir, "Templates");
                Directory.CreateDirectory(fullOutputDir);

                var filePath = Path.Combine(fullOutputDir, $"{fileName}.xlsx");
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                workbook.Close();

                AssetDatabase.Refresh();
                Debug.Log($"[ExcelToJSON] Excel 模板已生成: {filePath} (Sheet: {sheetName})");

                return filePath;
            }
            catch (Exception ex)
            {
                Debug.LogError($"[ExcelToJSON] 生成 Excel 模板失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Generate templates for ALL classes with [ExcelTable] attribute.
        /// </summary>
        public static List<string> GenerateAllTemplates(string outputDir)
        {
            var generated = new List<string>();
            var types = AttributeMapping.ScanExcelTableTypes();

            foreach (var kv in types)
            {
                var path = GenerateTemplate(kv.Key, outputDir);
                if (path != null)
                    generated.Add(path);
            }

            return generated;
        }

        // ============================================================
        // Helpers
        // ============================================================

        private static List<FieldInfo> GetExportableFields(Type type)
        {
            var fields = new List<FieldInfo>();
            foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
            {
                if (field.GetCustomAttribute<ExcelIgnoreAttribute>() != null)
                    continue;
                fields.Add(field);
            }
            return fields;
        }

        private static string InferExcelType(Type csType)
        {
            if (csType == typeof(int)) return "int";
            if (csType == typeof(float)) return "float";
            if (csType == typeof(double)) return "float";
            if (csType == typeof(long)) return "int";
            if (csType == typeof(string)) return "string";
            if (csType == typeof(bool)) return "bool";
            if (csType == typeof(int[])) return "int[]";
            if (csType == typeof(float[])) return "float[]";
            if (csType == typeof(string[])) return "string[]";
            if (csType == typeof(Vector2)) return "Vector2";
            if (csType == typeof(Vector3)) return "Vector3";
            if (csType == typeof(Color)) return "Color";
            return "string";
        }

        private static ICellStyle CreateHeaderStyle(XSSFWorkbook wb)
        {
            var style = wb.CreateCellStyle();
            var font = wb.CreateFont();
            font.FontName = "Microsoft YaHei";
            font.FontHeightInPoints = 11;
            font.IsBold = true;
            font.Color = IndexedColors.White.Index;
            style.SetFont(font);
            style.FillForegroundColor = IndexedColors.DarkBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;
            style.Alignment = HorizontalAlignment.Center;
            style.BorderBottom = BorderStyle.Thin;
            return style;
        }

        private static ICellStyle CreateTypeStyle(XSSFWorkbook wb)
        {
            var style = wb.CreateCellStyle();
            var font = wb.CreateFont();
            font.FontName = "Consolas";
            font.FontHeightInPoints = 10;
            font.Color = IndexedColors.Grey50Percent.Index;
            style.SetFont(font);
            style.Alignment = HorizontalAlignment.Center;
            style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
            style.BorderBottom = BorderStyle.Thin;
            return style;
        }

        private static ICellStyle CreateCommentStyle(XSSFWorkbook wb)
        {
            var style = wb.CreateCellStyle();
            var font = wb.CreateFont();
            font.FontName = "Microsoft YaHei";
            font.FontHeightInPoints = 9;
            font.Color = IndexedColors.Grey40Percent.Index;
            font.IsItalic = true;
            style.SetFont(font);
            style.BorderBottom = BorderStyle.Thin;
            return style;
        }
    }
}
