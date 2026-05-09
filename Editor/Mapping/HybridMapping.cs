using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Mapping
{
    /// <summary>
    /// Mode C (Hybrid): 每个 Sheet 独立选择模式 A 或 B。
    /// 同一 Excel 文件中不同 Sheet 可以各自使用不同模式。
    /// </summary>
    public static class HybridMapping
    {
        public enum SheetMode
        {
            /// <summary>Mode A: Auto-generate C# code from Excel</summary>
            ModeA_AutoGenerate,
            /// <summary>Mode B: Use existing C# class with [ExcelTable]</summary>
            ModeB_ReflectMatch,
        }

        /// <summary>
        /// Determine the mode for a given sheet.
        /// If there's a C# class marked with [ExcelTable(sheetName)], use Mode B.
        /// Otherwise, use Mode A.
        /// </summary>
        public static SheetMode DetectMode(string sheetName, Dictionary<string, Type> sheetTypeMap)
        {
            if (sheetTypeMap != null && sheetTypeMap.TryGetValue(sheetName, out var _))
                return SheetMode.ModeB_ReflectMatch;
            return SheetMode.ModeA_AutoGenerate;
        }

        /// <summary>
        /// Reconcile schema fields with C# class fields for partial matching.
        /// Returns the set of columns that should be active for data generation.
        /// - Columns present in both: mapped to C# field
        /// - Columns only in Excel: auto-generated in C# (Mode A fallback)
        /// - Columns only in C#: ignored in data parsing, set to default
        /// </summary>
        public static List<FieldMapping> ReconcileFields(
            TableSchema schema, Type mappedType)
        {
            var result = new List<FieldMapping>();
            var csharpFields = AttributeMapping.GetColumnMapping(mappedType);
            var usedCSharpFields = new HashSet<string>();

            foreach (var excelField in schema.Fields)
            {
                var mapping = new FieldMapping
                {
                    ExcelField = excelField,
                    CSharpField = null,
                };

                if (csharpFields.TryGetValue(excelField.Name, out var csField))
                {
                    mapping.CSharpField = csField;
                    usedCSharpFields.Add(excelField.Name);
                }

                result.Add(mapping);
            }

            return result;
        }
    }

    /// <summary>
    /// Mapping between an Excel column and a C# field.
    /// </summary>
    public class FieldMapping
    {
        public FieldDef ExcelField;
        public FieldInfo CSharpField;
        public bool HasCSharpMapping => CSharpField != null;
    }
}
