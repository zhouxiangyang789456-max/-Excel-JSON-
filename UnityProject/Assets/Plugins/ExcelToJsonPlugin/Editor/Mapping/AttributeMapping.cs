using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Runtime;

namespace ExcelToJsonPlugin.Editor.Mapping
{
    /// <summary>
    /// 模式 B：通过反射扫描 [ExcelTable] / [ExcelColumn] 标签，匹配 C# 类与 Excel Sheet。
    /// Sprint 3 实现完整功能。
    /// </summary>
    public static class AttributeMapping
    {
        /// <summary>
        /// 扫描所有已加载程序集中带 [ExcelTable] 的类，返回类→Sheet名映射。
        /// </summary>
        public static Dictionary<Type, string> ScanExcelTableTypes()
        {
            var result = new Dictionary<Type, string>();

            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    foreach (var type in asm.GetTypes())
                    {
                        var attr = type.GetCustomAttribute<ExcelTableAttribute>();
                        if (attr != null)
                        {
                            result[type] = attr.SheetName;
                        }
                    }
                }
                catch
                {
                    // 跳过无法加载的程序集
                }
            }

            return result;
        }

        /// <summary>
        /// 从 C# 类型中提取字段→Excel列的映射。
        /// </summary>
        public static Dictionary<string, FieldInfo> GetColumnMapping(Type type)
        {
            var mapping = new Dictionary<string, FieldInfo>();

            foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
            {
                if (field.GetCustomAttribute<ExcelIgnoreAttribute>() != null)
                    continue;

                var colAttr = field.GetCustomAttribute<ExcelColumnAttribute>();
                var colName = colAttr?.ColumnName ?? field.Name;
                mapping[colName] = field;
            }

            return mapping;
        }

        /// <summary>
        /// 检查 C# 类型是否兼容给定的 TableSchema。
        /// 返回不兼容的差异列表。
        /// </summary>
        public static List<string> ValidateCompatibility(Type type, TableSchema schema)
        {
            var issues = new List<string>();
            var fieldMap = GetColumnMapping(type);
            var schemaFields = new HashSet<string>();

            foreach (var field in schema.Fields)
            {
                schemaFields.Add(field.Name);
                if (!fieldMap.ContainsKey(field.Name))
                {
                    issues.Add($"Excel 有列 \"{field.Name}\" 但 C# 类 \"{type.Name}\" 中无对应字段");
                }
            }

            foreach (var kv in fieldMap)
            {
                if (!schemaFields.Contains(kv.Key))
                {
                    issues.Add($"C# 类 \"{type.Name}\" 有字段 \"{kv.Value.Name}\" 但 Excel 中无对应列");
                }
            }

            return issues;
        }
    }
}
