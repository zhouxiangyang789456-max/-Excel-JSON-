using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Runtime;

namespace ExcelToJsonPlugin.Editor.Mapping
{
    /// <summary>
    /// 模式 B：通过反射扫描 [ExcelTable] / [ExcelColumn] 标签，匹配 C# 类与 Excel Sheet。
    /// </summary>
    public static class AttributeMapping
    {
        /// <summary>
        /// 扫描所有已加载程序集中带 [ExcelTable] 的类，返回类型→Sheet名映射。
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
        /// 扫描类型映射（Sheet名→Type），与 ScanExcelTableTypes 相反。
        /// </summary>
        public static Dictionary<string, Type> ScanTableSheetMap()
        {
            var typeMap = ScanExcelTableTypes();
            var sheetMap = new Dictionary<string, Type>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var kv in typeMap)
            {
                if (!sheetMap.ContainsKey(kv.Value))
                    sheetMap[kv.Value] = kv.Key;
            }
            return sheetMap;
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
        /// 检查 Excel 声明的字段类型是否与 C# 字段类型兼容。
        /// Returns a warning string if incompatible, null if compatible.
        /// </summary>
        public static string CheckTypeCompatibility(string excelType, FieldInfo csharpField)
        {
            var csType = csharpField.FieldType;
            var normType = (excelType ?? "").Trim().ToLower();

            switch (normType)
            {
                case "int":
                    if (csType == typeof(int) || csType == typeof(long)
                        || csType == typeof(float) || csType == typeof(double)
                        || csType == typeof(string))
                        return null;
                    break;
                case "float":
                    if (csType == typeof(float) || csType == typeof(double)
                        || csType == typeof(int) || csType == typeof(string))
                        return null;
                    break;
                case "string":
                case "loc":
                    if (csType == typeof(string)) return null;
                    break;
                case "bool":
                    if (csType == typeof(bool) || csType == typeof(string)) return null;
                    break;
                case "int[]":
                    if (csType == typeof(int[])) return null;
                    break;
                case "float[]":
                    if (csType == typeof(float[]) || csType == typeof(double[])) return null;
                    break;
                case "string[]":
                    if (csType == typeof(string[])) return null;
                    break;
                case "Vector2":
                case "vec2":
                    if (csType == typeof(UnityEngine.Vector2)) return null;
                    break;
                case "Vector3":
                case "vec3":
                    if (csType == typeof(UnityEngine.Vector3)) return null;
                    break;
                case "Color":
                case "color":
                    if (csType == typeof(UnityEngine.Color)) return null;
                    break;
                case "json":
                    if (csType == typeof(string)) return null;
                    break;
                case "res":
                case "res:Prefab":
                case "res:Sprite":
                case "res:Material":
                case "res:Texture":
                case "res:AudioClip":
                    if (csType == typeof(string)) return null;
                    break;
                default:
                    // ref:xxx / enum:xxx → C# should be int
                    if (normType.StartsWith("ref:") || normType.StartsWith("enum:"))
                    {
                        if (csType == typeof(int) || csType == typeof(string)) return null;
                    }
                    break;
            }

            return $"类型不匹配: Excel 声明为 \"{excelType}\" 但 C# 字段类型为 \"{csType.Name}\"";
        }

        /// <summary>
        /// 检查 C# 类型是否兼容给定的 TableSchema。
        /// 返回不兼容的差异列表（ValidationError 级别）。
        /// </summary>
        public static List<ValidationError> ValidateCompatibility(
            Type type, TableSchema schema, string fileName)
        {
            var issues = new List<ValidationError>();
            var fieldMap = GetColumnMapping(type);
            var schemaFieldNames = new HashSet<string>();
            var schemaFields = new Dictionary<string, FieldDef>();

            foreach (var field in schema.Fields)
            {
                schemaFieldNames.Add(field.Name);
                schemaFields[field.Name] = field;
            }

            // Excel has column, C# class doesn't → Warning (missing column in class)
            foreach (var field in schema.Fields)
            {
                if (!fieldMap.ContainsKey(field.Name))
                {
                    issues.Add(new ValidationError
                    {
                        FileName = fileName,
                        SheetName = schema.TableName,
                        Row = 0,
                        ColumnName = field.Name,
                        RuleName = "AttributeMap",
                        Message = $"Excel 有列 \"{field.Name}\" 但 C# 类 \"{type.Name}\" 中无对应字段（模式 B 跳过此列）",
                        Level = ErrorLevel.Warning,
                    });
                }
            }

            // C# class has field, Excel doesn't → Warning
            foreach (var kv in fieldMap)
            {
                if (!schemaFieldNames.Contains(kv.Key))
                {
                    issues.Add(new ValidationError
                    {
                        FileName = fileName,
                        SheetName = schema.TableName,
                        Row = 0,
                        ColumnName = kv.Key,
                        RuleName = "AttributeMap",
                        Message = $"C# 类 \"{type.Name}\" 有字段 \"{kv.Value.Name}\" 但 Excel 中无对应列",
                        Level = ErrorLevel.Warning,
                    });
                }
            }

            // Type compatibility check for matching columns
            foreach (var field in schema.Fields)
            {
                if (fieldMap.TryGetValue(field.Name, out var csField))
                {
                    var compatIssue = CheckTypeCompatibility(field.RawType, csField);
                    if (compatIssue != null)
                    {
                        issues.Add(new ValidationError
                        {
                            FileName = fileName,
                            SheetName = schema.TableName,
                            Row = 0,
                            ColumnName = field.Name,
                            RuleName = "AttributeMap",
                            Message = compatIssue,
                            Level = ErrorLevel.Warning,
                        });
                    }
                }
            }

            return issues;
        }
    }
}
