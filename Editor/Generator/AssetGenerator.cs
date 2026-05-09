using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Runtime;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Generator
{
    /// <summary>
    /// 将解析后的数据生成 ScriptableObject .asset 文件。
    /// 使用反射动态创建 Row 对象并赋值，避免硬编码对类型的依赖。
    /// </summary>
    public static class AssetGenerator
    {
        /// <summary>
        /// 生成配置。
        /// </summary>
        public class Config
        {
            /// <summary>资产输出目录（相对于 Assets/）</summary>
            public string OutputDir { get; set; } = "Data";

            /// <summary>项目 Assets 根目录</summary>
            public string AssetsRoot { get; set; } = "Assets";

            /// <summary>Row 类所在程序集限定名</summary>
            public string RowTypeAssemblyQualified { get; set; } = null;

            /// <summary>Table 类所在程序集限定名</summary>
            public string TableTypeAssemblyQualified { get; set; } = null;
        }

        /// <summary>
        /// 生成 ScriptableObject 资产并写入 .asset 文件。
        /// </summary>
        /// <param name="data">已解析的数据行</param>
        /// <param name="schema">表结构定义</param>
        /// <param name="config">生成配置</param>
        /// <param name="rowType">Row 类的 Type（如果已知）</param>
        /// <param name="tableType">Table 类的 Type（如果已知）</param>
        /// <returns>生成的 .asset 文件路径</returns>
        public static string Generate(
            List<Dictionary<string, object>> data,
            TableSchema schema,
            Config config,
            Type rowType = null,
            Type tableType = null)
        {
            // 1. 查找或创建 Row 类型
            if (rowType == null)
                rowType = FindType(CodeGenerator.GetRowClassName(
                    schema.TableName, new CodeGenerator.Config()));

            if (rowType == null)
            {
                Debug.LogWarning(
                    $"[ExcelToJSON] 找不到 Row 类型 '{CodeGenerator.GetRowClassName(schema.TableName, new CodeGenerator.Config())}'。" +
                    "请先生成 C# 代码再导出。");
                return null;
            }

            // 2. 创建 Row 对象列表
            var rows = new List<object>();
            foreach (var rowDict in data)
            {
                var rowObj = CreateRowObject(rowType, rowDict, schema);
                if (rowObj != null)
                    rows.Add(rowObj);
            }

            // 3. 查找或创建 Table 类型
            if (tableType == null)
                tableType = FindType(CodeGenerator.GetTableClassName(
                    schema.TableName, new CodeGenerator.Config()));

            if (tableType == null)
            {
                Debug.LogWarning(
                    $"[ExcelToJSON] 找不到 Table 类型 " +
                    $"'{CodeGenerator.GetTableClassName(schema.TableName, new CodeGenerator.Config())}'。");
                return null;
            }

            // 4. 创建或更新 Table 资产
            var outputDir = Path.Combine(config.AssetsRoot, config.OutputDir);
            Directory.CreateDirectory(outputDir);

            var assetPath = Path.Combine(outputDir, $"{schema.TableName}.asset");
            assetPath = assetPath.Replace("\\", "/");

            var tableObj = AssetDatabase.LoadAssetAtPath(assetPath, tableType);
            bool isNew = tableObj == null;

            if (isNew)
            {
                tableObj = ScriptableObject.CreateInstance(tableType);
            }

            // 5. 调用 SetRows 填充数据
            var setRowsMethod = tableType.GetMethod("SetRows",
                BindingFlags.Public | BindingFlags.Instance);
            if (setRowsMethod != null)
            {
                // 需要把 List<object> 转换为 List<TRow>
                var typedList = CreateTypedList(rowType, rows);
                setRowsMethod.Invoke(tableObj, new[] { typedList });
            }
            else
            {
                // 回退：直接设置 rows 字段
                var rowsField = tableType.GetField("rows",
                    BindingFlags.NonPublic | BindingFlags.Instance);
                if (rowsField != null)
                {
                    var typedList = CreateTypedList(rowType, rows);
                    rowsField.SetValue(tableObj, typedList);
                }
            }

            // 6. 保存资产
            if (isNew)
            {
                AssetDatabase.CreateAsset(tableObj, assetPath);
            }
            else
            {
                EditorUtility.SetDirty(tableObj);
            }

            Debug.Log($"[ExcelToJSON] {(isNew ? "创建" : "更新")}: {assetPath} ({rows.Count} 行)");

            return assetPath;
        }

        /// <summary>
        /// 通过反射创建一个 Row 对象并填充字段值。
        /// </summary>
        private static object CreateRowObject(
            Type rowType,
            Dictionary<string, object> values,
            TableSchema schema)
        {
            var obj = Activator.CreateInstance(rowType);

            foreach (var kv in values)
            {
                var fieldName = kv.Key;
                var fieldValue = kv.Value;

                var field = rowType.GetField(fieldName,
                    BindingFlags.Public | BindingFlags.Instance);
                if (field != null && fieldValue != null)
                {
                    try
                    {
                        // 类型适配：如果字段类型与值类型不完全匹配，尝试转换
                        var convertedValue = ConvertToFieldType(fieldValue, field.FieldType);
                        field.SetValue(obj, convertedValue);
                    }
                    catch (Exception ex)
                    {
                        Debug.LogWarning(
                            $"[ExcelToJSON] 设置字段 '{rowType.Name}.{fieldName}' 失败: {ex.Message}");
                    }
                }
            }

            return obj;
        }

        /// <summary>
        /// 将 object 值转换为目标字段类型。
        /// </summary>
        private static object ConvertToFieldType(object value, Type targetType)
        {
            if (value == null) return null;

            var valueType = value.GetType();
            if (targetType.IsAssignableFrom(valueType))
                return value;

            // int → float
            if (targetType == typeof(float) && value is int intVal)
                return (float)intVal;

            // float → int
            if (targetType == typeof(int) && value is float floatVal)
                return Mathf.RoundToInt(floatVal);

            // int/float → string
            if (targetType == typeof(string))
                return value.ToString();

            // string → bool
            if (targetType == typeof(bool) && value is string strBool)
            {
                return strBool.ToLower() == "true" || strBool == "1";
            }

            // 数组类型
            if (targetType.IsArray && value is Array srcArray)
            {
                var elemType = targetType.GetElementType();
                var destArray = Array.CreateInstance(elemType, srcArray.Length);
                for (int i = 0; i < srcArray.Length; i++)
                {
                    destArray.SetValue(ConvertToFieldType(srcArray.GetValue(i), elemType), i);
                }
                return destArray;
            }

            // Unity Vector2
            if (targetType == typeof(Vector2) && value is Vector2 v2)
                return v2;

            // Unity Vector3
            if (targetType == typeof(Vector3) && value is Vector3 v3)
                return v3;

            // Unity Color
            if (targetType == typeof(Color) && value is Color c)
                return c;

            // 尝试 Convert
            try
            {
                return Convert.ChangeType(value, targetType);
            }
            catch
            {
                return value;
            }
        }

        /// <summary>
        /// 创建泛型 List&lt;TRow&gt; 并用给定数据填充。
        /// </summary>
        private static object CreateTypedList(Type elementType, List<object> items)
        {
            var listType = typeof(List<>).MakeGenericType(elementType);
            var list = Activator.CreateInstance(listType);
            var addMethod = listType.GetMethod("Add");

            foreach (var item in items)
            {
                addMethod?.Invoke(list, new[] { item });
            }

            return list;
        }

        /// <summary>
        /// 在已加载的程序集中按类名查找类型。
        /// </summary>
        private static Type FindType(string className)
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                foreach (var type in asm.GetTypes())
                {
                    if (type.Name == className && !type.IsAbstract)
                        return type;
                }
            }
            return null;
        }
    }
}
