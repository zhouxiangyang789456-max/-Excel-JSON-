using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelToJsonPlugin.Editor.Core.Models;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// 将 Excel 单元格字符串值按照声明类型转换为 C# 对象。
    /// 用于 Editor 下创建 ScriptableObject 和 JSON 导出。
    /// </summary>
    public static class TypeMapper
    {
        /// <summary>
        /// 转换单个单元格的值。
        /// </summary>
        /// <param name="rawValue">Excel 原始字符串</param>
        /// <param name="field">字段定义（含类型信息）</param>
        /// <param name="errorMessage">输出错误信息，成功时为空</param>
        /// <returns>转换后的对象</returns>
        public static object ConvertValue(string rawValue, FieldDef field, out string errorMessage)
        {
            errorMessage = null;

            // 空值处理
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return GetDefaultValue(field);
            }

            var value = rawValue.Trim();

            try
            {
                return field.NormalizedType switch
                {
                    "int" => ParseInt(value),
                    "float" => ParseFloat(value),
                    "string" => value,
                    "bool" => ParseBool(value),
                    "int[]" => ParseIntArray(value),
                    "float[]" => ParseFloatArray(value),
                    "string[]" => ParseStringArray(value),
                    "Vector2" => ParseVector2(value),
                    "Vector3" => ParseVector3(value),
                    "Color" => ParseColor(value),
                    "json" => ParseJson(value),
                    "loc" => value,
                    _ => HandleCompositeOrUnknown(value, field),
                };
            }
            catch (FormatException ex)
            {
                errorMessage = ex.Message;
                return GetDefaultValue(field);
            }
        }

        /// <summary>
        /// 获取字段类型的默认值（空值回退）。
        /// </summary>
        public static object GetDefaultValue(FieldDef field)
        {
            return field.NormalizedType switch
            {
                "int" => 0,
                "float" => 0f,
                "string" => "",
                "bool" => false,
                "int[]" => null,
                "float[]" => null,
                "string[]" => null,
                "Vector2" => Vector2.zero,
                "Vector3" => Vector3.zero,
                "Color" => Color.white,
                "json" => "",
                "loc" => "",
                _ => field.IsCompositeType && field.NormalizedType?.StartsWith("ref:") == true ? 0 :
                     field.IsCompositeType && field.NormalizedType?.StartsWith("enum:") == true ? 0 :
                     field.IsCompositeType && field.NormalizedType?.StartsWith("res") == true ? "" : (object)"",
            };
        }

        // ============================================================
        // 基础类型
        // ============================================================

        private static int ParseInt(string value)
        {
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
                return result;

            // 尝试解析浮点再截断
            if (float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var f))
                return (int)f;

            throw new FormatException($"无法将 \"{value}\" 转换为 int");
        }

        private static float ParseFloat(string value)
        {
            if (float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
                return result;

            // 兼容中文格式
            if (float.TryParse(value, NumberStyles.Float, CultureInfo.GetCultureInfo("zh-CN"), out result))
                return result;

            throw new FormatException($"无法将 \"{value}\" 转换为 float");
        }

        private static bool ParseBool(string value)
        {
            value = value.ToLower().Trim();

            if (value == "true" || value == "1" || value == "是" || value == "yes" || value == "y")
                return true;
            if (value == "false" || value == "0" || value == "否" || value == "no" || value == "n")
                return false;

            throw new FormatException($"无法将 \"{value}\" 转换为 bool（合法值: true/false/0/1/是/否）");
        }

        // ============================================================
        // 数组类型
        // ============================================================

        private static int[] ParseIntArray(string value)
        {
            var elements = ParseArrayElements(value);
            var result = new int[elements.Count];
            for (int i = 0; i < elements.Count; i++)
            {
                if (!int.TryParse(elements[i].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out result[i]))
                    throw new FormatException($"int[] 中包含无效元素: \"{elements[i]}\"");
            }
            return result;
        }

        private static float[] ParseFloatArray(string value)
        {
            var elements = ParseArrayElements(value);
            var result = new float[elements.Count];
            for (int i = 0; i < elements.Count; i++)
            {
                if (!float.TryParse(elements[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out result[i]))
                    throw new FormatException($"float[] 中包含无效元素: \"{elements[i]}\"");
            }
            return result;
        }

        private static string[] ParseStringArray(string value)
        {
            return ParseArrayElements(value).ToArray();
        }

        /// <summary>
        /// 解析数组元素，支持两种格式：
        /// 1. JSON 风格：[1, 2, 3] 或 ["A", "B"]
        /// 2. 管道风格（策划友好）：1|2|3
        /// </summary>
        private static List<string> ParseArrayElements(string value)
        {
            value = value.Trim();
            var elements = new List<string>();

            // JSON 数组格式 [ ... ]
            if (value.StartsWith("[") && value.EndsWith("]"))
            {
                value = value.Substring(1, value.Length - 2).Trim();
                if (string.IsNullOrEmpty(value))
                    return elements;

                // 简易 JSON 解析（不引入完整 JSON 库的情况下）
                // 支持 [1,2,3] 和 ["a","b"] 和 [1, 2, 3]
                bool inString = false;
                var current = new System.Text.StringBuilder();
                for (int i = 0; i < value.Length; i++)
                {
                    char c = value[i];
                    if (c == '"')
                    {
                        inString = !inString;
                        current.Append(c);
                    }
                    else if (c == ',' && !inString)
                    {
                        var elem = current.ToString().Trim().Trim('"');
                        if (!string.IsNullOrEmpty(elem) || current.Length > 0)
                            elements.Add(elem);
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                var last = current.ToString().Trim().Trim('"');
                if (!string.IsNullOrEmpty(last) || current.Length > 0)
                    elements.Add(last);
            }
            // 管道分隔格式 1|2|3
            else if (value.Contains("|"))
            {
                foreach (var part in value.Split('|'))
                {
                    var trimmed = part.Trim().Trim('"');
                    if (!string.IsNullOrEmpty(trimmed))
                        elements.Add(trimmed);
                }
            }
            // 单元素
            else
            {
                elements.Add(value);
            }

            return elements;
        }

        // ============================================================
        // Unity 类型
        // ============================================================

        private static Vector2 ParseVector2(string value)
        {
            var nums = ParseNumbers(value, 2);
            return new Vector2(nums[0], nums[1]);
        }

        private static Vector3 ParseVector3(string value)
        {
            var nums = ParseNumbers(value, 3);
            return new Vector3(nums[0], nums[1], nums[2]);
        }

        private static Color ParseColor(string value)
        {
            value = value.Trim();

            // Hex 格式: #FF0000 或 #FF0000FF
            if (value.StartsWith("#"))
            {
                value = value.TrimStart('#');
                if (value.Length == 6)
                    value += "FF"; // 默认不透明
                if (value.Length == 8)
                {
                    if (uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var hex))
                    {
                        return new Color(
                            ((hex >> 24) & 0xFF) / 255f,
                            ((hex >> 16) & 0xFF) / 255f,
                            ((hex >> 8) & 0xFF) / 255f,
                            (hex & 0xFF) / 255f
                        );
                    }
                }
            }

            // RGBA 数组格式: [1,0,0,1]
            var nums = ParseNumbers(value, 3, 4);
            return nums.Length == 4
                ? new Color(nums[0], nums[1], nums[2], nums[3])
                : new Color(nums[0], nums[1], nums[2], 1f);
        }

        /// <summary>
        /// 从 "[a,b,c]" 格式中解析数字列表。
        /// </summary>
        private static float[] ParseNumbers(string value, int minCount, int maxCount = -1)
        {
            if (maxCount < 0) maxCount = minCount;

            value = value.Trim();
            if (value.StartsWith("[") && value.EndsWith("]"))
                value = value.Substring(1, value.Length - 2);

            var parts = value.Split(',', '|');
            var nums = new List<float>();
            foreach (var p in parts)
            {
                var trimmed = p.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                if (float.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
                    nums.Add(n);
                else
                    throw new FormatException($"无法解析数字: \"{trimmed}\"");
            }

            if (nums.Count < minCount || nums.Count > maxCount)
                throw new FormatException(
                    $"需要 {minCount}" + (maxCount != minCount ? $"~{maxCount}" : "") + $" 个数字，实际 {nums.Count} 个");

            return nums.ToArray();
        }

        // ============================================================
        // 复合类型 & 特殊
        // ============================================================

        private static string ParseJson(string value)
        {
            // 简单的 JSON 格式校验
            value = value.Trim();
            if ((value.StartsWith("{") && value.EndsWith("}"))
                || (value.StartsWith("[") && value.EndsWith("]")))
            {
                // 基本格式正确，不强校验完整 JSON
                return value;
            }
            throw new FormatException($"json 类型必须为 JSON 对象或数组: \"{value}\"");
        }

        private static object HandleCompositeOrUnknown(string value, FieldDef field)
        {
            // ref:TableName → 存 int
            if (field.NormalizedType?.StartsWith("ref:") == true)
                return ParseInt(value);

            // enum:Type → 存 int
            if (field.NormalizedType?.StartsWith("enum:") == true)
                return ParseInt(value);

            // res / res:Sprite → 存 string
            if (field.NormalizedType?.StartsWith("res") == true)
                return value;

            // 兜底：返回原始字符串
            return value;
        }
    }
}
