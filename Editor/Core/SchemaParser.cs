using System;
using System.Collections.Generic;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// 从 Excel 原始行列数据中解析表结构（字段名、类型、注释）。
    /// 支持配置表头行号和跳过规则。
    /// </summary>
    public static class SchemaParser
    {
        /// <summary>
        /// 解析单个 Sheet 的结构定义。
        /// </summary>
        /// <param name="rows">Sheet 的原始行数据</param>
        /// <param name="sheetName">Sheet 名称</param>
        /// <param name="fileName">所属 Excel 文件名</param>
        /// <param name="headerRow">字段名所在行号（1-based）</param>
        /// <param name="typeRow">类型声明所在行号（1-based）</param>
        /// <param name="commentRow">注释行号（1-based，0 表示无注释行）</param>
        /// <param name="dataStartRow">数据起始行号（1-based）</param>
        /// <param name="skipSheetPrefixes">以这些字符串开头的 Sheet 会被跳过</param>
        /// <returns>解析后的 TableSchema，如果应跳过则返回 null</returns>
        public static TableSchema Parse(
            List<List<string>> rows,
            string sheetName,
            string fileName,
            int headerRow = 1,
            int typeRow = 2,
            int commentRow = 3,
            int dataStartRow = 4,
            string[] skipSheetPrefixes = null)
        {
            // 跳过特定前缀的 Sheet
            if (skipSheetPrefixes != null)
            {
                foreach (var prefix in skipSheetPrefixes)
                {
                    if (!string.IsNullOrEmpty(prefix) && sheetName.StartsWith(prefix))
                        return null;
                }
            }

            // 转换为 0-based
            int headerIdx = headerRow - 1;
            int typeIdx = typeRow - 1;
            int commentIdx = commentRow - 1;

            if (headerIdx >= rows.Count)
                throw new FormatException(
                    $"Sheet '{sheetName}' 行数不足：需要第 {headerRow} 行作为字段名，但只有 {rows.Count} 行");

            if (typeIdx >= rows.Count)
                throw new FormatException(
                    $"Sheet '{sheetName}' 行数不足：需要第 {typeRow} 行作为类型声明，但只有 {rows.Count} 行");

            var headerNames = rows[headerIdx];
            var typeNames = rows[typeIdx];

            var schema = new TableSchema
            {
                TableName = sheetName,
                FileName = fileName,
                DataStartRow = dataStartRow,
            };

            int maxCols = Math.Max(headerNames.Count, typeNames.Count);

            for (int ci = 0; ci < maxCols; ci++)
            {
                var rawName = ci < headerNames.Count ? headerNames[ci] : string.Empty;
                var rawType = ci < typeNames.Count ? typeNames[ci] : string.Empty;

                // 跳过字段名和类型都为空或仅含注释的列
                if (string.IsNullOrWhiteSpace(rawName) && string.IsNullOrWhiteSpace(rawType))
                    continue;

                if (string.IsNullOrWhiteSpace(rawName))
                {
                    // 有类型但没有字段名 → 用自动生成名
                    rawName = $"Column_{ci}";
                }

                var field = new FieldDef
                {
                    Name = rawName.Trim(),
                    RawType = rawType?.Trim() ?? "string",
                    ColumnIndex = ci,
                };

                // 生成 C# 合法字段名
                field.CsFieldName = ToCSharpFieldName(field.Name);

                // 标准化类型并映射到 C# 类型
                ParseType(field);

                // 注释（如果有）
                if (commentRow > 0 && commentIdx < rows.Count && ci < rows[commentIdx].Count)
                {
                    field.Comment = rows[commentIdx][ci];
                }

                schema.Fields.Add(field);
            }

            // 收集注释行原文
            if (commentRow > 0 && commentIdx < rows.Count)
            {
                schema.Comments = new List<string>(rows[commentIdx]);
            }

            return schema;
        }

        /// <summary>
        /// 解析类型声明，填充 NormalizedType / CsTypeName / IsCompositeType 等。
        /// </summary>
        private static void ParseType(FieldDef field)
        {
            var raw = field.RawType?.Trim().ToLower() ?? "string";

            // 复合类型检测
            if (raw.StartsWith("ref:"))
            {
                field.NormalizedType = raw;
                field.CsTypeName = "int";
                field.IsCompositeType = true;
                field.CompositeParam = raw.Substring(4).Trim();
                return;
            }

            if (raw.StartsWith("enum:"))
            {
                field.NormalizedType = raw;
                field.CsTypeName = "int";
                field.IsCompositeType = true;
                field.CompositeParam = raw.Substring(5).Trim();
                return;
            }

            if (raw.StartsWith("res"))
            {
                field.NormalizedType = raw;
                field.CsTypeName = "string";
                field.IsCompositeType = true;
                // res:Sprite → param = "Sprite"
                var colonIdx = raw.IndexOf(':');
                field.CompositeParam = colonIdx > 0 ? raw.Substring(colonIdx + 1).Trim() : "Object";
                return;
            }

            // 基础类型映射
            switch (raw)
            {
                case "int":
                    field.NormalizedType = "int";
                    field.CsTypeName = "int";
                    break;
                case "float":
                    field.NormalizedType = "float";
                    field.CsTypeName = "float";
                    break;
                case "string":
                    field.NormalizedType = "string";
                    field.CsTypeName = "string";
                    break;
                case "bool":
                case "boolean":
                    field.NormalizedType = "bool";
                    field.CsTypeName = "bool";
                    break;
                case "int[]":
                    field.NormalizedType = "int[]";
                    field.CsTypeName = "int[]";
                    break;
                case "float[]":
                    field.NormalizedType = "float[]";
                    field.CsTypeName = "float[]";
                    break;
                case "string[]":
                    field.NormalizedType = "string[]";
                    field.CsTypeName = "string[]";
                    break;
                case "vec2":
                case "vector2":
                    field.NormalizedType = "Vector2";
                    field.CsTypeName = "Vector2";
                    break;
                case "vec3":
                case "vector3":
                    field.NormalizedType = "Vector3";
                    field.CsTypeName = "Vector3";
                    break;
                case "color":
                    field.NormalizedType = "Color";
                    field.CsTypeName = "Color";
                    break;
                case "json":
                case "object":
                    field.NormalizedType = "json";
                    field.CsTypeName = "string";
                    field.IsCompositeType = true;
                    break;
                case "loc":
                    field.NormalizedType = "loc";
                    field.CsTypeName = "string";
                    field.IsCompositeType = true;
                    break;
                default:
                    field.NormalizedType = raw;
                    field.CsTypeName = "string"; // 兜底
                    break;
            }

            if (string.IsNullOrEmpty(field.CsTypeName))
                field.CsTypeName = "string";
        }

        private static readonly HashSet<string> CSharpKeywords = new HashSet<string>
        {
            "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char",
            "checked", "class", "const", "continue", "decimal", "default", "delegate", "do",
            "double", "else", "enum", "event", "explicit", "extern", "false", "finally",
            "fixed", "float", "for", "foreach", "goto", "if", "implicit", "in", "int",
            "interface", "internal", "is", "lock", "long", "namespace", "new", "null",
            "object", "operator", "out", "override", "params", "private", "protected",
            "public", "readonly", "ref", "return", "sbyte", "sealed", "short", "sizeof",
            "stackalloc", "static", "string", "struct", "switch", "this", "throw", "true",
            "try", "typeof", "uint", "ulong", "unchecked", "unsafe", "ushort", "using",
            "virtual", "void", "volatile", "while",
        };

        /// <summary>
        /// 将 Excel 字段名转换为合法的 C# 字段名。
        /// </summary>
        private static string ToCSharpFieldName(string excelName)
        {
            if (string.IsNullOrEmpty(excelName))
                return "Unknown";

            // 已经是合法 C# 标识符的直接返回（但检查是否为关键字）
            if (System.Text.RegularExpressions.Regex.IsMatch(excelName, @"^[a-zA-Z_][a-zA-Z0-9_]*$"))
            {
                if (CSharpKeywords.Contains(excelName))
                    return "_" + excelName;
                return excelName;
            }

            // 简单策略：移除非法字符，首字母保持小写
            var sb = new System.Text.StringBuilder();
            foreach (char c in excelName)
            {
                if (char.IsLetterOrDigit(c) || c == '_')
                    sb.Append(c);
                else if (c == ' ' || c == '-' || c == '（' || c == '）' || c == '(' || c == ')')
                    sb.Append('_');
                // 中文字符暂时保留（C# 支持 Unicode 标识符，但不推荐）
                // 这里简单跳过非 ASCII 或保留 Unicode
            }

            var result = sb.ToString().Trim('_');
            if (string.IsNullOrEmpty(result))
                return "Unknown";

            // 确保不以数字开头
            if (char.IsDigit(result[0]))
                result = "_" + result;

            return result;
        }
    }
}
