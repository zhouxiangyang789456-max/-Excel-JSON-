using System.Collections.Generic;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Validator
{
    /// <summary>
    /// 校验引擎调度器。Sprint 2 实现完整功能。
    /// </summary>
    public static class ValidationEngine
    {
        /// <summary>
        /// 执行三阶段校验。
        /// </summary>
        public static ValidationReport Validate(
            List<List<string>> rows,
            TableSchema schema,
            string excelFileName)
        {
            var report = new ValidationReport();

            // Stage 1: 结构校验
            ValidateStructure(schema, excelFileName, report);

            // Stage 2: 数据校验（Sprint 2 实现）
            // ValidateData(rows, schema, excelFileName, report);

            // Stage 3: 引用校验（Sprint 3 实现）
            // ValidateReferences(rows, schema, excelFileName, report);

            return report;
        }

        private static void ValidateStructure(
            TableSchema schema,
            string fileName,
            ValidationReport report)
        {
            // 字段名唯一性
            var seen = new HashSet<string>();
            foreach (var field in schema.Fields)
            {
                if (!string.IsNullOrEmpty(field.Name) && !seen.Add(field.Name))
                {
                    report.Add(fileName, schema.TableName, 1, field.Name,
                        field.Name, "FieldNameUnique",
                        $"字段名重复: \"{field.Name}\"", ErrorLevel.Error);
                }
            }

            // 不支持的类开警告
            var supportedTypes = new HashSet<string>
            {
                "int", "float", "string", "bool", "int[]", "float[]",
                "string[]", "Vector2", "Vector3", "Color", "json", "loc"
            };

            foreach (var field in schema.Fields)
            {
                var normalized = field.NormalizedType;
                if (normalized == null) continue;

                bool isComposite = normalized.StartsWith("ref:")
                    || normalized.StartsWith("enum:")
                    || normalized.StartsWith("res");

                if (!isComposite && !supportedTypes.Contains(normalized))
                {
                    report.Add(fileName, schema.TableName, 2, field.Name,
                        field.RawType, "TypeRecognized",
                        $"类型 \"{field.RawType}\" 不在支持列表中，将按 string 处理",
                        ErrorLevel.Warning);
                }
            }
        }
    }
}
