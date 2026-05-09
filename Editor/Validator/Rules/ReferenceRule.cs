using System.Collections.Generic;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Validator.Rules
{
    /// <summary>
    /// Stage 3 — 跨表引用完整性校验。
    /// 检查 ref:TargetTable 字段的值是否存在于目标表的 ID 列中。
    /// </summary>
    public class RefIntegrityRule : IStage3Rule
    {
        public List<ValidationError> Validate(
            List<List<string>> rows,
            TableSchema schema,
            string fileName,
            Dictionary<string, HashSet<int>> referencedIds)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            foreach (var field in schema.Fields)
            {
                if (!field.IsCompositeType || !field.NormalizedType?.StartsWith("ref:") == true)
                    continue;

                var targetTable = field.CompositeParam;
                if (string.IsNullOrEmpty(targetTable))
                    continue;

                // Check if target table data is available
                if (!referencedIds.TryGetValue(targetTable, out var targetIds)
                    || targetIds == null || targetIds.Count == 0)
                {
                    errors.Add(MakeWarning(fileName, schema.TableName, 0, field.Name, "",
                        "RefIntegrity",
                        $"引用目标表 \"{targetTable}\" 未找到或无数据，跳过引用完整性检查"));
                    continue;
                }

                for (int ri = dataStart; ri < rows.Count; ri++)
                {
                    var row = rows[ri];
                    var rawValue = field.ColumnIndex < row.Count ? row[field.ColumnIndex] : "";

                    if (string.IsNullOrWhiteSpace(rawValue))
                        continue;

                    // ref values are parsed as int
                    if (int.TryParse(rawValue.Trim(), out var refId))
                    {
                        if (!targetIds.Contains(refId))
                        {
                            errors.Add(new ValidationError
                            {
                                FileName = fileName,
                                SheetName = schema.TableName,
                                Row = ri + 1,
                                ColumnName = field.Name,
                                RawValue = rawValue,
                                RuleName = "RefIntegrity",
                                Message = $"引用 ID {refId} 在表 \"{targetTable}\" 中不存在",
                                Level = ErrorLevel.Error,
                            });
                        }
                    }
                }
            }

            return errors;
        }

        /// <summary>
        /// Extract all ID values from a data table (first column = id).
        /// </summary>
        public static HashSet<int> ExtractIds(List<List<string>> rows,
            TableSchema tableSchema)
        {
            var ids = new HashSet<int>();
            if (rows == null || tableSchema == null) return ids;

            // Find the ID field (first field is typically the id)
            int idCol = -1;
            foreach (var field in tableSchema.Fields)
            {
                var name = (field.Name ?? "").Trim().ToLower();
                if (name == "id")
                {
                    idCol = field.ColumnIndex;
                    break;
                }
            }
            if (idCol < 0 && tableSchema.Fields.Count > 0)
                idCol = tableSchema.Fields[0].ColumnIndex;

            int dataStart = tableSchema.DataStartRow - 1;
            for (int ri = dataStart; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                var rawValue = idCol < row.Count ? row[idCol] : "";
                if (!string.IsNullOrWhiteSpace(rawValue)
                    && int.TryParse(rawValue.Trim(), out var id))
                {
                    ids.Add(id);
                }
            }

            return ids;
        }

        private static ValidationError MakeWarning(
            string file, string sheet, int row, string col,
            string value, string rule, string msg)
        {
            return new ValidationError
            {
                FileName = file,
                SheetName = sheet,
                Row = row,
                ColumnName = col,
                RawValue = value,
                RuleName = rule,
                Message = msg,
                Level = ErrorLevel.Warning,
            };
        }
    }

    /// <summary>
    /// Stage 3 — 枚举值存在性校验。
    /// 检查 enum:TargetTable 字段的值是否在目标枚举表的 ID 列中存在。
    /// </summary>
    public class EnumExistenceRule : IStage3Rule
    {
        public List<ValidationError> Validate(
            List<List<string>> rows,
            TableSchema schema,
            string fileName,
            Dictionary<string, HashSet<int>> referencedIds)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            foreach (var field in schema.Fields)
            {
                if (!field.IsCompositeType || !field.NormalizedType?.StartsWith("enum:") == true)
                    continue;

                var targetTable = field.CompositeParam;
                if (string.IsNullOrEmpty(targetTable))
                    continue;

                if (!referencedIds.TryGetValue(targetTable, out var targetIds)
                    || targetIds == null || targetIds.Count == 0)
                {
                    errors.Add(new ValidationError
                    {
                        FileName = fileName,
                        SheetName = schema.TableName,
                        Row = 0,
                        ColumnName = field.Name,
                        RuleName = "EnumExistence",
                        Message = $"枚举目标表 \"{targetTable}\" 未找到或无数据，跳过枚举值校验",
                        Level = ErrorLevel.Warning,
                    });
                    continue;
                }

                for (int ri = dataStart; ri < rows.Count; ri++)
                {
                    var row = rows[ri];
                    var rawValue = field.ColumnIndex < row.Count ? row[field.ColumnIndex] : "";

                    if (string.IsNullOrWhiteSpace(rawValue))
                        continue;

                    if (int.TryParse(rawValue.Trim(), out var enumVal))
                    {
                        if (!targetIds.Contains(enumVal))
                        {
                            errors.Add(new ValidationError
                            {
                                FileName = fileName,
                                SheetName = schema.TableName,
                                Row = ri + 1,
                                ColumnName = field.Name,
                                RawValue = rawValue,
                                RuleName = "EnumExistence",
                                Message = $"枚举值 {enumVal} 在表 \"{targetTable}\" 中不存在",
                                Level = ErrorLevel.Error,
                            });
                        }
                    }
                }
            }

            return errors;
        }
    }
}
