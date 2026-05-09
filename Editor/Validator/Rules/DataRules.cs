using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Validator.Rules
{
    /// <summary>ID column must not be empty.</summary>
    public class IdRequiredRule : IDataRule
    {
        public List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            if (schema.Fields.Count == 0) return errors;

            var idField = schema.Fields[0];
            int dataStart = schema.DataStartRow - 1;

            for (int ri = dataStart; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                if (ExcelReader.IsRowEmpty(row)) continue;

                var idValue = row.Count > idField.ColumnIndex ? row[idField.ColumnIndex] : "";
                if (string.IsNullOrWhiteSpace(idValue))
                {
                    errors.Add(new ValidationError
                    {
                        FileName = fileName,
                        SheetName = schema.TableName,
                        Row = ri + 1,
                        ColumnName = idField.Name,
                        RawValue = "",
                        RuleName = "IdRequired",
                        Message = "ID is required and cannot be empty.",
                        Level = ErrorLevel.Error,
                    });
                }
            }

            return errors;
        }
    }

    /// <summary>ID values must be unique across the table.</summary>
    public class IdUniqueRule : IDataRule
    {
        public List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            if (schema.Fields.Count == 0) return errors;

            var idField = schema.Fields[0];
            int dataStart = schema.DataStartRow - 1;
            var seenIds = new HashSet<string>();

            for (int ri = dataStart; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                if (ExcelReader.IsRowEmpty(row)) continue;

                var idValue = row.Count > idField.ColumnIndex ? row[idField.ColumnIndex] : "";
                if (string.IsNullOrWhiteSpace(idValue)) continue; // handled by IdRequiredRule

                if (!seenIds.Add(idValue))
                {
                    errors.Add(new ValidationError
                    {
                        FileName = fileName,
                        SheetName = schema.TableName,
                        Row = ri + 1,
                        ColumnName = idField.Name,
                        RawValue = idValue,
                        RuleName = "IdUnique",
                        Message = $"Duplicate ID: {idValue}",
                        Level = ErrorLevel.Error,
                    });
                }
            }

            return errors;
        }
    }

    /// <summary>Each cell's value must match its declared type.</summary>
    public class TypeMatchRule : IDataRule
    {
        public List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            for (int ri = dataStart; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                if (ExcelReader.IsRowEmpty(row)) continue;

                for (int fi = 0; fi < schema.Fields.Count; fi++)
                {
                    var field = schema.Fields[fi];
                    var rawValue = field.ColumnIndex < row.Count ? row[field.ColumnIndex] : "";

                    if (string.IsNullOrWhiteSpace(rawValue)) continue;

                    TypeMapper.ConvertValue(rawValue, field, out var errorMsg);
                    if (errorMsg != null)
                    {
                        errors.Add(new ValidationError
                        {
                            FileName = fileName,
                            SheetName = schema.TableName,
                            Row = ri + 1,
                            ColumnName = field.Name,
                            RawValue = rawValue,
                            RuleName = "TypeMatch",
                            Message = errorMsg,
                            Level = ErrorLevel.Error,
                        });
                    }
                }
            }

            return errors;
        }
    }

    /// <summary>Detect formula cells in data rows (NPOI returns cached values).</summary>
    public class FormulaDetectionRule : IDataRule
    {
        public List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            for (int ri = dataStart; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                if (ExcelReader.IsRowEmpty(row)) continue;

                for (int fi = 0; fi < schema.Fields.Count; fi++)
                {
                    var field = schema.Fields[fi];
                    if (field.ColumnIndex >= row.Count) continue;

                    var raw = row[field.ColumnIndex];

                    // NPOI returns formula cells as their cached value prefixed or
                    // as-is, but if the raw value starts with '=' it's a formula string.
                    // In our pipeline, ExcelReader.GetCellValue resolves formulas to
                    // cached values. This is a best-effort detection.
                    if (!string.IsNullOrEmpty(raw) && raw.TrimStart().StartsWith("="))
                    {
                        errors.Add(new ValidationError
                        {
                            FileName = fileName,
                            SheetName = schema.TableName,
                            Row = ri + 1,
                            ColumnName = field.Name,
                            RawValue = raw,
                            RuleName = "FormulaDetected",
                            Message = $"Cell contains a formula \"{raw.Trim()}\". Exported value is the cached result — press Ctrl+S in Excel to refresh.",
                            Level = ErrorLevel.Warning,
                        });
                    }
                }
            }

            return errors;
        }
    }

    /// <summary>Validate that enum:Type values are within the declared enum sheet's ID range.
    /// Full cross-sheet validation happens in Sprint 3.</summary>
    public class EnumSanityRule : IDataRule
    {
        public List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            foreach (var field in schema.Fields)
            {
                if (!field.IsCompositeType || !(field.NormalizedType?.StartsWith("enum:") == true))
                    continue;

                for (int ri = dataStart; ri < rows.Count; ri++)
                {
                    var row = rows[ri];
                    if (ExcelReader.IsRowEmpty(row)) continue;

                    var rawValue = field.ColumnIndex < row.Count ? row[field.ColumnIndex] : "";
                    if (string.IsNullOrWhiteSpace(rawValue)) continue;

                    if (!int.TryParse(rawValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                    {
                        errors.Add(new ValidationError
                        {
                            FileName = fileName,
                            SheetName = schema.TableName,
                            Row = ri + 1,
                            ColumnName = field.Name,
                            RawValue = rawValue,
                            RuleName = "EnumSanity",
                            Message = $"enum column \"{field.Name}\" expects an integer value, got \"{rawValue}\"",
                            Level = ErrorLevel.Error,
                        });
                    }
                }
            }

            return errors;
        }
    }

    /// <summary>Validate res:Type columns have valid path format.</summary>
    public class ResPathRule : IDataRule
    {
        public List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            foreach (var field in schema.Fields)
            {
                if (!field.IsCompositeType || !(field.NormalizedType?.StartsWith("res") == true))
                    continue;

                for (int ri = dataStart; ri < rows.Count; ri++)
                {
                    var row = rows[ri];
                    if (ExcelReader.IsRowEmpty(row)) continue;

                    var rawValue = field.ColumnIndex < row.Count ? row[field.ColumnIndex] : "";
                    if (string.IsNullOrWhiteSpace(rawValue)) continue;

                    var trimmed = rawValue.Trim();

                    if (trimmed.Contains("\\"))
                    {
                        errors.Add(new ValidationError
                        {
                            FileName = fileName,
                            SheetName = schema.TableName,
                            Row = ri + 1,
                            ColumnName = field.Name,
                            RawValue = rawValue,
                            RuleName = "ResPathFormat",
                            Message = $"Resource path uses backslashes. Use forward slashes: \"{trimmed.Replace("\\", "/")}\"",
                            Level = ErrorLevel.Warning,
                        });
                    }

                    if (trimmed.StartsWith("/") || trimmed.EndsWith("/"))
                    {
                        errors.Add(new ValidationError
                        {
                            FileName = fileName,
                            SheetName = schema.TableName,
                            Row = ri + 1,
                            ColumnName = field.Name,
                            RawValue = rawValue,
                            RuleName = "ResPathFormat",
                            Message = "Resource path should not start or end with '/'",
                            Level = ErrorLevel.Warning,
                        });
                    }

                    if (trimmed.Contains(".."))
                    {
                        errors.Add(new ValidationError
                        {
                            FileName = fileName,
                            SheetName = schema.TableName,
                            Row = ri + 1,
                            ColumnName = field.Name,
                            RawValue = rawValue,
                            RuleName = "ResPathTraversal",
                            Message = $"Resource path contains '..' (path traversal): \"{trimmed}\"",
                            Level = ErrorLevel.Error,
                        });
                    }
                }
            }

            return errors;
        }
    }

    /// <summary>Check that required fields (ref, enum composite) are not empty when expected.</summary>
    public class RequiredFieldRule : IDataRule
    {
        public List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            for (int ri = dataStart; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                if (ExcelReader.IsRowEmpty(row)) continue;

                for (int fi = 0; fi < schema.Fields.Count; fi++)
                {
                    var field = schema.Fields[fi];
                    var rawValue = field.ColumnIndex < row.Count ? row[field.ColumnIndex] : "";

                    // ref:X fields should have a value (foreign key)
                    if (field.IsCompositeType && field.NormalizedType?.StartsWith("ref:") == true)
                    {
                        if (string.IsNullOrWhiteSpace(rawValue))
                        {
                            errors.Add(new ValidationError
                            {
                                FileName = fileName,
                                SheetName = schema.TableName,
                                Row = ri + 1,
                                ColumnName = field.Name,
                                RawValue = rawValue,
                                RuleName = "RefRequired",
                                Message = $"Foreign key reference \"{field.Name}\" is empty — this row won't link to {field.CompositeParam}",
                                Level = ErrorLevel.Warning,
                            });
                        }
                    }
                }
            }

            return errors;
        }
    }
}
