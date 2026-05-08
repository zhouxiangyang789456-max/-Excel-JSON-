using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Validator.Rules
{
    /// <summary>
    /// Applies validation rules defined in the #Rules Excel sheet.
    /// Supports: range, regex, multiple, not_empty, required, enum.
    /// Conditions are evaluated before the rule is applied.
    /// </summary>
    public class CustomDataRule : IDataRule
    {
        private readonly List<RuleConfig> configs;

        public CustomDataRule(List<RuleConfig> configs)
        {
            this.configs = configs ?? new List<RuleConfig>();
        }

        public List<ValidationError> Validate(
            List<List<string>> rows, TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            int dataStart = schema.DataStartRow - 1;

            foreach (var config in configs)
            {
                var field = FindField(schema, config.FieldName);
                if (field == null) continue;

                int condFieldIdx = -1;
                FieldDef condField = null;
                if (!string.IsNullOrEmpty(config.ConditionField))
                {
                    condField = FindField(schema, config.ConditionField);
                    if (condField != null)
                        condFieldIdx = condField.ColumnIndex;
                }

                for (int ri = dataStart; ri < rows.Count; ri++)
                {
                    var row = rows[ri];
                    if (ExcelReader.IsRowEmpty(row)) continue;

                    var rawValue = field.ColumnIndex < row.Count ? row[field.ColumnIndex] : "";

                    if (!ShouldApply(row, config, condField, condFieldIdx))
                        continue;

                    var error = ApplyRule(rawValue, field, config, fileName, sheetName, ri + 1);
                    if (error != null)
                        errors.Add(error);
                }
            }

            return errors;
        }

        private bool ShouldApply(List<string> row, RuleConfig config, FieldDef condField, int condFieldIdx)
        {
            if (string.IsNullOrEmpty(config.ConditionField))
                return true;

            var condValue = condFieldIdx >= 0 && condFieldIdx < row.Count
                ? row[condFieldIdx] : "";

            return config.ConditionOp switch
            {
                "="  => condValue == config.ConditionValue,
                "!=" => condValue != config.ConditionValue,
                ">=" => CompareNumeric(condValue, config.ConditionValue) >= 0,
                "<=" => CompareNumeric(condValue, config.ConditionValue) <= 0,
                ">"  => CompareNumeric(condValue, config.ConditionValue) > 0,
                "<"  => CompareNumeric(condValue, config.ConditionValue) < 0,
                _    => true,
            };
        }

        private int CompareNumeric(string a, string b)
        {
            if (float.TryParse(a, NumberStyles.Float, CultureInfo.InvariantCulture, out var va) &&
                float.TryParse(b, NumberStyles.Float, CultureInfo.InvariantCulture, out var vb))
            {
                return va.CompareTo(vb);
            }
            return string.Compare(a, b, System.StringComparison.Ordinal);
        }

        private ValidationError ApplyRule(
            string rawValue, FieldDef field, RuleConfig config,
            string fileName, string sheetName, int rowNum)
        {
            var msg = config.Message;

            switch (config.RuleName)
            {
                case "not_empty":
                case "required":
                    if (string.IsNullOrWhiteSpace(rawValue))
                    {
                        return MakeError(fileName, sheetName, rowNum, field.Name,
                            rawValue, config.RuleName,
                            msg ?? $"Field \"{field.Name}\" is required but is empty.");
                    }
                    break;

                case "range":
                    if (!string.IsNullOrWhiteSpace(rawValue) && config.Params != null)
                    {
                        var parts = config.Params.Split('~', '-', ',');
                        if (parts.Length >= 2 &&
                            float.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var min) &&
                            float.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var max))
                        {
                            if (float.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var val))
                            {
                                if (val < min || val > max)
                                {
                                    return MakeError(fileName, sheetName, rowNum, field.Name,
                                        rawValue, "RangeRule",
                                        msg ?? $"Value {val} is outside range [{min}, {max}].");
                                }
                            }
                        }
                    }
                    break;

                case "regex":
                    if (!string.IsNullOrWhiteSpace(rawValue) && !string.IsNullOrEmpty(config.Params))
                    {
                        try
                        {
                            if (!Regex.IsMatch(rawValue, config.Params))
                            {
                                return MakeError(fileName, sheetName, rowNum, field.Name,
                                    rawValue, "RegexRule",
                                    msg ?? $"Value does not match pattern \"{config.Params}\".");
                            }
                        }
                        catch (RegexParseException)
                        {
                            // Invalid regex — skip
                        }
                    }
                    break;

                case "multiple":
                    if (!string.IsNullOrWhiteSpace(rawValue) && !string.IsNullOrEmpty(config.Params))
                    {
                        if (int.TryParse(config.Params.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var factor) &&
                            factor != 0)
                        {
                            if (float.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var val))
                            {
                                if (val % factor != 0)
                                {
                                    return MakeError(fileName, sheetName, rowNum, field.Name,
                                        rawValue, "MultipleRule",
                                        msg ?? $"Value {val} is not a multiple of {factor}.");
                                }
                            }
                        }
                    }
                    break;

                case "enum":
                    if (!string.IsNullOrWhiteSpace(rawValue) && !string.IsNullOrEmpty(config.Params))
                    {
                        var allowed = new HashSet<string>();
                        foreach (var part in config.Params.Split('|', ','))
                        {
                            var trimmed = part.Trim();
                            if (!string.IsNullOrEmpty(trimmed))
                                allowed.Add(trimmed);
                        }
                        if (allowed.Count > 0 && !allowed.Contains(rawValue))
                        {
                            return MakeError(fileName, sheetName, rowNum, field.Name,
                                rawValue, "EnumRule",
                                msg ?? $"Value \"{rawValue}\" is not in allowed set: [{string.Join(", ", allowed)}].");
                        }
                    }
                    break;

                case "unique":
                    // Already handled by IdUniqueRule, but can be applied to other columns
                    // This is a simplified version — full uniqueness would need cross-row tracking
                    break;
            }

            return null;
        }

        private FieldDef FindField(TableSchema schema, string name)
        {
            name = (name ?? "").Trim().ToLower();
            foreach (var field in schema.Fields)
            {
                if (field.Name?.Trim().ToLower() == name || field.CsFieldName?.Trim().ToLower() == name)
                    return field;
            }
            return null;
        }

        private static ValidationError MakeError(
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
                Level = ErrorLevel.Error,
            };
        }

    }
}
