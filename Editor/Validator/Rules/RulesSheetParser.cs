using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToJsonPlugin.Editor.Validator.Rules
{
    /// <summary>
    /// Parses the #Rules sheet from an Excel file.
    /// Each row defines a validation rule for a specific column.
    ///
    /// Format (simple):
    ///   field | rule     | params
    ///   attack| range    | 0~9999
    ///   price | multiple | 10
    ///
    /// Format (extended):
    ///   field | rule    | condition | params  | message
    ///   heal  |required | type=2    |         | 治疗类物品必须填治疗量
    /// </summary>
    public static class RulesSheetParser
    {
        public static List<RuleConfig> Parse(List<List<string>> rows, string sheetName = "#Rules")
        {
            var rules = new List<RuleConfig>();
            if (rows == null || rows.Count < 2) return rules;

            // Find the header row (row 0 = field names)
            var header = rows[0];

            // Map column names to indices
            int fieldIdx = FindColumn(header, "field", "column", "target_field", "列名", "字段");
            int ruleIdx = FindColumn(header, "rule", "type", "rule_name", "规则", "校验");
            int paramsIdx = FindColumn(header, "params", "parameter", "config", "参数", "配置");
            int condIdx = FindColumn(header, "condition", "cond", "条件");
            int msgIdx = FindColumn(header, "message", "msg", "error_message", "说明", "提示");

            if (fieldIdx < 0 || ruleIdx < 0)
            {
                // Try alternative: first row might be data directly, columns by position
                if (header.Count >= 2)
                {
                    fieldIdx = 0;
                    ruleIdx = 1;
                    paramsIdx = header.Count > 2 ? 2 : -1;
                }
                else
                {
                    return rules;
                }
            }

            // Parse data rows (skip header)
            for (int ri = 1; ri < rows.Count; ri++)
            {
                var row = rows[ri];
                if (IsRowEmpty(row)) continue;

                var config = new RuleConfig();

                config.FieldName = GetCell(row, fieldIdx);
                config.RuleName = GetCell(row, ruleIdx)?.Trim().ToLower();
                config.Params = GetCell(row, paramsIdx);

                if (condIdx >= 0)
                {
                    var condRaw = GetCell(row, condIdx);
                    if (!string.IsNullOrEmpty(condRaw))
                        ParseCondition(condRaw, config);
                }

                if (msgIdx >= 0)
                    config.Message = GetCell(row, msgIdx);

                // Skip rows with empty field or rule
                if (string.IsNullOrWhiteSpace(config.FieldName) || string.IsNullOrWhiteSpace(config.RuleName))
                    continue;

                rules.Add(config);
            }

            return rules;
        }

        /// <summary>
        /// Parse a condition string like "type=2" or "rarity>=3" into components.
        /// </summary>
        private static void ParseCondition(string raw, RuleConfig config)
        {
            raw = raw.Trim();

            // Try each known operator
            foreach (var op in new[] { "!=", ">=", "<=", "=", ">", "<" })
            {
                var idx = raw.IndexOf(op, StringComparison.Ordinal);
                if (idx >= 0)
                {
                    config.ConditionField = raw.Substring(0, idx).Trim();
                    config.ConditionOp = op;
                    config.ConditionValue = raw.Substring(idx + op.Length).Trim();
                    return;
                }
            }

            // No operator found, treat whole string as condition field name (truthy check)
            config.ConditionField = raw;
            config.ConditionOp = "!=";
            config.ConditionValue = "";
        }

        /// <summary>
        /// Find a column index by its header name. Supports multiple aliases.
        /// </summary>
        private static int FindColumn(List<string> header, params string[] names)
        {
            for (int ci = 0; ci < header.Count; ci++)
            {
                var cell = header[ci]?.Trim().ToLower() ?? "";
                foreach (var name in names)
                {
                    if (cell == name.ToLower())
                        return ci;
                }
            }
            return -1;
        }

        private static string GetCell(List<string> row, int index)
        {
            if (index < 0 || index >= row.Count) return null;
            return row[index]?.Trim();
        }

        private static bool IsRowEmpty(List<string> row)
        {
            return row == null || row.All(string.IsNullOrEmpty);
        }
    }
}
