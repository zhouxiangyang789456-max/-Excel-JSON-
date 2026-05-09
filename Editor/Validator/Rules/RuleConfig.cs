namespace ExcelToJsonPlugin.Editor.Validator.Rules
{
    /// <summary>
    /// Parsed configuration for a single validation rule from a #Rules sheet.
    /// </summary>
    public class RuleConfig
    {
        /// <summary>Target field name (matches column header)</summary>
        public string FieldName { get; set; }

        /// <summary>Rule type: range, regex, multiple, not_empty, required, enum</summary>
        public string RuleName { get; set; }

        /// <summary>Rule parameters (e.g., "0~9999" for range, "10" for multiple)</summary>
        public string Params { get; set; }

        /// <summary>Optional condition: column name to check</summary>
        public string ConditionField { get; set; }

        /// <summary>Condition operator: =, !=, &gt;=, &lt;=, &gt;, &lt;</summary>
        public string ConditionOp { get; set; }

        /// <summary>Condition value to compare against</summary>
        public string ConditionValue { get; set; }

        /// <summary>Custom error message</summary>
        public string Message { get; set; }

        public override string ToString()
        {
            var cond = string.IsNullOrEmpty(ConditionField) ? "" : $" when {ConditionField}{ConditionOp}{ConditionValue}";
            return $"  {FieldName}: {RuleName}({Params}){cond}";
        }
    }
}
