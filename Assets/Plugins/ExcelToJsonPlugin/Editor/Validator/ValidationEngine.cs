using System.Collections.Generic;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Core.Models;
using ExcelToJsonPlugin.Editor.Validator.Rules;

namespace ExcelToJsonPlugin.Editor.Validator
{
    public class ValidationEngine
    {
        private static readonly List<IStructureRule> StructureRules = new List<IStructureRule>
        {
            new FieldNameUniqueRule(),
            new TypeValidityRule(),
            new HeaderCompletenessRule(),
            new FieldNameSanityRule(),
        };

        private static readonly List<IDataRule> DataRules = new List<IDataRule>
        {
            // Note: TypeMatch is handled by DataParser during export,
            // so it's excluded here to avoid double-reporting.
            // Use ValidateAll() for standalone use that includes TypeMatch.
            new IdRequiredRule(),
            new IdUniqueRule(),
            new RequiredFieldRule(),
            new EnumSanityRule(),
            new ResPathRule(),
            new FormulaDetectionRule(),
        };

        /// <summary>
        /// Standalone validation that includes TypeMatch rule (for validate-only flows
        /// that don't go through DataParser).
        /// </summary>
        public static ValidationReport ValidateAll(
            List<List<string>> rows,
            TableSchema schema,
            string fileName)
        {
            var report = Validate(rows, schema, fileName);

            var typeErrors = new TypeMatchRule().Validate(rows, schema, fileName);
            foreach (var err in typeErrors)
                report.Errors.Add(err);

            return report;
        }

        public static ValidationReport Validate(
            List<List<string>> rows,
            TableSchema schema,
            string fileName,
            List<RuleConfig> customRules = null)
        {
            var report = new ValidationReport();

            if (rows == null || rows.Count == 0)
            {
                report.Add(fileName, schema?.TableName ?? "", 0, "", "",
                    "EmptySheet", "Sheet has no rows.", ErrorLevel.Error);
                return report;
            }

            // Stage 1: Structure
            RunStage1(schema, fileName, report);

            // Stage 2: Data
            RunStage2(rows, schema, fileName, report, customRules);

            // Stage 3: References (Sprint 3)
            RunStage3(rows, schema, fileName, report);

            return report;
        }

        private static void RunStage1(TableSchema schema, string fileName, ValidationReport report)
        {
            foreach (var rule in StructureRules)
            {
                var errors = rule.Validate(schema, fileName);
                foreach (var err in errors)
                    report.Errors.Add(err);
            }
        }

        private static void RunStage2(
            List<List<string>> rows, TableSchema schema,
            string fileName, ValidationReport report,
            List<RuleConfig> customRules = null)
        {
            // Built-in data rules
            foreach (var rule in DataRules)
            {
                var errors = rule.Validate(rows, schema, fileName);
                foreach (var err in errors)
                    report.Errors.Add(err);
            }

            // Custom rules from #Rules sheet
            if (customRules != null && customRules.Count > 0)
            {
                var customRule = new CustomDataRule(customRules);
                var customErrors = customRule.Validate(rows, schema, fileName);
                foreach (var err in customErrors)
                    report.Errors.Add(err);
            }
        }

        private static void RunStage3(
            List<List<string>> rows, TableSchema schema,
            string fileName, ValidationReport report)
        {
            // Sprint 3: Cross-file reference integrity check
            // - ref:TableName → verify referenced ID exists in target table
            // - enum:TableName → verify value is in named enum Sheet
            // - Cycle detection in reference chains
        }

        /// <summary>
        /// Collect all unique sheet names referenced via ref:Type and enum:Type
        /// (used by Sprint 3 for preloading dependent tables).
        /// </summary>
        public static HashSet<string> GetReferencedTableNames(TableSchema schema)
        {
            var names = new HashSet<string>();
            foreach (var field in schema.Fields)
            {
                if (field.IsCompositeType && !string.IsNullOrEmpty(field.CompositeParam))
                {
                    if (field.NormalizedType?.StartsWith("ref:") == true ||
                        field.NormalizedType?.StartsWith("enum:") == true)
                    {
                        names.Add(field.CompositeParam);
                    }
                }
            }
            return names;
        }
    }
}
