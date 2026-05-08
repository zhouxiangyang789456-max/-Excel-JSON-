using System.Collections.Generic;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Validator.Rules
{
    /// <summary>Detect duplicate field names in the header row.</summary>
    public class FieldNameUniqueRule : IStructureRule
    {
        public List<ValidationError> Validate(TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();
            var seen = new HashSet<string>();

            foreach (var field in schema.Fields)
            {
                if (string.IsNullOrEmpty(field.Name)) continue;
                if (!seen.Add(field.Name))
                {
                    errors.Add(new ValidationError
                    {
                        FileName = fileName,
                        SheetName = schema.TableName,
                        Row = 1,
                        ColumnName = field.Name,
                        RawValue = field.Name,
                        RuleName = "FieldNameUnique",
                        Message = $"Duplicate field name: \"{field.Name}\"",
                        Level = ErrorLevel.Error,
                    });
                }
            }

            return errors;
        }
    }

    /// <summary>Verify every field declares a recognized type.</summary>
    public class TypeValidityRule : IStructureRule
    {
        private static readonly HashSet<string> ValidTypes = new HashSet<string>
        {
            "int", "float", "string", "bool", "boolean",
            "int[]", "float[]", "string[]",
            "vec2", "vector2", "vec3", "vector3", "color",
            "json", "object", "loc",
        };

        public List<ValidationError> Validate(TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();

            for (int fi = 0; fi < schema.Fields.Count; fi++)
            {
                var field = schema.Fields[fi];
                if (string.IsNullOrEmpty(field.RawType)) continue;

                var rawLower = field.RawType.Trim().ToLower();
                if (ValidTypes.Contains(rawLower)) continue;
                if (rawLower.StartsWith("ref:") || rawLower.StartsWith("enum:") || rawLower.StartsWith("res")) continue;

                errors.Add(new ValidationError
                {
                    FileName = fileName,
                    SheetName = schema.TableName,
                    Row = 2,
                    ColumnName = field.Name,
                    RawValue = field.RawType,
                    RuleName = "TypeValidity",
                    Message = $"Unrecognized type: \"{field.RawType}\". Supported: int, float, string, bool, int[], float[], string[], Vector2, Vector3, Color, json, loc, ref:Table, enum:Table, res[:Type]",
                    Level = ErrorLevel.Warning,
                });
            }

            return errors;
        }
    }

    /// <summary>Ensure header row has at least one non-empty field name.</summary>
    public class HeaderCompletenessRule : IStructureRule
    {
        public List<ValidationError> Validate(TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();

            if (schema.Fields.Count == 0 || schema.Fields.All(f => string.IsNullOrWhiteSpace(f.Name)))
            {
                errors.Add(new ValidationError
                {
                    FileName = fileName,
                    SheetName = schema.TableName,
                    Row = 1,
                    RuleName = "HeaderCompleteness",
                    Message = "Header row is empty — no field names found.",
                    Level = ErrorLevel.Error,
                });
            }

            return errors;
        }
    }

    /// <summary>Check that field names don't contain obviously invalid characters.</summary>
    public class FieldNameSanityRule : IStructureRule
    {
        public List<ValidationError> Validate(TableSchema schema, string fileName)
        {
            var errors = new List<ValidationError>();

            foreach (var field in schema.Fields)
            {
                if (string.IsNullOrEmpty(field.Name)) continue;

                // Check for Excel formula remnants or control characters
                if (field.Name.StartsWith("="))
                {
                    errors.Add(new ValidationError
                    {
                        FileName = fileName,
                        SheetName = schema.TableName,
                        Row = 1,
                        ColumnName = field.Name,
                        RawValue = field.Name,
                        RuleName = "FieldNameSanity",
                        Message = $"Field name \"{field.Name}\" looks like a formula — did you paste from Excel?",
                        Level = ErrorLevel.Error,
                    });
                }
            }

            return errors;
        }
    }
}
