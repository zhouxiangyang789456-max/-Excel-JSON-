using System.Collections.Generic;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Validator.Rules
{
    public interface IStructureRule
    {
        List<ValidationError> Validate(TableSchema schema, string fileName);
    }

    public interface IDataRule
    {
        List<ValidationError> Validate(List<List<string>> rows, TableSchema schema, string fileName);
    }
}
