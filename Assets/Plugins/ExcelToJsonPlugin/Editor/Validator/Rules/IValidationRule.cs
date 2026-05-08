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

    /// <summary>
    /// Stage 3 跨表引用校验规则接口。
    /// 接受被引用表的 ID 集合作为上下文。
    /// </summary>
    public interface IStage3Rule
    {
        /// <param name="rows">当前 Sheet 的数据行</param>
        /// <param name="schema">当前表的结构定义</param>
        /// <param name="fileName">当前 Excel 文件名</param>
        /// <param name="referencedIds">被引用表名 → 该表所有 ID 的集合</param>
        List<ValidationError> Validate(
            List<List<string>> rows,
            TableSchema schema,
            string fileName,
            Dictionary<string, HashSet<int>> referencedIds);
    }
}
