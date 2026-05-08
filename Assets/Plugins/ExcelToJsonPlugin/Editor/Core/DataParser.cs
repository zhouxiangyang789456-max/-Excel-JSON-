using System;
using System.Collections.Generic;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// 根据 TableSchema 解析数据行，将每行转换为 Dictionary。
    /// </summary>
    public static class DataParser
    {
        /// <summary>
        /// 解析一个 Sheet 的所有数据行。
        /// </summary>
        /// <param name="rows">Sheet 的原始行数据（包含表头）</param>
        /// <param name="schema">该 Sheet 的结构定义</param>
        /// <param name="errors">输出解析过程中的错误</param>
        /// <returns>数据行列表，每行是一个 Dictionary{字段名 → 值}</returns>
        public static List<Dictionary<string, object>> ParseData(
            List<List<string>> rows,
            TableSchema schema,
            List<ValidationError> errors = null)
        {
            var data = new List<Dictionary<string, object>>();
            errors ??= new List<ValidationError>();

            int dataStartIdx = schema.DataStartRow - 1; // 转 0-based

            // 检查数据起始行是否在范围内
            if (dataStartIdx >= rows.Count)
            {
                // 没有数据行，不算错误
                return data;
            }

            for (int ri = dataStartIdx; ri < rows.Count; ri++)
            {
                var rawRow = rows[ri];

                // 跳过全空行
                if (ExcelReader.IsRowEmpty(rawRow))
                    continue;

                var rowDict = new Dictionary<string, object>();
                bool rowHasError = false;

                for (int fi = 0; fi < schema.Fields.Count; fi++)
                {
                    var field = schema.Fields[fi];
                    var colIdx = field.ColumnIndex;

                    string rawValue = colIdx < rawRow.Count ? rawRow[colIdx] : string.Empty;

                    var converted = TypeMapper.ConvertValue(rawValue, field, out var errorMsg);

                    if (errorMsg != null)
                    {
                        errors.Add(new ValidationError
                        {
                            FileName = schema.FileName,
                            SheetName = schema.TableName,
                            Row = ri + 1, // 转回 1-based 行号
                            ColumnName = field.Name,
                            RawValue = rawValue,
                            RuleName = "TypeMatch",
                            Message = errorMsg,
                            Level = ErrorLevel.Error,
                        });
                        rowHasError = true;
                    }

                    rowDict[field.CsFieldName] = converted;
                }

                // 有错误的行也加入（使用默认值），便于部分导出
                data.Add(rowDict);
            }

            return data;
        }
    }
}
