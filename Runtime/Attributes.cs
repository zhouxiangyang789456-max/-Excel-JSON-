using System;

namespace ExcelToJsonPlugin.Runtime
{
    /// <summary>
    /// 标记一个类为 Excel 数据表对应的 C# 类。
    /// 用于模式 B（反射匹配），指定该类对应哪个 Sheet。
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, Inherited = false)]
    public class ExcelTableAttribute : Attribute
    {
        /// <summary>Excel 中的 Sheet 名称</summary>
        public string SheetName { get; }
        /// <summary>Excel 文件名（可选，默认自动查找）</summary>
        public string FileName { get; set; }

        public ExcelTableAttribute(string sheetName)
        {
            SheetName = sheetName;
        }
    }

    /// <summary>
    /// 标记一个字段对应 Excel 中的某列。
    /// 列名默认与字段名一致，可手动指定。
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, Inherited = false)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>Excel 列名</summary>
        public string ColumnName { get; }

        public ExcelColumnAttribute(string columnName = null)
        {
            ColumnName = columnName;
        }
    }

    /// <summary>
    /// 标记一个字段不参与 Excel 映射。
    /// 用于 C# 中有但在 Excel 中没有对应列的字段。
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, Inherited = false)]
    public class ExcelIgnoreAttribute : Attribute { }

    /// <summary>
    /// 字段校验：数值范围
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false)]
    public class ValidateRangeAttribute : Attribute
    {
        public double Min { get; }
        public double Max { get; }
        public string Message { get; set; }

        public ValidateRangeAttribute(double min, double max)
        {
            Min = min;
            Max = max;
        }
    }

    /// <summary>
    /// 字段校验：不能为空
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false)]
    public class ValidateNotEmptyAttribute : Attribute
    {
        public string Message { get; set; }
    }

    /// <summary>
    /// 字段校验：正则匹配
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false)]
    public class ValidateRegexAttribute : Attribute
    {
        public string Pattern { get; }
        public string Message { get; set; }

        public ValidateRegexAttribute(string pattern)
        {
            Pattern = pattern;
        }
    }
}
