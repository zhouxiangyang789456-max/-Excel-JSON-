using System.Collections.Generic;

namespace ExcelToJsonPlugin.Editor.Core.Models
{
    /// <summary>
    /// 一个 Sheet 的结构定义。
    /// </summary>
    public class TableSchema
    {
        /// <summary>Sheet 名称（也作为表名）</summary>
        public string TableName { get; set; }

        /// <summary>所属 Excel 文件名</summary>
        public string FileName { get; set; }

        /// <summary>字段定义列表（按列顺序排列）</summary>
        public List<FieldDef> Fields { get; set; } = new List<FieldDef>();

        /// <summary>注释行原文（第 3 行），可选</summary>
        public List<string> Comments { get; set; } = new List<string>();

        /// <summary>数据起始行号（1-based）</summary>
        public int DataStartRow { get; set; } = 4;

        public override string ToString()
        {
            return $"[{TableName}] {Fields.Count} fields, from {FileName}";
        }
    }

    /// <summary>
    /// 单个字段的元数据定义。
    /// </summary>
    public class FieldDef
    {
        /// <summary>字段名（Excel 第 1 行）</summary>
        public string Name { get; set; }

        /// <summary>C# 合法字段名（自动转换）</summary>
        public string CsFieldName { get; set; }

        /// <summary>Excel 声明的类型字符串（第 2 行）</summary>
        public string RawType { get; set; }

        /// <summary>标准化类型（如 "int", "string", "int[]", "ref:Skill"）</summary>
        public string NormalizedType { get; set; }

        /// <summary>映射到的 C# 类型名</summary>
        public string CsTypeName { get; set; }

        /// <summary>是否为复合类型（ref/enum/res/json/loc）</summary>
        public bool IsCompositeType { get; set; }

        /// <summary>复合类型的参数（如 ref:Skill 的 "Skill"）</summary>
        public string CompositeParam { get; set; }

        /// <summary>注释/说明（Excel 第 3 行对应列）</summary>
        public string Comment { get; set; }

        /// <summary>在 Excel 中的列索引（0-based）</summary>
        public int ColumnIndex { get; set; }

        public override string ToString()
        {
            return $"{Name} ({NormalizedType}) → {CsFieldName} : {CsTypeName}";
        }
    }
}
