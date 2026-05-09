using System.Collections.Generic;

namespace ExcelToJsonPlugin.Editor.Core.Models
{
    /// <summary>
    /// 单条校验结果。
    /// </summary>
    public class ValidationError
    {
        public string FileName;
        public string SheetName;
        public int Row;
        public string ColumnName;
        public string RawValue;
        public string RuleName;
        public string Message;
        public ErrorLevel Level;

        public override string ToString()
        {
            var prefix = Level == ErrorLevel.Error ? "[E]" : Level == ErrorLevel.Warning ? "[W]" : "[I]";
            return $"{prefix} {FileName}/{SheetName} Row={Row} Col={ColumnName}: {Message}";
        }
    }

    public enum ErrorLevel
    {
        Error,
        Warning,
        Info,
    }

    /// <summary>
    /// 一次校验的完整结果集合。
    /// </summary>
    public class ValidationReport
    {
        public List<ValidationError> Errors { get; set; } = new List<ValidationError>();

        public int ErrorCount => Errors.FindAll(e => e.Level == ErrorLevel.Error).Count;
        public int WarningCount => Errors.FindAll(e => e.Level == ErrorLevel.Warning).Count;
        public int InfoCount => Errors.FindAll(e => e.Level == ErrorLevel.Info).Count;
        public bool HasErrors => ErrorCount > 0;
        public bool HasWarnings => WarningCount > 0;

        public void Add(ValidationError error)
        {
            Errors.Add(error);
        }

        public void Add(string file, string sheet, int row, string col,
            string value, string rule, string msg, ErrorLevel level)
        {
            Errors.Add(new ValidationError
            {
                FileName = file,
                SheetName = sheet,
                Row = row,
                ColumnName = col,
                RawValue = value,
                RuleName = rule,
                Message = msg,
                Level = level,
            });
        }
    }
}
