using System;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// 读取 .xlsx / .xls 文件，输出结构化行列数据。
    /// 处理合并单元格、隐藏行列、公式缓存值。
    /// </summary>
    public static class ExcelReader
    {
        /// <summary>
        /// 读取结果：{ sheetName → [[cellString, ...], ...] }
        /// 每个 Sheet 的值是一个二维字符串列表，row[col] 获取单元格文本。
        /// </summary>
        public class ReadResult
        {
            /// <summary>Sheet 名称 → 该 Sheet 的所有行数据</summary>
            public Dictionary<string, List<List<string>>> Sheets { get; set; }
                = new Dictionary<string, List<List<string>>>();

            /// <summary>Sheet 名称列表（保持 Excel 中的顺序）</summary>
            public List<string> SheetNames { get; set; }
                = new List<string>();

            /// <summary>每个 Sheet 的隐藏状态</summary>
            public Dictionary<string, bool> SheetHidden { get; set; }
                = new Dictionary<string, bool>();

            /// <summary>原始文件名</summary>
            public string FileName { get; set; }
        }

        /// <summary>
        /// 读取一个 Excel 文件的所有 Sheet。
        /// </summary>
        /// <param name="filePath">.xlsx 或 .xls 文件路径</param>
        /// <param name="skipHiddenRows">是否跳过隐藏行</param>
        /// <param name="skipHiddenColumns">是否跳过隐藏列</param>
        /// <param name="skipEmptyRows">是否跳过全空行</param>
        public static ReadResult Read(string filePath,
            bool skipHiddenRows = true,
            bool skipHiddenColumns = true,
            bool skipEmptyRows = true)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Excel 文件不存在: {filePath}");

            var result = new ReadResult
            {
                FileName = Path.GetFileName(filePath),
            };

            IWorkbook workbook;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var ext = Path.GetExtension(filePath).ToLower();
                if (ext == ".xlsx")
                    workbook = new XSSFWorkbook(fs);
                else if (ext == ".xls")
                    workbook = new HSSFWorkbook(fs);
                else
                    throw new NotSupportedException($"不支持的 Excel 格式: {ext}，仅支持 .xlsx 和 .xls");
            }

            for (int si = 0; si < workbook.NumberOfSheets; si++)
            {
                var sheet = workbook.GetSheetAt(si);
                var sheetName = sheet.SheetName.Trim();

                result.SheetNames.Add(sheetName);
                result.SheetHidden[sheetName] = workbook.IsSheetHidden(si)
                    || workbook.IsSheetVeryHidden(si);

                var rows = ReadSheet(sheet, skipHiddenRows, skipHiddenColumns, skipEmptyRows);
                result.Sheets[sheetName] = rows;
            }

            workbook.Close();
            return result;
        }

        /// <summary>
        /// 读取单个 Sheet 的所有行。
        /// </summary>
        private static List<List<string>> ReadSheet(
            ISheet sheet,
            bool skipHiddenRows,
            bool skipHiddenColumns,
            bool skipEmptyRows)
        {
            var rows = new List<List<string>>();

            // 先确定该 Sheet 的最大列数（遍历所有行）
            int maxCols = 0;
            for (int ri = sheet.FirstRowNum; ri <= sheet.LastRowNum; ri++)
            {
                var row = sheet.GetRow(ri);
                if (row == null) continue;
                if (row.LastCellNum > maxCols)
                    maxCols = row.LastCellNum;
            }

            // 收集合并区域信息
            var mergedRegions = new Dictionary<(int row, int col), string>();
            for (int mi = 0; mi < sheet.NumMergedRegions; mi++)
            {
                var region = sheet.GetMergedRegion(mi);
                var topLeftValue = GetCellValue(sheet, region.FirstRow, region.FirstColumn);
                for (int r = region.FirstRow; r <= region.LastRow; r++)
                {
                    if (skipHiddenRows && sheet.GetRow(r)?.ZeroHeight == true) continue;
                    for (int c = region.FirstColumn; c <= region.LastColumn; c++)
                    {
                        if (skipHiddenColumns && sheet.IsColumnHidden(c)) continue;
                        mergedRegions[(r, c)] = topLeftValue;
                    }
                }
            }

            // 读取所有行
            for (int ri = sheet.FirstRowNum; ri <= sheet.LastRowNum; ri++)
            {
                var row = sheet.GetRow(ri);

                // 跳过隐藏行
                if (skipHiddenRows && row?.ZeroHeight == true) continue;

                var cells = new List<string>();
                bool allEmpty = true;

                for (int ci = 0; ci < maxCols; ci++)
                {
                    // 跳过隐藏列
                    if (skipHiddenColumns && sheet.IsColumnHidden(ci))
                        continue;

                    string cellValue;

                    // 合并单元格优先
                    if (mergedRegions.TryGetValue((ri, ci), out var mergedValue))
                    {
                        cellValue = mergedValue;
                    }
                    else
                    {
                        cellValue = row != null ? GetCellValue(row, ci) : string.Empty;
                    }

                    if (!string.IsNullOrEmpty(cellValue))
                        allEmpty = false;

                    cells.Add(cellValue ?? string.Empty);
                }

                // 跳过全空行
                if (skipEmptyRows && allEmpty)
                    continue;

                // 确保总是有数据（即使整行全空但 skipEmptyRows=false）
                rows.Add(cells);
            }

            // 去除尾部连续的空行（从最后一行往上找第一个非空行）
            if (skipEmptyRows)
            {
                while (rows.Count > 0 && IsRowEmpty(rows[rows.Count - 1]))
                {
                    rows.RemoveAt(rows.Count - 1);
                }
            }

            return rows;
        }

        /// <summary>
        /// 获取指定单元格的字符串值。
        /// 公式单元格返回缓存值。
        /// </summary>
        private static string GetCellValue(ISheet sheet, int rowIndex, int colIndex)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null) return string.Empty;
            return GetCellValue(row, colIndex);
        }

        /// <summary>
        /// 获取指定单元格的字符串值。
        /// </summary>
        private static string GetCellValue(IRow row, int colIndex)
        {
            var cell = row.GetCell(colIndex);
            if (cell == null) return string.Empty;

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    // 判断是否为日期格式
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try
                        {
                            // NPOI 2.7 DateCellValue returns DateTime
                            var dt = cell.DateCellValue;
                            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", dt);
                        }
                        catch
                        {
                            return cell.NumericCellValue.ToString();
                        }
                    }
                    // 判断是否为整数（避免 100.0 这种显示）
                    var numVal = cell.NumericCellValue;
                    if (Math.Abs(numVal - Math.Round(numVal)) < 1e-10)
                        return ((long)numVal).ToString();
                    return numVal.ToString();

                case CellType.String:
                    return cell.StringCellValue?.Trim() ?? string.Empty;

                case CellType.Boolean:
                    return cell.BooleanCellValue ? "true" : "false";

                case CellType.Formula:
                    // 公式 → 尝试获取缓存值
                    try
                    {
                        // 先尝试作为数值结果
                        var cachedNum = cell.NumericCellValue;
                        if (Math.Abs(cachedNum - Math.Round(cachedNum)) < 1e-10)
                            return ((long)cachedNum).ToString();
                        return cachedNum.ToString();
                    }
                    catch
                    {
                        try
                        {
                            return cell.StringCellValue?.Trim() ?? string.Empty;
                        }
                        catch
                        {
                            return cell.ToString()?.Trim() ?? string.Empty;
                        }
                    }

                case CellType.Blank:
                    return string.Empty;

                case CellType.Error:
                    return $"#ERROR:{(int)cell.ErrorCellValue}";

                default:
                    return cell.ToString()?.Trim() ?? string.Empty;
            }
        }

        /// <summary>
        /// 判断一行是否全为空字符串。
        /// </summary>
        private static bool IsRowEmpty(List<string> row)
        {
            foreach (var cell in row)
            {
                if (!string.IsNullOrEmpty(cell))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// 检查一行是否全为空（静态工具方法）
        /// </summary>
        public static bool IsRowEmpty(IList<string> row)
        {
            foreach (var cell in row)
            {
                if (!string.IsNullOrEmpty(cell))
                    return false;
            }
            return true;
        }
    }
}
