using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor
{
    /// <summary>
    /// Generates sample Excel files for demo/testing purposes.
    /// Menu: Window > Excel Data Manager > Generate Sample Data
    /// </summary>
    public static class SampleGenerator
    {
        private const string OutputDir = "Assets/Excel/";

        [MenuItem("Window/Excel Data Manager/Generate Sample Data", priority = 101)]
        public static void GenerateSampleData()
        {
            var outputDir = OutputDir;
            Directory.CreateDirectory(outputDir);

            GenerateItemData(outputDir);
            GenerateSkillEnum(outputDir);

            AssetDatabase.Refresh();
            EditorUtility.DisplayDialog(Core.Loc.Tr("sample_done_title"),
                Core.Loc.Tr("sample_done_msg", outputDir),
                "OK");

            Debug.Log("[ExcelToJSON] Sample data generated. Select the files in the Editor Window and export.");
        }

        private static void GenerateItemData(string outputDir)
        {
            var workbook = new XSSFWorkbook();

            // Sheet 1: Weapon
            CreateSheet(workbook, "Weapon",
                new[] { "id", "name", "attack", "defense", "price", "rare", "element", "desc" },
                new[] { "int", "string", "int", "int", "int", "enum:SkillEnum", "string", "string" },
                new[] { "ID", "名称", "攻击力", "防御力", "价格", "稀有度", "属性", "描述" },
                new[]
                {
                    new[] { "1001", "新手剑", "12", "0", "100", "1", "火", "一把普通的铁剑" },
                    new[] { "1002", "铁剑", "25", "5", "300", "2", "无", "经过锻造的坚固铁剑" },
                    new[] { "1003", "魔导杖", "8", "2", "500", "2", "冰", "可以释放寒冰法术的魔杖" },
                    new[] { "1004", "龙牙匕首", "35", "0", "800", "3", "火", "用龙牙打造，锋利无比" },
                    new[] { "1005", "圣光剑", "45", "10", "1500", "4", "光", "传说中圣骑士的佩剑" },
                    new[] { "1006", "暗影弓", "30", "0", "1200", "3", "暗", "暗影精灵的弓箭" },
                    new[] { "1007", "雷霆锤", "55", "15", "2000", "5", "雷", "雷电之神的战锤" },
                });

            // Sheet 2: Armor
            CreateSheet(workbook, "Armor",
                new[] { "id", "name", "defense", "price", "weight", "type" },
                new[] { "int", "string", "int", "int", "float", "string" },
                new[] { "ID", "名称", "防御力", "价格", "重量", "类型" },
                new[]
                {
                    new[] { "2001", "皮甲", "8", "200", "2.5", "轻甲" },
                    new[] { "2002", "锁子甲", "18", "600", "8.0", "中甲" },
                    new[] { "2003", "板甲", "35", "1500", "15.0", "重甲" },
                    new[] { "2004", "魔法袍", "5", "800", "3.0", "布甲" },
                    new[] { "2005", "龙鳞甲", "50", "3000", "20.0", "重甲" },
                });

            // Sheet 3: _Notes (hidden sheet, should be skipped)
            var notesSheet = workbook.CreateSheet("_Notes");
            notesSheet.CreateRow(0).CreateCell(0).SetCellValue("This is a hidden notes sheet — should be skipped by the plugin.");

            // Save
            var filePath = Path.Combine(outputDir, "Item.xlsx");
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                workbook.Write(fs);
            workbook.Close();

            Debug.Log($"[ExcelToJSON] Generated: {filePath}");
        }

        private static void GenerateSkillEnum(string outputDir)
        {
            var workbook = new XSSFWorkbook();
            CreateSheet(workbook, "SkillEnum",
                new[] { "id", "name", "color" },
                new[] { "int", "string", "Color" },
                new[] { "ID", "名称", "颜色" },
                new[]
                {
                    new[] { "1", "Common", "#808080" },
                    new[] { "2", "Uncommon", "#00FF00" },
                    new[] { "3", "Rare", "#0000FF" },
                    new[] { "4", "Epic", "#FF00FF" },
                    new[] { "5", "Legendary", "#FFD700" },
                });

            var filePath = Path.Combine(outputDir, "SkillEnum.xlsx");
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                workbook.Write(fs);
            workbook.Close();

            Debug.Log($"[ExcelToJSON] Generated: {filePath}");
        }

        private static ISheet CreateSheet(
            IWorkbook workbook, string sheetName,
            string[] fieldNames, string[] types,
            string[] comments, string[][] dataRows)
        {
            var sheet = workbook.CreateSheet(sheetName);

            // Styles
            var headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerFont.FontHeightInPoints = 11;
            var headerStyle = workbook.CreateCellStyle();
            headerStyle.SetFont(headerFont);
            headerStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;

            var typeFont = workbook.CreateFont();
            typeFont.FontName = "Consolas";
            typeFont.FontHeightInPoints = 10;
            typeFont.Color = IndexedColors.Grey50Percent.Index;
            var typeStyle = workbook.CreateCellStyle();
            typeStyle.SetFont(typeFont);

            // Row 1: Field names
            var nameRow = sheet.CreateRow(0);
            for (int i = 0; i < fieldNames.Length; i++)
            {
                var cell = nameRow.CreateCell(i);
                cell.SetCellValue(fieldNames[i]);
                cell.CellStyle = headerStyle;
                sheet.SetColumnWidth(i, 16 * 256);
            }

            // Row 2: Types
            var typeRow = sheet.CreateRow(1);
            for (int i = 0; i < types.Length; i++)
            {
                var cell = typeRow.CreateCell(i);
                cell.SetCellValue(types[i]);
                cell.CellStyle = typeStyle;
            }

            // Row 3: Comments
            var commentRow = sheet.CreateRow(2);
            for (int i = 0; i < comments.Length; i++)
                commentRow.CreateCell(i).SetCellValue(comments[i]);

            // Row 4+: Data
            for (int r = 0; r < dataRows.Length; r++)
            {
                var row = sheet.CreateRow(3 + r);
                for (int c = 0; c < dataRows[r].Length; c++)
                    row.CreateCell(c).SetCellValue(dataRows[r][c]);
            }

            return sheet;
        }
    }
}
