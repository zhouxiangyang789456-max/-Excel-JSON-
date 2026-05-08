using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections.Generic;

class GenCharSkill
{
    static string OutDir = @"D:\git\-Excel-JSON-\UnityProject\Assets\Excel\TestCases";
    static List<string> Strs;
    static Dictionary<string, int> Idx;

    static void Main()
    {
        Directory.CreateDirectory(OutDir);

        // Skill.xlsx - 技能表（被引用表）
        Make("Skill.xlsx", new[]{"Skill"}, new[]{
            D(new[]{"id","name","type","power","cost","desc"}),
            D(new[]{"int","string","string","int","int","string"}),
            D(new[]{"ID","技能名称","类型","威力","消耗","描述"}),
            D("1","重击","物理","150","30","集中力量进行一次强力攻击"),
            D("2","火球术","魔法","200","50","发射一颗灼热的火球"),
            D("3","旋风斩","物理","120","25","旋转攻击周围所有敌人"),
            D("4","治疗术","辅助","0","40","恢复目标生命值"),
            D("5","冰冻箭","魔法","180","45","射出寒冰箭矢减速敌人"),
            D("6","闪避","辅助","0","20","短时间内提高闪避率"),
        });

        // Character.xlsx - 人物表（引用技能表）
        Make("Character.xlsx", new[]{"Character"}, new[]{
            D(new[]{"id","name","job","hp","atk","skill_ids","ultimate_skill","avatar"}),
            D(new[]{"int","string","string","int","int","int[]","ref:Skill","res:Sprite"}),
            D(new[]{"ID","名称","职业","生命","攻击","拥有技能","大招","头像"}),
            D("101","亚瑟","战士","1500","120","[1,3]","1","Ui/Avatar/Arthur"),
            D("102","梅林","法师","800","200","[2,5]","2","Ui/Avatar/Merlin"),
            D("103","艾琳","牧师","700","80","[4,6]","4","Ui/Avatar/Eileen"),
            D("104","兰斯","骑士","2000","150","[1,3,5]","3","Ui/Avatar/Lance"),
        });

        Console.WriteLine("Created: Skill.xlsx + Character.xlsx");
        Console.WriteLine("Character.skill_ids type=int[]: stores [1,3] etc.");
        Console.WriteLine("Character.ultimate_skill type=ref:Skill: stores skill ID");
        Console.WriteLine("Runtime: load Character, get skill_ids[0], query Skill.Get(id)");
    }

    static string[] D(params string[] a) { return a; }

    static void Make(string fn, string[] sn, params string[][][] sheets)
    {
        string output = Path.Combine(OutDir, fn);
        string tmp = Path.Combine(Path.GetTempPath(), "xl_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);
        Strs = new List<string>(); Idx = new Dictionary<string, int>();
        foreach (var d in sheets) foreach (var r in d) foreach (var c in r) if (!IsNum(c)) AddS(c);

        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.AppendLine("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.AppendLine("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.AppendLine("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.AppendLine("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        for (int i = 0; i < sn.Length; i++)
            sb.AppendLine("<Override PartName=\"/xl/worksheets/sheet"+(i+1)+".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        sb.AppendLine("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
        sb.AppendLine("</Types>");
        W(tmp, "[Content_Types].xml", sb.ToString());

        W(tmp, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");

        sb.Clear();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.AppendLine("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
        for (int i = 0; i < sn.Length; i++)
            sb.AppendLine("<sheet name=\""+Esc(sn[i])+"\" sheetId=\""+(i+1)+"\" r:id=\"rId"+(i+1)+"\"/>");
        sb.AppendLine("</sheets></workbook>");
        W(tmp, "xl/workbook.xml", sb.ToString());

        sb.Clear();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.AppendLine("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (int i = 0; i < sn.Length; i++)
            sb.AppendLine("<Relationship Id=\"rId"+(i+1)+"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"+(i+1)+".xml\"/>");
        sb.AppendLine("<Relationship Id=\"rId_ss\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        sb.AppendLine("</Relationships>");
        W(tmp, "xl/_rels/workbook.xml.rels", sb.ToString());

        sb.Clear();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.AppendLine("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\""+Strs.Count+"\" uniqueCount=\""+Strs.Count+"\">");
        foreach (var s in Strs) sb.AppendLine("<si><t>"+Esc(s)+"</t></si>");
        sb.AppendLine("</sst>");
        W(tmp, "xl/sharedStrings.xml", sb.ToString());

        for (int si = 0; si < sheets.Length; si++)
        {
            sb.Clear();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.AppendLine("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            var d = sheets[si];
            for (int ri = 0; ri < d.Length; ri++)
            {
                sb.Append("<row r=\""+(ri+1)+"\">");
                for (int ci = 0; ci < d[ri].Length; ci++)
                {
                    string v = d[ri][ci];
                    char col = (char)('A'+ci);
                    string ref_ = col.ToString()+(ri+1);
                    if (IsNum(v))
                        sb.Append("<c r=\""+ref_+"\"><v>"+v+"</v></c>");
                    else
                        sb.Append("<c r=\""+ref_+"\" t=\"s\"><v>"+Idx[v]+"</v></c>");
                }
                sb.AppendLine("</row>");
            }
            sb.AppendLine("</sheetData></worksheet>");
            W(tmp, "xl/worksheets/sheet"+(si+1)+".xml", sb.ToString());
        }

        if (File.Exists(output)) File.Delete(output);
        ZipFile.CreateFromDirectory(tmp, output, CompressionLevel.Optimal, false);
        Directory.Delete(tmp, true);
    }

    static bool IsNum(string s) { if (string.IsNullOrEmpty(s)) return true; decimal x; return decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out x); }
    static void AddS(string s) { if (!Idx.ContainsKey(s)) { Idx[s]=Strs.Count; Strs.Add(s); } }
    static string Esc(string s) { return s.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace("\"","&quot;"); }
    static void W(string b, string r, string c) { string f=Path.Combine(b,r); Directory.CreateDirectory(Path.GetDirectoryName(f)); File.WriteAllText(f,c,Encoding.UTF8); }
}
