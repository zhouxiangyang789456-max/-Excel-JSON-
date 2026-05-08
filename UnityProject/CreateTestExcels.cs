using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections.Generic;

class CreateTestExcels
{
    static string OutDir = @"D:\git\-Excel-JSON-\UnityProject\Assets\Excel\TestCases";
    static List<string> Strs;
    static Dictionary<string, int> Idx;

    static void Main()
    {
        Directory.CreateDirectory(OutDir);
        // T1: Types
        Make("Test01_Types.xlsx", new[]{"Types"}, new[]{
            D(new[]{"id","int_val","float_val","str_val","bool_val","empty_col"}),
            D(new[]{"int","int","float","string","bool","string"}),
            D(new[]{"ID","???","???","???","???","??"}),
            D("1","100","3.14","hello","true",""),
            D("2","-50","0.0","","false",""),
            D("3","0","-1.5","!@#$%","?",""),
            D("4","99999","0.001","?? ?","?",""),
            D("5","-1","100.0","multi","1",""),
        });
        // T2: Arrays
        Make("Test02_Arrays.xlsx", new[]{"Arrays"}, new[]{
            D(new[]{"id","json_arr","pipe_arr","single","empty_arr","float_arr"}),
            D(new[]{"int","int[]","int[]","int[]","int[]","float[]"}),
            D(new[]{"ID","JSON??","????","????","???","?????"}),
            D("1","[1,2,3]","1|2|3","42","[]","[1.5,2.0,3.14]"),
            D("2","[10,20]","10|20|30|40","7","","[0.1,0.2]"),
            D("3","[100]","","99","[1]","[]"),
        });
        // T3: Unity
        Make("Test03_UnityTypes.xlsx", new[]{"UnityTypes"}, new[]{
            D(new[]{"id","pos2","pos3","color_hex","color_rgba"}),
            D(new[]{"int","Vector2","Vector3","Color","Color"}),
            D(new[]{"ID","????","????","??Hex","??RGBA"}),
            D("1","[1,2]","[1,2,3]","#FF0000","[0,1,0,1]"),
            D("2","[0.5,1.5]","[0,0,0]","#00FF00FF","[0.5,0.5,0.5,0.8]"),
            D("3","[-1,0]","[100,200,300]","#0000FF","[1,0,0,0.5]"),
        });
        // T4: Composite
        Make("Test04_Composite.xlsx", new[]{"Composite","RefTarget"}, new[]{
            D(new[]{"id","name","ref_val","enum_val","icon","prefab","sound"}),
            D(new[]{"int","string","ref:RefTarget","enum:Quality","res:Sprite","res:GameObject","res:AudioClip"}),
            D(new[]{"ID","??","??","??","??","???","??"}),
            D("1","???","100","1","Ui/Icon/Main","Prefabs/Main","Audio/SFX/Click"),
            D("2","???","101","2","Ui/Icon/Sub","",""),
            D("3","???","","3","","Prefabs/Empty",""),
        }, new[]{
            D(new[]{"id","name"}),
            D(new[]{"int","string"}),
            D(new[]{"ID","??"}),
            D("100","??A"),
            D("101","??B"),
            D("102","??C"),
        });
        // T5: Large
        var lr = new List<string[]>();
        lr.Add(D(new[]{"id","name","value","category"}));
        lr.Add(D(new[]{"int","string","int","string"}));
        lr.Add(D(new[]{"ID","??","??","??"}));
        for (int i = 1; i <= 200; i++)
            lr.Add(D(i.ToString(), "Entry_" + i, (i*10).ToString(), "Cat_" + ((i%10)+1)));
        Make("Test05_Large.xlsx", new[]{"LargeData"}, lr.ToArray());
        // T6: Skip sheets
        Make("Test06_SkipSheets.xlsx", new[]{"Weapon","_Notes","Armor","#Internal","Skill"}, new[]{
            D(new[]{"id","name"}), D(new[]{"int","string"}), D(new[]{"ID","??"}), D("1","??A"),
        }, new[]{
            D(new[]{"note"}), D(new[]{"string"}), D(new[]{"??"}), D("skip_me"),
        }, new[]{
            D(new[]{"id","name"}), D(new[]{"int","string"}), D(new[]{"ID","??"}), D("201","??A"),
        }, new[]{
            D(new[]{"internal"}), D(new[]{"string"}), D(new[]{"??"}), D("skip_too"),
        }, new[]{
            D(new[]{"id","name"}), D(new[]{"int","string"}), D(new[]{"ID","??"}), D("301","??A"),
        });
        // T7: Bool edge
        Make("Test07_BoolEdge.xlsx", new[]{"BoolTest"}, new[]{
            D(new[]{"id","b1","b2","b3","b4","b5","b6","b7","b8"}),
            D(new[]{"int","bool","bool","bool","bool","bool","bool","bool","bool"}),
            D(new[]{"ID","T/F","0/1","?/?","y/n","Y/N","True","TRUE","?"}),
            D("1","true","1","?","yes","y","True","TRUE",""),
            D("2","false","0","?","no","n","False","FALSE","true"),
            D("3","TRUE","0","?","YES","N","false","true","false"),
        });
        Console.WriteLine("All 7 test Excels created.");
    }

    static string[] D(params string[] a) { return a; }

    static void Make(string fn, string[] sn, params string[][][] sheets)
    {
        string output = Path.Combine(OutDir, fn);
        string tmp = Path.Combine(Path.GetTempPath(), "xl_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);

        Strs = new List<string>(); Idx = new Dictionary<string, int>();
        foreach (var d in sheets) foreach (var r in d) foreach (var c in r) if (!IsNum(c)) AddS(c);

        // [Content_Types].xml
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
        Console.WriteLine("  " + fn);
    }

    static bool IsNum(string s) { if (string.IsNullOrEmpty(s)) return true; decimal x; return decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out x); }
    static void AddS(string s) { if (!Idx.ContainsKey(s)) { Idx[s]=Strs.Count; Strs.Add(s); } }
    static string Esc(string s) { return s.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace("\"","&quot;"); }
    static void W(string b, string r, string c) { string f=Path.Combine(b,r); Directory.CreateDirectory(Path.GetDirectoryName(f)); File.WriteAllText(f,c,Encoding.UTF8); }
}
