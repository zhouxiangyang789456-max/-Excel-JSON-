using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelToJsonPlugin.Editor.Core.Models;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// Saves and loads schema snapshots for change detection.
    /// Snapshots stored at .cache/schema_snapshots/{TableName}.json
    /// </summary>
    public static class SchemaSnapshot
    {
        private const string SnapshotsDir = ".cache/schema_snapshots";

        [Serializable]
        public class Snapshot
        {
            public string excelFileName;
            public string sheetName;
            public int version = 1;
            public List<ColumnInfo> columns = new List<ColumnInfo>();

            [Serializable]
            public class ColumnInfo
            {
                public string name;
                public string type;
                public int index;
            }
        }

        public static Snapshot FromSchema(TableSchema schema, string excelFileName)
        {
            var snap = new Snapshot
            {
                excelFileName = excelFileName,
                sheetName = schema.TableName,
            };

            foreach (var field in schema.Fields)
            {
                snap.columns.Add(new Snapshot.ColumnInfo
                {
                    name = field.Name,
                    type = field.RawType ?? field.NormalizedType ?? "",
                    index = field.ColumnIndex,
                });
            }

            return snap;
        }

        public static void Save(Snapshot snapshot, string excelDir)
        {
            var dir = Path.Combine(excelDir, SnapshotsDir);
            Directory.CreateDirectory(dir);

            var path = Path.Combine(dir, $"{snapshot.sheetName}.json");
            var json = SerializeSnapshot(snapshot);
            File.WriteAllText(path, json, Encoding.UTF8);
        }

        public static Snapshot Load(string sheetName, string excelDir)
        {
            var path = Path.Combine(excelDir, SnapshotsDir, $"{sheetName}.json");
            if (!File.Exists(path)) return null;

            try
            {
                var json = File.ReadAllText(path);
                return DeserializeSnapshot(json);
            }
            catch (Exception ex)
            {
                Debug.LogWarning($"[ExcelToJSON] Failed to load schema snapshot for '{sheetName}': {ex.Message}");
                return null;
            }
        }

        private static string SerializeSnapshot(Snapshot snap)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine($"  \"excelFileName\": \"{snap.excelFileName}\",");
            sb.AppendLine($"  \"sheetName\": \"{snap.sheetName}\",");
            sb.AppendLine($"  \"version\": {snap.version},");
            sb.AppendLine("  \"columns\": [");
            for (int i = 0; i < snap.columns.Count; i++)
            {
                var c = snap.columns[i];
                var comma = i < snap.columns.Count - 1 ? "," : "";
                sb.AppendLine($"    {{\"name\": \"{c.name}\", \"type\": \"{c.type}\", \"index\": {c.index}}}{comma}");
            }
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private static Snapshot DeserializeSnapshot(string json)
        {
            var snap = new Snapshot();

            // Parse excelFileName
            var fnStart = json.IndexOf("\"excelFileName\"");
            if (fnStart >= 0)
            {
                var vStart = json.IndexOf('"', fnStart + 16) + 1;
                var vEnd = json.IndexOf('"', vStart);
                if (vStart > 0 && vEnd > vStart)
                    snap.excelFileName = json.Substring(vStart, vEnd - vStart);
            }

            // Parse sheetName
            var snStart = json.IndexOf("\"sheetName\"");
            if (snStart >= 0)
            {
                var vStart = json.IndexOf('"', snStart + 12) + 1;
                var vEnd = json.IndexOf('"', vStart);
                if (vStart > 0 && vEnd > vStart)
                    snap.sheetName = json.Substring(vStart, vEnd - vStart);
            }

            // Parse version
            var verStart = json.IndexOf("\"version\"");
            if (verStart >= 0)
            {
                var vStart = json.IndexOf(':', verStart) + 1;
                var vEnd = json.IndexOfAny(new[] { ',', '\n', '}' }, vStart);
                if (vStart > 0 && vEnd > vStart)
                    int.TryParse(json.Substring(vStart, vEnd - vStart).Trim(), out snap.version);
            }

            // Parse columns array
            var arrStart = json.IndexOf('[');
            var arrEnd = json.LastIndexOf(']');
            if (arrStart >= 0 && arrEnd > arrStart)
            {
                var arrJson = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
                var entries = SplitJsonObjects(arrJson);
                foreach (var entry in entries)
                {
                    var col = new Snapshot.ColumnInfo();
                    col.name = ExtractJsonString(entry, "name");
                    col.type = ExtractJsonString(entry, "type");
                    if (int.TryParse(ExtractJsonValue(entry, "index"), out var idx))
                        col.index = idx;
                    snap.columns.Add(col);
                }
            }

            return snap;
        }

        private static string ExtractJsonString(string json, string key)
        {
            var search = $"\"{key}\"";
            var idx = json.IndexOf(search);
            if (idx < 0) return "";
            var vStart = json.IndexOf('"', idx + search.Length) + 1;
            var vEnd = json.IndexOf('"', vStart);
            if (vStart > 0 && vEnd > vStart)
                return json.Substring(vStart, vEnd - vStart);
            return "";
        }

        private static string ExtractJsonValue(string json, string key)
        {
            var search = $"\"{key}\"";
            var idx = json.IndexOf(search);
            if (idx < 0) return "";
            var vStart = json.IndexOf(':', idx + search.Length) + 1;
            var vEnd = json.IndexOfAny(new[] { ',', '\n', '}', ' ' }, vStart);
            if (vStart > 0 && vEnd > vStart)
                return json.Substring(vStart, vEnd - vStart).Trim();
            return "";
        }

        private static List<string> SplitJsonObjects(string json)
        {
            var result = new List<string>();
            int depth = 0;
            var sb = new StringBuilder();
            foreach (char c in json)
            {
                if (c == '{') depth++;
                else if (c == '}') depth--;
                sb.Append(c);
                if (depth == 0 && sb.Length > 0)
                {
                    var trimmed = sb.ToString().Trim();
                    if (trimmed.StartsWith("{"))
                        result.Add(trimmed);
                    sb.Clear();
                }
            }
            return result;
        }
    }
}
