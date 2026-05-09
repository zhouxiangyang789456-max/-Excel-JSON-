using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// MD5-based incremental cache for Excel sheets.
    /// Only processes sheets whose data has actually changed since last export.
    /// Cache stored as JSON at Assets/Excel/.cache/hashes.json
    /// </summary>
    public static class IncrementalCache
    {
        private const string CacheFileName = "hashes.json";

        /// <summary>
        /// Compute MD5 hash of a sheet's string data rows (excluding header rows).
        /// </summary>
        public static string ComputeHash(List<List<string>> rows, int dataStartRow)
        {
            using (var md5 = MD5.Create())
            {
                var sb = new StringBuilder();
                int start = dataStartRow - 1;

                for (int ri = start; ri < rows.Count; ri++)
                {
                    var row = rows[ri];
                    // Skip empty rows
                    bool allEmpty = true;
                    foreach (var cell in row)
                    {
                        if (!string.IsNullOrEmpty(cell))
                        {
                            allEmpty = false;
                            break;
                        }
                    }
                    if (allEmpty) continue;

                    sb.AppendJoin("|", row);
                    sb.Append('\n');
                }

                var bytes = Encoding.UTF8.GetBytes(sb.ToString());
                var hash = md5.ComputeHash(bytes);
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Build a composite key for a file+sheet pair.
        /// </summary>
        public static string MakeKey(string fileName, string sheetName)
        {
            return $"{fileName}###{sheetName}";
        }

        /// <summary>
        /// Load cached hashes from disk.
        /// Returns empty dict if no cache exists.
        /// </summary>
        public static Dictionary<string, string> LoadCache(string excelDir)
        {
            var cache = new Dictionary<string, string>();
            var cacheFile = Path.Combine(excelDir, ".cache", CacheFileName);

            if (!File.Exists(cacheFile))
                return cache;

            try
            {
                var json = File.ReadAllText(cacheFile);
                // Simple JSON parse without external dependency
                // Expected format: {"key1": "hash1", "key2": "hash2"}
                json = json.Trim();
                if (json.StartsWith("{") && json.EndsWith("}"))
                {
                    json = json.Substring(1, json.Length - 2);
                    var entries = SplitJsonEntries(json);
                    foreach (var entry in entries)
                    {
                        var parts = entry.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            var key = parts[0].Trim().Trim('"');
                            var val = parts[1].Trim().Trim('"');
                            if (!string.IsNullOrEmpty(key))
                                cache[key] = val;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.LogWarning($"[ExcelToJSON] Failed to load hash cache: {ex.Message}");
            }

            return cache;
        }

        /// <summary>
        /// Save hashes to disk.
        /// </summary>
        public static void SaveCache(string excelDir, Dictionary<string, string> cache)
        {
            var cacheDir = Path.Combine(excelDir, ".cache");
            Directory.CreateDirectory(cacheDir);

            var cacheFile = Path.Combine(cacheDir, CacheFileName);

            var sb = new StringBuilder();
            sb.AppendLine("{");
            int count = 0;
            foreach (var kv in cache)
            {
                count++;
                var comma = count < cache.Count ? "," : "";
                sb.AppendLine($"  \"{kv.Key}\": \"{kv.Value}\"{comma}");
            }
            sb.AppendLine("}");
            File.WriteAllText(cacheFile, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Check if a sheet has changed since last export.
        /// Returns true if the sheet should be re-exported.
        /// </summary>
        public static bool HasChanged(
            string excelDir, string fileName, string sheetName,
            List<List<string>> rows, int dataStartRow,
            Dictionary<string, string> cache)
        {
            var key = MakeKey(fileName, sheetName);
            var hash = ComputeHash(rows, dataStartRow);

            if (cache.TryGetValue(key, out var cachedHash))
            {
                return cachedHash != hash;
            }

            // No cache entry — treat as changed
            return true;
        }

        /// <summary>
        /// Update cache entry for a sheet after successful export.
        /// </summary>
        public static void UpdateCache(
            string fileName, string sheetName,
            List<List<string>> rows, int dataStartRow,
            Dictionary<string, string> cache,
            string excelDir)
        {
            var key = MakeKey(fileName, sheetName);
            var hash = ComputeHash(rows, dataStartRow);
            cache[key] = hash;
            SaveCache(excelDir, cache);
        }

        /// <summary>
        /// Reference graph cache file path.
        /// </summary>
        private const string RefGraphFileName = "ref_graph.json";

        /// <summary>
        /// Save the table reference graph (table → set of tables it references).
        /// Used for ref chain invalidation in incremental export.
        /// </summary>
        public static void SaveRefGraph(string excelDir,
            Dictionary<string, HashSet<string>> refGraph)
        {
            var cacheDir = Path.Combine(excelDir, ".cache");
            Directory.CreateDirectory(cacheDir);
            var path = Path.Combine(cacheDir, RefGraphFileName);

            var sb = new StringBuilder();
            sb.AppendLine("{");
            int count = 0;
            foreach (var kv in refGraph)
            {
                count++;
                var comma = count < refGraph.Count ? "," : "";
                var refs = string.Join("\",\"", kv.Value);
                sb.AppendLine($"  \"{kv.Key}\": [\"{refs}\"]{comma}");
            }
            sb.AppendLine("}");
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Given a changed sheet name, returns all sheets that reference it
        /// (and should therefore be re-exported).
        /// </summary>
        public static HashSet<string> GetDependentSheets(
            string excelDir, string changedSheet)
        {
            var dependents = new HashSet<string>();
            var refGraph = LoadRefGraph(excelDir);

            // Simple reverse lookup
            foreach (var kv in refGraph)
            {
                if (kv.Value.Contains(changedSheet))
                    dependents.Add(kv.Key);
            }

            // Recursively find transitive dependents
            var visited = new HashSet<string>();
            var queue = new Queue<string>(dependents);
            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                if (!visited.Add(current)) continue;
                foreach (var kv in refGraph)
                {
                    if (kv.Value.Contains(current) && visited.Add(kv.Key))
                        queue.Enqueue(kv.Key);
                }
            }

            return visited;
        }

        private static Dictionary<string, HashSet<string>> LoadRefGraph(
            string excelDir)
        {
            var graph = new Dictionary<string, HashSet<string>>();
            var path = Path.Combine(excelDir, ".cache", RefGraphFileName);
            if (!File.Exists(path)) return graph;

            try
            {
                var json = File.ReadAllText(path).Trim();
                if (json.StartsWith("{") && json.EndsWith("}"))
                {
                    json = json.Substring(1, json.Length - 2);
                    var entries = SplitJsonEntries(json);
                    foreach (var entry in entries)
                    {
                        // Format: "table": ["ref1","ref2"]
                        var colonIdx = entry.IndexOf(':');
                        if (colonIdx < 0) continue;
                        var key = entry.Substring(0, colonIdx).Trim().Trim('"');
                        var arrPart = entry.Substring(colonIdx + 1).Trim();
                        arrPart = arrPart.Trim('[', ']');
                        var refs = new HashSet<string>();
                        foreach (var r in arrPart.Split(','))
                        {
                            var trimmed = r.Trim().Trim('"');
                            if (!string.IsNullOrEmpty(trimmed))
                                refs.Add(trimmed);
                        }
                        graph[key] = refs;
                    }
                }
            }
            catch { }

            return graph;
        }

        /// <summary>
        /// Simple JSON entry splitting (handles nested quotes).
        /// </summary>
        private static List<string> SplitJsonEntries(string json)
        {
            var entries = new List<string>();
            int depth = 0;
            bool inString = false;
            var sb = new StringBuilder();

            foreach (char c in json)
            {
                if (c == '"' && (sb.Length == 0 || sb[sb.Length - 1] != '\\'))
                    inString = !inString;

                if (!inString)
                {
                    if (c == '{' || c == '[') depth++;
                    else if (c == '}' || c == ']') depth--;
                    else if (c == ',' && depth == 0)
                    {
                        entries.Add(sb.ToString().Trim());
                        sb.Clear();
                        continue;
                    }
                }

                sb.Append(c);
            }

            var remaining = sb.ToString().Trim();
            if (!string.IsNullOrEmpty(remaining))
                entries.Add(remaining);

            return entries;
        }
    }
}
