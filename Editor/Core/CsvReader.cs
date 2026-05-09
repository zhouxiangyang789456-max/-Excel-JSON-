using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// Simple CSV reader as a lightweight alternative to Excel.
    /// CSV files can be git-diffed, merged, and edited in any text editor.
    ///
    /// Format:
    ///   First row: field names
    ///   Second row: type declarations
    ///   Third row: comments (optional)
    ///   Subsequent rows: data
    ///
    /// Columns separated by comma or tab.
    /// Values may be quoted with double-quotes.
    /// </summary>
    public static class CsvReader
    {
        /// <summary>
        /// Read a CSV file and return structured rows (same format as ExcelReader).
        /// </summary>
        public static List<List<string>> Read(string filePath)
        {
            var rows = new List<List<string>>();

            if (!File.Exists(filePath))
            {
                Debug.LogWarning($"[ExcelToJSON] CSV file not found: {filePath}");
                return rows;
            }

            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                if (line.TrimStart().StartsWith("#")) continue; // Skip comment lines

                var cells = ParseCsvLine(line);
                if (cells.Count > 0)
                    rows.Add(cells);
            }

            Debug.Log($"[ExcelToJSON] CSV read: {Path.GetFileName(filePath)} ({rows.Count} rows)");
            return rows;
        }

        /// <summary>
        /// Parse a single CSV line, handling quoted fields.
        /// </summary>
        private static List<string> ParseCsvLine(string line)
        {
            var cells = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            char delimiter = line.Contains('\t') ? '\t' : ',';

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == delimiter && !inQuotes)
                {
                    cells.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
            cells.Add(sb.ToString().Trim()); // Last cell

            return cells;
        }

        /// <summary>
        /// Check if a file is a CSV file.
        /// </summary>
        public static bool IsCsvFile(string filePath)
        {
            var ext = Path.GetExtension(filePath)?.ToLower();
            return ext == ".csv" || ext == ".tsv";
        }
    }
}
