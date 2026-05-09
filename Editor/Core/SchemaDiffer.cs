using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core.Models;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// Detects schema changes between current Excel structure and saved snapshot.
    /// Outputs a diff report with safe/unsafe change classification.
    /// </summary>
    public static class SchemaDiffer
    {
        public enum ChangeType
        {
            Added,      // New column added — safe, auto-migrate
            Removed,    // Column removed — dangerous, requires confirmation
            Renamed,    // Column renamed — dangerous, requires confirmation
            Retyped,    // Column type changed — dangerous, requires confirmation
            Reordered,  // Column order changed — safe, auto-migrate
        }

        public class Change
        {
            public ChangeType Type;
            public string ColumnName;
            public string OldValue;
            public string NewValue;
            public int Index;

            public bool IsDangerous => Type == ChangeType.Removed
                || Type == ChangeType.Renamed
                || Type == ChangeType.Retyped;

            public override string ToString()
            {
                return Type switch
                {
                    ChangeType.Added => $"+ 新增列: \"{ColumnName}\" ({NewValue})",
                    ChangeType.Removed => $"- 删除列: \"{ColumnName}\" (旧类型: {OldValue})",
                    ChangeType.Renamed => $"~ 重命名: \"{OldValue}\" → \"{ColumnName}\"",
                    ChangeType.Retyped => $"~ 类型变更: \"{ColumnName}\" ({OldValue} → {NewValue})",
                    ChangeType.Reordered => $"≈ 列顺序变更: \"{ColumnName}\"",
                    _ => ColumnName,
                };
            }
        }

        /// <summary>
        /// Compare the current schema with a saved snapshot.
        /// Returns list of detected changes.
        /// </summary>
        public static List<Change> Diff(TableSchema schema, SchemaSnapshot.Snapshot snapshot)
        {
            var changes = new List<Change>();

            var currentCols = new Dictionary<string, FieldDef>();
            var oldCols = new Dictionary<string, SchemaSnapshot.Snapshot.ColumnInfo>();

            foreach (var f in schema.Fields)
                currentCols[f.Name] = f;

            foreach (var c in snapshot.columns)
                oldCols[c.name] = c;

            // Check for removed/renamed/retyped columns
            foreach (var old in snapshot.columns)
            {
                // See if this column index now has a different name → renamed
                var currentByIndex = currentCols.Values.FirstOrDefault(
                    f => f.ColumnIndex == old.index);

                if (currentByIndex != null && currentByIndex.Name != old.name)
                {
                    // Index match, different name → potential rename
                    changes.Add(new Change
                    {
                        Type = ChangeType.Renamed,
                        ColumnName = currentByIndex.Name,
                        OldValue = old.name,
                        NewValue = currentByIndex.Name,
                        Index = old.index,
                    });
                }
                else if (!currentCols.ContainsKey(old.name))
                {
                    // Old column no longer exists by name or index → removed
                    changes.Add(new Change
                    {
                        Type = ChangeType.Removed,
                        ColumnName = old.name,
                        OldValue = old.type,
                        Index = old.index,
                    });
                }
                else if (currentCols.TryGetValue(old.name, out var currentField))
                {
                    // Same name, different type → retyped
                    var currentType = currentField.RawType ?? currentField.NormalizedType ?? "";
                    if (currentType != old.type && !string.IsNullOrEmpty(old.type))
                    {
                        changes.Add(new Change
                        {
                            Type = ChangeType.Retyped,
                            ColumnName = old.name,
                            OldValue = old.type,
                            NewValue = currentType,
                            Index = old.index,
                        });
                    }
                    // Same name, different index → reordered
                    else if (currentField.ColumnIndex != old.index)
                    {
                        changes.Add(new Change
                        {
                            Type = ChangeType.Reordered,
                            ColumnName = old.name,
                            Index = old.index,
                        });
                    }
                }
            }

            // Check for added columns (in current but not in old)
            foreach (var field in schema.Fields)
            {
                if (!oldCols.ContainsKey(field.Name))
                {
                    // Not renamed either (already handled)
                    bool isRenamed = changes.Any(c =>
                        c.Type == ChangeType.Renamed && c.ColumnName == field.Name);
                    if (!isRenamed)
                    {
                        changes.Add(new Change
                        {
                            Type = ChangeType.Added,
                            ColumnName = field.Name,
                            NewValue = field.RawType ?? field.NormalizedType ?? "",
                            Index = field.ColumnIndex,
                        });
                    }
                }
            }

            return changes;
        }

        /// <summary>
        /// Check if any changes are dangerous (require confirmation).
        /// </summary>
        public static bool HasDangerousChanges(List<Change> changes)
        {
            return changes.Any(c => c.IsDangerous);
        }

        /// <summary>
        /// Build a human-readable summary of changes.
        /// </summary>
        public static string BuildSummary(List<Change> changes, string sheetName)
        {
            if (changes.Count == 0) return null;

            var safeCount = changes.Count(c => !c.IsDangerous);
            var dangerCount = changes.Count(c => c.IsDangerous);

            var lines = new List<string>
            {
                $"Schema 变更检测: {sheetName}",
                $"共 {changes.Count} 处变更 ({safeCount} 安全, {dangerCount} 危险)",
                "",
            };
            foreach (var c in changes)
                lines.Add($"  {c}");

            return string.Join("\n", lines);
        }
    }
}
