using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    public class DashboardDrawer
    {
        private readonly ExcelDataWindow window;
        private Vector2 scroll;

        public DashboardDrawer(ExcelDataWindow window) { this.window = window; }

        public void Draw()
        {
            scroll = EditorGUILayout.BeginScrollView(scroll);

            EditorGUILayout.LabelField("Project Overview", EditorStyles.boldLabel);
            EditorGUILayout.Space(5);

            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Excel Directory:", GUILayout.Width(100));
            EditorGUILayout.LabelField(window.ExcelDir);
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Output Directory:", GUILayout.Width(100));
            EditorGUILayout.LabelField(window.OutputDir);
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField("Quick Status", EditorStyles.boldLabel);

            if (window.FileEntries.Count == 0)
            {
                EditorGUILayout.HelpBox(
                    "No Excel files found. Place .xlsx files in Assets/Excel/ or click '+ Add Excel Directory'.",
                    MessageType.Info);
            }
            else
            {
                // Summary table
                EditorGUILayout.BeginHorizontal(EditorStyles.toolbar);
                EditorGUILayout.LabelField("File", EditorStyles.miniBoldLabel, GUILayout.Width(180));
                EditorGUILayout.LabelField("Sheets", EditorStyles.miniBoldLabel, GUILayout.Width(60));
                EditorGUILayout.LabelField("Status", EditorStyles.miniBoldLabel);
                EditorGUILayout.EndHorizontal();

                foreach (var entry in window.FileEntries)
                {
                    EditorGUILayout.BeginHorizontal(EditorStyles.helpBox);

                    EditorGUILayout.LabelField(entry.FileName, GUILayout.Width(180));
                    EditorGUILayout.LabelField(entry.Sheets.Count.ToString(), GUILayout.Width(60));

                    // Check if output exists
                    bool hasOutput = false;
                    foreach (var sheet in entry.Sheets)
                    {
                        if (File.Exists($"Assets/Data/{sheet.Name}.asset") ||
                            File.Exists($"Assets/Data/{sheet.Name}.json"))
                        {
                            hasOutput = true;
                            break;
                        }
                    }

                    var status = hasOutput ? "✅ Exported" : "⚪ Not exported";
                    EditorGUILayout.LabelField(status);

                    // Quick actions
                    if (GUILayout.Button("Export", EditorStyles.miniButton, GUILayout.Width(50)))
                        window.RunExportSingle(entry.RelativePath, null);

                    if (GUILayout.Button("Open", EditorStyles.miniButton, GUILayout.Width(45)))
                        window.OpenExcel(entry.FullPath);

                    EditorGUILayout.EndHorizontal();
                }
            }

            EditorGUILayout.Space(15);

            // Quick actions
            EditorGUILayout.LabelField("Quick Actions", EditorStyles.boldLabel);
            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button("Export All", GUILayout.Height(35)))
                window.RunExportAll();
            if (GUILayout.Button("Validate All", GUILayout.Height(35)))
                window.RunValidate();
            if (GUILayout.Button("Refresh Files", GUILayout.Height(35)))
                window.RefreshFileList();
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(15);

            // Generated output
            EditorGUILayout.LabelField("Generated Output", EditorStyles.boldLabel);
            var dataDir = "Assets/Data";
            if (Directory.Exists(dataDir))
            {
                var files = Directory.GetFiles(dataDir)
                    .Where(f => f.EndsWith(".asset") || f.EndsWith(".json"))
                    .OrderBy(f => f).ToList();

                if (files.Count == 0)
                    EditorGUILayout.LabelField("(No generated files yet)", EditorStyles.miniLabel);
                else
                {
                    EditorGUILayout.LabelField($"  {files.Count} generated files:");
                    foreach (var f in files)
                        EditorGUILayout.LabelField($"    {Path.GetFileName(f)}");
                }
            }

            EditorGUILayout.EndScrollView();
        }
    }
}
