using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
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

            EditorGUILayout.LabelField(Loc.Tr("dashboard_title"), EditorStyles.boldLabel);
            EditorGUILayout.Space(5);

            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField(Loc.Tr("excel_dir_label"), GUILayout.Width(100));
            EditorGUILayout.LabelField(window.ExcelDir);
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField(Loc.Tr("output_dir_label"), GUILayout.Width(100));
            EditorGUILayout.LabelField(window.OutputDir);
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField(Loc.Tr("quick_status"), EditorStyles.boldLabel);

            if (window.FileEntries.Count == 0)
            {
                EditorGUILayout.HelpBox(Loc.Tr("no_excel_files"), MessageType.Info);
            }
            else
            {
                // Summary table
                EditorGUILayout.BeginHorizontal(EditorStyles.toolbar);
                EditorGUILayout.LabelField(Loc.Tr("file_col"), EditorStyles.miniBoldLabel, GUILayout.Width(180));
                EditorGUILayout.LabelField(Loc.Tr("sheets_col"), EditorStyles.miniBoldLabel, GUILayout.Width(60));
                EditorGUILayout.LabelField(Loc.Tr("status_col"), EditorStyles.miniBoldLabel);
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

                    var status = hasOutput ? Loc.Tr("status_exported") : Loc.Tr("status_not_exported");
                    EditorGUILayout.LabelField(status);

                    // Quick actions
                    if (GUILayout.Button(Loc.Tr("export_all"), EditorStyles.miniButton, GUILayout.Width(50)))
                        window.RunExportSingle(entry.RelativePath, null);

                    if (GUILayout.Button(Loc.Tr("open_excel"), EditorStyles.miniButton, GUILayout.Width(45)))
                        window.OpenExcel(entry.FullPath);

                    EditorGUILayout.EndHorizontal();
                }
            }

            EditorGUILayout.Space(15);

            // Quick actions
            EditorGUILayout.LabelField(Loc.Tr("quick_actions"), EditorStyles.boldLabel);
            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button(Loc.Tr("export_all_btn"), GUILayout.Height(35)))
                window.RunExportAll();
            if (GUILayout.Button(Loc.Tr("validate_all_btn"), GUILayout.Height(35)))
                window.RunValidate();
            if (GUILayout.Button(Loc.Tr("refresh_files_btn"), GUILayout.Height(35)))
                window.RefreshFileList();
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(15);

            // Generated output
            EditorGUILayout.LabelField(Loc.Tr("generated_output"), EditorStyles.boldLabel);
            var dataDir = "Assets/Data";
            if (Directory.Exists(dataDir))
            {
                var files = Directory.GetFiles(dataDir)
                    .Where(f => f.EndsWith(".asset") || f.EndsWith(".json"))
                    .OrderBy(f => f).ToList();

                if (files.Count == 0)
                    EditorGUILayout.LabelField(Loc.Tr("no_generated_files"), EditorStyles.miniLabel);
                else
                {
                    EditorGUILayout.LabelField($"  {files.Count} {Loc.Tr("gen_files_title")}:");
                    foreach (var f in files)
                        EditorGUILayout.LabelField($"    {Path.GetFileName(f)}");
                }
            }

            EditorGUILayout.EndScrollView();
        }
    }
}
