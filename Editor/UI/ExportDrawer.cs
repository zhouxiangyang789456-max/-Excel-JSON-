using System.Collections.Generic;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using ExcelToJsonPlugin.Editor.Mapping;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    /// <summary>Tab 3: Export controls and progress</summary>
    public class ExportDrawer
    {
        private readonly ExcelDataWindow window;
        private bool exportJson = true;
        private bool autoValidate = true;
        private bool blockOnError = false;
        private string lastExportSummary = "";
        private bool lastExportSuccess;

        public ExportDrawer(ExcelDataWindow window) { this.window = window; }

        public void Draw(string selectedFile, string selectedSheet, List<ExcelFileEntry> entries)
        {
            EditorGUILayout.LabelField(Loc.Tr("export_config"), EditorStyles.boldLabel);
            EditorGUILayout.Space(5);

            // Output format
            EditorGUILayout.LabelField(Loc.Tr("output_format"), EditorStyles.label);
            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Toggle(true, Loc.Tr("format_so"), EditorStyles.radioButton)) { }
            if (GUILayout.Toggle(false, Loc.Tr("format_json"), EditorStyles.radioButton)) { }
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(5);

            // Output path
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField(Loc.Tr("output_path"), GUILayout.Width(90));
            window.OutputDir = EditorGUILayout.TextField(window.OutputDir);
            if (GUILayout.Button("...", GUILayout.Width(30)))
            {
                var newDir = EditorUtility.OpenFolderPanel("Select Output Directory", "Assets", "");
                if (!string.IsNullOrEmpty(newDir))
                    window.OutputDir = newDir;
            }
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(10);

            // Options
            autoValidate = EditorGUILayout.ToggleLeft(Loc.Tr("auto_validate"), autoValidate);
            blockOnError = EditorGUILayout.ToggleLeft(Loc.Tr("block_on_error"), blockOnError);
            exportJson = EditorGUILayout.ToggleLeft(Loc.Tr("export_json_too"), exportJson);

            EditorGUILayout.Space(15);

            // Action buttons
            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button(Loc.Tr("export_all_btn_lg"), GUILayout.Height(40)))
                window.RunExportAll();

            if (!string.IsNullOrEmpty(selectedFile) && GUILayout.Button(Loc.Tr("export_selected_btn"), GUILayout.Height(40)))
                window.RunExportSingle(selectedFile, selectedSheet);
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(5);

            // Mode B: Export Excel template from C# classes
            if (GUILayout.Button(Loc.Tr("export_templates_btn"), GUILayout.Height(30)))
            {
                var paths = Mapping.TemplateExporter.GenerateAllTemplates("Excel");
                var count = paths?.Count ?? 0;
                if (count > 0)
                    EditorUtility.DisplayDialog("Templates", Loc.Tr("template_done", count), "OK");
                else
                    EditorUtility.DisplayDialog("Templates", Loc.Tr("no_template_classes"), "OK");
            }

            EditorGUILayout.Space(10);

            // Last export result
            if (!string.IsNullOrEmpty(lastExportSummary))
            {
                EditorGUILayout.HelpBox(lastExportSummary,
                    lastExportSuccess ? MessageType.Info : MessageType.Error);
            }

            // Generated files
            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField(Loc.Tr("gen_files_title"), EditorStyles.boldLabel);
            ShowGeneratedFiles();
        }

        private void ShowGeneratedFiles()
        {
            var dataDir = "Assets/Data";
            if (!System.IO.Directory.Exists(dataDir))
            {
                EditorGUILayout.LabelField(Loc.Tr("no_gen_files"), EditorStyles.miniLabel);
                return;
            }

            var files = System.IO.Directory.GetFiles(dataDir, "*.asset")
                .Concat(System.IO.Directory.GetFiles(dataDir, "*.json"));

            foreach (var f in files)
            {
                var fileName = System.IO.Path.GetFileName(f);
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField($"  {(f.EndsWith(".asset") ? "📦" : "📄")} {fileName}");

                if (GUILayout.Button(Loc.Tr("select_btn"), GUILayout.Width(50)))
                {
                    var asset = AssetDatabase.LoadAssetAtPath<Object>(f);
                    if (asset != null)
                        Selection.activeObject = asset;
                }
                EditorGUILayout.EndHorizontal();
            }
        }
    }
}
