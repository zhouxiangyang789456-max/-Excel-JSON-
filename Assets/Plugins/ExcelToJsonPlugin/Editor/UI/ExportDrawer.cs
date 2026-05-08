using System.Collections.Generic;
using System.Linq;
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
            EditorGUILayout.LabelField("Export Configuration", EditorStyles.boldLabel);
            EditorGUILayout.Space(5);

            // Output format
            EditorGUILayout.LabelField("Output Format:", EditorStyles.label);
            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Toggle(true, "ScriptableObject (.asset)", EditorStyles.radioButton)) { }
            if (GUILayout.Toggle(false, "JSON", EditorStyles.radioButton)) { }
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(5);

            // Output path
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Output Path:", GUILayout.Width(90));
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
            autoValidate = EditorGUILayout.ToggleLeft("Auto-validate before export", autoValidate);
            blockOnError = EditorGUILayout.ToggleLeft("Block export on validation error", blockOnError);
            exportJson = EditorGUILayout.ToggleLeft("Also export JSON (for hot-update)", exportJson);

            EditorGUILayout.Space(15);

            // Action buttons
            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button("Export All", GUILayout.Height(40)))
                window.RunExportAll();

            if (!string.IsNullOrEmpty(selectedFile) && GUILayout.Button("Export Selected", GUILayout.Height(40)))
                window.RunExportSingle(selectedFile, selectedSheet);
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.Space(10);

            // Last export result
            if (!string.IsNullOrEmpty(lastExportSummary))
            {
                EditorGUILayout.HelpBox(lastExportSummary,
                    lastExportSuccess ? MessageType.Info : MessageType.Error);
            }

            // Generated files
            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField("Generated Files", EditorStyles.boldLabel);
            ShowGeneratedFiles();
        }

        private void ShowGeneratedFiles()
        {
            var dataDir = "Assets/Data";
            if (!System.IO.Directory.Exists(dataDir))
            {
                EditorGUILayout.LabelField("(No generated files yet)", EditorStyles.miniLabel);
                return;
            }

            var files = System.IO.Directory.GetFiles(dataDir, "*.asset")
                .Concat(System.IO.Directory.GetFiles(dataDir, "*.json"));

            foreach (var f in files)
            {
                var fileName = System.IO.Path.GetFileName(f);
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField($"  {(f.EndsWith(".asset") ? "📦" : "📄")} {fileName}");

                if (GUILayout.Button("Select", GUILayout.Width(50)))
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
