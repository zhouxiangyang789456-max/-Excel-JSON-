using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    public class SettingsWindow : EditorWindow
    {
        // Settings
        private int headerRow = 1;
        private int typeRow = 2;
        private int commentRow = 3;
        private int dataStartRow = 4;
        private bool skipHiddenRows = true;
        private bool skipHiddenCols = true;
        private bool skipEmptyRows = true;
        private string skipPrefixes = "_,#";
        private bool strictMode = false;
        private bool enableAutoExport = false;
        private int debounceMs = 500;
        private string genPath = "Scripts/Generated";
        private string genNamespace = "Game.Data";

        public static void ShowWindow()
        {
            var w = GetWindow<SettingsWindow>("Plugin Settings");
            w.minSize = new Vector2(400, 500);
            w.Show();
        }

        private void OnEnable()
        {
            LoadSettings();
        }

        private void OnGUI()
        {
            var style = new GUIStyle(EditorStyles.boldLabel) { fontSize = 13 };

            EditorGUILayout.LabelField(Loc.Tr("excel_parse_settings"), style);
            EditorGUILayout.Space(5);

            headerRow = EditorGUILayout.IntField(Loc.Tr("header_row"), headerRow);
            typeRow = EditorGUILayout.IntField(Loc.Tr("type_row"), typeRow);
            commentRow = EditorGUILayout.IntField(Loc.Tr("comment_row"), commentRow);
            dataStartRow = EditorGUILayout.IntField(Loc.Tr("data_start_row"), dataStartRow);

            EditorGUILayout.Space(10);
            skipHiddenRows = EditorGUILayout.Toggle(Loc.Tr("skip_hidden_rows"), skipHiddenRows);
            skipHiddenCols = EditorGUILayout.Toggle(Loc.Tr("skip_hidden_cols"), skipHiddenCols);
            skipEmptyRows = EditorGUILayout.Toggle(Loc.Tr("skip_empty_rows"), skipEmptyRows);
            skipPrefixes = EditorGUILayout.TextField(Loc.Tr("skip_prefixes"), skipPrefixes);

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField(Loc.Tr("validation_section"), style);
            strictMode = EditorGUILayout.Toggle(Loc.Tr("strict_mode"), strictMode);

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField(Loc.Tr("auto_export_section"), style);
            enableAutoExport = EditorGUILayout.Toggle(Loc.Tr("watch_files"), enableAutoExport);
            if (enableAutoExport)
                debounceMs = EditorGUILayout.IntSlider(Loc.Tr("debounce"), debounceMs, 100, 3000);

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField(Loc.Tr("code_gen_section"), style);
            genPath = EditorGUILayout.TextField(Loc.Tr("output_path_setting"), genPath);
            genNamespace = EditorGUILayout.TextField(Loc.Tr("namespace_setting"), genNamespace);

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField(Loc.Tr("language_section"), style);
            var currentLang = Core.Loc.GetLanguage();
            var langNames = new[] { Loc.Tr("lang_zh"), Loc.Tr("lang_en") };
            var langIdx = currentLang == "en" ? 1 : 0;
            var newLangIdx = EditorGUILayout.Popup(Loc.Tr("language_label"), langIdx, langNames);
            if (newLangIdx != langIdx)
            {
                Core.Loc.SetLanguage(newLangIdx == 1 ? "en" : "zh");
            }

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField(Loc.Tr("sheet_mode_title"), style);
            EditorGUILayout.HelpBox(Loc.Tr("sheet_mode_hint"), MessageType.Info);
            DrawSheetModeOverrides();

            EditorGUILayout.Space(20);

            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button(Loc.Tr("save_btn"), GUILayout.Height(30)))
            {
                SaveSettings();
                EditorUtility.DisplayDialog(Loc.Tr("settings_title"), Loc.Tr("settings_saved"), "OK");
            }
            if (GUILayout.Button(Loc.Tr("reset_defaults_btn"), GUILayout.Height(30)))
                ResetDefaults();
            EditorGUILayout.EndHorizontal();
        }

        private void DrawSheetModeOverrides()
        {
            var excelDir = EditorPrefs.GetString("ExcelToJson.ExcelDir", "Assets/Excel");
            if (!System.IO.Directory.Exists(excelDir))
            {
                EditorGUILayout.LabelField("(No Excel directory configured)", EditorStyles.miniLabel);
                return;
            }

            // Collect all sheet names from Excel files
            var allSheets = new System.Collections.Generic.HashSet<string>();
            var files = System.IO.Directory.GetFiles(excelDir, "*.xlsx",
                System.IO.SearchOption.AllDirectories)
                .Concat(System.IO.Directory.GetFiles(excelDir, "*.xls",
                    System.IO.SearchOption.AllDirectories))
                .Where(f => !System.IO.Path.GetFileName(f).StartsWith("~$"));

            foreach (var file in files)
            {
                try
                {
                    var reader = Core.ExcelReader.Read(file, false, false, false);
                    foreach (var sn in reader.SheetNames)
                    {
                        if (!sn.StartsWith("_") && !sn.StartsWith("#"))
                            allSheets.Add(sn);
                    }
                }
                catch { }
            }

            if (allSheets.Count == 0)
            {
                EditorGUILayout.LabelField("(No sheets found in Excel directory)", EditorStyles.miniLabel);
                return;
            }

            var scrollPos = EditorGUILayout.BeginScrollView(
                Vector2.zero, GUILayout.Height(120));

            foreach (var sheet in allSheets.OrderBy(s => s))
            {
                var prefKey = $"ExcelToJson.SheetMode.{sheet}";
                var currentMode = EditorPrefs.GetInt(prefKey, 0); // 0=Auto, 1=ModeA, 2=ModeB

                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField(sheet, GUILayout.Width(140));
                var newMode = EditorGUILayout.Popup(currentMode,
                    new[] { "Auto", "Mode A", "Mode B" }, GUILayout.Width(90));
                EditorGUILayout.EndHorizontal();

                if (newMode != currentMode)
                    EditorPrefs.SetInt(prefKey, newMode);
            }

            EditorGUILayout.EndScrollView();
        }

        private void LoadSettings()
        {
            headerRow = EditorPrefs.GetInt("ExcelToJson.HeaderRow", 1);
            typeRow = EditorPrefs.GetInt("ExcelToJson.TypeRow", 2);
            commentRow = EditorPrefs.GetInt("ExcelToJson.CommentRow", 3);
            dataStartRow = EditorPrefs.GetInt("ExcelToJson.DataStartRow", 4);
            skipHiddenRows = EditorPrefs.GetBool("ExcelToJson.SkipHiddenRows", true);
            skipHiddenCols = EditorPrefs.GetBool("ExcelToJson.SkipHiddenCols", true);
            skipEmptyRows = EditorPrefs.GetBool("ExcelToJson.SkipEmptyRows", true);
            skipPrefixes = EditorPrefs.GetString("ExcelToJson.SkipPrefixes", "_,#");
            strictMode = EditorPrefs.GetBool("ExcelToJson.StrictMode", false);
            enableAutoExport = EditorPrefs.GetBool("ExcelToJson.AutoExport", false);
            debounceMs = EditorPrefs.GetInt("ExcelToJson.DebounceMs", 500);
            genPath = EditorPrefs.GetString("ExcelToJson.GenPath", "Scripts/Generated");
            genNamespace = EditorPrefs.GetString("ExcelToJson.Namespace", "Game.Data");
        }

        private void SaveSettings()
        {
            EditorPrefs.SetInt("ExcelToJson.HeaderRow", headerRow);
            EditorPrefs.SetInt("ExcelToJson.TypeRow", typeRow);
            EditorPrefs.SetInt("ExcelToJson.CommentRow", commentRow);
            EditorPrefs.SetInt("ExcelToJson.DataStartRow", dataStartRow);
            EditorPrefs.SetBool("ExcelToJson.SkipHiddenRows", skipHiddenRows);
            EditorPrefs.SetBool("ExcelToJson.SkipHiddenCols", skipHiddenCols);
            EditorPrefs.SetBool("ExcelToJson.SkipEmptyRows", skipEmptyRows);
            EditorPrefs.SetString("ExcelToJson.SkipPrefixes", skipPrefixes);
            EditorPrefs.SetBool("ExcelToJson.StrictMode", strictMode);
            EditorPrefs.SetBool("ExcelToJson.AutoExport", enableAutoExport);
            EditorPrefs.SetInt("ExcelToJson.DebounceMs", debounceMs);
            EditorPrefs.SetString("ExcelToJson.GenPath", genPath);
            EditorPrefs.SetString("ExcelToJson.Namespace", genNamespace);
        }

        private void ResetDefaults()
        {
            headerRow = 1; typeRow = 2; commentRow = 3; dataStartRow = 4;
            skipHiddenRows = true; skipHiddenCols = true; skipEmptyRows = true;
            skipPrefixes = "_,#"; strictMode = false; enableAutoExport = false;
            debounceMs = 500; genPath = "Scripts/Generated"; genNamespace = "Game.Data";
            Repaint();
        }
    }
}
