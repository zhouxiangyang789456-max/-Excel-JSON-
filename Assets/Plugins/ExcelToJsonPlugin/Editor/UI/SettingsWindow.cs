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

            EditorGUILayout.LabelField("Excel Parse Settings", style);
            EditorGUILayout.Space(5);

            headerRow = EditorGUILayout.IntField("Header Row (field names)", headerRow);
            typeRow = EditorGUILayout.IntField("Type Row", typeRow);
            commentRow = EditorGUILayout.IntField("Comment Row (0 = none)", commentRow);
            dataStartRow = EditorGUILayout.IntField("Data Start Row", dataStartRow);

            EditorGUILayout.Space(10);
            skipHiddenRows = EditorGUILayout.Toggle("Skip Hidden Rows", skipHiddenRows);
            skipHiddenCols = EditorGUILayout.Toggle("Skip Hidden Columns", skipHiddenCols);
            skipEmptyRows = EditorGUILayout.Toggle("Skip Empty Rows", skipEmptyRows);
            skipPrefixes = EditorGUILayout.TextField("Skip Sheet Prefixes (comma-separated)", skipPrefixes);

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField("Validation", style);
            strictMode = EditorGUILayout.Toggle("Strict Mode (warnings block export)", strictMode);

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField("Auto Export", style);
            enableAutoExport = EditorGUILayout.Toggle("Watch Excel files", enableAutoExport);
            if (enableAutoExport)
                debounceMs = EditorGUILayout.IntSlider("Debounce (ms)", debounceMs, 100, 3000);

            EditorGUILayout.Space(15);
            EditorGUILayout.LabelField("Code Generation", style);
            genPath = EditorGUILayout.TextField("Output Path (under Assets/)", genPath);
            genNamespace = EditorGUILayout.TextField("Namespace", genNamespace);

            EditorGUILayout.Space(20);

            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button("Save", GUILayout.Height(30)))
            {
                SaveSettings();
                EditorUtility.DisplayDialog("Settings", "Settings saved.", "OK");
            }
            if (GUILayout.Button("Reset Defaults", GUILayout.Height(30)))
                ResetDefaults();
            EditorGUILayout.EndHorizontal();
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
