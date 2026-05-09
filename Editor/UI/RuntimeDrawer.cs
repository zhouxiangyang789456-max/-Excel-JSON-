using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToJsonPlugin.Editor.Core;
using UnityEditor;
using UnityEngine;

namespace ExcelToJsonPlugin.Editor.UI
{
    public class RuntimeDrawer
    {
        private readonly ExcelDataWindow window;

        public RuntimeDrawer(ExcelDataWindow window) { this.window = window; }

        public void Draw(string selectedFile, string selectedSheet, List<ExcelFileEntry> entries)
        {
            EditorGUILayout.LabelField(Loc.Tr("runtime_title"), EditorStyles.boldLabel);
            EditorGUILayout.Space(5);

            // Quick start
            EditorGUILayout.LabelField(Loc.Tr("quick_start"), EditorStyles.boldLabel);
            EditorGUILayout.HelpBox(Loc.Tr("quick_start_text"), MessageType.Info);

            if (GUILayout.Button(Loc.Tr("gen_datamanager_btn"), GUILayout.Height(35)))
                GenerateDataManagerGO();

            EditorGUILayout.Space(15);

            // API reference
            EditorGUILayout.LabelField(Loc.Tr("api_ref"), EditorStyles.boldLabel);
            DrawApiBox(
                "GetTable<T>()",
                "var t = DataManager.Instance.GetTable<WeaponTable>();\nvar row = t.Get(1001);",
                "Get a typed table reference, then query by ID (O(1)).");

            DrawApiBox(
                "t.GetAll() / t.Find()",
                "foreach (var w in t.GetAll()) { ... }\nvar filtered = t.Find(w => w.atk > 30);",
                "Get all rows or filter by condition.");

            DrawApiBox(
                "t.HasId() / t.GetRandom() / t.GetByIds()",
                "if (t.HasId(id)) { ... }\nvar random = t.GetRandom(w => w.rare >= 3);\nvar batch = t.GetByIds(new[]{1,2,3});",
                "Check existence, random pick, batch query.");

            DrawApiBox(
                "BuildAllCachesAsync()",
                "yield return DataManager.Instance.BuildAllCachesAsync(5, (c,t) => progress = c/t);",
                "Async cache build with progress callback (no main thread freeze).");

            EditorGUILayout.Space(15);

            // Generated files
            EditorGUILayout.LabelField("Generated Files", EditorStyles.boldLabel);
            ShowGeneratedFiles("Assets/Scripts/Generated/Data", "Row classes");
            ShowGeneratedFiles("Assets/Scripts/Generated/Tables", "Table ScriptableObject classes");
            ShowGeneratedFiles("Assets/Data", "*.asset / *.json data files");
        }

        private void DrawApiBox(string title, string code, string desc)
        {
            EditorGUILayout.BeginVertical(EditorStyles.helpBox);
            EditorGUILayout.LabelField(title, EditorStyles.boldLabel);
            EditorGUILayout.LabelField(desc, EditorStyles.wordWrappedMiniLabel);
            EditorGUILayout.TextArea(code, GUILayout.Height(50));
            if (GUILayout.Button("Copy", GUILayout.Width(50)))
            {
                GUIUtility.systemCopyBuffer = code;
                Debug.Log($"Copied: {title}");
            }
            EditorGUILayout.EndVertical();
            EditorGUILayout.Space(5);
        }

        private void ShowGeneratedFiles(string dir, string label)
        {
            EditorGUILayout.LabelField($"  {label}:", EditorStyles.miniBoldLabel);
            if (!Directory.Exists(dir))
            {
                EditorGUILayout.LabelField("    (not yet generated)", EditorStyles.miniLabel);
                return;
            }

            foreach (var f in Directory.GetFiles(dir, "*.*")
                .Where(f => f.EndsWith(".cs") || f.EndsWith(".asset") || f.EndsWith(".json")))
            {
                var fn = Path.GetFileName(f);
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField($"    {fn}", EditorStyles.miniLabel);
                if (GUILayout.Button("Open", EditorStyles.miniButton, GUILayout.Width(45)))
                {
                    var asset = AssetDatabase.LoadAssetAtPath<Object>(f);
                    if (asset != null) AssetDatabase.OpenAsset(asset);
                }
                EditorGUILayout.EndHorizontal();
            }
        }

        private void GenerateDataManagerGO()
        {
            // Find or create DataManager in scene
            var existing = Object.FindObjectOfType<Runtime.DataManager>();
            if (existing != null)
            {
                Selection.activeGameObject = existing.gameObject;
                EditorGUIUtility.PingObject(existing);
                Debug.Log("DataManager already exists in scene.");
                return;
            }

            var go = new GameObject("DataManager");
            var dm = go.AddComponent<Runtime.DataManager>();

            // Auto-populate allTables
            var dataDir = "Assets/Data";
            var tables = new List<Runtime.BaseDataTable>();
            if (Directory.Exists(dataDir))
            {
                foreach (var f in Directory.GetFiles(dataDir, "*.asset"))
                {
                    var asset = AssetDatabase.LoadAssetAtPath<Runtime.BaseDataTable>(f);
                    if (asset != null) tables.Add(asset);
                }
            }
            dm.SetAllTables(tables);

            Selection.activeGameObject = go;
            Debug.Log($"DataManager GameObject created with {tables.Count} tables.");
        }
    }
}
