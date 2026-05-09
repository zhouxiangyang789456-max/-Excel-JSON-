using System.Collections;
using UnityEngine;
using ExcelToJsonPlugin.Runtime;

namespace ExcelToJsonPlugin.Demo
{
    /// <summary>
    /// Demo script showing how to use the DataManager API in game code.
    /// Attach this to a GameObject, press Play, and click the UI buttons.
    ///
    /// Prerequisites:
    ///   1. Export Excel data via Window > Excel Data Manager
    ///   2. Click "Auto-Generate DataManager GameObject" in Runtime API tab
    ///   3. Drag this script onto the DataManager GameObject (or any GameObject)
    /// </summary>
    public class DemoDataUsage : MonoBehaviour
    {
        [Header("UI")]
        public bool showDemoGUI = true;

        private DataManager dataManager;
        private string lastQueryResult = "";
        private int queryId = 1001;

        private void Start()
        {
            dataManager = DataManager.Instance;
            if (dataManager == null)
            {
                Debug.LogError("[Demo] DataManager not found! " +
                    "Use Runtime API tab → Auto-Generate DataManager GameObject.");
                return;
            }

            // Build all table caches
            dataManager.BuildAllCaches((done, total) =>
            {
                Debug.Log($"[Demo] Building cache: {done}/{total}");
            });

            Debug.Log($"[Demo] DataManager ready with {dataManager.TableCount} tables.");

            // Example: print all table names
            var tables = dataManager.GetAllTableAssets();
            foreach (var t in tables)
            {
                if (t != null)
                    Debug.Log($"[Demo] Table: {t.TableName} ({t.Count} rows)");
            }
        }

        private void OnGUI()
        {
            if (!showDemoGUI) return;
            if (dataManager == null)
            {
                GUI.Label(new Rect(10, 10, 400, 30), "DataManager not found. See Console.");
                return;
            }

            GUILayout.BeginArea(new Rect(10, 10, 400, 500));
            GUILayout.BeginVertical("box");

            GUILayout.Label("=== ExcelToJSON Demo ===", GUI.skin.box);
            GUILayout.Space(10);

            GUILayout.Label($"Tables loaded: {dataManager.TableCount}");
            foreach (var t in dataManager.GetAllTableAssets())
            {
                if (t != null)
                    GUILayout.Label($"  - {t.TableName}: {t.Count} rows");
            }

            GUILayout.Space(15);
            GUILayout.Label("Query by ID:");
            GUILayout.BeginHorizontal();
            queryId = int.Parse(GUILayout.TextField(queryId.ToString(), GUILayout.Width(80)));
            if (GUILayout.Button("Query"))
            {
                lastQueryResult = QueryById(queryId);
            }
            GUILayout.EndHorizontal();

            if (!string.IsNullOrEmpty(lastQueryResult))
                GUILayout.Label(lastQueryResult, GUI.skin.textArea);

            GUILayout.Space(15);

            if (GUILayout.Button("Get All Rows (first table)"))
            {
                var table = dataManager.GetAllTableAssets();
                if (table.Count > 0 && table[0] != null)
                {
                    lastQueryResult = $"First table has {table[0].Count} rows.";
                }
            }

            if (GUILayout.Button("Build Caches Async"))
            {
                StartCoroutine(dataManager.BuildAllCachesAsync(3, (done, total) =>
                {
                    Debug.Log($"[Demo] Async: {done}/{total}");
                }));
            }

            GUILayout.EndVertical();
            GUILayout.EndArea();
        }

        private string QueryById(int id)
        {
            var tables = dataManager.GetAllTableAssets();
            if (tables.Count == 0) return "No tables loaded.";

            foreach (var table in tables)
            {
                if (table == null) continue;

                // Try to query by ID for each table
                var hasIdMethod = table.GetType().GetMethod("HasId");
                if (hasIdMethod != null && (bool)hasIdMethod.Invoke(table, new object[] { id }))
                {
                    return $"ID {id} found in table '{table.TableName}'";
                }
            }

            return $"ID {id} not found in any table.";
        }
    }
}
