using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace ExcelToJsonPlugin.Runtime
{
    /// <summary>
    /// 游戏运行时数据中心。
    /// 持有所有数据表的引用，提供类型安全的查询 API。
    ///
    /// 使用方式：
    ///   var weapon = DataManager.Instance.GetTable&lt;WeaponTable&gt;().Get(1001);
    /// </summary>
    public class DataManager : MonoBehaviour
    {
        public static DataManager Instance { get; private set; }

        [SerializeField]
        [Tooltip("所有数据表资产，拖入或由一键生成自动填充")]
        private List<BaseDataTable> allTables = new List<BaseDataTable>();

        private readonly Dictionary<Type, BaseDataTable> tableByType = new Dictionary<Type, BaseDataTable>();
        private readonly Dictionary<string, BaseDataTable> tableByName = new Dictionary<string, BaseDataTable>();

        public int TableCount => allTables.Count;

        // ============================================================
        // 生命周期
        // ============================================================

        private void Awake()
        {
            if (Instance != null)
            {
                Destroy(gameObject);
                return;
            }

            Instance = this;
            DontDestroyOnLoad(gameObject);

            BuildTypeIndex();
        }

        /// <summary>
        /// 构建 Type→Table 索引（轻量，仅遍历表数量，O(表数)）
        /// 不调用 BuildCache，数据缓存延后到 InitializeAsync
        /// </summary>
        private void BuildTypeIndex()
        {
            tableByType.Clear();
            tableByName.Clear();

            foreach (var table in allTables)
            {
                if (table == null) continue;

                tableByType[table.GetType()] = table;

                if (!string.IsNullOrEmpty(table.TableName))
                {
                    tableByName[table.TableName] = table;
                }
            }
        }

        // ============================================================
        // 公共查询 API
        // ============================================================

        /// <summary>
        /// 按类型获取表（推荐）。
        /// 编译时类型安全，表不存在返回 null。
        /// </summary>
        public T GetTable<T>() where T : BaseDataTable
        {
            tableByType.TryGetValue(typeof(T), out var table);
            return table as T;
        }

        /// <summary>
        /// 按名称获取表（字符串方式）。
        /// 用于动态场景，如按配置的表名查询。
        /// </summary>
        public BaseDataTable GetTable(string tableName)
        {
            tableByName.TryGetValue(tableName, out var table);
            return table;
        }

        // ============================================================
        // 初始化 API
        // ============================================================

        /// <summary>
        /// 同步构建所有表的查询缓存。
        /// 表少可直接用（≤15 张），表多请用 BuildAllCachesAsync 或后台线程。
        /// </summary>
        public void BuildAllCaches(Action<int, int> onProgress = null)
        {
            for (int i = 0; i < allTables.Count; i++)
            {
                allTables[i]?.BuildCache();
                onProgress?.Invoke(i + 1, allTables.Count);
            }
        }

        /// <summary>
        /// 分帧异步构建所有缓存（不卡主线程）。
        /// 每帧处理 batchPerFrame 张表，适合加载界面显示进度条。
        /// </summary>
        public IEnumerator BuildAllCachesAsync(
            int batchPerFrame = 3,
            Action<int, int> onProgress = null)
        {
            for (int i = 0; i < allTables.Count; i++)
            {
                allTables[i]?.ClearCache();
                allTables[i]?.BuildCache();
                onProgress?.Invoke(i + 1, allTables.Count);

                if ((i + 1) % batchPerFrame == 0)
                    yield return null;
            }
        }

        /// <summary>
        /// 热重载：重新扫描 allTables 列表中的所有表并重建缓存。
        /// 仅 Editor 下生效。
        /// </summary>
        public void ReloadAllTables()
        {
            BuildTypeIndex();
            BuildAllCaches();
            Debug.Log($"[DataManager] 热重载完成，共 {allTables.Count} 张表");
        }

        // ============================================================
        // Editor 工具
        // ============================================================

#if UNITY_EDITOR
        /// <summary>
        /// 编辑器下用于一键生成时填充列表。
        /// 也用于手动调整后的验证。
        /// </summary>
        public void SetAllTables(List<BaseDataTable> tables)
        {
            allTables = tables ?? new List<BaseDataTable>();
            BuildTypeIndex();
        }

        /// <summary>
        /// 获取当前 all  tables 列表引用（仅供 Editor 用）
        /// </summary>
        public List<BaseDataTable> GetAllTableAssets()
        {
            return allTables;
        }
#endif
    }
}
