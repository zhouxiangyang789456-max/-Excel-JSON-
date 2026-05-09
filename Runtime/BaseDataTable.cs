using System;
using System.Collections.Generic;
using UnityEngine;

namespace ExcelToJsonPlugin.Runtime
{
    /// <summary>
    /// 所有数据表 ScriptableObject 的抽象基类。
    /// 提供表注册中心所需的公共接口。
    /// </summary>
    public abstract class BaseDataTable : ScriptableObject
    {
        /// <summary>表名（对应 Excel Sheet 名）</summary>
        public abstract string TableName { get; }

        /// <summary>数据总行数</summary>
        public abstract int Count { get; }

        /// <summary>构建 ID → Row 的查询缓存</summary>
        public abstract void BuildCache();

        /// <summary>清空缓存（用于热重载）</summary>
        public abstract void ClearCache();
    }

    /// <summary>
    /// 泛型版本，提供类型安全的行查询。
    /// 自动生成的 Table 类继承此泛型基类。
    /// </summary>
    public abstract class BaseDataTable<TRow> : BaseDataTable
    {
        [SerializeField]
        protected List<TRow> rows = new List<TRow>();

        protected Dictionary<int, TRow> lookup;
        protected Func<TRow, int> idSelector;

        public override int Count => rows?.Count ?? 0;

        public override void BuildCache()
        {
            if (rows == null)
            {
                lookup = new Dictionary<int, TRow>();
                return;
            }

            lookup = new Dictionary<int, TRow>(rows.Count);
            foreach (var row in rows)
            {
                if (row != null)
                {
                    var id = GetRowId(row);
                    lookup[id] = row;
                }
            }
        }

        public override void ClearCache()
        {
            lookup = null;
        }

        protected abstract int GetRowId(TRow row);

        /// <summary>按 ID 查询，O(1)</summary>
        public virtual TRow Get(int id)
        {
            if (lookup == null) BuildCache();
            lookup.TryGetValue(id, out var row);
            return row;
        }

        /// <summary>检查 ID 是否存在</summary>
        public virtual bool HasId(int id)
        {
            if (lookup == null) BuildCache();
            return lookup.ContainsKey(id);
        }

        /// <summary>获取所有数据行</summary>
        public virtual List<TRow> GetAll()
        {
            return new List<TRow>(rows ?? new List<TRow>());
        }

        /// <summary>按条件筛选</summary>
        public virtual List<TRow> Find(Predicate<TRow> match)
        {
            return rows?.FindAll(match) ?? new List<TRow>();
        }

        /// <summary>随机获取一行</summary>
        public virtual TRow GetRandom(Predicate<TRow> filter = null)
        {
            if (rows == null || rows.Count == 0) return default;

            var candidates = filter != null ? rows.FindAll(filter) : rows;
            if (candidates.Count == 0) return default;

            return candidates[UnityEngine.Random.Range(0, candidates.Count)];
        }

        /// <summary>批量按 ID 获取</summary>
        public virtual List<TRow> GetByIds(IEnumerable<int> ids)
        {
            var result = new List<TRow>();
            if (lookup == null) BuildCache();
            foreach (var id in ids)
            {
                if (lookup.TryGetValue(id, out var row))
                    result.Add(row);
            }
            return result;
        }

#if UNITY_EDITOR
        /// <summary>编辑器下填充数据</summary>
        public void SetRows(List<TRow> newRows)
        {
            rows = newRows ?? new List<TRow>();
            lookup = null;
        }
#endif
    }
}
