# 游戏 Excel 转 JSON 插件 — 设计文档 v3

> 一个 **Unity Editor 插件**，让策划在 Excel 中编辑数据，自动生成强类型的 ScriptableObject / JSON 资产，支持实时校验、代码生成、一键导出。
>
> **v3 更新：** 修复 ClosedXML 无法在 Unity 运行的问题；新增 Schema 迁移、Addressables 集成、热更新、跨列校验、partial class 等遗漏设计。

---

## 1. 产品定位 & 上架分析

### 1.1 产品形态

| 维度 | 决策 |
|------|------|
| **主产品** | Unity Editor 插件（.unitypackage），纯 C# 实现 |
| **辅助工具** | 独立 CLI 版本（.exe），用于 CI/CD 流水线 |
| **目标用户** | 主用户：Unity 手游项目组（策划 + 程序）；次用户：Unreal / Cocos 项目组 |
| **上架渠道** | Unity Asset Store、开源 GitHub、OpenUPM |

### 1.2 Unity Asset Store 上架可行性

| 审核要求 | 本方案满足情况 |
|----------|----------------|
| 必须是 C# 代码，运行在 Unity Editor 内 | ✅ 纯 C#，Editor 文件夹下运行 |
| 必须有 Editor Window 入口 | ✅ 提供 `Window > ExcelToJSON` 菜单 |
| 不能依赖外部 EXE / Python 运行时 | ✅ NPOI 是纯 C# 的 xlsx 读写库，零外部依赖 |
| 必须有文档和示例 | ✅ 附带 Demo 场景、示例 Excel、英文文档 |
| 不能用 AGPL 等传染性协议 | ✅ NPOI 使用 Apache 2.0 协议，允许商用 |

**结论：技术上完全可行。** 用 NPOI（Apache 2.0）作为 Excel 解析库。

### 1.3 Excel 库选型：为什么是 NPOI

| 库 | 能在 Unity 跑吗 | 原因 |
|----|:---:|------|
| **ClosedXML** | ❌ | 依赖 `System.Drawing.Common`，Unity Mono 运行时没有这个 DLL |
| **EPPlus v5+** | ❌ | 同样依赖 `System.Drawing.Common` |
| **EPPlus v4** | ⚠️ 勉强 | LGPL 协议，功能旧，且部分 API 也需要 System.Drawing |
| **ExcelDataReader** | ✅ | MIT 协议，纯 C#，极轻量。但**只能读不能写** |
| **MiniExcel** | ✅ | MIT 协议，纯 C#，性能最好。但功能偏简单，复杂格式解析弱 |
| **NPOI** | ✅ | Apache 2.0，纯 C#，**读写都支持**，经过 15 年+ 验证，Unity 社区多用 |

**最终选择：NPOI 2.5.6 net45** — 经过实际 Unity 2022.3 验证的版本。

**实测结论（2026-05-08）：**

| 尝试 | 结果 | 原因 |
|------|:---:|------|
| NPOI 2.7.0 netstandard2.0 | ❌ 编译通过但 DLL 加载失败 | Unity 缺少 System.Buffers/Memory 等依赖 |
| NPOI 2.7.0 net40 | ❌ 同版本无此目标 | — |
| NPOI 2.5.6 netstandard2.0 | ❌ 同上 | — |
| **NPOI 2.5.6 net45** | ✅ | .NET Framework 4.5  = Unity Editor 原生兼容 |
| NPOI 2.5.6 net45 + BouncyCastle.Crypto | ✅ | NPOI 必需依赖 |
| 不装 BouncyCastle.Crypto | ❌ | `Unable to resolve reference 'BouncyCastle.Crypto'` |

**最终 DLL 组合（5 个文件）：**
- `NPOI.dll`（2.5.6 net45）
- `NPOI.OOXML.dll`（2.5.6 net45）
- `NPOI.OpenXml4Net.dll`（2.5.6 net45）
- `NPOI.OpenXmlFormats.dll`（2.5.6 net45）
- `BouncyCastle.Crypto.dll`（1.8.9 netstandard2.0）

**Unity 项目设置要求：**
- API Compatibility Level: `.NET Framework`（不是 .NET Standard 2.1）
- Unity 版本: 2021.3 LTS 以上

### 1.3 竞品对比

| 工具 | 上架 Unity Store | 可视化面板 | 代码生成 | 数据校验 | 价格 |
|------|:---:|:---:|:---:|:---:|------|
| **Excel Importer (已上架)** | ✅ | ✅ | ❌ | ❌ | $35 |
| **xls2json (GitHub)** | ❌ | ❌ | ❌ | ❌ | 免费 |
| **Unity QuickSheet** | ✅ | ✅ | ✅ | ❌ | $45 |
| **本方案** | ✅ | ✅ | ✅ | ✅ | 开源 / 付费 |

---

## 2. 核心问题：Excel 如何与 C# 代码匹配

这是整个插件最关键的设计问题。**Excel 里的列名 `attack` 怎么变成 C# 里的 `public int attack`，然后在运行时被游戏逻辑调用？**

### 2.1 三种映射模式

本插件提供 **三种映射模式**，覆盖不同团队的工作流偏好：

---

#### 模式 A：Excel 驱动（全自动代码生成）

**适用场景：** 策划先定数据结构，程序不写数据类。适合原型阶段、小型项目。

```
Excel (Item.xlsx)                    自动生成的 C# 代码
┌──────────────────────┐            ┌─────────────────────────────┐
│ id │ name  │ attack  │            │ // Auto-generated            │
│ int│ string│ int     │   ──→      │ [System.Serializable]        │
│ ID │ 名称  │ 攻击力  │            │ public class WeaponRow {     │
│1001│ 新手剑│ 12      │            │     public int    id;        │
│1002│ 铁剑  │ 25      │            │     public string name;      │
└──────────────────────┘            │     public int    attack;    │
                                    │ }                            │
                                    └─────────────────────────────┘

Excel 第 1 行 (字段名) → C# 字段名 (自动 camelCase)
Excel 第 2 行 (类型)   → C# 类型
Excel 第 3 行 (说明)   → C# 注释 / Tooltip
```

**工作流：**
1. 策划在 Excel 中定义字段名和类型
2. 程序在 Editor Window 中点击 **"生成 C# 类"**
3. 插件生成 `WeaponRow.cs` 到 `Assets/Scripts/Generated/`
4. 插件生成对应的 ScriptableObject 数据资产
5. 运行时游戏直接加载 ScriptableObject 使用

**优缺点：**
- ✅ 策划完全自主，不依赖程序定义类
- ✅ 类型永不不匹配（同源生成）
- ❌ 生成的类不能包含业务逻辑（需要扩展类）
- ❌ Excel 类型表达能力有限（不支持嵌套对象、泛型等）

---

#### 模式 B：C# 驱动（反射匹配）

**适用场景：** 程序先定义数据类，策划按类字段填写 Excel。适合成熟项目、有复杂类型需求。

```
C# 手写类 (程序维护)                 Excel (策划按类填写)
┌─────────────────────────────┐     ┌──────────────────────┐
│ [ExcelTable("Weapon")]      │     │ id │ name  │ attack  │
│ public class WeaponRow {    │     │ int│ string│ int     │
│   [ExcelColumn("id")]       │     │ 1001│ 新手剑│ 12      │
│   public int id;            │     └──────────────────────┘
│                             │              ↕ 列名自动匹配
│   [ExcelColumn("name")]     │     Excel 列名 "id"  =  C# 属性上
│   public string name;       │     的 [ExcelColumn("id")] 标签
│                             │
│   [ExcelColumn("attack")]   │     如果 Excel 缺少某列 → 报错
│   public int attack;        │     如果 C# 缺少某字段 → Warning
│                             │
│   // 不在 Excel 中的字段    │
│   public int computedValue; │     [ExcelIgnore] 标记的字段
│   // 在运行时计算           │     Excel 不需要有对应列
│ }                           │
└─────────────────────────────┘
```

**工作流：**
1. 程序在 C# 中定义数据类 + `[ExcelColumn("列名")]` 标签
2. 程序在 Editor Window 中点击 **"导出 Excel 模板"**，插件生成只有表头的空 Excel
3. 策划在生成的模板中填写数据
4. 导出时，插件通过反射扫描所有带 `[ExcelTable]` 的类 → 读取对应 Excel → 按 `[ExcelColumn]` 名称匹配列 → 赋值到 C# 字段 → 生成 ScriptableObject

**优缺点：**
- ✅ 支持任意复杂 C# 类型（嵌套对象、Dictionary、自定义类）
- ✅ 数据类可以包含业务方法、计算属性
- ✅ 程序对数据结构有完全控制权
- ❌ 策划新增字段需要程序先在 C# 加属性
- ❌ 需要维护标签和列名的一致性

---

#### 模式 C：混合模式（⭐推荐正式项目使用）

**适用场景：** 大部分项目的最佳选择。简单表用模式 A，复杂表用模式 B。

```
                    ┌──────────────────┐
                    │   Excel 源文件    │
                    └────────┬─────────┘
                             │
                    ┌────────▼─────────┐
                    │  插件读取 Excel   │
                    └────────┬─────────┘
                             │
              ┌──────────────┼──────────────┐
              ▼              ▼              ▼
     ┌────────────┐  ┌────────────┐  ┌────────────┐
     │ 模式 A 匹配 │  │ 模式 B 匹配 │  │  复合匹配   │
     │ 无对应C#类  │  │ 有[C#属性]  │  │ 部分字段有  │
     │ → 自动生成  │  │ → 反射匹配  │  │ → 自动补全  │
     └──────┬─────┘  └──────┬─────┘  └──────┬─────┘
            │               │               │
            └───────────────┼───────────────┘
                            ▼
              ┌─────────────────────────┐
              │      类型映射引擎        │
              │  Excel类型 → C#类型      │
              │  int      → int          │
              │  float    → float        │
              │  string   → string       │
              │  int[]    → int[]        │
              │  vec3     → Vector3      │
              │  ref:X    → int (外键)   │
              │  enum:X   → X (枚举)     │
              └────────────┬────────────┘
                           ▼
              ┌─────────────────────────┐
              │     校验引擎 (见 §5)     │
              └────────────┬────────────┘
                           ▼
              ┌─────────────────────────┐
              │    ScriptableObject      │
              │    写入 .asset 文件       │
              └─────────────────────────┘
```

---

### 2.2 类型映射表（Excel → C#）

这是映射引擎的核心 — 每一对类型如何转换、校验、序列化：

| Excel 类型 | C# 类型 | Excel 中写法示例 | 校验规则 | 备注 |
|------------|---------|-----------------|----------|------|
| `int` | `int` | `100` | 整数，范围检查 | 最常用 |
| `float` | `float` | `3.14` | 浮点数 | `100` 也合法，自动转 `100.0f` |
| `string` | `string` | `新手剑` | — | 空值 → `""` |
| `bool` | `bool` | `true` / `1` / `是` | 仅接受 true/false/0/1 | 支持中文"是/否" |
| `int[]` | `int[]` | `[1,2,3]` | JSON 数组语法 | 也可用 `1\|2\|3`（策划友好） |
| `float[]` | `float[]` | `[1.5,2.0]` | JSON 数组语法 | |
| `string[]` | `string[]` | `["A","B"]` | JSON 数组语法 | |
| `Vector2` | `Vector2` | `[1.5,2.0]` | 必须恰好 2 个数字 | 自动映射到 Unity Vector2 |
| `Vector3` | `Vector3` | `[1,2,3]` | 必须恰好 3 个数字 | 自动映射到 Unity Vector3 |
| `Color` | `Color` | `#FF0000` 或 `[1,0,0,1]` | 支持 hex 和 RGBA 数组 | |
| `enum:X` | 对应的 C# enum | `Fire` 或 `1` | 值必须在 X 枚举中 | X 是另一个 Sheet 名或 C# enum 名 |
| `ref:X` | `int` | `1001` | 值必须在 X 表的 ID 列中存在 | **外键引用**，导出时校验完整性 |
| `res` | `string`（存资源路径） | `Ui/Icon/Item_001` | 路径格式校验；可选检查资源是否存在 | **见下方 §2.2.1 详解** |
| `res:Sprite` | `string` | `Ui/Icon/Item_001` | 同上 + 导出时可选验证 Unity 资源 | Sprite/Texture2D/AudioClip/GameObject 等 |
| `json` | `string`（存 JSON 文本） | `{"k":"v"}` | 必须是合法 JSON | 不推荐，尽量用结构化列 |
| `loc` | `string` | `ITEM_SWORD` | 本地化 key | 导出后去多语言表查值 |

### 2.2.1 资源类型 `res` 详解

这是策划最常用的类型之一。Excel 里填的是**资源在 Unity 项目中的相对路径**（不需要后缀名），插件负责校验路径合法性，并在生成的 C# 类中提供加载方法。

**Excel 中怎么写：**

```
| id   | name   | icon              | prefab                 | hit_sound           |
| int  | string | res:Sprite        | res:GameObject         | res:AudioClip       |
| ID   | 名称   | 图标              | 预制体                  | 受击音效              |
| 1001 | 新手剑 | Ui/Icon/Item_001  | Prefabs/Weapon/Sword_01 | Audio/SFX/Hit_Wood  |
| 1002 | 铁剑   | Ui/Icon/Item_002  | Prefabs/Weapon/Sword_02 | Audio/SFX/Hit_Metal |
| 1003 | 秘银剑 | Ui/Icon/Item_003  | Prefabs/Weapon/Sword_03 | Audio/SFX/Hit_Magic |
```

**路径规则：**

| 规则 | 示例 | 说明 |
|------|------|------|
| 路径分隔符用 `/` | `Ui/Icon/Item_001` | 不要用 `\` |
| **不带后缀名** | `Ui/Icon/Item_001` 而非 `Item_001.png` | Unity 资源引用不需要后缀 |
| 相对项目根目录 | `Assets/Resources/Ui/Icon/Item_001` → 填 `Ui/Icon/Item_001` | 如果加载方式为 Resources |
| 相对 Assets 目录 | `Assets/Ui/Prefabs/Enemy/Goblin` → 填 `Ui/Prefabs/Enemy/Goblin` | 如果加载方式为 Addressables |
| 大小写敏感 | `Ui/Icon/item_001` ≠ `Ui/Icon/Item_001` | 与文件系统一致 |

**类型声明语法：**

| Excel 类型 | C# 字段类型 | 可选的编辑器校验 | 运行时加载示例 |
|------------|------------|-----------------|---------------|
| `res` | `string` | 只校验路径格式 | 通用资源，加载时自行指定类型 |
| `res:Sprite` | `string` | 校验路径格式 + 检查该路径下是否有 Sprite 资源 | `Resources.Load<Sprite>(icon)` |
| `res:GameObject` | `string` | 同上，检查是否有 GameObject/Prefab | `Resources.Load<GameObject>(prefab)` |
| `res:AudioClip` | `string` | 同上 | `Resources.Load<AudioClip>(hitSound)` |
| `res:Texture2D` | `string` | 同上 | `Resources.Load<Texture2D>(texture)` |
| `res:Material` | `string` | 同上 | `Resources.Load<Material>(material)` |
| `res:AnimationClip` | `string` | 同上 | `Resources.Load<AnimationClip>(anim)` |
| `res:Scene` | `string` | 校验路径格式（场景路径特殊处理） | `SceneManager.LoadScene(scenePath)` |

**校验规则：**

```
Stage 1: 路径格式校验（必定执行）
  □ 不能包含 \ (反斜杠)
  □ 不能以 / 开头或结尾
  □ 不能包含 .. (路径穿越)
  □ 不能包含非法字符 (空格建议用下划线替代)
  □ 长度不超过 255 字符
  
Stage 2: 资源存在性校验（可选，仅 Editor，关闭以提升速度）
  □ res:Sprite → 检查该路径 + ".png/.jpg/.psd" 是否存在
  □ res:GameObject → 检查该路径 + ".prefab" 是否存在
  □ res:AudioClip → 检查该路径 + ".wav/.mp3/.ogg" 是否存在
  □ res:Scene → 检查该路径 + ".unity" 是否存在
  □ 纯 res → 不做存在性校验
```

**生成的 C# 类：**

```csharp
// 自动生成的数据行类
[Serializable]
public partial class WeaponRow
{
    public int id;
    public string name;
    public string icon;           // "Ui/Icon/Item_001"
    public string prefab;         // "Prefabs/Weapon/Sword_01"
    public string hitSound;       // "Audio/SFX/Hit_Wood"
}

// 配套的加载方法（自动生成或手动写在 partial class 中）
public partial class WeaponRow
{
    // Resources 方式加载
    public Sprite LoadIcon() => Resources.Load<Sprite>(icon);
    public GameObject LoadPrefab() => Resources.Load<GameObject>(prefab);
    public AudioClip LoadHitSound() => Resources.Load<AudioClip>(hitSound);

    // Addressables 方式加载
    public async Task<Sprite> LoadIconAsync() =>
        await Addressables.LoadAssetAsync<Sprite>(icon).Task;
}
```

**Editor Window 中的体验：**

```
数据预览 Tab 中，资源类型列的行为：

┌──────┬──────┬──────────────────────┬─────────────────────────────┐
│ id   │ name │ icon (res:Sprite)    │ prefab (res:GameObject)     │
├──────┼──────┼──────────────────────┼─────────────────────────────┤
│ 1001 │新手剑│ Ui/Icon/Item_001  ✅ │ Prefabs/Weapon/Sword_01  ✅ │
│ 1002 │铁剑  │ Ui/Icon/Item_002  ✅ │ Prefabs/Weapon/Sword_02  ✅ │
│ 1003 │木剑  │ Ui/Icon/Item_999  ❌ │ Prefabs/Weapon/WoodSword ✅ │
│      │      │ 资源不存在: Item_999  │                             │
└──────┴──────┴──────────────────────┴─────────────────────────────┘

✅ 绿色 = 路径格式正确 + 资源存在
⚠ 黄色 = 路径格式正确但未检查存在性（快速模式）
❌ 红色 = 路径格式错误或资源不存在

双击单元格 → 在 Project 窗口中定位到该资源
右键单元格 → "在 Finder/Explorer 中显示"
```

**策划填写建议：**

```
正确写法:
  Ui/Icon/Item_001          ← 最简，推荐
  Prefabs/Enemy/Goblin      ← 目录结构清晰
  Audio/BGM/Battle_01       ← 类别/用途/具体资源

错误写法:
  Ui\Icon\Item_001          ← 用了反斜杠
  /Ui/Icon/Item_001         ← 开头多了一个 /
  Ui/Icon/Item_001.png      ← 带了后缀名（插件会自动忽略，但建议不要写）
  Ui/Icon/Item 001          ← 文件名含空格
  ../OtherProject/Icon      ← 路径穿越到项目外
```

### 2.3 完整生命周期：从 Excel 到运行时调用

```
 策划修改 Excel                   程序/CI 处理                     运行时游戏调用
 ═══════════════                 ══════════════                   ════════════════

 编辑 Item.xlsx                  Editor 检测到变更                GameManager 启动
      │                              │                                │
      ▼                              ▼                                ▼
 保存文件                  ExcelParser.Read("Item.xlsx")      DataManager.Instance
      │                              │                          .Initialize()
      ▼                              ▼                                │
 (excel文件)               ┌─ 读取所有 Sheet                     ┌─────┴─────┐
                           │─ 识别字段名/类型/数据行             ▼           ▼
                           │─ 类型转换 & 校验             ScriptableObject  JSON
                           │                                    加载到内存   解析
                           ▼                                      │           │
                    校验通过?                                      ▼           ▼
                     ├─ No → 输出错误列表                  Dictionary<int,    Dictionary<int,
                     │        精确定位单元格                 WeaponRow>       WeaponRow>
                     │        阻止导出                              │           │
                     └─ Yes ↓                                      └─────┬─────┘
                     ┌──────────────┐                                   │
                     │ 映射到 C# 类  │                                   ▼
                     │ (创建对象)    │                          var sword = DataManager
                     └──────┬───────┘                            .Instance.Weapons.Get(1001);
                            │                                   Debug.Log(sword.attack); // 12
                            ▼                                   Debug.Log(sword.name);  // "新手剑"
                     ┌──────────────┐
                     │ Scriptable   │
                     │ Object 写入  │
                     │ .asset 文件  │
                     └──────┬───────┘
                            │
                            ▼
                  Assets/Data/Weapon.asset
                  (Unity 原生资产，可被
                   Addressables 直接引用)
```

**运行时数据访问代码示例：**

```csharp
// ===== DataManager.cs — 运行时数据管理器 =====
public class DataManager : MonoBehaviour
{
    public static DataManager Instance;

    // 在 Inspector 中拖入生成的 ScriptableObject
    public WeaponTable weaponTable;      // ScriptableObject
    public SkillTable skillTable;
    public QuestTable questTable;

    // 按 ID 查询（O(1)）
    public WeaponRow GetWeapon(int id) => weaponTable.Get(id);
}

// ===== 使用示例 =====
var weapon = DataManager.Instance.GetWeapon(1001);
player.Equip(weapon);
Debug.Log($"装备了 {weapon.name}, 攻击力 +{weapon.attack}");
```

---

## 3. 可视化编辑器面板设计

这是策划唯一需要的界面。所有操作都在 Unity Editor 内完成，不需要命令行。

### 3.1 主窗口布局

```
┌─────────────────────────────────────────────────────────────────┐
│  📊 Excel Data Manager                            [⚙] [📖] [✕] │
├────────────────┬────────────────────────────────────────────────┤
│                │  ┌─ 标签页 ──────────────────────────────────────────┐  │
│  📁 文件列表   │  │  📋总览 │ 📝数据预览 │ ✅校验 │ 📦导出 │ 🎮运行时用法 │  │
│                │  ├──────────────────────────────────────────┤  │
│  Excel/        │  │                                          │  │
│  ├─ 📄 Item    │  │  选中: Item.xlsx → Weapon (42 行, 8 列)  │  │
│  │  ├ Weapon   │  │                                          │  │
│  │  ├ Armor    │  │  映射模式: [▼ 模式A: 自动生成C#类      ] │  │
│  │  └ _备注    │  │  目标类:   WeaponRow (自动生成)          │  │
│  ├─ 📄 Skill   │  │                                          │  │
│  │  └ Skill    │  │  ┌────┬──────┬──────┬────┬──────┬─────┐  │  │
│  ├─ 📄 Quest   │  │  │ id │ name │ atk  │ hp │skills│desc │  │  │
│  │  ├ Main     │  │  ├────┼──────┼──────┼────┼──────┼─────┤  │  │
│  │  └ Side     │  │  │1001│新手剑│  12  │100 │[1,2] │ 一..│  │  │
│  ├─ 📄 Enemy   │  │  │1002│铁剑  │  25  │  0 │[1,3] │ 坚..│  │  │
│  └─ 📄 Buff    │  │  │1003│秘银剑│  45  │  0 │[2,4] │ 传..│  │  │
│                │  │  └────┴──────┴──────┴────┴──────┴─────┘  │  │
│  [+添加目录]   │  │                                          │  │
│                │  │  ⚠ 2 个警告: Row 5 "attack"=-5 超出范围  │  │
│                │  │  ✅ 引用完整性: 3/3 外键通过              │  │
│                │  │                                          │  │
│                │  └──────────────────────────────────────────┘  │
│                │                                                │
│                │  [🔄 刷新] [📝 导出 Excel 模板] [⚡ 生成 C# 类]  │
│                │  [✅ 校验全部] [📦 导出 ScriptableObject]      │
├────────────────┴────────────────────────────────────────────────┤
│  ✅ 导出完成: 4 文件 → 6 个 .asset | 230 行数据 | 0 错误 | 2 警告│
│  ⏱ 耗时 1.2s | 上次导出: 2026-05-08 14:30                        │
└─────────────────────────────────────────────────────────────────┘
```

### 3.2 五个标签页详情

#### Tab 1 — 📋 总览（Dashboard）

```
┌─────────────────────────────────────────────┐
│  项目: MyRPGGame                             │
│  Excel 目录: Assets/Excel/                    │
│  输出目录:   Assets/Data/                     │
│  输出格式:   [▼ ScriptableObject]            │
│                                              │
│  快速状态:                                    │
│  ┌──────────┬────────┬────────┬────────┐    │
│  │ Excel 文件│ 数据行 │ 上次导出│ 状态   │    │
│  ├──────────┼────────┼────────┼────────┤    │
│  │ Item.xlsx│ 60 行  │ 14:30  │ ✅ 通过 │    │
│  │ Skill.xsl│ 156 行 │ 14:30  │ ✅ 通过 │    │
│  │ Quest.xls│ 42 行  │ 14:28  │ ✅ 通过 │    │
│  │ Enemy.xls│ 88 行  │ 14:25  │ ⚠ 3警告│    │
│  │ Buff.xlsx│ -      │ 从未   │ ❌ 未配置│   │
│  └──────────┴────────┴────────┴────────┘    │
│                                              │
│  [一键导出全部] [仅校验] [导出选中]           │
└─────────────────────────────────────────────┘
```

#### Tab 2 — 📝 数据预览

- 左栏选择 Sheet，右侧显示完整的表格数据（只读）
- 表头高亮显示：**绿色**=校验通过、**黄色**=有警告、**红色**=有错误
- 点击单元格可查看详细错误信息（浮动 tooltip）
- 搜索栏：`🔍 id=1001` 快速定位行

#### Tab 3 — ✅ 校验

```
┌─────────────────────────────────────────────┐
│  校验结果: Item.xlsx                          │
│                                              │
│  ❌ Weapon (Sheet)  3 个错误                  │
│  ├ [Row 5,  "attack"] 值 -5 不在范围 [0,9999] │
│  ├ [Row 12, "id"]     ID 重复: 1001           │
│  └ [Row 20, "skill"]  外键引用失败:           │
│       Skill 表中不存在 id=999                  │
│                                              │
│  ⚠ Armor  (Sheet)  1 个警告                  │
│  └ [Row 8,  "defense"] 超出推荐范围 (值: 500)  │
│                                              │
│  ✅ Quality (Sheet) 通过                      │
│                                              │
│  双击错误 → 打开 Excel 定位到该单元格          │
│  [复制错误列表] [导出错误 CSV]                 │
└─────────────────────────────────────────────┘
```

#### Tab 4 — 📦 导出

```
┌─────────────────────────────────────────────┐
│  导出配置                                     │
│                                              │
│  输出格式: [▼ ScriptableObject (.asset)    ] │
│  输出路径: Assets/Data/                       │
│                                              │
│  ☑ 导出前自动校验                             │
│  ☑ 校验失败阻止导出                           │
│  ☑ 自动生成 C# 类型定义                       │
│  ☐ 同时导出 JSON (用于 AssetBundle)           │
│  ☐ 同时导出 MessagePack                       │
│                                              │
│  ┌──── 导出进度 ────────────────────────┐    │
│  │ Item.xlsx    ████████████ 100%       │    │
│  │ Skill.xlsx   ████████████ 100%       │    │
│  │ Quest.xlsx   ██████████   90%        │    │
│  │ Enemy.xlsx   ██           15%        │    │
│  └──────────────────────────────────────┘    │
│                                              │
│  [开始导出] [取消]                             │
└─────────────────────────────────────────────┘
```

#### Tab 5 — 🎮 运行时用法（Runtime API）

这是给**程序**看的页面。导出完成后，切到此 Tab 查看如何在游戏代码中使用这些数据。

```
┌─────────────────────────────────────────────────────────────────┐
│  运行时用法 — Weapon (Item.xlsx)                                  │
│                                                                  │
│  ┌─ 快速开始 ─────────────────────────────────────────────────┐  │
│  │                                                            │  │
│  │  📁 生成的文件:                                             │  │
│  │    Assets/Data/Weapon.asset     ← 拖到场景中 DataManager    │  │
│  │    Assets/Scripts/Generated/Data/WeaponRow.cs               │  │
│  │    Assets/Scripts/Generated/Tables/WeaponTable.cs           │  │
│  │                                                            │  │
│  │  [📋 复制 DataManager 模板] [📋 复制初始化代码] [🔗 打开生成文件]│  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  ┌─ API 参考 ────────────────────────────────────────────────┐  │
│  │                                                            │  │
│  │  var t = DataManager.Instance.GetTable<WeaponTable>();      │  │
│  │  t.Get(1001)          // 按 ID 获取，O(1)                   │  │
│  │  t.GetAll()           // 获取所有数据列表                    │  │
│  │  t.Find(w => w.atk>30)// 条件查找                           │  │
│  │  t.HasId(1001)        // 检查 ID 是否存在                   │  │
│  │  t.Count              // 数据总行数                         │  │
│  │  t.GetRandom()         // 随机获取一行                      │  │
│  │  t.GetByIds([1,2,3])  // 批量获取                          │  │
│  │                                                            │  │
│  │  新增表: 导出 .asset → 拖入 allTables → 代码零改动           │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  ┌─ 代码示例 ────────────────────────────────────────────────┐  │
│  │                                                            │  │
│  │  // 示例 1: 获取单个物品                                    │  │
│  │  var t = DataManager.Instance.GetTable<WeaponTable>();      │  │
│  │  var sword = t.Get(1001);                                  │  │
│  │  Debug.Log(sword.name); // "新手剑"                         │  │
│  │                                                            │  │
│  │  // 示例 2: 遍历所有武器                                    │  │
│  │  foreach (var w in t.GetAll())                             │  │
│  │      Debug.Log($"{w.id}: {w.name}");                       │  │
│  │                                                            │  │
│  │  // 示例 3: 按条件筛选                                      │  │
│  │  var legendaries = t.Find(w => w.quality >= 5);             │  │
│  │                                                            │  │
│  │  // 示例 4: 外键引用                                        │  │
│  │  var wt = DataManager.Instance.GetTable<WeaponTable>();     │  │
│  │  var st = DataManager.Instance.GetTable<SkillTable>();      │  │
│  │  var weapon = wt.Get(1001);                                │  │
│  │  var skill = st.Get(weapon.skills[0]);                     │  │
│  │                                                            │  │
│  │  [📋 复制示例 1] [📋 复制示例 2] [📋 复制示例 3] [📋 复制示例 4]│  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  ┌─ Inspector 设置 ───────────────────────────────────────────┐  │
│  │                                                            │  │
│  │  1. 在场景中创建空 GameObject，命名为 "DataManager"          │  │
│  │  2. 挂载 DataManager.cs 脚本                                 │  │
│  │  3. 将 Assets/Data/Weapon.asset 拖入 weaponTable 字段        │  │
│  │  4. 将 Assets/Data/Skill.asset  拖入 skillTable 字段         │  │
│  │  5. ... 依此类推                                            │  │
│  │                                                            │  │
│  │  [🎬 一键生成 DataManager GameObject]                       │  │
│  └────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
```

**"一键生成 DataManager GameObject"按钮的行为：**

插件扫描 `Assets/Data/` 下所有 `.asset` 文件 → 自动生成场景中的 DataManager GameObject → 自动填充所有 Table 引用 → 程序只需在代码中写 `DataManager.Instance.Weapon.Get(1001)` 即可使用。

### 3.3 设置面板（点击 ⚙ 图标打开）

```yaml
# 这些设置都通过 Editor Window 的 UI 控件操作，不需要手写 YAML

┌─────────────────────────────────────────────┐
│  插件设置                                     │
│                                              │
│  Excel 解析:                                  │
│  字段名行: [2▼]  (第几行是字段名)              │
│  类型行:   [3▼]  (第几行是类型声明)            │
│  数据起始: [5▼]  (第几行开始是数据)            │
│  ID 列:    [1▼]  (第几列是主键)               │
│                                              │
│  跳过规则:                                    │
│  ☑ 跳过隐藏行    ☑ 跳过隐藏列                 │
│  ☑ 跳过空行      跳过前缀: [_▼]  [#▼]         │
│                                              │
│  校验:                                        │
│  ☑ 启用类型校验  ☑ 启用ID唯一校验              │
│  ☑ 启用外键校验  严格模式: [☐]                │
│                                              │
│  文件监听:                                    │
│  ☑ 监听 Excel 变更自动导出                     │
│  监听延迟: [500ms ▼]                          │
│                                              │
│  代码生成:                                    │
│  生成路径: Assets/Scripts/Generated/Data/     │
│  命名空间: Game.Data                          │
│  类名后缀: [Row]                              │
│                                              │
│  [恢复默认] [导出配置] [导入配置]               │
└─────────────────────────────────────────────┘
```

---

## 4. 系统架构

### 4.1 整体架构

```
┌─────────────────────────────────────────────────────────────────┐
│                    Unity Editor 层                               │
│  ┌──────────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │ ExcelDataWindow  │  │ SettingsWindow│  │ AssetPostprocessor│ │
│  │ (主编编辑器面板)  │  │ (设置面板)    │  │ (自动导入触发)    │  │
│  └────────┬─────────┘  └──────┬───────┘  └────────┬─────────┘  │
│           │                   │                    │             │
│           └───────────────────┼────────────────────┘             │
│                               ▼                                  │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │                   ExcelDataPipeline                        │  │
│  │                   (核心调度器)                              │  │
│  └───────────────────────────────────────────────────────────┘  │
│                               │                                  │
│     ┌─────────────┬───────────┼───────────┬─────────────┐       │
│     ▼             ▼           ▼           ▼             ▼       │
│  ┌──────┐   ┌──────────┐ ┌──────┐  ┌──────────┐ ┌──────────┐  │
│  │Reader│   │Mapper    │ │Valid.│  │Generator │ │Exporter  │  │
│  │ 读取 │→  │ 映射匹配  │→│ 校验 │→ │ 代码生成  │→│ 资产导出  │  │
│  │Excel │   │Excel→C#  │ │ 引擎 │  │ .cs 文件  │ │ .asset    │  │
│  └──────┘   └──────────┘ └──────┘  └──────────┘ └──────────┘  │
│                                                                  │
├──────────────────────────────────────────────────────────────────┤
│                    Unity 运行时层                                 │
│  ┌──────────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │ DataManager      │  │ ScriptableObject│ │ JSON Loader     │  │
│  │ (数据查询API)    │  │ (生成的资产)   │  │ (AssetBundle用)  │  │
│  └──────────────────┘  └──────────────┘  └──────────────────┘  │
└──────────────────────────────────────────────────────────────────┘
```

### 4.2 目录结构（Unity 项目内）

```
Assets/
├── Excel/                              # ← 策划的 Excel 文件放这里
│   ├── Item.xlsx
│   ├── Skill.xlsx
│   ├── Quest.xlsx
│   └── Enemy.xlsx
│
├── Data/                               # ← 插件自动生成的资产
│   ├── Weapon.asset                    # ScriptableObject（Unity 原生）
│   ├── Armor.asset
│   ├── Skill.asset
│   └── Quest.asset
│
├── Scripts/
│   ├── Generated/                      # ← 插件自动生成的 C# 类
│   │   ├── Data/
│   │   │   ├── WeaponRow.cs            # Excel 模式 A 自动生成
│   │   │   ├── ArmorRow.cs
│   │   │   ├── SkillRow.cs
│   │   │   └── QuestRow.cs
│   │   └── Tables/
│   │       ├── WeaponTable.cs          # ScriptableObject 容器类
│   │       ├── ArmorTable.cs
│   │       └── SkillTable.cs
│   │
│   ├── Runtime/                        # ← 运行时加载代码（手写）
│   │   ├── DataManager.cs              # 数据管理器单例
│   │   ├── IDataTable.cs               # 数据表接口
│   │   └── DataExtensions.cs           # 扩展方法
│   │
│   └── Handwritten/                    # ← 手写数据类（模式 B）
│       ├── ComplexQuestData.cs         # 包含业务逻辑的数据类
│       └── EnemyAIData.cs
│
├── Plugins/
│   └── ExcelToJsonPlugin/              # ← 插件本体
│       ├── Editor/
│       │   ├── ExcelDataWindow.cs      # 主编辑窗口
│       │   ├── SettingsWindow.cs       # 设置窗口
│       │   ├── Core/
│       │   │   ├── ExcelReader.cs      # Excel 文件读取
│       │   │   ├── SchemaParser.cs     # 表头解析（字段/类型/说明）
│       │   │   ├── DataParser.cs       # 数据行解析
│       │   │   ├── TypeMapper.cs       # 类型映射引擎
│       │   │   └── Pipeline.cs         # 核心流水线调度
│       │   ├── Mapping/
│       │   │   ├── AutoMapping.cs      # 模式A：自动生成
│       │   │   ├── AttributeMapping.cs # 模式B：反射匹配
│       │   │   ├── HybridMapping.cs    # 模式C：混合匹配
│       │   │   └── MappingResult.cs    # 匹配结果数据模型
│       │   ├── Validator/
│       │   │   ├── ValidationEngine.cs # 校验引擎
│       │   │   ├── Rules/
│       │   │   │   ├── TypeRule.cs
│       │   │   │   ├── UniqueIdRule.cs
│       │   │   │   ├── RangeRule.cs
│       │   │   │   ├── RequiredRule.cs
│       │   │   │   ├── ForeignKeyRule.cs
│       │   │   │   └── EnumRule.cs
│       │   │   └── ValidationResult.cs
│       │   ├── Generator/
│       │   │   ├── CodeGenerator.cs    # C# 代码生成器
│       │   │   ├── AssetGenerator.cs   # ScriptableObject 生成器
│       │   │   └── JsonExporter.cs     # JSON 导出器
│       │   ├── Watcher/
│       │   │   ├── ExcelFileWatcher.cs # 文件监听
│       │   │   └── AssetPostprocessor.cs
│       │   └── UI/
│       │       ├── ExcelTreeView.cs    # 左侧文件树
│       │       ├── DataPreviewGrid.cs  # 数据预览表格
│       │       ├── ValidationPanel.cs  # 校验结果面板
│       │       └── ExportPanel.cs      # 导出面板
│       │
│       ├── Runtime/
│       │   ├── Attributes.cs           # [ExcelTable], [ExcelColumn] 等
│       │   ├── BaseDataTable.cs        # ScriptableObject 基类
│       │   └── DataTableExtensions.cs  # 扩展方法（Get, GetAll, Find）
│       │
│       └── Dependencies/
│           └── NPOI.dll                # Excel 读写库（Apache 2.0）
│
└── Demo/                               # ← 示例场景
    ├── DemoScene.unity
    ├── DemoExcel/
    │   └── SampleItem.xlsx
    └── DemoScripts/
        └── DemoDataUsage.cs            # 演示如何加载和使用数据
```

---

## 5. 校验引擎详细设计

### 5.1 三阶段校验流程

```
Excel 读取完毕
     │
     ▼
┌── Stage 1: 结构校验 ─────────────────────┐
│  □ 字段名是否重复                          │
│  □ 类型声明是否在支持列表中                 │
│  □ 表头行是否完整（字段名+类型+数据行）     │
│  □ Sheet 是否为空                          │
└──────────────────────────────────────────┘
     │ 通过
     ▼
┌── Stage 2: 数据校验 ─────────────────────┐
│  □ 每个单元格的值是否匹配声明的类型         │
│  □ ID 列是否唯一、非空                     │
│  □ 必填字段是否有值                        │
│  □ 数值是否在配置的范围内                   │
│  □ 字符串是否匹配正则（命名规范等）         │
└──────────────────────────────────────────┘
     │ 通过
     ▼
┌── Stage 3: 引用校验 ─────────────────────┐
│  □ ref:X 的值在 X 表中存在                │
│  □ enum:X 的值在枚举定义中存在             │
│  □ 跨文件引用完整性                        │
└──────────────────────────────────────────┘
     │ 通过
     ▼
   导出
```

### 5.2 校验结果数据结构

```csharp
public class ValidationError
{
    public string FileName;       // "Item.xlsx"
    public string SheetName;      // "Weapon"
    public int Row;               // 5 (Excel 行号)
    public string ColumnName;     // "attack"
    public string RawValue;       // "-5"
    public string RuleName;       // "RangeRule"
    public string Message;        // "攻击力不能为负数 (当前值: -5, 允许范围: 0~9999)"
    public ErrorLevel Level;      // Error | Warning | Info
}

public enum ErrorLevel
{
    Error,    // 阻止导出
    Warning,  // 允许导出但提示
    Info      // 纯提示
}
```

### 5.3 自定义校验规则

策划可以在 Excel 中新增一个名为 `#Rules` 的 Sheet（或通过 Editor Window 配置），按列施加校验：

**方式一：在 Excel 的 `#Rules` Sheet 中定义**

| field | rule | params |
|-------|------|--------|
| attack | range | 0~9999 |
| price | multiple | 10 |
| name | not_empty | |
| quality | enum | Quality |

**方式二：在 C# 中用 Attribute 标注（模式 B）**

```csharp
[ExcelTable("Weapon")]
public class WeaponRow
{
    [ExcelColumn("id")]
    public int id;

    [ExcelColumn("attack")]
    [ValidateRange(0, 9999, Message = "攻击力必须在 0~9999 之间")]
    public int attack;

    [ExcelColumn("name")]
    [ValidateNotEmpty]
    [ValidateRegex(@"^[一-龥a-zA-Z0-9_]+$")]
    public string name;
}
```

---

## 6. 代码生成详细设计

### 6.1 生成的 C# 数据行类

```csharp
// ===== 自动生成: Assets/Scripts/Generated/Data/WeaponRow.cs =====
// 源文件: Assets/Excel/Item.xlsx → Sheet: Weapon
// 生成时间: 2026-05-08 14:30
// ⚠ DO NOT EDIT — 重新导出会覆盖此文件
// 如需添加业务方法，请在另一个文件中写 partial class WeaponRow { }

using System;
using UnityEngine;

namespace Game.Data
{
    [Serializable]
    public partial class WeaponRow
    {
        [Tooltip("ID")]
        public int id;

        [Tooltip("名称")]
        public string name;

        [Tooltip("攻击力")]
        public int attack;

        [Tooltip("生命值")]
        public int hp;

        [Tooltip("技能ID列表")]
        public int[] skills;

        [Tooltip("价格")]
        public int price;

        [Tooltip("品质")]
        public int quality;

        [Tooltip("描述Key")]
        public string desc;

        [Tooltip("图标")]
        public string icon;           // res:Sprite → "Ui/Icon/Item_001"

        [Tooltip("预制体")]
        public string prefab;         // res:GameObject → "Prefabs/Weapon/Sword_01"

        [Tooltip("受击音效")]
        public string hitSound;       // res:AudioClip → "Audio/SFX/Hit_Wood"
    }
}
```

### 6.2 生成的 ScriptableObject 容器类

```csharp
// ===== 自动生成: Assets/Scripts/Generated/Tables/WeaponTable.cs =====
using System.Collections.Generic;
using System.Linq;
using UnityEngine;

namespace Game.Data
{
    [CreateAssetMenu(fileName = "Weapon", menuName = "Game Data/Weapon Table")]
    public class WeaponTable : ScriptableObject
    {
        [SerializeField]
        private List<WeaponRow> rows = new List<WeaponRow>();

        // 运行时缓存：ID → 数据行
        private Dictionary<int, WeaponRow> lookup;

        public void BuildCache()
        {
            lookup = new Dictionary<int, WeaponRow>();
            foreach (var row in rows)
            {
                if (row != null)
                    lookup[row.id] = row;
            }
        }

        public WeaponRow Get(int id)
        {
            if (lookup == null) BuildCache();
            return lookup.TryGetValue(id, out var row) ? row : null;
        }

        public List<WeaponRow> GetAll()
        {
            return new List<WeaponRow>(rows);
        }

        public List<WeaponRow> Find(System.Predicate<WeaponRow> match)
        {
            return rows.FindAll(match);
        }

#if UNITY_EDITOR
        // 编辑器下填充数据用
        public void SetRows(List<WeaponRow> newRows)
        {
            rows = newRows;
            lookup = null;
        }
#endif
    }
}
```

### 6.3 运行时数据管理器（可扩展架构）

项目有 5 张表时，逐个声明字段没问题。但 50 张、100 张表时，下面这个硬编码模式会直接炸掉。所以需要两套方案。

#### 方案 A：小型项目（≤15 张表）—— 硬编码字段

```csharp
// ===== 手写: Assets/Scripts/Runtime/DataManager.cs =====
using UnityEngine;
using Game.Data;

public class DataManager : MonoBehaviour
{
    public static DataManager Instance { get; private set; }

    // Inspector 拖入 .asset 文件
    public WeaponTable weaponTable;
    public ArmorTable armorTable;
    public SkillTable skillTable;
    public QuestTable questTable;

    private void Awake()
    {
        if (Instance != null) { Destroy(gameObject); return; }
        Instance = this;
        DontDestroyOnLoad(gameObject);

        // 仅 Awake 构建缓存（表少，不会卡）
        weaponTable?.BuildCache();
        armorTable?.BuildCache();
        skillTable?.BuildCache();
        questTable?.BuildCache();
    }

    // 便捷查询（编译时类型安全）
    public WeaponRow GetWeapon(int id) => weaponTable?.Get(id);
    public ArmorRow GetArmor(int id) => armorTable?.Get(id);
    public SkillRow GetSkill(int id) => skillTable?.Get(id);
    public QuestRow GetQuest(int id) => questTable?.Get(id);
}
```

#### 方案 B：中大型项目（>15 张表，⭐推荐）—— 表注册中心

**核心思路：** DataManager 不持有具体类型的字段，改为持有一个 `List<BaseDataTable>`，通过泛型按类型查询。新增表只需导出 .asset 并拖入列表，**零代码改动**。

```csharp
// ===== BaseDataTable.cs — 所有 Table 的基类 =====
public abstract class BaseDataTable : ScriptableObject
{
    // 子类实现：构建自己的 ID→Row 缓存
    public abstract void BuildCache();
    // 子类实现：返回缓存大小
    public abstract int Count { get; }
    // 子类实现：返回数据类型名
    public abstract string TableName { get; }
}
```

```csharp
// ===== WeaponTable.cs 改为继承 BaseDataTable =====
public class WeaponTable : BaseDataTable
{
    [SerializeField] private List<WeaponRow> rows;
    private Dictionary<int, WeaponRow> lookup;

    public override string TableName => "Weapon";
    public override int Count => rows?.Count ?? 0;

    public override void BuildCache()
    {
        lookup = new Dictionary<int, WeaponRow>(rows?.Count ?? 0);
        if (rows == null) return;
        foreach (var row in rows)
            if (row != null) lookup[row.id] = row;
    }

    public WeaponRow Get(int id)
    {
        if (lookup == null) BuildCache();
        lookup.TryGetValue(id, out var row);
        return row;
    }

    public List<WeaponRow> GetAll() => new List<WeaponRow>(rows ?? new List<WeaponRow>());
    public List<WeaponRow> Find(Predicate<WeaponRow> m) => rows?.FindAll(m);
}
```

```csharp
// ===== DataManager.cs — 注册中心版 =====
using System;
using System.Collections.Generic;
using UnityEngine;

public class DataManager : MonoBehaviour
{
    public static DataManager Instance { get; private set; }

    // 唯一的 Inspector 列表 —— 所有 .asset 都拖到这里
    [SerializeField]
    private List<BaseDataTable> allTables = new List<BaseDataTable>();

    // 类型 → Table 实例，O(1) 查询
    private Dictionary<Type, BaseDataTable> tableByType = new Dictionary<Type, BaseDataTable>();
    // 名称 → Table 实例（备选查询方式）
    private Dictionary<string, BaseDataTable> tableByName = new Dictionary<string, BaseDataTable>();

    public int TableCount => allTables.Count;

    private void Awake()
    {
        if (Instance != null) { Destroy(gameObject); return; }
        Instance = this;
        DontDestroyOnLoad(gameObject);

        // 构建类型索引
        foreach (var table in allTables)
        {
            if (table == null) continue;
            tableByType[table.GetType()] = table;
            tableByName[table.TableName] = table;
        }
    }

    // ===== 核心 API =====

    // 按类型获取 Table（推荐，编译时类型安全）
    public T GetTable<T>() where T : BaseDataTable
    {
        tableByType.TryGetValue(typeof(T), out var table);
        return table as T;
    }

    // 按名称获取 Table（字符串方式，灵活性高）
    public BaseDataTable GetTable(string tableName)
    {
        tableByName.TryGetValue(tableName, out var table);
        return table;
    }

    // 构建所有表的缓存（支持进度回调）
    public void BuildAllCaches(Action<int, int> onProgress = null)
    {
        for (int i = 0; i < allTables.Count; i++)
        {
            allTables[i]?.BuildCache();
            onProgress?.Invoke(i + 1, allTables.Count);
        }
    }

    // 异步构建所有缓存（分帧，避免卡顿）
    public System.Collections.IEnumerator BuildAllCachesAsync(
        int batchPerFrame = 3,
        Action<int, int> onProgress = null)
    {
        for (int i = 0; i < allTables.Count; i++)
        {
            allTables[i]?.BuildCache();
            onProgress?.Invoke(i + 1, allTables.Count);

            // 每处理 N 张表，等待一帧
            if ((i + 1) % batchPerFrame == 0)
                yield return null;
        }
    }
}
```

**使用对比：**

```csharp
// 方案 A（硬编码）:
var weapon = DataManager.Instance.GetWeapon(1001);

// 方案 B（注册中心）:
var weapon = DataManager.Instance.GetTable<WeaponTable>().Get(1001);
//                    ↑ 泛型参数指定要查哪张表，编译时类型安全

// 查询不存在的表 → 返回 null，不会崩溃
var missing = DataManager.Instance.GetTable<NonExistentTable>();
// missing == null ✅
```

**新增表零改动流程：**

```
策划新建 Enemy.xlsx → 导出 Enemy.asset
    │
    ▼
在 Inspector 中将 Enemy.asset 拖入 DataManager 的 allTables 列表
（"一键生成"按钮会自动完成这步）
    │
    ▼
代码中直接使用，无需修改 DataManager.cs：
    var enemy = DataManager.Instance.GetTable<EnemyTable>().Get(3001);
```

**对比总结：**

| | 方案 A（硬编码） | 方案 B（注册中心） |
|---|---|---|
| 适用规模 | ≤15 张表 | 任意规模 |
| 新增表改动 | 改 DataManager.cs + 拖 Inspector | 仅拖 Inspector（或自动） |
| 查询方式 | `Instance.GetWeapon(id)` | `Instance.GetTable<WeaponTable>().Get(id)` |
| 查询不存在的表 | 空引用异常 | 返回 null，安全 |
| 编译时类型检查 | ✅ | ✅ (泛型) |
| 表名冲突检测 | ❌ | ✅ |

> 插件默认生成方案 B 的代码。如果项目表很少，可以手动选择方案 A 的代码模板。

---

## 7. 运行时使用指南（程序必读）

这是导出完成之后的事情——**程序如何在游戏代码中加载和使用数据**。对应 Editor Window 的 Tab 5（🎮 运行时用法）。

### 7.1 整体流程（3 步）

```
导出完成          程序接入                     运行时使用
═════════        ══════════                   ══════════

插件生成:        编辑器中:                     游戏代码中:
 Weapon.asset    创建 DataManager GO          按 ID 查数据
 WeaponRow.cs    拖入所有 .asset 引用          遍历/筛选数据
 WeaponTable.cs  调整加载顺序 (可选)           外键跳转查询
```

### 7.2 第一步：场景中创建 DataManager

插件可以**一键生成**，也可以手动操作：

**自动方式（推荐）：** 在 Editor Window > Tab 5 中点击 `[一键生成 DataManager GameObject]`：

```
插件自动执行:
  1. 扫描 Assets/Data/ 下所有 .asset 文件
  2. 检查场景中是否已有 DataManager GameObject
     └── 已有 → 只更新 allTables 列表，不重建 GameObject
     └── 没有 → 新建名为 "DataManager" 的 GameObject
  3. 挂载/更新 DataManager.cs 脚本
  4. 自动将所有 .asset 拖入 allTables 列表（按字母排序）
  5. 打印日志:
     "DataManager 已就绪，注册了 12 张数据表"
     "  Weapon      (60 行)"
     "  Armor       (18 行)"
     "  Skill       (156 行)"
     "  ..."

结果 (方案 B 注册中心):
  Hierarchy:
    DataManager (GameObject)
       DataManager.cs
          All Tables  (List<BaseDataTable>)
          ├─ Element 0 → Weapon.asset
          ├─ Element 1 → Armor.asset
          ├─ Element 2 → Skill.asset
          ├─ ...
          └─ Element 11→ Buff.asset

关键区别:
  方案 B 只有一个 allTables 列表，新增表自动追加到列表末尾
  不需要为每张表单独声明 public 字段
  代码层面零改动
```

**手动方式：** 创建空 GameObject → Add Component → DataManager → 展开 allTables 列表 → 逐个拖入 .asset 文件。

### 7.3 第二步：游戏启动时初始化

**问题：** 50 张表 × 1000 行在 Awake 里同步 BuildCache 会卡住主线程 2-5 秒。

**解法：** Awake 只做轻量注册（构建类型索引），BuildCache 延后到异步初始化。

```csharp
// ===== 游戏入口脚本 =====
public class GameBootstrap : MonoBehaviour
{
    public LoadingScreen loadingScreen;  // 加载界面引用

    async void Start()
    {
        // Phase 0: Awake 已完成，DataManager 存在
        //          只做了类型索引构建（O(表数量)，几乎无开销）

        // Phase 1: 异步构建所有表的查询缓存
        await BuildAllCachesWithProgress();

        // Phase 2: 检查热更新（见 §7.5）
        await DataManager.Instance.CheckForHotfixAsync();

        // Phase 3: 进入游戏
        loadingScreen.Hide();
        EnterGame();
    }

    private async Task BuildAllCachesWithProgress()
    {
        var total = DataManager.Instance.TableCount;
        var completed = 0;

        await Task.Run(() =>
        {
            DataManager.Instance.BuildAllCaches((current, total) =>
            {
                completed = current;
            });
        });

        // 或者用分帧协程（不卡主线程）：
        // yield return DataManager.Instance.BuildAllCachesAsync(
        //     batchPerFrame: 5,
        //     onProgress: (c, t) => loadingScreen.SetProgress(c, t)
        // );
    }
}

// 加载界面示例
public class LoadingScreen : MonoBehaviour
{
    public Text tipText;
    public Slider progressBar;

    public void SetProgress(int current, int total)
    {
        progressBar.value = (float)current / total;
        tipText.text = $"正在加载数据... {current}/{total} 张表";
    }

    public void Hide() { gameObject.SetActive(false); }
}
```

**为什么 Awake 不会卡了：**

| | 旧方案（硬编码 Awake） | 新方案 |
|---|---|---|
| Awake 做的事 | 遍历所有 Row 构建 Dictionary | 只遍历 allTables 列表建 Type→Table 索引 |
| 操作量 | 表数量 × 每表行数 | 仅表数量 |
| 100 张 × 1000 行耗时 | 2~5 秒（卡） | < 1ms（不卡） |
| BuildCache 时机 | Awake 中阻塞 | 延后到异步初始化，有进度条 |

### 7.4 第三步：在游戏逻辑中使用数据

#### 7.4.1 基础查询

```csharp
// 获取 Table 引用（一次获取，反复使用）
var weaponTable = DataManager.Instance.GetTable<WeaponTable>();
var skillTable = DataManager.Instance.GetTable<SkillTable>();

// 有效性检查（表不存在时返回 null，不抛异常）
if (weaponTable == null)
{
    Debug.LogError("WeaponTable 未注册到 DataManager");
    return;
}

// 1. 按 ID 精确查询 —— 最常用，O(1)
var weapon = weaponTable.Get(1001);
if (weapon != null)
{
    player.atk += weapon.attack;
    ui.ShowItemName(weapon.name);
}

// 2. 检查 ID 是否存在
if (skillTable.HasId(skillId))
{
    player.LearnSkill(skillId);
}

// 3. 获取所有数据
foreach (var w in weaponTable.GetAll())
{
    shopPanel.AddItem(w.id, w.name, w.price);
}

// 4. 获取随机一行（掉落/抽卡用）
var randomLoot = weaponTable.GetRandom(
    w => w.quality >= 3  // 可选种子: 仅限稀有以上
);

// 5. 批量获取
var ids = new[] { 1001, 1002, 1003 };
var weapons = weaponTable.GetByIds(ids);
```

#### 7.4.2 条件筛选

```csharp
var weaponTable = DataManager.Instance.GetTable<WeaponTable>();
var skillTable = DataManager.Instance.GetTable<SkillTable>();

// 按条件筛选——内部遍历全表，数据量大时注意性能
var cheapWeapons = weaponTable.Find(w => w.price <= 500);
var fireSkills = skillTable.Find(s => s.element == Element.Fire);
```

#### 7.4.3 外键关联查询

```csharp
// Excel 中 weapon.skills 类型为 int[]（存的是 Skill 表的 id 列表）
// 运行时手动关联查询：

var weaponTable = DataManager.Instance.GetTable<WeaponTable>();
var skillTable = DataManager.Instance.GetTable<SkillTable>();

var weapon = weaponTable.Get(1001);

foreach (var skillId in weapon.skills)
{
    var skill = skillTable.Get(skillId);
    if (skill != null)
    {
        Debug.Log($"武器 {weapon.name} 拥有技能: {skill.name} (威力: {skill.power})");
    }
}
```

#### 7.4.4 联合查询封装

```csharp
// 推荐在 partial class 中封装关联查询方法
// 文件: WeaponRow.cs (partial class, 不会被自动生成覆盖)

public partial class WeaponRow
{
    // 获取该武器的所有技能对象
    public List<SkillRow> GetSkills()
    {
        var result = new List<SkillRow>();
        if (skills == null) return result;

        var skillTable = DataManager.Instance.GetTable<SkillTable>();
        if (skillTable == null) return result;

        foreach (var skillId in skills)
        {
            var skill = skillTable.Get(skillId);
            if (skill != null) result.Add(skill);
        }
        return result;
    }

    // 获取该武器的品质枚举
    public QualityType QualityType => (QualityType)quality;
}
```

### 7.5 热更新环境（手游上线后）

内建 ScriptableObject 作为**兜底数据**，CDN 上的 JSON 作为**最新数据**：

```csharp
public async Task InitializeAsync()
{
    // Phase 1: 从 SO 加载内建数据（保证游戏一定能跑）
    weaponTable?.BuildCache();
    skillTable?.BuildCache();

    // Phase 2: 检查 CDN 是否有更新
    var cdnUrl = RemoteConfig.Instance.DataCdnBaseUrl;
    var updater = new DataUpdater(cdnUrl);

    try
    {
        var manifest = await updater.FetchManifestAsync();
        foreach (var kv in manifest.Tables)
        {
            var tableName = kv.Key;       // "Weapon"
            var remoteHash = kv.Value;    // "a1b2c3d4"

            // 本地缓存的 hash
            var localHash = PlayerPrefs.GetString($"data_hash_{tableName}", "");

            if (remoteHash != localHash)
            {
                // 有新数据，下载并替换
                var json = await updater.DownloadTableJsonAsync(tableName);
                ReplaceFromJson(tableName, json);
                PlayerPrefs.SetString($"data_hash_{tableName}", remoteHash);
                Debug.Log($"数据更新: {tableName}");
            }
        }
    }
    catch (System.Exception ex)
    {
        Debug.LogWarning($"热更新数据获取失败，使用内建数据: {ex.Message}");
    }
}
```

### 7.6 Addressables 加载方式

如果项目用 Addressables，生成的数据会自动注册到 `GameData` Group：

```csharp
// 异步加载
var handle = Addressables.LoadAssetAsync<WeaponTable>("Data/Weapon");
await handle.Task;
var weaponTable = handle.Result;

// 或直接用 DataManager 封装
public async Task LoadTablesAsync()
{
    var tasks = new List<Task>();
    tasks.Add(LoadTable<WeaponTable>("Data/Weapon", t => weaponTable = t));
    tasks.Add(LoadTable<SkillTable>("Data/Skill", t => skillTable = t));
    await Task.WhenAll(tasks);
}
```

### 7.7 常见问题

| 问题 | 解答 |
|------|------|
| **新增表后需要改代码吗？** | 方案 B 不需要。导出 .asset → 拖入 allTables → 代码用 `GetTable<T>()`。零改动 |
| **表查询不到会崩吗？** | 不崩。`GetTable<T>()` 返回 null，加个判空即可 |
| **什么时候调用 BuildCache？** | Awake 只建索引。BuildCache 在 `InitializeAsync()` 中异步调用，有进度条 |
| **100 张表启动会卡多久？** | Awake <1ms（只建索引）。BuildCache 分帧或后台线程，主线程不卡 |
| **数据表能运行时修改吗？** | SO 在运行时可读写（不回写磁盘），修改仅本次运行有效 |
| **多场景怎么共享 DataManager？** | 已标记 `DontDestroyOnLoad`，多场景自动保持 |
| **Excel 改了数据，运行中游戏会更新吗？** | 不会自动更新。Editor 下用 Ctrl+Shift+R 热重载（见 §7.8） |

### 7.8 运行时热重载（Editor 开发用）

在 Editor 下运行游戏时，如果策划修改了 Excel 并重新导出，**不需要停止 Play Mode**：

```csharp
#if UNITY_EDITOR
void Update()
{
    // 检测快捷键 Ctrl+Shift+R 触发运行时数据重载
    if (Input.GetKey(KeyCode.LeftControl) &&
        Input.GetKey(KeyCode.LeftShift) &&
        Input.GetKeyDown(KeyCode.R))
    {
        DataManager.Instance.ReloadAllTablesFromAssets();
        Debug.Log("运行时数据已刷新");
    }
}
#endif
```

---

## 8. 文件监听 & 自动导出

```
Unity Editor 启动
       │
       ▼
ExcelFileWatcher 开始监听 Assets/Excel/ 目录
       │
       ▼
策划用 Excel 打开 Item.xlsx，修改 attack 值 12→15，Ctrl+S 保存
       │
       ▼
FileSystemWatcher 检测到 Item.xlsx 变更
       │
       ▼
DebounceTimer.Start(500ms)  ← 防抖（Excel 保存会触发多次变更事件）
       │
       ▼ (500ms 内无新变更)
触发 ExcelDataPipeline.Process("Item.xlsx")
       │
       ├─ 1. ExcelReader.Read()        读取数据
       ├─ 2. SchemaParser.Parse()      解析结构
       ├─ 3. MappingEngine.Map()       匹配 C# 类
       ├─ 4. ValidationEngine.Validate() 校验数据
       │     │
       │     ├─ 有错误 → 停止，Console 报错，弹窗提示
       │     └─ 无错误 ↓
       ├─ 5. AssetGenerator.Generate() 生成 .asset
       └─ 6. AssetDatabase.Refresh()   刷新 Unity
              │
              ▼
       Inspector 中 ScriptableObject 数据已更新
       游戏中使用 DataManager.GetWeapon(1001).attack 立即获得新值 15
```

**防抖实现要点：**

```csharp
// 每个文件独立的 debounce timer
private Dictionary<string, Timer> debounceTimers;

void OnFileChanged(string filePath)
{
    if (debounceTimers.TryGetValue(filePath, out var timer))
        timer.Reset(500);         // 重置计时器
    else
    {
        timer = new Timer(500);   // 新建计时器
        timer.Elapsed += () => ProcessFile(filePath);
        debounceTimers[filePath] = timer;
    }
    timer.Start();
}
```

---

## 9. CLI 版本（CI/CD 用）

Unity Editor 面板用于日常开发，CLI 版本用于 CI/CD 流水线：

```bash
# 将 CLI 打包为独立的 .exe（.NET 独立发布）
dotnet publish -c Release -r win-x64 --self-contained -o ./build

# CI 中使用
ExcelToJson.Cli export \
  --excel-dir ./Assets/Excel \
  --output-dir ./Assets/Data \
  --format json \
  --strict-validation

# 仅校验
ExcelToJson.Cli validate \
  --excel-dir ./Assets/Excel \
  --exit-code-on-error   # 有错误则 exit(1)，让 CI 失败
```

**CLI 与 Editor 共享同一套 Core 库**，只是入口不同：
- Editor 入口：`ExcelDataWindow.cs` → Unity Editor UI
- CLI 入口：`Program.cs` → 命令行

---

## 10. 实施路线图

### Phase 1 — MVP（4 周）
```
目标：策划能在 Excel 里编辑数据，程序能在 Unity 里用 ScriptableObject 获取数据

□ NPOI 集成，实现 ExcelReader（读取 .xlsx / .xls）
□ SchemaParser：字段名行 + 类型行 + 数据行 识别
□ DataParser：单元格类型转换（int/float/string/bool/int[]）
□ Mapping — 模式 A（自动生成 C# 类）
□ CodeGenerator：从 Schema 生成 C# Row 类 + Table 类
□ AssetGenerator：生成 ScriptableObject .asset 文件
□ Editor Window — 主窗口框架 + 文件列表 + 数据预览
□ 编辑器内导出按钮
```

### Phase 2 — 核心体验（3 周）
```
□ 校验引擎：类型匹配、ID唯一、范围检查、外键引用
□ 校验面板：错误列表 + 双击定位
□ Mapping — 模式 B（反射匹配 [ExcelColumn] 标签）
□ 导出 Excel 模板功能
□ 文件监听 + 自动导出
□ 增量导出（hash 缓存）
```

### Phase 3 — 上架准备（2 周）
```
□ 设置面板
□ Mapping — 模式 C（混合匹配）
□ enum:X 和 ref:X 复合类型
□ 错误处理 & 容错完善
□ Demo 场景 + 示例 Excel + 文档
□ Asset Store 提交准备（截图、描述、定价）
□ GitHub 开源仓库 + README
```

### Phase 4 — 扩展（按需）
```
□ JSON / MessagePack 导出（用于 AssetBundle）
□ CLI 独立版本
□ 自定义校验脚本（C# ScriptableObject 规则）
□ Unreal / Cocos 适配调研
□ VS Code 扩展（为策划提供轻量方案）
```

---

## 11. 关键技术问题 & 解决方案

### 10.1 为什么 NPOI 是唯一正确的选择

| 库 | Unity 可用 | 读 .xlsx | 写 .xlsx | 协议 | 结论 |
|----|:---:|:---:|:---:|------|------|
| **ClosedXML** | ❌ | ✅ | ✅ | MIT | `System.Drawing` 在 Unity 中缺失 |
| **EPPlus v5+** | ❌ | ✅ | ✅ | 商用付费 | 同上 |
| **EPPlus v4** | ⚠️ | ✅ | ⚠️ | LGPL | 旧版，部分 API 也需要 System.Drawing |
| **ExcelDataReader** | ✅ | ✅ | ❌ | MIT | 只读不写，无法生成 Excel 模板 |
| **MiniExcel** | ✅ | ✅ | ✅ | MIT | API 太简单，复杂格式解析弱 |
| **NPOI** | ✅ | ✅ | ✅ | Apache 2.0 | **唯一全功能在 Unity 中可用的方案** |

### 10.2 大文件性能

| 场景 | 预期表现 |
|------|----------|
| 10 个 Sheet，每 Sheet 500 行 | < 1s |
| 100 个 Sheet，每 Sheet 2000 行 | < 8s |
| 1000 个 Sheet，每 Sheet 5000 行 | < 60s |

优化策略：
- **异步导出**：Unity Editor 中使用 `EditorApplication.update` 分帧处理
- **增量缓存**：仅处理变更的 Sheet，未变更的直接复用上次结果
- **并行 Sheet**：同一文件内多个 Sheet 用 `Parallel.ForEach` 并行

### 10.3 合并单元格处理

Excel 中策划经常合并单元格（如"武器种类"合并多行）。处理策略：

```
┌────┬──────┬──────┐
│种类 │ 名称  │攻击力 │
├────┼──────┼──────┤   合并单元格 "武器" 跨 3 行
│武器 │ 铁剑  │ 25   │    ↓
├────┼──────┼──────┤   解析时展开为每行都有值:
│武器 │ 秘银剑│ 40   │   Row 1: 武器, 铁剑, 25
├────┼──────┼──────┤   Row 2: 武器, 秘银剑, 40
│武器 │ 短剑  │ 10   │   Row 3: 武器, 短剑, 10
└────┴──────┴──────┘
```

---

## 12. 附录：完整示例

### 11.1 策划操作流程

```
第 1 步：策划打开 Excel，新建 Item.xlsx
         按模板格式填写 字段名行 + 类型行 + 说明行 + 数据
         
第 2 步：策划把 Item.xlsx 放入项目的 Assets/Excel/ 目录
         或用 Unity 的拖拽功能拖入 Editor Window

第 3 步：策划在 Unity 中打开 Window > Excel Data Manager
         看到 Item.xlsx 出现在左侧文件列表中
         点击它，右侧预览表格数据

第 4 步：策划点击 [⚡ 生成 C# 类] 按钮
         插件自动生成 WeaponRow.cs 和 WeaponTable.cs

第 5 步：策划点击 [📦 导出] 按钮
         插件校验数据 → 生成 Weapon.asset 到 Assets/Data/

第 6 步：策划可以在 Inspector 中查看 Weapon.asset
         确认数据正确

之后每次修改 Excel：
  改数据 → Ctrl+S → 插件自动重新导出 → .asset 自动更新
  不需要再点任何按钮
```

### 11.2 程序使用数据流程

```csharp
// 1. 在场景中创建 GameObject，挂载 DataManager
// 2. 在 Inspector 中将生成的 Weapon.asset 拖入 DataManager 的 weaponTable 字段
// 3. 代码中直接使用：

public class PlayerEquipment : MonoBehaviour
{
    void Start()
    {
        // 按 ID 精确查询
        var sword = DataManager.Instance.GetWeapon(1001);
        if (sword != null)
        {
            Debug.Log($"获得武器: {sword.name}, 攻击力: {sword.attack}");
            // 输出: 获得武器: 新手剑, 攻击力: 12
        }

        // 获取所有武器列表
        var allWeapons = DataManager.Instance.weaponTable.GetAll();
        foreach (var w in allWeapons)
        {
            Debug.Log($"  {w.id}: {w.name}");
        }

        // 按条件查找
        var highAttack = DataManager.Instance.weaponTable
            .Find(w => w.attack > 30);
    }
}
```

### 11.3 外键引用示例

```
Item.xlsx → Sheet: Weapon
┌──────┬──────┬──────────┐
│ id   │ name │skill_ref │   ← skill_ref 类型声明为 ref:Skill
│ int  │string│ ref:Skill│
├──────┼──────┼──────────┤
│ 1001 │新手剑│ 1        │   ← 引用 Skill 表的 id=1
│ 1002 │铁剑  │ 3        │   ← 引用 Skill 表的 id=3
└──────┴──────┴──────────┘

Skill.xlsx → Sheet: Skill
┌──────┬──────┬──────┐
│ id   │ name │power │
│ int  │string│ int  │
├──────┼──────┼──────┤
│ 1    │ 重击 │ 150  │   ← id=1 存在 ✅
│ 2    │ 旋风 │ 80   │
│ 3    │ 火球 │ 200  │   ← id=3 存在 ✅
└──────┴──────┴──────┘

校验结果:
  ✅ Weapon Row 1001: skill_ref=1 → Skill 表中存在
  ✅ Weapon Row 1002: skill_ref=3 → Skill 表中存在

如果策划写了 skill_ref=99：
  ❌ Weapon Row 1003: skill_ref=99 → Skill 表中不存在 id=99
```

---

## 13. 风险 & 应对

| 风险 | 概率 | 影响 | 应对 |
|------|------|------|------|
| NPOI 大文件（100MB+ Excel）解析慢 | 低 | 用户等待久 | NPOI 支持流式读取（XSSF），逐行处理不全部加载；异步+分帧 |
| Excel 格式不统一 | 高 | 解析失败 | 提供可配置的表头行号，支持自定义布局模板 |
| 策划不遵守模板格式 | 高 | 导出报错 | 强校验+清晰错误提示+提供模板导出+**格式自动修复** |
| Unity Asset Store 审核拒 | 低 | 无法上架 | 提前阅读审核指南，NPOI 无外部依赖 |
| 竞品功能覆盖 | 中 | 竞争力弱 | 校验+代码生成+热更新+Addressables 的组合是差异化优势 |
| ScriptableObject 单文件过大导致 Editor 卡顿 | 中 | 编辑器体验差 | 单表超过 5000 行自动切换为 JSON 模式；提供分页预览 |
| Schema 变更导致旧代码编译失败 | 高 | 开发效率降低 | partial class + Schema 迁移系统 (见 §14) |
| 多人同时编辑 Excel 产生冲突 | 中 | 数据丢失 | 以 Excel 为源，git 无法 diff → 提供内置 diff 工具 (见 §15) |
| 热更新时数据版本不兼容 | 中 | 运行时崩溃 | Schema 版本号 + 数据迁移脚本机制 (见 §16) |
| NPOI 公式求值不完整 | 中 | 缓存值过期 | 导出前检测公式单元格，调用 Excel 刷新缓存或警告用户 |

---

## 14. Schema 迁移系统（新增）

**问题：** 策划在 Excel 中改了列名、类型、增删列，会导致：
- 模式 A：自动生成的 C# 类变了 → 所有引用旧字段的代码编译报错
- 模式 B：C# 标签名与 Excel 列名不匹配 → 导出失败  
- 旧的 ScriptableObject 资产与新类结构不一致 → 加载异常

### 13.1 变更检测

每次导出前，对比当前 Excel schema 与上次导出的 "schema 快照"：

```csharp
// .cache/schema_snapshots/Weapon.schema.json （插件自动维护）
{
  "excel": "Item.xlsx",
  "sheet": "Weapon",
  "version": 3,                    // schema 版本号
  "columns": [
    {"name": "id",     "type": "int",    "index": 0},
    {"name": "name",   "type": "string", "index": 1},
    // "desc" 列被删除了 ← 快照中没有这一列
    {"name": "attack", "type": "int",    "index": 2}
  ]
}
```

**变更类型识别：**

| 变更 | 检测方式 | 处理 |
|------|----------|------|
| 新增列 | 快照中无此列 | 自动在 C# 类中新增字段（默认值填充旧数据） |
| 删除列 | 快照中有、Excel 中无 | **阻止导出**，弹窗要求策划确认（防止误删） |
| 重命名列 | 快照和 Excel 列索引相同但名称不同 | 弹窗询问："检测到'atk'可能被重命名，是改名还是新增？" |
| 类型变更 | 同名列类型不同 | **阻止导出**，提示"id 类型从 int 变为 string，需要手动处理" |
| 列顺序变更 | 列名还在但 index 不同 | 自动适配，无需人工干预 |

### 13.2 自动迁移流程

```
检测到 Schema 变更
       │
       ▼
┌──────────────────────────────┐
│ 变更分析器输出差异报告        │
│ + added:    [crit_rate]     │
│ - removed:  [desc]          │
│ ~ renamed:  atk → attack    │
│ ~ retyped:  price int→float │
└──────────┬───────────────────┘
           │
     ┌─────┴─────┐
     ▼           ▼
  安全变更     危险变更
 (新增/改顺序) (删除/重命名/改类型)
     │           │
     ▼           ▼
  自动处理    弹窗要求策划/程序确认
     │           │
     └─────┬─────┘
           ▼
  ┌─────────────────────┐
  │ 确认后执行迁移       │
  │ 1. 更新 schema 快照  │
  │ 2. 重新生成 C# 类    │
  │ 3. 重新导出 .asset   │
  │ 4. 刷新 Unity        │
  └─────────────────────┘
```

### 13.3 Partial Class 防止代码覆盖

自动生成的 C# 类永远标记为 `partial`，**程序手写的业务逻辑放在独立的 partial 文件中**：

```csharp
// ====== 自动生成: WeaponRow.generated.cs （会被覆盖）======
public partial class WeaponRow
{
    public int id;
    public string name;
    public int attack;
}

// ====== 手写: WeaponRow.cs （不会被覆盖）======
public partial class WeaponRow
{
    // 业务计算属性
    public bool IsLegendary => quality >= 5;

    // 关联查询方法
    public SkillRow GetPrimarySkill()
    {
        return skills.Length > 0
            ? DataManager.Instance.GetSkill(skills[0])
            : null;
    }
}
```

---

## 15. Excel 版本控制 & Diff（新增）

**问题：** Excel (.xlsx) 是二进制文件，git diff 无法查看修改了什么，多人冲突无法解决。

### 14.1 解决方案分层

| 层级 | 方案 | 说明 |
|------|------|------|
| **提交前** | Git pre-commit hook 自动导出一份 JSON 副本 | 每次 commit Excel 时自动生成 JSON 放到旁边，JSON 可 diff |
| **代码审查** | PR 中同时包含 Excel + 其 JSON 导出 | reviewer 看 JSON diff 就能知道策划改了什么 |
| **冲突解决** | Editor Window 内置 "Excel Diff" 工具 | 逐 Sheet 逐行对比两个版本（通过都导出为表格对比） |
| **历史追溯** | CI 中生成数据变更日志 | 每次合并后自动生成 `CHANGELOG_DATA.md` |

### 14.2 内置 Diff 工具

```
┌─────────────────────────────────────────────────┐
│  Excel Diff: Item.xlsx (Weapon)                  │
│  Left: HEAD~1          Right: Working Copy       │
├─────────────────────────────────────────────────┤
│  ┌──────────┬──────────┬──────────┬──────────┐  │
│  │ Status   │ id       │ name     │ attack   │  │
│  ├──────────┼──────────┼──────────┼──────────┤  │
│  │ =        │ 1001     │ 新手剑   │ 12       │  │
│  │ ~ MOD    │ 1002     │ 铁剑     │ 25 → 30  │  │ ← 变更高亮
│  │ + ADD    │ 1003     │ 秘银剑   │ 45       │  │ ← 新增行
│  │ - DEL    │ 1004     │ 木剑     │ 5        │  │ ← 删除行
│  └──────────┴──────────┴──────────┴──────────┘  │
│                                                   │
│  变更: 1 新增, 1 修改, 1 删除                      │
│  [接受所有变更] [拒绝所有变更] [逐行选择]           │
└─────────────────────────────────────────────────┘
```

---

## 16. 热更新 & Addressables 集成（新增）

**问题：** 手游上线后需要更新数据（新增道具、调整数值），不能发新版本。

早期的 "所有数据装进 ScriptableObject" 方案有致命缺陷：SO 资源打包进 APK/IPA 后无法替换。

### 15.1 双模式架构

```
             游戏启动
                 │
                 ▼
        ┌────────────────┐
        │ DataManager    │
        │ .Initialize()  │
        └───────┬────────┘
                │
    ┌───────────┴───────────┐
    ▼                       ▼
开发模式 (Editor/内建)    生产模式 (有热更新)
    │                       │
    ▼                       ▼
ScriptableObject        先从本地缓存加载
.asset 文件             │
                        ▼
                   检查服务器是否有新版本
                        │
               ┌────────┴────────┐
               ▼                 ▼
            有更新              无更新
               │                 │
               ▼                 ▼
         下载新 JSON       使用缓存 JSON
         存入本地              │
               │                 │
               └────────┬────────┘
                        ▼
              反序列化为 Row 对象
              构建 Dictionary 缓存
                        │
                        ▼
              游戏正常使用数据
```

### 15.2 导出产物

```
打包前导出:
  Assets/Data/
  ├── Weapon.asset          # ScriptableObject (内建兜底)
  ├── Weapon.json           # JSON (上传 CDN 用于热更)
  └── Weapon.json.hash      # MD5 哈希 (版本比较用)

CDN 上的目录结构:
  https://cdn.example.com/game-data/v1.0.0/
  ├── version.json             # { "Weapon": "a1b2c3", "Skill": "d4e5f6" }
  ├── Weapon.a1b2c3.json       # 内容 hash 命名，天然去重
  └── Skill.d4e5f6.json
```

### 15.3 热更新加载代码

```csharp
public class DataManager : MonoBehaviour
{
    public WeaponTable weaponTable;  // SO 兜底

    private Dictionary<int, WeaponRow> weaponLookup;

    public async Task InitializeAsync()
    {
        // 1. 先从 SO 加载兜底数据
        weaponTable?.BuildCache();

        // 2. 检查 CDN 是否有更新
        var updater = new DataUpdater("https://cdn.example.com/game-data/");
        var updates = await updater.CheckForUpdatesAsync();

        foreach (var update in updates)
        {
            // 3. 下载最新 JSON → 反序列化 → 替换缓存
            var json = await updater.DownloadAsync(update);
            var rows = JsonUtility.FromJson<WeaponRowList>(json);
            BuildLookupFromList(rows.data);
        }
    }
}
```

### 15.4 Addressables 集成

生成的 ScriptableObject 自动加入 Addressables Group：

```csharp
#if UNITY_EDITOR
// 导出时自动注册到 Addressables
var settings = AddressableAssetSettingsDefaultObject.Settings;
var group = settings.FindGroup("GameData") ?? settings.CreateGroup("GameData", ...);
var guid = AssetDatabase.AssetPathToGUID("Assets/Data/Weapon.asset");
var entry = settings.CreateOrMoveEntry(guid, group);
entry.address = "Data/Weapon";  // 运行时加载 key
#endif
```

运行时通过 Addressables 加载：

```csharp
var handle = Addressables.LoadAssetAsync<WeaponTable>("Data/Weapon");
await handle.Task;
weaponTable = handle.Result;
```

---

## 17. 大表性能 & 分页（新增）

**问题：** 10000+ 行的 ScriptableObject .asset 在 Unity Inspector 中展开会卡死。

### 16.1 解决方案

| 策略 | 说明 |
|------|------|
| **单表行数上限** | 超过 5000 行自动切换为 JSON 模式（SO 仅存引用+元数据） |
| **分页预览** | Editor Window 中数据预览按 100 行/页分页，不一次性渲染 |
| **虚拟滚动** | 使用 Unity UI Toolkit 的 `ListView.virtualizationMethod` 实现 |
| **搜索优先** | 按 ID 搜索是 O(1) 的，不需要遍历全表 |
| **按需加载** | 运行时按 ID 查询，不从 SO 反序列化所有数据到内存 |

---

## 18. 跨列条件校验（新增）

**问题：** 当前校验只支持单列（range/regex/required），无法表达 "如果 type=1，则 attack 必须 > 100" 这种业务规则。

### 17.1 条件校验表达式

```csharp
// 在 Excel 的 #Rules Sheet 中定义：

| field   | rule      | condition         | params     | message              |
|---------|-----------|-------------------|------------|----------------------|
| attack  | range     |                   | 0~9999     | 攻击力超出范围        |
| heal    | required  | type=2            |            | 治疗类物品必须填治疗量 |
| element | enum      | rarity>=3         | Element    | 稀有以上必须填元素    |
```

**条件语法：** `column op value`
- `type=2` — 该行的 type 列等于 2
- `rarity>=3` — 该行的 rarity 列大于等于 3
- `category!=consumable` — 该行的 category 列不等于 "consumable"

### 17.2 实现

```csharp
public class ConditionalRule
{
    public string TargetField;       // "heal"
    public string RuleName;          // "required"
    public string ConditionColumn;   // "type"
    public string ConditionOp;       // "="
    public string ConditionValue;    // "2"
    public string RuleParams;        // (额外参数)

    public bool ShouldApply(Dictionary<string, object> row)
    {
        // 没有 condition → 始终应用
        if (string.IsNullOrEmpty(ConditionColumn)) return true;

        var actualValue = row[ConditionColumn]?.ToString();
        return ConditionOp switch
        {
            "="  => actualValue == ConditionValue,
            "!=" => actualValue != ConditionValue,
            ">=" => float.TryParse(actualValue, out var av) &&
                     float.TryParse(ConditionValue, out var cv) && av >= cv,
            // ... 更多操作符
            _ => true
        };
    }
}
```

---

## 19. 遗漏补充清单

下面这些是在 v2 设计中未涉及但实际开发中会遇到的问题，以及应对方案。

### 18.1 程序手写业务代码防覆盖 — 已解决

见 §14.3（partial class）。

### 18.2 多 Excel 目录支持

一个项目中 Excel 可能分散在不同目录（`Config/`、`Quest/`、`Item/`），设置面板中应允许添加多个监控目录。

### 18.3 键盘快捷键

| 快捷键 | 操作 |
|--------|------|
| `Ctrl+E` | 导出选中文件 |
| `Ctrl+Shift+E` | 导出全部 |
| `Ctrl+R` | 刷新文件列表 |
| `Ctrl+V` | 仅校验不导出 |
| `F5` | 执行当前 Sheet 的重新导出 |

### 18.4 Assembly Definition 隔离

```csharp
// Editor 代码只在 Editor 下编译
// Plugins/ExcelToJsonPlugin/Editor/ExcelToJsonPlugin.Editor.asmdef
{
  "name": "ExcelToJsonPlugin.Editor",
  "references": ["ExcelToJsonPlugin.Runtime"],
  "includePlatforms": ["Editor"],
  "excludePlatforms": []
}

// Runtime 代码可进 Build（仅包含 Attributes + BaseDataTable）
// Plugins/ExcelToJsonPlugin/Runtime/ExcelToJsonPlugin.Runtime.asmdef
{
  "name": "ExcelToJsonPlugin.Runtime",
  "references": [],
  "includePlatforms": [],
  "excludePlatforms": []
}

// NPOI 仅在 Editor 中引用，不进 Build
```

### 18.5 CSV 作为轻量备选格式

对于偏好纯文本的小团队，支持 CSV 输入：

```csv
# Item_Weapon.csv
"id","name","attack","hp"
1001,"新手剑",12,100
1002,"铁剑",25,0
```

CSV 优点：git 可 diff、可 merge、纯文本、编辑器无关。

### 18.6 Excel 公式警告

NPOI 能读取公式字符串，但不能像 Excel 一样实时求值。如果在数据区域检测到公式：

1. 读取 Excel 缓存值（上次 Excel 保存时的值）
2. 如果该格子有公式 → 出 Warning："Row 5, attack 包含公式 =A4*2，导出的是缓存值 24。请在 Excel 中按 Ctrl+S 刷新缓存。"

### 18.7 Unity 版本兼容策略

| Unity 版本 | .NET 版本 | NPOI 兼容 | 支持计划 |
|------------|-----------|-----------|----------|
| 2020.3 LTS | .NET Standard 2.0 | ⚠️ 需测试 | 尽力兼容 |
| 2021.3 LTS | .NET Standard 2.1 | ✅ | 完整支持 |
| 2022.3 LTS | .NET Standard 2.1 | ✅ | 完整支持 |
| Unity 6 | .NET Standard 2.1 / CoreCLR | ✅ | 完整支持 |

### 18.8 编辑器下即时预览

在 Editor Window 中修改数据后，直接写入 Excel 并重新导出：

```
编辑器中点击单元格 → 修改值 → 按 Enter → 插件调用 NPOI 写入 Excel → 自动重新导出
```

实现策划可以在 Unity 中修改数据，不一定要打开 Excel。但这要求 NPOI 写入功能（NPOI 支持）。

---

## 附录 C：更新后的技术决策汇总

| 决策项 | v2 错误决策 | v3 修正决策 | 原因 |
|--------|------------|------------|------|
| Excel 库 | ClosedXML | **NPOI** | ClosedXML 依赖 System.Drawing，Unity 不可用 |
| 生成类类型 | `class` | **`partial class`** | 防止覆盖程序员手写的业务方法 |
| 输出格式 | 仅 ScriptableObject | **SO + JSON 双输出** | SO 用于内建兜底，JSON 用于 CDN 热更新 |
| 校验能力 | 单列校验 | **单列 + 跨列条件校验** | 实际业务需要 "if type=2 then required" |
| 大表策略 | 无 | **5000 行自动切 JSON** | 防止 SO 资产过大导致 Editor 卡顿 |
| 版本控制 | 无 | **内置 Diff + pre-commit JSON 导出** | Excel 二进制无法 git diff |
| Schema 变更 | 无 | **自动检测 + 分级处理 + 确认弹窗** | 策划频繁改表头会破坏代码 |
| Addressables | 无 | **自动注册 + 分组** | 现代 Unity 项目标配 |
| asmdef 隔离 | 无 | **Editor/Runtime 分 assembly** | 防止 NPOI DLL 打入游戏包 |
| NPOI 版本 | 2.7.0 netstandard2.0 | **NPOI 2.5.6 net45** | 实测 2.7.0 DLL 在 Unity 中加载失败（缺 System.Buffers） |
| NPOI 依赖 | 只有 NPOI | **+ BouncyCastle.Crypto** | NPOI 硬依赖，否则 DLL 加载报 `Unable to resolve reference` |
| API 兼容性 | .NET Standard 2.1 | **.NET Framework** | NPOI net45 需要 .NET Framework，否则 DLL 加载失败 |
| Pipeline 两段式 | 一次 Pass 完成 | **Code Gen → Unity 编译 → Asset Gen** | 首次导出的 C# 类未编译完，反射找不到 |

---

## 附录 D：Sprint 1 实测记录（2026-05-08，最终版）

**环境：** Unity 2022.3.62f3c1, API `.NET Framework`, NPOI 2.5.6 net45 + BouncyCastle.Crypto 1.8.9

**产出：** 17 个 C# 源文件 + 5 个 DLL + 11 个测试 Excel

### D.1 类型覆盖测试（15/15 全通过）

| # | 类型 | 测试文件 | 边界覆盖 | 结果 |
|---|------|----------|----------|:---:|
| 1 | `int` | Test11 | 0, max(2147483647), min(-2147483648) | ✅ |
| 2 | `float` | Test11 | 0.000001, 999999.999, 整型浮点(100), -0.5 | ✅ |
| 3 | `string` | Test01 | 空值, 中文, 英文, 特殊字符(!@#$%) | ✅ |
| 4 | `bool` | Test01,07 | true/false, 0/1, yes/no, Y/N, 是/否 | ✅ |
| 5 | `int[]` | Test02,10 | 负数, 大值, 单元素, `[]`, JSON格式 | ✅ |
| 6 | `float[]` | Test02,10 | 管道分隔(`1.5\|2.0\|3.14`), 小数 | ✅ |
| 7 | `string[]` | Test08 | 中文, 英文, `[]`, 单元素 | ✅ |
| 8 | `Vector2` | Test03 | `[1,2]`, `[0.5,1.5]`, `[-1,0]` | ✅ |
| 9 | `Vector3` | Test03 | `[1,2,3]`, `[0,0,0]`, `[100,200,300]` | ✅ |
| 10 | `Color` | Test03 | Hex `#FF0000`, Hex+Alpha `#00FF00FF`, RGBA | ✅ |
| 11 | `enum:X` | Test04 | `enum:Quality` → int 存储 | ✅ |
| 12 | `ref:X` | Test04, Char+Skill | `ref:Skill`, `int[]` 批量引用 | ✅ |
| 13 | `res` | Test04,09 | Sprite/GameObject/AudioClip/Texture2D/Material, 纯res无子类型 | ✅ |
| 14 | `json` | Test08 | 嵌套`{"require":{"lv":50}}`, `{}`, 空 | ✅ |
| 15 | `loc` | Test08 | `ITEM_SWORD_001`, `QUEST_MAIN_001` | ✅ |

### D.2 功能测试

| # | 用例 | 行数 | 结果 | 说明 |
|---|------|------|:---:|------|
| 1 | Item.xlsx (Weapon+Armor) | 8 | ✅ | 基准：2 Sheet，9 列，含中文 |
| 2 | Character.xlsx + Skill.xlsx | 4+6 | ✅ | 跨表引用：`ref:Skill` + `int[]` 关联 |
| 3 | Test05_Large | 200 | ✅ | 200 行 0.1s，无性能退化 |
| 4 | Test06_SkipSheets | 4 | ✅ | `_Notes` 和 `#Internal` 正确跳过 |
| 5 | Test11_NumEdge | 2 | ✅ | 整型极值、浮点边界、零值 | 

**总计：11 个 Excel，14 个 Sheet，252 行数据，0 个类型转换错误。**

### D.3 已知限制

| 限制 | 影响 | 应对 |
|------|------|------|
| 首次导出需运行两次 | 首先生成 .cs → Unity 编译 → 再导出生成 .asset | CLI 中自动处理；Editor Window 中引导用户点两次 |
| NPOI 公式只能读缓存值 | 含公式的单元格导出的是上次 Excel 保存时的值 | 导出时检测公式并 Warning |
| 空单元格存储格式 | Excel 中空单元格若为数值型则 NPOI 读成 0 | 建议策划在模板中预设正确的单元格格式 |
| C# 关键字冲突 | 字段名若为 `internal`/`class` 等关键字会编译失败 | 已加关键字检测，自动加 `_` 前缀 |

### D.4 开发中发现的关键决策变更

| 决策 | 设计文档 v3 | 实测结论 |
|------|------------|----------|
| NPOI 版本 | 2.7.0 | **2.5.6** — 2.7.0 DLL 在 Unity 中无法加载 |
| 目标框架 | netstandard2.0 | **net45** — Unity Editor 运行在 .NET Framework 4.8 |
| 额外依赖 | 无 | **BouncyCastle.Crypto 1.8.9** — NPOI 硬依赖 |
| API 兼容 | .NET Standard 2.1 | **.NET Framework** — 非此模式 NPOI DLL 加载失败 |
| Pipeline | 单次 Pass | **两段式** — Code Gen 后需等 Unity 编译再 Asset Gen |
