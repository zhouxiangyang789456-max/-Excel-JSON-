# Excel To JSON — Unity 编辑器插件

[![中文文档](https://img.shields.io/badge/README-%E4%B8%AD%E6%96%87-blue?style=for-the-badge)](README.md)
[![English](https://img.shields.io/badge/README-English-2ea44f?style=for-the-badge)](README_en.md)

将 Excel 表格转换为强类型 ScriptableObject 资产，支持数据校验、代码生成、一键导出。

## 功能特性

- **Excel → ScriptableObject 流水线**：把 `.xlsx` 文件放入 `Assets/Excel/`，导出到 `Assets/Data/` 生成 `.asset` 文件
- **16 条内置校验规则**：结构校验、数据完整性、类型匹配、跨表引用
- **三种映射模式**：自动生成 C#（模式 A）、C# 反射匹配（模式 B）、混合模式（模式 C）
- **文件监听**：Excel 文件变化自动导出（防抖可配置）
- **增量导出**：仅重新处理变更的 Sheet
- **数据预览**：分页表格视图，支持搜索和按类型着色表头
- **运行时 API**：类型安全的 `DataManager`，支持 Get / Find / GetRandom / GetByIds

## 安装

### Git URL 安装（推荐）
1. 打开 Unity，进入 **Window > Package Manager**
2. 点击 **+** 按钮，选择 **"Add package from git URL"**
3. 输入：`https://github.com/zhouxiangyang789456-max/-Excel-JSON-.git`
4. 点击 **Add** — 插件会出现在 `Packages/com.github.excel-to-json`

### 本地安装
1. 克隆或下载本仓库
2. 在 Package Manager 中点击 **+** → **"Add package from disk"**
3. 选择仓库根目录下的 `package.json` 文件

### 导入示例（可选）
1. 在 Package Manager 中选中 **"Excel To JSON"**
2. 切换到 **Samples** 标签页
3. 点击 **Import** 导入 "Demo Data Usage"

## 环境要求

- **Unity 2021.3 LTS** 及以上
- **API Compatibility Level**：必须设置为 **.NET Framework**（Edit → Project Settings → Player → Other Settings → API Compatibility Level）
- 无需其他外部依赖（NPOI 已作为 DLL 内置）

## 快速开始

1. 将 `.xlsx` 文件放入项目的 `Assets/Excel/` 文件夹
2. 打开 **Window > Excel Data Manager**
3. 在文件树中选择一个 Sheet
4. 点击 **导出**（工具栏或导出标签页）
5. 代码生成到 `Assets/Scripts/Generated/`，资产生成到 `Assets/Data/`

## 支持的类型

| Excel 类型 | C# 类型 | 示例 |
|-----------|---------|------|
| `int` | `int` | `42` |
| `float` | `float` | `3.14` |
| `string` | `string` | `hello` |
| `bool` | `bool` | `true` / `false` / `是` / `否` |
| `int[]` | `int[]` | `1\|2\|3` 或 `[1,2,3]` |
| `float[]` | `float[]` | `1.0\|2.5\|3.0` |
| `string[]` | `string[]` | `a\|b\|c` |
| `Vector2` | `Vector2` | `[1,2]` |
| `Vector3` | `Vector3` | `[1,2,3]` |
| `Color` | `Color` | `#FF0000` 或 `[1,0,0,1]` |
| `ref:表名` | `int` | `1001`（外键引用） |
| `enum:表名` | `int` | `1`（枚举值） |
| `res:Sprite` | `string` | `Sprites/icon` |
| `json` | `string` | `{"key":"value"}` |
| `loc` | `string` | 本地化 key |

## Excel 格式

每个数据 Sheet 有 4 行表头：
- **第 1 行**：字段名（如 `id`、`name`、`attack`）
- **第 2 行**：类型声明（如 `int`、`string`、`ref:Weapon`）
- **第 3 行**：注释（可选，会作为生成代码的 Tooltip）
- **第 4 行起**：数据行

以 `_` 或 `#` 开头的 Sheet 会被跳过。`#Rules` Sheet 用于定义自定义校验规则。

## 映射模式

### 模式 A：Excel 驱动（自动代码生成）
- 策划在 Excel 中定义字段
- 插件自动生成 C# Row/Table 类
- 适合原型开发和新项目

### 模式 B：C# 驱动（反射匹配）
```csharp
[ExcelTable("Weapon")]
public class WeaponRow
{
    [ExcelColumn("id")]
    public int Id;

    [ExcelColumn("name")]
    public string Name;

    [ExcelColumn("attack")]
    public int Attack;

    [ExcelIgnore]
    public int ComputedValue; // 不在 Excel 中
}
```
- 程序员定义带属性的 C# 类
- 插件自动匹配 Excel 列到 C# 字段
- 适合已有项目和复杂类型

### 模式 C：混合模式
每个 Sheet 独立选择模式 A 或 B，在 Settings 中按 Sheet 配置。

## 校验规则

### 阶段 1 — 结构校验
- 字段名唯一性
- 类型声明有效性
- 表头完整性
- 字段名合法性（无特殊字符）

### 阶段 2 — 数据校验
- ID 必须存在且唯一
- 类型匹配（值 vs 声明类型）
- 必填字段检查
- 枚举格式校验
- 资源路径格式校验
- 公式检测

### 阶段 3 — 跨表引用校验
- ref 完整性：验证引用的 ID 在目标表中是否存在
- enum 存在性：验证枚举值在目标枚举表中是否存在

### 自定义规则（#Rules Sheet）
在 Excel 中直接定义：`range`、`regex`、`multiple`、`not_empty`、`required`、`enum`

## 运行时 API

```csharp
var dm = DataManager.Instance;

// 按类型获取表
var weaponTable = dm.GetTable<WeaponTable>();

// 按 ID 查询（O(1)）
var weapon = weaponTable.Get(1001);

// 获取所有行
foreach (var w in weaponTable.GetAll()) { ... }

// 条件筛选
var strong = weaponTable.Find(w => w.Attack > 30);

// 随机获取
var random = weaponTable.GetRandom(w => w.Rarity >= 3);
```

## 编辑器窗口

- **仪表盘（标签页 0）**：项目概览、快速状态、导出按钮
- **数据预览（标签页 1）**：分页表格视图，支持搜索
- **校验结果（标签页 2）**：按颜色图标显示错误列表，支持 CSV 导出
- **导出（标签页 3）**：导出配置、模板生成、已生成文件列表
- **运行时 API（标签页 4）**：API 参考、代码示例、DataManager 生成

## 许可证

MIT License — 个人和商业用途免费。
