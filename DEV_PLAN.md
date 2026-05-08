# Excel To JSON 插件 — 开发计划

> 基于 DESIGN.md v3，拆分可执行的开发展望。每轮 Sprint 有明确的交付物和验收标准。

---

## 总览

| 项目 | 说明 |
|------|------|
| 目标 | 一个上架 Unity Asset Store 的 Excel 转 ScriptableObject / JSON 编辑器插件 |
| 语言 | 纯 C#（Editor + Runtime），零外部进程依赖 |
| Excel 库 | NPOI（Apache 2.0，纯 C#，无 System.Drawing 依赖） |
| 最低 Unity | 2021.3 LTS |
| 总工期 | **10 周**（Sprint × 5，每轮 2 周） |
| 开发人数 | 1 人 |

---

## Sprint 1：核心流水线（第 1–2 周）

> 目标：一个 Excel 文件进去，一个 .asset 文件出来。没有 UI，没有校验。

### 1.1 环境搭建

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 1 | 创建 Unity 项目，建立目录结构 | `Assets/Plugins/ExcelToJsonPlugin/` 骨架 | 1h |
| 2 | 通过 UPM 或手动导入 NPOI 2.x DLL | `Dependencies/NPOI.dll` 引用就绪 | 1h |
| 3 | 创建 Editor + Runtime 两套 asmdef | `ExcelToJsonPlugin.Editor.asmdef`、`ExcelToJsonPlugin.Runtime.asmdef` | 0.5h |
| 4 | Runtime 放 Attributes（`[ExcelTable]`、`[ExcelColumn]`） | `Attributes.cs` | 0.5h |

**验收：** Unity 编译通过，NPOI 在 Editor 下可正常 `using NPOI.XSSF.UserModel`。

### 1.2 ExcelReader

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 5 | 实现 `ExcelReader.Open(string path)`，返回 `IWorkbook` | `ExcelReader.cs` | 2h |
| 6 | 处理 .xlsx（XSSFWorkbook）和 .xls（HSSFWorkbook）两种格式 | 同上 | 1h |
| 7 | 处理合并单元格：取左上角值填充到所有覆盖格 | 内置逻辑 | 2h |
| 8 | 跳过空行、跳过隐藏行（可选开关） | 内置逻辑 | 1h |
| 9 | 单元测试：准备 3 个测试 Excel（正常 / 含合并格 / 含空行） | `Tests/Editor/ExcelReaderTests.cs` | 2h |

**验收：** `ExcelReader.Read("Item.xlsx")` 返回 `{ sheetName: [[cell], ...] }` 结构化数据。

### 1.3 SchemaParser

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 10 | 根据配置的行号解析字段名行 | `SchemaParser.cs` | 1h |
| 11 | 解析类型声明行 | 同上 | 1h |
| 12 | 解析注释行（可选，作为 Tooltip） | 同上 | 0.5h |
| 13 | 以 `_` 开头的 Sheet 自动跳过 | 同上 | 0.5h |
| 14 | 输出 `TableSchema { string TableName, List<FieldDef> Fields }` | `TableSchema.cs` 数据模型 | 1h |
| 15 | 单元测试：不同表头行号配置下的解析 | `Tests/Editor/SchemaParserTests.cs` | 1.5h |

**验收：** 正确识别字段名、类型、注释行。`_` Sheet 不输出。

### 1.4 DataParser + TypeMapper

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 16 | 实现基础类型转换：int / float / string / bool | `TypeMapper.cs` | 3h |
| 17 | 实现数组类型转换：int[] / float[] / string[]（支持 JSON 数组和 `\|` 分隔符） | 同上 | 2h |
| 18 | 空值处理：各类型的空值默认行为 | 同上 | 1h |
| 19 | 按行解析数据，输出 `List<Dictionary<string, object>>` | `DataParser.cs` | 2h |
| 20 | 单元测试：各种类型 + 边界值 + 空值 | `Tests/Editor/DataParserTests.cs` | 2h |

**验收：** 所有支持的基础类型都能正确转换，空值有合理的默认值。

### 1.5 CodeGenerator（模式 A）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 21 | 从 `TableSchema` 生成 `Row` 类 C# 源码字符串 | `CodeGenerator.cs` | 4h |
| 22 | 生成 `Table` 容器类（继承 `BaseDataTable` / `ScriptableObject`） | 同上 | 3h |
| 23 | 支持 partial class、[Serializable]、[Tooltip] | 同上 | 1h |
| 24 | 可配置命名空间、类名后缀、生成路径 | 配置绑定 | 1h |
| 25 | 生成 `BaseDataTable.cs`（抽象基类，如果不存在则创建） | 同上 | 1h |
| 26 | 单元测试：验证生成的代码能通过 C# 编译 | `Tests/Editor/CodeGeneratorTests.cs` | 2h |

**验收：** 给一个 Excel Schema，能生成可编译的 `.cs` 文件。

### 1.6 AssetGenerator

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 27 | 从解析后的数据创建 Row 对象列表 | `AssetGenerator.cs` | 2h |
| 28 | 创建 ScriptableObject 实例，填充 rows → 写入 .asset | 同上 | 3h |
| 29 | 写入前检查是否已存在同名资产（覆盖 or 跳过可选） | 同上 | 0.5h |
| 30 | 调用 `AssetDatabase.Refresh()` + `AssetDatabase.SaveAssets()` | 同上 | 0.5h |
| 31 | 集成测试：全流程 Excel → .asset | `Tests/Editor/PipelineIntegrationTests.cs` | 2h |

**验收：** Excel 放在 `Assets/Excel/` → 运行流水线 → `Assets/Data/Weapon.asset` 出现，Inspector 中可查看。

### 1.7 Pipeline 调度器

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 32 | 串联 Reader → Parser → Mapper → Generator → 输出 | `Pipeline.cs` | 2h |
| 33 | 单文件处理入口 `Pipeline.ProcessFile(string excelPath)` | 同上 | 1h |
| 34 | 目录批量处理 `Pipeline.ProcessDirectory(string dir)` | 同上 | 1h |

**验收：** 调用 `Pipeline.ProcessFile("Assets/Excel/Item.xlsx")` 即可完成全部流程。

---

## Sprint 2：Editor Window + 校验引擎（第 3–4 周）

> 目标：有可视化面板了，策划能在 Unity 里看到数据，程序能看到错误。

### 2.1 Editor Window 骨架

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 35 | 创建 `ExcelDataWindow.cs`，注册 `Window > Excel Data Manager` | Editor Window 可打开 | 2h |
| 36 | 左侧文件树（`ExcelTreeView`），扫描 `Assets/Excel/` 递归列出 .xlsx | `UI/ExcelTreeView.cs` | 3h |
| 37 | 文件树支持：展开 Sheet、多选、右键菜单（导出/校验/打开） | 同上 | 2h |
| 38 | 右侧内容区，用 `TabView` 切换 5 个标签页 | 主窗口布局完成 | 3h |

**验收：** 打开 Window 能看到文件树和标签页，点击文件能切换。

### 2.2 数据预览面板（Tab 2）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 39 | 选中 Sheet 后，在表格中渲染数据（前 100 行，分页） | `UI/DataPreviewGrid.cs` | 4h |
| 40 | 表头：字段名行 + 类型行，颜色区分 | 同上 | 1h |
| 41 | 搜索功能：`id=1001` 快速跳转行 | 同上 | 1.5h |
| 42 | 虚拟滚动（ListView virtualization），大表不卡 | 同上 | 2h |

**验收：** 选中 Sheet 能预览数据，搜索能快速定位。

### 2.3 导出面板（Tab 4）+ 一键导出

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 43 | 导出配置 UI：输出格式下拉、路径选择、选项勾选框 | `UI/ExportPanel.cs` | 2h |
| 44 | 导出进度条（分文件/分 Sheet 展示进度） | 同上 | 2h |
| 45 | "导出全部"按钮：调用 Pipeline 批量处理 | 同上 | 1h |
| 46 | 导出后自动刷新 AssetDatabase | 集成 | 0.5h |

**验收：** 点"导出"按钮 → 进度条 → 导出完成提示 → .asset 生成。

### 2.4 校验引擎

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 47 | `ValidationEngine` 调度器：按 Stage 1→2→3 依次执行 | `ValidationEngine.cs` | 2h |
| 48 | Stage 1 规则：字段名唯一、类型合法、表头完整、Sheet 非空 | `Rules/StructureRules.cs` | 2h |
| 49 | Stage 2 规则：类型匹配、ID 唯一、ID 非空、必填检查 | `Rules/DataRules.cs` | 3h |
| 50 | 范围校验（range）、正则校验（regex）、倍数校验（multiple） | 同上 | 2h |
| 51 | `ValidationResult` 收集错误列表（Error / Warning / Info） | 数据模型 | 1h |
| 52 | 单元测试：每个规则独立 + 组合测试 | `Tests/Editor/ValidationTests.cs` | 3h |

**验收：** Excel 中故意写错数据 → 校验引擎输出精确到单元格的错误列表。

### 2.5 校验面板（Tab 3）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 53 | 错误列表展示：文件名 / Sheet / 行号 / 列名 / 错误信息 / 级别 | `UI/ValidationPanel.cs` | 3h |
| 54 | 颜色图标区分 Error（红）/ Warning（黄）/ Info（灰） | 同上 | 1h |
| 55 | 双击错误行 → 在项目窗口中高亮对应 Excel + 定位行号提示 | 同上 | 1.5h |
| 56 | "复制错误列表"→ 剪贴板 CSV | 同上 | 0.5h |

**验收：** 校验后错误清晰展示，双击能定位到 Excel 文件。

### 2.6 运行时用法面板（Tab 5）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 57 | 显示生成文件路径、API 参考、代码示例（静态） | `UI/RuntimePanel.cs` | 2h |
| 58 | [复制代码] 按钮：把示例代码拷到剪贴板 | 同上 | 0.5h |
| 59 | [一键生成 DataManager GameObject] 按钮 | 同上 | 1.5h |

**验收：** Tab 5 能展示运行时 API，按钮功能正常。

---

## Sprint 3：模式 B + 文件监听 + 复合类型（第 5–6 周）

> 目标：支持 C# 驱动映射，支持 ref/enum 等高级类型，保存 Excel 即自动导出。

### 3.1 模式 B（C# 驱动反射匹配）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 60 | 通过反射扫描所有带 `[ExcelTable]` 的类 | `Mapping/AttributeMapping.cs` | 3h |
| 61 | 解析 `[ExcelColumn("列名")]` 标签，构建 Field→Column 映射 | 同上 | 2h |
| 62 | 类型兼容性检查（Excel 列类型 vs C# 字段类型） | 同上 | 2h |
| 63 | `[ExcelIgnore]` 标签支持：该字段不参与 Excel 匹配 | 同上 | 0.5h |
| 64 | 缺少列 / 多余列的 Warning 生成 | 同上 | 1h |
| 65 | 从已有 C# 类反向生成 Excel 模板（只有表头行，数据行留空） | `TemplateExporter.cs` | 2h |
| 66 | 集成测试：用标签定义一个类 → 填 Excel → 导出 → 字段值正确 | 测试 | 2h |

**验收：** 给一个带标签的 C# 类，能匹配 Excel 列并导出。

### 3.2 模式 C（混合匹配）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 67 | 检测逻辑：没有对应 C# 类 → 模式 A；有标签 → 模式 B；部分字段有标签 → 自动补全 | `Mapping/HybridMapping.cs` | 3h |
| 68 | 映射模式选择 UI：每个 Sheet 可单独选择 A / B / C | 设置面板 | 1.5h |

**验收：** 同一文件中 Sheet A 用模式 A，Sheet B 用模式 B，都能正常导出。

### 3.3 复合类型：ref 和 enum

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 69 | `ref:TableName` 类型：解析为 int，校验时检查目标表 ID 列是否存在 | `TypeMapper.cs` 扩展 | 3h |
| 70 | 跨文件引用：校验时需加载被引用表的数据 | 同上 | 2h |
| 71 | `enum:TableName` 类型：解析为 int，校验值在枚举 Sheet 中存在 | 同上 | 2h |
| 72 | `res:Type` 类型：解析为 string，路径格式校验 | 同上 | 2h |
| 73 | 多种类型在数据预览中的显示颜色区分 | UI 适配 | 1h |
| 74 | 单元测试：ref 引用存在、不存在、跨文件、循环引用 | 测试 | 2h |

**验收：** ref 引用能正确校验完整性，enum 值校验通过。

### 3.4 文件监听 + 自动导出

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 75 | `ExcelFileWatcher`：使用 `FileSystemWatcher` 监听 Excel 目录 | `Watcher/ExcelFileWatcher.cs` | 2h |
| 76 | Debounce 机制：500ms 内同一文件多次变更只触发一次处理 | 同上 | 1h |
| 77 | 忽略 Excel 临时文件（`~$*`、`*.tmp`） | 同上 | 0.5h |
| 78 | 变更后自动调用 Pipeline.ProcessFile → 自动刷新 AssetDatabase | 集成 | 1h |
| 79 | Console 输出变更日志 + 导出结果 | 集成 | 0.5h |
| 80 | 开关：设置面板中可启用/禁用自动导出 | 设置面板 | 0.5h |

**验收：** 策划改 Excel 按 Ctrl+S → 500ms 后 .asset 自动更新。

### 3.5 增量导出

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 81 | 导出时为每个文件/Sheet 计算 MD5 hash，存入 `.cache/hashes.json` | 增量缓存 | 2h |
| 82 | 下次导出前对比 hash，仅处理变更的 Sheet | 同上 | 1.5h |
| 83 | ref 引用链追踪：被引用表变更 → 级联重新校验引用表 | 同上 | 1.5h |

**验收：** 改 1 个 Sheet 只导 1 个 Sheet，不改的不动。

---

## Sprint 4：设置面板 + Dashboard + 上架准备（第 7–8 周）

> 目标：可配置、有总览、有 Demo、文档齐全。

### 4.1 Dashboard（Tab 1）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 84 | 总览面板：项目信息 + Excel 目录 + 输出目录 + 导出格式展示 | `UI/DashboardPanel.cs` | 2h |
| 85 | 快速状态表：每个 Excel 的行数/上次导出时间/校验状态/操作按钮 | 同上 | 3h |
| 86 | "一键导出全部" / "仅校验" / "导出选中" 快捷按钮 | 同上 | 1h |

**验收：** 打开 Tab 1 一眼看到所有表的状态。

### 4.2 设置面板

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 87 | `SettingsWindow.cs` 独立设置窗口（也可做在 Editor Window 侧边栏） | `SettingsWindow.cs` | 2h |
| 88 | Excel 解析设置：表头行号、数据起始行、ID 列、跳过前缀 | UI 控件 | 2h |
| 89 | 校验设置：启用/禁用规则、严格模式开关 | UI 控件 | 1h |
| 90 | 代码生成设置：输出路径、命名空间、类名后缀 | UI 控件 | 1h |
| 91 | 文件监听设置：启用/禁用、模式、debounce 时长 | UI 控件 | 1h |
| 92 | 配置持久化：使用 `ScriptableObject` 或 JSON 保存设置到 `ProjectSettings/` | `PluginSettings.cs` | 2h |
| 93 | 恢复默认 + 导入/导出配置按钮 | UI | 1h |

**验收：** 所有设置项可用，关闭重开设置不丢失。

### 4.3 错误处理 & 容错

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 94 | Excel 文件损坏 → 捕获异常，输出 Fatal 级别错误，跳过该文件继续其他 | Pipeline 容错 | 1.5h |
| 95 | 类型转换失败 → 捕获并生成 Error，该行跳过 | DataParser 容错 | 1h |
| 96 | 部分成功策略：10 个文件中 3 个报错，其余 7 个正常导出 | Pipeline 容错 | 1h |
| 97 | 错误重试建议 UI：导出失败后给出"打开文件定位错误"引导 | UI 提示 | 1h |

**验收：** 故意损坏一个 Excel，不影响其他文件的导出。

### 4.4 Demo + 示例

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 98 | 创建 Demo 场景 `DemoScene.unity` | Demo 目录 | 1h |
| 99 | 创建示例 Excel：`SampleItem.xlsx`（含 Weapon / Armor 两 Sheet + _备注） | Demo 目录 | 1h |
| 100 | 创建示例使用脚本 `DemoDataUsage.cs`：展示 Get / GetAll / Find 用法 | Demo 目录 | 1.5h |
| 101 | 场景运行效果：点击 UI 按钮查询数据并显示 | Demo 目录 | 2h |

**验收：** 打开 Demo 场景 → 运行 → 能看到数据从 Excel → SO → 游戏界面的完整流程。

### 4.5 文档

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 102 | 英文 README.md（GitHub 用）：安装、快速开始、截图、API 概览 | `README.md` | 3h |
| 103 | 中文使用手册：策划篇（Excel 怎么写）+ 程序篇（代码怎么调） | `docs/USAGE.md` | 3h |
| 104 | API 文档（生成或手写） | `docs/API.md` | 2h |
| 105 | CHANGELOG.md | `CHANGELOG.md` | 0.5h |

**验收：** 新人按文档能在 10 分钟内跑通 Demo。

### 4.6 Asset Store 上架准备

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 106 | 打包 .unitypackage（含 Editor + Runtime + Demo + Docs） | `ExcelToJsonPlugin.unitypackage` | 1h |
| 107 | Asset Store 页面截图（Editor Window 各 Tab + Demo 运行效果） | 5-8 张截图 | 2h |
| 108 | Asset Store 描述文案（英文） | 文案 | 1h |
| 109 | 定价策略 + 许可证说明 | 定价说明 | 0.5h |
| 110 | 提交到 Unity Asset Store Publisher Portal | 提交 | 1h |

---

## Sprint 5：扩展 & 深化（第 9–10 周）

> 目标：大型项目可用，边缘场景覆盖。

### 5.1 Schema 迁移系统（DESIGN.md §14）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 111 | Schema 快照保存：导出时自动保存 `.cache/schema_snapshots/{Table}.json` | `SchemaSnapshot.cs` | 2h |
| 112 | 变更检测器：对比当前 Excel vs 快照，输出 added / removed / renamed / retyped 差异 | `SchemaDiffer.cs` | 3h |
| 113 | 安全变更自动处理（新增列 → 旧数据用默认值填充） | 自动迁移 | 2h |
| 114 | 危险变更弹窗确认（删除列 / 重命名 / 改类型 → 阻塞导出，需确认） | UI 弹窗 | 2h |

**验收：** 改列名后导出 → 弹窗询问 → 确认后迁移成功。

### 5.2 JSON 导出 + 热更新

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 115 | JSON 导出器：与 SO 同时生成 .json + .hash 文件 | `JsonExporter.cs` | 2h |
| 116 | `DataUpdater` 运行时类：从 CDN 下载 JSON → 反序列化 → 替换缓存 | `Runtime/DataUpdater.cs` | 3h |
| 117 | 版本清单 `version.json` 生成逻辑 | 导出器扩展 | 1h |

**验收：** 导出产生 SO + JSON 双文件，运行时能从 CDN 热更。

### 5.3 跨列条件校验（DESIGN.md §18）

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 118 | 条件解析器：`column op value` 语法（type=2 / rarity>=3） | `ConditionParser.cs` | 2h |
| 119 | 条件校验规则执行：先判断条件是否满足，再执行校验 | `Rules/ConditionalRule.cs` | 2h |
| 120 | `#Rules` Sheet 中支持条件列 | Excel 解析扩展 | 1h |

**验收：** 在 `#Rules` Sheet 中写 `heal / required / type=2` → 条件满足时才校验。

### 5.4 CLI 独立版本

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 121 | 创建 .NET Console 项目，引用共享的 Core 库 | `ExcelToJson.Cli/` | 2h |
| 122 | 实现 `export` / `validate` / `inspect` 子命令 | `Program.cs` | 3h |
| 123 | `--exit-code-on-error` 开关 | CLI 参数 | 0.5h |
| 124 | 打包脚本（win-x64 / osx-x64 / linux-x64 self-contained） | `build.ps1` / `build.sh` | 1h |

**验收：** `ExcelToJson.Cli validate -i ./Excel --exit-code-on-error` 在 CI 中可用。

### 5.5 遗漏补充

| # | 任务 | 产出 | 估时 |
|---|------|------|------|
| 125 | 多 Excel 目录支持：设置面板中可添加多个监控路径 | 设置扩展 | 1.5h |
| 126 | 键盘快捷键：Ctrl+E 导出、Ctrl+R 刷新、Ctrl+V 校验 | 快捷键注册 | 1h |
| 127 | 确保 NPOI DLL 不进入 Build（asmdef exclude + Editor 目录隔离确认） | Build 测试 | 1h |
| 128 | CSV 备选导入支持 | `CsvReader.cs` | 1.5h |
| 129 | Excel 公式单元格检测 + Warning 生成 | 校验扩展 | 1h |
| 130 | Editor 下运行时热重载：Ctrl+Shift+R 重新 BuildCache | `Runtime/HotReload.cs` | 1.5h |

---

## 里程碑

| 里程碑 | 时间 | 判定标准 |
|--------|------|----------|
| **M1: 流水线跑通** | Sprint 1 结束 | Excel → .asset 生成成功，单元测试全绿 |
| **M2: 有界面可用** | Sprint 2 结束 | Editor Window 5 个 Tab 功能完整，校验通过/报错 |
| **M3: 策划可日常用** | Sprint 3 结束 | 模式 A+B+C 可用，复合类型 OK，文件监听自动导出 OK |
| **M4: 可上架** | Sprint 4 结束 | Demo 可演示，文档齐全，.unitypackage 已打包 |
| **M5: 生产可用** | Sprint 5 结束 | Schema 迁移、JSON 热更、CLI、大表优化全部就绪 |

---

## 风险 & 缓解

| 风险 | 概率 | 缓解 |
|------|------|------|
| NPOI API 在特定 .xlsx 格式下异常（加密/保护/数据透视表） | 中 | Sprint 1 即覆盖 5+ 种 Excel 变体测试 |
| Unity Editor UI Toolkit 在 2021.3 上功能不全 | 低 | 降级使用 IMGUI（`EditorGUILayout`），UI Toolkit 仅 2022+ 启用 |
| Asset Store 审核周期长（2-4 周） | 高 | Sprint 4 即提交，审核期间继续 Sprint 5 开发 |
| NPOI 大文件（100MB+）解析慢 | 低 | 流式读取 + 分帧处理 + 增量导出 |
| 策划实际 Excel 格式与模板差异大 | 中 | 设置面板暴露所有表头行号配置，不强约束行布局 |

---

## 总工时估算

| Sprint | 内容 | 估时 |
|--------|------|------|
| Sprint 1 | 核心流水线 | 55h |
| Sprint 2 | Editor Window + 校验 | 48h |
| Sprint 3 | 模式 B + 监听 + 复合类型 | 40h |
| Sprint 4 | 设置 + Dashboard + Demo + 文档 + 上架 | 38h |
| Sprint 5 | 扩展 & 深化 | 30h |
| **合计** | | **211h** ≈ 10 周（1 人全职） |

---

## 附录：文件创建清单（Sprint 1）

按开发顺序排列的关键文件（Sprint 1）：

```
Assets/Plugins/ExcelToJsonPlugin/
├── Dependencies/
│   └── NPOI.dll                           [步骤 2]
├── Editor/
│   ├── ExcelToJsonPlugin.Editor.asmdef    [步骤 3]
│   ├── Core/
│   │   ├── ExcelReader.cs                 [步骤 5]
│   │   ├── SchemaParser.cs                [步骤 10]
│   │   ├── DataParser.cs                  [步骤 19]
│   │   ├── TypeMapper.cs                  [步骤 16]
│   │   ├── Pipeline.cs                    [步骤 32]
│   │   └── Models/
│   │       └── TableSchema.cs             [步骤 14]
│   ├── Generator/
│   │   ├── CodeGenerator.cs               [步骤 21]
│   │   └── AssetGenerator.cs              [步骤 27]
│   └── UI/                               (Sprint 2 开始)
├── Runtime/
│   ├── ExcelToJsonPlugin.Runtime.asmdef   [步骤 3]
│   ├── Attributes.cs                      [步骤 4]
│   ├── BaseDataTable.cs                   [步骤 25]
│   └── DataManager.cs                     [步骤 59]
└── Tests/
    └── Editor/
        ├── ExcelReaderTests.cs            [步骤 9]
        ├── SchemaParserTests.cs           [步骤 15]
        ├── DataParserTests.cs             [步骤 20]
        ├── CodeGeneratorTests.cs          [步骤 26]
        └── PipelineIntegrationTests.cs    [步骤 31]
```
