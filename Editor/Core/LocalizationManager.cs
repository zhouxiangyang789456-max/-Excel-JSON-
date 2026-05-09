using System.Collections.Generic;
using UnityEditor;

namespace ExcelToJsonPlugin.Editor.Core
{
    /// <summary>
    /// Simple localization for the editor UI.
    /// Stores language preference in EditorPrefs. Defaults to Chinese.
    /// Usage: Loc.Tr("key") → returns localized string.
    /// </summary>
    public static class Loc
    {
        private static readonly Dictionary<string, string> ZhDict = new Dictionary<string, string>
        {
            // ===== Main Window =====
            ["window_title"] = "Excel 数据管理器",
            ["window_menu"] = "Window/Excel 数据管理器",
            ["refresh"] = "刷新",
            ["export_all"] = "全部导出",
            ["validate"] = "校验",
            ["template"] = "模板",
            ["settings"] = "设置",
            ["ready"] = "就绪",

            // ===== Toolbar =====
            ["toolbar_refresh"] = "刷新",
            ["toolbar_export_all"] = "全部导出",
            ["toolbar_validate"] = "校验",
            ["toolbar_template"] = "模板",

            // ===== File tree =====
            ["files"] = "文件列表",
            ["search_filter"] = "搜索...",
            ["filter"] = "筛选",
            ["selected_count"] = "已选 {0} 个",
            ["export_selected"] = "导出选中",
            ["validate_selected"] = "校验选中",
            ["clear"] = "清除",
            ["add_excel_dir"] = "+ 添加 Excel 目录",
            ["export_this_file"] = "导出此文件",
            ["validate_this_file"] = "校验此文件",
            ["open_excel"] = "打开 Excel",
            ["add_to_selection"] = "加入选择",
            ["export_this_sheet"] = "导出此表",
            ["validate_this_sheet"] = "校验此表",

            // ===== Sheet info bar =====
            ["mapping_mode"] = "映射模式:",
            ["mode_auto"] = "自动",
            ["mode_a"] = "模式 A",
            ["mode_b"] = "模式 B",
            ["target_class"] = "目标类:",

            // ===== Tabs =====
            ["tab_dashboard"] = "总览",
            ["tab_data_preview"] = "数据预览",
            ["tab_validation"] = "校验",
            ["tab_export"] = "导出",
            ["tab_runtime"] = "运行时",

            // ===== Dashboard =====
            ["dashboard_title"] = "项目总览",
            ["excel_dir_label"] = "Excel 目录:",
            ["output_dir_label"] = "输出目录:",
            ["quick_status"] = "快速状态",
            ["no_excel_files"] = "没有找到 Excel 文件。请将 .xlsx 放入 Assets/Excel/ 或点击 '+ 添加 Excel 目录'。",
            ["file_col"] = "文件",
            ["sheets_col"] = "表",
            ["status_col"] = "状态",
            ["status_exported"] = "✅ 已导出",
            ["status_not_exported"] = "⚪ 未导出",
            ["quick_actions"] = "快捷操作",
            ["export_all_btn"] = "全部导出",
            ["validate_all_btn"] = "全部校验",
            ["refresh_files_btn"] = "刷新文件",
            ["generated_output"] = "生成结果",
            ["no_generated_files"] = "(暂无生成文件)",

            // ===== Data Preview =====
            ["select_sheet_hint"] = "从左侧文件树中选择一个 Sheet 预览数据。",
            ["no_data"] = "此 Sheet 没有数据。",
            ["search"] = "搜索:",
            ["find"] = "查找",
            ["found_at_row"] = "在第 {0} 行找到",
            ["page"] = "页",
            ["showing_rows"] = "显示 {0} / {1} 行（第 {2} 页）",

            // ===== Validation =====
            ["validation_title"] = "校验结果",
            ["no_errors"] = "未发现错误。",
            ["errors_count"] = "{0} 个错误",
            ["warnings_count"] = "{0} 个警告",
            ["infos_count"] = "{0} 个提示",
            ["filter_all"] = "全部",
            ["filter_errors"] = "错误",
            ["filter_warnings"] = "警告",
            ["copy_csv"] = "复制为 CSV",
            ["export_csv"] = "导出 CSV",

            // ===== Export =====
            ["export_config"] = "导出配置",
            ["output_format"] = "输出格式:",
            ["format_so"] = "ScriptableObject (.asset)",
            ["format_json"] = "JSON",
            ["output_path"] = "输出路径:",
            ["auto_validate"] = "导出前自动校验",
            ["block_on_error"] = "校验错误时阻止导出",
            ["export_json_too"] = "同时导出 JSON（用于热更新）",
            ["export_all_btn_lg"] = "全部导出",
            ["export_selected_btn"] = "导出选中",
            ["export_templates_btn"] = "从 C# 类导出 Excel 模板",
            ["template_done"] = "已生成 {0} 个 Excel 模板。",
            ["no_template_classes"] = "未找到 [ExcelTable] 类。",
            ["gen_files_title"] = "已生成文件",
            ["no_gen_files"] = "(暂无生成文件)",
            ["select_btn"] = "选中",

            // ===== Runtime =====
            ["runtime_title"] = "运行时 API 参考",
            ["quick_start"] = "快速开始",
            ["quick_start_text"] = "1. 点击下方「生成 DataManager」按钮\n2. 在游戏代码中:\n   var t = DataManager.Instance.GetTable<YourTable>();\n   var row = t.Get(id);",
            ["gen_datamanager_btn"] = "自动生成 DataManager GameObject",
            ["api_ref"] = "API 参考",

            // ===== Settings =====
            ["settings_title"] = "插件设置",
            ["excel_parse_settings"] = "Excel 解析设置",
            ["header_row"] = "字段名行",
            ["type_row"] = "类型行",
            ["comment_row"] = "注释行 (0 = 无)",
            ["data_start_row"] = "数据起始行",
            ["skip_hidden_rows"] = "跳过隐藏行",
            ["skip_hidden_cols"] = "跳过隐藏列",
            ["skip_empty_rows"] = "跳过空行",
            ["skip_prefixes"] = "跳过 Sheet 前缀 (逗号分隔)",
            ["validation_section"] = "校验",
            ["strict_mode"] = "严格模式 (警告也阻止导出)",
            ["auto_export_section"] = "自动导出",
            ["watch_files"] = "监听 Excel 文件变更",
            ["debounce"] = "防抖延迟 (ms)",
            ["code_gen_section"] = "代码生成",
            ["output_path_setting"] = "输出路径 (Assets/ 下)",
            ["namespace_setting"] = "命名空间",
            ["language_section"] = "语言",
            ["language_label"] = "界面语言:",
            ["lang_zh"] = "中文",
            ["lang_en"] = "English",
            ["save_btn"] = "保存",
            ["reset_defaults_btn"] = "恢复默认",
            ["settings_saved"] = "设置已保存。",

            // ===== Sheet mode overrides =====
            ["sheet_mode_title"] = "Sheet 模式覆盖",
            ["sheet_mode_hint"] = "为每个 Sheet 单独选择映射模式:\n• 自动 — 有 [ExcelTable] 类则用模式 B，否则用模式 A\n• 模式 A — 始终从 Excel 自动生成 C# 代码\n• 模式 B — 始终使用对应的 C# 类",

            // ===== Status bar =====
            ["status_export_complete"] = "导出完成: {0} 文件 → {1} 表 → {2} 行 ({3:F1}s)",
            ["status_export_error"] = "导出错误: {0} 错误, {1} 警告",
            ["status_validate_pass"] = "校验通过: {0} 表 OK",
            ["status_validate_error"] = "校验: {0} 错误, {1} 警告",
            ["status_exporting"] = "正在导出 {0}...",
            ["last_export"] = "上次导出:",
            ["shortcut_hint"] = "Ctrl+E 导出 | Ctrl+V 校验 | Ctrl+R 刷新",
            ["template_generated"] = "已生成 {0} 个 Excel 模板",

            // ===== Sample Generator =====
            ["sample_menu"] = "Window/Excel 数据管理器/生成示例数据",
            ["sample_done_title"] = "示例数据",
            ["sample_done_msg"] = "示例 Excel 文件已生成到 {0}\n\n• Item.xlsx — Weapon + Armor 表\n• SkillEnum.xlsx — 枚举引用表\n\n打开 Excel 数据管理器，选择 Item.xlsx，点击导出。",

            // ===== Color Legend =====
            ["legend_valid"] = "✅ 通过",
            ["legend_warning"] = "⚠ 警告",
            ["legend_error"] = "❌ 错误",
            ["legend_empty"] = "⚪ 空值",
        };

        private static readonly Dictionary<string, string> EnDict = new Dictionary<string, string>
        {
            ["window_title"] = "Excel Data Manager",
            ["window_menu"] = "Window/Excel Data Manager",
            ["refresh"] = "Refresh",
            ["export_all"] = "Export All",
            ["validate"] = "Validate",
            ["template"] = "Template",
            ["settings"] = "Settings",
            ["ready"] = "Ready",
            ["toolbar_refresh"] = "Refresh",
            ["toolbar_export_all"] = "Export All",
            ["toolbar_validate"] = "Validate",
            ["toolbar_template"] = "Template",
            ["files"] = "Files",
            ["search_filter"] = "Search...",
            ["filter"] = "Filter",
            ["selected_count"] = "{0} selected",
            ["export_selected"] = "Export Selected",
            ["validate_selected"] = "Validate Selected",
            ["clear"] = "Clear",
            ["add_excel_dir"] = "+ Add Excel Directory",
            ["export_this_file"] = "Export This File",
            ["validate_this_file"] = "Validate This File",
            ["open_excel"] = "Open Excel",
            ["add_to_selection"] = "Add to Selection",
            ["export_this_sheet"] = "Export This Sheet",
            ["validate_this_sheet"] = "Validate This Sheet",
            ["mapping_mode"] = "Mapping:",
            ["mode_auto"] = "Auto",
            ["mode_a"] = "Mode A",
            ["mode_b"] = "Mode B",
            ["target_class"] = "Target:",
            ["tab_dashboard"] = "Dashboard",
            ["tab_data_preview"] = "Data Preview",
            ["tab_validation"] = "Validation",
            ["tab_export"] = "Export",
            ["tab_runtime"] = "Runtime",
            ["dashboard_title"] = "Project Overview",
            ["excel_dir_label"] = "Excel Dir:",
            ["output_dir_label"] = "Output Dir:",
            ["quick_status"] = "Quick Status",
            ["no_excel_files"] = "No Excel files found. Place .xlsx files in Assets/Excel/ or click '+ Add Excel Directory'.",
            ["file_col"] = "File",
            ["sheets_col"] = "Sheets",
            ["status_col"] = "Status",
            ["status_exported"] = "✅ Exported",
            ["status_not_exported"] = "⚪ Not exported",
            ["quick_actions"] = "Quick Actions",
            ["export_all_btn"] = "Export All",
            ["validate_all_btn"] = "Validate All",
            ["refresh_files_btn"] = "Refresh Files",
            ["generated_output"] = "Generated Output",
            ["no_generated_files"] = "(No generated files yet)",
            ["select_sheet_hint"] = "Select a sheet from the file tree to preview data.",
            ["no_data"] = "No data in this sheet.",
            ["search"] = "Search:",
            ["find"] = "Find",
            ["found_at_row"] = "Found at row {0}",
            ["page"] = "Page",
            ["showing_rows"] = "Showing {0} of {1} rows (page {2})",
            ["validation_title"] = "Validation Results",
            ["no_errors"] = "No errors found.",
            ["errors_count"] = "{0} errors",
            ["warnings_count"] = "{0} warnings",
            ["infos_count"] = "{0} infos",
            ["filter_all"] = "All",
            ["filter_errors"] = "Errors",
            ["filter_warnings"] = "Warnings",
            ["copy_csv"] = "Copy as CSV",
            ["export_csv"] = "Export CSV",
            ["export_config"] = "Export Configuration",
            ["output_format"] = "Output Format:",
            ["format_so"] = "ScriptableObject (.asset)",
            ["format_json"] = "JSON",
            ["output_path"] = "Output Path:",
            ["auto_validate"] = "Auto-validate before export",
            ["block_on_error"] = "Block export on validation error",
            ["export_json_too"] = "Also export JSON (for hot-update)",
            ["export_all_btn_lg"] = "Export All",
            ["export_selected_btn"] = "Export Selected",
            ["export_templates_btn"] = "Export Excel Templates (from C# classes)",
            ["template_done"] = "Generated {0} Excel template(s).",
            ["no_template_classes"] = "No [ExcelTable] classes found.",
            ["gen_files_title"] = "Generated Files",
            ["no_gen_files"] = "(No generated files yet)",
            ["select_btn"] = "Select",
            ["runtime_title"] = "Runtime API Reference",
            ["quick_start"] = "Quick Start",
            ["quick_start_text"] = "1. Click 'Generate DataManager' button below\n2. In your game code:\n   var t = DataManager.Instance.GetTable<YourTable>();\n   var row = t.Get(id);",
            ["gen_datamanager_btn"] = "Auto-Generate DataManager GameObject",
            ["api_ref"] = "API Reference",
            ["settings_title"] = "Plugin Settings",
            ["excel_parse_settings"] = "Excel Parse Settings",
            ["header_row"] = "Header Row (field names)",
            ["type_row"] = "Type Row",
            ["comment_row"] = "Comment Row (0 = none)",
            ["data_start_row"] = "Data Start Row",
            ["skip_hidden_rows"] = "Skip Hidden Rows",
            ["skip_hidden_cols"] = "Skip Hidden Columns",
            ["skip_empty_rows"] = "Skip Empty Rows",
            ["skip_prefixes"] = "Skip Sheet Prefixes (comma-separated)",
            ["validation_section"] = "Validation",
            ["strict_mode"] = "Strict Mode (warnings block export)",
            ["auto_export_section"] = "Auto Export",
            ["watch_files"] = "Watch Excel files",
            ["debounce"] = "Debounce (ms)",
            ["code_gen_section"] = "Code Generation",
            ["output_path_setting"] = "Output Path (under Assets/)",
            ["namespace_setting"] = "Namespace",
            ["language_section"] = "Language",
            ["language_label"] = "Language:",
            ["lang_zh"] = "中文",
            ["lang_en"] = "English",
            ["save_btn"] = "Save",
            ["reset_defaults_btn"] = "Reset Defaults",
            ["settings_saved"] = "Settings saved.",
            ["sheet_mode_title"] = "Sheet Mode Overrides",
            ["sheet_mode_hint"] = "Choose mapping mode per sheet:\n• Auto — Mode B if [ExcelTable] class exists, else Mode A\n• Mode A — Always auto-generate C# from Excel\n• Mode B — Always use matching C# class",
            ["status_export_complete"] = "Export complete: {0} files, {1} sheets, {2} rows ({3:F1}s)",
            ["status_export_error"] = "Export errors: {0} errors, {1} warnings",
            ["status_validate_pass"] = "Validation passed: {0} sheets OK",
            ["status_validate_error"] = "Validation: {0} errors, {1} warnings",
            ["status_exporting"] = "Exporting {0}...",
            ["last_export"] = "Last export:",
            ["shortcut_hint"] = "Ctrl+E Export | Ctrl+V Validate | Ctrl+R Refresh",
            ["template_generated"] = "Generated {0} Excel template(s)",
            ["sample_menu"] = "Window/Excel Data Manager/Generate Sample Data",
            ["sample_done_title"] = "Sample Data",
            ["sample_done_msg"] = "Sample Excel files generated at {0}\n\n• Item.xlsx — Weapon + Armor sheets\n• SkillEnum.xlsx — Enum reference table\n\nOpen Excel Data Manager, select Item.xlsx, and export.",
            ["legend_valid"] = "✅ Valid",
            ["legend_warning"] = "⚠ Warning",
            ["legend_error"] = "❌ Error",
            ["legend_empty"] = "⚪ Empty",
        };

        public static bool IsZh => GetLanguage() == "zh";

        public static string Tr(string key)
        {
            var dict = IsZh ? ZhDict : EnDict;
            if (dict.TryGetValue(key, out var val))
                return val;
            // Fallback to key with ! prefix for missing translations
            return $"!{key}";
        }

        public static string Tr(string key, params object[] args)
        {
            return string.Format(Tr(key), args);
        }

        public static string GetLanguage()
        {
            return EditorPrefs.GetString("ExcelToJson.Language", "zh");
        }

        public static void SetLanguage(string lang)
        {
            EditorPrefs.SetString("ExcelToJson.Language", lang == "en" ? "en" : "zh");
        }
    }
}
