# Excel To JSON — Unity Editor Plugin

A Unity Editor plugin that converts Excel spreadsheets into strongly-typed ScriptableObject assets with validation, code generation, and one-click export.

## Features

- **Excel → ScriptableObject pipeline**: Place .xlsx files in `Assets/Excel/`, export to `Assets/Data/` as .asset files
- **16 built-in validation rules**: Structure, data integrity, type matching, cross-file references
- **Three mapping modes**: Auto-generate C# (Mode A), C# reflection (Mode B), Hybrid (Mode C)
- **File watcher**: Auto-export on Excel file changes (debounce configurable)
- **Incremental export**: Only re-process changed sheets
- **Data preview**: Paginated table view with search and type-colored headers
- **Runtime API**: Type-safe `DataManager` with Get/Find/GetRandom/GetByIds queries

## Installation

### Install via Git URL (recommended)
1. Open Unity, go to **Window > Package Manager**
2. Click the **+** button, select **"Add package from git URL"**
3. Enter: `https://github.com/zhouxiangyang789456-max/-Excel-JSON-.git`
4. Click **Add** — the plugin appears under `Packages/com.github.excel-to-json`

### Install from Disk
1. Clone or download this repository
2. In Package Manager, click **+** → **"Add package from disk"**
3. Select the `package.json` file in the repo root

### Import Demos (optional)
1. In Package Manager, select **"Excel To JSON"**
2. Switch to the **Samples** tab
3. Click **Import** on "Demo Data Usage"

## Quick Start

1. Put your `.xlsx` files in your project's `Assets/Excel/` folder
2. Open **Window > Excel Data Manager**
3. Select a sheet in the file tree
4. Click **Export** (toolbar or Export tab)
5. Code is generated to `Assets/Scripts/Generated/`, assets to `Assets/Data/`

## Supported Types

| Excel Type | C# Type | Example |
|-----------|---------|---------|
| `int` | `int` | `42` |
| `float` | `float` | `3.14` |
| `string` | `string` | `hello` |
| `bool` | `bool` | `true` / `false` / `是` / `否` |
| `int[]` | `int[]` | `1\|2\|3` or `[1,2,3]` |
| `float[]` | `float[]` | `1.0\|2.5\|3.0` |
| `string[]` | `string[]` | `a\|b\|c` |
| `Vector2` | `Vector2` | `[1,2]` |
| `Vector3` | `Vector3` | `[1,2,3]` |
| `Color` | `Color` | `#FF0000` or `[1,0,0,1]` |
| `ref:TableName` | `int` | `1001` (FK reference) |
| `enum:TableName` | `int` | `1` (enum value) |
| `res:Sprite` | `string` | `Sprites/icon` |
| `json` | `string` | `{"key":"value"}` |
| `loc` | `string` | Localization key |

## Excel Format

Each data sheet has 4 rows of header:
- **Row 1**: Field names (e.g., `id`, `name`, `attack`)
- **Row 2**: Type declarations (e.g., `int`, `string`, `ref:Weapon`)
- **Row 3**: Comments (optional, used as Tooltip in generated code)
- **Row 4+**: Data rows

Sheets prefixed with `_` or `#` are skipped. `#Rules` sheets define custom validation rules.

## Mapping Modes

### Mode A: Excel-Driven (Auto Code Generation)
- Designer defines fields in Excel
- Plugin generates C# Row/Table classes
- Best for prototyping and new projects

### Mode B: C#-Driven (Reflection Matching)
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
    public int ComputedValue; // Not in Excel
}
```
- Programmer defines C# classes with attributes
- Plugin matches Excel columns to C# fields
- Best for existing projects with complex types

### Mode C: Hybrid
Each sheet independently chooses Mode A or B. Set per-sheet mode in Settings.

## Validation Rules

### Stage 1 — Structure
- Field name uniqueness
- Type declaration validity
- Header completeness
- Field name sanity (no special chars)

### Stage 2 — Data
- ID required and unique
- Type matching (value vs declared type)
- Required field check
- Enum format sanity
- Resource path format
- Formula detection

### Stage 3 — Cross-Table References
- ref integrity: Verify referenced IDs exist in target tables
- enum existence: Verify enum values in target enum tables

### Custom Rules (#Rules sheet)
Define in Excel: `range`, `regex`, `multiple`, `not_empty`, `required`, `enum`

## Runtime API

```csharp
var dm = DataManager.Instance;

// Get table by type
var weaponTable = dm.GetTable<WeaponTable>();

// Query by ID (O(1))
var weapon = weaponTable.Get(1001);

// Get all rows
foreach (var w in weaponTable.GetAll()) { ... }

// Filter
var strong = weaponTable.Find(w => w.Attack > 30);

// Random pick
var random = weaponTable.GetRandom(w => w.Rarity >= 3);
```

## Editor Window

- **Dashboard (Tab 0)**: Project overview, quick status, export buttons
- **Data Preview (Tab 1)**: Paginated table view with search
- **Validation (Tab 2)**: Error list with color icons, double-click navigation, CSV export
- **Export (Tab 3)**: Export configuration, template generation, generated file list
- **Runtime API (Tab 4)**: API reference, code examples, DataManager generation

## Requirements

- Unity 2021.3 LTS or later
- API Compatibility Level: .NET Framework
- No external dependencies (NPOI bundled as DLLs)

## License

MIT License — free for personal and commercial use.
