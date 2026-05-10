# Changelog

## [1.0.6] — 2026-05-10
### Fixed
- Excel Data Manager: choosing an Excel folder outside the Unity project (e.g. Desktop) now stores the **full absolute path**. Previously the picker path was collapsed to the folder name only, so scans targeted an empty folder under the project root and listed no `.xlsx` files.

## [1.0.5] — 2026-05-10
### Fixed
- Remove root `Samples~.meta` to avoid Unity warnings about an empty/hidden `Samples~` folder under UPM; sample content under `Samples~/Demo` keeps its own `.meta` files.

## [1.0.4] — 2026-05-10
### Fixed
- UPM Git install: commit Unity `.meta` files for all package assets and folders. Without them, Unity reports *"has no meta file, but it's in an immutable folder"* and ignores scripts/DLLs, so the package appears empty.

## [1.0.3] — 2026-05-10
### Fixed
- UPM / Git URL import: move NPOI and related DLLs to `Editor/` root (same folder as the Editor `.asmdef`) so Unity detects plugin assemblies reliably.
- Editor assembly: set `overrideReferences` to `false` and drop explicit `precompiledReferences` so Unity can resolve framework assemblies together with bundled DLLs.

## [Sprint 3] — 2026-05-08
### Added
- Mode B (C# reflection matching): Export data using [ExcelTable] / [ExcelColumn] attributed classes
- Mode C (Hybrid matching): Different sheets in same file can use different mapping modes
- Template exporter: Generate empty Excel templates from C# classes
- Stage 3 validation: Cross-file ref/enum reference integrity checks
- File watcher: Auto-export on Excel file changes (debounce configurable)
- Incremental export: Only re-process sheets that have changed (MD5 hash based)
- Ref chain tracking: Cascading re-export when referenced tables change
- Sheet mode override UI: Per-sheet mapping mode selection in Settings
- Field type color coding in Data Preview header row

### Fixed
- Compilation error: `sheetName` undefined variable → `schema.TableName`
- Compilation error: `RegexParseException` not found → `System.ArgumentException`

## [Sprint 2] — 2026-05-08
### Added
- Editor Window with 5 tabs (Dashboard, Data Preview, Validation, Export, Runtime API)
- 16 built-in validation rules across 3 stages (Structure, Data, Custom)
- #Rules Sheet support for dynamic validation rules (range, regex, multiple, not_empty, enum)
- Progress bar for export/validate operations
- Ctrl/Shift multi-select in file tree
- Settings window with EditorPrefs persistence
- Dashboard with file overview and quick actions
- Data Preview with pagination and search
- Validation panel with error list, color icons, double-click navigation, CSV export
- Runtime API reference tab with code examples
- Auto-generate DataManager GameObject button

## [Sprint 1] — 2026-05-08
### Added
- NPOI 2.5.6 net45 integration for .xlsx/.xls reading
- ExcelReader with merged cell, hidden row/column, and empty row handling
- SchemaParser: Field names, types, comments from Excel header rows
- TypeMapper: int/float/string/bool/array/Vector2/Vector3/Color/ref/enum/res/json/loc
- DataParser: Row-level data parsing with error collection
- CodeGenerator: Auto-generate C# Row and Table classes from schema
- AssetGenerator: Create ScriptableObject .asset files via reflection
- Pipeline: Full Excel → .asset pipeline with progress reporting
- BaseDataTable<T> runtime base class with Get/Find/GetRandom/GetByIds
- DataManager runtime singleton with type-safe table queries
- 7 test Excel cases (Types, Arrays, UnityTypes, Composite, Large, SkipSheets, BoolEdge)
