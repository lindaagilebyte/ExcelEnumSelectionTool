# VBA Tool Specs

## Purpose
The VBA Tool provides an automated, user-friendly Dropdown Menu in Excel to select enum values defined in a central reference file, ensuring data entry consistency. It is deployed via a Python injection script.

## Capabilities
- `vba-tool`
- `vba-injector-script`

---

## Requirements

### Requirement: Data Loading (Reference File)
#### Scenario: Caching Scope
- **WHEN** the Data Workbook is open
- **THEN** the cache exists in memory for the duration of the Excel Session.

#### Scenario: Cold Start
- **WHEN** the tool is triggered for the first time
- **THEN** scan the Reference File entirely, index Enum Keys in Row 2, and store `SheetName` and `CellAddress` for each Key.

#### Scenario: Warm Start
- **WHEN** the tool is triggered with a known Key
- **THEN** reference the in-memory cache to jump directly to the target sheet/cell without re-scanning.

#### Scenario: Extracting List
- **WHEN** finding a matching Enum Key
- **THEN** extract the list under `定義(巨集顯示)` down to the next empty cell or block header.

### Requirement: Trigger Logic (`double-click-activation`)
#### Scenario: Selection in Data Area
- **WHEN** a user double-clicks a cell (`Workbook_SheetBeforeDoubleClick`)
- **THEN** check if `ActiveCell.Row >= 4`. If true, check the Row 2 Header for an Enum Key.
- **AND** set `Cancel = True` to prevent normal Excel cell editing.

#### Scenario: Valid Key Found
- **WHEN** the Row 2 Enum Key matches a known definition
- **THEN** display the Enum Selector UserForm.

### Requirement: UserForm UI (`explicit-confirmation`)
#### Scenario: Enum Selection UI
- **WHEN** the form is launched
- **THEN** title is "選擇數值: [EnumName]", header is "請為 [EnumName] 選擇一個數值:", and refresh button is "重新整理快取".
- **AND** `lstEnums` click only selects the item visually.
- **AND** `btnConfirm` ([確認]) applies the value, logs `Application.OnUndo`, and closes the form.
- **AND** `btnCancel` ([取消]) closes the form without applying.

### Requirement: VBA Injector Script (`vba-injector-script`)
#### Scenario: Code Injection
- **WHEN** user runs `python inject_vba.py <path_to_file>`
- **THEN** the script creates a new macro-enabled clone if the target is `.xlsx`.
- **OR** if the target is `.xlsm`, the script removes old injected components and saves the updated VBA in-place.

### Requirement: Feedback & Logging
#### Scenario: Status Bar and Warnings
- **WHEN** scanning large files, **THEN** display "正在讀取列舉定義..." in the Excel Status Bar.
- **WHEN** a definition is empty, **THEN** show a warning MsgBox: "找不到 [EnumName] 的資料定義...".
- **WHEN** debug mode is enabled, **THEN** output resolution paths to the Immediate Window.
