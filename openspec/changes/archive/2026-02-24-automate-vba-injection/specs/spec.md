# Specs: automate-vba-injection

## Capabilities
- `vba-injector-script`
- `status-bar-feedback`
- `missing-definition-warning`
- `debug-logging`
- `enum-selector-ui` (Updated)

---

## ADDED Requirements

### Requirement: vba-injector-script
The tool MUST automatically inject VBA source files into a target Excel workbook.
#### Scenario: Successful Injection
- **WHEN** the user runs `python inject_vba.py <path_to_xlsx>`
- **THEN** a new `<path_to_xlsx_basename>_MacroEnabled.xlsm` is generated containing the `Module_EnumSelector`, `Form_EnumSelect`, and `ThisWorkbook` code.

### Requirement: status-bar-feedback
The tool MUST provide non-intrusive feedback during heavy processing.
#### Scenario: Scanning Large Reference Files
- **WHEN** the tool begins scanning the reference workbook
- **THEN** the Excel Status Bar displays "正在讀取列舉定義... (Sheet X of Y)"
- **WHEN** the scan is complete
- **THEN** the Status Bar is reset to default.

### Requirement: missing-definition-warning
The tool MUST alert the user if a definition is incomplete.
#### Scenario: Incomplete Enum Block
- **WHEN** the user clicks an active column, and the Enum Key is found in the reference file, but the `定義(巨集顯示)` block is empty or missing
- **THEN** a `MsgBox` alerts the user: "找不到 [EnumName] 的資料定義，請檢查列舉參考檔。"

### Requirement: debug-logging
The tool MUST allow developers to trace its execution path.
#### Scenario: Path Resolution Tracing
- **WHEN** `CONST_DEBUG_MODE` is `True`
- **THEN** the tool outputs its path resolution attempts and cache load results to the VBA Immediate Window (`Debug.Print`).

## MODIFIED Requirements

### Requirement: enum-selector-ui
The tool MUST display all UI elements in Traditional Chinese.
#### Scenario: Opening the Form
- **WHEN** the form is launched
- **THEN** the form title is "選擇數值: [EnumName]".
- **THEN** the header label is "請為 [EnumName] 選擇一個數值:".
- **THEN** the refresh button is "重新整理快取".
