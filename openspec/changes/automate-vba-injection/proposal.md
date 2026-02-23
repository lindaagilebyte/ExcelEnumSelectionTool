# Proposal: automate-vba-injection

## Goal
To eliminate manual setup errors and provide a reliable, testable distribution mechanism for the Excel Enum Selection tool. Additionally, this change will fulfill the missing requirements from the original specification (Traditional Chinese UI, Status Bar feedback, error handling for missing definitions, and debug logging).

## New Capabilities
- `vba-injector-script`: A Python script using `win32com` to automatically inject the `.bas`, `.frm`, and `.cls` files into a target `.xlsx` file and save it as a macro-enabled `.xlsm` file.
- `debug-logging`: A `CONST_DEBUG_MODE` in the VBA module that outputs search paths and cache status to the Immediate Window.
- `status-bar-feedback`: Non-disruptive feedback in the Excel Status Bar during the reference file scanning process.
- `missing-definition-warning`: A clear alert if a column has an Enum ID in Row 2, but the reference file is missing the `定義(巨集顯示)` block for that ID.

## Impacted Capabilities
- `enum-selector-ui`: The UserForm UI needs to be localized to Traditional Chinese as per the original specification.
- `cache-management`: The cache needs to be explicitly cleared when the workbook closes to manage memory.

## Technical Scope
- **Development Tooling**: Creating a new Python script (`inject_vba.py`) in the `Source/` or `.agent/` directory.
- **VBA Module (`Module_EnumSelector.bas`)**: Adding debug constants, status bar updates (`Application.StatusBar`), and error handling for empty cache results.
- **VBA Workbook Hook (`ThisWorkbook.cls`)**: Adding the `Workbook_BeforeClose` event to clear the dictionary cache.
- **VBA UserForm (`Form_EnumSelect.frm`)**: Updating labels and captions to Traditional Chinese.
