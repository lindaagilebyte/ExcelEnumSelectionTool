# Design: automate-vba-injection

## Context
The VBA Enum Selection tool requires manual copying of `.bas`, `.cls`, and `.frm` files into an Excel workbook. This is prone to error and causes encoding issues on non-English locales when dealing with `.frm` and `.frx` files. A Python-based automation script is needed to securely inject these files into a pristine Excel environment. Additionally, we need to implement debugging features and localized UI components that were missed in the initial rollout.

## Goals / Non-Goals
**Goals:**
- Provide a 1-click script (`inject_vba.py`) to inject the `Source/DataWorkbook/` VBA files into any target `.xlsx` file and save it as `.xlsm`.
- Update the VBA code to support Traditional Chinese UI.
- Implement a `CONST_DEBUG_MODE` to aid in path resolution troubleshooting.
- Add status bar feedback to prevent users from thinking the tool is hanging during large sheet scans.
- Provide a clear warning when an Enum ID exists but its data block is missing.

**Non-Goals:**
- We are not rewriting the core Enum selection logic (which already works).
- We are not building a full Excel Add-in (`.xlam`) at this point, as the requirement is for the code to live inside the Data Workbooks.

## Decisions
1. **Python with `win32com`:** Python's `win32com.client` is the most robust way to interact with the Excel VBE (Visual Basic Extensibility) object model externally. It can seamlessly add modules and import forms.
2. **Late Binding for Dictionary (Kept from current state):** We will maintain the use of `CreateObject("Scripting.Dictionary")` to ensure the tool runs on any user's machine without requiring manual VBA Reference checks.
3. **Debug Logging:** A public constant `Public Const CONST_DEBUG_MODE As Boolean = True` will toggle explicit `Debug.Print` statements for path resolution and cache hits.
4. **Status Bar Usage:** `Application.StatusBar` will be updated during the `ScanWorksheet` loop, and reset to `False` once complete or on error.

## Risks / Trade-offs
- **VBE Access:** The user running the Python script *must* have "Trust access to the VBA project object model" enabled in their Excel Macro Settings. The script should detect if this fails and instruct the user how to enable it.
- **Form UI Locale:** Hardcoding Traditional Chinese strings into the `.frm` file might still pose an issue if the text file itself gets corrupted before injection. We will mitigate this by setting the `.Caption` properties via VBA code in `UserForm_Initialize()` rather than relying on the imported `.frm` properties.
