<artifact id="tasks" change="enum-selection-double-click" schema="spec-driven">

## Tasks

### 1. Update VBA Source Code (`Source/DataWorkbook/`)
- [ ] In `ThisWorkbook.cls`, replace `Workbook_SheetSelectionChange` with `Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)`.
- [ ] In `Module_EnumSelector.bas`, change `TryLaunchEnumSelector` to a `Function` returning `Boolean`. Return `True` right before `Form_EnumSelect.Show`.
- [ ] Inside `Workbook_SheetBeforeDoubleClick` hook, call `If Module_EnumSelector.TryLaunchEnumSelector(Target) Then Cancel = True`.
- [ ] In `Form_EnumSelect.frm`, move cell overwriting logic out of `lstEnums_Click`.
- [ ] In `Form_EnumSelect.frm`, create `btnConfirm_Click` to apply the value, trigger the custom Undo registration, and close the form.
- [ ] In `Form_EnumSelect.frm`, create `btnCancel_Click` to close the form without saving.

### 2. Update Python Injection Script (`inject_vba.py`)
- [ ] In `inject_vba.py`, under `# Inject UserForm programmatically`, modify the height of the user form `form_comp.Properties("Height").Value` to make room for new buttons.
- [ ] Add the `btnConfirm` and `btnCancel` command buttons dynamically in python via standard layout coordinates.
- [ ] Assign captions `[確認]` and `[取消]` in UI rendering.

### 3. Verification
- [ ] Run `python inject_vba.py "reference\Form\Test_Contacts.xlsx"`.
- [ ] Open `Test_Contacts_MacroEnabled.xlsm`.
- [ ] Verify single-click highlights cell normally without triggering macro.
- [ ] Verify double-click launches UserForm AND suppresses native edit mode.
- [ ] Verify Cancel closes without altering cell value.
- [ ] Verify Confirm writes the selected value and closes.
- [ ] Verify Ctrl+Z successfully undoes the selection using the custom handler.

</artifact>
