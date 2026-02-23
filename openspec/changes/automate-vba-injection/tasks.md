# Tasks: automate-vba-injection

## 1. Python Automation Script
- [ ] 1.1 Create `inject_vba.py` script.
- [ ] 1.2 Implement `win32com` hook to open target `.xlsx` file silently.
- [ ] 1.3 Implement logic to import `Module_EnumSelector.bas`, `Form_EnumSelect.frm`, and `ThisWorkbook.cls` into the active VBA Project.
- [ ] 1.4 Implement SaveAs logic to Output as `.xlsm` with `52` file format constant.

## 2. VBA Missing Features (UI & Feedback)
- [ ] 2.1 Update `Form_EnumSelect.frm` to force Traditional Chinese `.Caption` properties via `UserForm_Initialize()`.
- [ ] 2.2 Update `Module_EnumSelector.bas` to use `Application.StatusBar` during `ScanWorksheet()`.
- [ ] 2.3 Add `MsgBox` warning in `Module_EnumSelector.bas` when an Enum Key exists but returns an empty list.

## 3. VBA Missing Features (Debugging & Memory)
- [ ] 3.1 Introduce `CONST_DEBUG_MODE` in `Module_EnumSelector.bas` with `Debug.Print` outputs for path resolution.
- [ ] 3.2 Update `ThisWorkbook.cls` to include `Workbook_BeforeClose` event that sets the Cache to `Nothing`.

## 4. Verification
- [ ] 4.1 Copy `reference/Form/Contacts.xlsx` to a safe test location (e.g., `reference/Form/Test_Contacts.xlsx`) to avoid corrupting the real file.
- [ ] 4.2 Run `inject_vba.py` against `reference/Form/Test_Contacts.xlsx`.
- [ ] 4.3 Verify the output `Test_Contacts_MacroEnabled.xlsm` opens correctly without corruption and successfully triggers the VBA logic against `..\列舉定義(企劃用).xlsx`.
- [ ] 4.4 Verify all new VBA features (status bar, Traditional Chinese UI, debug logging, warning boxes, memory release) function as expected when active.
