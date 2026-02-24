# Tasks: Excel Enum Selection Tool

- [ ] **1. Core Logic (Module_EnumSelector)**
    - [ ] 1.1 Implement `GetReferenceFilePath()` with fallback logic
    - [ ] 1.2 Implement `LoadEnumDefinitions(Key)` with `Scripting.Dictionary` caching
    - [ ] 1.3 Implement `ScanReferenceWorkbook(Workbook, Key)` to parse `EnumDef`

- [ ] **2. UI (Form_EnumSelect)**
    - [ ] 2.1 Create UserForm layout (ListBox, Label)
    - [ ] 2.2 Implement `UserForm_Initialize` to populate list
    - [ ] 2.3 Implement selection logic (write to ActiveCell)

- [ ] **3. Integration (ThisWorkbook)**
    - [ ] 3.1 Hook `Workbook_SheetSelectionChange`
    - [ ] 3.2 Implement Row/Column validation (Row > 3, Row 2 Key check)

- [ ] **4. Verification**
    - [ ] 4.1 Manual Test: Copy code to `Contacts.xlsx` and verify trigger
    - [ ] 4.2 Manual Test: Verify data loading from `EnumDef`
