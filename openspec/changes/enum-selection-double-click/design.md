<artifact id="design" change="enum-selection-double-click" schema="spec-driven">

## Context
The VBA Enum Selection Tool currently triggers on `Workbook_SheetSelectionChange`. This prevents users from clicking a cell just to select or copy it. Furthermore, it automatically applies the value upon clicking an item in the ListBox, with no option to browse or cancel.

## Goals
- Move the event hook from `SheetSelectionChange` to `SheetBeforeDoubleClick`.
- Suppress Excel's default "edit cell" mode on double-click if the cell is valid for enum selection.
- Add "Confirm" and "Cancel" buttons to the UserForm.

## Non-Goals
- Changing how data is loaded from the reference file.
- Changing the specific cells the tool applies to.
- Solving other unrelated Excel bugs.

## Decisions
1. **Event Migration**: Change `Workbook_SheetSelectionChange` to `Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)`.
2. **Double-Click Suppression**: Set `Cancel = True` inside the double-click event before launching the form, so Excel does not enter cursor-edit mode on the cell.
3. **UserForm Buttons**: Add two CommandButtons to `Form_EnumSelect.frm`. Change the `lstEnums_Click` event to merely highlight the choice (no action), and move the cell write/undo logic into `btnConfirm_Click`. 
4. **Python Injection Updates**: The `inject_vba.py` script dynamically constructs the UserForm. It must be updated to insert `btnConfirm` and `btnCancel`, position them at the bottom, and resize the form's total height accordingly.

## Risks / Trade-offs
- Risk: Users might not immediately realize they need to double-click instead of single-click. 
- Mitigation: Double-click is standard for triggering edit menus in Excel, so it should be intuitive. 
- Trade-off: The UI takes slightly longer to confirm an entry now, but it's much safer than instantly overwriting values on accident.

</artifact>
