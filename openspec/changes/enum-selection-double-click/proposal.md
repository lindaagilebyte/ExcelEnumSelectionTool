<artifact id="proposal" change="enum-selection-double-click" schema="spec-driven">

## Objective
Change the Enum Selection Tool's activation from a single click to a double-click, and require explicit confirmation (or cancellation) on the UserForm before writing data to the cell.

## Requirements
- The Enum Selection Tool must only trigger when a user double-clicks an eligible cell (Row >= 4, Column has a valid Enum Key in Row 2).
- The UserForm must not automatically write the value and close when an item in the ListBox is clicked.
- The UserForm must have a "Confirm" (確認) button to write the selected value and close.
- The UserForm must have a "Cancel" (取消) button to close the form without writing anything.
- The standard Excel double-click edit mode must be suppressed when the form is launched on an eligible cell.

## Capabilities
- `double-click-activation`: The ability to trigger the tool via double-click instead of single-click.
- `explicit-confirmation`: The ability to select an item and explicitly confirm or cancel the operation via dedicated buttons.

## Impact
This solves the user experience issue where the tool aggressively interrupted normal worksheet navigation (single clicks) and provides a safe way to exit the menu if triggered accidentally. It affects `ThisWorkbook.cls` (event hook), `Module_EnumSelector.bas` (launch logic), `Form_EnumSelect.frm` (UI and event handlers), and `inject_vba.py` (dynamic form generation).

</artifact>
