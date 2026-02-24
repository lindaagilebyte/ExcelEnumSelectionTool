<artifact id="specs" change="enum-selection-double-click" schema="spec-driven">

## Requirements
### `double-click-activation`
- **Description**: The tool requires a double-click on a target cell to launch, avoiding single-click interference.
- **Rules**:
  - Event must be hooked via `Workbook_SheetBeforeDoubleClick`.
  - The `Cancel = True` parameter must be used to stop the default cell edit mode from engaging on valid target cells.

### `explicit-confirmation`
- **Description**: The UserForm must include "Confirm" and "Cancel" buttons for deliberate data entry.
- **Rules**:
  - `lstEnums` click event only selects the visual item.
  - `btnConfirm` applies the selected value, logs `Application.OnUndo`, and Unloads.
  - `btnCancel` Unloads without applying.
  - Form UI is extended to fit these new buttons.

</artifact>
