# Requirement: Trigger Logic

## Context
The tool monitors user clicks in the Data Workbook (e.g., `Contacts.xlsx`) and decides whether to show the selection form.

## Trigger Conditions

### Requirement: Row Validation
The tool **MUST** only activate when the user edits data rows, not headers.

#### Scenario: Selection in Data Area
- **WHEN** a cell is selected (ActiveCell)
- **THEN** check if `ActiveCell.Row >= 4`
- **AND** if true, proceed to Column Validation
- **AND** if false (Row 1-3), do nothing (exit immediately)

### Requirement: Column Validation (Enum Key)
The tool **MUST** uses **Row 2** of the active column to determine the "Enum Key".

#### Scenario: Check Enum Key
- **WHEN** the Row validation passes
- **THEN** read the value of `Cells(2, ActiveCell.Column)` (Row 2 Header)
- **AND** sanitize the value (Trim whitespace)
- **AND** check if this value exists in the `Reference Cache` (or Reference File)

#### Scenario: Valid Key Found
- **WHEN** the Row 2 value matches a known Enum Definition
- **THEN** display the UserForm
- **AND** populate the list with values for that Key

#### Scenario: Invalid or Missing Key
- **WHEN** the Row 2 value is empty OR does not exist in the Reference File
- **THEN** do nothing (allow normal Excel behavior)
