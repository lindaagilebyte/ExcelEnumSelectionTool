# Specs: Support Four Header Rows

## ADDED Requirements
These are new capabilities introduced by this change.

### Requirement: Targeted Caching Optimization
openspec/specs/vba-tool/spec.md => targeted-caching

The tool MUST optimize reference file caching by only loading data for Enum Keys that are actively used in the current Data file.

#### Scenario: Data File Initialization Pre-scan
- **GIVEN** a Data file containing multiple worksheets
- **WHEN** the tool initializes its Enum cache (e.g., on first double-click)
- **THEN** it MUST first scan Row 3 of all worksheets in the active Data file
- **AND** it MUST ignore worksheets whose names begin with `#`
- **AND** it MUST compile a unique list of all non-empty string values found in Row 3 of the valid sheets
- **AND** it MUST only extract and cache Enum definitions from the Reference file if their keyName matches a value in this compiled list.

## CHANGED Requirements
These modify existing behavior.

### Requirement: 4-Row Header Trigger Activation
openspec/specs/vba-tool/spec.md => four-row-header-trigger

The tool MUST activate its selection interface based on the new 4-row header layout, where data begins on Row 5.

#### Scenario: Double-Click on Data Row
- **GIVEN** a valid Data file with 4 header rows
- **WHEN** the user double-clicks a cell in Row 5 or below
- **AND** Row 3 of that same column contains a valid Enum Key
- **THEN** the Enum Selection menu MUST appear for that cell.

#### Scenario: Double-Click on Header Row
- **GIVEN** a valid Data file with 4 header rows
- **WHEN** the user double-clicks a cell in Row 1, Row 2, Row 3, or Row 4
- **THEN** the Enum Selection menu MUST NOT appear, allowing native Excel double-click behavior.

### Requirement: 4-Row Header Key Resolution
openspec/specs/vba-tool/spec.md => four-row-header-resolution

The tool MUST read the Enum Key from the correct row (Row 3) to determine which list to display.

#### Scenario: Resolving the Enum Key
- **GIVEN** the user double-clicks a valid data cell
- **WHEN** the tool attempts to look up the required Enum list
- **THEN** it MUST read the string value specifically from Row 3 of the target cell's column.
