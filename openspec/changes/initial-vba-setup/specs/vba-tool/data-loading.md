# Requirement: Data Loading (Reference File)

## Context
The tool needs to load valid enum definitions from a central reference file (`列舉定義(企劃用).xlsx`). The definitions are scattered across multiple sheets and layout blocks.

## Retrieval Strategy

The system **MUST** minimize file I/O by caching definitions in memory.

### Requirement: Caching Scope
- **Scope:** The cache exists in memory for the duration of the **Excel Session** (while the Data Workbook is open).
- **Persistence:** It is NOT saved to disk. Closing and reopening the workbook clears the cache.

### Requirement: Cold Start (First Execution in Session)
- **WHEN** the user triggers the tool for the first time after opening the workbook
- **OR** when the VBA project has been reset (e.g. after a crash or manual reset)
- **THEN** scan `The Reference File` entirely
- **AND** iterate through all sheets
- **AND** index every "Enum Key" found in **Row 2** (or the defined Header Row) of each block
- **AND** store the `SheetName` and `CellAddress` for each Key

### Requirement: Warm Start (Subsequent Runs)
- **WHEN** the tool is triggered with a Key
- **THEN** check the in-memory cache
- **AND** if the Key exists, jump directly to the target sheet/cell to read data
- **AND** DO NOT re-scan the file

#### Scenario: Finding Definition Column
- **WHEN** the Key cell is found (e.g., at `A10`)
- **THEN** look for the sub-header `定義(巨集顯示)` in the same block (e.g., `B11` or `A11`)
- **AND** identify the column index of that sub-header as the "Data Column"

#### Scenario: Extracting List
- **WHEN** the "Data Column" is identified
- **THEN** read values starting from the cell **below** the sub-header
- **AND** stop reading when an empty cell or a new block header is encountered
- **AND** ignore any cells that are empty strings

## Performance
- **WHEN** the reference file is loaded
- **THEN** cache the mapping `Key -> {Sheet, DataColumn, RowStart}` in a `Scripting.Dictionary`
- **AND** reuse this cache for subsequent clicks to avoid re-scanning
