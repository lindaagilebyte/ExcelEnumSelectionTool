# Proposal: Support Four Header Rows

## Summary

Update the Excel Enum Selection Tool to support a 4-row header format in Data files, shifting the Enum Key definition to Row 3 and the data start to Row 5. Additionally, optimize the Reference file caching mechanism to only load Enum Keys that are actually used by the current Data file, preventing unnecessary memory bloat as the Reference file grows.

## Motivation

1. **Header Format Change**: The Excel templates have been updated from 3 header rows to 4 header rows. The VBA macro needs to be updated to match the new structure so that it triggers correctly on data rows and reads the Enum Key from the correct row.
2. **Performance/Memory Optimization**: The current cache generation logic (`ScanReferenceFile` in VBA) blindly reads and caches every single Enum definition from the Reference file, regardless of whether the opened Data file needs it. As the Reference file expands to cover more systems and tables, this becomes increasingly inefficient and can cause memory/performance issues.

## Capabilities

1. **4-Row Header Support**: The macro will read the Enum Key from Row 3 (previously Row 2) and will only activate for double-clicks on Row 5 or below (previously Row 4).
2. **Targeted Caching**: A pre-scan step will be added to the Data file initialization. It will scan Row 3 of all relevant sheets (ignoring those starting with `#`) to build a list of "Required Keys". The Reference file scanner will then only cache data for keys that exist in this required list, ignoring all other definitions.

## Impact

- **`Module_EnumSelector.bas`**: 
  - `TryLaunchEnumSelector`: Update row indices (Row 3 for key, Row 5 for trigger).
  - Add logic to scan the active Data Workbook for required keys before reading the Reference file.
  - `ScanWorksheet`: Modify to accept a `Dictionary` of required keys and only cache matched keys.
- **Backwards Compatibility**: This change assumes all Data files injected with this tool will now use the 4-row header format. Files using the old 3-row format will read the wrong row for the Enum Key.
