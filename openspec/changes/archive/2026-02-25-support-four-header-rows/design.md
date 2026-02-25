# Design: Support Four Header Rows

## Context
The Excel Enum Selection Tool relies on VBA macros injected into Data files to offer dropdown choices from a shared Reference file. Data files are being updated from a 3-row to a 4-row header structure. The Enum Key indicator is moving from Row 2 to Row 3, and actual user data starts on Row 5 instead of 4.

Furthermore, the existing cache mechanism (`ScanReferenceFile` in `Module_EnumSelector.bas`) blindly loads all definitions from the Reference file. As the Reference file scales to support the entire game's databases, this approach will cause unacceptable memory consumption and initialization delays.

## Goals
- React smoothly to the new 4-row header format in Data files (Key on Row 3, trigger on Row 5).
- Drastically improve cache initialization performance by only loading keys that the active Data file actually uses.

## Non-Goals
- Changing the layout or format of the Reference file itself.
- Supporting mixed legacy files (3-headers) concurrently with the new 4-header files via dynamic detection (we assume all injected files going forward use the 4-row format).

## Decisions
1. **Hardcoded Row Shifts**: 
   - `TryLaunchEnumSelector` will strictly check `Target.Row < 5` and read the key from `Target.Worksheet.Cells(3, Target.Column).Value`. The hardcoded nature is acceptable as the Python injector handles distribution and versioning.
2. **Pre-scan for Used Keys**:
   - Before attempting to open the Reference file, the macro will iterate through all `Worksheets` in `ThisWorkbook` (the active Data file).
   - It will skip sheets whose names begin with `#` (metadata/instructions).
   - For valid sheets, it will scan Row 3 for non-empty string values. These string values will be added to a `Dictionary` object representing the "Required Items".
3. **Filter at the Source (`ScanWorksheet` in Reference File)**:
   - When scanning the Reference file, after resolving a `keyName` from the `"定義(巨集顯示)"` block, the code will check if `keyName` exists in the "Required Items" dictionary.
   - If it does, the column extraction proceeds, and it is cached. If it does not, the block is ignored.

## Risks / Trade-offs
- **Risk**: A Data file sheet might have a legitimate Enum Key in Row 3, but the Reference file might not have the definition.
  - **Mitigation**: The code will just not find the key in the Reference file, and therefore won't cache it. When the user double-clicks that key, `GetEnumList` will return `Null`, and the `TryLaunchEnumSelector` method will pop up the standard "找不到 [EnumName] 的資料定義" warning box, which is the exact same safe behavior as before.
- **Risk**: The pre-scan takes time on large Data files.
  - **Mitigation**: Scanning Row 3 of the active `UsedRange` is computationally trivial in memory compared to opening and parsing the entire remote external Reference workbook. The net performance gain will be massive.
