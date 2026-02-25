# Tasks: Support Four Header Rows

## 1. Implement 4-Row Header Activation
- [ ] 1.1 In `TryLaunchEnumSelector`, change `Target.Row < 4` to `Target.Row < 5` to only activate on Row 5 and below.
- [ ] 1.2 In `TryLaunchEnumSelector`, change the Enum Key look up to read from `Target.Worksheet.Cells(3, Target.Column).Value` instead of Row 2.

## 2. Implement Data File Pre-Scan Logic
- [ ] 2.1 In `TryLaunchEnumSelector` (or `GetEnumList`/`RefreshCache`), introduce logic to scan the `ActiveWorkbook` before initializing the cache.
- [ ] 2.2 Create a helper function `GetRequiredEnumKeys()` that iterates all `Worksheets` in `ThisWorkbook` (the Data file).
- [ ] 2.3 Make the helper function skip worksheets whose name starts with `#`.
- [ ] 2.4 Make the helper function scan Row 3 of the `UsedRange` for non-empty string values and collect them into a Dictionary containing the Required Keys.

## 3. Implement Targeted Caching in Reference Scanner
- [ ] 3.1 Modify `ScanReferenceFile` to call `GetRequiredEnumKeys()` and store the required keys dictionary.
- [ ] 3.2 Update `ScanWorksheet` to accept this dictionary as a parameter (e.g., `ScanWorksheet ws, requiredKeys`).
- [ ] 3.3 Inside `ScanWorksheet`, before calling `ExtractColumnData`, add a condition to check if `keyName` exists in `requiredKeys`. If not, skip extraction and caching for that block.

## 4. Testing
- [ ] 4.1 Ask the user which Data file they want to use for testing, and inject the updated VBA code into it.
- [ ] 4.2 Verify double-clicking on Row 1-4 does nothing.
- [ ] 4.3 Verify double-clicking on Row 5 onwards with an Enum Key in Row 3 correctly opens the selector.
- [ ] 4.4 Verify with Debug prints enabled that the cache only contains the keys actually used in the Data file, confirming the memory optimization.
