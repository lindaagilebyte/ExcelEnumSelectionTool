# Spec: temp-file-cleanup

## Purpose
TBD: Handles the delayed or graceful deletion of dynamic temporary reference files.

## Requirements

### Requirement: The temporary copy of the reference file MUST be deleted at the end of the `ScanReferenceFile` routine, regardless of success or failure.

#### Scenario: Normal Execution
- **GIVEN**: The Enum Selection Tool has successfully finished reading the temporary reference file.
- **WHEN**: The `sourceWb` is closed and resources are released.
- **THEN**: The script deletes the uniquely generated temporary file from `%TEMP%`.
- **AND**: The system `%TEMP%` directory is kept clean.

### Requirement: The deletion operation MUST fail gracefully (e.g., using `On Error Resume Next`) if the file cannot be deleted, preventing the tool from throwing an error over trivial cleanup failures.

#### Scenario: Interrupted Execution or Locked File
- **GIVEN**: An error happens during the parsing sequence or the user manually opens and locks the uniquely generated temporary reference file.
- **WHEN**: The VBA script attempts to delete the temporary file, or an unhandled exception causes early termination of the script.
- **THEN**: The cleanup operation will be skipped or silently fail, leaving a single unique file in the `%TEMP%` directory without crashing the main application flow.
