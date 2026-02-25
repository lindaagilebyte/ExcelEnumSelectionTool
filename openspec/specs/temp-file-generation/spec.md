# Spec: temp-file-generation

## Purpose
TBD: Handles generation of unique temporary file names for reference files to prevent file locking issues in typical Excel workflows.

## Requirements

### Requirement: The application MUST generate a collision-free unique temporary filename for each `ScanReferenceFile` execution.

#### Scenario: Tool Execution
- **GIVEN**: A user triggers the Enum Selection Tool from an active Excel instance.
- **WHEN**: The `Module_EnumSelector.bas` invokes `ScanReferenceFile`.
- **THEN**: The script generates a unique path string (e.g., using `Environ("TEMP")` combined with a timestamp and random number).
- **AND**: The script successfully copies the target reference file to this unique path.

### Requirement: The unique filename MUST be located within the user's `%TEMP%` directory.

### Requirement: The tool MUST NOT use the static hardcoded `%TEMP%\列舉定義(企劃用).xlsx` path.
