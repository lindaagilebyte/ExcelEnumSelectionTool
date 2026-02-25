# Spec: temp-file-generation

## Requirements
- **NEW**: The application MUST generate a collision-free unique temporary filename for each `ScanReferenceFile` execution.
- **NEW**: The unique filename MUST be located within the user's `%TEMP%` directory.
- **MODIFIED**: The tool MUST NOT use the static hardcoded `%TEMP%\列舉定義(企劃用).xlsx` path.

## Scenarios
### Scenario: Tool Execution
- **GIVEN**: A user triggers the Enum Selection Tool from an active Excel instance.
- **WHEN**: The `Module_EnumSelector.bas` invokes `ScanReferenceFile`.
- **THEN**: The script generates a unique path string (e.g., using `Environ("TEMP")` combined with a timestamp and random number).
- **AND**: The script successfully copies the target reference file to this unique path.
