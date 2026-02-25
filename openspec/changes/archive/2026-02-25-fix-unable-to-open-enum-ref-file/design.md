# Design: Dynamic Temporary Unique Filenames for Enum Reference

## Context
The Excel Enum Selection Tool caches enumeration definitions by opening a reference file. To avoid SVN read-only lock issues, it copies the reference file to `%TEMP%\列舉定義(企劃用).xlsx` and opens it. However, if Excel crashes, the zombie process holds a lock on this specific file, preventing any subsequent use of the tool in other Excel instances because `FileSystemObject.CopyFile` or `Workbooks.Open` fails.

## Architecture / Approach
We will modify the `ScanReferenceFile` in `Module_EnumSelector.bas` to dynamically generate a unique temporary filename on every launch rather than using a static one.

1.  **Unique Filename Generation**:
    *   Instead of `REF_FILE_NAME` directly in `%TEMP%`, we'll generate a GUID or use a timestamp + random number.
    *   Example implementation using `Timer` and `Rnd` functions (which are natively available in VBA without API declarations): `Environ("TEMP") & "\EnumRef_" & Format(Now, "yyyymmdd_hhmmss") & "_" & Int((9999 - 1000 + 1) * Rnd + 1000) & ".xlsx"`
    *   This ensures every execution has a unique target path.

2.  **Resource Cleanup**:
    *   The code already contains cleanup logic (`fso.DeleteFile tempPath, True`) wrapped in `On Error Resume Next`.
    *   This gracefully handles the deletion of the temporary file after it is no longer needed.
    *   Because the filename is unique, a failure to delete (e.g., if the user somehow locked the specific temporary Excel file themselves while it was open) will only leak a single file into `%TEMP%`, and it will not block subsequent executions of the tool.

## Data Model / API
No API or Data Model changes. The only change is local to the temporary file path variable (`tempPath`) within the `ScanReferenceFile` routine.

## Alternatives Considered

1.  **Finding and Killing the Zombie Process**:
    *   *Rationale*: Overly aggressive and risky. Terminating `EXCEL.EXE` processes via Windows Management Instrumentation (WMI) might close instances the user is actively working on.

2.  **Attempting to clear the specific file lock**:
    *   *Rationale*: Cannot be done purely in VBA without complex Windows API calls. Windows doesn't natively provide a simple way to forcibly release file locks held by other applications.

3.  **Using a GUID via `CreateObject("Scriptlet.TypeLib").Guid`**:
    *   *Rationale*: It's a clean way to generate a GUID, but the timestamp + random approach is slightly faster and easier to debug since the time of creation is visible in the filename. Either is acceptable. We will use `Scriptlet.TypeLib` as it guarantees cross-instance uniqueness better than `Rnd`.

## Risks / Trade-offs
*   **Leakage in `%TEMP%`**: If Excel crashes frequently during the `ScanReferenceFile` routine before the `Cleanup` step is reached, orphaned `.xlsx` files will accumulate in the user's `%TEMP%` folder. This is a very minor trade-off compared to a complete application lock-out, as Windows automatically manages disk space and users can clean their temp folders.
