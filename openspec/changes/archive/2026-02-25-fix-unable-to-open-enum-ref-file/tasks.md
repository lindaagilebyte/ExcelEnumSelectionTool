# Tasks: Fix Unable to Open Enum Reference File

## 1. Implement Dynamic Temporary Filename
- [x] 1.1 In `Module_EnumSelector.bas`, locate the `ScanReferenceFile` Subroutine.
- [x] 1.2 Modify the line `tempPath = Environ("TEMP") & "\" & REF_FILE_NAME` to generate a unique filename on every run.
- [x] 1.3 Use `Scriptlet.TypeLib` to generate a GUID, or `Timer`/`Rnd`, and append it to the filename (e.g., `Environ("TEMP") & "\" & Replace(Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36), "-", "") & ".xlsx"`).

## 2. Refine Resource Cleanup
- [x] 2.1 Verify that `fso.DeleteFile tempPath, True` is wrapped with `On Error Resume Next` and `On Error GoTo 0` to prevent unhandled exceptions during cleanup if the file is locked by the user.

## 3. Testing
- [ ] 3.1 Trigger the Enum Selection Tool from an Excel instance.
- [ ] 3.2 Verify that a unique temporary file is created in `%TEMP%` and no errors occur.
- [ ] 3.3 Verify that the unique temporary file is correctly deleted after the tool finishes loading the cache.
- [ ] 3.4 Open the generated temporary file manually (while the script is paused, or by inserting a breakpoint) to simulate a lock, and verify the cleanup error is swallowed gracefully without crashing the tool.
