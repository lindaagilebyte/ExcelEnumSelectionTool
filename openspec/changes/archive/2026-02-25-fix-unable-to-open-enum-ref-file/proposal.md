# Proposal: Fix Unable to Open Enum Reference File

## Goal

To resolve a fatal error ("無法開啟列舉參考檔") that occurs when the Excel Enum Selection Tool attempts to open a reference file. The problem restricts users from selecting enumerations in the target row when the static temporary reference file is locked by a zombie Excel process.

## Why

Currently, the `Module_EnumSelector.bas` code hardcodes the temporary path for the copied reference file as`%TEMP%\列舉定義(企劃用).xlsx`. When Excel crashes or doesn't exit cleanly, a zombie Excel process retains a read-lock on this file. When the user subsequently triggers the Enum Selection Tool from another Excel instance, `FileSystemObject.CopyFile` or `Workbooks.Open` inevitably fails due to the lock collision, throwing the "無法開啟列舉參考檔" error. Since this relies on a specific filename within `%TEMP%`, every instance of the tool will fail until the zombie process is killed or the computer is restarted. 

The fix will dynamically generate unique temporary filenames on each launch, avoiding such filesystem collisions.

## Capabilities

- `temp-file-generation`: The system needs a robust method to generate unique and collision-free temporary filenames when scanning the reference file instead of a hardcoded single target file.
- `temp-file-cleanup`: Unique temporary files should still be proactively cleaned up by the script to avoid `%TEMP%` bloat, while graceful failure on cleanup should be permitted if locked by the user.

## Impact

- **Affected code**: `ScanReferenceFile` in `DataWorkbook/Module_EnumSelector.bas`
- **Dependencies**: Native VBA `Scripting.FileSystemObject`, native VBA `Environ("TEMP")`, and VBA's random/guid generation functions.
- **Systems**: Modifying the way caching is initiated from the double-click event on the rows containing enum identifiers.
