---
description: How to develop, test, and deploy VBA code in this project
---

# VBA Development Workflow

This repository uses a unique development pipeline to avoid corrupted Excel binaries and SVN lock issues. **Never edit the macro file directly using Excel's built-in Visual Basic Editor for long-term storage.**

Instead, all VBA code is stored as raw text files, and a Python script injects them into the target Excel file (`.xlsm` or `.xlsx`).

When modifying the VBA tool, always follow these exact steps:

1. **Modify the Source Code:**
   - Make all VBA code changes exclusively to the raw text files located in `Source/DataWorkbook/`:
     - `Module_EnumSelector.bas` (Main logic)
     - `Form_EnumSelect.frm` (UserForm design and events)
     - `ThisWorkbook.cls` (Workbook event hooks)

2. **Ask the User for a Test File:**
   - You **must** ask the USER which `.xlsm` or `.xlsx` file they want to use for testing before running any scripts.
   - Example: "Which Excel file would you like me to inject the new code into for testing?"

3. **Run the Injection Script:**
   - Once the user provides the target file path, run the python injection script from the project root.
   - Example command: `python inject_vba.py "<path_to_target_file>"`
   - If `.xlsm`: The script saves the file in-place, preserving worksheets.
   - If `.xlsx`: The script clones the file and creates a new `_MacroEnabled.xlsm`.

4. **Ask the User to Verify:**
   - Wait for the user to open the resulting `.xlsm` file in Excel and confirm that the new changes are working correctly before committing anything to Git or closing the task.
