# Excel Enum Selection Tool

This tool automatically injects a custom VBA UserForm into Excel files to provide a dropdown selection interface for specific columns based on reference data. 

To improve the user experience and prevent accidental overwrites, the macro triggers on **Double-Click** and requires explicit confirmation.

> [!WARNING]
> **Undo History Limitation:** Modifying cells via VBA macros (like this tool) natively clears Excel's Undo history. We have implemented a custom `Ctrl+Z` patch, but it has specific limitations. Please read the **Undo Support** section below.

---

## Prerequisites
- Windows OS with Microsoft Excel installed.
- Python 3.x
- `pywin32` library (install via `pip install pywin32`)

---

## Part 1: Injecting the Macro into your Excel File

The `inject_vba.py` script automatically injects the necessary VBA code (UserForm and Modules) into your target Excel file. It handles both fresh `.xlsx` files and existing `.xlsm` files dynamically.

### Usage
Open your Command Prompt or Terminal, navigate to the project folder, and run:

```cmd
python inject_vba.py "<path_to_your_excel_file>"
```

### Behavior based on File Type:
- **If you provide a `.xlsx` file:** The script will clone the file, inject the macros, and save a brand new file named `<YourFile>_MacroEnabled.xlsm` in the same directory.
- **If you provide a `.xlsm` file:** The script will open the file, cleanly remove any old versions of the injected macros, inject the latest code, and save the file **in-place**. This perfectly preserves all your existing data, records, and tabs.

---

## Part 2: Using the Excel Dropdown Menu

Once you open your generated or updated `.xlsm` file in Excel, ensure that **Macros are Enabled**.

### 1. Triggering the Menu
- **Single-Clicking:** Clicking a cell normally will just highlight the cell. You can copy, paste, or drag without any menus popping up.
- **Double-Clicking:** Double-clicking an eligible cell (Row 4 or below, with a valid Enum Key in Row 2) will instantly launch the "Select Value" menu.

### 2. Making a Selection
When the menu opens:
1. Click an item in the list to highlight your choice.
2. Click **[確認] (Confirm)** to write the chosen value directly into the cell and close the menu.
3. Click **[取消] (Cancel)** to close the menu instantly without modifying the cell.

### 3. Undo Support
If you click Confirm by mistake, you can immediately hit `Ctrl + Z` (or click Undo in Excel's top bar) to instantly revert the cell back to its original value.

> [!IMPORTANT]
> **Understanding the Undo Patch Limitations**
> 
> Natively, running *any* VBA macro that modifies a worksheet instantly and permanently clears your entire Excel Undo stack. Because this tool uses a macro to write your selection into the cell, your previous history is naturally wiped out.
> 
> To mitigate this frustrating Excel quirk, this tool includes a custom `Ctrl+Z` Undo patch. **However, this patch only remembers the single, specific cell modification made by the dropdown menu.** 
> 
> **What this means for you:**
> 1. You **can** undo the immediate selection you just made with the tool.
> 2. You **cannot** press Undo multiple times to revert changes you made *before* using the tool. 
> 3. Your older history is still cleared every time you confirm a selection.

---

## Troubleshooting

- **"Cannot access VBA Project Object Model" Error:** 
  You must allow Python to interact with Excel's backend. In Excel, go to `File > Options > Trust Center > Trust Center Settings > Macro Settings` and check **"Trust access to the VBA project object model"**.

- **Console Hanging or Locked Files:**
  The Python injection script automatically forces background Excel processes to close upon completion. If you ever experience issues where the injection fails because a file is locked by a previously crashed session, you can manually clear your PC's memory by typing `taskkill /F /IM excel.exe` in your command prompt.
