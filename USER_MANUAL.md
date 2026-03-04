# Unity Excel Enum Selection Tool: User Manual

> 🌐 **Language / 語言:** [English](#unity-excel-enum-selection-tool-user-manual) | [繁體中文 (Traditional Chinese)](#unity-excel-列舉選擇工具使用手冊)

## Overview
The Enum Selection Tool is an Excel macro designed to help game planners quickly and accurately select predefined string values (such as IdentityType, PassiveAbilityType, or MissonType) from a dropdown menu while working in their data configuration spreadsheets. 

To provide these dropdown options, the tool reads from a central reference file named `列舉定義(企劃用).xlsx`. This manual explains how the tool scans that file so you can safely add, modify, or remove enumeration lists without breaking the system.

## How it Works (The Magic)
The tool does **not** rely on hardcoded row numbers or column letters. Instead, it uses **Spatial Scanning**. It searches all sheets in `列舉定義(企劃用).xlsx` for a specific anchor text:

**`定義(巨集顯示)`** (Definition (Macro Display))

Whenever the tool finds this exact text, it looks at the cells immediately surrounding it to figure out the **Enum Key** (the name of the list) and the **Enum Values** (the items in the dropdown).

### The Spatial Layout Rule
1. **The Anchor**: `定義(巨集顯示)`
2. **The Enum Key**: Must be placed exactly **one row above** and **one column to the left** of the anchor.
3. **The Enum Values**: Must start exactly **one row below** the anchor, in the same column as the anchor, going downwards.

```text
      [Column A]        [Column B]
[10]  Enum_Key_Name     (Empty or other text)
[11]  (Empty)           定義(巨集顯示)       <-- THE ANCHOR
[12]  (Empty)           Value_1
[13]  (Empty)           Value_2
[14]  (Empty)           Value_3
[15]  (Empty)           (Empty Cell)        <-- TOOL STOPS READING HERE
```

## Step-by-Step Guide: Adding a New Enum
To add a new dropdown list:

1. Open `列舉定義(企劃用).xlsx`.
2. Find an empty space on any sheet.
3. Type `定義(巨集顯示)` in a cell (e.g., cell `C5`).
4. Type your new Enum Key name diagonally above and to the left of your anchor (e.g., cell `B4`). This is the ID the data sheets will look for.
5. List your dropdown values starting directly below the anchor (e.g., starting at `C6` and going down to `C7`, `C8`, etc.).
6. **Important**: Do not leave any blank rows inside your list of values. The tool stops reading as soon as it hits an empty cell.
7. Save the file.

## Troubleshooting
- **Dropdown is empty or missing**: Check if you misspelled `定義(巨集顯示)`. It must be an exact match (no extra spaces).
- **Values are missing at the bottom**: Check if you accidentally left a blank row in the middle of your list.
- **Wrong list is showing up**: Check if your Enum Key name perfectly matches what is written in Row 3 of your data sheet.
- **Tool isn't finding it at all**: Make sure the Enum Key is in the correct diagonal position (top-left) relative to the anchor text.

---

# Unity Excel 列舉選擇工具：使用手冊

> 🌐 **Language / 語言:** [English](#unity-excel-enum-selection-tool-user-manual) | [繁體中文 (Traditional Chinese)](#unity-excel-列舉選擇工具使用手冊)

## 概述
列舉選擇工具是一個 Excel 巨集，旨在幫助遊戲企劃在編輯資料設定表時，能夠快速且準確地從下拉選單中選擇預先定義好的字串值（例如 IdentityType、PassiveAbilityType 或 MissonType 等）。

為了提供這些下拉選項，工具會讀取一個名為 `列舉定義(企劃用).xlsx` 的中央參考檔。本手冊將說明工具是如何掃描該檔案的，以便您能夠安全地新增、修改或移除列舉清單，而不會破壞系統運作。

## 運作原理（黑魔法）
本工具**不**依賴寫死的行號或欄位代碼。相反地，它使用**空間掃描 (Spatial Scanning)**。它會搜尋 `列舉定義(企劃用).xlsx` 內所有工作表中特定的定位點文字（Anchor Text）：

**`定義(巨集顯示)`**

每當工具找到這段完全一致的文字時，它就會觀察其周圍的儲存格，來判斷**列舉鍵值 (Enum Key)**（清單的名稱）以及**列舉值 (Enum Values)**（下拉選單中的項目）。

### 空間配置規則
1. **定位點**：`定義(巨集顯示)`
2. **列舉鍵值**：必須精準放置在定位點的**左上方**（上一列、左一欄）。
3. **列舉值**：必須從定位點的**正下方**開始（下一列、同一欄），並持續往下填寫。

```text
      [A 欄]            [B 欄]
[10]  列舉鍵值名稱        (空白或其他文字)
[11]  (空白)            定義(巨集顯示)       <-- 定位點
[12]  (空白)            數值_1
[13]  (空白)            數值_2
[14]  (空白)            數值_3
[15]  (空白)            (空白儲存格)         <-- 工具讀到這裡就會停止
```

## 逐步指南：新增一個列舉清單
若要新增一個下拉選單：

1. 打開 `列舉定義(企劃用).xlsx`。
2. 在任意工作表中找一個空白區域。
3. 在一個儲存格中輸入 `定義(巨集顯示)`（例如：`C5` 儲存格）。
4. 在定位點的左上方儲存格輸入您新的列舉鍵值名稱（例如：`B4` 儲存格）。這是資料表將會對應尋找的 ID。
5. 從定位點正下方的儲存格開始填寫您的下拉選單數值（例如：從 `C6` 開始往下的 `C7`, `C8` 等等）。
6. **注意**：在您的數值清單中，請勿留下任何空白列。工具只要一讀到空白儲存格就會立刻停止讀取。
7. 儲存檔案。

## 常見問題與排除
- **下拉選單是空的或不見了**：請檢查您是否拼錯了 `定義(巨集顯示)`。必須完全一致（不可有多餘的空白）。
- **最下面的數值不見了**：請檢查您是否不小心在清單中間留了空白列。
- **出現錯誤的清單**：請檢查您的列舉鍵值名稱，是否與資料表第 3 列所寫的完全相符。
- **工具完全找不到資料**：請確認您的列舉鍵值，相對於定位點文字，是否放在正確的對角線位置（左上方）。
