# 🟨 Excel Highlighter Tool (UserForm Edition)

This Excel VBA tool allows users to **highlight and populate cells** in multiple worksheets based on a centralized `HighlighterSetting` sheet using a **simple user-friendly form interface**.

---

## ✨ Features

- 🎯 Match items in Column A of target sheets
- 🎨 Apply cell background color based on your `HighlighterSetting`
- 📥 Dynamically insert values into specified columns
- ✅ Supports multi-sheet updates via a pop-up UserForm
- ⚡ No coding needed to update settings – just edit the sheet

---

## 📄 HighlighterSetting Sheet Format

| A (Items) | B (Color Cell) | C (A) | D (B) | E (C) | ... |
|-----------|----------------|-------|-------|-------|-----|
| Apple     | (filled blue)  | 1     | Cat   | Yes   | ... |
| Banana    | (filled green) | 2     | Dog   | No    | ... |

- **Column A**: Item names to match in your data sheets
- **Column B**: Cells with the background color to apply
- **Row 1 (from Column C onward)**: Excel column letters indicating where to insert values (e.g., `A`, `B`, `F`)
- **Rows 2+**: Values to insert for each matched item

---

## 📦 How to Use

1. **Open `Highlighter.xlsm`** containing the `HighlighterSetting` sheet.
2. **Open any target workbook(s)** you wish to apply formatting to.
3. Press `Alt + F8` and run:
   ```vb
   LaunchHighlighterForm
4. In the form:
   - Select the **workbook** to apply changes to
   - Select one or more **sheets** (multi-select supported)
   - Click **Apply Highlights**

---

## 🧠 Logic Overview

- Items in **Column A** of the `HighlighterSetting` sheet are matched against **Column A** of the target sheets.
- If matched:
  - The corresponding **fill color** (from **Column B**) is applied to the matched row’s Column A.
  - Values from Columns **C onward** are inserted into the corresponding columns, as defined by the **column letters in Row 1** (e.g., “B”, “D”, “F”).
- Matching is **case-sensitive** by default.

---

## 🛠 Developer Notes

- **Main UserForm**: `frmHighlighterSelector`
- **Launcher Macro**:
  ```vba
  Sub LaunchHighlighterForm()
      frmHighlighterSelector.Show
  End Sub
All logic runs from the workbook containing the `HighlighterSetting` sheet.

No external libraries required.

---

## 📋 Example

### HighlighterSetting Sheet

| A (Item) | B (Color) | C (A) | D (B) | E (C) |
|----------|-----------|--------|--------|--------|
| Apple    | (Blue)    | 1      | Cat    | Yes    |
| Banana   | (Green)   | 2      | Dog    | No     |

### Target Sheet (Before)

| A       | B | C | D |
|---------|---|---|---|
| Apple   |   |   |   |
| Banana  |   |   |   |

### Target Sheet (After Applying)

| A       | B | C   | D    |
|---------|---|-----|------|
| Apple   | 1 | Cat | Yes  |
| Banana  | 2 | Dog | No   |

---

## ✅ Requirements

- Microsoft Excel (with VBA support)
- Macros must be enabled
- Target sheets must contain item names in **Column A**

---

## 🚧 To-Do / Suggestions

- [ ] Add support for **case-insensitive** matching
- [ ] Allow **custom column** matching (not just Column A)
- [ ] Add **preview mode** before applying changes
- [ ] Support for **auto-exporting** modified sheets

---

## 📎 License

This software is not open source.  
Use, distribution, or modification requires a commercial license.  
For inquiries, contact the author.
