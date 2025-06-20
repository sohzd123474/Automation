# 🟨 Excel Highlighter Tool (UserForm Edition)

This Excel VBA utility lets you **highlight and populate cells** across one or many worksheets from a single, easy‑to‑edit `HighlighterSetting` sheet.  
Everything is driven from a pop‑up form—no code tweaks required after setup.

---

## ✨ Key Features
|  |  |
|--|--|
| 🔍 **Column‑letter search** &nbsp;| Put the column letter to search in **`HighlighterSetting!A1`** (e.g. `C`). The macro hunts that column in each selected sheet. |
| 🎨 **Colour transfer** &nbsp;| Each matched cell gets the fill colour stored in Column B of the setting sheet. |
| 📥 **Dynamic data insert** &nbsp;| Row 1 from **C1→** lists destination column letters. Rows 2↓ hold the values that will be dropped into those columns. |
| 🗂️ **Multi‑sheet update** &nbsp;| Select any number of open sheets via the UserForm and process them in one click. |
| ⚡ **No code edits** &nbsp;| Change the setting sheet, press the button—done. |

---

## 📄 HighlighterSetting Sheet Layout

| A (search col) | B (colour) | C → (target columns) |
|---------------|-----------|----------------------|
| **C** *(cell A1)* |   | **B** | **D** | **F** |
| Apple  | *(blue fill)*  | 1 | Cat | Yes |
| Banana | *(green fill)* | 2 | Dog | No  |

*Legend*

* **A1** – Column letter you want to search in every target sheet.  
* **Row 1 from C1→** – Letters of columns that will receive data (any number).  
* **Column A (row 2↓)** – Values to hunt for in the chosen search column.  
* **Column B** – Fill colours to copy to the matched cells.  
* **Columns C→** – Data that will be inserted into the columns named in Row 1.

---

## 📦 How to Use

1. **Open `Highlighter.xlsm`** (contains the macro and `HighlighterSetting` sheet).  
2. **Open the workbook(s)** you want to update.  
3. On the **`Main`** sheet click the **Highlighter** button *– or –* press **Alt + F8** and run:
   ```vb
   LaunchHighlighterForm
   ```
4. In the form:  
   - Choose the **workbook**  
   - Tick one or more **sheets** (Ctrl‑click for multi‑select)  
   - Click **Apply Highlights**

---

## 🧠 Logic Overview
1. Reads the **search column letter** from `HighlighterSetting!A1`.  
2. Finds that column in every selected sheet.  
3. Walks each value in Column A (row 2↓):  
   - Colours the matching cells with Column B’s fill colour.  
   - Inserts the extra values (Cols C→) into the columns named in Row 1.  
4. **No “first‑match only” limit**—every occurrence is processed.

---

## 🛠 Developer Notes
* **UserForm:** `frmHighlighterSelector`  
* **Launcher macro:**
  ```vb
  Sub LaunchHighlighterForm()
      frmHighlighterSelector.Show
  End Sub
  ```
* Runs entirely from **Highlighter.xlsm**—no external libraries required.

---

## 📋 Worked Example

### Setting Sheet
| A1 = **C** |   | C1 = **B** | D1 = **D** |
|------------|---|------------|------------|
| Apple      | 🔵 | 1 | Cat |
| Banana     | 🟢 | 2 | Dog |

### Target Sheet (before)
| A | B | **C (Items)** | D |
|---|---|---------------|---|
|   |   | Apple         |   |
|   |   | Banana        |   |

### Target Sheet (after)
| A | B | **C (Items)** | D  |
|---|---|---------------|----|
|   | 1 | *(blue)* Apple| Cat|
|   | 2 | *(green)*Banana|Dog|

---

## ✅ Requirements
- Excel for Windows / Mac **with VBA**  
- Macros enabled  
- Target sheets must contain the column letter specified in `A1`

---

## 🚧 Road‑map
- [ ] Optional **case‑insensitive** matching  
- [ ] **Preview / Dry‑run** mode  
- [ ] Automatic **undo / backup**  
- [ ] **Export** updated sheets to new files

---

## 📎 License
**Commercial – All rights reserved.**  
Redistribution or modification requires a commercial licence—contact the author for details.
