🟨 Excel Highlighter Tool — UserForm Edition
This tool lets you highlight and populate values in one or more open Excel sheets using a master HighlighterSetting sheet and a simple pop-up form.

🧩 Features
🎨 Highlight cells based on matching item names

📥 Insert values into any columns defined by the user

✅ Point-and-click selection of workbooks and sheets

💡 Dynamically driven by the contents of the HighlighterSetting sheet — no code changes needed to add more fields

📑 HighlighterSetting Sheet Structure
Located in your main workbook (e.g. Book1.xlsm):

A (Item)	B (Color Cell)	C	D	E	...
Apple	(filled blue)	1	Cat	Yes	
Banana	(filled green)	2	Dog	No	
...	...	...	...	...	

Row 1 (from column C onward): Target column letters in the destination sheets (A, B, C, etc.)

Rows 2+:

Column A: Item names to match (in Column A of destination sheet)

Column B: Cell fill color to apply to matched item row

Columns C+: Values to insert into the columns specified in Row 1

🧑‍💻 How to Use
Set up your HighlighterSetting sheet:

Fill item names in Column A

Fill color cells in Column B

Type column letters in Row 1 (C1, D1, E1, etc.)

Below each, type the value to insert for each item

Open your data workbook(s) — ensure the sheets you want to update are visible.

Press Alt + F8 and run:

vb
Copy
Edit
LaunchHighlighterForm
In the UserForm:

Select the workbook you want to update

Ctrl+click the sheets you want to process

Click Apply Highlights

⚙️ Requirements
The macro must be run from the workbook containing the HighlighterSetting sheet

Target sheets must have data in column A (item names)

Column letters in Row 1 (C1 onward) must match valid Excel columns in the target sheets

✅ Example
Suppose your target sheet has:

A	B	C	D	E
Apple				
Banana				

And your HighlighterSetting sheet has:

A (Item)	B (Color)	C (A)	D (B)	E (C)
Apple	(Blue)	1	Cat	Yes
Banana	(Green)	2	Dog	No

Then clicking “Apply Highlights” will:

Fill column A with the color

Write 1, Cat, and Yes into columns A, B, and C respectively — on the same row as the matched item

🛠 Developer Notes
Form name: frmHighlighterSelector

Launch macro: LaunchHighlighterForm

Populates dynamically based on sheet structure; no changes needed if columns are added

Uses Range(colLetter & "1").Column to resolve actual target columns
