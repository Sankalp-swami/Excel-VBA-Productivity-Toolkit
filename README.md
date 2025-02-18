# Excel-VBA-Productivity-Toolkit
Boost Excel efficiency with these VBA tools. Automate tasks, check dependencies, and streamline workflows.



Streamline your daily Excel workflows with a curated collection of powerful VBA tools. Designed to solve common workplace challenges, these scripts tackle everything from cross-sheet referencing to automation of tedious tasks. Ideal for analysts, engineers, and professionals seeking practical solutions to enhance efficiency in Excel. This collection simplifies complex or repetitive tasks, especially for those working with multi-tab spreadsheets. Open-source and continuously updated to address real-world problems, these tools are here to improve your productivity and make Excel work for you.

---

## 📜 Scripts Included

### 1. **Tab Reference Finder (List Tabs Only)**

**File Name**: `TabReferenceFinder_TabsOnly.vba`

**Description**:  
This script identifies and lists all the worksheet tabs in your Excel workbook that reference a specified tab. It is ideal for quickly determining dependencies between tabs without diving into individual cells.

**Usage**:
- Specify the tab name to check in the VBA code (variable `tabName`).
- Run the macro.
- The script will create a new sheet named `Referencing Tabs` and list all the tabs that reference the target tab.

---

### 2. **Tab Reference Finder (List Tabs and Cells)**

**File Name**: `TabReferenceFinder_TabsAndCells.vba`

**Description**:  
This script identifies all worksheet tabs that reference a specified tab and lists the exact cells containing those references. It is perfect for detailed impact analysis when modifying or deleting a tab.

**Usage**:
- Specify the tab name to check in the VBA code (variable `tabName`).
- Run the macro.
- The script will create a new sheet named `Detailed References` with a detailed list of tab names and cell addresses referencing the target tab.

3. Filter Rows by Highlighted Color in Pivot Tables
File Name: FilterRowsByColor.vba

Description:
This VBA function allows users to filter and analyze only the rows highlighted in a specific color (e.g., green) within an Excel dataset. Since Excel does not natively support filtering by cell color in Pivot Tables, this script provides a helper function to check whether an entire row is filled with a specific color. It then labels each row accordingly, enabling filtering in Pivot Tables or other data analysis.

Usage:

Download and Import the VBA Script:

Open the VBA Editor (ALT + F11).
Select File > Import File... and choose FilterRowsByColor.vba.
Apply the Function in Excel:

In a new column (e.g., "Row Color"), enter the formula:
excel
Copy
Edit
=IsRowGreen(A2:D2)
Adjust A2:D2 to match your dataset’s full row range.
Drag down to apply the formula across all rows.
Use in a Pivot Table:

Select the dataset (including the "Row Color" column).
Insert a Pivot Table (Insert > PivotTable).
Drag "Row Color" to the Filters area.
Filter the Pivot Table to show only "Green" rows.
Why Use This Script?
✅ Enables Pivot Table filtering based on row color.
✅ Automates color-based analysis without manual sorting.
✅ Helps in reviewing specific highlighted data efficiently.

4. Analyze Workbook References
File Name: AnalyzeWorkbookReferences.bas

Description:
This VBA script provides a comprehensive analysis of tab references in an Excel workbook. It identifies:

Tabs Referencing Each Tab: Lists all tabs that reference the current tab in their formulas.
Tabs Referenced by Each Tab: Lists all tabs referenced within the current tab's formulas.
The output is summarized in a new worksheet, "Workbook References Summary", with columns for the tab name, referencing tabs, and referenced tabs. Handles complex formulas with multiple tab references, self-references, and ensures compatibility with large workbooks.

🔧 How to Use These Scripts
Download the .vba files:
TabReferenceFinder_TabsOnly.vba
TabReferenceFinder_TabsAndCells.vba
FilterRowsByColor.vba
AnalyzeWorkbookReferences.bas
Import into Excel:
Open the VBA editor (Alt + F11).
Select File > Import File... and choose the .vba file.
Run the Macro:
Set the required parameters (like tabName) in the VBA code.
Press F5 to run the macro.
💡 Why Use These Scripts?
These tools are designed to:
✅ Simplify dependency analysis in large Excel workbooks.
✅ Save time by automating tedious manual checks.
✅ Minimize errors when modifying or deleting tabs in complex workbooks.
✅ Enable Pivot Table filtering based on row color.

📂 Repository Goals
This repository aims to provide practical Excel VBA tools to address real-world problems encountered by professionals. Contributions and feedback are welcome to enhance and expand this collection!

📜 License
This project is open-source and licensed under the MIT License.
