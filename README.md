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

---

## 🔧 How to Use These Scripts

1. **Download the `.vba` files**:
   - [TabReferenceFinder_TabsOnly.vba](#)
   - [TabReferenceFinder_TabsAndCells.vba](#)

2. **Import into Excel**:
   - Open the VBA editor (`Alt + F11`).
   - Select **File > Import File...** and choose the `.vba` file.

3. **Run the Macro**:
   - Set the `tabName` variable in the code to your target tab.
   - Press `F5` to run the macro.

---

## 💡 Why Use These Scripts?

These tools are designed to:
- Simplify dependency analysis in large Excel workbooks.
- Save time by automating tedious manual checks.
- Minimize errors when modifying or deleting tabs in complex workbooks.

---

## 📂 Repository Goals

This repository aims to provide practical Excel VBA tools to address real-world problems encountered by professionals. Contributions and feedback are welcome to enhance and expand this collection!

---

## 📜 License

This project is open-source and licensed under the [MIT License](LICENSE).
