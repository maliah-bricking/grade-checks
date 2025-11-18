# grade-checks

This repository contains an Excel template (macro-enabled) and helper scripts to build a workbook that:

- Maintains a StudentsTable with the following headers:
  - Student Name, Student Email, Advisors, Acad Year, Credits, Course No, Course Name, Student Current Grade, Section Average Grade, Average Daily Class Visits, Daily Class Visit, Grade
- Automatically removes duplicate rows and sorts the table by "Student Current Grade" ascending.
- Produces a filtered sheet (default name: `AtRisk`) containing only rows where `Student Current Grade` < 70 and only the columns: Student Name, Student Email, Advisors, Course No, Student Current Grade.
- Generates an outreach list to the right of the filtered table with two columns: "Students for Outreach" (unique student names, sorted A→Z) and "Email" (unique emails per student joined with `;`).

Files in this repo:
- `RemoveDuplicatesAndSort.bas` - the VBA module with the main routines.
- `SettingsSheet.md` - documentation for the Settings sheet layout and default values.
- `build_xlsm.ps1` - PowerShell script that creates a .xlsm workbook pre-populated with sheets, table, sample data, and the VBA module (requires Excel and trusted access to VB project object model).
- `sample_data.csv` - sample data to be imported into the StudentsTable.
- `.gitignore` - ignores common OS files.

Usage
1. Clone this repo locally.
2. If you want the ready .xlsm created automatically, run the PowerShell script `build_xlsm.ps1` from PowerShell on a Windows machine with Excel installed and with "Trust access to the VBA project object model" enabled.
3. Alternatively, open Excel, create a workbook, add the VBA module from `RemoveDuplicatesAndSort.bas` and create the `Settings` and `Data` sheets per `SettingsSheet.md` and import `sample_data.csv`.

Security note
- The PowerShell script programmatically writes VBA into the workbook. To allow this, Excel must be configured to allow programmatic access to the VBA project (Trust Center -> Macro Settings -> Trust access to the VBA project object model). Only run the script on a trusted machine.