# Settings sheet: layout and sample values (updated for filtered export)

Create a worksheet called exactly: Settings

Layout (column A = label, column B = value). Enter these exact labels in column A:

A2: Data sheet name
B2: Data

A3: Table name
B3: StudentsTable

A4: Sort column header
B4: Student Current Grade

A5: Dedupe columns
B5: (leave blank OR enter comma-separated indexes or header names)
    Examples:
      - leave blank  -> dedupe by whole row (all columns)
      - "1,3"        -> dedupe using columns 1 and 3 of the table
      - "Email,Name" -> dedupe using the columns named Email and Name

A6: Sort order
B6: Ascending   (or "Descending")

A7: Run on change
B7: Yes         (Yes/No — controls whether Worksheet_Change triggers the cleaner)

A8: Filter sheet name
B8: AtRisk      (sheet that will contain filtered subset)

A9: Filter table name
B9: AtRiskTable

A10: Filter columns
B10: Student Name,Student Email,Advisors,Course No,Student Current Grade
     (comma-separated header names in the order you want them on the filtered sheet)

A11: Filter column
B11: Student Current Grade
     (the column used for numeric comparison)

A12: Filter threshold
B12: 70
     (rows where Filter column < Filter threshold are included)

Notes:
- Labels in column A must match exactly (case-insensitive) for ReadSetting to find them.
- Filter columns must match table headers exactly. The code copies columns in the order listed in A10.
- If Filter column is missing or non-numeric values are present, those rows are skipped.