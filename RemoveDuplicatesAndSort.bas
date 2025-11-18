' VBA module: RemoveDuplicatesAndSort.bas
' Paste this into a standard Module (Insert -> Module) in the VBA editor.
Option Explicit

' Main routine: read settings from "Settings" sheet, remove duplicates, sort the table,
' and update the filtered sheet.
Public Sub CleanAndSortTableFromSettings()
    On Error GoTo ErrHandler
    Dim wsSettings As Worksheet
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long, idx As Long
    Dim cols() As Long
    Dim arrCols As Variant
    Dim dedupeColsSetting As String
    Dim sortColName As String, settingsDataSheetName As String, tableName As String
    Dim sortOrderSetting As String
    Dim lc As ListColumn
    Dim orderConst As XlSortOrder
    Dim foundCols As Long

    Const SETTINGS_SHEET_NAME As String = "Settings"

    ' Try to get the settings sheet
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo ErrHandler

    ' Read settings (use defaults when Settings sheet or entries are missing)
    If wsSettings Is Nothing Then
        settingsDataSheetName = "Data"
        tableName = "StudentsTable"
        sortColName = "Student Current Grade"
        dedupeColsSetting = ""
        sortOrderSetting = "Ascending"
    Else
        settingsDataSheetName = ReadSetting(wsSettings, "Data sheet name", "Data")
        tableName = ReadSetting(wsSettings, "Table name", "StudentsTable")
        sortColName = ReadSetting(wsSettings, "Sort column header", "Student Current Grade")
        dedupeColsSetting = ReadSetting(wsSettings, "Dedupe columns", "") ' comma-separated
        sortOrderSetting = ReadSetting(wsSettings, "Sort order", "Ascending")
    End If

    ' Get the data worksheet and table
    Set ws = ThisWorkbook.Worksheets(settingsDataSheetName)
    Set lo = ws.ListObjects(tableName)

    ' Build array of columns for RemoveDuplicates:
    If Trim(dedupeColsSetting) = "" Then
        ReDim cols(1 To lo.ListColumns.Count)
        For i = 1 To lo.ListColumns.Count
            cols(i) = i
        Next i
    Else
        arrCols = Split(dedupeColsSetting, ",")
        ReDim cols(1 To UBound(arrCols) + 1)
        idx = 1
        foundCols = 0
        For i = LBound(arrCols) To UBound(arrCols)
            Dim s As String
            s = Trim(arrCols(i))
            If s <> "" Then
                If IsNumeric(s) Then
                    Dim n As Long
                    n = CLng(s)
                    If n >= 1 And n <= lo.ListColumns.Count Then
                        cols(idx) = n
                        idx = idx + 1
                        foundCols = foundCols + 1
                    End If
                Else
                    ' Try to find by header name
                    On Error Resume Next
                    Set lc = lo.ListColumns(s)
                    On Error GoTo ErrHandler
                    If Not lc Is Nothing Then
                        cols(idx) = lc.Index
                        idx = idx + 1
                        foundCols = foundCols + 1
                    End If
                    Set lc = Nothing
                End If
            End If
        Next i
        If foundCols = 0 Then
            ' Fallback: use all columns
            ReDim cols(1 To lo.ListColumns.Count)
            For i = 1 To lo.ListColumns.Count
                cols(i) = i
            Next i
        Else
            ' Resize to actual used count
            ReDim Preserve cols(1 To idx - 1)
        End If
    End If

    ' Remove duplicates
    lo.Range.RemoveDuplicates Columns:=cols, Header:=xlYes

    ' Determine the column to sort by (by header name), fallback to first
    On Error Resume Next
    Set lc = lo.ListColumns(sortColName)
    On Error GoTo ErrHandler
    If lc Is Nothing Then
        Set lc = lo.ListColumns(1)
    End If

    ' Decide sort order
    If LCase(Left(Trim(sortOrderSetting), 1)) = "d" Then
        orderConst = xlDescending
    Else
        orderConst = xlAscending
    End If

    ' Apply sort
    lo.Sort.SortFields.Clear
    lo.Sort.SortFields.Add Key:=lc.DataBodyRange, _
        SortOn:=xlSortOnValues, Order:=orderConst, DataOption:=xlSortNormal
    With lo.Sort
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With

    ' Update filtered sheet after cleaning/sorting
    UpdateFilteredSheetFromSettings

    Exit Sub

ErrHandler:
    ' For debugging uncomment the next line
    ' MsgBox "CleanAndSortTableFromSettings error: " & Err.Description, vbExclamation
End Sub

' ------------------------------
' Filtered export routine
' Reads settings and creates/updates a sheet with a subset of columns and rows
' where a numeric column value is less than the threshold.
' Also adds a "Students for Outreach" and "Email" columns to the right of the filtered table
' containing unique Student Name values present in the filtered output table and their emails,
' sorted alphabetically by Student Name.
' ------------------------------
Public Sub UpdateFilteredSheetFromSettings()
    On Error GoTo ErrHandler
    Dim wsSettings As Worksheet
    Dim settingsDataSheetName As String, tableName As String
    Dim filterSheetName As String, filterTableName As String
    Dim filterColsSetting As String, filterColsArr As Variant
    Dim filterColumnName As String
    Dim filterThresholdSetting As String
    Dim threshold As Double
    Dim wsData As Worksheet, lo As ListObject
    Dim outWS As Worksheet
    Dim outLo As ListObject
    Dim i As Long, j As Long, rowsCopied As Long
    Dim sourceHeaders As Variant, outHeaders As Variant
    Dim colIndexes() As Long
    Dim dataArr As Variant, outArr() As Variant
    Dim val As Variant
    Dim rngBody As Range
    Dim lc As ListColumn

    Const SETTINGS_SHEET_NAME As String = "Settings"

    ' Get settings sheet (if absent we'll use defaults)
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo ErrHandler

    If wsSettings Is Nothing Then
        settingsDataSheetName = "Data"
        tableName = "StudentsTable"
        filterSheetName = "AtRisk"
        filterTableName = "AtRiskTable"
        filterColsSetting = "Student Name,Student Email,Advisors,Course No,Student Current Grade"
        filterColumnName = "Student Current Grade"
        filterThresholdSetting = "70"
    Else
        settingsDataSheetName = ReadSetting(wsSettings, "Data sheet name", "Data")
        tableName = ReadSetting(wsSettings, "Table name", "StudentsTable")
        filterSheetName = ReadSetting(wsSettings, "Filter sheet name", "AtRisk")
        filterTableName = ReadSetting(wsSettings, "Filter table name", "AtRiskTable")
        filterColsSetting = ReadSetting(wsSettings, "Filter columns", "Student Name,Student Email,Advisors,Course No,Student Current Grade")
        filterColumnName = ReadSetting(wsSettings, "Filter column", "Student Current Grade")
        filterThresholdSetting = ReadSetting(wsSettings, "Filter threshold", "70")
    End If

    ' Parse threshold
    If IsNumeric(Trim(filterThresholdSetting)) Then
        threshold = CDbl(Trim(filterThresholdSetting))
    Else
        threshold = 70
    End If

    ' Get data table
    Set wsData = ThisWorkbook.Worksheets(settingsDataSheetName)
    Set lo = wsData.ListObjects(tableName)

    ' Parse filter columns (by header names). Preserve the requested order.
    filterColsArr = Split(filterColsSetting, ",")
    ReDim colIndexes(1 To UBound(filterColsArr) + 1)
    For i = LBound(filterColsArr) To UBound(filterColsArr)
        Dim cname As String
        cname = Trim(filterColsArr(i))
        If cname <> "" Then
            On Error Resume Next
            Dim lcCol As ListColumn
            Set lcCol = lo.ListColumns(cname)
            On Error GoTo ErrHandler
            If Not lcCol Is Nothing Then
                colIndexes(i + 1) = lcCol.Index
            Else
                ' If a header not found, set index 0 to mark invalid
                colIndexes(i + 1) = 0
            End If
            Set lcCol = Nothing
        Else
            colIndexes(i + 1) = 0
        End If
    Next i

    ' Build array of output headers (use the exact names provided in filterColsArr)
    ReDim outHeaders(1 To UBound(filterColsArr) + 1)
    For i = LBound(filterColsArr) To UBound(filterColsArr)
        outHeaders(i + 1) = Trim(filterColsArr(i))
    Next i

    Application.ScreenUpdating = False

    ' Create or clear output sheet
    On Error Resume Next
    Set outWS = ThisWorkbook.Worksheets(filterSheetName)
    On Error GoTo ErrHandler
    If outWS Is Nothing Then
        Set outWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        outWS.Name = filterSheetName
    Else
        ' Clear contents and any existing tables
        Dim tbl As ListObject
        For Each tbl In outWS.ListObjects
            tbl.Delete
        Next tbl
        outWS.Cells.Clear
    End If

    ' If table has no data rows, create empty output table with headers and exit (still create outreach headers)
    If lo.DataBodyRange Is Nothing Then
        outWS.Range(outWS.Cells(1, 1), outWS.Cells(1, UBound(outHeaders))).Value = outHeaders
        Set outLo = outWS.ListObjects.Add(xlSrcRange, outWS.Range(outWS.Cells(1, 1), outWS.Cells(1, UBound(outHeaders))), , xlYes)
        outLo.Name = filterTableName
        ' Add Students for Outreach header (two columns: Name and Email)
        outWS.Cells(1, UBound(outHeaders) + 1).Value = "Students for Outreach"
        outWS.Cells(1, UBound(outHeaders) + 2).Value = "Email"
        GoTo Cleanup
    End If

    ' Read data into array for faster processing
    Set rngBody = lo.DataBodyRange
    dataArr = rngBody.Value ' 1-based 2D array: (row, col)

    ' Prepare output array with worst-case size (rows x outCols), then resize later
    Dim outCols As Long
    outCols = UBound(outHeaders)
    ReDim outArr(1 To UBound(dataArr, 1), 1 To outCols)
    rowsCopied = 0

    ' Find index of filterColumnName in source table
    Dim filterColIndex As Long
    On Error Resume Next
    Set lc = lo.ListColumns(filterColumnName)
    On Error GoTo ErrHandler
    If Not lc Is Nothing Then
        filterColIndex = lc.Index
    Else
        ' If not found, no rows match - create empty table
        filterColIndex = -1
    End If
    Set lc = Nothing

    ' Loop through rows and copy matching rows
    For i = 1 To UBound(dataArr, 1)
        Dim rawVal As Variant
        If filterColIndex = -1 Then
            ' no filter column found; skip all
        Else
            rawVal = dataArr(i, filterColIndex)
            ' Try numeric conversion
            Dim numVal As Double
            If IsNumeric(rawVal) Then
                numVal = CDbl(rawVal)
                If numVal < threshold Then
                    ' copy requested columns
                    rowsCopied = rowsCopied + 1
                    For j = 1 To outCols
                        Dim srcIndex As Long
                        srcIndex = colIndexes(j)
                        If srcIndex >= 1 Then
                            outArr(rowsCopied, j) = dataArr(i, srcIndex)
                        Else
                            outArr(rowsCopied, j) = "" ' header not found in source
                        End If
                    Next j
                End If
            End If
        End If
    Next i

    ' If no rowsCopied, create table with headers only (still create outreach headers)
    If rowsCopied = 0 Then
        outWS.Range(outWS.Cells(1, 1), outWS.Cells(1, outCols)).Value = outHeaders
        Set outLo = outWS.ListObjects.Add(xlSrcRange, outWS.Range(outWS.Cells(1, 1), outWS.Range(1, outCols)), , xlYes)
        outLo.Name = filterTableName
        outWS.Cells(1, outCols + 1).Value = "Students for Outreach"
        outWS.Cells(1, outCols + 2).Value = "Email"
    Else
        ' Write headers and data (resize to rowsCopied)
        outWS.Range(outWS.Cells(1, 1), outWS.Cells(1, outCols)).Value = outHeaders
        outWS.Range(outWS.Cells(2, 1), outWS.Cells(1 + rowsCopied, outCols)).Value = _
            Application.Index(outArr, Evaluate("ROW(1:" & rowsCopied & ")"), Evaluate("COLUMN(1:" & outCols & ")"))
        Set outLo = outWS.ListObjects.Add(xlSrcRange, outWS.Range(outWS.Cells(1, 1), outWS.Range(1 + rowsCopied, outCols)), , xlYes)
        outLo.Name = filterTableName

        ' Build unique Student Name -> Email(s) map from the filtered output table
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = vbTextCompare
        Dim outStudentColIndex As Long
        Dim outEmailColIndex As Long
        On Error Resume Next
        outStudentColIndex = outLo.ListColumns("Student Name").Index
        outEmailColIndex = outLo.ListColumns("Student Email").Index
        On Error GoTo ErrHandler

        If outStudentColIndex >= 1 Then
            Dim totalRows As Long
            totalRows = outLo.DataBodyRange.Rows.Count
            For i = 1 To totalRows
                Dim nm As String
                Dim em As String
                nm = Trim(CStr(outLo.DataBodyRange.Cells(i, outStudentColIndex).Value))
                If outEmailColIndex >= 1 Then
                    em = Trim(CStr(outLo.DataBodyRange.Cells(i, outEmailColIndex).Value))
                Else
                    em = ""
                End If
                If Len(nm) > 0 Then
                    If Not dict.Exists(nm) Then
                        If Len(em) > 0 Then
                            dict.Add nm, em
                        Else
                            dict.Add nm, ""
                        End If
                    Else
                        ' append email if new and non-empty
                        If Len(em) > 0 Then
                            Dim existing As String
                            existing = dict(nm)
                            If Len(existing) = 0 Then
                                dict(nm) = em
                            Else
                                ' check if em already present (case-insensitive)
                                Dim parts As Variant
                                parts = Split(existing, ";")
                                Dim already As Boolean: already = False
                                Dim p As Variant
                                For Each p In parts
                                    If StrComp(Trim(p), em, vbTextCompare) = 0 Then
                                        already = True
                                        Exit For
                                    End If
                                Next p
                                If Not already Then
                                    dict(nm) = existing & ";" & em
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If

        ' Sort the unique names alphabetically (case-insensitive)
        Dim keys As Variant
        keys = dict.Keys
        Dim kCount As Long
        kCount = dict.Count
        If kCount > 1 Then
            Dim a As Long, b As Long
            Dim tmpK As Variant
            For a = 0 To kCount - 2
                For b = a + 1 To kCount - 1
                    If StrComp(CStr(keys(a)), CStr(keys(b)), vbTextCompare) > 0 Then
                        tmpK = keys(a)
                        keys(a) = keys(b)
                        keys(b) = tmpK
                    End If
                Next b
            Next a
        End If

        ' Write the Students for Outreach header and the unique names + emails in the two columns immediately to the right of the table
        Dim outreachCol As Long
        outreachCol = outCols + 1
        outWS.Cells(1, outreachCol).Value = "Students for Outreach"
        outWS.Cells(1, outreachCol + 1).Value = "Email"
        If kCount > 0 Then
            For i = 0 To kCount - 1
                outWS.Cells(2 + i, outreachCol).Value = keys(i)
                outWS.Cells(2 + i, outreachCol + 1).Value = dict(keys(i))
            Next i
        End If
    End If

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    ' For debugging uncomment:
    ' MsgBox "UpdateFilteredSheetFromSettings error: " & Err.Description, vbExclamation
    Application.ScreenUpdating = True
End Sub

' Helper: find the label in column A of the settings sheet and return column B value or default
Public Function ReadSetting(ws As Worksheet, label As String, defaultValue As String) As String
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Columns(1).Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    If Not rng Is Nothing Then
        ReadSetting = CStr(ws.Cells(rng.Row, 2).Value)
    Else
        ReadSetting = defaultValue
    End If
End Function
