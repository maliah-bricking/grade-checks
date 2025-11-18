<#
PowerShell helper to build the .xlsm workbook with the Data sheet, StudentsTable, Settings sheet, sample data and VBA module.

Requirements:
- Windows with Excel installed.
- PowerShell run with permissions to access COM objects.
- In Excel: Trust Center -> Macro Settings -> Trust access to the VBA project object model must be enabled.

Usage: Open PowerShell and run:
    .\build_xlsm.ps1 -OutPath .\grade-checks.xlsm
#>
param(
    [string]$OutPath = ".\grade-checks.xlsm",
    [string]$SampleCsv = ".\sample_data.csv",
    [string]$VbaModulePath = ".\RemoveDuplicatesAndSort.bas"
)

# Create Excel COM objects
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()

# Remove default sheets
while ($workbook.Sheets.Count -gt 0) { $workbook.Sheets.Item(1).Delete() }

# Create Data sheet
$wsData = $workbook.Worksheets.Add()
$wsData.Name = "Data";

# Load sample data from CSV if present, otherwise create headers
if (Test-Path $SampleCsv) {
    $csv = Import-Csv -Path $SampleCsv
    $headers = $csv[0].psobject.properties.name
    # Write headers
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $wsData.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }
    # Write rows
    $r = 2
    foreach ($row in $csv) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $wsData.Cells.Item($r, $c + 1).Value2 = $row.$($headers[$c])
        }
        $r++
    }
} else {
    $headers = @("Student Name","Student Email","Advisors","Acad Year","Credits","Course No","Course Name","Student Current Grade","Section Average Grade","Average Daily Class Visits","Daily Class Visit","Grade")
    for ($c = 0; $c -lt $headers.Count; $c++) { $wsData.Cells.Item(1, $c + 1).Value2 = $headers[$c] }
}

# Create Table
$usedRange = $wsData.Range($wsData.Cells.Item(1,1), $wsData.Cells.Item($wsData.UsedRange.Rows.Count, $headers.Count))
$listObj = $wsData.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $usedRange, $null, 1)
$listObj.Name = "StudentsTable"

# Create Settings sheet and populate defaults
$wsSettings = $workbook.Worksheets.Add()
$wsSettings.Name = "Settings"
$settings = @(
    @("Data sheet name","Data"),
    @("Table name","StudentsTable"),
    @("Sort column header","Student Current Grade"),
    @("Dedupe columns",""),
    @("Sort order","Ascending"),
    @("Run on change","Yes"),
    @("Filter sheet name","AtRisk"),
    @("Filter table name","AtRiskTable"),
    @("Filter columns","Student Name,Student Email,Advisors,Course No,Student Current Grade"),
    @("Filter column","Student Current Grade"),
    @("Filter threshold","70")
)
$r = 2
foreach ($pair in $settings) {
    $wsSettings.Cells.Item($r,1).Value2 = $pair[0]
    $wsSettings.Cells.Item($r,2).Value2 = $pair[1]
    $r++
}

# Add VBA module (requires trust access to VB project)
if (-not (Test-Path $VbaModulePath)) {
    Write-Host "VBA module file $VbaModulePath not found. Please place RemoveDuplicatesAndSort.bas next to this script." -ForegroundColor Yellow
} else {
    $vbaCode = Get-Content -Raw -Path $VbaModulePath
    $vbProject = $workbook.VBProject
    $mod = $vbProject.VBComponents.Add(1) # vbext_ct_StdModule
    $mod.Name = "RemoveDuplicatesAndSort"
    $mod.CodeModule.AddFromString($vbaCode)

    # Add Worksheet_Change code to Data sheet module
    $dataComp = $vbProject.VBComponents.Item($wsData.CodeName)
    $sheetCode = @"
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler
    Application.EnableEvents = False
    Dim runOnChange As String
    On Error Resume Next
    runOnChange = ReadSetting(ThisWorkbook.Worksheets("Settings"), "Run on change", "Yes")
    On Error GoTo ExitHandler
    If LCase(Trim(runOnChange)) <> "no" And LCase(Trim(runOnChange)) <> "false" Then
        CleanAndSortTableFromSettings
    End If
ExitHandler:
    Application.EnableEvents = True
End Sub
"@
    $dataComp.CodeModule.AddFromString($sheetCode)

    # Add Workbook_Open to ThisWorkbook module
    $thisComp = $vbProject.VBComponents.Item($workbook.VBProject.VBComponents.Item(1).Name) -ErrorAction SilentlyContinue
    # Ensure ThisWorkbook component exists
    $wbComp = $vbProject.VBComponents.Item("ThisWorkbook")
    $wbCode = @"
Private Sub Workbook_Open()
    On Error Resume Next
    CleanAndSortTableFromSettings
End Sub
"@
    $wbComp.CodeModule.AddFromString($wbCode)
}

# Save workbook
$workbook.SaveAs((Resolve-Path $OutPath).Path), 52 # xlOpenXMLWorkbookMacroEnabled
$workbook.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Created $OutPath"
