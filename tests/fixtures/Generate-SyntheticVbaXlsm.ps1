# Generate-SyntheticVbaXlsm.ps1 — author tests/fixtures/synthetic-vba.xlsm via Excel COM.
#
# Why a script and not programmatic generation in test code?
#   DevExpress.Spreadsheet cannot author a VBA project. Real Excel can. We commit
#   the resulting .xlsm so tests don't need Excel installed at runtime; this script
#   only runs when a developer wants to regenerate or extend the fixture.
#
# Prereq:
#   File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings ->
#   "Trust access to the VBA project object model" must be ON.
#   (Dutch Excel: "Vertrouwen geven aan toegang tot het VBA-projectobjectmodel".)
#
# Locale note:
#   Sheet codenames are locale-dependent (Dutch Excel: "Blad1"). The Workbook
#   codename happens to stay "ThisWorkbook" across locales because that name is
#   a VBA-language built-in. Worksheet.CodeName / Workbook.CodeName return empty
#   on an unsaved workbook in COM, so we don't use them; instead we enumerate
#   VBProject.VBComponents directly and match the workbook component by name and
#   the remaining Type-100 component as the sheet. Standard- and class-module
#   names are set explicitly (Module1, Class1).

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'synthetic-vba.xlsm')
)

$ErrorActionPreference = 'Stop'

# Excel + VBE constants (we don't have access to the typed enums in PS).
$xlOpenXMLWorkbookMacroEnabled = 52
$vbext_ct_StdModule   = 1
$vbext_ct_ClassModule = 2
$vbext_ct_Document    = 100

$module1Source = @'
Option Explicit

Sub Main()
    Process Worksheets("Data").Range("A1")
End Sub

Sub Process(ByVal r As Range)
    r.Value = "x"
End Sub

Sub Variadic(ParamArray args() As Variant)
    Debug.Print UBound(args)
End Sub

Static Sub StatefulCount()
    Static n As Long
    n = n + 1
End Sub
'@

$thisWorkbookSource = @'
Option Explicit

Private Sub Workbook_Open()
    Main
End Sub
'@

$sheet1Source = @'
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Debug.Print Target.Address
End Sub
'@

$class1Source = @'
Option Explicit

Public Sub Greet(ByVal who As String)
    MsgBox "Hello, " & who
End Sub
'@

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    # Rename first sheet so Module1.Main can resolve Worksheets("Data") at runtime.
    $sheet1 = $wb.Worksheets.Item(1)
    $sheet1.Name = 'Data'

    try {
        $vbproj = $wb.VBProject
    } catch {
        throw "Could not access VBProject. Enable 'Trust access to the VBA project object model' in Excel's Trust Center -> Macro Settings, then retry. Underlying error: $($_.Exception.Message)"
    }

    # Identify the two pre-existing document VBComponents by enumeration:
    # - workbook component is always named "ThisWorkbook"
    # - the other Type-100 component is the sheet (locale-dependent name)
    $thisWbComp = $null
    $sheetComp  = $null
    foreach ($c in $vbproj.VBComponents) {
        if ($c.Type -ne $vbext_ct_Document) { continue }
        if ($c.Name -eq 'ThisWorkbook') { $thisWbComp = $c } else { $sheetComp = $c }
    }
    if ($null -eq $thisWbComp) { throw "Could not locate ThisWorkbook VBComponent." }
    if ($null -eq $sheetComp)  { throw "Could not locate sheet VBComponent." }
    Write-Host "Sheet codename:        $($sheetComp.Name)"
    Write-Host "Workbook component:    $($thisWbComp.Name)"

    [void]$sheetComp.CodeModule.AddFromString($sheet1Source)
    [void]$thisWbComp.CodeModule.AddFromString($thisWorkbookSource)

    # Add a standard module and inject code.
    $module1 = $vbproj.VBComponents.Add($vbext_ct_StdModule)
    $module1.Name = 'Module1'
    [void]$module1.CodeModule.AddFromString($module1Source)

    # Add a class module and inject code.
    $class1 = $vbproj.VBComponents.Add($vbext_ct_ClassModule)
    $class1.Name = 'Class1'
    [void]$class1.CodeModule.AddFromString($class1Source)

    $wb.SaveAs($OutputPath, $xlOpenXMLWorkbookMacroEnabled)
    Write-Host "Wrote: $OutputPath"
}
finally {
    if ($wb)    { try { $wb.Close($false) | Out-Null } catch {} ; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($excel) { try { $excel.Quit() }                catch {} ; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
