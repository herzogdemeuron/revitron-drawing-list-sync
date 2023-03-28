If WScript.Arguments.Count < 3 Then
    WScript.Echo "Error! Please specify all required arguments."
    Wscript.Quit
End If

Dim excel
Set excel = CreateObject("Excel.Application")

Dim workbook
Set workbook = excel.Workbooks.Open(Wscript.Arguments.Item(0))
Set sheetDrawingList = workbook.Sheets(WScript.Arguments.Item(1))
sheetDrawingList.Activate

Dim xlCellTypeConstants, xlDate
xlCellTypeConstants = 2
xlDate = 8

On Error Resume Next ' Start error handling

Dim dateCells
Set dateCells = sheetDrawingList.UsedRange.SpecialCells(xlCellTypeConstants, xlDate)

If Err.Number = 0 Then ' No error occurred
    Dim shell, shortDateFormat
    Set shell = CreateObject("WScript.Shell")
    shortDateFormat = shell.RegRead("HKEY_CURRENT_USER\Control Panel\International\sShortDate")

    dateCells.NumberFormat = shortDateFormat
Else
    Err.Clear ' Clear the error if any cells were not found
End If

On Error GoTo 0 ' End error handling

workbook.SaveAs WScript.Arguments.Item(2), 6
workbook.Close False
excel.Quit