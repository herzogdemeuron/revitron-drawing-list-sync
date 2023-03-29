if WScript.Arguments.Count < 3 Then
    WScript.Echo "Error! Please specify all required arguments."
    Wscript.Quit
End If
Dim excel
Set excel = CreateObject("Excel.Application")
Dim workbook
Set workbook = excel.Workbooks.Open(Wscript.Arguments.Item(0))
Set sheetDrawingList = workbook.Sheets(WScript.Arguments.Item(1))
sheetDrawingList.Activate
workbook.SaveAs WScript.Arguments.Item(2), 6
workbook.Close False
excel.Quit
