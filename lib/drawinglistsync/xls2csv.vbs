if WScript.Arguments.Count < 3 Then
    MsgBox "Error! Please specify all required arguments.", vbCritical, "HdM DT - Error"
    Wscript.Quit
End If

Dim excel, workbook, worksheetName, sheetExists, ws, sheetDrawingList

' Get the worksheet name from arguments
worksheetName = WScript.Arguments.Item(1)

' Create Excel application
On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Error! Could not start Excel application." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "HdM DT - Error"
    Wscript.Quit
End If
On Error Goto 0

' Open workbook
On Error Resume Next
Set workbook = excel.Workbooks.Open(Wscript.Arguments.Item(0))
If Err.Number <> 0 Then
    MsgBox "Error! Could not open the Excel file." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "HdM DT - Error"
    excel.Quit
    Wscript.Quit
End If
On Error Goto 0

' Check if worksheet exists
sheetExists = False
For Each ws In workbook.Sheets
    If ws.Name = worksheetName Then
        sheetExists = True
        Exit For
    End If
Next

If Not sheetExists Then
    Dim availableSheets, sheet
    availableSheets = ""
    For Each sheet In workbook.Sheets
        If availableSheets <> "" Then
            availableSheets = availableSheets & vbCrLf & "  - "
        Else
            availableSheets = "  - "
        End If
        availableSheets = availableSheets & sheet.Name
    Next
    
    MsgBox "Error! Worksheet '" & worksheetName & "' was not found in the Excel file." & vbCrLf & vbCrLf & _
           "Available worksheets are:" & vbCrLf & availableSheets & vbCrLf & vbCrLf & _
           "Please update your Drawing-List Sync configuration with the correct worksheet name.", _
           vbCritical, "HdM DT - Worksheet Not Found"
    
    workbook.Close False
    excel.Quit
    WScript.Quit
End If

' Process the worksheet
On Error Resume Next
Set sheetDrawingList = workbook.Sheets(worksheetName)
sheetDrawingList.Activate
workbook.SaveAs WScript.Arguments.Item(2), 6
workbook.Close False
excel.Quit

If Err.Number <> 0 Then
    MsgBox "Error! Could not save the worksheet as CSV." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "HdM DT - Error"
    WScript.Quit
End If