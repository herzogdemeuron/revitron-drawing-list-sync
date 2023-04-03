If WScript.Arguments.Count < 3 Then
    WScript.Echo "Error! Please specify all required arguments."
    Wscript.Quit
End If

Dim excel
Set excel = CreateObject("Excel.Application")
Dim workbook
Set workbook = excel.Workbooks.Open(Wscript.Arguments.Item(0))
Dim sheetDrawingList
Set sheetDrawingList = workbook.Sheets(WScript.Arguments.Item(1))

' Copy the sheet to a new workbook
Dim newWorkbook
Set newWorkbook = excel.Workbooks.Add
sheetDrawingList.Copy newWorkbook.Worksheets(1)

' Save the new workbook to a temporary tab-delimited text file
Dim tempTextFile
tempTextFile = WScript.Arguments.Item(2) & ".tmp"
newWorkbook.SaveAs tempTextFile, -4158 ' xlTextWindows (-4158) is the constant for tab-delimited text file format
newWorkbook.Close False

' Read the content of the temporary text file
Dim fso, inputFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set inputFile = fso.OpenTextFile(tempTextFile, 1)
Dim fileContent
fileContent = inputFile.ReadAll
inputFile.Close

' Replace tab characters with commas and save the content to the CSV file
Dim outputFile
Set outputFile = fso.CreateTextFile(WScript.Arguments.Item(2), True)
outputFile.Write Replace(fileContent, vbTab, ",")
outputFile.Close

' Delete the temporary text file
fso.DeleteFile tempTextFile

workbook.Close False
excel.Quit