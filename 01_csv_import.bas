Option Explicit

' ______________________________________________________________________________
Sub import_csv_file_line_by_line()

' ______________________________________________________________________________
Dim objFileSystem As Object
Dim obj_csv_file As Object
Dim intLine As Integer
Dim strLine As String
Dim wkSheet As Worksheet
' ______________________________________________________________________________

Set wkSheet = Table1

On Error GoTo err_msg
With wkSheet
  .UsedRange.Clear
  
  Set objFileSystem = CreateObject("Scripting.FileSystemObject")
  Set obj_csv_file = objFileSystem.OpenTextFile(ThisWorkbook.Path & "\file.csv")
  intLine = 1
  
  Do Until obj_csv_file.AtEndOfStream
    strLine = obj_csv_file.ReadLine
    .Cells(intLine, 1).Value = strLine
    intLine = intLine + 1
  Loop
  
  obj_csv_file.Close
  .Columns("A:A").TextToColumns Destination:=.Range("A1"), _
    DataType:=xlDelimited, semicolon:=True
    
End With

Exit Sub

' Error Handler:
err_msg:
MsgBox "File not found." & " Error Number: " & Err.Number & " - " & Err.Description
End Sub
