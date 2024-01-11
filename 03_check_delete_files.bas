Option Explicit

' ______________________________________________________________________________
Sub delete_csv_files()

Dim intCounter As Integer
Dim strFile As String
Dim varFiles As Variant

varFiles = Array("file_1.csv", "file_2.csv")

For intCounter = LBound(varFiles) To UBound(varFiles)
  strFile = ThisWorkbook.Path & "\" & varFiles(intCounter)
  If file_exists(strFile) Then
    SetAttr strFile, vbNormal
    Kill strFile
  End If
Next intCounter

End Sub

' ______________________________________________________________________________
Function file_exists(ByVal strFile As String) As Boolean
  file_exists = (Dir(strFile) <> "")
End Function
