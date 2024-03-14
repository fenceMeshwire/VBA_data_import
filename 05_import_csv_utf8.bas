Option Explicit

' ____________________________________________________________________________
Sub import_csv()

Dim intCounter As Integer, intRow As Integer
Dim objStream As Object
Dim strFileName As String
Dim strLineFromCSV As String
Dim varLineItems As Variant
Dim wkSheet As Worksheet

Set wkSheet = Sheet1

strFileName = "C:\Users\...\import.csv"

Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = "utf-8"
objStream.Open
objStream.LoadFromFile (strFileName)

With wkSheet
  .Rows.Delete
  intRow = 1
  
  Do Until objStream.EOS
      strLineFromCSV = objStream.ReadText(-2)
      varLineItems = Split(strLineFromCSV, ";")
      
      For intCounter = LBound(varLineItems) To UBound(varLineItems)
        .Cells(intRow, intCounter + 1).Value = varLineItems(intCounter)
      Next intCounter
  
      intRow = intRow + 1
  Loop
  
  .Columns.AutoFit
End With

Set objStream = Nothing

End Function
