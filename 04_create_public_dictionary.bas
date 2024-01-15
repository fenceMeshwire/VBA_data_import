Option Explicit

' Worksheets required: dict_country, queue, result
' ________________________________________________________________________________________________
Public dictCodeRegions As Object

' ________________________________________________________________________________________________
Sub translate_regions()

Dim operations As New cls_operations

CallByName operations, "copy_used_range", VbMethod
CallByName operations, "get_country_dict", VbMethod
CallByName operations, "translate_expressions", VbMethod

End Sub

' ________________________________________________________________________________________________
Public Function copy_used_range()

Dim intCol As Integer, intColMax As Integer

result.UsedRange.Clear
queue.UsedRange.Copy Destination:=result.Range("A1")

intColMax = result.UsedRange.Columns.Count

For intCol = 1 To intColMax
  result.Columns(intCol).ColumnWidth = 50
Next intCol

End Function
' ________________________________________________________________________________________________
Public Function get_country_dict()

Dim lngRow As Long, lngRowMax As Long
Dim strKey As String, strItem As String

Set dictCodeRegions = CreateObject("Scripting.Dictionary")

With dict_country
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 2 To lngRowMax
    strKey = .Cells(lngRow, 1).Value    ' Country Code
    strItem = .Cells(lngRow, 2).Value   ' Country Name
    dictCodeRegions.Add Key:=strKey, Item:=strItem
  Next lngRow
End With

End Function
' ________________________________________________________________________________________________
Public Function translate_expressions()

Dim intCounter As Integer
Dim lngRow, lngRowMax As Long
Dim strKey As String, strValue As String
Dim strExpression As String, strResult As String
Dim strMatch As String
Dim varKey As Variant, varExpressions As Variant

With result

  lngRowMax = .UsedRange.Rows.Count
  
  For lngRow = 3 To lngRowMax
    
    strExpression = .Cells(lngRow, 1).Value
    varExpressions = Split(strExpression, "/")
    
    For intCounter = LBound(varExpressions) To UBound(varExpressions)
      strKey = varExpressions(intCounter)
      
      For Each varKey In dictCodeRegions.keys
        If varKey = strKey Then
          strMatch = dictCodeRegions(varKey)
          Exit For
        End If
        If strKey = "*" Then  ' In case the key equals "*"
          strMatch = "*"      ' The match is equal to "*"
          Exit For
        End If
    
      Next varKey
      
      If strResult = "" Then
        strResult = strMatch
      Else
        strResult = strResult & "/" & strMatch
      End If
      
    Next intCounter
    
          strKey = ""
          strMatch = ""

    result.Cells(lngRow, 1).Value = strResult
    strResult = ""
  Next lngRow
  
End With

End Function
