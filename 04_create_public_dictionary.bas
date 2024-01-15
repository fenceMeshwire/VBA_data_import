Option Explicit

' ________________________________________________________________________________________________
Public dictCodeRegions As Object

' ________________________________________________________________________________________________
Sub translate_regions()

Dim operations As New cls_operations

Dim varKey As Variant

CallByName operations, "get_country_dict", VbMethod

For Each varKey In dictCodeRegions.keys
  Debug.Print varKey, dictCodeRegions(varKey)
Next varKey

End Sub

' ________________________________________________________________________________________________
Public Function get_country_dict()

Dim lngRow As Long, lngRowMax As Long
Dim strKey As String, strItem As String

Set dictCodeRegions = CreateObject("Scripting.Dictionary")

With dict_country
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 2 To lngRowMax
    strKey = .Cells(lngRow, 1).Value    ' Country Code
    strItem = .Cells(lngRow, 2).Value   ' Country
    dictCodeRegions.Add Key:=strKey, Item:=strItem
  Next lngRow
End With

End Function
