Sub open_worksheet()

Dim strFile As String

With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Please select a WorkBook"
    .InitialFileName = ThisWorkbook.Path
    .Filters.Add "Workbook", "*.xls*", 1
    If .Show = -1 Then
        strFile = .SelectedItems(1)
    End If
End With

If strFile <> "" Then
    Workbooks.Open strFile
End If

End Sub
