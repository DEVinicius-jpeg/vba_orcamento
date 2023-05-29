Dim qtevent As qtclass
Private Sub Workbook_Open()

    Set qtevent = New qtclass

    Set qtevent.HookedTable = ThisWorkbook.Worksheets("Estoque").ListObjects(1).QueryTable

End Sub
