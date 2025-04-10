Sub ClearAllFilters()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If ws.AutoFilterMode Then
                ws.AutoFilterMode = False
            End If
        End If
    Next ws
End Sub
