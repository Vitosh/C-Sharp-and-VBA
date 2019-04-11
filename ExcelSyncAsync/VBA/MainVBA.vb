Public Sub TestMe()

    Dim i As Long
    If Worksheets.Count < 33 Then
        For i = 1 To 33
            ThisWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = i
        Next i
    End If

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        For i = 1 To 1234
            ws.Cells(i, 1) = i + CLng((0 - 1000 + 1) * Rnd + 1000)
        Next i
    Next ws

End Sub