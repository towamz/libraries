Function getLastUsedRange(ws As Worksheet) As range
    Dim rngLastRow As range
    Dim rngLastCol As range

    Set rngLastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If rngLastRow Is Nothing Then
         Set getLastRange = Nothing
         Exit Function
    End If
    
    Set rngLastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    Set getLastUsedRange = ws.Cells(rngLastRow.Row, rngLastCol.Column)

End Function
