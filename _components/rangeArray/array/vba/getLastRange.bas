Function getLastRange(ws As Worksheet) As Range
    Dim rngLastRow As Range
    Dim rngLastCol As Range

    Set rngLastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If rngLastRow Is Nothing Then
         Set getLastRange = Nothing
         Exit Function
    End If
    
    Set rngLastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    Set getLastRange = ws.Cells(rngLastRow.Row, rngLastCol.Column)

End Function
