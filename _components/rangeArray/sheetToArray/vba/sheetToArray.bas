Function sheetToArray(Optional ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    sheetToArray = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
End Function