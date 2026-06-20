
Function sliceArray(rng As Range, Optional ws As Worksheet) As Variant
    Dim aryData As Variant
    Dim resultAry As Variant
    Dim rowOffset As Long
    Dim colOffset As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim j As Long
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    aryData = sheetToArray(ws)

    rowOffset = rng.Row - 1
    colOffset = rng.Column - 1
    rowCount = rng.Rows.Count
    colCount = rng.Columns.Count

    ReDim resultAry(1 To rowCount, 1 To colCount)

    For i = 1 To rowCount
        For j = 1 To colCount
            resultAry(i, j) = aryData(i + rowOffset, j + colOffset)
        Next j
    Next i
    
    sliceArray = resultAry

End Function