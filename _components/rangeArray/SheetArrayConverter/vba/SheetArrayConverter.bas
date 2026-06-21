Option Explicit

Private WsOrig_ As Worksheet
Private WsDest_ As Worksheet
Private ArySheet_ As Variant
Private AryDest_ As Variant

Public Property Set WorksheetOrig(ws As Worksheet)
    Set WsOrig_ = ws
End Property

Public Property Set WorksheetDest(ws As Worksheet)
    Set WsDest_ = ws
End Property



Private Sub checkBeforeExec()
    If WsOrig_ Is Nothing Then
        Set WsOrig_ = ActiveSheet
    End If
    
    If WsDest_ Is Nothing Then
        Set WsDest_ = WsOrig_
    End If

End Sub


Public Function loadArray() As Variant
    Dim rngLastRow As range
    Dim rngLastCol As range
    Dim rngLast As range
    
    Call checkBeforeExec
    
    Set rngLast = getLastRange

    'データがない(空シート)の時はemptyを返す
    If rngLast Is Nothing Then
        ArySheet_ = Empty
    'データがA1の時はスカラー値になるため2次元配列に変換
    ElseIf rngLast.Row = 1 And rngLast.Column = 1 Then
        Dim tmpAry As Variant
        ReDim tmpAry(1 To 1, 1 To 1)
        tmpAry(1, 1) = WsOrig_.Cells(1, 1).Value
        ArySheet_ = tmpAry
    '上記以外はデータ範囲を配列として取得する
    Else
        ArySheet_ = WsOrig_.range(WsOrig_.Cells(1, 1), rngLast).Value
    End If

End Function


Public Sub setArray(rng As range, aryData As Variant)
    Dim tmpData As Variant
    
    '書き込み先が不明な場合はそのままexit
    If rng Is Nothing Then
        Exit Sub
    End If
    
    Set rng = WsDest_.range(rng.Address)

    'スカラー値・1次元配列対応のためrebaseArrayを実行する
    tmpData = rebaseArray(aryData)

    rng.Resize(UBound(tmpData, 1) - LBound(tmpData, 1) + 1, UBound(tmpData, 2) - LBound(tmpData, 2) + 1).Value = tmpData

    '読込と書込シート同じときは再読込
    If WsOrig_ Is WsDest_ Then
        Call loadArray
    End If

End Sub


Public Function getArray(Optional rng As range) As Variant
    Call checkBeforeExec
    
    If rng Is Nothing Then
        getArray = ArySheet_
    Else
        getArray = sliceArray(rng)
    End If
End Function



Private Function sliceArray(rng As range) As Variant
    Dim resultAry As Variant
    Dim rowOffset As Long
    Dim colOffset As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim i As Long
    Dim j As Long
    
    rowOffset = rng.Row - 1
    colOffset = rng.Column - 1
    rowCount = rng.Rows.Count
    colCount = rng.Columns.Count

    ReDim resultAry(1 To rowCount, 1 To colCount)

    For i = 1 To rowCount
        For j = 1 To colCount
            rowIndex = i + rowOffset
            colIndex = j + colOffset
            '配列のインデックス外の場合はEmptyを代入する
            If rowIndex >= LBound(ArySheet_, 1) And rowIndex <= UBound(ArySheet_, 1) And _
               colIndex >= LBound(ArySheet_, 2) And colIndex <= UBound(ArySheet_, 2) Then
                resultAry(i, j) = ArySheet_(rowIndex, colIndex)
            Else
                resultAry(i, j) = Empty
            End If
        Next j
    Next i
    
    sliceArray = resultAry

End Function



Function getLastRange() As range
    Dim rngLastRow As range
    Dim rngLastCol As range

    Set rngLastRow = WsOrig_.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If rngLastRow Is Nothing Then
         Set getLastRange = Nothing
         Exit Function
    End If
    
    Set rngLastCol = WsOrig_.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    Set getLastRange = WsOrig_.Cells(rngLastRow.Row, rngLastCol.Column)

End Function

