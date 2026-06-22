Function sheetToArray(Optional ws As Worksheet) As Variant
    Dim lastUsedRange As range
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    '最終セルを取得する
    Set lastUsedRange = getLastUsedRange(ws)


    'データがない(空シート)の時はemptyを返す
    If lastUsedRange Is Nothing Then
        sheetToArray = Empty
    'データがA1の時はスカラー値になるため2次元配列に変換
    ElseIf lastUsedRange.Row = 1 And lastUsedRange.Column = 1 Then
        Dim tmpAry As Variant
        ReDim tmpAry(1 To 1, 1 To 1)
        tmpAry(1, 1) = ws.Cells(1, 1).Value
        sheetToArray = tmpAry
    '上記以外はデータ範囲を配列として取得する
    Else
        sheetToArray = ws.range(ws.Cells(1, 1), lastUsedRange).Value
    End If

End Function

