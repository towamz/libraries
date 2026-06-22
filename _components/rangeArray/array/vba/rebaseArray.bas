Function rebaseArray(aryData As Variant, Optional newRowBase As Long = 1, Optional newColBase As Long = 1) As Variant
    Dim resultAry As Variant
    Dim oldRowBase As Long
    Dim oldColBase As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim j As Long

    Select Case getArrayRank(aryData)
        Case 0
            Err.Raise vbObjectError + 1, "rebaseArray", "配列を指定してください"
        Case 1
            oldColBase = LBound(aryData, 1)
            
            'インデックス開始が同じであれば引数をそのまま返す
            If oldColBase = newColBase Then
                rebaseArray = aryData
                Exit Function
            End If
            
            colCount = UBound(aryData, 1) - oldColBase
            
            ReDim resultAry(newColBase To newColBase + colCount)
       
            For j = 0 To colCount
                resultAry(j + newColBase) = aryData(j + oldColBase)
            Next j
            rebaseArray = resultAry
        Case 2
            oldRowBase = LBound(aryData, 1)
            oldColBase = LBound(aryData, 2)
            'インデックス開始が同じであれば引数をそのまま返す
            If oldRowBase = newRowBase And oldColBase = newColBase Then
                rebaseArray = aryData
                Exit Function
            End If
            
            rowCount = UBound(aryData, 1) - oldRowBase
            colCount = UBound(aryData, 2) - oldColBase
            
            ReDim resultAry(newRowBase To newRowBase + rowCount, newColBase To newColBase + colCount)
            
            For i = 0 To rowCount
                For j = 0 To colCount
                    resultAry(i + newRowBase, j + newColBase) = aryData(i + oldRowBase, j + oldColBase)
                Next j
            Next i
            rebaseArray = resultAry
        Case Else
            Err.Raise vbObjectError + 1, "rebaseArray", "3次元以上の配列には対応していません"
    End Select

End Function