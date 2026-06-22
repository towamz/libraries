Function get1DArray(aryData As Variant, Optional newColBase As Long = 1) As Variant
    Dim resultAry As Variant
    Dim oldRowBase As Long
    Dim oldColBase As Long
    Dim oldRowCount As Long
    Dim oldColCount As Long
    Dim colCount As Long
    Dim colIndex As Long
    Dim i As Long
    Dim j As Long

    Select Case getArrayRank(aryData)
        Case 0
            ReDim resultAry(newColBase To newColBase)
            resultAry(newColBase) = aryData
            get1DArray = resultAry
        
        Case 1
            '1次元配列のベース変換はnewRowBaseは利用されないが第2引数のため1を仮に渡す
            get1DArray = rebaseArray(aryData, 1, newColBase)
        
        Case 2
            oldRowBase = LBound(aryData, 1)
            oldColBase = LBound(aryData, 2)
            
            oldRowCount = UBound(aryData, 1) - oldRowBase
            oldColCount = UBound(aryData, 2) - oldColBase
            colCount = (oldRowCount + 1) * (oldColCount + 1) - 1
            
            ReDim resultAry(newColBase To newColBase + colCount)
        
            colIndex = newColBase
            For i = 0 To oldRowCount
                For j = 0 To oldColCount
                    resultAry(colIndex) = aryData(i + oldRowBase, j + oldColBase)
                    colIndex = colIndex + 1
                Next j
            Next i
            get1DArray = resultAry
        
        Case Else
            Err.Raise vbObjectError + 1, "get1DArray", "3次元以上の配列には対応していません"
    
    End Select

End Function