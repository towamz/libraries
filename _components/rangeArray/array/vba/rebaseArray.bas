Function rebaseArray(aryData As Variant, Optional newRowBase As Long = 1, Optional newColBase As Long = 1) As Variant
    Dim resultAry As Variant
    Dim arrayRank As Long
    Dim oldRowBase As Long
    Dim oldColBase As Long
    Dim rowOffset As Long
    Dim colOffset As Long
    Dim i As Long
    Dim j As Long

    arrayRank = getArrayRank(aryData, 3)

    Select Case arrayRank
        Case 0
            ReDim resultAry(newRowBase To newRowBase, newColBase To newColBase)
            resultAry(newRowBase, newColBase) = aryData
            rebaseArray = resultAry
        Case 1
            oldColBase = LBound(aryData, 1)
            colOffset = UBound(aryData, 1) - oldColBase
            
            ReDim resultAry(newRowBase To newRowBase, newColBase To newColBase + colOffset)
       
            For j = 0 To colOffset
                resultAry(newRowBase, j + newColBase) = aryData(j + oldColBase)
            Next j
            rebaseArray = resultAry
        Case 2
            oldRowBase = LBound(aryData, 1)
            oldColBase = LBound(aryData, 2)
            rowOffset = UBound(aryData, 1) - oldRowBase
            colOffset = UBound(aryData, 2) - oldColBase
            
            ReDim resultAry(newRowBase To newRowBase + rowOffset, newColBase To newColBase + colOffset)
            
            For i = 0 To rowOffset
                For j = 0 To colOffset
                    resultAry(i + newRowBase, j + newColBase) = aryData(i + oldRowBase, j + oldColBase)
                Next j
            Next i
            rebaseArray = resultAry
        Case 3
            Err.Raise vbObjectError + 1, "rebaseArray", "3次元以上の配列には対応していません"
    End Select

End Function