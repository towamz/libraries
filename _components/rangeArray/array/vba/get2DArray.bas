Function get2DArray(aryData As Variant, Optional newRowBase As Long = 1, Optional newColBase As Long = 1) As Variant
    Dim resultAry As Variant
    Dim oldColBase As Long
    Dim colCount As Long
    Dim j As Long
    
    
    Select Case getArrayRank(aryData)
        Case 0
            ReDim resultAry(newRowBase To newRowBase, newColBase To newColBase)
            resultAry(newRowBase, newColBase) = aryData
            get2DArray = resultAry
        
        Case 1
            oldColBase = LBound(aryData, 1)
            colCount = UBound(aryData, 1) - oldColBase
            
            ReDim resultAry(newRowBase To newRowBase, newColBase To newColBase + colCount)
       
            For j = 0 To colCount
                resultAry(newRowBase, j + newColBase) = aryData(j + oldColBase)
            Next j
            get2DArray = resultAry
        Case 2
            get2DArray = rebaseArray(aryData, newRowBase, newColBase)
        
        Case Else
            Err.Raise vbObjectError + 1, "get2DArray", "3次元以上の配列には対応していません"
    
    End Select
End Function