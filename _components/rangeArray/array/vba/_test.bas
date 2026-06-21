Sub rebaseArraytest()
    Dim aryData1 As Variant
    Dim aryData2 As Variant

    'インデックス(0,0)開始
    ReDim aryData1(0 To 1, 0 To 2)
    
    aryData1(0, 0) = 1
    aryData1(0, 1) = 3
    aryData1(0, 2) = 5
    aryData1(1, 0) = 2
    aryData1(1, 1) = 4
    aryData1(1, 2) = 6

    aryData2 = rebaseArray(aryData1, 1, 2)
End Sub

Sub rebaseArraytest2()
    Dim aryData1 As Variant
    Dim aryData2 As Variant

    'インデックス(0,0)開始
    ReDim aryData1(0 To 2)
    
    aryData1(0) = 1
    aryData1(1) = 3
    aryData1(2) = 5

    aryData2 = rebaseArray(aryData1, 1, 2)
End Sub

Sub rebaseArraytest3()
    Dim aryData1 As Variant
    Dim aryData2 As Variant

    'インデックス(0,0)開始
    aryData1 = 5
    aryData2 = rebaseArray(aryData1, 1, 2)
End Sub

Sub sliceArrayTest()
    Dim ary As Variant

'    ary = sliceArray(Range("D5:E7"), Worksheets("Sheet2"))
    ary = sliceArray(Range("A2:C3"), Worksheets("Sheet2"))

    Stop
End Sub


Sub getLastUsedRangeTest()
    Dim rng As range
    
    Set rng = getLastUsedRange(Worksheets("Sheet2"))
    
    Stop
End Sub
