Option Explicit

Sub arrayToRangeSample()
    Dim rng11 As Range: Set rng11 = Range("B5")
    Dim rng12 As Range: Set rng12 = Range("B6")
    Dim rng21 As Range: Set rng21 = Range("B7")
    Dim rng22 As Range: Set rng22 = Range("B8")
    Dim aryData1 As Variant
    Dim aryData2 As Variant
    
    'インデックス(0,0)開始
    ReDim aryData1(0 To 0, 0 To 2)
    
    aryData1(0, 0) = 1
    aryData1(0, 1) = 3
    aryData1(0, 2) = 5
    
    'インデックス(1,1)開始
    ReDim aryData2(1 To 1, 1 To 3)
    
    aryData2(1, 1) = 2
    aryData2(1, 2) = 4
    aryData2(1, 3) = 6
    
    'インデックス開始に依存しない汎用的な書き方
    rng11.Resize(UBound(aryData1, 1) - LBound(aryData1, 1) + 1, UBound(aryData1, 2) - LBound(aryData1, 2) + 1).Value = aryData1
    rng21.Resize(UBound(aryData2, 1) - LBound(aryData2, 1) + 1, UBound(aryData2, 2) - LBound(aryData2, 2) + 1).Value = aryData2
    
    'インデックス開始が(0,0)の場合は下記でもOK
    rng12.Resize(UBound(aryData1, 1) + 1, UBound(aryData1, 2) + 1).Value = aryData1
    'インデックス開始が(1,1)の場合は下記でもOK
    rng22.Resize(UBound(aryData2, 1), UBound(aryData2, 2)).Value = aryData2

End Sub

