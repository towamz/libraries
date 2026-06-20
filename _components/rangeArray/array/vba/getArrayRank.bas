Function getArrayRank(aryData As Variant) As Long
    Dim n As Long
    Dim tmp As Variant
    
    On Error Resume Next
    For n = 1 To 60
        Err.Clear
        tmp = LBound(aryData, n)
        If Err.Number <> 0 Then
            getArrayRank = n - 1
            On Error GoTo 0
            Exit Function
        End If
    Next
    On Error GoTo 0
    'VBAの配列は最大60次元
    getArrayRank = 60

End Function
