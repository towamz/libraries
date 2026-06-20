Function getArrayRank(aryData As Variant, Optional maxBaseNumber As Long = 60) As Long
    Dim i As Long
    Dim tmp As Long
    
    If Not IsArray(aryData) Then
         getArrayRank = 0
         Exit Function
    End If
    
    For i = maxBaseNumber To 1 Step -1
        On Error Resume Next
        tmp = LBound(aryData, i)
        If Err.Number = 0 Then
            getArrayRank = i
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next
    
End Function
