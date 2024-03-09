Private Function getFilenameMultiByDir(argPattern As String) As Variant
    Dim aryFileName() As String
    Dim FileName As String
    Dim cnt As Integer
    
    cnt = 0

    'パターンに一致したファイルを処理対象にする
    FileName = Dir(argPattern, vbNormal)
    
    Do While FileName <> ""
        ReDim Preserve aryFileName(cnt)
        
        aryFileName(cnt) = FileName
        
        FileName = Dir()
        cnt = cnt + 1
    Loop

    getFilenameByDir = aryFileName

End Function
