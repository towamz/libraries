Function normalizeWhitespace(targetString As String) As String
    Dim REX As Object
    Set REX = CreateObject("VBScript.RegExp")
    
    REX.IgnoreCase = True
    REX.Global = True
    REX.Pattern = "(\s|　)+"

    '連続する全角半角スペースを半角スペース1つに置換
    '先頭・末尾空白はtrimで削除
    normalizeWhitespace = REX.Replace(Trim(targetString), Chr(32))

    Set REX = Nothing
End Function