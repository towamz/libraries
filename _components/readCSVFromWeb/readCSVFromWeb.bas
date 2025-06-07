Function readCSVFromWeb(url As String) As String
    Dim httpRequest As Object: Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    
    ' インターネットからCSVをダウンロード
    httpRequest.Open "GET", url, False
    httpRequest.Send
    
    ' ダウンロードした内容を取得
    With stream
       .Type = 1 ' バイナリを読み込む
       .Open
       .Write httpRequest.responseBody
       .Position = 0
       .Type = 2 ' テキストに変更
       .Charset = "shift_jis" '"utf-8"
       readCSVFromWeb = .ReadText
       .Close
    End With
    
    Set stream = Nothing
    Set httpRequest = Nothing
End Function