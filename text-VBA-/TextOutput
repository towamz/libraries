'テキストファイル出力
Public Sub TextOutput()
    Dim stdRange As Range
    Dim userName As String
    Dim filePath As String
    Dim fileName As String
    
    Set stdRange = Range("A1")
    userName = Range("A1").Value
    filePath = ThisWorkbook.Path & "\data\"
    fileName = Format(Now, "mmdd") & userName & ".txt"
    
    
    'ファイルを書き込みで開く(無ければ新規作成される、あれば上書き)
    Open filePath & fileName For Output As #1
    
    For j = 0 To 19 Step 1
        For i = 0 To 19 Step 1
        
            Print #1, Trim(stdRange.Offset(i, j).Value)
    
        Next
    Next

    '開いたファイルを閉じる
    Close #1
     
    '終わったのが分かるようにメッセージを出す
    MsgBox "完了！"
 
End Sub
