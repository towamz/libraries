Public Function getFilenameByDialog(ByRef argFilename As String) As Boolean
    '----------ファイルを開く----------
    argFilename = Application.GetOpenFilename(FileFilter:="テキストファイル,*.txt", _
                                        Title:="テキストファイルを選択", MultiSelect:=False)
    
    'ファイル名が取得できない時はFalse(配列でない)が帰ってくる
    If argFilename = "False" Then
        getFilenameByDialog = False
    Else
        getFilenameByDialog = True
    End If

End Function



