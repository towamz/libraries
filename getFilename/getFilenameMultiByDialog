Public Function getFilenameMultiByDialog(ByVal argPromptString As String, ByRef argFilenames As Variant) As Boolean
    '----------ファイルを開く(複数ファイル選択)----------
    argFilenames = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xlsx", _
                                        Title:=argPromptString, MultiSelect:=True)
    
    'ファイル名が取得できない時はFalse(配列でない)が帰ってくる
    If IsArray(argFilenames) Then
        getFilenameMultiByDialog = True
    Else
        getFilenameMultiByDialog = False
    End If

End Function
