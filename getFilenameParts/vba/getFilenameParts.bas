'フルパス保存用プロパティ
Private STR_FullFilename As String


'フルパスを設定する
Public Property Let setFullFilename(argFullFilename As String)

    STR_FullFilename = argFullFilename

End Property

'フルパス取得
Property Get getFullFilename()

    If STR_FullFilename = "" Then
        Err.Raise 1000, , "フルパスが設定されていません"
    End If

    getFullFilename = STR_FullFilename
    
End Property

'パス取得
Property Get getPath()

    If STR_FullFilename = "" Then
        Err.Raise 1000, , "フルパスが設定されていません"
    End If

    getPath = Left(STR_FullFilename, InStrRev(STR_FullFilename, "\"))
    
End Property

'ファイル名取得
Property Get getFilename()

    If STR_FullFilename = "" Then
        Err.Raise 1000, , "フルパスが設定されていません"
    End If

    getFilename = Mid(STR_FullFilename, InStrRev(STR_FullFilename, "\") + 1)
    
End Property

'ファイル名(拡張子なし)取得
Property Get getFilenameNoExt()
    Dim strFilename As String

    strFilename = getFilename
    
    getFilenameNoExt = Left(strFilename, InStrRev(strFilename, ".") - 1)
    
End Property

'拡張子取得
Property Get getExtension()
    Dim strFilename As String

    strFilename = getFilename
    
    getExtension = Mid(strFilename, InStrRev(strFilename, ".") + 1)
    
End Property

