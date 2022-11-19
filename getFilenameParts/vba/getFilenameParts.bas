'フルパス保存用プロパティ
Private STR_FullFilename As String
Private STR_Delimiter As String


Private Sub Class_Initialize()
    'デリミタの既定値は"\"(windowsパス)
    STR_Delimiter = "\"
End Sub

'デリミタを設定する
Public Property Let setDelimiter(argDelimiter As String)

    STR_Delimiter = argDelimiter

End Property


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

    getPath = Left(STR_FullFilename, InStrRev(STR_FullFilename, STR_Delimiter))
    
End Property

'ファイル名取得
Property Get getFilename()

    If STR_FullFilename = "" Then
        Err.Raise 1000, , "フルパスが設定されていません"
    End If

    getFilename = Mid(STR_FullFilename, InStrRev(STR_FullFilename, STR_Delimiter) + 1)
    
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

