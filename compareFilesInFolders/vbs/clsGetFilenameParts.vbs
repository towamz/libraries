Class clsGetFilenameParts
	'フルパス保存用プロパティ
	Private STR_FullFilename 
	Private STR_Delimiter

    Private Sub Class_Initialize()
	    'デリミタの既定値は"\"(windowsパス)
	    STR_Delimiter = "\"
    End Sub
    
    'デリミタを設定する
    Public Property Let setDelimiter(argDelimiter)

    	STR_Delimiter = argDelimiter

	End Property
    
    
	'フルパスを設定する
	Public Property Let setFullFilename(argFullFilename)

	    STR_FullFilename = argFullFilename

	End Property

	'フルパス取得
	Public Property Get getFullFilename()

	    If STR_FullFilename = "" Then
	        Err.Raise 1000, , "フルパスが設定されていません"
	    End If

	    getFullFilename = STR_FullFilename
	    
	End Property

	'パス取得
	Public Property Get getPath()

	    If STR_FullFilename = "" Then
	        Err.Raise 1000, , "フルパスが設定されていません"
	    End If

	    getPath = Left(STR_FullFilename, InStrRev(STR_FullFilename, STR_Delimiter))
	    
	End Property

	'ファイル名取得
	Public Property Get getFilename()

	    If STR_FullFilename = "" Then
	        Err.Raise 1000, , "フルパスが設定されていません"
	    End If

	    getFilename = Mid(STR_FullFilename, InStrRev(STR_FullFilename, STR_Delimiter) + 1)
	    
	End Property

	'ファイル名(拡張子なし)取得
	Public Property Get getFilenameNoExt()
	    Dim strFilename 

	    strFilename = getFilename
	    
	    getFilenameNoExt = Left(strFilename, InStrRev(strFilename, ".") - 1)
	    
	End Property

	'拡張子取得
	Public Property Get getExtension()
	    Dim strFilename 

	    strFilename = getFilename
	    
	    getExtension = Mid(strFilename, InStrRev(strFilename, ".") + 1)
	    
	End Property
End Class