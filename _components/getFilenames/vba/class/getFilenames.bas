Private FSO As Object

Private searchDirectory As String
Private searchPattern As String
Private redimNumber As Long

'検索ディレクトリを設定する
Public Property Let setDirectory(directoryString As String)
    '相対パスの場合 / if relative path
    If FSO.GetDriveName(directoryString) = "" Then
        'bookパス+相対パス+仮ファイル名を結合して親フォルダ名を取得する(末尾の\を付加しないに統一するため)
        'get pareant folder name from (book path + relative path + tempname) to remove "/" from the tail.
        searchDirectory = FSO.GetParentFolderName(FSO.BuildPath(FSO.BuildPath(ThisWorkbook.Path, directoryString), "tmp.txt"))
    '絶対パスの場合 / if absolute path
    Else
        '絶対パス+仮ファイル名を結合して親フォルダ名を取得する(末尾の \ を付加しないに統一するため)
        'get pareant folder name from (absolute path + tempname) to remove "/" from the tail.
        searchDirectory = FSO.GetParentFolderName(FSO.BuildPath(directoryString, "tmp.txt"))
    End If
End Property

'パターンを設定する
Public Property Let setPattern(patternString As String)
'    Dim tmp As String
'    tmp = Dir(FSO.BuildPath(ThisWorkbook.Path, patternString), vbDirectory)
    searchPattern = patternString
End Property

'設定されている検索ディレクトリを返す
Public Property Get getSearchDirectory()
    getSearchDirectory = searchDirectory
End Property

'設定されている検索ディレクトリの文字列長を返す
Public Property Get getSearchDirectoryLen()
    getSearchDirectory = Len(searchDirectory)
End Property

'設定されている検索ディレクトリを返す
Public Property Get getSearchPattern()
    getSearchPattern = searchPattern
End Property


Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    '仮の配列要素数を設定
    redimNumber = 1000
End Sub

Private Sub Class_Terminate()
    Set FSO = Nothing
End Sub


Public Function getFilenames() As String()
    Dim filenames() As String
    Dim filename As String
    Dim cnt As Long
    
    If searchDirectory = "" Or searchPattern = "" Then
        Err.Raise 1000
        Exit Function
    End If
    
    
    cnt = -1
    filename = Dir(FSO.BuildPath(searchDirectory, searchPattern), vbNormal)

    Do Until filename = ""
        cnt = cnt + 1
        
        '仮の配列要素数に達したら仮の要素数を追加する
        If cnt Mod redimNumber = 0 Then
            ReDim Preserve filenames(cnt + redimNumber)
        End If
        
        filenames(cnt) = filename
        
        filename = Dir()
       
    Loop
    
    
    'パターンマッチするファイルがない場合は未初期化配列を返す
    If cnt = -1 Then
        'getFilenames = ""
        Exit Function
    End If

    '実際のファイル数で配列要素数を再設定する
    'set index number to actual files number
    ReDim Preserve filenames(cnt)
    getFilenames = filenames

End Function


