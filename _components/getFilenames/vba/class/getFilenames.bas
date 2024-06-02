Option Explicit

Private FSO As Object
Private REP As Object

Private STR_Directory As String
Private STR_RegExp As String
Private STR_Delimiter As String

Private OBJ_Files As Object
Private DIC_FileNames As Object
Private ARY_FileNames() As String
Private STR_Filenames As String
Private LNG_FilesCnt As Long
Private IS_FirstMatchFileOnly As Boolean
Private IS_Exec As Boolean

'検索ディレクトリを設定する
Public Property Let setDirectory(argDirectory As String)
    Dim tmpDirectory As String
    
    '相対パスの場合 / if relative path
    If FSO.GetDriveName(argDirectory) = "" Then
        'bookパス+相対パス+仮ファイル名を結合して親フォルダ名を取得する(末尾の\を付加しないに統一するため)
        'get pareant folder name from (book path + relative path + tempname) to remove "/" from the tail.
        tmpDirectory = FSO.GetParentFolderName(FSO.BuildPath(FSO.BuildPath(ThisWorkbook.Path, argDirectory), "tmp.txt"))
    '絶対パスの場合 / if absolute path
    Else
        '絶対パス+仮ファイル名を結合して親フォルダ名を取得する(末尾の \ を付加しないに統一するため)
        'get pareant folder name from (absolute path + tempname) to remove "/" from the tail.
        tmpDirectory = FSO.GetParentFolderName(FSO.BuildPath(argDirectory, "tmp.txt"))
    End If
    
    'ディレクトリが実在するか確認する
    If Not FSO.FolderExists(tmpDirectory) Then
        Err.Raise 1000
    End If

    STR_Directory = tmpDirectory
End Property

'パターンを設定する
Public Property Let setRegExp(argRegExp As String)
    '不正な正規表現であればエラー発生 / an error occure if invalid RegExp
    REP.Pattern = argRegExp
    REP.test ("testExec")

    STR_RegExp = argRegExp
End Property

'文字列用デミリタを設定する
Public Property Let setDelimiter(argDelimiter As String)
    STR_Delimiter = argDelimiter
End Property

'最初にマッチしたファイル1つだけを返すか設定する
Public Property Let setIS_FirstMatchFileOnly(argIs As Boolean)
    IS_FirstMatchFileOnly = argIs
End Property



'設定されている検索ディレクトリを返す
Public Property Get getDirectory()
    getDirectory = STR_Directory
End Property

'設定されている検索ディレクトリの文字列長を返す
Public Property Get getDirectoryLen()
    getDirectoryLen = Len(STR_Directory)
End Property

'設定されている検索ディレクトリを返す
Public Property Get getRegExp()
    getRegExp = STR_RegExp
End Property

'ファイルオブジェクトを返す
Public Property Get getFilesObj()
    If Not IS_Exec Then
        Set OBJ_Files = FSO.GetFolder(STR_Directory).Files
    End If
    
    Set getFilesObj = OBJ_Files
End Property

'配列を返す
Public Property Get getFilenamesArray()
    If Not IS_Exec Then
        Call getFilenamesMain
    End If
    
    getFilenamesArray = ARY_FileNames
End Property

'ディクショナリを返す
Public Property Get getFilenamesDictionary()
    If Not IS_Exec Then
        Call getFilenamesMain
    End If
    Set getFilenamesDictionary = DIC_FileNames
End Property

'文字列を返す
Public Property Get getFilenamesString()
    If Not IS_Exec Then
        Call getFilenamesMain
    End If
    getFilenamesString = STR_Filenames
End Property

'ファイル数を返す
Public Property Get getFilesCnt()
    If Not IS_Exec Then
        Call getFilenamesMain
    End If
    getFilesCnt = LNG_FilesCnt
End Property


Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set REP = CreateObject("VBScript.RegExp")
    
    '既定値はカレントディレクトリ / defalut is the current directory
    STR_Directory = ThisWorkbook.Path

    '既定値はすべてのファイル / defalut is all files
    STR_RegExp = ".*"
    
    '既定値は"," / defalut is ","
    STR_Delimiter = ","
    
    '既定値はFalse / default is false
    IS_FirstMatchFileOnly = False
    
    IS_Exec = False
End Sub

'各変数を初期化
Public Sub initVariables()
    Set OBJ_Files = Nothing
    Set DIC_FileNames = Nothing
    Set DIC_FileNames = CreateObject("Scripting.Dictionary")
    Erase ARY_FileNames
    STR_Filenames = ""
    LNG_FilesCnt = 0
    IS_Exec = False
End Sub


Private Sub getFilenamesMain()
    Dim objFile
    
    '各変数を初期化
    Call initVariables
    
    Set OBJ_Files = getFilesObj()

    LNG_FilesCnt = 0
    '配列要素数をファイル数に設定する / set array index to fils count (possible max number)
    ReDim Preserve ARY_FileNames(OBJ_Files.Count)

    For Each objFile In OBJ_Files
        If REP.test(objFile.Name) Then
            DIC_FileNames.Add objFile.Name, objFile.Path
            ARY_FileNames(LNG_FilesCnt) = objFile.Name
            STR_Filenames = STR_Filenames & objFile.Name & STR_Delimiter
            
            LNG_FilesCnt = LNG_FilesCnt + 1

            If IS_FirstMatchFileOnly Then
                Exit For
            End If
        End If
    Next
    
    '後処理
    If LNG_FilesCnt = 0 Then
        ARY_FileNames = Array()
        STR_Filenames = ""
    Else
        '配列要素数を再設定する(正規表現に一致しないファイルがある可能性がある) /
        'reset array index (there may be files unmatch the regular expression)
        ReDim Preserve ARY_FileNames(LNG_FilesCnt - 1)
        '文字列から最後のデミリタを削除する
        STR_Filenames = Left(STR_Filenames, Len(STR_Filenames) - Len(STR_Delimiter))
    End If

    IS_Exec = True
End Sub

