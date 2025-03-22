Option Explicit

Private FSO As Object
Private REX As Object

Private TargetDirectory_ As String
Private Pattern_ As String
Private Delimiter_ As String

Private IsFirstMatchFileOnly_ As Boolean
Private IsExec_ As Boolean

Private Files_ As Object
Private DicFileNames_ As Object
Private AryFileNames_() As String
Private Ary2dFileNames_() As String
Private StrFilenames_ As String
Private LngFilesCnt_ As Long


'検索ディレクトリを設定する
Public Property Let TargetDirectory(arg1 As String)
    Dim tmpDirectory As String
    
    If FSO.GetDriveName(arg1) = "" Then
        tmpDirectory = FSO.GetParentFolderName(FSO.BuildPath(FSO.BuildPath(ThisWorkbook.Path, arg1), "tmp.txt"))
    Else
        tmpDirectory = FSO.GetParentFolderName(FSO.BuildPath(arg1, "tmp.txt"))
    End If
    
    If Not FSO.FolderExists(tmpDirectory) Then
        Err.Raise 1000
    End If

    TargetDirectory_ = tmpDirectory
End Property

Public Property Get TargetDirectory() As String
    getDirectory = TargetDirectory_
End Property

Public Property Get TargetDirectoryLen() As Long
    TargetDirectoryLen = Len(TargetDirectory_)
End Property


Public Property Let Pattern(arg1 As String)
    REX.Pattern = arg1
    REX.test ("testExec")
    Pattern_ = arg1
End Property

Public Property Get Pattern() As String
    Pattern = Pattern_
End Property


Public Property Let Delimiter(arg1 As String)
    Delimiter_ = arg1
End Property

Public Property Get Delimiter() As String
    Delimiter = Delimiter_
End Property


Public Property Let IsFirstMatchFileOnly(arg1 As Boolean)
    IsFirstMatchFileOnly_ = arg1
End Property

Public Property Get IsFirstMatchFileOnly() As Boolean
    IsFirstMatchFileOnly = IsFirstMatchFileOnly_
End Property


'実際のファイル名取得用getter
Public Property Get FilesObj() As Object
    If Not IsExec_ Then
        Set Files_ = FSO.GetFolder(TargetDirectory_).Files
    End If
    
    Set FilesObj = Files_
End Property

Public Property Get FilenamesArray() As String()
    If Not IsExec_ Then
        Call getFilenamesMain
    End If
    
    FilenamesArray = AryFileNames_
End Property

Public Property Get FilenamesArray2D() As String()
    If Not IsExec_ Then
        Call getFilenamesMain
    End If
    
    FilenamesArray2D = Ary2dFileNames_
End Property

Public Property Get FilenamesDictionary() As Object
    If Not IsExec_ Then
        Call getFilenamesMain
    End If
    Set FilenamesDictionary = DicFileNames_
End Property

Public Property Get FilenamesString() As String
    If Not IsExec_ Then
        Call getFilenamesMain
    End If
    FilenamesString = StrFilenames_
End Property

Public Property Get FilesCnt() As Long
    If Not IsExec_ Then
        Call getFilenamesMain
    End If
    FilesCnt = LngFilesCnt_
End Property


Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set REX = CreateObject("VBScript.RegExp")
    
    TargetDirectory_ = ThisWorkbook.Path
    Pattern_ = ".*"
    Delimiter_ = ","
    IsFirstMatchFileOnly_ = False
    IsExec_ = False
End Sub

Public Sub initVariables()
    Set Files_ = Nothing
    Set DicFileNames_ = Nothing
    Set DicFileNames_ = CreateObject("Scripting.Dictionary")
    Erase AryFileNames_
    Erase Ary2dFileNames_
    StrFilenames_ = ""
    LngFilesCnt_ = 0
    IsExec_ = False
End Sub

Private Sub getFilenamesMain()
    Dim File
    
    Call initVariables
    
    Set Files_ = FilesObj()
    LngFilesCnt_ = 0
    ReDim Preserve AryFileNames_(Files_.Count)
    ReDim Preserve Ary2dFileNames_(1, Files_.Count)
    
    For Each File In Files_
        If REX.test(File.Name) Then
            DicFileNames_.Add File.Name, File.Path
            AryFileNames_(LngFilesCnt_) = File.Name
            Ary2dFileNames_(0, LngFilesCnt_) = File.Name
            Ary2dFileNames_(1, LngFilesCnt_) = File.Path
            StrFilenames_ = StrFilenames_ & File.Name & Delimiter_
            
            LngFilesCnt_ = LngFilesCnt_ + 1
            
            If IsFirstMatchFileOnly_ Then
                Exit For
            End If
        End If
    Next
    
    If LngFilesCnt_ = 0 Then
        AryFileNames_ = Array()
        Ary2dFileNames_ = Array()
        StrFilenames_ = ""
    Else
        'パターンマッチしているのでフォルダにあるファイル数より少なくなる場合があるためインデックス番号を修正する
        ReDim Preserve AryFileNames_(LngFilesCnt_ - 1)
        ReDim Preserve Ary2dFileNames_(1, LngFilesCnt_ - 1)
        StrFilenames_ = Left(StrFilenames_, Len(StrFilenames_) - Len(Delimiter_))
    End If

    '実行済みフラグを立てる
    IsExec_ = True
End Sub

