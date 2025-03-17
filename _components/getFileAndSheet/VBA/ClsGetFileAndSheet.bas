Option Explicit
Private DefaultDirectory_ As String

Private WbOrig_ As Workbook
Private WbOrigAbsoluteFileName_ As String
Private WbOrigFileName_ As String

Private WsOrig_ As Worksheet
Private WsOrigSheetName_ As String
Private WsOrigSheetNames_() As String

Private GetOpenFilenameFileFilter_ As String
Private GetOpenFilenameTitle_  As String
    
Public Property Let DefaultDirectory(arg1 As String)
    DefaultDirectory_ = arg1
End Property
    
Public Property Get DefaultDirectory() As String
    DefaultDirectory = DefaultDirectory_
End Property

Public Property Let WbOrigAbsoluteFileName(arg1 As String)

'    With CreateObject("Scripting.FileSystemObject")
'        If Not .FileExists(arg1) Then
'            Err.Raise 1003, , "ファイルが存在しません。:" & arg1
'        End If
'    End With
    
    WbOrigAbsoluteFileName_ = arg1
End Property
    
Public Property Get WbOrigAbsoluteFileName() As String
    DefaultDirectory = WbOrigAbsoluteFileName_
End Property

Public Property Let WbOrigFileName(arg1 As String)
    WbOrigFileName_ = arg1
End Property
    
Public Property Get WbOrigFileName() As String
    DefaultDirectory = WbOrigFileName_
End Property

Public Property Let GetOpenFilenameFileFilter(arg1 As String)
'    With CreateObject("VBScript.RegExp")
'        .IgnoreCase = True
'        .Global = True
'        '.Pattern = "^([a-zA-Z0-9\s]+(?:\s?[a-zA-Z0-9]*\s?)*\s?\([a-zA-Z0-9\*\.;,\s]*\))(\|[a-zA-Z0-9\s]+(?:\s?[a-zA-Z0-9]*\s?)*\s?\([a-zA-Z0-9\*\.;,\s]*\))*$"
'        .Pattern = "^(.+)(\([a-zA-Z\*\.;,\s\-]*\))(\|(.+)(\([a-zA-Z\*\.;,\s\-]*\)))*$"
'        If Not .test(arg1) Then
'            Err.Raise 1002, , "フィルターが間違っています"
'        End If
'    End With
    
    GetOpenFilenameFileFilter_ = arg1
End Property
    
Public Property Get GetOpenFilenameFileFilter() As String
    DefaultDirectory = GetOpenFilenameFileFilter_
End Property


Private Sub Class_Initialize()
    GetOpenFilenameFileFilter_ = "ファイル,*.*"
    GetOpenFilenameTitle_ = "ファイルを選んでください"

End Sub

Private Sub Class_Terminate()
    If Not WbOrig_ Is Nothing Then
        WbOrig_.Close SaveChanges:=False
    End If
End Sub

Public Function getBook() As Workbook
    If WbOrig_ Is Nothing Then
        With CreateObject("Scripting.FileSystemObject")
            If Not .FileExists(WbOrigAbsoluteFileName_) Then
                GetOpenFilenameTitle_ = "指定されたファイルが見つかりませんでした:" & GetOpenFilenameTitle_
                Call getAbsoluteFileName
            End If
        End With
        
        Set WbOrig_ = Workbooks.Open(Filename:=WbOrigAbsoluteFileName_, ReadOnly:=True)
    End If

    Set getBook = WbOrig_

End Function
    
Public Function getAbsoluteFileName() As String
    'カレントディレクトリ変更  / change the current directory
    If DefaultDirectory <> "" Then
        With CreateObject("WScript.Shell")
            .CurrentDirectory = DefaultDirectory
        End With
    End If
    
    'ファイル名(フルパス)取得  / get filename(full path)
    WbOrigAbsoluteFileName_ = Application.GetOpenFilename(fileFilter:=GetOpenFilenameFileFilter_, Title:=GetOpenFilenameTitle_)
    
    'キャンセルしたときは中止 / abort when cancel was pushed
    If WbOrigAbsoluteFileName_ = "False" Then
        WbOrigAbsoluteFileName_ = ""
        Err.Raise 1001, , "ファイルが選ばれませんでした。"
    End If
    
'    ファイル名指定がある場合
    If WbOrigFileName_ <> "" Then
        With CreateObject("Scripting.FileSystemObject")
            If .GetFileName(WbOrigAbsoluteFileName_) <> WbOrigFileName_ Then
              Err.Raise 1004, , "選択したファイルが指定されたファイル名と一致しません"
            End If
        End With
    End If
    
    getAbsoluteFileName = WbOrigAbsoluteFileName_

End Function

Public Function getSheet() As Worksheet

    If WsOrig_ Is Nothing Then
        If WsOrigSheetName_ = "" Then
            Call getSheetName
        End If
    
        Set WsOrig_ = WbOrig_.Worksheets(WsOrigSheetName_)
    End If

    Set getSheet = WsOrig_

End Function

Public Function getSheetName() As String
    Dim wsString As String
    Dim i, res, errNum, loopExitNum As Long
    
    'シート名配列未取得の場合は取得する
    If (Not WsOrigSheetNames_) = -1 Then
        Call getSheetNames
    End If
    
    'シートが１枚のみの時はプロンプトを表示しないでそのまま返す
    If UBound(WsOrigSheetNames_) = 0 Then
        res = 0
    Else
        wsString = ""
        
        For i = 0 To UBound(WsOrigSheetNames_)
            wsString = wsString & i & ":" & WsOrigSheetNames_(i) & vbCrLf
        Next i
        
        '終了番号を取得(シート数が1~9の時:99, 10~99の時:999)
        loopExitNum = (10 ^ (Int(Log(UBound(WsOrigSheetNames_)) / Log(10)) + 2)) - 1
        
        Do
            On Error Resume Next
            res = CLng(InputBox(wsString, "選択するシート名の番号を入力してください。" & loopExitNum & "で終了します。"))
            'エラー処理を終わらせないとErr.Raiseできないのでエラー番号を取得
            errNum = Err.Number
            On Error GoTo 0
            
            Select Case errNum
                '整数が入力されたとき(正常)
                Case 0
                    '99が入力されたときは処理を中断
                    If res = loopExitNum Then
                        Err.Raise 1099, , "ユーザーによる中断"
                    End If
                '整数以外が入力されたとき(異常)
                Case 13
                    'resに0を代入してループさせる
                    res = 0
                Case Else
                    Err.Raise 9999, , "不明なエラー"
            End Select
        Loop While res < 0 Or UBound(WsOrigSheetNames_) < res
    End If
    
    WsOrigSheetName_ = WsOrigSheetNames_(res)
    
    getSheetName = WsOrigSheetName_

End Function

Public Function getSheetNames() As String()
    Dim i As Long
    
    If WbOrig_ Is Nothing Then
        Call getBook
    End If
    
    ReDim WsOrigSheetNames_(WbOrig_.Worksheets.Count - 1)
    
    For i = 1 To WbOrig_.Worksheets.Count
        WsOrigSheetNames_(i - 1) = WbOrig_.Worksheets(i).Name
    Next i

    getSheetNames = WsOrigSheetNames_

End Function


'Err.Raise 1001, , "ファイルが選ばれませんでした。"
'Err.Raise 1002, , "フィルターが間違っています"
'Err.Raise 1003, , "ファイルが存在しません。"
'Err.Raise 1004, , "選択したファイルが指定されたファイル名と一致しません"
'Err.Raise 1099, , "ユーザーによる中断"
'Err.Raise 9999, , "不明なエラー"
