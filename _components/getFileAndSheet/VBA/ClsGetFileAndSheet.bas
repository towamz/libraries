Option Explicit

Public DefaultDirectory As String

Private Wb_ As Workbook
Public AbsoluteFileName As String   'ブック名(フルパス)指定用変数 / 指定なしかファイルがない場合はダイアログ表示
Public FileName As String '選択されたファイルが指定のファイル名であるか確認するための変数(チェック不要であれば空白)

Private Ws_ As Worksheet
Public SheetName As String  'シート名指定用変数 / 指定なしかシートがない場合はダイアログ表示
Private SheetNames_() As String
Private SheetNamesString_ As String

Public DialogFileFilter As String
Public DialogTitle As String

Private Sub Class_Initialize()
    DialogFileFilter = "Excel,*.xls*"
    DialogTitle = "ファイルを選んでください。"
    SheetNamesString_ = ""
End Sub

Private Sub Class_Terminate()
    If Not Wb_ Is Nothing Then
        Wb_.Close SaveChanges:=False
    End If
End Sub

Public Function getBook() As Workbook
    If Wb_ Is Nothing Then
        With CreateObject("Scripting.FileSystemObject")
            If AbsoluteFileName = "" Then
                Call getAbsoluteFileName
            ElseIf Not .FileExists(AbsoluteFileName) Then
                DialogTitle = DialogTitle & " 指定されたファイルが見つかりませんでした:" & AbsoluteFileName
                Call getAbsoluteFileName
            End If
        End With
        
        Set Wb_ = Workbooks.Open(FileName:=AbsoluteFileName, ReadOnly:=True)
    End If

    Set getBook = Wb_

End Function

Public Function getAbsoluteFileName() As String
    'カレントディレクトリ変更  / change the current directory
    If DefaultDirectory <> "" Then
        With CreateObject("WScript.Shell")
            .CurrentDirectory = DefaultDirectory
        End With
    End If
    
    'ファイル名(フルパス)取得  / get filename(full path)
    AbsoluteFileName = Application.GetOpenFilename(fileFilter:=DialogFileFilter, Title:=DialogTitle)
    
    'キャンセルしたときは中止 / abort when cancel was pushed
    If AbsoluteFileName = "False" Then
        AbsoluteFileName = ""
        Err.Raise 1001, , "ファイルが選ばれませんでした。"
    End If
    
    'ファイル名指定がある場合
    If FileName <> "" Then
        With CreateObject("Scripting.FileSystemObject")
            If .GetFileName(AbsoluteFileName) <> FileName Then
                Err.Raise 1004, , "選択したファイルが指定されたファイル名と一致しません"
            End If
        End With
    End If
    
    getAbsoluteFileName = AbsoluteFileName

End Function

Public Function getSheet() As Worksheet
    'シート名直指定あり
    If SheetName <> "" Then
        'シート名指定があるときはブックを取得する
        Call getBook

        On Error Resume Next
        'シートを取得してみる
        Set Ws_ = Wb_.Worksheets(SheetName)
        On Error GoTo 0
    End If

    'シート名指定がない/シート名直指定でシートが存在しない場合
    If Ws_ Is Nothing Then
        If SheetName <> "" Then
            SheetNamesString_ = """" & SheetName & """" & "シートは見つかりませんでした。" & vbCrLf & SheetNamesString_
        End If
    
        'シート名選択プロンプト表示
        Call getSheetName
        Set Ws_ = Wb_.Worksheets(SheetName)
    End If

    Set getSheet = Ws_

End Function

Public Function getSheetName() As String
'    Dim wsString As String
    Dim i, res, errNum, loopExitNum As Long
    
    'シート名配列未取得の場合は取得する
    If (Not SheetNames_) = -1 Then
        Call getSheetNames
    End If
    
    'シートが１枚のみの時はそのまま返す
    If UBound(SheetNames_) = 0 Then
        'シート名指定があり一致していないときは警告を表示する
        If SheetName <> "" Then
            If SheetNames_(0) <> SheetName Then
                MsgBox """" & SheetName & """" & "シートは見つかりませんでした。シートが１枚のため" & """" & SheetNames_(0) & """" & "を選択します。"
            End If
        End If
        res = 0
    Else
        '終了番号を取得(シート数が1~9の時:99, 10~99の時:999)
        loopExitNum = (10 ^ (Int(Log(UBound(SheetNames_)) / Log(10)) + 2)) - 1
        
        Do
            On Error Resume Next
            res = CLng(InputBox(SheetNamesString_, "選択するシート名の番号を入力してください。" & loopExitNum & "で終了します。"))
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
        Loop While res < 0 Or UBound(SheetNames_) < res
    End If
    
    SheetName = SheetNames_(res)
    
    getSheetName = SheetName

End Function

Public Function getSheetNames() As String()
    Dim i As Long
    
    If Wb_ Is Nothing Then
        Call getBook
    End If
    
    ReDim SheetNames_(Wb_.Worksheets.Count - 1)
    
    For i = 1 To Wb_.Worksheets.Count
        SheetNames_(i - 1) = Wb_.Worksheets(i).Name
        SheetNamesString_ = SheetNamesString_ & (i - 1) & ":" & Wb_.Worksheets(i).Name & vbCrLf
    Next i

    getSheetNames = SheetNames_

End Function

'Err.Raise 1001, , "ファイルが選ばれませんでした。"
'Err.Raise 1002, , "フィルターが間違っています"
'Err.Raise 1003, , "ファイルが存在しません。"
'Err.Raise 1004, , "選択したファイルが指定されたファイル名と一致しません"
'Err.Raise 1011, , "シートが存在しません。:" & SheetName_
'Err.Raise 1099, , "ユーザーによる中断"
'Err.Raise 9999, , "不明なエラー"

