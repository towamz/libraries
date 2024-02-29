Private FSO As Scripting.FileSystemObject
Private STR_TEXTFILE_PATH As String
Private LNG_FILE_NO As Long 'ファイル番号は1~255の範囲なので0を設定するとエラーになる
Private ENM_OPENMODE_TYPE As openmodeType

Private STR_Filename As String
Const STR_DEFAULT_FILENAME As String = "newfile.txt"

Enum openmodeType
    asRead
    asWrite
    asAppend
    notOpened
End Enum

Private Sub Class_Initialize()
    Set FSO = New Scripting.FileSystemObject
    ENM_OPENMODE_TYPE = notOpened
    LNG_FILE_NO = 0
    STR_Filename = STR_DEFAULT_FILENAME
    'STR_TEXTFILE_PATH = ThisDocument.Path & "\" & STR_DEFAULT_FILENAME
    
End Sub

Private Sub Class_Terminate()
    If LNG_FILE_NO <> 0 Then
        Call closeTextfile
    End If
    Set FSO = Nothing
End Sub

Property Let setTextFilePath(argPath As String)
    If LNG_FILE_NO = 0 Then
        STR_TEXTFILE_PATH = argPath
    Else
        err.Raise 3, , "ファイルが開かれています。ファイルパスの変更はできません"
    End If
End Property

Property Let setFilename(argFilename As String)
    STR_Filename = argFilename
End Property

Public Function isEOF() As Boolean
    If LNG_FILE_NO = 0 Then
        err.Raise 2, , "ファイルは開かれていません"
    End If
    
    isEOF = EOF(LNG_FILE_NO)
End Function


Public Sub setCursorPosition(Optional argPosition As Long = 1)
    If LNG_FILE_NO = 0 Then
        err.Raise 2, , "ファイルは開かれていません"
    End If
    
    If argPosition < 1 Or argPosition > 2147483647 Then
        err.Raise 21, , "1 ～ 2,147,483,647を指定してください"
    End If
    
    Seek #LNG_FILE_NO, argPosition
End Sub

Public Sub writeTextfile(argWriteText As String)
    If LNG_FILE_NO = 0 Then
        Call openTextfile(asAppend)
    ElseIf Not (ENM_OPENMODE_TYPE = asAppend Or ENM_OPENMODE_TYPE = asWrite) Then
        err.Raise 13, , "書き込みモード以外で開かれています"
    End If
    
    Print #LNG_FILE_NO, argWriteText
End Sub

Public Function readTextfile(Optional argLineNumber As Long = 0) As String
    Dim lineText As String
    Dim lineCnt As Long
    
    If LNG_FILE_NO = 0 Then
        Call openTextfile(asRead)
    ElseIf ENM_OPENMODE_TYPE <> asRead Then
        err.Raise 12, , "読み込みモード以外で開かれています"
    End If

    '読み込み行数未指定/0の場合はEOFまで読み込む
    If argLineNumber = 0 Then
        Do Until EOF(LNG_FILE_NO)
            Line Input #LNG_FILE_NO, lineText
            readTextfile = readTextfile & lineText & vbCrLf
        Loop
    '読み込み行数指定の場合は指定の行数を読み込む(EOFで中断)
    Else
        For lineCnt = 1 To argLineNumber
            If EOF(LNG_FILE_NO) Then
                Exit For
            End If
            
            Line Input #LNG_FILE_NO, lineText
            readTextfile = readTextfile & lineText & vbCrLf
        Next
    End If
    
    '最後の改行記号を削除する
    If Right(readTextfile, 2) = vbCrLf Then
        readTextfile = Left(readTextfile, Len(readTextfile) - 2)
    End If
End Function

Public Sub openTextfile(argOpenmodeType As openmodeType)
    Dim isFirstLoopForWrite As Boolean
    Dim prevErrNumber As Long
    
    Select Case argOpenmodeType
        Case openmodeType.asRead
            'ファイルが存在することを確認 / confirm the file is exist
            If Not FSO.FileExists(STR_TEXTFILE_PATH) Then
                STR_TEXTFILE_PATH = getFilenamesByDialog(msoFileDialogOpen)
            End If
            
            LNG_FILE_NO = FreeFile()
            Open STR_TEXTFILE_PATH For Input As LNG_FILE_NO
            ENM_OPENMODE_TYPE = asRead
            
        Case openmodeType.asWrite
            'ファイルパスが存在するときは上書き警告を表示する / show alert when the file is exist
            If FSO.FileExists(STR_TEXTFILE_PATH) Then
                If MsgBox("指定されたテキストファイルが存在します。上書きしてもいいですか", vbOKCancel) = vbOK Then
                    Kill STR_TEXTFILE_PATH
                Else
                    err.Raise 91, , "キャンセルが押されました"
                End If
            End If
            
            isFirstLoopForWrite = True
            Do
                DoEvents    '無限ループ回避用
                
                LNG_FILE_NO = FreeFile()
                On Error Resume Next
                Open STR_TEXTFILE_PATH For Output As LNG_FILE_NO
                prevErrNumber = err.Number
                On Error GoTo 0
                
                '正常に開けた場合は、オープンモードを変更してループを抜ける
                If prevErrNumber = 0 Then
                    ENM_OPENMODE_TYPE = asWrite
                    Exit Do
                End If
                
                '異常な場合は、LNG_FILE_NO = 0に設定してから各処理をする
                LNG_FILE_NO = 0
                
                Select Case prevErrNumber
                    'ファイルが存在しない場合
                    Case 75
                        '初回はダイアログを表示してファイルを選択してもらう
                        If isFirstLoopForWrite Then
                            isFirstLoopForWrite = False
                            STR_TEXTFILE_PATH = getFilenamesByDialog(msoFileDialogSaveAs)
                        'ダイアログで選択したファイルが存在しない場合は例外を投げる
                        Else
                            err.Raise 1, , "指定されたパスが存在しません"
                        End If
                    Case Else
                        err.Raise 99, , "エラーが発生しました"
                End Select
            Loop
        
        Case openmodeType.asAppend
            'ファイルが存在することを確認 / confirm the file is exist
            If Not FSO.FileExists(STR_TEXTFILE_PATH) Then
                STR_TEXTFILE_PATH = getFilenamesByDialog(msoFileDialogOpen)
            End If
            
            LNG_FILE_NO = FreeFile()
            Open STR_TEXTFILE_PATH For Append As LNG_FILE_NO
            ENM_OPENMODE_TYPE = asAppend
        
        Case Else
            err.Raise 11, , "オープンモードの指定が間違っています"
    End Select
End Sub


Public Sub closeTextfile()
    If LNG_FILE_NO = 0 Then
        err.Raise 2, , "ファイルは開かれていません"
    Else
        Close LNG_FILE_NO
        'ファイル番号を0に設定してファイルが開いていないことを明示する/set 0 to clarify a file is not opened
        LNG_FILE_NO = 0
        ENM_OPENMODE_TYPE = notOpened
    End If
End Sub

Public Sub renewTextfile()
    Call closeTextfile
    Kill STR_TEXTFILE_PATH
    Call openTextfile(asWrite)
    Call closeTextfile
End Sub

'テキストファイルを作成する
Public Sub createTextfile()
    Call openTextfile(asWrite)
    Call closeTextfile
End Sub

Private Function getFilenamesByDialog(argFileDialogType As Long) As String
    Dim FD As FileDialog
    
    Set FD = Application.FileDialog(FileDialogType:=argFileDialogType)
    FD.AllowMultiSelect = False
    FD.InitialFileName = ThisDocument.Path & "\" & STR_Filename
    FD.FilterIndex = 13 '(.txt)
    FD.Show
    
    If FD.SelectedItems.Count = 0 Then
        err.Raise 91, , "キャンセルが押されました"
    End If

    getFilenamesByDialog = FD.SelectedItems.Item(1)

End Function


'err.Raise 1, , "指定されたパスが存在しません"
'err.Raise 2, , "ファイルは開かれていません"
'err.Raise 3, , "ファイルが開かれています。ファイルパスの変更はできません"
'err.Raise 11, , "オープンモードの指定が間違っています"
'err.Raise 12, , "読み込みモード以外で開かれています"
'err.Raise 13, , "書き込みモード以外で開かれています"
'err.Raise 21, , "1 ～ 2,147,483,647を指定してください"
'err.Raise 91, , "キャンセルが押されました"
'err.Raise 99, , "エラーが発生しました"
