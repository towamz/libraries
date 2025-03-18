Private FSO As Scripting.FileSystemObject
Private STR_TextfilePath As String
'ファイル番号は1~255の範囲なので0はファイルが閉じていることを示すことにする 
'0 indicate the file is not opened as fileno is between 1 and 255
Private LNG_FileNo As Long 
Private ENM_OpenmodeType As openmodeType
'ファイルを開くダイアログで表示するデフォルトファイル名 / the default filename for the openfile dialog
Private STR_DefaultFilename As String

Enum openmodeType
    asRead
    asWrite
    asAppend
    notOpened
End Enum


Private Sub Class_Initialize()
    Set FSO = New Scripting.FileSystemObject
    ENM_OpenmodeType = notOpened
    LNG_FileNo = 0
    STR_DefaultFilename = "newfile.txt"
    'STR_TextfilePath = ThisDocument.Path & "\" & STR_DEFAULT_FILENAME    
End Sub

Private Sub Class_Terminate()
    If LNG_FileNo <> 0 Then
        Call closeTextfile
    End If
    Set FSO = Nothing
End Sub


Property Let setTextFilePath(argPath As String)
    If LNG_FileNo = 0 Then
        STR_TextfilePath = argPath
    Else
        err.Raise 3, , "ファイルが開かれています。ファイルパスの変更はできません"
    End If
End Property

Property Let setFilename(argFilename As String)
    STR_DefaultFilename = argFilename
End Property


Public Function isEOF() As Boolean
    If LNG_FileNo = 0 Then
        err.Raise 2, , "ファイルは開かれていません"
    End If
    
    isEOF = EOF(LNG_FileNo)
End Function

Public Sub setCursorPosition(Optional argPosition As Long = 1)
    If LNG_FileNo = 0 Then
        err.Raise 2, , "ファイルは開かれていません"
    End If
    
    If argPosition < 1 Or argPosition > 2147483647 Then
        err.Raise 21, , "1 ～ 2,147,483,647を指定してください"
    End If
    
    Seek #LNG_FileNo, argPosition
End Sub

Public Sub writeTextfile(argWriteText As String)
    If LNG_FileNo = 0 Then
        Call openTextfile(asAppend)
    ElseIf Not (ENM_OpenmodeType = asAppend Or ENM_OpenmodeType = asWrite) Then
        err.Raise 13, , "書き込みモード以外で開かれています"
    End If
    
    Print #LNG_FileNo, argWriteText
End Sub

Public Function readTextfile(Optional argLineNumber As Long = 0) As String
    Dim lineText As String
    Dim lineCnt As Long
    
    If LNG_FileNo = 0 Then
        Call openTextfile(asRead)
    ElseIf ENM_OpenmodeType <> asRead Then
        err.Raise 12, , "読み込みモード以外で開かれています"
    End If

    '読み込み行数未指定/0の場合はEOFまで読み込む
    If argLineNumber = 0 Then
        Do Until EOF(LNG_FileNo)
            Line Input #LNG_FileNo, lineText
            readTextfile = readTextfile & lineText & vbCrLf
        Loop
    '読み込み行数指定の場合は指定の行数を読み込む(EOFで中断)
    Else
        For lineCnt = 1 To argLineNumber
            If EOF(LNG_FileNo) Then
                Exit For
            End If
            
            Line Input #LNG_FileNo, lineText
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
            'ファイルが存在することを確認 / confirm the file exists
            If Not FSO.FileExists(STR_TextfilePath) Then
                setFilenamesByDialog(msoFileDialogOpen)
            End If
            
            LNG_FileNo = FreeFile()
            Open STR_TextfilePath For Input As LNG_FileNo
            ENM_OpenmodeType = asRead
            
        Case openmodeType.asWrite
            'ファイルパスが存在するときは上書き警告を表示する / show alert when the file exists
            If FSO.FileExists(STR_TextfilePath) Then
                If MsgBox("指定されたテキストファイルが存在します。上書きしてもいいですか", vbOKCancel) = vbOK Then
                    Kill STR_TextfilePath
                Else
                    err.Raise 91, , "キャンセルが押されました"
                End If
            End If
            
            isFirstLoopForWrite = True
            Do
                DoEvents    '無限ループ回避用
                
                LNG_FileNo = FreeFile()
                On Error Resume Next
                Open STR_TextfilePath For Output As LNG_FileNo
                prevErrNumber = err.Number
                On Error GoTo 0
                
                '正常に開けた場合は、オープンモードを変更してループを抜ける
                If prevErrNumber = 0 Then
                    ENM_OpenmodeType = asWrite
                    Exit Do
                End If
                
                '異常な場合は、LNG_FileNo = 0に設定してから各処理をする
                LNG_FileNo = 0
                
                Select Case prevErrNumber
                    'ファイルが存在しない場合
                    Case 75
                        '初回はダイアログを表示してファイルを選択してもらう
                        If isFirstLoopForWrite Then
                            isFirstLoopForWrite = False
                            setFilenamesByDialog(msoFileDialogSaveAs)
                        'ダイアログで選択したファイルが存在しない場合は例外を投げる
                        Else
                            err.Raise 1, , "指定されたパスが存在しません"
                        End If
                    Case Else
                        err.Raise 99, , "エラーが発生しました"
                End Select
            Loop
        
        Case openmodeType.asAppend
            'ファイルが存在することを確認 / confirm the file exists
            If Not FSO.FileExists(STR_TextfilePath) Then
                setFilenamesByDialog(msoFileDialogOpen)
            End If
            
            LNG_FileNo = FreeFile()
            Open STR_TextfilePath For Append As LNG_FileNo
            ENM_OpenmodeType = asAppend
        
        Case Else
            err.Raise 11, , "オープンモードの指定が間違っています"

    End Select
End Sub

Public Sub closeTextfile()
    If LNG_FileNo = 0 Then
        err.Raise 2, , "ファイルは開かれていません"
    Else
        Close LNG_FileNo
        'ファイル番号を0に設定してファイルが開いていないことを明示する/set 0 to clarify a file is not opened
        LNG_FileNo = 0
        ENM_OpenmodeType = notOpened
    End If
End Sub

Public Sub renewTextfile(Optional argOpenmodeType As openmodeType = notOpened)
    Call closeTextfile
    Kill STR_TextfilePath
    Call openTextfile(asWrite)
    Call closeTextfile

    'openmodeがnotOpened以外の場合はそのモードで開く
    If argOpenmodeType <> notOpened Then
        Call openTextfile(argOpenmodeType)
    End If
End Sub

'テキストファイルを作成する
Public Sub createTextfile()
    Call openTextfile(asWrite)
    Call closeTextfile
End Sub

Private sub setFilenamesByDialog(argFileDialogType As Long) 
    Dim FD As FileDialog
    
    Set FD = Application.FileDialog(FileDialogType:=argFileDialogType)
    FD.AllowMultiSelect = False
    FD.InitialFileName = ThisDocument.Path & "\" & STR_DefaultFilename
    FD.FilterIndex = 13 '(.txt)
    FD.Show
    
    If FD.SelectedItems.Count = 0 Then
        err.Raise 91, , "キャンセルが押されました"
    End If

    STR_TextfilePath = FD.SelectedItems.Item(1)

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
