Private GFFI As clsGetFolderFileInfo
Private STR_TEXTFILE_PATH As String
Private LNG_FILE_NO As Long 'ファイル番号は1~255の範囲なので0を設定するとエラーになる
Private ENM_OPENMODE_TYPE As openmodeType

Enum openmodeType
    asRead
    asWrite
    asAppend
    notDefined
End Enum


Private Sub Class_Initialize()
    Set GFFI = New clsGetFolderFileInfo
    ENM_OPENMODE_TYPE = notDefined
    LNG_FILE_NO = 0
End Sub


Private Sub Class_Terminate()
    If LNG_FILE_NO <> 0 Then
        Call closeTextfile
    End If
End Sub


Property Let setTextFilePath(argPath As String)
    If LNG_FILE_NO = 0 Then
        STR_TEXTFILE_PATH = argPath
    Else
        err.Raise 9010, , "ファイルが開かれています。ファイルパスの変更はできません"
    End If
End Property


Public Sub setPositionFirst()
    If LNG_FILE_NO = 0 Then
        err.Raise 9011, , "ファイルは開かれていません"
    End If
    
    Seek #LNG_FILE_NO, 1

End Sub


Public Sub setPositionSpecified(argPosition As Long)
    
    If LNG_FILE_NO = 0 Then
        err.Raise 9011, , "ファイルは開かれていません"
    End If
    
    If argPosition < 1 Or argPosition > 2147483647 Then
        err.Raise 9031, , "1 ～ 2,147,483,647を指定してください"
    End If
    
    Seek #LNG_FILE_NO, argPosition

End Sub


Public Sub writeTextfile(argWriteText As String)

    If LNG_FILE_NO = 0 Then
        Call openTextfile(asAppend)
    ElseIf Not (ENM_OPENMODE_TYPE = asAppend Or ENM_OPENMODE_TYPE = asWrite) Then
        err.Raise 9012, , "書き込みモード以外で開かれています"
    End If
    
    Print #LNG_FILE_NO, argWriteText

End Sub


Public Function readTextfile(Optional argLineNumber As Long) As String
    Dim lineText As String
    Dim lineCnt As Long
    
    If LNG_FILE_NO = 0 Then
        Call openTextfile(asRead)
    ElseIf ENM_OPENMODE_TYPE <> asRead Then
        err.Raise 9011, , "読み込みモード以外で開かれています"
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
    If Right(readTextfile, 1) = vbCrLf Then
        readTextfile = Left(readTextfile, Len(readTextfile) - 1)
    End If
    
End Function


Public Sub openTextfile(argOpenmodeType As openmodeType)
    Select Case argOpenmodeType
        Case openmodeType.asRead
            'ファイルが存在することを確認 / confirm the file is exist
            If GFFI.isFileExist(STR_TEXTFILE_PATH) Then
                LNG_FILE_NO = FreeFile()
                Open STR_TEXTFILE_PATH For Input As LNG_FILE_NO
                ENM_OPENMODE_TYPE = asRead
            Else
                err.Raise 9002, , "指定されたパスが存在しません"
            End If
        Case openmodeType.asWrite
            'ファイルパスが存在するときは上書き警告を表示する / show alert when the file is exist
            If GFFI.isFileExist(STR_TEXTFILE_PATH) Then
                If MsgBox("指定されたテキストファイルが存在します。上書きしてもいいですか", vbOKCancel) = vbOK Then
                    Kill STR_TEXTFILE_PATH
                Else
                    err.Raise 9099, , "ユーザーによる中断"
                End If
            End If
            
            LNG_FILE_NO = FreeFile()
            Open STR_TEXTFILE_PATH For Output As LNG_FILE_NO
            ENM_OPENMODE_TYPE = asWrite
        
        Case openmodeType.asAppend
            'ファイルが存在することを確認 / confirm the file is exist
            If GFFI.isFileExist(STR_TEXTFILE_PATH) Then
                LNG_FILE_NO = FreeFile()
                Open STR_TEXTFILE_PATH For Append As LNG_FILE_NO
                ENM_OPENMODE_TYPE = asAppend
            Else
                err.Raise 9002, , "指定されたパスが存在しません"
            End If
    End Select
End Sub


Public Sub closeTextfile()

    If LNG_FILE_NO = 0 Then
        err.Raise 9011, , "ファイルは開かれていません"
    Else
        Close LNG_FILE_NO
        'ファイル番号を0に設定してファイルが開いていないことを明示する/set 0 to clarify a file is not opened
        LNG_FILE_NO = 0
        ENM_OPENMODE_TYPE = notDefined
    End If
End Sub

