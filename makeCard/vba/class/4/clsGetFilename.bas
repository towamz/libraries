Option Explicit

Private FSO As Scripting.FileSystemObject
Private WT As clsWriteTextfile

Private DIC_PICTURE_INFO As Object
Private STR_PICTURE_PATH As String
Private STR_USED_INFO_PATH As String

'画像保存フォルダ名
Const STR_PICTURE_PATH_NAME As String = "picture"
'前回実行時使用ファイル名保存ファイル
Const STR_USED_INFO_NAME As String = "used.txt"
'ファイル名から情報を分割するデミリタ
Const STR_PICTURE_INFO_DELIMITER As String = "#"


Private Sub Class_Initialize()
    Set FSO = New Scripting.FileSystemObject
    Set WT = New clsWriteTextfile
    
    STR_PICTURE_PATH = ThisDocument.Path & "\" & STR_PICTURE_PATH_NAME
    STR_USED_INFO_PATH = ThisDocument.Path & "\" & STR_USED_INFO_NAME
    
    WT.setTextFilePath = STR_USED_INFO_PATH
    
    Call checkFolderFilesFirst

End Sub

Private Sub Class_Terminate()
    Set FSO = Nothing
    Set WT = Nothing
End Sub


Property Let setPicturePath(argPath As String)
    '相対パスの場合 / when relative path
    If InStr(argPath, "\") = 0 Then
        STR_PICTURE_PATH = ThisDocument.Path & "\" & argPath
    '絶対パスの場合 / when absolute path
    Else
        STR_PICTURE_PATH = argPath
    End If

End Property


Private Sub checkFolderFilesFirst()
    Call checkFolderFiles
        
    If FSO.FileExists(STR_USED_INFO_PATH) Then
        If FSO.GetFile(STR_USED_INFO_PATH).Size <> 0 Then
            If MsgBox("前回の使用済みファイル名が保存されています。初期化しますか", vbYesNo + vbDefaultButton2) = vbYes Then
                Kill STR_USED_INFO_PATH
                WT.createTextfile
            End If
        End If
    '使用済みファイル情報のテキストがない場合はここで空ファイルを作る
    Else
        WT.createTextfile
    End If
End Sub

Private Sub checkFolderFiles()
    If Not FSO.FolderExists(STR_PICTURE_PATH) Then
        MsgBox "保存フォルダがありません。作成します", vbOKOnly + vbInformation
        FSO.CreateFolder (STR_PICTURE_PATH)
        err.Raise 1000, , "フォルダを作成しました。画像を保存して再実行してください"
    ElseIf FSO.GetFolder(STR_PICTURE_PATH).Files.Count = 0 Then
        err.Raise 1001, , "画像が1枚もありません。画像を保存して再実行してください"
    End If
End Sub

Private Sub makeDic(ByRef argDic As Object, ByRef argPath As String)
    'Dim filesPath As Variant
    Dim objFile As Object
    Dim filePath As Variant
    Dim usedFilePath As Variant
    Dim backupDic As Object
    
    '既存のディクショナリは破棄する
    Set argDic = Nothing
    Set argDic = CreateObject("Scripting.Dictionary")
    Set backupDic = CreateObject("Scripting.Dictionary")
    
    'フォルダ・ファイル存在確認
    Call checkFolderFiles
    
    'ファイルパスを取得 / get filepath
    For Each objFile In FSO.GetFolder(argPath).Files
        'Debug.Print objFile.Path
        argDic.Add objFile.Path, 0
        backupDic.Add objFile.Path, 0
    Next
    
    '使用済みファイルに入力のあるファイル名をディクショナリから削除する / remove filename in the usedfile from dic
    WT.openTextfile (asRead)
    Do Until WT.isEOF
        DoEvents
        usedFilePath = WT.readTextfile(1)
        'Debug.Print usedFilePath
        
        If argDic.Exists(usedFilePath) Then
            argDic.Remove usedFilePath
        End If
    Loop
    
    'テキストファイルからのデータを削除した結果、count=0になった場合はバックアップをコピー
    If argDic.Count = 0 Then
        Debug.Print "dic keys all deleted, recover now"
        Set argDic = backupDic
        'ファイルは削除する
        WT.renewTextfile (notOpened)
    Else
        WT.closeTextfile
    End If
End Sub

Private Function getInfo(ByRef argDic As Object, ByRef argPath As String) As Variant
    Dim aryTmp(1) As Variant
    Dim aryInfoFromFilename As Variant
    Dim lngRndNum As Long

    'ディクショナリがない・要素0の場合は再取得する
    If argDic Is Nothing Then
        Debug.Print "dic make"
        Call makeDic(argDic, argPath)
    End If
    
    If argDic.Count = 0 Then
        'すべてのファイルを読み込んだので、テキストファイルを削除し空ファイルを作成する
        Debug.Print "dic remake"
        WT.renewTextfile
        Call makeDic(argDic, argPath)
    End If

    '乱数生成
    lngRndNum = Rnd * argDic.Count
    
    'ファイル名を取得する / get filename
    On Error GoTo errLabel
    aryTmp(0) = argDic.Keys()(lngRndNum)
    On Error GoTo 0

    '-----ファイル名をデミリタで分割して、情報を取得する / split filename by the demiliter to get info-----
    'ファイル名から2個目の情報を格納する / get info from the 2nd part of the filename
'    aryInfoFromFilename = Split(aryTmp(0), STR_PICTURE_INFO_DELIMITER)
'    aryTmp(1) = aryInfoFromFilename(1)
    '-----(終わり/End)ファイル名をデミリタで分割して、情報を取得する / split filename by the demiliter to get info-----
        
    '使用済みなのでディクショナリから削除する / remove  the filename from dic as it was used
    argDic.Remove aryTmp(0)
    '使用済みなのでファイル名をテキストファイルに書き込む / write the filename as it was used
    WT.writeTextfile (aryTmp(0))

    getInfo = aryTmp

    Exit Function

errLabel:
    lngRndNum = Rnd * argDic.Count
    Resume

End Function

Public Function getPictureInfo() As Variant

    getPictureInfo = getInfo(DIC_PICTURE_INFO, STR_PICTURE_PATH)

End Function
