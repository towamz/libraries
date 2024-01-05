Option Explicit

Private GFFI As clsGetFolderFileInfo
Private DIC_PICTURE_PATH As Object
Private STR_PICTURE_PATH As String

'画像保存フォルダ名
Const STR_PICTURE_PATH_NAME As String = "picture"

Private Sub Class_Initialize()
    Set GFFI = New clsGetFolderFileInfo
    
    STR_PICTURE_PATH = ThisDocument.Path & "\" & STR_PICTURE_PATH_NAME
    
    Call checkFolderFiles

End Sub


Private Sub Class_Terminate()
    Set GFFI = Nothing
End Sub


Private Sub checkFolderFiles()

    If Not GFFI.isFolderExist(STR_PICTURE_PATH) Then
        MsgBox "保存フォルダがありません。作成します", vbOKOnly + vbInformation
        GFFI.makeFolder (STR_PICTURE_PATH)
        err.Raise 1000, , "フォルダを作成しました。画像を保存して再実行してください"
    ElseIf GFFI.getPictureCount(STR_PICTURE_PATH) = 0 Then
        err.Raise 1001, , "画像が1枚もありません。画像を保存して再実行してください"
    End If

End Sub


Private Sub makeDic(ByRef argDic As Object, ByRef argPath As String)
    Dim filesPath As Variant
    Dim filePath As Variant
    
    '既存のディクショナリは破棄する
    Set argDic = Nothing
    Set argDic = CreateObject("Scripting.Dictionary")
    
    'フォルダ・ファイル存在確認
    Call checkFolderFiles
    
    'ファイルパスを取得する
    filesPath = GFFI.getPictureArray(STR_PICTURE_PATH)

    'ディクショナリにファイルパスを格納する
    For Each filePath In filesPath
        argDic.Add filePath, 0
    Next

End Sub


Private Function getInfo(ByRef argDic As Object, ByRef argPath As String) As Variant
    Dim aryTmp(1) As Variant
    Dim lngRndNum As Long

    'ディクショナリがない・要素0の場合は再取得する
    If argDic Is Nothing Then
        Call makeDic(argDic, argPath)
    ElseIf argDic.Count = 0 Then
        Call makeDic(argDic, argPath)
    End If

    '乱数生成
    lngRndNum = Rnd * argDic.Count
    
    On Error GoTo errLabel
    aryTmp(0) = argDic.Keys()(lngRndNum)
    On Error GoTo 0

    aryTmp(1) = argDic.Items()(lngRndNum)
    argDic.Remove aryTmp(0)

    getInfo = aryTmp

    Exit Function

errLabel:
    lngRndNum = Rnd * argDic.Count
    Resume

End Function


Public Function getPictureInfo() As Variant

    getPictureInfo = getInfo(DIC_PICTURE_PATH, STR_PICTURE_PATH)

End Function
