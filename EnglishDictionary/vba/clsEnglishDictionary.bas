Option Explicit

'Windows APIの宣言
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
     ByVal szURL As String, _
     ByVal szFileName As String, _
     ByVal dwReserved As Long, _
     ByVal lpfnCB As Long _
    ) As Long
    

Const dictionaryURL As String = "https://www.oxfordlearnersdictionaries.com/definition/english/"

Private sinfo As secretInfo
Private saveDir As String

Private Driver As Selenium.WebDriver
Private targetUrl As String


Private filename As String  'ファイル名(URLの最後の/以降)
Private mp3URL As String 'mp3のソースurl

Private columnEngWord As Integer
Private columnPos As Integer
Private columnJpnWord As Integer
Private columnEngDesc As Integer
Private columnEngEx As Integer
Private columnJpnEx As Integer
Private columnMP3Anki As Integer
Private columnMP3Url As Integer



Private Sub Class_Initialize()
    Set sinfo = New secretInfo
    Set Driver = New Selenium.WebDriver
    saveDir = sinfo.getSaveDir
    
    '各列の定義
    columnEngWord = 2
    columnPos = 3
    columnJpnWord = 4
    columnEngDesc = 5
    columnEngEx = 6
    columnJpnEx = 7
    columnMP3Anki = 8
    columnMP3Url = 9

End Sub



Public Sub setDriverOfPageBySelenium()
    '品詞=Part of speech
    Dim word As String
    Dim pos As String
    
    Dim posElement  As WebElement
    Dim posCnt As Integer
    Dim isDefLoopExit As Boolean
    
    word = Trim(Cells(ActiveCell.Row, columnEngWord).Value)
    
    If word = "" Then
        Err.Raise 1000, , "単語が設定されていません"
    End If
    
    '品詞の英訳を取得
    pos = getPosTranslation(Cells(ActiveCell.Row, columnPos).Value)
    
    
    Do
        'Set Driver = Nothing
        'Set Driver = New Selenium.WebDriver
        '最初のループではcloseでエラーになるので処理を続行するようにする
        On Error Resume Next
        Driver.Close
        On Error GoTo 0
        
        posCnt = posCnt + 1
        
        'URLをプロパティ変数に格納する
        targetUrl = dictionaryURL & word & "_" & posCnt
    
        Driver.AddArgument "headless"
        Driver.Start "Chrome"
        Driver.Get targetUrl
        
    
        Set posElement = Driver.FindElementByClass("webtop").FindElementByClass("pos")
        
        '品詞未入力の場合は都度メッセージボックスで確認する
        If pos = "" Then
            If MsgBox(posElement.Text, vbYesNo) = vbYes Then
                isDefLoopExit = True
            Else
                isDefLoopExit = False
            End If
        '品詞入力済みの場合は自動で比較する
        Else
            If posElement.Text = pos Then
                isDefLoopExit = True
            Else
                isDefLoopExit = False
            End If
        End If
    
    Loop Until isDefLoopExit

End Sub



Public Sub getMP3BySelenium()
    Dim elements  As WebElements
    
    Dim gfp As New getFilenameParts
    
    If Driver Is Nothing Then
        Err.Raise 1000, , "Driverが設定されていません"
    End If
    
    Set elements = Driver.FindElementsByCss(".sound.audio_play_button.pron-us.icon-audio")
    
    mp3URL = elements.Item(1).Attribute("data-src-mp3")
    
    gfp.setFullFilename = mp3URL
    filename = gfp.getFilename
 
    If Dir(saveDir & filename) = "" Then
        
        Cells(ActiveCell.Row, columnMP3Url).Value = mp3URL
        Call URLDownloadToFile(0, mp3URL, saveDir & filename, 0, 0)
        Cells(ActiveCell.Row, columnMP3Anki).Value = "[sound:" & filename & "]"
        
        Call playMP3
        
        
        'MsgBox "ダウンロードしました"
    Else
        Call playMP3
        MsgBox "すでにダウンロード済みです"
    End If
 
 
 
End Sub



Public Sub getWordDefinitionBySelenium()
    Dim defElements  As WebElements
    
    Dim pos As String   '品詞=Part of speech
    Dim defs As String
    
    Dim defCnt As Integer
    Dim ans As Integer
    
    Dim isDefLoopExit As Boolean

    If Driver Is Nothing Then
        Err.Raise 1000, , "Driverが設定されていません"
    End If






    Set defElements = Driver.FindElementsByClass("def")
    
    'すべての説明を結合する
    For defCnt = 1 To defElements.Count
        defs = defs & defCnt & vbCrLf & defElements.Item(defCnt).Text & vbCrLf
    Next
    
    
    
    
    
    Do
        ans = InputBox("入力する定義を選択してください" & vbCrLf & _
                                "96=何もしないで終了" & vbCrLf & _
                                "97=webページを開く" & vbCrLf & _
                                "98=全ての定義を入力して終了" & vbCrLf & _
                                "99=全ての定義を入力してwebページを開く" & vbCrLf & vbCrLf & _
                                defs)
        
        Select Case ans
            Case 96
                '何もしない
                isDefLoopExit = True
            Case 97
                Call openEnglishDictionary
                isDefLoopExit = True
            Case 98
                Cells(ActiveCell.Row, columnEngDesc).Value = defs
                Cells(ActiveCell.Row, columnEngDesc).font.Size = 8
                isDefLoopExit = True
            Case 99
                Cells(ActiveCell.Row, columnEngDesc).Value = defs
                Cells(ActiveCell.Row, columnEngDesc).font.Size = 8
                Call openEnglishDictionary
                isDefLoopExit = True
            Case Else
                isDefLoopExit = True
                On Error GoTo errLabelDef
                Cells(ActiveCell.Row, columnEngDesc).Value = defElements.Item(ans).Text
                Cells(ActiveCell.Row, columnEngDesc).font.Size = 8

        End Select
    
    Loop Until isDefLoopExit
    
    Exit Sub

errLabelDef:
    Select Case Err.Number
        Case -2146233080
            MsgBox "範囲外の番号が指定されました。選択しなおしてください"
            isDefLoopExit = False
            Resume Next
        Case Else
            Stop
    End Select
End Sub





Private Function getPosTranslation(argPos As String) As String

    Select Case argPos
        Case "modal verb"
            getPosTranslation = "助動詞"
        Case "助動詞"
            getPosTranslation = "modal verb"
            
        Case "adjective"
            getPosTranslation = "形容詞"
        Case "形容詞"
            getPosTranslation = "adjective"
            
        Case "verb"
            getPosTranslation = "動詞"
        Case "動詞"
            getPosTranslation = "verb"
        
        Case "noun"
            getPosTranslation = "名詞"
        Case "名詞"
            getPosTranslation = "noun"

        'Case ""
        '    getPosTranslation = ""
        'Case ""
        '    getPosTranslation = ""
        
        'Case ""
        '    getPosTranslation = ""
        'Case ""
        '    getPosTranslation = ""

        'Case ""
        '    getPosTranslation = ""
        'Case ""
        '    getPosTranslation = ""
        
        'Case ""
        '    getPosTranslation = ""
        'Case ""
        '    getPosTranslation = ""
        
        'Case ""
        '    getPosTranslation = ""
        'Case ""
        '    getPosTranslation = ""
        
        'Case ""
        '    getPosTranslation = ""
        'Case ""
        '    getPosTranslation = ""
        
        'Case ""
        '    getPosTranslation = ""
        'Case ""
        '    getPosTranslation = ""
        
        Case Else
            getPosTranslation = InputBox("未登録の品詞/品詞が未入力です。英語の品詞を入力してください。未入力の場合はページ毎に品詞の確認をします")
            
            'デバッグ用コマンド
            If LCase(getPosTranslation) = "stop" Then
                Stop
            End If
            
    End Select

End Function


Public Sub playMP3()
    Dim gfp As New getFilenameParts
    Dim wsh As Object
    
    'mp3ファイルを取得するとfilenameメンバ変数に格納するので値が設定されているか確認する
    If filename = "" Then
        'エクセルシート上にmp3のurlがない場合は例外を投げる
        If Cells(ActiveCell.Row, columnMP3Url).Value = "" Then
            Err.Raise 1000, , "ファイル名が不明です"
        'ある場合は、ファイル名を取得する
        Else
            gfp.setFullFilename = Cells(ActiveCell.Row, columnMP3Url).Value
            filename = gfp.getFilename
        End If
    End If
    
    Set wsh = CreateObject("Wscript.Shell")
    wsh.Run saveDir & filename
    Application.Wait (Now + TimeValue("00:00:05"))
    wsh.SendKeys "%{F4}"

End Sub



Public Sub openEnglishDictionary()
    
    On Error GoTo errLabel
    CreateObject("WScript.Shell").Run ("chrome.exe -url " & Driver.Url)

    Exit Sub
errLabel:
    Select Case Err.Number
        'webDriverスタートしていない場合は取得する
        Case 57
            Call setDriverOfPageBySelenium
            Resume
    
        Case Else
            Debug.Print Err.Number
            Debug.Print Err.Description
    
            Stop
    End Select
        
End Sub


