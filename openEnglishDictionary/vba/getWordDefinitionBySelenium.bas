Option Explicit

Const baseURL As String = "https://www.oxfordlearnersdictionaries.com/definition/english/"
Const saveDir As String = "C:\anki\collection.media\"

Sub getWordDefinitionBySelenium()
    Dim Driver As New Selenium.WebDriver
    Dim posElement  As WebElement
    Dim defElements  As WebElements
    
    Dim word As String
    Dim pos As String   '品詞=Part of speech
    Dim filename As String
    Dim defs As String
    Dim gfp As New getFilenameParts
    
    Dim posCnt As Integer
    Dim defCnt As Integer
    Dim ans As Integer
    
    Dim isPosLoopExit As Boolean
    Dim isDefLoopExit As Boolean
    
    word = ActiveCell.Value
    
    If word = "" Then
        Exit Sub
    End If
    
    '品詞の英訳を取得
    pos = getPosTranslation(Cells(ActiveCell.Row, 3).Value)
    
    
    Do
        'Set Driver = Nothing
        'Set Driver = New Selenium.WebDriver
        '最初のループではcloseでエラーになるので処理を続行するようにする
        On Error Resume Next
        Driver.Close
        On Error GoTo 0
        
        posCnt = posCnt + 1
        
    
        Driver.AddArgument "headless"
        Driver.Start "Chrome"
        Driver.Get baseURL & word & "_" & posCnt
        
    
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
    
    
    
    Set defElements = Driver.FindElementsByClass("def")
    
    
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
                Cells(ActiveCell.Row, 5).Value = defs
                Cells(ActiveCell.Row, 5).font.Size = 8
                isDefLoopExit = True
            Case 99
                Cells(ActiveCell.Row, 5).Value = defs
                Cells(ActiveCell.Row, 5).font.Size = 8
                Call openEnglishDictionary
                isDefLoopExit = True
            Case Else
                isDefLoopExit = True
                On Error GoTo errLabelDef
                Cells(ActiveCell.Row, 5).Value = defElements.Item(ans).Text
                Cells(ActiveCell.Row, 5).font.Size = 8

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


Function getPosTranslation(pos As String) As String

    Select Case pos
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


        Case Else
            Stop
    End Select




End Function
