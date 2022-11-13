Option Explicit

Const baseURL As String = "https://www.oxfordlearnersdictionaries.com/definition/english/"
Const saveDir As String = "C:\anki\collection.media\"

Sub getWordDefinitionBySelenium()
    Dim Driver As New Selenium.WebDriver
    Dim posElement  As WebElement
    Dim defElements  As WebElements
    
    Dim word As String
    Dim filename As String
    Dim defs As String
    Dim gfp As New getFilenameParts
    
    Dim posCnt As Integer
    Dim defCnt As Integer
    Dim ans As Integer
    
    Dim isLoopExit As Boolean
    
    word = ActiveCell.Value
    
    If word = "" Then
        Exit Sub
    End If
    
    
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
    
    Loop Until MsgBox(posElement.Text, vbYesNo) = vbYes
    
    
    
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
                isLoopExit = True
            Case 97
                Call openEnglishDictionary
                isLoopExit = True
            Case 98
                Cells(ActiveCell.Row, 5).Value = defs
                isLoopExit = True
            Case 99
                Cells(ActiveCell.Row, 5).Value = defs
                Call openEnglishDictionary
                isLoopExit = True
            Case Else
                isLoopExit = True
                On Error GoTo errLabel
                Cells(ActiveCell.Row, 5).Value = defElements.Item(ans).Text

        End Select
    
    Loop Until isLoopExit
    
    Exit Sub

errLabel:
    Select Case Err.Number
        Case -2146233080
            MsgBox "範囲外の番号が指定されました。選択しなおしてください"
            isLoopExit = False
            Resume Next
        Case Else
            Stop
    End Select
End Sub
