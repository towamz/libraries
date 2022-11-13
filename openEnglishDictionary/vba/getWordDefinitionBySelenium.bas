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
        defs = defs & defElements.Item(defCnt).Text & vbCrLf & "---" & vbCrLf
    Next
    
    Cells(ActiveCell.Row, 5).Value = defs
    
    If MsgBox("ページを開きますか?" & vbCrLf & vbCrLf & defs, vbYesNo) = vbYes Then
        Call openEnglishDictionary
    End If

 
End Sub
