Option Explicit
'Windows APIの宣言
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, _
     ByVal szURL As String, _
     ByVal szFileName As String, _
     ByVal dwReserved As Long, _
     ByVal lpfnCB As Long _
    ) As Long

Const baseURL As String = "https://www.oxfordlearnersdictionaries.com/definition/english/"
Const saveDir As String = "C:\Users\forwa\AppData\Roaming\Anki2\英検準1級単熟語EX2400\collection.media\"

Sub getMP3BySelenium()
    Dim Driver As New Selenium.WebDriver
    Dim elements  As WebElements
    
    Dim word As String
    Dim filename As String
    Dim sourceURL As String
    Dim gfp As New getFilenameParts
    
    word = ActiveCell.Value
    
    If word = "" Then
        Exit Sub
    End If
 
    Driver.Start "Chrome"
    Driver.Get baseURL & word
    
    Set elements = Driver.FindElementsByCss(".sound.audio_play_button.pron-us.icon-audio")
    
    sourceURL = elements.Item(1).Attribute("data-src-mp3")
    
    gfp.setFullFilename = sourceURL
    filename = gfp.getFilename
 
    If Dir(saveDir & filename) = "" Then
        
        Cells(ActiveCell.Row, 9).Value = sourceURL
        Call URLDownloadToFile(0, sourceURL, saveDir & filename, 0, 0)
        Call getMP3filename
    Else
        MsgBox "すでにダウンロード済みです"
    End If

End Sub
