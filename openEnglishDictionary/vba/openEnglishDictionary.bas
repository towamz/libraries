Const baseURL As String = "https://www.oxfordlearnersdictionaries.com/definition/english/"

Sub openEnglishDictionary()
    
    word = ActiveCell.Value
    
    If word = "" Then
        Exit Sub
    End If

    CreateObject("WScript.Shell").Run ("chrome.exe -url " & baseURL & word)
 
End Sub

