'https://y-moride.com/vba/dailog-folder-picker.html

Dim FOLDER
Dim path

With CreateObject("Shell.Application")        
    'フォルダ選択ダイアログを表示
    '&H10 =  選択中のフォルダ名をテキストボックスに表示する
    '&H200 = 「新しいフォルダの作成」を表示しない
    Set FOLDER = .BrowseForFolder(0, "フォルダを選んでください" , &H200 , "C:\")
End With
If FOLDER Is Nothing Then WScript.Quit

path = FOLDER.Self.Path
wscript.echo path

