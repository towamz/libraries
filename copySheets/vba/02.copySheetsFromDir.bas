Option Explicit

Const STR_DEFAULT_DIR = "C:\"

'ディレクトリ内のすべてのcsvをシートとしてコピーする / copy all csvs that are in the specified directory 
Sub importCSV()

    Dim getDir As String
    Dim filename As String
    
    Dim wb As Workbook
    Dim wbData As Workbook
    
    Dim ws As Worksheet
    Dim wsData As Worksheet
    
    'カレントディレクトリ変更  / change the current directory
    With CreateObject("WScript.Shell")
        .CurrentDirectory = STR_DEFAULT_DIR
    End With
    
    'フォルダ選択ボックス / show a folder picker
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            getDir = .SelectedItems(1)
        Else
            MsgBox "フォルダが選ばれませんでした。処理を中断します。"
            Exit Sub
        End If
    End With
    
    'カレントディレクトリ再変更  / rechange the current directory
    With CreateObject("WScript.Shell")
        .CurrentDirectory = getDir
    End With
    
    'csvシートをのコピー先ワークブック / set a workbook that gather csvs.
    Set wb = ThisWorkbook
    
    filename = Dir(getDir & "\*")
    Do While filename <> ""
        
        'データファイルを開く / open the data file
        Set wbData = Workbooks.Open(filename:=filename, ReadOnly:=True)
    
        'データシートを取得(csvはシート1枚のみ) / get target sheet(csv has only one sheet)
        Set wsData = wbData.Worksheets(1)
    
        'ファイル内容チェック/ check data contents
        If wsData.Range("A4").Value <> "年月日" Then
            MsgBox "ファイルの内容が間違っています"
            Exit Sub
        ElseIf wsData.Range("B4").Value <> "平均気温(℃)" Then
            MsgBox "ファイルの内容が間違っています"
            Exit Sub
        End If
        
        '末尾にシートをコピーする/ copy the sheet to the end of the book
        wsData.Copy after:=wb.Sheets(wb.Sheets.Count)
        
        wbData.Close
    
        filename = Dir()
    Loop

End Sub

