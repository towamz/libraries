Option Explicit

Const STR_DIR = "C:\sampleMacro"

Const STR_DATA_BOOK As String = "newly_confirmed_cases_daily.csv"
'コピー元シート名
Const STR_DATA_SHEET As String = "newly_confirmed_cases_daily"

'新シート名(空白の場合は、シート名を変更しない) / new sheet name (if no need to change, specify blank)
Const STR_NEW_SHEET As String = "daily"




'ファイル1個用 / for one file
Sub openCopyFile()
    Dim fullFilename As String
    Dim filename As String
    
    Dim wb As Workbook
    Dim wbData As Workbook
    
    Dim ws As Worksheet
    Dim wsData As Worksheet
    
    'カレントディレクトリ変更  / change the current directory
    With CreateObject("WScript.Shell")
        .CurrentDirectory = STR_DIR
    End With

    'ファイル名(フルパス)取得  / get filename(full path)
    fullFilename = Application.GetOpenFilename(FileFilter:="ファイル,*.*", Title:="ファイルを選んでください")
    
    'キャンセルしたときは中止 / abort when cancel was pushed
    If filename = "False" Then
        MsgBox "ファイルが選ばれませんでした"
        Exit Sub
    End If
    
    
    'ファイル名(パスなし)取得  / get filename(no path)
    filename = Mid(fullFilename, InStrRev(fullFilename, "\") + 1)
    
    
    'ファイル名確認 / check the filename
    'ファイル名が固定の時使う / use this when the filename is fixed.
    If filename <> STR_DATA_BOOK Then
        MsgBox "間違ったファイルが選択されました"
        Exit Sub
    End If
    
    'ファイル名が可変部がある時使う / use this when the filename has variable part.
'    If InStr(filename, STR_DATA_BOOK) = 0 Then
'        MsgBox "間違ったファイルが選択されました"
'        Exit Sub
'
'    End If

    


    Set wb = ThisWorkbook

    'データファイルを開く / open the data file
    Set wbData = Workbooks.Open(filename:=fullFilename, ReadOnly:=True)

    'データシートを取得 / get target sheet
    'シート名が固定の時使う / use this when the sheetname is fixed.
    Set wsData = wbData.Worksheets(STR_DATA_SHEET)
    
    
    'シート名が可変部がある時使う / use this when the sheetname has variable part.
'    For Each ws In wbData.Worksheets
'        If InStr(ws.Name, STR_DATA_SHEET) <> 0 Then
'            Set wsData = ws
'            Exit For
'        End If
'    Next
    
    'シートが見つからなかった場合は警告を表示してデータファイルを閉じる処理へ飛ぶ
    ' / if there is no specified sheet, show alert and jump to finally to close the data file.
    If wsData Is Nothing Then
        MsgBox "指定のシートが見つかりませんでした"
        GoTo finally
    
    End If
    
    
    
    'ファイル内容チェック/ check data contents
    If wsData.Range("A1").Value <> "Date" Then
        MsgBox "ファイルの内容が間違っています"
        Exit Sub
    ElseIf wsData.Range("B1").Value <> "ALL" Then
        MsgBox "ファイルの内容が間違っています"
        Exit Sub
    ElseIf wsData.Range("C1").Value <> "Hokkaido" Then
        MsgBox "ファイルの内容が間違っています"
        Exit Sub
    End If



    '末尾にシートをコピーする/ copy the sheet to the end of the book
    wsData.Copy after:=wb.Sheets(wb.Sheets.Count)
    
    'シート名を変更(指定があるとき) / change the sheet name (when it was specified)
    If STR_NEW_SHEET <> "" Then
        ActiveSheet.Name = STR_NEW_SHEET
    End If
    
    
finally:
    'データファイルを閉じる/ close the data file
    wbData.Close SaveChanges:=False


End Sub

