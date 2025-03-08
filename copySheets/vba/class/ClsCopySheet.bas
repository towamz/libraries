Option Explicit

Private DefaultDirectory As String
Private DataBook As String
Private DataBookFullname As String
Private DataSheet As String
Private DataSheetRename As String
Private DataSheetCheckDictionary As Object

'デフォルトディレクトリを設定する(任意)
Public Property Let setDefaultDirectory(arg1 As String)
    DefaultDirectory = arg1
End Property

'コピー対象のブック名を設定する(任意)
Public Property Let setDataBook(arg1 As String)
    DataBook = arg1
End Property

'コピー対象のブック名(フルパス)を設定する(任意)
Public Property Let setDataBookFullname(arg1 As String)
    DataBookFullname = arg1
End Property

'コピー対象のシート名を設定する
Public Property Let setDataSheet(arg1 As String)
    DataSheet = arg1
End Property

'新シート名を設定する(任意)
Public Property Let setDataSheetRename(arg1 As String)
    DataSheetRename = arg1
End Property

Private Sub Class_Initialize()
    'データチェック用ディクショナリの初期化
    Set DataSheetCheckDictionary = CreateObject("Scripting.Dictionary")
End Sub

'データシートの列名チェックを設定する(任意)
Public Sub setDataSheetCheckKey(arg1 As String, arg2 As String)
    Dim testRange As Range
    Dim errNumber As Long
    'セルの有効性確認
    On Error Resume Next
    Set testRange = Range(arg1)
    
    'エラー処理無効にしないと例外を投げられないのでエラー番号を取得する
    errNumber = Err.Number
    On Error GoTo 0
    
    Select Case errNumber
        Case 0
            '何もしない
        Case 1004
            Err.Raise 1001, , "有効なセル番地を設定してください"
        Case Else
            Err.Raise 9999, , "不明なエラーが発生しました"
    End Select

    DataSheetCheckDictionary.Add arg1, arg2
End Sub

'ファイル1個用 / for one file
Public Sub copySheet()
    Dim fullFilename As String
    Dim filename As String
    
    Dim wb As Workbook
    Dim wbData As Workbook
    
    Dim ws As Worksheet
    Dim wsData As Worksheet
    
    Dim dicKey As Variant
    
    'コピー対象ブックが存在するか確認する
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(DataBookFullname) Then
            fullFilename = DataBookFullname
        End If
    End With
    
    'ブック(フルパス)が設定されていない、または存在しない場合はファイルを開くダイアログを表示する
    If fullFilename = "" Then
        'カレントディレクトリ変更  / change the current directory
        If DefaultDirectory <> "" Then
            With CreateObject("WScript.Shell")
                .CurrentDirectory = DefaultDirectory
            End With
        End If
    
        'ファイル名(フルパス)取得  / get filename(full path)
        fullFilename = Application.GetOpenFilename(FileFilter:="ファイル,*.*", Title:="ファイルを選んでください")
        
        'キャンセルしたときは中止 / abort when cancel was pushed
        If fullFilename = "False" Then
            MsgBox "ファイルが選ばれませんでした"
            Exit Sub
        End If
    End If
    
    'ファイル名(パスなし)取得  / get filename(no path)
    filename = Mid(fullFilename, InStrRev(fullFilename, "\") + 1)
    
    'ファイル名確認 / check the filename
    'ファイル名が固定の時使う / use this when the filename is fixed.
    If DataBook <> "" Then
        If filename <> DataBook Then
            MsgBox "間違ったファイルが選択されました"
            Exit Sub
        End If
    End If
    
    'ファイル名が可変部がある時使う / use this when the filename has variable part.
'    If DataBook <> "" Then
'        If InStr(filename, STR_DATA_BOOK) = 0 Then
'            MsgBox "間違ったファイルが選択されました"
'            Exit Sub
'        End If
'    End If

    Set wb = ThisWorkbook

    'データファイルを開く / open the data file
    Set wbData = Workbooks.Open(filename:=fullFilename, ReadOnly:=True)

    'データシートを取得 / get target sheet
    If DataSheet = "" Then
        'データシート名の指定がないときは一番左のシートをコピーする/ copy the left sheet when not specified
        Set wsData = wbData.Worksheets(1)
    Else
        'シート名が固定の時使う / use this when the sheetname is fixed.
        Set wsData = wbData.Worksheets(DataSheet)
        
        'シート名が可変部がある時使う / use this when the sheetname has variable part.
    '    For Each ws In wbData.Worksheets
    '        If InStr(ws.Name, DataSheet) <> 0 Then
    '            Set wsData = ws
    '            Exit For
    '        End If
    '    Next
    End If
    
    'シートが見つからなかった場合は警告を表示してデータファイルを閉じる処理へ飛ぶ
    ' / if there is no specified sheet, show alert and jump to finally to close the data file.
    If wsData Is Nothing Then
        MsgBox "指定のシートが見つかりませんでした"
        GoTo finally
    End If
    
    'ファイル内容チェック/ check data contents
    For Each dicKey In DataSheetCheckDictionary
        If wsData.Range(dicKey).Value <> DataSheetCheckDictionary.Item(dicKey) Then
            Err.Raise 2001, , "ファイルの内容が間違っています"
            Exit Sub
        End If
    Next

    '末尾にシートをコピーする/ copy the sheet to the end of the book
    wsData.Copy after:=wb.Sheets(wb.Sheets.Count)
    
    'シート名を変更(指定があるとき) / change the sheet name (when it was specified)
    If DataSheetRename <> "" Then
        ActiveSheet.Name = DataSheetRename
    End If


finally:
    'データファイルを閉じる/ close the data file
    wbData.Close SaveChanges:=False

End Sub

