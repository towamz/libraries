Option Explicit

Private FSO As Object
Private GDR As getDataRows
Private GFN As getFilenames

Private templateWorkbookName As String
Private templateWorksheetName As String
Private margeWorksheetName As String
Private dataWorksheetName As String

Private margeWorksheet As Worksheet

Private workbookNameColumn As Long  'データワークブック名を挿入する列(任意)
Private serialNumbersColumn As Long  '連番付与列

'テンプレートブック名
Public Property Let setTemplateWorkbookName(workbookString As String)
    templateWorkbookName = workbookString
End Property

'テンプレートシート名
Public Property Let setTemplateWorksheetName(worksheetString As String)
    templateWorksheetName = worksheetString
End Property

'マージシート名(各シートからデータを取りまとめるシート)
Public Property Let setMargeWorksheetName(worksheetString As String)
    margeWorksheetName = worksheetString
End Property

'マージ対象のシート名
Public Property Let setDataWorksheetName(worksheetString As String)
    dataWorksheetName = worksheetString
End Property

'データワークブック名を挿入する列
Public Property Let setWorkbookNameColumn(ColumnString As String)
    '列(英字)を列(数値)に変更 / get row number from row alphabet
    workbookNameColumn = Columns(ColumnString).Column
End Property

'連番付与列
Public Property Let setSerialNumbersColumn(ColumnString As String)
    '列(英字)を列(数値)に変更 / get row number from row alphabet
    serialNumbersColumn = Columns(ColumnString).Column
End Property

'--------getDataRowsクラスのプロパティ--------
Public Property Let setSearchFirstRow(rowNumber As Long)
    GDR.setSearchFirstRow = rowNumber
End Property

Public Property Let setSearchLastRow(rowNumber As Long)
    GDR.setSearchLastRow = rowNumber
End Property

Public Property Let setSearchColumn(ColumnString As String)
    GDR.setSearchColumn = ColumnString
End Property


'--------getDataRowsクラスのプロパティ終わり--------



'--------getFilenamesクラスのプロパティ--------
'検索ディレクトリを設定する
Public Property Let setDirectory(directoryString As String)
    GFN.setDirectory = directoryString
End Property

'パターンを設定する
Public Property Let setPattern(patternString As String)
    GFN.setPattern = patternString
End Property

'--------getFilenamesクラスのプロパティ終わり--------


Private Sub Class_Initialize()
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set GDR = New getDataRows
    Set GFN = New getFilenames

    'ワークブック名を挿入する列は0(挿入しない)
    workbookNameColumn = 0
End Sub

Private Sub Class_Terminate()
    Set FSO = Nothing
    Set GDR = Nothing
    Set GFN = Nothing
End Sub

Public Sub printSettings()
    'Debug.Print ":" & vbTab & ""
    Debug.Print "templateWorkbookName:" & vbTab & templateWorkbookName
    Debug.Print "templateWorksheetName:" & vbTab & templateWorksheetName
    Debug.Print "margeWorksheetName:" & vbTab & margeWorksheetName
    Debug.Print "dataWorksheetName:" & vbTab & dataWorksheetName
    Debug.Print "workbookNameColumn:" & vbTab & Split(Columns(workbookNameColumn).Address, "$")(2)
    Debug.Print "serialsNumberColumn:" & vbTab & Split(Columns(serialNumbersColumn).Address, "$")(2)
    
    Debug.Print "SearchFirstRow:" & vbTab & GDR.getSearchFirstRow
    Debug.Print "SearchLastRow:" & vbTab & GDR.getSearchLastRow
    Debug.Print "SearchColumn:" & vbTab & GDR.getSearchColumns
    
    Debug.Print "SearchDirectory:" & vbTab & GFN.getSearchDirectory
    Debug.Print "SearchPattern:" & vbTab & GFN.getSearchPattern

End Sub

Public Sub margeData()
    Dim dataWorkbook As Workbook
    Dim dataRows As Range
    Dim dataFilenames() As String
    Dim dataFilename As Variant
    Dim dataNextRowNumber As Long
    Dim dataCurrentEndRowNumber As Long

    '■■■■■ログ出力■■■■■
    Debug.Print "プログラム開始時間:" & Format(Now(), "yyyymmdd-hhmmss")
    '■■■■■ログ出力終了■■■■■


    '■■■■■初期確認■■■■■
    Call checkBeforeExec
    '■■■■■初期確認終了■■■■■


    '■■■■■1.データファイル名を取得■■■■■
    dataFilenames = GFN.getFilenames
    
    'ファイルがないとき
    If (Not dataFilenames) = -1 Then
        Err.Raise 1091, , "対象のデータファイルがありません"
    End If
    '■■■■■1.データファイル名を取得終了■■■■■
    
    
    '■■■■■2.ひな形シートをコピー■■■■■
    Call makeMargeSheet
    '■■■■■2.ひな形シートをコピー終了■■■■■

    
    '■■■■■3.データファイルを開いてデータを取得する■■■■■
    'データ開始位置を取得
    dataNextRowNumber = GDR.getSearchFirstRow

    '各作業ファイルのマクロを実行しないようにEnableEventsをFalseに設定する
    Application.EnableEvents = False
    For Each dataFilename In dataFilenames
        
        'マージ対象データブックを取得
        Set dataWorkbook = Workbooks.Open(filename:=dataFilename)
        'マージ対象データシートを取得
        Set GDR.setSearchWorksheet = dataWorkbook.Worksheets(dataWorksheetName)
        
        Set dataRows = GDR.getDataRows
        
        'データが存在しない場合はログ出力だけする
        If dataRows Is Nothing Then
            Debug.Print "データなしファイル名:" & dataFilename
        
        'データが存在する場合はデータをマージシートに貼り付ける
        Else
            dataCurrentEndRowNumber = dataNextRowNumber + dataRows.Rows.Count - 1
    
            'データを値貼り付けする
            margeWorksheet.Rows(dataNextRowNumber & ":" & dataCurrentEndRowNumber).Value = dataRows.Value
            
            'ファイル名挿入列の指定がある場合は挿入する
            If workbookNameColumn <> 0 Then
                margeWorksheet.Range(margeWorksheet.Cells(dataNextRowNumber, workbookNameColumn), margeWorksheet.Cells(dataCurrentEndRowNumber, workbookNameColumn)).Value _
                    = FSO.GetFileName(dataFilename)
            End If
    
            '次の貼り付け開始行番号を取得する
            dataNextRowNumber = dataCurrentEndRowNumber + 1
        
            'ログ出力
            Debug.Print "マージ対象ファイル名,行数:" & dataFilename & "," & dataRows.Rows.Count
        
        End If
        
        dataWorkbook.Close SaveChanges:=False
    
    Next
    Application.EnableEvents = True
    
    If dataNextRowNumber = GDR.getSearchFirstRow Then
        If MsgBox("マージ対象のデータがありませんでした。マージ用シートを削除しますか", vbYesNo) = vbYes Then
            Application.DisplayAlerts = False
            margeWorksheet.Delete
            Application.DisplayAlerts = True
        End If
    End If
    '■■■■■3.データファイルを開いてデータを取得する終了■■■■■


    '■■■■■4.データ付与■■■■■
    'シリアル番号付与指定がある場合
    If serialNumbersColumn > 0 Then
        margeWorksheet.Cells(GDR.getSearchFirstRow, serialNumbersColumn) = 1
        margeWorksheet.Cells(GDR.getSearchFirstRow, serialNumbersColumn).AutoFill _
            Destination:=margeWorksheet.Range(margeWorksheet.Cells(GDR.getSearchFirstRow, serialNumbersColumn), _
                                              margeWorksheet.Cells(dataNextRowNumber - 1, serialNumbersColumn)), _
            Type:=xlFillSeries
    End If
    '■■■■■4.データ付与■■■■■


    '■■■■■ログ出力■■■■■
    Debug.Print "プログラム終了時間:" & Format(Now(), "yyyymmdd-hhmmss")
    '■■■■■ログ出力終了■■■■■
    
End Sub

Private Sub checkBeforeExec()
    Dim errString As String
    Dim searchColumns As Variant
    Dim searchColumn As Variant
    Dim searchColumnNumber As Long

    If dataWorksheetName = "" Then
        errString = errString & "データシート名が設定されていません" & vbCrLf
    End If
    
    'マージ先シート名が指定されていない場合は、データシート名と同じにする
    If margeWorksheetName = "" Then
        margeWorksheetName = dataWorksheetName
    End If
    
    'テンプレートシート名が指定されていない場合は、データシート名と同じにする
    If templateWorksheetName = "" Then
        templateWorksheetName = dataWorksheetName
    End If
    
    'テンプレートブック名が指定されていない場合は、マクロがあるブックと仮定する
    If templateWorkbookName = "" Then
        templateWorkbookName = ThisWorkbook.Name
    End If


    If GFN.getSearchDirectory = "" Then
        errString = errString & "マージ対象ファイル保存フォルダが指定されていません" & vbCrLf
    End If
    
    If GFN.getSearchPattern = "" Then
        errString = errString & "マージ対象ファイルパターンが指定されていません" & vbCrLf
    End If
    
    
    '列の確認
    If GDR.getSearchColumns = "" Then
        errString = errString & "キーとなる列が指定されていません" & vbCrLf

    Else
        If Not (workbookNameColumn = 0 And serialNumbersColumn = 0) Then
            If workbookNameColumn = serialNumbersColumn Then
                errString = errString & "データワークブック名と連番を付与する列が同じです" & vbCrLf
            Else
                searchColumns = Split(GDR.getSearchColumns, vbCrLf)
                
'                '最後に空白行があるので配列最後の要素を削除する→GDR.getSearchColumnsで最後の改行を削除したのでここでは削除しない
'                'delete the last array index as it is blank
'                ReDim Preserve searchColumns(UBound(searchColumns) - 1)
            
                For Each searchColumn In searchColumns
                    
                    '列(英字)を列(数値)に変更 / get row number from row alphabet
                    searchColumnNumber = Columns(searchColumn).Column
                    
                    If workbookNameColumn = searchColumnNumber Or serialNumbersColumn = searchColumnNumber Then
                        errString = errString & "データ挿入列がキーとなる列と同じです" & vbCrLf
                        Exit For
                    End If
                Next
            End If
        End If
    End If

    'エラーがあった場合はエラーを投げる
    If errString <> "" Then
        Err.Raise 1099, , "パラメータの指定が間違っています。" & vbCrLf & vbCrLf & errString
    End If
End Sub

Private Sub makeMargeSheet()
    Workbooks(templateWorkbookName).Worksheets(templateWorksheetName).Copy After:=ThisWorkbook.Worksheets(Worksheets.Count)
    Set margeWorksheet = ActiveSheet
    margeWorksheet.Name = dataWorksheetName & Format(Now(), "yyyymmdd-hhmmss")
End Sub
