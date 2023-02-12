Option Explicit

'定数宣言（設定保存セル）
Private Const CONTROL_SHEET_NAME As String = "設定"
Private Const FORMAT_SHEET_NAME_RANGE_STR As String = "B2"
Private Const DATA_SHEET_NAME_RANGE_STR As String = "B3"
Private Const DATA_START_RANGE_STR As String = "B4"
Private Const MAX_BLANK_CNT_RANGE_STR As String = "B5"
Private Const TARGET_DATE_RANGE_STR As String = "B6"
Private Const SEARCH_DIRECTORYE_RANGE_STR As String = "B7"



Public Function readFileByDay() As Boolean
    Dim objDataFile As Workbook
    Dim objControlFile As Workbook
    
    Dim varFileNames() As Variant
    Dim currentFileName As Variant
    
    Dim dataRangeCnt As Integer
    Dim blankCnt As Integer
    Dim nextRow As Long
    
    Dim strFormatSheetName As String
    Dim strDataSheetName As String
    Dim strDataStartRange As String
    Dim strTargetDirectory As String
    
    
    '----------各パラメータを取得----------
    'とりまとめブックオブジェクト
    Set objControlFile = ThisWorkbook
    'ひな形名取得
    strFormatSheetName = objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(FORMAT_SHEET_NAME_RANGE_STR).Value
    'データ・とりまとめシート名取得
    strDataSheetName = objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(DATA_SHEET_NAME_RANGE_STR).Value
    '検索対象ディレクトリ
    strTargetDirectory = objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(SEARCH_DIRECTORYE_RANGE_STR).Value 'データ保存ディレクトリを取得
    'データ開始セル
    strDataStartRange = objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(DATA_START_RANGE_STR).Value).Value
    
    
    
    '----------対象ファイル名を取得する----------
    varFileNames = getFileNames
    
    'ファイル名が取得できない時はFalse(配列でない)が帰ってくる
    If Not IsArray(varFileNames) Then
        readFileByDay = False
        Exit Function
    End If
    
    
    If MsgBox("下記のファイルを処理します。実行しますか?" & vbCrLf & Join(varFileNames, vbCrLf), vbOKCancel) = vbCancel Then
        readFileByDay = False
        Exit Function
    End If
    
    
    '----------ひな形シートをコピーする----------
    On Error Resume Next
    objControlFile.Worksheets(strDataSheetName).Activate

    If Err.Number = 0 Then
        If MsgBox(strDataSheetName & "は存在します。シートを削除し処理を続行しますか", vbYesNo) = vbYes Then
            Application.DisplayAlerts = False
            objControlFile.Worksheets(strDataSheetName).Delete
            Application.DisplayAlerts = True
        Else
            readFileByDay = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    objControlFile.Worksheets(strFormatSheetName).Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = strDataSheetName
    
    
    
    
    '----------各作業ファイルから作業結果をコピーする----------
    'とりまとめ用のデータ開始行を取得する
    nextRow = objControlFile.Worksheets(strDataSheetName).Range(strDataStartRange).Row
    
    '各作業ファイルのマクロを実行しないようにEnableEventsをFalseに設定する
    Application.EnableEvents = False
    For Each currentFileName In varFileNames
        '作業ファイル用のカウンタ変数を初期化
        dataRangeCnt = 0
        blankCnt = 0
        
        '作業ファイルを取得
        Set objDataFile = Workbooks.Open(FileName:=strTargetDirectory & currentFileName)
        

        objDataFile.Worksheets(strDataSheetName).Select
        objDataFile.Worksheets(strDataSheetName).Range(strDataStartRange).Activate
        
        'データ範囲を取得(xlEndだと正しく取得できない場合があるため)、blankCntで連続空白数を指定
        Do Until blankCnt > objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(MAX_BLANK_CNT_RANGE_STR).Value
            If objDataFile.Worksheets(strDataSheetName).Range(strDataStartRange).Offset(dataRangeCnt + blankCnt, 0) = "" Then
                blankCnt = blankCnt + 1
            Else
                dataRangeCnt = dataRangeCnt + blankCnt + 1
                blankCnt = 0
            End If
        Loop
        
        'データが1件もなかった場合はコピーしないでそのまま閉じる
        If dataRangeCnt <> 0 Then
            '取得したデータ範囲をコピー
            objDataFile.Worksheets(strDataSheetName).Rows(Worksheets(strDataSheetName).Range(strDataStartRange).Row & ":" & objDataFile.Worksheets(strDataSheetName).Range(strDataStartRange).Offset(dataRangeCnt - 1, 0).Row).Copy
            'MsgBox "コピー範囲を確認"
            
            'コピーしたデータをとりまとめシートに貼り付け(各作業者が書式を変えている可能性があるので、値貼り付けで書式はコピーしない)
            objControlFile.Worksheets(strDataSheetName).Rows(nextRow).PasteSpecial Paste:=xlPasteValues
        
            'とりまとめシートの次の貼り付け位置を保存
            nextRow = nextRow + dataRangeCnt
        End If
        
        '各作業ファイルを閉じる
        Application.DisplayAlerts = False
        objDataFile.Close False
        Application.DisplayAlerts = True
        
        objControlFile.Worksheets(strDataSheetName).Rows(nextRow).Activate
        'MsgBox "次の貼り付け先を確認"
    Next
    Application.EnableEvents = True

    'フォーマット書式をコピー（各作業者が書式を変えている可能性があるので、ひな形の書式をコピーする）
    objControlFile.Worksheets(strFormatSheetName).Rows(objControlFile.Worksheets(strFormatSheetName).Range(strDataStartRange).Row).Copy
    objControlFile.Worksheets(strDataSheetName).Rows(objControlFile.Worksheets(strDataSheetName).Range(strDataStartRange).Row & ":" & nextRow - 1).PasteSpecial Paste:=xlPasteFormats

    '正常に取り込めたのでtrueを返す
    readFileByDay = True


End Function


Private Function getFileNames() As Variant
    Dim objControlFile As Workbook
    Dim strTargetDirectory As String
    Dim strTargetData As String
    
    Dim aryFileNames() As Variant
    Dim FileName As String
    Dim cnt As Integer
    
    
    '----------各パラメータを取得----------
    'とりまとめブックオブジェクト
    Set objControlFile = ThisWorkbook
    
    strTargetDirectory = objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(SEARCH_DIRECTORYE_RANGE_STR).Value 'データ保存ディレクトリを取得
    strTargetData = objControlFile.Worksheets(CONTROL_SHEET_NAME).Range(TARGET_DATE_RANGE_STR).Value
    
    cnt = 0

    '指定フォルダの[yymm*.xlsx]形式のファイルを処理対象にする
    Debug.Print strTargetDirectory & Format(strTargetData, "mmdd") & "*.xlsx"
    FileName = Dir(strTargetDirectory & Format(strTargetData, "mmdd") & "*.xlsx", vbNormal)
    
    Do While FileName <> ""
        ReDim Preserve aryFileNames(cnt)
        
        aryFileNames(cnt) = FileName
        
        FileName = Dir()
        cnt = cnt + 1
    Loop

    getFileNames = aryFileNames

End Function

