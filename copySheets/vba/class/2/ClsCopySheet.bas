Option Explicit

Private OrigGFAS As ClsGetFileAndSheet

Public OrigSheetNameNew As String
Private OrigSheetCheckDictionary_ As Object

'デフォルトディレクトリを設定する(任意)
Public Property Let DefaultDirectory(arg1 As String)
    OrigGFAS.DefaultDirectory = arg1
End Property

Public Property Get DefaultDirectory() As String
    DefaultDirectory = OrigGFAS.DefaultDirectory
End Property


'コピー対象のブック名を設定する(任意) / 選択したブック名とここで設定したブック名が一致しないとエラーになる
Public Property Let OrigFileName(arg1 As String)
    OrigGFAS.FileName = arg1
End Property

Public Property Get OrigFileName() As String
    OrigFileName = OrigGFAS.FileName
End Property


'コピー対象のブック名(フルパス)を設定する(任意) / 設定しないとファイルを開くダイアログを表示
Public Property Let OrigAbsoluteFileName(arg1 As String)
    OrigGFAS.AbsoluteFileName = arg1
End Property

Public Property Get OrigAbsoluteFileName() As String
    OrigAbsoluteFileName = OrigGFAS.AbsoluteFileName
End Property


'コピー対象のシート名を設定する
Public Property Let OrigSheetName(arg1 As String)
    OrigGFAS.SheetName = arg1
End Property

Public Property Get OrigSheetName() As String
    OrigSheetName = OrigGFAS.SheetName
End Property


'ファイル選択ダイアログのタイトルを設定する
Public Property Let OrigDialogTitle(arg1 As String)
    OrigGFAS.DialogTitle = arg1
End Property

Public Property Get OrigDialogTitle() As String
    OrigDialogTitle = OrigGFAS.DialogTitle
End Property


'コピー対象のファイル名フィルタを設定する
Public Property Let OrigDialogFileFilter(arg1 As String)
    OrigGFAS.DialogFileFilter = arg1
End Property

Public Property Get OrigDialogFileFilter() As String
    OrigDialogFileFilter = OrigGFAS.DialogFileFilter
End Property


Private Sub Class_Initialize()
    'データチェック用ディクショナリの初期化
    Set OrigSheetCheckDictionary_ = CreateObject("Scripting.Dictionary")
    Set OrigGFAS = New ClsGetFileAndSheet
End Sub

Private Sub Class_Terminate()
    Set OrigGFAS = Nothing
End Sub


'データシートの列名チェックを設定する(任意)
Public Sub setOrigSheetCheckKey(arg1 As String, arg2 As String)
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

    OrigSheetCheckDictionary_.Add arg1, arg2
End Sub

'ファイル1個用 / for one file
Public Sub copySheet()
    Dim wbDest As Workbook
    Dim wsOrig As Worksheet
    Dim dicKey As Variant
    
    Set wbDest = ThisWorkbook
    Set wsOrig = OrigGFAS.getSheet
    
    'ファイル内容チェック/ check data contents
    For Each dicKey In OrigSheetCheckDictionary_
        If wsOrig.Range(dicKey).Value <> OrigSheetCheckDictionary_.Item(dicKey) Then
            Err.Raise 2001, , "ファイルの内容が間違っています"
            Exit Sub
        End If
    Next

    'シート名の指定がある場合は変更する(コピー先で変更するとシート名が重複しているとエラーとなるため、コピー元で変更する)
    If OrigSheetNameNew <> "" Then
        wsOrig.Name = OrigSheetNameNew
    End If
    
    '末尾にシートをコピーする/ copy the sheet to the end of the book
    wsOrig.Copy after:=wbDest.Sheets(wbDest.Sheets.Count)

End Sub


