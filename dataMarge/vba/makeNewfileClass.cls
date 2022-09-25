VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "makeNewfileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CONTROL_SHEET_NAME As String = "設定"

Private DATA_SHEET_NAME As String
Private SAVE_DIRECTORY As String
Private FILENAME As String

Private Sub Class_Initialize()

    DATA_SHEET_NAME = Worksheets(CONTROL_SHEET_NAME).Range("B3").Value
    
    '保存ディレクトリ、\で終わってない場合は付加する
    If Right(Worksheets(CONTROL_SHEET_NAME).Range("B7").Value, 1) <> "\" Then
        Worksheets(CONTROL_SHEET_NAME).Range("B7").Value = Worksheets(CONTROL_SHEET_NAME).Range("B7").Value & "\"
    End If

    SAVE_DIRECTORY = Worksheets(CONTROL_SHEET_NAME).Range("B7").Value
    
    FILENAME = Worksheets(CONTROL_SHEET_NAME).Range("B13").Value & Format(Now(), "mmdd") & ".xlsx"

End Sub





Public Function makeNewfile() As Boolean

    On Error Resume Next
    Sheets(DATA_SHEET_NAME).Select

    If Err.Number <> 0 Then
        MsgBox "マージが実行されていません", vbOKOnly + vbExclamation
        makeNewfile = False
        Exit Function
    End If


    'ひな形シートを新規ブックにコピー
    Sheets(DATA_SHEET_NAME).Copy
    
    '指定フォルダに保存する
    ActiveWorkbook.SaveAs FILENAME:=SAVE_DIRECTORY & FILENAME


End Function
