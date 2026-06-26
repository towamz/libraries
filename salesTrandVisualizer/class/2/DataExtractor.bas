'データ抽出用Class
'日付:ファイル名(yyyymmdd.csv)
'データ:1行データ:

Option Explicit

Private GFN As getFilenames
Private DateDataMap_ As Object
Private TargetDirectory_ As String


Private Sub Class_Initialize()
    Set GFN = New getFilenames
    Set DateDataMap_ = CreateObject("Scripting.Dictionary")
    
    GFN.TargetDirectory = "C:\sampleMacro\salesTrendVisualizer\csv"
    GFN.Pattern = "\.csv"
End Sub

Public Function getDateDataMap() As Object
    Dim fn As Object
    Dim k As Variant
    Dim tDateString As String
    Dim tDate As Date
    Dim tData As Variant


    Set fn = GFN.FilenamesDictionary
    
    Application.ScreenUpdating = False
    For Each k In fn.Keys
        On Error Resume Next
        tDate = DateSerial(Left(k, 4), Mid(k, 5, 2), Mid(k, 7, 2))
        On Error GoTo 0
        tDateString = Left(k, 8)
        
        '当サンプルではyyyymmdd.csvを想定しているので
        'len(k)=12(ファイル名の長さ12文字)を条件に入れている
        If Format(tDate, "yyyymmdd") = tDateString And Len(k) = 12 Then
        
            tData = getArrayFromCSV(fn(k))
            
            '■■■■■取得したデータの正常性確認■■■■■
            Dim rc As Long: rc = UBound(tData, 1) - LBound(tData, 1) + 1
            Dim cc As Long: cc = UBound(tData, 2) - LBound(tData, 2) + 1
            
            If rc <> 1 Then
                Err.Raise 9999, "", "データが不正です" & vbCrLf & "ファイル名:" & k & vbCrLf & "行数:" & rc & vbCrLf & "列数:" & cc
            End If
            '■■■■■取得したデータの正常性確認終わり■■■■■
            
            '■■■■■取得したデータの要素数調整■■■■■
            '配列要素数を調整するにはReDim Preserveを使う
            'ReDim Preserve myArray(0 To 0, 0 To 1)
            '■■■■■取得したデータの要素数調整終わり■■■■■
        
            
            '■■■■■データをディクショナリに追加■■■■■
            '重複データがあった場合に上書きか例外を投げるかは状況により判断する
            If DateDataMap_.Exists(tDate) Then
                Err.Raise 9999, "", "データが重複しています"
            End If
        
            DateDataMap_.Add tDateString, tData
            '■■■■■データをディクショナリに追加終わり■■■■■
        Else
            Err.Raise 9999, "", "データが不正です" & vbCrLf & "ファイル名:" & k
        End If
    
    Next k
    Application.ScreenUpdating = True

    Set getDateDataMap = DateDataMap_

End Function

Private Function getArrayFromCSV(targetWorkbook As String) As Variant
    Dim wb As Workbook
    Dim srcData As Variant
    
    On Error GoTo errHandler
    Workbooks.OpenText Filename:=targetWorkbook, DataType:=xlDelimited, Comma:=True
      Set wb = Workbooks(Workbooks.Count)
    If StrComp(wb.Name, Dir$(targetWorkbook), vbTextCompare) <> 0 Then
        Err.Raise 9999, "", "ファイル名が一致していません"
    End If
    
    srcData = wb.Sheets(1).UsedRange.Value
    
    'データがない場合はempty値のまま
    If IsEmpty(srcData) Then
        '何もしない
    'A1のみデータがある場合は2次元配列に変換する
    ElseIf Not IsArray(srcData) Then
        Dim tmpAry() As Variant: ReDim tmpAry(1 To 1, 1 To 1)
        tmpAry(1, 1) = srcData
        srcData = tmpAry
    End If
    
    getArrayFromCSV = srcData
    GoTo finally

errHandler:
    MsgBox Err.Number & ":" & Err.Description, vbOKOnly + vbCritical
    
finally:
    On Error GoTo 0
    If Not wb Is Nothing Then
        wb.Close savechanges:=False
    End If

End Function
