' グローバル定数の定義
Const OFT_FILE_PATH As String = "C:\Path\To\Your\File.oft" ' OFTファイルのパス
Const SHEET_NAME As String = "Sheet1" ' Excelのワークシート名（適宜変更）

Sub createEmailsFromOFT()
    Dim ExcelWB As Workbook
    Dim ExcelWS As Worksheet

    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim MailSubject As String
    Dim MailBody As String

    Dim i As Long
    Dim j As Long

    ' Outlookアプリケーションを取得
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    If OutlookApp Is Nothing Then
        MsgBox "outlookを正常に操作できません", vbOk + vbCritical
    End If
    ' Excelの現在のブックとシートを取得
    Set ExcelWB = ThisWorkbook
    Set ExcelWS = ExcelWB.Sheets(SHEET_NAME) ' グローバル定数からシート名を取得

    ' 行ループ（paramSub1が空白かつparam1が空白になるまで）
    i = 2 ' データは2行目から開始
    Do Until ExcelWS.Cells(i, 4).Value = "" And ExcelWS.Cells(i, 9).Value = "" ' paramSub1が空白かつparam1が空白
        ' OFTファイルをベースにメール作成
        Set MailItem = OutlookApp.CreateItemFromTemplate(OFT_FILE_PATH) ' グローバル定数からOFTファイルのパスを取得

        ' 本文と件名の取得
        MailSubject = MailItem.Subject
        MailBody = MailItem.Body

        ' メールタイトル用プレースホルダ {paramSub1}～{paramSub5} の置き換え
        For j = 1 To 5
            If ExcelWS.Cells(i, 3 + j).Value = "" Then
                Exit For
            End If
            MailSubject = Replace(MailSubject, "{paramSub" & j & "}", ExcelWS.Cells(i, 3 + j).Value)
        Next j

        ' メール本文用プレースホルダ {param1}～{param10} の置き換え
        For j = 1 To 10
            If ExcelWS.Cells(i, 8 + j).Value = "" Then
                Exit For
            End If
            MailBody = Replace(MailBody, "{param" & j & "}", ExcelWS.Cells(i, 8 + j).Value)
        Next j

        ' メールに情報を設定（アドレスを直接Excelから取得）
        MailItem.To = ExcelWS.Cells(i, 1).Value
        MailItem.CC = ExcelWS.Cells(i, 2).Value
        MailItem.BCC = ExcelWS.Cells(i, 3).Value
        MailItem.Subject = MailSubject
        MailItem.Body = MailBody

        ' メールを保存
        MailItem.Save ' 下書きフォルダに保存

        i = i + 1
    Loop

    ' 終了処理
    MsgBox "メール作成完了（下書きに保存済み）", vbInformation
    Set MailItem = Nothing
    Set OutlookApp = Nothing
    Set ExcelWS = Nothing
    Set ExcelWB = Nothing
End Sub