Sub ExportOutlookEmailsToExcel()

    Dim OutlookApp As Object
    Dim Namespace As Object
    Dim FolderReceived As Object
    Dim FolderSent As Object
    Dim Items As Object
    Dim Mail As Object
    Dim i As Long
    Dim ws As Worksheet
    Dim dateFrom As Date
    Dim dateUntil As Date
    Dim filter As String

    ' 日付範囲の設定
    dateFrom = "2025/1/20"
    dateUntil = "2025/1/25"

    ' Excelのワークシートを設定
    Set ws = ThisWorkbook.Sheets("メール")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "key"
    ws.Cells(1, 2).Value = "タイムスタンプ"
    ws.Cells(1, 3).Value = "from"
    ws.Cells(1, 4).Value = "to"
    ws.Cells(1, 5).Value = "cc"
    ws.Cells(1, 6).Value = "bcc"
    ws.Cells(1, 7).Value = "件名"
    ws.Cells(1, 8).Value = "本文"

    ' Outlookアプリケーションを取得
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    ' Namespaceを取得
    Set Namespace = OutlookApp.GetNamespace("MAPI")
    
    ' 受信フォルダの「対応」フォルダを取得
    Set FolderReceived = Namespace.GetDefaultFolder(olFolderInbox).Folders("対応")
    ' 送信済トレイフォルダを取得
    Set FolderSent = Namespace.GetDefaultFolder(olFolderSentMail)

    i = 2 ' データの開始行
    
    ' 日付範囲のフィルタ作成
    filter = "[ReceivedTime] >= '" & Format(dateFrom, "yyyy/mm/dd hh:nn AMPM") & "' AND [ReceivedTime] <= '" & Format(dateUntil, "yyyy/mm/dd hh:nn AMPM") & "'"

    ' 受信フォルダの「対応」フォルダのメールを取得しフィルタを適用
    Set Items = FolderReceived.Items.Restrict(filter)
    Items.Sort "[ReceivedTime]", True ' 受信日時でソート（降順）
    
    For Each Mail In Items
        If Mail.Class = olMail Then
            ws.Cells(i, 1).Value = Mail.SenderEmailAddress ' key (from)
            ws.Cells(i, 2).Value = Mail.ReceivedTime ' タイムスタンプ
            ws.Cells(i, 3).Value = Mail.SenderEmailAddress ' from
            ws.Cells(i, 4).Value = Mail.To ' to
            ws.Cells(i, 5).Value = Mail.CC ' cc
            ws.Cells(i, 6).Value = Mail.BCC ' bcc
            ws.Cells(i, 7).Value = Mail.Subject ' 件名
            ws.Cells(i, 8).Value = Mail.Body ' 本文
            i = i + 1
        End If
    Next Mail

    ' 送信済トレイフォルダのメールを取得しフィルタを適用
    filter = "[SentOn] >= '" & Format(dateFrom, "yyyy/mm/dd hh:nn AMPM") & "' AND [SentOn] <= '" & Format(dateUntil, "yyyy/mm/dd hh:nn AMPM") & "'"
    Set Items = FolderSent.Items.Restrict(filter)
    Items.Sort "[SentOn]", True ' 送信日時でソート（降順）
    
    For Each Mail In Items
        If Mail.Class = olMail Then
            ws.Cells(i, 1).Value = Mail.To ' key (to)
            ws.Cells(i, 2).Value = Mail.SentOn ' タイムスタンプ
            ws.Cells(i, 3).Value = Mail.SenderEmailAddress ' from
            ws.Cells(i, 4).Value = Mail.To ' to
            ws.Cells(i, 5).Value = Mail.CC ' cc
            ws.Cells(i, 6).Value = Mail.BCC ' bcc
            ws.Cells(i, 7).Value = Mail.Subject ' 件名
            ws.Cells(i, 8).Value = Mail.Body ' 本文
            i = i + 1
        End If
    Next Mail

    ' A列昇順、B列降順でソート
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(1), Order:=xlAscending
        .SortFields.Add Key:=ws.Columns(2), Order:=xlDescending
        .SetRange ws.UsedRange
        .Header = xlYes
        .Apply
    End With

    MsgBox "メールデータを取得しました。", vbInformation
End Sub