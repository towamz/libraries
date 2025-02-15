Option Explicit

Private sh As Worksheet
Private FolderTarget As Object

Private FolderName As String
Private SheetName As String
Private DateFrom As Date
Private DateUntil As Date

Private KeyType As EnumKeyType

Public Enum EnumKeyType
    MailFrom
    MailTo
    MailCC
    MailBCC
End Enum

Public Property Let setFolderName(string1 As String)
    FolderName = string1
End Property

Public Property Let setSheetName(string1 As String)
    SheetName = string1
End Property

Public Property Let setDateFrom(date1 As Date)
    DateFrom = date1
End Property

Public Property Let setDateUntil(date1 As Date)
    DateUntil = date1
End Property

Public Property Let setKeyType(KeyType1 As EnumKeyType)
    KeyType = KeyType1
End Property

Public Sub setTitles()
    sh.Cells.Clear
    sh.Cells(1, 1).Value = "key"
    sh.Cells(1, 2).Value = "タイムスタンプ"
    sh.Cells(1, 3).Value = "from"
    sh.Cells(1, 4).Value = "to"
    sh.Cells(1, 5).Value = "cc"
    sh.Cells(1, 6).Value = "bcc"
    sh.Cells(1, 7).Value = "件名"
    sh.Cells(1, 8).Value = "本文"
End Sub

Public Sub setFolderTarget()
    Dim OutlookApp As Object
    Dim Namespace As Object
    Dim olFolderInbox As Object

    ' Outlookアプリケーションを取得
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    ' Namespaceを取得
    Set Namespace = OutlookApp.GetNamespace("MAPI")
    
    ' 受信箱フォルダ（olFolderInbox）を取得
    Set olFolderInbox = Namespace.GetDefaultFolder(6)
    Set FolderTarget = olFolderInbox.Folders(FolderName)
End Sub

Public Sub inputMessages()
    Dim OutlookApp As Object
    Dim Namespace As Object
    Dim olFolderInbox As Object
    Dim Items As Object
    Dim Mail As Object
    Dim i As Long
    Dim ws As Worksheet
    Dim filter As String

    If SheetName = "" Then
        Err.Raise 1000, , "シート名が設定されていません"
    End If
    
    If FolderName = "" Then
        Err.Raise 1000, , "メールフォルダ名が設定されていません"
    End If

    If DateFrom = 0 Then
        DateFrom = Date
    End If
    
    If DateUntil = 0 Then
        DateUntil = DateAdd("d", 1, DateFrom)
    End If

    Set sh = Worksheets(SheetName)

    Call setTitles
    Call setFolderTarget

    i = 2 ' データの開始行

    ' 送信済トレイフォルダのメールを取得しフィルタを適用
    filter = "[SentOn] >= '" & Format(DateFrom, "yyyy/mm/dd hh:nn AMPM") & "' AND [SentOn] <= '" & Format(DateUntil, "yyyy/mm/dd hh:nn AMPM") & "'"
    Set Items = FolderTarget.Items.Restrict(filter)
    'Set Items = FolderSent.Items
    Items.Sort "[SentOn]", True ' 送信日時でソート（降順）
    
    For Each Mail In Items
            If KeyType = MailFrom Then
                sh.Cells(i, 1).Value = GetSenderEmailAddress(Mail) ' key (to)
            Else
                sh.Cells(i, 1).Value = GetRecipientsAddresses(Mail, CInt(KeyType)) ' key (to)
            End If
            sh.Cells(i, 2).Value = Mail.SentOn ' タイムスタンプ
            sh.Cells(i, 3).Value = GetSenderEmailAddress(Mail) ' from
            sh.Cells(i, 4).Value = GetRecipientsAddresses(Mail, 1) ' to
            sh.Cells(i, 5).Value = GetRecipientsAddresses(Mail, 2) ' cc
            sh.Cells(i, 6).Value = GetRecipientsAddresses(Mail, 3) ' bcc
            sh.Cells(i, 7).Value = Mail.Subject ' 件名
            'sh.Cells(i, 8).Value = Mail.Body ' 本文
            i = i + 1
    Next Mail

    MsgBox "メールデータを取得しました。", vbInformation
End Sub

' メールの受信者リストを取得するゲッター関数（To, CC, BCC共通）
Private Function GetRecipientsAddresses(Mail As Object, recipientType As Integer) As String
    Dim recipient As Object
    Dim recipientsList As String
    recipientsList = ""
    
    ' 受信者リストをループ
    For Each recipient In Mail.Recipients
        ' 受信者のタイプが一致する場合に処理
        If recipient.Type = recipientType Then
            ' Exchangeユーザーの場合、名前を追加、それ以外はメールアドレスを追加
            If Not recipient.AddressEntry Is Nothing Then
                If recipient.AddressEntry.Type = "EX" Then
                    recipientsList = recipientsList & recipient.Name & ";"
                Else
                    recipientsList = recipientsList & recipient.Address & ";"
                End If
            End If
        End If
    Next recipient
    
    ' 最後のセミコロンを削除（オプション）
    If Len(recipientsList) > 0 Then
        recipientsList = Left(recipientsList, Len(recipientsList) - 1)
    End If
    
    GetRecipientsAddresses = recipientsList
End Function

Private Function GetSenderEmailAddress(Mail As Object) As String
    Dim senderEmail As String
    
    ' Mail.Senderが存在するか確認
    If Not Mail.Sender Is Nothing Then
        If Mail.Sender.Type = "EX" Then
            GetSenderEmailAddress = Mail.Sender.Name
        Else
            GetSenderEmailAddress = Mail.Sender.Address
        End If

    End If
End Function
