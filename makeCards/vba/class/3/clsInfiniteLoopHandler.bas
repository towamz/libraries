Private DAT_Next_Due_Time As Date
Private LNG_HandlerCallCnt As Long

'定数(変更なし)
Const LNG_TIMEINTERVAL_HOURS As Long = 0
Const LNG_TIMEINTERVAL_MINUTES As Long = 5
Const LNG_TIMEINTERVAL_SECONDS As Long = 0

Private Sub Class_Initialize()
    'ハンドラーを呼ばれた回数を保存
    LNG_HandlerCallCnt = 0
    
    '処理を止める時刻を取得
    DAT_Next_Due_Time = getNextDueTime()

    Debug.Print "-----処理開始-----"
    print2

End Sub


Public Sub infiniteLoopHandler()
    'カウントアップ
    LNG_HandlerCallCnt = LNG_HandlerCallCnt + 1
    print3
    
    If DAT_Next_Due_Time < Now() Then
        If MsgBox("処理に長時間かかっています。処理を継続しますか", vbOKCancel) = vbCancel Then
            Debug.Print "-----処理中断-----"
            print1
            err.Raise 1000, , "ユーザーによる中断"
        Else
            DAT_Next_Due_Time = getNextDueTime()
            Debug.Print "-----処理続行-----"
            print2
        End If
    End If

End Sub


Private Sub Class_Terminate()
    Debug.Print "-----処理終了-----"
    print1
End Sub


Private Sub print1()
    Debug.Print LNG_HandlerCallCnt, Now()
End Sub


Private Sub print2()
    Debug.Print LNG_HandlerCallCnt, Now(), DAT_Next_Due_Time, TimeSerial(LNG_TIMEINTERVAL_HOURS, LNG_TIMEINTERVAL_MINUTES, LNG_TIMEINTERVAL_SECONDS)
End Sub


Private Sub print3()
    Debug.Print LNG_HandlerCallCnt, Now(), DAT_Next_Due_Time
End Sub


Private Function getNextDueTime()

    getNextDueTime = Now() + TimeSerial(LNG_TIMEINTERVAL_HOURS, LNG_TIMEINTERVAL_MINUTES, LNG_TIMEINTERVAL_SECONDS)

End Function

