Private ENM_HandleBy As handleBy

'時間によるハンドラー用変数 / variables for hadle by time
Private DAT_NextDuetime As Date
Private LNG_TimeIntervalSeconds As Long

'回数によるハンドラー用変数 / variables for hadle by cnt
Private LNG_HandlerCallCnt As Long
Private LNG_AbortionPer As Long


Enum handleBy
    byTime
    byCnt
End Enum


Private Sub Class_Initialize()
    '時間でハンドルが初期設定
    ENM_HandleBy = byTime
    
    '時間によるハンドラーの初期設定
    LNG_TimeIntervalSeconds = 300
    Call setNextDueTime
    
    '回数によるハンドラーの初期設定
    LNG_AbortionPer = 5
    LNG_HandlerCallCnt = 0

    Debug.Print "-----処理開始-----"
    Call printLog(0)
End Sub

Private Sub Class_Terminate()
    Debug.Print "-----処理終了-----"
    Call printLog(1)
End Sub


Property Let setTimeInterval(argTimeInterval As Long)
    If argTimeInterval < 60 Then
        err.Raise 1, , "60秒以上を設定してください"
    End If
    
    LNG_TimeIntervalSeconds = argTimeInterval
    Call setNextDueTime
    Debug.Print "-----設定変更-----"
    Call printLog(0)
End Property

Property Let setAbortionPer(argAbortionPer As Long)
    If argAbortionPer < 5 Then
        err.Raise 2, , "5回以上を設定してください"
    End If
    
    LNG_AbortionPer = argAbortionPer
    
    Debug.Print "-----設定変更-----"
    Call printLog(0)
End Property

Property Let setHandleBy(argHandleBy As handleBy)
    ENM_HandleBy = argHandleBy
    
    Debug.Print "-----設定変更-----"
    Call printLog(0)
End Property


Public Sub infiniteLoopHandler()
    DoEvents
    'カウントアップ
    LNG_HandlerCallCnt = LNG_HandlerCallCnt + 1
    
    Select Case ENM_HandleBy
        Case handleBy.byTime
            Call printLog(2)
            If DAT_NextDuetime < Now() Then
                If MsgBox("処理に長時間かかっています。処理を継続しますか", vbOKCancel) = vbCancel Then
                    Debug.Print "-----処理中断-----"
                    Call printLog(1)
                    err.Raise 99, , "ユーザーによる中断"
                Else
                    setNextDueTime
                    Debug.Print "-----処理続行-----"
                    Call printLog(0)
                End If
            End If
    
        Case handleBy.byCnt
            Call printLog(11)
            If LNG_HandlerCallCnt Mod LNG_AbortionPer = 0 Then
                If MsgBox("処理が指定の回数に達しました。処理を継続しますか" & vbCrLf & "指定回数" & LNG_AbortionPer, vbOKCancel) = vbCancel Then
                    Debug.Print "-----処理中断-----"
                    Call printLog(11)
                    err.Raise 99, , "ユーザーによる中断"
                Else
                    Debug.Print "-----処理続行-----"
                    Call printLog(11)
                End If
            End If
    
    End Select
End Sub


Private Sub printLog(argLogType As Long)
    Select Case argLogType
        Case 0
            Debug.Print LNG_HandlerCallCnt, Now(), DAT_NextDuetime, TimeSerial(0, 0, LNG_TimeIntervalSeconds), "AbortionPer:" & LNG_AbortionPer, "handleBy:" & ENM_HandleBy
        Case 1
            Debug.Print LNG_HandlerCallCnt, Now()
        Case 2
            Debug.Print LNG_HandlerCallCnt, Now(), DAT_NextDuetime
        Case 11
            Debug.Print LNG_HandlerCallCnt, Now(), LNG_AbortionPer
        Case Else
            Debug.Print "unknown logtype specified"
            err.Raise 11, , "ログの出力指定が間違っています"
    End Select
End Sub


Private Sub setNextDueTime()
    DAT_NextDuetime = Now() + TimeSerial(0, 0, LNG_TimeIntervalSeconds)
End Sub


'err.Raise 1, , "60秒以上を設定してください"
'err.Raise 2, , "5回以上を設定してください"
'err.Raise 11, , "ログの出力指定が間違っています"
'err.Raise 99, , "ユーザーによる中断"
