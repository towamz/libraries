Option Explicit

'mainで下記のクラスファイルを読み込む
'Include("clsPadding.vbs")	'クラスファイルの読み込み

Class clsPaddingNumber

Private paddingDigit
Private paddingString

'桁数
Public Property Let setPaddingDigit(argNumber)
	paddingDigit=argNumber
End Property

'文字
Public Property Let setPaddingString(argString)
	paddingString=argString
End Property

Private Sub Class_Initialize()
    paddingDigit = 2
    paddingString = "0"
End Sub

Private Sub Class_Terminate()

End Sub

Public Function getPaddingNumber(number)
	dim paddingStrings
	dim paddingCnt

	'数字以外が入力されたらそのまま返す
	if not IsNumeric(number) then
		getPaddingNumber=number
		Exit Function
	End if

	'指定の桁数より大きいときはそのまま返す
	if Len(number)>paddingDigit then
		getPaddingNumber=number
		Exit Function
	End if

	for paddingCnt = 0 to paddingDigit
		paddingStrings = paddingStrings & paddingString
	next

	getPaddingNumber = right(paddingStrings & number,paddingDigit)

End Function

End Class
