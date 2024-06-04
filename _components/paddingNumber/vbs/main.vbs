Option Explicit

'カレントディレクトリをこのファイルがあるディレクトリに変更
'change current directory to where this file is
Dim FSO, SHL

Set FSO = Wscript.createObject("Scripting.FileSystemObject")
Set SHL = Wscript.CreateObject("WScript.shell")

SHL.CurrentDirectory = FSO.getParentFolderName(WScript.ScriptFullName)

'外部ファイル読み込み参照サイト
'http://cloiwan.com/?p=272
Function Include(strFile)
	Dim OTF

	Set OTF = FSO.OpenTextFile(strFile)

	ExecuteGlobal OTF.ReadAll()
	OTF.Close

	Set OTF = Nothing
End Function


'mainで下記のクラスファイルを読み込む
Include("clsPaddingNumber.vbs")	'クラスファイルの読み込み


Dim PAD 
Set PAD = New clsPaddingNumber

PAD.setPaddingDigit=5
PAD.setPaddingString="-"

MsgBox PAD.getPaddingNumber(11)
