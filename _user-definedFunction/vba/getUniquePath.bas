'---------------------------------------------------------------------------------------
' Procedure : getUniquePath
' Purpose   : To generate a unique path with a serial suffix.
'             The suffix is minimum lack number
'
' Parameters:
'  - targetPath : The base file path.
'  - digit      : [Optional]The digit of serial number
'  - startNumber: [Optional][ByRef] The start suffix number
'
' Returns   : String - The unique file path.
' Error     : 1000 - The start number exceed the digit.
'             1001 - A unique number was not found within the specified digit.
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' 関数名：getUniquePath
' 目的　：重複しないファイルパスを、連番（サフィックス）を付与して生成する。
' 備考　：
'   1. フォルダ内に「虫食い（1.txt, 3.txt...）」がある場合、最小の空き番号を優先して返す。
'   2. 第3引数(startNumber)に変数を通すことで、前回見つかった番号が書き戻される。
'      これを次回呼び出し時に再利用（+1）することで、大量処理時の速度を劇的に向上させる。
'
' 引数　：
'   targetPath  - ベースとなるファイルパス
'   digit       - [Optional]連番の桁数（1?9桁。Long型の制限により最大9）
'   startNumber - [Optional][ByRef] 検索を開始する番号。
'
' 戻り値：String - 重複しない新しいファイルパス
' 例外    : 1000 - 開始番号が桁数制限を超えた場合。
'           1001 - 指定桁数内で空き番号が見つからなかった場合。
'---------------------------------------------------------------------------------------

Function getUniquePath(targetPath As String, Optional digit As Long = 2, Optional ByRef startNumber As Long = 1) As String
    Dim FSO As Object
    Dim suffix As Long
    Dim maxSuffix As Long
    Dim filePath As String
    Dim parentFolderName As String
    Dim baseName As String
    Dim extensionName As String
    
    '桁は1~9にする(1以上なので1桁以上、longの最大値が9桁のため9桁以下)
    'the digit is between 1 to 9 as the suffix number is above 1, and max number of long is 9 digit.
    If digit < 1 Then digit = 1
    If digit > 9 Then digit = 9
    
    maxSuffix = CLng(10 ^ digit) - 1
    
    If startNumber > maxSuffix Then Err.Raise 1000, , "開始番号が指定の桁数を超えています"
    
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    '指定のファイルがない場合はそのまま返す。
    'if the speficied targetPath is not exist, return the path directory
    If Not FSO.FileExists(targetPath) Then
        getUniquePath = targetPath
        startNumber = 0 'byRefで受け取っているため指定ファイル名がそのまま見つかったので0を返す
        Exit Function
    End If
    

    parentFolderName = FSO.GetParentFolderName(targetPath)
    baseName = FSO.GetBaseName(targetPath)
    extensionName = FSO.GetExtensionName(targetPath)
    
    For suffix = startNumber To maxSuffix
    
        filePath = FSO.BuildPath(parentFolderName, _
                            baseName & _
                            "_" & Format(suffix, String(digit, "0")) & "." & _
                            extensionName)
    
        If Not FSO.FileExists(filePath) Then
            If extensionName = "" And Right(filePath, 1) = "." Then
                filePath = Left(filePath, Len(filePath) - 1)
            End If
            getUniquePath = filePath
            startNumber = suffix    'byRefで受け取っているため今回発行した番号を代入する
            Exit Function
        End If

    Next
    
    Err.Raise 1001, , "重複しないファイル名はありません"

End Function
