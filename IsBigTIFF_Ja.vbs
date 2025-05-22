'--------------------------------------------------------------------------
'    任意のTIFファイルをドラッグ＆ドロップして BigTIFF か否かを判定する
'                    Created by Sewege11 at 2025-05-21
'--------------------------------------------------------------------------
Option Explicit
' ファイルのフルパスを取得
Dim uFullpath
uFullpath = WScript.Arguments.Item(0)
Dim uExt3, uExt4
uExt3 = UCase(Right(uFullpath, 3))
uExt4 = UCase(Right(uFullpath, 4))
If uExt3 <> "TIF" Then
  If uExt4 <> "TIFF" Then
    Msgbox "拡張子が .tif か .tiff のファイルをドロップして下さい。", 16, "エラー"
    WScript.Quit
  End If
End If
' ADODB.Streamで3バイト目を取得
Dim oStream
Set oStream = CreateObject("ADODB.Stream")
Dim uBuf
With oStream
	.Type = 1
	.Open
	On Error Resume Next
	.LoadFromFile uFullpath
	If Err.Number <> 0 Then
		.Close
		MsgBox "ファイルが読み込めませんでした。", 48, "エラー"
		WScript.Quit
	End If
	On Error GoTo 0
	.Position = 2
	uBuf = .Read(2)
	.Close
End With
' 判定
Dim uRetVal
Dim eMsgStyle: eMsgStyle = 64
Select Case Hex(Asc(uBuf))
	Case "2A" ' TIFF
		uRetVal = uFullpath & vbCrLf & "これは 普通のTIFF です。(2A)"
	Case "2B" ' BigTIFF
		uRetVal = uFullpath & vbCrLf & "これは BigTIFF です。(2B)"
	Case Else
		uRetVal = uFullpath & vbCrLf & "これはTIFFではなさそうです"
	  eMsgStyle = 48
End Select
' 結果表示
MsgBox uRetVal, eMsgStyle, "TIFF / BigTIFF 判定VBScript"
' オブジェクト変数の解放
Set oStream = Nothing
'[EOF]