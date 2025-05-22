'--------------------------------------------------------------------------
' Drag and drop any TIF file to determine whether it is a BigTIFF or not.
'                    Created by Sewege11 at 2025-05-21
'--------------------------------------------------------------------------
Option Explicit
' Get fullpath of file by drag and drop
Dim uFullpath
uFullpath = WScript.Arguments.Item(0)
Dim uExt3, uExt4
uExt3 = UCase(Right(uFullpath, 3))
uExt4 = UCase(Right(uFullpath, 4))
If uExt3 <> "TIF" Then
  If uExt4 <> "TIFF" Then
    Msgbox "Please drop a file with the extension .tif or .tiff.", 16, "Error"
    WScript.Quit
  End If
End If
' Get the third byte by ADODB.Stream
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
		MsgBox "The file could not be read.", 48, "Error"
		WScript.Quit
	End If
	On Error GoTo 0
	.Position = 2
	uBuf = .Read(2)
	.Close
End With
' Identify the type of TIFF file
Dim uRetVal
Dim eMsgStyle: eMsgStyle = 64
Select Case Hex(Asc(uBuf))
	Case "2A" ' TIFF
		uRetVal = uFullpath & vbCrLf & "This is a normal TIFF. (2A)"
	Case "2B" ' BigTIFF
		uRetVal = uFullpath & vbCrLf & "This is a BigTIFF. (2B)"
	Case Else
		uRetVal = uFullpath & vbCrLf & "This doesn't seem to be a TIFF."
	  eMsgStyle = 48
End Select
' Show result
MsgBox uRetVal, eMsgStyle, "TIFF or BigTIFF judgement script"
' Release memory
Set oStream = Nothing
'[EOF]