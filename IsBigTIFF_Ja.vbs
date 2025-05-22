'--------------------------------------------------------------------------
'    �C�ӂ�TIF�t�@�C�����h���b�O���h���b�v���� BigTIFF ���ۂ��𔻒肷��
'                    Created by Sewege11 at 2025-05-21
'--------------------------------------------------------------------------
Option Explicit
' �t�@�C���̃t���p�X���擾
Dim uFullpath
uFullpath = WScript.Arguments.Item(0)
Dim uExt3, uExt4
uExt3 = UCase(Right(uFullpath, 3))
uExt4 = UCase(Right(uFullpath, 4))
If uExt3 <> "TIF" Then
  If uExt4 <> "TIFF" Then
    Msgbox "�g���q�� .tif �� .tiff �̃t�@�C�����h���b�v���ĉ������B", 16, "�G���["
    WScript.Quit
  End If
End If
' ADODB.Stream��3�o�C�g�ڂ��擾
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
		MsgBox "�t�@�C�����ǂݍ��߂܂���ł����B", 48, "�G���["
		WScript.Quit
	End If
	On Error GoTo 0
	.Position = 2
	uBuf = .Read(2)
	.Close
End With
' ����
Dim uRetVal
Dim eMsgStyle: eMsgStyle = 64
Select Case Hex(Asc(uBuf))
	Case "2A" ' TIFF
		uRetVal = uFullpath & vbCrLf & "����� ���ʂ�TIFF �ł��B(2A)"
	Case "2B" ' BigTIFF
		uRetVal = uFullpath & vbCrLf & "����� BigTIFF �ł��B(2B)"
	Case Else
		uRetVal = uFullpath & vbCrLf & "�����TIFF�ł͂Ȃ������ł�"
	  eMsgStyle = 48
End Select
' ���ʕ\��
MsgBox uRetVal, eMsgStyle, "TIFF / BigTIFF ����VBScript"
' �I�u�W�F�N�g�ϐ��̉��
Set oStream = Nothing
'[EOF]