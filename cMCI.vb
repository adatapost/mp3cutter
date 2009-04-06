Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cMCI_NET.cMCI")> Public Class cMCI
	
	Private Declare Function mciSendString Lib "winmm.dll"  Alias "mciSendStringA"(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
	Private Declare Function mciGetErrorString Lib "winmm.dll"  Alias "mciGetErrorStringA"(ByVal dwError As Integer, ByVal lpstrBuffer As String, ByVal uLength As Integer) As Integer
	Dim pFileName As String
	Dim lngError As Integer

	Public Sub OpenMCI()
		'-----------------
		If FileName <> "" Then
			lngError = mciSendString("Open " & FileName & " Alias CurrMp3", CStr(0), 0, 0)
			lngError = mciSendString("Set CurrMp3 time format 0", CStr(0), 0, 0)
		End If
		'-----------------
	End Sub
	'=============================================================
	Public Sub SetVolume(ByRef Volume As Short)
		'----------------
		lngError = mciSendString("Setaudio CurrMp3 Volume to " & Volume, CStr(0), 0, 0)
		'----------------
	End Sub
	 Public Function ShowError(ByRef Error_Renamed As Integer) As String
        '-----------------
        Dim Buffer As New VB6.FixedLengthString(256)
        '-----------------
        mciGetErrorString(Error_Renamed, Buffer.Value, Len(Buffer.Value))
        Buffer.Value = Left(Buffer.Value, InStr(Buffer.Value, Chr(0)) - 1)
        ShowError = Buffer.Value
        '-----------------
    End Function
	 Public Sub CloseMCI()
        '--------------
        lngError = mciSendString("Close All", CStr(0), 0, 0)
        '-------------
    End Sub
	 Public Sub Play(Optional ByRef From As Single = 0)
        '----------------
        lngError = mciSendString("Play CurrMp3 from " & From, CStr(0), 0, 0)
        '---------------
    End Sub
	 Public Sub StopMCI()
        '-----------
        lngError = mciSendString("Stop CurrMp3", CStr(0), 0, 0)
        '-----------
    End Sub
 	Public Sub SeekTo(ByRef MilliSec As Single)
        '-------------
        lngError = mciSendString("Seek CurrMp3 To " & MilliSec, CStr(0), 0, 0)
        '-------------
    End Sub
 	Public Sub Pause()
        '-------------
        lngError = mciSendString("Pause CurrMp3", CStr(0), 0, 0)
        '------------
    End Sub
 	Public Sub ResumeMCI()
        '-----------------
        lngError = mciSendString("Resume CurrMp3", CStr(0), 0, 0)
        '-----------------
    End Sub
 	Public Function Mode() As String
        '-----------------
        Dim Ret As New VB6.FixedLengthString(255)
        lngError = mciSendString("Status CurrMp3 Mode", Ret.Value, Len(Ret.Value), 0)
        Ret.Value = Left(Ret.Value, InStr(Ret.Value, Chr(0)) - 1)
        Mode = Ret.Value
        '-----------------
    End Function
 	Public Function Length() As String
        '-------------
        Dim Ret As New VB6.FixedLengthString(255)
        lngError = mciSendString("Status CurrMp3 Length", Ret.Value, Len(Ret.Value), 0)
        Ret.Value = Left(Ret.Value, InStr(Ret.Value, Chr(0)) - 1)
        Length = Ret.Value
        '------------
    End Function
 	Public Function Position() As String
        '-------------
        Dim Ret As New VB6.FixedLengthString(255)
        lngError = mciSendString("Status CurrMp3 Position", Ret.Value, Len(Ret.Value), 0)
        Ret.Value = Left(Ret.Value, InStr(Ret.Value, Chr(0)) - 1)
        Position = Ret.Value
        '------------
    End Function
 	Public Property FileName() As String
        Get
            '----------------
            FileName = pFileName
            '----------------
        End Get
        Set(ByVal Value As String)
            '-----------------
            pFileName = Chr(34) & Value & Chr(34)
            '-----------------
        End Set
    End Property
End Class