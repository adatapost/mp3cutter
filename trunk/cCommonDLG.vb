Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cCommonDLG_NET.cCommonDLG")> Public Class cCommonDLG
	
	'===============================================================
	'=============================================================
	'--------------------/\API/\
	'UPGRADE_WARNING: Structure OPENFILENAME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetOpenFileName Lib "comdlg32.dll"  Alias "GetOpenFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'UPGRADE_WARNING: Structure OPENFILENAME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetSaveFileName Lib "comdlg32.dll"  Alias "GetSaveFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'-------------------/\Type/\
	Private Structure OPENFILENAME
		Dim lStructSize As Integer
		Dim hWndOwner As Integer
		Dim hInstance As Integer
		Dim lpstrFilter As String
		Dim lpstrCustomFilter As String
		Dim nMaxCustFilter As Integer
		Dim nFilterIndex As Integer
		Dim lpstrFile As String
		Dim nMaxFile As Integer
		Dim lpstrFileTitle As String
		Dim nMaxFileTitle As Integer
		Dim lpstrInitialDir As String
		Dim lpstrTitle As String
		Dim Flags As Integer
		Dim nFileOffset As Short
		Dim nFileExtension As Short
		Dim lpstrDefExt As String
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
	End Structure
	'-----------------------/\Constants/\
	Private Const OFN_ALLOWMULTISELECT As Short = &H200s
	Private Const OFN_EXPLORER As Integer = &H80000 '  new look commdlg
	Private Const OFN_OVERWRITEPROMPT As Short = &H2s
	Private Const OFN_FILEMUSTEXIST As Short = &H1000s
	Private Const cdlCancel As Short = 32755
	'----------------------/\Variable/\
	Dim pFileName As String
	Dim pFlag As String
	Dim pHwnd As Integer
	Dim OFName As OPENFILENAME
	'=============================================================
	Private Sub Init()
		'-----------------
		With OFName
			'------------
			.hInstance = VB6.GetHInstance.ToInt32
			.lpstrFile = Space(4095)
			.nMaxFile = 4096
			'---------
		End With
		'----------
	End Sub
	Public Sub ParseMultiFileName(ByVal sFileName As String, ByRef dFiles() As String)
		'-----------------
		If Len(sFileName) > 3 Then
			'------
			sFileName = Left(sFileName, InStr(sFileName, Chr(0) & Chr(0)) - 1)
			dFiles = Split(sFileName, Chr(0))
			'------
		End If
		'----------------
	End Sub
	'============================================================
	Public Sub ShowOpen()
		'--------------
		On Error Resume Next
		'--------------
		Call Init()
		OFName.Flags = CInt(Flags)
		OFName.lpstrFilter = "Mp3 File" & Chr(0) & "*.Mp3" & Chr(0) & "All Files" & Chr(0) & "*.*" & Chr(0)
		If GetOpenFileName(OFName) Then
			'--------
			FileName = OFName.lpstrFile
			'----
		Else
			'---
			FileName = ""
			'--------------
		End If
		'---------------
	End Sub
	'=============================================================
	Public Sub ShowSave()
		'--------------
		On Error Resume Next
		'--------------
		Call Init()
		OFName.Flags = CInt(Flags)
		OFName.lpstrFilter = "Mp3 File" & Chr(0) & "*.Mp3" & Chr(0)
		If GetSaveFileName(OFName) Then
			'--------
			With OFName
				If LCase(Mid(.lpstrFile, InStrRev(.lpstrFile, ".") + 1, 3)) = "mp3" Then
					FileName = Left(.lpstrFile, InStr(.lpstrFile, Chr(0)) - 1)
				Else
					FileName = Left(.lpstrFile, InStr(.lpstrFile, Chr(0)) - 1) & ".mp3"
				End If
			End With
			'----
		Else
			'---
			FileName = ""
			'--------
		End If
		'--------------
	End Sub
	'=============================================================
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'----------
		OFName.lStructSize = Len(OFName)
		'----------
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	'=============================================================
	'All Properties
	'=============================================================
	'=============================================================
	Public Property FileName() As String
		Get
			'----------------
			FileName = pFileName
			'----------------
		End Get
		Set(ByVal Value As String)
			'--------------
			pFileName = Value
			'--------------
		End Set
	End Property
	'=============================================================
	'=============================================================
	Public Property HWND() As Integer
		Get
			'--------------
			HWND = pHwnd
			'--------------
		End Get
		Set(ByVal Value As Integer)
			'-------------
			pHwnd = Value
			OFName.hWndOwner = pHwnd
			'-------------
		End Set
	End Property
	'=============================================================
	'=============================================================
	Public Property Flags() As String
		Get
			'-----------------
			Flags = pFlag
			'-----------------
		End Get
		Set(ByVal Value As String)
			'-----------------
			pFlag = Value
			'-----------------
		End Set
	End Property
	'=============================================================
End Class