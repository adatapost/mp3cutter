Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cBrowseForFolder_NET.cBrowseForFolder")> Public Class cBrowseForFolder
	
 	Private Structure BrowseInfo
        Dim hWndOwner As Integer
        Dim pIDLRoot As Integer
        Dim pszDisplayName As Integer
        Dim lpszTitle As Integer
        Dim ulFlags As Integer
        Dim lpfnCallback As Integer
        Dim lParam As Integer
        Dim iImage As Integer
    End Structure
	'---------------------------/\Constants/\
	Const BIF_RETURNONLYFSDIRS As Short = 1
	Const MAX_PATH As Short = 260
	'------------------------------/\API/\
	Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Integer)
	Private Declare Function lstrcat Lib "kernel32"  Alias "lstrcatA"(ByVal lpString1 As String, ByVal lpString2 As String) As Integer

	Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef lpbi As BrowseInfo) As Integer
	Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Integer, ByVal lpBuffer As String) As Integer

	Public Function BrowseForFolder(ByRef HWND As Integer) As String
        Dim iNull As Short
		Dim lpIDList, lResult As Integer
		Dim sPath As String
		Dim udtBI As BrowseInfo
		'----------------
		With udtBI
			'Set the owner window
			.hWndOwner = HWND
			'lstrcat appends the two strings and returns the memory address
            .lpszTitle = lstrcat("Select a directory", "")
			'Return only if the user selected a directory
			.ulFlags = BIF_RETURNONLYFSDIRS
		End With
		'----------------
		'Show the 'Browse for folder' dialog
		lpIDList = SHBrowseForFolder(udtBI)
		If lpIDList Then
			sPath = New String(Chr(0), MAX_PATH)
			'Get the path from the IDList
			SHGetPathFromIDList(lpIDList, sPath)
			'free the block of memory
			CoTaskMemFree(lpIDList)
			iNull = InStr(sPath, vbNullChar)
			If iNull Then
				sPath = Left(sPath, iNull - 1)
			End If
		End If
		'-----------------
		BrowseForFolder = sPath
		'-----------------
	End Function

End Class