Option Strict Off
Option Explicit On
Friend Class frmMenu
	Inherits System.Windows.Forms.Form
	
	'This form just hold menu for frmMain.frm
	Private Sub RemoveItemFromList(Optional ByRef SelectedItem As Boolean = True, Optional ByRef All As Boolean = False)
		'---------------
		Dim i As Integer
		'--------------
		If All = True Then
			frmMain.lstPlayList.Items.Clear()
			frmMain.lstPlayListPath.Items.Clear()
			Exit Sub
		End If
		'--------------
Re_search: 
		For i = 0 To frmMain.lstPlayList.Items.Count - 1
			'----------------
			If frmMain.lstPlayList.GetSelected(i) = SelectedItem Then
				frmMain.lstPlayList.Items.RemoveAt((i))
				frmMain.lstPlayListPath.Items.RemoveAt((i))
				GoTo Re_search
			End If
			'-----------
			System.Windows.Forms.Application.DoEvents()
		Next i
		'---------------
	End Sub
	Private Sub FixNumbers()
		'---------------
		Dim i As Integer
		'---------------
		With frmMain.lstPlayList
			For i = 0 To .Items.Count - 1
				VB6.SetItemString(frmMain.lstPlayList, i, i + 1 & ". " & Mid(VB6.GetItemString(frmMain.lstPlayList, i), InStr(VB6.GetItemString(frmMain.lstPlayList, i), ".") + 1))
				System.Windows.Forms.Application.DoEvents()
			Next i
		End With
		'---------------
	End Sub
	
	'==============================================================
	Public Sub mnuABS_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuABS.Click
		'---------------
		RemoveItemFromList(False)
		FixNumbers()
		'---------------
	End Sub
	'==============================================================
	Public Sub mnuAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAll.Click
		'------------------
		RemoveItemFromList( , True)
		'------------------
	End Sub
	'==============================================================
	Public Sub mnuSelected_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSelected.Click
		'---------------
		RemoveItemFromList(True)
		FixNumbers()
		'---------------
	End Sub
	'===============================================================
	Public Sub mnuTimeElapsed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTimeElapsed.Click
		'----------------
		If mnuTimeElapsed.Checked = False Then frmMain.TimeSetting()
		'----------------
	End Sub
	'===========================================================
	Public Sub mnuTimeRemaining_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTimeRemaining.Click
		'----------------
		If mnuTimeRemaining.Checked = False Then frmMain.TimeSetting()
		'----------------
	End Sub
	'============================================================

    Private Sub frmMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class