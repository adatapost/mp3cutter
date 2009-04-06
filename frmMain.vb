Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
    Public kint As Integer

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Object) As Integer
	Private Declare Sub ReleaseCapture Lib "user32" ()


	Const WM_NCLBUTTONDOWN As Short = &HA1s
	Const HTCAPTION As Short = 2
	Const Chunk As Integer = 2 ^ 20 '1 MB

	Private Const OFN_ALLOWMULTISELECT As Short = &H200s
	Private Const OFN_EXPLORER As Integer = &H80000 '  new look commdlg
	Private Const OFN_OVERWRITEPROMPT As Short = &H2s
	Private Const OFN_FILEMUSTEXIST As Short = &H1000s
	'-----------------------/\varibales/\
	Dim CD As New cCommonDLG
	Dim MCI As New cMCI
	Dim BFF As cBrowseForFolder
	Dim TimeElapsed As Boolean
	'=========================================================
	Public Sub TimeSetting()
		'-----------------
		Negative.Visible = Not (Negative.Visible)
		TimeElapsed = Not (TimeElapsed)
		If TimeElapsed = True Then
			frmMenu.mnuTimeElapsed.Checked = True
			frmMenu.mnuTimeRemaining.Checked = False
		Else
			frmMenu.mnuTimeElapsed.Checked = False
			frmMenu.mnuTimeRemaining.Checked = True
		End If
		'-----------------
	End Sub
	 Private Sub SplitFile(ByRef Path As String, ByRef strArray() As String, ByRef nRecord As Single, ByRef Size_Renamed As Single)
        '--------------------
        Dim nDivide As Short 'Number of divide
        Dim sngResult As Single
        Dim i As Short 'loop variable
        '-------------------
        If Size_Renamed < Chunk Then 'less than 1MB
            ReDim strArray(1)
            strArray(1) = GetFromFile(Path, nRecord, Size_Renamed)
        Else
            nDivide = Size_Renamed \ Chunk
            ReDim strArray(nDivide)
            '---------
            For i = 1 To nDivide
                strArray(i) = GetFromFile(Path, (i - 1) * Chunk + nRecord, Chunk)
                System.Windows.Forms.Application.DoEvents()
            Next i
            '---------
            If (nDivide * Chunk) < Size_Renamed Then
                sngResult = Size_Renamed - (nDivide * Chunk)
                ReDim Preserve strArray(nDivide + 1)
                strArray(nDivide + 1) = GetFromFile(Path, nDivide * Chunk, sngResult)
            End If
        End If
        '--------------------
    End Sub
	Private Function GetFromFile(ByRef Path As String, ByRef nRecord As Single, ByRef nByte As Single) As String
        '------------------
        Dim FN As Short 'file number
        Dim strBuffer As String 'buffer for reading from file
        '------------------
        FN = FreeFile()
        strBuffer = Space(nByte)
        FileOpen(FN, Path, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
        FileGet(FN, strBuffer, nRecord, True)
        FileClose(FN)
        GetFromFile = strBuffer
        '------------------
    End Function
	'=========================================================
	Private Sub PutToFile(ByRef Path As String, ByRef nRecord As Single, ByRef Data As String)
		'-----------------
		Dim FN As Short 'file number
		'------------------
		FN = FreeFile
		FileOpen(FN, Path, OpenMode.Binary, OpenAccess.Write)
        FilePut(FN, Data, CInt(nRecord))
		FileClose(FN)
		'-----------------
	End Sub
	'===============================================================
	Private Sub CutMp3(ByRef Source As String, ByRef Destination As String, Optional ByRef nDivide As Integer = 0)
		'--------------
		Dim sngTemp As Single
		Dim sngFileSize As Single
		Dim sngRecordNumber As Single
		Dim sngResult As Single
		Dim intNum As Short
		Dim nLoop As Short
		Dim ProgAdd As Integer
		Dim i As Short
		Dim j As Short
		'--------------
		If nDivide > 0 Then
			If VB.Right(Destination, 1) <> "\" Then Destination = Destination & "\"
			sngFileSize = FileLen(Source)
			sngTemp = sngFileSize \ nDivide
			If sngTemp < Chunk Then
				ProgAdd = 100 \ nDivide
				For i = 1 To nDivide
					PutToFile(Destination & "MyControls" & i & ".mp3", 1, GetFromFile(Source, (i - 1) * sngTemp + 1, sngTemp))
					Progress.Value = Progress.Value + ProgAdd
					System.Windows.Forms.Application.DoEvents()
				Next i
				MsgBox("Your file has been divided successfully", MsgBoxStyle.Information, "Finish")
				Progress.Value = 0
			Else
				sngRecordNumber = 1
				For i = 1 To nDivide
					intNum = sngTemp \ Chunk
					ProgAdd = 100 \ (nDivide * intNum)
					For j = 1 To intNum
						PutToFile(Destination & "MyControls" & i & ".mp3", (j - 1) * Chunk + 1, GetFromFile(Source, sngRecordNumber, Chunk))
						sngRecordNumber = sngRecordNumber + Chunk
						Progress.Value = Progress.Value + ProgAdd
						System.Windows.Forms.Application.DoEvents()
					Next j
				Next i
				If sngRecordNumber < sngFileSize Then
					sngResult = sngFileSize - sngRecordNumber
					PutToFile(Destination & "MyControls" & i - 1 & ".mp3", (j - 1) * Chunk + 1, GetFromFile(Source, sngRecordNumber, sngResult))
				End If
				MsgBox("Your file has been divided successfully", MsgBoxStyle.Information, "Finish")
				Progress.Value = 0
			End If
			'---------------------------------
		Else
			sngFileSize = FileLen(Source) 'Get file size
			sngRecordNumber = (Int(Pos.ClipMarkLeft) * sngFileSize) \ (CDbl(MCI.Length) \ 1000)
			If sngRecordNumber <= 0 Then sngRecordNumber = 1
			sngTemp = (Int(Pos.ClipMarkLeft + Pos.ClipMarkWidth) * sngFileSize) \ (CDbl(MCI.Length) \ 1000)
			sngTemp = sngTemp - sngRecordNumber
			'---------
			If sngTemp < Chunk Then 'less than 1 MB
				PutToFile(Destination, 1, GetFromFile(Source, sngRecordNumber, sngTemp))
				Progress.Value = 100
				MsgBox("Part of music has been saved successfully", MsgBoxStyle.Information, "Finish")
				Progress.Value = 0
			Else
				intNum = sngTemp \ Chunk

				sngResult = sngTemp Mod Chunk
				ProgAdd = 100 \ intNum
				For i = 1 To intNum
					PutToFile(Destination, (i - 1) * Chunk + 1, GetFromFile(Source, sngRecordNumber + (i - 1) * Chunk, Chunk))
					Progress.Value = Progress.Value + intNum
					System.Windows.Forms.Application.DoEvents()
				Next i
				If sngResult <> 0 Then PutToFile(Destination, (i - 1) * Chunk + 1, GetFromFile(Source, sngRecordNumber + (i - 1) * Chunk, sngResult))
				MsgBox("Part of music has been saved successfully", MsgBoxStyle.Information, "Finish")
				Progress.Value = 0
			End If
		End If
		'--------------------
	End Sub
    Private Sub ShowTabContents(ByRef TabIndex_Renamed As Short, ByRef Show_Renamed As Boolean)
        '--------------
        Dim strHelp As String
        '--------------
        kint = TabIndex_Renamed

        Select Case TabIndex_Renamed
            '----------

            Case 0
                strHelp = "1-Specify begining of part of music by clicking:" & vbCrLf & vbCrLf
                strHelp = strHelp & "2-Specify end of part of music by clicking:" & vbCrLf & vbCrLf
                strHelp = strHelp & "3-Click 'Cut and save as' button"
                lblHelp.Text = strHelp
                lblHelp.SetBounds(16, 220, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                lblHelp.Visible = Show_Renamed
                imgHelp(0).SetBounds(235, 219, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                imgHelp(0).Visible = Show_Renamed
                imgHelp(1).SetBounds(213, 245, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                imgHelp(1).Visible = Show_Renamed
                optClipMark(0).SetBounds(VB6.TwipsToPixelsX(10), VB6.TwipsToPixelsY(120), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                optClipMark(0).Visible = Show_Renamed
                optClipMark(1).SetBounds(VB6.TwipsToPixelsX(480), VB6.TwipsToPixelsY(120), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                optClipMark(1).Visible = Show_Renamed
                fltCutAndSave.SetBounds(96, 176, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                fltCutAndSave.Visible = Show_Renamed
                Frame1.SetBounds(16, 170, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Frame1.Visible = Show_Renamed
                Progress.SetBounds(16, 312, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Progress.Visible = Show_Renamed

                '-----------
            Case 1
                Label(0).SetBounds(16, 176, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Label(0).Text = "select your original file"
                Label(0).Visible = Show_Renamed
                Label(1).SetBounds(16, 216, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Label(1).Text = "select your attachment file"
                Label(1).Visible = Show_Renamed
                fltBrowse(0).SetBounds(208, 192, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                fltBrowse(0).Visible = Show_Renamed
                fltBrowse(1).SetBounds(208, 232, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                fltBrowse(1).Visible = Show_Renamed
                txtSource1.SetBounds(16, 192, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

                txtSource1.Text = ""
                txtSource1.Visible = Show_Renamed
                txtDir1.SetBounds(16, 232, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

                txtDir1.Text = ""
                txtDir1.Visible = Show_Renamed
                optMerge(0).SetBounds(16, 272, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                optMerge(0).Visible = Show_Renamed
                optMerge(1).SetBounds(16, 288, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                optMerge(1).Visible = Show_Renamed
                optMerge(2).SetBounds(16, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                optMerge(2).Visible = Show_Renamed
                fltMerge.SetBounds(128, 280, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                fltMerge.Visible = Show_Renamed
                txtTime(0).SetBounds(136, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                txtTime(0).Visible = Show_Renamed
                txtTime(1).SetBounds(192, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                txtTime(1).Visible = Show_Renamed
                Label2(0).SetBounds(160, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Label2(0).Visible = Show_Renamed
                Label2(1).SetBounds(216, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Label2(1).Visible = Show_Renamed

                Progress.SetBounds(16, 312, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Progress.Visible = Show_Renamed
                '-----------
            Case 2
                Label(0).SetBounds(16, 176, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Label(0).Text = "Source Mp3 file name"
                Label(0).Visible = Show_Renamed
                Label(1).SetBounds(16, 216, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Label(1).Text = "Destination dir to save"
                Label(1).Visible = Show_Renamed
                Label(2).SetBounds(16, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Label(2).Visible = Show_Renamed
                fltBrowse(0).SetBounds(208, 192, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                fltBrowse(0).Visible = Show_Renamed
                fltBrowse(1).SetBounds(208, 232, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                fltBrowse(1).Visible = Show_Renamed
                txtSource1.SetBounds(16, 192, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

                txtSource1.Text = ""
                txtSource1.Visible = Show_Renamed
                txtDir1.SetBounds(16, 232, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

                txtDir1.Text = ""
                txtDir1.Visible = Show_Renamed
                txtNumDivide.SetBounds(112, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

                txtNumDivide.Text = ""
                txtNumDivide.Visible = Show_Renamed
                fltDivide.SetBounds(176, 256, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                fltDivide.Visible = Show_Renamed
                Progress.SetBounds(16, 312, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Progress.Visible = Show_Renamed
                '------------
            Case 3
                lblAbout(0).SetBounds(30, 190, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                lblAbout(0).Visible = Show_Renamed
                lblAbout(1).SetBounds(24, 230, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                lblAbout(1).Visible = Show_Renamed
                lblAbout(2).SetBounds(42, 262, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                lblAbout(2).Visible = Show_Renamed
                Progress.SetBounds(16, 312, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                Progress.Visible = False
        End Select
        '--------------
    End Sub

    Private Sub Init()
        '-------------Volume Setting----------
        Volume.Length = 1000
        Volume.SeekTo(Val(GetSetting("MyControls", "Setting", "Volume", CStr(500))))
        MCI.SetVolume((Volume.Value))
        '------------Time Setting------------
        TimeElapsed = CBool(GetSetting("MyControls", "Setting", "TimeElapsed", "True"))
        Negative.Visible = Not (TimeElapsed)
        If TimeElapsed = True Then
            frmMenu.mnuTimeElapsed.Checked = True
            frmMenu.mnuTimeRemaining.Checked = False
        Else
            frmMenu.mnuTimeElapsed.Checked = False
            frmMenu.mnuTimeRemaining.Checked = True
        End If
        '-----------Tab Setting-----------
        fltTab_ClickEvent(fltTab.Item(CShort(Val(GetSetting("MyControls", "Setting", "TabIndex", CStr(0))))), New System.EventArgs())
        '------------progressBar
        With Progress
            .Min = 0
            .Max = 100
            .Value = 0
        End With
    
    End Sub
    Private Sub chkMute_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMute.CheckStateChanged

        If chkMute.CheckState = 1 Then
            MCI.SetVolume(0)
        Else
            MCI.SetVolume(Int(Volume.Value))
        End If

    End Sub

    Private Sub CloseApp_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CloseApp.ClickEvent
        Me.Opacity = 0.6
        MCI.CloseMCI()
        CD = Nothing
        MCI = Nothing
        End

    End Sub

    Private Sub ShowTime(ByRef strTime As String)
        Dim i As Byte
        For i = 1 To Len(strTime)
            If MusicTime(i - 1).PictureId <> Val(Mid(strTime, i, 1)) + 11 Then MusicTime(i - 1).PictureId = Val(Mid(strTime, i, 1)) + 11
            System.Windows.Forms.Application.DoEvents()
        Next i
        '------------
    End Sub

    Private Function MilliSecToMin(ByVal MilliSec As Single) As String
        Dim Second_Renamed As Single
        Dim Min As Single
        Second_Renamed = MilliSec \ 1000
        Min = Second_Renamed \ 60
        Second_Renamed = Second_Renamed Mod 60
        MilliSecToMin = VB6.Format(Str(Min), "00") & VB6.Format(Str(Second_Renamed), "00")
    End Function

    Private Sub Colon_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMyControls.__CustomControl_MouseDownEvent) Handles Colon.MouseDownEvent
        If eventArgs.Button = VB6.MouseButtonConstants.LeftButton Then
            TimeSetting()
        Else

            frmMenu.mnuTime.ShowDropDown()
        End If
    End Sub

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Shell(System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(Application.ExecutablePath)) & "\iexplore.exe", AppWinStyle.MaximizedFocus)
    End Sub

    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        VideoForm.Show()


        frmMenu.Visible = False
 
    End Sub

     Private Sub fltBrowse_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fltBrowse.ClickEvent
        Dim Index As Short = fltBrowse.GetIndex(eventSender)


        If Index = 0 Or kint = 1 Then
            CD.HWND = Me.Handle.ToInt32
            CD.Flags = CStr(OFN_EXPLORER Or OFN_FILEMUSTEXIST)
            CD.ShowOpen()
            If CD.FileName <> "" Then
                If Index = 0 Then
                    txtSource1.Text = CD.FileName
                Else
                    txtDir1.Text = CD.FileName
                End If
            End If
        ElseIf Index = 1 And CDbl(fltTab(0).Tag) <> 1 Then
            BFF = New cBrowseForFolder
            txtDir1.Text = BFF.BrowseForFolder(Me.Handle.ToInt32)
            BFF = Nothing
        End If
    End Sub
    Private Sub fltDivide_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fltDivide.ClickEvent
        On Error GoTo Err_Handle
        If Val(txtNumDivide.Text) <= 0 Then
            MsgBox("Please inter divide number", MsgBoxStyle.Critical, "Error")
            txtNumDivide.Text = ""
            txtNumDivide.Focus()
            Exit Sub
        ElseIf txtSource1.Text = "" Then
            MsgBox("Please specify source mp3 file name", MsgBoxStyle.Critical, "Error")
            Exit Sub
        ElseIf txtDir1.Text = "" Then
            MsgBox("Please specify destination directory", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
        Progress.Value = 0
        CutMp3((txtSource1.Text), (txtDir1.Text), Val(txtNumDivide.Text))
        Exit Sub
Err_Handle:
        MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
        Progress.Value = 0
    End Sub
    Private Sub fltMerge_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fltMerge.ClickEvent
        On Error GoTo Err_Handle
        Dim strTemp1() As String
        Dim strTemp2() As String
        Dim strTemp3() As String
        Dim sngFileSize As Single
        Dim sngRecordNumber As Single
        Dim TotalTime As Single
        Dim MergeTime As Single
        Dim ProgAdd As Integer
        Dim i As Short
        Progress.Value = 0
        If txtSource1.Text = "" Then
            MsgBox("Please specify original mp3 file name", MsgBoxStyle.Critical, "Error")
            Exit Sub
        ElseIf txtDir1.Text = "" Then
            MsgBox("Please specify attachment file name", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
        If optMerge(0).Checked = True Then 'insert at start
            SplitFile((txtSource1.Text), strTemp1, 1, FileLen(txtSource1.Text))
            Progress.Value = 10
            SplitFile((txtDir1.Text), strTemp2, 1, FileLen(txtDir1.Text))
            Progress.Value = 20
            Kill(txtSource1.Text)
            ProgAdd = 80 \ (UBound(strTemp2) + UBound(strTemp1))
            For i = 1 To UBound(strTemp2)
                PutToFile((txtSource1.Text), (i - 1) * Chunk + 1, strTemp2(i))
                Progress.Value = Progress.Value + ProgAdd
                System.Windows.Forms.Application.DoEvents()
            Next i
            sngFileSize = FileLen(txtSource1.Text) + 1
            For i = 1 To UBound(strTemp1)
                PutToFile((txtSource1.Text), (i - 1) * Chunk + sngFileSize, strTemp1(i))
                Progress.Value = Progress.Value + ProgAdd
                System.Windows.Forms.Application.DoEvents()
            Next i
            MsgBox("File has been attached as a begining successfully", MsgBoxStyle.Information, "Finish")
            Progress.Value = 0
            Exit Sub
        End If
        '----------------------
        If optMerge(1).Checked = True Then 'insert at end
            SplitFile((txtDir1.Text), strTemp2, 1, FileLen(txtDir1.Text))
            Progress.Value = 50
            sngFileSize = FileLen(txtSource1.Text) + 1
            ProgAdd = 50 \ UBound(strTemp2)
            For i = 1 To UBound(strTemp2)
                PutToFile((txtSource1.Text), sngFileSize + (i - 1) * Chunk, strTemp2(i))
                Progress.Value = Progress.Value + ProgAdd
                System.Windows.Forms.Application.DoEvents()
            Next i
            MsgBox("File has been attached as a ending successfully", MsgBoxStyle.Information, "Finish")
            Progress.Value = 0
            Exit Sub
        End If
        Dim MCI2 As New cMCI
        If optMerge(2).Checked = True Then 'insert at middle
            MCI2.FileName = txtSource1.Text
            MCI2.OpenMCI()
            TotalTime = CDbl(MCI2.Length) \ 1000
            MergeTime = Val(txtTime(0).Text) * 60 + Val(txtTime(1).Text)
            If (MergeTime > TotalTime Or Val(txtTime(0).Text) < 0 Or Val(txtTime(1).Text) < 0) Or (Val(txtTime(0).Text) = 0 And Val(txtTime(1).Text) = 0) Then
                MsgBox("out of real time", MsgBoxStyle.Critical, "Error")
                txtTime(0).Text = ""
                txtTime(1).Text = ""
                txtTime(0).Focus()
                Exit Sub
            End If
            MCI2.CloseMCI()
            MCI2 = Nothing
            sngFileSize = FileLen(txtSource1.Text)
            sngRecordNumber = MergeTime * sngFileSize \ TotalTime
            SplitFile((txtSource1.Text), strTemp1, 1, sngRecordNumber)
            Progress.Value = 10
            SplitFile((txtSource1.Text), strTemp2, sngRecordNumber + 1, sngFileSize - sngRecordNumber + 1)
            Progress.Value = 20
            SplitFile((txtDir1.Text), strTemp3, 1, FileLen(txtDir1.Text))
            Progress.Value = 30
            Kill(txtSource1.Text)
            System.Windows.Forms.Application.DoEvents()
            ProgAdd = 70 \ (UBound(strTemp1) + UBound(strTemp2) + UBound(strTemp3))
            For i = 1 To UBound(strTemp1)

                PutToFile((txtSource1.Text), (i - 1) * Chunk + 1, strTemp1(i))
                Progress.Value = Progress.Value + ProgAdd
                System.Windows.Forms.Application.DoEvents()
            Next i
            sngFileSize = FileLen(txtSource1.Text)
            For i = 1 To UBound(strTemp3)
                PutToFile((txtSource1.Text), (i - 1) * Chunk + sngFileSize, strTemp3(i))
                Progress.Value = Progress.Value + ProgAdd
                System.Windows.Forms.Application.DoEvents()
            Next i
            sngFileSize = FileLen(txtSource1.Text)
            For i = 1 To UBound(strTemp2)
                PutToFile((txtSource1.Text), (i - 1) * Chunk + sngFileSize, strTemp2(i))
                Progress.Value = Progress.Value + ProgAdd
                System.Windows.Forms.Application.DoEvents()
            Next i
            MsgBox("File has been attached as a middling successfully", MsgBoxStyle.Information, "Finish")
            Progress.Value = 0
        End If
        Exit Sub
Err_Handle:
        If Err.Number = 75 Then
            MsgBox("File already in use.Close it and try again.", MsgBoxStyle.Critical, "Error")
        Else
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
        End If
        Progress.Value = 0
    End Sub
    Private Sub fltRemove_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMyControls.__FlatButton_MouseDownEvent) Handles fltRemove.MouseDownEvent
        Dim k As New frmShow
        

        k.ShowDialog()
    End Sub
    Private Sub fltCutAndSave_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fltCutAndSave.ClickEvent
        If Pos.ClipMarkVisible = False Then
            MsgBox("Please specify cut mark", MsgBoxStyle.Critical, "Error")
            Exit Sub
        ElseIf MCI.FileName = "" Then
            MsgBox("Please play your mp3 first", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
        SaveFileDialog1.ShowDialog()
        Progress.Value = 0
        If SaveFileDialog1.FileName <> "" Then Call CutMp3(Mid(MCI.FileName, 2, Len(MCI.FileName) - 2), (SaveFileDialog1.FileName))
        Exit Sub

    End Sub
	Private Sub fltTab_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fltTab.ClickEvent
        Dim Index As Short = fltTab.GetIndex(eventSender)
        Dim i As Short
        For i = 0 To fltTab.Count - 1
            If fltTab(i).Top = 144 Then
                With fltTab(i)
                    .SetBounds(.Left, 152, .Width, 17)
                    ShowTabContents(i, False)
                    Exit For
                End With
            End If
            System.Windows.Forms.Application.DoEvents()
        Next i

        With fltTab(Index)
            .SetBounds(.Left, 144, .Width, 25)
        End With

        Progress.Visible = True

        ShowTabContents(Index, True)
        fltTab(0).Tag = Index

    End Sub

	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		Call Init()

	End Sub

	Private Sub Forward_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Forward.ClickEvent

		With lstPlayList
			If .Items.Count = 0 Then Exit Sub
			If .SelectedIndex <> .Items.Count - 1 Then
				.SetSelected(.SelectedIndex, False)
				.SelectedIndex = .SelectedIndex + 1
				.SetSelected(.SelectedIndex, True)
			Else
				.SetSelected(.SelectedIndex, False)
				.SelectedIndex = 0 'First item
				.SetSelected(.SelectedIndex, True)
			End If
		End With
		Play_ClickEvent(Play, New System.EventArgs())

	End Sub

	Private Sub imgBar_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgBar.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)

		If Button = VB6.MouseButtonConstants.LeftButton Then

			ReleaseCapture()
			SendMessage(Me.Handle.ToInt32, WM_NCLBUTTONDOWN, HTCAPTION, 0)

		End If

	End Sub
	
    Private Sub lstPlayList_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstPlayList.DoubleClick
        MCI.CloseMCI()
        Call Play_ClickEvent(Play, New System.EventArgs())

    End Sub

	Private Sub Minimized_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Minimized.ClickEvent

		WindowState = System.Windows.Forms.FormWindowState.Minimized

	End Sub

	Private Sub MusicTime_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMyControls.__CustomControl_MouseDownEvent) Handles MusicTime.MouseDownEvent
		Dim Index As Short = MusicTime.GetIndex(eventSender)

		If eventArgs.Button = VB6.MouseButtonConstants.LeftButton Then
			TimeSetting()
		Else

            frmMenu.mnuTime.ShowDropDown()

		End If

	End Sub

	Private Sub Negative_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMyControls.__CustomControl_MouseDownEvent) Handles Negative.MouseDownEvent

		If eventArgs.Button = VB6.MouseButtonConstants.LeftButton Then
			TimeSetting()
		Else

            frmMenu.mnuTime.ShowDropDown()

		End If

	End Sub

	Private Sub OpenFiles_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OpenFiles.ClickEvent
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.ShowDialog()
        Dim s() As String = OpenFileDialog1.FileNames()
        Dim i As Integer
        For i = 0 To s.Length - 1
            lstPlayListPath.Items.Add(s(i))
            lstPlayList.Items.Add(System.IO.Path.GetFileNameWithoutExtension(s(i)))

        Next
	End Sub
    Private Sub optClipMark_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optClipMark.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optClipMark.GetIndex(eventSender)
            Pos.StartClipMark = Not (CBool(Index))
        End If
    End Sub
    Private Sub optMerge_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMerge.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optMerge.GetIndex(eventSender)
            '------------------
            If Index = 2 Then
                _txtTime_0.Enabled = True
                _txtTime_1.Enabled = True

                _Label2_0.Enabled = True
                _Label2_1.Enabled = True

            Else
                _Label2_0.Enabled = False
                _Label2_1.Enabled = False

            End If
            '------------------
        End If
    End Sub

	Private Sub Pause_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Pause.ClickEvent

		MCI.Pause()
		Timer1.Enabled = False

	End Sub

	Private Sub Play_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Play.ClickEvent

		Dim LI As Integer
		Dim TotalTime As String

		If lstPlayListPath.Items.Count = 0 Then
			Timer1.Enabled = False
			MCI.CloseMCI()
			Pos.Length = 0
			Pos.SeekTo(0)
			OpenFiles_ClickEvent(OpenFiles, New System.EventArgs())
			Exit Sub
		End If

		If LCase(VB.Left(MCI.Mode, 6)) = "paused" Then
			MCI.ResumeMCI()
			Timer1.Enabled = True

		ElseIf LCase(VB.Left(MCI.Mode, 7)) = "stopped" Then 
			Pos.ClipMarkVisible = False
			optClipMark(0).Checked = False
			optClipMark(1).Checked = False
			MCI.Play(Int(Pos.Value) * 1000)
			Timer1.Enabled = True
		Else

			With MCI
				.CloseMCI()

				If lstPlayList.SelectedIndex >= 0 Then
					LI = lstPlayList.SelectedIndex
				Else
					LI = 0
				End If
				'------------
				.FileName = VB6.GetItemString(lstPlayListPath, LI)
				.OpenMCI()
				Pos.Length = Val(MCI.Length) \ 1000
				Pos.SeekTo(0)
				Pos.ClipMarkVisible = False
				optClipMark(0).Checked = False
				optClipMark(1).Checked = False
				.Play()
				TotalTime = MilliSecToMin(Val(MCI.Length))
				'--------------
				If VB.Right(VB6.GetItemString(lstPlayList, LI), 10) <> " - [" & VB.Left(TotalTime, 2) & ":" & VB.Right(TotalTime, 2) & "]" Then
					VB6.SetItemString(lstPlayList, LI, VB6.GetItemString(lstPlayList, LI) & " - [" & VB.Left(TotalTime, 2) & ":" & VB.Right(TotalTime, 2) & "]")
				End If
				'----------
				Timer1.Enabled = True
				'--------------
			End With
		End If
		'----------------
	End Sub

	Private Sub Pos_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Pos.Scroll

		If LCase(VB.Left(MCI.Mode, 7)) = "stopped" Or LCase(VB.Left(MCI.Mode, 6)) = "paused" Then

			MCI.SeekTo(Int(Pos.Value) * 1000)

		Else
            MCI.Play(Int(Pos.Value) * 1000)

		End If
    End Sub

	Private Sub BlankSpace_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BlankSpace.ClickEvent
        TimeSetting()
    End Sub

	Private Sub Rewind_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Rewind.ClickEvent
		'-----------------
		With lstPlayList
			If .Items.Count = 0 Then Exit Sub
			If .SelectedIndex <> 0 Then
				.SetSelected(.SelectedIndex, False)
				.SelectedIndex = .SelectedIndex - 1
				.SetSelected(.SelectedIndex, True)
			Else
				.SetSelected(.SelectedIndex, False)
				.SelectedIndex = .Items.Count - 1
				.SetSelected(.SelectedIndex, True)
			End If
		End With
		Play_ClickEvent(Play, New System.EventArgs())

	End Sub

	Private Sub Stop_Renamed_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Stop_Renamed.ClickEvent

		Timer1.Enabled = False
		MCI.StopMCI()
		Pos.SeekTo(0)
		ShowTime("0000")

	End Sub

	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick

		Pos.SeekTo(Val(MCI.Position) \ 1000)
		'-------------
		If TimeElapsed = True Then
			ShowTime(MilliSecToMin(Val(MCI.Position)))
		Else
			ShowTime(MilliSecToMin(Val(MCI.Length) - Val(MCI.Position)))
		End If
		'-------------
		If Val(MCI.Position) = Val(MCI.Length) Then Stop_Renamed_ClickEvent(Stop_Renamed, New System.EventArgs())
		'---------------
	End Sub
	
    Private Sub Volume_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Volume.Scroll
        '--------------
        MCI.SetVolume(Int(Volume.Value))
        '-------------
    End Sub
	 
    Private Sub fltRemove_ClickEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fltRemove.ClickEvent

    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub
End Class