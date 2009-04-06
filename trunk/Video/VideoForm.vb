Public Class VideoForm
    Inherits System.Windows.Forms.Form


    Dim mCls As csMovieLibrary = New csMovieLibrary()

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents gbScreen As System.Windows.Forms.GroupBox
    Friend WithEvents gbCommands As System.Windows.Forms.GroupBox
    Friend WithEvents btnPlay As System.Windows.Forms.Button
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnPause As System.Windows.Forms.Button
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents gbStats As System.Windows.Forms.GroupBox
    Friend WithEvents lblPframes As System.Windows.Forms.Label
    Friend WithEvents lblPms As System.Windows.Forms.Label
    Friend WithEvents lblPsec As System.Windows.Forms.Label
    Friend WithEvents lblDframes As System.Windows.Forms.Label
    Friend WithEvents lblDms As System.Windows.Forms.Label
    Friend WithEvents lblDsec As System.Windows.Forms.Label
    Friend WithEvents t As System.Windows.Forms.Timer
    Friend WithEvents lblfps As System.Windows.Forms.Label
    Friend WithEvents lblnfr As System.Windows.Forms.Label
    Friend WithEvents lbldeviceName As System.Windows.Forms.Label
    Friend WithEvents lblDeviceVersion As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblStat As System.Windows.Forms.Label
    Friend WithEvents lblPStat As System.Windows.Forms.Label
    Friend WithEvents lblCSize As System.Windows.Forms.Label
    Friend WithEvents lblFPos As System.Windows.Forms.Label
    Friend WithEvents lblDF As System.Windows.Forms.Label
    Friend WithEvents btnRewind As System.Windows.Forms.Button
    Friend WithEvents btnForward As System.Windows.Forms.Button
    Friend WithEvents lblVStat As System.Windows.Forms.Label
    Friend WithEvents chkHideMovie As System.Windows.Forms.CheckBox
    Friend WithEvents chkMute As System.Windows.Forms.CheckBox
    Friend WithEvents lblMInput As System.Windows.Forms.Label
    Friend WithEvents lblMOutput As System.Windows.Forms.Label
    Friend WithEvents lblPRate As System.Windows.Forms.Label
    Friend WithEvents hsRate As System.Windows.Forms.HScrollBar
    Friend WithEvents lblDReady As System.Windows.Forms.Label
    Friend WithEvents chkFull As System.Windows.Forms.CheckBox
    Friend WithEvents btnRestartMovie As System.Windows.Forms.Button
    Friend WithEvents hsPosition As System.Windows.Forms.HScrollBar
    Friend WithEvents lblVolBar As System.Windows.Forms.Label
    Friend WithEvents lblSpeed As System.Windows.Forms.Label
    Friend WithEvents lblBPP As System.Windows.Forms.Label
    Friend WithEvents hsVol As System.Windows.Forms.HScrollBar
    Friend WithEvents lblFilename As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.gbScreen = New System.Windows.Forms.GroupBox
        Me.gbCommands = New System.Windows.Forms.GroupBox
        Me.lblFilename = New System.Windows.Forms.Label
        Me.btnRestartMovie = New System.Windows.Forms.Button
        Me.chkFull = New System.Windows.Forms.CheckBox
        Me.chkMute = New System.Windows.Forms.CheckBox
        Me.chkHideMovie = New System.Windows.Forms.CheckBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOpen = New System.Windows.Forms.Button
        Me.btnPause = New System.Windows.Forms.Button
        Me.btnStop = New System.Windows.Forms.Button
        Me.btnPlay = New System.Windows.Forms.Button
        Me.btnRewind = New System.Windows.Forms.Button
        Me.btnForward = New System.Windows.Forms.Button
        Me.lblCSize = New System.Windows.Forms.Label
        Me.lblPStat = New System.Windows.Forms.Label
        Me.gbStats = New System.Windows.Forms.GroupBox
        Me.lblVolBar = New System.Windows.Forms.Label
        Me.lblBPP = New System.Windows.Forms.Label
        Me.hsPosition = New System.Windows.Forms.HScrollBar
        Me.hsVol = New System.Windows.Forms.HScrollBar
        Me.lblPRate = New System.Windows.Forms.Label
        Me.lblDF = New System.Windows.Forms.Label
        Me.lblFPos = New System.Windows.Forms.Label
        Me.lblnfr = New System.Windows.Forms.Label
        Me.lblfps = New System.Windows.Forms.Label
        Me.lblDsec = New System.Windows.Forms.Label
        Me.lblDms = New System.Windows.Forms.Label
        Me.lblDframes = New System.Windows.Forms.Label
        Me.lblPsec = New System.Windows.Forms.Label
        Me.lblPms = New System.Windows.Forms.Label
        Me.lblPframes = New System.Windows.Forms.Label
        Me.hsRate = New System.Windows.Forms.HScrollBar
        Me.lblSpeed = New System.Windows.Forms.Label
        Me.t = New System.Windows.Forms.Timer(Me.components)
        Me.lbldeviceName = New System.Windows.Forms.Label
        Me.lblDeviceVersion = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblDReady = New System.Windows.Forms.Label
        Me.lblMOutput = New System.Windows.Forms.Label
        Me.lblMInput = New System.Windows.Forms.Label
        Me.lblVStat = New System.Windows.Forms.Label
        Me.lblStat = New System.Windows.Forms.Label
        Me.gbCommands.SuspendLayout()
        Me.gbStats.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbScreen
        '
        Me.gbScreen.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.gbScreen.Location = New System.Drawing.Point(8, 199)
        Me.gbScreen.Name = "gbScreen"
        Me.gbScreen.Size = New System.Drawing.Size(560, 249)
        Me.gbScreen.TabIndex = 0
        Me.gbScreen.TabStop = False
        '
        'gbCommands
        '
        Me.gbCommands.Controls.Add(Me.lblFilename)
        Me.gbCommands.Controls.Add(Me.btnRestartMovie)
        Me.gbCommands.Controls.Add(Me.chkFull)
        Me.gbCommands.Controls.Add(Me.chkMute)
        Me.gbCommands.Controls.Add(Me.chkHideMovie)
        Me.gbCommands.Controls.Add(Me.btnClose)
        Me.gbCommands.Controls.Add(Me.btnOpen)
        Me.gbCommands.Controls.Add(Me.btnPause)
        Me.gbCommands.Controls.Add(Me.btnStop)
        Me.gbCommands.Controls.Add(Me.btnPlay)
        Me.gbCommands.Controls.Add(Me.btnRewind)
        Me.gbCommands.Controls.Add(Me.btnForward)
        Me.gbCommands.Controls.Add(Me.lblCSize)
        Me.gbCommands.Controls.Add(Me.lblPStat)
        Me.gbCommands.Location = New System.Drawing.Point(8, 0)
        Me.gbCommands.Name = "gbCommands"
        Me.gbCommands.Size = New System.Drawing.Size(296, 123)
        Me.gbCommands.TabIndex = 1
        Me.gbCommands.TabStop = False
        '
        'lblFilename
        '
        Me.lblFilename.Location = New System.Drawing.Point(8, 104)
        Me.lblFilename.Name = "lblFilename"
        Me.lblFilename.Size = New System.Drawing.Size(280, 16)
        Me.lblFilename.TabIndex = 16
        Me.lblFilename.Text = "Filename: "
        '
        'btnRestartMovie
        '
        Me.btnRestartMovie.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRestartMovie.Location = New System.Drawing.Point(152, 40)
        Me.btnRestartMovie.Name = "btnRestartMovie"
        Me.btnRestartMovie.Size = New System.Drawing.Size(64, 23)
        Me.btnRestartMovie.TabIndex = 15
        Me.btnRestartMovie.Text = "Restart"
        '
        'chkFull
        '
        Me.chkFull.Location = New System.Drawing.Point(192, 69)
        Me.chkFull.Name = "chkFull"
        Me.chkFull.Size = New System.Drawing.Size(88, 16)
        Me.chkFull.TabIndex = 14
        Me.chkFull.Text = "Full Screen"
        '
        'chkMute
        '
        Me.chkMute.Location = New System.Drawing.Point(16, 69)
        Me.chkMute.Name = "chkMute"
        Me.chkMute.Size = New System.Drawing.Size(88, 16)
        Me.chkMute.TabIndex = 13
        Me.chkMute.Text = "Mute Sound"
        '
        'chkHideMovie
        '
        Me.chkHideMovie.Location = New System.Drawing.Point(104, 69)
        Me.chkHideMovie.Name = "chkHideMovie"
        Me.chkHideMovie.Size = New System.Drawing.Size(88, 16)
        Me.chkHideMovie.TabIndex = 12
        Me.chkHideMovie.Text = "Hide Movie"
        '
        'btnClose
        '
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Location = New System.Drawing.Point(224, 40)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 23)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        '
        'btnOpen
        '
        Me.btnOpen.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOpen.Location = New System.Drawing.Point(224, 16)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(64, 23)
        Me.btnOpen.TabIndex = 5
        Me.btnOpen.Text = "Open"
        '
        'btnPause
        '
        Me.btnPause.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPause.Location = New System.Drawing.Point(152, 16)
        Me.btnPause.Name = "btnPause"
        Me.btnPause.Size = New System.Drawing.Size(64, 23)
        Me.btnPause.TabIndex = 4
        Me.btnPause.Text = "Pause"
        '
        'btnStop
        '
        Me.btnStop.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnStop.Location = New System.Drawing.Point(80, 16)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(64, 23)
        Me.btnStop.TabIndex = 3
        Me.btnStop.Text = "Stop"
        '
        'btnPlay
        '
        Me.btnPlay.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPlay.Location = New System.Drawing.Point(8, 16)
        Me.btnPlay.Name = "btnPlay"
        Me.btnPlay.Size = New System.Drawing.Size(64, 23)
        Me.btnPlay.TabIndex = 2
        Me.btnPlay.Text = "Play"
        '
        'btnRewind
        '
        Me.btnRewind.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRewind.Location = New System.Drawing.Point(8, 40)
        Me.btnRewind.Name = "btnRewind"
        Me.btnRewind.Size = New System.Drawing.Size(64, 23)
        Me.btnRewind.TabIndex = 10
        Me.btnRewind.Text = "Rewind"
        '
        'btnForward
        '
        Me.btnForward.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnForward.Location = New System.Drawing.Point(80, 40)
        Me.btnForward.Name = "btnForward"
        Me.btnForward.Size = New System.Drawing.Size(64, 23)
        Me.btnForward.TabIndex = 11
        Me.btnForward.Text = "Forward"
        '
        'lblCSize
        '
        Me.lblCSize.Location = New System.Drawing.Point(128, 87)
        Me.lblCSize.Name = "lblCSize"
        Me.lblCSize.Size = New System.Drawing.Size(160, 16)
        Me.lblCSize.TabIndex = 8
        Me.lblCSize.Text = "Current Size Loc: "
        '
        'lblPStat
        '
        Me.lblPStat.Location = New System.Drawing.Point(8, 87)
        Me.lblPStat.Name = "lblPStat"
        Me.lblPStat.Size = New System.Drawing.Size(128, 16)
        Me.lblPStat.TabIndex = 10
        Me.lblPStat.Text = "Playing Status: stopped "
        '
        'gbStats
        '
        Me.gbStats.Controls.Add(Me.lblVolBar)
        Me.gbStats.Controls.Add(Me.lblBPP)
        Me.gbStats.Controls.Add(Me.hsPosition)
        Me.gbStats.Controls.Add(Me.hsVol)
        Me.gbStats.Controls.Add(Me.lblPRate)
        Me.gbStats.Controls.Add(Me.lblDF)
        Me.gbStats.Controls.Add(Me.lblFPos)
        Me.gbStats.Controls.Add(Me.lblnfr)
        Me.gbStats.Controls.Add(Me.lblfps)
        Me.gbStats.Controls.Add(Me.lblDsec)
        Me.gbStats.Controls.Add(Me.lblDms)
        Me.gbStats.Controls.Add(Me.lblDframes)
        Me.gbStats.Controls.Add(Me.lblPsec)
        Me.gbStats.Controls.Add(Me.lblPms)
        Me.gbStats.Controls.Add(Me.lblPframes)
        Me.gbStats.Controls.Add(Me.hsRate)
        Me.gbStats.Controls.Add(Me.lblSpeed)
        Me.gbStats.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.gbStats.Location = New System.Drawing.Point(8, 123)
        Me.gbStats.Name = "gbStats"
        Me.gbStats.Size = New System.Drawing.Size(560, 77)
        Me.gbStats.TabIndex = 2
        Me.gbStats.TabStop = False
        '
        'lblVolBar
        '
        Me.lblVolBar.Location = New System.Drawing.Point(240, 16)
        Me.lblVolBar.Name = "lblVolBar"
        Me.lblVolBar.Size = New System.Drawing.Size(100, 16)
        Me.lblVolBar.TabIndex = 18
        Me.lblVolBar.Text = "Volume Bar"
        '
        'lblBPP
        '
        Me.lblBPP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblBPP.Location = New System.Drawing.Point(160, 56)
        Me.lblBPP.Name = "lblBPP"
        Me.lblBPP.Size = New System.Drawing.Size(72, 16)
        Me.lblBPP.TabIndex = 17
        Me.lblBPP.Text = "BPP: "
        '
        'hsPosition
        '
        Me.hsPosition.LargeChange = 1
        Me.hsPosition.Location = New System.Drawing.Point(240, 56)
        Me.hsPosition.Maximum = 0
        Me.hsPosition.Name = "hsPosition"
        Me.hsPosition.Size = New System.Drawing.Size(312, 16)
        Me.hsPosition.TabIndex = 16
        '
        'hsVol
        '
        Me.hsVol.Location = New System.Drawing.Point(240, 32)
        Me.hsVol.Maximum = 1000
        Me.hsVol.Name = "hsVol"
        Me.hsVol.Size = New System.Drawing.Size(144, 16)
        Me.hsVol.TabIndex = 15
        Me.hsVol.Value = 800
        '
        'lblPRate
        '
        Me.lblPRate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPRate.Location = New System.Drawing.Point(160, 40)
        Me.lblPRate.Name = "lblPRate"
        Me.lblPRate.Size = New System.Drawing.Size(72, 16)
        Me.lblPRate.TabIndex = 10
        Me.lblPRate.Text = "Rate: "
        '
        'lblDF
        '
        Me.lblDF.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDF.Location = New System.Drawing.Point(80, 56)
        Me.lblDF.Name = "lblDF"
        Me.lblDF.Size = New System.Drawing.Size(80, 16)
        Me.lblDF.TabIndex = 9
        Me.lblDF.Text = "Length: "
        '
        'lblFPos
        '
        Me.lblFPos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFPos.Location = New System.Drawing.Point(8, 56)
        Me.lblFPos.Name = "lblFPos"
        Me.lblFPos.Size = New System.Drawing.Size(72, 16)
        Me.lblFPos.TabIndex = 8
        Me.lblFPos.Text = "Pos: "
        '
        'lblnfr
        '
        Me.lblnfr.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblnfr.Location = New System.Drawing.Point(160, 24)
        Me.lblnfr.Name = "lblnfr"
        Me.lblnfr.Size = New System.Drawing.Size(72, 16)
        Me.lblnfr.TabIndex = 7
        Me.lblnfr.Text = "NFPS"
        '
        'lblfps
        '
        Me.lblfps.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblfps.Location = New System.Drawing.Point(160, 8)
        Me.lblfps.Name = "lblfps"
        Me.lblfps.Size = New System.Drawing.Size(72, 16)
        Me.lblfps.TabIndex = 6
        Me.lblfps.Text = "FPS: "
        '
        'lblDsec
        '
        Me.lblDsec.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDsec.Location = New System.Drawing.Point(80, 40)
        Me.lblDsec.Name = "lblDsec"
        Me.lblDsec.Size = New System.Drawing.Size(80, 16)
        Me.lblDsec.TabIndex = 5
        Me.lblDsec.Text = "Seconds: "
        '
        'lblDms
        '
        Me.lblDms.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDms.Location = New System.Drawing.Point(80, 24)
        Me.lblDms.Name = "lblDms"
        Me.lblDms.Size = New System.Drawing.Size(80, 16)
        Me.lblDms.TabIndex = 4
        Me.lblDms.Text = "Milli-Secs"
        '
        'lblDframes
        '
        Me.lblDframes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDframes.Location = New System.Drawing.Point(80, 8)
        Me.lblDframes.Name = "lblDframes"
        Me.lblDframes.Size = New System.Drawing.Size(80, 16)
        Me.lblDframes.TabIndex = 3
        Me.lblDframes.Text = "Frames: "
        '
        'lblPsec
        '
        Me.lblPsec.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPsec.Location = New System.Drawing.Point(8, 40)
        Me.lblPsec.Name = "lblPsec"
        Me.lblPsec.Size = New System.Drawing.Size(72, 16)
        Me.lblPsec.TabIndex = 2
        Me.lblPsec.Text = "Second: "
        '
        'lblPms
        '
        Me.lblPms.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPms.Location = New System.Drawing.Point(8, 24)
        Me.lblPms.Name = "lblPms"
        Me.lblPms.Size = New System.Drawing.Size(72, 16)
        Me.lblPms.TabIndex = 1
        Me.lblPms.Text = "Milli-Sec: "
        '
        'lblPframes
        '
        Me.lblPframes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPframes.Location = New System.Drawing.Point(8, 8)
        Me.lblPframes.Name = "lblPframes"
        Me.lblPframes.Size = New System.Drawing.Size(72, 16)
        Me.lblPframes.TabIndex = 0
        Me.lblPframes.Text = "Frame: "
        '
        'hsRate
        '
        Me.hsRate.Location = New System.Drawing.Point(392, 32)
        Me.hsRate.Maximum = 2000
        Me.hsRate.Name = "hsRate"
        Me.hsRate.Size = New System.Drawing.Size(160, 16)
        Me.hsRate.TabIndex = 14
        Me.hsRate.Value = 1000
        '
        'lblSpeed
        '
        Me.lblSpeed.Location = New System.Drawing.Point(392, 16)
        Me.lblSpeed.Name = "lblSpeed"
        Me.lblSpeed.Size = New System.Drawing.Size(100, 16)
        Me.lblSpeed.TabIndex = 4
        Me.lblSpeed.Text = "Speed Bar"
        '
        't
        '
        Me.t.Interval = 1000
        '
        'lbldeviceName
        '
        Me.lbldeviceName.Location = New System.Drawing.Point(8, 72)
        Me.lbldeviceName.Name = "lbldeviceName"
        Me.lbldeviceName.Size = New System.Drawing.Size(152, 16)
        Me.lbldeviceName.TabIndex = 8
        Me.lbldeviceName.Text = "Device Name: "
        '
        'lblDeviceVersion
        '
        Me.lblDeviceVersion.Location = New System.Drawing.Point(8, 88)
        Me.lblDeviceVersion.Name = "lblDeviceVersion"
        Me.lblDeviceVersion.Size = New System.Drawing.Size(152, 16)
        Me.lblDeviceVersion.TabIndex = 9
        Me.lblDeviceVersion.Text = "Device Version: "
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblDReady)
        Me.GroupBox1.Controls.Add(Me.lblMOutput)
        Me.GroupBox1.Controls.Add(Me.lblMInput)
        Me.GroupBox1.Controls.Add(Me.lblVStat)
        Me.GroupBox1.Controls.Add(Me.lblStat)
        Me.GroupBox1.Controls.Add(Me.lblDeviceVersion)
        Me.GroupBox1.Controls.Add(Me.lbldeviceName)
        Me.GroupBox1.Location = New System.Drawing.Point(304, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(264, 123)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'lblDReady
        '
        Me.lblDReady.Location = New System.Drawing.Point(8, 104)
        Me.lblDReady.Name = "lblDReady"
        Me.lblDReady.Size = New System.Drawing.Size(152, 16)
        Me.lblDReady.TabIndex = 14
        Me.lblDReady.Text = "Device Ready: "
        '
        'lblMOutput
        '
        Me.lblMOutput.Location = New System.Drawing.Point(160, 88)
        Me.lblMOutput.Name = "lblMOutput"
        Me.lblMOutput.Size = New System.Drawing.Size(96, 16)
        Me.lblMOutput.TabIndex = 13
        Me.lblMOutput.Text = "Movie Output: "
        '
        'lblMInput
        '
        Me.lblMInput.Location = New System.Drawing.Point(160, 72)
        Me.lblMInput.Name = "lblMInput"
        Me.lblMInput.Size = New System.Drawing.Size(96, 16)
        Me.lblMInput.TabIndex = 12
        Me.lblMInput.Text = "Movie Input: "
        '
        'lblVStat
        '
        Me.lblVStat.Location = New System.Drawing.Point(160, 104)
        Me.lblVStat.Name = "lblVStat"
        Me.lblVStat.Size = New System.Drawing.Size(96, 16)
        Me.lblVStat.TabIndex = 11
        Me.lblVStat.Text = "Video Status: on"
        '
        'lblStat
        '
        Me.lblStat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblStat.Location = New System.Drawing.Point(8, 16)
        Me.lblStat.Name = "lblStat"
        Me.lblStat.Size = New System.Drawing.Size(248, 56)
        Me.lblStat.TabIndex = 3
        Me.lblStat.Text = "MCI Status: "
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(578, 458)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gbStats)
        Me.Controls.Add(Me.gbCommands)
        Me.Controls.Add(Me.gbScreen)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "M Cutter"
        Me.gbCommands.ResumeLayout(False)
        Me.gbStats.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Dim openDLG As OpenFileDialog = New OpenFileDialog()
        Dim dlgResult As DialogResult

        openDLG.Filter = "Movie Files (*.mp3, *.avi, *.mov, *.mpg, *.mpeg, *.wmv)|*.mp3;*.avi;*.mov;*.mpg;*.mpeg;*.wmv"
        openDLG.Title = "Please select the movie file you want to play."

        dlgResult = openDLG.ShowDialog()

        If dlgResult = DialogResult.Cancel Then Exit Sub

        mCls.filename = openDLG.FileName
        lblFilename.Text = "Filename: " & openDLG.FileName
        mCls.openMovieInWindow(gbScreen.Handle.ToInt32)

        'display various stats and features 
        lblStat.Text = "MCI Status: " & mCls.returnCommandInterfaceStatus
        lblDReady.Text = "Device Ready: " & mCls.isDeviceReady
        lbldeviceName.Text = "Device Name: " & mCls.deviceName
        lblDeviceVersion.Text = "Device Version: " & mCls.deviceVersion
        lblMInput.Text = "Movie Input: " & mCls.movieInputSource
        lblMOutput.Text = "Movie Output: " & mCls.movieOutputSource

        If mCls.bitsPerPixel = 0 Then

            lblBPP.Text = "BPP: n/a"

        Else
            lblBPP.Text = "BPP: " & mCls.bitsPerPixel

        End If

        openDLG.Dispose()

        mCls.sizeLocateMovie(10, 15, gbScreen.Width - 20, gbScreen.Height - 20)

        lblCSize.Text = "Current Size Loc: " & mCls.getCurrentSizeLocation
    End Sub

    Private Sub btnPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlay.Click
        mCls.playMovie()

        lblStat.Text = "MCI Status: " & mCls.returnCommandInterfaceStatus
        lblDframes.Text = "Frames: " & mCls.durationInFrames
        lblDms.Text = "Milli: " & mCls.durationInMS
        lblDsec.Text = "Sec: " & mCls.durationInSec
        lblDF.Text = "Length: " & mCls.formatDuration

        lblfps.Text = "FPS: " & mCls.framePerSecondRate
        lblnfr.Text = "NFPS: " & mCls.nominalFrameRate
        lblPRate.Text = "Rate: " & mCls.moviePlayRate

        lblMInput.Text = "Input Source: " & mCls.movieInputSource
        lblMOutput.Text = "Output Source: " & mCls.movieOutputSource

        lblVStat.Text = "Video Status: " & mCls.getVideoStatus
        lblPStat.Text = "Playing Status: " & mCls.getPlayingStatus

        hsPosition.Maximum = mCls.durationInSec

        t.Enabled = True

    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        mCls.stopMovie()
        lblStat.Text = "MCI Status: " & mCls.returnCommandInterfaceStatus
        t.Enabled = False
        lblPStat.Text = "Playing Status: " & mCls.getPlayingStatus
    End Sub

    Private Sub btnPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPause.Click
        If btnPause.Text = "Pause" Then
            btnPause.Text = "Resume"
            mCls.pauseMovie()
            lblStat.Text = "MCI Status: " & mCls.returnCommandInterfaceStatus
            t.Enabled = False
            lblPStat.Text = "Playing Status: " & mCls.getPlayingStatus
        Else
            btnPause.Text = "Pause"
            mCls.resumeMovie()
            lblStat.Text = "MCI Status: " & mCls.returnCommandInterfaceStatus
            t.Enabled = True
            lblPStat.Text = "Playing Status: " & mCls.getPlayingStatus
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        mCls.closeMovie()
        lblStat.Text = "MCI Status: " & mCls.returnCommandInterfaceStatus
        t.Enabled = False
    End Sub

    Private Sub t_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t.Tick
        lblPframes.Text = "Frame: " & mCls.positionInFrames
        lblPms.Text = "MS: " & mCls.positionInMS
        lblPsec.Text = "Sec: " & mCls.positionInSec
        lblFPos.Text = "Pos: " & mCls.formatPosition

        hsPosition.Value = mCls.positionInSec

    End Sub

    Private Sub btnRewind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRewind.Click

        'rewind 10 seconds
        mCls.rewindBySeconds = 10

    End Sub

    Private Sub btnForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnForward.Click

        'fast forward 10 seconds
        mCls.forwardBySeconds = 10

    End Sub

    Private Sub chkHideMovie_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHideMovie.CheckedChanged

        'whether to hide or show the movie

        If chkHideMovie.CheckState = CheckState.Checked Then

            mCls.hideMovie = True

        Else

            mCls.hideMovie = False

        End If
    End Sub

    Private Sub chkMute_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMute.CheckedChanged

        If chkMute.CheckState = CheckState.Checked Then

            mCls.muteSoundOutput = True

        Else

            mCls.muteSoundOutput = False

        End If

    End Sub

    Private Sub hsRate_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hsRate.Scroll

        'change the play rate of the movie. some movie formats may not support
        'this feature

        mCls.moviePlayRate = hsRate.Value
        lblPRate.Text = "Rate: " & mCls.moviePlayRate

    End Sub

    Private Sub chkFull_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFull.CheckedChanged

        'whether to change to fullscreen or normal

        If chkFull.CheckState = CheckState.Checked Then

            mCls.playFullScreen()

        Else

            mCls.restoreOriginalMovieState()

        End If

    End Sub

    Private Sub btnRestartMovie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestartMovie.Click

        'restart the movie
        mCls.restartMovie()

    End Sub

    Private Sub hsPosition_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hsPosition.Scroll

        'change the playing position
        mCls.positionInSec = hsPosition.Value

    End Sub

    Private Sub hsVol_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hsVol.Scroll


        mCls.volumeLevel = hsVol.Value

    End Sub

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        'be sure to always call this command
        mCls.closeMovie()

    End Sub

    Private Sub gbScreen_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbScreen.Enter

    End Sub
End Class
