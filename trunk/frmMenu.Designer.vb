<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMenu
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents mnuSelected As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuABS As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAll As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRPL As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTimeElapsed As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTimeRemaining As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTime As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuRPL = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuSelected = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuABS = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAll = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTime = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTimeElapsed = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTimeRemaining = New System.Windows.Forms.ToolStripMenuItem
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRPL, Me.mnuTime})
        Me.MainMenu1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.MainMenu1.Size = New System.Drawing.Size(312, 4)
        Me.MainMenu1.TabIndex = 0
        '
        'mnuRPL
        '
        Me.mnuRPL.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuSelected, Me.mnuABS, Me.mnuAll})
        Me.mnuRPL.Name = "mnuRPL"
        Me.mnuRPL.Size = New System.Drawing.Size(94, 17)
        Me.mnuRPL.Text = "RemovePlayList"
        Me.mnuRPL.Visible = False
        '
        'mnuSelected
        '
        Me.mnuSelected.Name = "mnuSelected"
        Me.mnuSelected.Size = New System.Drawing.Size(158, 22)
        Me.mnuSelected.Text = "Selected"
        '
        'mnuABS
        '
        Me.mnuABS.Name = "mnuABS"
        Me.mnuABS.Size = New System.Drawing.Size(158, 22)
        Me.mnuABS.Text = "All but selected"
        '
        'mnuAll
        '
        Me.mnuAll.Name = "mnuAll"
        Me.mnuAll.Size = New System.Drawing.Size(158, 22)
        Me.mnuAll.Text = "All"
        '
        'mnuTime
        '
        Me.mnuTime.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTimeElapsed, Me.mnuTimeRemaining})
        Me.mnuTime.Name = "mnuTime"
        Me.mnuTime.Size = New System.Drawing.Size(67, 17)
        Me.mnuTime.Text = "TimeMenu"
        Me.mnuTime.Visible = False
        '
        'mnuTimeElapsed
        '
        Me.mnuTimeElapsed.Name = "mnuTimeElapsed"
        Me.mnuTimeElapsed.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
        Me.mnuTimeElapsed.Size = New System.Drawing.Size(195, 22)
        Me.mnuTimeElapsed.Text = "Time elapsed"
        '
        'mnuTimeRemaining
        '
        Me.mnuTimeRemaining.Name = "mnuTimeRemaining"
        Me.mnuTimeRemaining.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.mnuTimeRemaining.Size = New System.Drawing.Size(195, 22)
        Me.mnuTimeRemaining.Text = "Time remaining"
        '
        'frmMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(312, 206)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmMenu"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Form2"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class