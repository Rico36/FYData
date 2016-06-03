<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.ProgressBarStyle = New System.Windows.Forms.StatusStrip()
        Me.lblStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pbarProgress = New System.Windows.Forms.ToolStripProgressBar()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.ProgressBarStyle.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Controls.Add(Me.btnImport)
        Me.Panel1.Location = New System.Drawing.Point(-1, 364)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(797, 89)
        Me.Panel1.TabIndex = 0
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(671, 4)
        Me.btnImport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(111, 40)
        Me.btnImport.TabIndex = 0
        Me.btnImport.Text = "Run"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'ProgressBarStyle
        '
        Me.ProgressBarStyle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBarStyle.AutoSize = False
        Me.ProgressBarStyle.Dock = System.Windows.Forms.DockStyle.None
        Me.ProgressBarStyle.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ProgressBarStyle.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatus, Me.pbarProgress})
        Me.ProgressBarStyle.Location = New System.Drawing.Point(-1, 415)
        Me.ProgressBarStyle.Name = "ProgressBarStyle"
        Me.ProgressBarStyle.Padding = New System.Windows.Forms.Padding(18, 0, 1, 0)
        Me.ProgressBarStyle.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ProgressBarStyle.Size = New System.Drawing.Size(797, 38)
        Me.ProgressBarStyle.SizingGrip = False
        Me.ProgressBarStyle.TabIndex = 1
        Me.ProgressBarStyle.Text = "StatusStrip1"
        '
        'lblStatus
        '
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 33)
        '
        'pbarProgress
        '
        Me.pbarProgress.AutoSize = False
        Me.pbarProgress.ForeColor = System.Drawing.Color.Lime
        Me.pbarProgress.Name = "pbarProgress"
        Me.pbarProgress.Size = New System.Drawing.Size(300, 32)
        Me.pbarProgress.Visible = False
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'txtOutput
        '
        Me.txtOutput.Location = New System.Drawing.Point(-1, 1)
        Me.txtOutput.Margin = New System.Windows.Forms.Padding(4)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(797, 359)
        Me.txtOutput.TabIndex = 1
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(798, 455)
        Me.Controls.Add(Me.ProgressBarStyle)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmMain"
        Me.Text = "FY2018 IT Assessment Data Processor"
        Me.Panel1.ResumeLayout(False)
        Me.ProgressBarStyle.ResumeLayout(False)
        Me.ProgressBarStyle.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents btnImport As Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBarStyle As Windows.Forms.StatusStrip
    Friend WithEvents pbarProgress As Windows.Forms.ToolStripProgressBar
    Friend WithEvents lblStatus As Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtOutput As Windows.Forms.TextBox
End Class
