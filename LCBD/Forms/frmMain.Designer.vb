<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.picHeader = New System.Windows.Forms.PictureBox()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtFileName = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSaveFolder = New System.Windows.Forms.Label()
        Me.btnOutputFolder = New System.Windows.Forms.Button()
        Me.btnExecute = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.lnk1 = New System.Windows.Forms.LinkLabel()
        Me.lnk2 = New System.Windows.Forms.LinkLabel()
        Me.picBottom = New System.Windows.Forms.PictureBox()
        Me.pnlWip = New System.Windows.Forms.Panel()
        Me.lblWip = New System.Windows.Forms.Label()
        Me.pgbWip = New System.Windows.Forms.ProgressBar()
        Me.btnStop = New System.Windows.Forms.Button()
        CType(Me.picHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picBottom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlWip.SuspendLayout()
        Me.SuspendLayout()
        '
        'picHeader
        '
        Me.picHeader.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.picHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.picHeader.Location = New System.Drawing.Point(0, 0)
        Me.picHeader.Name = "picHeader"
        Me.picHeader.Size = New System.Drawing.Size(834, 127)
        Me.picHeader.TabIndex = 0
        Me.picHeader.TabStop = False
        '
        'picLogo
        '
        Me.picLogo.BackgroundImage = Global.LCBD.My.Resources.Resources.logo
        Me.picLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.picLogo.Location = New System.Drawing.Point(12, 12)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(292, 101)
        Me.picLogo.TabIndex = 1
        Me.picLogo.TabStop = False
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(689, 163)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(79, 35)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Sfoglia"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtFileName
        '
        Me.txtFileName.BackColor = System.Drawing.SystemColors.Window
        Me.txtFileName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFileName.Location = New System.Drawing.Point(128, 163)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(555, 35)
        Me.txtFileName.TabIndex = 3
        Me.txtFileName.Text = "txtFileName"
        Me.txtFileName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(29, 163)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 35)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "File CSV"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(29, 208)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 35)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Scarica in"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtSaveFolder
        '
        Me.txtSaveFolder.BackColor = System.Drawing.SystemColors.Window
        Me.txtSaveFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSaveFolder.Location = New System.Drawing.Point(128, 208)
        Me.txtSaveFolder.Name = "txtSaveFolder"
        Me.txtSaveFolder.Size = New System.Drawing.Size(555, 35)
        Me.txtSaveFolder.TabIndex = 7
        Me.txtSaveFolder.Text = "txtSaveFolder"
        Me.txtSaveFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnOutputFolder
        '
        Me.btnOutputFolder.Location = New System.Drawing.Point(689, 208)
        Me.btnOutputFolder.Name = "btnOutputFolder"
        Me.btnOutputFolder.Size = New System.Drawing.Size(79, 35)
        Me.btnOutputFolder.TabIndex = 6
        Me.btnOutputFolder.Text = "Scegli"
        Me.btnOutputFolder.UseVisualStyleBackColor = True
        '
        'btnExecute
        '
        Me.btnExecute.Location = New System.Drawing.Point(353, 258)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.Size = New System.Drawing.Size(105, 51)
        Me.btnExecute.TabIndex = 9
        Me.btnExecute.Text = "Esegui"
        Me.btnExecute.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'lnk1
        '
        Me.lnk1.ActiveLinkColor = System.Drawing.Color.White
        Me.lnk1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lnk1.ForeColor = System.Drawing.Color.White
        Me.lnk1.LinkColor = System.Drawing.Color.White
        Me.lnk1.Location = New System.Drawing.Point(721, 12)
        Me.lnk1.Name = "lnk1"
        Me.lnk1.Size = New System.Drawing.Size(108, 25)
        Me.lnk1.TabIndex = 10
        Me.lnk1.TabStop = True
        Me.lnk1.Tag = "http://www.webartigiani.it/"
        Me.lnk1.Text = "WebArtigiani.it"
        Me.lnk1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lnk1.VisitedLinkColor = System.Drawing.Color.White
        '
        'lnk2
        '
        Me.lnk2.ActiveLinkColor = System.Drawing.Color.White
        Me.lnk2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lnk2.ForeColor = System.Drawing.Color.White
        Me.lnk2.LinkColor = System.Drawing.Color.White
        Me.lnk2.Location = New System.Drawing.Point(9, 522)
        Me.lnk2.Name = "lnk2"
        Me.lnk2.Size = New System.Drawing.Size(254, 25)
        Me.lnk2.TabIndex = 11
        Me.lnk2.TabStop = True
        Me.lnk2.Tag = "http://www.webartigiani.it/"
        Me.lnk2.Text = "Copyright ©  2020 WebArtigiani"
        Me.lnk2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lnk2.VisitedLinkColor = System.Drawing.Color.White
        '
        'picBottom
        '
        Me.picBottom.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.picBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.picBottom.Location = New System.Drawing.Point(0, 518)
        Me.picBottom.Name = "picBottom"
        Me.picBottom.Size = New System.Drawing.Size(834, 33)
        Me.picBottom.TabIndex = 12
        Me.picBottom.TabStop = False
        '
        'pnlWip
        '
        Me.pnlWip.Controls.Add(Me.lblWip)
        Me.pnlWip.Controls.Add(Me.pgbWip)
        Me.pnlWip.Location = New System.Drawing.Point(128, 333)
        Me.pnlWip.Name = "pnlWip"
        Me.pnlWip.Size = New System.Drawing.Size(555, 104)
        Me.pnlWip.TabIndex = 13
        '
        'lblWip
        '
        Me.lblWip.Location = New System.Drawing.Point(3, 45)
        Me.lblWip.Name = "lblWip"
        Me.lblWip.Size = New System.Drawing.Size(549, 44)
        Me.lblWip.TabIndex = 1
        Me.lblWip.Text = "lblWip"
        Me.lblWip.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pgbWip
        '
        Me.pgbWip.Location = New System.Drawing.Point(3, 12)
        Me.pgbWip.Name = "pgbWip"
        Me.pgbWip.Size = New System.Drawing.Size(549, 24)
        Me.pgbWip.TabIndex = 0
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(464, 258)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(105, 51)
        Me.btnStop.TabIndex = 14
        Me.btnStop.Text = "Interrompi"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(834, 551)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.pnlWip)
        Me.Controls.Add(Me.lnk2)
        Me.Controls.Add(Me.lnk1)
        Me.Controls.Add(Me.btnExecute)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSaveFolder)
        Me.Controls.Add(Me.btnOutputFolder)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.picHeader)
        Me.Controls.Add(Me.picBottom)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(850, 590)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(850, 590)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmMain"
        CType(Me.picHeader, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picBottom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlWip.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents picHeader As PictureBox
    Friend WithEvents picLogo As PictureBox
    Friend WithEvents btnBrowse As Button
    Friend WithEvents txtFileName As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtSaveFolder As Label
    Friend WithEvents btnOutputFolder As Button
    Friend WithEvents btnExecute As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents lnk1 As LinkLabel
    Friend WithEvents lnk2 As LinkLabel
    Friend WithEvents picBottom As PictureBox
    Friend WithEvents pnlWip As Panel
    Friend WithEvents lblWip As Label
    Friend WithEvents pgbWip As ProgressBar
    Friend WithEvents btnStop As Button
End Class
