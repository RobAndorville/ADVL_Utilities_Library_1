<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewProject
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtProjectDirectoryName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDirectoryPath = New System.Windows.Forms.TextBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtProjectFileName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnFindProjFileDir = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtPFDirectoryPath = New System.Windows.Forms.TextBox()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.TabControl2 = New System.Windows.Forms.TabControl()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.txtHPDirectoryName = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btnFindHybridProjDir = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtHPDirectoryPath = New System.Windows.Forms.TextBox()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.cmbHPSettingsType = New System.Windows.Forms.ComboBox()
        Me.Label70 = New System.Windows.Forms.Label()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.txtHPSettingsName = New System.Windows.Forms.TextBox()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.cmbHPDataType = New System.Windows.Forms.ComboBox()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.txtHPDataName = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.TabPage7 = New System.Windows.Forms.TabPage()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.cmbHPSystemType = New System.Windows.Forms.ComboBox()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.txtHPSystemName = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtAuthorDescription = New System.Windows.Forms.TextBox()
        Me.txtAuthorContact = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtAuthorName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rbHybrid = New System.Windows.Forms.RadioButton()
        Me.rbProjectFile = New System.Windows.Forms.RadioButton()
        Me.rbProjectDir = New System.Windows.Forms.RadioButton()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtProjectDescription = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtProjectName = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage6.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.TabPage7.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(572, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 19
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(12, 12)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(64, 22)
        Me.btnCreate.TabIndex = 20
        Me.btnCreate.Text = "Create"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(12, 307)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(624, 309)
        Me.TabControl1.TabIndex = 21
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtProjectDirectoryName)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.btnFind)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.txtDirectoryPath)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(616, 283)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Project Directory"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtProjectDirectoryName
        '
        Me.txtProjectDirectoryName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProjectDirectoryName.Location = New System.Drawing.Point(127, 59)
        Me.txtProjectDirectoryName.Name = "txtProjectDirectoryName"
        Me.txtProjectDirectoryName.Size = New System.Drawing.Size(483, 20)
        Me.txtProjectDirectoryName.TabIndex = 30
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(115, 13)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Project directory name:"
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(6, 25)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(64, 22)
        Me.btnFind.TabIndex = 28
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "Path:"
        '
        'txtDirectoryPath
        '
        Me.txtDirectoryPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDirectoryPath.Location = New System.Drawing.Point(86, 6)
        Me.txtDirectoryPath.Multiline = True
        Me.txtDirectoryPath.Name = "txtDirectoryPath"
        Me.txtDirectoryPath.Size = New System.Drawing.Size(524, 47)
        Me.txtDirectoryPath.TabIndex = 26
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label17)
        Me.TabPage2.Controls.Add(Me.txtProjectFileName)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.btnFindProjFileDir)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.txtPFDirectoryPath)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(616, 283)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Project Archive"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(97, 83)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(140, 13)
        Me.Label17.TabIndex = 40
        Me.Label17.Text = "(File extension: .AdvlProject)"
        '
        'txtProjectFileName
        '
        Me.txtProjectFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProjectFileName.Location = New System.Drawing.Point(122, 60)
        Me.txtProjectFileName.Name = "txtProjectFileName"
        Me.txtProjectFileName.Size = New System.Drawing.Size(488, 20)
        Me.txtProjectFileName.TabIndex = 39
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 63)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(110, 13)
        Me.Label5.TabIndex = 38
        Me.Label5.Text = "Project archive name:"
        '
        'btnFindProjFileDir
        '
        Me.btnFindProjFileDir.Location = New System.Drawing.Point(6, 25)
        Me.btnFindProjFileDir.Name = "btnFindProjFileDir"
        Me.btnFindProjFileDir.Size = New System.Drawing.Size(64, 22)
        Me.btnFindProjFileDir.TabIndex = 37
        Me.btnFindProjFileDir.Text = "Find"
        Me.btnFindProjFileDir.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 36
        Me.Label6.Text = "Path:"
        '
        'txtPFDirectoryPath
        '
        Me.txtPFDirectoryPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPFDirectoryPath.Location = New System.Drawing.Point(86, 6)
        Me.txtPFDirectoryPath.Multiline = True
        Me.txtPFDirectoryPath.Name = "txtPFDirectoryPath"
        Me.txtPFDirectoryPath.Size = New System.Drawing.Size(524, 48)
        Me.txtPFDirectoryPath.TabIndex = 35
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.TabControl2)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(616, 283)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Hybrid Project"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Location = New System.Drawing.Point(3, 3)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(610, 277)
        Me.TabControl2.TabIndex = 0
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.txtHPDirectoryName)
        Me.TabPage4.Controls.Add(Me.Label9)
        Me.TabPage4.Controls.Add(Me.btnFindHybridProjDir)
        Me.TabPage4.Controls.Add(Me.Label10)
        Me.TabPage4.Controls.Add(Me.txtHPDirectoryPath)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(602, 251)
        Me.TabPage4.TabIndex = 0
        Me.TabPage4.Text = "Project Directory"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'txtHPDirectoryName
        '
        Me.txtHPDirectoryName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHPDirectoryName.Location = New System.Drawing.Point(121, 60)
        Me.txtHPDirectoryName.Name = "txtHPDirectoryName"
        Me.txtHPDirectoryName.Size = New System.Drawing.Size(475, 20)
        Me.txtHPDirectoryName.TabIndex = 39
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 63)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(115, 13)
        Me.Label9.TabIndex = 38
        Me.Label9.Text = "Project directory name:"
        '
        'btnFindHybridProjDir
        '
        Me.btnFindHybridProjDir.Location = New System.Drawing.Point(6, 25)
        Me.btnFindHybridProjDir.Name = "btnFindHybridProjDir"
        Me.btnFindHybridProjDir.Size = New System.Drawing.Size(64, 22)
        Me.btnFindHybridProjDir.TabIndex = 37
        Me.btnFindHybridProjDir.Text = "Find"
        Me.btnFindHybridProjDir.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 9)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(32, 13)
        Me.Label10.TabIndex = 36
        Me.Label10.Text = "Path:"
        '
        'txtHPDirectoryPath
        '
        Me.txtHPDirectoryPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHPDirectoryPath.Location = New System.Drawing.Point(86, 6)
        Me.txtHPDirectoryPath.Multiline = True
        Me.txtHPDirectoryPath.Name = "txtHPDirectoryPath"
        Me.txtHPDirectoryPath.Size = New System.Drawing.Size(510, 48)
        Me.txtHPDirectoryPath.TabIndex = 35
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.cmbHPSettingsType)
        Me.TabPage6.Controls.Add(Me.Label70)
        Me.TabPage6.Controls.Add(Me.Label74)
        Me.TabPage6.Controls.Add(Me.Label69)
        Me.TabPage6.Controls.Add(Me.txtHPSettingsName)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(602, 251)
        Me.TabPage6.TabIndex = 2
        Me.TabPage6.Text = "Settings"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'cmbHPSettingsType
        '
        Me.cmbHPSettingsType.FormattingEnabled = True
        Me.cmbHPSettingsType.Location = New System.Drawing.Point(126, 32)
        Me.cmbHPSettingsType.Name = "cmbHPSettingsType"
        Me.cmbHPSettingsType.Size = New System.Drawing.Size(135, 21)
        Me.cmbHPSettingsType.TabIndex = 293
        '
        'Label70
        '
        Me.Label70.AutoSize = True
        Me.Label70.Location = New System.Drawing.Point(3, 35)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(111, 13)
        Me.Label70.TabIndex = 292
        Me.Label70.Text = "Settings location type:"
        '
        'Label74
        '
        Me.Label74.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label74.AutoSize = True
        Me.Label74.Location = New System.Drawing.Point(576, 9)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(23, 13)
        Me.Label74.TabIndex = 291
        Me.Label74.Text = ".zip"
        '
        'Label69
        '
        Me.Label69.AutoSize = True
        Me.Label69.Location = New System.Drawing.Point(3, 9)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(117, 13)
        Me.Label69.TabIndex = 46
        Me.Label69.Text = "Settings location name:"
        '
        'txtHPSettingsName
        '
        Me.txtHPSettingsName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHPSettingsName.Location = New System.Drawing.Point(126, 6)
        Me.txtHPSettingsName.Name = "txtHPSettingsName"
        Me.txtHPSettingsName.Size = New System.Drawing.Size(444, 20)
        Me.txtHPSettingsName.TabIndex = 45
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.Label57)
        Me.TabPage5.Controls.Add(Me.cmbHPDataType)
        Me.TabPage5.Controls.Add(Me.Label73)
        Me.TabPage5.Controls.Add(Me.txtHPDataName)
        Me.TabPage5.Controls.Add(Me.Label16)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(602, 251)
        Me.TabPage5.TabIndex = 1
        Me.TabPage5.Text = "Data"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.Location = New System.Drawing.Point(6, 35)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(96, 13)
        Me.Label57.TabIndex = 291
        Me.Label57.Text = "Data location type:"
        '
        'cmbHPDataType
        '
        Me.cmbHPDataType.FormattingEnabled = True
        Me.cmbHPDataType.Location = New System.Drawing.Point(114, 32)
        Me.cmbHPDataType.Name = "cmbHPDataType"
        Me.cmbHPDataType.Size = New System.Drawing.Size(135, 21)
        Me.cmbHPDataType.TabIndex = 290
        '
        'Label73
        '
        Me.Label73.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label73.AutoSize = True
        Me.Label73.Location = New System.Drawing.Point(573, 9)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(23, 13)
        Me.Label73.TabIndex = 289
        Me.Label73.Text = ".zip"
        '
        'txtHPDataName
        '
        Me.txtHPDataName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHPDataName.Location = New System.Drawing.Point(114, 6)
        Me.txtHPDataName.Name = "txtHPDataName"
        Me.txtHPDataName.Size = New System.Drawing.Size(453, 20)
        Me.txtHPDataName.TabIndex = 42
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(6, 9)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(102, 13)
        Me.Label16.TabIndex = 40
        Me.Label16.Text = "Data location name:"
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.Label71)
        Me.TabPage7.Controls.Add(Me.cmbHPSystemType)
        Me.TabPage7.Controls.Add(Me.Label75)
        Me.TabPage7.Controls.Add(Me.Label72)
        Me.TabPage7.Controls.Add(Me.txtHPSystemName)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Size = New System.Drawing.Size(602, 251)
        Me.TabPage7.TabIndex = 3
        Me.TabPage7.Text = "System"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'Label71
        '
        Me.Label71.AutoSize = True
        Me.Label71.Location = New System.Drawing.Point(3, 35)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(107, 13)
        Me.Label71.TabIndex = 297
        Me.Label71.Text = "System location type:"
        '
        'cmbHPSystemType
        '
        Me.cmbHPSystemType.FormattingEnabled = True
        Me.cmbHPSystemType.Location = New System.Drawing.Point(122, 32)
        Me.cmbHPSystemType.Name = "cmbHPSystemType"
        Me.cmbHPSystemType.Size = New System.Drawing.Size(135, 21)
        Me.cmbHPSystemType.TabIndex = 296
        '
        'Label75
        '
        Me.Label75.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label75.AutoSize = True
        Me.Label75.Location = New System.Drawing.Point(576, 9)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(23, 13)
        Me.Label75.TabIndex = 295
        Me.Label75.Text = ".zip"
        '
        'Label72
        '
        Me.Label72.AutoSize = True
        Me.Label72.Location = New System.Drawing.Point(3, 9)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(113, 13)
        Me.Label72.TabIndex = 293
        Me.Label72.Text = "System location name:"
        '
        'txtHPSystemName
        '
        Me.txtHPSystemName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHPSystemName.Location = New System.Drawing.Point(122, 6)
        Me.txtHPSystemName.Name = "txtHPSystemName"
        Me.txtHPSystemName.Size = New System.Drawing.Size(448, 20)
        Me.txtHPSystemName.TabIndex = 292
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.rbHybrid)
        Me.GroupBox1.Controls.Add(Me.rbProjectFile)
        Me.GroupBox1.Controls.Add(Me.rbProjectDir)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtProjectDescription)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtProjectName)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(624, 261)
        Me.GroupBox1.TabIndex = 22
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "New Project:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.txtAuthorDescription)
        Me.GroupBox2.Controls.Add(Me.txtAuthorContact)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.txtAuthorName)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 125)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(612, 130)
        Me.GroupBox2.TabIndex = 35
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Author:"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 48)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(63, 13)
        Me.Label12.TabIndex = 36
        Me.Label12.Text = "Description:"
        '
        'txtAuthorDescription
        '
        Me.txtAuthorDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAuthorDescription.Location = New System.Drawing.Point(84, 45)
        Me.txtAuthorDescription.Multiline = True
        Me.txtAuthorDescription.Name = "txtAuthorDescription"
        Me.txtAuthorDescription.Size = New System.Drawing.Size(522, 50)
        Me.txtAuthorDescription.TabIndex = 33
        '
        'txtAuthorContact
        '
        Me.txtAuthorContact.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAuthorContact.Location = New System.Drawing.Point(84, 101)
        Me.txtAuthorContact.Name = "txtAuthorContact"
        Me.txtAuthorContact.Size = New System.Drawing.Size(522, 20)
        Me.txtAuthorContact.TabIndex = 32
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 104)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(47, 13)
        Me.Label11.TabIndex = 31
        Me.Label11.Text = "Contact:"
        '
        'txtAuthorName
        '
        Me.txtAuthorName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAuthorName.Location = New System.Drawing.Point(84, 19)
        Me.txtAuthorName.Name = "txtAuthorName"
        Me.txtAuthorName.Size = New System.Drawing.Size(522, 20)
        Me.txtAuthorName.TabIndex = 30
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 22)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(38, 13)
        Me.Label7.TabIndex = 29
        Me.Label7.Text = "Name:"
        '
        'rbHybrid
        '
        Me.rbHybrid.AutoSize = True
        Me.rbHybrid.Location = New System.Drawing.Point(302, 102)
        Me.rbHybrid.Name = "rbHybrid"
        Me.rbHybrid.Size = New System.Drawing.Size(55, 17)
        Me.rbHybrid.TabIndex = 34
        Me.rbHybrid.TabStop = True
        Me.rbHybrid.Text = "Hybrid"
        Me.ToolTip1.SetToolTip(Me.rbHybrid, "Store settings in a project directory and data in one or more files")
        Me.rbHybrid.UseVisualStyleBackColor = True
        '
        'rbProjectFile
        '
        Me.rbProjectFile.AutoSize = True
        Me.rbProjectFile.Location = New System.Drawing.Point(199, 102)
        Me.rbProjectFile.Name = "rbProjectFile"
        Me.rbProjectFile.Size = New System.Drawing.Size(97, 17)
        Me.rbProjectFile.TabIndex = 33
        Me.rbProjectFile.TabStop = True
        Me.rbProjectFile.Text = "Project Archive"
        Me.ToolTip1.SetToolTip(Me.rbProjectFile, "Store all settings and data in a single project file")
        Me.rbProjectFile.UseVisualStyleBackColor = True
        '
        'rbProjectDir
        '
        Me.rbProjectDir.AutoSize = True
        Me.rbProjectDir.Location = New System.Drawing.Point(90, 102)
        Me.rbProjectDir.Name = "rbProjectDir"
        Me.rbProjectDir.Size = New System.Drawing.Size(103, 17)
        Me.rbProjectDir.TabIndex = 32
        Me.rbProjectDir.TabStop = True
        Me.rbProjectDir.Text = "Project Directory"
        Me.ToolTip1.SetToolTip(Me.rbProjectDir, "Store all settings and data in a single project directory")
        Me.rbProjectDir.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(34, 13)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Type:"
        '
        'txtProjectDescription
        '
        Me.txtProjectDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProjectDescription.Location = New System.Drawing.Point(90, 46)
        Me.txtProjectDescription.Multiline = True
        Me.txtProjectDescription.Name = "txtProjectDescription"
        Me.txtProjectDescription.Size = New System.Drawing.Size(527, 50)
        Me.txtProjectDescription.TabIndex = 28
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Description:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Name:"
        '
        'txtProjectName
        '
        Me.txtProjectName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProjectName.Location = New System.Drawing.Point(90, 19)
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(527, 20)
        Me.txtProjectName.TabIndex = 25
        '
        'frmNewProject
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(648, 628)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmNewProject"
        Me.Text = "New Project"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage6.ResumeLayout(False)
        Me.TabPage6.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        Me.TabPage7.ResumeLayout(False)
        Me.TabPage7.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents txtProjectDirectoryName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDirectoryPath As System.Windows.Forms.TextBox
    Friend WithEvents txtProjectFileName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindProjFileDir As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPFDirectoryPath As System.Windows.Forms.TextBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents txtHPDirectoryName As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnFindHybridProjDir As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtHPDirectoryPath As System.Windows.Forms.TextBox
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents txtHPDataName As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtAuthorName As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtProjectDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtProjectName As System.Windows.Forms.TextBox
    Friend WithEvents rbHybrid As System.Windows.Forms.RadioButton
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents rbProjectFile As System.Windows.Forms.RadioButton
    Friend WithEvents rbProjectDir As System.Windows.Forms.RadioButton
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtAuthorContact As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtAuthorDescription As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TabPage6 As Windows.Forms.TabPage
    Friend WithEvents cmbHPSettingsType As Windows.Forms.ComboBox
    Friend WithEvents Label70 As Windows.Forms.Label
    Friend WithEvents Label74 As Windows.Forms.Label
    Friend WithEvents Label69 As Windows.Forms.Label
    Friend WithEvents txtHPSettingsName As Windows.Forms.TextBox
    Friend WithEvents Label57 As Windows.Forms.Label
    Friend WithEvents cmbHPDataType As Windows.Forms.ComboBox
    Friend WithEvents Label73 As Windows.Forms.Label
    Friend WithEvents TabPage7 As Windows.Forms.TabPage
    Friend WithEvents Label71 As Windows.Forms.Label
    Friend WithEvents cmbHPSystemType As Windows.Forms.ComboBox
    Friend WithEvents Label75 As Windows.Forms.Label
    Friend WithEvents Label72 As Windows.Forms.Label
    Friend WithEvents txtHPSystemName As Windows.Forms.TextBox
End Class
