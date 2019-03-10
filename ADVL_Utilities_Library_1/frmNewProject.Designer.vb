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
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.chkProjectDir = New System.Windows.Forms.CheckBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtDataFileName = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.btnFindDataFileDir = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtDataFileDirectory = New System.Windows.Forms.TextBox()
        Me.txtDataDescription = New System.Windows.Forms.TextBox()
        Me.txtDataName = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
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
        Me.TabPage5.SuspendLayout()
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
        Me.TabControl1.Size = New System.Drawing.Size(624, 246)
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
        Me.TabPage1.Size = New System.Drawing.Size(616, 220)
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
        Me.TabPage2.Size = New System.Drawing.Size(616, 220)
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
        Me.TabPage3.Size = New System.Drawing.Size(616, 220)
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
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Location = New System.Drawing.Point(3, 3)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(610, 214)
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
        Me.TabPage4.Size = New System.Drawing.Size(602, 188)
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
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.chkProjectDir)
        Me.TabPage5.Controls.Add(Me.Label18)
        Me.TabPage5.Controls.Add(Me.txtDataFileName)
        Me.TabPage5.Controls.Add(Me.Label13)
        Me.TabPage5.Controls.Add(Me.btnFindDataFileDir)
        Me.TabPage5.Controls.Add(Me.Label14)
        Me.TabPage5.Controls.Add(Me.txtDataFileDirectory)
        Me.TabPage5.Controls.Add(Me.txtDataDescription)
        Me.TabPage5.Controls.Add(Me.txtDataName)
        Me.TabPage5.Controls.Add(Me.Label15)
        Me.TabPage5.Controls.Add(Me.Label16)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(602, 188)
        Me.TabPage5.TabIndex = 1
        Me.TabPage5.Text = "Data"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'chkProjectDir
        '
        Me.chkProjectDir.AutoSize = True
        Me.chkProjectDir.Location = New System.Drawing.Point(6, 136)
        Me.chkProjectDir.Name = "chkProjectDir"
        Me.chkProjectDir.Size = New System.Drawing.Size(75, 17)
        Me.chkProjectDir.TabIndex = 23
        Me.chkProjectDir.Text = "Project Dir"
        Me.chkProjectDir.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(6, 169)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(60, 13)
        Me.Label18.TabIndex = 23
        Me.Label18.Text = "(.AdvlData)"
        '
        'txtDataFileName
        '
        Me.txtDataFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataFileName.Location = New System.Drawing.Point(90, 159)
        Me.txtDataFileName.Name = "txtDataFileName"
        Me.txtDataFileName.Size = New System.Drawing.Size(506, 20)
        Me.txtDataFileName.TabIndex = 48
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 156)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(78, 13)
        Me.Label13.TabIndex = 47
        Me.Label13.Text = "Data file name:"
        '
        'btnFindDataFileDir
        '
        Me.btnFindDataFileDir.Location = New System.Drawing.Point(42, 113)
        Me.btnFindDataFileDir.Name = "btnFindDataFileDir"
        Me.btnFindDataFileDir.Size = New System.Drawing.Size(42, 22)
        Me.btnFindDataFileDir.TabIndex = 46
        Me.btnFindDataFileDir.Text = "Find"
        Me.btnFindDataFileDir.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 116)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(32, 13)
        Me.Label14.TabIndex = 45
        Me.Label14.Text = "Path:"
        '
        'txtDataFileDirectory
        '
        Me.txtDataFileDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataFileDirectory.Location = New System.Drawing.Point(90, 113)
        Me.txtDataFileDirectory.Multiline = True
        Me.txtDataFileDirectory.Name = "txtDataFileDirectory"
        Me.txtDataFileDirectory.Size = New System.Drawing.Size(506, 40)
        Me.txtDataFileDirectory.TabIndex = 44
        '
        'txtDataDescription
        '
        Me.txtDataDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataDescription.Location = New System.Drawing.Point(90, 32)
        Me.txtDataDescription.Multiline = True
        Me.txtDataDescription.Name = "txtDataDescription"
        Me.txtDataDescription.Size = New System.Drawing.Size(506, 75)
        Me.txtDataDescription.TabIndex = 43
        '
        'txtDataName
        '
        Me.txtDataName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataName.Location = New System.Drawing.Point(90, 6)
        Me.txtDataName.Name = "txtDataName"
        Me.txtDataName.Size = New System.Drawing.Size(506, 20)
        Me.txtDataName.TabIndex = 42
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 35)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(63, 13)
        Me.Label15.TabIndex = 41
        Me.Label15.Text = "Description:"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(6, 9)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 13)
        Me.Label16.TabIndex = 40
        Me.Label16.Text = "Data Name:"
        '
        'TabPage6
        '
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(602, 188)
        Me.TabPage6.TabIndex = 2
        Me.TabPage6.Text = "Settings"
        Me.TabPage6.UseVisualStyleBackColor = True
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
        Me.ClientSize = New System.Drawing.Size(648, 565)
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
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
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
    Friend WithEvents txtDataFileName As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btnFindDataFileDir As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtDataFileDirectory As System.Windows.Forms.TextBox
    Friend WithEvents txtDataDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtDataName As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
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
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents chkProjectDir As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage6 As Windows.Forms.TabPage
End Class
