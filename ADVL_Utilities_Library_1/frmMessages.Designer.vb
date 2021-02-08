<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMessages
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
        Dim XmlHtmDisplaySettings1 As ADVL_Utilities_Library_1.XmlHtmDisplaySettings = New ADVL_Utilities_Library_1.XmlHtmDisplaySettings()
        Dim TextSettings1 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings2 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings3 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings4 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings5 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings6 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings7 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings8 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings9 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings10 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings11 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings12 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings13 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings14 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings15 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim XmlDisplaySettings1 As ADVL_Utilities_Library_1.XmlDisplaySettings = New ADVL_Utilities_Library_1.XmlDisplaySettings()
        Dim TextSettings16 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings17 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings18 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings19 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings20 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings21 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMessages))
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnClearMessages = New System.Windows.Forms.Button()
        Me.btnClearInstructions = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnShowXMsgTextTypes = New System.Windows.Forms.Button()
        Me.btnShowMsgTextTypes = New System.Windows.Forms.Button()
        Me.chkShowXMessages = New System.Windows.Forms.CheckBox()
        Me.chkShowSysMessages = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.XmlHtmDisplay1 = New ADVL_Utilities_Library_1.XmlHtmDisplay(Me.components)
        Me.XmlDisplay = New ADVL_Utilities_Library_1.XmlDisplay(Me.components)
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(657, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 18
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnClearMessages
        '
        Me.btnClearMessages.Location = New System.Drawing.Point(12, 12)
        Me.btnClearMessages.Name = "btnClearMessages"
        Me.btnClearMessages.Size = New System.Drawing.Size(96, 22)
        Me.btnClearMessages.TabIndex = 19
        Me.btnClearMessages.Text = "Clear Messages"
        Me.ToolTip1.SetToolTip(Me.btnClearMessages, "Clear all messages from the Messages tab")
        Me.btnClearMessages.UseVisualStyleBackColor = True
        '
        'btnClearInstructions
        '
        Me.btnClearInstructions.Location = New System.Drawing.Point(114, 12)
        Me.btnClearInstructions.Name = "btnClearInstructions"
        Me.btnClearInstructions.Size = New System.Drawing.Size(115, 22)
        Me.btnClearInstructions.TabIndex = 22
        Me.btnClearInstructions.Text = "Clear Instructions"
        Me.ToolTip1.SetToolTip(Me.btnClearInstructions, "Clear all instructions from the XMessage Instructions tab")
        Me.btnClearInstructions.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 47)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(709, 365)
        Me.TabControl1.TabIndex = 23
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.XmlHtmDisplay1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(701, 339)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Messages"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.XmlDisplay)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(701, 339)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "XMessage Instructions"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnShowXMsgTextTypes)
        Me.TabPage2.Controls.Add(Me.btnShowMsgTextTypes)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(701, 339)
        Me.TabPage2.TabIndex = 3
        Me.TabPage2.Text = "Settings"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnShowXMsgTextTypes
        '
        Me.btnShowXMsgTextTypes.Location = New System.Drawing.Point(6, 34)
        Me.btnShowXMsgTextTypes.Name = "btnShowXMsgTextTypes"
        Me.btnShowXMsgTextTypes.Size = New System.Drawing.Size(235, 22)
        Me.btnShowXMsgTextTypes.TabIndex = 1
        Me.btnShowXMsgTextTypes.Text = "Show Named XMessage Text Settings"
        Me.btnShowXMsgTextTypes.UseVisualStyleBackColor = True
        '
        'btnShowMsgTextTypes
        '
        Me.btnShowMsgTextTypes.Location = New System.Drawing.Point(6, 6)
        Me.btnShowMsgTextTypes.Name = "btnShowMsgTextTypes"
        Me.btnShowMsgTextTypes.Size = New System.Drawing.Size(235, 22)
        Me.btnShowMsgTextTypes.TabIndex = 0
        Me.btnShowMsgTextTypes.Text = "Show Named Message Text Settings"
        Me.btnShowMsgTextTypes.UseVisualStyleBackColor = True
        '
        'chkShowXMessages
        '
        Me.chkShowXMessages.AutoSize = True
        Me.chkShowXMessages.Location = New System.Drawing.Point(235, 16)
        Me.chkShowXMessages.Name = "chkShowXMessages"
        Me.chkShowXMessages.Size = New System.Drawing.Size(111, 17)
        Me.chkShowXMessages.TabIndex = 24
        Me.chkShowXMessages.Text = "Show XMessages"
        Me.ToolTip1.SetToolTip(Me.chkShowXMessages, "Show the XMessages - used to exchange information between applications")
        Me.chkShowXMessages.UseVisualStyleBackColor = True
        '
        'chkShowSysMessages
        '
        Me.chkShowSysMessages.AutoSize = True
        Me.chkShowSysMessages.Location = New System.Drawing.Point(352, 16)
        Me.chkShowSysMessages.Name = "chkShowSysMessages"
        Me.chkShowSysMessages.Size = New System.Drawing.Size(141, 17)
        Me.chkShowSysMessages.TabIndex = 25
        Me.chkShowSysMessages.Text = "Show System Messages"
        Me.ToolTip1.SetToolTip(Me.chkShowSysMessages, "Show the system generated XMessages")
        Me.chkShowSysMessages.UseVisualStyleBackColor = True
        '
        'XmlHtmDisplay1
        '
        Me.XmlHtmDisplay1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XmlHtmDisplay1.Font = New System.Drawing.Font("Arial", 10.0!)
        Me.XmlHtmDisplay1.Location = New System.Drawing.Point(3, 3)
        Me.XmlHtmDisplay1.Name = "XmlHtmDisplay1"
        TextSettings1.Bold = False
        TextSettings1.Color = System.Drawing.Color.Black
        TextSettings1.ColorIndex = 4
        TextSettings1.FontIndex = 1
        TextSettings1.FontName = "Arial"
        TextSettings1.HalfPointSize = 20
        TextSettings1.Italic = False
        TextSettings1.PointSize = 10.0!
        XmlHtmDisplaySettings1.DefaultText = TextSettings1
        TextSettings2.Bold = False
        TextSettings2.Color = System.Drawing.Color.Red
        TextSettings2.ColorIndex = 3
        TextSettings2.FontIndex = 1
        TextSettings2.FontName = "Arial"
        TextSettings2.HalfPointSize = 20
        TextSettings2.Italic = False
        TextSettings2.PointSize = 10.0!
        XmlHtmDisplaySettings1.HAttribute = TextSettings2
        TextSettings3.Bold = False
        TextSettings3.Color = System.Drawing.Color.Blue
        TextSettings3.ColorIndex = 1
        TextSettings3.FontIndex = 1
        TextSettings3.FontName = "Arial"
        TextSettings3.HalfPointSize = 20
        TextSettings3.Italic = False
        TextSettings3.PointSize = 10.0!
        XmlHtmDisplaySettings1.HChar = TextSettings3
        TextSettings4.Bold = False
        TextSettings4.Color = System.Drawing.Color.Green
        TextSettings4.ColorIndex = 6
        TextSettings4.FontIndex = 1
        TextSettings4.FontName = "Arial"
        TextSettings4.HalfPointSize = 20
        TextSettings4.Italic = False
        TextSettings4.PointSize = 10.0!
        XmlHtmDisplaySettings1.HComment = TextSettings4
        TextSettings5.Bold = False
        TextSettings5.Color = System.Drawing.Color.DarkRed
        TextSettings5.ColorIndex = 2
        TextSettings5.FontIndex = 1
        TextSettings5.FontName = "Arial"
        TextSettings5.HalfPointSize = 20
        TextSettings5.Italic = False
        TextSettings5.PointSize = 10.0!
        XmlHtmDisplaySettings1.HElement = TextSettings5
        TextSettings6.Bold = False
        TextSettings6.Color = System.Drawing.Color.Purple
        TextSettings6.ColorIndex = 7
        TextSettings6.FontIndex = 1
        TextSettings6.FontName = "Arial"
        TextSettings6.HalfPointSize = 20
        TextSettings6.Italic = False
        TextSettings6.PointSize = 10.0!
        XmlHtmDisplaySettings1.HStyle = TextSettings6
        TextSettings7.Bold = False
        TextSettings7.Color = System.Drawing.Color.Black
        TextSettings7.ColorIndex = 4
        TextSettings7.FontIndex = 1
        TextSettings7.FontName = "Arial"
        TextSettings7.HalfPointSize = 20
        TextSettings7.Italic = False
        TextSettings7.PointSize = 10.0!
        XmlHtmDisplaySettings1.HText = TextSettings7
        TextSettings8.Bold = False
        TextSettings8.Color = System.Drawing.Color.Blue
        TextSettings8.ColorIndex = 1
        TextSettings8.FontIndex = 1
        TextSettings8.FontName = "Arial"
        TextSettings8.HalfPointSize = 20
        TextSettings8.Italic = False
        TextSettings8.PointSize = 10.0!
        XmlHtmDisplaySettings1.HValue = TextSettings8
        TextSettings9.Bold = False
        TextSettings9.Color = System.Drawing.Color.Black
        TextSettings9.ColorIndex = 4
        TextSettings9.FontIndex = 1
        TextSettings9.FontName = "Arial"
        TextSettings9.HalfPointSize = 20
        TextSettings9.Italic = False
        TextSettings9.PointSize = 10.0!
        XmlHtmDisplaySettings1.PlainText = TextSettings9
        TextSettings10.Bold = False
        TextSettings10.Color = System.Drawing.Color.Red
        TextSettings10.ColorIndex = 3
        TextSettings10.FontIndex = 1
        TextSettings10.FontName = "Arial"
        TextSettings10.HalfPointSize = 20
        TextSettings10.Italic = False
        TextSettings10.PointSize = 10.0!
        XmlHtmDisplaySettings1.XAttributeKey = TextSettings10
        TextSettings11.Bold = False
        TextSettings11.Color = System.Drawing.Color.Blue
        TextSettings11.ColorIndex = 1
        TextSettings11.FontIndex = 1
        TextSettings11.FontName = "Arial"
        TextSettings11.HalfPointSize = 20
        TextSettings11.Italic = False
        TextSettings11.PointSize = 10.0!
        XmlHtmDisplaySettings1.XAttributeValue = TextSettings11
        TextSettings12.Bold = False
        TextSettings12.Color = System.Drawing.Color.Gray
        TextSettings12.ColorIndex = 5
        TextSettings12.FontIndex = 1
        TextSettings12.FontName = "Arial"
        TextSettings12.HalfPointSize = 20
        TextSettings12.Italic = False
        TextSettings12.PointSize = 10.0!
        XmlHtmDisplaySettings1.XComment = TextSettings12
        TextSettings13.Bold = False
        TextSettings13.Color = System.Drawing.Color.DarkRed
        TextSettings13.ColorIndex = 2
        TextSettings13.FontIndex = 1
        TextSettings13.FontName = "Arial"
        TextSettings13.HalfPointSize = 20
        TextSettings13.Italic = False
        TextSettings13.PointSize = 10.0!
        XmlHtmDisplaySettings1.XElement = TextSettings13
        XmlHtmDisplaySettings1.XIndentSpaces = 4
        XmlHtmDisplaySettings1.XmlLargeFileSizeLimit = 1000000
        TextSettings14.Bold = False
        TextSettings14.Color = System.Drawing.Color.Blue
        TextSettings14.ColorIndex = 1
        TextSettings14.FontIndex = 1
        TextSettings14.FontName = "Arial"
        TextSettings14.HalfPointSize = 20
        TextSettings14.Italic = False
        TextSettings14.PointSize = 10.0!
        XmlHtmDisplaySettings1.XTag = TextSettings14
        TextSettings15.Bold = False
        TextSettings15.Color = System.Drawing.Color.Black
        TextSettings15.ColorIndex = 4
        TextSettings15.FontIndex = 1
        TextSettings15.FontName = "Arial"
        TextSettings15.HalfPointSize = 20
        TextSettings15.Italic = False
        TextSettings15.PointSize = 10.0!
        XmlHtmDisplaySettings1.XValue = TextSettings15
        Me.XmlHtmDisplay1.Settings = XmlHtmDisplaySettings1
        Me.XmlHtmDisplay1.Size = New System.Drawing.Size(695, 333)
        Me.XmlHtmDisplay1.TabIndex = 20
        Me.XmlHtmDisplay1.Text = ""
        Me.XmlHtmDisplay1.WordWrap = False
        '
        'XmlDisplay
        '
        Me.XmlDisplay.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XmlDisplay.Location = New System.Drawing.Point(3, 3)
        Me.XmlDisplay.Name = "XmlDisplay"
        TextSettings16.Bold = False
        TextSettings16.Color = System.Drawing.Color.Red
        TextSettings16.ColorIndex = 3
        TextSettings16.FontIndex = 1
        TextSettings16.FontName = "Arial"
        TextSettings16.HalfPointSize = 20
        TextSettings16.Italic = False
        TextSettings16.PointSize = 10.0!
        XmlDisplaySettings1.AttributeKey = TextSettings16
        TextSettings17.Bold = False
        TextSettings17.Color = System.Drawing.Color.Blue
        TextSettings17.ColorIndex = 1
        TextSettings17.FontIndex = 1
        TextSettings17.FontName = "Arial"
        TextSettings17.HalfPointSize = 20
        TextSettings17.Italic = False
        TextSettings17.PointSize = 10.0!
        XmlDisplaySettings1.AttributeValue = TextSettings17
        TextSettings18.Bold = False
        TextSettings18.Color = System.Drawing.Color.Gray
        TextSettings18.ColorIndex = 5
        TextSettings18.FontIndex = 1
        TextSettings18.FontName = "Arial"
        TextSettings18.HalfPointSize = 20
        TextSettings18.Italic = False
        TextSettings18.PointSize = 10.0!
        XmlDisplaySettings1.Comment = TextSettings18
        TextSettings19.Bold = False
        TextSettings19.Color = System.Drawing.Color.DarkRed
        TextSettings19.ColorIndex = 2
        TextSettings19.FontIndex = 1
        TextSettings19.FontName = "Arial"
        TextSettings19.HalfPointSize = 20
        TextSettings19.Italic = False
        TextSettings19.PointSize = 10.0!
        XmlDisplaySettings1.Element = TextSettings19
        XmlDisplaySettings1.IndentSpaces = 4
        TextSettings20.Bold = False
        TextSettings20.Color = System.Drawing.Color.Blue
        TextSettings20.ColorIndex = 1
        TextSettings20.FontIndex = 1
        TextSettings20.FontName = "Arial"
        TextSettings20.HalfPointSize = 20
        TextSettings20.Italic = False
        TextSettings20.PointSize = 10.0!
        XmlDisplaySettings1.Tag = TextSettings20
        TextSettings21.Bold = False
        TextSettings21.Color = System.Drawing.Color.Black
        TextSettings21.ColorIndex = 4
        TextSettings21.FontIndex = 1
        TextSettings21.FontName = "Arial"
        TextSettings21.HalfPointSize = 20
        TextSettings21.Italic = False
        TextSettings21.PointSize = 10.0!
        XmlDisplaySettings1.Value = TextSettings21
        XmlDisplaySettings1.XmlLargeFileSizeLimit = 1000000
        Me.XmlDisplay.Settings = XmlDisplaySettings1
        Me.XmlDisplay.Size = New System.Drawing.Size(695, 333)
        Me.XmlDisplay.TabIndex = 0
        Me.XmlDisplay.Text = ""
        Me.XmlDisplay.WordWrap = False
        '
        'frmMessages
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(733, 424)
        Me.Controls.Add(Me.chkShowSysMessages)
        Me.Controls.Add(Me.chkShowXMessages)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnClearInstructions)
        Me.Controls.Add(Me.btnClearMessages)
        Me.Controls.Add(Me.btnExit)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMessages"
        Me.Text = "Messages"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnClearMessages As System.Windows.Forms.Button
    Friend WithEvents btnClearInstructions As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As Windows.Forms.TabPage
    Friend WithEvents XmlDisplay As XmlDisplay
    Friend WithEvents TabPage2 As Windows.Forms.TabPage
    Friend WithEvents btnShowMsgTextTypes As Windows.Forms.Button
    Friend WithEvents XmlHtmDisplay1 As XmlHtmDisplay
    Friend WithEvents btnShowXMsgTextTypes As Windows.Forms.Button
    Friend WithEvents chkShowXMessages As Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents chkShowSysMessages As Windows.Forms.CheckBox
End Class
