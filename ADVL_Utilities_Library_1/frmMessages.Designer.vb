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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMessages))
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnClearMessages = New System.Windows.Forms.Button()
        Me.btnClearInstructions = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.rtbMessages = New System.Windows.Forms.RichTextBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.rtbInstructions = New System.Windows.Forms.RichTextBox()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(595, 12)
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
        Me.btnClearMessages.UseVisualStyleBackColor = True
        '
        'btnClearInstructions
        '
        Me.btnClearInstructions.Location = New System.Drawing.Point(114, 12)
        Me.btnClearInstructions.Name = "btnClearInstructions"
        Me.btnClearInstructions.Size = New System.Drawing.Size(115, 22)
        Me.btnClearInstructions.TabIndex = 22
        Me.btnClearInstructions.Text = "Clear Instructions"
        Me.btnClearInstructions.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 47)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(647, 410)
        Me.TabControl1.TabIndex = 23
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.rtbMessages)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(639, 384)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Messages"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'rtbMessages
        '
        Me.rtbMessages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbMessages.Location = New System.Drawing.Point(6, 6)
        Me.rtbMessages.Name = "rtbMessages"
        Me.rtbMessages.Size = New System.Drawing.Size(627, 372)
        Me.rtbMessages.TabIndex = 19
        Me.rtbMessages.Text = ""
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.rtbInstructions)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(639, 384)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "XMessage Instructions"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'rtbInstructions
        '
        Me.rtbInstructions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbInstructions.Location = New System.Drawing.Point(6, 6)
        Me.rtbInstructions.Name = "rtbInstructions"
        Me.rtbInstructions.Size = New System.Drawing.Size(627, 372)
        Me.rtbInstructions.TabIndex = 22
        Me.rtbInstructions.Text = ""
        '
        'frmMessages
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(671, 469)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnClearInstructions)
        Me.Controls.Add(Me.btnClearMessages)
        Me.Controls.Add(Me.btnExit)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMessages"
        Me.Text = "Messages"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnClearMessages As System.Windows.Forms.Button
    Friend WithEvents btnClearInstructions As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents rtbMessages As System.Windows.Forms.RichTextBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents rtbInstructions As System.Windows.Forms.RichTextBox
End Class
