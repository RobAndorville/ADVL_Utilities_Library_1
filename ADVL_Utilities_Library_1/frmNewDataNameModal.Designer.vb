<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewDataNameModal
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtNewFileName = New System.Windows.Forms.TextBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtDataName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDataDescr = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtWarning = New System.Windows.Forms.TextBox()
        Me.txtDataLabel = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "New file name:"
        '
        'txtNewFileName
        '
        Me.txtNewFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNewFileName.Location = New System.Drawing.Point(95, 40)
        Me.txtNewFileName.Name = "txtNewFileName"
        Me.txtNewFileName.Size = New System.Drawing.Size(657, 20)
        Me.txtNewFileName.TabIndex = 1
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(12, 12)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(64, 22)
        Me.btnOK.TabIndex = 27
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(688, 12)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 22)
        Me.btnCancel.TabIndex = 28
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'txtDataName
        '
        Me.txtDataName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataName.Location = New System.Drawing.Point(95, 66)
        Me.txtDataName.Name = "txtDataName"
        Me.txtDataName.Size = New System.Drawing.Size(657, 20)
        Me.txtDataName.TabIndex = 29
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Data name:"
        '
        'txtDataDescr
        '
        Me.txtDataDescr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataDescr.Location = New System.Drawing.Point(95, 118)
        Me.txtDataDescr.Multiline = True
        Me.txtDataDescr.Name = "txtDataDescr"
        Me.txtDataDescr.Size = New System.Drawing.Size(657, 60)
        Me.txtDataDescr.TabIndex = 31
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 121)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Description:"
        '
        'txtWarning
        '
        Me.txtWarning.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWarning.Location = New System.Drawing.Point(12, 184)
        Me.txtWarning.Multiline = True
        Me.txtWarning.Name = "txtWarning"
        Me.txtWarning.ReadOnly = True
        Me.txtWarning.Size = New System.Drawing.Size(740, 60)
        Me.txtWarning.TabIndex = 33
        '
        'txtDataLabel
        '
        Me.txtDataLabel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataLabel.Location = New System.Drawing.Point(95, 92)
        Me.txtDataLabel.Name = "txtDataLabel"
        Me.txtDataLabel.Size = New System.Drawing.Size(657, 20)
        Me.txtDataLabel.TabIndex = 34
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 95)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 35
        Me.Label4.Text = "Data label:"
        '
        'frmNewDataFileModal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(764, 258)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtDataLabel)
        Me.Controls.Add(Me.txtWarning)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDataDescr)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtDataName)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtNewFileName)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmNewDataFileModal"
        Me.Text = "New Data File"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txtNewFileName As Windows.Forms.TextBox
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents txtDataName As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents txtDataDescr As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents txtWarning As Windows.Forms.TextBox
    Friend WithEvents txtDataLabel As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
End Class
