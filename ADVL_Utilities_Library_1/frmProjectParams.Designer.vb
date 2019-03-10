<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjectParams
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.rbParameters = New System.Windows.Forms.RadioButton()
        Me.rbParentParameters = New System.Windows.Forms.RadioButton()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(566, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 18
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 40)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(618, 478)
        Me.DataGridView1.TabIndex = 24
        '
        'rbParameters
        '
        Me.rbParameters.AutoSize = True
        Me.rbParameters.Location = New System.Drawing.Point(12, 12)
        Me.rbParameters.Name = "rbParameters"
        Me.rbParameters.Size = New System.Drawing.Size(114, 17)
        Me.rbParameters.TabIndex = 25
        Me.rbParameters.TabStop = True
        Me.rbParameters.Text = "Project Parameters"
        Me.rbParameters.UseVisualStyleBackColor = True
        '
        'rbParentParameters
        '
        Me.rbParentParameters.AutoSize = True
        Me.rbParentParameters.Location = New System.Drawing.Point(132, 12)
        Me.rbParentParameters.Name = "rbParentParameters"
        Me.rbParentParameters.Size = New System.Drawing.Size(148, 17)
        Me.rbParentParameters.TabIndex = 26
        Me.rbParentParameters.TabStop = True
        Me.rbParentParameters.Text = "Parent Project Parameters"
        Me.rbParentParameters.UseVisualStyleBackColor = True
        '
        'frmProjectParams
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(642, 530)
        Me.Controls.Add(Me.rbParentParameters)
        Me.Controls.Add(Me.rbParameters)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmProjectParams"
        Me.Text = "Project Parameters"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Windows.Forms.Button
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents rbParameters As Windows.Forms.RadioButton
    Friend WithEvents rbParentParameters As Windows.Forms.RadioButton
End Class
