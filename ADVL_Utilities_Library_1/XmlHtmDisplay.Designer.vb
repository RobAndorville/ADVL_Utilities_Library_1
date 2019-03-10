Partial Class XmlHtmDisplay
    'Inherits System.ComponentModel.Component

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

        'Debug.Print("Running New(ByVal container As System.ComponentModel.IContainer)")
        'Debug.Print("Settings.HAttribute.Color.Name = " & Settings.HAttribute.Color.Name)

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Debug.Print("Running New()")
        'Debug.Print("Settings.HAttribute.Color.Name = " & Settings.HAttribute.Color.Name)

        Settings.ClearAllTextTypes()

        Settings.AddNewTextType("Message")
        Settings.TextType("Message").Color = System.Drawing.Color.Black
        Settings.TextType("Message").PointSize = 10

        Settings.AddNewTextType("Heading")
        Settings.TextType("Heading").Bold = True
        Settings.TextType("Heading").Color = System.Drawing.Color.Black
        Settings.TextType("Heading").PointSize = 12

        Settings.AddNewTextType("Warning")
        Settings.TextType("Warning").Bold = True
        Settings.TextType("Warning").Color = System.Drawing.Color.Red
        Settings.TextType("Warning").PointSize = 10

        'Me.Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point) 'Set initial font.
        Me.Font = New Drawing.Font("Arial", 10, Drawing.FontStyle.Regular) 'Set initial font.

        Settings.UpdateFontIndexes()
        Settings.UpdateColorIndexes()

        'Debug.Print("Finished Running New()")
        'Debug.Print("Settings.HAttribute.Color.Name = " & Settings.HAttribute.Color.Name)


    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        'Settings.UpdateFontIndexes()
        'Settings.UpdateColorIndexes()
        'Debug.Print("Running InitializeComponent()")
        'Debug.Print("Settings.HAttribute.Color.Name = " & Settings.HAttribute.Color.Name)
    End Sub

End Class
