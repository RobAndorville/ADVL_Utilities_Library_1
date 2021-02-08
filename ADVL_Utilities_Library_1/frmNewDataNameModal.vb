Public Class frmNewDataNameModal
    'The NewDataFileModal form is used to get the name of a new data file.
    'Properties can be set for the following options:
    '  FileExtension
    '  GetDataName
    '  GetDataDescription
    'The form can be displayed as a modal form.

#Region " Variable Declarations - All the variables used in this form." '--------------------------------------------------------------------------------------------------------------------

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public DataLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store data.
    Public ApplicationName As String 'The name of the application using this form.

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '----------------------------------------------------------------------------------------------------------

    Private _title As String = "New Data Name" 'The title to show at the top of the form.
    Property Title As String
        Get
            Return _title
        End Get
        Set(value As String)
            _title = value
        End Set
    End Property

    Private _entryName As String = "" 'The new data file Entry Name - Form display settings are saved for each unique EntryName - Examples: MonteCarloModel, CorrelationMatrix, CalculationTree etc
    Property EntryName As String
        Get
            Return _entryName
        End Get
        Set(value As String)
            _entryName = value
        End Set
    End Property

    Private _fileExtension As String = ""
    Property FileExtension As String 'File extension of file to include in the list. (If FileExtension = "" include all files.)
        Get
            Return _fileExtension
        End Get
        Set(value As String)
            _fileExtension = value
        End Set
    End Property

    Private _getFileName As Boolean = False 'If True the form will show a text box to enter the File Name
    Property GetFileName As Boolean
        Get
            Return _getFileName
        End Get
        Set(value As Boolean)
            _getFileName = value
        End Set
    End Property

    Private _fileName As String = "" 'The name of the new file.
    Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
        End Set
    End Property

    Private _getDataName As Boolean = False 'If True the form will show a text box to enter the Data Name
    Property GetDataName As Boolean
        Get
            Return _getDataName
        End Get
        Set(value As Boolean)
            _getDataName = value
        End Set
    End Property

    Private _dataName As String = "" 'The name of the dataset contained in the file.
    Property DataName As String
        Get
            Return _dataName
        End Get
        Set(value As String)
            _dataName = value
        End Set
    End Property

    Private _getDataLabel As Boolean = False 'If True the form will show a text box to enter the Data Label
    Property GetDataLabel As Boolean
        Get
            Return _getDataLabel
        End Get
        Set(value As Boolean)
            _getDataLabel = value
        End Set
    End Property

    Private _dataLabel As String = "" 'The label used to identify the DataSet on a chart.
    Property DataLabel As String
        Get
            Return _dataLabel
        End Get
        Set(value As String)
            _dataLabel = value
        End Set
    End Property

    Private _getDataDescription As Boolean = False 'If True the form will show a text box to enter a description of the data.
    Property GetDataDescription As Boolean
        Get
            Return _getDataDescription
        End Get
        Set(value As Boolean)
            _getDataDescription = value
        End Set
    End Property

    Private _dataDescription As String = "" 'A description of the dataset contained in the file.
    Property DataDescription As String
        Get
            Return _dataDescription
        End Get
        Set(value As String)
            _dataDescription = value
        End Set
    End Property

    Private _entryError As Boolean = False 'EntryError is True if an error was found in the entry fields. This will be due to the file name already in use or an incorrect file extension.
    Property EntryError As Boolean
        Get
            Return _entryError
        End Get
        Set(value As Boolean)
            _entryError = value
        End Set
    End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '-----------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings as an XML document.

        Dim Settings = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Form settings for Project form.-->
                       <FormSettings>
                           <Left><%= Me.Left %></Left>
                           <Top><%= Me.Top %></Top>
                           <Width><%= Me.Width %></Width>
                           <!---->
                       </FormSettings>

        '   <Height><%= Me.Height %></Height>

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & "_" & EntryName & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & "DataNameEntry" & "_" & EntryName & ".xml"
        SettingsLocn.SaveXmlData(SettingsFileName, Settings)

    End Sub

    Public Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & "_" & EntryName & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & "DataNameEntry" & "_" & EntryName & ".xml"

        Dim Settings As System.Xml.Linq.XDocument

        SettingsLocn.ReadXmlData(SettingsFileName, Settings)

        If Settings Is Nothing Then
            Exit Sub
        End If

        'Restore form position and size:
        If Settings.<FormSettings>.<Left>.Value = Nothing Then
            'Form setting not saved.
        Else
            Me.Left = Settings.<FormSettings>.<Left>.Value
        End If

        If Settings.<FormSettings>.<Top>.Value = Nothing Then
            'Form setting not saved.
        Else
            Me.Top = Settings.<FormSettings>.<Top>.Value
        End If

        'If Settings.<FormSettings>.<Height>.Value = Nothing Then
        '    'Form setting not saved.
        'Else
        '    Me.Height = Settings.<FormSettings>.<Height>.Value
        'End If

        If Settings.<FormSettings>.<Width>.Value = Nothing Then
            'Form setting not saved.
        Else
            Me.Width = Settings.<FormSettings>.<Width>.Value
        End If
        CheckFormPos()
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New System.Drawing.Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As System.Drawing.Rectangle = System.Windows.Forms.Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
        End If
    End Sub

#End Region 'Process XML Files --------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '--------------------------------------------------------------------------------------------------------------------------

    Private Sub frmProject_Load(sender As Object, e As EventArgs) Handles Me.Load

        'The application that opens this form also initialises the form.

        'ShowDataEntryBoxes()
        RestoreFormSettings()
        ShowDataEntryBoxes()
        Me.Text = Title
    End Sub

    'Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
    '    SaveFormSettings() 'Save the form settings.
    '    Me.Close()
    'End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        'File selection has been cancelled
        SaveFormSettings()
        FileName = ""
        DataName = ""
        DataDescription = ""
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        'New file information has been entered.

        'FileName = txtNewFileName.Text.Trim
        SaveFormSettings()

        If EntryError = False Then
            DataName = txtDataName.Text.Trim
            DataLabel = txtDataLabel.Text.Trim
            DataDescription = txtDataDescr.Text.Trim
            Me.DialogResult = DialogResult.OK
        Else
            Beep()
            'RaiseEvent ErrorMessage("The specified file name is not valid." & vbCrLf)
            txtWarning.Text = "The specified file name is not valid."
        End If
    End Sub

#End Region 'Form Display Subroutines -------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Form methods - The main actions performed by this form" ' --------------------------------------------------------------------------------------------------------------------------

    'Public Sub GetFileList()
    '    'Get the list of entries in the archive.

    '    ListBox1.Items.Clear()

    '    'If ZipArchivePath = "" Then
    '    '    RaiseEvent ErrorMessage("No zip archive selected." & vbCrLf)
    '    '    Exit Sub
    '    'End If

    '    Dim Zip As New ZipComp
    '    Zip.ArchivePath = ZipArchivePath
    '    Dim ValidExtension As String = FileExtension
    '    If ValidExtension.StartsWith(".") Then
    '    Else
    '        ValidExtension = "." & ValidExtension
    '    End If
    '    Dim Index As Integer
    '    Dim EntryName As String
    '    For Index = 1 To Zip.NEntries
    '        If FileExtension = "" Then
    '            ListBox1.Items.Add(Zip.GetEntryName(Index - 1))
    '        Else
    '            EntryName = Zip.GetEntryName(Index - 1)
    '            If EntryName.EndsWith(ValidExtension) Then
    '                ListBox1.Items.Add(Zip.GetEntryName(Index - 1))
    '            End If
    '        End If
    '    Next
    'End Sub

    'Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
    '    RaiseEvent FileSelected(ListBox1.Text)
    'End Sub

    Public Sub ShowDataEntryBoxes()
        'Show the data entry boxes corresponding to the GetDataName and GetDataDescription property settings.
        If GetFileName = True Then
            Label1.Enabled = True
            Label1.Visible = True
            Label1.Top = 43
            txtNewFileName.Enabled = True
            txtNewFileName.Visible = True
            txtNewFileName.Top = 40
            If GetDataName = True Then 'FileName - Name ...
                Label2.Enabled = True
                Label2.Visible = True
                Label2.Top = 69
                txtDataName.Enabled = True
                txtDataName.Visible = True
                txtDataName.Top = 66
                If GetDataLabel = True Then 'FileName - Name - Label ...
                    Label4.Enabled = True
                    Label4.Visible = True
                    Label4.Top = 95
                    txtDataLabel.Enabled = True
                    txtDataLabel.Visible = True
                    txtDataLabel.Top = 92
                    If GetDataDescription = True Then 'FileName - Name - Label - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        txtDataDescr.Top = 118
                        Me.Height = 297
                        txtDataDescr.Height = 60
                    Else 'FileName - Name - Label - (no Description) 
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 60 - 6
                    End If
                Else 'FileName - Name - (no Label) ...
                    Label4.Enabled = False
                    Label4.Visible = False
                    txtDataLabel.Enabled = False
                    txtDataLabel.Visible = False
                    If GetDataDescription = True Then 'FileName - Name - (no Label) - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121 - 20 - 6
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        'Me.Height = 297 - 20 - 7
                        txtDataDescr.Top = 118 - 20 - 6
                        Me.Height = 297 - 20 - 6
                        txtDataDescr.Height = 60
                    Else 'FileName - Name - (no Label) - (no Decription)
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 20 - 6 - 60 - 6
                    End If
                End If
            Else 'FileName - (no Name) ...
                Label2.Enabled = False
                Label2.Visible = False
                txtDataName.Enabled = False
                txtDataName.Visible = False
                If GetDataLabel = True Then
                    Label4.Enabled = True
                    Label4.Visible = True
                    Label4.Top = 95 - 20 - 6
                    txtDataLabel.Enabled = True
                    txtDataLabel.Visible = True
                    txtDataLabel.Top = 92 - 20 - 6
                    If GetDataDescription = True Then 'FileName - (no Name) - Label - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121 - 20 - 6
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        txtDataDescr.Top = 118 - 20 - 6
                        Me.Height = 297 - 20 - 6
                        txtDataDescr.Height = 60
                    Else 'FileName - (no Name) - Label - (no Description)
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 20 - 6 - 60 - 6
                    End If
                Else
                    Label4.Enabled = False
                    Label4.Visible = False
                    txtDataLabel.Enabled = False
                    txtDataLabel.Visible = False
                    If GetDataDescription = True Then 'FileName - (no Name) - (no Label) - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121 - 20 - 6 - 20 - 6
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        txtDataDescr.Top = 118 - 20 - 6 - 20 - 6
                        Me.Height = 297 - 20 - 6 - 20 - 6
                        txtDataDescr.Height = 60
                    Else 'FileName - (no Name) - (no Label) - (no Description)
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 20 - 6 - 20 - 6 - 60 - 6
                    End If
                End If
            End If

        Else
            Label1.Enabled = False
            Label1.Visible = False
            txtNewFileName.Enabled = False
            txtNewFileName.Visible = False
            If GetDataName = True Then '(no FileName) - Name ...
                Label2.Enabled = True
                Label2.Visible = True
                Label2.Top = 69 - 20 - 6
                txtDataName.Enabled = True
                txtDataName.Visible = True
                txtDataName.Top = 66 - 20 - 6
                If GetDataLabel = True Then '(no FileName) - Name - Label ...
                    Label4.Enabled = True
                    Label4.Visible = True
                    Label4.Top = 95 - 20 - 6
                    txtDataLabel.Enabled = True
                    txtDataLabel.Visible = True
                    txtDataLabel.Top = 92 - 20 - 6
                    If GetDataDescription = True Then '(no FileName) - Name - Label - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121 - 20 - 6
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        txtDataDescr.Top = 118 - 20 - 6
                        Me.Height = 297 - 20 - 6
                        txtDataDescr.Height = 60
                    Else '(no FileName) - Name - Label - (no Description) 
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 20 - 6 - 60 - 6
                    End If
                Else '(no FileName) - Name - (no Label) ...
                    Label4.Enabled = False
                    Label4.Visible = False
                    txtDataLabel.Enabled = False
                    txtDataLabel.Visible = False
                    If GetDataDescription = True Then '(no FileName) - Name - (no Label) - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121 - 20 - 6 - 20 - 6
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        'Me.Height = 297 - 20 - 7
                        txtDataDescr.Top = 118 - 20 - 6 - 20 - 6
                        Me.Height = 297 - 20 - 6 - 20 - 6
                        txtDataDescr.Height = 60
                    Else '(no FileName) - Name - (no Label) - (no Decription)
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 20 - 6 - 20 - 6 - 60 - 6
                    End If
                End If
            Else '(no FileName) - (no Name) ...
                Label2.Enabled = False
                Label2.Visible = False
                txtDataName.Enabled = False
                txtDataName.Visible = False
                If GetDataLabel = True Then
                    Label4.Enabled = True
                    Label4.Visible = True
                    Label4.Top = 95 - 20 - 6 - 20 - 6
                    txtDataLabel.Enabled = True
                    txtDataLabel.Visible = True
                    txtDataLabel.Top = 92 - 20 - 6 - 20 - 6
                    If GetDataDescription = True Then '(no FileName) - (no Name) - Label - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121 - 20 - 6 - 20 - 6
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        txtDataDescr.Top = 118 - 20 - 6 - 20 - 6
                        Me.Height = 297 - 20 - 6 - 20 - 6
                        txtDataDescr.Height = 60
                    Else '(no FileName) - (no Name) - Label - (no Description)
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 20 - 6 - 20 - 6 - 60 - 6
                    End If
                Else
                    Label4.Enabled = False
                    Label4.Visible = False
                    txtDataLabel.Enabled = False
                    txtDataLabel.Visible = False
                    If GetDataDescription = True Then '(no FileName) - (no Name) - (no Label) - Description
                        Label3.Enabled = True
                        Label3.Visible = True
                        Label3.Top = 121 - 20 - 6 - 20 - 6 - 20 - 6
                        txtDataDescr.Enabled = True
                        txtDataDescr.Visible = True
                        txtDataDescr.Top = 118 - 20 - 6 - 20 - 6 - 20 - 6
                        Me.Height = 297 - 20 - 6 - 20 - 6 - 20 - 6
                        txtDataDescr.Height = 60
                    Else '(no FileName) - (no Name) - (no Label) - (no Description)
                        Label3.Enabled = False
                        Label3.Visible = False
                        txtDataDescr.Enabled = False
                        txtDataDescr.Visible = False
                        Me.Height = 297 - 20 - 6 - 20 - 6 - 20 - 6 - 60 - 6
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub CheckFileName()
        'Check the file name entered in txtFileName

        Dim LCaseExt As String = LCase(FileExtension.Trim)
        'Dim FileName As String = txtNewFileName.Text.Trim
        FileName = txtNewFileName.Text.Trim

        If LCase(FileName).EndsWith("." & LCaseExt) Then
            FileName = IO.Path.GetFileNameWithoutExtension(FileName) & "." & FileExtension.Trim
            txtNewFileName.Text = FileName
            txtWarning.Text = ""
            EntryError = False
        ElseIf FileName.Contains(".") Then
            txtWarning.Text = "Incorrect file extension!"
            EntryError = True
        Else
            FileName = FileName & "." & FileExtension.Trim
            txtNewFileName.Text = FileName
            txtWarning.Text = ""
            EntryError = False
        End If

        If EntryError = False Then
            'Check if the file name is blank:
            If FileName = "." & FileExtension Then
                txtWarning.Text = "The file name is blank!"
                EntryError = True
            Else
                'Check if the file name is in use:
                If DataLocn.FileExists(FileName) Then
                    txtWarning.Text = "The file name is already in use!"
                    EntryError = True
                Else
                    'The file name is available for use.
                End If
            End If
        End If
    End Sub

    Private Sub txtNewFileName_TextChanged(sender As Object, e As EventArgs) Handles txtNewFileName.TextChanged

    End Sub

    Private Sub txtNewFileName_LostFocus(sender As Object, e As EventArgs) Handles txtNewFileName.LostFocus
        CheckFileName()
    End Sub



#End Region 'Form Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event FileSelected(ByVal FileName As String) 'Raise an event passing the name of the selected file.
    Public Event ErrorMessage(ByVal Message As String) 'Send an error message.
    'Public Event Message(ByVal Message As String) 'Send a message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class