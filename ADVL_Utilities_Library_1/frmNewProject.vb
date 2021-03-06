﻿Public Class frmNewProject
    'The New Project form is used to create a new project for an Andorville Labs application.

#Region " Variable Declarations - All the variables used in this form." '--------------------------------------------------------------------------------------------------------------------

    'Dim NewProjectInfo As New ADVL_Utilities.ProjectInfo 'This object stores information about the selected project.
    Dim NewProject As New ADVL_Utilities_Library_1.Project 'This object stores information about the selected project.

    Public ApplicationSummary As New ApplicationSummary

    'Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public ProjectLocn As ADVL_Utilities_Library_1.FileLocation 'The project location - used to store form settings.
    Public ApplicationName As String 'The name of the application using the message form.

    'Public SettingsLocn As New Location 'This is a directory or archive where settings are stored. NOTE: USE ApplicationDir TYO STORE SETTINGS.

#End Region 'Variable Declarations -----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form" '--------------------------------------------------------------------------------------------------------------------------------

    Private _applicationDir As String = "" 'The path to the directory used to store application data.
    Public Property ApplicationDir As String
        Get
            Return _applicationDir
        End Get
        Set(ByVal value As String)
            _applicationDir = value
            ApplicationSummary.ReadFile(_applicationDir) 'Update the Application Summary by reading the Application_Info.xml in the Application Directory.
        End Set
    End Property

    Private _projectDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments 'The default directory to create a new project.
    Property ProjectDir As String
        Get
            Return _projectDir
        End Get
        Set(value As String)
            _projectDir = value
            txtDirectoryPath.Text = _projectDir
        End Set
    End Property

    Private _projectFileDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments 'The default directory to create a new project file.
    Property ProjectFileDir As String
        Get
            Return _projectFileDir
        End Get
        Set(value As String)
            _projectFileDir = value
            txtPFDirectoryPath.Text = _projectFileDir
        End Set
    End Property

    Private _hybridProjectDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments 'The default directory to create a new Hybrid project.
    Property HybridProjectDir As String
        Get
            Return _hybridProjectDir
        End Get
        Set(value As String)
            _hybridProjectDir = value
            txtHPDirectoryPath.Text = _hybridProjectDir
        End Set
    End Property

    Private _dataFileDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments 'The default directory to create a new Data file in a Hybrid project.
    Property DataFileDir As String
        Get
            Return _dataFileDir
        End Get
        Set(value As String)
            _dataFileDir = value
        End Set
    End Property

    'Private _dataFileInProjectDir As Boolean = True 'If True, the DataFile of a Hybrid Project will be placed in the Project Directory
    'Property DataFileInProjectDir As Boolean
    '    Get
    '        Return _dataFileInProjectDir
    '    End Get
    '    Set(value As Boolean)
    '        _dataFileInProjectDir = value
    '        If _dataFileInProjectDir = True Then
    '            chkProjectDir.Checked = True
    '        Else
    '            chkProjectDir.Checked = False
    '        End If
    '    End Set
    'End Property

#End Region 'Properties ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Process XML files - Read and write XML files." '------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings as an XML document.

        Dim ProjectType As String = ""
        If rbProjectDir.Checked Then
            ProjectType = "Directory"
        ElseIf rbProjectFile.Checked Then
            ProjectType = "File"
        ElseIf rbHybrid.Checked Then
            ProjectType = "Hybrid"
        End If

        Dim Settings = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Form settings for Project form.-->
                       <FormSettings>
                           <Left><%= Me.Left %></Left>
                           <Top><%= Me.Top %></Top>
                           <Width><%= Me.Width %></Width>
                           <Height><%= Me.Height %></Height>
                           <ProjectDir><%= ProjectDir %></ProjectDir>
                           <ProjectFileDir><%= ProjectFileDir %></ProjectFileDir>
                           <HybridProjectDir><%= HybridProjectDir %></HybridProjectDir>
                           <DataFileDir><%= DataFileDir %></DataFileDir>
                           <ProjectType><%= ProjectType %></ProjectType>
                           <!---->
                       </FormSettings>

        '                           <DataFileInProjectDir><%= DataFileInProjectDir %></DataFileInProjectDir>
        'Dim SettingsName As String = ApplicationSummary.Name & "_" & "FormSettings_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        'Settings.Save(ApplicationDir & "\" & SettingsName)
        'SettingsLocn.SaveXmlData(SettingsFileName, Settings)
        ProjectLocn.SaveXmlData(SettingsFileName, Settings)

    End Sub

    Public Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsName As String = ApplicationSummary.Name & "_" & "FormSettings_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"

        'If System.IO.File.Exists(ApplicationDir & "\" & SettingsName) Then
        Dim Settings As System.Xml.Linq.XDocument
        'Settings = XDocument.Load(ApplicationDir & "\" & SettingsName)
        'SettingsLocn.ReadXmlData(SettingsFileName, Settings)
        ProjectLocn.ReadXmlData(SettingsFileName, Settings)

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

        If Settings.<FormSettings>.<Height>.Value = Nothing Then
            'Form setting not saved.
        Else
            Me.Height = Settings.<FormSettings>.<Height>.Value
        End If

        If Settings.<FormSettings>.<Width>.Value = Nothing Then
            'Form setting not saved.
        Else
            Me.Width = Settings.<FormSettings>.<Width>.Value
        End If

        If Settings.<FormSettings>.<ProjectDir>.Value = Nothing Then
            'Form setting not saved.
        Else
            ProjectDir = Settings.<FormSettings>.<ProjectDir>.Value
        End If

        If Settings.<FormSettings>.<ProjectFileDir>.Value = Nothing Then
            'Form setting not saved.
        Else
            ProjectFileDir = Settings.<FormSettings>.<ProjectFileDir>.Value
        End If

        If Settings.<FormSettings>.<HybridProjectDir>.Value = Nothing Then
            'Form setting not saved.
        Else
            HybridProjectDir = Settings.<FormSettings>.<HybridProjectDir>.Value
        End If

        If Settings.<FormSettings>.<DataFileDir>.Value = Nothing Then
            'Form setting not saved.
        Else
            DataFileDir = Settings.<FormSettings>.<DataFileDir>.Value
        End If

        'If Settings.<FormSettings>.<DataFileInProjectDir>.Value = Nothing Then
        '    'Form setting not saved.
        'Else
        '    Select Case Settings.<FormSettings>.<DataFileInProjectDir>.Value
        '        Case "true"
        '            DataFileInProjectDir = True
        '        Case "false"
        '            DataFileInProjectDir = False
        '    End Select
        'End If

        If Settings.<FormSettings>.<ProjectType>.Value = Nothing Then
            'Form setting not saved.
        Else
            Select Case Settings.<FormSettings>.<ProjectType>.Value
                Case "Directory"
                    rbProjectDir.Checked = True
                    ChangeProjectType()
                Case "File"
                    rbProjectFile.Checked = True
                    ChangeProjectType()
                Case "Hybrid"
                    rbHybrid.Checked = True
                    ChangeProjectType()
            End Select
        End If
        CheckFormPos()

        'Else
        ''Settings file not found.
        'End If

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

#End Region 'Process XML Files ---------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        SaveFormSettings() 'Save the form settings.
        Me.Close()
    End Sub


    Private Sub frmNewProject_Load(sender As Object, e As EventArgs) Handles Me.Load
        rbProjectDir.Checked = True

        'Set Up Hybrid Project options:
        txtHPSettingsName.Text = "Settings"
        cmbHPSettingsType.Items.Clear()
        cmbHPSettingsType.Items.Add("Directory")
        cmbHPSettingsType.Items.Add("Archive")
        cmbHPSettingsType.SelectedIndex = 0

        txtHPDataName.Text = "Data"
        cmbHPDataType.Items.Clear()
        cmbHPDataType.Items.Add("Directory")
        cmbHPDataType.Items.Add("Archive")
        cmbHPDataType.SelectedIndex = 1

        txtHPSystemName.Text = "System"
        cmbHPSystemType.Items.Clear()
        cmbHPSystemType.Items.Add("Directory")
        cmbHPSystemType.Items.Add("Archive")
        cmbHPSystemType.SelectedIndex = 1

    End Sub


#End Region 'Form Display Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Form methods - The main actions performed by this form" ' ---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        'Create the new project:
        If TabControl1.SelectedTab Is TabPage1 Then 'A new Project Directory will be created. ----------------------------------------------------------------

            Dim DirectoryPath As String = txtDirectoryPath.Text 'The new project directory will be created in this directory.
            'Check if the ProjectPath exists:
            If System.IO.Directory.Exists(DirectoryPath) Then
                'ProjectPath exists.
            Else
                RaiseEvent CreateProjectError("The new project cannot be created. The specified directory does not exist: " & DirectoryPath)
                Exit Sub
            End If

            Dim NewProjectDirectoryName As String = txtProjectDirectoryName.Text 'The name of the new project directory.
            Dim NewProjectDirectoryPath As String = DirectoryPath & "\" & NewProjectDirectoryName
            'Check if the NewProjectDirectoryPath already exists:
            If System.IO.Directory.Exists(NewProjectDirectoryPath) Then
                RaiseEvent CreateProjectError("The new project cannot be created. The specified new project directory already exists: " & NewProjectDirectoryPath)
                Exit Sub
            Else
                'Check if a file exists at that path:
                If System.IO.File.Exists(NewProjectDirectoryPath) Then
                    RaiseEvent CreateProjectError("The new project cannot be created. A file with the same name as the project directory already exists: " & NewProjectDirectoryPath)
                    Exit Sub
                End If
                System.IO.Directory.CreateDirectory(NewProjectDirectoryPath) 'The new project directory has been created.
            End If

            NewProject.Type = Project.Types.Directory 'ProjectInfo.Types.Directory
            NewProject.Path = NewProjectDirectoryPath '29Jul18
            NewProject.Name = txtProjectName.Text
            NewProject.Description = txtProjectDescription.Text
            NewProject.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")

            'Added 24Aug18:
            Dim IDString As String = NewProject.Name & " " & Format(NewProject.CreationDate, "d-MMM-yyyy H:mm:ss")
            NewProject.ID = IDString.GetHashCode

            NewProject.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
            NewProject.SettingsLocn.Path = NewProjectDirectoryPath
            NewProject.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
            NewProject.DataLocn.Path = NewProjectDirectoryPath

            'Add 3Nov18:
            NewProject.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
            NewProject.SystemLocn.Path = NewProjectDirectoryPath

            NewProject.Author.Name = txtAuthorName.Text
            NewProject.Author.Description = txtAuthorDescription.Text
            NewProject.Author.Contact = txtAuthorContact.Text
            'NewProject.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
            'NewProject.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            'NewProject.Usage.FirstUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            NewProject.Usage.FirstUsed = NewProject.CreationDate
            NewProject.Usage.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")


        ElseIf TabControl1.SelectedTab Is TabPage2 Then 'A new Project Archive File will be created. --------------------------------------------------------------
            Dim DirectoryPath As String = txtPFDirectoryPath.Text 'The new project file will be created in this directory.
            'Check if the ProjectPath exists:
            If System.IO.Directory.Exists(DirectoryPath) Then
                'ProjectPath exists.
            Else
                RaiseEvent CreateProjectError("The new project cannot be created. The specified directory does not exist: " & DirectoryPath)
                Exit Sub
            End If

            Dim NewProjectFileName As String = txtProjectFileName.Text 'The name of the new project file.
            Dim NewProjectFilePath As String = DirectoryPath & "\" & NewProjectFileName & ".AdvlProject"
            'Check if the NewProjectFilePath already exists:
            'If System.IO.Directory.Exists(NewProjectFilePath) Then
            If System.IO.File.Exists(NewProjectFilePath) Then
                RaiseEvent CreateProjectError("The new project cannot be created. The specified new project file already exists: " & NewProjectFilePath)
                Exit Sub
            Else
                'System.IO.Directory.CreateDirectory(NewProjectFilePath) 'The new project directory has been created.
                Dim Zip As New ZipComp
                Zip.NewArchivePath = NewProjectFilePath
                Zip.CreateArchive() 'The new project file has been created.
            End If

            NewProject.Type = Project.Types.Archive 'ProjectInfo.Types.Archive
            NewProject.Path = NewProjectFilePath '29Jul18
            NewProject.Name = txtProjectName.Text
            NewProject.Description = txtProjectDescription.Text
            NewProject.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")

            'Added 24Aug18:
            Dim IDString As String = NewProject.Name & " " & Format(NewProject.CreationDate, "d-MMM-yyyy H:mm:ss")
            NewProject.ID = IDString.GetHashCode

            NewProject.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
            NewProject.SettingsLocn.Path = NewProjectFilePath
            NewProject.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
            NewProject.DataLocn.Path = NewProjectFilePath

            'Add 3Nov18:
            NewProject.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
            NewProject.SystemLocn.Path = NewProjectFilePath

            NewProject.Author.Name = txtAuthorName.Text
            NewProject.Author.Description = txtAuthorDescription.Text
            NewProject.Author.Contact = txtAuthorContact.Text
            'NewProject.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
            'NewProject.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            'NewProject.Usage.FirstUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            NewProject.Usage.FirstUsed = NewProject.CreationDate
            NewProject.Usage.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")


        ElseIf TabControl1.SelectedTab Is TabPage3 Then 'A new hybrid project will be created. ------------------------------------------------------------
            'Dim ProjectPath As String = txtHPDirectoryPath.Text 'The new project directory will be created in this directory.
            ''Check if the ProjectPath exists:
            'If System.IO.Directory.Exists(ProjectPath) Then
            '    'ProjectPath exists.
            'Else
            '    RaiseEvent CreateProjectError("The new project cannot be created. The specified directory does not exist: " & ProjectPath)
            '    Exit Sub
            'End If

            'Dim NewProjectDirectoryName As String = txtHPDirectoryName.Text 'The name of the new project directory.
            'Dim NewProjectDirectoryPath As String = ProjectPath & "\" & NewProjectDirectoryName
            ''Check if the NewProjectDirectoryPath already exists:
            'If System.IO.Directory.Exists(NewProjectDirectoryPath) Then
            '    RaiseEvent CreateProjectError("The new project cannot be created. The specified new project directory already exists: " & NewProjectDirectoryPath)
            '    Exit Sub
            'Else
            '    System.IO.Directory.CreateDirectory(NewProjectDirectoryPath) 'The new project directory has been created.
            'End If

            ''Create the Data File: --------------------------------------------------------------------------------------------------------------------------------
            'Dim NewDataFileName As String 'The name of the new project file.
            'Dim NewDataFilePath As String
            'Dim DataFileDirectory As String = txtDataFileDirectory.Text 'The new project file will be created in this directory.
            ''Check if the DataFileDirectory exists:
            'If System.IO.Directory.Exists(DataFileDirectory) Then
            '    'ProjectPath exists.
            '    NewDataFileName = txPDataFileName.Text 'The name of the new project file.
            '    NewDataFilePath = DataFileDirectory & "\" & NewDataFileName & ".AdvlData"
            '    'Check if the NewDataFilePath already exists:
            '    If System.IO.Directory.Exists(NewDataFilePath) Then
            '        RaiseEvent CreateProjectError("The new project cannot be created. The specified new project file already exists: " & NewDataFilePath)
            '        Exit Sub
            '    Else
            '        Dim Zip As New ZipComp
            '        Zip.NewArchivePath = NewDataFilePath
            '        Zip.CreateArchive() 'The new data file has been created.
            '    End If
            'Else
            '    RaiseEvent CreateProjectError("The new data file cannot be created. The specified directory does not exist: " & DataFileDirectory)
            '    DataFileDirectory = ""
            '    NewDataFileName = ""
            '    NewDataFilePath = ""
            '    'The new hybrid project will be created without a data file.
            'End If

            'NewProject.Type = Project.Types.Hybrid 'ProjectInfo.Types.Hybrid
            'NewProject.Path = NewProjectDirectoryPath '29Jul18
            'NewProject.Name = txtProjectName.Text
            'NewProject.Description = txtProjectDescription.Text
            'NewProject.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")

            ''Added 24Aug18:
            'Dim IDString As String = NewProject.Name & " " & Format(NewProject.CreationDate, "d-MMM-yyyy H:mm:ss")
            'NewProject.ID = IDString.GetHashCode

            'NewProject.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
            'NewProject.SettingsLocn.Path = NewProjectDirectoryPath
            'NewProject.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
            'NewProject.DataLocn.Path = NewDataFilePath
            'NewProject.Author.Name = txtAuthorName.Text
            'NewProject.Author.Description = txtAuthorDescription.Text
            'NewProject.Author.Contact = txtAuthorContact.Text
            'NewProject.Usage.FirstUsed = NewProject.CreationDate
            'NewProject.Usage.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")


            Dim ProjectPath As String = txtHPDirectoryPath.Text 'The new project directory will be created in this directory.
            'Check if the ProjectPath exists:
            If System.IO.Directory.Exists(ProjectPath) Then
                'ProjectPath exists.
            Else
                RaiseEvent CreateProjectError("The new project cannot be created. The specified directory does not exist: " & ProjectPath)
                Exit Sub
            End If

            Dim NewProjectDirectoryName As String = txtHPDirectoryName.Text 'The name of the new project directory.
            Dim NewProjectDirectoryPath As String = ProjectPath & "\" & NewProjectDirectoryName
            'Check if the NewProjectDirectoryPath already exists:
            If System.IO.Directory.Exists(NewProjectDirectoryPath) Then
                RaiseEvent CreateProjectError("The new project cannot be created. The specified new project directory already exists: " & NewProjectDirectoryPath)
                Exit Sub
            Else
                System.IO.Directory.CreateDirectory(NewProjectDirectoryPath) 'The new project directory has been created.
            End If

            'OK to create the new Hybrid project:
            'Dim NewProject As New ADVL_Utilities_Library_1.Project 'This object stores information about the selected project. THIS IS DECLARED AT THE START OF THE FORM CODE!

            NewProject.Type = Project.Types.Hybrid
            NewProject.Path = NewProjectDirectoryPath
            'NewProject.RelativePath = txtHybridProjectRelativePath.Text
            NewProject.Name = txtProjectName.Text
            NewProject.Description = txtProjectDescription.Text
            NewProject.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")

            Dim IDString As String = NewProject.Name & " " & Format(NewProject.CreationDate, "d-MMM-yyyy H:mm:ss")
            NewProject.ID = IDString.GetHashCode

            'Set up the settings location:
            If Trim(txtHPSettingsName.Text) = "" Then
                txtHPSettingsName.Text = "Settings"
            End If
            If cmbHPSettingsType.SelectedItem.ToString = "Directory" Then
                'Set up the Settings relative location:
                NewProject.SettingsRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SettingsRelLocn.Path = "\" & Trim(txtHPSettingsName.Text)
                'Set up the Settings location:
                NewProject.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SettingsLocn.Path = NewProjectDirectoryPath & NewProject.SettingsRelLocn.Path
                'Create the Settings Directory:
                System.IO.Directory.CreateDirectory(NewProject.SettingsLocn.Path)
            ElseIf cmbHPSettingsType.SelectedItem.ToString = "Archive" Then
                'Set up the Settings relative location:
                NewProject.SettingsRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                NewProject.SettingsRelLocn.Path = "\" & Trim(txtHPSettingsName.Text) & ".zip"
                'Set up the Settings location:
                NewProject.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                NewProject.SettingsLocn.Path = NewProjectDirectoryPath & NewProject.SettingsRelLocn.Path
                'Create the Settings Archive:
                Dim Zip As New ADVL_Utilities_Library_1.ZipComp
                Zip.NewArchivePath = NewProject.SettingsLocn.Path
                Zip.CreateArchive() 'The new project file has been created.
            Else 'Unknown settings location type.
                'Message.AddWarning("Unknown settings location type: " & cmbHPSettingsType.SelectedItem.ToString & vbCrLf)
                RaiseEvent CreateProjectError("Unknown settings location type: " & cmbHPSettingsType.SelectedItem.ToString & vbCrLf)
                'Message.AddWarning("A settings directory location will be created." & vbCrLf)
                RaiseEvent CreateProjectError("A settings directory location will be created." & vbCrLf)
                'Set up the Settings relative location:
                NewProject.SettingsRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SettingsRelLocn.Path = "\" & Trim(txtHPSettingsName.Text)
                'Set up the Settings location:
                NewProject.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SettingsLocn.Path = NewProjectDirectoryPath & NewProject.SettingsRelLocn.Path
                'Create the Settings Directory:
                System.IO.Directory.CreateDirectory(NewProject.SettingsLocn.Path)
            End If

            'Set up the Data Location:
            If Trim(txtHPDataName.Text) = "" Then
                txtHPDataName.Text = "Data"
            End If
            If cmbHPDataType.SelectedItem.ToString = "Directory" Then
                'Set up the Data relative location:
                NewProject.DataRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.DataRelLocn.Path = "\" & Trim(txtHPDataName.Text)
                'Set up the Data location:
                NewProject.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.DataLocn.Path = NewProjectDirectoryPath & NewProject.DataRelLocn.Path
                'Create the Data Directory:
                System.IO.Directory.CreateDirectory(NewProject.DataLocn.Path)
            ElseIf cmbHPDataType.SelectedItem.ToString = "Archive" Then
                'Set up the Data relative location:
                NewProject.DataRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                NewProject.DataRelLocn.Path = "\" & Trim(txtHPDataName.Text) & ".zip"
                'Set up the Data location:
                NewProject.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                NewProject.DataLocn.Path = NewProjectDirectoryPath & NewProject.DataRelLocn.Path
                'Create the Data Archive:
                Dim Zip As New ADVL_Utilities_Library_1.ZipComp
                Zip.NewArchivePath = NewProject.DataLocn.Path
                Zip.CreateArchive() 'The new project file has been created.
            Else 'Unknown Data location type.
                'Message.AddWarning("Unknown settings location type: " & cmbHPDataType.SelectedItem.ToString & vbCrLf)
                RaiseEvent CreateProjectError("Unknown settings location type: " & cmbHPDataType.SelectedItem.ToString & vbCrLf)
                'Message.AddWarning("A settings directory location will be created." & vbCrLf)
                RaiseEvent CreateProjectError("A settings directory location will be created." & vbCrLf)
                'Set up the Data relative location:
                NewProject.DataRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.DataRelLocn.Path = "\" & Trim(txtHPDataName.Text)
                'Set up the Data location:
                NewProject.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.DataLocn.Path = NewProjectDirectoryPath & NewProject.DataRelLocn.Path
                'Create the Data Directory:
                System.IO.Directory.CreateDirectory(NewProject.DataLocn.Path)
            End If

            'Set up the System Location:
            If Trim(txtHPSystemName.Text) = "" Then
                txtHPSystemName.Text = "System"
            End If
            If cmbHPSystemType.SelectedItem.ToString = "Directory" Then
                'Set up the System relative location:
                NewProject.SystemRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SystemRelLocn.Path = "\" & Trim(txtHPSystemName.Text)
                'Set up the System location:
                NewProject.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SystemLocn.Path = NewProjectDirectoryPath & NewProject.SystemRelLocn.Path
                'Create the System Directory:
                System.IO.Directory.CreateDirectory(NewProject.SystemLocn.Path)
            ElseIf cmbHPSystemType.SelectedItem.ToString = "Archive" Then
                'Set up the System relative location:
                NewProject.SystemRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                NewProject.SystemRelLocn.Path = "\" & Trim(txtHPSystemName.Text) & ".zip"
                'Set up the System location:
                NewProject.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                NewProject.SystemLocn.Path = NewProjectDirectoryPath & NewProject.SystemRelLocn.Path
                'Create the System Archive:
                Dim Zip As New ADVL_Utilities_Library_1.ZipComp
                Zip.NewArchivePath = NewProject.SystemLocn.Path
                Zip.CreateArchive() 'The new project file has been created.
            Else 'Unknown System location type.
                'Message.AddWarning("Unknown settings location type: " & cmbHPSystemType.SelectedItem.ToString & vbCrLf)
                RaiseEvent CreateProjectError("Unknown settings location type: " & cmbHPSystemType.SelectedItem.ToString & vbCrLf)
                'Message.AddWarning("A settings directory location will be created." & vbCrLf)
                RaiseEvent CreateProjectError("A settings directory location will be created." & vbCrLf)
                'Set up the System relative location:
                NewProject.SystemRelLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SystemRelLocn.Path = "\" & Trim(txtHPSystemName.Text)
                'Set up the System location:
                NewProject.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                NewProject.SystemLocn.Path = NewProjectDirectoryPath & NewProject.SystemRelLocn.Path
                'Create the System Directory:
                System.IO.Directory.CreateDirectory(NewProject.SystemLocn.Path)
            End If

            NewProject.Author.Name = txtAuthorName.Text
            NewProject.Author.Description = txtAuthorDescription.Text
            NewProject.Author.Contact = txtAuthorContact.Text
            NewProject.Usage.FirstUsed = NewProject.CreationDate
            NewProject.Usage.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")

        Else
            'A valid project type was not selected.
            Exit Sub
        End If

        'Find the Application Info: ---------------------------------------------------------------------------------------------------------------------------------------------------------
        'SEE UPDATED CODE BELOW!!!
        ''Check if the Application Info file exists:
        ''If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.xml") Then
        'If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info_ADVL_2.xml") Then

        'Else
        '    RaiseEvent CreateProjectError("The Application Information file (Application_Info.xml) is missing from the Application Directory: " & ApplicationDir)
        '    Exit Sub
        'End If

        ''Read the Application Information:
        ''Dim ApplicationInfo As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Application_Info.xml")
        'Dim ApplicationInfo As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Application_Info_ADVL_2.xml")

        'If ApplicationInfo.<Application>.<Name>.Value = Nothing Then
        '    'NewProject.ApplicationSummary.Name = ""
        '    'NewProject.HostApplication.Name = ""
        '    NewProject.Application.Name = ""
        'Else
        '    'NewProject.HostApplication.Name = ApplicationInfo.<Application>.<Name>.Value
        '    NewProject.Application.Name = ApplicationInfo.<Application>.<Name>.Value
        'End If

        'If ApplicationInfo.<Application>.<Description>.Value = Nothing Then
        '    'NewProject.HostApplication.Description = ""
        '    NewProject.Application.Description = ""
        'Else
        '    'NewProject.HostApplication.Description = ApplicationInfo.<Application>.<Description>.Value
        '    NewProject.Application.Description = ApplicationInfo.<Application>.<Description>.Value
        'End If

        'If ApplicationInfo.<Application>.<CreationDate>.Value = Nothing Then
        '    'NewProject.HostApplication.CreationDate = "1-Jan-2000 12:00:00"
        '    NewProject.Application.CreationDate = "1-Jan-2000 12:00:00"
        'Else
        '    'NewProject.ApplicationSummary.CreationDate = ApplicationInfo.<Application>.<CreationDate>.Value
        '    'NewProject.ApplicationSummary.CreationDate = Format(ApplicationInfo.<Application>.<CreationDate>.Value, "d-MMM-yyyy H:mm:ss") 'ERROR
        '    'NewProject.HostApplication.CreationDate = ApplicationInfo.<Application>.<CreationDate>.Value
        '    NewProject.Application.CreationDate = ApplicationInfo.<Application>.<CreationDate>.Value
        'End If

        'If ApplicationInfo.<Application>.<Version>.<Major>.Value = Nothing Then
        '    'NewProject.HostApplication.Version.Major = 1
        '    NewProject.Application.Version.Major = 1
        'Else
        '    'NewProject.HostApplication.Version.Major = ApplicationInfo.<Application>.<Version>.<Major>.Value
        '    NewProject.Application.Version.Major = ApplicationInfo.<Application>.<Version>.<Major>.Value
        'End If

        'If ApplicationInfo.<Application>.<Version>.<Minor>.Value = Nothing Then
        '    'NewProject.HostApplication.Version.Minor = 0
        '    NewProject.Application.Version.Minor = 0
        'Else
        '    'NewProject.HostApplication.Version.Minor = ApplicationInfo.<Application>.<Version>.<Minor>.Value
        '    NewProject.Application.Version.Minor = ApplicationInfo.<Application>.<Version>.<Minor>.Value
        'End If

        'If ApplicationInfo.<Application>.<Version>.<Build>.Value = Nothing Then
        '    'NewProject.HostApplication.Version.Build = 1
        '    NewProject.Application.Version.Build = 1
        'Else
        '    'NewProject.HostApplication.Version.Build = ApplicationInfo.<Application>.<Version>.<Build>.Value
        '    NewProject.Application.Version.Build = ApplicationInfo.<Application>.<Version>.<Build>.Value
        'End If

        'If ApplicationInfo.<Application>.<Version>.<Revision>.Value = Nothing Then
        '    'NewProject.HostApplication.Version.Revision = 0
        '    NewProject.Application.Version.Revision = 0
        'Else
        '    'NewProject.HostApplication.Version.Revision = ApplicationInfo.<Application>.<Version>.<Revision>.Value
        '    NewProject.Application.Version.Revision = ApplicationInfo.<Application>.<Version>.<Revision>.Value
        'End If

        'If ApplicationInfo.<Application>.<Author>.<Name>.Value = Nothing Then
        '    'NewProject.HostApplication.Author.Name = ""
        '    NewProject.Application.Author.Name = ""
        'Else
        '    'NewProject.HostApplication.Author.Name = ApplicationInfo.<Application>.<Author>.<Name>.Value
        '    NewProject.Application.Author.Name = ApplicationInfo.<Application>.<Author>.<Name>.Value
        'End If

        'If ApplicationInfo.<Application>.<Author>.<Description>.Value = Nothing Then
        '    'NewProject.HostApplication.Author.Description = ""
        '    NewProject.Application.Author.Description = ""
        'Else
        '    'NewProject.HostApplication.Author.Description = ApplicationInfo.<Application>.<Author>.<Description>.Value
        '    NewProject.Application.Author.Description = ApplicationInfo.<Application>.<Author>.<Description>.Value
        'End If

        'If ApplicationInfo.<Application>.<Author>.<Contact>.Value = Nothing Then
        '    'NewProject.HostApplication.Author.Contact = ""
        '    NewProject.Application.Author.Contact = ""
        'Else
        '    'NewProject.HostApplication.Author.Contact = ApplicationInfo.<Application>.<Author>.<Contact>.Value
        '    NewProject.Application.Author.Contact = ApplicationInfo.<Application>.<Author>.<Contact>.Value
        'End If

        'UPDATED CODE:
        Dim NewProjectAppInfo As New ADVL_Utilities_Library_1.ApplicationInfo
        NewProjectAppInfo.ApplicationDir = ApplicationDir
        NewProjectAppInfo.ReadFile()
        NewProject.ApplicationDir = ApplicationDir
        NewProject.Application.Name = NewProjectAppInfo.Name
        NewProject.Application.Description = NewProjectAppInfo.Description
        NewProject.Application.CreationDate = NewProjectAppInfo.CreationDate
        NewProject.Application.Version.Major = NewProjectAppInfo.Version.Major
        NewProject.Application.Version.Minor = NewProjectAppInfo.Version.Minor
        NewProject.Application.Version.Build = NewProjectAppInfo.Version.Build
        NewProject.Application.Version.Revision = NewProjectAppInfo.Version.Revision
        NewProject.Application.Author.Name = NewProjectAppInfo.Author.Name
        NewProject.Application.Author.Description = NewProjectAppInfo.Author.Description
        NewProject.Application.Author.Contact = NewProjectAppInfo.Author.Contact

        'Save the Project Information:
        NewProject.SaveProjectInfoFile()

        'Return the new project summary to the Project form:
        Dim NewProjectSummary As New ADVL_Utilities_Library_1.ProjectSummary

        'Name
        'Description
        'Type
        'CreationDate
        'SettingsLocationType
        'SettingsLocationPath
        'DataLocationType
        'DataLocationPath
        'AuthorName
        'ApplicationName
        'Extension

        NewProjectSummary.Name = NewProject.Name
        NewProjectSummary.Description = NewProject.Description
        NewProjectSummary.Type = NewProject.Type
        'NewProjectSummary.CreationDate = NewProject.CreationDate

        'NewProjectSummary.SettingsLocnType = NewProject.SettingsLocn.Type
        'NewProjectSummary.SettingsLocnPath = NewProject.SettingsLocn.Path
        NewProjectSummary.Type = NewProject.Type
        NewProjectSummary.Path = NewProject.Path

        'NewProjectSummary.DataLocnType = NewProject.DataLocn.Type
        'NewProjectSummary.DataLocnPath = NewProject.DataLocn.Path
        NewProjectSummary.AuthorName = NewProject.Author.Name
        'NewProjectSummary.ApplicationName = NewProject.HostApplication.Name
        'NewProjectSummary.Extension = NewProject.ApplicationSummary.Extension

        NewProjectSummary.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")




        RaiseEvent NewProjectCreated(NewProjectSummary) 'Send the NewProjectSummary to the Project form.

        'Close the New Project form:
        SaveFormSettings() 'Save the form settings.
        Me.Close()

    End Sub

    Private Sub cmbHPSettingsType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHPSettingsType.SelectedIndexChanged
        If cmbHPSettingsType.SelectedItem.ToString = "Archive" Then
            Label74.Text = ".zip"
        Else '
            Label74.Text = ""
        End If
    End Sub

    Private Sub cmbHPDataType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHPDataType.SelectedIndexChanged
        If cmbHPDataType.SelectedItem.ToString = "Archive" Then
            Label73.Text = ".zip"
        Else '
            Label73.Text = ""
        End If
    End Sub

    Private Sub cmbHPSystemType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbHPSystemType.SelectedIndexChanged
        If cmbHPSystemType.SelectedItem.ToString = "Archive" Then
            Label75.Text = ".zip"
        Else '
            Label75.Text = ""
        End If
    End Sub


    Private Sub rbProjectDir_Click(sender As Object, e As EventArgs) Handles rbProjectDir.Click
        ChangeProjectType()
    End Sub

    Private Sub rbProjectFile_Click(sender As Object, e As EventArgs) Handles rbProjectFile.Click
        ChangeProjectType()
    End Sub

    Private Sub rbHybrid_Click(sender As Object, e As EventArgs) Handles rbHybrid.Click
        ChangeProjectType()
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        rbProjectDir.Checked = True
    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        rbProjectFile.Checked = True
    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        rbHybrid.Checked = True
    End Sub

    Private Sub ChangeProjectType()
        If rbProjectDir.Checked Then
            TabControl1.SelectedIndex = 0
        ElseIf rbProjectFile.Checked Then
            TabControl1.SelectedIndex = 1
        ElseIf rbHybrid.Checked Then
            TabControl1.SelectedIndex = 2
        End If
    End Sub

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        'Select a directory for the the project:

        FolderBrowserDialog1.SelectedPath = ProjectDir
        FolderBrowserDialog1.ShowDialog()
        ProjectDir = FolderBrowserDialog1.SelectedPath
    End Sub

    'Private Sub chkProjectDir_CheckedChanged(sender As Object, e As EventArgs)
    '    If chkProjectDir.Checked Then
    '        _dataFileInProjectDir = True
    '        txtDataFileDirectory.Enabled = False
    '        btnFindDataFileDir.Enabled = False
    '    Else
    '        _dataFileInProjectDir = True
    '        txtDataFileDirectory.Enabled = True
    '        btnFindDataFileDir.Enabled = True
    '    End If
    'End Sub

    Private Sub btnFindProjFileDir_Click(sender As Object, e As EventArgs) Handles btnFindProjFileDir.Click
        'Select a directory to store the new project file.
        FolderBrowserDialog1.SelectedPath = ProjectFileDir
        FolderBrowserDialog1.ShowDialog()
        ProjectFileDir = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub btnFindHybridProjDir_Click(sender As Object, e As EventArgs) Handles btnFindHybridProjDir.Click
        'Select a directory to store the new hybrid project.
        FolderBrowserDialog1.SelectedPath = HybridProjectDir
        FolderBrowserDialog1.ShowDialog()
        HybridProjectDir = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub btnFindDataFileDir_Click(sender As Object, e As EventArgs)
        'Select a directory to store the new hybrid project.
        FolderBrowserDialog1.SelectedPath = DataFileDir
        FolderBrowserDialog1.ShowDialog()
        DataFileDir = FolderBrowserDialog1.SelectedPath
    End Sub


#End Region 'Form Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Event CreateProjectError(ByVal ErrorMessage As String) 'Raise an event if there is an error while creating a new project.
    Public Event NewProjectCreated(ByVal ProjectSummary As ADVL_Utilities_Library_1.ProjectSummary) 'Raise an event to indicate that a new project has been created.



#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------











End Class

