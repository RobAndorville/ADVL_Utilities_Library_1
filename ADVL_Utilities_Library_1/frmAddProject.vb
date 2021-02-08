

Public Class frmAddProject
    'The Add Project form is used to add an existing project to the Project List.

#Region " Variable Declarations - All the variables used in this form." '---------------------------------------------------------------------------------------------------------------------

    Public ApplicationSummary As New ApplicationSummary

    'Dim folders() As String = IO.Directory.GetDirectories(Dir)

    Dim RootPath As String

    'Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public ProjectLocn As ADVL_Utilities_Library_1.FileLocation 'The Project location - used to store this form's settings.

    'NOTE: ApplicationName is now a property.
    'Public ApplicationName As String 'The name of the application using the message form. This value is set in the method that calls the AddProject form.


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

    Private _applicationName As String 'The name of the application using the message form. This value is set in the method that calls the AddProject form.
    Public Property ApplicationName As String
        Get
            Return _applicationName
        End Get
        Set(value As String)
            _applicationName = value
            txtThisAppName.Text = _applicationName
        End Set
    End Property


    Private _searchDirectory As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments   'The directory used to search for project files.
    Property SearchDirectory As String
        Get
            Return _searchDirectory
        End Get
        Set(value As String)
            _searchDirectory = value
            txtSearchDirectory.Text = _searchDirectory
        End Set
    End Property

#End Region 'Properties ----------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Process XML files - Read and write XML files." '------------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings as an XML document.
        Dim Settings = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Form settings for Project form.-->
                       <FormSettings>
                           <Left><%= Me.Left %></Left>
                           <Top><%= Me.Top %></Top>
                           <Width><%= Me.Width %></Width>
                           <Height><%= Me.Height %></Height>
                           <SearchDirectory><%= SearchDirectory %></SearchDirectory>
                           <!---->
                       </FormSettings>

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

        If Settings.<FormSettings>.<SearchDirectory>.Value = Nothing Then
            'Form setting not saved.
        Else
            SearchDirectory = Settings.<FormSettings>.<SearchDirectory>.Value
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

    Private Sub frmAddProject_Load(sender As Object, e As EventArgs) Handles Me.Load

        Label12.Visible = False 'Hide "Searching ..." label
        chkSearchSubDirectories.Checked = True
        chkShowThisAppProjects.Checked = True

        If SearchDirectory <> "" Then
            txtSearchDirectory.Text = SearchDirectory
        End If

        'For Each drv As System.IO.DriveInfo In My.Computer.FileSystem.Drives
        '    TreeView1.Nodes.Add(drv.Name)
        '    TreeView1.Nodes(TreeView1.Nodes.Count - 1).Tag = drv.Name
        '    TreeView1.Nodes(TreeView1.Nodes.Count - 1).Nodes.Add("*DUMMY*")
        'Next

    End Sub

    'Private Sub TreeView1_BeforeCollapse(sender As Object, e As Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeCollapse
    '    ' clear the node that is being collapsed
    '    e.Node.Nodes.Clear()
    '    ' add a dummy TreeNode to the node being collapsed so it is expandable
    '    e.Node.Nodes.Add("*DUMMY*")
    'End Sub

    'Private Sub TreeView1_BeforeExpand(sender As Object, e As Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand
    '    ' clear the expanding node so we can re-populate it, or else we end up with duplicate nodes
    '    e.Node.Nodes.Clear()
    '    ' get the directory representing this node
    '    Dim mNodeDirectory As IO.DirectoryInfo
    '    mNodeDirectory = New IO.DirectoryInfo(e.Node.Tag.ToString)
    '    ' add each subdirectory from the file system to the expanding node as a child node
    '    For Each mDirectory As IO.DirectoryInfo In mNodeDirectory.GetDirectories
    '        ' declare a child TreeNode for the next subdirectory
    '        Dim mDirectoryNode As New System.Windows.Forms.TreeNode
    '        ' store the full path to this directory in the child TreeNode's Tag property
    '        mDirectoryNode.Tag = mDirectory.FullName
    '        ' set the child TreeNodes's display text
    '        mDirectoryNode.Text = mDirectory.Name
    '        ' add a dummy TreeNode to this child TreeNode to make it expandable
    '        mDirectoryNode.Nodes.Add("*DUMMY*")
    '        ' add this child TreeNode to the expanding TreeNode
    '        e.Node.Nodes.Add(mDirectoryNode)
    '    Next
    'End Sub

   

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        SaveFormSettings() 'Save the form settings.
        Me.Close()
    End Sub

#End Region 'Form Display Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Form Methods - The main actions performed by this form" ' ---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click

        FolderBrowserDialog2.SelectedPath = SearchDirectory
        'FolderBrowserDialog2.ShowDialog()
        If FolderBrowserDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then
            SearchDirectory = FolderBrowserDialog2.SelectedPath
        End If

    End Sub


    Private Sub btnFindFiles_Click(sender As Object, e As EventArgs) Handles btnFindFiles.Click
        ListBox1.Items.Clear()
        'For Each foundFile As String In My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "Project_Info.xml")

        Label12.Visible = True 'Show "Searching ..." label
        Me.Refresh()

        'Timer1.Interval = 500 'ms
        'Timer1.Enabled = True
        'Timer1.Start()



        With ProgressBar1
            'Try
            If chkSearchSubDirectories.Checked = True Then
                '.Maximum = My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "ADVL_Project_Info.xml").Count
                .Maximum = My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "Project_Info_ADVL_2.xml").Count 'Search for the latest file version.
                '.Maximum = My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "ADVL_Project_Info.xml").Where(Function(file) (file.Contai
            Else
                '.Maximum = My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchTopLevelOnly, "ADVL_Project_Info.xml").Count
                .Maximum = My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchTopLevelOnly, "Project_Info_ADVL_2.xml").Count
            End If

                'Catch ex As Exception
                '    RaiseEvent ErrorMessage(ex.Message)
                'End Try

                .Minimum = 0
            .Step = 1
            '.Value = 1
            If .Maximum = 0 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With

        Dim ProjectInfoXDoc As System.Xml.Linq.XDocument
        Dim ProjectAppName As String

        Dim Zip As New ZipComp

        If chkSearchSubDirectories.Checked = True Then
            'For Each foundFile As String In My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "ADVL_Project_Info.xml")
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "Project_Info_ADVL_2.xml")
                Try
                    ProgressBar1.PerformStep()
                    If chkShowThisAppProjects.Checked Then
                        'Only show projects belonging the this application.
                        ProjectInfoXDoc = XDocument.Load(foundFile)
                        ProjectAppName = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                        'If ProjectAppName = ApplicationSummary.Name Then ListBox1.Items.Add(foundFile)
                        If ProjectAppName = ApplicationName Then ListBox1.Items.Add(foundFile)
                    Else
                        ListBox1.Items.Add(foundFile)
                    End If
                Catch ex As Exception
                    RaiseEvent ErrorMessage(ex.Message)
                End Try
            Next
        Else
            'For Each foundFile As String In My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchTopLevelOnly, "ADVL_Project_Info.xml")
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchTopLevelOnly, "Project_Info_ADVL_2.xml")
                Try
                    ProgressBar1.PerformStep()
                    If chkShowThisAppProjects.Checked Then
                        'Only show projects belonging the this application.
                        ProjectInfoXDoc = XDocument.Load(foundFile)
                        ProjectAppName = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                        'If ProjectAppName = ApplicationSummary.Name Then ListBox1.Items.Add(foundFile)
                        If ProjectAppName = ApplicationName Then ListBox1.Items.Add(foundFile)
                    Else
                        ListBox1.Items.Add(foundFile)
                    End If
                Catch ex As Exception
                    RaiseEvent ErrorMessage(ex.Message)
                End Try
            Next
        End If

        With ProgressBar2
            'Try
            If chkSearchSubDirectories.Checked = True Then
                    .Maximum = My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "*.AdvlProject").Count
                Else
                    .Maximum = My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchTopLevelOnly, "*.AdvlProject").Count
                End If
                'Catch ex As Exception
                '    RaiseEvent ErrorMessage(ex.Message)
                'End Try


                .Minimum = 0
            .Step = 1
            If .Maximum = 0 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With

        'Try
        If chkSearchSubDirectories.Checked = True Then
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchAllSubDirectories, "*.AdvlProject")
                Try
                    ProgressBar2.PerformStep()
                    If chkShowThisAppProjects.Checked Then
                        'Only show projects belonging the this application.
                        Zip.ArchivePath = foundFile
                        If Zip.ArchiveExists Then
                            If Zip.EntryExists("ADVL_Project_Info.xml") Then
                                ProjectInfoXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("ADVL_Project_Info.xml"))
                                ProjectAppName = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                                'If ProjectAppName = ApplicationSummary.Name Then ListBox1.Items.Add(foundFile)
                                If ProjectAppName = ApplicationName Then ListBox1.Items.Add(foundFile)

                            ElseIf Zip.EntryExists("Project_Info_ADVL_2.xml") Then
                                ProjectInfoXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
                                ProjectAppName = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                                'If ProjectAppName = ApplicationSummary.Name Then ListBox1.Items.Add(foundFile)
                                If ProjectAppName = ApplicationName Then ListBox1.Items.Add(foundFile)

                            Else
                                RaiseEvent ErrorMessage("Project information file not found in archive: " & ListBox1.Text & vbCrLf)
                            End If
                        Else
                            RaiseEvent ErrorMessage("Archive file does not exist: " & foundFile & vbCrLf)
                        End If
                    Else
                        ListBox1.Items.Add(foundFile)
                    End If
                Catch ex As Exception
                    RaiseEvent ErrorMessage(ex.Message)
                End Try
            Next
        Else
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(SearchDirectory, FileIO.SearchOption.SearchTopLevelOnly, "*.AdvlProject")
                Try
                    ProgressBar2.PerformStep()
                    If chkShowThisAppProjects.Checked Then
                        'Only show projects belonging the this application.
                        Zip.ArchivePath = foundFile
                        If Zip.ArchiveExists Then
                            If Zip.EntryExists("ADVL_Project_Info.xml") Then
                                ProjectInfoXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("ADVL_Project_Info.xml"))
                                ProjectAppName = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                                'If ProjectAppName = ApplicationSummary.Name Then ListBox1.Items.Add(foundFile)
                                If ProjectAppName = ApplicationName Then ListBox1.Items.Add(foundFile)

                            ElseIf Zip.EntryExists("Project_Info_ADVL_2.xml") Then
                                ProjectInfoXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
                                ProjectAppName = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                                'If ProjectAppName = ApplicationSummary.Name Then ListBox1.Items.Add(foundFile)
                                If ProjectAppName = ApplicationName Then ListBox1.Items.Add(foundFile)

                            Else
                                RaiseEvent ErrorMessage("Project information file not found in archive: " & ListBox1.Text & vbCrLf)
                            End If
                        Else
                            RaiseEvent ErrorMessage("Archive file does not exist: " & foundFile & vbCrLf)
                        End If
                    Else
                        ListBox1.Items.Add(foundFile)
                    End If
                Catch ex As Exception
                    RaiseEvent ErrorMessage(ex.Message)
                End Try
            Next
        End If
            'Catch ex As Exception
            '    RaiseEvent ErrorMessage(ex.Message)
            'End Try
            Label12.Visible = False
        'Timer1.Stop()

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        'Display the project properties in the ADVL_Project_Info.xml file:

        'If ListBox1.Text.EndsWith("ADVL_Project_Info.xml") Then
        If ListBox1.Text.EndsWith("Project_Info_ADVL_2.xml") Then
            'Directory project.
            Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ListBox1.Text)
            txtProjectName.Text = ProjectInfoXDoc.<Project>.<Name>.Value
            txtDescription.Text = ProjectInfoXDoc.<Project>.<Description>.Value
            txtType.Text = ProjectInfoXDoc.<Project>.<Type>.Value
            txtCreationDate.Text = ProjectInfoXDoc.<Project>.<CreationDate>.Value
            txtAuthor.Text = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
            txtSettingsRelLocnType.Text = ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Type>.Value
            txtSettingsRelLocnPath.Text = ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value
            txtDataRelLocnType.Text = ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Type>.Value
            txtDataRelLocnPath.Text = ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value
            txtApplicationName.Text = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
            Dim Directory As String
            Directory = System.IO.Path.GetDirectoryName(ListBox1.Text)
            'If Trim(txtSettingsRelLocnPath.Text) = Trim(Directory) Then
            '    'Project is in the same location as the saved Settings Path
            '    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
            '    txtComments.ForeColor = Drawing.Color.Black
            '    txtComments.Text = "Project location OK."
            'Else
            '    'Project has been moved from the saved Settings Path
            '    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
            '    txtComments.ForeColor = Drawing.Color.Red
            '    txtComments.Text = "Project has moved from the saved location!"
            'End If
        ElseIf ListBox1.Text.EndsWith(".AdvlProject") Then
            'Archive project.
            'RaiseEvent Message(".AdvlProject type selected." & vbCrLf)
            Dim Zip As New ZipComp
            Zip.ArchivePath = ListBox1.Text
            If Zip.ArchiveExists Then
                If Zip.EntryExists("Project_Info_ADVL_2.xml") Then
                    'Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(Zip.GetText("ADVL_Project_Info.xml"))
                    Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
                    txtProjectName.Text = ProjectInfoXDoc.<Project>.<Name>.Value
                    txtDescription.Text = ProjectInfoXDoc.<Project>.<Description>.Value
                    txtType.Text = ProjectInfoXDoc.<Project>.<Type>.Value
                    txtCreationDate.Text = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                    txtAuthor.Text = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                    txtSettingsRelLocnType.Text = ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Type>.Value
                    txtSettingsRelLocnPath.Text = ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value
                    txtDataRelLocnType.Text = ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Type>.Value
                    txtDataRelLocnPath.Text = ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value
                    txtApplicationName.Text = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                    'If Trim(txtSettingsRelLocnPath.Text) = Trim(ListBox1.Text) Then
                    '    'Project is in the same location as the saved Settings Path
                    '    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
                    '    txtComments.ForeColor = Drawing.Color.Black
                    '    txtComments.Text = "Project location OK."
                    'Else
                    '    'Project has been moved from the saved Settings Path
                    '    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
                    '    txtComments.ForeColor = Drawing.Color.Red
                    '    txtComments.Text = "Project has moved from the saved location!"
                    'End If
                ElseIf Zip.EntryExists("ADVL_Project_Info.xml") Then
                    Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("ADVL_Project_Info.xml"))
                    txtProjectName.Text = ProjectInfoXDoc.<Project>.<Name>.Value
                    txtDescription.Text = ProjectInfoXDoc.<Project>.<Description>.Value
                    txtType.Text = ProjectInfoXDoc.<Project>.<Type>.Value
                    txtCreationDate.Text = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                    txtAuthor.Text = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                    txtSettingsRelLocnType.Text = ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Type>.Value
                    txtSettingsRelLocnPath.Text = ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value
                    txtDataRelLocnType.Text = ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Type>.Value
                    txtDataRelLocnPath.Text = ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value
                    txtApplicationName.Text = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                Else
                    RaiseEvent ErrorMessage("Project information file not found in archive: " & ListBox1.Text & vbCrLf)
                End If
            Else
                RaiseEvent ErrorMessage("Selected archive does not exist: " & ListBox1.Text & vbCrLf)
            End If

        Else
            RaiseEvent ErrorMessage("Unknown project type. Project selected: " & ListBox1.Text & vbCrLf)
        End If

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Add the selected project to the project list.

        'If ListBox1.Text.EndsWith("ADVL_Project_Info.xml") Then
        If ListBox1.Text.EndsWith("Project_Info_ADVL_2.xml") Then
            'Directory project or Hybrid project.
            Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ListBox1.Text)
            'Dim PathChanged As Boolean = False
            Dim ProjectSummary As New ADVL_Utilities_Library_1.ProjectSummary

            If ProjectInfoXDoc Is Nothing Then
                RaiseEvent ErrorMessage("No Project Information was found." & vbCrLf & vbCrLf)
                Exit Sub
            End If

            If ProjectInfoXDoc.<Project>.<Application>.<Name>.Value <> ApplicationName Then
                RaiseEvent ErrorMessage("The Project Application Name is: " & ProjectInfoXDoc.<Project>.<Application>.<Name>.Value & vbCrLf)
                RaiseEvent ErrorMessage("This does not match the current Application Name: " & ApplicationName & vbCrLf & vbCrLf)
                Exit Sub
            End If

            Select Case ProjectInfoXDoc.<Project>.<Type>.Value
                Case "Directory"
                    'Check if the SettingsPath is the same as the current Project Directory path:
                    Dim ProjectDirPath As String = System.IO.Path.GetDirectoryName(ListBox1.Text)

                    'NOTE: Only the relative location of the SettingsLocn and DataLocn are now stored! - No need for the following checks:
                    'If ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value = ProjectDirPath Then
                    '    'Saved Settings Path is correct.
                    'Else
                    '    'Saved Settings Path must be updated:
                    '    'ProjectInfoXDoc.<Project>.<SettingsLocation.<Path>.Value = ProjectDirPath
                    '    'ProjectInfoXDoc.Document.Element("Project/SettingsLocation").SetElementValue("Path", ProjectDirPath)
                    '    ProjectInfoXDoc.Element("Project").Element("SettingsLocation").SetElementValue("Path", ProjectDirPath)
                    '    PathChanged = True
                    'End If
                    'If ProjectInfoXDoc.<Project>.<DataLocation>.<Path>.Value = ProjectDirPath Then
                    '    'Saved Data Path is correct.
                    'Else
                    '    'Saved Data Path must be updated:
                    '    'ProjectInfoXDoc.Document.Element("Project/DataLocation").SetElementValue("Path", ProjectDirPath)
                    '    ProjectInfoXDoc.Element("Project").Element("DataLocation").SetElementValue("Path", ProjectDirPath)
                    '    PathChanged = True
                    'End If
                    'If PathChanged = True Then
                    '    ProjectInfoXDoc.Save(ProjectDirPath & "\" & "ADVL_Project_Info.xml")
                    'End If

                    'ProjectSummary.ApplicationName = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Name>.Value
                    ProjectSummary.AuthorName = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                    ProjectSummary.CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                    ProjectSummary.Description = ProjectInfoXDoc.<Project>.<Description>.Value
                    ProjectSummary.Name = ProjectInfoXDoc.<Project>.<Name>.Value
                    'ProjectSummary.SettingsLocnPath = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value
                    'ProjectSummary.Path = ProjectInfoXDoc.<Project>.<Path>.Value
                    ProjectSummary.Path = ProjectDirPath

                    'ProjectSummary.SettingsLocnType = FileLocation.Types.Directory
                    'ProjectSummary.Type = FileLocation.Types.Directory

                    ProjectSummary.Type = Project.Types.Directory

                    RaiseEvent ProjectToAdd(ProjectSummary)
                Case "Hybrid"

                    ''FOR TESTING:
                    'RaiseEvent Message("Hybrid Project:" & vbCrLf)
                    'Dim ProjectDirPath As String = System.IO.Path.GetDirectoryName(ListBox1.Text)
                    'RaiseEvent Message("Project directory path: " & ProjectDirPath & vbCrLf)
                    'RaiseEvent Message("ProjectSummary.AuthorName: " & ProjectInfoXDoc.<Project>.<Author>.<Name>.Value & vbCrLf)
                    'RaiseEvent Message("ProjectSummary.CreationDate: " & ProjectInfoXDoc.<Project>.<CreationDate>.Value & vbCrLf)
                    'RaiseEvent Message("ProjectSummary.Description: " & ProjectInfoXDoc.<Project>.<Description>.Value & vbCrLf)
                    'RaiseEvent Message("ProjectSummary.Name: " & ProjectInfoXDoc.<Project>.<Name>.Value & vbCrLf)
                    'RaiseEvent Message("ProjectSummary.Path: " & ProjectDirPath & vbCrLf)
                    'RaiseEvent Message("ProjectSummary.Type: " & "Hybrid" & vbCrLf)

                    Dim ProjectDirPath As String = System.IO.Path.GetDirectoryName(ListBox1.Text)
                    ProjectSummary.AuthorName = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                    ProjectSummary.CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                    ProjectSummary.Description = ProjectInfoXDoc.<Project>.<Description>.Value
                    ProjectSummary.Name = ProjectInfoXDoc.<Project>.<Name>.Value
                    ProjectSummary.Path = ProjectDirPath
                    ProjectSummary.Type = Project.Types.Hybrid
                    RaiseEvent ProjectToAdd(ProjectSummary)
                Case Else

            End Select

        ElseIf ListBox1.Text.EndsWith(".AdvlProject") Then
            'Archive project.
            Dim Zip As New ZipComp
            Zip.ArchivePath = ListBox1.Text
            If Zip.ArchiveExists Then
                'If Zip.EntryExists("ADVL_Project_Info.xml") Then
                If Zip.EntryExists("Project_Info_ADVL_2.xml") Then
                    'Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("ADVL_Project_Info.xml"))
                    Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
                    'Dim PathChanged As Boolean = False
                    Dim ProjectSummary As New ADVL_Utilities_Library_1.ProjectSummary

                    'NOTE: Only the relative location of the SettingsLocn and DataLocn are now stored! - No need for the following checks:
                    ''Check if the SettingsPath is the same as the current Project Directory path:
                    'If ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value = ListBox1.Text Then
                    '    'Saved Settings Path is correct.
                    'Else
                    '    'Saved Settings Path must be updated:
                    '    ProjectInfoXDoc.Element("Project").Element("SettingsLocation").SetElementValue("Path", ListBox1.Text)
                    '    PathChanged = True
                    'End If
                    ''Check if the DataPath is the same as the current Project Directory path:
                    'If ProjectInfoXDoc.<Project>.<DataLocation>.<Path>.Value = ListBox1.Text Then
                    '    'Saved Data Path is correct.
                    'Else
                    '    'Saved Data Path must be updated:
                    '    ProjectInfoXDoc.Element("Project").Element("DataLocation").SetElementValue("Path", ListBox1.Text)
                    '    PathChanged = True
                    'End If

                    'If PathChanged = True Then
                    '    Zip.AddText("ADVL_Project_Info.xml", ProjectInfoXDoc.ToString)
                    'End If

                    'ProjectSummary.ApplicationName = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Name>.Value
                    ProjectSummary.AuthorName = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                    ProjectSummary.CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                    ProjectSummary.Description = ProjectInfoXDoc.<Project>.<Description>.Value
                    ProjectSummary.Name = ProjectInfoXDoc.<Project>.<Name>.Value
                    'ProjectSummary.SettingsLocnPath = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value
                    'ProjectSummary.Path = ProjectInfoXDoc.<Project>.<Path>.Value
                    ProjectSummary.Path = Zip.ArchivePath

                    'ProjectSummary.SettingsLocnType = FileLocation.Types.Archive

                    ProjectSummary.Type = Project.Types.Archive

                    RaiseEvent ProjectToAdd(ProjectSummary)
                End If

            End If
        End If








        'Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ListBox1.Text)

        'Dim ProjectSummary As New ADVL_System_Utilities.ProjectSummary
        'ProjectSummary.ApplicationName = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Name>.Value
        'ProjectSummary.AuthorName = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
        'ProjectSummary.CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
        'ProjectSummary.Description = ProjectInfoXDoc.<Project>.<Description>.Value
        'ProjectSummary.Name = ProjectInfoXDoc.<Project>.<Name>.Value
        'ProjectSummary.SettingsLocnPath = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value

        'Select Case ProjectInfoXDoc.<Project>.<SettingsLocation>.<Type>.Value
        '    Case "Directory"
        '        ProjectSummary.SettingsLocnType = FileLocation.Types.Directory
        '    Case "Archive"
        '        ProjectSummary.SettingsLocnType = FileLocation.Types.Archive
        '    Case Else
        '        RaiseEvent ErrorMessage("Settings location type not recognized: " & ProjectInfoXDoc.<Project>.<SettingsLocation>.<Type>.Value & vbCrLf)
        'End Select
        ''ProjectSummary.SettingsLocnType = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Type>.Value
        'Select Case ProjectInfoXDoc.<Project>.<Type>.Value
        '    Case "Directory"
        '        ProjectSummary.Type = Project.Types.Directory
        '    Case "Archive"
        '        ProjectSummary.Type = Project.Types.Archive
        '    Case "Hybrid"
        '        ProjectSummary.Type = Project.Types.Hybrid
        '    Case "None"
        '        ProjectSummary.Type = Project.Types.None
        '    Case Else
        '        RaiseEvent ErrorMessage("Project type not recognized: " & ProjectInfoXDoc.<Project>.<Type>.Value & vbCrLf)
        'End Select
        ''ProjectSummary.Type = ProjectInfoXDoc.<Project>.<Type>.Value

        'RaiseEvent ProjectToAdd(ProjectSummary)

    End Sub

    Private Sub frmAddProject_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        'Form is being resized:
        'Dim FormWidth As Integer = Me.Width

        'ProgressBar1.Right = Me.Width / 2
        ProgressBar1.Width = Me.Width / 2 - 25
        ProgressBar2.Left = ProgressBar1.Width + 20
        ProgressBar2.Width = Me.Width / 2 - 25

    End Sub

    'Private Sub Timer1_Tick(sender As Object, e As System.EventArgs) Handles Timer1.Tick

    '    If Label12.Visible = True Then
    '        Label12.Visible = False
    '        'Me.Refresh()
    '    Else
    '        Label12.Visible = True
    '        'Me.Refresh()
    '    End If
    'End Sub


#End Region 'Form Methods --------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Event ProjectToAdd(ByVal ProjectSummary As ADVL_Utilities_Library_1.ProjectSummary) 'Raise an event to indicate that a project has been selected to add to the Project List.
    Public Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Public Event Message(ByVal Message As String) 'Send a message.

#End Region


    'btnTestXmlUpdate removed 30Jul16. This was used for testing code.

    'Private Sub btnTestXmlUpdate_Click(sender As Object, e As EventArgs) Handles btnTestXmlUpdate.Click
    '    'Test XML update code:

    '    'Read the ADVL_Project_Info.xml file

    '    If ListBox1.Text.EndsWith("ADVL_Project_Info.xml") Then
    '        'Directory project.
    '        Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ListBox1.Text)
    '        Dim OldSettingsPath As String = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value
    '        Dim NewSettingsPath As String = System.IO.Path.GetDirectoryName(ListBox1.Text)

    '        RaiseEvent Message("Project type: " & ProjectInfoXDoc.<Project>.<Type>.Value & vbCrLf)
    '        RaiseEvent Message("Old settings path: " & OldSettingsPath & vbCrLf)
    '        RaiseEvent Message("New settings path: " & NewSettingsPath & vbCrLf & vbCrLf)

    '    ElseIf ListBox1.Text.EndsWith(".AdvlProject") Then
    '        'Archive project.
    '        Dim Zip As New ZipComp
    '        Zip.ArchivePath = ListBox1.Text
    '        If Zip.ArchiveExists Then
    '            If Zip.EntryExists("ADVL_Project_Info.xml") Then
    '                Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("ADVL_Project_Info.xml"))
    '                Dim OldSettingsPath As String = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value
    '                Dim OldSettingsDir As String = System.IO.Path.GetDirectoryName(OldSettingsPath)
    '                Dim OldSettingsArchiveName As String = System.IO.Path.GetFileName(OldSettingsPath)

    '                RaiseEvent Message("Project type: " & ProjectInfoXDoc.<Project>.<Type>.Value & vbCrLf)
    '                RaiseEvent Message("Old settings path: " & OldSettingsPath & vbCrLf)
    '                RaiseEvent Message("New settings path: " & ListBox1.Text & vbCrLf & vbCrLf)

    '            End If
    '        End If
    '    End If
    'End Sub

End Class