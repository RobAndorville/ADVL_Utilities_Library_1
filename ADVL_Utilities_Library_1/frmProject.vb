Public Class frmProject
    'The Project form is used to select a project for an Andorville Labs application or add or remove a project from the list of available projects.

#Region " Variable Declarations - All the variables used in this form." '--------------------------------------------------------------------------------------------------------------------

    'Dim ProjectInfo As New ADVL_Utilities.ProjectInfo 'This object stores information about the selected project.
    'Dim Project As New ADVL_Utilities.Project 'This object stores information about the selected project.
    Dim ProjectSummary As New ProjectSummary 'A project summary is provided for each project inthe Project List.
    'Dim NewProjectSummary As New ProjectSummary

    Public ApplicationSummary As New ApplicationSummary 'The Application Summary is updated by the Project class that calls the Project Form. This summary is needed when a new project is created.

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public ApplicationName As String 'The name of the application using the message form.

    'Public SettingsLocn As New Location 'This is a directory or archive where settings are stored.

    Public WithEvents NewProjectForm As frmNewProject
    Public WithEvents AddProjectForm As frmAddProject

    'Dim ProjectList As New List(Of ProjectInfo) 'List of projects
    Dim ProjectList As New List(Of ProjectSummary) 'List of projects
    'Dim NProjects As Integer = 0 'The number of projects in the list
    Dim CurrentRecordNo As Integer 'The number of the current record displayed

    Dim SelectedProjectNo As Integer 'Stores the selected row number in DataGridView1

#End Region 'Variable Declarations -----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '-----------------------------------------------------------------------------------------------------------

    Private _applicationDir As String = "" 'The path to the directory used to store application data.
    Public Property ApplicationDir As String
        Get
            Return _applicationDir
        End Get
        Set(ByVal value As String)
            _applicationDir = value
        End Set
    End Property

    'THIS INFORMATION IS NOW INCLUDED IN THE APPLICATIONSUMMARY OBJECT.
    'Private _applicationName As String = "" 'The Application Name. This is used as part of the name of a settings file.
    'Property ApplicationName As String
    '    Get
    '        Return _applicationName
    '    End Get
    '    Set(value As String)
    '        _applicationName = value
    '    End Set
    'End Property

    'THIS INFORMATION IS NOW INCLUDED IN THE APPLICATIONSUMMARY OBJECT.
    'Private _projectFileExt = "" 'The file extension for a new Project File.
    'Property ProjectFileExt As String
    '    Get
    '        Return _projectFileExt
    '    End Get
    '    Set(value As String)
    '        _projectFileExt = value
    '    End Set
    'End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Process XML files - Read and write XML files." '-----------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveFormSettings()
        'Save the form settings as an XML document.
        '
        Dim RowNo As Integer
        If DataGridView1.SelectedRows.Count > 0 Then
            RowNo = DataGridView1.SelectedRows(0).Index
        Else
            RowNo = -1
        End If

        Dim Settings = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Form settings for Project form.-->
                       <FormSettings>
                           <Left><%= Me.Left %></Left>
                           <Top><%= Me.Top %></Top>
                           <Width><%= Me.Width %></Width>
                           <Height><%= Me.Height %></Height>
                           <SelectedProjectNo><%= RowNo %></SelectedProjectNo>
                           <!---->
                       </FormSettings>

        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        SettingsLocn.SaveXmlData(SettingsFileName, Settings)

    End Sub

    Public Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"

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

        If Settings.<FormSettings>.<SelectedProjectNo>.Value = Nothing Then
            'Form setting not saved.
            SelectedProjectNo = 0
            DataGridView1.ClearSelection()
            DataGridView1.Rows(SelectedProjectNo).Selected = True
            ShowProjectInfo(SelectedProjectNo)
        Else
            SelectedProjectNo = Settings.<FormSettings>.<SelectedProjectNo>.Value
            DataGridView1.ClearSelection()
            If SelectedProjectNo > -1 Then
                If SelectedProjectNo + 1 > DataGridView1.Rows.Count Then
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Selected = True
                Else
                    DataGridView1.Rows(SelectedProjectNo).Selected = True
                End If

            End If

            ShowProjectInfo(SelectedProjectNo)
        End If

        'Else
        ''Settings file not found.
        'End If

    End Sub

#End Region 'Process XML Files --------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '--------------------------------------------------------------------------------------------------------------------------

    Private Sub frmProject_Load(sender As Object, e As EventArgs) Handles Me.Load
        'RaiseEvent FrmLoading(Me.Text)
        'This form is opened by the Project.SelectProject method, which also initialises the form.

        'Set up DataGridView1 to display a summary of each project:
        DataGridView1.Rows.Clear()
        DataGridView1.ColumnCount = 3
        DataGridView1.Columns(0).HeaderText = "Name"
        DataGridView1.Columns(1).HeaderText = "Type"
        DataGridView1.Columns(2).HeaderText = "Description"
        DataGridView1.AutoResizeColumns()

        ReadProjectList()
        'NOTE: ReadProjectDir must be called after the project has loaded. ApplicationDir is not set until then.

        'DataGridView1.Rows(SelectedProjectNo).Selected = True


    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        SaveFormSettings() 'Save the form settings.
        WriteProjectList()
        Me.Close()
    End Sub

#End Region 'Form Display Subroutines -------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and close forms - Code used to open and close other forms." '-----------------------------------------------------------------------------------------------------------------

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        'Show the NewProject form:
        If IsNothing(NewProjectForm) Then
            NewProjectForm = New frmNewProject
            NewProjectForm.ApplicationName = ApplicationName
            NewProjectForm.SettingsLocn = SettingsLocn
            NewProjectForm.Show()
            NewProjectForm.ApplicationDir = ApplicationDir     'Pass the name of the Application Directory.
            'NewProjectForm.ApplicationName = ApplicationName   'Pass the name of the Application Name.
            'NewProjectForm.ProjectFileExt = ProjectFileExt     'Pass the project file extension.
            'NewProjectForm.SettingsLocn = SettingsLocn
            NewProjectForm.RestoreFormSettings()
        Else
            NewProjectForm.Show()
            NewProjectForm.BringToFront()
        End If
    End Sub


    Private Sub NewProjectForm_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles NewProjectForm.FormClosed
        NewProjectForm = Nothing
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Show the Add Project form:
        If IsNothing(AddProjectForm) Then
            AddProjectForm = New frmAddProject
            AddProjectForm.ApplicationName = ApplicationName
            AddProjectForm.SettingsLocn = SettingsLocn
            AddProjectForm.Show()
            AddProjectForm.ApplicationDir = ApplicationDir 'Pass the name of the Application Directory.
            AddProjectForm.RestoreFormSettings()
        Else
            AddProjectForm.Show()
            AddProjectForm.BringToFront()
        End If
    End Sub

    Private Sub AddProjectForm_ErrorMessage(Message As String) Handles AddProjectForm.ErrorMessage
        RaiseEvent ErrorMessage(Message)
    End Sub

    Private Sub AddProjectForm_Message(Message As String) Handles AddProjectForm.Message
        RaiseEvent Message(Message)
    End Sub

    Private Sub AddProjectForm_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles AddProjectForm.FormClosed
        AddProjectForm = Nothing
    End Sub

#End Region


#Region "Form methods - The main actions performed by this form" ' --------------------------------------------------------------------------------------------------------------------------

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        'Select the project.

        'If DataGridView1.SelectedRows(0).Index > -1 Then
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim RowNo As Integer = DataGridView1.SelectedRows(0).Index

            'Check if the project file or directory is missing:
            Select Case ProjectList(RowNo).Type
                Case Project.Types.Archive
                    txtType.Text = "Archive"
                    If System.IO.File.Exists(ProjectList(RowNo).SettingsLocnPath) Then
                        'Archive found.
                    Else
                        RaiseEvent ErrorMessage("Project Archive not found!" & vbCrLf)
                        Exit Sub
                    End If
                Case Project.Types.Directory
                    txtType.Text = "Directory"
                    If System.IO.Directory.Exists(ProjectList(RowNo).SettingsLocnPath) Then
                        'Directory found.
                    Else
                        RaiseEvent ErrorMessage("Project Directory not found!" & vbCrLf)
                        Exit Sub
                    End If
                Case Project.Types.Hybrid
                    txtType.Text = "Hybrid"
                    If System.IO.Directory.Exists(ProjectList(RowNo).SettingsLocnPath) Then
                        'Directory found.
                    Else
                         RaiseEvent ErrorMessage("Project Directory not found!" & vbCrLf)
                        Exit Sub
                    End If
                Case Project.Types.None
                    txtType.Text = "None"
            End Select

            RaiseEvent Message("Project number: " & RowNo & " selected." & vbCrLf)
            'RaiseEvent Message("Project name: " & RowNo & ProjectList(RowNo).Name & vbCrLf)
            RaiseEvent Message("Project name: " & ProjectList(RowNo).Name & vbCrLf)
            'RaiseEvent Message("Settings location type: " & RowNo & ProjectList(RowNo).SettingsLocnType & vbCrLf)
            RaiseEvent Message("Settings location type: " & ProjectList(RowNo).SettingsLocnType.ToString & vbCrLf)
            'RaiseEvent Message("Settings location path: " & RowNo & ProjectList(RowNo).SettingsLocnPath & vbCrLf & vbCrLf)
            RaiseEvent Message("Settings location path: " & ProjectList(RowNo).SettingsLocnPath & vbCrLf & vbCrLf)
            RaiseEvent ProjectSelected(ProjectList(RowNo))
        Else
            RaiseEvent ErrorMessage("No project selected." & vbCrLf)
        End If


    End Sub


    Private Sub NewProjectForm_CreateProjectError(ErrorMessage As String) Handles NewProjectForm.CreateProjectError
        'There was an error while attempting to create a new project.
        RaiseEvent ErrorMessage(ErrorMessage) 'Pass the error message to the Main form.
    End Sub

    Private Sub NewProjectForm_NewProjectCreated(ProjectSummary As ProjectSummary) Handles NewProjectForm.NewProjectCreated
        'A new project has been created.
        ProjectList.Add(ProjectSummary) 'Add the project summary to the list.
        'NProjects += 1                  'Increment the number of projects
        UpdateProjectGrid()             'Update the list of projects displayed in DataGridView1
    End Sub

    Private Sub WriteProjectList()
        'Write the ProjectList data to the ProjectList.xml file in the Application Directory.

        Dim ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Project List File-->
                              <ProjectList>
                                  <%= From item In ProjectList _
                                      Select _
                                      <Project>
                                          <Name><%= item.Name %></Name>
                                          <Description><%= item.Description %></Description>
                                          <Type><%= item.Type %></Type>
                                          <CreationDate><%= Format(item.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                          <SettingsLocationType><%= item.SettingsLocnType %></SettingsLocationType>
                                          <SettingsLocationPath><%= item.SettingsLocnPath %></SettingsLocationPath>
                                          <AuthorName><%= item.AuthorName %></AuthorName>
                                          <ApplicationName><%= item.ApplicationName %></ApplicationName>
                                      </Project>
                                  %>
                              </ProjectList>

        ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")

    End Sub

    Private Sub ReadProjectList()
        'Read the Project_List.xml file in the Application Directory.

        If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'Read the Project List.
            Dim ProjectListXDoc As System.Xml.Linq.XDocument
            ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List.xml")

            Dim Projects = From item In ProjectListXDoc.<ProjectList>.<Project>

            ProjectList.Clear()
            For Each item In Projects
                Dim NewProject As New ProjectSummary
                NewProject.Name = item.<Name>.Value
                NewProject.Description = item.<Description>.Value
                Select Case item.<Type>.Value
                    Case "None"
                        NewProject.Type = Project.Types.None
                    Case "Directory"
                        NewProject.Type = Project.Types.Directory
                    Case "Archive"
                        NewProject.Type = Project.Types.Archive
                    Case "Hybrid"
                        NewProject.Type = Project.Types.Hybrid
                End Select
                NewProject.CreationDate = item.<CreationDate>.Value
                Select Case item.<SettingsLocationType>.Value
                    Case "Directory"
                        NewProject.SettingsLocnType = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                    Case "Archive"
                        NewProject.SettingsLocnType = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                End Select
                NewProject.SettingsLocnPath = item.<SettingsLocationPath>.Value
                'Select Case item.<DataLocationType>.Value
                '    Case "Directory"
                '        NewProject.DataLocnPath = ADVL_Utilities.Location.Types.Directory
                '    Case "Archive"
                '        NewProject.DataLocnPath = ADVL_Utilities.Location.Types.Archive
                'End Select
                'NewProject.DataLocnPath = item.<SettingsLocationPath>.Value
                NewProject.AuthorName = item.<AuthorName>.Value
                NewProject.ApplicationName = item.<ApplicationName>.Value

                ProjectList.Add(NewProject)
            Next

            'NProjects = ProjectList.Count
            UpdateProjectGrid()
        Else 'Create a Project List.


        End If
    End Sub

    Private Sub UpdateProjectGrid()
        'Update DataGridView1 - Show the Name, Type and Description of each project.

        DataGridView1.Rows.Clear()

        Dim NProjects As Integer = ProjectList.Count

        If NProjects = 0 Then
            Exit Sub
        End If

        Dim Index As Integer

        For Index = 0 To NProjects - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(Index).Cells(0).Value = ProjectList(Index).Name
            DataGridView1.Rows(Index).Cells(1).Value = ProjectList(Index).Type
            DataGridView1.Rows(Index).Cells(2).Value = ProjectList(Index).Description
        Next

        DataGridView1.AutoResizeColumns()

    End Sub

    Private Sub DisplayRecordNo(ByVal RecordNo As Integer)
        'Displays a record from ProjectList.

        If RecordNo < 1 Then
            'Main.MessageStyleWarningSet()
            'Main.MessageAdd("Cannot display Unit Of Measure data. Selected record number is too small." & vbCrLf)
            RaiseEvent ErrorMessage("Cannot display Unit Of Measure data. Selected record number is too small." & vbCrLf)
            Exit Sub
        End If

        'If RecordNo > NProjects Then
        If RecordNo > ProjectList.Count Then
            'Main.MessageStyleWarningSet()
            'Main.MessageAdd("Cannot display Unit Of Measure data. Selected record number is too large." & vbCrLf)
            RaiseEvent ErrorMessage("Cannot display Unit Of Measure data. Selected record number is too large." & vbCrLf)
            Exit Sub
        End If

        txtProjectName.Text = ProjectList(RecordNo - 1).Name
        txtDescription.Text = ProjectList(RecordNo - 1).Description
        txtAuthor.Text = ProjectList(RecordNo - 1).AuthorName
        txtType.Text = ProjectList(RecordNo - 1).Type
        txtCreationDate.Text = ProjectList(RecordNo - 1).CreationDate
        txtSettingsLocnType.Text = ProjectList(RecordNo - 1).SettingsLocnType
        txtSettingsLocnPath.Text = ProjectList(RecordNo - 1).SettingsLocnPath
        'txtDataLocnType.Text = ProjectList(RecordNo - 1).DataLocnType
        'txtDataLocnPath.Text = ProjectList(RecordNo - 1).DataLocnPath


    End Sub

  

    Private Sub AddProjectForm_ProjectToAdd(ProjectSummary As ProjectSummary) Handles AddProjectForm.ProjectToAdd
        ProjectList.Add(ProjectSummary) 'Add the project summary to the list.
        'NProjects += 1                  'Increment the number of projects
        UpdateProjectGrid()             'Update the list of projects displayed in DataGridView1
    End Sub



    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        Dim RowNo As Integer = e.RowIndex

        DataGridView1.Rows(RowNo).Selected = True
        ShowProjectInfo(RowNo)

    End Sub

    Private Sub ShowProjectInfo(ByVal RecordNo As Integer)

        If RecordNo >= ProjectList.Count Then
            RaiseEvent ErrorMessage("Error displaying project information. Record number too large." & vbCrLf)
            Exit Sub
        End If

        If RecordNo = -1 Then
            'No project selected.
            Exit Sub
        End If

        txtProjectName.Text = ProjectList(RecordNo).Name
        txtDescription.Text = ProjectList(RecordNo).Description
        Select Case ProjectList(RecordNo).Type
            Case Project.Types.Archive
                txtType.Text = "Archive"
                If System.IO.File.Exists(ProjectList(RecordNo).SettingsLocnPath) Then
                    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
                    txtComments.ForeColor = Drawing.Color.Black
                    txtComments.Text = "Archive found."
                Else
                    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
                    txtComments.ForeColor = Drawing.Color.Red
                    txtComments.Text = "Archive not found!"
                End If
            Case Project.Types.Directory
                txtType.Text = "Directory"
                If System.IO.Directory.Exists(ProjectList(RecordNo).SettingsLocnPath) Then
                    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
                    txtComments.ForeColor = Drawing.Color.Black
                    txtComments.Text = "Directory found."
                Else
                    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
                    txtComments.ForeColor = Drawing.Color.Red
                    txtComments.Text = "Directory not found!"
                End If
            Case Project.Types.Hybrid
                txtType.Text = "Hybrid"
                If System.IO.Directory.Exists(ProjectList(RecordNo).SettingsLocnPath) Then
                    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
                    txtComments.ForeColor = Drawing.Color.Black
                    txtComments.Text = "Directory found."
                Else
                    txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
                    txtComments.ForeColor = Drawing.Color.Red
                    txtComments.Text = "Directory not found!"
                End If
            Case Project.Types.None
                txtType.Text = "None"
        End Select
        txtCreationDate.Text = ProjectList(RecordNo).CreationDate
        txtAuthor.Text = ProjectList(RecordNo).AuthorName
        Select Case ProjectList(RecordNo).SettingsLocnType
            Case FileLocation.Types.Archive
                txtSettingsLocnType.Text = "Archive"
            Case FileLocation.Types.Directory
                txtSettingsLocnType.Text = "Directory"
        End Select
        txtSettingsLocnPath.Text = ProjectList(RecordNo).SettingsLocnPath
        txtApplicationName.Text = ProjectList(RecordNo).ApplicationName

    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        'Remove the selected project from the list.

        Dim RowNo As Integer

        If DataGridView1.SelectedRows.Count > 0 Then
            RowNo = DataGridView1.SelectedRows(0).Index
            ProjectList.RemoveAt(RowNo)
            UpdateProjectGrid()
        Else
            RaiseEvent ErrorMessage("No project has been selected for removal!" & vbCrLf)
        End If
    End Sub



#End Region 'Form Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Event ProjectSelected(ByRef ProjectSummary As ProjectSummary) 'Raise an event passing the selected project information.
    Public Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Public Event Message(ByVal Message As String) 'Send a message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------















End Class