Public Class frmProject
    'The Project form is used to select a project for an Andorville(TM) application or add or remove a project from the list of available projects.
    'This form is opened from the Project class using the SelectProject() method.

#Region " Variable Declarations - All the variables used in this form." '--------------------------------------------------------------------------------------------------------------------

    'Dim ProjectInfo As New ADVL_Utilities.ProjectInfo 'This object stores information about the selected project.
    'Dim Project As New ADVL_Utilities.Project 'This object stores information about the selected project.
    Dim ProjectSummary As New ProjectSummary 'A project summary is provided for each project in the Project List.
    'Dim NewProjectSummary As New ProjectSummary

    Public ApplicationSummary As New ApplicationSummary 'The Application Summary is updated by the Project class that calls the Project Form. This summary is needed when a new project is created.

    'Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings. 'UPDATE 29Jul18 - Replaced by ProjectLocn
    'Public ProjectLocn As ADVL_Utilities_Library_1.FileLocation 'The project location.
    Public ProjectLocn As New ADVL_Utilities_Library_1.FileLocation 'The project location.

    'NOTE: Application Name now extracted from ApplicationSummary (23Apr19)
    'Public ApplicationName As String 'The name of the application associated with the projects.

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
        'Dim RowNo As Integer
        'Dim SelProjNo As Integer
        Dim SelRowNo As Integer
        If DataGridView1.SelectedRows.Count > 0 Then
            'RowNo = DataGridView1.SelectedRows(0).Index
            'SelProjNo = Val(DataGridView1.SelectedRows(0).Cells(0).Value)
            SelRowNo = DataGridView1.SelectedRows(0).Index
        Else
            'RowNo = -1
            'SelProjNo = -1
            SelRowNo = -1
        End If

        Dim Settings = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Form settings for Project form.-->
                       <FormSettings>
                           <Left><%= Me.Left %></Left>
                           <Top><%= Me.Top %></Top>
                           <Width><%= Me.Width %></Width>
                           <Height><%= Me.Height %></Height>
                           <SelectedProjectNo><%= SelRowNo %></SelectedProjectNo>
                           <!---->
                       </FormSettings>
        ' <SelectedProjectNo><%= SelProjNo %></SelectedProjectNo>
        '    <SelectedProjectNo><%= RowNo %></SelectedProjectNo>
        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationSummary.Name & "_" & Me.Text & ".xml"

        'SettingsLocn.SaveXmlData(SettingsFileName, Settings)
        ProjectLocn.SaveXmlData(SettingsFileName, Settings)

    End Sub

    Public Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationSummary.Name & "_" & Me.Text & ".xml"

        Dim Settings As System.Xml.Linq.XDocument

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

#End Region 'Process XML Files --------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '--------------------------------------------------------------------------------------------------------------------------

    Private Sub frmProject_Load(sender As Object, e As EventArgs) Handles Me.Load
        'RaiseEvent FrmLoading(Me.Text)
        'This form is opened by the Project.SelectProject method, which also initialises the form.

        'Set up DataGridView1 to display a summary of each project:
        DataGridView1.Rows.Clear()
        'DataGridView1.ColumnCount = 3
        DataGridView1.ColumnCount = 4
        'DataGridView1.Columns(0).HeaderText = "Name"
        'DataGridView1.Columns(1).HeaderText = "Type"
        'DataGridView1.Columns(2).HeaderText = "Description"
        'DataGridView1.Columns(3).HeaderText = "No"
        DataGridView1.Columns(0).HeaderText = "No"
        DataGridView1.Columns(1).HeaderText = "Name"
        DataGridView1.Columns(2).HeaderText = "Type"
        DataGridView1.Columns(3).HeaderText = "Description"
        DataGridView1.AutoResizeColumns()
        DataGridView1.AllowUserToAddRows = False



        'Set up DataGridView2 to display a summary of each recycled project:
        DataGridView2.Rows.Clear()
        'DataGridView2.ColumnCount = 3
        DataGridView2.ColumnCount = 4
        'DataGridView2.Columns(0).HeaderText = "Name"
        'DataGridView2.Columns(1).HeaderText = "Type"
        'DataGridView2.Columns(2).HeaderText = "Description"
        'DataGridView2.Columns(3).HeaderText = "No"
        DataGridView2.Columns(0).HeaderText = "No"
        DataGridView2.Columns(1).HeaderText = "Name"
        DataGridView2.Columns(2).HeaderText = "Type"
        DataGridView2.Columns(3).HeaderText = "Description"
        DataGridView2.AutoResizeColumns()
        DataGridView2.AllowUserToAddRows = False

        TabControl1.SelectTab("TabPage1")
        'Project List Tab selected.
        btnDelete.Text = "Recycle"
        btnRestore.Enabled = False
        chkDeleteData.Enabled = False

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
            'NewProjectForm.ApplicationName = ApplicationName
            NewProjectForm.ApplicationName = ApplicationSummary.Name
            'NewProjectForm.SettingsLocn = SettingsLocn
            NewProjectForm.ProjectLocn = ProjectLocn
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
            'AddProjectForm.ApplicationName = ApplicationName
            AddProjectForm.ApplicationName = ApplicationSummary.Name
            'AddProjectForm.SettingsLocn = SettingsLocn
            AddProjectForm.ProjectLocn = ProjectLocn
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

        If DataGridView1.SelectedRows.Count > 0 Then
            Dim SelProjNo As Integer = Val(DataGridView1.SelectedRows(0).Cells(0).Value)

            'Check if the project file or directory is missing:
            Select Case ProjectList(SelProjNo).Type
                Case Project.Types.Archive
                    txtType.Text = "Archive"
                    If System.IO.File.Exists(ProjectList(SelProjNo).Path) Then
                        'Archive found.
                    Else
                        RaiseEvent ErrorMessage("Project Archive not found!" & vbCrLf)
                        Exit Sub
                    End If
                Case Project.Types.Directory
                    txtType.Text = "Directory"
                    If System.IO.Directory.Exists(ProjectList(SelProjNo).Path) Then
                        'Directory found.
                    Else
                        RaiseEvent ErrorMessage("Project Directory not found!" & vbCrLf)
                        Exit Sub
                    End If
                Case Project.Types.Hybrid
                    txtType.Text = "Hybrid"
                    If System.IO.Directory.Exists(ProjectList(SelProjNo).Path) Then
                        'Directory found.
                    Else
                        RaiseEvent ErrorMessage("Project Directory not found!" & vbCrLf)
                        Exit Sub
                    End If
                Case Project.Types.None
                    txtType.Text = "None"
            End Select

            RaiseEvent Message("Project number: " & SelProjNo & " selected." & vbCrLf)
            RaiseEvent Message("Project name: " & ProjectList(SelProjNo).Name & vbCrLf)
            RaiseEvent Message("Project type: " & ProjectList(SelProjNo).Type.ToString & vbCrLf)
            RaiseEvent Message("Project path: " & ProjectList(SelProjNo).Path & vbCrLf & vbCrLf)
            RaiseEvent ProjectSelected(ProjectList(SelProjNo))
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

        RaiseEvent NewProjectCreated(ProjectSummary.Path)
    End Sub

    Private Sub WriteProjectList()
        'Write the ProjectList data to the ProjectList.xml file in the Application Directory.

        Dim ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Project List File-->
                              <ProjectList>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <ApplicationName><%= ApplicationSummary.Name %></ApplicationName>
                                  <%= From item In ProjectList
                                      Select
                                      <Project>
                                          <Name><%= item.Name %></Name>
                                          <Description><%= item.Description %></Description>
                                          <Type><%= item.Type %></Type>
                                          <Path><%= item.Path %></Path>
                                          <CreationDate><%= Format(item.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                          <AuthorName><%= item.AuthorName %></AuthorName>
                                          <Status><%= item.Status %></Status>
                                      </Project>
                                  %>
                              </ProjectList>
        '                                          <Status><%= item.Status %></Status> - ADDED 30Mar20 - Status is OK or Recycled - OK: the project is on the Project List, Recycled: on the Recycled List.
        '     <ApplicationName><%= ApplicationName %></ApplicationName>

        'ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")
        'ProjectListXDoc.Save(ApplicationDir & "\Project_List_ADVL_2.xml")
        'NOTE: Leave Project List file name as Project_List.xml. <FormatCode>ADVL_2</FormatCode> has been added but no need to change the file name.
        'ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")
        ProjectListXDoc.Save(ApplicationDir & "\Project_List_ADVL_2.xml") 'ÚPDATE THE FILE NAME!

    End Sub

    Private Sub ReadProjectList()
        'Read the Project_List.xml file in the Application Directory.

        'If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'The latest ADVL_2 format version of the Project List file exists.
        'NOTE: Leave Project List file name as Project_List.xml. <FormatCode>ADVL_2</FormatCode> has been added but no need to change the file name.
        'NOTE: IGNOTE THE COMMENT ABOVE: Project_List.xml files will be converted to Project_List_ADVL_2.xml files.
        'If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'The latest ADVL_2 format version of the Project List file exists.
        If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'The latest ADVL_2 format version of the Project List file exists.
            Dim ProjectListXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")
            ReadProjectListAdvl_2(ProjectListXDoc)
        Else
            If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'The original ADVL_1 format version of the Project List file exists.
                RaiseEvent Message("Converting Project_List.xml to Project_List_ADVL_2.xml." & vbCrLf)
                'Convert the file to the latest ADVL_2 format:
                Dim Conversion As New ADVL_Utilities_Library_1.FormatConvert.ProjectListFileConversion
                Conversion.DirectoryPath = ApplicationDir
                Conversion.InputFileName = "Project_List.xml"
                Conversion.InputFormatCode = FormatConvert.ProjectListFileConversion.FormatCodes.ADVL_1
                Conversion.OutputFormatCode = FormatConvert.ProjectListFileConversion.FormatCodes.ADVL_2
                Conversion.Convert()
                If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then
                    ReadProjectList() 'Try ReadProjectList again. This time Project_List_ADVL_2.xml should be found
                Else
                    RaiseEvent ErrorMessage("Error converting Project_List.xml to Project_List_ADVL_2.xml." & vbCrLf)
                End If
            Else
                RaiseEvent ErrorMessage("No versions of the Project List were found." & vbCrLf)
            End If
        End If
    End Sub

    Private Sub ReadProjectList_Old()
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

                NewProject.Path = item.<Path>.Value

                NewProject.CreationDate = item.<CreationDate>.Value

                'Select Case item.<SettingsLocationType>.Value
                '    Case "Directory"
                '        NewProject.SettingsLocnType = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                '    Case "Archive"
                '        NewProject.SettingsLocnType = ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'End Select
                'NewProject.SettingsLocnPath = item.<SettingsLocationPath>.Value

                'Select Case item.<DataLocationType>.Value
                '    Case "Directory"
                '        NewProject.DataLocnPath = ADVL_Utilities.Location.Types.Directory
                '    Case "Archive"
                '        NewProject.DataLocnPath = ADVL_Utilities.Location.Types.Archive
                'End Select
                'NewProject.DataLocnPath = item.<SettingsLocationPath>.Value
                NewProject.AuthorName = item.<AuthorName>.Value
                'NewProject.ApplicationName = item.<ApplicationName>.Value

                ProjectList.Add(NewProject)
            Next

            'NProjects = ProjectList.Count
            UpdateProjectGrid()
        Else 'Create a Project List.

        End If
    End Sub

    Private Sub ReadProjectListAdvl_2(ByRef XDoc As System.Xml.Linq.XDocument)
        'Readt the Project List XDocument (ADVL_2 format version).

        Dim Projects = From item In XDoc.<ProjectList>.<Project>

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
            NewProject.Path = item.<Path>.Value
            NewProject.CreationDate = item.<CreationDate>.Value
            NewProject.AuthorName = item.<AuthorName>.Value
            If item.<Status>.Value = Nothing Then
                'The Project list file records do not contain the Status field.
            Else
                NewProject.Status = item.<Status>.Value
            End If
            ProjectList.Add(NewProject)
        Next
        UpdateProjectGrid()

    End Sub

    Private Sub UpdateProjectGrid()
        'Update DataGridView1 - Show the Name, Type and Description of each project.

        DataGridView1.Rows.Clear() 'Clear the list of Projects
        DataGridView2.Rows.Clear() 'Clear the list of Recycled Projects

        Dim NProjects As Integer = ProjectList.Count

        If NProjects = 0 Then
            Exit Sub
        End If

        Dim Index As Integer

        For Index = 0 To NProjects - 1
            If ProjectList(Index).Status = "Recycled" Then
                'Add the Project to the Recycled Project List:
                'DataGridView2.Rows.Add()
                'DataGridView2.Rows(Index).Cells(0).Value = ProjectList(Index).Name
                'DataGridView2.Rows(Index).Cells(1).Value = ProjectList(Index).Type
                'DataGridView2.Rows(Index).Cells(2).Value = ProjectList(Index).Description

                'DataGridView2.Rows.Add(ProjectList(Index).Name, ProjectList(Index).Type, ProjectList(Index).Description, Index)

                DataGridView2.Rows.Add(Index, ProjectList(Index).Name, ProjectList(Index).Type, ProjectList(Index).Description)
            Else
                'Add the Project to the Project List:
                'DataGridView1.Rows.Add()
                'DataGridView1.Rows(Index).Cells(0).Value = ProjectList(Index).Name
                'DataGridView1.Rows(Index).Cells(1).Value = ProjectList(Index).Type
                'DataGridView1.Rows(Index).Cells(2).Value = ProjectList(Index).Description

                'DataGridView1.Rows.Add(ProjectList(Index).Name, ProjectList(Index).Type, ProjectList(Index).Description, Index)

                DataGridView1.Rows.Add(Index, ProjectList(Index).Name, ProjectList(Index).Type, ProjectList(Index).Description)
            End If
        Next

        DataGridView1.AutoResizeColumns()
        DataGridView2.AutoResizeColumns()

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
        'txtType.Text = ProjectList(RecordNo - 1).SettingsLocnType
        txtType.Text = ProjectList(RecordNo - 1).Type
        'txtPath.Text = ProjectList(RecordNo - 1).SettingsLocnPath
        txtPath.Text = ProjectList(RecordNo - 1).Path
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
        If RowNo = -1 Then Exit Sub
        Dim SelProjNo = Val(DataGridView1.SelectedRows(0).Cells(0).Value)

        DataGridView1.Rows(RowNo).Selected = True
        'ShowProjectInfo(RowNo)
        ShowProjectInfo(SelProjNo)

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
                'If System.IO.File.Exists(ProjectList(RecordNo).SettingsLocnPath) Then
                If System.IO.File.Exists(ProjectList(RecordNo).Path) Then
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
                'If System.IO.Directory.Exists(ProjectList(RecordNo).SettingsLocnPath) Then
                If System.IO.Directory.Exists(ProjectList(RecordNo).Path) Then
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
                'If System.IO.Directory.Exists(ProjectList(RecordNo).SettingsLocnPath) Then
                If System.IO.Directory.Exists(ProjectList(RecordNo).Path) Then
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
        'Select Case ProjectList(RecordNo).SettingsLocnType
        'Select Case ProjectList(RecordNo).Type
        '    Case FileLocation.Types.Archive
        '        txtType.Text = "Archive"
        '    Case FileLocation.Types.Directory
        '        txtType.Text = "Directory"
        'End Select
        'txtPath.Text = ProjectList(RecordNo).SettingsLocnPath
        txtPath.Text = ProjectList(RecordNo).Path
        'txtApplicationName.Text = ProjectList(RecordNo).ApplicationName

    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        'Remove the selected project from the list.

        'Dim RowNo As Integer

        If TabControl1.SelectedTab.Name = "TabPage1" Then
            'Project List Tab selected.
            If DataGridView1.SelectedRows.Count > 0 Then
                'Recycle the selected project.
                Dim SelProjNo As Integer = Val(DataGridView1.SelectedRows(0).Cells(0).Value)
                ProjectList(SelProjNo).Status = "Recycled"
                UpdateProjectGrid()
            Else
                RaiseEvent ErrorMessage("No project has been selected to Delete!" & vbCrLf)
            End If
        Else
            'Recycled Tab selected.
            Dim SelProjNo As Integer = Val(DataGridView2.SelectedRows(0).Cells(0).Value)
            Dim SelProjName As String = ProjectList(SelProjNo).Name
            Dim SelProjPath As String = ProjectList(SelProjNo).Path
            If chkDeleteData.Checked Then DeleteProject(SelProjName, SelProjPath)
            ProjectList.RemoveAt(SelProjNo)
            UpdateProjectGrid()

        End If


        'If DataGridView1.SelectedRows.Count > 0 Then
        '    RowNo = DataGridView1.SelectedRows(0).Index
        '    ProjectList.RemoveAt(RowNo)
        '    UpdateProjectGrid()
        'Else
        '    RaiseEvent ErrorMessage("No project has been selected for removal!" & vbCrLf)
        'End If
    End Sub

    Private Sub DeleteProject(ByVal Name As String, ByVal Path As String)
        'Delete the project with the specified Name at the specified Path.
        RaiseEvent Message("Project name: " & Name & vbCrLf & "Path: " & Path & vbCrLf)
        'RaiseEvent ErrorMessage("Delete Project Data code not yet complete!" & vbCrLf)

        If MsgBox("Are you sure you want to delete the project data?", MsgBoxStyle.YesNoCancel, "Warning") = MsgBoxResult.Yes Then
            RaiseEvent ErrorMessage("Delete Project Data code not yet complete!" & vbCrLf)
        End If

    End Sub

    Private Sub btnRestore_Click(sender As Object, e As EventArgs) Handles btnRestore.Click
        'Restore the selected project to the Project List.

        If TabControl1.SelectedTab.Name = "TabPage1" Then
            'Wrong tab selected. Only projects shown in the Recycled Project List can be recycled
        Else
            'Recycled tab selected
            If DataGridView2.SelectedRows.Count > 0 Then
                'Restore the selected project.
                Dim SelProjNo As Integer = Val(DataGridView2.SelectedRows(0).Cells(0).Value)
                ProjectList(SelProjNo).Status = "OK"
                UpdateProjectGrid()
            End If
        End If

    End Sub


    'Private Sub btnSelectDefault_Click(sender As Object, e As EventArgs)

    '    'Check if ProjectList contains a project named Default
    '    Dim SelectedProject As ProjectSummary
    '    SelectedProject = ProjectList.Find(Function(p) p.Name = "Default")

    '    If SelectedProject.Name = "Default" Then
    '        'The Default project is already on the list.
    '    Else
    '        RaiseEvent CreateDefaultProject() 'Create the Default project. This also adds the project to the list.
    '        ReadProjectList() 'Read the updated project list.
    '        UpdateProjectGrid() 'Úpdate the project grid.
    '    End If

    'End Sub

    Private Sub btnSelDefault_Click(sender As Object, e As EventArgs) Handles btnSelDefault.Click

        'Check if ProjectList contains a project named Default
        Dim SelectedProject As ProjectSummary
        SelectedProject = ProjectList.Find(Function(p) p.Name = "Default")

        'If SelectedProject.Name = "Default" Then
        If SelectedProject Is Nothing Then
            'The Default project does not exist.
            RaiseEvent CreateDefaultProject() 'Create the Default project. This also adds the project to the list.
            ReadProjectList() 'Read the updated project list.
            UpdateProjectGrid() 'Úpdate the project grid.
        Else
            'The Default project is already on the list.

        End If
    End Sub

    'Private Sub TabPage1_GotFocus(sender As Object, e As EventArgs) Handles TabPage1.GotFocus
    '    'Project List Tab selected.
    '    btnDelete.Text = "Recycle"
    '    btnRestore.Enabled = False
    '    chkDeleteData.Enabled = False

    'End Sub

    'Private Sub TabPage2_GotFocus(sender As Object, e As EventArgs) Handles TabPage2.GotFocus
    '    'Recycled Tab selected.
    '    btnDelete.Text = "Delete"
    '    btnRestore.Enabled = True
    '    chkDeleteData.Enabled = True

    'End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        'A tab has been selected:

        If TabControl1.SelectedTab.Name = "TabPage1" Then
            'Project List Tab selected.
            btnDelete.Text = "Recycle"
            btnRestore.Enabled = False
            chkDeleteData.Enabled = False

            btnSelect.Enabled = True
            btnNew.Enabled = True
            btnAdd.Enabled = True
            btnSaveAs.Enabled = True
            btnSelDefault.Enabled = True
        Else
            'Recycled Tab selected.
            btnDelete.Text = "Delete"
            btnRestore.Enabled = True
            chkDeleteData.Enabled = True

            btnSelect.Enabled = False
            btnNew.Enabled = False
            btnAdd.Enabled = False
            btnSaveAs.Enabled = False
            btnSelDefault.Enabled = False
        End If
    End Sub

    'Private Sub btnSelectDefault_Click(sender As Object, e As EventArgs) Handles btnSelectDefault.Click
    '    'Select the Default project.

    '    'The Default project is a directory project named 'Default_Project' in the Application Directory.

    '    'Check if the Default_Project exists:
    '    'CreateDefaultProject

    '    'If System.IO.Directory.Exists(ApplicationDir & "\" & "Default_Project") Then
    '    '    'The Default_Project exists
    '    '    'Check if it is on the list:

    '    'Else
    '    '    'Default_Project does not exist.
    '    '    RaiseEvent CreateDefaultProject() 'Create the Default_Project.
    '    '    ReadProjectList() 'Re-read the project list. The Default_Project should now be listed.
    '    'End If

    '    'RaiseEvent OpenDefaultProject()

    '    'Check if ProjectList contains a project named Default
    '    Dim SelectedProject As ProjectSummary
    '    SelectedProject = ProjectList.Find(Function(p) p.Name = "Default")

    '    If SelectedProject.Name = "Default" Then
    '        'The Default project is already on the list.
    '    Else
    '        RaiseEvent CreateDefaultProject() 'Create the Default project. This also adds the project to the list.
    '        ReadProjectList() 'Read the updated project list.
    '        UpdateProjectGrid() 'Úpdate the project grid.
    '    End If


    '    'RaiseEvent CreateDefaultProject()

    '    'Search DataGridView1 for a project named Default
    '    'Search ProjectList for a project named Default



    'End Sub



#End Region 'Form Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Event ProjectSelected(ByRef ProjectSummary As ProjectSummary) 'Raise an event passing the selected project information.
    Public Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Public Event Message(ByVal Message As String) 'Send a message.
    'Public Event CreateDefaultProject() '
    Public Event OpenDefaultProject() 'Raise an event to open the Default_Project.
    Public Event CreateDefaultProject() 'Raise an event to create the Default_Project (if it does not already exist).

    Public Event NewProjectCreated(ByVal ProjectPath As String) 'Raise an event to indicate that a new project has been created at the specified path.






#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class