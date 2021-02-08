Public Class frmProjectParams
    'The Project Parameters form is used to display the project parameters for an Andorville(TM) application.
    'This form is opened from the Project class using the ProjectParams() method.

#Region " Variable Declarations - All the variables used in this form." '--------------------------------------------------------------------------------------------------------------------


    'Dim ProjectSummary As New ProjectSummary 'A project summary is provided for each project in the Project List.


    'Public ApplicationSummary As New ApplicationSummary 'The Application Summary is updated by the Project class that calls the Project Form. This summary is needed when a new project is created.


    Public ProjectLocn As New ADVL_Utilities_Library_1.FileLocation 'The project location.
    'Public ProjectPath As String

    Public ApplicationName As String 'The name of the application associated with the project.

    Public ParentProjectLocn As New ADVL_Utilities_Library_1.FileLocation 'The parent project location.
    'Public ParentProjectPath As String


    'ADDED 3Feb19
    Public Parameter As New Dictionary(Of String, ParamInfo) 'Dictionary of Project Parameters.
    Public ParentParameter As New Dictionary(Of String, ParamInfo) 'Dictionary of Parent Project Parameters


    'Public WithEvents NewProjectForm As frmNewProject
    'Public WithEvents AddProjectForm As frmAddProject


    'Dim ProjectList As New List(Of ProjectSummary) 'List of projects

    'Dim CurrentRecordNo As Integer 'The number of the current record displayed

    'Dim SelectedProjectNo As Integer 'Stores the selected row number in DataGridView1

#End Region 'Variable Declarations -----------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '-----------------------------------------------------------------------------------------------------------

    'Private _applicationDir As String = "" 'The path to the directory used to store application data.
    'Public Property ApplicationDir As String
    '    Get
    '        Return _applicationDir
    '    End Get
    '    Set(ByVal value As String)
    '        _applicationDir = value
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
                           <!---->
                       </FormSettings>

        '<SelectedProjectNo><%= RowNo %></SelectedProjectNo>

        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        ProjectLocn.SaveXmlData(SettingsFileName, Settings)

    End Sub

    Public Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"

        Dim Settings As System.Xml.Linq.XDocument

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
        CheckFormPos()
        'If Settings.<FormSettings>.<SelectedProjectNo>.Value = Nothing Then
        '    'Form setting not saved.
        '    SelectedProjectNo = 0
        '    DataGridView1.ClearSelection()
        '    DataGridView1.Rows(SelectedProjectNo).Selected = True
        '    ShowProjectInfo(SelectedProjectNo)
        'Else
        '    SelectedProjectNo = Settings.<FormSettings>.<SelectedProjectNo>.Value
        '    DataGridView1.ClearSelection()
        '    If SelectedProjectNo > -1 Then
        '        If SelectedProjectNo + 1 > DataGridView1.Rows.Count Then
        '            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Selected = True
        '        Else
        '            DataGridView1.Rows(SelectedProjectNo).Selected = True
        '        End If

        '    End If

        '    ShowProjectInfo(SelectedProjectNo)
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

        'Set up DataGridView1 to display each project parameter:
        DataGridView1.Rows.Clear()
        DataGridView1.ColumnCount = 3
        DataGridView1.Columns(0).HeaderText = "Name"
        'DataGridView1.Columns(1).HeaderText = "Type"
        DataGridView1.Columns(1).HeaderText = "Value"
        DataGridView1.Columns(2).HeaderText = "Description"

        'DataGridView1.AutoResizeColumns(Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells)

        DataGridView1.AutoResizeColumns()

        rbParameters.Checked = True
        'ReadProjectParameters()
        ShowProjectParameters() 'ADDED 3Feb19
        'DataGridView1.AutoResizeColumns()

        'ReadProjectList()
        'NOTE: ReadProjectDir must be called after the project has loaded. ApplicationDir is not set until then.

        'DataGridView1.Rows(SelectedProjectNo).Selected = True



    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        SaveFormSettings() 'Save the form settings.
        'WriteProjectList()
        Me.Close()
    End Sub

#End Region 'Form Display Subroutines -------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and close forms - Code used to open and close other forms." '-----------------------------------------------------------------------------------------------------------------


#End Region 'Open and close forms -----------------------------------------------------------------------------------------------------------------------------------------------------------




#Region "Form methods - The main actions performed by this form" ' --------------------------------------------------------------------------------------------------------------------------

    'Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
    '    'Select the project.

    '    If DataGridView1.SelectedRows.Count > 0 Then
    '        Dim RowNo As Integer = DataGridView1.SelectedRows(0).Index

    '        'Check if the project file or directory is missing:
    '        Select Case ProjectList(RowNo).Type
    '            Case Project.Types.Archive
    '                txtType.Text = "Archive"
    '                If System.IO.File.Exists(ProjectList(RowNo).Path) Then
    '                    'Archive found.
    '                Else
    '                    RaiseEvent ErrorMessage("Project Archive not found!" & vbCrLf)
    '                    Exit Sub
    '                End If
    '            Case Project.Types.Directory
    '                txtType.Text = "Directory"
    '                If System.IO.Directory.Exists(ProjectList(RowNo).Path) Then
    '                    'Directory found.
    '                Else
    '                    RaiseEvent ErrorMessage("Project Directory not found!" & vbCrLf)
    '                    Exit Sub
    '                End If
    '            Case Project.Types.Hybrid
    '                txtType.Text = "Hybrid"
    '                If System.IO.Directory.Exists(ProjectList(RowNo).Path) Then
    '                    'Directory found.
    '                Else
    '                    RaiseEvent ErrorMessage("Project Directory not found!" & vbCrLf)
    '                    Exit Sub
    '                End If
    '            Case Project.Types.None
    '                txtType.Text = "None"
    '        End Select

    '        RaiseEvent Message("Project number: " & RowNo & " selected." & vbCrLf)
    '        RaiseEvent Message("Project name: " & ProjectList(RowNo).Name & vbCrLf)
    '        RaiseEvent Message("Project type: " & ProjectList(RowNo).Type.ToString & vbCrLf)
    '        RaiseEvent Message("Project path: " & ProjectList(RowNo).Path & vbCrLf & vbCrLf)
    '        RaiseEvent ProjectSelected(ProjectList(RowNo))
    '    Else
    '        RaiseEvent ErrorMessage("No project selected." & vbCrLf)
    '    End If

    'End Sub


    'Private Sub NewProjectForm_CreateProjectError(ErrorMessage As String) Handles NewProjectForm.CreateProjectError
    '    'There was an error while attempting to create a new project.
    '    RaiseEvent ErrorMessage(ErrorMessage) 'Pass the error message to the Main form.
    'End Sub

    'Private Sub NewProjectForm_NewProjectCreated(ProjectSummary As ProjectSummary) Handles NewProjectForm.NewProjectCreated
    '    'A new project has been created.
    '    ProjectList.Add(ProjectSummary) 'Add the project summary to the list.
    '    UpdateProjectGrid()             'Update the list of projects displayed in DataGridView1
    'End Sub

    'Private Sub WriteProjectList()
    '    'Write the ProjectList data to the ProjectList.xml file in the Application Directory.

    '    Dim ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
    '                          <!---->
    '                          <!--Project List File-->
    '                          <ProjectList>
    '                              <FormatCode>ADVL_2</FormatCode>
    '                              <ApplicationName><%= ApplicationName %></ApplicationName>
    '                              <%= From item In ProjectList
    '                                  Select
    '                                  <Project>
    '                                      <Name><%= item.Name %></Name>
    '                                      <Description><%= item.Description %></Description>
    '                                      <Type><%= item.Type %></Type>
    '                                      <Path><%= item.Path %></Path>
    '                                      <CreationDate><%= Format(item.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
    '                                      <AuthorName><%= item.AuthorName %></AuthorName>
    '                                  </Project>
    '                              %>
    '                          </ProjectList>

    '    'NOTE: Leave Project List file name as Project_List.xml. <FormatCode>ADVL_2</FormatCode> has been added but no need to change the file name.
    '    ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")

    'End Sub

    'UPDATE: ShowProjectParameters is used now.
    Public Sub ReadProjectParameters()
        'Read the Project_Params_ADVL_2.xml file in the Project directory.

        Dim XmlDoc As System.Xml.Linq.XDocument

        DataGridView1.Rows.Clear()

        Select Case ProjectLocn.Type
            Case FileLocation.Types.Directory
                'Read the Xml data document in the directory at ProjectLocn.Path
                If System.IO.File.Exists(ProjectLocn.Path & "\" & "Project_Params_ADVL_2.xml") Then
                    Try
                        XmlDoc = XDocument.Load(ProjectLocn.Path & "\" & "Project_Params_ADVL_2.xml")
                    Catch ex As Exception
                        'RaiseEvent ErrorMessage("Error reading XML file. " & ex.Message & vbCrLf)
                    End Try

                Else
                    XmlDoc = Nothing
                End If

            Case FileLocation.Types.Archive
                'Read the Xml data document in the archive at ProjectLocn.Path

                Dim Zip As New ZipComp
                Zip.ArchivePath = ProjectLocn.Path
                If Zip.EntryExists("Project_Params_ADVL_2.xml") Then
                    XmlDoc = XDocument.Parse(Zip.GetText("Project_Params_ADVL_2.xml"))
                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing

        End Select


        If XmlDoc Is Nothing Then
            'The Project Parameters could not be read from the Project_Params_ADVL_2.xml file.
        Else
            'Display the list of Project Parameters in DataGridView1:

            Dim Params = From item In XmlDoc.<ProjectParameterList>.<Parameter>

            For Each item In Params
                DataGridView1.Rows.Add(item.<Name>.Value, item.<Value>.Value, item.<Description>.Value)
            Next
            DataGridView1.AutoResizeColumns()
        End If

    End Sub

    'ADDED 3Feb19
    Public Sub ShowProjectParameters()
        'Show the contents of the Parameter list in DataGridView1:

        DataGridView1.Rows.Clear()

        For Each item In Parameter
            DataGridView1.Rows.Add(item.Key, item.Value.Value, item.Value.Description)
        Next
    End Sub

    'UPDATE: ShowParentProjectParameters is used now.
    Public Sub ReadParentProjectParameters()
        'Read the Parent Project Project_Params_ADVL_2.xml file in the Project directory.

        Dim XmlDoc As System.Xml.Linq.XDocument

        DataGridView1.Rows.Clear()

        Select Case ParentProjectLocn.Type
            Case FileLocation.Types.Directory
                'Read the Xml data document in the directory at ProjectLocn.Path
                If System.IO.File.Exists(ParentProjectLocn.Path & "\" & "Project_Params_ADVL_2.xml") Then
                    Try
                        XmlDoc = XDocument.Load(ParentProjectLocn.Path & "\" & "Project_Params_ADVL_2.xml")
                    Catch ex As Exception
                        'RaiseEvent ErrorMessage("Error reading XML file. " & ex.Message & vbCrLf)
                    End Try

                Else
                    XmlDoc = Nothing
                End If

            Case FileLocation.Types.Archive
                'Read the Xml data document in the archive at ProjectLocn.Path

                Dim Zip As New ZipComp
                Zip.ArchivePath = ParentProjectLocn.Path
                If Zip.EntryExists("Project_Params_ADVL_2.xml") Then
                    XmlDoc = XDocument.Parse(Zip.GetText("Project_Params_ADVL_2.xml"))
                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing

        End Select


        If XmlDoc Is Nothing Then
            'The Project Parameters could not be read from the Project_Params_ADVL_2.xml file.
        Else
            'Display the list of Project Parameters in DataGridView1:

            Dim Params = From item In XmlDoc.<ProjectParameterList>.<Parameter>

            For Each item In Params
                DataGridView1.Rows.Add(item.<Name>.Value, item.<Value>.Value, item.<Description>.Value)
            Next
            DataGridView1.AutoResizeColumns()
        End If
    End Sub

    'ADDED 3Feb19
    Public Sub ShowParentProjectParameters()
        'Show the contents of the ParentParameter list in DataGridView1:

        DataGridView1.Rows.Clear()

        For Each item In ParentParameter
            DataGridView1.Rows.Add(item.Key, item.Value.Value, item.Value.Description)
        Next
        DataGridView1.AutoResizeColumns()
    End Sub

    Private Sub rbParameters_CheckedChanged(sender As Object, e As EventArgs) Handles rbParameters.CheckedChanged
        If rbParameters.Checked Then
            'ReadProjectParameters()
            ShowProjectParameters()
        Else
            'ReadParentProjectParameters()
            ShowParentProjectParameters()
        End If
    End Sub

    Private Sub rbParentParameters_CheckedChanged(sender As Object, e As EventArgs) Handles rbParentParameters.CheckedChanged
        If rbParameters.Checked Then
            'ReadProjectParameters()
            ShowProjectParameters()
        Else
            'ReadParentProjectParameters()
            ShowParentProjectParameters()
        End If
        DataGridView1.AutoResizeColumns()
    End Sub

    'Private Sub ReadProjectList()
    '    'Read the Project_List.xml file in the Application Directory.

    '    'If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'The latest ADVL_2 format version of the Project List file exists.
    '    'NOTE: Leave Project List file name as Project_List.xml. <FormatCode>ADVL_2</FormatCode> has been added but no need to change the file name.
    '    'NOTE: IGNOTE THE COMMENT ABOVE: Project_List.xml files will be converted to Project_List_ADVL_2.xml files.
    '    'If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'The latest ADVL_2 format version of the Project List file exists.
    '    If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'The latest ADVL_2 format version of the Project List file exists.
    '        Dim ProjectListXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")
    '        ReadProjectListAdvl_2(ProjectListXDoc)
    '    Else
    '        If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'The original ADVL_1 format version of the Project List file exists.
    '            RaiseEvent Message("Converting Project_List.xml to Project_List_ADVL_2.xml." & vbCrLf)
    '            'Convert the file to the latest ADVL_2 format:
    '            Dim Conversion As New ADVL_Utilities_Library_1.FormatConvert.ProjectListFileConversion
    '            Conversion.DirectoryPath = ApplicationDir
    '            Conversion.InputFileName = "Project_List.xml"
    '            Conversion.InputFormatCode = FormatConvert.ProjectListFileConversion.FormatCodes.ADVL_1
    '            Conversion.OutputFormatCode = FormatConvert.ProjectListFileConversion.FormatCodes.ADVL_2
    '            Conversion.Convert()
    '            If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then
    '                ReadProjectList() 'Try ReadProjectList again. This time Project_List_ADVL_2.xml should be found
    '            Else
    '                RaiseEvent ErrorMessage("Error converting Project_List.xml to Project_List_ADVL_2.xml." & vbCrLf)
    '            End If
    '        Else
    '            RaiseEvent ErrorMessage("No versions of the Project List were found." & vbCrLf)
    '        End If
    '    End If
    'End Sub

    'Private Sub ReadProjectList_Old()
    '    'Read the Project_List.xml file in the Application Directory.

    '    If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'Read the Project List.
    '        Dim ProjectListXDoc As System.Xml.Linq.XDocument
    '        ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List.xml")

    '        Dim Projects = From item In ProjectListXDoc.<ProjectList>.<Project>

    '        ProjectList.Clear()
    '        For Each item In Projects
    '            Dim NewProject As New ProjectSummary
    '            NewProject.Name = item.<Name>.Value
    '            NewProject.Description = item.<Description>.Value
    '            Select Case item.<Type>.Value
    '                Case "None"
    '                    NewProject.Type = Project.Types.None
    '                Case "Directory"
    '                    NewProject.Type = Project.Types.Directory
    '                Case "Archive"
    '                    NewProject.Type = Project.Types.Archive
    '                Case "Hybrid"
    '                    NewProject.Type = Project.Types.Hybrid
    '            End Select

    '            NewProject.Path = item.<Path>.Value

    '            NewProject.CreationDate = item.<CreationDate>.Value

    '            NewProject.AuthorName = item.<AuthorName>.Value

    '            ProjectList.Add(NewProject)
    '        Next

    '        UpdateProjectGrid()
    '    Else 'Create a Project List.

    '    End If
    'End Sub

    'Private Sub ReadProjectListAdvl_2(ByRef XDoc As System.Xml.Linq.XDocument)
    '    'Readt the Project List XDocument (ADVL_2 format version).

    '    Dim Projects = From item In XDoc.<ProjectList>.<Project>

    '    ProjectList.Clear()
    '    For Each item In Projects
    '        Dim NewProject As New ProjectSummary
    '        NewProject.Name = item.<Name>.Value
    '        NewProject.Description = item.<Description>.Value
    '        Select Case item.<Type>.Value
    '            Case "None"
    '                NewProject.Type = Project.Types.None
    '            Case "Directory"
    '                NewProject.Type = Project.Types.Directory
    '            Case "Archive"
    '                NewProject.Type = Project.Types.Archive
    '            Case "Hybrid"
    '                NewProject.Type = Project.Types.Hybrid
    '        End Select

    '        NewProject.Path = item.<Path>.Value

    '        NewProject.CreationDate = item.<CreationDate>.Value

    '        NewProject.AuthorName = item.<AuthorName>.Value

    '        ProjectList.Add(NewProject)
    '    Next

    '    UpdateProjectGrid()

    'End Sub

    'Private Sub UpdateProjectGrid()
    '    'Update DataGridView1 - Show the Name, Type and Description of each project.

    '    DataGridView1.Rows.Clear()

    '    Dim NProjects As Integer = ProjectList.Count

    '    If NProjects = 0 Then
    '        Exit Sub
    '    End If

    '    Dim Index As Integer

    '    For Index = 0 To NProjects - 1
    '        DataGridView1.Rows.Add()
    '        DataGridView1.Rows(Index).Cells(0).Value = ProjectList(Index).Name
    '        DataGridView1.Rows(Index).Cells(1).Value = ProjectList(Index).Type
    '        DataGridView1.Rows(Index).Cells(2).Value = ProjectList(Index).Description
    '    Next

    '    DataGridView1.AutoResizeColumns()

    'End Sub

    'Private Sub DisplayRecordNo(ByVal RecordNo As Integer)
    '    'Displays a record from ProjectList.

    '    If RecordNo < 1 Then
    '        RaiseEvent ErrorMessage("Cannot display Unit Of Measure data. Selected record number is too small." & vbCrLf)
    '        Exit Sub
    '    End If

    '    If RecordNo > ProjectList.Count Then
    '        RaiseEvent ErrorMessage("Cannot display Unit Of Measure data. Selected record number is too large." & vbCrLf)
    '        Exit Sub
    '    End If

    '    txtProjectName.Text = ProjectList(RecordNo - 1).Name
    '    txtDescription.Text = ProjectList(RecordNo - 1).Description
    '    txtAuthor.Text = ProjectList(RecordNo - 1).AuthorName
    '    txtType.Text = ProjectList(RecordNo - 1).Type
    '    txtCreationDate.Text = ProjectList(RecordNo - 1).CreationDate
    '    txtType.Text = ProjectList(RecordNo - 1).Type
    '    txtPath.Text = ProjectList(RecordNo - 1).Path


    'End Sub



    'Private Sub AddProjectForm_ProjectToAdd(ProjectSummary As ProjectSummary) Handles AddProjectForm.ProjectToAdd
    '    ProjectList.Add(ProjectSummary) 'Add the project summary to the list.
    '    'NProjects += 1                  'Increment the number of projects
    '    UpdateProjectGrid()             'Update the list of projects displayed in DataGridView1
    'End Sub



    'Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    '    Dim RowNo As Integer = e.RowIndex

    '    DataGridView1.Rows(RowNo).Selected = True
    '    ShowProjectInfo(RowNo)

    'End Sub

    'Private Sub ShowProjectInfo(ByVal RecordNo As Integer)

    '    If RecordNo >= ProjectList.Count Then
    '        RaiseEvent ErrorMessage("Error displaying project information. Record number too large." & vbCrLf)
    '        Exit Sub
    '    End If

    '    If RecordNo = -1 Then
    '        'No project selected.
    '        Exit Sub
    '    End If

    '    txtProjectName.Text = ProjectList(RecordNo).Name
    '    txtDescription.Text = ProjectList(RecordNo).Description
    '    Select Case ProjectList(RecordNo).Type
    '        Case Project.Types.Archive
    '            txtType.Text = "Archive"
    '            If System.IO.File.Exists(ProjectList(RecordNo).Path) Then
    '                txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
    '                txtComments.ForeColor = Drawing.Color.Black
    '                txtComments.Text = "Archive found."
    '            Else
    '                txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
    '                txtComments.ForeColor = Drawing.Color.Red
    '                txtComments.Text = "Archive not found!"
    '            End If
    '        Case Project.Types.Directory
    '            txtType.Text = "Directory"
    '            If System.IO.Directory.Exists(ProjectList(RecordNo).Path) Then
    '                txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
    '                txtComments.ForeColor = Drawing.Color.Black
    '                txtComments.Text = "Directory found."
    '            Else
    '                txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
    '                txtComments.ForeColor = Drawing.Color.Red
    '                txtComments.Text = "Directory not found!"
    '            End If
    '        Case Project.Types.Hybrid
    '            txtType.Text = "Hybrid"
    '            If System.IO.Directory.Exists(ProjectList(RecordNo).Path) Then
    '                txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
    '                txtComments.ForeColor = Drawing.Color.Black
    '                txtComments.Text = "Directory found."
    '            Else
    '                txtComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Bold)
    '                txtComments.ForeColor = Drawing.Color.Red
    '                txtComments.Text = "Directory not found!"
    '            End If
    '        Case Project.Types.None
    '            txtType.Text = "None"
    '    End Select
    '    txtCreationDate.Text = ProjectList(RecordNo).CreationDate
    '    txtAuthor.Text = ProjectList(RecordNo).AuthorName
    '    txtPath.Text = ProjectList(RecordNo).Path

    'End Sub

    'Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
    '    'Remove the selected project from the list.

    '    Dim RowNo As Integer

    '    If DataGridView1.SelectedRows.Count > 0 Then
    '        RowNo = DataGridView1.SelectedRows(0).Index
    '        ProjectList.RemoveAt(RowNo)
    '        UpdateProjectGrid()
    '    Else
    '        RaiseEvent ErrorMessage("No project has been selected for removal!" & vbCrLf)
    '    End If
    'End Sub



    'Private Sub btnSelDefault_Click(sender As Object, e As EventArgs) Handles btnSelDefault.Click

    '    'Check if ProjectList contains a project named Default
    '    Dim SelectedProject As ProjectSummary
    '    SelectedProject = ProjectList.Find(Function(p) p.Name = "Default")

    '    'If SelectedProject.Name = "Default" Then
    '    If SelectedProject Is Nothing Then
    '        'The Default project does not exist.
    '        RaiseEvent CreateDefaultProject() 'Create the Default project. This also adds the project to the list.
    '        ReadProjectList() 'Read the updated project list.
    '        UpdateProjectGrid() 'Úpdate the project grid.
    '    Else
    '        'The Default project is already on the list.

    '    End If
    'End Sub





#End Region 'Form Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event ProjectSelected(ByRef ProjectSummary As ProjectSummary) 'Raise an event passing the selected project information.
    'Public Event ErrorMessage(ByVal Message As String) 'Send an error message.
    'Public Event Message(ByVal Message As String) 'Send a message.
    'Public Event OpenDefaultProject() 'Raise an event to open the Default_Project.
    'Public Event CreateDefaultProject() 'Raise an event to create the Default_Project (if it does not already exist).

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class