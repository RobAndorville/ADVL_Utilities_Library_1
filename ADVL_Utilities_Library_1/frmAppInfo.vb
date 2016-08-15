Public Class frmAppInfo
    'The AppInfo form is used to display information about an Andorville Labs application.

#Region " Variable Declarations - All the variables used in this form." '--------------------------------------------------------------------------------------------------------------------

    'Public SettingsLocn As New Location 'This is a directory or archive where settings are stored.

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public ApplicationName As String 'The name of the application using the message form.

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form" '-------------------------------------------------------------------------------------------------------------------------------

    Private _nAssoc As Integer 'The number of File Association records.
    Property NAssoc As Integer
        Get
            Return _nAssoc
        End Get
        Set(value As Integer)
            _nAssoc = value
            txtNAssns.Text = _nAssoc
        End Set
    End Property

    Private _selectedAssocNo As Integer 'The number of the selected file association record.
    Property SelectedAssocNo As Integer
        Get
            Return _selectedAssocNo
        End Get
        Set(value As Integer)
            If _nAssoc = 0 Then 'There are no records to display.
                _selectedAssocNo = 0
                txtExtension.Text = ""
                txtFileAssocationDesc.Text = ""
                txtAssnNo.Text = "0"
            Else
                If value < 1 Then 'Display the first record.
                    _selectedAssocNo = 1
                    txtAssnNo.Text = "1"
                    RaiseEvent DisplayAssociation(1) 'Display the first File Association record.
                Else
                    If value > _nAssoc Then 'Display the last record.
                        _selectedAssocNo = _nAssoc
                        txtAssnNo.Text = _nAssoc
                        RaiseEvent DisplayAssociation(_nAssoc) 'Display the first File Association record.
                    Else
                        _selectedAssocNo = value
                        txtAssnNo.Text = value
                        RaiseEvent DisplayAssociation(value) 'Display the first File Association record.
                    End If
                End If
            End If
        End Set
    End Property

    Private _nLibraries As Integer 'The number of software libraries.
    Property NLibraries As Integer
        Get
            Return _nLibraries
        End Get
        Set(value As Integer)
            _nLibraries = value
            txtNLibs.Text = value
        End Set
    End Property

    Private _selectedLibraryNo As Integer 'The number of the selected library.
    Property SelectedLibraryNo As Integer
        Get
            Return _selectedLibraryNo
        End Get
        Set(value As Integer)

            If _nLibraries = 0 Then 'There are no records to display.
                _selectedLibraryNo = 0
                txtExtension.Text = ""
                txtFileAssocationDesc.Text = ""
                txtLibNo.Text = "0"
            Else
                If value < 1 Then 'Display the first record.
                    _selectedLibraryNo = 1
                    txtLibNo.Text = "1"
                    RaiseEvent DisplayLibrary(1) 'Display the first File Association record.
                Else
                    If value > _nLibraries Then 'Display the last record.
                        _selectedLibraryNo = _nLibraries
                        txtLibNo.Text = _nLibraries
                        RaiseEvent DisplayLibrary(_nLibraries) 'Display the first File Association record.
                    Else
                        _selectedLibraryNo = value
                        txtLibNo.Text = value
                        RaiseEvent DisplayLibrary(value) 'Display the first File Association record.
                    End If
                End If
            End If




        End Set
    End Property

    Private _nClasses As Integer 'The number of classes in the current library
    Property NClasses As Integer
        Get
            Return _nClasses
        End Get
        Set(value As Integer)
            _nClasses = value
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
                           <Height><%= Me.Height %></Height>
                           <!---->
                       </FormSettings>

        'Dim SettingsName As String = "FormSettings_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        SettingsLocn.SaveXmlData(SettingsFileName, Settings)

    End Sub

    Public Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsName As String = "FormSettings_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"

        'If SettingsLocn.FileExists(SettingsName) Then
        Dim Settings As System.Xml.Linq.XDocument
        'SettingsLocn.ReadXmlData(SettingsName, Settings)
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
        'Else
        ''Settings file not found.
        'End If
    End Sub

#End Region 'Process XML Files --------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '--------------------------------------------------------------------------------------------------------------------------

    Private Sub frmAppInfo_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        SaveFormSettings() 'Save the form settings.
        Me.Close()


        'rtbLicenceNotice.Sel
    End Sub

#End Region 'Form Display Methods -----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Form methods - The main actions performed by this form" ' --------------------------------------------------------------------------------------------------------------------------

    Private Sub btnPrevExt_Click(sender As Object, e As EventArgs) Handles btnPrevExt.Click
        'Show the previous file association record:
        SelectedAssocNo -= 1
    End Sub

    Private Sub btnNextExt_Click(sender As Object, e As EventArgs) Handles btnNextExt.Click
        'Show the next file assoication record:
        SelectedAssocNo += 1
    End Sub

    Private Sub btnPrevClass_Click(sender As Object, e As EventArgs) Handles btnPrevClass.Click
        'Show the previous software library in the list:
        SelectedLibraryNo -= 1
    End Sub

    Private Sub btnNextClass_Click(sender As Object, e As EventArgs) Handles btnNextClass.Click
        'Show the next software library in the list:
        SelectedLibraryNo += 1
    End Sub



#End Region 'Form Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------------



   
 
    'Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
    '    RaiseEvent RunTest()
    'End Sub

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event RunTest() 'Raise an event to run the test code.

    Public Event DisplayAssociation(ByVal RecordNo As Integer) 'Display the File Association record.
    Public Event DisplayLibrary(ByVal RecordNo As Integer) 'Display the Library record.

#End Region



  
   
  
  
End Class