Public Class frmMessages

#Region " Variable Declarations - All the variables used in this form and this application." '-----------------------------------------------------------------------------------------------

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public ApplicationName As String 'The name of the application using the message form.

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Process XML files - Read and write XML files." '-----------------------------------------------------------------------------------------------------------------------------------

    Private Sub SaveXmlFormSettings()
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

        'RaiseEvent SaveFormSettings(Me.Text, Settings)
        'Dim DataName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & "Messages" & ".xml"
        'SettingsLocn.SaveXmlData(DataName, Settings)
        SettingsLocn.SaveXmlData(SettingsFileName, Settings)
    End Sub

    Public Sub RestoreXmlFormSettings()
        'Read the form settings from an XML document.

        'Dim DataName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & "Messages" & ".xml"
        Dim Settings As System.Xml.Linq.XDocument
        'SettingsLocn.ReadXmlData(DataName, Settings)
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

    End Sub



#End Region 'Process XML Files --------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Subroutines - Code used to display this form." '------------------------------------------------------------------------------------------------------------------------------

    Private Sub frmMessages_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreXmlFormSettings()
        Me.Text = ApplicationName & " Messages"
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        'SaveXmlFormSettings() 'Save the form settings NOTE: the form can close without the exit button pressed! Save the settings when the form is closing (see code below).
        Me.Close()
    End Sub

    Private Sub frmMessages_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Form closing: save the form settings.

        SaveXmlFormSettings() 'Save the form settings
    End Sub

#End Region 'Form Subroutines ---------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '---------------------------------------------------------------------------------------------------------------------------

    Private Sub btnClearMessages_Click(sender As Object, e As EventArgs) Handles btnClearMessages.Click
        rtbMessages.Clear()
    End Sub

    Private Sub btnClearInstructions_Click(sender As Object, e As EventArgs) Handles btnClearInstructions.Click
        rtbInstructions.Clear()
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event SaveFormSettings(ByVal FormName As String, ByRef Settings As System.Xml.Linq.XDocument) 'Raise an event to save the form settings. The settings are contained in the Settings XML document.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------






End Class