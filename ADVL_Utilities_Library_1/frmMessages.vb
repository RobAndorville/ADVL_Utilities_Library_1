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
        'Debug.Print("Running: frmMessages_Load")
        RestoreXmlFormSettings()
        Me.Text = ApplicationName & " Messages"

        'Set up the XML Display:
        'XML formatting adjustments:
        XmlDisplay.Settings.IndentSpaces = 4
        XmlDisplay.Settings.Value.Bold = True
        XmlDisplay.Settings.Comment.Color = System.Drawing.Color.Gray

        'Message text settings:
        XmlDisplay.Settings.AddNewTextType("XmlReceivedNotice")
        XmlDisplay.Settings.TextType("XmlReceivedNotice").Bold = True
        XmlDisplay.Settings.TextType("XmlReceivedNotice").Color = System.Drawing.Color.Blue
        XmlDisplay.Settings.TextType("XmlReceivedNotice").PointSize = 10

        XmlDisplay.Settings.AddNewTextType("XmlSentNotice")
        XmlDisplay.Settings.TextType("XmlSentNotice").Bold = True
        XmlDisplay.Settings.TextType("XmlSentNotice").Color = System.Drawing.Color.Red
        XmlDisplay.Settings.TextType("XmlSentNotice").PointSize = 10

        XmlDisplay.Settings.AddNewTextType("Normal")
        XmlDisplay.Settings.TextType("Normal").Bold = False
        XmlDisplay.Settings.TextType("Normal").Color = System.Drawing.Color.Black
        XmlDisplay.Settings.TextType("Normal").PointSize = 10

        XmlDisplay.Settings.UpdateFontIndexes()
        XmlDisplay.Settings.UpdateColorIndexes()

        'Set up the Message display:

        'XML formatting adjustments:
        XmlHtmDisplay1.Settings.XIndentSpaces = 4
        XmlHtmDisplay1.Settings.XValue.Bold = True
        XmlHtmDisplay1.Settings.XComment.Color = System.Drawing.Color.Gray

        'Add new named text settings:
        XmlHtmDisplay1.Settings.AddNewTextType("Normal")
        XmlHtmDisplay1.Settings.TextType("Normal").Bold = False
        XmlHtmDisplay1.Settings.TextType("Normal").Color = System.Drawing.Color.Black
        XmlHtmDisplay1.Settings.TextType("Normal").PointSize = 10

        XmlHtmDisplay1.Settings.AddNewTextType("Warning")
        XmlHtmDisplay1.Settings.TextType("Warning").Bold = True
        XmlHtmDisplay1.Settings.TextType("Warning").Color = System.Drawing.Color.Red
        XmlHtmDisplay1.Settings.TextType("Warning").PointSize = 10

        XmlHtmDisplay1.Settings.AddNewTextType("Heading")
        XmlHtmDisplay1.Settings.TextType("Heading").Bold = True
        XmlHtmDisplay1.Settings.TextType("Heading").Color = System.Drawing.Color.Black
        XmlHtmDisplay1.Settings.TextType("Heading").PointSize = 12

        XmlHtmDisplay1.Settings.UpdateFontIndexes()
        XmlHtmDisplay1.Settings.UpdateColorIndexes()



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
        'rtbMessages.Clear()
        XmlHtmDisplay1.Clear()
    End Sub

    Private Sub btnClearInstructions_Click(sender As Object, e As EventArgs) Handles btnClearInstructions.Click
        'rtbInstructions.Clear()
        XmlDisplay.Clear()
    End Sub

    Private Sub rtbMessages_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub XmlDisplay_GotFocus(sender As Object, e As EventArgs) Handles XmlDisplay.GotFocus

    End Sub

    Private Sub XmlDisplay_Message(Msg As String) Handles XmlDisplay.Message
        RaiseEvent Message(Msg)
    End Sub

    Private Sub XmlDisplay_ErrorMessage(Msg As String) Handles XmlDisplay.ErrorMessage
        RaiseEvent ErrorMessage(Msg)
    End Sub

    Private Sub XmlHtmDisplay1_Message(Msg As String) Handles XmlHtmDisplay1.Message
        RaiseEvent Message(Msg)
    End Sub

    Private Sub XmlHtmDisplay1_ErrorMessage(Msg As String) Handles XmlHtmDisplay1.ErrorMessage
        RaiseEvent ErrorMessage(Msg)
    End Sub

    'Private Sub btnShowMsgTextTypes_Click(sender As Object, e As EventArgs) Handles btnShowMsgTextTypes.Click
    '    'The the text types defined for the Messages window.
    'THIS IS DONE IS SYSTEM_UTILTIIES MESSAGE Class

    '    Dim NTypes As Integer = XmlHtmDisplay1.Settings.TextType.Count
    '    Dim I As Integer

    '    For I = 0 To NTypes - 1
    '        'XmlHtmDisplay1.Selected
    '    Next


    'End Sub

    Private Sub btnShowMsgTextTypes_Click(sender As Object, e As EventArgs) Handles btnShowMsgTextTypes.Click
        RaiseEvent ShowTextTypes()
    End Sub

    Private Sub btnShowXMsgTextTypes_Click(sender As Object, e As EventArgs) Handles btnShowXMsgTextTypes.Click
        RaiseEvent XShowTextTypes()
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event SaveFormSettings(ByVal FormName As String, ByRef Settings As System.Xml.Linq.XDocument) 'Raise an event to save the form settings. The settings are contained in the Settings XML document.
    Event ErrorMessage(ByVal Message As String)
    Event Message(ByVal Message As String)
    Event ShowTextTypes()
    Event XShowTextTypes()









#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------






End Class