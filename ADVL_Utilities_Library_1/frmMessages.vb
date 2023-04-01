Imports System.Windows.Forms

Public Class frmMessages

#Region " Variable Declarations - All the variables used in this form and this application." '-----------------------------------------------------------------------------------------------

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    'Public ProjectLocn As New ADVL_Utilities_Library_1.FileLocation 'The project location.
    Public ApplicationName As String 'The name of the application using the message form.

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------



#Region " Process XML files - Read and write XML files." '-----------------------------------------------------------------------------------------------------------------------------------

    'Private Sub SaveXmlFormSettings()
    '    'Save the form settings as an XML document.
    '    Dim Settings = <?xml version="1.0" encoding="utf-8"?>
    '                   <!---->
    '                   <!--Form settings for Project form.-->
    '                   <FormSettings>
    '                       <Left><%= Me.Left %></Left>
    '                       <Top><%= Me.Top %></Top>
    '                       <Width><%= Me.Width %></Width>
    '                       <Height><%= Me.Height %></Height>
    '                       <!---->
    '                   </FormSettings>

    '    'RaiseEvent SaveFormSettings(Me.Text, Settings)
    '    'Dim DataName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
    '    'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
    '    Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & "Messages" & ".xml"
    '    'SettingsLocn.SaveXmlData(DataName, Settings)
    '    SettingsLocn.SaveXmlData(SettingsFileName, Settings)
    'End Sub

    'Private Sub SaveFormSettings()
    Public Sub SaveFormSettings()
        'Save the form settings as an XML document.
        '

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

        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & Me.Text & ".xml"
        SettingsLocn.SaveXmlData(SettingsFileName, Settings)
        'ProjectLocn.SaveXmlData(SettingsFileName, Settings)

        Debug.Print("SaveFormSettings")
        Debug.Print("SettingsFileName: " & SettingsFileName)

    End Sub

    'Public Sub RestoreXmlFormSettings()
    '    'Read the form settings from an XML document.

    '    'Dim DataName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
    '    'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"
    '    Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & "Messages" & ".xml"
    '    Dim Settings As System.Xml.Linq.XDocument
    '    'SettingsLocn.ReadXmlData(DataName, Settings)
    '    SettingsLocn.ReadXmlData(SettingsFileName, Settings)

    '    If Settings Is Nothing Then
    '        Exit Sub
    '    End If

    '    'Restore form position and size:
    '    If Settings.<FormSettings>.<Left>.Value = Nothing Then
    '        'Form setting not saved.
    '    Else
    '        Me.Left = Settings.<FormSettings>.<Left>.Value
    '    End If

    '    If Settings.<FormSettings>.<Top>.Value = Nothing Then
    '        'Form setting not saved.
    '    Else
    '        Me.Top = Settings.<FormSettings>.<Top>.Value
    '    End If

    '    If Settings.<FormSettings>.<Height>.Value = Nothing Then
    '        'Form setting not saved.
    '    Else
    '        Me.Height = Settings.<FormSettings>.<Height>.Value
    '    End If

    '    If Settings.<FormSettings>.<Width>.Value = Nothing Then
    '        'Form setting not saved.
    '    Else
    '        Me.Width = Settings.<FormSettings>.<Width>.Value
    '    End If
    '    CheckFormPos()
    'End Sub

    Public Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "Formsettings_" & ApplicationName & "_" & Me.Text & ".xml"

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationName & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & Me.Text & ".xml"

        Dim Settings As System.Xml.Linq.XDocument

        If SettingsLocn Is Nothing Then Exit Sub 'This can occur when attempting to write a message before the Message form is set up.

        SettingsLocn.ReadXmlData(SettingsFileName, Settings)
        'ProjectLocn.ReadXmlData(SettingsFileName, Settings)
        Debug.Print("RestoreFormSettings")
        Debug.Print("SettingsFileName: " & SettingsFileName)

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
        'Else
        ''Settings file not found.
        'End If

    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        'Dim MinWidthVisible As Integer = 48 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        'Dim MinWidthVisible As Integer = 128 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        'Dim MinHeightVisible As Integer = 48 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.
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


#Region " Form Subroutines - Code used to display this form." '------------------------------------------------------------------------------------------------------------------------------

    Private Sub frmMessages_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Debug.Print("Running: frmMessages_Load")
        'RestoreXmlFormSettings()
        'RestoreFormSettings()
        'Me.Text = ApplicationName & " Messages"
        Me.Text = ApplicationName & " - Messages"

        btnCopyHtmlMarkup.Enabled = False

        RestoreFormSettings() 'Move this to after Me.Text is changed - otherwise the wrong settings file name is used!

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

        XmlHtmDisplay1.Settings.AddNewTextType("Bold")
        XmlHtmDisplay1.Settings.TextType("Bold").Bold = True
        XmlHtmDisplay1.Settings.TextType("Bold").Color = System.Drawing.Color.Black
        XmlHtmDisplay1.Settings.TextType("Bold").PointSize = 10

        XmlHtmDisplay1.Settings.AddNewTextType("Warning")
        XmlHtmDisplay1.Settings.TextType("Warning").Bold = True
        XmlHtmDisplay1.Settings.TextType("Warning").Color = System.Drawing.Color.Red
        XmlHtmDisplay1.Settings.TextType("Warning").PointSize = 10

        XmlHtmDisplay1.Settings.AddNewTextType("Heading")
        XmlHtmDisplay1.Settings.TextType("Heading").Bold = True
        XmlHtmDisplay1.Settings.TextType("Heading").Color = System.Drawing.Color.Black
        XmlHtmDisplay1.Settings.TextType("Heading").PointSize = 12

        XmlHtmDisplay1.Settings.AddNewTextType("Heading 11pt")
        XmlHtmDisplay1.Settings.TextType("Heading 11pt").Bold = True
        XmlHtmDisplay1.Settings.TextType("Heading 11pt").Color = System.Drawing.Color.Black
        XmlHtmDisplay1.Settings.TextType("Heading 11pt").PointSize = 11

        XmlHtmDisplay1.Settings.UpdateFontIndexes()
        XmlHtmDisplay1.Settings.UpdateColorIndexes()



    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        'SaveXmlFormSettings() 'Save the form settings NOTE: the form can close without the exit button pressed! Save the settings when the form is closing (see code below).
        Me.Close()
    End Sub

    Private Sub frmMessages_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Form closing: save the form settings.

        'SaveXmlFormSettings() 'Save the form settings
        SaveFormSettings() 'Save the form settings
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

    Private Sub chkShowXMessages_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowXMessages.CheckedChanged
        If chkShowXMessages.Focused Then
            If chkShowXMessages.Checked Then
                RaiseEvent ShowXMessages(True)
            Else
                RaiseEvent ShowXMessages(False)
            End If
        End If
    End Sub

    Private Sub chkShowSysMessages_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowSysMessages.CheckedChanged
        If chkShowSysMessages.Focused Then
            If chkShowSysMessages.Checked Then
                RaiseEvent ShowSysMessages(True)
            Else
                RaiseEvent ShowSysMessages(False)
            End If
        End If
    End Sub

    Private Sub btnCopyHtmlMarkup_Click(sender As Object, e As EventArgs) Handles btnCopyHtmlMarkup.Click
        'Copy the selected XML code to the Clipboard as Html markup.

        Dim SelText As String = XmlDisplay.SelectedText
        If SelText.Trim = "" Then
            RaiseEvent ErrorMessage("No XML code has been selected." & vbCrLf)
        Else
            Dim SelHtml As String = XmlDisplay.XmlToHtml(SelText, False)
            If SelHtml = "" Then
                RaiseEvent ErrorMessage("Generated HTML code is blank!" & vbCrLf)
            Else
                Clipboard.SetText(SelHtml, TextDataFormat.Text)
            End If
        End If
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        'The SelectedIndex of TabControl1 has changed.
        If TabControl1.SelectedIndex = 0 Then 'Messages tab selected.
            btnCopyHtmlMarkup.Enabled = False
        ElseIf TabControl1.SelectedIndex = 1 Then 'XMessage Instructions tab selected.
            btnCopyHtmlMarkup.Enabled = True
        ElseIf TabControl1.SelectedIndex = 2 Then 'Settings tab selected.
            btnCopyHtmlMarkup.Enabled = False
        End If
    End Sub

    Private Sub btnCopyText_Click(sender As Object, e As EventArgs) Handles btnCopyText.Click
        'Copy the selected text.
        If TabControl1.SelectedIndex = 0 Then 'Messages tab selected.
            Clipboard.SetText(XmlHtmDisplay1.SelectedText, TextDataFormat.Text)
        ElseIf TabControl1.SelectedIndex = 1 Then 'XMessage Instructions tab selected.
            Clipboard.SetText(XmlDisplay.SelectedText, TextDataFormat.Text)
        ElseIf TabControl1.SelectedIndex = 2 Then 'Settings tab selected.

        End If
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event SaveFormSettings(ByVal FormName As String, ByRef Settings As System.Xml.Linq.XDocument) 'Raise an event to save the form settings. The settings are contained in the Settings XML document.
    Event ErrorMessage(ByVal Message As String)
    Event Message(ByVal Message As String)
    Event ShowTextTypes()
    Event XShowTextTypes()
    Event ShowXMessages(ByVal Show As Boolean)
    Event ShowSysMessages(ByVal Show As Boolean)








#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------






End Class