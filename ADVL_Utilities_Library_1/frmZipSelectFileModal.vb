﻿Public Class frmZipSelectFileModal
    'The ZipSelectFileModal form is used to select a file from the specified Zip archive.
    'It can be displayed as a modal form.

#Region " Variable Declarations - All the variables used in this form." '--------------------------------------------------------------------------------------------------------------------

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public ApplicationName As String 'The name of the application using this form.

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '----------------------------------------------------------------------------------------------------------

    Private _zipArchivePath As String
    Property ZipArchivePath As String 'The path of the selected Zip archive file.
        Get
            Return _zipArchivePath
        End Get
        Set(value As String)
            _zipArchivePath = value
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

    Private _fileName As String 'The name of the selected file.
    Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
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

    'Private Sub frmProject_Load(sender As Object, e As EventArgs) Handles Me.Load

    '    'The application that opens this form also initialises the form.

    'End Sub

    'Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
    '    SaveFormSettings() 'Save the form settings.
    '    Me.Close()
    'End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        'File selection has been cancelled
        FileName = ""
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        'A file has been selected.
        FileName = ListBox1.Text
        Me.DialogResult = DialogResult.OK
    End Sub

#End Region 'Form Display Subroutines -------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Form methods - The main actions performed by this form" ' --------------------------------------------------------------------------------------------------------------------------

    Public Sub GetFileList()
        'Get the list of entries in the archive.

        ListBox1.Items.Clear()

        'If ZipArchivePath = "" Then
        '    RaiseEvent ErrorMessage("No zip archive selected." & vbCrLf)
        '    Exit Sub
        'End If

        Dim Zip As New ZipComp
        Zip.ArchivePath = ZipArchivePath
        Dim ValidExtension As String = FileExtension
        If ValidExtension.StartsWith(".") Then
        Else
            ValidExtension = "." & ValidExtension
        End If
        Dim Index As Integer
        Dim EntryName As String
        For Index = 1 To Zip.NEntries
            If FileExtension = "" Then
                ListBox1.Items.Add(Zip.GetEntryName(Index - 1))
            Else
                EntryName = Zip.GetEntryName(Index - 1)
                If EntryName.EndsWith(ValidExtension) Then
                    ListBox1.Items.Add(Zip.GetEntryName(Index - 1))
                End If
            End If
        Next
    End Sub

    'Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
    '    RaiseEvent FileSelected(ListBox1.Text)
    'End Sub

#End Region 'Form Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event FileSelected(ByVal FileName As String) 'Raise an event passing the name of the selected file.
    'Public Event ErrorMessage(ByVal Message As String) 'Send an error message.
    'Public Event Message(ByVal Message As String) 'Send a message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------



End Class