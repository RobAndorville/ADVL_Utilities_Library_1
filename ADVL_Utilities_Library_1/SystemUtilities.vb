﻿'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Imports System
Imports System.IO.Compression
Imports System.Windows.Forms
'Note: to access .ZipFile, use Project \ Add Reference \ Assemblies \  Framework \ System.IO.Compression
'Note: to access .ZipFile, use Project \ Add Reference \ Assemblies \  Framework \ System.IO.Compression.FileSystem
'Note: To access FontStyle, use Project \ Add Reference \ Assemblies \  Framework \ System.Drawing

'The ZipComp class is used to compress files into and extract files from a zip file.
Public Class ZipComp '------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable declarations - All the variables used in this class." '--------------------------------------------------------------------------------------------------------------------

    Public WithEvents SelectFileForm As frmZipSelectFile 'Used to select a project.

#End Region 'Variable Declarations -----------------------------------------------------------------------------------------------------------------------------------------------------------

    'PROPERTIES:
    'ArchivePath
    '  ArchiveDirectory
    '  ArchiveName
    'ArchiveExists
    'AddFileDirectory
    'AddFileName
    'ExtractFileDirectory
    'ExtractFileName

    'METHODS:
    'CreateArchive - Create a new archive. ArchivePath property is used to specify the archive name and directory.
    'AddFile       - Add a file to the archive. AddFileDirectory and AddFileName properties are used to specify the file name and directory.
    'ExtractFile   - Extract a file from the archive. ExtractFileDirectory and ExtractFileName properties are used to specify the file to extract and its destination directory.

#Region " Properties" '----------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private _archiveDirectory As String = "" 'The directory containing the selected archive. 
    Property ArchiveDirectory As String
        Get
            Return _archiveDirectory
        End Get
        Set(value As String)
            _archiveDirectory = value
            If _archiveName <> "" Then
                _archivePath = _archiveDirectory & "\" & _archiveName
                _archiveExists = IO.File.Exists(_archivePath)
            Else
                _archiveExists = False
            End If
        End Set
    End Property

    Private _archiveName As String = "" 'The name of the selected archive. 
    Property ArchiveName As String
        Get
            Return _archiveName
        End Get
        Set(value As String)
            _archiveName = value
            If _archiveDirectory = "" Then
                _archiveExists = False
            Else
                _archivePath = _archiveDirectory & "\" & _archiveName
                _archiveExists = IO.File.Exists(_archivePath)
            End If
        End Set
    End Property

    Private _archivePath As String = "" 'The path of the selected archive.
    Property ArchivePath As String
        Get
            Return _archivePath
        End Get
        Set(value As String)
            _archivePath = value
            If _archivePath = "" Then
                _archiveExists = False
            Else
                _archiveDirectory = System.IO.Path.GetDirectoryName(_archivePath)
                _archiveName = System.IO.Path.GetFileName(_archivePath)
                _archiveExists = IO.File.Exists(_archivePath)
            End If
        End Set
    End Property

    Private _archiveExists As Boolean = False 'ArchiveExists is True if the archive file having path ArchivePath exists.
    ReadOnly Property ArchiveExists As Boolean
        Get
            Return _archiveExists
        End Get
    End Property

    Private _newArchiveDirectory As String = "" 'When a new archive is created, it will be created in this directory.
    Property NewArchiveDirectory As String
        Get
            Return _newArchiveDirectory
        End Get
        Set(value As String)
            _newArchiveDirectory = value
        End Set
    End Property

    Private _newArchiveName As String = "" 'The name of the new archive to be created.
    Property NewArchiveName As String
        Get
            Return _newArchiveName
        End Get
        Set(value As String)
            _newArchiveName = value
            If _newArchiveDirectory = "" Then
                _newArchiveExists = False
            Else
                _newArchivePath = _newArchiveDirectory & "\" & _newArchiveName
                _newArchiveExists = IO.File.Exists(_newArchivePath)
            End If
        End Set
    End Property

    Private _newArchivePath As String = "" 'The path of the new archive to be created.
    Property NewArchivePath As String
        Get
            Return _newArchivePath
        End Get
        Set(value As String)
            _newArchivePath = value
            If _newArchivePath = "" Then
                _newArchiveExists = False
            Else
                _newArchiveDirectory = System.IO.Path.GetDirectoryName(_newArchivePath)
                _newArchiveName = System.IO.Path.GetFileName(_newArchivePath)
                _newArchiveExists = IO.File.Exists(_newArchivePath)
            End If
        End Set
    End Property

    Private _newArchiveExists As Boolean = False 'NewArchiveExists is True if the archive file having path NewArchivePath exists.
    ReadOnly Property NewArchiveExists As Boolean
        Get
            Return _newArchiveExists
        End Get
    End Property

    Private _addFileDirectory As String = "" 'The directory containing a file to add to the archive.
    Property AddFileDirectory As String
        Get
            Return _addFileDirectory
        End Get
        Set(value As String)
            _addFileDirectory = value
        End Set
    End Property

    Private _addFileName As String = "" 'The name of a file to add to the directory.
    Property AddFileName As String
        Get
            Return _addFileName
        End Get
        Set(value As String)
            _addFileName = value
        End Set
    End Property

    Private _extractFileDirectory As String = "" 'The directory to place a file extracted from the archive.
    Property ExtractFileDirectory As String
        Get
            Return _extractFileDirectory
        End Get
        Set(value As String)
            _extractFileDirectory = value
        End Set
    End Property

    Private _extractFileName As String = "" 'The name of a file to extract from the archive.
    Property ExtractFileName As String
        Get
            Return _extractFileName
        End Get
        Set(value As String)
            _extractFileName = value
        End Set
    End Property

    Private _selectedFile As String = "" 'The name of a file selected on the ZipSelectFile form.
    Property SelectedFile As String
        Get
            Return _selectedFile
        End Get
        Set(value As String)
            _selectedFile = value
        End Set
    End Property



    'Private _zipFilePath As String = "" 'The path of the zip file.
    'Property ZipFilePath As String
    '    Get
    '        Return _zipFilePath
    '    End Get
    '    Set(value As String)
    '        _zipFilePath = Trim(value)
    '        If _zipFilePath.EndsWith("\") Then
    '            RaiseEvent ErrorMsg("Zip file path is not valid: ends with \.")
    '            _zipFilePath = ""
    '        End If
    '        _zipFileExists = IO.File.Exists(_zipFilePath)
    '    End Set
    'End Property

    'Private _zipFileExists As Boolean = False 'ZipFileExists is true if the  zip file having path ZipFilePath exists
    ''_zipFileExists is set to True or False when the ZipFilePath property is set.
    'Public ReadOnly Property ZipFileExists As Boolean   'ZipFileExists is true if the  zip file having path ZipFilePath exists
    '    Get
    '        Return _zipFileExists
    '    End Get
    'End Property




#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub CreateArchive()
        'Create a new archive at the path ArchivePath.

        'If ArchiveExists Then
        If NewArchiveExists Then
            'Specified archive already exists.
            RaiseEvent ErrorMessage("Archive already exists.")
        Else
            'Dim archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Create)
            Dim archive As ZipArchive = ZipFile.Open(NewArchivePath, ZipArchiveMode.Create)
            archive.Dispose()
        End If
    End Sub

    Public Sub AddFile()
        'Add the specified file to the archive.
        'The file is in directory AddFileDirectory.
        'The file has the name AddFileName.

        Try
            Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Update)
                'Dim AddFileName As String = System.IO.Path.GetFileName(AddFilePath)
                'archive.CreateEntryFromFile(AddFilePath, AddFileName, CompressionLevel.Fastest)
                archive.CreateEntryFromFile(AddFileDirectory & "\" & AddFileName, AddFileName, CompressionLevel.Fastest)
            End Using

        Catch ex As Exception
            'MessageBox.Show("Error adding file to archive: " & ex.Message)
            RaiseEvent ErrorMessage("Error adding file to archive: " & ex.Message)
        End Try

        'Try
        '    Using archive As ZipArchive = ZipFile.Open(ZipFilePath, ZipArchiveMode.Update)
        '        Dim AddFileName As String = System.IO.Path.GetFileName(AddFilePath)
        '        archive.CreateEntryFromFile(AddFilePath, AddFileName, CompressionLevel.Fastest)
        '    End Using

        'Catch ex As Exception
        '    MessageBox.Show("Error adding file to archive: " & ex.Message)
        'End Try

    End Sub

    Public Sub ExtractFile()
        'Extract the specified file from the directory.
        'The file to be extracted has the name ExtractFileName.
        'The file will be extracted into the directory ExtractFileDirectory.

        Try
            Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                Dim entry As ZipArchiveEntry = archive.GetEntry(ExtractFileName)
                If System.IO.File.Exists(System.IO.Path.Combine(ExtractFileDirectory, entry.FullName)) Then
                    RaiseEvent ErrorMessage("File already exists at location: " & System.IO.Path.Combine(ExtractFileDirectory, entry.FullName))
                Else
                    entry.ExtractToFile(System.IO.Path.Combine(ExtractFileDirectory, entry.FullName))
                End If
            End Using
        Catch ex As Exception
            RaiseEvent ErrorMessage("Error extracting file from archive: " & ex.Message)
        End Try
    End Sub

    Public Sub AddText(ByVal TextFileName As String, ByRef TextData As String)
        'Add an entry to the archive with the name TextFileName and containing TextData.

        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank.")
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Update)
                    'First remove existing entries of the same name:
                    Do
                        Dim entry As ZipArchiveEntry = archive.GetEntry(TextFileName)
                        If IsNothing(entry) Then
                            Exit Do
                        Else
                            entry.Delete()
                        End If
                    Loop

                    'Add the new entry:
                    Dim newEntry As ZipArchiveEntry = archive.CreateEntry(TextFileName)
                    'Dim bytesToAdd As Byte() = System.Text.Encoding.ASCII.GetBytes(TextData)
                    Dim bytesToAdd As Byte() = System.Text.Encoding.UTF8.GetBytes(TextData)
                    newEntry.Open.Write(bytesToAdd, 0, bytesToAdd.Length)
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error with AddText: " & vbCrLf & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub AddData(ByVal DataFileName As String, ByRef Stream As IO.Stream)
        'Save the data in the Stream to an entry named DataFileName in the Archive file.

        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank.")
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Update)
                    'First remove existing entries of the same name:
                    Do
                        Dim entry As ZipArchiveEntry = archive.GetEntry(DataFileName)
                        If IsNothing(entry) Then
                            Exit Do
                        Else
                            entry.Delete()
                        End If
                    Loop

                    'Add the new entry:
                    Dim newEntry As ZipArchiveEntry = archive.CreateEntry(DataFileName)

                    'newEntry.Open.Write(Stream, 0, Stream.Length)
                    'Dim bytesToAdd As Byte() = Stream.
                    'newEntry.Open.Write()
                    'Dim writer As IO.StreamWriter = New IO.StreamWriter(Stream)

                    'Dim writer2 As New IO.StreamWriter(newEntry.Open)
                    'writer2.Write(Stream)

                    'Dim entryStream As IO.Stream()
                    Stream.CopyTo(newEntry.Open)

                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error writing Data to Archive:" & ex.Message & vbCrLf)
            End Try
        End If

    End Sub

    Public Function GetText(ByVal TextFileName As String) As String
        'Get the text from an entry with the name TextFileName.

        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank.")
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                    Dim entry As ZipArchiveEntry = archive.GetEntry(TextFileName)
                    'Dim bytesToGet(entry.Length) As Byte
                    Dim bytesToGet(entry.Length - 1) As Byte 'THIS WORKS - Temp Disable for testing
                    entry.Open.Read(bytesToGet, 0, entry.Length)
                    'entry.Open.Read(bytesToGet, 0, entry.Length - 1)


                    GetText = System.Text.Encoding.UTF8.GetString(bytesToGet)


                    'GetText start with <!---->
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error with GetText: " & vbCrLf & ex.Message & vbCrLf)
            End Try
        End If

    End Function

    Public Function GetData(ByVal DataFilename As String) As IO.Stream
        'Get the data from an entry witht he name DataFileName.

        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank.")
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                    Dim entry As ZipArchiveEntry = archive.GetEntry(DataFilename)

                    'GetData = entry.Open
                    'Dim myStream As IO.Stream
                    'entry.Open().CopyTo(GetData)
                    'entry.Open().CopyTo(myStream)
                    'Return myStream
                    'entry.Open()
                    'entry.Open.CopyTo(myStream)
                    'GetData = myStream

                    'entry.Open.CopyTo(GetData)
                    'entry.Open.CopyToAsync(GetData)

                    'GetData = entry.Open

                    'Dim myStream As IO.Stream = entry.Open
                    'myStream.Position = 0
                    'GetData = myStream


                    'GetData = entry.Open 'Unable to cast object of type 'System.IO.Compression.DeflateStream' to type 'System.IO.MemoryStream'.
                    'Dim defStream As IO.Compression.DeflateStream = entry.Open
                    'Dim defStream As IO.Compression.DeflateStream
                    'defStream = entry.Open

                    'entry.Open.CopyTo(defStream) 'Error 
                    'entry.Open().CopyTo(defStream) 'Error

                    'If defStream Is Nothing Then
                    '    Exit Function
                    'End If

                    'Dim Len As Integer = defStream.Length

                    'defStream.Position = 0 'Error
                    'defStream.CopyTo(GetData) 'Error

                    'Dim myStream As IO.Stream
                    'defStream.CopyTo(myStream)

                    'Return myStream

                    'Dim reader As IO.StreamReader = New IO.StreamReader(entry.Open)

                    'Dim Len As Long = reader.BaseStream.Length 'Operation not supported

                    'reader.BaseStream.CopyTo(GetData) 'Value cannot be a null

                    'Dim stream As IO.Stream
                    'stream = entry.Open
                    'Dim reader As New IO.StreamReader(stream)

                    'GetData = reader.ReadToEnd
                    'reader.BaseStream.CopyTo(GetData)

                    'Return reader.BaseStream

                    Dim bytesToGet(entry.Length - 1) As Byte
                    entry.Open.Read(bytesToGet, 0, entry.Length)

                    'Dim stream As IO.Stream = MemoryStream(bytesToGet)

                    Dim myStream As IO.MemoryStream = New IO.MemoryStream(bytesToGet)

                    'GetData = bytesToGet.
                    GetData = myStream

                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error with GetData: " & vbCrLf & ex.Message & vbCrLf)
            End Try

        End If

    End Function

    Public Function EntryExists(ByVal EntryName As String) As Boolean
        'Return True if the archive contains an entry with the name EntryName.
        If ArchivePath = "" Then
            Return False
        Else
            Dim Result As Boolean = False
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                    'Dim entry As ZipArchiveEntry = archive.GetEntry(EntryName)
                    'entry.
                    For Each entry As ZipArchiveEntry In archive.Entries
                        If entry.FullName = EntryName Then
                            Result = True
                            Exit For
                        End If
                    Next
                    Return Result
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error checking if entry exists: " & vbCrLf & ex.Message & vbCrLf)
            End Try
        End If
    End Function

    'NOTE: The Entry creation date is not stored in a zip file!
    'Public Function EntryCreationDate(ByVal EntryName As String) As Date
    '    'Returns the Creation Date of the entry if it exists.
    '    If ArchivePath = "" Then
    '        'Return False
    '        Return DateValue("1900-01-01")
    '    Else
    '        Dim Result As Boolean = False
    '        Try
    '            Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
    '                Dim CreationDate As Date
    '                For Each entry As ZipArchiveEntry In archive.Entries
    '                    If entry.FullName = EntryName Then
    '                        Result = True
    '                        'CreationDate = entry.
    '                        Exit For
    '                    End If
    '                Next
    '                'Return Result
    '            End Using
    '        Catch ex As Exception
    '            RaiseEvent ErrorMessage("Error checking if entry exists: " & vbCrLf & ex.Message & vbCrLf)
    '        End Try
    '    End If
    'End Function

    Public Function EntryLastEditDate(ByVal EntryName As String) As Date
        'Returns the Creation Date of the entry if it exists.
        If ArchivePath = "" Then
            'Return False
            Return DateValue("1900-01-01")
        Else
            Dim Result As Boolean = False
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                    'Dim LastEditDate As Date
                    Dim LastEditDateTimeOffset As DateTimeOffset
                    For Each entry As ZipArchiveEntry In archive.Entries
                        If entry.FullName = EntryName Then
                            Result = True
                            LastEditDateTimeOffset = entry.LastWriteTime
                            Exit For
                        End If
                    Next
                    'Return Result
                    If Result = True Then 'Entry found
                        Return LastEditDateTimeOffset.DateTime
                    Else 'Entry not found
                        Return DateValue("1900-01-01")
                    End If
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error checking if entry exists: " & vbCrLf & ex.Message & vbCrLf)
                Return DateValue("1900-01-01")
            End Try
        End If
    End Function

    Public Sub RemoveEntry(ByVal EntryName As String)
        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank." & vbCrLf)
        Else
            Try
                ''Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                'Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Update)
                '    For Each entry As ZipArchiveEntry In archive.Entries
                '        If entry.FullName = EntryName Then
                '            entry.Delete()
                '            'Exit For
                '            'Continue - Entry name may be repeated!
                '        End If
                '    Next
                'End Using

                'NOTE: Cannot update the collection (of Entries) while in the For Each loop! (Produces an error!)
                '      Use a For loop instead:
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Update)
                    Dim NEntries As Integer = archive.Entries.Count
                    Dim I As Integer
                    'For Each entry As ZipArchiveEntry In archive.Entries
                    For I = 0 To NEntries - 1
                        'If entry.FullName = EntryName Then
                        If archive.Entries(I).FullName = EntryName Then
                            'entry.Delete()
                            archive.Entries(I).Delete()
                            'Exit For
                            'Continue - Entry name may be repeated!
                        End If
                    Next
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error checking if entry exists: " & vbCrLf & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub AddEntry(ByVal EntryName As String)
        'Create an entry name without data.
        'The can be used to create lock files to indicate an archive is in use.
        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank." & vbCrLf)
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Update)
                    Dim newEntry As ZipArchiveEntry = archive.CreateEntry(EntryName)
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error creating new entry named: " & EntryName & ". Error message: " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RenameEntry(ByVal Oldname As String, ByVal NewName As String)
        'Rename an entry.
        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank." & vbCrLf)
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Update)
                    Dim OldEntry As ZipArchiveEntry = archive.GetEntry(Oldname)
                    If IsNothing(OldEntry) Then
                        RaiseEvent ErrorMessage("Entry name not found in the archive: " & Oldname & vbCrLf)
                    Else
                        'OldEntry found. Create entry with the new name.
                        'First delete any existing entries with the new name:
                        Do
                            Dim FindEntry As ZipArchiveEntry = archive.GetEntry(NewName)
                            If IsNothing(FindEntry) Then
                                Exit Do
                            Else
                                FindEntry.Delete()
                            End If
                        Loop

                        'Add the new entry:
                        Dim newEntry As ZipArchiveEntry = archive.CreateEntry(NewName)

                        'Copy the old entry to the new entry:
                        Dim oldStream As IO.Stream = OldEntry.Open()
                        Dim newStream As IO.Stream = newEntry.Open
                        oldStream.CopyTo(newStream)

                        'Delete the old entry:
                        OldEntry.Delete()
                    End If
                End Using
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Function NEntries() As Integer
        'Return the number of entries in the archive.
        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank." & vbCrLf)
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                    Return archive.Entries.Count
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error getting number of entries in the archive." & " Error message: " & ex.Message & vbCrLf)
            End Try
        End If
    End Function

    Public Function GetEntryName(ByVal EntryNumber As Integer) As String
        'Return the name of entry number EntryNumber.
        If ArchivePath = "" Then
            RaiseEvent ErrorMessage("Archive path is blank." & vbCrLf)
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                    If EntryNumber < archive.Entries.Count Then
                        Return archive.Entries(EntryNumber).Name
                    Else
                        RaiseEvent ErrorMessage("EntryNumber too large.")
                    End If

                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("Error getting name of entry in the archive." & " Error message: " & ex.Message & vbCrLf)
            End Try
        End If
    End Function

    Public Sub SelectFile()
        'Show the ZipSelectFile form.

        SelectedFile = ""

        If IsNothing(SelectFileForm) Then
            SelectFileForm = New frmZipSelectFile
            SelectFileForm.ZipArchivePath = ArchivePath
        Else
            SelectFileForm.Show()
            SelectFileForm.BringToFront()
        End If
    End Sub

    Public Function SelectFileModal(ByVal FileExtension As String) As String
        'Show the frmZipSelectFile form as modal and return the selected file.

        'Dim SelectFileForm As New frmZipSelectFile
        Dim SelectFileForm As New frmZipSelectFileModal
        SelectFileForm.ZipArchivePath = ArchivePath
        SelectFileForm.FileExtension = FileExtension
        SelectFileForm.GetFileList()

        If SelectFileForm.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Return SelectFileForm.FileName
        Else
            Return ""
        End If

    End Function


    Private Sub SelectFileForm_FileSelected(FileName As String) Handles SelectFileForm.FileSelected
        SelectedFile = FileName
        RaiseEvent FileSelected(FileName)
    End Sub

    Private Sub SelectFileForm_ErrorMessage(Message As String) Handles SelectFileForm.ErrorMessage
        RaiseEvent ErrorMessage(Message)
    End Sub

    Private Sub SelectFileForm_Message(Message As String) Handles SelectFileForm.Message
        RaiseEvent Message(Message)
    End Sub

    Public Sub GetEntryList(ByRef FileNames As ArrayList, ByVal FileExtension As String)
        'Return a list of file names with the specified file extension.

        Dim ValidExtension As String = FileExtension
        If ValidExtension.StartsWith(".") Then
        Else
            ValidExtension = "." & ValidExtension
        End If

        If ArchivePath = "" Then
            'No archive specified.
            RaiseEvent ErrorMessage("No archive path specified." & vbCrLf)
        Else
            Try
                Using archive As ZipArchive = ZipFile.Open(ArchivePath, ZipArchiveMode.Read)
                    For Each entry As ZipArchiveEntry In archive.Entries
                        If entry.FullName.EndsWith(ValidExtension) Then
                            FileNames.Add(entry.FullName)
                        End If
                    Next
                End Using
            Catch ex As Exception
                RaiseEvent ErrorMessage("GetEntryList error: " & ex.Message & vbCrLf)
            End Try
        End If

    End Sub


#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Event Message(ByVal Message As String) 'Send a message.
    Public Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Public Event FileSelected(ByVal FileName As String) 'Send the name of the selected file.


#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'ZipComp ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



'The XSequence class is used to run XSequence files.

'An XSequence file is an AL-H7 (TM) Information Vector Sequence stored in an XML format.
'AL-H7 (TM) is the name of a programming system that uses sequences of information and location value pairs to store data items or processing steps.
'A single information and location value pair is called a knowledge element (or noxel).
'Any computer program, mathematical expression or data set can be expressed as an Information Vector Sequence.

Public Class XSequence '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Runs an XML Sequence file.

#Region " Variable declarations"

    Public NewParameter As New ParamNameDescVal 'Store the Name, Description and Value of a new parameter.

    'Public Parameter As New Dictionary(Of String, String) 'The Parameter dictionary.
    Public Parameter As New Dictionary(Of String, ParamDescVal) 'The Parameter dictionary.
    'Add new parameter example: Prameter.Add("InputQuery", "Select * From My_Data_Table")
    'Get a parameter value: InputDataQuery = Parameter("InputQuery").Value

    'Structure used  for recording and running a processing sequence:
    Structure strucSequenceStatus
        Dim Recording As Boolean
        Dim Running As Boolean
        Dim Interactive As Boolean
    End Structure

    'Structure used to hold Information Vector values
    Structure strucPropSeq
        Dim Path As String
        Dim Value As String
        Dim NextItem As Integer
    End Structure

    'Structure used to hold Loop information
    Structure strucLoopInfo
        Dim Name As String        'The name of the loop
        Dim StartIndex As Integer 'The Sequence() index number of the start of the loop
        Dim EndIndex As Integer   'The Sequence() index number of the end of the loop
    End Structure

    'Dim Sequence() As strucPropSeq 'Now using Information Vector Sequence()...
    'Dim SequenceCount As Integer 'The number of entries in Seqeunce()

    'Dim LoopInfo() As strucLoopInfo 'Stores information about loops. Used for processing _PropSeq()
    Dim LoopOrGroupInfo() As strucLoopInfo 'Stores information about loops or groups. Used for processing _PropSeq()
    'Dim LoopInfoCount As Integer 'The number of entries stored in LoopInfo()
    Dim LoopOrGroupInfoCount As Integer 'The number of entries stored in LoopOrGroupInfo()

    'Private _sequenceStatus As String 'Valid status strings: "No_more_input_files", "At_end_of_file"

    'Variables used to store and process a Information Vector Sequence:

    Dim LoopInfo() As strucLoopInfo 'Stores information about loops. Used for processing _PropSeq()
    Dim LoopInfoCount As Integer 'The number of entries stored in LoopInfo()

    'Public SequenceStatus As String
    'Public SequenceStatus As Collection
    'This collection holds sequence status strings.
    'Examples: "No_more_input_files" indicates that there are no more text files left to process.
    '          "At_end_of_file" indicates that the end of the text file has been reached
    'The collection of status string is now passed from the calling program!


#End Region

#Region " Properties"

    'List of Properties:
    'SequenceName               The name of a processing sequence.
    'SequenceDescription        A description of the processing sequence.
    'Sequence()                 Array of processing sequence instructions.
    'SequenceStatus             A string indicating the status of the processing sequence. Strings include: "No_more_input_files" "At_end_of_file"

    Private mSequenceName As String = ""
    Property SequenceName() As String
        Get
            Return mSequenceName
        End Get
        Set(ByVal value As String)
            mSequenceName = value
        End Set
    End Property

    Private mSequenceDescription As String = ""
    Property SequenceDescription() As String
        Get
            Return mSequenceDescription
        End Get
        Set(ByVal value As String)
            mSequenceDescription = value
        End Set
    End Property

    Private mSequence() As strucPropSeq 'Information Vector Sequence (Property Sequence) elements: Path, Value, NextItem
    Private mSequenceCount As Integer = 0 'The number of entries in _seqeunce() 'Added initialsation to zero on 4Jan14
    Property Sequence(ByVal Index As Integer) As strucPropSeq
        Get
            If Index > mSequenceCount - 1 Then 'Indes points past the last entry in _PropSeq()
                'Invalid index number!
            Else
                Return mSequence(Index)
            End If
        End Get
        Set(ByVal value As strucPropSeq)
            If Index > mSequenceCount - 1 Then
                If Index = mSequenceCount Then
                    'Append an element to the sequence:
                    Dim ArrayLen As Integer
                    ArrayLen = mSequence.GetLength(0)
                    If ArrayLen < Index Then 'Need to increase the size of the array:
                        ReDim Preserve mSequence(0 To ArrayLen + 128)
                    End If
                    mSequenceCount = mSequenceCount + 1
                    mSequence(Index) = value
                Else
                    'Cannot add a new element here!
                End If
            Else
                mSequence(Index) = value
            End If
        End Set
    End Property


#End Region

#Region " Methods"

    'List of methods:


    'RunXSequence(XSeq)     Runs an XML Information Vector Sequence file.

    'ExtractPropSeq(XSeq)   Extracts the Information Vector sequence from an XML Sequence file. The sequence is stored in the Sequence() array.
    '  ClearPropSeq           Clears the Sequence() array. The Sequence array contains Path, Value and NextItem fields.
    '  AppendPropSeqItem      Adds a new property setting to the Sequence() array.
    '  ScanChildNodes         Scans the child nodes of the XML Sequence file

    'ProcessPropSeqNextItems    Calculates the NextItem fields in the Infromation Vector Sequence array Sequence().

    'RunPropSeq                 Runs the Information Vector sequence (Property Sequence) statements contained in the Sequence() array.
    '  PreScanInstruction       Initial processing of a PInformation Vector Sequence instruction.
    '  ProcessLoops             Checks for loops in Information Vector Sequence instructions.

    'Public Sub RunXSequence(ByRef XSeq As System.Xml.XmlDocument, ByRef SequenceStatus As String)
    'Public Sub RunXSequence(ByRef XSeq As System.Xml.XmlDocument, ByRef Status As Collection)
    Public Sub RunXSequence(ByRef XSeq As System.Xml.XmlDocument, ByRef Status As System.Collections.Specialized.StringCollection)
        'Debug.Print("Running RunXSequence(XSeq, Status)")
        'Debug.Print("Starting ExtractPropSeq(XSeq)")
        ExtractPropSeq(XSeq)
        'Debug.Print("Starting ProcessPropSeqNextItems()")
        ProcessPropSeqNextItems()
        'Debug.Print("RunPropSeq(Status)")
        RunPropSeq(Status)
        'Debug.Print("Finished RunXSequence")
    End Sub


    Private Sub ClearPropSeq()
        'Clears the Information Vector Sequence (Property Sequence) array:

        'ReDim Sequence(0 To 0)
        'SequenceCount = 0

        ReDim mSequence(0 To 0)
        mSequenceCount = 0
    End Sub

    Private Sub AppendPropSeqItem(ByVal Path As String, ByVal Value As String, ByVal NextItem As Integer)
        'Appends a Property Sequence item to the end of the Sequence() array.

        'Property Sequence items hare stored in the Sequence() array
        'SequenceCount stores number of Property Sequence items
        'Sequence() is zero based so the last item is stored at index number PropSeqCount - 1

        'Reize PropSeq() if required:
        Dim ArrayLen As Integer
        'ArrayLen = Sequence.GetLength(0)
        ArrayLen = mSequence.GetLength(0)
        If ArrayLen < mSequenceCount + 1 Then
            'ReDim Preserve Sequence(0 To ArrayLen + 128)
            ReDim Preserve mSequence(0 To ArrayLen + 128) 'Expand mSeqeunce by 128 elements at a time - should be quicker than resizing every time an item is added.
        End If

        'Append the new item to the array:
        mSequenceCount = mSequenceCount + 1
        mSequence(mSequenceCount - 1).Path = Path
        mSequence(mSequenceCount - 1).Value = Value
        mSequence(mSequenceCount - 1).NextItem = NextItem

    End Sub

    Private Sub ExtractPropSeq(ByRef XSeq As System.Xml.XmlDocument)
        'Extracts the property sequence from an XML Sequence file

        Try
            ClearPropSeq() 'Clear the Property Sequence array.

            Dim node As System.Xml.XmlNode
            node = XSeq.DocumentElement

            Dim CurrentPath As String = ""
            Dim OldPath As String
            Dim NodeNo As Integer
            Dim NodeVal As String

            Dim I As Integer 'Loop index

            CurrentPath = node.Name
            NodeNo = 1
            If node.ChildNodes.Count = 0 Then
                NodeVal = node.Value
                AppendPropSeqItem(CurrentPath, NodeVal, 0)
            Else
                For I = 1 To node.ChildNodes.Count
                    If node.ChildNodes(I - 1).NodeType = System.Xml.XmlNodeType.Element Then
                        OldPath = CurrentPath
                        CurrentPath = CurrentPath & ":" & node.ChildNodes(I - 1).Name
                        NodeNo = NodeNo + 1
                        ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                        CurrentPath = OldPath
                    Else
                        NodeNo = NodeNo + 1
                        ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                    End If
                Next
            End If

            'Add End_Of_Sequence item:
            AppendPropSeqItem("End_Of_Sequence", "", 0)

        Catch ex As Exception
            RaiseEvent ErrorMsg(ex.Message & vbCrLf)
        End Try
        'For Debugging:
        'Debug.Print("")
        'Debug.Print("ExtractPropSeq(XSeq)")
        'Debug.Print("Contents of the Property Sequence array: -------------------------------------------------------")
        'Debug.Print("Path       Value         Next Item")
        'For I = 0 To mSequenceCount - 1
        '    Debug.Print(mSequence(I).Path & " " & mSequence(I).Value & " " & mSequence(I).NextItem)
        'Next
        'Debug.Print("End of the Property Sequence -------------------------------------------------------------------")
        'Debug.Print("")

    End Sub

    Private Sub ScanChildNodes(ByRef node As System.Xml.XmlNode, ByRef NodeNo As Integer, ByRef CurrentPath As String)
        'Scan the child nodes.
        'This is called by the ExtractPropSeq subroutine.
        'This subroutine calls itself (recursion) as it scans the nodes in an XML file.

        Dim I As Integer 'Loop index
        Dim OldPath As String
        Dim NodeVal As String
        Dim LoopName As String
        LoopName = ""
        Dim GroupName As String
        GroupName = ""

        If node.ChildNodes.Count = 0 Then
            NodeVal = node.Value
            If node.NodeType <> System.Xml.XmlNodeType.Comment Then
                AppendPropSeqItem(CurrentPath, NodeVal, 0)
            End If
        Else
            'Scan the child nodes:
            If node.Name = "Loop" Then
                If node.Attributes.Count > 0 Then
                    LoopName = node.Attributes(0).Value.ToString
                Else
                    LoopName = ""
                End If
                AppendPropSeqItem("Start_Loop", LoopName, 0)
            End If

            If node.Name = "Group" Then
                If node.Attributes.Count > 0 Then
                    GroupName = node.Attributes(0).Value.ToString
                Else
                    GroupName = ""
                End If
                AppendPropSeqItem("Start_Group", GroupName, 0)
            End If

            For I = 1 To node.ChildNodes.Count
                If node.ChildNodes(I - 1).NodeType = System.Xml.XmlNodeType.Element Then
                    OldPath = CurrentPath
                    CurrentPath = CurrentPath & ":" & node.ChildNodes(I - 1).Name
                    NodeNo = NodeNo + 1
                    ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                    CurrentPath = OldPath
                Else
                    NodeNo = NodeNo + 1
                    ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                End If
            Next
            'Finished scanning the child nodes.
            'In  a Signalworks Sequence file, we need to record details of nodes named "Loop" and "Group".
            If node.Name = "Loop" Then
                AppendPropSeqItem("End_Loop", LoopName, 0)
            End If
            If node.Name = "Group" Then
                AppendPropSeqItem("End_Group", GroupName, 0)
            End If
        End If
    End Sub

    Private Sub ProcessPropSeqNextItems()
        'Calculates the Next Item fields in the Property Sequence array Sequence().

        'Set the NextItem fields for each item in Sequence()
        'End_Loop items shound point back to the corresponding Start_Loop
        'Exit_Loop_If items sould point past the corresponding End_Loop
        'All other items should point to the next item.

        ReDim LoopOrGroupInfo(0 To 64)
        LoopOrGroupInfoCount = 0
        Dim I As Integer 'Loop index
        Dim J As Integer 'Loop index
        Dim LoopOrGroupNumber As Integer 'The sequential number of the current loop or group.
        LoopOrGroupNumber = 0
        Dim LoopOrGroupLevel As Integer 'The nesting level of the current loop or group.
        LoopOrGroupLevel = 0
        Dim LoopOrGroupName As String
        Dim ArraySize As Integer
        LoopOrGroupName = ""

        'Scan the Sequence() array and extract Loop and group information:
        For I = 0 To mSequenceCount - 1
            If mSequence(I).Path = "Start_Loop" Then
                LoopOrGroupNumber = LoopOrGroupNumber + 1
                LoopOrGroupLevel = LoopOrGroupLevel + 1
                LoopOrGroupName = mSequence(I).Value 'The name of the loop
                'Increase the size of the LoopOrGroupInfo() array if required:
                ArraySize = LoopOrGroupInfo.GetLength(0)
                If LoopOrGroupInfoCount > ArraySize Then
                    ReDim Preserve LoopOrGroupInfo(0 To ArraySize + 64) 'Resize by 64 elements.
                End If
                LoopOrGroupInfo(LoopOrGroupNumber - 1).StartIndex = I 'The Sequence() index number of the start of the loop
                LoopOrGroupInfo(LoopOrGroupNumber - 1).Name = mSequence(I).Value 'The name of the loop
                LoopOrGroupInfoCount = LoopOrGroupInfoCount + 1

            ElseIf Sequence(I).Path = "End_Loop" Then
                LoopOrGroupLevel = LoopOrGroupLevel - 1
                LoopOrGroupName = mSequence(I).Value
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        LoopOrGroupInfo(J).EndIndex = I 'The Sequence() index number of the end of the loop
                        Exit For
                    End If
                Next
            ElseIf Sequence(I).Path = "Start_Group" Then
                LoopOrGroupNumber = LoopOrGroupNumber + 1
                LoopOrGroupLevel = LoopOrGroupLevel + 1
                LoopOrGroupName = Sequence(I).Value
                'Increase the size of the LoopInfo() array if required:
                ArraySize = LoopOrGroupInfo.GetLength(0)
                If LoopOrGroupInfoCount > ArraySize Then
                    ReDim Preserve LoopOrGroupInfo(0 To ArraySize + 64) 'Resize by 64 elements.
                End If
                LoopOrGroupInfo(LoopOrGroupNumber - 1).StartIndex = I 'The Sequence() index number of the start of the group
                LoopOrGroupInfo(LoopOrGroupNumber - 1).Name = mSequence(I).Value 'The name of the group
                LoopOrGroupInfoCount = LoopOrGroupInfoCount + 1
            ElseIf mSequence(I).Path = "End_Group" Then
                LoopOrGroupLevel = LoopOrGroupLevel - 1
                LoopOrGroupName = mSequence(I).Value
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        LoopOrGroupInfo(J).EndIndex = I 'The Sequence() index number of the end of the group
                        Exit For
                    End If
                Next
            Else

            End If
        Next

        'Scan the Sequence() array and update the NextItems field:
        For I = 0 To mSequenceCount - 1
            If mSequence(I).Path = "Start_Loop" Then
                mSequence(I).NextItem = I + 1 'The index number of the next item
                LoopOrGroupName = Sequence(I).Value
            ElseIf mSequence(I).Path = "End_Loop" Then
                LoopOrGroupName = Sequence(I).Value
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        mSequence(I).NextItem = LoopOrGroupInfo(J).StartIndex 'The index number of the next item
                        Exit For
                    End If
                Next
                'ElseIf Sequence(I).Path.Substring(Sequence(I).Path.Length - 10, 10) = "ExitLoopIf" Then
            ElseIf mSequence(I).Path.EndsWith("ExitLoopIf") Then
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        mSequence(I).NextItem = LoopOrGroupInfo(J).EndIndex + 1
                        Exit For
                    End If
                Next
            ElseIf mSequence(I).Path = "Start_Group" Then
                mSequence(I).NextItem = I + 1 'The index number of the next item
                LoopOrGroupName = Sequence(I).Value
            ElseIf mSequence(I).Path = "End_Group" Then
                LoopOrGroupName = mSequence(I).Value
                mSequence(I).NextItem = I + 1
            ElseIf mSequence(I).Path.EndsWith("ExitGroupIf") Then
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        mSequence(I).NextItem = LoopOrGroupInfo(J).EndIndex + 1
                        Exit For
                    End If
                Next
            Else
                mSequence(I).NextItem = I + 1
            End If
        Next

        'For Debugging:
        'Debug.Print("")
        'Debug.Print("ProcessPropSeqNextItems()")
        'Debug.Print("Contents of the Property Sequence array: -------------------------------------------------------")
        'Debug.Print("Index     Path       Value         Next Item")
        'For I = 0 To mSequenceCount - 1
        '    Debug.Print(I & " " & mSequence(I).Path & " " & mSequence(I).Value & " " & mSequence(I).NextItem)
        'Next
        'Debug.Print("End of the Property Sequence -------------------------------------------------------------------")
        'Debug.Print("")

    End Sub

    'Sub RunPropSeq(ByRef SequenceStatus As String)
    'Sub RunPropSeq(ByRef Status As Collection)
    Sub RunPropSeq(ByRef Status As System.Collections.Specialized.StringCollection)
        'Runs the property sequence statements contained in the Sequence() array.
        'Status is a collection passed by the program calling this class.
        '  It can contain 0, 1 or more strings indicating the status of the processing such as "No_more_input_files" or "At_end_of_file".

        Dim LineNo As Integer = 0
        Dim Path As String = ""
        Dim Value As String = ""
        Dim NextItem As Integer = 0

        'SequenceStatus variable may also be modified in the code that calls this class!
        '  It will be declared in that code and passes by reference to this class.
        'Dim SequenceStatus As String = "" 'Valid status strings: "No_more_input_files", "At_end_of_file"

        'Debug.Print("Starting RunPropSeq(Status)")

        'Debug.Print("Starting: While Path <> End_Of_Sequence")

        While Path <> "End_Of_Sequence"



            Path = mSequence(LineNo).Path
            'Debug.Print("Path = " & Path)

            Value = mSequence(LineNo).Value
            'Debug.Print("Value = " & Value)

            NextItem = mSequence(LineNo).NextItem
            If Path <> "End_Of_Sequence" Then
                If Path = "Start_Loop" Then
                    LineNo = NextItem
                ElseIf Path = "End_Loop" Then
                    LineNo = NextItem 'The NextItem for an End_Loop statement is the Start_Loop line.
                ElseIf Path.EndsWith("ExitLoopIf") Then
                    If IsNothing(Status) Then 'Continue to next line
                        LineNo = LineNo + 1
                    ElseIf Status.Contains(Value) Then 'Exit loop
                        LineNo = NextItem 'The NextItem for an ExitLoopIf statement in the exit line.
                    Else 'Continue to the next line.
                        LineNo = LineNo + 1
                    End If
                    'If Value = Status Then 'Exit loop
                    '    LineNo = NextItem 'The NextItem for an ExitLoopIf statement is the exit line.
                    'Else
                    '    LineNo = LineNo + 1 'Continue to next line.
                    'End If
                    'Added May 5th:
                ElseIf Path.EndsWith("ExitGroupIf") Then
                    If IsNothing(Status) Then 'Continue to next line
                        LineNo = LineNo + 1
                    ElseIf Status.Contains(Value) Then 'Exit group
                        LineNo = NextItem
                    Else
                        LineNo = LineNo + 1
                    End If
                    'If Value = SequenceStatus Then 'Exit group
                    '    LineNo = NextItem
                    'Else
                    '    LineNo = LineNo + 1
                    'End If
                Else
                    PreScanInstruction(Path, Value)
                    LineNo = NextItem
                End If
            End If
        End While

        'RaiseEvent Instruction("EndOfSequence", "")
        RaiseEvent Instruction("", "EndOfSequence")


        'At end of sequence.
        'SequenceStatus = ""
        If IsNothing(Status) Then
        Else
            Status.Clear()
        End If


    End Sub

    Sub DisplaySequenceArray(ByRef XSeq As System.Xml.XmlDocument)
        'Display the contents of the Sequence() array for debugging:

        ExtractPropSeq(XSeq)
        ProcessPropSeqNextItems()

        Dim I As Integer 'Loop index

        'Debug.Print("DisplaySequenceArray(XSeq)")
        'Debug.Print("Contents of the Sequence() array:")
        'Debug.WriteLine(String.Format("{0, -8}{1,-96}{2,-64}{3,-8}", "Item", "Path", "Value", "Next Item"))

        'For I = 1 To mSequenceCount
        '    Debug.WriteLine(String.Format("{0,-8}{1,-96}{2,-64}{3,-8}", I - 1, mSequence(I - 1).Path, mSequence(I - 1).Value, mSequence(I - 1).NextItem))
        'Next

    End Sub

    Private Sub PreScanInstruction(ByVal path As String, ByVal prop As String)
        'Applies a Processing Sequence Instruction - Initial Processing
        'An instruction consists of a path and a property value
        'The path points to an item on a form
        'The item is set to the property value (prop).

        'Valid paths have the form:
        '    ProcessingSequence:<CreationDate>
        '    ProcessingSequence:<Description>
        '    ProcessingSequence:Process:<Path>
        '    Start_Loop
        '    End_Loop
        '    Start_Group
        '    End_Group
        '    End_Of_Sequence
        'Any other path returns an "Unknown Instruction" warning.

        Dim Path2 As String = "" 'The second part of the Instruction path
        Dim Path3 As String = "" 'The third part of the instruction path

        If path.StartsWith("ProcessingSequence:") Then
            Path2 = path.Substring(19, path.Length - 19) 'Get path string with ProcessingSequence: removed from the start.
            If Path2.StartsWith("Process:") Then
                Path3 = Path2.Substring(8, Path2.Length - 8) 'Get path string with Process: removed from the start.
                'Instruction has the form ProcessingSequence:Process:<Path3>
                ProcessLoops(Path3, prop)
            Else
                'Instruction has the form ProcessingSequence:<Path2>
                Select Case Path2
                    Case "CreationDate"
                        'The creation date of the instruction sequence.
                    Case "Description"
                        'A description of the instruction sequence.
                End Select
            End If
        Else
            'Instruction may be Start_Loop or End_Loop or End_Of_Sequence or <unknown>
            Select Case path
                Case "Start_Loop"
                    'At the start of a loop. Continue processing.
                Case "End_Loop"
                    'Point back to the start of the loop.

                Case "Start_Group"
                    'At the start of a group. Continue processing"
                Case "End_Group"


                Case "End_Of_Sequence"
                    'Stop processing!

                Case Else
                    'Unknown instruction!
                    'RaiseEvent ErrorMsg("XSequence warning: Unknown instruction: " & path & "  Property value: " & prop & vbCrLf)

                    'Send as an instruction: (This works OK if the XML file is not a Processing Sequence - the elements are just sent in sequence.)
                    RaiseEvent Instruction(prop, path)

            End Select
        End If
    End Sub

    Private Sub ProcessLoops(ByVal path As String, ByVal prop As String)
        'Applies a Processing Sequence Instruction - Check for Loop: in the path
        'An instruction consists of a path and a property value.
        'The path either points to an item on a form or a property or method in theMain form.
        'The item is set to the property value.

        'Some statements have a multi level path:
        'ProcessingSequence:Process:DataPaths:TextFilesToProcess:TextFile  Value: C:\Users\Robert\Documents\Shares\asxdata\20090102.txt 

        'Some statements have a single level path:
        'ProcessingSequence:Process:Loop:ReadTextCommand  Value: ReadNextLine

        'Consider using single level path statements to simplify the instruction sequence code.

        Dim PathN As String = ""

        'Debug.Print("In method ProcessLoops(). Path = " & path)

        If path.StartsWith("Loop:") Then
            'Debug.Print("In method ProcessLoops. Path starts with Loop:. Path = " & path)
            'Instruction has the form Loop:<PathN>
            PathN = path.Substring(5, path.Length - 5)
            ProcessLoops(PathN, prop) 'Call the subroutine recursively until all Loop: and Group: components of the path are removed.
        ElseIf path.StartsWith("Group:") Then
            'Instruction has the form Group:<PathN>
            PathN = path.Substring(6, path.Length - 6)
            ProcessLoops(PathN, prop) 'Call the subroutine recursively until all Loop: and Group: components of the path are removed.
        Else
            'Instruction has the form <Path> <Value>
            'Apply the instruction:
            'RaiseEvent Instruction(path, prop)
            RaiseEvent Instruction(prop, path)
        End If
    End Sub

    Public Sub AddParameter()
        'Add a new parameter entry to Parameter.

        Dim ParamData As New ParamDescVal

        ParamData.Value = NewParameter.Value
        ParamData.Description = NewParameter.Description
        Parameter.Add(NewParameter.Name, ParamData)

        'Reset NewParameter for next Parameter
        NewParameter.Name = ""
        NewParameter.Description = ""
        NewParameter.Value = ""
    End Sub


#Region " Sequence Status methods"

    'Adding Sequence Sttus strings:
    'SequenceStatus.Add("No_more_input_files")
    'SequenceStatus.Add("At_end_of_file")

    'Public Sub SequenceStatusRemove(ByVal Status As String)
    '    'Remove the specified Status string from the SequenceStatus collection if it is present:

    '    If IsNothing(SequenceStatus) Then
    '        'SequenceStatus contains no Status strings
    '    Else
    '        SequenceStatus.Remove(Status)
    '    End If
    'End Sub

    'Public Function SequenceStatusContains(ByVal Status As String) As Boolean
    '    'Check if the SequenceStatus collection contains the specified Status string:

    '    If IsNothing(SequenceStatus) Then
    '        Return False
    '    Else
    '        If SequenceStatus.Contains(Status) Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    End If
    'End Function

#End Region


#End Region

#Region " Events"

    Public Event ErrorMsg(ByVal ErrMsg As String) 'Send an error message.

    'Public Event Instruction(ByVal Path As String, ByVal Prop As String) 'Send an instruction to be executed.
    'Public Event Instruction(ByVal Locn As String, ByVal Info As String) 'Send an instruction to be executed.
    Public Event Instruction(ByVal Info As String, ByVal Locn As String) 'Send an instruction to be executed.

#End Region

#Region " Classes "


#End Region

End Class 'XSequence -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Class ParamNameDescVal
    'ParamInfo is used to store the Name, Description and Value of a parameter.

    Private _name As String = "" 'The name of a parameter.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A description of the parameter
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _value As String = "" 'The value of a parameter.
    Property Value As String
        Get
            Return _value
        End Get
        Set(value As String)
            _value = value
        End Set
    End Property

End Class

Public Class ParamDescVal
    'ParamInfo is used to store the Description and Value of a parameter.

    Private _description As String = "" 'A description of the parameter
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _value As String = "" 'The value of a parameter.
    Property Value As String
        Get
            Return _value
        End Get
        Set(value As String)
            _value = value
        End Set
    End Property

End Class

'The XMessage class is used to read an XML Message (XMessage).
'An XMessage is a simplified XSequence. It is used to exchange information between Andorville (TM) applications.
Public Class XMessage '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Runs an XML Message.

#Region " Variable declarations"

    'Structure used  for recording and running a processing sequence:
    Structure strucSequenceStatus
        Dim Recording As Boolean
        Dim Running As Boolean
        Dim Interactive As Boolean
    End Structure

    'Structure used to hold Property Sequence values
    Structure strucPropSeq
        Dim Path As String
        Dim Value As String
        Dim NextItem As Integer
    End Structure

    'Structure used to hold Loop information
    Structure strucLoopInfo
        Dim Name As String        'The name of the loop
        Dim StartIndex As Integer 'The Sequence() index number of the start of the loop
        Dim EndIndex As Integer   'The Sequence() index number of the end of the loop
    End Structure

    Dim LoopOrGroupInfo() As strucLoopInfo 'Stores information about loops or groups. Used for processing _PropSeq()
    Dim LoopOrGroupInfoCount As Integer 'The number of entries stored in LoopOrGroupInfo()

    'Private _sequenceStatus As String 'Valid status strings: "No_more_input_files", "At_end_of_file"

    'Variables used to store and process a Property Sequence:
    Dim LoopInfo() As strucLoopInfo 'Stores information about loops. Used for processing _PropSeq()
    Dim LoopInfoCount As Integer 'The number of entries stored in LoopInfo()

    'Public SequenceStatus As String
    'Public SequenceStatus As Collection
    'This collection holds sequence status strings.
    'Examples: "No_more_input_files" indicates that there are no more text files left to process.
    '          "At_end_of_file" indicates that the end of the text file has been reached
    'The collection of status string is now passed from the calling program!

#End Region 'Variable declarations

#Region " Properties"

    'List of Properties:
    'SequenceName               The name of a processing sequence.
    'SequenceDescription        A description of the processing sequence.
    'Sequence()                 Array of processing sequence instructions.
    'SequenceStatus             A string indicating the status of the processing sequence. Strings include: "No_more_input_files" "At_end_of_file"

    Private mSequenceName As String = ""
    Property SequenceName() As String
        Get
            Return mSequenceName
        End Get
        Set(ByVal value As String)
            mSequenceName = value
        End Set
    End Property

    Private mSequenceDescription As String = ""
    Property SequenceDescription() As String
        Get
            Return mSequenceDescription
        End Get
        Set(ByVal value As String)
            mSequenceDescription = value
        End Set
    End Property

    Private mSequence() As strucPropSeq 'Property Sequence elements: Path, Value, NextItem
    Private mSequenceCount As Integer = 0 'The number of entries in _seqeunce() 'Added initialsation to zero on 4Jan14
    Property Sequence(ByVal Index As Integer) As strucPropSeq
        Get
            If Index > mSequenceCount - 1 Then 'Indes points past the last entry in _PropSeq()
                'Invalid index number!
            Else
                Return mSequence(Index)
            End If
        End Get
        Set(ByVal value As strucPropSeq)
            If Index > mSequenceCount - 1 Then
                If Index = mSequenceCount Then
                    'Append an element to the sequence:
                    Dim ArrayLen As Integer
                    ArrayLen = mSequence.GetLength(0)
                    If ArrayLen < Index Then 'Need to increase the size of the array:
                        ReDim Preserve mSequence(0 To ArrayLen + 128)
                    End If
                    mSequenceCount = mSequenceCount + 1
                    mSequence(Index) = value
                Else
                    'Cannot add a new element here!
                End If
            Else
                mSequence(Index) = value
            End If
        End Set
    End Property


#End Region 'Properties

#Region " Methods"

    'List of methods:


    'Run(XSeq, Status)     Runs an XML Property Sequence file.

    'ExtractPropSeq(XSeq)   Extracts the property sequence from an XML Sequence file. The sequence is stored in the Sequence() array.
    '  ClearPropSeq           Clears the Sequence() array. The Sequence array contains Path, Value and NextItem fields.
    '  AppendPropSeqItem      Adds a new property setting to the Sequence() array.
    '  ScanChildNodes         Scans the child nodes of the XML Sequence file

    'ProcessPropSeqNextItems    Calculates the NextItem fields in the Property Sequence array Sequence().

    'RunPropSeq                 Runs the property sequence statements contained in the Sequence() array.
    '  PreScanInstruction       Initial processing of a Property Sequence instruction.
    '  ProcessLoops             Checks for loops in Property Sequence instructions.

    'Public Sub RunXSequence(ByRef XSeq As System.Xml.XmlDocument, ByRef Status As System.Collections.Specialized.StringCollection)
    Public Sub Run(ByRef XSeq As System.Xml.XmlDocument, ByRef Status As System.Collections.Specialized.StringCollection)
        ExtractPropSeq(XSeq)
        ProcessPropSeqNextItems()
        RunPropSeq(Status)
    End Sub

    Private Sub ClearPropSeq()
        'Clears the Property Sequence array:
        ReDim mSequence(0 To 0)
        mSequenceCount = 0
    End Sub

    Private Sub AppendPropSeqItem(ByVal Path As String, ByVal Value As String, ByVal NextItem As Integer)
        'Appends a Property Sequence item to the end of the Sequence() array.

        'Property Sequence items hare stored in the Sequence() array
        'SequenceCount stores number of Property Sequence items
        'Sequence() is zero based so the last item is stored at index number PropSeqCount - 1

        'Reize PropSeq() if required:
        Dim ArrayLen As Integer
        ArrayLen = mSequence.GetLength(0)
        If ArrayLen < mSequenceCount + 1 Then
            ReDim Preserve mSequence(0 To ArrayLen + 128) 'Expand mSeqeunce by 128 elements at a time - should be quicker than resizing every time an item is added.
        End If

        'Append the new item to the array:
        mSequenceCount = mSequenceCount + 1
        mSequence(mSequenceCount - 1).Path = Path
        mSequence(mSequenceCount - 1).Value = Value
        mSequence(mSequenceCount - 1).NextItem = NextItem

    End Sub

    Private Sub ExtractPropSeq(ByRef XSeq As System.Xml.XmlDocument)
        'Extracts the property sequence from an XML Sequence file

        ClearPropSeq() 'Clear the Property Sequence array.

        Dim node As System.Xml.XmlNode
        node = XSeq.DocumentElement

        Dim CurrentPath As String = ""
        Dim OldPath As String
        Dim NodeNo As Integer
        Dim NodeVal As String

        Dim I As Integer 'Loop index

        CurrentPath = node.Name
        NodeNo = 1
        If node.ChildNodes.Count = 0 Then
            NodeVal = node.Value
            AppendPropSeqItem(CurrentPath, NodeVal, 0)
        Else
            For I = 1 To node.ChildNodes.Count
                If node.ChildNodes(I - 1).NodeType = System.Xml.XmlNodeType.Element Then
                    OldPath = CurrentPath
                    CurrentPath = CurrentPath & ":" & node.ChildNodes(I - 1).Name
                    NodeNo = NodeNo + 1
                    ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                    CurrentPath = OldPath
                Else
                    NodeNo = NodeNo + 1
                    ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                End If
            Next
        End If

        'Add End_Of_Sequence item:
        AppendPropSeqItem("End_Of_Sequence", "", 0)

        'For Debugging:
        'Debug.Print("")
        'Debug.Print("ExtractPropSeq() completed")
        'Debug.Print("Contents of the Property Sequence array: -------------------------------------------------------")
        'Debug.Print("Path       Value         Next Item")
        'For I = 0 To mSequenceCount - 1
        '    Debug.Print(mSequence(I).Path & " " & mSequence(I).Value & " " & mSequence(I).NextItem)
        'Next
        'Debug.Print("End of the Property Sequence -------------------------------------------------------------------")
        'Debug.Print("")

    End Sub

    Private Sub ScanChildNodes(ByRef node As System.Xml.XmlNode, ByRef NodeNo As Integer, ByRef CurrentPath As String)
        'Scan the child nodes.
        'This is called by the ExtractPropSeq subroutine.
        'This subroutine calls itself (recursion) as it scans the nodes in an XML file.

        Dim I As Integer 'Loop index
        Dim OldPath As String
        Dim NodeVal As String
        Dim LoopName As String
        LoopName = ""
        Dim GroupName As String
        GroupName = ""

        If node.ChildNodes.Count = 0 Then
            NodeVal = node.Value
            If node.NodeType <> System.Xml.XmlNodeType.Comment Then
                AppendPropSeqItem(CurrentPath, NodeVal, 0)
            End If
        Else
            'Scan the child nodes:
            If node.Name = "Loop" Then
                If node.Attributes.Count > 0 Then
                    LoopName = node.Attributes(0).Value.ToString
                Else
                    LoopName = ""
                End If
                AppendPropSeqItem("Start_Loop", LoopName, 0)
            End If

            If node.Name = "Group" Then
                If node.Attributes.Count > 0 Then
                    GroupName = node.Attributes(0).Value.ToString
                Else
                    GroupName = ""
                End If
                AppendPropSeqItem("Start_Group", GroupName, 0)
            End If

            For I = 1 To node.ChildNodes.Count
                If node.ChildNodes(I - 1).NodeType = System.Xml.XmlNodeType.Element Then
                    OldPath = CurrentPath
                    CurrentPath = CurrentPath & ":" & node.ChildNodes(I - 1).Name
                    NodeNo = NodeNo + 1
                    ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                    CurrentPath = OldPath
                Else
                    NodeNo = NodeNo + 1
                    ScanChildNodes(node.ChildNodes(I - 1), NodeNo, CurrentPath)
                End If
            Next
            'Finished scanning the child nodes.
            'In  a Signalworks Sequence file, we need to record details of nodes named "Loop" and "Group".
            If node.Name = "Loop" Then
                AppendPropSeqItem("End_Loop", LoopName, 0)
            End If
            If node.Name = "Group" Then
                AppendPropSeqItem("End_Group", GroupName, 0)
            End If
        End If
    End Sub

    Private Sub ProcessPropSeqNextItems()
        'Calculates the Next Item fields in the Property Sequence array Sequence().

        'Set the NextItem fields for each item in Sequence()
        'End_Loop items shound point back to the corresponding Start_Loop
        'Exit_Loop_If items sould point past the corresponding End_Loop
        'All other items should point to the next item.

        ReDim LoopOrGroupInfo(0 To 64)
        LoopOrGroupInfoCount = 0
        Dim I As Integer 'Loop index
        Dim J As Integer 'Loop index
        Dim LoopOrGroupNumber As Integer 'The sequential number of the current loop or group.
        LoopOrGroupNumber = 0
        Dim LoopOrGroupLevel As Integer 'The nesting level of the current loop or group.
        LoopOrGroupLevel = 0
        Dim LoopOrGroupName As String
        Dim ArraySize As Integer
        LoopOrGroupName = ""

        'Scan the Sequence() array and extract Loop and group information:
        For I = 0 To mSequenceCount - 1
            If mSequence(I).Path = "Start_Loop" Then
                LoopOrGroupNumber = LoopOrGroupNumber + 1
                LoopOrGroupLevel = LoopOrGroupLevel + 1
                LoopOrGroupName = mSequence(I).Value 'The name of the loop
                'Increase the size of the LoopOrGroupInfo() array if required:
                ArraySize = LoopOrGroupInfo.GetLength(0)
                If LoopOrGroupInfoCount > ArraySize Then
                    ReDim Preserve LoopOrGroupInfo(0 To ArraySize + 64) 'Resize by 64 elements.
                End If
                LoopOrGroupInfo(LoopOrGroupNumber - 1).StartIndex = I 'The Sequence() index number of the start of the loop
                LoopOrGroupInfo(LoopOrGroupNumber - 1).Name = mSequence(I).Value 'The name of the loop
                LoopOrGroupInfoCount = LoopOrGroupInfoCount + 1

            ElseIf Sequence(I).Path = "End_Loop" Then
                LoopOrGroupLevel = LoopOrGroupLevel - 1
                LoopOrGroupName = mSequence(I).Value
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        LoopOrGroupInfo(J).EndIndex = I 'The Sequence() index number of the end of the loop
                        Exit For
                    End If
                Next
            ElseIf Sequence(I).Path = "Start_Group" Then
                LoopOrGroupNumber = LoopOrGroupNumber + 1
                LoopOrGroupLevel = LoopOrGroupLevel + 1
                LoopOrGroupName = Sequence(I).Value
                'Increase the size of the LoopInfo() array if required:
                ArraySize = LoopOrGroupInfo.GetLength(0)
                If LoopOrGroupInfoCount > ArraySize Then
                    ReDim Preserve LoopOrGroupInfo(0 To ArraySize + 64) 'Resize by 64 elements.
                End If
                LoopOrGroupInfo(LoopOrGroupNumber - 1).StartIndex = I 'The Sequence() index number of the start of the group
                LoopOrGroupInfo(LoopOrGroupNumber - 1).Name = mSequence(I).Value 'The name of the group
                LoopOrGroupInfoCount = LoopOrGroupInfoCount + 1
            ElseIf mSequence(I).Path = "End_Group" Then
                LoopOrGroupLevel = LoopOrGroupLevel - 1
                LoopOrGroupName = mSequence(I).Value
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        LoopOrGroupInfo(J).EndIndex = I 'The Sequence() index number of the end of the group
                        Exit For
                    End If
                Next
            Else

            End If
        Next

        'Scan the Sequence() array and update the NextItems field:
        For I = 0 To mSequenceCount - 1
            If mSequence(I).Path = "Start_Loop" Then
                mSequence(I).NextItem = I + 1 'The index number of the next item
                LoopOrGroupName = Sequence(I).Value
            ElseIf mSequence(I).Path = "End_Loop" Then
                LoopOrGroupName = Sequence(I).Value
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        mSequence(I).NextItem = LoopOrGroupInfo(J).StartIndex 'The index number of the next item
                        Exit For
                    End If
                Next
            ElseIf mSequence(I).Path.EndsWith("ExitLoopIf") Then
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        mSequence(I).NextItem = LoopOrGroupInfo(J).EndIndex + 1
                        Exit For
                    End If
                Next
            ElseIf mSequence(I).Path = "Start_Group" Then
                mSequence(I).NextItem = I + 1 'The index number of the next item
                LoopOrGroupName = Sequence(I).Value
            ElseIf mSequence(I).Path = "End_Group" Then
                LoopOrGroupName = mSequence(I).Value
                mSequence(I).NextItem = I + 1
            ElseIf mSequence(I).Path.EndsWith("ExitGroupIf") Then
                'Find the matching loop entry in LoopOrGroupInfo():
                For J = 0 To LoopOrGroupNumber - 1
                    If LoopOrGroupInfo(J).Name = LoopOrGroupName Then
                        mSequence(I).NextItem = LoopOrGroupInfo(J).EndIndex + 1
                        Exit For
                    End If
                Next
            Else
                mSequence(I).NextItem = I + 1
            End If
        Next

        'For Debugging:
        'Debug.Print("")
        'Debug.Print("ProcessPropSeqNextItems() completed")
        'Debug.Print("Contents of the Property Sequence array: -------------------------------------------------------")
        'Debug.Print("Index     Path       Value         Next Item")
        'For I = 0 To mSequenceCount - 1
        '    Debug.Print(I & " " & mSequence(I).Path & " " & mSequence(I).Value & " " & mSequence(I).NextItem)
        'Next
        'Debug.Print("End of the Property Sequence -------------------------------------------------------------------")
        'Debug.Print("")

    End Sub

    'Sub RunPropSeq(ByRef SequenceStatus As String)
    'Sub RunPropSeq(ByRef Status As Collection)
    Sub RunPropSeq(ByRef Status As System.Collections.Specialized.StringCollection)
        'Runs the property sequence statements contained in the Sequence() array.
        'Status is a collection passed by the program calling this class.
        '  It can contain 0, 1 or more strings indicating the status of the processing such as "No_more_input_files" or "At_end_of_file".

        Dim LineNo As Integer = 0
        Dim Path As String = ""
        Dim Value As String = ""
        Dim NextItem As Integer = 0

        'SequenceStatus variable may also be modified in the code that calls this class!
        '  It will be declared in that code and passes by reference to this class.
        'Dim SequenceStatus As String = "" 'Valid status strings: "No_more_input_files", "At_end_of_file"

        While Path <> "End_Of_Sequence"
            Path = mSequence(LineNo).Path
            Value = mSequence(LineNo).Value
            NextItem = mSequence(LineNo).NextItem
            If Path <> "End_Of_Sequence" Then
                If Path = "Start_Loop" Then
                    LineNo = NextItem
                ElseIf Path = "End_Loop" Then
                    LineNo = NextItem 'The NextItem for an End_Loop statement is the Start_Loop line.
                ElseIf Path.EndsWith("ExitLoopIf") Then
                    If IsNothing(Status) Then 'Continue to next line
                        LineNo = LineNo + 1
                    ElseIf Status.Contains(Value) Then 'Exit loop
                        LineNo = NextItem 'The NextItem for an ExitLoopIf statement in the exit line.
                    Else 'Continue to the next line.
                        LineNo = LineNo + 1
                    End If
                    'If Value = Status Then 'Exit loop
                    '    LineNo = NextItem 'The NextItem for an ExitLoopIf statement is the exit line.
                    'Else
                    '    LineNo = LineNo + 1 'Continue to next line.
                    'End If
                    'Added May 5th:
                ElseIf Path.EndsWith("ExitGroupIf") Then
                    If IsNothing(Status) Then 'Continue to next line
                        LineNo = LineNo + 1
                    ElseIf Status.Contains(Value) Then 'Exit group
                        LineNo = NextItem
                    Else
                        LineNo = LineNo + 1
                    End If
                    'If Value = SequenceStatus Then 'Exit group
                    '    LineNo = NextItem
                    'Else
                    '    LineNo = LineNo + 1
                    'End If
                Else
                    PreScanInstruction(Path, Value)
                    LineNo = NextItem
                End If
            End If
        End While

        'RaiseEvent Instruction("EndOfSequence", "")
        RaiseEvent Instruction("", "EndOfSequence")


        'At end of sequence.
        'SequenceStatus = ""
        If IsNothing(Status) Then
        Else
            Status.Clear()
        End If


    End Sub

    Sub DisplaySequenceArray(ByRef XSeq As System.Xml.XmlDocument)
        'Display the contents of the Sequence() array for debugging:

        ExtractPropSeq(XSeq)
        ProcessPropSeqNextItems()

        Dim I As Integer 'Loop index

        'Debug.Print("Contents of the Sequence() array:")
        'Debug.WriteLine(String.Format("{0, -8}{1,-96}{2,-64}{3,-8}", "Item", "Path", "Value", "Next Item"))

        'For I = 1 To mSequenceCount
        '    Debug.WriteLine(String.Format("{0,-8}{1,-96}{2,-64}{3,-8}", I - 1, mSequence(I - 1).Path, mSequence(I - 1).Value, mSequence(I - 1).NextItem))
        'Next

    End Sub

    Private Sub PreScanInstruction(ByVal path As String, ByVal prop As String)
        'Applies a Processing Sequence Instruction - Initial Processing
        'An instruction consists of a path and a property value
        'The path points to an item on a form
        'The item is set to the property value (prop).

        Dim Path2 As String = "" 'The second part of the Instruction path
        Dim Path3 As String = "" 'The third part of the instruction path

        'For comparison:
        'Valid XSequence paths have the form:
        '    ProcessingSequence:<CreationDate>
        '    ProcessingSequence:<Description>
        '    ProcessingSequence:Process:<Path>
        '    Start_Loop
        '    End_Loop
        '    Start_Group
        '    End_Group
        '    End_Of_Sequence
        'Any other instruction returns an "Unknown Instruction" warning.

        'Value XMessage instructions have the form:
        '    XMsg:<Path>
        '    Start_Loop
        '    End_Loop
        '    Start_Group
        '    End_Group
        '    End_Of_Sequence
        'Any other paths returns an "Unknown Instruction" warning.

        ' If path.StartsWith("ProcessingSequence:") Then
        'If path.StartsWith("XMsg:") Then
        If path.StartsWith("XMsg:") Or path.StartsWith("XSys:") Then
            'Path2 = path.Substring(19, path.Length - 19) 'Get path string with ProcessingSequence: removed from the start.
            Path2 = path.Substring(5, path.Length - 5) 'Get path string with ProcessingSequence: removed from the start.
            'Path2 = path.Substring(startIndex:=, length)
            ProcessLoops(Path2, prop)
            'If Path2.StartsWith("Process:") Then
            '    Path3 = Path2.Substring(8, Path2.Length - 8) 'Get path string with Process: removed from the start.
            '    'Instruction has the form ProcessingSequence:Process:<Path3>
            '    ProcessLoops(Path3, prop)
            'Else
            '    'Instruction has the form ProcessingSequence:<Path2>
            '    Select Case Path2
            '        Case "CreationDate"
            '            'The creation date of the instruction sequence.
            '        Case "Description"
            '            'A description of the instruction sequence.
            '    End Select
            'End If
        Else
            'Instruction may be Start_Loop or End_Loop or End_Of_Sequence or <unknown>
            Select Case path
                Case "Start_Loop"
                    'At the start of a loop. Continue processing.
                Case "End_Loop"
                    'Point back to the start of the loop.

                Case "Start_Group"
                    'At the start of a group. Continue processing"
                Case "End_Group"


                Case "End_Of_Sequence"
                    'Stop processing!

                Case Else
                    'Unknown instruction!
                    RaiseEvent ErrorMsg("XMessage warning: Unknown instruction: " & path)
            End Select
        End If
    End Sub

    Private Sub ProcessLoops(ByVal path As String, ByVal prop As String)
        'Applies a Processing Sequence Instruction - Check for Loop: in the path
        'An instruction consists of a path and a property value.
        'The path either points to an item on a form or a property or method in theMain form.
        'The item is set to the property value.

        'Some statements have a multi level path:
        'ProcessingSequence:Process:DataPaths:TextFilesToProcess:TextFile  Value: C:\Users\Robert\Documents\Shares\asxdata\20090102.txt 

        'Some statements have a single level path:
        'ProcessingSequence:Process:Loop:ReadTextCommand  Value: ReadNextLine

        'Consider using single level path statements to simplify the instruction sequence code.

        Dim PathN As String = ""

        'Debug.Print("In method ProcessLoops(). Path = " & path)

        If path.StartsWith("Loop:") Then
            'Debug.Print("In method ProcessLoops. Path starts with Loop:. Path = " & path)
            'Instruction has the form Loop:<PathN>
            PathN = path.Substring(5, path.Length - 5)
            ProcessLoops(PathN, prop) 'Call the subroutine recursively until all Loop: and Group: components of the path are removed.
        ElseIf path.StartsWith("Group:") Then
            'Instruction has the form Group:<PathN>
            PathN = path.Substring(6, path.Length - 6)
            ProcessLoops(PathN, prop) 'Call the subroutine recursively until all Loop: and Group: components of the path are removed.
        Else
            'Instruction has the form <Path> <Value>
            'Apply the instruction:
            'RaiseEvent Instruction(path, prop)
            RaiseEvent Instruction(prop, path)
        End If
    End Sub



#Region " Sequence Status methods"

    'Adding Sequence Status strings:
    'SequenceStatus.Add("No_more_input_files")
    'SequenceStatus.Add("At_end_of_file")

    'Public Sub SequenceStatusRemove(ByVal Status As String)
    '    'Remove the specified Status string from the SequenceStatus collection if it is present:

    '    If IsNothing(SequenceStatus) Then
    '        'SequenceStatus contains no Status strings
    '    Else
    '        SequenceStatus.Remove(Status)
    '    End If
    'End Sub

    'Public Function SequenceStatusContains(ByVal Status As String) As Boolean
    '    'Check if the SequenceStatus collection contains the specified Status string:

    '    If IsNothing(SequenceStatus) Then
    '        Return False
    '    Else
    '        If SequenceStatus.Contains(Status) Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    End If
    'End Function

#End Region


#End Region 'Methods

#Region " Events"

    Public Event ErrorMsg(ByVal ErrMsg As String) 'Send an error message.

    'Public Event Instruction(ByVal Path As String, ByVal Prop As String) 'Send an instruction to be executed.
    'Public Event Instruction(ByVal Locn As String, ByVal Info As String) 'Send an instruction to be executed.
    Public Event Instruction(ByVal Info As String, ByVal Locn As String) 'Send an instruction to be executed.

#End Region 'Events

End Class 'XMessage -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Class FileLocation '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'A location is the path to a Directory or Archive used to store data.

#Region " Location Properties - Properties used to store location information." '-------------------------------------------------------------------------------------------------------------

    Public Enum Types
        Directory
        Archive
    End Enum

    Private _type As Types = Types.Directory 'The type of location (Directory or Archive).
    Property Type As Types
        Get
            Return _type
        End Get
        Set(value As Types)
            _type = value
        End Set
    End Property

    Private _path As String = "" 'The path to the location.
    Property Path As String
        Get
            Return _path
        End Get
        Set(value As String)
            _path = value
        End Set
    End Property

#End Region 'Location Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Location Methods" '------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub ReadXmlData(ByVal DataName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Read XML data in the Location stored with the name DataName. The data is read into XmlDoc.

        Select Case Type
            Case Types.Directory
                If System.IO.File.Exists(Path & "\" & DataName) Then
                    XmlDoc = XDocument.Load(Path & "\" & DataName)
                Else
                    XmlDoc = Nothing
                End If
            Case Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = Path
                If Zip.EntryExists(DataName) Then
                    Try
                        XmlDoc = XDocument.Parse(Zip.GetText(DataName))
                    Catch ex As Exception
                        RaiseEvent ErrorMessage(ex.Message)
                    End Try

                Else
                    XmlDoc = Nothing
                End If
        End Select
    End Sub


    Public Sub SaveXmlData(ByVal DataFileName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Save XML data in the Location with the name DataName. The data is read from XmlDoc.

        If DataFileName = "" Then
            RaiseEvent ErrorMessage("DataFileName is blank." & vbCrLf)
            Exit Sub
        End If

        If IsNothing(XmlDoc) Then
            RaiseEvent ErrorMessage("Error saving XML Data to file: " & DataFileName & vbCrLf)
            RaiseEvent ErrorMessage("  - No XML document to save." & DataFileName & vbCrLf)
        Else
            Select Case Type
                Case Types.Directory
                    If System.IO.Directory.Exists(Path) Then
                        XmlDoc.Save(Path & "\" & DataFileName)
                    Else

                    End If

                Case Types.Archive
                    If System.IO.File.Exists(Path) Then
                        Dim Zip As New ZipComp
                        Zip.ArchivePath = Path
                        Zip.AddText(DataFileName, XmlDoc.ToString)
                    Else
                        RaiseEvent ErrorMessage("Specified archive not found: " & Path & vbCrLf)
                    End If

            End Select
        End If
    End Sub

    Public Function FileExists(ByVal FileName As String) As Boolean
        'Returns True if a file with name FileName is found in the Location.

        Select Case Type
            Case Types.Directory
                Return System.IO.File.Exists(Path & "\" & FileName)
            Case Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = Path
                Return Zip.EntryExists(FileName)
        End Select
    End Function

    Public Sub ReadXmlDocData(ByVal DataFileName As String, ByRef XmlDoc As System.Xml.XmlDocument)
        'Version of the ReadXmlData that outputs the data into an XmlDocument (instead of and XDocument).

        If IsNothing(XmlDoc) Then
            RaiseEvent ErrorMessage("ReadXmlDocData error: The XmlDocument is blank." & vbCrLf)
        Else
            Select Case Type
                Case FileLocation.Types.Directory
                    'Read the Xml data document in the directory at DataLocn.Path
                    If System.IO.File.Exists(Path & "\" & DataFileName) Then
                        'XmlDoc = XDocument.Load(DataLocn.Path & "\" & DataFileName)
                        XmlDoc.Load(Path & "\" & DataFileName)
                    Else
                        XmlDoc = Nothing
                    End If

                Case FileLocation.Types.Archive
                    'Read the Xml Ddata document in the archive at DataLocn.Path
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = Path
                    If Zip.EntryExists(DataFileName) Then
                        XmlDoc.LoadXml(Zip.GetText(DataFileName))
                    Else
                        XmlDoc = Nothing
                    End If
                    Zip = Nothing
            End Select
        End If
    End Sub



#End Region 'Location Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Location Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Event ErrorMessage(ByVal Message As String)
    Event Message(ByVal Message As String)

#End Region 'Location Events ------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'Location --------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The Project class contains properties and methods used to manage an Andorville (TM) Application Project.
Public Class Project '------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable declarations - All the variables used in this class." '-------------------------------------------------------------------------------------------------------------------

    'The relative locations allow the SettingsLocn and DataLocn to be updated when the project file or archive is moved:
    Public SettingsRelLocn As New FileLocation 'This is the Settings location relative to the Project Location.
    Public DataRelLocn As New FileLocation 'This is the Data location relative to the Project Location.
    Public DataDirRelLocn As New FileLocation 'This is the DataDir location realtive to the Project Location.
    Public SystemRelLocn As New FileLocation 'This is the System location relative to the Project Location.
    'For a Directory project type: SettingsRelLocn.Type = Directory, SettingsRelLocn.Path = ""
    '                              DataRelLocn.Type = Directory,     DataRelLocn.Path = ""
    'For an Archive project type:  SettingsRelLocn.Type = Archive,   SettingsRelLocn.Path = ""
    '                              DataRelLocn.Type = Archive,       DataRelLocn.Path = ""
    'For a Hybrid project type:    The Project is stored in a Directory
    '                              SettingsRelLocn.Type = Directory or Archive, SettingsRelLocn.Path = <The name of the directory or archive within the Project Directory>
    '                              DataRelLocn.Type = Directory or Archive,     DataRelLocn.Path = <The name of the directory or archive within the Project Directory>
    'These settings and the Project Type and Path are used to generate the SettingsLocn and DataLocn.
    'Public ProjectLocn As New FileLocation 'This the the Project directory or Archive. (ADDED 29Jul18.) This contains the Settings Location and Data Location. UPDATE: USE THE TYPE AND PATH PROPERTIES>

    Public SettingsLocn As New FileLocation 'This is a directory or archive where settings are stored.
    'SettingsLocn.Type (Directory or Archive), SettingsLocn.Path
    Public DataLocn As New FileLocation 'This is a directory or archive where data is stored.
    'DataLocn.Type (Directory or Archive), DataLocn.Path
    Public DataDirLocn As New FileLocation 'This is a directory created to store data that is not suitable for a Data archive. (eg. pdf and xlsx files that are already compressed.)
    Public SystemLocn As New FileLocation 'This is a directory or archive where system data is stored. (System data can be preserved when settings or other data is deleted.)

    Public WithEvents ProjectForm As frmProject 'Used to select a project.
    Public WithEvents ProjectParamsForm As frmProjectParams 'Used to view the Project Parameters.

    'Public ProjectParam As New Dictionary(Of String, ParamInfo) 'Dictionary of Project Parameters.
    Public Parameter As New Dictionary(Of String, ParamInfo) 'Dictionary of Project Parameters.
    Public ParentParameter As New Dictionary(Of String, ParamInfo) 'Dictionary of Parent Project Parameters

    'USE Application.Name
    'Public ApplicationName As String 'The name of the application using the project.

    'Public ApplicationSummary As New ApplicationSummary 'This stores information about the application that created the project.
    'Public HostApplication As New ApplicationSummary 'This stores information about the application that created the project.
    Public Application As New ApplicationSummary 'This stores information about the application that created the project.
    Public Author As New Author 'Stores information about the project author.
    Public Version As New Version ' Stores the Project format version (based on the ADVL_Project_Utilities project version).
    Public Usage As New Usage 'Stores project usage information.

    'File conversion variables:
    Dim WithEvents ProjInfoConversion As ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion

#End Region 'Variable Declarations -----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Project Properties - Properties used to store project information." '---------------------------------------------------------------------------------------------------------------

    Private _applicationDir As String = "" 'The path to the directory used to store application data.
    Public Property ApplicationDir As String
        Get
            Return _applicationDir
        End Get
        Set(ByVal value As String)
            _applicationDir = value
        End Set
    End Property

    Private _name As String = "" 'The name of the current project.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A description of the current project.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    'Private _formatCode As String = "2" 'The Format Code identifies the format of the project. From 20 July 2018, the Format Code is "2".
    Private _formatCode As String = "ADVL_2" 'The Format Code identifies the format of the project. From 20 July 2018, the Format Code is "ADVL_2".
    Property FormatCode As String
        Get
            Return _formatCode
        End Get
        Set(value As String)
            _formatCode = value
        End Set
    End Property

    Private _creationDate As DateTime = "1-Jan-2000 12:00:00" 'The creation date of the current project.
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    Private _iD As String = "" 'Project ID. This is generated from the .GetHashCode of the string: <ProjectName> & " " & <CreationDate>
    Property ID As String
        Get
            Return _iD
        End Get
        Set(value As String)
            _iD = value
        End Set
    End Property

    Public Enum Types
        None      'All files are stored in the Application Directory.
        Directory 'All files are stored in the Project Directory.
        Archive      'All files are stored in the Project Archive.
        Hybrid    'Settings data is stored in the Project Directory. Data files are stored in one or more Project Files.
    End Enum

    Private _type As Types = Types.None 'The type of project (None, Directory, Archive, Hybrid).
    Property Type As Types
        Get
            Return _type
        End Get
        Set(value As Types)
            _type = value
            If _type = Types.Archive Then
                _locnType = FileLocation.Types.Archive
            ElseIf _type = Types.Directory Then
                _locnType = FileLocation.Types.Directory
            ElseIf _type = Types.Hybrid Then
                _locnType = FileLocation.Types.Directory
            ElseIf _type = Types.None Then
                _locnType = FileLocation.Types.Directory
            End If
        End Set
    End Property

    Private _locnType As FileLocation.Types = FileLocation.Types.Directory 'The location type of the Project location (Directory or Archive).
    ReadOnly Property LocnType As FileLocation.Types
        Get
            Return _locnType
        End Get
    End Property

    Private _path As String = "" 'The path to the project directory or archive.
    Property Path As String
        Get
            Return _path
        End Get
        Set(value As String)
            _path = value
        End Set
    End Property

    Private _relativePath As String = "" 'The path to the project directory (or archive?) relative to the Parent Project. (Currently Archive projects cannot contain child projects.)
    Property RelativePath As String
        Get
            Return _relativePath
        End Get
        Set(value As String)
            _relativePath = value
        End Set
    End Property

    'Private _hostProjectName As String = "" 'The name of the host project. (This is blank if there is no host project.)
    'Property HostProjectName As String
    '    Get
    '        Return _hostProjectName
    '    End Get
    '    Set(value As String)
    '        _hostProjectName = value
    '    End Set
    'End Property

    Private _parentProjectName As String = "" 'The name of the host project. (This is blank if there is no host project.)
    Property ParentProjectName As String
        Get
            Return _parentProjectName
        End Get
        Set(value As String)
            _parentProjectName = value
        End Set
    End Property

    'Private _hostProjectDirectoryName As String = "" 'The name of the host project directory.
    'Property HostProjectDirectoryName As String
    '    Get
    '        Return _hostProjectDirectoryName
    '    End Get
    '    Set(value As String)
    '        _hostProjectDirectoryName = value
    '    End Set
    'End Property

    Private _parentProjectType As Types = Types.None  'The type of parent project (None, Directory, Archive, Hybrid).
    Property ParentProjectType As Types
        Get
            Return _parentProjectType
        End Get
        Set(value As Types)
            _parentProjectType = value
        End Set
    End Property

    Private _parentProjectPath As String = "" 'The path of the Parent Project.
    Property ParentProjectPath As String
        Get
            Return _parentProjectPath
        End Get
        Set(value As String)
            _parentProjectPath = value
        End Set
    End Property

    Private _parentProjectDirectoryName As String = "" 'The name of the Parent Project directory.
    Property ParentProjectDirectoryName As String
        Get
            Return _parentProjectDirectoryName
        End Get
        Set(value As String)
            _parentProjectDirectoryName = value
        End Set
    End Property

    'Private _hostProjectCreationDate As DateTime = "1-Jan-2000 12:00:00" 'The creation date of the host project.
    'Property HostProjectCreationDate As Date
    '    Get
    '        Return _hostProjectCreationDate
    '    End Get
    '    Set(value As Date)
    '        _hostProjectCreationDate = value
    '    End Set
    'End Property

    Private _parentProjectCreationDate As DateTime = "1-Jan-2000 12:00:00" 'The creation date of the host project.
    Property ParentProjectCreationDate As Date
        Get
            Return _parentProjectCreationDate
        End Get
        Set(value As Date)
            _parentProjectCreationDate = value
        End Set
    End Property

    'Private _hostProjectID As String = "" 'The host project ID. (This is a hash of the host project name and creation date.)
    'Property HostProjectID As String
    '    Get
    '        Return _hostProjectID
    '    End Get
    '    Set(value As String)
    '        _hostProjectID = value
    '    End Set
    'End Property

    Private _parentProjectID As String = "" 'The host project ID. (This is a hash of the host project name and creation date.)
    Property ParentProjectID As String
        Get
            Return _parentProjectID
        End Get
        Set(value As String)
            _parentProjectID = value
        End Set
    End Property

    'ADDED 20Feb19:
    Private _connectOnOpen As Boolean = True 'It True, the application will connect to the Message Service when the Project is opened.
    Property ConnectOnOpen As Boolean
        Get
            Return _connectOnOpen
        End Get
        Set(value As Boolean)
            _connectOnOpen = value
        End Set
    End Property

#End Region 'Project Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Project Methods" '------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Function SettingsLocked() As Boolean
        'Returns True if a lock file is found in the settings location.
        Return SettingsFileExists("Settings.Lock")
    End Function

    Public Function DataLocked() As Boolean
        'Returns True if a lock file is found in the data location.
        Return DataFileExists("Data.Lock")
    End Function

    Public Function SystemLocked() As Boolean
        'Returns True is a lock file is found in the system data location.
        Return SystemFileExists("System.Lock")
    End Function

    Public Function ProjectLocked() As Boolean
        'Returns True if a project lock file is found in the settings location.
        'Return SettingsFileExists("Project.Lock")
        'Return ProjectFileExists("Project.Lock")
        Return ProjectFileExists("Project.Lock")
    End Function

    Public Function LockedAtPath(ByVal Path As String) As Boolean
        'Returns True if the project at Path contains a "Project.Lock" file.

        If Path.EndsWith(".AdvlProject") Then 'Archive Project
            'Check if the File "Project.Lock" exists in the project archive:
            Dim Zip As New ZipComp
            Zip.ArchivePath = Path
            If Zip.ArchiveExists Then
                Return Zip.EntryExists("Project.Lock")
            Else
                Return False
            End If
        Else 'Directory or Hybrid or Default Project.
            Return System.IO.File.Exists(Path & "\Project.Lock")
        End If
    End Function

    Public Sub SaveLastProjectInfo()
        'Save information about the last project used.
        'This is saved in the Application Directory.

        Dim ProjectInfoXDoc = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Last Project Information-->
                              <!---->
                              <Project>
                                  <Name><%= Name %></Name>
                                  <Description><%= Description %></Description>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <Location>
                                      <Type><%= Type %></Type>
                                      <Path><%= Path %></Path>
                                  </Location>
                              </Project>

        'ProjectInfoXDoc.Save(ApplicationDir & "\" & "Last_Project_Info.xml")
        ProjectInfoXDoc.Save(ApplicationDir & "\" & "Last_Project_Info_ADVL_2.xml")
    End Sub

    Public Sub SaveLastProjectInfo_Old()
        'Save information about the last project used.
        'This is saved in the Application Directory.
        'UPDATE: 17Jul18: The Project Information file: ADVL_Project_Info.xml is now stored in the Project Location (not the SettingsLocn).

        Dim SettingsLocationType As String
        Select Case SettingsLocn.Type
            Case FileLocation.Types.Directory
                SettingsLocationType = "Directory"
            Case FileLocation.Types.Archive
                SettingsLocationType = "Archive"
        End Select

        Dim ProjectInfoXDoc = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Last Project Information for Application: ADVL_Zip-->
                              <!---->
                              <Project>
                                  <Name><%= Name %></Name>
                                  <Description><%= Description %></Description>
                                  <Location>
                                      <Type><%= Type %></Type>
                                      <Path><%= Path %></Path>
                                  </Location>
                              </Project>

        ProjectInfoXDoc.Save(ApplicationDir & "\" & "Last_Project_Info.xml")
    End Sub

    Public Sub ReadLastProjectInfo()
        'Read the Last Project Information.
        '  The Last Project Information file is in the Application Directory.

        If System.IO.File.Exists(ApplicationDir & "\" & "Last_Project_Info_ADVL_2.xml") Then 'The Last Project Info file exists.
            Dim ProjInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\" & "Last_Project_Info_ADVL_2.xml")
            ReadLastProjectInfoAdvl_2(ProjInfoXDoc)
        Else
            If System.IO.File.Exists(ApplicationDir & "\" & "Last_Project_Info.xml") Then 'The original ADVL_1 version
                RaiseEvent Message("Converting Last_Project_Info.xml to Last_Project_Info_ADVL_2.xml." & vbCrLf)
                'Convert the file to the latest ADVL_2 format:
                Dim Conversion As New ADVL_Utilities_Library_1.FormatConvert.LastProjectInfoFileConversion
                Conversion.DirectoryPath = ApplicationDir
                Conversion.InputFileName = "Last_Project_Info.xml"
                Conversion.InputFormatCode = FormatConvert.LastProjectInfoFileConversion.FormatCodes.ADVL_1
                Conversion.OutputFormatCode = FormatConvert.LastProjectInfoFileConversion.FormatCodes.ADVL_2
                Conversion.Convert()
                If System.IO.File.Exists(ApplicationDir & "\" & "Last_Project_Info_ADVL_2.xml") Then
                    ReadLastProjectInfo() 'Try to read the file again.
                Else
                    RaiseEvent ErrorMessage("Error converting Last_Project_Info.xml to Last_Project_Info_ADVL_2.xml." & vbCrLf)
                End If
            Else
                RaiseEvent ErrorMessage("No versions of the Last Project Information file were found." & vbCrLf)
                RaiseEvent ErrorMessage("The Default project will be used." & vbCrLf)
                UseDefaultProject()
            End If
        End If
    End Sub

    Private Sub ReadLastProjectInfoAdvl_2(ByRef XDoc As System.Xml.Linq.XDocument)
        'Read the the ADVL_2 format version of the Last Project Information file.

        'Read the Project Name:
        If XDoc.<Project>.<Name>.Value <> Nothing Then
            Name = XDoc.<Project>.<Name>.Value
        Else
            RaiseEvent ErrorMessage("The Project Name is not specified in the Last Project Information file." & vbCrLf)
            RaiseEvent ErrorMessage("The Default project will be opened." & vbCrLf)
            UseDefaultProject()
            Exit Sub
        End If

        'Read the Project Description:
        If XDoc.<Project>.<Description>.Value <> Nothing Then
            Description = XDoc.<Project>.<Description>.Value
        Else
            RaiseEvent ErrorMessage("The Project Description is not specified in the Last Project Information file." & vbCrLf)
        End If

        'Read the Project Location Type:
        If XDoc.<Project>.<Location>.<Type>.Value <> Nothing Then
            Select Case XDoc.<Project>.<Location>.<Type>.Value '(None, Directory, Archive or Hybrid)
                Case "None"
                    Type = Types.None
                    Usage.SaveLocn.Type = FileLocation.Types.Directory
                Case "Directory"
                    Type = Types.Directory
                    Usage.SaveLocn.Type = FileLocation.Types.Directory
                Case "Archive"
                    Type = Types.Archive
                    Usage.SaveLocn.Type = FileLocation.Types.Archive
                Case "Hybrid"
                    Type = Types.Hybrid
                    Usage.SaveLocn.Type = FileLocation.Types.Directory
            End Select
        Else
            RaiseEvent ErrorMessage("The Project Location Type is not specified in the Last Project Information file." & vbCrLf)
        End If

        'Read the Project Location Path:
        If XDoc.<Project>.<Location>.<Path>.Value <> Nothing Then
            Path = XDoc.<Project>.<Location>.<Path>.Value
            Usage.SaveLocn.Path = Path
        Else
            RaiseEvent ErrorMessage("The Project Location Path is not specified in the Last Project Information file." & vbCrLf)
        End If

    End Sub

    Private Sub UseDefaultProject()
        'Open the Default project.
        'Create the Default project if it doesn't already exist.

        If System.IO.Directory.Exists(ApplicationDir & "\" & "Default_Project") Then

        Else
            CreateDefaultProject()
        End If
        'Set setting to Default Project:
        'SettingsLocn.Type = FileLocation.Types.Directory
        Type = Types.Directory
        Usage.SaveLocn.Type = FileLocation.Types.Directory
        'SettingsLocn.Path = ApplicationDir & "\" & "Default_Project"
        Path = ApplicationDir & "\" & "Default_Project"
        Usage.SaveLocn.Path = ApplicationDir & "\" & "Default_Project"
        Name = "Default"
        Description = "Default project. Data and settings are stored in the Default_Project directory in the Application Directory."
        Type = Types.None
    End Sub

    Public Function ProjectInfoFileExists() As Boolean
        'Return System.IO.File.Exists(ApplicationDir & "\" & "ADVL_Project_Info.xml")
        Return SettingsFileExists("ADVL_Project_Info.xml")
    End Function

    Public Sub ReadProjectInfoFile()
        'Read the Project information in a Project Information file.

        'First check if the project exists:
        Select Case Type
            Case Types.Archive
                If System.IO.File.Exists(Path) Then
                    'Archive file found.
                Else
                    UseDefaultProject()
                    Exit Sub
                End If
            Case Types.Directory
                If System.IO.Directory.Exists(Path) Then
                    'The Project directory found.
                Else
                    UseDefaultProject()
                    Exit Sub
                End If
            Case Types.Hybrid
                If System.IO.Directory.Exists(Path) Then
                    'The Project directory found.
                Else
                    UseDefaultProject()
                    Exit Sub
                End If
            Case Types.None
                'Use the default project.
                'Check if the Default_Project directory exists in the Application Directory:
                If System.IO.Directory.Exists(ApplicationDir & "\Default_Project") Then
                    'Default_Project directory found.
                Else
                    UseDefaultProject()
                    Exit Sub
                End If
        End Select

        If ProjectFileExists("Project_Info_ADVL_2.xml") Then 'The Project Information file exists (ADVL_2 format version).
            Dim ProjectInfoXDoc As System.Xml.Linq.XDocument
            ReadXmlProjectFile("Project_Info_ADVL_2.xml", ProjectInfoXDoc)
            ReadProjectInfoFileAdvl_2(ProjectInfoXDoc)
        Else
            If ProjectFileExists("ADVL_Project_Info.xml") Then 'The original ADVL_1 format version of the Project Information file exists.
                RaiseEvent Message("Converting ADVL_Project_Info.xml to Project_Info_ADVL_2.xml." & vbCrLf)
                'Convert the file to the latest ADVL_2 format:
                ProjInfoConversion = New ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion
                'ProjInfoConversion.ProjectType = Type
                'NOTE: PROJECTTYPE DOES NOT USE THE SAME ENUMS AS TYPE!!!
                Select Case Type
                    Case Types.Archive
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Archive
                    Case Types.Directory
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Directory
                    Case Types.Hybrid
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Directory
                    Case Types.None
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Directory
                End Select
                RaiseEvent Message("Project Path: " & Path & vbCrLf)
                ProjInfoConversion.ProjectPath = Path
                ProjInfoConversion.InputFileName = "ADVL_Project_Info.xml"
                ProjInfoConversion.InputFormatCode = FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_1
                ProjInfoConversion.OutputFormatCode = FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_2
                ProjInfoConversion.Convert()
                If ProjectFileExists("Project_Info_ADVL_2.xml") Then
                    ReadProjectInfoFile() 'Try ReadProjectInfoFile again. This time Project_Info_ADVL_2.xml should be found
                Else
                    RaiseEvent ErrorMessage("Error converting ADVL_Project_Info.xml to Project_Info_ADVL_2.xml." & vbCrLf)
                End If
            Else
                RaiseEvent ErrorMessage("No versions of the Project Information file were found." & vbCrLf)
            End If
        End If
    End Sub

    Public Function ReadProjectInfoFileAtPath() As Boolean
        'Read the Project Information file at the location in Path.

        If Path.EndsWith(".AdvlProject") Then 'Archive project.
            If System.IO.File.Exists(Path) Then
                'Archive file found.
                Type = Types.Archive
            Else
                UseDefaultProject()
                'Exit Sub
                Return False 'Project Info file could not be read at Path
                Exit Function
            End If
        Else 'Default, Directory or Hybrid project.
            If System.IO.Directory.Exists(Path) Then
                'The Project directory found.
                Type = Types.Directory 'Assume the project type is directory. This can be updated after the project info file is read.
            Else
                UseDefaultProject()
                Return False 'Project Info file could not be read at Path
                'Exit Sub
                Exit Function
            End If
        End If

        If ProjectFileExists("Project_Info_ADVL_2.xml") Then 'The Project Information file exists (ADVL_2 format version).
            Dim ProjectInfoXDoc As System.Xml.Linq.XDocument
            ReadXmlProjectFile("Project_Info_ADVL_2.xml", ProjectInfoXDoc)
            ReadProjectInfoFileAdvl_2(ProjectInfoXDoc)
            Return True 'Project Info file could be read at Path
        Else
            If ProjectFileExists("ADVL_Project_Info.xml") Then 'The original ADVL_1 format version of the Project Information file exists.
                RaiseEvent Message("Converting ADVL_Project_Info.xml to Project_Info_ADVL_2.xml." & vbCrLf)
                'Convert the file to the latest ADVL_2 format:
                ProjInfoConversion = New ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion
                'ProjInfoConversion.ProjectType = Type
                'NOTE: PROJECTTYPE DOES NOT USE THE SAME ENUMS AS TYPE!!!
                Select Case Type
                    Case Types.Archive
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Archive
                    Case Types.Directory
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Directory
                    Case Types.Hybrid
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Directory
                    Case Types.None
                        ProjInfoConversion.ProjectType = FormatConvert.ProjectInfoFileConversion.ProjectTypes.Directory
                End Select
                RaiseEvent Message("Project Path: " & Path & vbCrLf)
                ProjInfoConversion.ProjectPath = Path
                ProjInfoConversion.InputFileName = "ADVL_Project_Info.xml"
                ProjInfoConversion.InputFormatCode = FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_1
                ProjInfoConversion.OutputFormatCode = FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_2
                ProjInfoConversion.Convert()
                If ProjectFileExists("Project_Info_ADVL_2.xml") Then
                    ReadProjectInfoFile() 'Try ReadProjectInfoFile again. This time Project_Info_ADVL_2.xml should be found
                Else
                    RaiseEvent ErrorMessage("Error converting ADVL_Project_Info.xml to Project_Info_ADVL_2.xml." & vbCrLf)
                    Return False 'Project Info file could not be read at Path
                End If
            Else
                RaiseEvent ErrorMessage("No versions of the Project Information file were found." & vbCrLf)
                Return False 'Project Info file could not be read at Path
            End If
        End If

    End Function

    Private Sub ReadProjectInfoFileAdvl_2(ByRef XDoc As System.Xml.Linq.XDocument)
        'Read the Project Information XDocument (ADVL_2 format version).

        FormatCode = "ADVL_2"
        'Read the Project Name:
        If XDoc.<Project>.<Name>.Value = Nothing Then
            Name = ""
        Else
            Name = XDoc.<Project>.<Name>.Value
        End If

        'Read the Project Type
        If XDoc.<Project>.<Type>.Value = Nothing Then
            Type = Types.None
        Else
            Select Case XDoc.<Project>.<Type>.Value
                Case "None"
                    Type = Types.None
                    'SettingsLocn = ProjectLocn 'Use Default project: Settings location = Project location.
                    'DataLocn = ProjectLocn     'Use Default project: Data location = Project location.
                    'Note: If Type = None, the Default project will be used.
                    SettingsLocn.Type = FileLocation.Types.Directory
                    SettingsLocn.Path = Path
                    DataLocn.Type = FileLocation.Types.Directory
                    DataLocn.Path = Path
                    DataDirLocn.Type = FileLocation.Types.Directory 'For the Default project, the DataDirLocn is the same as the DataLocn. (It provides an alternative uncompressed location to an archive Data location.)
                    DataDirLocn.Path = Path
                    SystemLocn.Type = FileLocation.Types.Directory
                    SystemLocn.Path = Path
                    Usage.SaveLocn.Type = FileLocation.Types.Directory
                    Usage.SaveLocn.Path = Path
                Case "Archive"
                    Type = Types.Archive
                    'SettingsLocn = ProjectLocn 'For Archive type, Settings location is the same as the Project location.
                    'DataLocn = ProjectLocn     'For Archive type, Data location is the same as the Project location.
                    'Note: If Type = Archive, the SettingsLocn and DataLocn are both the project archive file.
                    SettingsLocn.Type = FileLocation.Types.Archive
                    SettingsLocn.Path = Path
                    DataLocn.Type = FileLocation.Types.Archive
                    DataLocn.Path = Path
                    'Use the Archive location for the DataDirLocn - it is not possible to use a Directory in this case.
                    'DataDirLocn.Type = FileLocation.Types.Directory
                    DataDirLocn.Type = FileLocation.Types.Archive
                    'DataDirLocn.Path = "" 'There is no DataDirLocn in an Archive project.
                    DataDirLocn.Path = Path 'The only option is to use the Archive path!.
                    SystemLocn.Type = FileLocation.Types.Archive
                    SystemLocn.Path = Path
                    Usage.SaveLocn.Type = FileLocation.Types.Archive
                    Usage.SaveLocn.Path = Path
                Case "Directory"
                    Type = Types.Directory
                    'SettingsLocn = ProjectLocn 'For Directory type, Settings location is the same as the Project location.
                    'DataLocn = ProjectLocn     'For Directory type, Data location is the same as the Project location.
                    'Note: If Type = Directory, the SettingsLocn and DataLocn are both the project directory.
                    SettingsLocn.Type = FileLocation.Types.Directory
                    SettingsLocn.Path = Path
                    DataLocn.Type = FileLocation.Types.Directory
                    DataLocn.Path = Path
                    DataDirLocn.Type = FileLocation.Types.Directory 'For a Directory project, the DataDirLocn is the same as the DataLocn. (It provides an alternative uncompressed location to an archive Data location.)
                    DataDirLocn.Path = Path
                    SystemLocn.Type = FileLocation.Types.Directory
                    SystemLocn.Path = Path
                    Usage.SaveLocn.Type = FileLocation.Types.Directory
                    Usage.SaveLocn.Path = Path
                Case "Hybrid"
                    Type = Types.Hybrid
                    'A Hybrid project is stored in a directory with the settings stored in the directory, a sub-directory or a sub-archive.
                    If XDoc.<Project>.<SettingsRelativeLocation>.<Type>.Value = Nothing Then
                        RaiseEvent ErrorMessage("No Settings relative location type is specified for the Hybrid project." & vbCrLf)
                    Else
                        Select Case XDoc.<Project>.<SettingsRelativeLocation>.<Type>.Value
                            Case "Directory"
                                SettingsLocn.Type = FileLocation.Types.Directory
                                If XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No Settings relative location path is specified for the Hybrid project." & vbCrLf)
                                Else
                                    If XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = "" Then
                                        SettingsLocn.Path = Path 'The Settings Path is the same as the Project Directory
                                    Else
                                        'SettingsLocn.Path = Path & "\" & XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value 'The Settings Path is a sub-directory of the Project Directory.
                                        SettingsLocn.Path = Path & XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value 'The Settings Path is a sub-directory of the Project Directory.
                                    End If
                                End If

                            Case "Archive"
                                SettingsLocn.Type = FileLocation.Types.Archive
                                If XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No Settings relative location path is specified for the Hybrid project." & vbCrLf)
                                Else
                                    If XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = "" Then
                                        SettingsLocn.Path = Path 'The Settings Path is the same as the Project Directory
                                        RaiseEvent ErrorMessage("The Settings relative location path is an Archive with a blank name specified." & vbCrLf)
                                        RaiseEvent ErrorMessage("The Project path will be used." & vbCrLf)
                                    Else
                                        'SettingsLocn.Path = Path & "\" & XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value 'The Settings Path is an Archive in the Project Directory.
                                        SettingsLocn.Path = Path & XDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value 'The Settings Path is an Archive in the Project Directory.
                                    End If
                                End If

                        End Select
                    End If

                    'A Hybrid project is stored in a directory with the data stored in the directory, a sub-directory or a sub-archive.
                    If XDoc.<Project>.<DataRelativeLocation>.<Type>.Value = Nothing Then
                        RaiseEvent ErrorMessage("No Data relative location type is specified for the Hybrid project." & vbCrLf)
                    Else
                        Select Case XDoc.<Project>.<DataRelativeLocation>.<Type>.Value
                            Case "Directory"
                                DataLocn.Type = FileLocation.Types.Directory
                                If XDoc.<Project>.<DataRelativeLocation>.<Path>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No Data relative location path is specified for the Hybrid project." & vbCrLf)
                                Else
                                    If XDoc.<Project>.<DataRelativeLocation>.<Path>.Value = "" Then
                                        DataLocn.Path = Path 'The Data Path is the same as the Project Directory
                                    Else
                                        DataLocn.Path = Path & XDoc.<Project>.<DataRelativeLocation>.<Path>.Value 'The Data Path is a sub-directory of the Project Directory.
                                    End If
                                End If
                                'For a Hybrid project with a Directory Data Location, the DataDirLocn is the same as the DataLocn. (It provides an alternative uncompressed location to an archive Data location.)
                                DataDirLocn.Type = FileLocation.Types.Directory
                                DataDirLocn.Path = DataLocn.Path

                            Case "Archive"
                                DataLocn.Type = FileLocation.Types.Archive
                                DataDirLocn.Type = FileLocation.Types.Directory
                                If XDoc.<Project>.<DataRelativeLocation>.<Path>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No Data relative location path is specified for the Hybrid project." & vbCrLf)
                                Else
                                    If XDoc.<Project>.<DataRelativeLocation>.<Path>.Value = "" Then
                                        DataLocn.Path = Path 'The Data Path is the same as the Project Directory
                                        RaiseEvent ErrorMessage("The Data relative location path is an Archive with a blank name specified." & vbCrLf)
                                        RaiseEvent ErrorMessage("The Project path will be used." & vbCrLf)
                                    Else
                                        DataLocn.Path = Path & XDoc.<Project>.<DataRelativeLocation>.<Path>.Value 'The Data Path is an Archive in the Project Directory.
                                        Dim DataDirName As String = System.IO.Path.GetFileNameWithoutExtension(DataLocn.Path) 'The Data directory name is the same as the Archive file name with the file extension removed.
                                        DataDirLocn.Path = Path & "\" & DataDirName
                                    End If
                                End If

                        End Select
                    End If

                    'A Hybrid project is stored in a directory with the system data stored in the directory, a sub-directory or a sub-archive.
                    If XDoc.<Project>.<SystemRelativeLocation>.<Type>.Value = Nothing Then
                        RaiseEvent ErrorMessage("No System relative location type is specified for the Hybrid project." & vbCrLf)
                    Else
                        Select Case XDoc.<Project>.<SystemRelativeLocation>.<Type>.Value
                            Case "Directory"
                                SystemLocn.Type = FileLocation.Types.Directory
                                If XDoc.<Project>.<SystemRelativeLocation>.<Path>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No System relative location path is specified for the Hybrid project." & vbCrLf)
                                Else
                                    If XDoc.<Project>.<SystemRelativeLocation>.<Path>.Value = "" Then
                                        SystemLocn.Path = Path 'The Data Path is the same as the Project Directory
                                    Else
                                        SystemLocn.Path = Path & XDoc.<Project>.<SystemRelativeLocation>.<Path>.Value 'The System Path is a sub-directory of the Project Directory.
                                    End If
                                End If

                            Case "Archive"
                                SystemLocn.Type = FileLocation.Types.Archive
                                If XDoc.<Project>.<SystemRelativeLocation>.<Path>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No System relative location path is specified for the Hybrid project." & vbCrLf)
                                Else
                                    If XDoc.<Project>.<SystemRelativeLocation>.<Path>.Value = "" Then
                                        SystemLocn.Path = Path 'The System Path is the same as the Project Directory
                                        RaiseEvent ErrorMessage("The System relative location path is an Archive with a blank name specified." & vbCrLf)
                                        RaiseEvent ErrorMessage("The Project path will be used." & vbCrLf)
                                    Else
                                        SystemLocn.Path = Path & XDoc.<Project>.<SystemRelativeLocation>.<Path>.Value 'The Data Path is an Archive in the Project Directory.
                                    End If
                                End If

                        End Select
                    End If

                Case Else
                    Type = Types.None
                    RaiseEvent ErrorMessage("Unknown Project Type: " & XDoc.<Project>.<Type>.Value & vbCrLf)
            End Select
        End If

        'Read the project Description
        If XDoc.<Project>.<Description>.Value = Nothing Then
            Description = ""
        Else
            Description = XDoc.<Project>.<Description>.Value
        End If

        'Read the project Creation Date:
        If XDoc.<Project>.<CreationDate>.Value = Nothing Then
            CreationDate = "1-Jan-2000 12:00:00"
        Else
            CreationDate = XDoc.<Project>.<CreationDate>.Value
        End If

        'The Project Information File Format Code has already beeen read.

        'Read Project ID:
        If XDoc.<Project>.<ID>.Value = Nothing Then
            'Leave the ID blank.
        Else
            ID = XDoc.<Project>.<ID>.Value
        End If

        'Read the Parent Project Name:
        If XDoc.<Project>.<ParentProject>.<Name>.Value = Nothing Then
            ParentProjectName = ""
        Else
            ParentProjectName = XDoc.<Project>.<ParentProject>.<Name>.Value
        End If

        'Read the Parent Project Directory Name:
        If XDoc.<Project>.<ParentProject>.<DirectoryName>.Value = Nothing Then
            ParentProjectDirectoryName = ""
        Else
            ParentProjectDirectoryName = XDoc.<Project>.<ParentProject>.<DirectoryName>.Value
        End If

        'Read the Parent Project CreationDate:
        If XDoc.<Project>.<ParentProject>.<CreationDate>.Value = Nothing Then
            ParentProjectCreationDate = "1-Jan-2000 12:00:00"
        Else
            ParentProjectCreationDate = XDoc.<Project>.<ParentProject>.<CreationDate>.Value
        End If

        'Read the Parent Project ID:
        If XDoc.<Project>.<ParentProject>.<ID>.Value = Nothing Then
            ParentProjectID = ""
        Else
            ParentProjectID = XDoc.<Project>.<ParentProject>.<ID>.Value
        End If

        'Read the Parent Project Path:
        If XDoc.<Project>.<ParentProject>.<Path>.Value = Nothing Then
            ParentProjectPath = ""
        Else
            ParentProjectPath = XDoc.<Project>.<ParentProject>.<Path>.Value
        End If

        'Read the Project Author information:
        If XDoc.<Project>.<Author>.<Name>.Value = Nothing Then
            Author.Name = ""
        Else
            Author.Name = XDoc.<Project>.<Author>.<Name>.Value
        End If
        If XDoc.<Project>.<Author>.<Description>.Value = Nothing Then
            Author.Description = ""
        Else
            Author.Description = XDoc.<Project>.<Author>.<Description>.Value
        End If
        If XDoc.<Project>.<Author>.<Contact>.Value = Nothing Then
            Author.Contact = ""
        Else
            Author.Contact = XDoc.<Project>.<Author>.<Contact>.Value
        End If

        'Read the Application information: 
        If XDoc.<Project>.<Application>.<Name>.Value = Nothing Then
            Application.Name = ""
        Else
            Application.Name = XDoc.<Project>.<Application>.<Name>.Value
        End If
        If XDoc.<Project>.<Application>.<Description>.Value = Nothing Then
            Application.Description = ""
        Else
            Application.Description = XDoc.<Project>.<Application>.<Description>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<CreationDate>.Value = Nothing Then
            Application.CreationDate = "1-Jan-2000 12:00:00"
        Else
            Application.CreationDate = XDoc.<Project>.<Application>.<CreationDate>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<Version>.<Major>.Value = Nothing Then
            Application.Version.Major = 0
        Else
            Application.Version.Major = XDoc.<Project>.<Application>.<Version>.<Major>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<Version>.<Minor>.Value = Nothing Then
            Application.Version.Minor = 0
        Else
            Application.Version.Minor = XDoc.<Project>.<Application>.<Version>.<Minor>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<Version>.<Build>.Value = Nothing Then
            Application.Version.Build = 0
        Else
            Application.Version.Build = XDoc.<Project>.<Application>.<Version>.<Build>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<Version>.<Revision>.Value = Nothing Then
            Application.Version.Revision = 0
        Else
            Application.Version.Revision = XDoc.<Project>.<Application>.<Version>.<Revision>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<Author>.<Name>.Value = Nothing Then
            Application.Author.Name = ""
        Else
            Application.Author.Name = XDoc.<Project>.<Application>.<Author>.<Name>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<Author>.<Description>.Value = Nothing Then
            Application.Author.Description = ""
        Else
            Application.Author.Description = XDoc.<Project>.<Application>.<Author>.<Description>.Value
        End If
        If XDoc.<Project>.<HostApplication>.<Author>.<Contact>.Value = Nothing Then
            Application.Author.Contact = ""
        Else
            Application.Author.Contact = XDoc.<Project>.<Application>.<Author>.<Contact>.Value
        End If

        If XDoc.<Project>.<ConnectOnOpen>.Value = Nothing Then
            ConnectOnOpen = True
        Else
            ConnectOnOpen = XDoc.<Project>.<ConnectOnOpen>.Value
        End If

        Usage.RestoreUsageInfo()
        Usage.StartTime = Now

    End Sub

    Public Sub ReadProjectInfoFile_Old2()
        'Read the Project information in a Project Information file.

        If ProjectFileExists("ADVL_Project_Info.xml") Then 'The Project Information file exists.
            Dim ProjectInfoXDoc As System.Xml.Linq.XDocument
            ReadXmlProjectFile("ADVL_Project_Info.xml", ProjectInfoXDoc)

            If ProjectInfoXDoc Is Nothing Then
                RaiseEvent ErrorMessage("Error reading the ADVL_Project_Info.xml file." & vbCrLf)
            Else
                'The ADVL_Project_Info.xml file was read.
                If ProjectInfoXDoc.<Project>.<FormatCode>.Value = Nothing Then
                    RaiseEvent ErrorMessage("The Project Information file has no format code. Run the Format Convert application to update the format." & vbCrLf)
                ElseIf ProjectInfoXDoc.<Project>.<FormatCode>.Value = "ADVL_2" Then 'Read the Project Information.
                    FormatCode = "ADVL_2"
                    'Read the Project Name:
                    If ProjectInfoXDoc.<Project>.<Name>.Value = Nothing Then
                        Name = ""
                    Else
                        Name = ProjectInfoXDoc.<Project>.<Name>.Value
                    End If

                    'Read the Project Type
                    If ProjectInfoXDoc.<Project>.<Type>.Value = Nothing Then
                        Type = Types.None
                    Else
                        Select Case ProjectInfoXDoc.<Project>.<Type>.Value
                            Case "None"
                                Type = Types.None
                                'SettingsLocn = ProjectLocn 'Use Default project: Settings location = Project location.
                                'DataLocn = ProjectLocn     'Use Default project: Data location = Project location.
                                'Note: If Type = None, the Default project will be used.
                                SettingsLocn.Type = FileLocation.Types.Directory
                                SettingsLocn.Path = Path
                                DataLocn.Type = FileLocation.Types.Directory
                                DataLocn.Path = Path
                            Case "Archive"
                                Type = Types.Archive
                                'SettingsLocn = ProjectLocn 'For Archive type, Settings location is the same as the Project location.
                                'DataLocn = ProjectLocn     'For Archive type, Data location is the same as the Project location.
                                'Note: If Type = Archive, the SettingsLocn and DataLocn are both the project archive file.
                                SettingsLocn.Type = FileLocation.Types.Archive
                                SettingsLocn.Path = Path
                                DataLocn.Type = FileLocation.Types.Archive
                                DataLocn.Path = Path
                            Case "Directory"
                                Type = Types.Directory
                                'SettingsLocn = ProjectLocn 'For Directory type, Settings location is the same as the Project location.
                                'DataLocn = ProjectLocn     'For Directory type, Data location is the same as the Project location.
                                'Note: If Type = Directory, the SettingsLocn and DataLocn are both the project directory.
                                SettingsLocn.Type = FileLocation.Types.Directory
                                SettingsLocn.Path = Path
                                DataLocn.Type = FileLocation.Types.Directory
                                DataLocn.Path = Path
                            Case "Hybrid"
                                Type = Types.Hybrid
                                'A Hybrid project is stored in a directory with the settings stored in the directory, a sub-directory or a sub-archive.
                                If ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Type>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No Settings relative location type is specified for the Hybrid project." & vbCrLf)
                                Else
                                    Select Case ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Type>.Value
                                        Case "Directory"
                                            SettingsLocn.Type = FileLocation.Types.Directory
                                            If ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = Nothing Then
                                                RaiseEvent ErrorMessage("No Settings relative location path is specified for the Hybrid project." & vbCrLf)
                                            Else
                                                If ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = "" Then
                                                    SettingsLocn.Path = Path 'The Settings Path is the same as the Project Directory
                                                Else
                                                    SettingsLocn.Path = Path & "\" & ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value 'The Settings Path is a sub-directory of the Project Directory.
                                                End If
                                            End If

                                        Case "Archive"
                                            SettingsLocn.Type = FileLocation.Types.Archive
                                            If ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = Nothing Then
                                                RaiseEvent ErrorMessage("No Settings relative location path is specified for the Hybrid project." & vbCrLf)
                                            Else
                                                If ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value = "" Then
                                                    SettingsLocn.Path = Path 'The Settings Path is the same as the Project Directory
                                                    RaiseEvent ErrorMessage("The Settings relative location path is an Archive with a blank name specified." & vbCrLf)
                                                    RaiseEvent ErrorMessage("The Project path will be used." & vbCrLf)
                                                Else
                                                    SettingsLocn.Path = Path & "\" & ProjectInfoXDoc.<Project>.<SettingsRelativeLocation>.<Path>.Value 'The Settings Path is an Archive in the Project Directory.
                                                End If
                                            End If

                                    End Select
                                End If
                                'A Hybrid project is stored in a directory with the data stored in the directory, a sub-directory or a sub-archive.
                                If ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Type>.Value = Nothing Then
                                    RaiseEvent ErrorMessage("No Data relative location type is specified for the Hybrid project." & vbCrLf)
                                Else
                                    Select Case ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Type>.Value
                                        Case "Directory"
                                            DataLocn.Type = FileLocation.Types.Directory
                                            If ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value = Nothing Then
                                                RaiseEvent ErrorMessage("No Data relative location path is specified for the Hybrid project." & vbCrLf)
                                            Else
                                                If ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value = "" Then
                                                    DataLocn.Path = Path 'The Data Path is the same as the Project Directory
                                                Else
                                                    DataLocn.Path = Path & "\" & ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value 'The Data Path is a sub-directory of the Project Directory.
                                                End If
                                            End If

                                        Case "Archive"
                                            DataLocn.Type = FileLocation.Types.Archive
                                            If ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value = Nothing Then
                                                RaiseEvent ErrorMessage("No Data relative location path is specified for the Hybrid project." & vbCrLf)
                                            Else
                                                If ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value = "" Then
                                                    DataLocn.Path = Path 'The Data Path is the same as the Project Directory
                                                    RaiseEvent ErrorMessage("The Data relative location path is an Archive with a blank name specified." & vbCrLf)
                                                    RaiseEvent ErrorMessage("The Project path will be used." & vbCrLf)
                                                Else
                                                    DataLocn.Path = Path & "\" & ProjectInfoXDoc.<Project>.<DataRelativeLocation>.<Path>.Value 'The Data Path is an Archive in the Project Directory.
                                                End If
                                            End If

                                    End Select
                                End If

                            Case Else
                                Type = Types.None
                                RaiseEvent ErrorMessage("Unknown Project Type: " & ProjectInfoXDoc.<Project>.<Type>.Value & vbCrLf)
                        End Select
                    End If

                    'Read the project Description
                    If ProjectInfoXDoc.<Project>.<Description>.Value = Nothing Then
                        Description = ""
                    Else
                        Description = ProjectInfoXDoc.<Project>.<Description>.Value
                    End If

                    'Read the project Creation Date:
                    If ProjectInfoXDoc.<Project>.<CreationDate>.Value = Nothing Then
                        CreationDate = ""
                    Else
                        CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                    End If

                    'The Project Information File Format Code has already beeen read.

                    'Read the Project Author information:
                    If ProjectInfoXDoc.<Project>.<Author>.<Name>.Value = Nothing Then
                        Author.Name = ""
                    Else
                        Author.Name = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<Author>.<Description>.Value = Nothing Then
                        Author.Description = ""
                    Else
                        Author.Description = ProjectInfoXDoc.<Project>.<Author>.<Description>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<Author>.<Contact>.Value = Nothing Then
                        Author.Contact = ""
                    Else
                        Author.Contact = ProjectInfoXDoc.<Project>.<Author>.<Contact>.Value
                    End If

                    'Read the Application information: 
                    If ProjectInfoXDoc.<Project>.<Application>.<Name>.Value = Nothing Then
                        Application.Name = ""
                    Else
                        Application.Name = ProjectInfoXDoc.<Project>.<Application>.<Name>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<Application>.<Description>.Value = Nothing Then
                        Application.Description = ""
                    Else
                        Application.Description = ProjectInfoXDoc.<Project>.<Application>.<Description>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<CreationDate>.Value = Nothing Then
                        Application.CreationDate = "1-Jan-2000 12:00:00"
                    Else
                        Application.CreationDate = ProjectInfoXDoc.<Project>.<Application>.<CreationDate>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Major>.Value = Nothing Then
                        Application.Version.Major = 0
                    Else
                        Application.Version.Major = ProjectInfoXDoc.<Project>.<Application>.<Version>.<Major>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Minor>.Value = Nothing Then
                        Application.Version.Minor = 0
                    Else
                        Application.Version.Minor = ProjectInfoXDoc.<Project>.<Application>.<Version>.<Minor>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Build>.Value = Nothing Then
                        Application.Version.Build = 0
                    Else
                        Application.Version.Build = ProjectInfoXDoc.<Project>.<Application>.<Version>.<Build>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Revision>.Value = Nothing Then
                        Application.Version.Revision = 0
                    Else
                        Application.Version.Revision = ProjectInfoXDoc.<Project>.<Application>.<Version>.<Revision>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Name>.Value = Nothing Then
                        Application.Author.Name = ""
                    Else
                        Application.Author.Name = ProjectInfoXDoc.<Project>.<Application>.<Author>.<Name>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Description>.Value = Nothing Then
                        Application.Author.Description = ""
                    Else
                        Application.Author.Description = ProjectInfoXDoc.<Project>.<Application>.<Author>.<Description>.Value
                    End If
                    If ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Contact>.Value = Nothing Then
                        Application.Author.Contact = ""
                    Else
                        Application.Author.Contact = ProjectInfoXDoc.<Project>.<Application>.<Author>.<Contact>.Value
                    End If

                    Usage.RestoreUsageInfo()

                Else
                    RaiseEvent ErrorMessage("Unknown Project Information format: " & ProjectInfoXDoc.<Project>.<FormatCode>.Value & vbCrLf)
                    RaiseEvent ErrorMessage("Run the Format Convert application to update the format." & vbCrLf)
                End If
            End If
        Else 'The Project Information file was not found.
            RaiseEvent ErrorMessage("The ADVL_Project_Info.xml file dwas not found." & vbCrLf)
        End If

    End Sub

    'Public Sub ReadProjectInfoFile_Old()
    '    'Read project settings in a project information file.
    '    'This file is located in the project Settings Location.
    '    'The project settings location is initially obtained from the current project file or from the project list.

    '    'UPDATE 13AUG16
    '    'SettingsLocn and DataLocn is the same as the Project Directory location or the Project Archive location.
    '    'These are now left blank in the ADVL_Project_Info.xml file.
    '    'DataLocn has a value only for Hybrid projects.
    '    'Future project types may use these location settings or add new location settings.

    '    'UPDATE 19JUL18
    '    'The ADVL_Project_Info.xml file is now stored in the project directory or archive file specified by Project.Type and Project.Path.

    '    'If Project.Path = "" Then the project is an old pre-19JUL18 version.

    '    If Path = "" Then 'Pre-19JUL18 Project version.

    '    Else 'Post-19JUL18 Project version.


    '    End If

    '    'If SettingsFileExists("ADVL_Project_Info.xml") Then
    '    If ProjectFileExists("ADVL_Project_Info.xml") Then
    '        Dim ProjectInfoXDoc As System.Xml.Linq.XDocument
    '        'ReadXmlSettings("ADVL_Project_Info.xml", ProjectInfoXDoc)
    '        ReadXmlProjectFile("ADVL_Project_Info.xml", ProjectInfoXDoc)

    '        'NOTE: saved values (SavedSettingsLocn etc) stored the values recorded in the ADVL_Project_Info.xml file.
    '        'These were used to handle the case where the project directory or archive has been moved and the saved values need to be updated.
    '        'These values are no longer stored for Directory and Archive project types. (The values are the same as the location of the Directory or Archive file.)

    '        'Read the project Name:
    '        'Dim SavedName As String = ""
    '        If ProjectInfoXDoc.<Project>.<Name>.Value = Nothing Then
    '            Name = ""
    '            'SavedName = ""
    '        Else
    '            Name = ProjectInfoXDoc.<Project>.<Name>.Value
    '            'SavedName = ProjectInfoXDoc.<Project>.<Name>.Value
    '        End If

    '        'If SavedName <> Name Then
    '        '    RaiseEvent ErrorMessage("The last used project name (" & Name & ") is different from the name in the Project Info file (" & SavedName & ")." & vbCrLf)
    '        'End If

    '        'Read the project Description
    '        'Dim SavedDescription As String = ""
    '        If ProjectInfoXDoc.<Project>.<Description>.Value = Nothing Then
    '            Description = ""
    '            'SavedDescription = ""
    '        Else
    '            Description = ProjectInfoXDoc.<Project>.<Description>.Value
    '            'SavedDescription = ProjectInfoXDoc.<Project>.<Description>.Value
    '        End If

    '        'If SavedDescription <> Description Then
    '        '    RaiseEvent ErrorMessage("The last used project description (" & Description & ") is different from the description in the Project Info file (" & SavedDescription & ")." & vbCrLf)
    '        'End If

    '        'Read the project Creation Date:
    '        If ProjectInfoXDoc.<Project>.<CreationDate>.Value = Nothing Then
    '            CreationDate = ""
    '        Else
    '            CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
    '        End If

    '        'Read the project version:
    '        If ProjectInfoXDoc.<Project>.<Version>.<Major>.Value = Nothing Then
    '            Version.Major = 0
    '        Else
    '            Version.Major = ProjectInfoXDoc.<Project>.<Version>.<Major>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<Version>.<Minor>.Value = Nothing Then
    '            Version.Minor = 0
    '        Else
    '            Version.Minor = ProjectInfoXDoc.<Project>.<Version>.<Minor>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<Version>.<Build>.Value = Nothing Then
    '            Version.Build = 0
    '        Else
    '            Version.Build = ProjectInfoXDoc.<Project>.<Version>.<Build>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<Version>.<Revision>.Value = Nothing Then
    '            Version.Revision = 0
    '        Else
    '            Version.Revision = ProjectInfoXDoc.<Project>.<Version>.<Revision>.Value
    '        End If

    '        'Read the Project Author information:
    '        If ProjectInfoXDoc.<Project>.<Author>.<Name>.Value = Nothing Then
    '            Author.Name = ""
    '        Else
    '            Author.Name = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<Author>.<Description>.Value = Nothing Then
    '            Author.Description = ""
    '        Else
    '            Author.Description = ProjectInfoXDoc.<Project>.<Author>.<Description>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<Author>.<Contact>.Value = Nothing Then
    '            Author.Contact = ""
    '        Else
    '            Author.Contact = ProjectInfoXDoc.<Project>.<Author>.<Contact>.Value
    '        End If

    '        'Read the project Type
    '        If ProjectInfoXDoc.<Project>.<Type>.Value = Nothing Then
    '            'Author = ""
    '            Type = Types.None
    '        Else
    '            Select Case ProjectInfoXDoc.<Project>.<Type>.Value
    '                Case "None"
    '                    Type = Types.None
    '                    DataLocn.Type = SettingsLocn.Type
    '                    DataLocn.Path = SettingsLocn.Path
    '                Case "Directory"
    '                    Type = Types.Directory
    '                    DataLocn.Type = SettingsLocn.Type
    '                    DataLocn.Path = SettingsLocn.Path
    '                Case "Archive"
    '                    Type = Types.Archive
    '                    DataLocn.Type = SettingsLocn.Type
    '                    DataLocn.Path = SettingsLocn.Path
    '                Case "Hybrid"
    '                    Type = Types.Hybrid
    '                    Select Case ProjectInfoXDoc.<Project>.<DataLocation>.<Type>.Value
    '                        Case "Directory"
    '                            DataLocn.Type = FileLocation.Types.Directory
    '                        Case "Archive"
    '                            DataLocn.Type = FileLocation.Types.Archive
    '                    End Select
    '                    DataLocn.Path = DataLocn.Path = ProjectInfoXDoc.<Project>.<DataLocation>.<Path>.Value
    '            End Select

    '        End If

    '        ''If the Default project is selected, the Type should be "None". Change the Type to "None" if required:
    '        'If Name = "Default" Then
    '        '    If Type = Types.None Then
    '        '        'Correct Type for the Default project.
    '        '    Else
    '        '        RaiseEvent ErrorMessage("The Default project is selected. The saved project type was: " & Type.ToString & " The project type has been changed to 'None'" & vbCrLf)
    '        '        Type = Types.None
    '        '    End If
    '        'End If

    '        ''Read the Saved SettingsLocn.Type but dont overwite the current value.
    '        'Dim SavedSettingsLocn As New FileLocation

    '        'If ProjectInfoXDoc.<Project>.<SettingsLocation>.<Type>.Value = Nothing Then
    '        '    'SettingsLocn.Type = FileLocation.Types.Directory 
    '        'Else
    '        '    Select Case ProjectInfoXDoc.<Project>.<SettingsLocation>.<Type>.Value
    '        '        Case "Directory"
    '        '            'SettingsLocn.Type = FileLocation.Types.Directory
    '        '            SavedSettingsLocn.Type = FileLocation.Types.Directory
    '        '        Case "Archive"
    '        '            'SettingsLocn.Type = FileLocation.Types.Archive
    '        '            SavedSettingsLocn.Type = FileLocation.Types.Archive
    '        '    End Select
    '        'End If

    '        'If SavedSettingsLocn.Type <> SettingsLocn.Type Then
    '        '    RaiseEvent ErrorMessage("The last used project SettingsLocn.Type (" & SettingsLocn.Type.ToString & ") is different from the type in the Project Info file (" & SavedSettingsLocn.Type.ToString & ")." & vbCrLf)
    '        'End If

    '        ''Read the Saved SettingsLocn.Path but dont overwite the current value.
    '        'If ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value = Nothing Then
    '        '    'SettingsLocn.Path = ApplicationDir
    '        'Else
    '        '    'SettingsLocn.Path = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value
    '        '    SavedSettingsLocn.Path = ProjectInfoXDoc.<Project>.<SettingsLocation>.<Path>.Value
    '        'End If


    '        'If SavedSettingsLocn.Path <> SettingsLocn.Path Then
    '        '    RaiseEvent ErrorMessage("The last used project SettingsLocn.Path (" & SettingsLocn.Path & ") is different from the path in the Project Info file (" & SavedSettingsLocn.Path & ")." & vbCrLf)
    '        'End If

    '        'If ProjectInfoXDoc.<Project>.<DataLocation>.<Type>.Value = Nothing Then
    '        '    DataLocn.Type = FileLocation.Types.Directory
    '        'Else
    '        '    Select Case ProjectInfoXDoc.<Project>.<DataLocation>.<Type>.Value
    '        '        Case "Directory"
    '        '            DataLocn.Type = FileLocation.Types.Directory
    '        '        Case "Archive"
    '        '            DataLocn.Type = FileLocation.Types.Archive
    '        '    End Select
    '        'End If
    '        'If ProjectInfoXDoc.<Project>.<DataLocation>.<Path>.Value = Nothing Then
    '        '    DataLocn.Path = ApplicationDir
    '        'Else
    '        '    DataLocn.Path = ProjectInfoXDoc.<Project>.<DataLocation>.<Path>.Value
    '        'End If

    '        ''If the Default project is being used, change the DataLocn.Type to Directory and DataLocn.Path to Application
    '        'If Name = "Default" Then
    '        '    If DataLocn.Type = FileLocation.Types.Directory Then
    '        '        'Correct DataLocn.Type for the Default project read.
    '        '    Else
    '        '        RaiseEvent ErrorMessage("The Default project is selected. The saved DataLocn.Type was: " & DataLocn.Type.ToString & " The project DataLocn.Type has been changed to 'Directory'" & vbCrLf)
    '        '        DataLocn.Type = FileLocation.Types.Directory
    '        '    End If
    '        '    If DataLocn.Path = ApplicationDir & "\" & "Default_Project" Then
    '        '        'Correct DataLocn.Path for the Default project read.
    '        '    Else
    '        '        RaiseEvent ErrorMessage("The Default project is selected. The saved DataLocn.Path was: " & DataLocn.Path & " The project DataLocn.Type has been changed to " & ApplicationDir & "\" & "Default_Project" & vbCrLf)
    '        '        DataLocn.Path = ApplicationDir & "\" & "Default_Project"
    '        '    End If
    '        'End If





    '        'Read the Application Summary NOTE: THIS CODE REMAINS TO HANDLE OLD PROJECT VERSIONS. IN NEW VERSIONS, <ApplicationSummary> IS REPLACED BY <HostApplication>.
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Name>.Value = Nothing Then
    '            'ApplicationSummary.Name = ""
    '            'HostApplication.Name = ""
    '            Application.Name = ""
    '        Else
    '            'ApplicationSummary.Name = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Name>.Value
    '            'HostApplication.Name = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Name>.Value
    '            Application.Name = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Name>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Description>.Value = Nothing Then
    '            'ApplicationSummary.Description = ""
    '            'HostApplication.Description = ""
    '            Application.Description = ""
    '        Else
    '            'ApplicationSummary.Description = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Description>.Value
    '            'HostApplication.Description = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Description>.Value
    '            Application.Description = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Description>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<CreationDate>.Value = Nothing Then
    '            'ApplicationSummary.CreationDate = ""
    '            'HostApplication.CreationDate = "1-Jan-2000 12:00:00"
    '            Application.CreationDate = "1-Jan-2000 12:00:00"

    '        Else
    '            'ApplicationSummary.CreationDate = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<CreationDate>.Value
    '            'HostApplication.CreationDate = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<CreationDate>.Value
    '            Application.CreationDate = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<CreationDate>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Major>.Value = Nothing Then
    '            'ApplicationSummary.Version.Major = ""
    '            'HostApplication.Version.Major = ""
    '            'HostApplication.Version.Major = 0
    '            Application.Version.Major = 0
    '        Else
    '            'ApplicationSummary.Version.Major = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Major>.Value
    '            'HostApplication.Version.Major = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Major>.Value
    '            Application.Version.Major = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Major>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Minor>.Value = Nothing Then
    '            'ApplicationSummary.Version.Minor = ""
    '            'HostApplication.Version.Minor = ""
    '            'HostApplication.Version.Minor = 0
    '            Application.Version.Minor = 0
    '        Else
    '            'ApplicationSummary.Version.Minor = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Minor>.Value
    '            'HostApplication.Version.Minor = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Minor>.Value
    '            Application.Version.Minor = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Minor>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Build>.Value = Nothing Then
    '            'ApplicationSummary.Version.Build = ""
    '            'HostApplication.Version.Build = ""
    '            'HostApplication.Version.Build = 0
    '            Application.Version.Build = 0
    '        Else
    '            'ApplicationSummary.Version.Build = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Build>.Value
    '            'HostApplication.Version.Build = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Build>.Value
    '            Application.Version.Build = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Build>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Revision>.Value = Nothing Then
    '            'ApplicationSummary.Version.Revision = ""
    '            'HostApplication.Version.Revision = ""
    '            'HostApplication.Version.Revision = 0
    '            Application.Version.Revision = 0
    '        Else
    '            'ApplicationSummary.Version.Revision = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Revision>.Value
    '            'HostApplication.Version.Revision = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Revision>.Value
    '            Application.Version.Revision = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Version>.<Revision>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Name>.Value = Nothing Then
    '            'ApplicationSummary.Author.Name = ""
    '            'HostApplication.Author.Name = ""
    '            Application.Author.Name = ""
    '        Else
    '            'ApplicationSummary.Author.Name = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Name>.Value
    '            'HostApplication.Author.Name = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Name>.Value
    '            Application.Author.Name = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Name>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Description>.Value = Nothing Then
    '            'ApplicationSummary.Author.Description = ""
    '            'HostApplication.Author.Description = ""
    '            Application.Author.Description = ""
    '        Else
    '            'ApplicationSummary.Author.Description = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Description>.Value
    '            'HostApplication.Author.Description = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Description>.Value
    '            Application.Author.Description = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Description>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Contact>.Value = Nothing Then
    '            'ApplicationSummary.Author.Contact = ""
    '            'HostApplication.Author.Contact = ""
    '            Application.Author.Contact = ""
    '        Else
    '            'ApplicationSummary.Author.Contact = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Contact>.Value
    '            'HostApplication.Author.Contact = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Contact>.Value
    '            Application.Author.Contact = ProjectInfoXDoc.<Project>.<ApplicationSummary>.<Author>.<Contact>.Value
    '        End If

    '        'Read the Host Application information: 
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Name>.Value = Nothing Then
    '            'HostApplication.Name = ""
    '        Else
    '            'HostApplication.Name = ProjectInfoXDoc.<Project>.<HostApplication>.<Name>.Value
    '            Application.Name = ProjectInfoXDoc.<Project>.<HostApplication>.<Name>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Description>.Value = Nothing Then
    '            'HostApplication.Description = ""
    '        Else
    '            'HostApplication.Description = ProjectInfoXDoc.<Project>.<HostApplication>.<Description>.Value
    '            Application.Description = ProjectInfoXDoc.<Project>.<HostApplication>.<Description>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<CreationDate>.Value = Nothing Then
    '            'HostApplication.CreationDate = ""
    '        Else
    '            'HostApplication.CreationDate = ProjectInfoXDoc.<Project>.<HostApplication>.<CreationDate>.Value
    '            Application.CreationDate = ProjectInfoXDoc.<Project>.<HostApplication>.<CreationDate>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Major>.Value = Nothing Then
    '            'HostApplication.Version.Major = ""
    '        Else
    '            'HostApplication.Version.Major = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Major>.Value
    '            Application.Version.Major = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Major>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Minor>.Value = Nothing Then
    '            'HostApplication.Version.Minor = ""
    '        Else
    '            'HostApplication.Version.Minor = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Minor>.Value
    '            Application.Version.Minor = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Minor>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Build>.Value = Nothing Then
    '            'HostApplication.Version.Build = ""
    '        Else
    '            'HostApplication.Version.Build = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Build>.Value
    '            Application.Version.Build = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Build>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Revision>.Value = Nothing Then
    '            'HostApplication.Version.Revision = ""
    '        Else
    '            'HostApplication.Version.Revision = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Revision>.Value
    '            Application.Version.Revision = ProjectInfoXDoc.<Project>.<HostApplication>.<Version>.<Revision>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Name>.Value = Nothing Then
    '            'HostApplication.Author.Name = ""
    '        Else
    '            'HostApplication.Author.Name = ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Name>.Value
    '            Application.Author.Name = ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Name>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Description>.Value = Nothing Then
    '            'HostApplication.Author.Description = ""
    '        Else
    '            'HostApplication.Author.Description = ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Description>.Value
    '            Application.Author.Description = ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Description>.Value
    '        End If
    '        If ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Contact>.Value = Nothing Then
    '            'HostApplication.Author.Contact = ""
    '        Else
    '            'HostApplication.Author.Contact = ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Contact>.Value
    '            Application.Author.Contact = ProjectInfoXDoc.<Project>.<HostApplication>.<Author>.<Contact>.Value
    '        End If

    '        Usage.RestoreUsageInfo()

    '    Else
    '        'No Project_Info.xml file found.
    '        'OpenDefaultProject()
    '        UseDefaultProject()
    '        Usage.RestoreUsageInfo()
    '    End If
    'End Sub

    Public Sub SaveProjectInfoFile()
        'Write the project settings to the project information file.
        'This file is written to the Settings Location.

        Dim ProjectInfoXDoc = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Project Information File-->
                              <!---->
                              <Project>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <Name><%= Name %></Name>
                                  <Type><%= Type %></Type>
                                  <Description><%= Description %></Description>
                                  <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                  <ID><%= ID %></ID>
                                  <Relativepath><%= RelativePath %></Relativepath>
                                  <ParentProject>
                                      <Name><%= ParentProjectName %></Name>
                                      <DirectoryName><%= ParentProjectDirectoryName %></DirectoryName>
                                      <CreationDate><%= Format(ParentProjectCreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                      <ID><%= ParentProjectID %></ID>
                                      <Path><%= ParentProjectPath %></Path>
                                  </ParentProject>
                                  <Author>
                                      <Name><%= Author.Name %></Name>
                                      <Description><%= Author.Description %></Description>
                                      <Contact><%= Author.Contact %></Contact>
                                  </Author>
                                  <SettingsRelativeLocation>
                                      <Type><%= SettingsRelLocn.Type %></Type>
                                      <Path><%= SettingsRelLocn.Path %></Path>
                                  </SettingsRelativeLocation>
                                  <DataRelativeLocation>
                                      <Type><%= DataRelLocn.Type %></Type>
                                      <Path><%= DataRelLocn.Path %></Path>
                                  </DataRelativeLocation>
                                  <SystemRelativeLocation>
                                      <Type><%= SystemRelLocn.Type %></Type>
                                      <Path><%= SystemRelLocn.Path %></Path>
                                  </SystemRelativeLocation>
                                  <Application>
                                      <Name><%= Application.Name %></Name>
                                      <Description><%= Application.Description %></Description>
                                      <CreationDate><%= Format(Application.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                      <Version>
                                          <Major><%= Application.Version.Major %></Major>
                                          <Minor><%= Application.Version.Minor %></Minor>
                                          <Build><%= Application.Version.Build %></Build>
                                          <Revision><%= Application.Version.Revision %></Revision>
                                      </Version>
                                      <Author>
                                          <Name><%= Application.Author.Name %></Name>
                                          <Description><%= Application.Author.Description %></Description>
                                          <Contact><%= Application.Author.Contact %></Contact>
                                      </Author>
                                  </Application>
                                  <ConnectOnOpen><%= ConnectOnOpen %></ConnectOnOpen>
                              </Project>

        SaveXmlProjectFile("Project_Info_ADVL_2.xml", ProjectInfoXDoc)

    End Sub

    Public Function SettingsFileExists(ByVal SettingsName As String) As Boolean
        'Returns True if the settings file exists.
        Select Case SettingsLocn.Type
            Case FileLocation.Types.Directory
                'Check if the SettingsName exists in the project directory:
                Return System.IO.File.Exists(SettingsLocn.Path & "\" & SettingsName)
            Case FileLocation.Types.Archive
                'Check if the SettingsName exists in the project archive:
                Dim Zip As New ZipComp
                Zip.ArchivePath = SettingsLocn.Path
                If Zip.ArchiveExists Then
                    Return Zip.EntryExists(SettingsName)
                Else
                    Return False
                End If
        End Select
    End Function

    Public Sub LockSettings()
        'Add a lock file to the settings location to indicate that it is in use.
        If SettingsLocked() Then
            RaiseEvent ErrorMessage("Settings are already locked." & vbCrLf)
        Else
            Select Case SettingsLocn.Type
                Case FileLocation.Types.Directory
                    'Create a lock file:
                    Dim fs = System.IO.File.Create(SettingsLocn.Path & "\" & "Settings.Lock")
                    fs.Dispose() 'This closes the file.
                Case FileLocation.Types.Archive
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = SettingsLocn.Path
                    Zip.AddEntry("Settings.Lock")
            End Select
        End If
    End Sub

    Public Sub UnlockSettings()
        'Remove a lock file from the settings location to indicate that it is no longer in use.
        If SettingsLocked() Then
            Select Case SettingsLocn.Type
                Case FileLocation.Types.Directory
                    System.IO.File.Delete(SettingsLocn.Path & "\" & "Settings.Lock")
                Case FileLocation.Types.Archive
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = SettingsLocn.Path
                    Zip.RemoveEntry("Settings.Lock")
            End Select
        Else
            RaiseEvent ErrorMessage("Settings are not locked." & vbCrLf)
        End If
    End Sub

    Public Function DataFileExists(ByVal DataName As String) As Boolean
        'Returns True if the data file exists.
        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                'Check if the DataName exists in the project directory:
                Return System.IO.File.Exists(DataLocn.Path & "\" & DataName)
            Case FileLocation.Types.Archive
                'Check if the DataName exists in the project archive:
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                If Zip.ArchiveExists Then
                    Return Zip.EntryExists(DataName)
                Else
                    Return False
                End If
        End Select
    End Function

    Public Function DataDirFileExists(ByVal DataName As String) As Boolean
        'Returns True if the data file exists.
        Select Case DataDirLocn.Type
            Case FileLocation.Types.Directory
                'Check if the DataName exists in the project directory:
                Return System.IO.File.Exists(DataDirLocn.Path & "\" & DataName)
            Case FileLocation.Types.Archive
                'Check if the DataName exists in the project archive:
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataDirLocn.Path
                If Zip.ArchiveExists Then
                    Return Zip.EntryExists(DataName)
                Else
                    Return False
                End If
        End Select
    End Function

    Public Function DataFileCreationDate(ByVal DataName As String) As Date
        'Returns the Creation Date of the file if it exists.
        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                'Check if the DataName exists in the project directory:
                If System.IO.File.Exists(DataLocn.Path & "\" & DataName) Then
                    Dim myFileInfo As New System.IO.FileInfo(DataLocn.Path & "\" & DataName)
                    Return myFileInfo.CreationTime
                Else
                    Return DateValue("1900-01-01")
                End If
            Case FileLocation.Types.Archive
                'Check if the DataName exists in the project archive:
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                If Zip.ArchiveExists Then
                    'Return Zip.EntryExists(DataName)
                    If Zip.EntryExists(DataName) Then
                        'Zip files do not the entry creation date!
                        Return DateValue("1900-01-01")
                    Else
                        Return DateValue("1900-01-01")
                    End If
                Else
                    'Return False
                    Return DateValue("1900-01-01")
                End If
        End Select
    End Function

    Public Function DataFileLastEditDate(ByVal DataName) As Date
        'Returns the Last Edit Date of the file if it exists.
        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                'Check if the DataName exists in the project directory:
                'Return System.IO.File.Exists(DataLocn.Path & "\" & DataName)
                If System.IO.File.Exists(DataLocn.Path & "\" & DataName) Then
                    Dim myFileInfo As New System.IO.FileInfo(DataLocn.Path & "\" & DataName)
                    Return myFileInfo.LastWriteTime
                Else
                    Return DateValue("1900-01-01")
                End If
            Case FileLocation.Types.Archive
                'Check if the DataName exists in the project archive:
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                If Zip.ArchiveExists Then
                    'Return Zip.EntryExists(DataName)
                    If Zip.EntryExists(DataName) Then
                        Return Zip.EntryLastEditDate(DataName)
                    Else
                        Return DateValue("1900-01-01")
                    End If
                Else
                    'Return False
                    Return DateValue("1900-01-01")
                End If
        End Select
    End Function

    Public Sub LockData()
        'Add a lock file to the Data location to indicate that it is in use.
        If DataLocked() Then
            RaiseEvent ErrorMessage("Data is already locked." & vbCrLf)
        Else
            Select Case DataLocn.Type
                Case FileLocation.Types.Directory
                    'Create a lock file:
                    Dim fs = System.IO.File.Create(DataLocn.Path & "\" & "Data.Lock")
                    fs.Dispose() 'This closes the file.
                Case FileLocation.Types.Archive
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = DataLocn.Path
                    Zip.AddEntry("Data.Lock")
            End Select
        End If
    End Sub

    Public Sub UnlockData()
        'Remove a lock file from the Data location to indicate that it is no longer in use.
        If DataLocked() Then
            Select Case DataLocn.Type
                Case FileLocation.Types.Directory
                    System.IO.File.Delete(DataLocn.Path & "\" & "Data.Lock")
                Case FileLocation.Types.Archive
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = DataLocn.Path
                    Zip.RemoveEntry("Data.Lock")
            End Select
        Else
            RaiseEvent ErrorMessage("Data is not locked." & vbCrLf)
        End If
    End Sub

    Public Function SystemFileExists(ByVal SystemName As String) As Boolean
        'Returns True if the system file exists.
        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory
                'Check if the SystemName exists in the project directory:
                Return System.IO.File.Exists(SystemLocn.Path & "\" & SystemName)
            Case FileLocation.Types.Archive
                'Check if the SystemName exists in the project archive:
                Dim Zip As New ZipComp
                Zip.ArchivePath = SystemLocn.Path
                If Zip.ArchiveExists Then
                    Return Zip.EntryExists(SystemName)
                Else
                    Return False
                End If
        End Select
    End Function

    Public Sub LockSystem()
        'Add a lock file to the System location to indicate that it is in use.
        If SystemLocked() Then
            RaiseEvent ErrorMessage("System is already locked." & vbCrLf)
        Else
            Select Case SystemLocn.Type
                Case FileLocation.Types.Directory
                    'Create a lock file:
                    Dim fs = System.IO.File.Create(SystemLocn.Path & "\" & "System.Lock")
                    fs.Dispose() 'This closes the file.
                Case FileLocation.Types.Archive
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = SystemLocn.Path
                    Zip.AddEntry("System.Lock")
            End Select
        End If
    End Sub

    Public Sub UnlockSystem()
        'Remove a lock file from the System location to indicate that it is no longer in use.
        If SystemLocked() Then
            Select Case SystemLocn.Type
                Case FileLocation.Types.Directory
                    System.IO.File.Delete(SystemLocn.Path & "\" & "System.Lock")
                Case FileLocation.Types.Archive
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = SystemLocn.Path
                    Zip.RemoveEntry("System.Lock")
            End Select
        Else
            RaiseEvent ErrorMessage("System is not locked." & vbCrLf)
        End If
    End Sub

    Public Function ProjectFileExists(ByVal FileName As String) As Boolean
        'Returns True if the Project File exists.
        'The Project Location is defined by Type and Path.

        Select Case Type
            Case Types.None
                'Note: When the Type is None, the Default project is being used (Type = Directory, Path = Application Directory \ Default_Project
                'Check if the FileName exists in the project directory:
                Return System.IO.File.Exists(Path & "\" & FileName)
            Case Types.Directory
                'Check if the FileName exists in the project directory:
                Return System.IO.File.Exists(Path & "\" & FileName)
            Case Types.Archive
                'Check if the FileName exists in the project archive:
                Dim Zip As New ZipComp
                Zip.ArchivePath = Path
                If Zip.ArchiveExists Then
                    'Return Zip.EntryExists(Name)
                    Return Zip.EntryExists(FileName)
                Else
                    Return False
                End If
            Case Types.Hybrid
                'Note: When the Type is Hybrid, the Project Directory is specified by Type (= Directory) and Path.
                'Check if the FileName exists in the project directory:
                Return System.IO.File.Exists(Path & "\" & FileName)
        End Select

    End Function

    Public Sub LockProject()
        'Add a Project lock file to the settings location to indicate that the project is in use.
        If ProjectLocked() Then
            RaiseEvent ErrorMessage("The project is already locked." & vbCrLf)
        Else
            'Select Case SettingsLocn.Type
            Select Case Type
                'Case FileLocation.Types.Directory
                Case Types.Directory
                    'Create a project lock file:
                    'Dim fs = System.IO.File.Create(SettingsLocn.Path & "\" & "Project.lock")
                    Dim fs = System.IO.File.Create(Path & "\" & "Project.Lock")
                    fs.Dispose() 'This closes the file.
                    'Case FileLocation.Types.Archive
                Case Types.Archive
                    Dim Zip As New ZipComp
                    'Zip.ArchivePath = SettingsLocn.Path
                    Zip.ArchivePath = Path
                    Zip.AddEntry("Project.Lock")
                Case Types.Hybrid
                    'RaiseEvent ErrorMessage("ADD CODE TO CREATE A LOCK FILE TO A HYBRID PROJECT!" & vbCrLf)
                    Dim fs = System.IO.File.Create(Path & "\" & "Project.Lock")
                    fs.Dispose() 'This closes the file.
                Case Types.None
                    'Create a project lock file:
                    Dim fs = System.IO.File.Create(Path & "\" & "Project.Lock")
                    fs.Dispose() 'This closes the file.
            End Select
        End If
    End Sub

    Public Sub UnlockProject()
        'Remove a Project lock file from the settings location to indicate that it is no longer in use.
        If ProjectLocked() Then
            'Select Case SettingsLocn.Type
            Select Case Type
                'Case FileLocation.Types.Directory
                Case Types.Directory
                    'System.IO.File.Delete(SettingsLocn.Path & "\" & "Project.lock")
                    System.IO.File.Delete(Path & "\" & "Project.Lock")
                    'Case FileLocation.Types.Archive
                Case Types.Archive
                    Dim Zip As New ZipComp
                    'Zip.ArchivePath = SettingsLocn.Path
                    Zip.ArchivePath = Path
                    Zip.RemoveEntry("Project.Lock")
                Case Types.Hybrid
                    'RaiseEvent ErrorMessage("ADD CODE TO REMOVE A LOCK FILE FROM A HYBRID PROJECT!" & vbCrLf)
                    System.IO.File.Delete(Path & "\" & "Project.Lock")
                Case Types.None
                    System.IO.File.Delete(Path & "\" & "Project.Lock")
            End Select
        Else
            RaiseEvent ErrorMessage("Project is not locked." & vbCrLf)
        End If
    End Sub

    Public Sub SaveXmlSettings(ByVal SettingsName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Save the XML settings in the settings location.

        Select Case SettingsLocn.Type
            Case FileLocation.Types.Directory
                'Save the settings document in the directory at SettingsLocn.Path
                XmlDoc.Save(SettingsLocn.Path & "\" & SettingsName)
            Case FileLocation.Types.Archive
                'Save the settings document in the archive at SettingsLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = SettingsLocn.Path
                Zip.AddText(SettingsName, XmlDoc.ToString)
        End Select
    End Sub

    Public Sub RenameSettingsFile(ByVal OldFilename, ByVal NewFilename)
        'Rename a settings file.
        Select Case SettingsLocn.Type
            Case FileLocation.Types.Directory
                'Remove any file named NewFileName:
                If System.IO.File.Exists(SettingsLocn.Path & "\" & NewFilename) Then
                    System.IO.File.Delete(SettingsLocn.Path & "\" & NewFilename)
                End If
                If System.IO.File.Exists(SettingsLocn.Path & "\" & OldFilename) Then
                    System.IO.File.Move(SettingsLocn.Path & "\" & OldFilename, SettingsLocn.Path & "\" & NewFilename)
                End If
            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = SettingsLocn.Path
                Zip.RenameEntry(OldFilename, NewFilename)
        End Select
    End Sub

    Public Sub DeleteSettings(ByVal SettingsName As String)
        Select Case SettingsLocn.Type
            Case FileLocation.Types.Directory
                Try
                    System.IO.File.Delete(SettingsLocn.Path & "\" & SettingsName)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error deleting settings file: " & SettingsName & "  Error message: " & ex.Message & vbCrLf)
                End Try
            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = SettingsLocn.Path
                Zip.RemoveEntry(SettingsName)
        End Select
    End Sub

    Public Sub ReadXmlSettings(ByVal SettingsName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Read the XML settings from the settings location.

        Select Case SettingsLocn.Type
            Case FileLocation.Types.Directory
                'Read the settings document in the directory at SettingsLocn.Path
                If System.IO.File.Exists(SettingsLocn.Path & "\" & SettingsName) Then
                    XmlDoc = XDocument.Load(SettingsLocn.Path & "\" & SettingsName)
                Else
                    XmlDoc = Nothing
                End If

            Case FileLocation.Types.Archive
                'Read the settings document in the archive at SettingsLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = SettingsLocn.Path
                If Zip.EntryExists(SettingsName) Then
                    'Debugging:
                    'Debug.Print(Zip.GetText(SettingsName))
                    Dim Temp As String = Zip.GetText(SettingsName)

                    XmlDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText(SettingsName))

                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing
        End Select

    End Sub

    Public Sub ReadXmlProjectFile(ByVal FileName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Read the XML File from the Project location.

        Select Case Type
            Case Types.None
                'Read the Project document in the directory at ProjectLocn.Path
                If System.IO.File.Exists(Path & "\" & FileName) Then
                    XmlDoc = XDocument.Load(Path & "\" & FileName)
                Else
                    XmlDoc = Nothing
                End If

            Case Types.Directory
                'Read the Project document in the directory at ProjectLocn.Path
                If System.IO.File.Exists(Path & "\" & FileName) Then
                    XmlDoc = XDocument.Load(Path & "\" & FileName)
                Else
                    XmlDoc = Nothing
                End If

            Case Types.Archive
                'Read the Project document in the archive at ProjectLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = Path
                If Zip.EntryExists(FileName) Then
                    'Debugging:
                    'Debug.Print(Zip.GetText(SettingsName))
                    'Dim Temp As String = Zip.GetText(FileName)

                    XmlDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText(FileName))
                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing

            Case Types.Hybrid
                'Read the Project document in the directory at ProjectLocn.Path
                If System.IO.File.Exists(Path & "\" & FileName) Then
                    XmlDoc = XDocument.Load(Path & "\" & FileName)
                Else
                    XmlDoc = Nothing
                End If

        End Select

    End Sub

    'End Sub

    'Read and write Data files: -------------------------------------------------------------------------------------
    'ReadXmlData     - Read data into an XDocument
    'ReadXmlDocData  - Read data into an XmlDocument
    'SaveData        - Save data from a Stream into a Data file
    'SaveXmlData     - Save data from an XDocument
    'RenameDataFile  - Rename a data file
    'DeleteData      - Delete a data file
    'SelectDataFile  - Display a list of data files for selection
    'GetDataFileList - Returns a list of data files
    'ReadData        - Read the data in a file into a Stream
    'CopyArchiveDataToProject - Extract a data file from the Data Archive to the Project directory.


    Public Sub ReadXmlData(ByVal DataFileName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Read the XML data from the data location.

        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                'Read the Xml data document in the directory at DataLocn.Path
                If System.IO.File.Exists(DataLocn.Path & "\" & DataFileName) Then
                    Try
                        XmlDoc = XDocument.Load(DataLocn.Path & "\" & DataFileName)
                    Catch ex As Exception
                        RaiseEvent ErrorMessage("Error reading XML file. " & ex.Message & vbCrLf)
                    End Try

                Else
                    XmlDoc = Nothing
                End If

            Case FileLocation.Types.Archive
                'Read the Xml data document in the archive at DataLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                If Zip.EntryExists(DataFileName) Then
                    Dim XmlStr As String = Zip.GetText(DataFileName)
                    Dim BOM As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble)
                    If XmlStr.StartsWith(BOM, StringComparison.Ordinal) Then 'REMOVE THE BOM!!!
                        XmlDoc = XDocument.Parse(XmlStr.Remove(0, BOM.Length)) 'Otherwise this erro is raised: System.Xml.XmlException: 'Data at the root level is invalid. Line 1, position 1.'
                    Else
                        XmlDoc = XDocument.Parse(XmlStr)
                    End If
                    'XmlDoc = XDocument.Parse(XmlStr)
                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing
        End Select
    End Sub

    Public Sub ReadXmlDocData(ByVal DataFileName As String, ByRef XmlDoc As System.Xml.XmlDocument)
        'Version of the ReadXmlData that outputs the data into an XmlDocument (instead of an XDocument).

        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                'Read the Xml data document in the directory at DataLocn.Path
                If System.IO.File.Exists(DataLocn.Path & "\" & DataFileName) Then
                    'XmlDoc = XDocument.Load(DataLocn.Path & "\" & DataFileName)
                    XmlDoc.Load(DataLocn.Path & "\" & DataFileName)
                Else
                    XmlDoc = Nothing
                End If

            Case FileLocation.Types.Archive
                'Read the Xml Ddata document in the archive at DataLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                If Zip.EntryExists(DataFileName) Then
                    'XmlDoc = XDocument.Parse(Zip.GetText(DataFileName))
                    XmlDoc.LoadXml(Zip.GetText(DataFileName))
                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing
        End Select

    End Sub

    Public Sub SaveData(ByVal DataFileName As String, ByRef Stream As IO.Stream)
        'Save the data in the Stream to a file named DataFileName in the data location.

        Select Case DataLocn.Type
            Case FileLocation.Types.Directory

                'Dim fileStream As New IO.FileStream(DataLocn.Path & "\" & DataFileName, IO.FileMode.OpenOrCreate) 'THIS RESULTED IN SOME OLD DATA IN THE ORIGINAL FILE REMAINING.
                Dim fileStream As New IO.FileStream(DataLocn.Path & "\" & DataFileName, IO.FileMode.Create) 'This creates a new file or completely overwrites an existing file.

                If IsNothing(Stream) Then
                    RaiseEvent ErrorMessage("Project.SaveData error: Stream is Nothing." & vbCrLf)
                Else
                    Stream.Position = 0
                    Dim streamReader As New IO.BinaryReader(Stream)
                    fileStream.Write(streamReader.ReadBytes(Stream.Length), 0, Stream.Length)
                End If

                fileStream.Close()

            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                Stream.Position = 0
                Zip.AddData(DataFileName, Stream)
        End Select

    End Sub

    Public Sub SaveXmlData(ByVal DataFileName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Save the XML data in the data location.

        If DataFileName = "" Then
            RaiseEvent ErrorMessage("DataFileName is blank." & vbCrLf)
            Exit Sub
        End If

        If IsNothing(XmlDoc) Then
            RaiseEvent ErrorMessage("Error saving XML Data to file: " & DataFileName & vbCrLf)
            RaiseEvent ErrorMessage("  - No XML document to save." & DataFileName & vbCrLf)
        Else
            Select Case DataLocn.Type
                Case FileLocation.Types.Directory
                    Try
                        'Save the data XML document in the directory at DataLocn.Path
                        XmlDoc.Save(DataLocn.Path & "\" & DataFileName)
                    Catch ex As Exception
                        RaiseEvent ErrorMessage("Error saving XML file. " & ex.Message & vbCrLf)
                    End Try

                Case FileLocation.Types.Archive
                    'Save the data XML document in the archive at DataLocn.Path
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = DataLocn.Path
                    'Zip.AddText(DataFileName, XmlDoc.ToString)
                    Zip.AddText(DataFileName, "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & XmlDoc.ToString) '14 Mar 2021 - XML files saved to archives were not including the XML header!

            End Select
        End If
    End Sub

    Public Sub RenameDataFile(ByVal OldFilename, ByVal NewFilename)
        'Rename a data file.
        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                'Remove any file named NewFileName:
                If System.IO.File.Exists(DataLocn.Path & "\" & NewFilename) Then
                    System.IO.File.Delete(DataLocn.Path & "\" & NewFilename)
                End If
                If System.IO.File.Exists(DataLocn.Path & "\" & OldFilename) Then
                    System.IO.File.Move(DataLocn.Path & "\" & OldFilename, DataLocn.Path & "\" & NewFilename)
                End If
            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                Zip.RenameEntry(OldFilename, NewFilename)
        End Select
    End Sub

    Public Sub DeleteData(ByVal DataName As String)
        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                Try
                    System.IO.File.Delete(DataLocn.Path & "\" & DataName)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error deleting data file: " & DataName & "  Error message: " & ex.Message & vbCrLf)
                End Try
            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                Zip.RemoveEntry(DataName)
        End Select
    End Sub

    Public Sub DeleteDataDirFile(ByVal DataName As String)
        'Delete the specified file from the DataDirLocn
        If DataDirLocn.Path = "" Then
            'The path does not exist - nothing to delete.
        ElseIf System.IO.Directory.Exists(DataDirLocn.Path) Then
            'Check if the file exists:
            If System.IO.File.Exists(DataDirLocn.Path & "\" & DataName) Then
                'File found. Delete it:
                Try
                    System.IO.File.Delete(DataDirLocn.Path & "\" & DataName)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error deleting data file: " & DataName & "  Error message: " & ex.Message & vbCrLf)
                End Try
            Else
                'The specified file does not exist.
            End If
        Else
            'The path does not exist - nothing to delete.
        End If
    End Sub

    Public Sub CreateDataDir()
        'Create the Data Directory if it doesnt already exist.
        'The DataDirLocn is a directory location that is used to store data that is not suitable for storing the an archive DataLocn.
        'It is only created when it is needed. Most applications do not need this directory. The Document library needs it to make storage of .xlsx and .pdf files easier. These files are already compressed.
        If DataDirLocn.Path = "" Then
            'This project can not contain a separate Data Directory. (It is probably an archive project.)
        Else
            If System.IO.Directory.Exists(DataDirLocn.Path) Then
                'The Data Directory already exists.
            Else
                System.IO.Directory.CreateDirectory(DataDirLocn.Path)
            End If
        End If
    End Sub

    Public Sub DeleteProjectFile(ByVal FileName As String)
        Select Case LocnType
            Case FileLocation.Types.Directory
                Try
                    System.IO.File.Delete(Path & "\" & FileName)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error deleting project file: " & FileName & "  Error message: " & ex.Message & vbCrLf)
                End Try
            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = Path
                Zip.RemoveEntry(FileName)
        End Select
    End Sub

    Public Function SelectDataFile(ByVal DataFileType As String, ByVal DataFileExtension As String) As String
        'Displays a list of data files with the specified extension for selection.
        'The DataFileType string is used to display a description of the type of data file.
        'Eg: SelectDataFile("Text files", "txt")

        'Select Case Type
        Select Case DataLocn.Type
            'Case Types.Archive
            Case FileLocation.Types.Archive
                'The data files are stored in an archive file.
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                'Zip.SelectFile()
                'Zip.frm
                'Dim OpenFileDialog As New Zip.frmZipSelectFile
                Return Zip.SelectFileModal(DataFileExtension)

                'Case Types.Directory
            Case FileLocation.Types.Directory
                'The data files are stored in a directory.
                Dim OpenFileDialog As New System.Windows.Forms.OpenFileDialog
                OpenFileDialog.InitialDirectory = DataLocn.Path
                OpenFileDialog.Filter = DataFileType & " | *." & DataFileExtension
                If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Return System.IO.Path.GetFileName(OpenFileDialog.FileName)
                Else
                    Return ""
                End If

                'Case Types.None
                '    'The data files are stored in the application directory.
                '    Dim OpenFileDialog As New System.Windows.Forms.OpenFileDialog
                '    OpenFileDialog.InitialDirectory = DataLocn.Path
                '    OpenFileDialog.Filter = DataFileType & " | *." & DataFileExtension
                '    If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                '        Return System.IO.Path.GetFileName(OpenFileDialog.FileName)
                '    Else
                '        Return ""
                '    End If

                'Case Types.Hybrid

        End Select

    End Function

    Public Function SelectDataDirFile(ByVal DataFileType As String, ByVal DataFileExtension As String) As String
        'Displays a list of data files in DataDirLocn with the specified extension for selection.

        'NOTE: The DataDirLocn directory created to store data that is not suitable for a Data archive. (eg. pdf and xlsx files that are already compressed.)
        '      Check that you should be using SelectDataFile instead!

        'The DataFileType string is used to display a description of the type of data file.
        'Eg: SelectDataFile("Text files", "txt")

        Select Case DataDirLocn.Type
            Case FileLocation.Types.Archive
                'The data files are stored in an archive file.
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                Return Zip.SelectFileModal(DataFileExtension)

            Case FileLocation.Types.Directory
                'The data files are stored in a directory.
                Dim OpenFileDialog As New System.Windows.Forms.OpenFileDialog
                OpenFileDialog.InitialDirectory = DataLocn.Path
                OpenFileDialog.Filter = DataFileType & " | *." & DataFileExtension
                If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Return System.IO.Path.GetFileName(OpenFileDialog.FileName)
                Else
                    Return ""
                End If
        End Select
    End Function

    Public Function GetDataFileList(ByVal DataFileExtension As String, ByRef FilenameList As ArrayList)
        'Returns a list of data files with the specified extension.

        If IsNothing(FilenameList) Then
            RaiseEvent ErrorMessage("The filename array list = Nothing" & vbCrLf)
        Else
            'Select Case Type
            Select Case DataLocn.Type
            'Case Types.Archive
                Case FileLocation.Types.Archive
                    'The data files are stored in an archive file.
                    Dim Zip As New ZipComp
                    Zip.ArchivePath = DataLocn.Path

                'Case Types.Directory
                Case FileLocation.Types.Directory
                    'The data files are stored in a directory.
                    FilenameList.Clear()
                    For Each foundFile As String In IO.Directory.GetFiles(DataLocn.Path, "*." & DataFileExtension)
                        'FilenameList.Add(foundFile)
                        FilenameList.Add(IO.Path.GetFileName(foundFile))
                    Next

                    'Case Types.None
                    '    'The data files are stored in the application directory.
                    '    FilenameList.Clear()
                    '    For Each foundFile As String In IO.Directory.GetFiles(DataLocn.Path, "*." & DataFileExtension)
                    '        'FilenameList.Add(foundFile)
                    '        FilenameList.Add(IO.Path.GetFileName(foundFile))
                    '    Next

            End Select
        End If
    End Function

    'Public Function ReadData(ByVal DataName) As String
    '    'Read the data from the data location.

    '    Select Case DataLocn.Type
    '        Case FileLocation.Types.Directory

    '        Case FileLocation.Types.Archive

    '    End Select

    'End Function

    Public Sub ReadData(ByVal DataFileName As String, ByRef Stream As IO.Stream)
        'Read the data in the file named DataFileName in the data location into the Stream.

        Select Case DataLocn.Type
            Case FileLocation.Types.Directory
                Try
                    Dim fileStream As New IO.FileStream(DataLocn.Path & "\" & DataFileName, System.IO.FileMode.Open)

                    Dim myData(fileStream.Length) As Byte

                    fileStream.Read(myData, 0, fileStream.Length) 'Additional information: Buffer cannot be null.

                    Dim streamWriter As New IO.BinaryWriter(Stream)
                    streamWriter.Write(myData)

                    fileStream.Close()

                Catch ex As Exception
                    RaiseEvent ErrorMessage(ex.Message & vbCrLf)
                End Try

            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path
                Stream = Zip.GetData(DataFileName)

        End Select

    End Sub

    Public Sub CopyArchiveDataToProject(ByVal FileName As String)
        'Extract a data file from the Data Archive to the Project directory.
        If DataLocn.Type = FileLocation.Types.Archive Then 'The Data Location is an archive.
            If LocnType = FileLocation.Types.Directory Then 'The Project Location is a directory.
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path 'The path of the Data Archive file.
                Zip.ExtractFileName = FileName 'The name of the data file to be extracted.
                Zip.ExtractFileDirectory = Path 'The data file will be extracted to this directory path.
                Zip.ExtractFile() '
            Else
                RaiseEvent ErrorMessage("The Project Location is not a directory." & vbCrLf)
            End If
        Else
            RaiseEvent ErrorMessage("The Data Location is not an archive." & vbCrLf)
        End If
    End Sub

    Public Sub CopyArchiveDataToProjectDir(ByVal FileName As String, ByVal DirName As String)
        'Extract a data file from the Data Archive to a Project sub-directory.
        'If the DirName sub-directory does not exist, then it is created.
        If DataLocn.Type = FileLocation.Types.Archive Then 'The Data Location is an archive.
            If LocnType = FileLocation.Types.Directory Then 'The Project Location is a directory.
                Dim DestDirPath As String = Path & "\" & DirName
                If System.IO.Directory.Exists(DestDirPath) Then 'The directory exists
                Else 'Create the directory
                    My.Computer.FileSystem.CreateDirectory(DestDirPath)
                End If
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path 'The path of the Data Archive file.
                Zip.ExtractFileName = FileName 'The name of the data file to be extracted.
                Zip.ExtractFileDirectory = DestDirPath 'The data file will be extracted to this directory path.
                Zip.ExtractFile()
            Else
                RaiseEvent ErrorMessage("The Project Location is not a directory." & vbCrLf)
            End If
        Else
            RaiseEvent ErrorMessage("The Data Location is not an archive." & vbCrLf)
        End If
    End Sub

    Public Sub CopyCheckArchiveDataToProjectDir(ByVal FileName As String, ByVal DirName As String)
        'Extract a data file from the Data Archive to a Project sub-directory.
        'Check if the extracted file already exists.
        'If the DirName sub-directory does not exist, then it is created.
        If DataLocn.Type = FileLocation.Types.Archive Then 'The Data Location is an archive.
            If LocnType = FileLocation.Types.Directory Then 'The Project Location is a directory.
                Dim DestDirPath As String = Path & "\" & DirName
                If System.IO.Directory.Exists(DestDirPath) Then 'The directory exists
                    'Check if the file has already been extracted:
                    If System.IO.File.Exists(DestDirPath & "\" & FileName) Then Exit Sub 'Exit if the file has already been extracted.
                Else 'Create the directory
                    My.Computer.FileSystem.CreateDirectory(DestDirPath)
                End If
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path 'The path of the Data Archive file.
                Zip.ExtractFileName = FileName 'The name of the data file to be extracted.
                Zip.ExtractFileDirectory = DestDirPath 'The data file will be extracted to this directory path.
                Zip.ExtractFile()
            Else
                RaiseEvent ErrorMessage("The Project Location is not a directory." & vbCrLf)
            End If
        Else
            RaiseEvent ErrorMessage("The Data Location is not an archive." & vbCrLf)
        End If
    End Sub

    'Read and write System files: -------------------------------------------------------------------------------------
    'ReadXmlSystemFile     - Read data into an XDocument
    'ReadXmlDocSystemFile  - Read data into an XmlDocument
    'SaveSystemFile        - Save data from a Stream into a Data file
    'SaveXmlSystemFile     - Save data from an XDocument
    'RenameSystemFile      - Rename a data file
    'DeleteSystemFile      - Delete a data file
    'SelectSystemFile      - Display a list of data files for selection
    'GetSystemFileList     - Returns a list of data files
    'ReadSystemFile        - Read the data in a file into a Stream

    Public Sub ReadXmlSystemFile(ByVal SystemFileName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Read the XML data from the system location.

        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory
                'Read the Xml data document in the directory at SystemLocn.Path
                If System.IO.File.Exists(SystemLocn.Path & "\" & SystemFileName) Then
                    Try
                        XmlDoc = XDocument.Load(SystemLocn.Path & "\" & SystemFileName)
                    Catch ex As Exception
                        RaiseEvent ErrorMessage("Error reading XML file. " & ex.Message & vbCrLf)
                    End Try

                Else
                    XmlDoc = Nothing
                End If

            Case FileLocation.Types.Archive
                'Read the Xml data document in the archive at SystemLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = SystemLocn.Path
                If Zip.EntryExists(SystemFileName) Then
                    XmlDoc = XDocument.Parse(Zip.GetText(SystemFileName))
                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing
        End Select

    End Sub

    Public Sub ReadXmlDocSystemFile(ByVal SystemFileName As String, ByRef XmlDoc As System.Xml.XmlDocument)
        'Version of the ReadXmlSystemFile that outputs the data into an XmlDocument (instead of and XDocument).

        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory
                'Read the Xml data document in the directory at SystemLocn.Path
                If System.IO.File.Exists(SystemLocn.Path & "\" & SystemFileName) Then
                    'XmlDoc = XDocument.Load(DataLocn.Path & "\" & DataFileName)
                    XmlDoc.Load(SystemLocn.Path & "\" & SystemFileName)
                Else
                    XmlDoc = Nothing
                End If

            Case FileLocation.Types.Archive
                'Read the Xml data document in the archive at SystemLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = SystemLocn.Path
                If Zip.EntryExists(SystemFileName) Then
                    'XmlDoc = XDocument.Parse(Zip.GetText(DataFileName))
                    XmlDoc.LoadXml(Zip.GetText(SystemFileName))
                Else
                    XmlDoc = Nothing
                End If
                Zip = Nothing
        End Select

    End Sub

    Public Sub SaveSystemFile(ByVal SystemFileName As String, ByRef Stream As IO.Stream)
        'Save the data in the Stream to a file named SystemFileName in the System location.

        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory

                'Dim fileStream As New IO.FileStream(DataLocn.Path & "\" & DataFileName, IO.FileMode.OpenOrCreate) 'THIS RESULTED IN SOME OLD DATA IN THE ORIGINAL FILE REMAINING.
                Dim fileStream As New IO.FileStream(SystemLocn.Path & "\" & SystemFileName, IO.FileMode.Create) 'This creates a new file or completely overwrites an existing file.

                If IsNothing(Stream) Then
                    RaiseEvent ErrorMessage("Project.SaveSystemFile error: Stream is Nothing." & vbCrLf)
                Else
                    Stream.Position = 0
                    Dim streamReader As New IO.BinaryReader(Stream)
                    fileStream.Write(streamReader.ReadBytes(Stream.Length), 0, Stream.Length)
                End If

                fileStream.Close()

            Case FileLocation.Types.Archive

        End Select

    End Sub

    Public Sub SaveXmlSystemFile(ByVal SystemFileName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Save the XML data in the System location.

        If SystemFileName = "" Then
            RaiseEvent ErrorMessage("SystemFileName is blank." & vbCrLf)
            Exit Sub
        End If

        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory
                Try
                    'Save the data XML document in the directory at DataLocn.Path
                    XmlDoc.Save(SystemLocn.Path & "\" & SystemFileName)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error saving XML file. " & ex.Message & vbCrLf)
                End Try

            Case FileLocation.Types.Archive
                'Save the data XML document in the archive at SystemLocn.Path
                Dim Zip As New ZipComp
                Zip.ArchivePath = SystemLocn.Path
                Zip.AddText(SystemFileName, XmlDoc.ToString)
        End Select
    End Sub

    Public Sub RenameSystemFile(ByVal OldFilename, ByVal NewFilename)
        'Rename a System file.
        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory
                'Remove any file named NewFileName:
                If System.IO.File.Exists(SystemLocn.Path & "\" & NewFilename) Then
                    System.IO.File.Delete(SystemLocn.Path & "\" & NewFilename)
                End If
                If System.IO.File.Exists(SystemLocn.Path & "\" & OldFilename) Then
                    System.IO.File.Move(SystemLocn.Path & "\" & OldFilename, SystemLocn.Path & "\" & NewFilename)
                End If
            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = SystemLocn.Path
                Zip.RenameEntry(OldFilename, NewFilename)
        End Select
    End Sub

    Public Sub DeleteSystemFile(ByVal SystemFileName As String)
        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory
                Try
                    System.IO.File.Delete(SystemLocn.Path & "\" & SystemFileName)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error deleting System file: " & SystemFileName & "  Error message: " & ex.Message & vbCrLf)
                End Try
            Case FileLocation.Types.Archive
                Dim Zip As New ZipComp
                Zip.ArchivePath = SystemLocn.Path
                Zip.RemoveEntry(SystemFileName)
        End Select
    End Sub

    Public Function SelectSystemFile(ByVal SystemFileType As String, ByVal SystemFileExtension As String) As String
        'Displays a list of System files with the specified extension for selection.
        'The SystemFileType string is used to display a description of the type of System file.
        'Eg: SelectSystemFile("Text files", "txt")

        'Select Case Type
        Select Case SystemLocn.Type
            'Case Types.Archive
            Case FileLocation.Types.Archive
                'The data files are stored in an archive file.
                Dim Zip As New ZipComp
                Zip.ArchivePath = SystemLocn.Path
                'Zip.SelectFile()
                'Zip.frm
                'Dim OpenFileDialog As New Zip.frmZipSelectFile
                Return Zip.SelectFileModal(SystemFileExtension)

                'Case Types.Directory
            Case FileLocation.Types.Directory
                'The data files are stored in a directory.
                Dim OpenFileDialog As New System.Windows.Forms.OpenFileDialog
                OpenFileDialog.InitialDirectory = SystemLocn.Path
                OpenFileDialog.Filter = SystemFileType & " | *." & SystemFileExtension
                If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Return System.IO.Path.GetFileName(OpenFileDialog.FileName)
                Else
                    Return ""
                End If

                'Case Types.None
                '    'The data files are stored in the application directory.
                '    Dim OpenFileDialog As New System.Windows.Forms.OpenFileDialog
                '    OpenFileDialog.InitialDirectory = DataLocn.Path
                '    OpenFileDialog.Filter = DataFileType & " | *." & DataFileExtension
                '    If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                '        Return System.IO.Path.GetFileName(OpenFileDialog.FileName)
                '    Else
                '        Return ""
                '    End If

                'Case Types.Hybrid

        End Select

    End Function

    Public Function GetSystemFileList(ByVal DataFileExtension As String, ByRef FilenameList As ArrayList)
        'Returns a list of data files with the specified extension.

        'Select Case Type
        Select Case SystemLocn.Type
            'Case Types.Archive
            Case FileLocation.Types.Archive
                'The data files are stored in an archive file.
                Dim Zip As New ZipComp
                Zip.ArchivePath = DataLocn.Path

                'Case Types.Directory
            Case FileLocation.Types.Directory
                'The data files are stored in a directory.
                FilenameList.Clear()
                For Each foundFile As String In IO.Directory.GetFiles(DataLocn.Path, "*." & DataFileExtension)
                    'FilenameList.Add(foundFile)
                    FilenameList.Add(IO.Path.GetFileName(foundFile))
                Next

                'Case Types.None
                '    'The data files are stored in the application directory.
                '    FilenameList.Clear()
                '    For Each foundFile As String In IO.Directory.GetFiles(DataLocn.Path, "*." & DataFileExtension)
                '        'FilenameList.Add(foundFile)
                '        FilenameList.Add(IO.Path.GetFileName(foundFile))
                '    Next

        End Select

    End Function

    Public Sub ReadSystemFile(ByVal SystemFileName As String, ByRef Stream As IO.Stream)
        'Read the data in the file named SystemFileName in the System location into the Stream.

        Select Case SystemLocn.Type
            Case FileLocation.Types.Directory
                Try
                    Dim fileStream As New IO.FileStream(SystemLocn.Path & "\" & SystemFileName, System.IO.FileMode.Open)

                    'Dim myData As New Byte()
                    Dim myData(fileStream.Length) As Byte

                    fileStream.Read(myData, 0, fileStream.Length) 'Additional information: Buffer cannot be null.

                    Dim streamWriter As New IO.BinaryWriter(Stream)
                    streamWriter.Write(myData)

                    fileStream.Close()

                    'fileStream3.Read(Stream, 0, fileStream3.Length)

                    'fileStream2.Read(Stream, 0, fileStream.Length)



                    'Dim fileStream As New IO.FileStream(DataLocn.Path & "\" & DataFileName, IO.FileMode.Open)

                    'Dim streamWriter As New IO.BinaryWriter(Stream)

                    'Dim myData As Byte()

                    'fileStream.Read(myData, 0, fileStream.Length) 'ERROR: Additional information: Buffer cannot be null.
                    'streamWriter.Write(myData)

                    'fileStream.Read(Stream, 0, fileStream.Length)
                Catch ex As Exception
                    RaiseEvent ErrorMessage(ex.Message & vbCrLf)
                End Try

            Case FileLocation.Types.Archive

        End Select

    End Sub

    'Read and write Project files: -------------------------------------------------------------------------------------
    'ReadXmlProjectFile     - Read data into an XDocument
    'ReadXmlDocProjectFile  - Read data into an XmlDocument
    'SaveProjectFile        - Save data from a Stream into a Data file
    'SaveXmlProjectFile     - Save data from an XDocument
    'RenameProjectFile      - Rename a data file
    'DeleteProjectFile      - Delete a data file
    'SelectProjectFile      - Display a list of data files for selection
    'GetProjectFileList     - Returns a list of data files
    'ReadProjectFile        - Read the data in a file into a Stream

    Public Sub SaveXmlProjectFile(ByVal ProjectFileName As String, ByRef XmlDoc As System.Xml.Linq.XDocument)
        'Save the XML data in the Project location.

        If ProjectFileName = "" Then
            RaiseEvent ErrorMessage("ProjectFileName is blank." & vbCrLf)
            Exit Sub
        End If

        'Select Case SystemLocn.Type 'As FileLocation.Types
        Select Case LocnType 'As FileLocation.Types
            Case FileLocation.Types.Directory
                Try
                    'Save the XML document in the directory at Path
                    XmlDoc.Save(Path & "\" & ProjectFileName)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error saving XML file. " & ex.Message & vbCrLf)
                End Try

            Case FileLocation.Types.Archive
                'Save the data XML document in the archive at Path
                Dim Zip As New ZipComp
                'Zip.ArchivePath = SystemLocn.Path
                Zip.ArchivePath = Path
                Zip.AddText(ProjectFileName, XmlDoc.ToString)
        End Select
    End Sub


    Public Function OpenProject(ByVal ProjectName As String) As Boolean
        'Open the project with the name ProjectName.

        Dim ProjectFound As Boolean = False

        Dim ProjectListXDoc As System.Xml.Linq.XDocument
        Dim ProjectListFound As Boolean = False

        If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'The latest ADVL_2 format version of the Project 
            'Dim ProjectListXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")
            'ReadProjectListAdvl_2(ProjectListXDoc)
            ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")
            ProjectListFound = True

        ElseIf System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'The original ADVL_1 format version of the Project 
            RaiseEvent Message("Converting Project_List.xml to Project_List_ADVL_2.xml." & vbCrLf)
            'Convert the file to the latest ADVL_2 format:
            Dim Conversion As New ADVL_Utilities_Library_1.FormatConvert.ProjectListFileConversion
            Conversion.DirectoryPath = ApplicationDir
            Conversion.InputFileName = "Project_List.xml"
            Conversion.InputFormatCode = FormatConvert.ProjectListFileConversion.FormatCodes.ADVL_1
            Conversion.OutputFormatCode = FormatConvert.ProjectListFileConversion.FormatCodes.ADVL_2
            Conversion.Convert()
            If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then
                'ReadProjectList() 'Try ReadProjectList again. This time Project_List_ADVL_2.xml should be found
                ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")
                ProjectListFound = True
            Else
                RaiseEvent ErrorMessage("Error converting Project_List.xml to Project_List_ADVL_2.xml." & vbCrLf)
                ProjectListFound = False
            End If
        Else
            RaiseEvent ErrorMessage("No versions of the Project List were found." & vbCrLf)
            ProjectListFound = False
        End If

        If ProjectListFound = True Then
            Dim Projects = From item In ProjectListXDoc.<ProjectList>.<Project>

            For Each item In Projects
                If item.<Name>.Value = ProjectName Then
                    'Project name found in the project list.

                    'Unlock the current project:
                    If ProjectLocked() Then
                        UnlockProject()
                    End If

                    'Get the project settings.
                    'Only need to read the SettingsLocnType and SettingsLocnPath and then run ReadProjectInfoFile()
                    'ReadProjectInfoFile reads all the other project settings.
                    'SettingsLocn.Path = item.<SettingsLocationPath>.Value
                    'ProjectLocn.Path = item.<Path>.Value
                    Path = item.<Path>.Value
                    'Select Case item.<SettingsLocationType>.Value
                    Select Case item.<Type>.Value
                        Case "Directory"
                            'SettingsLocn.Type = FileLocation.Types.Directory
                            Type = Types.Directory
                            'Usage.SaveLocn.Path = SettingsLocn.Path
                            Usage.SaveLocn.Path = Path
                        Case "Archive"
                            'SettingsLocn.Type = FileLocation.Types.Archive
                            Type = Types.Archive
                            'Usage.SaveLocn.Path = SettingsLocn.Path
                            Usage.SaveLocn.Path = Path
                        Case "None"
                            Type = Types.None
                            Usage.SaveLocn.Path = Path
                        Case "Hybrid"
                            Type = Types.Hybrid
                            Usage.SaveLocn.Path = Path
                    End Select
                    ReadProjectInfoFile()
                    'LockProject() 'Lock the project. NOTE: THIS IS LOCKED LATER!
                    ProjectFound = True
                    Exit For
                End If
            Next
            If ProjectFound = False And ProjectName = "Default" Then
                'The Default project is contained in the directory "Default_Project" in the Application Directory.
                If System.IO.Directory.Exists(ApplicationDir & "\" & "Default_Project") Then
                    'Default project directory exists.
                    'SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                    'SettingsLocn.Path = ApplicationDir & "\" & "Default_Project"
                    'ProjectLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                    'Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
                    Type = Types.Directory
                    'ProjectLocn.Path = ApplicationDir & "\" & "Default_Project"
                    Path = ApplicationDir & "\" & "Default_Project"

                    'Check if the Default project is locked:
                    If ProjectLocked() Then
                        RaiseEvent ErrorMessage("The Default project is locked." & vbCrLf)
                        ProjectFound = False
                    Else
                        'Check if the settings file exists:
                        If ProjectFileExists("Project_Info_ADVL_2.xml") Then 'The Project Information file exists (ADVL_2 format version).
                            ReadProjectInfoFile()
                            ProjectFound = True
                        ElseIf ProjectFileExists("ADVL_Project_Info.xml") Then ''The original ADVL_1 format version of the Project Information file exists.
                            ReadProjectInfoFile() 'This will automatically convert the ADVL_Project_Info.xml file to the Project_Info_ADVL_2.xml version.
                            ProjectFound = True
                        Else
                            'The project directory is missing the settings file!
                            ProjectFound = False
                            RaiseEvent ErrorMessage("The Default project directory was found but it is missing the information file." & vbCrLf)
                        End If
                        'xxx
                        'If SettingsFileExists("ADVL_Project_Info.xml") Then
                        '    ReadProjectInfoFile()
                        '    ProjectFound = True
                        'Else
                        '    'The project directory is missing the settings file!
                        '    ProjectFound = False
                        '    RaiseEvent ErrorMessage("The Default project is not in the Project list." & vbCrLf)
                        '    RaiseEvent ErrorMessage("The Default project directory was found but it is missing the settings file." & vbCrLf)
                        'End If
                    End If
                Else
                    'Create the Default project.
                    CreateDefaultProject() 'This also adds the project to the project list.
                    Type = Types.Directory
                    Path = ApplicationDir & "\" & "Default_Project"
                    Usage.SaveLocn.Path = SettingsLocn.Path
                    ReadProjectInfoFile()
                    ProjectFound = True
                End If
            End If
            Return ProjectFound
        Else
            RaiseEvent ErrorMessage("There is no project list." & vbCrLf)
            Return False
        End If


        'ORIGINAL CODE BELOW:
        'If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'Read the Project List.
        '    Dim ProjectListXDoc As System.Xml.Linq.XDocument
        '    ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List.xml")

        '    Dim Projects = From item In ProjectListXDoc.<ProjectList>.<Project>

        '    For Each item In Projects
        '        If item.<Name>.Value = ProjectName Then
        '            'Project name found in the project list.

        '            'Unlock the current project:
        '            If ProjectLocked() Then
        '                UnlockProject()
        '            End If

        '            'Get the project settings.
        '            'Only need to read the SettingsLocnType and SettingsLocnPath and then run ReadProjectInfoFile()
        '            'ReadProjectInfoFile reads all the other project settings.
        '            'SettingsLocn.Path = item.<SettingsLocationPath>.Value
        '            'ProjectLocn.Path = item.<Path>.Value
        '            Path = item.<Path>.Value
        '            'Select Case item.<SettingsLocationType>.Value
        '            Select Case item.<Type>.Value
        '                Case "Directory"
        '                    'SettingsLocn.Type = FileLocation.Types.Directory
        '                    Type = Types.Directory
        '                    'Usage.SaveLocn.Path = SettingsLocn.Path
        '                    Usage.SaveLocn.Path = Path
        '                Case "Archive"
        '                    'SettingsLocn.Type = FileLocation.Types.Archive
        '                    Type = Types.Archive
        '                    'Usage.SaveLocn.Path = SettingsLocn.Path
        '                    Usage.SaveLocn.Path = Path
        '                Case "None"
        '                    Type = Types.None
        '                    Usage.SaveLocn.Path = Path
        '                Case "Hybrid"
        '                    Type = Types.Hybrid
        '                    Usage.SaveLocn.Path = Path
        '            End Select
        '            ReadProjectInfoFile()
        '            LockProject() 'Lock the project
        '            ProjectFound = True
        '            Exit For
        '        End If
        '    Next
        '    If ProjectFound = False And ProjectName = "Default" Then
        '        'The Default project is contained in the directory "Default_Project" in the Application Directory.
        '        If System.IO.Directory.Exists(ApplicationDir & "\" & "Default_Project") Then
        '            'Default project directory exists.
        '            'SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '            'SettingsLocn.Path = ApplicationDir & "\" & "Default_Project"
        '            'ProjectLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '            'Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '            Type = Types.Directory
        '            'ProjectLocn.Path = ApplicationDir & "\" & "Default_Project"
        '            Path = ApplicationDir & "\" & "Default_Project"

        '            'Check if the Default project is locked:
        '            If ProjectLocked() Then
        '                RaiseEvent ErrorMessage("The Default project is locked." & vbCrLf)
        '                ProjectFound = False
        '            Else
        '                'Check if tghe settings file exists:
        '                If SettingsFileExists("ADVL_Project_Info.xml") Then
        '                    ReadProjectInfoFile()
        '                    ProjectFound = True
        '                Else
        '                    'The project directory is missing the settings file!
        '                    ProjectFound = False
        '                    RaiseEvent ErrorMessage("The Default project is not in the Project list." & vbCrLf)
        '                    RaiseEvent ErrorMessage("The Default project directory was found but it is missing the settings file." & vbCrLf)
        '                End If
        '            End If
        '        Else
        '            'Create the Default project.
        '            CreateDefaultProject() 'This also adds the project to the project list.
        '            'SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '            'SettingsLocn.Path = ApplicationDir & "\" & "Default_Project"
        '            'ProjectLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '            'Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '            Type = Types.Directory
        '            'ProjectLocn.Path = ApplicationDir & "\" & "Default_Project"
        '            Path = ApplicationDir & "\" & "Default_Project"
        '            Usage.SaveLocn.Path = SettingsLocn.Path
        '            ReadProjectInfoFile()
        '            ProjectFound = True
        '        End If
        '    End If
        '    Return ProjectFound
        'Else
        '    RaiseEvent ErrorMessage("There is no project list." & vbCrLf)
        '    Return False
        'End If

    End Function

    Public Function OpenProjectPath(ByVal ProjectPath As String) As Boolean
        'Open the project at the location ProjectPath.

        'Dim ProjectFound As Boolean = False

        'If ProjectPath.EndsWith(".AdvlProject") Then 'Archive project.

        'Else 'Default, Directory or Hybrid project. 

        'End If

        Path = ProjectPath
        Return ReadProjectInfoFileAtPath()


    End Function

    Public Sub SelectProject() 'Open the SelectProject form to select a project from the Project List.
        'Show the Project form:
        If IsNothing(ProjectForm) Then
            ProjectForm = New frmProject
            'ProjectForm.ApplicationName = ApplicationName
            'ProjectForm.ApplicationName = Application.Name
            ProjectForm.ApplicationSummary = Application
            'ProjectForm.SettingsLocn = SettingsLocn
            'ProjectForm.ProjectLocn = ProjectLocn
            'ProjectForm.ProjectLocn.Type = Type
            ProjectForm.ProjectLocn.Type = LocnType
            ProjectForm.ProjectLocn.Path = Path
            ProjectForm.ApplicationDir = ApplicationDir
            ProjectForm.Show()
            ProjectForm.RestoreFormSettings()
        Else
            ProjectForm.Show()
            ProjectForm.BringToFront()
        End If
    End Sub

    Public Sub ShowParameters() 'Open the ProjectParams Form an display the Project Parameters.
        'Show the Project Params form:
        If IsNothing(ProjectParamsForm) Then
            ProjectParamsForm = New frmProjectParams
            'ProjectParamsForm.ApplicationName = ApplicationName
            ProjectParamsForm.ApplicationName = Application.Name
            ProjectParamsForm.ProjectLocn.Type = LocnType
            ProjectParamsForm.ProjectLocn.Path = Path
            'ProjectParamsForm.ApplicationDir = ApplicationDir

            'ProjectParamsForm.ParentProjectLocn.Type = FileLocation.Types.Directory
            ProjectParamsForm.ParentProjectLocn.Type = Project.Types.Directory
            ProjectParamsForm.ParentProjectLocn.Path = ParentProjectPath

            ProjectParamsForm.Show()
            ProjectParamsForm.RestoreFormSettings()
            'UPDATED 3Feb19:
            'ProjectParamsForm.ReadProjectParameters()
            ProjectParamsForm.Parameter = Parameter
            ProjectParamsForm.ParentParameter = ParentParameter
            ProjectParamsForm.ShowProjectParameters()
            ProjectParamsForm.DataGridView1.AutoResizeColumns()
        Else
            ProjectParamsForm.Show()
            ProjectParamsForm.BringToFront()
            'UPDATED 3Feb19:
            'ProjectParamsForm.ReadProjectParameters()
            ProjectParamsForm.ShowProjectParameters()
            ProjectParamsForm.DataGridView1.AutoResizeColumns()
        End If

    End Sub

    Public Sub SaveParameters()
        'Save the Project Parameters in the Parameter dictionary in the Project_Params_ADVL_2.xml file.

        Dim ProjectParamsXDoc = <?xml version="1.0" encoding="utf-8"?>
                                <!---->
                                <!--Project Parameter File-->
                                <!---->
                                <ProjectParameterList>
                                    <%= From item In Parameter
                                        Select
                                            <Parameter>
                                                <Name><%= item.Key %></Name>
                                                <Value><%= item.Value.Value %></Value>
                                                <Description><%= item.Value.Description %></Description>
                                            </Parameter>
                                    %>
                                </ProjectParameterList>

        '<%= From item In ProjectParam

        SaveXmlProjectFile("Project_Params_ADVL_2.xml", ProjectParamsXDoc)

    End Sub

    Public Sub ReadParameters()
        'Read the Project Parameters from the Project_Params_ADVL_2.xml file into the Parameter dictionary.

        If ProjectFileExists("Project_Params_ADVL_2.xml") Then  'The Project Parameter file exists (ADVL_2 format version).
            Dim ProjectParamsXDoc As System.Xml.Linq.XDocument
            'ReadXmlProjectFile("Project_Info_ADVL_2.xml", ProjectParamsXDoc)
            ReadXmlProjectFile("Project_Params_ADVL_2.xml", ProjectParamsXDoc)

            Dim Params = From item In ProjectParamsXDoc.<ProjectParameterList>.<Parameter>
            Dim Name As String
            Parameter.Clear()

            For Each item In Params
                Name = item.<Name>.Value
                Parameter.Add(Name, New ParamInfo)
                Parameter(Name).Value = item.<Value>.Value
                Parameter(Name).Description = item.<Description>.Value
                'RaiseEvent Message("Parameter read: " & Name & vbCrLf)
                'Debug.Print("Parameter read: " & Name)
            Next
        End If
    End Sub

    Public Sub ReadParentParameters()
        'Read the Parent Project Parameters into the ParentParameter dictionary.

        'ParentProjectPath contains the directory path of the parent project.
        If System.IO.File.Exists(ParentProjectPath & "\Project_Params_ADVL_2.xml") Then
            Dim ParentProjectParamsXDoc As System.Xml.Linq.XDocument

            ParentProjectParamsXDoc = XDocument.Load(ParentProjectPath & "\Project_Params_ADVL_2.xml")

            Dim Params = From item In ParentProjectParamsXDoc.<ProjectParameterList>.<Parameter>
            Dim Name As String
            ParentParameter.Clear()

            For Each item In Params
                Name = item.<Name>.Value
                ParentParameter.Add(Name, New ParamInfo)
                ParentParameter(Name).Value = item.<Value>.Value
                ParentParameter(Name).Description = item.<Description>.Value
            Next

        Else
            'Parent Project Parameter file not found.
        End If

    End Sub

    Public Sub AddParameter(ByVal Name As String, ByVal Value As String, ByVal Description As String)
        'Add a parameter to the ProjectParam dictionary.

        'If ProjectParam.ContainsKey(Name) Then
        If Parameter.ContainsKey(Name) Then
            Parameter(Name).Value = Value
            Parameter(Name).Description = Description
        Else
            Parameter.Add(Name, New ParamInfo)
            Parameter(Name).Value = Value
            Parameter(Name).Description = Description
        End If

    End Sub

    Public Sub AddParameter(ByVal Name As String, ByVal Value As String)
        'Version of the AddParameter method that does not specify a description.
        'This can be used to update the value of an existing parameter.

        'If ProjectParam.ContainsKey(Name) Then
        If Parameter.ContainsKey(Name) Then
            Parameter(Name).Value = Value
        Else
            Parameter.Add(Name, New ParamInfo)
            Parameter(Name).Value = Value
        End If

    End Sub

    Public Sub RemoveParameter(ByVal Name As String)
        'Remove a parameter to the ProjectParam dictionary.
        If Parameter.ContainsKey(Name) Then
            Parameter.Remove(Name)
        Else
            'Parameter dictionary does not contain the Name Key.
        End If

    End Sub

    Public Function ParameterExists(ByVal Name As String) As Boolean
        'True if a Parameter with the specified Name exists.
        Return Parameter.ContainsKey(Name)
    End Function

    Public Function GetParameter(ByVal Name As String) As String
        'Return the parameter value with the specified name.
        'If there is no parameter with that name, return ""
        If ParameterExists(Name) Then
            Return Parameter(Name).Value
        Else
            Return ""
        End If
    End Function

    Public Function ParentParameterExists(ByVal Name As String) As Boolean
        'True if a Parent Parameter with the specified Name exists.
        Return ParentParameter.ContainsKey(Name)
    End Function

    Public Function GetParentParameter(ByVal Name As String) As String
        'Return the parent parameter value with the specified name.
        'If there is no parameter with that name, return ""
        If ParentParameterExists(Name) Then
            Return ParentParameter(Name).Value
        Else
            Return ""
        End If
    End Function

    'Private Sub ProjectForm_ErrorMessage(Message As String) Handles ProjectForm.ErrorMessage
    Private Sub ProjectForm_ErrorMessage(Msg As String) Handles ProjectForm.ErrorMessage
        RaiseEvent ErrorMessage(Msg)
    End Sub

    'Private Sub ProjectForm_Message(Message As String) Handles ProjectForm.Message
    Private Sub ProjectForm_Message(Msg As String) Handles ProjectForm.Message
        RaiseEvent Message(Msg)
    End Sub

    Private Sub ProjectForm_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles ProjectForm.FormClosed
        ProjectForm = Nothing
    End Sub

    Private Sub ProjectParamsForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ProjectParamsForm.FormClosed
        ProjectParamsForm = Nothing
    End Sub

    Private Sub ProjectForm_ProjectSelected(ByRef ProjectSummary As ProjectSummary) Handles ProjectForm.ProjectSelected
        'A project has been selected.
        RaiseEvent Closing() 'This event indicates that the current project is closing.
        Type = ProjectSummary.Type
        Path = ProjectSummary.Path
        ReadProjectInfoFile()
        RaiseEvent Selected()
    End Sub

    Public Sub CreateDefaultProject()
        'Create the default project.
        If System.IO.Directory.Exists(ApplicationDir & "\" & "Default_Project") Then
            'The default project directory already exists.
            'Add the Default project to the project list if it is not already present.
            AddDefaultProjectToList()
        Else
            System.IO.Directory.CreateDirectory(ApplicationDir & "\" & "Default_Project")

            'Set the project properties --------------------------------------------------------------
            Name = "Default"
            Type = Types.None
            Path = ApplicationDir & "\" & "Default_Project"
            'Type = Types.Directory

            Description = "Default project. Data and settings are stored in the Application Directory."
            CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'ADDED 22Aug18

            'Generate the Project ID:
            'Dim IDString As String = Name & " " & CreationDate
            Dim IDString As String = Name & " " & Format(CreationDate, "d-MMM-yyyy H:mm:ss")
            ID = IDString.GetHashCode

            ''Host Project information code added 22Aug18
            'HostProjectName = ""
            'HostProjectDirectoryName = ""
            'HostProjectCreationDate = "1-Jan-2000 12:00:00"
            'HostProjectID = ""

            'Parent Project information code added 24Sep18
            ParentProjectName = ""
            ParentProjectDirectoryName = ""
            ParentProjectCreationDate = "1-Jan-2000 12:00:00" 'Use this as a default blank value.
            ParentProjectID = ""

            Author.Name = ""
            Author.Description = ""
            Author.Contact = ""
            Type = ADVL_Utilities_Library_1.Project.Types.None
            'Project.CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
            'Project.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            SettingsRelLocn.Type = FileLocation.Types.Directory
            SettingsRelLocn.Path = ""
            SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
            SettingsLocn.Path = ApplicationDir & "\" & "Default_Project"
            DataRelLocn.Type = FileLocation.Types.Directory
            DataRelLocn.Path = ""
            DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
            DataLocn.Path = ApplicationDir & "\" & "Default_Project"

            'Added 3Nov18:
            SystemRelLocn.Type = FileLocation.Types.Directory
            SystemRelLocn.Path = ""
            SystemLocn.Type = FileLocation.Types.Directory
            SystemLocn.Path = ApplicationDir & "\" & "Default_Project"

            Usage.SaveLocn.Type = FileLocation.Types.Directory
            Usage.SaveLocn.Path = ApplicationDir & "\" & "Default_Project"
            Usage.FirstUsed = Format(Now, "d-MMM-yyyy H:mm:ss")

            'Added 3Nov18:
            Usage.LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")

            Usage.StartTime = Format(Now, "d-MMM-yyyy H:mm:ss")
            Usage.SaveUsageInfo()

            'Set the ApplicationInfo properties ------------------------------------------------------
            'Check if the Application Info file exists:
            'If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.xml") Then
            If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info_ADVL_2.xml") Then
                'Read the Application Information:
                'Dim ApplicationInfo As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Application_Info.xml")
                Dim ApplicationInfo As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Application_Info_ADVL_2.xml")

                If ApplicationInfo.<Application>.<Name>.Value = Nothing Then
                    'ApplicationSummary.Name = ""
                    'HostApplication.Name = ""
                    Application.Name = ""
                Else
                    'HostApplication.Name = ApplicationInfo.<Application>.<Name>.Value
                    Application.Name = ApplicationInfo.<Application>.<Name>.Value
                End If
                If ApplicationInfo.<Application>.<Description>.Value = Nothing Then
                    'HostApplication.Description = ""
                    Application.Description = ""
                Else
                    'HostApplication.Description = ApplicationInfo.<Application>.<Description>.Value
                    Application.Description = ApplicationInfo.<Application>.<Description>.Value
                End If
                If ApplicationInfo.<Application>.<CreationDate>.Value = Nothing Then
                    'HostApplication.CreationDate = "1-Jan-2000 12:00:00"
                    Application.CreationDate = "1-Jan-2000 12:00:00"
                Else
                    'HostApplication.CreationDate = ApplicationInfo.<Application>.<CreationDate>.Value
                    Application.CreationDate = ApplicationInfo.<Application>.<CreationDate>.Value
                End If
                If ApplicationInfo.<Application>.<Version>.<Major>.Value = Nothing Then
                    'HostApplication.Version.Major = 1
                    Application.Version.Major = 1
                Else
                    'HostApplication.Version.Major = ApplicationInfo.<Application>.<Version>.<Major>.Value
                    Application.Version.Major = ApplicationInfo.<Application>.<Version>.<Major>.Value
                End If
                If ApplicationInfo.<Application>.<Version>.<Minor>.Value = Nothing Then
                    'HostApplication.Version.Minor = 0
                    Application.Version.Minor = 0
                Else
                    'HostApplication.Version.Minor = ApplicationInfo.<Application>.<Version>.<Minor>.Value
                    Application.Version.Minor = ApplicationInfo.<Application>.<Version>.<Minor>.Value
                End If
                If ApplicationInfo.<Application>.<Version>.<Build>.Value = Nothing Then
                    'HostApplication.Version.Build = 1
                    Application.Version.Build = 1
                Else
                    'HostApplication.Version.Build = ApplicationInfo.<Application>.<Version>.<Build>.Value
                    Application.Version.Build = ApplicationInfo.<Application>.<Version>.<Build>.Value
                End If
                If ApplicationInfo.<Application>.<Version>.<Revision>.Value = Nothing Then
                    'HostApplication.Version.Revision = 0
                    Application.Version.Revision = 0
                Else
                    'HostApplication.Version.Revision = ApplicationInfo.<Application>.<Version>.<Revision>.Value
                    Application.Version.Revision = ApplicationInfo.<Application>.<Version>.<Revision>.Value
                End If
                If ApplicationInfo.<Application>.<Author>.<Name>.Value = Nothing Then
                    'HostApplication.Author.Name = ""
                    Application.Author.Name = ""
                Else
                    'HostApplication.Author.Name = ApplicationInfo.<Application>.<Author>.<Name>.Value
                    Application.Author.Name = ApplicationInfo.<Application>.<Author>.<Name>.Value
                End If
                If ApplicationInfo.<Application>.<Author>.<Description>.Value = Nothing Then
                    'HostApplication.Author.Description = ""
                    Application.Author.Description = ""
                Else
                    'HostApplication.Author.Description = ApplicationInfo.<Application>.<Author>.<Description>.Value
                    Application.Author.Description = ApplicationInfo.<Application>.<Author>.<Description>.Value
                End If
                If ApplicationInfo.<Application>.<Author>.<Contact>.Value = Nothing Then
                    'HostApplication.Author.Contact = ""
                    Application.Author.Contact = ""
                Else
                    'HostApplication.Author.Contact = ApplicationInfo.<Application>.<Author>.<Contact>.Value
                    Application.Author.Contact = ApplicationInfo.<Application>.<Author>.<Contact>.Value
                End If
            Else
                'RaiseEvent CreateProjectError("The Application Information file (Application_Info.xml) is missing from the Application Directory: " & ApplicationDir)
                'Exit Sub
                'HostApplication.Name = ""
                'HostApplication.Description = ""
                'HostApplication.CreationDate = "1-Jan-2000 12:00:00"
                'HostApplication.Version.Major = 1
                'HostApplication.Version.Minor = 0
                'HostApplication.Version.Build = 1
                'HostApplication.Version.Revision = 0
                'HostApplication.Author.Name = ""
                'HostApplication.Author.Description = ""
                'HostApplication.Author.Contact = ""

                RaiseEvent ErrorMessage("The Application_Info_ADVL_2.xml file was not found. Blank values will be shown in the Project Information file." & vbCrLf)

                Application.Name = ""
                Application.Description = ""
                Application.CreationDate = "1-Jan-2000 12:00:00"
                Application.Version.Major = 1
                Application.Version.Minor = 0
                Application.Version.Build = 1
                Application.Version.Revision = 0
                Application.Author.Name = ""
                Application.Author.Description = ""
                Application.Author.Contact = ""
            End If
            SaveProjectInfoFile()

            'Add the project to the project list:
            Dim ProjectSummary As New ProjectSummary
            ProjectSummary.Name = Name
            ProjectSummary.Description = Description
            ProjectSummary.Type = Type
            'ProjectSummary.SettingsLocnType = SettingsLocn.Type
            'ProjectSummary.SettingsLocnPath = SettingsLocn.Path
            'ProjectSummary.Type = ProjectLocn.Type
            ProjectSummary.Type = Type
            'ProjectSummary.Path = ProjectLocn.Path
            ProjectSummary.Path = Path
            ProjectSummary.AuthorName = Author.Name
            'ProjectSummary.ApplicationName = HostApplication.Name 'NO LONGER USED - 29Jul18 - Application Name is stored once in the project list, not in each project entry.

            Dim ProjectList As New List(Of ProjectSummary)

            'If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'Add the Default project to the existing Project List.
            If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'Add the Default project to the existing Project List.
                Dim ProjectListXDoc As System.Xml.Linq.XDocument
                'ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List.xml")
                ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")

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

                    NewProject.Path = item.<Path>.Value
                    NewProject.AuthorName = item.<AuthorName>.Value

                    ProjectList.Add(NewProject)
                Next

                ProjectList.Add(ProjectSummary)

                ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                                  <!---->
                                  <!--Project List File-->
                                  <ProjectList>
                                      <FormatCode>ADVL_2</FormatCode>
                                      <ApplicationName><%= Application.Name %></ApplicationName>
                                      <%= From item In ProjectList
                                          Select
                                      <Project>
                                          <Name><%= item.Name %></Name>
                                          <Description><%= item.Description %></Description>
                                          <CreationDate><%= Format(item.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                          <Type><%= item.Type %></Type>
                                          <Path><%= item.Path %></Path>
                                          <AuthorName><%= item.AuthorName %></AuthorName>
                                      </Project>
                                      %>
                                  </ProjectList>

                '<ApplicationName><%= ApplicationName %></ApplicationName>

                'ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")
                ProjectListXDoc.Save(ApplicationDir & "\Project_List_ADVL_2.xml")

            Else 'Create a Project List containing the Default project.

                Dim ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                                      <!---->
                                      <!--Project List File-->
                                      <ProjectList>
                                          <FormatCode>ADVL_2</FormatCode>
                                          <ApplicationName><%= Application.Name %></ApplicationName>
                                          <Project>
                                              <Name><%= ProjectSummary.Name %></Name>
                                              <Description><%= ProjectSummary.Description %></Description>
                                              <CreationDate><%= Format(ProjectSummary.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                              <Type><%= ProjectSummary.Type %></Type>
                                              <Path><%= ProjectSummary.Path %></Path>
                                              <AuthorName><%= ProjectSummary.AuthorName %></AuthorName>
                                          </Project>
                                      </ProjectList>

                ' <ApplicationName><%= ApplicationName %></ApplicationName>

                'ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")
                ProjectListXDoc.Save(ApplicationDir & "\Project_List_ADVL_2.xml")
            End If
        End If
    End Sub

    Private Sub AddDefaultProjectToList()
        'Add the Default project to the list.

        'Get the Default project information:
        If System.IO.Directory.Exists(ApplicationDir & "\" & "Default_Project") Then 'The default project directory exists.
            If System.IO.File.Exists(ApplicationDir & "\" & "Default_Project\Project_Info_ADVL_2.xml") Then 'The Project information file exists.
                Dim ProjectSummary As New ProjectSummary
                Dim ProjectInfoXDoc As System.Xml.Linq.XDocument
                ProjectInfoXDoc = XDocument.Load(ApplicationDir & "\" & "Default_Project\Project_Info_ADVL_2.xml")

                If ProjectInfoXDoc.<Project>.<Name>.Value = Nothing Then
                    ProjectSummary.Name = ""
                Else
                    ProjectSummary.Name = ProjectInfoXDoc.<Project>.<Name>.Value
                End If

                If ProjectInfoXDoc.<Project>.<Description>.Value = Nothing Then
                    ProjectSummary.Description = ""
                Else
                    ProjectSummary.Description = ProjectInfoXDoc.<Project>.<Description>.Value
                End If

                If ProjectInfoXDoc.<Project>.<CreationDate>.Value = Nothing Then
                    ProjectSummary.CreationDate = "1-Jan-2000 12:00:00"
                Else
                    ProjectSummary.CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                End If

                ProjectSummary.Path = ApplicationDir & "\" & "Default_Project"

                If ProjectInfoXDoc.<Project>.<Author>.<Name>.Value = Nothing Then
                    ProjectSummary.AuthorName = ""
                Else
                    ProjectSummary.AuthorName = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                End If

                ProjectSummary.Type = Types.None 'A Default project always has the Type = None.


                Dim ProjectList As New List(Of ProjectSummary)

                'If System.IO.File.Exists(ApplicationDir & "\Project_List.xml") Then 'Add the Default project to the existing Project List.
                If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'Add the Default project to the existing Project List.
                    Dim ProjectListXDoc As System.Xml.Linq.XDocument
                    'ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List.xml")
                    ProjectListXDoc = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")

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
                        NewProject.Path = item.<Path>.Value
                        NewProject.AuthorName = item.<AuthorName>.Value

                        ProjectList.Add(NewProject)
                    Next

                    ProjectList.Add(ProjectSummary)

                    ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                                      <!---->
                                      <!--Project List File-->
                                      <ProjectList>
                                          <FormatCode>ADVL_2</FormatCode>
                                          <ApplicationName><%= Application.Name %></ApplicationName>
                                          <%= From item In ProjectList
                                              Select
                                      <Project>
                                          <Name><%= item.Name %></Name>
                                          <Description><%= item.Description %></Description>
                                          <CreationDate><%= Format(item.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                          <Type><%= item.Type %></Type>
                                          <Path><%= item.Path %></Path>
                                          <AuthorName><%= item.AuthorName %></AuthorName>
                                      </Project>
                                          %>
                                      </ProjectList>
                    '<ApplicationName><%= ApplicationName %></ApplicationName>

                    'ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")
                    ProjectListXDoc.Save(ApplicationDir & "\Project_List_ADVL_2.xml")

                Else 'Create a Project List containing the Default project.
                    Dim ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                                          <!---->
                                          <!--Project List File-->
                                          <ProjectList>
                                              <FormatCode>ADVL_2</FormatCode>
                                              <ApplicationName><%= Application.Name %></ApplicationName>
                                              <Project>
                                                  <Name><%= ProjectSummary.Name %></Name>
                                                  <Description><%= ProjectSummary.Description %></Description>
                                                  <CreationDate><%= Format(ProjectSummary.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                                  <Type><%= ProjectSummary.Type %></Type>
                                                  <Path><%= ProjectSummary.Path %></Path>
                                                  <AuthorName><%= ProjectSummary.AuthorName %></AuthorName>
                                              </Project>
                                          </ProjectList>

                    '<ApplicationName><%= ApplicationName %></ApplicationName>

                    'ProjectListXDoc.Save(ApplicationDir & "\Project_List.xml")
                    ProjectListXDoc.Save(ApplicationDir & "\Project_List_ADVL_2.xml")
                End If

            Else
                'The Project information file does not exist.

            End If
        Else
            'The Default project directory does not exist.
        End If

    End Sub

    'UPDATE: Instead use: USEDEFAULTPROJECT() !!!
    Public Sub OpenDefaultProject_Old()
        'Open the default project.
        'RaiseEvent ProjectChanging()
        RaiseEvent Closing() 'This event indicates that the current project is closing.

        If System.IO.Directory.Exists(ApplicationDir & "\" & "Default_Project") Then 'Open the default project.
            'SettingsLocn.Type = FileLocation.Types.Directory
            'SettingsLocn.Path = ApplicationDir & "\" & "Default_Project"
            'ProjectLocn.Type = FileLocation.Types.Directory
            'Type = FileLocation.Types.Directory
            Type = Types.Directory
            'ProjectLocn.Path = ApplicationDir & "\" & "Default_Project"
            Path = ApplicationDir & "\" & "Default_Project"
            ReadProjectInfoFile()
            'RaiseEvent ProjectSelected()
            RaiseEvent Selected()
        Else 'Create the default project and open it.
            CreateDefaultProject()
            'SettingsLocn.Type = FileLocation.Types.Directory
            'SettingsLocn.Path = ApplicationDir & "\" & "Default_Project"
            'ProjectLocn.Type = FileLocation.Types.Directory
            Type = Types.Directory
            'ProjectLocn.Path = ApplicationDir & "\" & "Default_Project"
            Path = ApplicationDir & "\" & "Default_Project"
            ReadProjectInfoFile()
            'RaiseEvent ProjectSelected()
            RaiseEvent Selected()
        End If
    End Sub

    Public Sub AddProjectToList(ByVal ProjectPath As String)
        'Add the Project at ProjectPath to the Project List.

        If ProjectPath.EndsWith(".AdvlProject") Then
            'The Project to add is an Archive project.
            'To Do:
        Else
            Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\Project_Info_ADVL_2.xml")
            Dim Summary As New ProjectSummary

            If ProjectInfoXDoc Is Nothing Then
                RaiseEvent ErrorMessage("No project information was found. The project was not added to the list." & vbCrLf)
            Else
                If ProjectInfoXDoc.<Project>.<Application>.<Name>.Value <> Application.Name Then
                    RaiseEvent ErrorMessage("The Project Application Name is: " & ProjectInfoXDoc.<Project>.<Application>.<Name>.Value & vbCrLf)
                    RaiseEvent ErrorMessage("This does not match the current Application Name: " & Application.Name & vbCrLf)
                Else
                    Select Case ProjectInfoXDoc.<Project>.<Type>.Value
                        Case "Directory"
                            Summary.AuthorName = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                            Summary.CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                            Summary.Description = ProjectInfoXDoc.<Project>.<Description>.Value
                            Summary.Name = ProjectInfoXDoc.<Project>.<Name>.Value
                            Summary.Path = ProjectPath
                            Summary.Type = Project.Types.Directory
                            AddProject(Summary)

                        Case "Hybrid"
                            Summary.AuthorName = ProjectInfoXDoc.<Project>.<Author>.<Name>.Value
                            Summary.CreationDate = ProjectInfoXDoc.<Project>.<CreationDate>.Value
                            Summary.Description = ProjectInfoXDoc.<Project>.<Description>.Value
                            Summary.Name = ProjectInfoXDoc.<Project>.<Name>.Value
                            Summary.Path = ProjectPath
                            Summary.Type = Project.Types.Hybrid
                            AddProject(Summary)

                        Case Else

                    End Select
                End If
            End If
        End If
    End Sub

    Private Sub AddProject(ByRef Summary As ADVL_Utilities_Library_1.ProjectSummary)
        'Add the Project summary information to the project list.

        Dim ProjectList As New List(Of ADVL_Utilities_Library_1.ProjectSummary) 'List of projects

        'Read the Project list:
        If System.IO.File.Exists(ApplicationDir & "\Project_List_ADVL_2.xml") Then 'The latest ADVL_2 format version of the Project List file exists.
            Dim ProjectListXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\Project_List_ADVL_2.xml")
            'ReadProjectListAdvl_2(ProjectListXDoc)
            Dim ProjectFound As Boolean = False 'True if the Project to be added was found in the list.
            Dim Projects = From item In ProjectListXDoc.<ProjectList>.<Project>
            For Each item In Projects
                Dim NewProject As New ADVL_Utilities_Library_1.ProjectSummary
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
                If Summary.Path = NewProject.Path Then ProjectFound = True 'The Project to be added is already in the list.
                NewProject.CreationDate = item.<CreationDate>.Value
                NewProject.AuthorName = item.<AuthorName>.Value
                If item.<Status>.Value = Nothing Then
                    'The Project list file records do not contain the Status field.
                Else
                    NewProject.Status = item.<Status>.Value
                End If
                ProjectList.Add(NewProject)
            Next

            If ProjectFound Then
                RaiseEvent ErrorMessage("The project is already in the list." & vbCrLf)
            Else
                'Add the new project to the list:
                ProjectList.Add(Summary)

                'Write the Project list:
                Dim UpdatedProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                                             <!---->
                                             <!--Project List File-->
                                             <ProjectList>
                                                 <FormatCode>ADVL_2</FormatCode>
                                                 <ApplicationName><%= Application.Name %></ApplicationName>
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

                UpdatedProjectListXDoc.Save(ApplicationDir & "\Project_List_ADVL_2.xml")
            End If
        Else
            'Message.AddWarning("The project list was not found. The project was not added." & vbCrLf)
            RaiseEvent ErrorMessage("The project list was not found. The project was not added." & vbCrLf)
            'To Do: Handle case where list not found. Create a new list?
        End If
    End Sub

    Public Sub SelectProject(ByVal ProjectPath As String)
        'Select the project at path ProjectPath.

        If ProjectPath.EndsWith(".AdvlProject") Then
            'The Project to add is an Archive project.
            'To Do:
            Path = ProjectPath
            Type = Types.Archive
            ReadProjectInfoFile()

        Else
            Dim ProjectInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\Project_Info_ADVL_2.xml")
            If ProjectInfoXDoc Is Nothing Then
                RaiseEvent ErrorMessage("No project information was found. The project was not selected." & vbCrLf)
            Else
                Path = ProjectPath
                Select Case ProjectInfoXDoc.<Project>.<Type>.Value
                    Case "Directory"
                        Type = Types.Directory
                        ReadProjectInfoFile()

                    Case "Hybrid"
                        Type = Types.Hybrid
                        ReadProjectInfoFile()

                    Case "None"
                        Type = Types.None
                        ReadProjectInfoFile()

                    Case Else
                        RaiseEvent ErrorMessage("Unknown project type: " & ProjectInfoXDoc.<Project>.<Type>.Value & vbCrLf)

                End Select
            End If
        End If

    End Sub

    Private Sub ProjectForm_OpenDefaultProject() Handles ProjectForm.OpenDefaultProject
        'OpenDefaultProject()
        UseDefaultProject()
    End Sub

    Private Sub ProjectForm_CreateDefaultProject() Handles ProjectForm.CreateDefaultProject
        CreateDefaultProject()
    End Sub


    Private Sub ProjInfoConversion_ErrorMessage(Message As String) Handles ProjInfoConversion.ErrorMessage
        RaiseEvent ErrorMessage(Message)
    End Sub

    Private Sub ProjInfoConversion_Message(Message As String) Handles ProjInfoConversion.Message
        RaiseEvent Message(Message)
    End Sub

    Private Sub Project_Selected() Handles Me.Selected

    End Sub

    Private Sub ProjectForm_NewProjectCreated(ProjectPath As String) Handles ProjectForm.NewProjectCreated
        RaiseEvent NewProjectCreated(ProjectPath) 'Raise an event to indicate that a new project has been created at the specified path.
    End Sub


#End Region 'Project Methods -----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Project Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Event ErrorMessage(ByVal Message As String)
    Event ErrorMessage(ByVal Msg As String)
    'Event Message(ByVal Message As String)
    Event Message(ByVal Msg As String)
    'Event ProjectChanging() 'The project is changing. This event tells the Main form to save the old project parameters.
    Event Closing() 'This event indicates that the current project is closing. The parameters of the current project should be saved.
    'Event ProjectSelected() 'A project has been selected
    Event Selected() 'This event indicates that a new project has been selected. 

    'Public Event NewProjectCreated(ByVal ProjectPath As String) 'Raise an event to indicate that a new project has been created at the specified path.
    Event NewProjectCreated(ByVal ProjectPath As String) 'Raise an event to indicate that a new project has been created at the specified path.


#End Region 'Project Events ------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class 'Project----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Class ParamInfo
    'Stores information about a Project parameter.

    'NOTE: The Name of the parameter is used as a dictionary key and is not included in the Parameter list.

    Private _description As String = "" 'A description of the project parameter.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _value As String = "" 'The value of the project parameter.
    Property Value As String
        Get
            Return _value
        End Get
        Set(value As String)
            _value = value
        End Set
    End Property

End Class

'ProjectSummary stores a summary of an Andorville (TM) Application Project. ------------------------------------------------------------------------------------------------------------------
Public Class ProjectSummary
    'Name
    'Description
    'Type                 (None, Directory, File, Hybrid)
    'Path                 Added 29July18 - Replaces SettingsLocationPath
    'CreationDate
    'SettingsLocationType 'NOT USED - USE TYPE
    'SettingsLocationPath 'NOT USED = USE PATH
    'DataLocationType 'DONT STORE THIS - IT MAY CHANGE 
    'DataLocationPath 'DONT STORE THIS - IT MAY CHANGE 
    'AuthorName
    'ApplicationName  'NOT USED


    Private _name As String = "" 'The name of the current project.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A description of the current project.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _type As Project.Types = Project.Types.None 'The type of project (None, Directory, File, Hybrid).
    Property Type As Project.Types
        Get
            Return _type
        End Get
        Set(value As Project.Types)
            _type = value
        End Set
    End Property

    Private _path As String = "" 'The path to the project directory or archive.
    Property Path As String
        Get
            Return _path
        End Get
        Set(value As String)
            _path = value
        End Set
    End Property

    Private _creationDate As DateTime = "1-Jan-2000 12:00:00" 'The creation date of the current project.
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    'Private _settingsLocnType As FileLocation.Types = FileLocation.Types.Directory 'The location type used to store settings data (Directory or Archive)
    'Property SettingsLocnType As FileLocation.Types
    '    Get
    '        Return _settingsLocnType
    '    End Get
    '    Set(value As FileLocation.Types)
    '        _settingsLocnType = value
    '    End Set
    'End Property

    'Private _settingsLocnPath As String = "" 'The path of the location used to store settings data.
    'Property SettingsLocnPath As String
    '    Get
    '        Return _settingsLocnPath
    '    End Get
    '    Set(value As String)
    '        _settingsLocnPath = value
    '    End Set
    'End Property

    Private _authorName As String = "" 'The name of the author of the project.
    Property AuthorName As String
        Get
            Return _authorName
        End Get
        Set(value As String)
            _authorName = value
        End Set
    End Property

    'Private _applicationName As String = "" 'The name of the application used to create the project.
    'Property ApplicationName As String
    '    Get
    '        Return _applicationName
    '    End Get
    '    Set(value As String)
    '        _applicationName = value
    '    End Set
    'End Property

    Private _status As String = "OK" 'The status of the project: OK if it is on the Project List, Recycled if it is on the Recycled list.
    Property Status As String
        Get
            Return _status
        End Get
        Set(value As String)
            _status = value
        End Set
    End Property

End Class 'ProjectSummary -------------------------------------------------------------------------------------------------------------------------------------------------------------------

'DataFileInfo stores information about an Andorville (TM) Data File. ------------------------------------------------------------------------------------------------------------------------
Public Class DataFileInfo

#Region " Variable Declarations - All the variables used in this class." '-------------------------------------------------------------------------------------------------------------------
    Public ProjectSummary As New ProjectSummary 'Contains a summary of the Andorville Labs Project that contains this Data File.
#End Region


#Region " Properties - Properties used to store data file information." '--------------------------------------------------------------------------------------------------------------------

    Private _dataSetName As String = "" 'The name of the data set contained in the data file.
    Property DataSetName As String
        Get
            Return _dataSetName
        End Get
        Set(value As String)
            _dataSetName = value
        End Set
    End Property

    Private _dataSetDescription As String = "" 'A description of the data set.
    Property DataSetDescription As String
        Get
            Return _dataSetDescription
        End Get
        Set(value As String)
            _dataSetDescription = value
        End Set
    End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "DataFileInfo Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub Save(ByVal DataFilePath As String)
        'Save the Data File Info in the ADVL_DataFile_Info.xml file in the data file.
        'TO DO:

    End Sub

    Public Sub Restore(ByVal DataFilePath As String)
        'Restore the Data File Info from the ADVL_DataFile_Info.xml file in the data file.
        'TO DO:

    End Sub

#End Region 'DataFileInfo Methods -----------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'DataFileInfo ---------------------------------------------------------------------------------------------------------------------------------------------------------------------




'The Message class stores text properties used to display a message in an Andorville (TM) Application.
Public Class Message '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable declarations - All the variables used in this form and this application." '-----------------------------------------------------------------------------------------------

    Public WithEvents MessageForm As frmMessages 'The form used to display messages.

    Public SettingsLocn As ADVL_Utilities_Library_1.FileLocation 'The location used to store settings.
    Public ApplicationName As String 'The name of the application using the message form.

#End Region

#Region " Message Properties - Properties used to display application messages." '------------------------------------------------------------------------------------------------------------

    Private _fontName As String = "Arial" 'The name of the font used to display the message.
    Property FontName As String
        Get
            Return _fontName
        End Get
        Set(value As String)
            _fontName = value
        End Set
    End Property

    Private _fontSize As Single = 10 'The font size used to display the message.
    Property FontSize As Single
        Get
            Return _fontSize
        End Get
        Set(value As Single)
            _fontSize = value
        End Set
    End Property

    Private _fontStyle As System.Drawing.FontStyle = Drawing.FontStyle.Regular 'The font style used to display the message.
    Property FontStyle As System.Drawing.FontStyle 'Bold, Italic, Regular, Strikeout, Underline
        Get
            Return _fontStyle
        End Get
        Set(value As System.Drawing.FontStyle)
            If value = Drawing.FontStyle.Regular Then
                _fontStyle = Drawing.FontStyle.Regular
            Else
                _fontStyle = _fontStyle Or value 'FontStyle is a flags enumeration so Or is used to combine Styles
            End If
        End Set
    End Property

    Private _color As System.Drawing.Color = Drawing.Color.Black
    Property Color As System.Drawing.Color
        Get
            Return _color
        End Get
        Set(value As System.Drawing.Color)
            _color = value
        End Set
    End Property

    Private _showXMessages As Boolean 'If True, XMessages will be displayed on the Messages form.
    Property ShowXMessages As Boolean
        Get
            Return _showXMessages
        End Get
        Set(value As Boolean)
            _showXMessages = value
            MessageForm.chkShowXMessages.Checked = value
        End Set
    End Property

    Private _showSysMessages As Boolean 'If True, System XMessages will be displayed on the Messages form.
    Property ShowSysMessages As Boolean
        Get
            Return _showSysMessages
        End Get
        Set(value As Boolean)
            _showSysMessages = value
            MessageForm.chkShowSysMessages.Checked = value
        End Set
    End Property

#End Region 'Message properties -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Message Methods - Methods used to display application messages." '-----------------------------------------------------------------------------------------------------------------

    Public Sub SetNormalStyle()
        'Set the message text properties for displaying normal messages:
        SetFontName(FontList.Arial)
        Color = Drawing.Color.Black
        FontStyle = Drawing.FontStyle.Regular 'This removes the following font settings: Bold, Italic, Regular, Strikeout, Underline
    End Sub

    Public Sub SetWarningStyle()
        'Set the message text properties for displaying warning messages:
        SetFontName(FontList.Arial)
        Color = Drawing.Color.Red
        FontStyle = Drawing.FontStyle.Regular 'This removes the following font settings: Bold, Italic, Regular, Strikeout, Underline
        FontStyle = Drawing.FontStyle.Bold

    End Sub

    'Public Sub SetBoldStyle()
    '    'Set the message text to Bold:
    '    FontStyle = Drawing.FontStyle.Bold
    'End Sub

    Public Sub SetBoldStyle()
        'Set the message text to Bold:
        FontStyle = FontStyle Or Drawing.FontStyle.Bold
    End Sub

    Public Sub SetNotBoldStyle()
        'Set the message text to Bold:
        'FontStyle = FontStyle Or Not Drawing.FontStyle.Bold
        FontStyle = FontStyle Xor Drawing.FontStyle.Bold
    End Sub

    'Public Sub SetItalicStyle()
    '    'Set the message text to Italic:
    '    FontStyle = Drawing.FontStyle.Italic
    'End Sub

    Public Sub SetItalicStyle()
        'Set the message text to Italic:
        FontStyle = FontStyle Or Drawing.FontStyle.Italic
    End Sub

    Public Sub SetNotItalicStyle()
        'Set the message text to Italic:
        'FontStyle = FontStyle Or Not Drawing.FontStyle.Italic
        FontStyle = FontStyle Xor Drawing.FontStyle.Italic
    End Sub

    'Public Sub SetUnderlineStyle()
    '    'Set the message text to Underline:
    '    FontStyle = Drawing.FontStyle.Underline
    'End Sub

    Public Sub SetUnderlineStyle()
        'Set the message text to Underline:
        FontStyle = FontStyle Or Drawing.FontStyle.Underline
    End Sub

    Public Sub SetNotUnderlineStyle()
        'Set the message text to Underline:
        'FontStyle = FontStyle Or Not Drawing.FontStyle.Underline
        FontStyle = FontStyle Xor Drawing.FontStyle.Underline
    End Sub

    'Public Sub SetStrikeoutStyle()
    '    'Set the message text to Underline:
    '    FontStyle = Drawing.FontStyle.Strikeout
    'End Sub

    Public Sub SetStrikeoutStyle()
        'Set the message text to Underline:
        FontStyle = FontStyle Or Drawing.FontStyle.Strikeout
    End Sub

    Public Sub SetNotStrikeoutStyle()
        'Set the message text to Underline:
        'FontStyle = FontStyle Or Not Drawing.FontStyle.Strikeout
        FontStyle = FontStyle Xor Drawing.FontStyle.Strikeout
    End Sub

    'A list of text fonts. Used to set the MessageFontName property.
    Public Enum FontList
        Arial
        Calibri
        Courier
        Lucida_Sans_Typewriter
        OCR_A_Extended
        Times_New_Roman
        Courier_New
    End Enum

    Public Sub SetFontName(ByVal Name As FontList)
        'Set the name of the font used to display text in the message window.
        Select Case Name
            Case FontList.Arial
                FontName = "Ariel"
            Case FontList.Calibri
                FontName = "Calibri"
            Case FontList.Courier
                FontName = "Courier"
            Case FontList.Lucida_Sans_Typewriter
                FontName = "Lucida Sans Typewriter"
            Case FontList.OCR_A_Extended
                FontName = "OCR A Extended"
            Case FontList.Times_New_Roman
                FontName = "Times New Roman"
            Case FontList.Courier_New
                FontName = "Courier New"
            Case Else
                RaiseEvent ErrorMessage("Unknown font name: " & Name & vbCrLf)
        End Select
    End Sub

    Public Sub Show()
        'Show the Message form:

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            'MessageForm.ProjectLocn = ProjectLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.RestoreFormSettings()
            MessageForm.BringToFront()
        Else
            MessageForm.Show()
            MessageForm.SettingsLocn = SettingsLocn 'Added 18May19
            MessageForm.RestoreFormSettings() 'Added 18May19
            MessageForm.BringToFront()
        End If
    End Sub

    Public Sub Add(ByVal StrMsg As String)
        'Add a message to the Message form:

        If SettingsLocn Is Nothing Then Exit Sub 'This can occur when attempting to write a message before the Message form is set up.


        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            'MessageForm.ProjectLocn = ProjectLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            MessageForm.BringToFront()
        Else
            MessageForm.Show()
            MessageForm.BringToFront()
        End If

        If IsNothing(StrMsg) Then
            Exit Sub
        End If

        ''Old code - uses the rtbMessages richtextbox control on the MessageForm:
        'Dim StrLen As Integer
        'Dim StrStart As Integer

        'StrStart = MessageForm.rtbMessages.TextLength
        'StrLen = StrMsg.Length

        'MessageForm.rtbMessages.AppendText(StrMsg)
        'MessageForm.rtbMessages.Select(StrStart, StrLen)
        'MessageForm.rtbMessages.SelectionColor = Color
        'MessageForm.rtbMessages.SelectionFont = New System.Drawing.Font(FontName, FontSize, FontStyle)

        ''Move to the end of the message
        'MessageForm.rtbMessages.SelectionStart = MessageForm.rtbMessages.Text.Length
        'MessageForm.rtbMessages.SelectionLength = 0
        'MessageForm.rtbMessages.ScrollToCaret()

        'New code - uses the XmlHtmDisplay1:
        'StrStart = MessageForm.XmlHtmDisplay1.TextLength
        'StrLen = StrMsg.Length
        'MessageForm.XmlHtmDisplay1.SelectedRtf = MessageForm.XmlDisplay.TextToRtf(StrMsg, "Normal")

        'MessageForm.XmlHtmDisplay1.SelectedRtf = MessageForm.XmlHtmDisplay1.TextToRtf(StrMsg, "Normal") 'TEST 23May18
        Dim StrLen As Integer
        Dim StrStart As Integer
        StrStart = MessageForm.XmlHtmDisplay1.TextLength
        StrLen = StrMsg.Length
        MessageForm.XmlHtmDisplay1.AppendText(StrMsg)
        MessageForm.XmlHtmDisplay1.Select(StrStart, StrLen)
        MessageForm.XmlHtmDisplay1.SelectionColor = Color
        MessageForm.XmlHtmDisplay1.SelectionFont = New System.Drawing.Font(FontName, FontSize, FontStyle)

        'Move to the end of the message
        MessageForm.XmlHtmDisplay1.SelectionStart = MessageForm.XmlHtmDisplay1.Text.Length
        MessageForm.XmlHtmDisplay1.SelectionLength = 0
        MessageForm.XmlHtmDisplay1.ScrollToCaret()


    End Sub

    Public Sub AddWarning(ByVal StrMsg As String)
        'Add a warning message to the Message form:

        ''Old code - uses the rtbMessages richtextbox control on the MessageForm:
        ''Save the Current settings:
        'Dim OrigFontName As String = FontName
        'Dim OrigColor As System.Drawing.Color = Color
        'Dim OrigFontStyle As System.Drawing.FontStyle = FontStyle

        ''Set the message text properties for displaying warning messages:
        'SetFontName(FontList.Arial)
        'Color = Drawing.Color.Red
        'FontStyle = Drawing.FontStyle.Regular 'This removes the following font settings: Bold, Italic, Regular, Strikeout, Underline
        'FontStyle = Drawing.FontStyle.Bold

        ''Add(StrMsg) 'Add the message.

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            MessageForm.BringToFront()
        Else
            MessageForm.Show()
            MessageForm.BringToFront()
        End If

        If IsNothing(StrMsg) Then
            Exit Sub
        End If
        'Dim StrLen As Integer
        'Dim StrStart As Integer

        'StrStart = MessageForm.rtbMessages.TextLength
        'StrLen = StrMsg.Length

        'MessageForm.rtbMessages.AppendText(StrMsg)
        'MessageForm.rtbMessages.Select(StrStart, StrLen)
        'MessageForm.rtbMessages.SelectionColor = Color
        'MessageForm.rtbMessages.SelectionFont = New System.Drawing.Font(FontName, FontSize, FontStyle)

        ''Move to the end of the message
        'MessageForm.rtbMessages.SelectionStart = MessageForm.rtbMessages.Text.Length
        'MessageForm.rtbMessages.SelectionLength = 0
        'MessageForm.rtbMessages.ScrollToCaret()


        ''Restore the original settings:
        'FontName = OrigFontName
        'Color = OrigColor
        '_fontStyle = OrigFontStyle

        'New code - uses the XmlHtmDisplay1:
        MessageForm.XmlHtmDisplay1.SelectedRtf = MessageForm.XmlHtmDisplay1.TextToRtf(StrMsg, "Warning")
        'Move to the end of the message
        MessageForm.XmlHtmDisplay1.SelectionStart = MessageForm.XmlHtmDisplay1.Text.Length
        MessageForm.XmlHtmDisplay1.SelectionLength = 0
        MessageForm.XmlHtmDisplay1.ScrollToCaret()

    End Sub

    Public Sub AddXml(ByRef XDoc As System.Xml.XmlDocument)
        'Add an XML message to the XMessage tab on the Message form.

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            MessageForm.BringToFront()
        Else
            MessageForm.Show()
            MessageForm.BringToFront()
        End If

        Dim StrLen As Integer
        Dim StrStart As Integer

        StrStart = MessageForm.XmlHtmDisplay1.TextLength
        StrLen = XDoc.InnerXml.Length

        MessageForm.XmlHtmDisplay1.SelectedRtf = MessageForm.XmlHtmDisplay1.XmlToRtf(XDoc, False)

        'Move to the end of the message
        MessageForm.XmlHtmDisplay1.SelectionStart = MessageForm.XmlHtmDisplay1.Text.Length
        MessageForm.XmlHtmDisplay1.SelectionLength = 0
        MessageForm.XmlHtmDisplay1.ScrollToCaret()

    End Sub

    Public Sub AddXml(ByRef XmlText As String)
        'Version of the XAddXml subroutine that processes an XML string instead of an XMLDocument.
        'Add an XML message to the XMessage tab on the Message form.

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        Else
            MessageForm.Show()
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        End If

        Dim StrLen As Integer
        Dim StrStart As Integer

        StrStart = MessageForm.XmlHtmDisplay1.TextLength
        StrLen = XmlText.Length 'NOT SURE IF THIS WILL BE THE CORRECT LENGTH AFTER THE TEXT HAS BEEN CONVERTED TO AN XMLDOCUMENT AND WRITTEN TO THE RICHTEXTBOX!!!

        MessageForm.XmlHtmDisplay1.SelectedRtf = MessageForm.XmlHtmDisplay1.XmlToRtf(XmlText, False)

        'Move to the end of the message
        MessageForm.XmlHtmDisplay1.SelectionStart = MessageForm.XmlHtmDisplay1.Text.Length
        MessageForm.XmlHtmDisplay1.SelectionLength = 0
        MessageForm.XmlHtmDisplay1.ScrollToCaret()

    End Sub

    'Public Sub XAdd(ByVal StrMsg As String)
    '    'Add an XMessage to the Message form.

    '    If IsNothing(MessageForm) Then
    '        MessageForm = New frmMessages
    '        MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
    '        MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
    '        MessageForm.Show()
    '    Else
    '        MessageForm.Show()
    '    End If

    '    Dim StrLen As Integer
    '    Dim StrStart As Integer

    '    StrStart = MessageForm.rtbInstructions.TextLength
    '    StrLen = StrMsg.Length

    '    MessageForm.rtbInstructions.AppendText(StrMsg)
    '    MessageForm.rtbInstructions.Select(StrStart, StrLen)
    '    MessageForm.rtbInstructions.SelectionColor = Color ' Drawing.Color.Black 'Color
    '    MessageForm.rtbInstructions.SelectionFont = New System.Drawing.Font(FontName, FontSize, FontStyle)

    '    'Move to the end of the message
    '    MessageForm.rtbInstructions.SelectionStart = MessageForm.rtbInstructions.Text.Length
    '    MessageForm.rtbInstructions.SelectionLength = 0
    '    MessageForm.rtbInstructions.ScrollToCaret()


    '    StrStart = MessageForm.XmlDisplay.TextLength
    '    StrLen = StrMsg.Length

    '    'MessageForm.XmlDisplay.AppendText(StrMsg)
    '    MessageForm.XmlDisplay.TextToRtf(StrMsg, "Message")


    '    MessageForm.XmlDisplay.Select(StrStart, StrLen)
    '    MessageForm.XmlDisplay.SelectionColor = Color ' Drawing.Color.Black 'Color
    '    MessageForm.XmlDisplay.SelectionFont = New System.Drawing.Font(FontName, FontSize, FontStyle)

    '    'Move to the end of the message
    '    MessageForm.XmlDisplay.SelectionStart = MessageForm.rtbInstructions.Text.Length
    '    MessageForm.XmlDisplay.SelectionLength = 0
    '    MessageForm.XmlDisplay.ScrollToCaret()

    'End Sub

    Public Sub XAddXml(ByRef XDoc As System.Xml.XmlDocument)
        'Add an XML message to the XMessage tab on the Message form.

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        Else
            MessageForm.Show()
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        End If

        Dim StrLen As Integer
        Dim StrStart As Integer

        'StrStart = MessageForm.rtbInstructions.TextLength
        '     StrLen = StrMsg.Length
        StrStart = MessageForm.XmlDisplay.TextLength
        'StrLen = StrMsg.Length
        'StrLen = XDoc.ToString.Length
        StrLen = XDoc.InnerXml.Length

        'Add("XAddXml: StrStart = " & StrStart & vbCrLf)
        ''Add("XAddXml: StrMsg = " & StrMsg & vbCrLf)
        ''Add("XAddXml: XDoc.ToString = " & XDoc.ToString & vbCrLf) 'NOTE: THIS DISPLAYS: System.Xml.XmlDocument
        'Add("XAddXml: XDoc.InnerXml = " & XDoc.InnerXml & vbCrLf)
        'Add("XAddXml: StrLen = " & StrLen & vbCrLf)
        ''Add("XAddXml: TextType = " & TextType & vbCrLf & vbCrLf)

        'FOR DEBUGGING:
        'SHOW THE RTF in the message tab:
        'Add("MessageForm.XmlDisplay.XmlToRtf(XDoc, False) = " & MessageForm.XmlDisplay.XmlToRtf(XDoc, False) & vbCrLf)

        MessageForm.XmlDisplay.SelectedRtf = MessageForm.XmlDisplay.XmlToRtf(XDoc, False)

        'Move to the end of the message
        MessageForm.XmlDisplay.SelectionStart = MessageForm.XmlDisplay.Text.Length
        MessageForm.XmlDisplay.SelectionLength = 0
        MessageForm.XmlDisplay.ScrollToCaret()

    End Sub

    Public Sub XAddXml(ByRef XmlText As String)
        'Version of the XAddXml subroutine that processes an XML string instead of an XMLDocument.
        'Add an XML message to the XMessage tab on the Message form.

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        Else
            MessageForm.Show()
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        End If

        Dim StrLen As Integer
        Dim StrStart As Integer

        'StrStart = MessageForm.rtbInstructions.TextLength
        '     StrLen = StrMsg.Length
        StrStart = MessageForm.XmlDisplay.TextLength
        'StrLen = StrMsg.Length
        'StrLen = XDoc.ToString.Length
        StrLen = XmlText.Length 'NOT SURE IF THIS WILL BE THE CORRECT LENGTH AFTER THE TEXT HAS BEEN CONVERTED TO AN XMLDOCUMENT AND WRITTEN TO THE RICHTEXTBOX!!!

        'Add("XAddXml: StrStart = " & StrStart & vbCrLf)
        ''Add("XAddXml: StrMsg = " & StrMsg & vbCrLf)
        ''Add("XAddXml: StrMsg = " & XDoc.ToString & vbCrLf)
        'Add("XAddXml: StrMsg = " & XmlText & vbCrLf)
        'Add("XAddXml: StrLen = " & StrLen & vbCrLf)
        ''Add("XAddXml: TextType = " & TextType & vbCrLf & vbCrLf)

        'FOR DEBUGGING:
        'SHOW THE RTF in the message tab:
        'Add("MessageForm.XmlDisplay.XmlToRtf(XmlText, False) = " & MessageForm.XmlDisplay.XmlToRtf(XmlText, False) & vbCrLf)

        MessageForm.XmlDisplay.SelectedRtf = MessageForm.XmlDisplay.XmlToRtf(XmlText, False)

        'Move to the end of the message
        MessageForm.XmlDisplay.SelectionStart = MessageForm.XmlDisplay.Text.Length
        MessageForm.XmlDisplay.SelectionLength = 0
        MessageForm.XmlDisplay.ScrollToCaret()


    End Sub

    Public Sub XAddText(ByVal StrMsg As String, ByVal TextType As String)
        'Add a Text message to the XMessage tab on the Message form.

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        Else
            MessageForm.Show()
            'MessageForm.BringToFront() 'To keep the focus on the app form, dont bring the message form to the front!
        End If

        Dim StrLen As Integer
        Dim StrStart As Integer

        'StrStart = MessageForm.rtbInstructions.TextLength
        StrStart = MessageForm.XmlDisplay.TextLength
        StrLen = StrMsg.Length

        'Debug.Print("XAddText: StrStart = " & StrStart)
        'Debug.Print("XAddText: StrMsg = " & StrMsg)
        'Debug.Print("XAddText: StrLen = " & StrLen)
        'Debug.Print("XAddText: TextType = " & TextType)

        'Add("XAddText: StrStart = " & StrStart & vbCrLf)
        'Add("XAddText: StrMsg = " & StrMsg & vbCrLf)
        'Add("XAddText: StrLen = " & StrLen & vbCrLf)
        'Add("XAddText: TextType = " & TextType & vbCrLf & vbCrLf)

        MessageForm.XmlDisplay.SelectedRtf = MessageForm.XmlDisplay.TextToRtf(StrMsg, TextType)

        'Move to the end of the message
        MessageForm.XmlDisplay.SelectionStart = MessageForm.XmlDisplay.Text.Length
        MessageForm.XmlDisplay.SelectionLength = 0
        MessageForm.XmlDisplay.ScrollToCaret()

    End Sub

    Public Sub AddText(ByVal StrMsg As String, ByVal TextType As String)
        'Add a Text message to the Message tab on the Message form.

        If IsNothing(MessageForm) Then
            MessageForm = New frmMessages
            MessageForm.ApplicationName = ApplicationName 'Pass the Application Name to the Message Form.
            MessageForm.SettingsLocn = SettingsLocn       'Pass the Settings Location to the Message Form.
            MessageForm.Show()
            MessageForm.chkShowXMessages.Checked = ShowXMessages
            MessageForm.BringToFront()
        Else
            MessageForm.Show()
            MessageForm.BringToFront()
        End If

        'Dim StrLen As Integer
        'Dim StrStart As Integer

        ''StrStart = MessageForm.rtbInstructions.TextLength
        'StrStart = MessageForm.rtbMessages.TextLength
        'StrLen = StrMsg.Length

        ''Debug.Print("XAddText: StrStart = " & StrStart)
        ''Debug.Print("XAddText: StrMsg = " & StrMsg)
        ''Debug.Print("XAddText: StrLen = " & StrLen)
        ''Debug.Print("XAddText: TextType = " & TextType)

        ''Add("XAddText: StrStart = " & StrStart & vbCrLf)
        ''Add("XAddText: StrMsg = " & StrMsg & vbCrLf)
        ''Add("XAddText: StrLen = " & StrLen & vbCrLf)
        ''Add("XAddText: TextType = " & TextType & vbCrLf & vbCrLf)

        'MessageForm.rtbMessages.SelectedRtf = MessageForm.XmlDisplay.TextToRtf(StrMsg, TextType)

        ''Move to the end of the message
        'MessageForm.rtbMessages.SelectionStart = MessageForm.XmlDisplay.Text.Length
        'MessageForm.rtbMessages.SelectionLength = 0
        'MessageForm.rtbMessages.ScrollToCaret()

        'New code - uses the XmlHtmDisplay1:
        ''If MessageForm.XmlHtmDisplay1.SelectedRtf Is Nothing Then
        'If IsNothing(MessageForm.XmlHtmDisplay1.SelectedRtf) Then
        '    'Move to the end of the message
        '    MessageForm.XmlHtmDisplay1.SelectionStart = MessageForm.XmlHtmDisplay1.Text.Length
        '    MessageForm.XmlHtmDisplay1.SelectionLength = 0
        '    MessageForm.XmlHtmDisplay1.ScrollToCaret()
        'End If
        MessageForm.XmlHtmDisplay1.SelectedRtf = MessageForm.XmlHtmDisplay1.TextToRtf(StrMsg, TextType)
        'Move to the end of the message
        MessageForm.XmlHtmDisplay1.SelectionStart = MessageForm.XmlHtmDisplay1.Text.Length
        MessageForm.XmlHtmDisplay1.SelectionLength = 0
        MessageForm.XmlHtmDisplay1.ScrollToCaret()


    End Sub

    Public Sub ShowTextTypes()
        'Show the list of text types in Settings.TestType

        Dim NTypes As Integer = MessageForm.XmlHtmDisplay1.Settings.TextType.Count

        Dim I As Integer

        Add(vbCrLf & "List of Message text types:" & vbCrLf)

        For I = 0 To NTypes - 1
            AddText(MessageForm.XmlHtmDisplay1.Settings.TextType.Keys(I) & vbCrLf, MessageForm.XmlHtmDisplay1.Settings.TextType.Keys(I))
        Next

    End Sub

    Public Sub XShowTextTypes()
        'Show the list of text types in Settings.TestType

        Dim NTypes As Integer = MessageForm.XmlDisplay.Settings.TextType.Count

        Dim I As Integer

        'Add(vbCrLf & "List of text types:" & vbCrLf)
        XAddText(vbCrLf & "List of XMessage text types:" & vbCrLf, "Normal")


        For I = 0 To NTypes - 1
            'AddText(MessageForm.XmlDisplay.Settings.TextType.Keys(I) & vbCrLf, MessageForm.XmlDisplay.Settings.TextType.Keys(I))
            XAddText(MessageForm.XmlDisplay.Settings.TextType.Keys(I) & vbCrLf, MessageForm.XmlDisplay.Settings.TextType.Keys(I))
        Next

    End Sub

    ''Public Sub XAddXml(ByVal XmlDoc As System.Xml.XmlDocument)
    'Public Sub XmlAdd(ByVal XmlDoc As System.Xml.XmlDocument)
    '    'MessageForm.XmlDisplay.Rtf = MessageForm.XmlDisplay.XmlToRtf(XmlDoc, False)

    '    'MessageForm.XmlDisplay.SelectedRtf = MessageForm.XmlDisplay.XmlToRtf(XmlDoc, False)
    '    MessageForm.XmlDisplay.Rtf = MessageForm.XmlDisplay.XmlToRtf(XmlDoc, False)

    'End Sub

    'Public Sub XmlAdd(ByVal XmlDoc As System.Xml.Linq.XDocument)
    '    'Version of XmlAdd that processes and XDocument instead of an XmlDocument.

    '    Dim XDoc As New System.Xml.XmlDocument
    '    XDoc.LoadXml(XmlDoc.ToString)

    '    MessageForm.XmlDisplay.Rtf = MessageForm.XmlDisplay.XmlToRtf(XDoc, False)

    'End Sub

    Private Sub MessageForm_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles MessageForm.FormClosed
        MessageForm = Nothing
    End Sub

    Private Sub MessageForm_ErrorMessage(Message As String) Handles MessageForm.ErrorMessage
        RaiseEvent ErrorMessage(Message)
    End Sub

    Private Sub MessageForm_Message(Message As String) Handles MessageForm.Message
        RaiseEvent Message(Message)
    End Sub

    Private Sub MessageForm_ShowTextTypes() Handles MessageForm.ShowTextTypes
        ShowTextTypes()
    End Sub

    Private Sub MessageForm_XShowTextTypes() Handles MessageForm.XShowTextTypes
        XShowTextTypes()
    End Sub

    Private Sub MessageForm_ShowXMessages(Show As Boolean) Handles MessageForm.ShowXMessages
        'The ShowXMessages settings has been changed.
        _showXMessages = Show 'Update the property.
        RaiseEvent ShowXMessagesChanged(Show) 'Notify the host application of the change.
    End Sub

    Private Sub MessageForm_ShowSysMessages(Show As Boolean) Handles MessageForm.ShowSysMessages
        'The ShowSysMessages settings has been changed.
        _showSysMessages = Show 'Update the property.
        RaiseEvent ShowSysMessagesChanged(Show) 'Notify the host application of the change.
    End Sub

    'Private Sub MessageForm_SaveFormSettings(FormName As String, ByRef Settings As XDocument) Handles MessageForm.SaveFormSettings
    '    RaiseEvent SaveFormSettings(FormName, Settings)
    'End Sub

#End Region 'Message methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Public Event SaveFormSettings(ByVal FormName As String, ByRef Settings As System.Xml.Linq.XDocument) 'Raise an event to save the form settings. The settings are contained in the Settings XML document.
    Event ErrorMessage(ByVal Message As String)
    Event Message(ByVal Message As String)
    Event ShowXMessagesChanged(ByVal Show As Boolean) 'This event is triggered when the ShowXMessages property has changed.
    Event ShowSysMessagesChanged(ByVal Show As Boolean) 'This event is triggered when the ShowSysMessages property has changed.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'Message------------------------------------------------------------------------------------------------------------------------------------------------------------


'The ApplicationSummary class stores a summary of the Application that created an Andorville (TM) project.
'This information is contained in the ApplicationSummary section within the Project_Info.xml file.

Public Class ApplicationSummary '------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables used in this class." '-------------------------------------------------------------------------------------------------------------------

    Public Version As New Version 'Stores application version information (Major, Minor, Build, Revision)
    Public Author As New Author 'Stores application author information (Name, Description, Contact)

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - Properties used to store application information." '------------------------------------------------------------------------------------------------------------------

    Private _name As String = "" 'The name of the application
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A description of the application.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _creationDate As DateTime = "31-Dec-2015 22:00:00" 'The release date of the application.
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    Private _licenseNotice As String = "" 'The application software notice.
    Property LicenseNotice As String
        Get
            Return _licenseNotice
        End Get
        Set(value As String)
            _licenseNotice = value
        End Set
    End Property

    Private _copyrightNotice As String = "" 'The application copyright notice.
    Property CopyrightNotice As String
        Get
            Return _copyrightNotice
        End Get
        Set(value As String)
            _copyrightNotice = value
        End Set
    End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'ProjectApplicationInfo data is included in the ProjectInfo class.
    'The ProjectInfo class contains methods to read and write project information including the Project Application Info data.

    Public Sub ReadFile(ByVal ApplicationDir As String)
        'Read the Application Summary properties from the Application_Info.xml file.
        'This file is in the Application Directory.

        If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.xml") Then
            'The Application_Info.xml file exists.
            'Read the Appllication information:
            Dim applicationInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\" & "Application_Info.xml")

            'Read Application Name:
            If applicationInfoXDoc.<Application>.<Name>.Value = Nothing Then
                Name = ""
            Else
                Name = applicationInfoXDoc.<Application>.<Name>.Value
            End If

            'Read Application Description:
            If applicationInfoXDoc.<Application>.<Description>.Value = Nothing Then
                Description = ""
            Else
                Description = applicationInfoXDoc.<Application>.<Description>.Value
            End If

            'Read Application Creation Date:
            If applicationInfoXDoc.<Application>.<CreationDate>.Value = Nothing Then
                CreationDate = "1-Jan-2000 12:00:00"
            Else
                CreationDate = applicationInfoXDoc.<Application>.<CreationDate>.Value
            End If

            'Read Author Information ------------------------------------------------------------------------------
            'Read Application Author - Name
            If applicationInfoXDoc.<Application>.<Author>.<Name>.Value = Nothing Then
                Author.Name = ""
            Else
                Author.Name = applicationInfoXDoc.<Application>.<Author>.<Name>.Value
            End If

            'Read Application Author - Description
            If applicationInfoXDoc.<Application>.<Author>.<Description>.Value = Nothing Then
                Author.Description = ""
            Else
                Author.Description = applicationInfoXDoc.<Application>.<Author>.<Description>.Value
            End If

            'Read Application Author - Contact
            If applicationInfoXDoc.<Application>.<Author>.<Contact>.Value = Nothing Then
                Author.Contact = ""
            Else
                Author.Contact = applicationInfoXDoc.<Application>.<Author>.<Contact>.Value
            End If

            'Read Version Information ---------------------------------------------------------------------------
            'Read Application Version - Major
            If applicationInfoXDoc.<Application>.<Version>.<Major>.Value = Nothing Then
                Version.Major = 1
            Else
                Version.Major = applicationInfoXDoc.<Application>.<Version>.<Major>.Value
            End If

            'Read Application Version - Minor
            If applicationInfoXDoc.<Application>.<Version>.<Minor>.Value = Nothing Then
                Version.Minor = 0
            Else
                Version.Minor = applicationInfoXDoc.<Application>.<Version>.<Minor>.Value
            End If

            'Read Application Version - Build
            If applicationInfoXDoc.<Application>.<Version>.<Build>.Value = Nothing Then
                Version.Build = 1
            Else
                Version.Build = applicationInfoXDoc.<Application>.<Version>.<Build>.Value
            End If

            'Read Application Version - Revision
            If applicationInfoXDoc.<Application>.<Version>.<Revision>.Value = Nothing Then
                Version.Revision = 0
            Else
                Version.Revision = applicationInfoXDoc.<Application>.<Version>.<Revision>.Value
            End If


            'Read Copyright Information ------------------------------------------------------------------------
            'Read Copyright Owner Name
            'If applicationInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value = Nothing Then
            '    Copyright.OwnerName = ""
            'Else
            '    Copyright.OwnerName = applicationInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value
            'End If

            ''Read Copyright Publication Year
            'If applicationInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value = Nothing Then
            '    Copyright.PublicationYear = ""
            'Else
            '    Copyright.PublicationYear = applicationInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value
            'End If

            'Read Copyright Notice
            'NOTE: Copyright Notice is generated from the OwnerName and the PublicationYear.
            Dim PublicationYear As String = ""
            If applicationInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value = Nothing Then
                PublicationYear = ""
            Else
                PublicationYear = applicationInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value
            End If
            Dim OwnerName As String = ""
            If applicationInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value = Nothing Then
                OwnerName = ""
            Else
                OwnerName = applicationInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value
            End If
            CopyrightNotice = "© " & OwnerName & " " & PublicationYear

            'Read License Information ------------------------------------------------------------------------
            'Read License Notice
            'If applicationInfoXDoc.<Application>.<License>.<Notice>.Value = Nothing Then
            '    License.Notice = ""
            'Else
            '    License.Notice = applicationInfoXDoc.<Application>.<License>.<Notice>.Value
            'End If

            ''Read License Text
            'If applicationInfoXDoc.<Application>.<License>.<Text>.Value = Nothing Then
            '    License.Text = ""
            'Else
            '    License.Text = applicationInfoXDoc.<Application>.<License>.<Text>.Value
            'End If

            If applicationInfoXDoc.<Application>.<License>.<Notice>.Value = Nothing Then
                LicenseNotice = ""
            Else
                LicenseNotice = applicationInfoXDoc.<Application>.<License>.<Notice>.Value
            End If


            'Read Source Code Information ---------------------------------------------------------------------
            'Read Source Code Language
            'If applicationInfoXDoc.<Application>.<SourceCode>.<Language>.Value = Nothing Then
            '    SourceCode.Language = ""
            'Else
            '    SourceCode.Language = applicationInfoXDoc.<Application>.<SourceCode>.<Language>.Value
            'End If

            ''Read Source Code File Name
            'If applicationInfoXDoc.<Application>.<SourceCode>.<FileName>.Value = Nothing Then
            '    SourceCode.FileName = ""
            'Else
            '    SourceCode.FileName = applicationInfoXDoc.<Application>.<SourceCode>.<FileName>.Value
            'End If

            ''Read Source Code File Size
            'If applicationInfoXDoc.<Application>.<SourceCode>.<FileSize>.Value = Nothing Then
            '    SourceCode.FileSize = 0
            'Else
            '    SourceCode.FileSize = applicationInfoXDoc.<Application>.<SourceCode>.<FileSize>.Value
            'End If

            ''Read Source Code File Hash
            'If applicationInfoXDoc.<Application>.<SourceCode>.<FileHash>.Value = Nothing Then
            '    SourceCode.FileHash = ""
            'Else
            '    SourceCode.FileHash = applicationInfoXDoc.<Application>.<SourceCode>.<FileHash>.Value
            'End If

            ''Read Source Code Web Link
            'If applicationInfoXDoc.<Application>.<SourceCode>.<WebLink>.Value = Nothing Then
            '    SourceCode.WebLink = ""
            'Else
            '    SourceCode.WebLink = applicationInfoXDoc.<Application>.<SourceCode>.<WebLink>.Value
            'End If

            ''Read Source Code Contact
            'If applicationInfoXDoc.<Application>.<SourceCode>.<Contact>.Value = Nothing Then
            '    SourceCode.Contact = ""
            'Else
            '    SourceCode.Contact = applicationInfoXDoc.<Application>.<SourceCode>.<Contact>.Value
            'End If

            ''Read Source Code Comments
            'If applicationInfoXDoc.<Application>.<SourceCode>.<Comments>.Value = Nothing Then
            '    SourceCode.Comments = ""
            'Else
            '    SourceCode.Comments = applicationInfoXDoc.<Application>.<SourceCode>.<Comments>.Value
            'End If

            'End If
        Else 'The Application_Info.xml file does not exist.

        End If
    End Sub

#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'ApplicationSummary ---------------------------------------------------------------------------------------------------------------------------------------------------------------

'The LibrarySummary class stores a summary of a software library used in an application.
Public Class LibrarySummary '----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables used in this class." '-------------------------------------------------------------------------------------------------------------------

    Public Version As New Version 'Stores library version information (Major, Minor, Build, Revision)
    Public Author As New Author 'Stores library author information (Name, Description, Contact)
    Public Classes As New List(Of ClassSummary)     'Stores information about classes contained within the library.

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - Properties used to store software library information." '-------------------------------------------------------------------------------------------------------------

    Private _name As String = "" 'The name of the software library
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A description of the library.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _creationDate As DateTime = "31-Dec-2015 22:00:00" 'The release date of the library.
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    Private _licenseNotice As String = "" 'The library licence notice.
    Property LicenseNotice As String
        Get
            Return _licenseNotice
        End Get
        Set(value As String)
            _licenseNotice = value
        End Set
    End Property

    Private _copyrightNotice As String = "" 'The library copyright notice.
    Property CopyrightNotice As String
        Get
            Return _copyrightNotice
        End Get
        Set(value As String)
            _copyrightNotice = value
        End Set
    End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------



End Class 'LibrarySummary -------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The ClassSummary class stores a summary of a class contained in a software library.
Public Class ClassSummary '------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private _name As String = "" 'The name of the class
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A description of the class.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

End Class 'ClassSummary ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The ModificationSummary class stores a summary of any modifications made to an application or library.
Public Class ModificationSummary '-----------------------------------------------------------------------------------------------------------------------------------------------------------

    Public BaseCodeVersion As New Version 'The version of the application or software library that was modified to produce the current software.

    Private _baseCodeName As String = "" 'The name of the application or software library that was modified to produce the current software.
    Property BaseCodeName As String
        Get
            Return _baseCodeName
        End Get
        Set(value As String)
            _baseCodeName = value
        End Set
    End Property

    Private _baseCodeDescription As String = "" 'A description of the application or software that was modified to produce the current software.
    Property BaseCodeDescription As String
        Get
            Return _baseCodeDescription
        End Get
        Set(value As String)
            _baseCodeDescription = value
        End Set
    End Property

    Private _description As String = "" 'A description of the modication(s) that have been made to the base code.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

End Class 'ModificationSummary --------------------------------------------------------------------------------------------------------------------------------------------------------------



'The ApplicationInfo class stores information about an Andorville (TM) application.
Public Class ApplicationInfo '---------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables used in this class." '-------------------------------------------------------------------------------------------------------------------

    Public Version As New Version       'Stores application version information (Major, Minor, Build, Revision).
    Public Author As New Author         'Stores information about the author of the application (Name, Description, Contact).
    Public Copyright As New Copyright   'Stores information about the application copyright.
    Public Trademarks As New List(Of Trademark)     'Stores information about trademarks owned by the application author or other owners.
    Public License As New License       'Stores details of the application license.
    Public SourceCode As New SourceCode 'Stores information about the application source code.
    Public FileAssociations As New List(Of FileAssociation) 'Stores information of the file type(s) associated with this application.
    Public Libraries As New List(Of LibrarySummary) 'Stores a summary of any software library used within the application.
    Public ModificationSummary As New ModificationSummary 'A summary of any modifications that have been applied to base code to produce the current version of the software.

    Public WithEvents AppInfoForm As frmAppInfo
    Public SettingsLocn As New FileLocation 'This is a directory or archive where settings are stored for a project. This is used to save the AppInfoForm settings.

    'Public TextDisplay As TextDisplay 'Class used to define the text displayed in a richtextbox.

#End Region 'Variable Declarations ----------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - Properties used to store application information." '------------------------------------------------------------------------------------------------------------------

    Private _applicationDir As String = "" 'The path to the directory used to store application data.
    Public Property ApplicationDir As String
        Get
            Return _applicationDir
        End Get
        Set(ByVal value As String)
            _applicationDir = value
        End Set
    End Property

    Private _name As String = "" 'The name of the application
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    'Private _exeFileName As String = "" 'The name of the application executable file name.
    'Property ExeFileName As String
    '    Get
    '        Return _exeFileName
    '    End Get
    '    Set(value As String)
    '        _exeFileName = value
    '    End Set
    'End Property

    Private _executablePath As String = "" 'The path of the applications executable file.
    Property ExecutablePath As String
        Get
            Return _executablePath
        End Get
        Set(value As String)
            _executablePath = value
        End Set
    End Property

    Private _description As String = "" 'A description of the application.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _creationDate As DateTime = "31-Dec-2015 22:00:00" 'The release date of the application.
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    'ADDED 20Feb19
    Private _connectOnStartup As Boolean = True 'If True, the application will connect to the Message Service on startup.
    Property ConnectOnStartup As Boolean
        Get
            Return _connectOnStartup
        End Get
        Set(value As Boolean)
            _connectOnStartup = value
        End Set
    End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Function ApplicationLocked() As Boolean
        'Returns True if a lock file is found.
        'If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.lock") Then
        If System.IO.File.Exists(ApplicationDir & "\" & "Application.Lock") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub LockApplication()
        If ApplicationLocked() Then
            RaiseEvent ErrorMessage("Application already locked." & vbCrLf)
        Else
            'Create a lock file:
            'Dim fs = System.IO.File.Create(ApplicationDir & "\" & "Application_Info.lock")
            Dim fs = System.IO.File.Create(ApplicationDir & "\" & "Application.Lock")
            fs.Dispose() 'This closes the file.
        End If
    End Sub

    Public Sub UnlockApplication()
        If ApplicationLocked() Then
            'System.IO.File.Delete(ApplicationDir & "\" & "Application_Info.lock")
            System.IO.File.Delete(ApplicationDir & "\" & "Application.Lock")
        Else
            RaiseEvent ErrorMessage("Application already unlocked." & vbCrLf)
        End If
    End Sub

    Public Function FileExists() As Boolean
        'Returns True if the Application Info File exists in the Application Directory.
        'Return System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.xml")
        'Return System.IO.File.Exists(ApplicationDir & "\" & "Application.xml")

        'Return System.IO.File.Exists(ApplicationDir & "\" & "Application_Info_ADVL_2.xml")

        'UPDATE 26May21: Need to also check if the file has been corrupted.
        'Return False if the file can not be opened.
        If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info_ADVL_2.xml") Then 'File Found
            Try
                Dim AppInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\" & "Application_Info_ADVL_2.xml")
                Return True
            Catch ex As Exception
                RaiseEvent ErrorMessage("The Application Information file can not be read: " & vbCrLf & ex.Message & vbCrLf)
                Return False
            End Try
        Else
            Return False
        End If

    End Function


    '1 Aug 2016
    'Noticed that the ApplicationInfoXDoc function code was corrupted.
    'The code was replaced with code retrieved from an archive dated 30 Jun 2016

    'Public Function ApplicationInfoXDoc() As System.Xml.Linq.XDocument
    '    'Return the Application Information in an XDocument.

    '    ApplicationInfoXDoc = <?xml version="1.0" encoding="utf-8"?>


    'End Function



    Public Function ApplicationInfoAdvl_1XDoc() As System.Xml.Linq.XDocument
        'Return the Application Information in an XDocument. (ADVL_1 format.)

        ApplicationInfoAdvl_1XDoc = <?xml version="1.0" encoding="utf-8"?>
                                    <!---->
                                    <!--Application Information for Application: ADVL_Zip-->
                                    <!---->
                                    <Application>
                                        <Name><%= Name %></Name>
                                        <ExecutablePath><%= ExecutablePath %></ExecutablePath>
                                        <Description><%= Description %></Description>
                                        <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                        <FileAssociationList>
                                            <%= From item In FileAssociations
                                                Select
                                          <FileAssociation>
                                              <Extension><%= item.Extension %></Extension>
                                              <Description><%= item.Description %></Description>
                                          </FileAssociation>
                                            %>
                                        </FileAssociationList>
                                        <Author>
                                            <Name><%= Author.Name %></Name>
                                            <Description><%= Author.Description %></Description>
                                            <Contact><%= Author.Contact %></Contact>
                                        </Author>
                                        <Copyright>
                                            <OwnerName><%= Copyright.OwnerName %></OwnerName>
                                            <PublicationYear><%= Copyright.PublicationYear %></PublicationYear>
                                            <Notice><%= Copyright.Notice %></Notice>
                                        </Copyright>
                                        <TrademarkList>
                                            <%= From item In Trademarks
                                                Select
                                          <Trademark>
                                              <Text><%= item.Text %></Text>
                                              <OwnerName><%= item.OwnerName %></OwnerName>
                                              <Registered><%= item.Registered %></Registered>
                                              <GenericTerm><%= item.GenericTerm %></GenericTerm>
                                          </Trademark>
                                            %>
                                        </TrademarkList>
                                        <License>
                                            <Code><%= License.Code.ToString %></Code>
                                            <Notice><%= License.Notice %></Notice>
                                            <Text><%= License.Text %></Text>
                                        </License>
                                        <SourceCode>
                                            <Language><%= SourceCode.Language %></Language>
                                            <FileName><%= SourceCode.FileName %></FileName>
                                            <FileSize><%= SourceCode.FileSize %></FileSize>
                                            <FileHash><%= SourceCode.FileHash %></FileHash>
                                            <WebLink><%= SourceCode.WebLink %></WebLink>
                                            <Contact><%= SourceCode.Contact %></Contact>
                                            <Comments><%= SourceCode.Comments %></Comments>
                                        </SourceCode>
                                        <ModificationSummary>
                                            <BaseCodeName><%= ModificationSummary.BaseCodeName %></BaseCodeName>
                                            <BaseCodeDescription><%= ModificationSummary.BaseCodeDescription %></BaseCodeDescription>
                                            <BaseCodeVersion>
                                                <Major><%= ModificationSummary.BaseCodeVersion.Major %></Major>
                                                <Minor><%= ModificationSummary.BaseCodeVersion.Minor %></Minor>
                                                <Build><%= ModificationSummary.BaseCodeVersion.Build %></Build>
                                                <Revision><%= ModificationSummary.BaseCodeVersion.Revision %></Revision>
                                            </BaseCodeVersion>
                                            <Description><%= ModificationSummary.Description %></Description>
                                        </ModificationSummary>
                                        <LibraryList>
                                            <%= From item In Libraries
                                                Select
                                          <Library>
                                              <Name><%= item.Name %></Name>
                                              <Description><%= item.Description %></Description>
                                              <CreationDate><%= Format(item.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                              <LicenseNotice><%= item.LicenseNotice %></LicenseNotice>
                                              <CopyrightNotice><%= item.CopyrightNotice %></CopyrightNotice>
                                              <Version>
                                                  <Major><%= item.Version.Major %></Major>
                                                  <Minor><%= item.Version.Minor %></Minor>
                                                  <Build><%= item.Version.Build %></Build>
                                                  <Revision><%= item.Version.Revision %></Revision>
                                              </Version>
                                              <Author>
                                                  <Name><%= item.Author.Name %></Name>
                                                  <Description><%= item.Author.Description %></Description>
                                                  <Contact><%= item.Author.Contact %></Contact>
                                              </Author>
                                              <ClassList>
                                                  <%= From classItem In item.Classes
                                                      Select
                                                      <Class>
                                                          <Name><%= classItem.Name %></Name>
                                                          <Description><%= classItem.Description %></Description>
                                                      </Class>
                                                  %>
                                              </ClassList>
                                          </Library>
                                            %>
                                        </LibraryList>
                                    </Application>

        '</FileAssociationList>
        '<Version>
        '    <Major><%= Version.Major %></Major>
        '    <Minor><%= Version.Minor %></Minor>
        '    <Build><%= Version.Build %></Build>
        '    <Revision><%= Version.Revision %></Revision>
        '</Version>

    End Function

    Public Function ApplicationInfoAdvl_2XDoc() As System.Xml.Linq.XDocument
        'Return the Application Information in an XDocument. (ADVL_2 format.)

        ApplicationInfoAdvl_2XDoc = <?xml version="1.0" encoding="utf-8"?>
                                    <!---->
                                    <!--Application Information File-->
                                    <!---->
                                    <Application>
                                        <FormatCode>ADVL_2</FormatCode>
                                        <Name><%= Name %></Name>
                                        <ExecutablePath><%= ExecutablePath %></ExecutablePath>
                                        <Description><%= Description %></Description>
                                        <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                        <FileAssociationList>
                                            <%= From item In FileAssociations
                                                Select
                                          <FileAssociation>
                                              <Extension><%= item.Extension %></Extension>
                                              <Description><%= item.Description %></Description>
                                          </FileAssociation>
                                            %>
                                        </FileAssociationList>
                                        <Author>
                                            <Name><%= Author.Name %></Name>
                                            <Description><%= Author.Description %></Description>
                                            <Contact><%= Author.Contact %></Contact>
                                        </Author>
                                        <Copyright>
                                            <OwnerName><%= Copyright.OwnerName %></OwnerName>
                                            <PublicationYear><%= Copyright.PublicationYear %></PublicationYear>
                                            <Notice><%= Copyright.Notice %></Notice>
                                        </Copyright>
                                        <TrademarkList>
                                            <%= From item In Trademarks
                                                Select
                                          <Trademark>
                                              <Text><%= item.Text %></Text>
                                              <OwnerName><%= item.OwnerName %></OwnerName>
                                              <Registered><%= item.Registered %></Registered>
                                              <GenericTerm><%= item.GenericTerm %></GenericTerm>
                                          </Trademark>
                                            %>
                                        </TrademarkList>
                                        <License>
                                            <Code><%= License.Code.ToString %></Code>
                                            <Notice><%= License.Notice %></Notice>
                                            <Text><%= License.Text %></Text>
                                        </License>
                                        <SourceCode>
                                            <Language><%= SourceCode.Language %></Language>
                                            <FileName><%= SourceCode.FileName %></FileName>
                                            <FileSize><%= SourceCode.FileSize %></FileSize>
                                            <FileHash><%= SourceCode.FileHash %></FileHash>
                                            <WebLink><%= SourceCode.WebLink %></WebLink>
                                            <Contact><%= SourceCode.Contact %></Contact>
                                            <Comments><%= SourceCode.Comments %></Comments>
                                        </SourceCode>
                                        <ModificationSummary>
                                            <BaseCodeName><%= ModificationSummary.BaseCodeName %></BaseCodeName>
                                            <BaseCodeDescription><%= ModificationSummary.BaseCodeDescription %></BaseCodeDescription>
                                            <BaseCodeVersion>
                                                <Major><%= ModificationSummary.BaseCodeVersion.Major %></Major>
                                                <Minor><%= ModificationSummary.BaseCodeVersion.Minor %></Minor>
                                                <Build><%= ModificationSummary.BaseCodeVersion.Build %></Build>
                                                <Revision><%= ModificationSummary.BaseCodeVersion.Revision %></Revision>
                                            </BaseCodeVersion>
                                            <Description><%= ModificationSummary.Description %></Description>
                                        </ModificationSummary>
                                        <LibraryList>
                                            <%= From item In Libraries
                                                Select
                                          <Library>
                                              <Name><%= item.Name %></Name>
                                              <Description><%= item.Description %></Description>
                                              <CreationDate><%= Format(item.CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                                              <LicenseNotice><%= item.LicenseNotice %></LicenseNotice>
                                              <CopyrightNotice><%= item.CopyrightNotice %></CopyrightNotice>
                                              <Version>
                                                  <Major><%= item.Version.Major %></Major>
                                                  <Minor><%= item.Version.Minor %></Minor>
                                                  <Build><%= item.Version.Build %></Build>
                                                  <Revision><%= item.Version.Revision %></Revision>
                                              </Version>
                                              <Author>
                                                  <Name><%= item.Author.Name %></Name>
                                                  <Description><%= item.Author.Description %></Description>
                                                  <Contact><%= item.Author.Contact %></Contact>
                                              </Author>
                                              <ClassList>
                                                  <%= From classItem In item.Classes
                                                      Select
                                                      <Class>
                                                          <Name><%= classItem.Name %></Name>
                                                          <Description><%= classItem.Description %></Description>
                                                      </Class>
                                                  %>
                                              </ClassList>
                                          </Library>
                                            %>
                                        </LibraryList>
                                        <ConnectOnStartup><%= ConnectOnStartup %></ConnectOnStartup>
                                    </Application>

        'Version information is now obtained directory from My.Application.Info.Version.
        '</FileAssociationList>
        '<Version>
        '    <Major><%= Version.Major %></Major>
        '    <Minor><%= Version.Minor %></Minor>
        '    <Build><%= Version.Build %></Build>
        '    <Revision><%= Version.Revision %></Revision>
        '</Version>

        '<ConnectOnStartup><%= ConnectOnStartup %></ConnectOnStartup> ADDED 20Feb19

    End Function

    Public Sub ReadFile()
        'Read the Application Info properties from the Application_Info.xml file.
        'This file is in the Application Directory.

        'If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.xml") Then
        If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info_ADVL_2.xml") Then
            'The Application_Info_ADVL_2.xml file exists.
            'Check if it is locked:
            'If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.lock") Then
            If System.IO.File.Exists(ApplicationDir & "\" & "Application.Lock") Then
                'The Application has been locked.
                'It may be being used by another instance of the application.
                'Wait a while and try again.

                'RaiseEvent ErrorMessage("The Application Info file is locked. Check if the application is already running." & vbCrLf)
                RaiseEvent ErrorMessage("The Application is locked. Check if the application is already running." & vbCrLf)
                'RaiseEvent ErrorMessage("If required delete the 'Application_Info.lock' file in the application directory: " & ApplicationDir & vbCrLf)
                RaiseEvent ErrorMessage("If required delete the 'Application.Lock' file in the application directory: " & ApplicationDir & vbCrLf)

            Else
                'The Application_Info.xml file is not locked.
                'Read the Appllication information:
                'Dim applicationInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\" & "Application_Info.xml")
                Dim AppInfoXDoc As System.Xml.Linq.XDocument = XDocument.Load(ApplicationDir & "\" & "Application_Info_ADVL_2.xml")

                ReadAppInfoAdvl_2(AppInfoXDoc)

            End If
        Else 'The Application_Info_ADVL_2.xml file does not exist.
            'Check if the previous format version of the file exists:
            If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info.xml") Then 'The orignal ADVL_1 format version found.
                RaiseEvent Message("Converting Application_Info.xml to Application_Info_ADVL_2.xml." & vbCrLf)
                'Convert the file to the latest ADVL_2 format:
                Dim Conversion As New ADVL_Utilities_Library_1.FormatConvert.AppInfoFileConversion
                Conversion.DirectoryPath = ApplicationDir
                Conversion.InputFileName = "Application_Info.xml"
                Conversion.InputFormatCode = FormatConvert.AppInfoFileConversion.FormatCodes.ADVL_1
                Conversion.OutputFormatCode = FormatConvert.AppInfoFileConversion.FormatCodes.ADVL_2
                Conversion.Convert()
                If System.IO.File.Exists(ApplicationDir & "\" & "Application_Info_ADVL_2.xml") Then
                    ReadFile() 'Try ReadFile again. This time Application_Info_ADVL_2.xml should be found
                Else
                    RaiseEvent ErrorMessage("Error converting Application_Info.xml to Application_Info_ADVL_2.xml." & vbCrLf)
                End If
            End If
        End If
    End Sub

    Private Sub ReadAppInfoAdvl_1(ByRef AppInfoXDoc As System.Xml.Linq.XDocument)
        'Read the the ADVL_1 format version of the Application Information file.

        'Read Application Name:
        If AppInfoXDoc.<Application>.<Name>.Value = Nothing Then
            Name = ""
        Else
            Name = AppInfoXDoc.<Application>.<Name>.Value
        End If

        'Read the Applications Executable Path:
        If AppInfoXDoc.<Application>.<ExecutablePath>.Value = Nothing Then
            ExecutablePath = ""
        Else
            ExecutablePath = AppInfoXDoc.<Application>.<ExecutablePath>.Value
        End If

        'Read Application Description:
        If AppInfoXDoc.<Application>.<Description>.Value = Nothing Then
            Description = ""
        Else
            Description = AppInfoXDoc.<Application>.<Description>.Value
        End If

        'Read Application Creation Date:
        If AppInfoXDoc.<Application>.<CreationDate>.Value = Nothing Then
            CreationDate = "1-Jan-2000 12:00:00"
        Else
            CreationDate = AppInfoXDoc.<Application>.<CreationDate>.Value
        End If

        'Read File Association(s): -------------------------------------------------------------------------------------
        If AppInfoXDoc.<Application>.<FileAssociationList>.<FileAssociation> Is Nothing Then

        Else
            Dim Assns = From item In AppInfoXDoc.<Application>.<FileAssociationList>.<FileAssociation>

            FileAssociations.Clear()
            For Each item In Assns
                Dim NewAssoc As New FileAssociation
                NewAssoc.Extension = item.<Extension>.Value
                NewAssoc.Description = item.<Description>.Value
                FileAssociations.Add(NewAssoc)
            Next
        End If

        'Read Author Information ------------------------------------------------------------------------------
        'Read Application Author - Name
        If AppInfoXDoc.<Application>.<Author>.<Name>.Value = Nothing Then
            Author.Name = ""
        Else
            Author.Name = AppInfoXDoc.<Application>.<Author>.<Name>.Value
        End If

        'Read Application Author - Description
        If AppInfoXDoc.<Application>.<Author>.<Description>.Value = Nothing Then
            Author.Description = ""
        Else
            Author.Description = AppInfoXDoc.<Application>.<Author>.<Description>.Value
        End If

        'Read Application Author - Contact
        If AppInfoXDoc.<Application>.<Author>.<Contact>.Value = Nothing Then
            Author.Contact = ""
        Else
            Author.Contact = AppInfoXDoc.<Application>.<Author>.<Contact>.Value
        End If

        'Read Version Information ---------------------------------------------------------------------------
        'Read Application Version - Major
        'If AppInfoXDoc.<Application>.<Version>.<Major>.Value = Nothing Then
        '    Version.Major = 1
        'Else
        '    Version.Major = AppInfoXDoc.<Application>.<Version>.<Major>.Value
        'End If

        ''Read Application Version - Minor
        'If AppInfoXDoc.<Application>.<Version>.<Minor>.Value = Nothing Then
        '    Version.Minor = 0
        'Else
        '    Version.Minor = AppInfoXDoc.<Application>.<Version>.<Minor>.Value
        'End If

        ''Read Application Version - Build
        'If AppInfoXDoc.<Application>.<Version>.<Build>.Value = Nothing Then
        '    Version.Build = 1
        'Else
        '    Version.Build = AppInfoXDoc.<Application>.<Version>.<Build>.Value
        'End If

        ''Read Application Version - Revision
        'If AppInfoXDoc.<Application>.<Version>.<Revision>.Value = Nothing Then
        '    Version.Revision = 0
        'Else
        '    Version.Revision = AppInfoXDoc.<Application>.<Version>.<Revision>.Value
        'End If

        'NOTE: THIS APPEARS TO GET THE SystemUtilties version info:
        'Version.Major = My.Application.Info.Version.Major.ToString
        'Version.Minor = My.Application.Info.Version.Minor.ToString
        'Version.Build = My.Application.Info.Version.Build.ToString
        'Version.Revision = My.Application.Info.Version.Revision.ToString

        'THIS IS NOW UPDATED DIRECTLY IN THE APPLICATION USING My.Application.Info.Version.

        'Read Copyright Information ------------------------------------------------------------------------
        'Read Copyright Owner Name
        If AppInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value = Nothing Then
            Copyright.OwnerName = ""
        Else
            Copyright.OwnerName = AppInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value
        End If

        'Read Copyright Publication Year
        If AppInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value = Nothing Then
            Copyright.PublicationYear = ""
        Else
            Copyright.PublicationYear = AppInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value
        End If

        'Read Copyright Notice
        'NOTE: Copyright Notice is read only (it is automatically generated from the OwnerName and the PublicationYear.)

        'Read Trademarks ---------------------------------------------------------------------------------
        If AppInfoXDoc.<Application>.<TrademarkList>.<Trademark> Is Nothing Then

        Else
            Dim TMarks = From item In AppInfoXDoc.<Application>.<TrademarkList>.<Trademark>

            Trademarks.Clear()
            For Each item In TMarks
                Dim NewTMark As New Trademark
                NewTMark.Text = item.<Text>.Value
                NewTMark.OwnerName = item.<OwnerName>.Value
                Select Case item.<Registered>.Value
                    Case "True"
                        NewTMark.Registered = True
                    Case "False"
                        NewTMark.Registered = False
                End Select
                NewTMark.GenericTerm = item.<GenericTerm>.Value
                Trademarks.Add(NewTMark)
            Next
        End If

        'Read License Information ------------------------------------------------------------------------
        'Read License Code:
        If AppInfoXDoc.<Application>.<License>.<Code>.Value = Nothing Then
            License.Code = ADVL_Utilities_Library_1.License.Codes.Unknown
        Else
            Select Case AppInfoXDoc.<Application>.<License>.<Code>.Value
                Case "None"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.None
                Case "Apache_License_2_0"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
                Case "GNU_GPL_V3_0"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.GNU_GPL_V3_0
                Case "Unknown"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.Unknown
                Case Else
                    License.Code = ADVL_Utilities_Library_1.License.Codes.Unknown
            End Select
        End If

        'Read License Notice
        If AppInfoXDoc.<Application>.<License>.<Notice>.Value = Nothing Then
            License.Notice = ""
        Else
            License.Notice = AppInfoXDoc.<Application>.<License>.<Notice>.Value
        End If

        'Read License Text
        If AppInfoXDoc.<Application>.<License>.<Text>.Value = Nothing Then
            License.Text = ""
        Else
            License.Text = AppInfoXDoc.<Application>.<License>.<Text>.Value
        End If

        'Read Source Code Information ---------------------------------------------------------------------
        'Read Source Code Language
        If AppInfoXDoc.<Application>.<SourceCode>.<Language>.Value = Nothing Then
            SourceCode.Language = ""
        Else
            SourceCode.Language = AppInfoXDoc.<Application>.<SourceCode>.<Language>.Value
        End If

        'Read Source Code File Name
        If AppInfoXDoc.<Application>.<SourceCode>.<FileName>.Value = Nothing Then
            SourceCode.FileName = ""
        Else
            SourceCode.FileName = AppInfoXDoc.<Application>.<SourceCode>.<FileName>.Value
        End If

        'Read Source Code File Size
        If AppInfoXDoc.<Application>.<SourceCode>.<FileSize>.Value = Nothing Then
            SourceCode.FileSize = 0
        Else
            SourceCode.FileSize = AppInfoXDoc.<Application>.<SourceCode>.<FileSize>.Value
        End If

        'Read Source Code File Hash
        If AppInfoXDoc.<Application>.<SourceCode>.<FileHash>.Value = Nothing Then
            SourceCode.FileHash = ""
        Else
            SourceCode.FileHash = AppInfoXDoc.<Application>.<SourceCode>.<FileHash>.Value
        End If

        'Read Source Code Web Link
        If AppInfoXDoc.<Application>.<SourceCode>.<WebLink>.Value = Nothing Then
            SourceCode.WebLink = ""
        Else
            SourceCode.WebLink = AppInfoXDoc.<Application>.<SourceCode>.<WebLink>.Value
        End If

        'Read Source Code Contact
        If AppInfoXDoc.<Application>.<SourceCode>.<Contact>.Value = Nothing Then
            SourceCode.Contact = ""
        Else
            SourceCode.Contact = AppInfoXDoc.<Application>.<SourceCode>.<Contact>.Value
        End If

        'Read Source Code Comments
        If AppInfoXDoc.<Application>.<SourceCode>.<Comments>.Value = Nothing Then
            SourceCode.Comments = ""
        Else
            SourceCode.Comments = AppInfoXDoc.<Application>.<SourceCode>.<Comments>.Value
        End If

        'Read Modification Summary: -----------------------------------------------------------------------
        'Read Base Code Name:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeName>.Value = Nothing Then
            ModificationSummary.BaseCodeName = ""
        Else
            ModificationSummary.BaseCodeName = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeName>.Value
        End If
        'Read Base Code Description:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeDescription>.Value = Nothing Then
            ModificationSummary.BaseCodeDescription = ""
        Else
            ModificationSummary.BaseCodeDescription = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeDescription>.Value
        End If
        'Read Base Code Version:
        'Major:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Major>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Major = 0
        Else
            ModificationSummary.BaseCodeVersion.Major = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Major>.Value
        End If
        'Minor:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Minor>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Minor = 0
        Else
            ModificationSummary.BaseCodeVersion.Minor = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Minor>.Value
        End If
        'Build:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Build>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Build = 0
        Else
            ModificationSummary.BaseCodeVersion.Build = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Build>.Value
        End If
        'Revision:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Revision>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Revision = 0
        Else
            ModificationSummary.BaseCodeVersion.Revision = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Revision>.Value
        End If
        'Read Description:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<Description>.Value = Nothing Then
            ModificationSummary.Description = ""
        Else
            ModificationSummary.Description = AppInfoXDoc.<Application>.<ModificationSummary>.<Description>.Value
        End If

        'Read Library List information:
        If AppInfoXDoc.<Application>.<LibraryList>.<Library> Is Nothing Then

        Else
            Dim Libs = From item In AppInfoXDoc.<Application>.<LibraryList>.<Library>

            Libraries.Clear()
            For Each item In Libs
                Dim NewLib As New LibrarySummary
                NewLib.Name = item.<Name>.Value
                NewLib.Description = item.<Description>.Value
                NewLib.CreationDate = item.<CreationDate>.Value
                NewLib.LicenseNotice = item.<LicenseNotice>.Value
                NewLib.CopyrightNotice = item.<CopyrightNotice>.Value
                NewLib.Version.Major = item.<Version>.<Major>.Value
                NewLib.Version.Minor = item.<Version>.<Minor>.Value
                NewLib.Version.Build = item.<Version>.<Build>.Value
                NewLib.Version.Revision = item.<Version>.<Revision>.Value
                NewLib.Author.Name = item.<Author>.<Name>.Value
                NewLib.Author.Description = item.<Author>.<Description>.Value
                NewLib.Author.Contact = item.<Author>.<Contact>.Value
                For Each classItem In item.<ClassList>.<Class>
                    Dim NewClass As New ClassSummary
                    NewClass.Name = classItem.<Name>.Value
                    NewClass.Description = classItem.<Description>.Value
                    NewLib.Classes.Add(NewClass)
                Next
                Libraries.Add(NewLib)
            Next
        End If
    End Sub

    Private Sub ReadAppInfoAdvl_2(ByRef AppInfoXDoc As System.Xml.Linq.XDocument)
        'Read the the ADVL_2 format version of the Application Information file.

        'Read Application Name:
        If AppInfoXDoc.<Application>.<Name>.Value = Nothing Then
            Name = ""
        Else
            Name = AppInfoXDoc.<Application>.<Name>.Value
        End If

        'Read the Applications Executable Path:
        If AppInfoXDoc.<Application>.<ExecutablePath>.Value = Nothing Then
            ExecutablePath = ""
        Else
            ExecutablePath = AppInfoXDoc.<Application>.<ExecutablePath>.Value
        End If

        'Read Application Description:
        If AppInfoXDoc.<Application>.<Description>.Value = Nothing Then
            Description = ""
        Else
            Description = AppInfoXDoc.<Application>.<Description>.Value
        End If

        'Read Application Creation Date:
        If AppInfoXDoc.<Application>.<CreationDate>.Value = Nothing Then
            CreationDate = "1-Jan-2000 12:00:00"
        Else
            CreationDate = AppInfoXDoc.<Application>.<CreationDate>.Value
        End If

        'Read File Association(s): -------------------------------------------------------------------------------------
        If AppInfoXDoc.<Application>.<FileAssociationList>.<FileAssociation> Is Nothing Then

        Else
            Dim Assns = From item In AppInfoXDoc.<Application>.<FileAssociationList>.<FileAssociation>

            FileAssociations.Clear()
            For Each item In Assns
                Dim NewAssoc As New FileAssociation
                NewAssoc.Extension = item.<Extension>.Value
                NewAssoc.Description = item.<Description>.Value
                FileAssociations.Add(NewAssoc)
            Next
        End If

        'Read Author Information ------------------------------------------------------------------------------
        'Read Application Author - Name
        If AppInfoXDoc.<Application>.<Author>.<Name>.Value = Nothing Then
            Author.Name = ""
        Else
            Author.Name = AppInfoXDoc.<Application>.<Author>.<Name>.Value
        End If

        'Read Application Author - Description
        If AppInfoXDoc.<Application>.<Author>.<Description>.Value = Nothing Then
            Author.Description = ""
        Else
            Author.Description = AppInfoXDoc.<Application>.<Author>.<Description>.Value
        End If

        'Read Application Author - Contact
        If AppInfoXDoc.<Application>.<Author>.<Contact>.Value = Nothing Then
            Author.Contact = ""
        Else
            Author.Contact = AppInfoXDoc.<Application>.<Author>.<Contact>.Value
        End If

        'THIS IS NOW UPDATED DIRECTLY IN THE APPLICATION USING My.Application.Info.Version.
        ''Read Version Information ---------------------------------------------------------------------------
        ''Read Application Version - Major
        'If AppInfoXDoc.<Application>.<Version>.<Major>.Value = Nothing Then
        '    Version.Major = 1
        'Else
        '    Version.Major = AppInfoXDoc.<Application>.<Version>.<Major>.Value
        'End If

        ''Read Application Version - Minor
        'If AppInfoXDoc.<Application>.<Version>.<Minor>.Value = Nothing Then
        '    Version.Minor = 0
        'Else
        '    Version.Minor = AppInfoXDoc.<Application>.<Version>.<Minor>.Value
        'End If

        ''Read Application Version - Build
        'If AppInfoXDoc.<Application>.<Version>.<Build>.Value = Nothing Then
        '    Version.Build = 1
        'Else
        '    Version.Build = AppInfoXDoc.<Application>.<Version>.<Build>.Value
        'End If

        ''Read Application Version - Revision
        'If AppInfoXDoc.<Application>.<Version>.<Revision>.Value = Nothing Then
        '    Version.Revision = 0
        'Else
        '    Version.Revision = AppInfoXDoc.<Application>.<Version>.<Revision>.Value
        'End If

        'Read Copyright Information ------------------------------------------------------------------------
        'Read Copyright Owner Name
        If AppInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value = Nothing Then
            Copyright.OwnerName = ""
        Else
            Copyright.OwnerName = AppInfoXDoc.<Application>.<Copyright>.<OwnerName>.Value
        End If

        'Read Copyright Publication Year
        If AppInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value = Nothing Then
            Copyright.PublicationYear = ""
        Else
            Copyright.PublicationYear = AppInfoXDoc.<Application>.<Copyright>.<PublicationYear>.Value
        End If

        'Read Copyright Notice
        'NOTE: Copyright Notice is read only (it is automatically generated from the OwnerName and the PublicationYear.)

        'Read Trademarks ---------------------------------------------------------------------------------
        If AppInfoXDoc.<Application>.<TrademarkList>.<Trademark> Is Nothing Then

        Else
            Dim TMarks = From item In AppInfoXDoc.<Application>.<TrademarkList>.<Trademark>

            Trademarks.Clear()
            For Each item In TMarks
                Dim NewTMark As New Trademark
                NewTMark.Text = item.<Text>.Value
                NewTMark.OwnerName = item.<OwnerName>.Value
                Select Case item.<Registered>.Value
                    Case "True"
                        NewTMark.Registered = True
                    Case "False"
                        NewTMark.Registered = False
                End Select
                NewTMark.GenericTerm = item.<GenericTerm>.Value
                Trademarks.Add(NewTMark)
            Next
        End If

        'Read License Information ------------------------------------------------------------------------
        'Read License Code:
        If AppInfoXDoc.<Application>.<License>.<Code>.Value = Nothing Then
            License.Code = ADVL_Utilities_Library_1.License.Codes.Unknown
        Else
            Select Case AppInfoXDoc.<Application>.<License>.<Code>.Value
                Case "None"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.None
                Case "Apache_License_2_0"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
                Case "GNU_GPL_V3_0"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.GNU_GPL_V3_0
                Case "Unknown"
                    License.Code = ADVL_Utilities_Library_1.License.Codes.Unknown
                Case Else
                    License.Code = ADVL_Utilities_Library_1.License.Codes.Unknown
            End Select
        End If

        'Read License Notice
        If AppInfoXDoc.<Application>.<License>.<Notice>.Value = Nothing Then
            License.Notice = ""
        Else
            License.Notice = AppInfoXDoc.<Application>.<License>.<Notice>.Value
        End If

        'Read License Text
        If AppInfoXDoc.<Application>.<License>.<Text>.Value = Nothing Then
            License.Text = ""
        Else
            License.Text = AppInfoXDoc.<Application>.<License>.<Text>.Value
        End If

        'Read Source Code Information ---------------------------------------------------------------------
        'Read Source Code Language
        If AppInfoXDoc.<Application>.<SourceCode>.<Language>.Value = Nothing Then
            SourceCode.Language = ""
        Else
            SourceCode.Language = AppInfoXDoc.<Application>.<SourceCode>.<Language>.Value
        End If

        'Read Source Code File Name
        If AppInfoXDoc.<Application>.<SourceCode>.<FileName>.Value = Nothing Then
            SourceCode.FileName = ""
        Else
            SourceCode.FileName = AppInfoXDoc.<Application>.<SourceCode>.<FileName>.Value
        End If

        'Read Source Code File Size
        If AppInfoXDoc.<Application>.<SourceCode>.<FileSize>.Value = Nothing Then
            SourceCode.FileSize = 0
        Else
            SourceCode.FileSize = AppInfoXDoc.<Application>.<SourceCode>.<FileSize>.Value
        End If

        'Read Source Code File Hash
        If AppInfoXDoc.<Application>.<SourceCode>.<FileHash>.Value = Nothing Then
            SourceCode.FileHash = ""
        Else
            SourceCode.FileHash = AppInfoXDoc.<Application>.<SourceCode>.<FileHash>.Value
        End If

        'Read Source Code Web Link
        If AppInfoXDoc.<Application>.<SourceCode>.<WebLink>.Value = Nothing Then
            SourceCode.WebLink = ""
        Else
            SourceCode.WebLink = AppInfoXDoc.<Application>.<SourceCode>.<WebLink>.Value
        End If

        'Read Source Code Contact
        If AppInfoXDoc.<Application>.<SourceCode>.<Contact>.Value = Nothing Then
            SourceCode.Contact = ""
        Else
            SourceCode.Contact = AppInfoXDoc.<Application>.<SourceCode>.<Contact>.Value
        End If

        'Read Source Code Comments
        If AppInfoXDoc.<Application>.<SourceCode>.<Comments>.Value = Nothing Then
            SourceCode.Comments = ""
        Else
            SourceCode.Comments = AppInfoXDoc.<Application>.<SourceCode>.<Comments>.Value
        End If

        'Read Modification Summary: -----------------------------------------------------------------------
        'Read Base Code Name:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeName>.Value = Nothing Then
            ModificationSummary.BaseCodeName = ""
        Else
            ModificationSummary.BaseCodeName = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeName>.Value
        End If
        'Read Base Code Description:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeDescription>.Value = Nothing Then
            ModificationSummary.BaseCodeDescription = ""
        Else
            ModificationSummary.BaseCodeDescription = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeDescription>.Value
        End If
        'Read Base Code Version:
        'Major:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Major>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Major = 0
        Else
            ModificationSummary.BaseCodeVersion.Major = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Major>.Value
        End If
        'Minor:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Minor>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Minor = 0
        Else
            ModificationSummary.BaseCodeVersion.Minor = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Minor>.Value
        End If
        'Build:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Build>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Build = 0
        Else
            ModificationSummary.BaseCodeVersion.Build = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Build>.Value
        End If
        'Revision:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Revision>.Value = Nothing Then
            ModificationSummary.BaseCodeVersion.Revision = 0
        Else
            ModificationSummary.BaseCodeVersion.Revision = AppInfoXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Revision>.Value
        End If
        'Read Description:
        If AppInfoXDoc.<Application>.<ModificationSummary>.<Description>.Value = Nothing Then
            ModificationSummary.Description = ""
        Else
            ModificationSummary.Description = AppInfoXDoc.<Application>.<ModificationSummary>.<Description>.Value
        End If

        'Read Library List information:
        If AppInfoXDoc.<Application>.<LibraryList>.<Library> Is Nothing Then

        Else
            Dim Libs = From item In AppInfoXDoc.<Application>.<LibraryList>.<Library>

            Libraries.Clear()
            For Each item In Libs
                Dim NewLib As New LibrarySummary
                NewLib.Name = item.<Name>.Value
                NewLib.Description = item.<Description>.Value
                NewLib.CreationDate = item.<CreationDate>.Value
                NewLib.LicenseNotice = item.<LicenseNotice>.Value
                NewLib.CopyrightNotice = item.<CopyrightNotice>.Value
                NewLib.Version.Major = item.<Version>.<Major>.Value
                NewLib.Version.Minor = item.<Version>.<Minor>.Value
                NewLib.Version.Build = item.<Version>.<Build>.Value
                NewLib.Version.Revision = item.<Version>.<Revision>.Value
                NewLib.Author.Name = item.<Author>.<Name>.Value
                NewLib.Author.Description = item.<Author>.<Description>.Value
                NewLib.Author.Contact = item.<Author>.<Contact>.Value
                For Each classItem In item.<ClassList>.<Class>
                    Dim NewClass As New ClassSummary
                    NewClass.Name = classItem.<Name>.Value
                    NewClass.Description = classItem.<Description>.Value
                    NewLib.Classes.Add(NewClass)
                Next
                Libraries.Add(NewLib)
            Next
        End If

        'ADDED 20Feb18
        'Read ConnectOnStartup:
        If AppInfoXDoc.<Application>.<ConnectOnStartup>.Value Is Nothing Then
            ConnectOnStartup = True
        Else
            ConnectOnStartup = AppInfoXDoc.<Application>.<ConnectOnStartup>.Value
            Debug.Print("ConnectOnStartup:  " & ConnectOnStartup.ToString)
        End If

    End Sub

    Public Sub WriteFile()
        'Write the Application Info propserties to the Application_Info.xml file.

        'ApplicationInfoAdvl_1XDoc.Save(ApplicationDir & "\" & "Application_Info.xml")
        ApplicationInfoAdvl_2XDoc.Save(ApplicationDir & "\" & "Application_Info_ADVL_2.xml")

    End Sub

    Public Sub ShowInfo()
        'Open the AppInfo Form and display the application information.

        If IsNothing(AppInfoForm) Then
            AppInfoForm = New frmAppInfo
            AppInfoForm.ApplicationName = Name
            AppInfoForm.SettingsLocn = SettingsLocn
            AppInfoForm.Show()
            AppInfoForm.SettingsLocn = SettingsLocn
            AppInfoForm.RestoreFormSettings()

            ''Show the Application Summary information: ------------------------------------------------------
            'AppInfoForm.txtAppName.Text = Name
            'AppInfoForm.txtExecutablePath.Text = ExecutablePath
            'AppInfoForm.txtDescription.Text = Description
            'AppInfoForm.txtCreationDate.Text = CreationDate
            'AppInfoForm.txtMajor.Text = Version.Major
            'AppInfoForm.txtMinor.Text = Version.Minor
            'AppInfoForm.txtBuild.Text = Version.Build
            'AppInfoForm.txtRevision.Text = Version.Revision
            'AppInfoForm.txtAuthorName.Text = Author.Name
            'AppInfoForm.rtbAuthorDesc.Text = Author.Description
            'AppInfoForm.txtAuthorContact.Text = Author.Contact
            'AppInfoForm.txtCopyrightNotice.Text = Copyright.Notice

            ''Show the first file association record: ------------------------------------------------------
            'AppInfoForm.NAssoc = FileAssociations.Count
            'If FileAssociations.Count > 0 Then
            '    AppInfoForm.SelectedAssocNo = 1
            'Else

            'End If

            ''Show the Application License information: ----------------------------------------------------
            'AppInfoForm.rtbLicenceNotice.Text = License.Notice
            'AppInfoForm.rtbLicenseText.Text = License.Text

            ''Show trademark information: ------------------------------------------------------------------
            'AppInfoForm.rtbTrademarks.Clear()
            ''Add the Trademarks title to the page
            'AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 11, Drawing.FontStyle.Bold Or Drawing.FontStyle.Underline)
            'AppInfoForm.rtbTrademarks.SelectionColor = Drawing.Color.Black
            'AppInfoForm.rtbTrademarks.SelectedText = "Trademarks" & vbCrLf & vbCrLf

            'For Each item In Trademarks
            '    AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
            '    AppInfoForm.rtbTrademarks.SelectionColor = Drawing.Color.Black
            '    AppInfoForm.rtbTrademarks.SelectedText = item.Text
            '    If item.Registered = True Then
            '        'Add the ® superscript:
            '        AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
            '        AppInfoForm.rtbTrademarks.SelectionCharOffset = 10
            '        AppInfoForm.rtbTrademarks.SelectedText = "® "
            '        'Add the trademark owner name:
            '        AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Regular)
            '        AppInfoForm.rtbTrademarks.SelectionCharOffset = 0
            '        AppInfoForm.rtbTrademarks.SelectedText = "is a registered trademark of " & item.OwnerName & vbCrLf & vbCrLf
            '    Else
            '        'Add the TM superscript:
            '        AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 7, Drawing.FontStyle.Bold)
            '        AppInfoForm.rtbTrademarks.SelectionCharOffset = 10
            '        AppInfoForm.rtbTrademarks.SelectedText = "TM "
            '        'Add the trademark owner name:
            '        AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Regular)
            '        AppInfoForm.rtbTrademarks.SelectionCharOffset = 0
            '        AppInfoForm.rtbTrademarks.SelectedText = "is a trademark of " & item.OwnerName & vbCrLf & vbCrLf

            '    End If
            'Next

            ''Show Source Code information: --------------------------------------------------------------------------------
            'AppInfoForm.txtLanguage.Text = SourceCode.Language
            'AppInfoForm.txtSourceFileName.Text = SourceCode.FileName
            'AppInfoForm.txtSourceFileSize.Text = SourceCode.FileSize
            'AppInfoForm.txtSourceFileHash.Text = SourceCode.FileHash
            'AppInfoForm.txtSourceWebLink.Text = SourceCode.WebLink
            'AppInfoForm.txtSourceContact.Text = SourceCode.Contact
            'AppInfoForm.txtSourceComments.Text = SourceCode.Comments

            ''Show Modifications: -----------------------------------------------------------------------------------------
            'AppInfoForm.txtBaseCodeName.Text = ModificationSummary.BaseCodeName
            'AppInfoForm.txtBaseCodeDescription.Text = ModificationSummary.BaseCodeDescription
            'AppInfoForm.txtBaseCodeVersionMajor.Text = ModificationSummary.BaseCodeVersion.Major
            'AppInfoForm.txtBaseCodeVersionMinor.Text = ModificationSummary.BaseCodeVersion.Minor
            'AppInfoForm.txtBaseCodeVersionBuild.Text = ModificationSummary.BaseCodeVersion.Build
            'AppInfoForm.txtBaseCodeVersionRevision.Text = ModificationSummary.BaseCodeVersion.Revision
            'AppInfoForm.rtbModifications.Text = ModificationSummary.Description

            ''Show the first software library record: ------------------------------------------------------
            'AppInfoForm.NLibraries = Libraries.Count
            'If Libraries.Count > 0 Then
            '    AppInfoForm.SelectedLibraryNo = 1
            'Else

            'End If

        Else
            AppInfoForm.Show()
            AppInfoForm.BringToFront()
        End If

        'UPDATE THE INFORMATION EVERY TIME SHOWINFO() IS CALLED.
        'WHEN APPINTO.RESTOREDEFAULTS IS CALLED, THE DISPLAYED PARAMETERS WILL NEED TO BE UPDATED.

        'Show the Application Summary information: ------------------------------------------------------
        AppInfoForm.txtAppName.Text = Name
        AppInfoForm.txtExecutablePath.Text = ExecutablePath
        AppInfoForm.txtDescription.Text = Description
        AppInfoForm.txtCreationDate.Text = CreationDate
        AppInfoForm.txtVersion.Text = Version.Number
        AppInfoForm.txtMajor.Text = Version.Major
        AppInfoForm.txtMinor.Text = Version.Minor
        AppInfoForm.txtBuild.Text = Version.Build
        AppInfoForm.txtRevision.Text = Version.Revision
        AppInfoForm.txtSource.Text = Version.Source
        AppInfoForm.txtAuthorName.Text = Author.Name
        AppInfoForm.rtbAuthorDesc.Text = Author.Description
        AppInfoForm.txtAuthorContact.Text = Author.Contact
        AppInfoForm.txtCopyrightNotice.Text = Copyright.Notice

        'Show the first file association record: ------------------------------------------------------
        AppInfoForm.NAssoc = FileAssociations.Count
        If FileAssociations.Count > 0 Then
            AppInfoForm.SelectedAssocNo = 1
        Else

        End If

        'Show the Application License information: ----------------------------------------------------
        AppInfoForm.rtbLicenceNotice.Text = License.Notice
        AppInfoForm.rtbLicenseText.Text = License.Text

        'Show trademark information: ------------------------------------------------------------------
        AppInfoForm.rtbTrademarks.Clear()
        'Add the Trademarks title to the page
        AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 11, Drawing.FontStyle.Bold Or Drawing.FontStyle.Underline)
        AppInfoForm.rtbTrademarks.SelectionColor = Drawing.Color.Black
        AppInfoForm.rtbTrademarks.SelectedText = "Trademarks" & vbCrLf & vbCrLf

        For Each item In Trademarks
            AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
            AppInfoForm.rtbTrademarks.SelectionColor = Drawing.Color.Black
            AppInfoForm.rtbTrademarks.SelectedText = item.Text
            If item.Registered = True Then
                'Add the ® superscript:
                AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
                AppInfoForm.rtbTrademarks.SelectionCharOffset = 10
                AppInfoForm.rtbTrademarks.SelectedText = "® "
                'Add the trademark owner name:
                AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Regular)
                AppInfoForm.rtbTrademarks.SelectionCharOffset = 0
                AppInfoForm.rtbTrademarks.SelectedText = "is a registered trademark of " & item.OwnerName & vbCrLf & vbCrLf
            Else
                'Add the TM superscript:
                AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 7, Drawing.FontStyle.Bold)
                AppInfoForm.rtbTrademarks.SelectionCharOffset = 10
                AppInfoForm.rtbTrademarks.SelectedText = "TM "
                'Add the trademark owner name:
                AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Regular)
                AppInfoForm.rtbTrademarks.SelectionCharOffset = 0
                AppInfoForm.rtbTrademarks.SelectedText = "is a trademark of " & item.OwnerName & vbCrLf & vbCrLf

            End If
        Next

        'Show Source Code information: --------------------------------------------------------------------------------
        AppInfoForm.txtLanguage.Text = SourceCode.Language
        AppInfoForm.txtSourceFileName.Text = SourceCode.FileName
        AppInfoForm.txtSourceFileSize.Text = SourceCode.FileSize
        AppInfoForm.txtSourceFileHash.Text = SourceCode.FileHash
        AppInfoForm.txtSourceWebLink.Text = SourceCode.WebLink
        AppInfoForm.txtSourceContact.Text = SourceCode.Contact
        AppInfoForm.txtSourceComments.Text = SourceCode.Comments

        'Show Modifications: -----------------------------------------------------------------------------------------
        AppInfoForm.txtBaseCodeName.Text = ModificationSummary.BaseCodeName
        AppInfoForm.txtBaseCodeDescription.Text = ModificationSummary.BaseCodeDescription
        AppInfoForm.txtBaseCodeVersionMajor.Text = ModificationSummary.BaseCodeVersion.Major
        AppInfoForm.txtBaseCodeVersionMinor.Text = ModificationSummary.BaseCodeVersion.Minor
        AppInfoForm.txtBaseCodeVersionBuild.Text = ModificationSummary.BaseCodeVersion.Build
        AppInfoForm.txtBaseCodeVersionRevision.Text = ModificationSummary.BaseCodeVersion.Revision
        AppInfoForm.rtbModifications.Text = ModificationSummary.Description

        'Show the first software library record: ------------------------------------------------------
        AppInfoForm.NLibraries = Libraries.Count
        If Libraries.Count > 0 Then
            AppInfoForm.SelectedLibraryNo = 1
        Else

        End If

        'Show the ConnectOnStartup setting: ------------------------------------------------------------
        If ConnectOnStartup Then
            AppInfoForm.chkConnect.Checked = True
        Else
            AppInfoForm.chkConnect.Checked = False
        End If

    End Sub

    Private Sub ShowFileAssociation(ByVal RecordNo As Integer)
        'Show the File Association record.

        If RecordNo = 0 Then
            Exit Sub
        End If

        If RecordNo > FileAssociations.Count Then
            Exit Sub
        End If

        AppInfoForm.txtExtension.Text = FileAssociations(RecordNo - 1).Extension
        AppInfoForm.txtFileAssocationDesc.Text = FileAssociations(RecordNo - 1).Description
        AppInfoForm.txtAssnNo.Text = RecordNo

    End Sub

    Private Sub AppInfoForm_DisplayAssociation(RecordNo As Integer) Handles AppInfoForm.DisplayAssociation
        ShowFileAssociation(RecordNo)
    End Sub

    Private Sub ShowLibraryInfo(ByVal RecordNo As Integer)
        'Show the Software Library Information record.

        If RecordNo = 0 Then
            Exit Sub
        End If

        If RecordNo > Libraries.Count Then
            Exit Sub
        End If

        'Show general software library information:
        AppInfoForm.txtLibraryName.Text = Libraries(RecordNo - 1).Name
        AppInfoForm.txtLibraryDescription.Text = Libraries(RecordNo - 1).Description
        AppInfoForm.txtLibCreationDate.Text = Libraries(RecordNo - 1).CreationDate
        AppInfoForm.txtLibraryVersionMajor.Text = Libraries(RecordNo - 1).Version.Major
        AppInfoForm.txtLibraryVersionMinor.Text = Libraries(RecordNo - 1).Version.Minor
        AppInfoForm.txtLibraryVersionBuild.Text = Libraries(RecordNo - 1).Version.Build
        AppInfoForm.txtLibraryVersionRevision.Text = Libraries(RecordNo - 1).Version.Revision
        AppInfoForm.txtLibraryAuthorName.Text = Libraries(RecordNo - 1).Author.Name
        'AppInfoForm.txtLibraryAuthorDesc.Text = Libraries(RecordNo - 1).Author.Description
        AppInfoForm.rtbLibraryAuthorDesc.Text = Libraries(RecordNo - 1).Author.Description
        AppInfoForm.txtLibraryAuthorContact.Text = Libraries(RecordNo - 1).Author.Contact
        AppInfoForm.txtLibraryCopyrightNotice.Text = Libraries(RecordNo - 1).CopyrightNotice
        AppInfoForm.rtbLibraryLicenseNotice.Text = Libraries(RecordNo - 1).LicenseNotice

        'Show list of classes in the library:
        AppInfoForm.rtbClasses.Clear()
        'Add the Classes title to the page:
        AppInfoForm.rtbClasses.SelectionFont = New Drawing.Font("Arial", 11, Drawing.FontStyle.Bold Or Drawing.FontStyle.Underline)
        AppInfoForm.rtbClasses.SelectionColor = Drawing.Color.Black
        AppInfoForm.rtbClasses.SelectedText = "List of Classes" & vbCrLf & vbCrLf
        'Display the name and description of each class:
        For Each item In Libraries(RecordNo - 1).Classes
            'Class name:
            AppInfoForm.rtbClasses.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
            AppInfoForm.rtbClasses.SelectionColor = Drawing.Color.Black
            AppInfoForm.rtbClasses.SelectedText = item.Name & vbCrLf
            'Class description:
            AppInfoForm.rtbClasses.SelectionFont = New Drawing.Font("Arial", 10, Drawing.FontStyle.Regular)
            AppInfoForm.rtbClasses.SelectionCharOffset = 0
            AppInfoForm.rtbClasses.SelectedText = item.Description & vbCrLf & vbCrLf
        Next

    End Sub

    Private Sub AppInfoForm_DisplayLibrary(RecordNo As Integer) Handles AppInfoForm.DisplayLibrary
        ShowLibraryInfo(RecordNo)
    End Sub



    Private Sub AppInfoForm_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles AppInfoForm.FormClosed
        AppInfoForm = Nothing
    End Sub

    'Public Sub DisplayMessageTest()
    '    If IsNothing(AppInfoForm) Then

    '    Else
    '        AppInfoForm.rtbTrademarks.Clear() 'Clear the Trademarks richtextbox
    '        AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 12, Drawing.FontStyle.Bold Or Drawing.FontStyle.Italic Or Drawing.FontStyle.Underline)
    '        AppInfoForm.rtbTrademarks.SelectionColor = Drawing.Color.DarkBlue
    '        AppInfoForm.rtbTrademarks.SelectedText = "Test"

    '        AppInfoForm.rtbTrademarks.SelectionColor = Drawing.Color.Red
    '        AppInfoForm.rtbTrademarks.SelectedText = " Red text"
    '        AppInfoForm.rtbTrademarks.SelectedText = vbCrLf

    '        AppInfoForm.rtbTrademarks.SelectionFont = New Drawing.Font("Arial", 12, Drawing.FontStyle.Regular)
    '        AppInfoForm.rtbTrademarks.SelectionColor = Drawing.Color.Black
    '        AppInfoForm.rtbTrademarks.SelectedText = "Black text"


    '    End If
    'End Sub

    Public Function Summary() As ApplicationSummary
        'Return a summary of the Application Information

        Summary.Name = Name
        Summary.Description = Description
        Summary.CreationDate = CreationDate
        Summary.Version = Version
        Summary.Author = Author
    End Function

    Private Sub AppInfoForm_UpdateExePath() Handles AppInfoForm.UpdateExePath
        'Update the Executable path.
        RaiseEvent UpdateExePath()
        AppInfoForm.txtExecutablePath.Text = ExecutablePath
    End Sub

    Private Sub AppInfoForm_RestoreDefaults() Handles AppInfoForm.RestoreDefaults
        'Restore the default parameters.
        RaiseEvent RestoreDefaults()
        ShowInfo()
    End Sub

    Private Sub AppInfoForm_ConnectOnStartupChanged(Connect As Boolean) Handles AppInfoForm.ConnectOnStartupChanged
        ConnectOnStartup = Connect
        WriteFile() 'Save the updated settings.
    End Sub







#End Region 'Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    'Private Sub AppInfoForm_RunTest() Handles AppInfoForm.RunTest
    '    'RunTest event raised:
    '    DisplayMessageTest()
    'End Sub

#Region "Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Event ErrorMessage(ByVal Message As String)
    Event Message(ByVal Message As String)
    Event UpdateExePath()
    Event RestoreDefaults()

#End Region 'Events ------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'ApplicationInfo-------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The Version class stores application, library or project version information.
Public Class Version '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'http://programmers.stackexchange.com/questions/24987/what-exactly-is-the-build-number-in-major-minor-buildnumber-revision
    'I've never seen it written out in that form. Where I work, we are using the form 
    'MAJOR.MINOR.REVISION.BUILDNUMBER, where MAJOR is a major release (usually many new 
    'features or changes to the UI or underlying OS), MINOR is a minor release (perhaps some new 
    'features) on a previous major release, REVISION is usually a fix for a previous minor release (no 
    'new functionality), and BUILDNUMBER is incremented for each latest build of a revision.
    '
    'For example, a revision may be released to QA (quality control), and they come back with an issue 
    'which requires a change. The bug would be fixed, and released back to QA with the same 
    'REVISION number, but an incremented BUILDNUMBER.

    'https://msdn.microsoft.com/en-us/library/system.version(v=vs.110).aspx
    'Major: Assemblies with the same name but different major versions are not interchangeable. 
    '       A higher version number might indicate a major rewrite of a product where backward compatibility cannot be assumed.
    'Minor: If the name and major version number on two assemblies are the same, but the minor version number is different, 
    '       this indicates significant enhancement with the intention of backward compatibility. This higher minor version number 
    '       might indicate a point release of a product or a fully backward-compatible new version of a product.
    'Build: A difference in build number represents a recompilation of the same source. 
    '       Different build numbers might be used when the processor, platform, or compiler changes.
    'Revision: Assemblies with the same name, major, and minor version numbers but different revisions are intended to be fully interchangeable. 
    '       A higher revision number might be used in a build that fixes a security hole in a previously released assembly.



#Region " Properties - Properties used to store version information." '----------------------------------------------------------------------------------------------------------------------

    Private _source As String 'The source of the Application Version information. This is either Publish or Assembly.
    'The Publish value is set on the Publish section of the Application properties. (Double-click MyProject on Solution Explorer in Visual Studio.)
    'The Assembly value is set on the Application section of the Application properties. Click the Assembly Information button to view and edit the data.
    'The Assembly version number is only shown when the Publish number is not available.
    'Both numbers should be set when the Application is published. (Note that Publish Revision number may be set to increment automatically with each publish.)
    Property Source As String
        Get
            Return _source
        End Get
        Set(value As String)
            _source = value
        End Set
    End Property
    Private _number As String 'The Application Version number. This is a string containing the Major, Minor, Build and Revision numbers.
    Property Number As String
        Get
            Return _number
        End Get
        Set(value As String)
            _number = value
        End Set
    End Property

    Private _major As Integer = 1 '
    Property Major As Integer
        Get
            Return _major
        End Get
        Set(value As Integer)
            _major = value
        End Set
    End Property

    Private _minor As Integer = 0 '
    Property Minor As Integer
        Get
            Return _minor
        End Get
        Set(value As Integer)
            _minor = value
        End Set
    End Property

    Private _build As Integer = 1 '
    Property Build As Integer
        Get
            Return _build
        End Get
        Set(value As Integer)
            _build = value
        End Set
    End Property

    Private _revision As Integer = 0 '
    Property Revision As Integer
        Get
            Return _revision
        End Get
        Set(value As Integer)
            _revision = value
        End Set
    End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'Version --------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The Author class stores information about an Author.
Public Class Author '------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private _name As String = "" 'The name of the author
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A description of the author.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _contact As String = "" 'Author contact information.
    Property Contact As String
        Get
            Return _contact
        End Get
        Set(value As String)
            _contact = value
        End Set
    End Property

End Class 'Author ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The FileAssociation class stores the file association extension and description.
Public Class FileAssociation

    Private _extension As String 'The application file association extension.
    Property Extension As String
        Get
            Return _extension
        End Get
        Set(value As String)
            _extension = value
        End Set
    End Property

    Private _description As String 'A description of the associated file type.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

End Class


'The Copyright class stores copyright information.
Public Class Copyright '---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'http://copyright.com.au/about-copyright/website-terms/
    'http://www.copyright.org.au/acc_prod/ACC/Public_Content/Find_an_Answer.aspx
    'http://www.copyright.org.au/acc_prod/ACC/Information_Sheets/An_Introduction_to_Copyright_in_Australia.aspx

    Private _ownerName As String 'The name of the copywrite owner.
    Property OwnerName As String
        Get
            Return _ownerName
        End Get
        Set(value As String)
            _ownerName = value
            '_notice = "© " & _ownerName & " " & _publicationYear
            _notice = "© " & _publicationYear & " " & _ownerName
        End Set
    End Property

    'If you are regularly updating a work (such as material on a website, or a computer program), you can include all the years from 
    '    first publication to the present: for example, “© Gus O’Donnell 1998–2012”.

    Private _publicationYear As String 'The year of first publication or all the years from first publication to the present: for example 1998-2012. 
    Property PublicationYear As String
        Get
            Return _publicationYear
        End Get
        Set(value As String)
            _publicationYear = value
            '_notice = "© " & _ownerName & " " & _publicationYear
            _notice = "© " & _publicationYear & " " & _ownerName
        End Set
    End Property

    Private _notice As String 'The copyright notice.
    ReadOnly Property Notice As String
        Get
            Return _notice
        End Get
    End Property

End Class 'Copyright ------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The License class stores license information.
Public Class License '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'http://choosealicense.com/
    'http://www.apache.org/licenses/
    'http://opensource.org/
    'https://creativecommons.org/choose/

    Public dictLicenses As New Dictionary(Of Codes, LicenseInfo) 'Dictionary of information about each license. (Type, Abbreviation, Name)

#Region "License Properties" '===============================================================================================================================================================

    'Public Enum Types
    '    None               'No license type defined.
    '    Unknown            'Unknown license type
    '    Apache_License_2_0 'Apache License 2.0.
    '    GNU_GPL_V3_0       'GNU General Public License v3.0
    '    MIT_License        'MIT License
    '    The_Unlicense      'The Unlicense
    'End Enum

    'NOTE: Using an Enum for license names limits the characters that can be used in the name. (The following characters cannot be used: " ", "-", "&", ".")
    'However, this makes some of the coding easier.
    'The Enum will be changed from "Names" to "Codes".
    'A separate dictionary will include the license names and other information.
    'Public Enum Names
    Public Enum Codes
        None               'No license defined.
        Unknown            'Unknown license
        'Software licenses
        Apache_License_2_0 'Apache License 2.0.
        GNU_GPL_V3_0       'GNU General Public License v3.0
        MIT_License        'MIT License
        The_Unlicense      'The Unlicense
        'Data licenses:
        CC0_1_0            'Creative Commons Zero v1.0 Universal
        CC_BY_4_0          'Creative Commons Attribution 4.0
        CC_BY_SA_4_0       'Creative Commons Attribution Share Alike 4.0
    End Enum

    Public Enum Types 'The types of license.
        Software
        Data
        Any
    End Enum

    'Private _type As Types 'The software license type.
    'Property Type As Types
    '    Get
    '        Return _type
    '    End Get
    '    Set(value As Types)
    '        _type = value
    '    End Set
    'End Property

    'Private _name As Names 'The software license name (Abbreviated code).
    'Property Name As Names
    '    Get
    '        Return _name
    '    End Get
    '    Set(value As Names)
    '        _name = value
    '    End Set
    'End Property

    Private _code As Codes = Codes.Unknown 'The software license code.
    Property Code As Codes
        Get
            Return _code
        End Get
        Set(value As Codes)
            _code = value
        End Set
    End Property

    'NOTE: The license type is now stored in dictLicenses
    'USAGE: LicenseType = dictLicenses(Code).Type
    'Private _type As Types = Types.Any 'The type of license (Software, Data or Any)
    'Property Type As Types
    '    Get
    '        Return _type
    '    End Get
    '    Set(value As Types)
    '        _type = value
    '    End Set
    'End Property

    'NOTE: Fix all earlier software to handle the changed license properties!!!
    'Private _name As String = "" 'The name of the license.
    'Property Name As String
    '    Get
    '        Return _name
    '    End Get
    '    Set(value As String)
    '        _name = value
    '    End Set
    'End Property


    Private _copyrightOwnerName As String 'The name of the copywrite owner.
    Property CopyrightOwnerName As String
        Get
            Return _copyrightOwnerName
        End Get
        Set(value As String)
            _copyrightOwnerName = value
        End Set
    End Property

    'If you are regularly updating a work (such as material on a website, or a computer program), you can include all the years from 
    '    first publication to the present: for example, “© Gus O’Donnell 1998–2012”.

    Private _publicationYear As String 'The year of first publication or all the years from first publication to the present: for example 1998-2012. 
    Property PublicationYear As String
        Get
            Return _publicationYear
        End Get
        Set(value As String)
            _publicationYear = value
        End Set
    End Property

    'NOTE: The license notice can be generated by the GetNotice function.
    Private _notice As String 'License notice.
    Property Notice As String
        Get
            Return _notice
        End Get
        Set(value As String)
            _notice = value
        End Set
    End Property

    'LICENSE NOTICE EXAMPLE:
    'Copyright [yyyy] [name of copyright owner]
    '
    'Licensed under the Apache License, Version 2.0 (the "License");
    'you may not use this file except in compliance with the License.
    'You may obtain a copy of the License at
    '
    '    http://www.apache.org/licenses/LICENSE-2.0
    '
    'Unless required by applicable law or agreed to in writing, software
    'distributed under the License is distributed on an "AS IS" BASIS,
    'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    'See the License for the specific language governing permissions and
    'limitations under the License.

    'NOTE: The license text can be generated by the GetText function.
    Private _text As String 'The license text.
    Property Text As String
        Get
            Return _text
        End Get
        Set(value As String)
            _text = value
        End Set
    End Property

#End Region 'License Properties -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "License Methods" '------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Function GetNotice()
        'Returns the License Notice corresponding to the License Code in the Code property.

        Select Case Code
            Case Codes.Apache_License_2_0
                Return ApacheLicenseNotice()
            Case Codes.CC0_1_0
                Return "To Do"
            Case Codes.CC_BY_4_0
                Return "To Do"
            Case Codes.CC_BY_SA_4_0
                Return "To Do"
            Case Codes.GNU_GPL_V3_0
                Return "To Do"
            Case Codes.MIT_License
                Return MITLicenseNotice()
            Case Codes.None
                Return "No license specified"
            Case Codes.The_Unlicense
                Return UnLicenseNotice()
            Case Codes.Unknown
                Return "Unknown license"
        End Select
    End Function

    Public Function GetText()
        'Returns the License Text corresponding to the License Code in the Code property.

        Select Case Code
            Case Codes.Apache_License_2_0
                Return ApacheLicenseText()
            Case Codes.CC0_1_0
                Return "To Do"
            Case Codes.CC_BY_4_0
                Return "To Do"
            Case Codes.CC_BY_SA_4_0
                Return "To Do"
            Case Codes.GNU_GPL_V3_0
                Return "To Do"
            Case Codes.MIT_License
                Return MITLicenseText()
            Case Codes.None
                Return "No license specified"
            Case Codes.The_Unlicense
                Return UnLicenseText()
            Case Codes.Unknown
                Return "Unknown license"
        End Select
    End Function

    Public Function ApacheLicenseNotice() As String
        'Returns the Apache License notice.
        '    The copyright year is obtained from the PublicationYear property.
        '    The copyright owner is obtained from the CopyrightOwnerName property.

        Dim sb As New System.Text.StringBuilder
        sb.Append("Copyright " & PublicationYear & " " & CopyrightOwnerName & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf)
        sb.Append("you may not use this file except in compliance with the License." & vbCrLf)
        sb.Append("You may obtain a copy of the License at" & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("Unless required by applicable law or agreed to in writing, software" & vbCrLf)
        sb.Append("distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf)
        sb.Append("WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf)
        sb.Append("See the License for the specific language governing permissions and" & vbCrLf)
        sb.Append("limitations under the License." & vbCrLf)

        Return sb.ToString

    End Function

    Public Function ApacheLicenseText() As String
        'Returns the Apache License text.
        '    http://choosealicense.com/licenses/apache-2.0/
        '

        Dim sb As New System.Text.StringBuilder
        sb.Append("                                 Apache License" & vbCrLf)
        sb.Append("                           Version 2.0, January 2004" & vbCrLf)
        sb.Append("                        http://www.apache.org/licenses/" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   TERMS AND CONDITIONS FOR USE, REPRODUCTION, AND DISTRIBUTION" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   1. Definitions." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""License"" shall mean the terms and conditions for use, reproduction," & vbCrLf)
        sb.Append("      and distribution as defined by Sections 1 through 9 of this document." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Licensor"" shall mean the copyright owner or entity authorized by" & vbCrLf)
        sb.Append("      the copyright owner that is granting the License." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Legal Entity"" shall mean the union of the acting entity and all" & vbCrLf)
        sb.Append("      other entities that control, are controlled by, or are under common" & vbCrLf)
        sb.Append("      control with that entity. For the purposes of this definition," & vbCrLf)
        sb.Append("      ""control"" means (i) the power, direct or indirect, to cause the" & vbCrLf)
        sb.Append("      direction or management of such entity, whether by contract or" & vbCrLf)
        sb.Append("      otherwise, or (ii) ownership of fifty percent (50%) or more of the" & vbCrLf)
        sb.Append("      outstanding shares, or (iii) beneficial ownership of such entity." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""You"" (or ""Your"") shall mean an individual or Legal Entity" & vbCrLf)
        sb.Append("      exercising permissions granted by this License." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Source"" form shall mean the preferred form for making modifications," & vbCrLf)
        sb.Append("      including but not limited to software source code, documentation" & vbCrLf)
        sb.Append("      source, and configuration files." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Object"" form shall mean any form resulting from mechanical" & vbCrLf)
        sb.Append("      transformation or translation of a Source form, including but" & vbCrLf)
        sb.Append("      not limited to compiled object code, generated documentation," & vbCrLf)
        sb.Append("      and conversions to other media types." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Work"" shall mean the work of authorship, whether in Source or" & vbCrLf)
        sb.Append("      Object form, made available under the License, as indicated by a" & vbCrLf)
        sb.Append("      copyright notice that is included in or attached to the work" & vbCrLf)
        sb.Append("      (an example is provided in the Appendix below)." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Derivative Works"" shall mean any work, whether in Source or Object" & vbCrLf)
        sb.Append("      form, that is based on (or derived from) the Work and for which the" & vbCrLf)
        sb.Append("      editorial revisions, annotations, elaborations, or other modifications" & vbCrLf)
        sb.Append("      represent, as a whole, an original work of authorship. For the purposes" & vbCrLf)
        sb.Append("      of this License, Derivative Works shall not include works that remain" & vbCrLf)
        sb.Append("      separable from, or merely link (or bind by name) to the interfaces of," & vbCrLf)
        sb.Append("      the Work and Derivative Works thereof." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Contribution"" shall mean any work of authorship, including" & vbCrLf)
        sb.Append("      the original version of the Work and any modifications or additions" & vbCrLf)
        sb.Append("      to that Work or Derivative Works thereof, that is intentionally" & vbCrLf)
        sb.Append("      submitted to Licensor for inclusion in the Work by the copyright owner" & vbCrLf)
        sb.Append("      or by an individual or Legal Entity authorized to submit on behalf of" & vbCrLf)
        sb.Append("      the copyright owner. For the purposes of this definition, ""submitted""" & vbCrLf)
        sb.Append("      means any form of electronic, verbal, or written communication sent" & vbCrLf)
        sb.Append("      to the Licensor or its representatives, including but not limited to" & vbCrLf)
        sb.Append("      communication on electronic mailing lists, source code control systems," & vbCrLf)
        sb.Append("      and issue tracking systems that are managed by, or on behalf of, the" & vbCrLf)
        sb.Append("      Licensor for the purpose of discussing and improving the Work, but" & vbCrLf)
        sb.Append("      excluding communication that is conspicuously marked or otherwise" & vbCrLf)
        sb.Append("      designated in writing by the copyright owner as ""Not a Contribution.""" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      ""Contributor"" shall mean Licensor and any individual or Legal Entity" & vbCrLf)
        sb.Append("      on behalf of whom a Contribution has been received by Licensor and" & vbCrLf)
        sb.Append("      subsequently incorporated within the Work." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   2. Grant of Copyright License. Subject to the terms and conditions of" & vbCrLf)
        sb.Append("      this License, each Contributor hereby grants to You a perpetual," & vbCrLf)
        sb.Append("      worldwide, non-exclusive, no-charge, royalty-free, irrevocable" & vbCrLf)
        sb.Append("      copyright license to reproduce, prepare Derivative Works of," & vbCrLf)
        sb.Append("      publicly display, publicly perform, sublicense, and distribute the" & vbCrLf)
        sb.Append("      Work and such Derivative Works in Source or Object form." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   3. Grant of Patent License. Subject to the terms and conditions of" & vbCrLf)
        sb.Append("      this License, each Contributor hereby grants to You a perpetual," & vbCrLf)
        sb.Append("      worldwide, non-exclusive, no-charge, royalty-free, irrevocable" & vbCrLf)
        sb.Append("      (except as stated in this section) patent license to make, have made," & vbCrLf)
        sb.Append("      use, offer to sell, sell, import, and otherwise transfer the Work," & vbCrLf)
        sb.Append("      where such license applies only to those patent claims licensable" & vbCrLf)
        sb.Append("      by such Contributor that are necessarily infringed by their" & vbCrLf)
        sb.Append("      Contribution(s) alone or by combination of their Contribution(s)" & vbCrLf)
        sb.Append("      with the Work to which such Contribution(s) was submitted. If You" & vbCrLf)
        sb.Append("      institute patent litigation against any entity (including a" & vbCrLf)
        sb.Append("      cross-claim or counterclaim in a lawsuit) alleging that the Work" & vbCrLf)
        sb.Append("      or a Contribution incorporated within the Work constitutes direct" & vbCrLf)
        sb.Append("      or contributory patent infringement, then any patent licenses" & vbCrLf)
        sb.Append("      granted to You under this License for that Work shall terminate" & vbCrLf)
        sb.Append("      as of the date such litigation is filed." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   4. Redistribution. You may reproduce and distribute copies of the" & vbCrLf)
        sb.Append("      Work or Derivative Works thereof in any medium, with or without" & vbCrLf)
        sb.Append("      modifications, and in Source or Object form, provided that You" & vbCrLf)
        sb.Append("      meet the following conditions:" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      (a) You must give any other recipients of the Work or" & vbCrLf)
        sb.Append("          Derivative Works a copy of this License; and" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      (b) You must cause any modified files to carry prominent notices" & vbCrLf)
        sb.Append("          stating that You changed the files; and" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      (c) You must retain, in the Source form of any Derivative Works" & vbCrLf)
        sb.Append("          that You distribute, all copyright, patent, trademark, and" & vbCrLf)
        sb.Append("          attribution notices from the Source form of the Work," & vbCrLf)
        sb.Append("          excluding those notices that do not pertain to any part of" & vbCrLf)
        sb.Append("          the Derivative Works; and" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      (d) If the Work includes a ""NOTICE"" text file as part of its" & vbCrLf)
        sb.Append("          distribution, then any Derivative Works that You distribute must" & vbCrLf)
        sb.Append("          include a readable copy of the attribution notices contained" & vbCrLf)
        sb.Append("          within such NOTICE file, excluding those notices that do not" & vbCrLf)
        sb.Append("          pertain to any part of the Derivative Works, in at least one" & vbCrLf)
        sb.Append("          of the following places: within a NOTICE text file distributed" & vbCrLf)
        sb.Append("          as part of the Derivative Works; within the Source form or" & vbCrLf)
        sb.Append("          documentation, if provided along with the Derivative Works; or," & vbCrLf)
        sb.Append("          within a display generated by the Derivative Works, if and" & vbCrLf)
        sb.Append("          wherever such third-party notices normally appear. The contents" & vbCrLf)
        sb.Append("          of the NOTICE file are for informational purposes only and" & vbCrLf)
        sb.Append("          do not modify the License. You may add Your own attribution" & vbCrLf)
        sb.Append("          notices within Derivative Works that You distribute, alongside" & vbCrLf)
        sb.Append("          or as an addendum to the NOTICE text from the Work, provided" & vbCrLf)
        sb.Append("          that such additional attribution notices cannot be construed" & vbCrLf)
        sb.Append("          as modifying the License." & vbCrLf)
        sb.Append("      You may add Your own copyright statement to Your modifications and" & vbCrLf)
        sb.Append("      may provide additional or different license terms and conditions" & vbCrLf)
        sb.Append("      for use, reproduction, or distribution of Your modifications, or" & vbCrLf)
        sb.Append("      for any such Derivative Works as a whole, provided Your use," & vbCrLf)
        sb.Append("      reproduction, and distribution of the Work otherwise complies with" & vbCrLf)
        sb.Append("      the conditions stated in this License." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   5. Submission of Contributions. Unless You explicitly state otherwise," & vbCrLf)
        sb.Append("      any Contribution intentionally submitted for inclusion in the Work" & vbCrLf)
        sb.Append("      by You to the Licensor shall be under the terms and conditions of" & vbCrLf)
        sb.Append("      this License, without any additional terms or conditions." & vbCrLf)
        sb.Append("      Notwithstanding the above, nothing herein shall supersede or modify" & vbCrLf)
        sb.Append("      the terms of any separate license agreement you may have executed" & vbCrLf)
        sb.Append("      with Licensor regarding such Contributions." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   6. Trademarks. This License does not grant permission to use the trade" & vbCrLf)
        sb.Append("      names, trademarks, service marks, or product names of the Licensor," & vbCrLf)
        sb.Append("      except as required for reasonable and customary use in describing the" & vbCrLf)
        sb.Append("      origin of the Work and reproducing the content of the NOTICE file." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   7. Disclaimer of Warranty. Unless required by applicable law or" & vbCrLf)
        sb.Append("      agreed to in writing, Licensor provides the Work (and each" & vbCrLf)
        sb.Append("      Contributor provides its Contributions) on an ""AS IS"" BASIS," & vbCrLf)
        sb.Append("      WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or" & vbCrLf)
        sb.Append("      implied, including, without limitation, any warranties or conditions" & vbCrLf)
        sb.Append("      of TITLE, NON-INFRINGEMENT, MERCHANTABILITY, or FITNESS FOR A" & vbCrLf)
        sb.Append("      PARTICULAR PURPOSE. You are solely responsible for determining the" & vbCrLf)
        sb.Append("      appropriateness of using or redistributing the Work and assume any" & vbCrLf)
        sb.Append("      risks associated with Your exercise of permissions under this License." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   8. Limitation of Liability. In no event and under no legal theory," & vbCrLf)
        sb.Append("      whether in tort (including negligence), contract, or otherwise," & vbCrLf)
        sb.Append("      unless required by applicable law (such as deliberate and grossly" & vbCrLf)
        sb.Append("      negligent acts) or agreed to in writing, shall any Contributor be" & vbCrLf)
        sb.Append("      liable to You for damages, including any direct, indirect, special," & vbCrLf)
        sb.Append("      incidental, or consequential damages of any character arising as a" & vbCrLf)
        sb.Append("      result of this License or out of the use or inability to use the" & vbCrLf)
        sb.Append("      Work (including but not limited to damages for loss of goodwill," & vbCrLf)
        sb.Append("      work stoppage, computer failure or malfunction, or any and all" & vbCrLf)
        sb.Append("      other commercial damages or losses), even if such Contributor" & vbCrLf)
        sb.Append("      has been advised of the possibility of such damages." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   9. Accepting Warranty or Additional Liability. While redistributing" & vbCrLf)
        sb.Append("      the Work or Derivative Works thereof, You may choose to offer," & vbCrLf)
        sb.Append("      and charge a fee for, acceptance of support, warranty, indemnity," & vbCrLf)
        sb.Append("      or other liability obligations and/or rights consistent with this" & vbCrLf)
        sb.Append("      License. However, in accepting such obligations, You may act only" & vbCrLf)
        sb.Append("      on Your own behalf and on Your sole responsibility, not on behalf" & vbCrLf)
        sb.Append("      of any other Contributor, and only if You agree to indemnify," & vbCrLf)
        sb.Append("      defend, and hold each Contributor harmless for any liability" & vbCrLf)
        sb.Append("      incurred by, or claims asserted against, such Contributor by reason" & vbCrLf)
        sb.Append("      of your accepting any such warranty or additional liability." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   END OF TERMS AND CONDITIONS" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   APPENDIX: How to apply the Apache License to your work." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("      To apply the Apache License to your work, attach the following" & vbCrLf)
        sb.Append("      boilerplate notice, with the fields enclosed by brackets ""{}""" & vbCrLf)
        sb.Append("      replaced with your own identifying information. (Don't include" & vbCrLf)
        sb.Append("      the brackets!)  The text should be enclosed in the appropriate" & vbCrLf)
        sb.Append("      comment syntax for the file format. We also recommend that a" & vbCrLf)
        sb.Append("      file or class name and description of purpose be included on the" & vbCrLf)
        sb.Append("      same ""printed page"" as the copyright notice for easier" & vbCrLf)
        sb.Append("      identification within third-party archives." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   Copyright {yyyy} {name of copyright owner}" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf)
        sb.Append("   you may not use this file except in compliance with the License." & vbCrLf)
        sb.Append("   You may obtain a copy of the License at" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("       http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("   Unless required by applicable law or agreed to in writing, software" & vbCrLf)
        sb.Append("   distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf)
        sb.Append("   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf)
        sb.Append("   See the License for the specific language governing permissions and" & vbCrLf)
        sb.Append("   limitations under the License." & vbCrLf)
        sb.Append("" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function MITLicenseNotice() As String
        'Return the MIT License notice.
        '    The copyright year is obtained from the PublicationYear property.
        '    The copyright owner is obtained from the CopyrightOwnerName property.

        Dim sb As New System.Text.StringBuilder
        sb.Append("Copyright " & PublicationYear & " " & CopyrightOwnerName & vbCrLf)
        sb.Append("Licensed under the MIT License. " & vbCrLf)

        Return sb.ToString

    End Function

    Public Function MITLicenseText() As String
        'Returns the MIT License text.
        '    http://choosealicense.com/licenses/mit/
        '

        Dim sb As New System.Text.StringBuilder
        sb.Append("MIT License" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("Copyright (c) " & PublicationYear & " " & CopyrightOwnerName & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("Permission is hereby granted, free of charge, to any person obtaining a copy" & vbCrLf)
        sb.Append("of this software and associated documentation files (the ""Software""), to deal" & vbCrLf)
        sb.Append("in the Software without restriction, including without limitation the rights" & vbCrLf)
        sb.Append("to use, copy, modify, merge, publish, distribute, sublicense, and/or sell" & vbCrLf)
        sb.Append("copies of the Software, and to permit persons to whom the Software is" & vbCrLf)
        sb.Append("furnished to do so, subject to the following conditions:" & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("The above copyright notice and this permission notice shall be included in all" & vbCrLf)
        sb.Append("copies or substantial portions of the Software." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR" & vbCrLf)
        sb.Append("IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY," & vbCrLf)
        sb.Append("FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE" & vbCrLf)
        sb.Append("AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER" & vbCrLf)
        sb.Append("LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM," & vbCrLf)
        sb.Append("OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE" & vbCrLf)
        sb.Append("SOFTWARE." & vbCrLf)

        Return sb.ToString

    End Function

    Public Function UnLicenseNotice() As String
        'Return the UnLicense text.

        Dim sb As New System.Text.StringBuilder
        sb.Append("This is free and unencumbered software released into the public domain." & vbCrLf)

        Return sb.ToString

    End Function

    Public Function UnLicenseText() As String
        'Return the UnLicense notice.

        Dim sb As New System.Text.StringBuilder
        sb.Append("This is free and unencumbered software released into the public domain." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("Anyone is free to copy, modify, publish, use, compile, sell, or" & vbCrLf)
        sb.Append("distribute this software, either in source code form or as a compiled" & vbCrLf)
        sb.Append("binary, for any purpose, commercial or non-commercial, and by any" & vbCrLf)
        sb.Append("means." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("In jurisdictions that recognize copyright laws, the author or authors" & vbCrLf)
        sb.Append("of this software dedicate any and all copyright interest in the" & vbCrLf)
        sb.Append("software to the public domain. We make this dedication for the benefit" & vbCrLf)
        sb.Append("of the public at large and to the detriment of our heirs and" & vbCrLf)
        sb.Append("successors. We intend this dedication to be an overt act of" & vbCrLf)
        sb.Append("relinquishment in perpetuity of all present and future rights to this" & vbCrLf)
        sb.Append("software under copyright law." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND," & vbCrLf)
        sb.Append("EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF" & vbCrLf)
        sb.Append("MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT." & vbCrLf)
        sb.Append("IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR" & vbCrLf)
        sb.Append("OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE," & vbCrLf)
        sb.Append("ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR" & vbCrLf)
        sb.Append("OTHER DEALINGS IN THE SOFTWARE." & vbCrLf)
        sb.Append("" & vbCrLf)
        sb.Append("For more information, please refer to <http://unlicense.org>" & vbCrLf)

        Return sb.ToString

    End Function



    Public Sub LoadDefaultLicenseDictionary()
        'Constructs the default license dictionary.

        dictLicenses.Clear() 'Clear the license dictionary.

        'Add the None license (No license specified.)
        dictLicenses.Add(Codes.None, New LicenseInfo)
        dictLicenses(Codes.None).Name = "None"
        dictLicenses(Codes.None).Abbreviation = "None"
        dictLicenses(Codes.None).Type = Types.Any
        dictLicenses(Codes.None).Description = "No license specified."

        'Add the Unknown license (License not known)
        dictLicenses.Add(Codes.Unknown, New LicenseInfo)
        dictLicenses(Codes.Unknown).Name = "Unknown"
        dictLicenses(Codes.Unknown).Abbreviation = "Unknown"
        dictLicenses(Codes.Unknown).Type = Types.Any
        dictLicenses(Codes.Unknown).Description = "License not known."

        'Add the Apache License 2.0
        dictLicenses.Add(Codes.Apache_License_2_0, New LicenseInfo)
        dictLicenses(Codes.Apache_License_2_0).Name = "Apache License 2.0"
        dictLicenses(Codes.Apache_License_2_0).Abbreviation = "Apache License 2.0"
        dictLicenses(Codes.Apache_License_2_0).Type = Types.Software
        dictLicenses(Codes.Apache_License_2_0).Description = "A permissive license whose main conditions require preservation of copyright and license notices. Contributors provide an express grant of patent rights. Licensed works, modifications, and larger works may be distributed under different terms and without source code. License description provided by GitHub, Inc. at https://choosealicense.com/ under the Creative Commons Atribution 3.0 Unported License: https://creativecommons.org/licenses/by/3.0/"

        'Add the GNU General Public License v3.0
        dictLicenses.Add(Codes.GNU_GPL_V3_0, New LicenseInfo)
        dictLicenses(Codes.GNU_GPL_V3_0).Name = "GNU General Public License v3.0"
        dictLicenses(Codes.GNU_GPL_V3_0).Abbreviation = "GNU GPLv3"
        dictLicenses(Codes.GNU_GPL_V3_0).Type = Types.Software
        dictLicenses(Codes.GNU_GPL_V3_0).Description = "Permissions of this strong copyleft license are conditioned on making available complete source code of licensed works and modifications, which include larger works using a licensed work, under the same license. Copyright and license notices must be preserved. Contributors provide an express grant of patent rights. License description provided by GitHub, Inc. at https://choosealicense.com/ under the Creative Commons Atribution 3.0 Unported License: https://creativecommons.org/licenses/by/3.0/"

        'Add the MIT License
        dictLicenses.Add(Codes.MIT_License, New LicenseInfo)
        dictLicenses(Codes.MIT_License).Name = "MIT License"
        dictLicenses(Codes.MIT_License).Abbreviation = "MIT License"
        dictLicenses(Codes.MIT_License).Type = Types.Software
        dictLicenses(Codes.MIT_License).Description = "A short and simple permissive license with conditions only requiring preservation of copyright and license notices. Licensed works, modifications, and larger works may be distributed under different terms and without source code. License description provided by GitHub, Inc. at https://choosealicense.com/ under the Creative Commons Atribution 3.0 Unported License: https://creativecommons.org/licenses/by/3.0/"

        'Add the Unicense
        dictLicenses.Add(Codes.The_Unlicense, New LicenseInfo)
        dictLicenses(Codes.The_Unlicense).Name = "The Unlicense"
        dictLicenses(Codes.The_Unlicense).Abbreviation = "The Unlicense"
        dictLicenses(Codes.The_Unlicense).Type = Types.Software
        dictLicenses(Codes.The_Unlicense).Description = "A license with no conditions whatsoever which dedicates works to the public domain. Unlicensed works, modifications, and larger works may be distributed under different terms and without source code. License description provided by GitHub, Inc. at https://choosealicense.com/ under the Creative Commons Atribution 3.0 Unported License: https://creativecommons.org/licenses/by/3.0/"

        'Add the CC0 1.0 Universal
        dictLicenses.Add(Codes.CC0_1_0, New LicenseInfo)
        dictLicenses(Codes.CC0_1_0).Name = "Creative Commons Zero v1.0 Universal"
        dictLicenses(Codes.CC0_1_0).Abbreviation = "CC0 1.0 Universal"
        dictLicenses(Codes.CC0_1_0).Type = Types.Data
        dictLicenses(Codes.CC0_1_0).Description = "The Creative Commons CC0 Public Domain Dedication waives copyright interest in any a work you've created and dedicates it to the world-wide public domain. Use CC0 to opt out of copyright entirely and ensure your work has the widest reach. As with the Unlicense and typical software licenses, CC0 disclaims warranties. CC0 is very similar to the Unlicense. License description provided by GitHub, Inc. at https://choosealicense.com/ under the Creative Commons Atribution 3.0 Unported License: https://creativecommons.org/licenses/by/3.0/"

        'Add the CC-BY-4.0
        dictLicenses.Add(Codes.CC_BY_4_0, New LicenseInfo)
        dictLicenses(Codes.CC_BY_4_0).Name = "Creative Commons Attribution 4.0"
        dictLicenses(Codes.CC_BY_4_0).Abbreviation = "CC-BY-4.0"
        dictLicenses(Codes.CC_BY_4_0).Type = Types.Data
        dictLicenses(Codes.CC_BY_4_0).Description = "Permits almost any use subject to providing credit and license notice. Frequently used for media assets and educational materials. The most common license for Open Access scientific publications. Not recommended for software. License description provided by GitHub, Inc. at https://choosealicense.com/ under the Creative Commons Atribution 3.0 Unported License: https://creativecommons.org/licenses/by/3.0/"

        'Add the CC-BY-SA-4.0
        dictLicenses.Add(Codes.CC_BY_SA_4_0, New LicenseInfo)
        dictLicenses(Codes.CC_BY_SA_4_0).Name = "Creative Commons Attribution Share Alike 4.0"
        dictLicenses(Codes.CC_BY_SA_4_0).Abbreviation = "CC-BY-SA-4.0"
        dictLicenses(Codes.CC_BY_SA_4_0).Type = Types.Data
        dictLicenses(Codes.CC_BY_SA_4_0).Description = "Similar to CC-BY-4.0 but requires derivatives be distributed under the same or a similar, compatible license. Frequently used for media assets and educational materials. A previous version is the default license for Wikipedia and other Wikimedia projects. Not recommended for software. License description provided by GitHub, Inc. at https://choosealicense.com/ under the Creative Commons Atribution 3.0 Unported License: https://creativecommons.org/licenses/by/3.0/"
    End Sub


#End Region 'License Methods ----------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'License --------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Class LicenseInfo
    'Information about software and data licenses.
    'This is used in the dictionary of licenses in the License class.

    Private _type As License.Types 'The type of license (Software, Data or Any).
    Property Type As License.Types
        Get
            Return _type
        End Get
        Set(value As License.Types)
            _type = value
        End Set
    End Property

    Private _abbreviation As String = "" 'The abbreviated name of the license. (eg CC BY-SA)
    Property Abbreviation As String
        Get
            Return _abbreviation
        End Get
        Set(value As String)
            _abbreviation = value
        End Set
    End Property

    Private _name As String = "" 'The name of the license. (eg Creative Commons Attribution-ShareAlike)
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = "" 'A desription of the license.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    'NOTE: This is now generated within the License class.
    'Public Function LicenseNotice() As String
    '    'Returns the license notice corresponding to the license specified in the Abbreviation property.

    '    'For testing only: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '    'Return ""

    '    Select Case Abbreviation
    '        Case "None"
    '            Return "No license specified."
    '        Case "Unknown"
    '            Return "License is unknown."
    '        Case "Apache License 2.0"

    '            Return License.ApacheLicenseText
    '        Case "GNU GPLv3"

    '        Case "MIT License"

    '        Case "The Unlicense"

    '        Case "CC0 1.0 Universal"

    '        Case "CC-BY-4.0"

    '        Case "CC-BY-SA-4.0"

    '    End Select

    'End Function

End Class


'The SourceCode class stores information about the source code for the application.
Public Class SourceCode '--------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "SourceCode Properties" '------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private _language 'The source code programming language.
    Property Language As String
        Get
            Return _language
        End Get
        Set(value As String)
            _language = value
        End Set
    End Property

    Private _fileName As String 'The source code distribution file name.
    Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
        End Set
    End Property

    Private _fileSize As Integer ' The source code distribution file size in bytes.
    Property FileSize As Integer
        Get
            Return _fileSize
        End Get
        Set(value As Integer)
            _fileSize = value
        End Set
    End Property

    Private _fileHash As String 'The hashcode of the source code distribution file.
    Property FileHash As String
        Get
            Return _fileHash
        End Get
        Set(value As String)
            _fileHash = value
        End Set
    End Property

    Private _webLink As String 'The link to a website containing the source code distribution file.
    Property WebLink As String
        Get
            Return _webLink
        End Get
        Set(value As String)
            _webLink = value
        End Set
    End Property

    Private _contact As String 'Contact information for enquiries about the source code.
    Property Contact As String
        Get
            Return _contact
        End Get
        Set(value As String)
            _contact = value
        End Set
    End Property

    Private _comments As String 'Comments about the source code.
    Property Comments As String
        Get
            Return _comments
        End Get
        Set(value As String)
            _comments = value
        End Set
    End Property

#End Region 'SourceCode Properties ----------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'SourceCode -----------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The Usage class stores information about application or project usage.
Public Class Usage '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Usage Declarations" '-----------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public SaveLocn As New FileLocation 'The location on disk where the Usage information is saved.

#End Region 'Usage Declarations ---------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region "Usage Properties" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------


    Private _firstUsed As DateTime 'The date and time that the application, project or data was first used.
    Property FirstUsed As DateTime
        Get
            Return _firstUsed
        End Get
        Set(value As DateTime)
            _firstUsed = value
        End Set
    End Property

    Private _lastUsed As DateTime 'The date and time that the application, project or data was last used.
    'This value is saved at the end of the current usage.
    'The class containing the Usage object contains the method to save the value.
    Property LastUsed As DateTime
        Get
            Return _lastUsed
        End Get
        Set(value As DateTime)
            _lastUsed = value
        End Set
    End Property

    Private _startTime As DateTime 'The start time of the current usage.
    'This value is set to the Now value at the start of each usage.
    Property StartTime As DateTime
        Get
            Return _startTime
        End Get
        Set(value As DateTime)
            _startTime = value
        End Set
    End Property

    Private _currentDuration As TimeSpan 'The duration of the current usage. Units can be TotalDays (CurrentDuration.TotalDays), TotalHours, TotalMinutes, TotalSeconds or TotalMilliseconds.
    ReadOnly Property CurrentDuration As TimeSpan
        Get
            _currentDuration = Now.Subtract(StartTime)
            Return _currentDuration
        End Get
    End Property

    Private _lastDuration As TimeSpan 'The duration of the last usage.
    'This value is saved at the end of the current usage. The CurrentDuration value becomes the saved LastDuration value.
    'The class containing the Usage object contains the method to save the updated value.
    Property LastDuration As TimeSpan
        Get
            Return _lastDuration
        End Get
        Set(value As TimeSpan)
            _lastDuration = value
        End Set
    End Property

    Private _totalDuration As TimeSpan 'The total duration of usage. (Previous usage + current duration)
    'This value can be read at any time.
    'The updated value is saved at the end of the current usage.
    'The class containing the Usage object contains the method to save the updated value.
    Property TotalDuration As TimeSpan
        Get
            Return _totalDuration.Add(_currentDuration) 'NOTE: _totalDuration is not changed
        End Get
        Set(value As TimeSpan)
            _totalDuration = value
        End Set
    End Property

#End Region 'Usage Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region "Usage Methods" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub SaveUsageInfo()
        'Save the Usage information as an XML file named Usage_Info.xml in the location SaveLocn.

        'UPDATE GET THE OLD TOTALDURATION AND ADD THE CURRENTDURATION TO PRODUCE THE NEW TOTALDURATION
        '  This will produce the correct total duration for the case where multiple instances of the application or project are used at the same time.

        Dim OldUsageXml As System.Xml.Linq.XDocument

        SaveLocn.ReadXmlData("Usage_Info.xml", OldUsageXml)

        If OldUsageXml Is Nothing Then
            Dim UsageXml = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Usage Information-->
                           <!---->
                           <Usage>
                               <FirstUsed><%= Format(FirstUsed, "d-MMM-yyyy H:mm:ss") %></FirstUsed>
                               <LastUsed><%= Format(Now, "d-MMM-yyyy H:mm:ss") %></LastUsed>
                               <LastDuration><%= CurrentDuration.ToString("c") %></LastDuration>
                               <TotalDuration><%= TotalDuration.ToString("c") %></TotalDuration>
                           </Usage>

            '<LastDuration><%= CurrentDuration.ToString %></LastDuration>
            '<TotalDuration><%= TotalDuration.ToString %></TotalDuration>

            'The "c" format specifies a constant (invariant) format for the TimeSpan string.
            'The "G" format spacified a General Long Format: [-]d’:’hh’:’mm’:’ss.fffffff
            'The "[-]{ d | [d.]hh:mm[:ss[.ff]] }" format is not valid
            'The "d.hh:mm:ss.ff" format is not valid
            '"{d\\.hh\\:mm\\:ss\\.ff}" not valid
            '"d'.'hh':'mm':'ss'.'ff" 
            '"d':'hh':'mm':'ss':'ff" not valid
            '<LastDuration><%= CurrentDuration.ToString("d'.'hh':'mm':'ss'.'ff") %></LastDuration>
            '<TotalDuration><%= TotalDuration.ToString("d'.'hh':'mm':'ss'.'ff") %></TotalDuration>
            SaveLocn.SaveXmlData("Usage_Info.xml", UsageXml)
        Else
            Dim OldTotalDuration As TimeSpan = System.TimeSpan.Parse(OldUsageXml.<Usage>.<TotalDuration>.Value)
            Dim NewTotalDuration As TimeSpan = OldTotalDuration.Add(CurrentDuration)
            Dim UsageXml = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Usage Information-->
                           <!---->
                           <Usage>
                               <FirstUsed><%= Format(FirstUsed, "d-MMM-yyyy H:mm:ss") %></FirstUsed>
                               <LastUsed><%= Format(Now, "d-MMM-yyyy H:mm:ss") %></LastUsed>
                               <LastDuration><%= CurrentDuration.ToString("c") %></LastDuration>
                               <TotalDuration><%= NewTotalDuration.ToString("c") %></TotalDuration>
                           </Usage>
            SaveLocn.SaveXmlData("Usage_Info.xml", UsageXml)
        End If

        '<LastDuration><%= CurrentDuration.ToString %></LastDuration>
        '<TotalDuration><%= NewTotalDuration.ToString %></TotalDuration>

    End Sub

    Public Sub RestoreUsageInfo()
        'Restore the Usage information from an XML file named Usage_Info.xml in the location SaveLocn.

        Dim UsageXml As System.Xml.Linq.XDocument
        SaveLocn.ReadXmlData("Usage_Info.xml", UsageXml)

        If UsageXml Is Nothing Then
            'No usage information has been recorded.
            'Create new usage data:
            FirstUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            StartTime = Format(Now, "d-MMM-yyyy H:mm:ss")
            LastDuration = System.TimeSpan.Zero
            TotalDuration = System.TimeSpan.Zero
        Else
            'Get usage information from the Xml file:
            Dim culture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture
            If UsageXml.<Usage>.<FirstUsed>.Value = Nothing Then
                FirstUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            Else
                FirstUsed = UsageXml.<Usage>.<FirstUsed>.Value
            End If
            If UsageXml.<Usage>.<LastUsed>.Value = Nothing Then
                LastUsed = Format(Now, "d-MMM-yyyy H:mm:ss")
            Else
                LastUsed = UsageXml.<Usage>.<LastUsed>.Value
            End If
            If UsageXml.<Usage>.<LastDuration>.Value = Nothing Then
                LastDuration = System.TimeSpan.Zero
            Else
                'LastDuration = System.TimeSpan.Parse(UsageXml.<Usage>.<LastDuration>.Value)
                LastDuration = System.TimeSpan.ParseExact(UsageXml.<Usage>.<LastDuration>.Value, "c", System.Globalization.DateTimeFormatInfo.InvariantInfo)


                '[ws][-]{ d | [d.]hh:mm[:ss[.ff]] }[ws] TimeSpan format required.
                'LastDuration = System.TimeSpan.ParseExact(UsageXml.<Usage>.<LastDuration>.Value, "c", culture)
                'LastDuration = System.TimeSpan.TryParse(UsageXml.<Usage>.<LastDuration>.Value, )
                'If System.TimeSpan.TryParse(UsageXml.<Usage>.<LastDuration>.Value, LastDuration) Then
                '    Debug.Print("Parsed " & UsageXml.<Usage>.<LastDuration>.Value & " OK" & vbCrLf)
                'Else
                '    Debug.Print("Error parsing " & UsageXml.<Usage>.<LastDuration>.Value & vbCrLf)
                'End If
            End If
            If UsageXml.<Usage>.<TotalDuration>.Value = Nothing Then
                TotalDuration = System.TimeSpan.Zero
            Else
                'TotalDuration = System.TimeSpan.Parse(UsageXml.<Usage>.<TotalDuration>.Value)
                TotalDuration = System.TimeSpan.ParseExact(UsageXml.<Usage>.<TotalDuration>.Value, "c", System.Globalization.DateTimeFormatInfo.InvariantInfo)

                'TotalDuration = System.TimeSpan.ParseExact(UsageXml.<Usage>.<TotalDuration>.Value, "c", culture)
                'If System.TimeSpan.TryParse(UsageXml.<Usage>.<TotalDuration>.Value, TotalDuration) Then
                '    Debug.Print("Parsed " & UsageXml.<Usage>.<TotalDuration>.Value & " OK" & vbCrLf)
                'Else
                '    Debug.Print("Error parsing " & UsageXml.<Usage>.<TotalDuration>.Value & vbCrLf)
                'End If
            End If

        End If

    End Sub

#End Region 'Usage Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class 'Usage ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'The Trademark class stored information about a trademark used by the author of an application or data.
Public Class Trademark '---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'http://www.inta.org/TrademarkBasics/FactSheets/Pages/MarkingRequirementsFactSheet.aspx
    'http://www.ggmark.com/guide.html
    'http://www.bpmlegal.com/tmdodont.html
    'http://developer.android.com/legal.html



    'ANDORVILLE (TM) software
    'ANDORVILLE LABS (TM) software research
    'THE ANDORVILLE PROJECT (TM) software development system


    Private _text As String = "" 'The text of the trademark.
    Property Text As String
        Get
            Return _text
        End Get
        Set(value As String)
            _text = value
        End Set
    End Property

    Private _ownerName As String = "" 'The name of the owner of the trademark.
    Property OwnerName As String
        Get
            Return _ownerName
        End Get
        Set(value As String)
            _ownerName = value
        End Set
    End Property

    Private _registered As Boolean = False 'True if the trademark is a registered trademark.
    Property Registered As Boolean
        Get
            Return _registered
        End Get
        Set(value As Boolean)
            _registered = value
        End Set
    End Property

    Private _genericTerm As String 'Goods or Services (noun) that the trademark (adjective) describes. The word "brand" can also be used as the generic term.
    'Examples: Kleenex (R) tissue, Hoover (R) brand.
    Property GenericTerm As String
        Get
            Return _genericTerm
        End Get
        Set(value As String)
            _genericTerm = value
        End Set
    End Property

End Class 'Trademark ------------------------------------------------------------------------------------------------------------------------------------------------------------------------







''The TextDisplay class stores text display properties.
''REDUNDANT
'Public Class TextDisplay '-------------------------------------------------------------------------------------------------------------------------------------------------------------------

'#Region " Properties - Properties used to store text display information." '-----------------------------------------------------------------------------------------------------------------

'    Private _fontName As String = "Ariel" 'The name of the font used to display the message.
'    Property FontName As String
'        Get
'            Return _fontName
'        End Get
'        Set(value As String)
'            _fontName = value
'        End Set
'    End Property

'    Private _fontSize As Single = 10 'The font size used to display the message.
'    Property FontSize As Single
'        Get
'            Return _fontSize
'        End Get
'        Set(value As Single)
'            _fontSize = value
'        End Set
'    End Property

'    Private _fontStyle As System.Drawing.FontStyle = Drawing.FontStyle.Regular 'The font style used to display the message.
'    Property FontStyle As System.Drawing.FontStyle 'Bold, Italic, Regular, Strikeout, Underline
'        Get
'            Return _fontStyle
'        End Get
'        Set(value As System.Drawing.FontStyle)
'            If value = Drawing.FontStyle.Regular Then
'                _fontStyle = Drawing.FontStyle.Regular
'            Else
'                _fontStyle = _fontStyle Or value 'FontStyle is a flags enumeration so Or is used to combine Styles
'            End If
'        End Set
'    End Property

'    Private _color As System.Drawing.Color = Drawing.Color.Black
'    Property Color As System.Drawing.Color
'        Get
'            Return _color
'        End Get
'        Set(value As System.Drawing.Color)
'            _color = value
'        End Set
'    End Property

'    Private _charOffset As Integer = 0 'Text offset above or below the baseline. Superscipt +ve offset eg +10. Subscript -ve offset eg -10.
'    Property CharOffset As Integer
'        Get
'            Return _charOffset
'        End Get
'        Set(value As Integer)
'            _charOffset = value
'        End Set
'    End Property

'    Private _superscriptCharOffset As Integer = 10 'The offset to use for superscript text.
'    Property SuperscriptCharOffset As Integer
'        Get
'            Return _superscriptCharOffset
'        End Get
'        Set(value As Integer)
'            _superscriptCharOffset = value
'        End Set
'    End Property

'    Private _superscriptSizeReduction As Single = 0.6 'The font size reduction (multiplier) to apply to superscript text.
'    Property SuperscriptSizeReduction As Single
'        Get
'            Return _superscriptSizeReduction
'        End Get
'        Set(value As Single)
'            _superscriptSizeReduction = value
'        End Set
'    End Property

'    Private _subscriptCharOffset As Integer = -10 'The offset to use for subscript text.
'    Property SubscriptCharOffset As Integer
'        Get
'            Return _subscriptCharOffset
'        End Get
'        Set(value As Integer)
'            _subscriptCharOffset = value
'        End Set
'    End Property

'    Private _subscriptSizeReduction As Single = 0.6 'The font size reduction (multiplier) to apply to subscript text.
'    Property SubscriptSizeReduction As Single
'        Get
'            Return _subscriptSizeReduction
'        End Get
'        Set(value As Single)
'            _subscriptSizeReduction = value
'        End Set
'    End Property



'#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


'#Region " Methods - Methods used to format the displayed text" '------------------------------------------------------------------------------------------------------------------------------------------

'    Public Sub SetNormalStyle()
'        'Set the message text properties for displaying normal messages:
'        SetFontName(FontList.Arial)
'        Color = Drawing.Color.Black
'        FontStyle = Drawing.FontStyle.Regular 'This removes the following font settings: Bold, Italic, Regular, Strikeout, Underline
'    End Sub

'    'A list of text fonts. Used to set the MessageFontName property.
'    Public Enum FontList
'        Arial
'        Calibri
'        Courier
'        Lucida_Sans_Typewriter
'        OCR_A_Extended
'        Times_New_Roman
'    End Enum


'    Public Sub SetFontName(ByVal Name As FontList)
'        'Set the name of the font used to display text in the message window.
'        Select Case Name
'            Case FontList.Arial
'                FontName = "Ariel"
'            Case FontList.Calibri
'                FontName = "Calibri"
'            Case FontList.Courier
'                FontName = "Courier"
'            Case FontList.Lucida_Sans_Typewriter
'                FontName = "Lucida Sans Typewriter"
'            Case FontList.OCR_A_Extended
'                FontName = "OCR A Extended"
'            Case FontList.Times_New_Roman
'                FontName = "Times New Roman"
'        End Select
'    End Sub

'    'Public Sub DisplayTest()

'    'End Sub


'#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'End Class 'TextDisplay ----------------------------------------------------------------------------------------------------------------------------------------------------------------------





''The Trademarks class stores trademarks used by theauthor of an application or data. 
'REDUNDANT
'Public Class Trademarks

'    Public List As List(Of Trademark) 'Stores the list of trademarks

'    Private _ownerName As String
'    Property OwnerName As String
'        Get
'            Return _ownerName
'        End Get
'        Set(value As String)
'            _ownerName = value
'        End Set
'    End Property



'    'http://www.software-research.com/Corporate/copyright.html
'    'http://stackoverflow.com/questions/23068794/set-a-formatted-string-to-a-richtextbox-in-wpf
'    'http://www.biblioscape.com/rtf15_spec.htm


'    'Rich Text format strings:
'    Dim rtArialHeader As String = "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang3081{\fonttbl{\f0\fnil\fcharset0 Arial;}}" & vbCrLf & "{\*\generator Riched20 6.3.9600}\viewkind4\uc1 " & vbCrLf
'    Dim rtBoldTrademarkLine As String = ""



'End Class



''The Settings class stores project settings for an Andorville Labs application.
'Public Class clsProject

'    'Public Info As New clsProjectInfo 'Project information: Name, Description, Type, Path, CreationData, LastUsed.

'#Region " Project Properties - Properties used to store project settings."

'    Private _applicationDirectory As String = "" 'The path to the Application directory.
'    Property ApplicationDirectory As String
'        Get
'            Return _applicationDirectory
'        End Get
'        Set(value As String)
'            _applicationDirectory = value
'        End Set
'    End Property

'    Private _settingsPath As String = "" 'The path to the project directory or project file used to store application settings.
'    Property SettingsPath As String
'        Get
'            Return _settingsPath
'        End Get
'        Set(value As String)
'            _settingsPath = value
'        End Set
'    End Property

'    Private _dataPath As String = "" 'The path to the project directory or project file used to store application data.
'    Property DataPath As String
'        Get
'            Return _dataPath
'        End Get
'        Set(value As String)
'            _dataPath = value
'        End Set
'    End Property

'#End Region 'Settings properties


'#Region " Project Methods - Methods used to save and read project settings and data."


'#End Region

'End Class