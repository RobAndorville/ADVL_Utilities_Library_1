Public Class FormatConvert
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    'Copyright 2018 Signalworks Pty Ltd, ABN 26 066 681 598

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

    'Format Convert Classes.
    'The Format Convert classes are used to convert Andorville(TM) system files to the current file formats.
    '
    'AppListFileConversion         - This class converts the Application List file to a new format.
    'AppInfoFileConversion         - This class converts the Application Information file.
    'ProjectListFileConversion     - Converts the Project List File to the current file format.
    'ProjectInfoFileConversion     - Converts the Project Information File to the current file format.
    'LastProjectInfoFileConversion - Converts the Last Project Information File to the current file format.


    Public Class AppListFileConversion
        'This class is used to convert an Andorville(TM) Application List file to a version using a new format.
        '  An ADVL_1 format Application List file is named Application_List.xml.
        '  An ADVL_2 format Application List file is named Application_List_ADVL_2.xml.
        '  These files are stored in the Application Directory.
        'NOTE: THE APP LIST FILE IS ONLY USED IN THE APPLICATION NETWORK APP. THIS CODE WILL BE MOVED TO THIS APP.

        'Properties:
        '  DirectoryPath
        '  InputFilename
        '  InputFormatCode
        '  OutputFormatCode

        'Methods:
        '  Convert                       Converts the Input file to a version using the format corresponding to the OutputFormatCode.
        '  InputXDoc                     Returns the Input XDocument
        '  ConvertedXDoc                 Returns the converted XDocument
        '  GetInputFormatCode(InputXDoc) Sets the InputFormatCode property based on the InputXDoc.
        '  GetInputFormatCode()          Sets the InputFormatCode property based on the XDocument in the DirectoryPath with InputFileName.

#Region " Properties" '----------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Private _directoryPath As String = "" 'The path to the directory containing the Andorville(TM) system file.
        Property DirectoryPath As String
            Get
                Return _directoryPath
            End Get
            Set(value As String)
                _directoryPath = value
            End Set
        End Property

        Private _inputFileName As String = "" 'The input file name of the Andorville(TM) system file to be converted.
        Property InputFileName As String
            Get
                Return _inputFileName
            End Get
            Set(value As String)
                _inputFileName = value
            End Set
        End Property

        Enum FormatCodes
            Unknown
            ADVL_1
            ADVL_2
        End Enum

        Private _inputFormatCode As FormatCodes = FormatCodes.Unknown 'The format code of the input file.
        Property InputFormatCode As FormatCodes
            Get
                Return _inputFormatCode
            End Get
            Set(value As FormatCodes)
                _inputFormatCode = value
            End Set
        End Property


        Private _outputFormatCode As FormatCodes = FormatCodes.Unknown 'The required format code of the converted output file.
        Property OutputFormatCode As FormatCodes
            Get
                Return _outputFormatCode
            End Get
            Set(value As FormatCodes)
                _outputFormatCode = value
            End Set
        End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Sub Convert()
            'Convert the Application List file to the new format.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Application List.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Application List file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            GetInputFormatCode()

            If InputFormatCode = FormatCodes.Unknown Then
                RaiseEvent ErrorMessage("The Application List Xml document has an unknown format." & vbCrLf)
            ElseIf InputFormatCode = FormatCodes.ADVL_1 Then
                If OutputFormatCode = FormatCodes.Unknown Then
                    RaiseEvent ErrorMessage("The required converted format is unknown." & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_1 Then
                    RaiseEvent Message("The required format is ADVL_1. This is the current format. No conversion is necessary" & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_2 Then
                    'Convert ADVL_1 to ADVL_2.
                    'Check if the Application_Info_ADVL_2.xml file already exists:
                    If System.IO.File.Exists(DirectoryPath & "\Application_List_ADVL_2.xml") Then
                        RaiseEvent ErrorMessage("The required converted format file already exists: Application_List_ADVL_2.xml" & vbCrLf)
                    Else
                        ConvertADVL1ToADVL2(InputXDoc)
                    End If
                Else
                    RaiseEvent ErrorMessage("The required format is not defined." & vbCrLf)
                End If


            ElseIf InputFormatCode = FormatCodes.ADVL_2 Then
                RaiseEvent Message("The Application List Xml document uses the current ADVL_2 format. No conversion is necessary." & vbCrLf)
            End If

        End Sub

        Private Sub ConvertADVL1ToADVL2(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Convert the Application List file in format ADVL_1 to a the format ADVL_2.

            'Check if the ADVL_2 format version of the Applicvation List file already exists:
            If System.IO.File.Exists(DirectoryPath & "\Application_List_ADVL_2.xml") Then
                RaiseEvent Message("The ADVL_2 format version of the Application List Xml document already exisits. No conversion is necessary." & vbCrLf)
                Exit Sub
            End If

            Try
                Dim newList = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Application List File-->
                              <ApplicationList>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <%= From item In InputXDoc.<ApplicationList>.<Application>
                                      Select
                                  <Application>
                                      <Name><%= item.<Name>.Value %></Name>
                                      <Description><%= item.<Description>.Value %></Description>
                                      <Directory><%= item.<Directory>.Value %></Directory>
                                      <ExecutablePath><%= item.<ExecutablePath>.Value %></ExecutablePath>
                                  </Application> %>
                              </ApplicationList>

                'Save the new version of the Project List:
                newList.Save(DirectoryPath & "\Application_List_ADVL_2.xml")

            Catch ex As Exception
                'Main.Message.AddWarning("Application List format conversion error: " & ex.Message & vbCrLf)
                RaiseEvent ErrorMessage("Application List format conversion error: " & ex.Message & vbCrLf)
            End Try
        End Sub

        Public Function InputXDoc() As System.Xml.Linq.XDocument
            'Return the Input Application Information XDocument to be converted:
            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                Return XDocument.Load(DirectoryPath & "\" & InputFileName)
            End If
        End Function

        Public Function ConvertedXDoc() As System.Xml.Linq.XDocument
            'Return the converted Application List XDocument:
            If System.IO.File.Exists(DirectoryPath & "\Application_Info_ADVL_2.xml") Then
                Return XDocument.Load(DirectoryPath & "\Application_Info_ADVL_2.xml")
            End If
        End Function

        Public Sub GetInputFormatCode(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Set the FormatCode of the Input Application Information file.

            If InputXDoc.<ApplicationList>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Application List Xml document does not contain the ApplicationList element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<ApplicationList>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Application List Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<ApplicationList>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

        Public Sub GetInputFormatCode()
            'Version of the GetInputFormatCode that reads the Application List into InputXDoc.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Application List.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Application List file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            If InputXDoc.<ApplicationList>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Application List Xml document does not contain the ApplicationList element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<ApplicationList>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Application List Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<ApplicationList>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Event Message(ByVal Message As String) 'Send a message.
        Public Event ErrorMessage(ByVal Message As String) 'Send an error message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Class

    Public Class AppInfoFileConversion
        'This class is used to convert an Andorville(TM) Application Information file to a version using a new format.
        '  An ADVL_1 format Application Information file is named Application_Info.xml.
        '  An ADVL_2 format Application Information file is named Application_Info_ADVL_2.xml.
        '  These files are stored in the Application Directory.

        'Properties:
        '  DirectoryPath
        '  InputFilename
        '  InputFormatCode
        '  OutputFormatCode

        'Methods:
        '  Convert                       Converts the Input file to a version using the format corresponding to the OutputFormatCode.
        '  InputXDoc                     Returns the Input XDocument
        '  ConvertedXDoc                 Returns the converted XDocument
        '  GetInputFormatCode(InputXDoc) Sets the InputFormatCode property based on the InputXDoc.
        '  GetInputFormatCode()          Sets the InputFormatCode property based on the XDocument in the DirectoryPath with InputFileName.



#Region " Properties" '----------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Private _directoryPath As String = "" 'The path to the directory containing the Andorville(TM) system file.
        Property DirectoryPath As String
            Get
                Return _directoryPath
            End Get
            Set(value As String)
                _directoryPath = value
            End Set
        End Property

        Private _inputFileName As String = "" 'The input file name of the Andorville(TM) system file to be converted.
        Property InputFileName As String
            Get
                Return _inputFileName
            End Get
            Set(value As String)
                _inputFileName = value
            End Set
        End Property

        Enum FormatCodes
            Unknown
            ADVL_1
            ADVL_2
        End Enum

        Private _inputFormatCode As FormatCodes = FormatCodes.Unknown 'The format code of the input file.
        Property InputFormatCode As FormatCodes
            Get
                Return _inputFormatCode
            End Get
            Set(value As FormatCodes)
                _inputFormatCode = value
            End Set
        End Property

        Private _outputFormatCode As FormatCodes = FormatCodes.Unknown 'The required format code of the converted output file.
        Property OutputFormatCode As FormatCodes
            Get
                Return _outputFormatCode
            End Get
            Set(value As FormatCodes)
                _outputFormatCode = value
            End Set
        End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Sub Convert()
            'Convert the Application Information file to the new format.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Application Information.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Application Information file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            GetInputFormatCode()

            If InputFormatCode = FormatCodes.Unknown Then
                RaiseEvent ErrorMessage("The Application Information Xml document has an unknown format." & vbCrLf)
            ElseIf InputFormatCode = FormatCodes.ADVL_1 Then
                If OutputFormatCode = FormatCodes.Unknown Then
                    RaiseEvent ErrorMessage("The required converted format is unknown." & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_1 Then
                    RaiseEvent Message("The required format is ADVL_1. This is the current format. No conversion is necessary" & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_2 Then
                    'Convert ADVL_1 to ADVL_2.
                    'Check if the Application_Info_ADVL_2.xml file already exists:
                    If System.IO.File.Exists(DirectoryPath & "\Application_Info_ADVL_2.xml") Then
                        RaiseEvent ErrorMessage("The required converted format file already exists: Application_Info_ADVL_2.xml" & vbCrLf)
                    Else
                        ConvertADVL1ToADVL2(InputXDoc)
                    End If
                Else
                    RaiseEvent ErrorMessage("The required format is not defined." & vbCrLf)
                End If


            ElseIf InputFormatCode = FormatCodes.ADVL_2 Then
                RaiseEvent Message("The Application Information Xml document uses the current ADVL_2 format. No conversion is necessary." & vbCrLf)
            End If

        End Sub

        Private Sub ConvertADVL1ToADVL2(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Convert the Application Info file in format ADVL_1 to a the format ADVL_2.

            'Check if the ADVL_2 format version of the Applicvation Information file already exists:
            If System.IO.File.Exists(DirectoryPath & "\Application_Info_ADVL_2.xml") Then
                RaiseEvent Message("The ADVL_2 format version of the Application Information Xml document already exisits. No conversion is necessary." & vbCrLf)
                Exit Sub
            End If

            Try
                Dim newInfo = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Application Information File-->
                              <!---->
                              <Application>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <Name><%= InputXDoc.<Application>.<Name>.Value %></Name>
                                  <ExecutablePath><%= InputXDoc.<Application>.<ExecutablePath>.Value %></ExecutablePath>
                                  <Description><%= InputXDoc.<Application>.<Description>.Value %></Description>
                                  <CreationDate><%= InputXDoc.<Application>.<CreationDate>.Value %></CreationDate>
                                  <FileAssociationList>
                                      <%= From item In InputXDoc.<Application>.<FileAssociationList>.<FileAssociation>
                                          Select
                                      <FileAssociation>
                                          <Extension><%= item.<Extension>.Value %></Extension>
                                          <Description><%= item.<Description>.Value %></Description>
                                      </FileAssociation> %>
                                  </FileAssociationList>
                                  <Version>
                                      <Major><%= InputXDoc.<Application>.<Version>.<Major>.Value %></Major>
                                      <Minor><%= InputXDoc.<Application>.<Version>.<Minor>.Value %></Minor>
                                      <Build><%= InputXDoc.<Application>.<Version>.<Build>.Value %></Build>
                                      <Revision><%= InputXDoc.<Application>.<Version>.<Revision>.Value %></Revision>
                                  </Version>
                                  <Author>
                                      <Name><%= InputXDoc.<Application>.<Author>.<Name>.Value %></Name>
                                      <Description><%= InputXDoc.<Application>.<Author>.<Description>.Value %></Description>
                                      <Contact><%= InputXDoc.<Application>.<Author>.<Contact>.Value %></Contact>
                                  </Author>
                                  <Copyright>
                                      <OwnerName><%= InputXDoc.<Application>.<Copyright>.<OwnerName>.Value %></OwnerName>
                                      <PublicationYear><%= InputXDoc.<Application>.<Copyright>.<PublicationYear>.Value %></PublicationYear>
                                      <Notice><%= InputXDoc.<Application>.<Copyright>.<Notice>.Value %></Notice>
                                  </Copyright>
                                  <TrademarkList>
                                      <%= From item In InputXDoc.<Application>.<TrademarkList>.<Trademark>
                                          Select
                                      <Trademark>
                                          <Text><%= item.<Text>.Value %></Text>
                                          <OwnerName><%= item.<OwnerName>.Value %></OwnerName>
                                          <Registered><%= item.<Registered>.Value %></Registered>
                                          <GenericTerm><%= item.<GenericTerm>.Value %></GenericTerm>
                                      </Trademark> %>
                                  </TrademarkList>
                                  <License>
                                      <Code><%= InputXDoc.<Application>.<License>.<Code>.Value %></Code>
                                      <Notice><%= InputXDoc.<Application>.<License>.<Notice>.Value %></Notice>
                                      <Text><%= InputXDoc.<Application>.<License>.<Text>.Value %></Text>
                                  </License>
                                  <SourceCode>
                                      <Language><%= InputXDoc.<Application>.<SourceCode>.<Language>.Value %></Language>
                                      <FileName><%= InputXDoc.<Application>.<SourceCode>.<FileName>.Value %></FileName>
                                      <FileSize><%= InputXDoc.<Application>.<SourceCode>.<FileSize>.Value %></FileSize>
                                      <FileHash><%= InputXDoc.<Application>.<SourceCode>.<FileHash>.Value %></FileHash>
                                      <WebLink><%= InputXDoc.<Application>.<SourceCode>.<WebLink>.Value %></WebLink>
                                      <Contact><%= InputXDoc.<Application>.<SourceCode>.<Contact>.Value %></Contact>
                                      <Comments><%= InputXDoc.<Application>.<SourceCode>.<Comments>.Value %></Comments>
                                  </SourceCode>
                                  <ModificationSummary>
                                      <BaseCodeName><%= InputXDoc.<Application>.<ModificationSummary>.<BaseCodeName>.Value %></BaseCodeName>
                                      <BaseCodeDescription><%= InputXDoc.<Application>.<ModificationSummary>.<BaseCodeDescription>.Value %></BaseCodeDescription>
                                      <BaseCodeVersion>
                                          <Major><%= InputXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Major>.Value %></Major>
                                          <Minor><%= InputXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Minor>.Value %></Minor>
                                          <Build><%= InputXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Build>.Value %></Build>
                                          <Revision><%= InputXDoc.<Application>.<ModificationSummary>.<BaseCodeVersion>.<Revision>.Value %></Revision>
                                      </BaseCodeVersion>
                                      <Description><%= InputXDoc.<Application>.<ModificationSummary>.<Description>.Value %></Description>
                                  </ModificationSummary>
                                  <LibraryList>
                                      <%= From item In InputXDoc.<Application>.<LibraryList>.<Library>
                                          Select
                                          <Library>
                                              <Name><%= item.<Name>.Value %></Name>
                                              <Description><%= item.<Description>.Value %></Description>
                                              <CreationDate><%= item.<CreationDate>.Value %></CreationDate>
                                              <LicenseNotice><%= item.<LicenseNotice>.Value %></LicenseNotice>
                                              <CopyrightNotice><%= item.<CopyrightNotice>.Value %></CopyrightNotice>
                                              <Version>
                                                  <Major><%= item.<Version>.<Major>.Value %></Major>
                                                  <Minor><%= item.<Version>.<Minor>.Value %></Minor>
                                                  <Build><%= item.<Version>.<Build>.Value %></Build>
                                                  <Revision><%= item.<Version>.<Revision>.Value %></Revision>
                                              </Version>
                                              <Author>
                                                  <Name><%= item.<Author>.<Name>.Value %></Name>
                                                  <Description><%= item.<Author>.<Description>.Value %></Description>
                                                  <Contact><%= item.<Author>.<Contact>.Value %></Contact>
                                              </Author>
                                              <ClassList>
                                                  <%= From listItem In item.<ClassList>.<Class>
                                                      Select
                                                      <Class>
                                                          <Name><%= listItem.<Name>.Value %></Name>
                                                          <Description><%= listItem.<Description>.Value %></Description>
                                                      </Class> %>
                                              </ClassList>
                                          </Library>
                                      %>
                                  </LibraryList>
                              </Application>


                'Save the new version of the Application Information file:
                newInfo.Save(DirectoryPath & "\Application_Info_ADVL_2.xml")

            Catch ex As Exception
                'Main.Message.AddWarning("Application Information format conversion error: " & ex.Message & vbCrLf)
                RaiseEvent ErrorMessage("Application Information format conversion error: " & ex.Message & vbCrLf)
            End Try
        End Sub

        Public Function InputXDoc() As System.Xml.Linq.XDocument
            'Return the Input Application Information XDocument to be converted:
            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                Return XDocument.Load(DirectoryPath & "\" & InputFileName)
            End If
        End Function

        Public Function ConvertedXDoc() As System.Xml.Linq.XDocument
            'Return the converted Application Information XDocument:
            If System.IO.File.Exists(DirectoryPath & "\Application_Info_ADVL_2.xml") Then
                Return XDocument.Load(DirectoryPath & "\Application_Info_ADVL_2.xml")
            End If
        End Function

        Public Sub GetInputFormatCode(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Set the FormatCode of the Input Application Information file.

            If InputXDoc.<Application>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Application Information Xml document does not contain the Application element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<Application>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Application Information Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<Application>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

        Public Sub GetInputFormatCode()
            'Version of the GetInputFormatCode that reads the Application Information into InputXDoc.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Application Information.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Application Information file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            If InputXDoc.<Application>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Application Information Xml document does not contain the Application element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<Application>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Application Information Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<Application>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Event Message(ByVal Message As String) 'Send a message.
        Public Event ErrorMessage(ByVal Message As String) 'Send an error message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Class

    Public Class ProjectListFileConversion
        'This class is used to convert an Andorville(TM) Project List file to a version using a new format.
        '  An ADVL_1 format Application List file is named Project_List.xml.
        '  An ADVL_2 format Application List file is named Project_List_ADVL_2.xml.
        '  These files are stored in the Application Directory.

        'Properties:
        '  DirectoryPath
        '  InputFilename
        '  InputFormatCode
        '  OutputFormatCode

        'Methods:
        '  Convert                       Converts the Input file to a version using the format corresponding to the OutputFormatCode.
        '  InputXDoc                     Returns the Input XDocument
        '  ConvertedXDoc                 Returns the converted XDocument
        '  GetInputFormatCode(InputXDoc) Sets the InputFormatCode property based on the InputXDoc.
        '  GetInputFormatCode()          Sets the InputFormatCode property based on the XDocument in the DirectoryPath with InputFileName.

#Region " Properties" '----------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Private _directoryPath As String = "" 'The path to the directory containing the Andorville(TM) system file.
        Property DirectoryPath As String
            Get
                Return _directoryPath
            End Get
            Set(value As String)
                _directoryPath = value
            End Set
        End Property

        Private _inputFileName As String = "" 'The input file name of the Andorville(TM) system file to be converted.
        Property InputFileName As String
            Get
                Return _inputFileName
            End Get
            Set(value As String)
                _inputFileName = value
            End Set
        End Property

        Enum FormatCodes
            Unknown
            ADVL_1
            ADVL_2
        End Enum

        Private _inputFormatCode As FormatCodes = FormatCodes.Unknown 'The format code of the input file.
        Property InputFormatCode As FormatCodes
            Get
                Return _inputFormatCode
            End Get
            Set(value As FormatCodes)
                _inputFormatCode = value
            End Set
        End Property


        Private _outputFormatCode As FormatCodes = FormatCodes.Unknown 'The required format code of the converted output file.
        Property OutputFormatCode As FormatCodes
            Get
                Return _outputFormatCode
            End Get
            Set(value As FormatCodes)
                _outputFormatCode = value
            End Set
        End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Sub Convert()
            'Convert the Application List file to the new format.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Project List.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Project List file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            GetInputFormatCode()

            If InputFormatCode = FormatCodes.Unknown Then
                RaiseEvent ErrorMessage("The Project List Xml document has an unknown format." & vbCrLf)
            ElseIf InputFormatCode = FormatCodes.ADVL_1 Then
                If OutputFormatCode = FormatCodes.Unknown Then
                    RaiseEvent ErrorMessage("The required converted format is unknown." & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_1 Then
                    RaiseEvent Message("The required format is ADVL_1. This is the current format. No conversion is necessary" & vbCrLf)
                    'ElseIf OutputFormatCode = FormatCodes.ADVL_2 Then
                    'ElseIf OutputFormatCode = FormatCodes.Unknown Then
                ElseIf OutputFormatCode = FormatCodes.ADVL_2 Then
                    'Convert ADVL_1 to ADVL_2.
                    'Check if the Project_List_ADVL_2.xml file already exists:
                    If System.IO.File.Exists(DirectoryPath & "\Project_List_ADVL_2.xml") Then
                        RaiseEvent ErrorMessage("The required converted format file already exists: Project_List_ADVL_2.xml" & vbCrLf)
                    Else
                        ConvertADVL1ToADVL2(InputXDoc)
                    End If
                Else
                    RaiseEvent ErrorMessage("The required format is not defined." & vbCrLf)
                End If


            ElseIf InputFormatCode = FormatCodes.ADVL_2 Then
                RaiseEvent Message("The Project List Xml document uses the current ADVL_2 format. No conversion is necessary." & vbCrLf)
            End If

        End Sub

        Private Sub ConvertADVL1ToADVL2(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Convert the Project List file in format ADVL_1 to a the format ADVL_2.

            'Check if the ADVL_2 format version of the Project List file already exists:
            If System.IO.File.Exists(DirectoryPath & "\Project_List_ADVL_2.xml") Then
                RaiseEvent Message("The ADVL_2 format version of the Project List Xml document already exisits. No conversion is necessary." & vbCrLf)
                Exit Sub
            End If

            Try
                Dim newList = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Project List File-->
                              <ProjectList>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <ApplicationName><%= InputXDoc.<ProjectList>.<Project>.<ApplicationName>.Value %></ApplicationName>
                                  <%= From item In InputXDoc.<ProjectList>.<Project>
                                      Select
                                  <Project>
                                      <Name><%= item.<Name>.Value %></Name>
                                      <AppNetName><%= "" %></AppNetName>
                                      <Description><%= item.<Description>.Value %></Description>
                                      <CreationDate><%= item.<CreationDate>.Value %></CreationDate>
                                      <Type><%= item.<Type>.Value %></Type>
                                      <Path><%= item.<SettingsLocationPath>.Value %></Path>
                                      <AuthorName><%= item.<AuthorName>.Value %></AuthorName>
                                  </Project> %>
                              </ProjectList>


                ' <AppNetName><%= "" %></AppNetName> 'ADDED 10Feb19 - AppNetName field added to the Project entry. (This field is not present in earlier format versions.)

                'Save the new version of the Project List:
                newList.Save(DirectoryPath & "\Project_List_ADVL_2.xml")

            Catch ex As Exception
                'Main.Message.AddWarning("Project List format conversion error: " & ex.Message & vbCrLf)
                RaiseEvent ErrorMessage("Project List format conversion error: " & ex.Message & vbCrLf)
            End Try
        End Sub

        Public Function InputXDoc() As System.Xml.Linq.XDocument
            'Return the Input Project List XDocument to be converted:
            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                Return XDocument.Load(DirectoryPath & "\" & InputFileName)
            End If
        End Function

        Public Function ConvertedXDoc() As System.Xml.Linq.XDocument
            'Return the converted Project List XDocument:
            If System.IO.File.Exists(DirectoryPath & "\Project_List_ADVL_2.xml") Then
                Return XDocument.Load(DirectoryPath & "\Project_List_ADVL_2.xml")
            End If
        End Function

        Public Sub GetInputFormatCode(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Set the FormatCode of the Input Application Information file.

            If InputXDoc.<ProjectList>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Project List Xml document does not contain the ProjectList element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<ProjectList>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Project List Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<ProjectList>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

        Public Sub GetInputFormatCode()
            'Version of the GetInputFormatCode that reads the Application List into InputXDoc.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Application List.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Project List file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            If InputXDoc.<ProjectList>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Project List Xml document does not contain the ProjectList element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<ProjectList>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Project List Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<ProjectList>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Event Message(ByVal Message As String) 'Send a message.
        Public Event ErrorMessage(ByVal Message As String) 'Send an error message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    End Class

    Public Class LastProjectInfoFileConversion
        'This class is used to convert an Andorville(TM) Last Project Info file to a version using a new format.
        '  An ADVL_1 format Application List file is named Last_Project_Info.xml.
        '  An ADVL_2 format Application List file is named Last_Project_Info_ADVL_2.xml.
        '  These files are stored in the Application Directory.

        'Properties:
        '  DirectoryPath
        '  InputFilename
        '  InputFormatCode
        '  OutputFormatCode

        'Methods:
        '  Convert                       Converts the Input file to a version using the format corresponding to the OutputFormatCode.
        '  InputXDoc                     Returns the Input XDocument
        '  ConvertedXDoc                 Returns the converted XDocument
        '  GetInputFormatCode(InputXDoc) Sets the InputFormatCode property based on the InputXDoc.
        '  GetInputFormatCode()          Sets the InputFormatCode property based on the XDocument in the DirectoryPath with InputFileName.

#Region " Properties" '----------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Private _directoryPath As String = "" 'The path to the directory containing the Andorville(TM) system file.
        Property DirectoryPath As String
            Get
                Return _directoryPath
            End Get
            Set(value As String)
                _directoryPath = value
            End Set
        End Property

        Private _inputFileName As String = "" 'The input file name of the Andorville(TM) system file to be converted.
        Property InputFileName As String
            Get
                Return _inputFileName
            End Get
            Set(value As String)
                _inputFileName = value
            End Set
        End Property

        Enum FormatCodes
            Unknown
            ADVL_1
            ADVL_2
        End Enum

        Private _inputFormatCode As FormatCodes = FormatCodes.Unknown 'The format code of the input file.
        Property InputFormatCode As FormatCodes
            Get
                Return _inputFormatCode
            End Get
            Set(value As FormatCodes)
                _inputFormatCode = value
            End Set
        End Property

        Private _outputFormatCode As FormatCodes = FormatCodes.Unknown 'The required format code of the converted output file.
        Property OutputFormatCode As FormatCodes
            Get
                Return _outputFormatCode
            End Get
            Set(value As FormatCodes)
                _outputFormatCode = value
            End Set
        End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Sub Convert()
            'Convert the Last Project Information file to the new format.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Last Project Information.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Last Project Information file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            GetInputFormatCode()

            If InputFormatCode = FormatCodes.Unknown Then
                RaiseEvent ErrorMessage("The Last Project Information Xml document has an unknown format." & vbCrLf)
            ElseIf InputFormatCode = FormatCodes.ADVL_1 Then
                If OutputFormatCode = FormatCodes.Unknown Then
                    RaiseEvent ErrorMessage("The required converted format is unknown." & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_1 Then
                    RaiseEvent Message("The required format is ADVL_1. This is the current format. No conversion is necessary" & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_2 Then
                    'Convert ADVL_1 to ADVL_2.
                    'Check if the Application_Info_ADVL_2.xml file already exists:
                    If System.IO.File.Exists(DirectoryPath & "\Last_Project_Info_ADVL_2.xml") Then
                        RaiseEvent ErrorMessage("The required converted format file already exists: Last_Project_Info_ADVL_2.xml" & vbCrLf)
                    Else
                        ConvertADVL1ToADVL2(InputXDoc)
                    End If
                Else
                    RaiseEvent ErrorMessage("The required format is not defined." & vbCrLf)
                End If

            ElseIf InputFormatCode = FormatCodes.ADVL_2 Then
                RaiseEvent Message("The Last Project Information Xml document uses the current ADVL_2 format. No conversion is necessary." & vbCrLf)
            End If

        End Sub

        Private Sub ConvertADVL1ToADVL2(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Convert the Last Project Info file in format ADVL_1 to a the format ADVL_2.

            'Check if the ADVL_2 format version of the Last Project Information file already exists:
            If System.IO.File.Exists(DirectoryPath & "\Last_Project_Info_ADVL_2.xml") Then
                RaiseEvent Message("The ADVL_2 format version of the Last Project Information Xml document already exisits. No conversion is necessary." & vbCrLf)
                Exit Sub
            End If

            Try
                'Dim newFile = <?xml version="1.0" encoding="utf-8"?>
                '              <!---->
                '              <!--Last Project Information-->
                '              <Project>
                '                  <Name><%= InputXDoc.<Project>.<Name>.Value %></Name>
                '                  <Description><%= InputXDoc.<Project>.<Description>.Value %></Description>
                '                  <FormatCode>ADVL_2</FormatCode>
                '                  <Location>
                '                      <Type><%= InputXDoc.<Project>.<Location>.<Type>.Value %></Type>
                '                      <Path><%= InputXDoc.<Project>.<Location>.<Path>.Value %></Path>
                '                  </Location>
                '              </Project>

                Dim newFile = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Last Project Information-->
                              <Project>
                                  <Name><%= InputXDoc.<Project>.<Name>.Value %></Name>
                                  <Description><%= InputXDoc.<Project>.<Description>.Value %></Description>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <Location>
                                      <Type><%= InputXDoc.<Project>.<SettingsLocation>.<Type>.Value %></Type>
                                      <Path><%= InputXDoc.<Project>.<SettingsLocation>.<Path>.Value %></Path>
                                  </Location>
                              </Project>


                'Save the new version of the Last Project Information file:
                newFile.Save(DirectoryPath & "\Last_Project_Info_ADVL_2.xml")

            Catch ex As Exception
                'Main.Message.AddWarning("Last Project Information format conversion error: " & ex.Message & vbCrLf)
                RaiseEvent ErrorMessage("Last Project Information format conversion error: " & ex.Message & vbCrLf)
            End Try
        End Sub

        Public Function InputXDoc() As System.Xml.Linq.XDocument
            'Return the Last Project Information XDocument to be converted:
            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                Return XDocument.Load(DirectoryPath & "\" & InputFileName)
            End If
        End Function

        Public Function ConvertedXDoc() As System.Xml.Linq.XDocument
            'Return the converted Last Project Information XDocument:
            If System.IO.File.Exists(DirectoryPath & "\Last_Project_Info_ADVL_2.xml") Then
                Return XDocument.Load(DirectoryPath & "\Last_Project_Info_ADVL_2.xml")
            End If
        End Function

        Public Sub GetInputFormatCode(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Set the FormatCode of the Input Last Project Information file.

            If InputXDoc.<Project>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Last Project Information Xml document does not contain the Project element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<Project>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Last Project Information Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<Project>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

        Public Sub GetInputFormatCode()
            'Version of the GetInputFormatCode that reads the Last Project Information into InputXDoc.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Application Information.

            If System.IO.File.Exists(DirectoryPath & "\" & InputFileName) Then
                InputXDoc = XDocument.Load(DirectoryPath & "\" & InputFileName)
            Else
                RaiseEvent ErrorMessage("The Last Project Information file was not found: " & DirectoryPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            If InputXDoc.<Project>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Last Project Information Xml document does not contain the Project element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<Project>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Last project Information Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<Project>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Event Message(ByVal Message As String) 'Send a message.
        Public Event ErrorMessage(ByVal Message As String) 'Send an error message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    End Class

    Public Class ProjectInfoFileConversion
        'This class is used to convert an Andorville(TM) Project Information file to a version using a new format.
        '  An ADVL_1 format Application Information file is named ADVL_Project_Info.xml.
        '  An ADVL_2 format Application Information file is named Project_Info_ADVL_2.xml.
        '  These files are stored in the Project Directory or Archive.

        'Properties:
        '  ProjectType
        '  ProjectPath
        '  InputFilename
        '  InputFormatCode
        '  OutputFormatCode

        'Methods:
        '  Convert                       Converts the Input file to a version using the format corresponding to the OutputFormatCode.
        '  InputXDoc                     Returns the Input XDocument
        '  ConvertedXDoc                 Returns the converted XDocument
        '  GetInputFormatCode(InputXDoc) Sets the InputFormatCode property based on the InputXDoc.
        '  GetInputFormatCode()          Sets the InputFormatCode property based on the XDocument in the DirectoryPath with InputFileName.

        Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp


#Region " Properties" '----------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Enum ProjectTypes
            Directory
            Archive
            Unknown
        End Enum

        Private _projectType As ProjectTypes = ProjectTypes.Unknown 'The type of project containing the Project Information file.
        Property ProjectType As ProjectTypes
            Get
                Return _projectType
            End Get
            Set(value As ProjectTypes)
                _projectType = value
            End Set
        End Property

        Private _projectPath As String = "" 'The path to the directory containing the Andorville(TM) system file.
        Property ProjectPath As String
            Get
                Return _projectPath
            End Get
            Set(value As String)
                _projectPath = value
            End Set
        End Property

        Private _inputFileName As String = "" 'The input file name of the Andorville(TM) system file to be converted.
        Property InputFileName As String
            Get
                Return _inputFileName
            End Get
            Set(value As String)
                _inputFileName = value
            End Set
        End Property

        Enum FormatCodes
            Unknown
            ADVL_1
            ADVL_2
        End Enum

        Private _inputFormatCode As FormatCodes = FormatCodes.Unknown 'The format code of the input file.
        Property InputFormatCode As FormatCodes
            Get
                Return _inputFormatCode
            End Get
            Set(value As FormatCodes)
                _inputFormatCode = value
            End Set
        End Property

        Private _outputFormatCode As FormatCodes = FormatCodes.Unknown 'The required format code of the converted output file.
        Property OutputFormatCode As FormatCodes
            Get
                Return _outputFormatCode
            End Get
            Set(value As FormatCodes)
                _outputFormatCode = value
            End Set
        End Property

#End Region 'Properties ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Sub Convert()
            'Convert the Project Information file to the new format.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Project Information.

            If InputXDocExists() Then
                GetInputXDoc(InputXDoc)
            Else
                RaiseEvent ErrorMessage("The Project Information file was not found: " & ProjectPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            GetInputFormatCode()

            If InputFormatCode = FormatCodes.Unknown Then
                RaiseEvent ErrorMessage("The Project Information Xml document has an unknown format." & vbCrLf)
            ElseIf InputFormatCode = FormatCodes.ADVL_1 Then
                If OutputFormatCode = FormatCodes.Unknown Then
                    RaiseEvent ErrorMessage("The required converted format is unknown." & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_1 Then
                    RaiseEvent Message("The required format is ADVL_1. This is the current format. No conversion is necessary" & vbCrLf)
                ElseIf OutputFormatCode = FormatCodes.ADVL_2 Then
                    'Convert ADVL_1 to ADVL_2.
                    'Check if the Application_Info_ADVL_2.xml file already exists:
                    If FileExists("Project_Info_ADVL_2.xml") Then
                        RaiseEvent ErrorMessage("The required converted format file already exists: Application_Info_ADVL_2.xml" & vbCrLf)
                    Else
                        ConvertADVL1ToADVL2(InputXDoc)
                    End If
                Else
                    RaiseEvent ErrorMessage("The required format is not defined." & vbCrLf)
                End If

            ElseIf InputFormatCode = FormatCodes.ADVL_2 Then
                RaiseEvent Message("The Project Information Xml document uses the current ADVL_2 format. No conversion is necessary." & vbCrLf)
            End If

        End Sub

        Private Function InputXDocExists() As Boolean
            'Returns True if the InputXDoc specified by ProjectType, ProjectPath and InputFileName exists.

            Select Case ProjectType
                Case ProjectTypes.Unknown
                    RaiseEvent ErrorMessage("The Project Type is unknown." & vbCrLf)
                    Return False
                Case ProjectTypes.Directory
                    If System.IO.File.Exists(ProjectPath & "\" & InputFileName) Then
                        Return True
                    Else
                        Return False
                    End If
                Case ProjectTypes.Archive
                    If System.IO.File.Exists(ProjectPath) Then 'Project Archive found.
                        Zip = New ADVL_Utilities_Library_1.ZipComp
                        Zip.ArchivePath = ProjectPath
                        Return Zip.EntryExists(InputFileName)
                    Else 'Project Archive not found.
                        RaiseEvent ErrorMessage("The Project Archive was not found: " & ProjectPath & vbCrLf)
                        Return False
                    End If
            End Select

        End Function

        Private Function FileExists(ByVal FileName As String) As Boolean
            'Returns True if FileName exists in the Project.

            Select Case ProjectType
                Case ProjectTypes.Unknown
                    RaiseEvent ErrorMessage("The Project Type is unknown." & vbCrLf)
                    Return False
                Case ProjectTypes.Directory
                    If System.IO.File.Exists(ProjectPath & "\" & FileName) Then
                        Return True
                    Else
                        Return False
                    End If
                Case ProjectTypes.Archive
                    If System.IO.File.Exists(ProjectPath) Then 'Project Archive found.
                        Zip = New ADVL_Utilities_Library_1.ZipComp
                        Zip.ArchivePath = ProjectPath
                        Return Zip.EntryExists(FileName)
                    Else 'Project Archive not found.
                        RaiseEvent ErrorMessage("The Project Archive was not found: " & ProjectPath & vbCrLf)
                        Return False
                    End If
            End Select
        End Function

        Private Sub GetInputXDoc(ByRef InputXDoc As System.Xml.Linq.XDocument)
            'Get the XDocument specified by ProjectType, ProjectPath and InputFileName.

            Select Case ProjectType
                Case ProjectTypes.Unknown
                    RaiseEvent ErrorMessage("The Project Type is unknown." & vbCrLf)
                    InputXDoc = Nothing
                Case ProjectTypes.Directory
                    If System.IO.File.Exists(ProjectPath & "\" & InputFileName) Then
                        InputXDoc = XDocument.Load(ProjectPath & "\" & InputFileName)
                    Else
                        InputXDoc = Nothing
                    End If
                Case ProjectTypes.Archive
                    If System.IO.File.Exists(ProjectPath) Then 'Project Archive found.
                        Zip = New ADVL_Utilities_Library_1.ZipComp
                        Zip.ArchivePath = ProjectPath
                        If Zip.EntryExists(InputFileName) Then
                            InputXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText(InputFileName))
                        Else
                            RaiseEvent ErrorMessage("The Project Information file was not found in the Archive: " & InputFileName & vbCrLf)
                            InputXDoc = Nothing
                        End If
                    Else 'Project Archive not found.
                        RaiseEvent ErrorMessage("The Project Archive was not found: " & ProjectPath & vbCrLf)
                        InputXDoc = Nothing
                    End If
            End Select

        End Sub

        Private Sub SaveXDoc(ByRef XDoc As System.Xml.Linq.XDocument, ByVal FileName As String)
            'Saves the XDocument in the Project as FileName.

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("The XDocument is blank." & vbCrLf)
            Else
                Select Case ProjectType
                    Case ProjectTypes.Unknown
                        RaiseEvent ErrorMessage("The Project Type is unknown." & vbCrLf)
                    Case ProjectTypes.Directory
                        XDoc.Save(ProjectPath & "\" & FileName)
                    Case ProjectTypes.Archive
                        Zip = New ADVL_Utilities_Library_1.ZipComp
                        Zip.ArchivePath = ProjectPath
                        Zip.AddText(FileName, XDoc.ToString)
                End Select
            End If

        End Sub

        Private Sub ConvertADVL1ToADVL2(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Convert the Project Information file in format ADVL_1 to a the format ADVL_2.

            'Check if the ADVL_2 format version of the Project Information file already exists:
            If FileExists("Project_Info_ADVL_2.xml") Then
                RaiseEvent Message("The ADVL_2 format version of the Project Information Xml document already exisits. No conversion is necessary." & vbCrLf)
                Exit Sub
            End If

            'CODE ADDED 24Aug18:
            'Generate the Project ID:
            'Dim IDString As String = Name & " " & Format(CreationDate, "d-MMM-yyyy H:mm:ss")
            Dim IDString As String = InputXDoc.<Project>.<Name>.Value & " " & Format(InputXDoc.<Project>.<CreationDate>.Value, "d-MMM-yyyy H:mm:ss")

            'CODE ADDED 25Aug18:
            'A Directory project will have RelativeLocationType = Directory and Path = ""
            'An Archive project will have RelativeLocationType = Archive and Path = ""
            'There are no ADVL_1 format version Hybrid projects so no conversion of this type is required.
            'The relative location code has been updated from the follwing version:
            '<SettingsRelativeLocation>
            '    <Type>Directory</Type>
            '    <Path></Path>
            '</SettingsRelativeLocation>
            '<DataRelativeLocation>
            '    <Type>Directory</Type>
            '    <Path></Path>
            '</DataRelativeLocation>


            'NOTE: Host Project section added 22Aug18 - This info is no included in ADVL_1 files - blank values added during conversion.
            Try
                Dim newFile = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Project Information File-->
                              <Project>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <Name><%= InputXDoc.<Project>.<Name>.Value %></Name>
                                  <Type><%= InputXDoc.<Project>.<Type>.Value %></Type>
                                  <Description><%= InputXDoc.<Project>.<Description>.Value %></Description>
                                  <CreationDate><%= InputXDoc.<Project>.<CreationDate>.Value %></CreationDate>
                                  <ID><%= IDString.GetHashCode %></ID>
                                  <HostProject>
                                      <Name></Name>
                                      <DirectoryName></DirectoryName>
                                      <CreationDate>1-Jan-2000 12:00:00</CreationDate>
                                      <ID></ID>
                                  </HostProject>
                                  <Author>
                                      <Name><%= InputXDoc.<Project>.<Author>.<Name>.Value %></Name>
                                      <Description><%= InputXDoc.<Project>.<Author>.<Description>.Value %></Description>
                                      <Contact><%= InputXDoc.<Project>.<Author>.<Contact>.Value %></Contact>
                                  </Author>
                                  <SettingsRelativeLocation>
                                      <Type><%= InputXDoc.<Project>.<Type>.Value %></Type>
                                      <Path></Path>
                                  </SettingsRelativeLocation>
                                  <DataRelativeLocation>
                                      <Type><%= InputXDoc.<Project>.<Type>.Value %></Type>
                                      <Path></Path>
                                  </DataRelativeLocation>
                                  <Application>
                                      <Name><%= InputXDoc.<Project>.<ApplicationSummary>.<Name>.Value %></Name>
                                      <Description><%= InputXDoc.<Project>.<ApplicationSummary>.<Description>.Value %></Description>
                                      <CreationDate><%= InputXDoc.<Project>.<ApplicationSummary>.<CreationDate>.Value %></CreationDate>
                                      <Version>
                                          <Major><%= InputXDoc.<Project>.<ApplicationSummary>.<Version>.<Major>.Value %></Major>
                                          <Minor><%= InputXDoc.<Project>.<ApplicationSummary>.<Version>.<Minor>.Value %></Minor>
                                          <Build><%= InputXDoc.<Project>.<ApplicationSummary>.<Version>.<Build>.Value %></Build>
                                          <Revision><%= InputXDoc.<Project>.<ApplicationSummary>.<Version>.<Revision>.Value %></Revision>
                                      </Version>
                                      <Author>
                                          <Name><%= InputXDoc.<Project>.<ApplicationSummary>.<Author>.<Name>.Value %></Name>
                                          <Description><%= InputXDoc.<Project>.<ApplicationSummary>.<Author>.<Description>.Value %></Description>
                                          <Contact><%= InputXDoc.<Project>.<ApplicationSummary>.<Author>.<Contact>.Value %></Contact>
                                      </Author>
                                  </Application>
                              </Project>

                'OLD CODE BELOW. THIS CODE WAS UPDATED 24AUG18 - Local changed to Directory (Local is not a valid LocationType)
                '<SettingsRelativeLocation>
                '    <Type>Local</Type>
                '    <Path></Path>
                '</SettingsRelativeLocation>
                '<DataRelativeLocation>
                '    <Type>Local</Type>
                '    <Path></Path>
                '</DataRelativeLocation>

                'Save the new version of the Project Information file:
                SaveXDoc(newFile, "Project_Info_ADVL_2.xml")

            Catch ex As Exception
                'Main.Message.AddWarning("Project Information format conversion error: " & ex.Message & vbCrLf)
                RaiseEvent ErrorMessage("Project Information format conversion error: " & ex.Message & vbCrLf)
            End Try
        End Sub

        Public Function InputXDoc() As System.Xml.Linq.XDocument
            'Return the Input Project Information XDocument to be converted:

            Select Case ProjectType
                Case ProjectTypes.Unknown
                    RaiseEvent ErrorMessage("The Project Type is unknown." & vbCrLf)
                    Return Nothing
                Case ProjectTypes.Directory
                    If System.IO.File.Exists(ProjectPath & "\" & InputFileName) Then
                        Return XDocument.Load(ProjectPath & "\" & InputFileName)
                    Else
                        Return Nothing
                    End If
                Case ProjectTypes.Archive
                    If System.IO.File.Exists(ProjectPath) Then 'Project Archive found.
                        Zip = New ADVL_Utilities_Library_1.ZipComp
                        Zip.ArchivePath = ProjectPath
                        If Zip.EntryExists(InputFileName) Then
                            'Return XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
                            Return XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText(InputFileName))
                        Else
                            RaiseEvent ErrorMessage("The Project Information file was not found in the Archive: " & InputFileName & vbCrLf)
                            Return Nothing
                        End If
                    Else 'Project Archive not found.
                        RaiseEvent ErrorMessage("The Project Archive was not found: " & ProjectPath & vbCrLf)
                        Return Nothing
                    End If
            End Select

        End Function

        Public Function ConvertedXDoc() As System.Xml.Linq.XDocument
            'Return the converted Project Information XDocument:
            Select Case ProjectType
                Case ProjectTypes.Unknown
                    RaiseEvent ErrorMessage("The Project Type is unknown." & vbCrLf)
                    Return Nothing
                Case ProjectTypes.Directory
                    If System.IO.File.Exists(ProjectPath & "\Project_Info_ADVL_2.xml") Then
                        Return XDocument.Load(ProjectPath & "\Project_Info_ADVL_2.xml")
                    Else
                        Return Nothing
                    End If
                Case ProjectTypes.Archive
                    If System.IO.File.Exists(ProjectPath) Then 'Project Archive found.
                        Zip = New ADVL_Utilities_Library_1.ZipComp
                        Zip.ArchivePath = ProjectPath
                        If Zip.EntryExists("Project_Info_ADVL_2.xml") Then
                            Return XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
                        Else
                            RaiseEvent ErrorMessage("The Project Information file was not found in the Archive: " & "Project_Info_ADVL_2.xml" & vbCrLf)
                            Return Nothing
                        End If
                    Else 'Project Archive not found.
                        RaiseEvent ErrorMessage("The Project Archive was not found: " & ProjectPath & vbCrLf)
                        Return Nothing
                    End If
            End Select
        End Function

        Public Sub GetInputFormatCode(ByVal InputXDoc As System.Xml.Linq.XDocument)
            'Set the FormatCode of the Input Application Information file.

            If InputXDoc.<Application>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Application Information Xml document does not contain the Application element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<Application>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Application Information Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<Application>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

        Public Sub GetInputFormatCode()
            'Version of the GetInputFormatCode that reads the Application Information into InputXDoc.

            Dim InputXDoc As System.Xml.Linq.XDocument 'The Input XDocument containing the Application Information.

            If FileExists(InputFileName) Then
                GetInputXDoc(InputXDoc)
            Else
                RaiseEvent ErrorMessage("The Project Information file was not found: " & ProjectPath & "\" & InputFileName & vbCrLf)
                Exit Sub
            End If

            If InputXDoc.<Project>.Value = Nothing Then
                RaiseEvent ErrorMessage("The Project Information Xml document does not contain the Project element." & vbCrLf)
                InputFormatCode = FormatCodes.Unknown
            ElseIf InputXDoc.<Project>.<FormatCode>.Value = Nothing Then
                RaiseEvent Message("No FormatCode element found in the Project Information Xml document. ADVL_1 format assumed." & vbCrLf)
                InputFormatCode = FormatCodes.ADVL_1
            Else
                Select Case InputXDoc.<Project>.<FormatCode>.Value
                    Case "ADVL_1"
                        InputFormatCode = FormatCodes.ADVL_1
                    Case "ADVL_2"
                        InputFormatCode = FormatCodes.ADVL_2
                    Case Else
                        InputFormatCode = FormatCodes.Unknown
                End Select
            End If

        End Sub

#End Region 'Methods ------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Public Event Message(ByVal Message As String) 'Send a message.
        Public Event ErrorMessage(ByVal Message As String) 'Send an error message.

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Class


End Class
