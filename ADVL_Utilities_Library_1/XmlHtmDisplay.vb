Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms

Public Class XmlHtmDisplay
    'Rich Text Box control with XML document and HTML document formatting.

    Inherits RichTextBox

    Public WithEvents Settings As New XmlHtmDisplaySettings


    Public Sub ReadXmlFile(ByVal FilePath As String, ByVal ShowDecl As Boolean)
        'Read and display the XML data in the file.
        'Debug.Print("Starting XmlHtmDisplay.ReadXmlFile")

        If System.IO.File.Exists(FilePath) Then
            'Dim FileSize As Integer
            Dim myFileInfo As New System.IO.FileInfo(FilePath)
            If myFileInfo.Length >= Settings.XmlLargeFileSizeLimit Then
                If MessageBox.Show("File: " & FilePath & " (" & myFileInfo.Length & " bytes)" & " is larger than the large XML file limit of: " & Settings.XmlLargeFileSizeLimit & vbCrLf & "Are you sure you want to open this?", "Warning", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    'Dim xmlDoc As New System.Xml.XmlDocument
                    'xmlDoc.Load(FilePath)
                    'Me.Rtf = XmlToRtf(xmlDoc, ShowDecl)
                    Try
                        Dim xmlDoc As New System.Xml.XmlDocument
                        xmlDoc.Load(FilePath)
                        Me.Rtf = XmlToRtf(xmlDoc, ShowDecl)
                    Catch ex As Exception
                        RaiseEvent ErrorMessage("Error reading XML file: " & FilePath & " " & vbCrLf & ex.Message & vbCrLf)
                    End Try
                Else
                    RaiseEvent ErrorMessage("Read XML file cancelled because of the large file size." & vbCrLf)
                End If
            Else
                Try
                    Dim xmlDoc As New System.Xml.XmlDocument
                    xmlDoc.Load(FilePath)
                    Me.Rtf = XmlToRtf(xmlDoc, ShowDecl)
                Catch ex As Exception
                    RaiseEvent ErrorMessage("Error reading XML file: " & FilePath & " " & vbCrLf & ex.Message & vbCrLf)
                End Try

            End If
        Else
            RaiseEvent ErrorMessage("Xml file does not exist: " & FilePath & vbCrLf)
        End If

    End Sub

    Public Function SaveXmlFile(ByVal FilePath As String) As Boolean
        'Save the XML data to the specified file path.
        'Returns True if the file was saved OK.
        'Debug.Print("Starting XmlHtmDisplay.SaveXmlFile")

        Try
            Dim xmlDoc As New System.Xml.XmlDocument
            'xmlDoc.LoadXml(Me.Text)
            xmlDoc.LoadXml(Me.Text.Replace("&", "&amp;"))
            xmlDoc.Save(FilePath)
            Return True
        Catch ex As Exception
            RaiseEvent ErrorMessage("Error saving XML file: " & FilePath & " " & vbCrLf & ex.Message & vbCrLf)
            Return False
        End Try
    End Function

    'Note: The LoadFile method of the inherited RichTextBox is used to read an RTF file.

    Public Function SaveRtfFile(ByVal FilePath As String) As Boolean
        'Save the RTF data to the specified file path.
        'Returns True if the fiule was saved OK.
        'Debug.Print("Starting XmlHtmDisplay.SaveRtfFile")

        Try
            Me.SaveFile(FilePath)
            Return True
        Catch ex As Exception
            RaiseEvent ErrorMessage("Error saving RTF file: " & FilePath & " " & vbCrLf & ex.Message & vbCrLf)
            Return False
        End Try
    End Function

    Public Function TextToRtf(ByVal InputText As String, ByVal TypeName As String) As String
        'Create a RTF string to display the input text using the text type indicated by TypeName.
        'Debug.Print("Starting XmlHtmDisplay.TextToRtf")

        Try
            If Settings.TextType.ContainsKey(TypeName) Then
                'Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                '          & "{{\fonttbl{{{0}}}}}" _
                '          & "{{\colortbl ;{1}}}" _
                '          & "\viewkind4\uc1\pard\lang1033" _
                '          & Settings.RtfTextSettings(Settings.XTag) _
                '          & "{2}}}"
                Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                          & "{{\fonttbl{{{0}}}}}" _
                          & "{{\colortbl ;{1}}}" _
                          & "\viewkind4\uc1\pard\lang1033" _
                          & "{2}}}"

                'RaiseEvent Message("rtfFormat = " & rtfFormat & vbCrLf & vbCrLf)

                Dim textRtfContent As New StringBuilder()
                textRtfContent.Append(Settings.RtfTextSettings(Settings.TextType(TypeName)))
                'Code to escape \ characters:
                InputText = InputText.Replace("\", "\\")
                'Code to add CrLf:
                InputText = InputText.Replace(vbCrLf, "\line ")

                textRtfContent.Append(InputText)

                'RaiseEvent Message("textRtfContent = " & textRtfContent.ToString & vbCrLf & vbCrLf)

                'RaiseEvent Message("TextToRtf = " & String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), textRtfContent.ToString()) & vbCrLf & vbCrLf)

                Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), textRtfContent.ToString())
            Else
                RaiseEvent ErrorMessage("Text type not found: " & TypeName & vbCrLf)
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function PlainTextToRtf(ByVal InputText As String) As String
        'Create a RTF string to display the input text using Plain Text settings.

        Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
          & "{{\fonttbl{{{0}}}}}" _
          & "{{\colortbl ;{1}}}" _
          & "\viewkind4\uc1\pard\lang1033" _
          & "{2}}}"

        Dim textRtfContent As New StringBuilder()
        textRtfContent.Append(Settings.RtfTextSettings(Settings.PlainText))

        'Code to escape \ characters:
        InputText = InputText.Replace("\", "\\")
        'Code to add CrLf:
        InputText = InputText.Replace(vbCrLf, "\line ")

        textRtfContent.Append(InputText)

        Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), textRtfContent.ToString())

    End Function

    Public Function PlainTextFont() As Font
        'Return the Plain Text Font

        Dim myStyle As FontStyle = FontStyle.Regular
        If Settings.PlainText.Bold = True Then
            myStyle = myStyle & FontStyle.Bold
        End If
        If Settings.PlainText.Italic = True Then
            myStyle = myStyle & FontStyle.Italic
        End If
        Dim myFont As Font = New Font(Settings.PlainText.FontName, Settings.PlainText.PointSize, myStyle)
        Return myFont
    End Function

    Public Function DefaultTextToRtf(ByVal InputText As String) As String
        'Create a RTF string to display the input text using Default Text settings.

        Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
          & "{{\fonttbl{{{0}}}}}" _
          & "{{\colortbl ;{1}}}" _
          & "\viewkind4\uc1\pard\lang1033" _
          & "{2}}}"

        Dim textRtfContent As New StringBuilder()
        textRtfContent.Append(Settings.RtfTextSettings(Settings.DefaultText))

        'Code to escape \ characters:
        InputText = InputText.Replace("\", "\\")
        'Code to add CrLf:
        InputText = InputText.Replace(vbCrLf, "\line ")

        textRtfContent.Append(InputText)

        Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), textRtfContent.ToString())
    End Function

    Public Function DefaultTextFont() As Font
        'Return the Default Text Font

        Dim myStyle As FontStyle = FontStyle.Regular
        If Settings.DefaultText.Bold = True Then
            myStyle = myStyle & FontStyle.Bold
        End If
        If Settings.DefaultText.Italic = True Then
            myStyle = myStyle & FontStyle.Italic
        End If

        Dim myFont As Font = New Font(Settings.DefaultText.FontName, Settings.DefaultText.PointSize, myStyle)

        Return myFont

    End Function

    Public Function XmlToRtf(ByRef xmlDoc As System.Xml.XmlDocument, ByVal ShowDecl As Boolean) As String
        'Return the RTF code corresponding to the XML.
        'If ShowDecl is true, show the XMK declaration line.
        'Debug.Print("Starting XmlHtmDisplay.XmlToRtf")

        If xmlDoc Is Nothing Then
            RaiseEvent ErrorMessage("xmlDoc is nothing" & vbCrLf)
            Exit Function
        End If

        Try
            ' The Rtf contains 2 parts, header and content. The colortbl is a part of
            ' the header, and the {1} will be replaced with the content.

            Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                                      & "{{\fonttbl{{{0}}}}}" _
                                      & "{{\colortbl ;{1}}}" _
                                      & "\viewkind4\uc1\pard\lang1033" _
                                      & Settings.RtfTextSettings(Settings.XTag) _
                                      & "{2}}}"

            'NOTE: TO USE Settings.RtfHeader in a Format statement, "{" characters should be replaced by "{{" characters.

            Dim xmlRtfContent As New StringBuilder()

            Dim I As Integer
            Dim J As Integer
            For I = 0 To xmlDoc.ChildNodes().Count - 1
                If xmlDoc.ChildNodes(I).Attributes Is Nothing Then
                    'RaiseEvent Message("No attributes." & vbCrLf)
                End If

                Select Case xmlDoc.ChildNodes(I).NodeType
                    Case Xml.XmlNodeType.XmlDeclaration
                        Dim XDec As System.Xml.XmlDeclaration
                        XDec = xmlDoc.ChildNodes(I)

                        'If Settings.ShowDeclaration = True Then
                        If ShowDecl = True Then
                            ' The constants in XMLViewerSettings are used to specify the order 
                            ' in colortbl of the Rtf.
                            xmlRtfContent.AppendFormat(
                           "{0} <?{1} xml {2} version\cf{0} =\cf0 ""{3} {4}\cf0 "" " _
                           & "{2} encoding{0} =\cf0 ""{3} {5}\cf0 ""{0} ?>\par",
                           Settings.RtfTextSettings(Settings.XTag),
                           Settings.RtfTextSettings(Settings.XElement),
                           Settings.RtfTextSettings(Settings.XAttributeKey),
                           Settings.RtfTextSettings(Settings.XAttributeValue),
                           XDec.Version,
                           XDec.Encoding)
                        End If

                    Case Xml.XmlNodeType.Comment
                        xmlRtfContent.AppendFormat(
                        "{0}{1} <!--{2} {3}{1} -->\par",
                        "",
                        Settings.RtfTextSettings(Settings.XTag),
                        Settings.RtfTextSettings(Settings.XComment),
                        xmlDoc.ChildNodes(I).Value)

                    Case Xml.XmlNodeType.Element
                        'This should be the root element!
                        ' Get the Rtf of the root element.
                        'RaiseEvent Message("Case Xml.XmlNodeType.Element" & vbCrLf)
                        'RaiseEvent Message("I = " & I & vbCrLf)
                        'RaiseEvent Message("Settings.Comment.Color.Name  = " & Settings.Comment.Color.Name & vbCrLf)
                        'RaiseEvent Message("xmlDoc.ChildNodes.Count  = " & xmlDoc.ChildNodes.Count & vbCrLf)
                        'RaiseEvent Message("xmlDoc.ChildNodes(I).Name  = " & xmlDoc.ChildNodes(I).Name & vbCrLf)

                        Dim rootRtfContent As String = ProcessNode(xmlDoc.ChildNodes(I), 0, Settings)
                        xmlRtfContent.Append(rootRtfContent)
                End Select
            Next

            Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), xmlRtfContent.ToString())

            'Catch xmlException As System.Xml.XmlException
            '    Throw New ApplicationException("Problem with input Xml. Error:" & xmlException.Message, xmlException)
            'Catch
            '    Throw
        Catch ex As Exception
            'RaiseEvent Message("XmlToRtf error: " & ex.Message & vbCrLf)
            RaiseEvent ErrorMessage("XmlToRtf error: " & ex.Message & vbCrLf)
        End Try

    End Function

    Public Function XmlToRtf(ByRef xDoc As System.Xml.Linq.XDocument, ByVal ShowDecl As Boolean) As String
        'Version of the XmlToRtf function that processes an XDocument instead of an XmlDocument.
        'Return the RTF code corresponding to the XML.
        'If ShowDecl is true, show the XMK declaration line.
        'Debug.Print("Starting XmlHtmDisplay.XmlToRtf")

        'If xmlDoc Is Nothing Then
        If xDoc Is Nothing Then
            'RaiseEvent Message("xmlDoc is nothing" & vbCrLf)
            'RaiseEvent ErrorMessage("xmlDoc is nothing" & vbCrLf)
            RaiseEvent ErrorMessage("xDoc is nothing" & vbCrLf)
            Exit Function
        End If

        Dim xmlDoc As New System.Xml.XmlDocument
        xmlDoc.LoadXml(xDoc.ToString)

        Try
            ' The Rtf contains 2 parts, header and content. The colortbl is a part of
            ' the header, and the {1} will be replaced with the content.

            Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                                      & "{{\fonttbl{{{0}}}}}" _
                                      & "{{\colortbl ;{1}}}" _
                                      & "\viewkind4\uc1\pard\lang1033" _
                                      & Settings.RtfTextSettings(Settings.XTag) _
                                      & "{2}}}"

            'NOTE: TO USE Settings.RtfHeader in a Format statement, "{" characters should be replaced by "{{" characters.

            Dim xmlRtfContent As New StringBuilder()

            Dim I As Integer
            Dim J As Integer
            For I = 0 To xmlDoc.ChildNodes().Count - 1
                If xmlDoc.ChildNodes(I).Attributes Is Nothing Then
                    'RaiseEvent Message("No attributes." & vbCrLf)
                End If

                Select Case xmlDoc.ChildNodes(I).NodeType
                    Case Xml.XmlNodeType.XmlDeclaration
                        Dim XDec As System.Xml.XmlDeclaration
                        XDec = xmlDoc.ChildNodes(I)

                        'If Settings.ShowDeclaration = True Then
                        If ShowDecl = True Then
                            ' The constants in XMLViewerSettings are used to specify the order 
                            ' in colortbl of the Rtf.
                            xmlRtfContent.AppendFormat(
                           "{0} <?{1} xml {2} version\cf{0} =\cf0 ""{3} {4}\cf0 "" " _
                           & "{2} encoding{0} =\cf0 ""{3} {5}\cf0 ""{0} ?>\par",
                           Settings.RtfTextSettings(Settings.XTag),
                           Settings.RtfTextSettings(Settings.XElement),
                           Settings.RtfTextSettings(Settings.XAttributeKey),
                           Settings.RtfTextSettings(Settings.XAttributeValue),
                           XDec.Version,
                           XDec.Encoding)
                        End If

                    Case Xml.XmlNodeType.Comment
                        xmlRtfContent.AppendFormat(
                        "{0}{1} <!--{2} {3}{1} -->\par",
                        "",
                        Settings.RtfTextSettings(Settings.XTag),
                        Settings.RtfTextSettings(Settings.XComment),
                        xmlDoc.ChildNodes(I).Value)

                    Case Xml.XmlNodeType.Element
                        'This should be the root element!
                        ' Get the Rtf of the root element.
                        'RaiseEvent Message("Case Xml.XmlNodeType.Element" & vbCrLf)
                        'RaiseEvent Message("I = " & I & vbCrLf)
                        'RaiseEvent Message("Settings.Comment.Color.Name  = " & Settings.Comment.Color.Name & vbCrLf)
                        'RaiseEvent Message("xmlDoc.ChildNodes.Count  = " & xmlDoc.ChildNodes.Count & vbCrLf)
                        'RaiseEvent Message("xmlDoc.ChildNodes(I).Name  = " & xmlDoc.ChildNodes(I).Name & vbCrLf)

                        Dim rootRtfContent As String = ProcessNode(xmlDoc.ChildNodes(I), 0, Settings)
                        xmlRtfContent.Append(rootRtfContent)
                End Select
            Next

            Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), xmlRtfContent.ToString())

            'Catch xmlException As System.Xml.XmlException
            '    Throw New ApplicationException("Problem with input Xml. Error:" & xmlException.Message, xmlException)
            'Catch
            '    Throw
        Catch ex As Exception
            'RaiseEvent Message("XmlToRtf error: " & ex.Message & vbCrLf)
            RaiseEvent ErrorMessage("XmlToRtf error: " & ex.Message & vbCrLf)
        End Try

    End Function

    Public Function XmlToRtf(ByRef xmlText As String, ByVal ShowDecl As Boolean) As String
        'Version of the XmlToRtf function that processes an xml string instead of an XmlDocument.
        'Return the RTF code corresponding to the XML string.
        'If ShowDecl is true, show the XMK declaration line.
        'Debug.Print("Starting XmlHtmDisplay.XmlToRtf - xml string version")

        'RaiseEvent Message("Running XmlToRtf()" & vbCrLf)

        Dim xmlDoc As New System.Xml.XmlDocument

        Try
            xmlDoc.LoadXml(xmlText)



        Catch ex As Exception
            RaiseEvent ErrorMessage("Problem converting text to XML: " & vbCrLf & ex.Message & vbCrLf)
        End Try


        If xmlDoc Is Nothing Then
            RaiseEvent ErrorMessage("xmlDoc is nothing" & vbCrLf)
            Exit Function
        End If

        'RaiseEvent Message("xmlDoc = " & vbCrLf & xmlDoc.ToString & vbCrLf) 'FOR TESTING *************************************************************
        'RaiseEvent Message("xmlDoc = " & vbCrLf & xmlDoc.Value & vbCrLf) 'FOR TESTING *************************************************************
        'RaiseEvent Message("xmlDoc = " & vbCrLf & xmlDoc.InnerXml.ToString & vbCrLf) 'FOR TESTING *************************************************************

        Try
            ' The Rtf contains 2 parts, header and content. The colortbl is a part of
            ' the header, and the {1} will be replaced with the content.

            Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                                      & "{{\fonttbl{{{0}}}}}" _
                                      & "{{\colortbl ;{1}}}" _
                                      & "\viewkind4\uc1\pard\lang1033" _
                                      & Settings.RtfTextSettings(Settings.XTag) _
                                      & "{2}}}"

            'NOTE: TO USE Settings.RtfHeader in a Format statement, "{" characters should be replaced by "{{" characters.

            Dim xmlRtfContent As New StringBuilder()

            Dim I As Integer
            Dim J As Integer
            For I = 0 To xmlDoc.ChildNodes().Count - 1
                If xmlDoc.ChildNodes(I).Attributes Is Nothing Then
                    'RaiseEvent Message("No attributes." & vbCrLf)
                End If

                Select Case xmlDoc.ChildNodes(I).NodeType
                    Case Xml.XmlNodeType.XmlDeclaration
                        Dim XDec As System.Xml.XmlDeclaration
                        XDec = xmlDoc.ChildNodes(I)

                        'If Settings.ShowDeclaration = True Then
                        If ShowDecl = True Then
                            ' The constants in XMLViewerSettings are used to specify the order 
                            ' in colortbl of the Rtf.
                            xmlRtfContent.AppendFormat(
                           "{0} <?{1} xml {2} version\cf{0} =\cf0 ""{3} {4}\cf0 "" " _
                           & "{2} encoding{0} =\cf0 ""{3} {5}\cf0 ""{0} ?>\par",
                           Settings.RtfTextSettings(Settings.XTag),
                           Settings.RtfTextSettings(Settings.XElement),
                           Settings.RtfTextSettings(Settings.XAttributeKey),
                           Settings.RtfTextSettings(Settings.XAttributeValue),
                           XDec.Version,
                           XDec.Encoding)
                        End If

                    Case Xml.XmlNodeType.Comment
                        xmlRtfContent.AppendFormat(
                        "{0}{1} <!--{2} {3}{1} -->\par",
                        "",
                        Settings.RtfTextSettings(Settings.XTag),
                        Settings.RtfTextSettings(Settings.XComment),
                        xmlDoc.ChildNodes(I).Value)

                    Case Xml.XmlNodeType.Element
                        'This should be the root element!
                        ' Get the Rtf of the root element.

                        Dim rootRtfContent As String = ProcessNode(xmlDoc.ChildNodes(I), 0, Settings)
                        xmlRtfContent.Append(rootRtfContent)
                End Select
            Next

            Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), xmlRtfContent.ToString())

        Catch ex As Exception
            RaiseEvent ErrorMessage("XmlToRtf error: " & ex.Message & vbCrLf)
        End Try

    End Function

    'TRYING to write an XmlToRtf function that works when displaying an element containg a RegEx: <FileNameRegEx>(?<FileNameMatch>[\w\d]+)\.</FileNameRegex>
    'CURRENTLY the RegEx above does not appear correctly when displayed in an XmlHtmDisplay.
    'Public Function XmlToRtfAlt(ByRef xmlText As String, ByVal ShowDecl As Boolean) As String
    '    'Version of the XmlToRtf function that processes an xml string instead of an XmlDocument.
    '    'Return the RTF code corresponding to the XML string.
    '    'If ShowDecl is true, show the XMK declaration line.
    '    'Debug.Print("Starting XmlHtmDisplay.XmlToRtf - xml string version")

    '    Dim XDoc As New System.Xml.Linq.XDocument

    '    Try
    '        'xmlDoc.LoadXml(xmlText)
    '        XDoc = System.Xml.Linq.XDocument.Parse(xmlText)
    '    Catch ex As Exception
    '        RaiseEvent ErrorMessage("Problem converting text to XML: " & vbCrLf & ex.Message & vbCrLf)
    '    End Try

    '    If XDoc Is Nothing Then
    '        RaiseEvent ErrorMessage("XDoc is nothing" & vbCrLf)
    '        Exit Function
    '    End If


    '    Try
    '        ' The Rtf contains 2 parts, header and content. The colortbl is a part of
    '        ' the header, and the {1} will be replaced with the content.

    '        Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
    '                                  & "{{\fonttbl{{{0}}}}}" _
    '                                  & "{{\colortbl ;{1}}}" _
    '                                  & "\viewkind4\uc1\pard\lang1033" _
    '                                  & Settings.RtfTextSettings(Settings.XTag) _
    '                                  & "{2}}}"

    '        'NOTE: TO USE Settings.RtfHeader in a Format statement, "{" characters should be replaced by "{{" characters.

    '        Dim xmlRtfContent As New StringBuilder()

    '        Dim I As Integer
    '        Dim J As Integer
    '        'For I = 0 To xmlDoc.ChildNodes().Count - 1
    '        'For I = 0 To XDoc.DescendantNodes().Count - 1
    '        For I = 0 To XDoc.Nodes().Count - 1
    '            'If xmlDoc.ChildNodes(I).Attributes Is Nothing Then
    '            '    'RaiseEvent Message("No attributes." & vbCrLf)
    '            'End If

    '            'Select Case xmlDoc.ChildNodes(I).NodeType
    '            'Select Case XDoc.DescendantNodes(I).NodeType
    '            Select Case XDoc.Nodes(I).NodeType
    '                Case Xml.XmlNodeType.XmlDeclaration
    '                    Dim XDec As System.Xml.XmlDeclaration
    '                    'XDec = xmlDoc.ChildNodes(I)
    '                    'XDec = XDoc.DescendantNodes(I).Value
    '                    XDec = XDoc.Nodes(I).ElementsAfterSelf(0)

    '                    'If Settings.ShowDeclaration = True Then
    '                    If ShowDecl = True Then
    '                        ' The constants in XMLViewerSettings are used to specify the order 
    '                        ' in colortbl of the Rtf.
    '                        xmlRtfContent.AppendFormat(
    '                       "{0} <?{1} xml {2} version\cf{0} =\cf0 ""{3} {4}\cf0 "" " _
    '                       & "{2} encoding{0} =\cf0 ""{3} {5}\cf0 ""{0} ?>\par",
    '                       Settings.RtfTextSettings(Settings.XTag),
    '                       Settings.RtfTextSettings(Settings.XElement),
    '                       Settings.RtfTextSettings(Settings.XAttributeKey),
    '                       Settings.RtfTextSettings(Settings.XAttributeValue),
    '                       XDec.Version,
    '                       XDec.Encoding)
    '                    End If

    '                Case Xml.XmlNodeType.Comment
    '                    xmlRtfContent.AppendFormat(
    '                    "{0}{1} <!--{2} {3}{1} -->\par",
    '                    "",
    '                    Settings.RtfTextSettings(Settings.XTag),
    '                    Settings.RtfTextSettings(Settings.XComment),
    '                    'xmlDoc.ChildNodes(I).Value)
    '                    'XDoc.DescendantNodes(I).ElementsAfterSelf)
    '                    XDoc.Nodes(I).ElementsAfterSelf(0)

    '                Case Xml.XmlNodeType.Element
    '                    'This should be the root element!
    '                    ' Get the Rtf of the root element.

    '                    'Dim rootRtfContent As String = ProcessNode(xmlDoc.ChildNodes(I), 0, Settings)
    '                    'Dim rootRtfContent As String = ProcessNode(XDoc.DescendantNodes(I), 0, Settings)
    '                    'Dim rootRtfContent As String = ProcessNode(XDoc.Nodes(I), 0, Settings)
    '                    Dim rootRtfContent As String = ProcessNodeAlt(XDoc.Nodes(I), 0, Settings)
    '                    xmlRtfContent.Append(rootRtfContent)
    '            End Select
    '        Next

    '        Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), xmlRtfContent.ToString())

    '    Catch ex As Exception
    '        RaiseEvent ErrorMessage("XmlToRtf error: " & ex.Message & vbCrLf)
    '    End Try


    'End Function

    'Private Function ProcessNode(ByVal myNode As System.Xml.XmlNode, ByVal level As Integer, ByRef Settings As XmlDisplaySettings) As String
    Private Function ProcessNode(ByVal myNode As System.Xml.XmlNode, ByVal level As Integer, ByRef Settings As XmlHtmDisplaySettings) As String
        ' This viewer does not support the Xml file that has a Namespace.
        'Debug.Print("Starting XmlHtmDisplay.ProcessNode")

        If myNode.NodeType = Xml.XmlNodeType.Element Then
            If Not String.IsNullOrEmpty(myNode.NamespaceURI) Then
                Throw New ApplicationException(
                    "The Xml viewer does not support the Xml file that has Namespace.")
            End If
        End If

        Dim elementRtfFormat As String = String.Empty
        Dim childElementsRtfContent As New StringBuilder()
        Dim attributesRtfContent As New StringBuilder()

        'Set the indent spaces.
        Dim indent As New String(" "c, Settings.XIndentSpaces * level)

        'RaiseEvent Message(vbCrLf & "ProcessNode: myNode.Name  = " & myNode.Name & vbCrLf)
        'RaiseEvent Message("ProcessNode: myNode.NodeType.ToString  = " & myNode.NodeType.ToString & vbCrLf)
        'RaiseEvent Message("ProcessNode: myNode.ChildNodes.Count  = " & myNode.ChildNodes.Count & vbCrLf)
        'RaiseEvent Message("ProcessNode: myNode.ChildNodes(0).NodeType.ToString  = " & myNode.ChildNodes(0).NodeType.ToString & vbCrLf) 'NOTE THIS PRODUCES AN ERROR IF THE CHILD NODE IS EMPTY

        'RaiseEvent Message("ProcessNode: Location 1" & vbCrLf)

        If myNode.NodeType = Xml.XmlNodeType.XmlDeclaration Then '-------------------------------------------------------------1
            'NOTE: ProcessElement should not find any XML Declarations!
            'RaiseEvent Message("ProcessNode: Location 1" & vbCrLf)
        ElseIf myNode.NodeType = Xml.XmlNodeType.Comment Then
            'RaiseEvent Message("ProcessNode: Location 2" & vbCrLf)
            'elementRtfFormat = String.Format("{0}{1} <!--{2} {{0}}{1} -->\par",
            '              indent,
            '              Settings.RtfTextSettings(Settings.Tag),
            '              Settings.RtfTextSettings(Settings.Element))
            elementRtfFormat = String.Format("{0}{1} <!--{2} {3}{1} -->\par",
                          indent,
                          Settings.RtfTextSettings(Settings.XTag),
                          Settings.RtfTextSettings(Settings.XComment),
                          myNode.Value)

            'Dim childElementRtfContent As String = ProcessNode(myNode.NextSibling, level, Settings)
            'childElementsRtfContent.Append(childElementRtfContent)
            If myNode.NextSibling Is Nothing Then
            Else
                Dim childElementRtfContent As String = ProcessNode(myNode.NextSibling, level, Settings)
                childElementsRtfContent.Append(childElementRtfContent)
            End If

            'Process CDATA node:
        ElseIf myNode.NodeType = Xml.XmlNodeType.CDATA Then

            Dim charIndent As New String(" "c, Settings.XIndentSpaces * (level + 1)) 'The characters are displayed with an extra indent.

            elementRtfFormat = String.Format("{0}{1} <![CDATA[\line {4}{2} {3}{1} \line{0}]]>\par",
              indent,
              Settings.RtfTextSettings(Settings.XTag),
              Settings.RtfTextSettings(Settings.XValue),
              myNode.Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line " & charIndent),
              charIndent)

        ElseIf myNode.NodeType = Xml.XmlNodeType.Element Then
            'RaiseEvent Message("ProcessNode: Location 3" & vbCrLf)

            'If the element has child elements or value, then add the element to the Rtf.
            '{{0}} will be replaced with the attributes and {{1}} will be replaced
            'with the child elements or value.
            If myNode.HasChildNodes Then 'One or more child nodes. --------------------------------------------------------------
                If myNode.ChildNodes.Count = 1 Then '-------------------------------------------------------------------------------
                    If myNode.ChildNodes(0).NodeType = Xml.XmlNodeType.Text Then '-----------------------------------------------------
                        'RaiseEvent Message("Processing text node: " & myNode.ChildNodes(0).Name & " with one child node." & vbCrLf)
                        'If myNode.ChildNodes(0).Value = Nothing Then
                        '    RaiseEvent Message("The text node: " & myNode.ChildNodes(0).Name & " has no value!!!" & vbCrLf)
                        'End If
                        elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >" & "{{1}}" & "{1} </{2} {3}{1} >\par",
                          indent,
                          Settings.RtfTextSettings(Settings.XTag),
                          Settings.RtfTextSettings(Settings.XElement),
                          myNode.Name)

                        childElementsRtfContent.AppendFormat(
                          "{0} {1}",
                          Settings.RtfTextSettings(Settings.XValue),
                          myNode.ChildNodes(0).Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line "))
                        'Beep()

                        'myNode.ChildNodes(0).Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line "))
                        'myNode.ChildNodes(0).Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line ").Replace("<", " <"))
                        'myNode.ChildNodes(0).Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line ").Replace("<", "\<"))
                        'myNode.ChildNodes(0).Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line ").Replace("<", "&lt;").Replace(">", "&gt;")) 'Doesnt work: <FileNameRegEx>(?&lt;FileNameMatch&gt;[\w\d]+)\.</FileNameRegEx>
                        'NOTE: "\\" is used to display a "\" in rich text format. "\line " is used for a new line in rich text format.
                        'The Replace functions in the last line apply these changes to produce the correct rich text format.

                        'myNode.ChildNodes(0).Value.Trim())
                    Else 'Non-text child node.

                        'TESTING CODE TO HANDLE A NON-TEXT NODE: =======================================================================================
                        elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >\par" & "{{1}}" & "{0}{1} </{2} {3}{1} >\par",
                              indent,
                              Settings.RtfTextSettings(Settings.XTag),
                              Settings.RtfTextSettings(Settings.XElement),
                              myNode.Name)
                        'Dim childElementRtfContent As String = ProcessNode(myNode.ChildNodes(0), level + 1, Settings)
                        Dim childElementRtfContent As String = ProcessNode(myNode.ChildNodes(0), level + 1, Settings)
                        childElementsRtfContent.Append(childElementRtfContent)
                        'END OF TEST CODE ==============================================================================================================

                    End If '-----------------------------------------------------------------------------------------------------------
                Else 'Two or more child nodes.
                    elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >\par" & "{{1}}" & "{0}{1} </{2} {3}{1} >\par",
                          indent,
                          Settings.RtfTextSettings(Settings.XTag),
                          Settings.RtfTextSettings(Settings.XElement),
                          myNode.Name)
                    For Each Node In myNode.ChildNodes
                        Dim childElementRtfContent As String = ProcessNode(Node, level + 1, Settings)
                        childElementsRtfContent.Append(childElementRtfContent)
                    Next

                End If '------------------------------------------------------------------------------------------------------------
            Else
                elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} />\par",
                                     indent,
                                     Settings.RtfTextSettings(Settings.XTag),
                                     Settings.RtfTextSettings(Settings.XElement),
                                     myNode.Name)
            End If '-------------------------------------------------------------------------------------------------------------
            ' Construct the Rtf of the attributes.
            If myNode.Attributes.Count > 0 Then
                'For Each attribute As XAttribute In myNode.Attributes
                For Each attribute As Xml.XmlAttribute In myNode.Attributes
                    Dim attributeRtfContent As String =
                        String.Format(" {0} {3}{1} =\cf0 ""{2} {4}\cf0 """,
                                      Settings.RtfTextSettings(Settings.XAttributeKey),
                                      Settings.RtfTextSettings(Settings.XTag),
                                      Settings.RtfTextSettings(Settings.XAttributeValue),
                                      attribute.Name,
                                      attribute.Value)

                    attributesRtfContent.Append(attributeRtfContent)
                Next attribute
                attributesRtfContent.Append(" ")
            Else
                'RaiseEvent Message("No Attributes. " & vbCrLf & vbCrLf)
            End If

        End If '---------------------------------------------------------------------------------------------------------------

        Return String.Format(elementRtfFormat, attributesRtfContent, childElementsRtfContent)

    End Function

    ''The alternative version of ProcessNode: uses System.Xml.Linq.XDocument foramt instead of System.Xml.XmlDocument
    ''Private Function ProcessNodeAlt(ByVal myNode As System.Xml.XmlNode, ByVal level As Integer, ByRef Settings As XmlHtmDisplaySettings) As String
    'Private Function ProcessNodeAlt(ByVal myNode As System.Xml.Linq.XNode, ByVal level As Integer, ByRef Settings As XmlHtmDisplaySettings) As String
    '    ' This viewer does not support the Xml file that has a Namespace.

    '    'If myNode.NodeType = Xml.XmlNodeType.Element Then
    '    '    If Not String.IsNullOrEmpty(myNode.NamespaceURI) Then
    '    '        Throw New ApplicationException(
    '    '            "The Xml viewer does not support the Xml file that has Namespace.")
    '    '    End If
    '    'End If

    '    Dim elementRtfFormat As String = String.Empty
    '    Dim childElementsRtfContent As New StringBuilder()
    '    Dim attributesRtfContent As New StringBuilder()

    '    'Set the indent spaces.
    '    Dim indent As New String(" "c, Settings.XIndentSpaces * level)

    '    If myNode.NodeType = Xml.XmlNodeType.XmlDeclaration Then '-------------------------------------------------------------1
    '        'NOTE: ProcessElement should not find any XML Declarations!
    '    ElseIf myNode.NodeType = Xml.XmlNodeType.Comment Then
    '        elementRtfFormat = String.Format("{0}{1} <!--{2} {3}{1} -->\par",
    '                      indent,
    '                      Settings.RtfTextSettings(Settings.XTag),
    '                      Settings.RtfTextSettings(Settings.XComment),
    '                      myNode.ElementsAfterSelf(0))
    '        '      myNode.Value)

    '        'If myNode.NextSibling Is Nothing Then
    '        If myNode.NextNode Is Nothing Then
    '        Else
    '            'Dim childElementRtfContent As String = ProcessNode(myNode.NextSibling, level, Settings)
    '            Dim childElementRtfContent As String = ProcessNodeAlt(myNode.NextNode, level, Settings)
    '            childElementsRtfContent.Append(childElementRtfContent)
    '        End If

    '        'Process CDATA node:
    '    ElseIf myNode.NodeType = Xml.XmlNodeType.CDATA Then

    '        Dim charIndent As New String(" "c, Settings.XIndentSpaces * (level + 1)) 'The characters are displayed with an extra indent.

    '        elementRtfFormat = String.Format("{0}{1} <![CDATA[\line {4}{2} {3}{1} \line{0}]]>\par",
    '          indent,
    '          Settings.RtfTextSettings(Settings.XTag),
    '          Settings.RtfTextSettings(Settings.XValue),
    '          myNode.ElementsAfterSelf(0).ToString.Trim().Replace("\", "\\").Replace(vbCrLf, "\line " & charIndent),
    '          charIndent)
    '        '              myNode.Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line " & charIndent),
    '    ElseIf myNode.NodeType = Xml.XmlNodeType.Element Then
    '        'RaiseEvent Message("ProcessNode: Location 3" & vbCrLf)

    '        'If the element has child elements or value, then add the element to the Rtf.
    '        '{{0}} will be replaced with the attributes and {{1}} will be replaced
    '        'with the child elements or value.
    '        'If myNode.HasChildNodes Then 'One or more child nodes. --------------------------------------------------------------
    '        If myNode.NodesAfterSelf.Count > 0 Then 'One or more child nodes. --------------------------------------------------------------
    '            'If myNode.ChildNodes.Count = 1 Then '-------------------------------------------------------------------------------
    '            If myNode.NodesAfterSelf.Count = 1 Then '-------------------------------------------------------------------------------
    '                'If myNode.ChildNodes(0).NodeType = Xml.XmlNodeType.Text Then '-----------------------------------------------------
    '                If myNode.NodesAfterSelf(0).NodeType = Xml.XmlNodeType.Text Then '-----------------------------------------------------
    '                    elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >" & "{{1}}" & "{1} </{2} {3}{1} >\par",
    '                      indent,
    '                      Settings.RtfTextSettings(Settings.XTag),
    '                      Settings.RtfTextSettings(Settings.XElement),
    '                      myNode.Name)
    '                    'myNode.Name)

    '                    childElementsRtfContent.AppendFormat(
    '                      "{0} {1}",
    '                      Settings.RtfTextSettings(Settings.XValue),
    '                      myNode.ChildNodes(0).Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line "))
    '                    'NOTE: "\\" is used to display a "\" in rich text format. "\line " is used for a new line in rich text format.
    '                    'The Replace functions in the last line apply these changes to produce the correct rich text format.

    '                Else 'Non-text child node.

    '                    'TESTING CODE TO HANDLE A NON-TEXT NODE: =======================================================================================
    '                    elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >\par" & "{{1}}" & "{0}{1} </{2} {3}{1} >\par",
    '                          indent,
    '                          Settings.RtfTextSettings(Settings.XTag),
    '                          Settings.RtfTextSettings(Settings.XElement),
    '                          myNode.Name)
    '                    'Dim childElementRtfContent As String = ProcessNode(myNode.ChildNodes(0), level + 1, Settings)
    '                    Dim childElementRtfContent As String = ProcessNode(myNode.ChildNodes(0), level + 1, Settings)
    '                    childElementsRtfContent.Append(childElementRtfContent)
    '                    'END OF TEST CODE ==============================================================================================================

    '                End If '-----------------------------------------------------------------------------------------------------------
    '            Else 'Two or more child nodes.
    '                elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >\par" & "{{1}}" & "{0}{1} </{2} {3}{1} >\par",
    '                      indent,
    '                      Settings.RtfTextSettings(Settings.XTag),
    '                      Settings.RtfTextSettings(Settings.XElement),
    '                      myNode.Name)
    '                For Each Node In myNode.ChildNodes
    '                    Dim childElementRtfContent As String = ProcessNode(Node, level + 1, Settings)
    '                    childElementsRtfContent.Append(childElementRtfContent)
    '                Next

    '            End If '------------------------------------------------------------------------------------------------------------
    '        Else
    '            elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} />\par",
    '                                 indent,
    '                                 Settings.RtfTextSettings(Settings.XTag),
    '                                 Settings.RtfTextSettings(Settings.XElement),
    '                                 myNode.Name)
    '        End If '-------------------------------------------------------------------------------------------------------------
    '        ' Construct the Rtf of the attributes.
    '        If myNode.Attributes.Count > 0 Then
    '            'For Each attribute As XAttribute In myNode.Attributes
    '            For Each attribute As Xml.XmlAttribute In myNode.Attributes
    '                Dim attributeRtfContent As String =
    '                    String.Format(" {0} {3}{1} =\cf0 ""{2} {4}\cf0 """,
    '                                  Settings.RtfTextSettings(Settings.XAttributeKey),
    '                                  Settings.RtfTextSettings(Settings.XTag),
    '                                  Settings.RtfTextSettings(Settings.XAttributeValue),
    '                                  attribute.Name,
    '                                  attribute.Value)

    '                attributesRtfContent.Append(attributeRtfContent)
    '            Next attribute
    '            attributesRtfContent.Append(" ")
    '        Else
    '            'RaiseEvent Message("No Attributes. " & vbCrLf & vbCrLf)
    '        End If

    '    End If '---------------------------------------------------------------------------------------------------------------

    '    Return String.Format(elementRtfFormat, attributesRtfContent, childElementsRtfContent)
    'End Function


    Public Function FixXmlText(XmlText As String) As String
        'Fix an XML string so that it can be loaded correcly using the LoadXml method of a System.Xml.XmlDocument
        'Replace "<" in an element value with "&lt;"
        'Replace ">" in an element value with "&gt;"

        'XML Terminology:
        'XML declaration <?xml version="1.0" encoding="UTF-8"?>
        'Comments begin with <!-- and end with -->.
        'Start-tag <Element>
        'End-tag </Element>
        'Empty-element tag (Element />
        'Content  The characters between the start-tag and end-tag, if any, are the element's content, and may contain markup, including other elements, which are called child elements.
        'Predefined entities:
        '&lt; represents "<";
        '&gt; represents ">";
        '&amp; represents "&";
        '&apos; represents "'";
        '&quot; represents '"'.

        Dim FixedXmlText As New System.Text.StringBuilder
        Dim StartPos As Integer
        Dim EndPos As Integer
        Dim ScanPos As Integer = 0
        Dim LastPos As Integer = XmlText.Length

        If XmlText.Trim.StartsWith("<?xml") Then
            StartPos = XmlText.IndexOf("<?xml")
            EndPos = XmlText.IndexOf("?>", StartPos)
            FixedXmlText.Append(XmlText.Substring(StartPos, EndPos - StartPos + 2))
            ScanPos = EndPos + 2
        End If
        FixedXmlText.Append(ProcessContent(XmlText, ScanPos, LastPos))
        Return FixedXmlText.ToString
    End Function


    Private Function ProcessContent(ByRef XmlText As String, FromIndex As Integer, ToIndex As Integer) As String
        'Process the XML content in the XmlText string between FromIndex and ToIndex.
        'THIS VERSION SEARCHES FOR MATCHING End-Tags
        '
        'Content alternatives:
        'Content only
        '<!---->                                        One or more comments
        '<Element />                                    One or more empty element tags
        '<Element></Element>                            One or more empty elements
        '<Element>Content</Element>                     One or more elements containing content
        '<Element>                                      One or more elements containing child elements
        '  <ChildElement>ChildContent</ChildElement>
        '</Element>

        Dim StartScan As Integer = FromIndex 'The start of the current content scan
        Dim ScanIndex As Integer = FromIndex 'The current scan position
        Dim LtCharIndex As Integer 'The index position of the next < character
        Dim GtCharIndex As Integer 'The index position of the next > character
        Dim FixedXmlText As New System.Text.StringBuilder 'This is used to build the fixed XML text for the content if it contains XML tags
        Dim StartTagText As String = "" 'The text of a found Start-tag. The text may include attributes following the name.
        Dim EndNameIndex As Integer 'The index position of the end of the StartTagName. If the StartTagText contains attributes, the StartTagName will be followed by a space then the attributes.
        Dim StartTagName As String = "" 'The name of a found Start-tag
        Dim EndTagIndex As Integer 'The index of an End-tag
        Dim StartTagCount As Integer = 1 'The nesting level of the StartTag
        Dim EndTagCount As Integer = 1 'The nesting level of the EndTag
        Dim StartSearch As Integer 'StartSearch index used for counting other Start-Tags named StartTagName
        Dim NextSearch As Integer 'Search for the next Start-Tag named StartTagName
        Dim SearchIndex As Integer
        Dim Match As Boolean
        Dim TagLevelMatch As Boolean 'If True, the Start-Tag and End-Tag have matching levels.
        Dim SearchEndTagFrom As Integer 'The index to start the End-tag search from
        Dim ElementFound As Boolean = False 'True if an Element or a Comment was found.
        Dim EndSearch As Boolean 'If True, End the Search to find the End-Tag

        'While ScanIndex <= ToIndex
        While ScanIndex < ToIndex
            'Find the first pair of < > characters
            LtCharIndex = XmlText.IndexOf("<", ScanIndex) 'Find the start of the next Element
            If LtCharIndex = -1 Then '< char not found
                If ToIndex - ScanIndex = 2 Then
                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                        Exit While
                    End If
                End If
                'The characters between FromIndex and ToIndex are Content
                'NOTE: StartScan and FromIndex should be the same here: StartScan only advances if the Content contains one or more comments or elements.
                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                FixedXmlText.Append(Content)
                ScanIndex = ToIndex + 1
            ElseIf LtCharIndex >= ToIndex Then
                'Check if the remaining characters are CrLf:
                If ToIndex - ScanIndex = 2 Then
                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                        Exit While
                    End If
                End If
                'Check if the remaining characters are blank:
                If XmlText.Substring(ScanIndex, ToIndex - ScanIndex).Trim = "" Then
                    Exit While
                End If
                'Check if the remaining characters with blanks removed are CrLf:
                'Check if the remaining characters are blank:
                If XmlText.Substring(ScanIndex, ToIndex - ScanIndex).Trim = vbCrLf Then
                    Exit While
                End If
                'The characters between FromIndex and ToIndex are Content
                'NOTE: StartScan and FromIndex should be the same here
                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                FixedXmlText.Append(Content)
                ScanIndex = ToIndex + 1
            Else
                'The < character is within the Content range
                'Search for a > character
                GtCharIndex = XmlText.IndexOf(">", LtCharIndex + 1)
                If GtCharIndex = -1 Then '> char not found
                    If ToIndex - ScanIndex = 2 Then
                        If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                            Exit While
                        End If
                    End If
                    'The characters between FromIndex and ToIndex are Content
                    'NOTE: StartScan and FromIndex should be the same here
                    Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                    FixedXmlText.Append(Content)
                    ScanIndex = ToIndex + 1
                ElseIf GtCharIndex > ToIndex Then
                    If ToIndex - ScanIndex = 2 Then
                        If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
                            Exit While
                        End If
                    End If
                    'The characters between FromIndex and ToIndex are Content
                    'NOTE: StartScan and FromIndex should be the same here
                    Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                    FixedXmlText.Append(Content)
                    ScanIndex = ToIndex + 1
                Else
                    'A start-tagChar and end-tagChar <> pair has been found.
                    'The <> characters will contain a comment, an element name or be part of the element content.
                    If XmlText.Substring(LtCharIndex, 4) = "<!--" Then 'This is the start of a comment
                        If XmlText.Substring(GtCharIndex - 2, 3) = "-->" Then 'This is the end of a comment ---------------------  <--Comment-->  ---------------------------------
                            FixedXmlText.Append(XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1) & vbCrLf) 'Add the Comment to the Fixed XML Text
                            ScanIndex = GtCharIndex + 1
                            StartScan = GtCharIndex + 1
                        Else
                            'This is not a comment.
                            'The whole content must be the content of a single element
                            'The characters between FromIndex and ToIndex are Content
                            'NOTE: StartScan and FromIndex should be the same here
                            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                            FixedXmlText.Append(Content)
                            ScanIndex = ToIndex + 1
                        End If
                    Else 'This is a start-tag, empty element or content of a single element
                        If XmlText.Chars(GtCharIndex - 1) = "/" Then 'This is an empty element
                            FixedXmlText.Append(XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1))
                            ScanIndex = GtCharIndex + 1
                            StartScan = GtCharIndex + 1
                        Else
                            StartTagCount = 1
                            EndTagCount = 0
                            StartSearch = GtCharIndex + 1
                            EndSearch = False
                            While StartTagCount > EndTagCount And EndSearch = False
                                'Continue searching for StartTag-EndTag tag pairs with the name StartTagName until matching tags are found (StartTagCount = EndTagCount).
                                StartTagText = XmlText.Substring(LtCharIndex + 1, GtCharIndex - LtCharIndex - 1) 'This is the text of the Start-tag
                                EndNameIndex = StartTagText.IndexOf(" ")
                                If EndNameIndex = -1 Then 'There is no space in StartTagText so it contains no attributes.
                                    StartTagName = StartTagText
                                Else
                                    StartTagName = StartTagText.Substring(0, EndNameIndex)
                                End If
                                'Find the matching End-tag - The matching End-tag must have a matching TagName and a matching level.
                                TagLevelMatch = False
                                SearchEndTagFrom = GtCharIndex + 1
                                While TagLevelMatch = False
                                    EndTagIndex = XmlText.IndexOf("</" & StartTagName & ">", SearchEndTagFrom)
                                    If EndTagIndex = -1 Then 'There is no matching End-tag
                                        'This is not an element.
                                        'The whole content must be the content of a single element
                                        'The characters between FromIndex and ToIndex are Content
                                        'NOTE: StartScan and FromIndex should be the same here
                                        Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                                        FixedXmlText.Append(Content)
                                        ScanIndex = ToIndex + 1
                                        EndSearch = True 'End the search for the End-Tag
                                        Exit While 'There is no matching End-tag!
                                    ElseIf EndTagIndex > ToIndex - StartTagName.Length - 1 Then 'The matching tag is outside of the Content
                                        'This is not an element.
                                        'The whole content must be the content of a single element
                                        'The characters between FromIndex and ToIndex are Content
                                        'NOTE: StartScan and FromIndex should be the same here
                                        Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
                                        FixedXmlText.Append(Content)
                                        ScanIndex = ToIndex + 1
                                        EndSearch = True 'End the search for the End-Tag
                                        Exit While 'There is no matching End-tag withing the Content range.
                                    Else 'Matching End-tag found at EndTagIndex. 
                                        EndTagCount += 1 'Increment the End Tag Count
                                        'Search for any other Start-Tags named StartTagName between LtCharIndex and EndTagIndex
                                        Match = True
                                        NextSearch = StartSearch 'Search for <StartTagName> (without attributes)
                                        While Match = True 'Search for Start-Tags of the form: <StartTagName>
                                            SearchIndex = XmlText.IndexOf("<" & StartTagName & ">", NextSearch, EndTagIndex - NextSearch)
                                            If SearchIndex = -1 Then
                                                Match = False
                                            Else
                                                NextSearch = SearchIndex + StartTagName.Length
                                                StartTagCount += 1
                                            End If
                                        End While
                                        Match = True
                                        NextSearch = StartSearch 'Set NextSearch back to StartSearch to search the same chars for <StartTagName ...(with attributes)
                                        While Match = True 'Search for Start-Tags of the form: <StartTagName ...> (Start-Tag with attributes)
                                            SearchIndex = XmlText.IndexOf("<" & StartTagName & " ", NextSearch, EndTagIndex - NextSearch)
                                            If SearchIndex = -1 Then
                                                Match = False
                                            Else
                                                NextSearch = SearchIndex + StartTagName.Length
                                                StartTagCount += 1
                                            End If
                                        End While
                                        StartSearch = EndTagIndex + 1 'All Start-Tags named StartTagName have been found to EndTagIndex : Update StartSearch - If more searches are needed, they will start from here.
                                        If StartTagCount = EndTagCount Then
                                            TagLevelMatch = True
                                            FixedXmlText.Append("<" & StartTagText & ">" & ProcessContent(XmlText, GtCharIndex + 1, EndTagIndex) & "</" & StartTagName & ">" & vbCrLf)
                                            ScanIndex = EndTagIndex + StartTagName.Length + 3
                                            ElementFound = True
                                        Else
                                            SearchEndTagFrom = EndTagIndex + StartTagName.Length + 3
                                        End If
                                    End If
                                End While
                            End While
                        End If
                    End If
                End If
            End If
        End While
        If ElementFound Then
            Return vbCrLf & FixedXmlText.ToString
        Else
            Return FixedXmlText.ToString
        End If
    End Function

    'Private Function ProcessContent(ByRef XmlText As String, FromIndex As Integer, ToIndex As Integer) As String
    '    'Process the XML content in the XmlText string between FromIndex and ToIndex.
    '    'THIS VERSION SEARCHES FOR MATCHING End-Tags
    '    '
    '    'Content alternatives:
    '    'Content only
    '    '<!---->                                        One or more comments
    '    '<Element />                                    One or more empty element tags
    '    '<Element></Element>                            One or more empty elements
    '    '<Element>Content</Element>                     One or more elements containing content
    '    '<Element>                                      One or more elements containing child elements
    '    '  <ChildElement>ChildContent</ChildElement>
    '    '</Element>

    '    'NOTE: The commented-out Message code was used for testing the function within the ADVL_Import_1 application

    '    Dim StartScan As Integer = FromIndex 'The start of the current content scan
    '    Dim ScanIndex As Integer = FromIndex 'The current scan position
    '    Dim LtCharIndex As Integer 'The index position of the next < character
    '    Dim GtCharIndex As Integer 'The index position of the next > character
    '    Dim FixedXmlText As New System.Text.StringBuilder 'This is used to build the fixed XML text for the content if it contains XML tags
    '    Dim StartTagText As String = "" 'The text of a found Start-tag. The text may include attributes following the name.
    '    Dim EndNameIndex As Integer 'The index position of the end of the StartTagName. If the StartTagText contains attributes, the StartTagName will be followed by a space then the attributes.
    '    Dim StartTagName As String = "" 'The name of a found Start-tag
    '    Dim EndTagIndex As Integer 'The index of an End-tag
    '    Dim StartTagCount As Integer = 1 'The nesting level of the StartTag
    '    Dim EndTagCount As Integer = 1 'The nesting level of the EndTag
    '    Dim StartSearch As Integer 'StartSearch index used for counting other Start-Tags named StartTagName
    '    Dim NextSearch As Integer 'Search for the next Start-Tag named StartTagName
    '    Dim SearchIndex As Integer
    '    Dim Match As Boolean
    '    Dim TagLevelMatch As Boolean 'If True, the Start-Tag and End-Tag have matching levels.
    '    Dim SearchEndTagFrom As Integer 'The index to start the End-tag search from
    '    Dim ElementFound As Boolean = False 'True if an Element or a Comment was found.

    '    While ScanIndex <= ToIndex
    '        'Find the first pair of < > characters
    '        LtCharIndex = XmlText.IndexOf("<", ScanIndex) 'Find the start of the next Element
    '        If LtCharIndex = -1 Then '< char not found
    '            If ToIndex - ScanIndex = 2 Then
    '                If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                    'Message.Add("At end of line. Exit While" & vbCrLf)
    '                    'FixedXmlText.Append(vbCrLf)
    '                    Exit While
    '                End If
    '            End If
    '            'The characters between FromIndex and ToIndex are Content
    '            'NOTE: StartScan and FromIndex should be the same here: StartScan only advances if the Content contains one or more comments or elements.
    '            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '            FixedXmlText.Append(Content)
    '            'Message.Add("< char not found. Content string returned: " & Content & " ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '            ScanIndex = ToIndex + 1
    '        ElseIf LtCharIndex >= ToIndex Then
    '            'Check if the remaining characters are CrLf:
    '            If ToIndex - ScanIndex = 2 Then
    '                If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                    Exit While
    '                End If
    '            End If
    '            'Check if the remaining characters are blank:
    '            If XmlText.Substring(ScanIndex, ToIndex - ScanIndex).Trim = "" Then
    '                'Message.Add("The remaining characters are blank." & vbCrLf)
    '                Exit While
    '            End If
    '            'Check if the remaining characters with blanks removed are CrLf:
    '            'Check if the remaining characters are blank:
    '            If XmlText.Substring(ScanIndex, ToIndex - ScanIndex).Trim = vbCrLf Then
    '                'Message.Add("The remaining characters with blanks removed are CrLf." & vbCrLf)
    '                Exit While
    '            End If
    '            'The characters between FromIndex and ToIndex are Content
    '            'NOTE: StartScan and FromIndex should be the same here
    '            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '            FixedXmlText.Append(Content)
    '            'Message.Add("< char not found within Content range. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " FromIndex = " & FromIndex & " ToIndex = " & ToIndex & " LtCharIndex = " & LtCharIndex & vbCrLf)
    '            'Message.Add("String between ScanIndex and ToIndex: " & XmlText.Substring(ScanIndex, ToIndex - ScanIndex) & vbCrLf)
    '            ScanIndex = ToIndex + 1
    '        Else
    '            'The < character is within the Content range
    '            'Search for a > character
    '            GtCharIndex = XmlText.IndexOf(">", LtCharIndex + 1)
    '            If GtCharIndex = -1 Then '> char not found
    '                If ToIndex - ScanIndex = 2 Then
    '                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                        Exit While
    '                    End If
    '                End If
    '                'The characters between FromIndex and ToIndex are Content
    '                'NOTE: StartScan and FromIndex should be the same here
    '                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                FixedXmlText.Append(Content)
    '                'Message.Add("> char not found. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                ScanIndex = ToIndex + 1
    '            ElseIf GtCharIndex > ToIndex Then
    '                If ToIndex - ScanIndex = 2 Then
    '                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                        Exit While
    '                    End If
    '                End If
    '                'The characters between FromIndex and ToIndex are Content
    '                'NOTE: StartScan and FromIndex should be the same here
    '                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                FixedXmlText.Append(Content)
    '                'Message.Add("> char not found within Content range. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                ScanIndex = ToIndex + 1
    '            Else
    '                'A start-tagChar and end-tagChar <> pair has been found.
    '                'The <> characters will contain a comment, an element name or be part of the element content.
    '                If XmlText.Substring(LtCharIndex, 4) = "<!--" Then 'This is the start of a comment
    '                    If XmlText.Substring(GtCharIndex - 2, 3) = "-->" Then 'This is the end of a comment -------------------------------------------------------------  <--Comment-->  --------------------------------------------------------------
    '                        FixedXmlText.Append(XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1) & vbCrLf) 'Add the Comment to the Fixed XML Text
    '                        ScanIndex = GtCharIndex + 1
    '                        StartScan = GtCharIndex + 1
    '                        'Message.Add("Comment found and returned: " & XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1) & " ScanIndex = " & ScanIndex & vbCrLf)
    '                    Else
    '                        'This is not a comment.
    '                        'The whole content must be the content of a single element
    '                        'The characters between FromIndex and ToIndex are Content
    '                        'NOTE: StartScan and FromIndex should be the same here
    '                        Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                        FixedXmlText.Append(Content)
    '                        'Message.Add("Comment not found. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                        ScanIndex = ToIndex + 1
    '                    End If
    '                Else 'This is a start-tag, empty element or content of a single element
    '                    If XmlText.Chars(GtCharIndex - 1) = "/" Then 'This is an empty element
    '                        FixedXmlText.Append(XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1))
    '                        ScanIndex = GtCharIndex + 1
    '                        StartScan = GtCharIndex + 1
    '                        'Message.Add("Empty element found and returned. ScanIndex = " & ScanIndex & vbCrLf)
    '                    Else
    '                        StartTagCount = 1
    '                        EndTagCount = 0
    '                        StartSearch = GtCharIndex + 1
    '                        While StartTagCount > EndTagCount
    '                            'Continue searching for Start-End tag pairs with the name StartTagName until matching tags are found (StartTagCount = EndTagCount).
    '                            StartTagText = XmlText.Substring(LtCharIndex + 1, GtCharIndex - LtCharIndex - 1) 'This is the text of the Start-tag
    '                            EndNameIndex = StartTagText.IndexOf(" ")
    '                            If EndNameIndex = -1 Then 'There is no space in StartTagText so it contains no attributes.
    '                                StartTagName = StartTagText
    '                            Else
    '                                StartTagName = StartTagText.Substring(0, EndNameIndex)
    '                            End If
    '                            'Message.Add("StartTagName = " & StartTagName & vbCrLf)

    '                            'Find the matching End-tag - The matching End-tag must have a matching TagName and a matching level.
    '                            TagLevelMatch = False
    '                            SearchEndTagFrom = GtCharIndex + 1
    '                            While TagLevelMatch = False
    '                                'Message.Add("Searching for End-tag: " & "</" & StartTagName & ">" & vbCrLf)
    '                                EndTagIndex = XmlText.IndexOf("</" & StartTagName & ">", SearchEndTagFrom)
    '                                If EndTagIndex = -1 Then 'There is no matching End-tag
    '                                    'This is not an element.
    '                                    'The whole content must be the content of a single element
    '                                    'The characters between FromIndex and ToIndex are Content
    '                                    'NOTE: StartScan and FromIndex should be the same here
    '                                    Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                                    FixedXmlText.Append(Content)
    '                                    'Message.Add("Start-tag with no end-tag. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                                    ScanIndex = ToIndex + 1
    '                                ElseIf EndTagIndex > ToIndex - StartTagName.Length - 1 Then 'The matching tag is outside of the Content
    '                                    'This is not an element.
    '                                    'The whole content must be the content of a single element
    '                                    'The characters between FromIndex and ToIndex are Content
    '                                    'NOTE: StartScan and FromIndex should be the same here
    '                                    Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                                    FixedXmlText.Append(Content)
    '                                    'Message.Add("Start-tag with no end-tag within Content range. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                                    ScanIndex = ToIndex + 1
    '                                Else 'Matching End-tag found at EndTagIndex. 
    '                                    EndTagCount += 1 'Increment the End Tag Count
    '                                    'Search for any other Start-Tags named StartTagName between LtCharIndex and EndTagIndex
    '                                    Match = True
    '                                    NextSearch = StartSearch 'Search for <StartTagName> (without attributes)
    '                                    While Match = True 'Search for Start-Tags of the form: <StartTagName>
    '                                        SearchIndex = XmlText.IndexOf("<" & StartTagName & ">", NextSearch, EndTagIndex - NextSearch)
    '                                        If SearchIndex = -1 Then
    '                                            Match = False
    '                                        Else
    '                                            NextSearch = SearchIndex + StartTagName.Length
    '                                            StartTagCount += 1
    '                                        End If
    '                                    End While
    '                                    Match = True
    '                                    NextSearch = StartSearch 'Set NextSearch back to StartSearch to search the same chars for <StartTagName ...(with attributes)
    '                                    While Match = True 'Search for Start-Tags of the form: <StartTagName ...> (Start-Tag with attributes)
    '                                        SearchIndex = XmlText.IndexOf("<" & StartTagName & " ", NextSearch, EndTagIndex - NextSearch)
    '                                        If SearchIndex = -1 Then
    '                                            Match = False
    '                                        Else
    '                                            NextSearch = SearchIndex + StartTagName.Length
    '                                            StartTagCount += 1
    '                                        End If
    '                                    End While
    '                                    StartSearch = EndTagIndex + 1 'All Start-Tags named StartTagName have been found to EndTagIndex : Update StartSearch - If more searches are needed, they will start from here.

    '                                    If StartTagCount = EndTagCount Then
    '                                        'Message.Add("TagLevelMatch is True." & vbCrLf)
    '                                        TagLevelMatch = True
    '                                        'Message.Add("Processing Content of <" & StartTagName & "> FromIndex: " & GtCharIndex + 1 & " ToIndex: " & EndTagIndex & vbCrLf)
    '                                        FixedXmlText.Append("<" & StartTagText & ">" & ProcessContent(XmlText, GtCharIndex + 1, EndTagIndex) & "</" & StartTagName & ">" & vbCrLf)
    '                                        ScanIndex = EndTagIndex + StartTagName.Length + 3
    '                                        ElementFound = True
    '                                    Else
    '                                        'Message.Add("TagLevelMatch is False. Search for next matching End-tag." & vbCrLf)
    '                                        SearchEndTagFrom = EndTagIndex + StartTagName.Length + 3
    '                                    End If
    '                                End If
    '                            End While
    '                        End While
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End While
    '    If ElementFound Then
    '        Return vbCrLf & FixedXmlText.ToString
    '    Else
    '        Return FixedXmlText.ToString
    '    End If
    'End Function






    'Private Function ProcessContent(ByRef XmlText As String, FromIndex As Integer, ToIndex As Integer) As String
    '    'Process the XML content in the XmlText string between FromIndex and ToIndex.
    '    'NOTE: THis code was originally written and tested in the ADVL_Import_1 application.
    '    '
    '    'Content alternatives:
    '    'Content only
    '    '<!---->                                        One or more comments
    '    '<Element />                                    One or more empty element tags
    '    '<Element></Element>                            One or more empty elements
    '    '<Element>Content</Element>                     One or more elements containing content
    '    '<Element>                                      One or more elements containing child elements
    '    '  <ChildElement>ChildContent</ChildElement>
    '    '</Element>

    '    Dim StartScan As Integer = FromIndex 'The start of the current content scan
    '    Dim ScanIndex As Integer = FromIndex 'The current scan position
    '    Dim LtCharIndex As Integer 'The index position of the next < character
    '    Dim GtCharIndex As Integer 'The index position of the next > character
    '    Dim FixedXmlText As New System.Text.StringBuilder 'This is used to build the fixed XML text for the content if it contains XML tags
    '    Dim StartTagName As String = "" 'The name of a found Start-tag
    '    Dim EndTagIndex As Integer 'The index of an End-tag

    '    'Message.Add("ProcessContent: FromIndex = " & FromIndex & " ToIndex = " & ToIndex & vbCrLf)

    '    While ScanIndex <= ToIndex
    '        'Find the first pair of < > characters
    '        LtCharIndex = XmlText.IndexOf("<", ScanIndex) 'Find the start of the next Element
    '        If LtCharIndex = -1 Then '< char not found
    '            If ToIndex - ScanIndex = 2 Then
    '                If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                    'Message.Add("At end of line." & vbCrLf)
    '                    FixedXmlText.Append(vbCrLf)
    '                    Exit While
    '                End If
    '            End If
    '            'The characters between FromIndex and ToIndex are Content
    '            'NOTE: StartScan and FromIndex should be the same here: StartScan only advances if the Content contains one or more comments or elements.
    '            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '            FixedXmlText.Append(Content)
    '            'Message.Add("< char not found. Content string returned: " & Content & " ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '            ScanIndex = ToIndex + 1
    '        ElseIf LtCharIndex >= ToIndex Then
    '            If ToIndex - ScanIndex = 2 Then
    '                If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                    'Message.Add("At end of line." & vbCrLf)
    '                    FixedXmlText.Append(vbCrLf)
    '                    Exit While
    '                End If
    '            End If
    '            'The characters between FromIndex and ToIndex are Content
    '            'NOTE: StartScan and FromIndex should be the same here
    '            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '            FixedXmlText.Append(Content)
    '            'Message.Add("< char not found within Content range. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '            ScanIndex = ToIndex + 1
    '        Else
    '            'The < character is within the Content range
    '            'Search for a > character
    '            GtCharIndex = XmlText.IndexOf(">", LtCharIndex + 1)
    '            If GtCharIndex = -1 Then '> char not found
    '                If ToIndex - ScanIndex = 2 Then
    '                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                        'Message.Add("At end of line." & vbCrLf)
    '                        FixedXmlText.Append(vbCrLf)
    '                        Exit While
    '                    End If
    '                End If
    '                'The characters between FromIndex and ToIndex are Content
    '                'NOTE: StartScan and FromIndex should be the same here
    '                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                FixedXmlText.Append(Content)
    '                'Message.Add("> char not found. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                ScanIndex = ToIndex + 1
    '            ElseIf GtCharIndex > ToIndex Then
    '                If ToIndex - ScanIndex = 2 Then
    '                    If XmlText.Substring(ScanIndex, 2) = vbCrLf Then
    '                        'Message.Add("At end of line." & vbCrLf)
    '                        FixedXmlText.Append(vbCrLf)
    '                        Exit While
    '                    End If
    '                End If
    '                'The characters between FromIndex and ToIndex are Content
    '                'NOTE: StartScan and FromIndex should be the same here
    '                Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                FixedXmlText.Append(Content)
    '                'Message.Add("> char not found within Content range. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                ScanIndex = ToIndex + 1
    '            Else
    '                'A start-tag and end-tag pair has been found.
    '                'The <> characters will contain a comment, an element name or be part of the element content.
    '                If XmlText.Substring(LtCharIndex, 4) = "<!--" Then 'This is the start of a comment
    '                    If XmlText.Substring(GtCharIndex - 2, 3) = "-->" Then 'This is the end of a comment
    '                        FixedXmlText.Append(vbCrLf & XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1)) 'Add the Comment to the Fixed XML Text
    '                        ScanIndex = GtCharIndex + 1
    '                        StartScan = GtCharIndex + 1
    '                        'Message.Add("Comment found and returned: " & XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1) & " ScanIndex = " & ScanIndex & vbCrLf)
    '                    Else
    '                        'This is not a comment.
    '                        'The whole content must be the content of a single element
    '                        'The characters between FromIndex and ToIndex are Content
    '                        'NOTE: StartScan and FromIndex should be the same here
    '                        Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                        FixedXmlText.Append(Content)
    '                        'Message.Add("Comment not found. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                        ScanIndex = ToIndex + 1
    '                    End If
    '                Else 'This is a start-tag, empty element or content of a single element
    '                    If XmlText.Chars(GtCharIndex - 1) = "/" Then 'This is an empty element
    '                        FixedXmlText.Append(XmlText.Substring(LtCharIndex, GtCharIndex - LtCharIndex + 1))
    '                        ScanIndex = GtCharIndex + 1
    '                        StartScan = GtCharIndex + 1
    '                        'Message.Add("Empty element found and returned. ScanIndex = " & ScanIndex & vbCrLf)
    '                    Else
    '                        StartTagName = XmlText.Substring(LtCharIndex + 1, GtCharIndex - LtCharIndex - 1) 'This is the name of the Start-tag
    '                        'Find the matching End-tag
    '                        'Message.Add("Searching for End-tag: " & "</" & StartTagName & ">" & vbCrLf)
    '                        EndTagIndex = XmlText.IndexOf("</" & StartTagName & ">", GtCharIndex + 1)
    '                        If EndTagIndex = -1 Then 'There is no matching End-tag
    '                            'This is not an element.
    '                            'The whole content must be the content of a single element
    '                            'The characters between FromIndex and ToIndex are Content
    '                            'NOTE: StartScan and FromIndex should be the same here
    '                            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                            FixedXmlText.Append(Content)
    '                            'Message.Add("Start-tag with no end-tag. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                            ScanIndex = ToIndex + 1
    '                        ElseIf EndTagIndex > ToIndex - StartTagName.Length - 1 Then 'The matching tag is outside of the Content
    '                            'This is not an element.
    '                            'The whole content must be the content of a single element
    '                            'The characters between FromIndex and ToIndex are Content
    '                            'NOTE: StartScan and FromIndex should be the same here
    '                            Dim Content As String = XmlText.Substring(FromIndex, ToIndex - FromIndex).Replace("<", "&lt;").Replace(">", "&gt;")
    '                            FixedXmlText.Append(Content)
    '                            'Message.Add("Start-tag with no end-tag within Content range. Content string returned: " & Content & "  ScanIndex = " & ScanIndex & " ToIndex = " & ToIndex & vbCrLf)
    '                            ScanIndex = ToIndex + 1
    '                        Else 'Matching End-tag found.
    '                            'Message.Add("Processing Content of <" & StartTagName & "> FromIndex: " & GtCharIndex + 1 & " ToIndex: " & EndTagIndex & vbCrLf)
    '                            FixedXmlText.Append(vbCrLf & "<" & StartTagName & ">" & ProcessContent(XmlText, GtCharIndex + 1, EndTagIndex) & "</" & StartTagName & ">")
    '                            ScanIndex = EndTagIndex + StartTagName.Length + 3
    '                            'Message.Add("<" & StartTagName & "> " & "Start-tag and end-tag Found and processed. Content returned. ScanIndex = " & ScanIndex & vbCrLf)
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End While
    '    Return FixedXmlText.ToString
    'End Function





    Public Function XmlToHtml(ByRef xmlText As String, ByVal ShowDecl As Boolean) As String
        'Converts the XML code to HTML code to display the formatted XML on a web page.
        'The HTML code can be pasted into a HTML document to display the formatted XML data on the web page.
        'If ShowDecl is true, show the XMK declaration line.

        'XML Display Settings:
        Dim StartXComment As String = "<Span style=""color:Gray"">"
        Dim EndXComment As String = "</Span>"
        Dim StartXTag As String = "<Span style=""color:Blue"">"
        Dim EndXTag As String = "</Span>"
        Dim StartXElement As String = "<Span style=""color:DarkRed"">"
        Dim EndXElement As String = "</Span>"
        Dim StartXAttribKey As String = "<Span style=""color:Red"">"
        Dim EndXAttribKey As String = "</Span>"
        Dim StartXAttribVal As String = "<Span style=""color:Blue"">"
        Dim EndXAttribVal As String = "</Span>"
        Dim StartXValue As String = "<Span style=""color:Black""><b>"
        Dim EndXValue As String = "</b></Span>"

        Dim xmlDoc As New System.Xml.XmlDocument

        Try
            xmlDoc.LoadXml(xmlText)
        Catch ex As Exception
            RaiseEvent ErrorMessage("Problem converting text to XML: " & vbCrLf & ex.Message & vbCrLf)
        End Try

        Try
            Dim HtmlString As New StringBuilder()

            Dim I As Integer
            Dim J As Integer
            For I = 0 To xmlDoc.ChildNodes().Count - 1
                If xmlDoc.ChildNodes(I).Attributes Is Nothing Then

                End If

                Select Case xmlDoc.ChildNodes(I).NodeType
                    Case Xml.XmlNodeType.XmlDeclaration
                        Dim XDec As System.Xml.XmlDeclaration
                        XDec = xmlDoc.ChildNodes(I)

                        If ShowDecl = True Then
                            HtmlString.Append(StartXTag & "&lt;?" & EndXTag & StartXElement & "xml" & EndXElement & StartXAttribKey & " version" & EndXAttribKey & StartXTag & "=" & EndXTag & StartXAttribVal & """" & XDec.Version & """" & EndXAttribVal)
                            HtmlString.Append(StartXAttribKey & " encoding" & EndXAttribKey & StartXTag & "=" & EndXTag & StartXAttribVal & """" & XDec.Encoding & """" & EndXAttribVal & StartXTag & "?&gt;" & EndXTag & "</br>" & vbCrLf)
                        End If

                    Case Xml.XmlNodeType.Comment
                        HtmlString.Append(StartXTag & "&lt;!--" & EndXTag & StartXComment & xmlDoc.ChildNodes(I).Value & EndXComment & StartXTag & "--&gt;" & EndXTag & "</br>" & vbCrLf)

                    Case Xml.XmlNodeType.Element
                        'This should be the root element!
                        ' Get the Html of the root element.
                        'Dim rootHtmlContent As String = ProcessXmlNodeToHtml(xmlDoc.ChildNodes(I), 0, Settings)
                        Dim rootHtmlContent As String = ProcessXmlNodeToHtml(xmlDoc.ChildNodes(I), 0)
                        HtmlString.Append(rootHtmlContent)

                End Select
            Next
            Return HtmlString.ToString
        Catch ex As Exception
            RaiseEvent ErrorMessage("XmlToHtml error: " & ex.Message & vbCrLf)
        End Try
    End Function

    'Private Function ProcessXmlNodeToHtml(ByVal myNode As System.Xml.XmlNode, ByVal level As Integer, ByRef Settings As XmlHtmDisplaySettings) As String
    Private Function ProcessXmlNodeToHtml(ByVal myNode As System.Xml.XmlNode, ByVal level As Integer) As String
        'This function does not support an XML file that has a Namespace.
        '
        'XML Display Settings:

        'Select the number of indent spaces:
        'Dim XIndent As String = "&nbsp;" '1 space
        'Dim XIndent As String = "&ensp;" '2 spaces
        'Dim XIndent As String = "&ensp; " '3 spaces
        'Dim XIndent As String = "&emsp;" '4 spaces
        Dim XIndent As String = "&emsp; " '5 spaces
        'Dim XIndent As String = "&ensp;&ensp;" '6 spaces
        'Dim XIndent As String = "&ensp;&ensp; " '7 spaces
        'Dim XIndent As String = "&emsp;&emsp;" '8 spaces
        'Dim XIndent As String = "&emsp;&emsp; " '9 spaces
        'Dim XIndent As String = "&emsp; &emsp; " '10 spaces

        Dim StartXComment As String = "<Span style=""color:Gray"">"
        Dim EndXComment As String = "</Span>"
        Dim StartXTag As String = "<Span style=""color:Blue"">"
        Dim EndXTag As String = "</Span>"
        Dim StartXElement As String = "<Span style=""color:DarkRed"">"
        Dim EndXElement As String = "</Span>"
        Dim StartXAttribKey As String = "<Span style=""color:Red"">"
        Dim EndXAttribKey As String = "</Span>"
        Dim StartXAttribVal As String = "<Span style=""color:Blue"">"
        Dim EndXAttribVal As String = "</Span>"
        Dim StartXValue As String = "<Span style=""color:Black""><b>"
        Dim EndXValue As String = "</b></Span>"

        If myNode.NodeType = Xml.XmlNodeType.Element Then
            If Not String.IsNullOrEmpty(myNode.NamespaceURI) Then
                Throw New ApplicationException(
                    "The Xml to Html processor does not support an Xml file that has Namespace.")
            End If
        End If

        Dim elementHtmlFormat As String = String.Empty
        Dim childElementsHtmlContent As New StringBuilder()
        Dim attributesHtmlContent As New StringBuilder()

        'Set the indent spaces.
        Dim indent As String = Replace(Space(level), " ", XIndent)

        If myNode.NodeType = Xml.XmlNodeType.XmlDeclaration Then
            'NOTE: ProcessXmlNodeToHtml should not find any XML Declarations!

        ElseIf myNode.NodeType = Xml.XmlNodeType.Comment Then
            elementHtmlFormat = indent & StartXTag & "&lt;!--" & EndXTag & StartXComment & myNode.Value & EndXComment & StartXTag & "--&gt;" & EndXTag & "</br>" & vbCrLf

        ElseIf myNode.NodeType = Xml.XmlNodeType.CDATA Then
            Dim charIndent As String = Replace(Space((level + 1)), " ", XIndent)
            elementHtmlFormat = indent & StartXTag & "&lt;![CDATA[" & EndXTag & vbCrLf & charIndent & StartXValue & myNode.Value.Trim().Replace("\", "\\") & vbCrLf & indent & StartXTag & "]]&gt;" & EndXTag & "</br>" & vbCrLf

        ElseIf myNode.NodeType = Xml.XmlNodeType.Element Then
            If myNode.HasChildNodes Then 'One or more child nodes. 
                If myNode.ChildNodes.Count = 1 Then
                    If myNode.ChildNodes(0).NodeType = Xml.XmlNodeType.Text Then
                        elementHtmlFormat = indent & StartXTag & "&lt;" & EndXTag & StartXElement & myNode.Name & EndXElement & "{0}" & StartXTag & "&gt;" & EndXTag & "{1}" & StartXTag & "&lt;/" & EndXTag & StartXElement & myNode.Name & EndXElement & StartXTag & "&gt;" & EndXTag & "</br>" & vbCrLf
                        childElementsHtmlContent.Append(StartXValue & myNode.ChildNodes(0).Value.Trim().Replace(vbCrLf, "<br>") & EndXValue)
                    Else 'Non-text child node
                        elementHtmlFormat = indent & StartXTag & "&lt;" & EndXTag & StartXElement & myNode.Name & EndXElement & "{0}" & StartXTag & "&gt;" & EndXTag & "</br>" & vbCrLf &
                        "{1}" & indent & StartXTag & "&lt;/" & EndXTag & StartXElement & myNode.Name & EndXElement & StartXTag & "&gt;" & EndXTag & "</br>" & vbCrLf
                        'Dim childElementHtmlContent As String = ProcessXmlNodeToHtml(myNode.ChildNodes(0), level + 1, Settings)
                        Dim childElementHtmlContent As String = ProcessXmlNodeToHtml(myNode.ChildNodes(0), level + 1)
                        childElementsHtmlContent.Append(childElementHtmlContent)
                    End If
                Else 'Two or more child nodes.
                    elementHtmlFormat = indent & StartXTag & "&lt;" & EndXTag & StartXElement & myNode.Name & EndXElement & "{0}" & StartXTag & "&gt;" & EndXTag & "</br>" & vbCrLf &
                            "{1}" & indent & StartXTag & "&lt;/" & EndXTag & StartXElement & myNode.Name & EndXElement & StartXTag & "&gt;" & EndXTag & "</br>" & vbCrLf
                    For Each Node In myNode.ChildNodes
                        'Dim childElementHtmlContent As String = ProcessXmlNodeToHtml(Node, level + 1, Settings)
                        Dim childElementHtmlContent As String = ProcessXmlNodeToHtml(Node, level + 1)
                        childElementsHtmlContent.Append(childElementHtmlContent)
                    Next
                End If
            Else
                elementHtmlFormat = indent & StartXTag & "&lt;" & EndXTag & StartXElement & myNode.Name & EndXElement & "{0}" & StartXTag & "/&gt;" & EndXTag & "</br>" & vbCrLf
            End If
            'Construct the Html of the attributes.
            If myNode.Attributes.Count > 0 Then
                For Each attribute As Xml.XmlAttribute In myNode.Attributes
                    Dim attributeHtmlContent As String = StartXAttribKey & attribute.Name & EndXAttribKey & StartXTag & "=" & EndXTag & """" & StartXAttribVal & attribute.Value & EndXAttribVal & """"
                    attributesHtmlContent.Append(attributeHtmlContent)
                Next attribute
            End If
        End If
        Return String.Format(elementHtmlFormat, attributesHtmlContent, childElementsHtmlContent)
    End Function


    'NOTE: Now using pre-defined formatting strings defined within the method.
    'Private Function StartXTag() As String
    '    Dim StartBold As String = ""
    '    If Settings.XTag.Bold = True Then StartBold = "<b>"
    '    'Return "<Span style=""color:" & Settings.XTag.Color.ToString & """>" & StartBold
    '    Return "<Span style=""color:" & Settings.XTag.Color.Name & """>" & StartBold
    'End Function
    'Private Function EndXTag() As String
    '    Dim EndBold As String = ""
    '    If Settings.XTag.Bold = True Then EndBold = "</b>"
    '    Return EndBold & "</Span>"
    'End Function

    'Private Function StartXElement() As String
    '    Dim StartBold As String = ""
    '    If Settings.XElement.Bold = True Then StartBold = "<b>"
    '    Return "<Span style=""color:" & Settings.XElement.Color.Name & """>" & StartBold
    'End Function
    'Private Function EndXElement() As String
    '    Dim EndBold As String = ""
    '    If Settings.XElement.Bold = True Then EndBold = "</b>"
    '    Return EndBold & "</Span>"
    'End Function

    'Private Function StartXAttribKey() As String
    '    Dim StartBold As String = ""
    '    If Settings.XAttributeKey.Bold = True Then StartBold = "<b>"
    '    Return "<Span style=""color:" & Settings.XAttributeKey.Color.Name & """>" & StartBold
    'End Function
    'Private Function EndXAttribKey() As String
    '    Dim EndBold As String = ""
    '    If Settings.XAttributeKey.Bold = True Then EndBold = "</b>"
    '    Return EndBold & "</Span>"
    'End Function

    'Private Function StartXAttribVal() As String
    '    Dim StartBold As String = ""
    '    If Settings.XAttributeValue.Bold = True Then StartBold = "<b>"
    '    Return "<Span style=""color:" & Settings.XAttributeValue.Color.Name & """>" & StartBold
    'End Function
    'Private Function EndXAttribVal() As String
    '    Dim EndBold As String = ""
    '    If Settings.XAttributeValue.Bold = True Then EndBold = "</b>"
    '    Return EndBold & "</Span>"
    'End Function

    'Private Function StartXValue() As String
    '    Dim StartBold As String = ""
    '    If Settings.XValue.Bold = True Then StartBold = "<b>"
    '    Return "<Span style=""color:" & Settings.XValue.Color.Name & """>" & StartBold
    'End Function
    'Private Function EndXValue() As String
    '    Dim EndBold As String = ""
    '    If Settings.XValue.Bold = True Then EndBold = "</b>"
    '    Return EndBold & "</Span>"
    'End Function

    'Private Function StartXComment() As String
    '    Dim StartBold As String = ""
    '    If Settings.XComment.Bold = True Then StartBold = "<b>"
    '    Return "<Span style=""color:" & Settings.XComment.Color.Name & """>" & StartBold
    'End Function
    'Private Function EndXComment() As String
    '    Dim EndBold As String = ""
    '    If Settings.XComment.Bold = True Then EndBold = "</b>"
    '    Return EndBold & "</Span>"
    'End Function


    Public Function HmlToRtf(ByRef htmText As String) As String
        'Return the RTF code corresponding to the HTML string.
        'Debug.Print("Starting XmlHtmDisplay.HmlToRtf")

        If htmText Is Nothing Then
            RaiseEvent ErrorMessage("No HTML text to process." & vbCrLf)
            Exit Function
        End If

        If Trim(htmText) = "" Then
            RaiseEvent ErrorMessage("No HTML text to process." & vbCrLf)
            Exit Function
        End If

        Try
            ' The Rtf contains 2 parts, header and content. The colortbl is a part of
            ' the header, and the {1} will be replaced with the content.

            'HTML Display Settings: -------------------------------------------
            'HText          HTML Text                           <Text>
            'HElement       HTML Elements                       <Elem>
            'HAttribute     HTML Attributes                     <Attr>
            'HComment       HTML Comments                       <Comm>
            'HChar          HTML Special Characters             <Char>
            'HValue         HTML Attribute or Styles Values     <Valu>
            'HStyle         HTML Style Attribute                <Styl>

            Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                                      & "{{\fonttbl{{{0}}}}}" _
                                      & "{{\colortbl ;{1}}}" _
                                      & "\viewkind4\uc1\pard\lang1033" _
                                      & Settings.RtfTextSettings(Settings.HChar) _
                                      & "{2}}}"

            'NOTE: TO USE Settings.RtfHeader in a Format statement, "{" characters should be replaced by "{{" characters.

            'Code to escape \ characters:
            htmText = htmText.Replace("\", "\\")

            Dim htmRtfContent As New StringBuilder()
            'RaiseEvent Message("Starting HmlToRtf()" & vbCrLf) 
            Dim I As Integer = 0
            Dim ProcessingComment As Boolean = False
            Dim ProcessingTag As Boolean = False
            Dim ProcessingAttributeValue As Boolean = False

            Dim ProcessingScript As Boolean = False 'NEW for processing JavaScript code.
            Dim ProcessingJSComment As Boolean = False 'If True, processing JavaScript comment.

            While I < htmText.Length
                Dim Char1 = htmText(I)
                Dim Char2 As Char
                If htmText.Length > I + 1 Then
                    Char2 = htmText(I + 1)
                Else
                    Char2 = Chr(0)
                End If
                'RaiseEvent Message("I = " & I & ", Char1 = '" & Char1 & "', Char2 = '" & Char2 & "'" & vbCrLf) 



                If Not ProcessingComment AndAlso Char1 = "<" Then
                    'Look for <!-- (comments tag)
                    'RaiseEvent Message("STATE: NotProcComm And Char1 = '<'" & vbCrLf) 
                    If htmText.Length > I + 3 AndAlso Char2 = "!" AndAlso htmText(I + 2) = "-" AndAlso htmText(I + 3) = "-" Then 'Comments tag
                        'RaiseEvent Message("STATE: Chars = '<!-'" & vbCrLf) 
                        'RaiseEvent Message("<Comm>, Append Char1: '" & Char1 & "'" & vbCrLf) 
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HComment)).Append(" ").Append(Char1)
                        ProcessingComment = True

                        '<Script> Code ====================================================

                        'Look for </script>
                    ElseIf ProcessingScript AndAlso Char1 = "<" AndAlso htmText.Length > I + 8 AndAlso Char2 = "/" AndAlso LCase(htmText(I + 2)) = "s" AndAlso htmText(I + 3) = "c" AndAlso htmText(I + 4) = "r" AndAlso htmText(I + 5) = "i" AndAlso htmText(I + 6) = "p" AndAlso htmText(I + 7) = "t" AndAlso htmText(I + 8) = ">" Then
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1) 'Append "<" with Special Char font.
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HElement)).Append(" ")
                        htmRtfContent.Append("/script")                                                           'Append "script" with Element font.
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(">")   'Append ">" with Special Char font.
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")               'Set font to Text.
                        ProcessingScript = False                                                                 'Set JavaScript ProcessingScript flag to True.
                        I += 8
                        'RaiseEvent Message("ProcessingScript = False" & vbCrLf)
                        '</Script> Code ---------------------------------------------------

                        'ElseIf ProcessingScript AndAlso Not ProcessingJSComment AndAlso Char1 = "/" AndAlso Char2 = "/" Then 'JavaScript comments tag
                        '    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HComment)).Append(" ").Append(Char1)
                        '    ProcessingJSComment = True

                        'ElseIf ProcessingScript AndAlso ProcessingJSComment AndAlso Char1 = vbLf Then 'End of JavaScript comments tag
                        '    htmRtfContent.Append(Char1).Append("\par ").Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                        '    ProcessingJSComment = False


                        '<Script> Code ====================================================

                        'Look for <script>
                    ElseIf htmText.Length > I + 7 AndAlso LCase(Char2) = "s" AndAlso htmText(I + 2) = "c" AndAlso htmText(I + 3) = "r" AndAlso htmText(I + 4) = "i" AndAlso htmText(I + 5) = "p" AndAlso htmText(I + 6) = "t" AndAlso htmText(I + 7) = ">" Then
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1) 'Append "<" with Special Char font.
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HElement)).Append(" ")
                        htmRtfContent.Append("script")                                                           'Append "script" with Element font.
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(">")   'Append ">" with Special Char font.
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")               'Set font to Text.
                        ProcessingScript = True                                                                  'Set JavaScript ProcessingScript flag to True.
                        I += 7                                                                                   'Skip the next 8 characters (Past script>)
                        'RaiseEvent Message("ProcessingScript = True" & vbCrLf)
                        '</Script> Code ---------------------------------------------------
                    Else
                        'RaiseEvent Message("STATE: Chars <> '<!-'" & vbCrLf) 
                        'RaiseEvent Message("<Char>, Append Char1: '" & Char1 & "'" & vbCrLf) 
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1)
                        If Char2 = "/" Then
                            'RaiseEvent Message("STATE: Chars = '</'" & vbCrLf) 
                            'RaiseEvent Message("Append Char2 = '" & Char2 & "'" & vbCrLf) 
                            htmRtfContent.Append(Char2)
                            I += 1
                        End If
                        'RaiseEvent Message("<Elem>, ProcTag" & vbCrLf) 
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HElement)).Append(" ")
                        ProcessingTag = True
                    End If

                    '<Script> Code ====================================================
                ElseIf ProcessingScript AndAlso Not ProcessingJSComment AndAlso Char1 = "/" AndAlso Char2 = "/" Then 'JavaScript comments tag 
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HComment)).Append(" ").Append(Char1)
                    ProcessingJSComment = True

                ElseIf ProcessingScript AndAlso ProcessingJSComment AndAlso Char1 = vbLf Then 'End of JavaScript comments tag
                    htmRtfContent.Append(Char1).Append("\par ").Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                    ProcessingJSComment = False

                ElseIf ProcessingScript AndAlso Not ProcessingJSComment AndAlso Char1 = vbLf Then
                    htmRtfContent.Append(Char1).Append("\par ").Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                    '</Script> Code ---------------------------------------------------

                ElseIf Char1 = ">" Then 'Look for Comments tags:
                    'RaiseEvent Message("STATE: Char1 = '>'" & vbCrLf) 
                    If ProcessingComment AndAlso htmText(I - 1) = "-" AndAlso htmText(I - 2) = "-" Then
                        'RaiseEvent Message("STATE: ProcComm And Chars = '-->'" & vbCrLf) 
                        'RaiseEvent Message("Append Char1: '" & Char1 & "' <Text>, NotProcComm" & vbCrLf) 
                        htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                        ProcessingComment = False
                    ElseIf Not ProcessingComment Then
                        'RaiseEvent Message("STATE: NotProcComm" & vbCrLf) 
                        'RaiseEvent Message("<Char> Char1: '" & Char1 & "' <Text>, NotProcTag, NotProcAttrVal" & vbCrLf) 
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                        ProcessingTag = False
                        ProcessingAttributeValue = False
                    End If
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Char1 = "/" AndAlso Char2 = ">" Then
                    'RaiseEvent Message("STATE: ProcTag And NotProcComm And Chars = '/>'" & vbCrLf) 
                    'RaiseEvent Message("<Char> Append Char1 and Char2: '" & Char1 & Char2 & "', <Text>, NotProcTag" & vbCrLf) 
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Char2).Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                    ProcessingTag = False
                    I += 1
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Not ProcessingAttributeValue AndAlso Char1 = " " Then
                    'RaiseEvent Message("STATE: ProcTag And NotProcComment And NotProcValu And Char1 = ' '" & vbCrLf) 
                    'RaiseEvent Message("Append Char1: '" & Char1 & "', <Attr>" & vbCrLf)
                    htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HAttribute)).Append(" ")


                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Not ProcessingAttributeValue AndAlso Char1 = vbLf Then
                    'RaiseEvent Message("STATE: ProcTag And NotProcComment And NotProcValu And Char1 = 'vbLf'" & vbCrLf) 
                    'RaiseEvent Message("Append Char1: '" & Char1 & "', <Attr>" & vbCrLf)
                    htmRtfContent.Append(Char1).Append("\par ").Append(Settings.RtfTextSettings(Settings.HAttribute)).Append(" ")

                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Char1 = "=" Then
                    'RaiseEvent Message("STATE: ProcTag And NotProcComment And Char1 = '='" & vbCrLf) 
                    'RaiseEvent Message("<Char> Append Char1: '" & Char1 & "', <Style>" & vbCrLf)
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle)).Append(" ")
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso ProcessingAttributeValue AndAlso Char1 = ":" Then
                    'RaiseEvent Message("STATE: ProcTag And NotProcComment And ProcValu And Char1 = ':'" & vbCrLf) 
                    'RaiseEvent Message("<Text> Append Char1: '" & Char1 & "', <Valu>" & vbCrLf)
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HText)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HValue)).Append(" ")
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso ProcessingAttributeValue AndAlso Char1 = ";" Then
                    'RaiseEvent Message("STATE: ProcTag And NotProcComm And ProcValu And Char1 = ';'" & vbCrLf) 
                    'RaiseEvent Message("<Text> Append Char1: '" & Char1 & "', <Styl>" & vbCrLf)
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HText)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle)).Append(" ")
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso (Char1 = """" OrElse Char1 = "'") Then
                    'RaiseEvent Message("STATE: ProcTag And NotProcComm And Char1 = '""'" & vbCrLf) 
                    'RaiseEvent Message("<Char> Append Char1: '" & Char1 & "', <Styl> Not(ProcValu)" & vbCrLf)
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle)).Append(" ")
                    ProcessingAttributeValue = Not ProcessingAttributeValue

                    '<Script> Code ====================================================
                    'ElseIf ProcessingScript AndAlso (Char1 = """" OrElse Char1 = "'") Then
                ElseIf ProcessingScript AndAlso Not ProcessingJSComment AndAlso (Char1 = """" OrElse Char1 = "'") Then
                    If ProcessingAttributeValue Then
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                    Else
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle)).Append(" ")
                    End If
                    ProcessingAttributeValue = Not ProcessingAttributeValue
                    '</Script> Code ---------------------------------------------------

                ElseIf Char1 = vbLf Then
                    'RaiseEvent Message("STATE: Char1 = 'vbLf'" & vbCrLf) 
                    'RaiseEvent Message("Append Char1: '" & Char1 & "' \par" & vbCrLf)
                    htmRtfContent.Append(Char1).Append("\par ")
                ElseIf (Char1 = "{" OrElse Char1 = "}") Then
                    'RaiseEvent Message("STATE: Char1 = '{' OrElse '}'" & vbCrLf) 
                    'RaiseEvent Message("Append '\' Append Char1: '" & Char1 & "'" & vbCrLf)
                    htmRtfContent.Append("\").Append(Char1)
                Else
                    htmRtfContent.Append(Char1)
                End If
                I += 1

            End While

            Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), htmRtfContent.ToString)

        Catch ex As Exception
            RaiseEvent ErrorMessage("HmlToRtf error: " & ex.Message & vbCrLf)
        End Try

    End Function

    Public Function HmlToRtfDebug(ByRef htmText As String) As String
        'Return the RTF code corresponding to the HTML string.
        'Debug.Print("Starting XmlHtmDisplay.HmlToRtf")

        If htmText Is Nothing Then
            RaiseEvent ErrorMessage("No HTML text to process." & vbCrLf)
            Exit Function
        End If

        If Trim(htmText) = "" Then
            RaiseEvent ErrorMessage("No HTML text to process." & vbCrLf)
            Exit Function
        End If

        Try
            ' The Rtf contains 2 parts, header and content. The colortbl is a part of
            ' the header, and the {1} will be replaced with the content.

            'HTML Display Settings: -------------------------------------------
            'HText          HTML Text                           <Text>
            'HElement       HTML Elements                       <Elem>
            'HAttribute     HTML Attributes                     <Attr>
            'HComment       HTML Comments                       <Comm>
            'HChar          HTML Special Characters             <Char>
            'HValue         HTML Attribute or Styles Values     <Valu>
            'HStyle         HTML Style Attribute                <Styl>

            Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                                      & "{{\fonttbl{{{0}}}}}" _
                                      & "{{\colortbl ;{1}}}" _
                                      & "\viewkind4\uc1\pard\lang1033" _
                                      & Settings.RtfTextSettings(Settings.HChar) _
                                      & "{2}}}"

            'NOTE: TO USE Settings.RtfHeader in a Format statement, "{" characters should be replaced by "{{" characters.

            Dim htmRtfContent As New StringBuilder()
            RaiseEvent Message("Starting HmlToRtf()" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Dim I As Integer = 0
            Dim ProcessingComment As Boolean = False
            Dim ProcessingTag As Boolean = False
            Dim ProcessingAttributeValue As Boolean = False

            While I < htmText.Length
                Dim Char1 = htmText(I)
                Dim Char2 As Char
                If htmText.Length > I + 1 Then
                    Char2 = htmText(I + 1)
                Else
                    Char2 = Chr(0)
                End If
                RaiseEvent Message("I = " & I & ", Char1 = '" & Char1 & "', Char2 = '" & Char2 & "'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>

                If Not ProcessingComment AndAlso Char1 = "<" Then
                    RaiseEvent Message("STATE: NotProcComm And Char1 = '<'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    If htmText.Length > I + 3 AndAlso Char2 = "!" AndAlso htmText(I + 3) = "-" Then 'Comments tag
                        RaiseEvent Message("STATE: Chars = '<!-'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        RaiseEvent Message("<Comm>, Append Char1: '" & Char1 & "'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        'htmRtfContent.Append(Settings.RtfTextSettings(Settings.HComment)).Append(Char1)
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HComment)).Append(" ").Append(Char1)
                        ProcessingComment = True
                    Else
                        RaiseEvent Message("STATE: Chars <> '<!-'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        RaiseEvent Message("<Char>, Append Char1: '" & Char1 & "'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        'htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(Char1)
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1)
                        If Char2 = "/" Then
                            RaiseEvent Message("STATE: Chars = '</'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                            RaiseEvent Message("Append Char2 = '" & Char2 & "'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                            htmRtfContent.Append(Char2)
                            I += 1
                        End If
                        RaiseEvent Message("<Elem>, ProcTag" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        'htmRtfContent.Append(Settings.RtfTextSettings(Settings.HElement))
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HElement)).Append(" ")
                        ProcessingTag = True
                    End If
                ElseIf Char1 = ">" Then 'Look for Comments tags:
                    RaiseEvent Message("STATE: Char1 = '>'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    If ProcessingComment AndAlso htmText(I - 1) = "-" AndAlso htmText(I - 2) = "-" Then
                        RaiseEvent Message("STATE: ProcComm And Chars = '-->'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        RaiseEvent Message("Append Char1: '" & Char1 & "' <Text>, NotProcComm" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        'htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HText))
                        htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                        ProcessingComment = False
                    ElseIf Not ProcessingComment Then
                        RaiseEvent Message("STATE: NotProcComm" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        RaiseEvent Message("<Char> Char1: '" & Char1 & "' <Text>, NotProcTag, NotProcAttrVal" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                        'htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(Char1).Append(Settings.RtfTextSettings(Settings.HText))
                        htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                        ProcessingTag = False
                        ProcessingAttributeValue = False
                    End If
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Char1 = "/" AndAlso Char2 = ">" Then
                    RaiseEvent Message("STATE: ProcTag And NotProcComm And Chars = '/>'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("<Char> Append Char1 and Char2: '" & Char1 & Char2 & "', <Text>, NotProcTag" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Char2).Append(Settings.RtfTextSettings(Settings.HText)).Append(" ")
                    ProcessingTag = False
                    I += 1
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Not ProcessingAttributeValue AndAlso Char1 = " " Then
                    RaiseEvent Message("STATE: ProcTag And NotProcComment And NotProcValu And Char1 = ' '" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("Append Char1: '" & Char1 & "', <Attr>" & vbCrLf)
                    'htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HAttribute))
                    htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HAttribute)).Append(" ")


                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Not ProcessingAttributeValue AndAlso Char1 = vbLf Then
                    RaiseEvent Message("STATE: ProcTag And NotProcComment And NotProcValu And Char1 = 'vbLf'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("Append Char1: '" & Char1 & "', <Attr>" & vbCrLf)
                    'htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HAttribute))
                    'htmRtfContent.Append(Char1).Append(Settings.RtfTextSettings(Settings.HAttribute)).Append(" ")
                    htmRtfContent.Append(Char1).Append("\par ").Append(Settings.RtfTextSettings(Settings.HAttribute)).Append(" ")


                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso Char1 = "=" Then
                    RaiseEvent Message("STATE: ProcTag And NotProcComment And Char1 = '='" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("<Char> Append Char1: '" & Char1 & "', <Style>" & vbCrLf)
                    'htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle))
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle)).Append(" ")
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso ProcessingAttributeValue AndAlso Char1 = ":" Then
                    RaiseEvent Message("STATE: ProcTag And NotProcComment And ProcValu And Char1 = ':'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("<Text> Append Char1: '" & Char1 & "', <Valu>" & vbCrLf)
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HText)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HValue)).Append(" ")
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso ProcessingAttributeValue AndAlso Char1 = ";" Then
                    RaiseEvent Message("STATE: ProcTag And NotProcComm And ProcValu And Char1 = ';'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("<Text> Append Char1: '" & Char1 & "', <Styl>" & vbCrLf)
                    'htmRtfContent.Append(Settings.RtfTextSettings(Settings.HText)).Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle))
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HText)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle)).Append(" ")
                ElseIf ProcessingTag AndAlso Not ProcessingComment AndAlso (Char1 = """" OrElse Char1 = "'") Then
                    RaiseEvent Message("STATE: ProcTag And NotProcComm And Char1 = '""'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("<Char> Append Char1: '" & Char1 & "', <Styl> Not(ProcValu)" & vbCrLf)
                    'htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle))
                    htmRtfContent.Append(Settings.RtfTextSettings(Settings.HChar)).Append(" ").Append(Char1).Append(Settings.RtfTextSettings(Settings.HStyle)).Append(" ")
                    ProcessingAttributeValue = Not ProcessingAttributeValue
                ElseIf Char1 = vbLf Then
                    RaiseEvent Message("STATE: Char1 = 'vbLf'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("Append Char1: '" & Char1 & "' \par" & vbCrLf)
                    'htmRtfContent.Append(Char1).Append("\par")
                    htmRtfContent.Append(Char1).Append("\par ")
                ElseIf (Char1 = "{" OrElse Char1 = "}") Then
                    RaiseEvent Message("STATE: Char1 = '{' OrElse '}'" & vbCrLf) 'FOR DEBUGGING>>>>>>>>>>>>>>>>>>>>>>>
                    RaiseEvent Message("Append '\' Append Char1: '" & Char1 & "'" & vbCrLf)
                    htmRtfContent.Append("\").Append(Char1)
                    '_UnicodeTest(0) = Char1
                Else
                    htmRtfContent.Append(Char1)
                End If
                I += 1

            End While

            Return String.Format(rtfFormat, Settings.RtfFontTableFormatString(), Settings.RtfColorTableFormatString(), htmRtfContent.ToString)

        Catch ex As Exception
            RaiseEvent ErrorMessage("HmlToRtf error: " & ex.Message & vbCrLf)
        End Try

    End Function

    Public Function ReadTextFile(ByVal FilePath As String) As Boolean
        'Read the text data at the specified file path.
        'Returns True if the file was read OK.
        'NOTE: Project.ReadData can be used instead of this! (ReadData into a MemoryStream, then use the MemoryStream.)

        Try
            If System.IO.File.Exists(FilePath) Then
                Me.LoadFile(FilePath, RichTextBoxStreamType.PlainText)
                Return True
            Else
                RaiseEvent ErrorMessage("Text file does not exist: " & FilePath & vbCrLf)
                Return False
            End If
        Catch ex As Exception
            RaiseEvent ErrorMessage("Error reading text file: " & FilePath & " " & vbCrLf & ex.Message & vbCrLf)
            Return False
        End Try

    End Function

    Public Function ReadPlainTextFile(ByVal FilePath As String) As Boolean
        'Read the text data at the specified file path.
        'Display the text using the Plain Text settings.
        'Returns True if the file was read OK.

        Try
            If System.IO.File.Exists(FilePath) Then
                Me.Font = PlainTextFont()
                Me.LoadFile(FilePath, RichTextBoxStreamType.PlainText)
                Return True
            Else
                RaiseEvent ErrorMessage("Text file does not exist: " & FilePath & vbCrLf)
                Return False
            End If
        Catch ex As Exception
            RaiseEvent ErrorMessage("Error reading text file: " & FilePath & " " & vbCrLf & ex.Message & vbCrLf)
            Return False
        End Try

    End Function

    Public Function SaveTextFile(ByVal FilePath As String) As Boolean
        'Save the text data to the specified file path.
        'Returns True if the file was saved OK.
        'NOTE: Project.SaveData can be used instead of this! (Save the data to a MemoryStream, then use .SaveData to save the MemoryStream.)

        Try
            Me.SaveFile(FilePath, RichTextBoxStreamType.PlainText)
            Return True
        Catch ex As Exception
            RaiseEvent ErrorMessage("Error saving text file: " & FilePath & " " & vbCrLf & ex.Message & vbCrLf)
            Return False
        End Try


    End Function

    Public Sub DefaultXmlSettings()
        'Set the XML display settings to default.

        'XTag
        Settings.XTag.FontName = "Arial"
        Settings.XTag.Color = Color.Blue
        Settings.XTag.HalfPointSize = 20
        Settings.XTag.Bold = False
        Settings.XTag.Italic = False

        'XValue
        Settings.XValue.FontName = "Arial"
        Settings.XValue.Color = Color.Black
        Settings.XValue.HalfPointSize = 20
        Settings.XValue.Bold = True
        Settings.XValue.Italic = False

        'XComment
        Settings.XComment.FontName = "Arial"
        Settings.XComment.Color = Color.Gray
        Settings.XComment.HalfPointSize = 20
        Settings.XComment.Bold = False
        Settings.XComment.Italic = False

        'XElement
        Settings.XElement.FontName = "Arial"
        Settings.XElement.Color = Color.DarkRed
        Settings.XElement.HalfPointSize = 20
        Settings.XElement.Bold = False
        Settings.XElement.Italic = False

        'XAttributeKey
        Settings.XAttributeKey.FontName = "Arial"
        Settings.XAttributeKey.Color = Color.Red
        Settings.XAttributeKey.HalfPointSize = 20
        Settings.XAttributeKey.Bold = False
        Settings.XAttributeKey.Italic = False

        'XAttributeValue
        Settings.XAttributeValue.FontName = "Arial"
        Settings.XAttributeValue.Color = Color.Blue
        Settings.XAttributeValue.HalfPointSize = 20
        Settings.XAttributeValue.Bold = False
        Settings.XAttributeValue.Italic = False

        Settings.UpdateColorIndexes()
        Settings.UpdateFontIndexes()

    End Sub

    Public Sub DefaultHtmlSettings()
        'Set the HTML display settings to default.

        'HText
        Settings.HText.FontName = "Arial"
        Settings.HText.Color = Color.Black
        Settings.HText.HalfPointSize = 20
        Settings.HText.Bold = False
        Settings.HText.Italic = False

        'HElement
        Settings.HElement.FontName = "Arial"
        Settings.HElement.Color = Color.DarkRed
        Settings.HElement.HalfPointSize = 20
        Settings.HElement.Bold = False
        Settings.HElement.Italic = False

        'HAttribute
        Settings.HAttribute.FontName = "Arial"
        Settings.HAttribute.Color = Color.Red
        Settings.HAttribute.HalfPointSize = 20
        Settings.HAttribute.Bold = False
        Settings.HAttribute.Italic = False

        'HComment
        Settings.HComment.FontName = "Arial"
        Settings.HComment.Color = Color.Green
        Settings.HComment.HalfPointSize = 20
        Settings.HComment.Bold = False
        Settings.HComment.Italic = False

        'HChar
        Settings.HChar.FontName = "Arial"
        Settings.HChar.Color = Color.Blue
        Settings.HChar.HalfPointSize = 20
        Settings.HChar.Bold = False
        Settings.HChar.Italic = False

        'HValue
        Settings.HValue.FontName = "Arial"
        Settings.HValue.Color = Color.Blue
        Settings.HValue.HalfPointSize = 20
        Settings.HValue.Bold = False
        Settings.HValue.Italic = False

        'HStyle
        Settings.HStyle.FontName = "Arial"
        Settings.HStyle.Color = Color.Purple
        Settings.HStyle.HalfPointSize = 20
        Settings.HStyle.Bold = False
        Settings.HStyle.Italic = False

        Settings.UpdateColorIndexes()
        Settings.UpdateFontIndexes()

    End Sub

    Public Sub DefaultPlainTextSettings()
        'Set the Plain Text display settings to default.

        'PlainText
        Settings.PlainText.FontName = "Arial"
        Settings.PlainText.Color = Color.Black
        'Settings.PlainText.HalfPointSize = 20
        Settings.PlainText.HalfPointSize = 24
        Settings.PlainText.Bold = False
        Settings.PlainText.Italic = False

        Settings.UpdateColorIndexes()
        Settings.UpdateFontIndexes()

        'Me.DefaultFont = PlainTextFont()
        Me.Font = PlainTextFont()
    End Sub



    'Public Sub ShowTextTypes()
    '    'Show the text types stored in the TextType Dictionary:
    'THIS IS DONE IN THE SYSTEM_UTILITES MESSAGE CLASS

    'End Sub

    Private Sub Settings_ErrorMessage(Message As String) Handles Settings.ErrorMessage
        RaiseEvent ErrorMessage(Message)
    End Sub

    Event Message(ByVal Msg As String)
    Event ErrorMessage(ByVal Msg As String)

End Class


