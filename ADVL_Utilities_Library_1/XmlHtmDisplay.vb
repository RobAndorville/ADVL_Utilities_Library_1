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
            'RaiseEvent Message("xmlDoc is nothing" & vbCrLf)
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


