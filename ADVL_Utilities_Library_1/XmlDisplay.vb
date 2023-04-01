Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms

Public Class XmlDisplay
    'Rich Text Box control with XML document formatting.

    Inherits RichTextBox

    Public WithEvents Settings As New XmlDisplaySettings

    Public Sub ReadXmlFile(ByVal FilePath As String, ByVal ShowDecl As Boolean)
        'Read and display the XML data in the file.

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

        Try
            If Settings.TextType.ContainsKey(TypeName) Then
                Dim rtfFormat As String = "{{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052" _
                          & "{{\fonttbl{{{0}}}}}" _
                          & "{{\colortbl ;{1}}}" _
                          & "\viewkind4\uc1\pard\lang1033" _
                          & Settings.RtfTextSettings(Settings.Tag) _
                          & "{2}}}"

                RaiseEvent Message("rtfFormat = " & rtfFormat & vbCrLf & vbCrLf)

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

    Public Function XmlToRtf(ByRef xmlDoc As System.Xml.XmlDocument, ByVal ShowDecl As Boolean) As String
        'Return the RTF code corresponding to the XML.
        'If ShowDecl is true, show the XMK declaration line.

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
                                      & Settings.RtfTextSettings(Settings.Tag) _
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
                           Settings.RtfTextSettings(Settings.Tag),
                           Settings.RtfTextSettings(Settings.Element),
                           Settings.RtfTextSettings(Settings.AttributeKey),
                           Settings.RtfTextSettings(Settings.AttributeValue),
                           XDec.Version,
                           XDec.Encoding)
                        End If

                    Case Xml.XmlNodeType.Comment
                        xmlRtfContent.AppendFormat(
                        "{0}{1} <!--{2} {3}{1} -->\par",
                        "",
                        Settings.RtfTextSettings(Settings.Tag),
                        Settings.RtfTextSettings(Settings.Comment),
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


        Dim xmlDoc As New System.Xml.XmlDocument

        Try
            xmlDoc.LoadXml(xmlText)
        Catch ex As Exception
            RaiseEvent ErrorMessage("Problem converting text to XML: " & vbCrLf & ex.Message & vbCrLf)
        End Try


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
                                      & Settings.RtfTextSettings(Settings.Tag) _
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
                           Settings.RtfTextSettings(Settings.Tag),
                           Settings.RtfTextSettings(Settings.Element),
                           Settings.RtfTextSettings(Settings.AttributeKey),
                           Settings.RtfTextSettings(Settings.AttributeValue),
                           XDec.Version,
                           XDec.Encoding)
                        End If

                    Case Xml.XmlNodeType.Comment
                        xmlRtfContent.AppendFormat(
                        "{0}{1} <!--{2} {3}{1} -->\par",
                        "",
                        Settings.RtfTextSettings(Settings.Tag),
                        Settings.RtfTextSettings(Settings.Comment),
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

    Private Function ProcessNode(ByVal myNode As System.Xml.XmlNode, ByVal level As Integer, ByRef Settings As XmlDisplaySettings) As String
        ' This viewer does not support the Xml file that has a Namespace.
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
        Dim indent As New String(" "c, Settings.IndentSpaces * level)

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
                          Settings.RtfTextSettings(Settings.Tag),
                          Settings.RtfTextSettings(Settings.Comment),
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
                          Settings.RtfTextSettings(Settings.Tag),
                          Settings.RtfTextSettings(Settings.Element),
                          myNode.Name)

                        childElementsRtfContent.AppendFormat(
                          "{0} {1}",
                          Settings.RtfTextSettings(Settings.Value),
                          myNode.ChildNodes(0).Value.Trim().Replace("\", "\\").Replace(vbCrLf, "\line "))
                        'NOTE: "\\" is used to display a "\" in rich text format. "\line " is used for a new line in rich text format.
                        'The Replace functions in the last line apply these changes to produce the correct rich text format.

                        'myNode.ChildNodes(0).Value.Trim())
                    Else 'Non-text child node.

                        'TESTING CODE TO HANDLE A NON-TEXT NODE: =======================================================================================
                        elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >\par" & "{{1}}" & "{0}{1} </{2} {3}{1} >\par",
                              indent,
                              Settings.RtfTextSettings(Settings.Tag),
                              Settings.RtfTextSettings(Settings.Element),
                              myNode.Name)
                        'Dim childElementRtfContent As String = ProcessNode(myNode.ChildNodes(0), level + 1, Settings)
                        Dim childElementRtfContent As String = ProcessNode(myNode.ChildNodes(0), level + 1, Settings)
                        childElementsRtfContent.Append(childElementRtfContent)
                        'END OF TEST CODE ==============================================================================================================

                    End If '-----------------------------------------------------------------------------------------------------------
                Else 'Two or more child nodes.
                    elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} >\par" & "{{1}}" & "{0}{1} </{2} {3}{1} >\par",
                          indent,
                          Settings.RtfTextSettings(Settings.Tag),
                          Settings.RtfTextSettings(Settings.Element),
                          myNode.Name)
                    For Each Node In myNode.ChildNodes
                        Dim childElementRtfContent As String = ProcessNode(Node, level + 1, Settings)
                        childElementsRtfContent.Append(childElementRtfContent)
                    Next

                End If '------------------------------------------------------------------------------------------------------------
            Else
                elementRtfFormat = String.Format("{0}{1} <{2} {3}{{0}}{1} />\par",
                                     indent,
                                     Settings.RtfTextSettings(Settings.Tag),
                                     Settings.RtfTextSettings(Settings.Element),
                                     myNode.Name)
            End If '-------------------------------------------------------------------------------------------------------------
            ' Construct the Rtf of the attributes.
            If myNode.Attributes.Count > 0 Then
                For Each attribute As XAttribute In myNode.Attributes
                    Dim attributeRtfContent As String =
                        String.Format(" {0} {3}{1} =\cf0 ""{2} {4}\cf0 """,
                                      Settings.RtfTextSettings(Settings.AttributeKey),
                                      Settings.RtfTextSettings(Settings.Tag),
                                      Settings.RtfTextSettings(Settings.AttributeValue),
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

    Public Function XmlToHtml(ByRef xmlText As String, ByVal ShowDecl As Boolean) As String
        'Converts the XML code to HTML code to display the formatted XML on a web page.
        'The HTML code can be pasted into a HTML document to display the formatted XML data on the web page.
        'If ShowDecl is true, show the XML declaration line.

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
                        Dim childElementHtmlContent As String = ProcessXmlNodeToHtml(myNode.ChildNodes(0), level + 1)
                        childElementsHtmlContent.Append(childElementHtmlContent)
                    End If
                Else 'Two or more child nodes.
                    elementHtmlFormat = indent & StartXTag & "&lt;" & EndXTag & StartXElement & myNode.Name & EndXElement & "{0}" & StartXTag & "&gt;" & EndXTag & "</br>" & vbCrLf &
                            "{1}" & indent & StartXTag & "&lt;/" & EndXTag & StartXElement & myNode.Name & EndXElement & StartXTag & "&gt;" & EndXTag & "</br>" & vbCrLf
                    For Each Node In myNode.ChildNodes
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

    Private Sub Settings_ErrorMessage(Message As String) Handles Settings.ErrorMessage
        RaiseEvent ErrorMessage(Message)
    End Sub

    Event Message(ByVal Msg As String)
    Event ErrorMessage(ByVal Msg As String)

End Class
