Imports System.Text
Imports System.Drawing

Public Class XmlDisplaySettings
    'Settings used to display formatted XML data and specified RTF text types.

    Public NFonts As Integer 'The number of unique fonts used to display the XML data and other text types.
    Public NColors As Integer 'The number of unique colors used to display the XML data and other text types.
    Public TextType As New Dictionary(Of String, TextSettings) 'Dictionary of other text settings.

    'NOTE: The followng code sectiosn were attemps to automatically detect Font Name and Color changes so that UpdateFontIndexes and UpdateColorIndexes can be applied automatically.
    '  So far no technique has been found to detect changes made to the TextType dictionary.
    'Public WithEvents TextType As New Dictionary(Of String, TextSettings) 'Dictionary of other text settings. 'THIS HAS NO EVENTS.
    'Private _textType As New Dictionary(Of String, TextSettings) 'Dictionary of other text settings.
    'Public SettingsChanged As Boolean = False 'True if settings have been changed and Font and Color indexes need to be updated.
    'Private FontSettingsChanged As Boolean = False
    'Private ColorSettingsChanged As Boolean = False

    'Private _showDeclaration As Boolean = True 'If True, the XML declaration line is shown.
    '''' <summary>
    '''' If True, the XML declaration is shown.
    '''' </summary>
    'Public Property ShowDeclaration As Boolean
    '    Get
    '        Return _showDeclaration
    '    End Get
    '    Set(value As Boolean)
    '        _showDeclaration = value
    '    End Set
    'End Property

    'List of Properties:
    'XML Display Settings: ---------------------------------------------
    'XmlLargeFileSizeLimit
    'IndentSpaces
    'Tag
    'Element
    'AttributeKey
    'AttributeValue
    'Value
    'Comment

    Private _xmlLargeFileSizeLimit As Integer = 1000000
    ''' <summary>
    ''' The arbitrary upper file size that can be displayed quickly using the XmlDisplay. 
    ''' </summary>
    Property XmlLargeFileSizeLimit As Integer
        Get
            Return _xmlLargeFileSizeLimit
        End Get
        Set(value As Integer)
            _xmlLargeFileSizeLimit = value
        End Set
    End Property



    Private _indentSpaces As Integer = 4 'The number of spaces used to indent the XML code.
    ''' <summary>
    ''' The number of spaces used to indent the XML code.
    ''' </summary>
    Public Property IndentSpaces As Integer
        Get
            Return _indentSpaces
        End Get
        Set(value As Integer)
            _indentSpaces = value
        End Set
    End Property

    Private _tag As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 1, .Color = Color.Blue, .ColorIndex = 1, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML tag.
    ''' <summary>
    ''' The display properties of the XML tags.
    ''' </summary>
    Property Tag As TextSettings
        Get
            Return _tag
        End Get
        Set(value As TextSettings)
            Dim OldFontName As String = _tag.FontName
            Dim OldColor As Color = _tag.Color
            _tag = value
            'Update the indexes if a color or font has been changed.
            If _tag.FontName <> OldFontName Then
                UpdateFontIndexes()
            End If
            If _tag.Color <> OldColor Then
                UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _element As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 1, .Color = Color.DarkRed, .ColorIndex = 2, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML element.
    ''' <summary>
    ''' The display properties of the XML elements.
    ''' </summary>
    Property Element As TextSettings
        Get
            Return _element
        End Get
        Set(value As TextSettings)
            Dim OldFontName As String = _element.FontName
            Dim OldColor As Color = _element.Color
            _element = value
            'Update the indexes if a color or font has been changed.
            If _element.FontName <> OldFontName Then
                UpdateFontIndexes()
            End If
            If _element.Color <> OldColor Then
                UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _attributeKey As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 1, .Color = Color.Red, .ColorIndex = 3, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML attribute key.
    ''' <summary>
    ''' The display properties of the XML attribute keys.
    ''' </summary>
    Property AttributeKey As TextSettings
        Get
            Return _attributeKey
        End Get
        Set(value As TextSettings)
            Dim OldFontName As String = _attributeKey.FontName
            Dim OldColor As Color = _attributeKey.Color
            _attributeKey = value
            'Update the indexes if a color or font has been changed.
            If _attributeKey.FontName <> OldFontName Then
                UpdateFontIndexes()
            End If
            If _attributeKey.Color <> OldColor Then
                UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _attributeValue As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 1, .Color = Color.Blue, .ColorIndex = 4, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML attribute value.
    ''' <summary>
    ''' The display properties of the XML attribute values.
    ''' </summary>
    Property AttributeValue As TextSettings
        Get
            Return _attributeValue
        End Get
        Set(value As TextSettings)
            Dim OldFontName As String = _attributeValue.FontName
            Dim OldColor As Color = _attributeValue.Color
            _attributeValue = value
            'Update the indexes if a color or font has been changed.
            If _attributeValue.FontName <> OldFontName Then
                UpdateFontIndexes()
            End If
            If _attributeValue.Color <> OldColor Then
                UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _value As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 1, .Color = Color.Black, .ColorIndex = 5, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML element value.
    ''' <summary>
    ''' The display properties of the XML element values.
    ''' </summary>
    Property Value As TextSettings
        Get
            Return _value
        End Get
        Set(value As TextSettings)
            Dim OldFontName As String = _value.FontName
            Dim OldColor As Color = _value.Color
            _value = value
            'Update the indexes if a color or font has been changed.
            If _value.FontName <> OldFontName Then
                UpdateFontIndexes()
            End If
            If _value.Color <> OldColor Then
                UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _comment As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 1, .Color = Color.Gray, .ColorIndex = 6, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML comment.
    ''' <summary>
    ''' The display properties of the XML comments.
    ''' </summary>
    Property Comment As TextSettings
        Get
            Return _comment
        End Get
        Set(value As TextSettings)
            Dim OldFontName As String = _comment.FontName
            Dim OldColor As Color = _comment.Color
            _comment = value
            'Update the indexes if a color or font has been changed.
            If _comment.FontName <> OldFontName Then
                UpdateFontIndexes()
            End If
            If _comment.Color <> OldColor Then
                UpdateColorIndexes()
            End If
        End Set
    End Property

    ''Property TextType(ByVal TypeName As String) As Dictionary(Of String, TextSettings)
    'Property TextType(ByVal TypeName As String) As TextSettings
    '    Get
    '        Return _textType(TypeName)
    '    End Get
    '    Set(value As TextSettings)
    '        _textType(TypeName) = value
    '    End Set
    'End Property
    '    Get
    '        Return _textType(TypeName)
    '    End Get
    '    Set(value As Dictionary(Of String, TextSettings))

    '    End Set
    'End Property

    ''' <summary>
    ''' Updates the index numbers of each unique font in the XML Text Settings properties and the TextType dictionary.
    ''' </summary>
    Public Sub UpdateFontIndexes()
        'Update the Font indexes stored in the XML item properties.
        'Update NFonts

        'XML types:
        '  Tag
        '  Element
        '  AttributeKey
        '  AttributeValue
        '  Value
        '  Comment

        'Debug.Print("Running UpdateFontIndexes.")

        'Process Tag:
        Tag.FontIndex = 1
        NFonts = 1

        'Process Element:
        If Element.FontName = Tag.FontName Then
            Element.FontIndex = 1
        Else
            Element.FontIndex = 2
            NFonts = 2
        End If

        'Process AttributeKey:
        If AttributeKey.FontName = Tag.FontName Then
            AttributeKey.FontIndex = 1
        ElseIf AttributeKey.FontName = Element.FontName Then
            AttributeKey.FontIndex = Element.FontIndex
        Else
            NFonts = NFonts + 1
            AttributeKey.FontIndex = NFonts
        End If

        'Process AttribiteValue:
        If AttributeValue.FontName = Tag.FontName Then
            AttributeValue.FontIndex = 1
        ElseIf AttributeValue.FontName = Element.FontName Then
            AttributeValue.FontIndex = Element.FontIndex
        ElseIf AttributeValue.FontName = AttributeKey.FontName Then
            AttributeValue.FontIndex = AttributeKey.FontIndex
        Else
            NFonts = NFonts + 1
            AttributeValue.FontIndex = NFonts
        End If

        'Process Value:
        If Value.FontName = Tag.FontName Then
            Value.FontIndex = 1
        ElseIf Value.FontName = Element.FontName Then
            Value.FontIndex = Element.FontIndex
        ElseIf Value.FontName = AttributeKey.FontName Then
            Value.FontIndex = AttributeKey.FontIndex
        ElseIf Value.FontName = AttributeValue.FontName Then
            Value.FontIndex = AttributeValue.FontIndex
        Else
            NFonts = NFonts + 1
            Value.FontIndex = NFonts
        End If

        'Process Comment:
        If Comment.FontName = Tag.FontName Then
            Comment.FontIndex = 1
        ElseIf Comment.FontName = Element.FontName Then
            Comment.FontIndex = Element.FontIndex
        ElseIf Comment.FontName = AttributeKey.FontName Then
            Comment.FontIndex = AttributeKey.FontIndex
        ElseIf Comment.FontName = AttributeValue.FontName Then
            Comment.FontIndex = AttributeValue.FontIndex
        ElseIf Comment.FontName = Value.FontName Then
            Comment.FontIndex = Value.FontIndex
        Else
            NFonts = NFonts + 1
            Comment.FontIndex = NFonts
        End If

        'Process Other Text Types:
        Dim I As Integer
        For I = 1 To TextType.Count
            If TextType.ElementAt(I - 1).Value.FontName = Tag.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = 1
            ElseIf TextType.ElementAt(I - 1).Value.FontName = Element.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = Element.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = AttributeKey.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = AttributeKey.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = AttributeValue.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = AttributeValue.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = Value.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = Value.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = Comment.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = Comment.FontIndex
            Else
                Dim J As Integer
                Dim FontFound As Boolean = False
                For J = 1 To I - 1
                    If TextType.ElementAt(I - 1).Value.FontName = TextType.ElementAt(J - 1).Value.FontName Then
                        FontFound = True
                        TextType.ElementAt(I - 1).Value.FontIndex = TextType.ElementAt(J - 1).Value.FontIndex
                        Exit For
                    End If
                Next
                If FontFound = False Then
                    NFonts = NFonts + 1
                    TextType.ElementAt(I - 1).Value.FontIndex = NFonts
                End If
            End If
        Next

    End Sub

    ''' <summary>
    ''' Updates the index numbers of each unique color in the XML Text Settings properties and the TextType dictionary.
    ''' </summary>
    Public Sub UpdateColorIndexes()
        'Update the Color indexes stored in the XML item properties.
        'Update NColors

        'XML types:
        '  Tag
        '  Element
        '  AttributeKey
        '  AttributeValue
        '  Value
        '  Comment

        'Debug.Print("Running UpdateColorIndexes.")

        'Process Tag:
        Tag.ColorIndex = 1
        NColors = 1

        'Process Element:
        If Element.Color = Tag.Color Then
            Element.ColorIndex = 1
        Else
            Element.ColorIndex = 2
            NColors = 2
        End If

        'Process AttributeKey:
        If AttributeKey.Color = Tag.Color Then
            AttributeKey.ColorIndex = 1
        ElseIf AttributeKey.Color = Element.Color Then
            AttributeKey.ColorIndex = Element.ColorIndex
        Else
            NColors = NColors + 1
            AttributeKey.ColorIndex = NColors
        End If

        'Process AttribiteValue:
        If AttributeValue.Color = Tag.Color Then
            AttributeValue.ColorIndex = 1
        ElseIf AttributeValue.Color = Element.Color Then
            AttributeValue.ColorIndex = Element.ColorIndex
        ElseIf AttributeValue.Color = AttributeKey.Color Then
            AttributeValue.ColorIndex = AttributeKey.ColorIndex
        Else
            NColors = NColors + 1
            AttributeValue.ColorIndex = NColors
        End If

        'Process Value:
        If Value.Color = Tag.Color Then
            Value.ColorIndex = 1
        ElseIf Value.Color = Element.Color Then
            Value.ColorIndex = Element.ColorIndex
        ElseIf Value.Color = AttributeKey.Color Then
            Value.ColorIndex = AttributeKey.ColorIndex
        ElseIf Value.Color = AttributeValue.Color Then
            Value.ColorIndex = AttributeValue.ColorIndex
        Else
            NColors = NColors + 1
            Value.ColorIndex = NColors
        End If

        'Process Comment:
        If Comment.Color = Tag.Color Then
            Comment.ColorIndex = 1
        ElseIf Comment.Color = Element.Color Then
            Comment.ColorIndex = Element.ColorIndex
        ElseIf Comment.Color = AttributeKey.Color Then
            Comment.ColorIndex = AttributeKey.ColorIndex
        ElseIf Comment.Color = AttributeValue.Color Then
            Comment.ColorIndex = AttributeValue.ColorIndex
        ElseIf Comment.Color = Value.Color Then
            Comment.ColorIndex = Value.ColorIndex
        Else
            NColors = NColors + 1
            Comment.ColorIndex = NColors
        End If

        'Process Other Text Types:
        Dim I As Integer
        For I = 1 To TextType.Count
            If TextType.ElementAt(I - 1).Value.Color = Tag.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = 1
            ElseIf TextType.ElementAt(I - 1).Value.Color = Element.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = Element.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = AttributeKey.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = AttributeKey.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = AttributeValue.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = AttributeValue.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = Value.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = Value.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = Comment.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = Comment.ColorIndex
            Else
                Dim J As Integer
                Dim ColorFound As Boolean = False
                For J = 1 To I - 1
                    If TextType.ElementAt(I - 1).Value.Color = TextType.ElementAt(J - 1).Value.Color Then
                        ColorFound = True
                        TextType.ElementAt(I - 1).Value.ColorIndex = TextType.ElementAt(J - 1).Value.ColorIndex
                        Exit For
                    End If
                Next
                If ColorFound = False Then
                    NColors = NColors + 1
                    TextType.ElementAt(I - 1).Value.FontIndex = NColors
                End If
            End If
        Next

    End Sub


    ''' <summary>
    ''' Returns the RichTextFormat header code, including the font table and color table.
    ''' </summary>
    Public Function RtfHeader() As String
        'Return the rtf header.

        Dim rtfHeaderString As New StringBuilder()

        '\rtf1              The RTF version.
        '\ansi              The character set.
        '\ansicpg1252       ANSI code page 1252 - Western European.
        '\deff0             Default font is \f0
        '\deflang1033       Default language is 1033 - English (United States)
        '\deflangfe2052     Default language for East Asian text - 2052 - Chinese (PRC)
        '\viewkind4         Document view mode - 4 - Draft view.
        '\uc1               The number of bytes following a \uN Unicode character.
        '\pard              Resets to default paragraph properties.
        rtfHeaderString.Append("{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe2052")
        rtfHeaderString.Append("{\fonttbl{" & RtfFontTableFormatString() & "}")
        rtfHeaderString.Append("}")
        rtfHeaderString.Append("{\colortbl ;" & RtfColorTableFormatString() & "}")
        rtfHeaderString.Append("\viewkind4\uc1\pard")

        Return rtfHeaderString.ToString

    End Function

    ''' <summary>
    ''' Returns the RichTextFormat font table code.
    ''' </summary>
    Public Function RtfFontTableFormatString() As String
        'Return the rtf Font Table.
        'For example:
        '{\fonttbl{\f0\fnil Courier New;}}

        Dim format As String = "{{\f{0}\fnil {1};}}"
        Dim rtfFormatString As New StringBuilder()

        'The first font is the Tag font:
        rtfFormatString.AppendFormat(format, 0, Tag.FontName)

        If NFonts > 1 Then
            Dim I As Integer
            For I = 2 To NFonts
                If Element.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, Element.FontName)
                ElseIf AttributeKey.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, AttributeKey.FontName)
                ElseIf AttributeValue.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, AttributeValue.FontName)
                ElseIf Value.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, Value.FontName)
                ElseIf Comment.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, Comment.FontName)
                Else
                    Dim J As Integer
                    Dim FontFound As Boolean = False
                    For J = 1 To TextType.Count
                        If TextType.ElementAt(J - 1).Value.FontIndex = I Then
                            rtfFormatString.AppendFormat(format, I - 1, TextType.ElementAt(J - 1).Value.FontName)
                            FontFound = True
                            Exit For
                        End If
                    Next
                    If FontFound = False Then
                        'Error: Font Index not found!
                    End If
                End If
            Next
        End If

        Return rtfFormatString.ToString()

    End Function

    ''' <summary>
    ''' Returns the RichTextFormat color table code.
    ''' </summary>
    Public Function RtfColorTableFormatString() As String
        'Return the rtf Color Table.

        'XML types:
        '  Tag
        '  Element
        '  AttributeKey
        '  AttributeValue
        '  Value
        '  Comment

        Dim format As String = "\red{0}\green{1}\blue{2};"
        Dim rtfFormatString As New StringBuilder()

        'The first color table defintion is ";" to indicate "default text color".
        'This is set up by the code that calls RtfColorTableFormatString()

        'The first non-default color is the Tag color:
        rtfFormatString.AppendFormat(format, Tag.Color.R, Tag.Color.G, Tag.Color.B)

        If NColors > 1 Then 'There are more colors to add to the ColorTable:
            Dim I As Integer
            For I = 2 To NColors
                If Element.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, Element.Color.R, Element.Color.G, Element.Color.B)
                ElseIf AttributeKey.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, AttributeKey.Color.R, AttributeKey.Color.G, AttributeKey.Color.B)
                ElseIf AttributeValue.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, AttributeValue.Color.R, AttributeValue.Color.G, AttributeValue.Color.B)
                ElseIf Value.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, Value.Color.R, Value.Color.G, Value.Color.B)
                ElseIf Comment.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, Comment.Color.R, Comment.Color.G, Comment.Color.B)
                Else
                    Dim J As Integer
                    Dim ColorFound As Boolean = False
                    For J = 1 To TextType.Count
                        If TextType.ElementAt(J - 1).Value.ColorIndex = I Then
                            rtfFormatString.AppendFormat(format, TextType.ElementAt(J - 1).Value.Color.R, TextType.ElementAt(J - 1).Value.Color.G, TextType.ElementAt(J - 1).Value.Color.B)
                            ColorFound = True
                            Exit For
                        End If
                    Next
                    If ColorFound = False Then
                        'Error: Color Index not found!
                    End If
                End If
            Next
        End If

        Return rtfFormatString.ToString()

    End Function

    ''' <summary>
    ''' Returns the RichTextFormat settings code corresponding to the specified text settings.
    ''' </summary>
    Public Function RtfTextSettings(ByVal Settings As TextSettings) As String
        'Returns the RTF format string corresponding to the specified text settings.

        Dim rtfString As New StringBuilder

        rtfString.Append("\f" & Settings.FontIndex - 1) 'Set font index.
        rtfString.Append("\cf" & Settings.ColorIndex) 'Set color index.
        rtfString.Append("\fs" & Settings.HalfPointSize) 'Set half point size.

        If Settings.Bold Then
            rtfString.Append("\b") 'Set bold text.
        Else
            rtfString.Append("\b0") 'Set non-bold text.
        End If

        If Settings.Italic Then
            rtfString.Append("\i") 'Set italic text.
        Else
            rtfString.Append("\i0") 'Set non-italic text.
        End If

        Return rtfString.ToString()

    End Function

    ''' <summary>
    ''' Adds a new entry to the TextType dictionary.
    ''' </summary>
    Public Sub AddNewTextType(ByVal TypeName As String)
        'Add a new text type to the TextType dictionary.
        If TextType.ContainsKey(TypeName) Then
            RaiseEvent ErrorMessage("Text type already exists: " & TypeName & vbCrLf)
        Else
            TextType.Add(TypeName, New TextSettings)
            'SettingsChanged = True
        End If

    End Sub

    ''' <summary>
    ''' Clears all the entries from the TextType dictionary.
    ''' </summary>
    Public Sub ClearAllTextTypes()
        'Clear all the text types from the TextType dictionary.
        TextType.Clear()
        'SettingsChanged = True
    End Sub

    ''' <summary>
    ''' Removes the specified entry from the TextType dictionary.
    ''' </summary>
    Public Sub RemoveTextType(ByVal TypeName As String)
        'Remove the selected text type, with name TypeName, from the TextType dictionary.
        If TextType.ContainsKey(TypeName) Then
            TextType.Remove(TypeName)
            'SettingsChanged = True
        Else
            'TypeName entry not found.
        End If
    End Sub

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Event ErrorMessage(ByVal Message As String)
    Event Message(ByVal Message As String)

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class

Public Class TextSettings
    'Settings used to display text.

    Private _fontName As String = "Arial" 'The name of the font used to display the message.
    Property FontName As String
        Get
            Return _fontName
        End Get
        Set(value As String)
            _fontName = value
        End Set
    End Property

    Private _fontIndex As Integer 'The index number of the font in the rich text font table.
    Property FontIndex As Integer
        Get
            Return _fontIndex
        End Get
        Set(value As Integer)
            _fontIndex = value
        End Set
    End Property

    Private _color As Color = Color.Black 'The text color.
    Property Color As Color
        Get
            Return _color
        End Get
        Set(value As Color)
            _color = value
        End Set
    End Property

    Private _colorIndex As Integer 'The index number of the color in the rich text color table.
    Property ColorIndex As Integer
        Get
            Return _colorIndex
        End Get
        Set(value As Integer)
            _colorIndex = value
        End Set
    End Property

    Private _halfPointSize As Integer = 20 'The font size in half-points.
    Property HalfPointSize As Integer
        Get
            Return _halfPointSize
        End Get
        Set(value As Integer)
            _halfPointSize = value
            _pointSize = _halfPointSize / 2
        End Set
    End Property

    Private _pointSize As Single = 10 'The font size in points. (The rtf size is specified in integer half points. The point size will be adjusted to correspond to integer half points.)
    Property PointSize As Single
        Get
            Return _pointSize
        End Get
        Set(value As Single)
            _pointSize = value
            _halfPointSize = Math.Round(PointSize * 2) 'Set HalfPointSize to the nearest integral half point number.
            _pointSize = _halfPointSize / 2 'Calculate the corresponding PointSize.
        End Set
    End Property

    Private _bold As Boolean = False 'If True, text is bold.
    Property Bold As Boolean
        Get
            Return _bold
        End Get
        Set(value As Boolean)
            _bold = value
        End Set
    End Property

    Private _italic As Boolean = False 'If True, text is italic.
    Property Italic As Boolean
        Get
            Return _italic
        End Get
        Set(value As Boolean)
            _italic = value
        End Set
    End Property

End Class
