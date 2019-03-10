Imports System.Text
Imports System.Drawing

Public Class XmlHtmDisplaySettings
    'Settings used to display formatted XML data, HTML data and specified RTF text types.

    Public NFonts As Integer 'The number of unique fonts used to display the XML data and other text types.
    Public NColors As Integer 'The number of unique colors used to display the XML data and other text types.
    Public TextType As New Dictionary(Of String, TextSettings) 'Dictionary of other text settings.

    'List of Properties:

    'XML Display Settings: ---------------------------------------------
    'XmlLargeFileSizeLimit
    'XIndentSpaces
    'XTag
    'XElement
    'XAttributeKey
    'XAttributeValue
    'XValue
    'XComment

    'HTML Display Settings: -------------------------------------------
    'HText          HTML Text
    'HElement       HTML Elements
    'HAttribute     HTML Attributes
    'HComment       HTML Comments
    'HChar          HTML Special Characters
    'HValue         HTML Attribute or Styles Values 
    'HStyle         HTML Style Attribute


#Region "XML Display Properties" '===========================================================================================================================================

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



    Private _xIndentSpaces As Integer = 4 'The number of spaces used to indent the XML code.
    ''' <summary>
    ''' The number of spaces used to indent the XML code.
    ''' </summary>
    Public Property XIndentSpaces As Integer
        Get
            Return _xIndentSpaces
        End Get
        Set(value As Integer)
            _xIndentSpaces = value
        End Set
    End Property

    Private _xTag As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 1, .Color = Color.Blue, .ColorIndex = 1, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML tag.
    ''' <summary>
    ''' The display properties of the XML tags.
    ''' </summary>
    Property XTag As TextSettings
        Get
            Return _xTag
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting XTag Color value to: " & value.Color.Name)
            Dim OldFontName As String = _xTag.FontName
            Dim OldColor As Color = _xTag.Color
            _xTag = value
            'Update the indexes if a color or font has been changed.
            If _xTag.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _xTag.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _xElement As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 2, .Color = Color.DarkRed, .ColorIndex = 2, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML element.
    ''' <summary>
    ''' The display properties of the XML elements.
    ''' </summary>
    Property XElement As TextSettings
        Get
            Return _xElement
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting XElement Color value to: " & value.Color.Name)
            Dim OldFontName As String = _xElement.FontName
            Dim OldColor As Color = _xElement.Color
            _xElement = value
            'Update the indexes if a color or font has been changed.
            If _xElement.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _xElement.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _xAttributeKey As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 3, .Color = Color.Red, .ColorIndex = 3, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML attribute key.
    ''' <summary>
    ''' The display properties of the XML attribute keys.
    ''' </summary>
    Property XAttributeKey As TextSettings
        Get
            Return _xAttributeKey
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting XAttributeKey Color value to: " & value.Color.Name)
            Dim OldFontName As String = _xAttributeKey.FontName
            Dim OldColor As Color = _xAttributeKey.Color
            _xAttributeKey = value
            'Update the indexes if a color or font has been changed.
            If _xAttributeKey.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _xAttributeKey.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _xAttributeValue As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 4, .Color = Color.Blue, .ColorIndex = 4, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML attribute value.
    ''' <summary>
    ''' The display properties of the XML attribute values.
    ''' </summary>
    Property XAttributeValue As TextSettings
        Get
            Return _xAttributeValue
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting XAttributeValue Color value to: " & value.Color.Name)
            Dim OldFontName As String = _xAttributeValue.FontName
            Dim OldColor As Color = _xAttributeValue.Color
            _xAttributeValue = value
            'Update the indexes if a color or font has been changed.
            If _xAttributeValue.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _xAttributeValue.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _xValue As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 5, .Color = Color.Black, .ColorIndex = 5, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML element value.
    ''' <summary>
    ''' The display properties of the XML element values.
    ''' </summary>
    Property XValue As TextSettings
        Get
            Return _xValue
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting XValue Color value to: " & value.Color.Name)
            Dim OldFontName As String = _xValue.FontName
            Dim OldColor As Color = _xValue.Color
            _xValue = value
            'Update the indexes if a color or font has been changed.
            If _xValue.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _xValue.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _xComment As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 6, .Color = Color.Gray, .ColorIndex = 6, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display an XML comment.
    ''' <summary>
    ''' The display properties of the XML comments.
    ''' </summary>
    Property XComment As TextSettings
        Get
            Return _xComment
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting XComment Color value to: " & value.Color.Name)
            Dim OldFontName As String = _xComment.FontName
            Dim OldColor As Color = _xComment.Color
            _xComment = value
            'Update the indexes if a color or font has been changed.
            If _xComment.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _xComment.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property




#End Region 'XML Display Properties -----------------------------------------------------------------------------------------------------------------------------------------


#Region "HTML Display Properties" '==========================================================================================================================================

    Private _hText As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 7, .Color = Color.Black, .ColorIndex = 7, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display HTML text.
    ''' <summary>
    ''' The display properties of HTML text.
    ''' </summary>
    Property HText As TextSettings
        Get
            Return _hText
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HText Color value to: " & value.Color.Name)
            Dim OldFontName As String = _hText.FontName
            Dim OldColor As Color = _hText.Color
            _hText = value
            'Update the indexes if a color or font has been changed.
            If _hText.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _hText.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _hElement As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 8, .Color = Color.DarkRed, .ColorIndex = 8, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display a HTML element.
    ''' <summary>
    ''' The display properties of the HTML elements.
    ''' </summary>
    Property HElement As TextSettings
        Get
            Return _hElement
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HElement Color value to: " & value.Color.Name)
            Dim OldFontName As String = _hElement.FontName
            Dim OldColor As Color = _hElement.Color
            _hElement = value
            'Update the indexes if a color or font has been changed.
            If _hElement.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _hElement.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _hAttribute As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 9, .Color = Color.Red, .ColorIndex = 9, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display a HTML attribute value.
    ''' <summary>
    ''' The display properties of the HTML attributes.
    ''' </summary>
    Property HAttribute As TextSettings
        Get
            Return _hAttribute
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HAttribute Color value to: " & value.Color.Name)
            Dim OldFontName As String = _hAttribute.FontName
            Dim OldColor As Color = _hAttribute.Color
            _hAttribute = value
            'Update the indexes if a color or font has been changed.
            If _hAttribute.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _hAttribute.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _hComment As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 10, .Color = Color.Green, .ColorIndex = 10, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display a HTML comment.
    ''' <summary>
    ''' The display properties of the HTML comments.
    ''' </summary>
    Property HComment As TextSettings
        Get
            Return _hComment
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HComment Color value to: " & value.Color.Name)
            Dim OldFontName As String = _hComment.FontName
            Dim OldColor As Color = _hComment.Color
            _hComment = value
            'Update the indexes if a color or font has been changed.
            If _hComment.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _hComment.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _hChar As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 11, .Color = Color.Blue, .ColorIndex = 11, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display a HTML comment.
    ''' <summary>
    ''' The display properties of the HTML Special Characters.
    ''' </summary>
    Property HChar As TextSettings
        Get
            Return _hChar
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HChar Color value to: " & value.Color.Name)
            Dim OldFontName As String = _hChar.FontName
            Dim OldColor As Color = _hChar.Color
            _hChar = value
            'Update the indexes if a color or font has been changed.
            If _hChar.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _hChar.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _hValue As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 12, .Color = Color.Blue, .ColorIndex = 12, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display a HTML element value.
    ''' <summary>
    ''' The display properties of the HTML attribute or style values.
    ''' </summary>
    Property HValue As TextSettings
        Get
            Return _hValue
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HValue Color value to: " & value.Color.Name)
            Dim OldFontName As String = _hValue.FontName
            Dim OldColor As Color = _hValue.Color
            _hValue = value
            'Update the indexes if a color or font has been changed.
            If _hValue.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _hValue.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    Private _hStyle As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 13, .Color = Color.Purple, .ColorIndex = 13, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display a HTML Style.
    ''' <summary>
    ''' The display properties of the HTML style attributes.
    ''' </summary>
    Property HStyle As TextSettings
        Get
            Return _hStyle
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HStyle Color value to: " & value.Color.Name)
            Dim OldFontName As String = _hStyle.FontName
            Dim OldColor As Color = _hStyle.Color
            _hStyle = value
            'Update the indexes if a color or font has been changed.
            If _hStyle.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _hStyle.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property


#End Region 'HTML Display Properties ----------------------------------------------------------------------------------------------------------------------------------------

    'Private _plainText As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 7, .Color = Color.Black, .ColorIndex = 7, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display plain text.
    Private _plainText As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 7, .Color = Color.Black, .ColorIndex = 7, .PointSize = 12, .Bold = False, .Italic = False} 'The text settings to display plain text.
    ''' <summary>
    ''' The display properties of Plain text.
    ''' </summary>
    Property PlainText As TextSettings
        Get
            Return _plainText
        End Get
        Set(value As TextSettings)
            'Debug.Print("Setting HText Color value to: " & value.Color.Name)
            Dim OldFontName As String = _plainText.FontName
            Dim OldColor As Color = _plainText.Color
            _plainText = value
            'Update the indexes if a color or font has been changed.
            If _plainText.FontName <> OldFontName Then
                'UpdateFontIndexes()
            End If
            If _plainText.Color <> OldColor Then
                'UpdateColorIndexes()
            End If
        End Set
    End Property

    'Private _defaultText As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 7, .Color = Color.Black, .ColorIndex = 7, .PointSize = 10, .Bold = False, .Italic = False} 'The text settings to display default text.
    Private _defaultText As TextSettings = New TextSettings With {.FontName = "Arial", .FontIndex = 7, .Color = Color.Black, .ColorIndex = 7, .PointSize = 12, .Bold = False, .Italic = False} 'The text settings to display default text.
    ''' <summary>
    ''' The display properties of Default text.
    ''' </summary>
    Property DefaultText As TextSettings
        Get
            Return _defaultText
        End Get
        Set(value As TextSettings)
            _defaultText = value
        End Set
    End Property


    ''' <summary>
    ''' Updates the index numbers of each unique font in the XML Text Settings properties, HTML Text Settings properties and the TextType dictionary.
    ''' </summary>
    Public Sub UpdateFontIndexes()
        'Update the Font indexes stored in the XML item properties.
        'Update NFonts

        'XML types:
        '  XTag
        '  XElement
        '  XAttributeKey
        '  XAttributeValue
        '  XValue
        '  XComment

        'HTML types:
        '  HText          HTML Text
        '  HElement       HTML Elements
        '  HAttribute     HTML Attributes
        '  HComment       HTML Comments
        '  HCharacters    HTML Special Characters
        '  HValues        HTML Attribute or Styles Values 
        '  HStyle         HTML Style Attribute

        'Debug.Print("Running UpdateFontIndexes.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        'Process XML Fonts: ------------------------------------------------------------

        'Process Tag:
        _xTag.FontIndex = 1
        NFonts = 1

        'Process XML Element:
        If _xElement.FontName = _xTag.FontName Then
            _xElement.FontIndex = 1
        Else
            _xElement.FontIndex = 2
            NFonts = 2
        End If

        'Process XML AttributeKey:
        If _xAttributeKey.FontName = _xTag.FontName Then
            _xAttributeKey.FontIndex = 1
        ElseIf XAttributeKey.FontName = _XElement.FontName Then
            _xAttributeKey.FontIndex = _xElement.FontIndex
        Else
            NFonts = NFonts + 1
            _xAttributeKey.FontIndex = NFonts
        End If

        'Process XML AttribiteValue:
        If _xAttributeValue.FontName = _xTag.FontName Then
            _xAttributeValue.FontIndex = 1
        ElseIf XAttributeValue.FontName = _XElement.FontName Then
            _xAttributeValue.FontIndex = _xElement.FontIndex
        ElseIf XAttributeValue.FontName = _XAttributeKey.FontName Then
            _xAttributeValue.FontIndex = _xAttributeKey.FontIndex
        Else
            NFonts = NFonts + 1
            _xAttributeValue.FontIndex = NFonts
        End If

        'Process XML Value:
        If _xValue.FontName = XTag.FontName Then
            _xValue.FontIndex = 1
        ElseIf _XValue.FontName = XElement.FontName Then
            _xValue.FontIndex = XElement.FontIndex
        ElseIf _XValue.FontName = XAttributeKey.FontName Then
            _xValue.FontIndex = XAttributeKey.FontIndex
        ElseIf _XValue.FontName = XAttributeValue.FontName Then
            _xValue.FontIndex = XAttributeValue.FontIndex
        Else
            NFonts = NFonts + 1
            _xValue.FontIndex = NFonts
        End If

        'Process XML Comment:
        If _xComment.FontName = _xTag.FontName Then
            _xComment.FontIndex = 1
        ElseIf _xComment.FontName = _XElement.FontName Then
            _xComment.FontIndex = _xElement.FontIndex
        ElseIf _xComment.FontName = _XAttributeKey.FontName Then
            _xComment.FontIndex = _xAttributeKey.FontIndex
        ElseIf _xComment.FontName = _XAttributeValue.FontName Then
            _xComment.FontIndex = _xAttributeValue.FontIndex
        ElseIf _xComment.FontName = _XValue.FontName Then
            _xComment.FontIndex = _xValue.FontIndex
        Else
            NFonts = NFonts + 1
            _xComment.FontIndex = NFonts
        End If

        'Process HTML Fonts: -------------------------------------------------------------

        'Process HTML Text:
        If _hText.FontName = _xTag.FontName Then
            _hText.FontIndex = 1
        ElseIf _HText.FontName = _XElement.FontName Then
            _hText.FontIndex = _xElement.FontIndex
        ElseIf HText.FontName = _XAttributeKey.FontName Then
            _hText.FontIndex = _xAttributeKey.FontIndex
        ElseIf _HText.FontName = _XAttributeValue.FontName Then
            _hText.FontIndex = _xAttributeValue.FontIndex
        ElseIf _HText.FontName = _XValue.FontName Then
            _hText.FontIndex = _xValue.FontIndex
        ElseIf _HText.FontName = _XComment.FontName Then
            _hText.FontIndex = _xComment.FontIndex
        Else
            NFonts = NFonts + 1
            _hText.FontIndex = NFonts
        End If

        'Process HTML Element:
        If _hElement.FontName = _xTag.FontName Then
            _hElement.FontIndex = 1
        ElseIf _HElement.FontName = _XElement.FontName Then
            _hElement.FontIndex = _xElement.FontIndex
        ElseIf _HElement.FontName = _XAttributeKey.FontName Then
            _hElement.FontIndex = _xAttributeKey.FontIndex
        ElseIf _HElement.FontName = _XAttributeValue.FontName Then
            _hElement.FontIndex = _xAttributeValue.FontIndex
        ElseIf _HElement.FontName = _XValue.FontName Then
            _hElement.FontIndex = _xValue.FontIndex
        ElseIf _HElement.FontName = _XComment.FontName Then
            _hElement.FontIndex = _xComment.FontIndex
        ElseIf _HElement.FontName = _HText.FontName Then
            _hElement.FontIndex = _hText.FontIndex
        Else
            NFonts = NFonts + 1
            _hElement.FontIndex = NFonts
        End If

        'Process HTML Attribute:
        If _hAttribute.FontName = _xTag.FontName Then
            _hAttribute.FontIndex = 1
        ElseIf _hAttribute.FontName = _XElement.FontName Then
            _hAttribute.FontIndex = _xElement.FontIndex
        ElseIf _hAttribute.FontName = _XAttributeKey.FontName Then
            _hAttribute.FontIndex = _xAttributeKey.FontIndex
        ElseIf _hAttribute.FontName = _XAttributeValue.FontName Then
            _hAttribute.FontIndex = _xAttributeValue.FontIndex
        ElseIf _hAttribute.FontName = _XValue.FontName Then
            _hAttribute.FontIndex = _xValue.FontIndex
        ElseIf _hAttribute.FontName = _XComment.FontName Then
            _hAttribute.FontIndex = _xComment.FontIndex
        ElseIf _hAttribute.FontName = _HText.FontName Then
            _hAttribute.FontIndex = _hText.FontIndex
        ElseIf _hAttribute.FontName = _HElement.FontName Then
            _hAttribute.FontIndex = _hElement.FontIndex
        Else
            NFonts = NFonts + 1
            _hAttribute.FontIndex = NFonts
        End If

        'Process HTML Comment:
        If _hComment.FontName = _xTag.FontName Then
            _hComment.FontIndex = 1
        ElseIf _hComment.FontName = _XElement.FontName Then
            _hComment.FontIndex = _xElement.FontIndex
        ElseIf _hComment.FontName = _XAttributeKey.FontName Then
            _hComment.FontIndex = _xAttributeKey.FontIndex
        ElseIf _hComment.FontName = _XAttributeValue.FontName Then
            _hComment.FontIndex = _xAttributeValue.FontIndex
        ElseIf _hComment.FontName = _XValue.FontName Then
            _hComment.FontIndex = _xValue.FontIndex
        ElseIf _hComment.FontName = _XComment.FontName Then
            _hComment.FontIndex = _xComment.FontIndex
        ElseIf _hComment.FontName = _HText.FontName Then
            _hComment.FontIndex = _hText.FontIndex
        ElseIf _hComment.FontName = _HElement.FontName Then
            _hComment.FontIndex = _hElement.FontIndex
        ElseIf _hComment.FontName = _HAttribute.FontName Then
            _hComment.FontIndex = _hAttribute.FontIndex
        Else
            NFonts = NFonts + 1
            _hComment.FontIndex = NFonts
        End If


        'Process HTML Special Character:
        If _hChar.FontName = _xTag.FontName Then
            _hChar.FontIndex = 1
        ElseIf _hChar.FontName = _XElement.FontName Then
            _hChar.FontIndex = _xElement.FontIndex
        ElseIf _hChar.FontName = _XAttributeKey.FontName Then
            _hChar.FontIndex = _xAttributeKey.FontIndex
        ElseIf _hChar.FontName = _XAttributeValue.FontName Then
            _hChar.FontIndex = _xAttributeValue.FontIndex
        ElseIf _hChar.FontName = _XValue.FontName Then
            _hChar.FontIndex = _xValue.FontIndex
        ElseIf _hChar.FontName = _XComment.FontName Then
            _hChar.FontIndex = _xComment.FontIndex
        ElseIf _hChar.FontName = _HText.FontName Then
            _hChar.FontIndex = _hText.FontIndex
        ElseIf _hChar.FontName = _HElement.FontName Then
            _hChar.FontIndex = _hElement.FontIndex
        ElseIf _hChar.FontName = _HAttribute.FontName Then
            _hChar.FontIndex = _hAttribute.FontIndex
        ElseIf _hChar.FontName = _HComment.FontName Then
            _hChar.FontIndex = _hComment.FontIndex
        Else
            NFonts = NFonts + 1
            _hChar.FontIndex = NFonts
        End If

        'Process HTML Value:
        If _hValue.FontName = XTag.FontName Then
            _hValue.FontIndex = 1
        ElseIf _HValue.FontName = _XElement.FontName Then
            _hValue.FontIndex = _xElement.FontIndex
        ElseIf _HValue.FontName = _XAttributeKey.FontName Then
            _hValue.FontIndex = _xAttributeKey.FontIndex
        ElseIf _HValue.FontName = _XAttributeValue.FontName Then
            _hValue.FontIndex = _xAttributeValue.FontIndex
        ElseIf _HValue.FontName = _XValue.FontName Then
            _hValue.FontIndex = _xValue.FontIndex
        ElseIf _HValue.FontName = _XComment.FontName Then
            _hValue.FontIndex = _xComment.FontIndex
        ElseIf _HValue.FontName = _HText.FontName Then
            _hValue.FontIndex = _hText.FontIndex
        ElseIf _HValue.FontName = _HElement.FontName Then
            _hValue.FontIndex = _hElement.FontIndex
        ElseIf _HValue.FontName = _HAttribute.FontName Then
            _hValue.FontIndex = _hAttribute.FontIndex
        ElseIf _HValue.FontName = _HComment.FontName Then
            _hValue.FontIndex = _hComment.FontIndex
        ElseIf _HValue.FontName = _HChar.FontName Then
            _hValue.FontIndex = _hChar.FontIndex
        Else
            NFonts = NFonts + 1
            _hValue.FontIndex = NFonts
        End If

        'Process HTML Style:
        If _hStyle.FontName = _xTag.FontName Then
            _hStyle.FontIndex = 1
        ElseIf _hStyle.FontName = _XElement.FontName Then
            _hStyle.FontIndex = _xElement.FontIndex
        ElseIf _hStyle.FontName = _XAttributeKey.FontName Then
            _hStyle.FontIndex = _xAttributeKey.FontIndex
        ElseIf _hStyle.FontName = _XAttributeValue.FontName Then
            _hStyle.FontIndex = _xAttributeValue.FontIndex
        ElseIf _hStyle.FontName = _XValue.FontName Then
            _hStyle.FontIndex = _xValue.FontIndex
        ElseIf _hStyle.FontName = _XComment.FontName Then
            _hStyle.FontIndex = _xComment.FontIndex
        ElseIf _hStyle.FontName = _HText.FontName Then
            _hStyle.FontIndex = _hText.FontIndex
        ElseIf _hStyle.FontName = _HElement.FontName Then
            _hStyle.FontIndex = _hElement.FontIndex
        ElseIf _hStyle.FontName = _HAttribute.FontName Then
            _hStyle.FontIndex = _hAttribute.FontIndex
        ElseIf _hStyle.FontName = _HComment.FontName Then
            _hStyle.FontIndex = _hComment.FontIndex
        ElseIf _hStyle.FontName = _HChar.FontName Then
            _hStyle.FontIndex = _hChar.FontIndex
        ElseIf _hStyle.FontName = _HValue.FontName Then
            _hStyle.FontIndex = _hValue.FontIndex
        Else
            NFonts = NFonts + 1
            _hStyle.FontIndex = NFonts
        End If

        'Process Plain Text:
        If _plainText.FontName = _xTag.FontName Then
            _plainText.FontIndex = 1
        ElseIf _plainText.FontName = _xElement.FontName Then
            _plainText.FontIndex = _xElement.FontIndex
        ElseIf _plainText.FontName = _xAttributeKey.FontName Then
            _plainText.FontIndex = _xAttributeKey.FontIndex
        ElseIf _plainText.FontName = _xAttributeValue.FontName Then
            _plainText.FontIndex = _xAttributeValue.FontIndex
        ElseIf _plainText.FontName = _xValue.FontName Then
            _plainText.FontIndex = _xValue.FontIndex
        ElseIf _plainText.FontName = _xComment.FontName Then
            _plainText.FontIndex = _xComment.FontIndex
        ElseIf _plainText.FontName = _hText.FontName Then
            _plainText.FontIndex = _hText.FontIndex
        ElseIf _plainText.FontName = _hElement.FontName Then
            _plainText.FontIndex = _hElement.FontIndex
        ElseIf _plainText.FontName = _hAttribute.FontName Then
            _plainText.FontIndex = _hAttribute.FontIndex
        ElseIf _plainText.FontName = _hComment.FontName Then
            _plainText.FontIndex = _hComment.FontIndex
        ElseIf _plainText.FontName = _hChar.FontName Then
            _plainText.FontIndex = _hChar.FontIndex
        ElseIf _plainText.FontName = _hValue.FontName Then
            _plainText.FontIndex = _hValue.FontIndex
        ElseIf _plainText.FontName = _hStyle.FontName Then
            _plainText.FontIndex = _hStyle.FontIndex
        Else
            NFonts = NFonts + 1
            _plainText.FontIndex = NFonts
        End If

        'Process Default Text:
        If _defaultText.FontName = _xTag.FontName Then
            _defaultText.FontIndex = 1
        ElseIf _defaultText.FontName = _xElement.FontName Then
            _defaultText.FontIndex = _xElement.FontIndex
        ElseIf _defaultText.FontName = _xAttributeKey.FontName Then
            _defaultText.FontIndex = _xAttributeKey.FontIndex
        ElseIf _defaultText.FontName = _xAttributeValue.FontName Then
            _defaultText.FontIndex = _xAttributeValue.FontIndex
        ElseIf _defaultText.FontName = _xValue.FontName Then
            _defaultText.FontIndex = _xValue.FontIndex
        ElseIf _defaultText.FontName = _xComment.FontName Then
            _defaultText.FontIndex = _xComment.FontIndex
        ElseIf _defaultText.FontName = _hText.FontName Then
            _defaultText.FontIndex = _hText.FontIndex
        ElseIf _defaultText.FontName = _hElement.FontName Then
            _defaultText.FontIndex = _hElement.FontIndex
        ElseIf _defaultText.FontName = _hAttribute.FontName Then
            _defaultText.FontIndex = _hAttribute.FontIndex
        ElseIf _defaultText.FontName = _hComment.FontName Then
            _defaultText.FontIndex = _hComment.FontIndex
        ElseIf _defaultText.FontName = _hChar.FontName Then
            _defaultText.FontIndex = _hChar.FontIndex
        ElseIf _defaultText.FontName = _hValue.FontName Then
            _defaultText.FontIndex = _hValue.FontIndex
        ElseIf _defaultText.FontName = _hStyle.FontName Then
            _defaultText.FontIndex = _hStyle.FontIndex
        ElseIf _defaultText.FontName = _plainText.FontName Then
            _defaultText.FontIndex = _plainText.FontIndex
        Else
            NFonts = NFonts + 1
            _defaultText.FontIndex = NFonts
        End If

        'Process Other Text Types: -------------------------------------------------------
        Dim I As Integer
        For I = 1 To TextType.Count
            If TextType.ElementAt(I - 1).Value.FontName = _xTag.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = 1
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _XElement.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _xElement.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _XAttributeKey.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _xAttributeKey.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _XAttributeValue.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _xAttributeValue.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _XValue.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _xValue.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _XComment.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _xComment.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _HText.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _hText.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _HElement.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _hElement.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _HAttribute.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _hAttribute.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _HComment.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _hComment.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _HChar.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _hChar.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _HValue.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _hValue.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _hStyle.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _hStyle.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _plainText.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _plainText.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = _defaultText.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = _defaultText.FontIndex
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

    Public Sub UpdateFontIndexes_Old()
        'Update the Font indexes stored in the XML item properties.
        'Update NFonts

        'XML types:
        '  XTag
        '  XElement
        '  XAttributeKey
        '  XAttributeValue
        '  XValue
        '  XComment

        'HTML types:
        '  HText          HTML Text
        '  HElement       HTML Elements
        '  HAttribute     HTML Attributes
        '  HComment       HTML Comments
        '  HCharacters    HTML Special Characters
        '  HValues        HTML Attribute or Styles Values 
        '  HStyle         HTML Style Attribute

        'Debug.Print("Running UpdateFontIndexes.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        'Process XML Fonts: ------------------------------------------------------------

        'Process Tag:
        XTag.FontIndex = 1
        NFonts = 1

        'Process XML Element:
        If XElement.FontName = XTag.FontName Then
            XElement.FontIndex = 1
        Else
            XElement.FontIndex = 2
            NFonts = 2
        End If

        'Process XML AttributeKey:
        If XAttributeKey.FontName = XTag.FontName Then
            XAttributeKey.FontIndex = 1
        ElseIf XAttributeKey.FontName = XElement.FontName Then
            XAttributeKey.FontIndex = XElement.FontIndex
        Else
            NFonts = NFonts + 1
            XAttributeKey.FontIndex = NFonts
        End If

        'Process XML AttribiteValue:
        If XAttributeValue.FontName = XTag.FontName Then
            XAttributeValue.FontIndex = 1
        ElseIf XAttributeValue.FontName = XElement.FontName Then
            XAttributeValue.FontIndex = XElement.FontIndex
        ElseIf XAttributeValue.FontName = XAttributeKey.FontName Then
            XAttributeValue.FontIndex = XAttributeKey.FontIndex
        Else
            NFonts = NFonts + 1
            XAttributeValue.FontIndex = NFonts
        End If

        'Process XML Value:
        If XValue.FontName = XTag.FontName Then
            XValue.FontIndex = 1
        ElseIf XValue.FontName = XElement.FontName Then
            XValue.FontIndex = XElement.FontIndex
        ElseIf XValue.FontName = XAttributeKey.FontName Then
            XValue.FontIndex = XAttributeKey.FontIndex
        ElseIf XValue.FontName = XAttributeValue.FontName Then
            XValue.FontIndex = XAttributeValue.FontIndex
        Else
            NFonts = NFonts + 1
            XValue.FontIndex = NFonts
        End If

        'Process XML Comment:
        If XComment.FontName = XTag.FontName Then
            XComment.FontIndex = 1
        ElseIf XComment.FontName = XElement.FontName Then
            XComment.FontIndex = XElement.FontIndex
        ElseIf XComment.FontName = XAttributeKey.FontName Then
            XComment.FontIndex = XAttributeKey.FontIndex
        ElseIf XComment.FontName = XAttributeValue.FontName Then
            XComment.FontIndex = XAttributeValue.FontIndex
        ElseIf XComment.FontName = XValue.FontName Then
            XComment.FontIndex = XValue.FontIndex
        Else
            NFonts = NFonts + 1
            XComment.FontIndex = NFonts
        End If

        'Process HTML Fonts: -------------------------------------------------------------

        'Process HTML Text:
        If HText.FontName = XTag.FontName Then
            HText.FontIndex = 1
        ElseIf HText.FontName = XElement.FontName Then
            HText.FontIndex = XElement.FontIndex
        ElseIf HText.FontName = XAttributeKey.FontName Then
            HText.FontIndex = XAttributeKey.FontIndex
        ElseIf HText.FontName = XAttributeValue.FontName Then
            HText.FontIndex = XAttributeValue.FontIndex
        ElseIf HText.FontName = XValue.FontName Then
            HText.FontIndex = XValue.FontIndex
        ElseIf HText.FontName = XComment.FontName Then
            HText.FontIndex = XComment.FontIndex
        Else
            NFonts = NFonts + 1
            HText.FontIndex = NFonts
        End If

        'Process HTML Element:
        If HElement.FontName = XTag.FontName Then
            HElement.FontIndex = 1
        ElseIf HElement.FontName = XElement.FontName Then
            HElement.FontIndex = XElement.FontIndex
        ElseIf HElement.FontName = XAttributeKey.FontName Then
            HElement.FontIndex = XAttributeKey.FontIndex
        ElseIf HElement.FontName = XAttributeValue.FontName Then
            HElement.FontIndex = XAttributeValue.FontIndex
        ElseIf HElement.FontName = XValue.FontName Then
            HElement.FontIndex = XValue.FontIndex
        ElseIf HElement.FontName = XComment.FontName Then
            HElement.FontIndex = XComment.FontIndex
        ElseIf HElement.FontName = HText.FontName Then
            HElement.FontIndex = HText.FontIndex
        Else
            NFonts = NFonts + 1
            HElement.FontIndex = NFonts
        End If

        'Process HTML Attribute:
        If HAttribute.FontName = XTag.FontName Then
            HAttribute.FontIndex = 1
        ElseIf HAttribute.FontName = XElement.FontName Then
            HAttribute.FontIndex = XElement.FontIndex
        ElseIf HAttribute.FontName = XAttributeKey.FontName Then
            HAttribute.FontIndex = XAttributeKey.FontIndex
        ElseIf HAttribute.FontName = XAttributeValue.FontName Then
            HAttribute.FontIndex = XAttributeValue.FontIndex
        ElseIf HAttribute.FontName = XValue.FontName Then
            HAttribute.FontIndex = XValue.FontIndex
        ElseIf HAttribute.FontName = XComment.FontName Then
            HAttribute.FontIndex = XComment.FontIndex
        ElseIf HAttribute.FontName = HText.FontName Then
            HAttribute.FontIndex = HText.FontIndex
        ElseIf HAttribute.FontName = HElement.FontName Then
            HAttribute.FontIndex = HElement.FontIndex
        Else
            NFonts = NFonts + 1
            HAttribute.FontIndex = NFonts
        End If

        'Process HTML Comment:
        If HComment.FontName = XTag.FontName Then
            HComment.FontIndex = 1
        ElseIf HComment.FontName = XElement.FontName Then
            HComment.FontIndex = XElement.FontIndex
        ElseIf HComment.FontName = XAttributeKey.FontName Then
            HComment.FontIndex = XAttributeKey.FontIndex
        ElseIf HComment.FontName = XAttributeValue.FontName Then
            HComment.FontIndex = XAttributeValue.FontIndex
        ElseIf HComment.FontName = XValue.FontName Then
            HComment.FontIndex = XValue.FontIndex
        ElseIf HComment.FontName = XComment.FontName Then
            HComment.FontIndex = XComment.FontIndex
        ElseIf HComment.FontName = HText.FontName Then
            HComment.FontIndex = HText.FontIndex
        ElseIf HComment.FontName = HElement.FontName Then
            HComment.FontIndex = HElement.FontIndex
        ElseIf HComment.FontName = HAttribute.FontName Then
            HComment.FontIndex = HAttribute.FontIndex
        Else
            NFonts = NFonts + 1
            HComment.FontIndex = NFonts
        End If


        'Process HTML Special Character:
        If HChar.FontName = XTag.FontName Then
            HChar.FontIndex = 1
        ElseIf HChar.FontName = XElement.FontName Then
            HChar.FontIndex = XElement.FontIndex
        ElseIf HChar.FontName = XAttributeKey.FontName Then
            HChar.FontIndex = XAttributeKey.FontIndex
        ElseIf HChar.FontName = XAttributeValue.FontName Then
            HChar.FontIndex = XAttributeValue.FontIndex
        ElseIf HChar.FontName = XValue.FontName Then
            HChar.FontIndex = XValue.FontIndex
        ElseIf HChar.FontName = XComment.FontName Then
            HChar.FontIndex = XComment.FontIndex
        ElseIf HChar.FontName = HText.FontName Then
            HChar.FontIndex = HText.FontIndex
        ElseIf HChar.FontName = HElement.FontName Then
            HChar.FontIndex = HElement.FontIndex
        ElseIf HChar.FontName = HAttribute.FontName Then
            HChar.FontIndex = HAttribute.FontIndex
        ElseIf HChar.FontName = HComment.FontName Then
            HChar.FontIndex = HComment.FontIndex
        Else
            NFonts = NFonts + 1
            HChar.FontIndex = NFonts
        End If

        'Process HTML Value:
        If HValue.FontName = XTag.FontName Then
            HValue.FontIndex = 1
        ElseIf HValue.FontName = XElement.FontName Then
            HValue.FontIndex = XElement.FontIndex
        ElseIf HValue.FontName = XAttributeKey.FontName Then
            HValue.FontIndex = XAttributeKey.FontIndex
        ElseIf HValue.FontName = XAttributeValue.FontName Then
            HValue.FontIndex = XAttributeValue.FontIndex
        ElseIf HValue.FontName = XValue.FontName Then
            HValue.FontIndex = XValue.FontIndex
        ElseIf HValue.FontName = XComment.FontName Then
            HValue.FontIndex = XComment.FontIndex
        ElseIf HValue.FontName = HText.FontName Then
            HValue.FontIndex = HText.FontIndex
        ElseIf HValue.FontName = HElement.FontName Then
            HValue.FontIndex = HElement.FontIndex
        ElseIf HValue.FontName = HAttribute.FontName Then
            HValue.FontIndex = HAttribute.FontIndex
        ElseIf HValue.FontName = HComment.FontName Then
            HValue.FontIndex = HComment.FontIndex
        ElseIf HValue.FontName = HChar.FontName Then
            HValue.FontIndex = HChar.FontIndex
        Else
            NFonts = NFonts + 1
            HValue.FontIndex = NFonts
        End If

        'Process HTML Style:
        If HStyle.FontName = XTag.FontName Then
            HStyle.FontIndex = 1
        ElseIf HStyle.FontName = XElement.FontName Then
            HStyle.FontIndex = XElement.FontIndex
        ElseIf HStyle.FontName = XAttributeKey.FontName Then
            HStyle.FontIndex = XAttributeKey.FontIndex
        ElseIf HStyle.FontName = XAttributeValue.FontName Then
            HStyle.FontIndex = XAttributeValue.FontIndex
        ElseIf HStyle.FontName = XValue.FontName Then
            HStyle.FontIndex = XValue.FontIndex
        ElseIf HStyle.FontName = XComment.FontName Then
            HStyle.FontIndex = XComment.FontIndex
        ElseIf HStyle.FontName = HText.FontName Then
            HStyle.FontIndex = HText.FontIndex
        ElseIf HStyle.FontName = HElement.FontName Then
            HStyle.FontIndex = HElement.FontIndex
        ElseIf HStyle.FontName = HAttribute.FontName Then
            HStyle.FontIndex = HAttribute.FontIndex
        ElseIf HStyle.FontName = HComment.FontName Then
            HStyle.FontIndex = HComment.FontIndex
        ElseIf HStyle.FontName = HChar.FontName Then
            HStyle.FontIndex = HChar.FontIndex
        ElseIf HStyle.FontName = HValue.FontName Then
            HStyle.FontIndex = HValue.FontIndex
        Else
            NFonts = NFonts + 1
            HStyle.FontIndex = NFonts
        End If



        'Process Other Text Types: -------------------------------------------------------
        Dim I As Integer
        For I = 1 To TextType.Count
            If TextType.ElementAt(I - 1).Value.FontName = XTag.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = 1
            ElseIf TextType.ElementAt(I - 1).Value.FontName = XElement.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = XElement.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = XAttributeKey.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = XAttributeKey.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = XAttributeValue.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = XAttributeValue.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = XValue.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = XValue.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = XComment.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = XComment.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = HText.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = HText.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = HElement.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = HElement.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = HAttribute.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = HAttribute.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = HComment.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = HComment.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = HChar.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = HChar.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = HValue.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = HValue.FontIndex
            ElseIf TextType.ElementAt(I - 1).Value.FontName = HStyle.FontName Then
                TextType.ElementAt(I - 1).Value.FontIndex = HStyle.FontIndex
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
    ''' Updates the index numbers of each unique color in the XML Text Settings, HTML Text Settings properties and the TextType dictionary.
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

        'HTML types:
        '  HText          HTML Text
        '  HElement       HTML Elements
        '  HAttribute     HTML Attributes
        '  HComment       HTML Comments
        '  HCharacters    HTML Special Characters
        '  HValues        HTML Attribute or Styles Values 
        '  HStyle         HTML Style Attribute

        'Debug.Print("Running UpdateColorIndexes.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        'Process XML Colors: ------------------------------------------------------------

        'Process XML Tag:
        _xTag.ColorIndex = 1
        NColors = 1

        'Process XML Element:
        If _xElement.Color = _xTag.Color Then
            _xElement.ColorIndex = 1
        Else
            _xElement.ColorIndex = 2
            NColors = 2
        End If

        'Process XML AttributeKey:
        If _xAttributeKey.Color = _xTag.Color Then
            _xAttributeKey.ColorIndex = 1
        ElseIf _XAttributeKey.Color = _XElement.Color Then
            _xAttributeKey.ColorIndex = _xElement.ColorIndex
        Else
            NColors = NColors + 1
            _xAttributeKey.ColorIndex = NColors
        End If

        'Process XML AttribiteValue:
        If _xAttributeValue.Color = _xTag.Color Then
            _xAttributeValue.ColorIndex = 1
        ElseIf _XAttributeValue.Color = _XElement.Color Then
            _xAttributeValue.ColorIndex = _xElement.ColorIndex
        ElseIf _XAttributeValue.Color = _XAttributeKey.Color Then
            _xAttributeValue.ColorIndex = _xAttributeKey.ColorIndex
        Else
            NColors = NColors + 1
            _xAttributeValue.ColorIndex = NColors
        End If

        'Process XML Value:
        If _xValue.Color = _xTag.Color Then
            _xValue.ColorIndex = 1
        ElseIf _XValue.Color = _XElement.Color Then
            _xValue.ColorIndex = _xElement.ColorIndex
        ElseIf _XValue.Color = _XAttributeKey.Color Then
            _xValue.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _XValue.Color = _XAttributeValue.Color Then
            _xValue.ColorIndex = _xAttributeValue.ColorIndex
        Else
            NColors = NColors + 1
            _xValue.ColorIndex = NColors
        End If

        'Process XML Comment:
        If _xComment.Color = _xTag.Color Then
            _xComment.ColorIndex = 1
        ElseIf _XComment.Color = _XElement.Color Then
            _xComment.ColorIndex = _xElement.ColorIndex
        ElseIf _XComment.Color = _XAttributeKey.Color Then
            _xComment.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _XComment.Color = _XAttributeValue.Color Then
            _xComment.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _XComment.Color = _XValue.Color Then
            _xComment.ColorIndex = _xValue.ColorIndex
        Else
            NColors = NColors + 1
            _xComment.ColorIndex = NColors
        End If


        'Process HTML Colors: -------------------------------------------------------------

        'Process HTML Text:
        If _hText.Color = _xTag.Color Then
            _hText.ColorIndex = 1
        ElseIf _HText.Color = _XElement.Color Then
            _hText.ColorIndex = _xElement.ColorIndex
        ElseIf _HText.Color = _XAttributeKey.Color Then
            _hText.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _HText.Color = _XAttributeValue.Color Then
            _hText.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _HText.Color = _XValue.Color Then
            _hText.ColorIndex = _xValue.ColorIndex
        ElseIf _HText.Color = _XComment.Color Then
            _hText.ColorIndex = _xComment.ColorIndex
        Else
            NColors = NColors + 1
            _hText.ColorIndex = NColors
        End If

        'Process HTML Element:
        If _hElement.Color = _xTag.Color Then
            _hElement.ColorIndex = 1
        ElseIf _hElement.Color = _xElement.Color Then
            _hElement.ColorIndex = _xElement.ColorIndex
        ElseIf _hElement.Color = _xAttributeKey.Color Then
            _hElement.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _hElement.Color = _xAttributeValue.Color Then
            _hElement.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _hElement.Color = _xValue.Color Then
            _hElement.ColorIndex = _xValue.ColorIndex
        ElseIf _hElement.Color = _xComment.Color Then
            _hElement.ColorIndex = _xComment.ColorIndex
        ElseIf _hElement.Color = _hText.Color Then
            _hElement.ColorIndex = _hText.ColorIndex
        Else
            NColors = NColors + 1
            _hElement.ColorIndex = NColors
        End If

        'Process HTML Attribute:
        If _hAttribute.Color = _xTag.Color Then
            _hAttribute.ColorIndex = 1
        ElseIf _hAttribute.Color = _xElement.Color Then
            _hAttribute.ColorIndex = _xElement.ColorIndex
        ElseIf _hAttribute.Color = _xAttributeKey.Color Then
            _hAttribute.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _hAttribute.Color = _xAttributeValue.Color Then
            _hAttribute.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _hAttribute.Color = _xValue.Color Then
            _hAttribute.ColorIndex = _xValue.ColorIndex
        ElseIf _hAttribute.Color = _xComment.Color Then
            _hAttribute.ColorIndex = _xComment.ColorIndex
        ElseIf _hAttribute.Color = _hText.Color Then
            _hAttribute.ColorIndex = _hText.ColorIndex
        ElseIf _hAttribute.Color = _hElement.Color Then
            _hAttribute.ColorIndex = _hElement.ColorIndex
        Else
            NColors = NColors + 1
            _hAttribute.ColorIndex = NColors
        End If

        'Process HTML Comment:
        If _hComment.Color = _xTag.Color Then
            _hComment.ColorIndex = 1
        ElseIf _hComment.Color = _XElement.Color Then
            _hComment.ColorIndex = _xElement.ColorIndex
        ElseIf _hComment.Color = _XAttributeKey.Color Then
            _hComment.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _hComment.Color = _XAttributeValue.Color Then
            _hComment.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _hComment.Color = _XValue.Color Then
            _hComment.ColorIndex = _xValue.ColorIndex
        ElseIf _hComment.Color = _XComment.Color Then
            _hComment.ColorIndex = _xComment.ColorIndex
        ElseIf _hComment.Color = _HText.Color Then
            _hComment.ColorIndex = _hText.ColorIndex
        ElseIf _hComment.Color = _HElement.Color Then
            _hComment.ColorIndex = _hElement.ColorIndex
        ElseIf _hComment.Color = _HAttribute.Color Then
            _hComment.ColorIndex = _hAttribute.ColorIndex
        Else
            NColors = NColors + 1
            _hComment.ColorIndex = NColors
        End If

        'Process HTML Special Character:
        If _hChar.Color = _xTag.Color Then
            _hChar.ColorIndex = 1
        ElseIf _HChar.Color = _XElement.Color Then
            _hChar.ColorIndex = _xElement.ColorIndex
        ElseIf _HChar.Color = _XAttributeKey.Color Then
            _hChar.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _HChar.Color = _XAttributeValue.Color Then
            _hChar.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _HChar.Color = _XValue.Color Then
            _hChar.ColorIndex = _xValue.ColorIndex
        ElseIf _HChar.Color = _XComment.Color Then
            _hChar.ColorIndex = _xComment.ColorIndex
        ElseIf _HChar.Color = _HText.Color Then
            _hChar.ColorIndex = _hText.ColorIndex
        ElseIf _HChar.Color = _HElement.Color Then
            _hChar.ColorIndex = _hElement.ColorIndex
        ElseIf _HChar.Color = _HAttribute.Color Then
            _hChar.ColorIndex = _hAttribute.ColorIndex
        ElseIf _HChar.Color = _HComment.Color Then
            _hChar.ColorIndex = _hComment.ColorIndex
        Else
            NColors = NColors + 1
            _hChar.ColorIndex = NColors
        End If

        'Process HTML Value:
        If _hValue.Color = _xTag.Color Then
            _hValue.ColorIndex = 1
        ElseIf _hValue.Color = _xElement.Color Then
            _hValue.ColorIndex = _xElement.ColorIndex
        ElseIf _hValue.Color = _xAttributeKey.Color Then
            _hValue.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _hValue.Color = _xAttributeValue.Color Then
            _hValue.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _hValue.Color = _xValue.Color Then
            _hValue.ColorIndex = _xValue.ColorIndex
        ElseIf _hValue.Color = _xComment.Color Then
            _hValue.ColorIndex = _xComment.ColorIndex
        ElseIf _hValue.Color = _hText.Color Then
            _hValue.ColorIndex = _hText.ColorIndex
        ElseIf _hValue.Color = _hElement.Color Then
            _hValue.ColorIndex = _hElement.ColorIndex
        ElseIf _hValue.Color = _hAttribute.Color Then
            _hValue.ColorIndex = _hAttribute.ColorIndex
        ElseIf _hValue.Color = _hComment.Color Then
            _hValue.ColorIndex = _hComment.ColorIndex
        ElseIf _hValue.Color = _hChar.Color Then
            _hValue.ColorIndex = _hChar.ColorIndex
        Else
            NColors = NColors + 1
            _hValue.ColorIndex = NColors
        End If

        'Process HTML Style:
        If _hStyle.Color = _xTag.Color Then
            _hStyle.ColorIndex = 1
        ElseIf _hStyle.Color = _xElement.Color Then
            _hStyle.ColorIndex = _xElement.ColorIndex
        ElseIf _hStyle.Color = _xAttributeKey.Color Then
            _hStyle.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _hStyle.Color = _xAttributeValue.Color Then
            _hStyle.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _hStyle.Color = _xValue.Color Then
            _hStyle.ColorIndex = _xValue.ColorIndex
        ElseIf _hStyle.Color = _xComment.Color Then
            _hStyle.ColorIndex = _xComment.ColorIndex
        ElseIf _hStyle.Color = _hText.Color Then
            _hStyle.ColorIndex = _hText.ColorIndex
        ElseIf _hStyle.Color = _hElement.Color Then
            _hStyle.ColorIndex = _hElement.ColorIndex
        ElseIf _hStyle.Color = _hAttribute.Color Then
            _hStyle.ColorIndex = _hAttribute.ColorIndex
        ElseIf _hStyle.Color = _hComment.Color Then
            _hStyle.ColorIndex = _hComment.ColorIndex
        ElseIf _hStyle.Color = _hChar.Color Then
            _hStyle.ColorIndex = _hChar.ColorIndex
        ElseIf _hStyle.Color = _hValue.Color Then
            _hStyle.ColorIndex = _hValue.ColorIndex
        Else
            NColors = NColors + 1
            _hStyle.ColorIndex = NColors
        End If

        'Process Plain Text:
        If _plainText.Color = _xTag.Color Then
            _plainText.ColorIndex = 1
        ElseIf _plainText.Color = _xElement.Color Then
            _plainText.ColorIndex = _xElement.ColorIndex
        ElseIf _plainText.Color = _xAttributeKey.Color Then
            _plainText.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _plainText.Color = _xAttributeValue.Color Then
            _plainText.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _plainText.Color = _xValue.Color Then
            _plainText.ColorIndex = _xValue.ColorIndex
        ElseIf _plainText.Color = _xComment.Color Then
            _plainText.ColorIndex = _xComment.ColorIndex
        ElseIf _plainText.Color = _hText.Color Then
            _plainText.ColorIndex = _hText.ColorIndex
        ElseIf _plainText.Color = _hElement.Color Then
            _plainText.ColorIndex = _hElement.ColorIndex
        ElseIf _plainText.Color = _hAttribute.Color Then
            _plainText.ColorIndex = _hAttribute.ColorIndex
        ElseIf _plainText.Color = _hComment.Color Then
            _plainText.ColorIndex = _hComment.ColorIndex
        ElseIf _plainText.Color = _hChar.Color Then
            _plainText.ColorIndex = _hChar.ColorIndex
        ElseIf _plainText.Color = _hValue.Color Then
            _plainText.ColorIndex = _hValue.ColorIndex
        ElseIf _plainText.Color = _hStyle.Color Then
            _plainText.ColorIndex = _hStyle.ColorIndex
        Else
            NColors = NColors + 1
            _plainText.ColorIndex = NColors
        End If

        'Process Default Text:
        If _defaultText.Color = _xTag.Color Then
            _defaultText.ColorIndex = 1
        ElseIf _defaultText.Color = _xElement.Color Then
            _defaultText.ColorIndex = _xElement.ColorIndex
        ElseIf _defaultText.Color = _xAttributeKey.Color Then
            _defaultText.ColorIndex = _xAttributeKey.ColorIndex
        ElseIf _defaultText.Color = _xAttributeValue.Color Then
            _defaultText.ColorIndex = _xAttributeValue.ColorIndex
        ElseIf _defaultText.Color = _xValue.Color Then
            _defaultText.ColorIndex = _xValue.ColorIndex
        ElseIf _defaultText.Color = _xComment.Color Then
            _defaultText.ColorIndex = _xComment.ColorIndex
        ElseIf _defaultText.Color = _hText.Color Then
            _defaultText.ColorIndex = _hText.ColorIndex
        ElseIf _defaultText.Color = _hElement.Color Then
            _defaultText.ColorIndex = _hElement.ColorIndex
        ElseIf _defaultText.Color = _hAttribute.Color Then
            _defaultText.ColorIndex = _hAttribute.ColorIndex
        ElseIf _defaultText.Color = _hComment.Color Then
            _defaultText.ColorIndex = _hComment.ColorIndex
        ElseIf _defaultText.Color = _hChar.Color Then
            _defaultText.ColorIndex = _hChar.ColorIndex
        ElseIf _defaultText.Color = _hValue.Color Then
            _defaultText.ColorIndex = _hValue.ColorIndex
        ElseIf _defaultText.Color = _hStyle.Color Then
            _defaultText.ColorIndex = _hStyle.ColorIndex
        ElseIf _defaultText.Color = _plainText.Color Then
            _defaultText.ColorIndex = _plainText.ColorIndex
        Else
            NColors = NColors + 1
            _defaultText.ColorIndex = NColors
        End If

        'Process Other Text Types: -------------------------------------------------------
        Dim I As Integer
        For I = 1 To TextType.Count
            If TextType.ElementAt(I - 1).Value.Color = _xTag.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = 1
            ElseIf TextType.ElementAt(I - 1).Value.Color = _XElement.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _xElement.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _XAttributeKey.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _xAttributeKey.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _XAttributeValue.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _xAttributeValue.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _XValue.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _xValue.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _XComment.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _xComment.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _HText.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _hText.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _HElement.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _hElement.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _HAttribute.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _hAttribute.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _HComment.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _hComment.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _HChar.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _hChar.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _HValue.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _hValue.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _HStyle.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _hStyle.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _plainText.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _plainText.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = _defaultText.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = _defaultText.ColorIndex
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
    ''' Updates the index numbers of each unique color in the XML Text Settings, HTML Text Settings properties and the TextType dictionary.
    ''' </summary>
    Public Sub UpdateColorIndexes_Old()
        'Update the Color indexes stored in the XML item properties.
        'Update NColors

        'XML types:
        '  Tag
        '  Element
        '  AttributeKey
        '  AttributeValue
        '  Value
        '  Comment

        'HTML types:
        '  HText          HTML Text
        '  HElement       HTML Elements
        '  HAttribute     HTML Attributes
        '  HComment       HTML Comments
        '  HCharacters    HTML Special Characters
        '  HValues        HTML Attribute or Styles Values 
        '  HStyle         HTML Style Attribute

        'Debug.Print("Running UpdateColorIndexes.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        'Process XML Colors: ------------------------------------------------------------

        'Process XML Tag:
        XTag.ColorIndex = 1
        NColors = 1

        'Process XML Element:
        If XElement.Color = XTag.Color Then
            XElement.ColorIndex = 1
        Else
            XElement.ColorIndex = 2
            NColors = 2
        End If

        'Process XML AttributeKey:
        If XAttributeKey.Color = XTag.Color Then
            XAttributeKey.ColorIndex = 1
        ElseIf XAttributeKey.Color = XElement.Color Then
            XAttributeKey.ColorIndex = XElement.ColorIndex
        Else
            NColors = NColors + 1
            XAttributeKey.ColorIndex = NColors
        End If

        'Process XML AttribiteValue:
        If XAttributeValue.Color = XTag.Color Then
            XAttributeValue.ColorIndex = 1
        ElseIf XAttributeValue.Color = XElement.Color Then
            XAttributeValue.ColorIndex = XElement.ColorIndex
        ElseIf XAttributeValue.Color = XAttributeKey.Color Then
            XAttributeValue.ColorIndex = XAttributeKey.ColorIndex
        Else
            NColors = NColors + 1
            XAttributeValue.ColorIndex = NColors
        End If

        'Process XML Value:
        If XValue.Color = XTag.Color Then
            XValue.ColorIndex = 1
        ElseIf XValue.Color = XElement.Color Then
            XValue.ColorIndex = XElement.ColorIndex
        ElseIf XValue.Color = XAttributeKey.Color Then
            XValue.ColorIndex = XAttributeKey.ColorIndex
        ElseIf XValue.Color = XAttributeValue.Color Then
            XValue.ColorIndex = XAttributeValue.ColorIndex
        Else
            NColors = NColors + 1
            XValue.ColorIndex = NColors
        End If

        'Process XML Comment:
        If XComment.Color = XTag.Color Then
            XComment.ColorIndex = 1
        ElseIf XComment.Color = XElement.Color Then
            XComment.ColorIndex = XElement.ColorIndex
        ElseIf XComment.Color = XAttributeKey.Color Then
            XComment.ColorIndex = XAttributeKey.ColorIndex
        ElseIf XComment.Color = XAttributeValue.Color Then
            XComment.ColorIndex = XAttributeValue.ColorIndex
        ElseIf XComment.Color = XValue.Color Then
            XComment.ColorIndex = XValue.ColorIndex
        Else
            NColors = NColors + 1
            XComment.ColorIndex = NColors
        End If


        'Process HTML Colors: -------------------------------------------------------------

        'Process HTML Text:
        If HText.Color = XTag.Color Then
            HText.ColorIndex = 1
        ElseIf HText.Color = XElement.Color Then
            HText.ColorIndex = XElement.ColorIndex
        ElseIf HText.Color = XAttributeKey.Color Then
            HText.ColorIndex = XAttributeKey.ColorIndex
        ElseIf HText.Color = XAttributeValue.Color Then
            HText.ColorIndex = XAttributeValue.ColorIndex
        ElseIf HText.Color = XValue.Color Then
            HText.ColorIndex = XValue.ColorIndex
        ElseIf HText.Color = XComment.Color Then
            HText.ColorIndex = XComment.ColorIndex
        Else
            NColors = NColors + 1
            HText.ColorIndex = NColors
        End If

        'Process HTML Element:
        If HElement.Color = XTag.Color Then
            HElement.ColorIndex = 1
        ElseIf HElement.Color = XElement.Color Then
            HElement.ColorIndex = XElement.ColorIndex
        ElseIf HElement.Color = XAttributeKey.Color Then
            HElement.ColorIndex = XAttributeKey.ColorIndex
        ElseIf HElement.Color = XAttributeValue.Color Then
            HElement.ColorIndex = XAttributeValue.ColorIndex
        ElseIf HElement.Color = XValue.Color Then
            HElement.ColorIndex = XValue.ColorIndex
        ElseIf HElement.Color = XComment.Color Then
            HElement.ColorIndex = XComment.ColorIndex
        ElseIf HElement.Color = HText.Color Then
            HElement.ColorIndex = HText.ColorIndex
        Else
            NColors = NColors + 1
            HElement.ColorIndex = NColors
        End If

        'Process HTML Attribute:
        If HAttribute.Color = XTag.Color Then
            HAttribute.ColorIndex = 1
        ElseIf HAttribute.Color = XElement.Color Then
            HAttribute.ColorIndex = XElement.ColorIndex
        ElseIf HAttribute.Color = XAttributeKey.Color Then
            HAttribute.ColorIndex = XAttributeKey.ColorIndex
        ElseIf HAttribute.Color = XAttributeValue.Color Then
            HAttribute.ColorIndex = XAttributeValue.ColorIndex
        ElseIf HAttribute.Color = XValue.Color Then
            HAttribute.ColorIndex = XValue.ColorIndex
        ElseIf HAttribute.Color = XComment.Color Then
            HAttribute.ColorIndex = XComment.ColorIndex
        ElseIf HAttribute.Color = HText.Color Then
            HAttribute.ColorIndex = HText.ColorIndex
        ElseIf HAttribute.Color = HElement.Color Then
            HAttribute.ColorIndex = HElement.ColorIndex
        Else
            NColors = NColors + 1
            HAttribute.ColorIndex = NColors
        End If

        'Process HTML Comment:
        If HComment.Color = XTag.Color Then
            HComment.ColorIndex = 1
        ElseIf HComment.Color = XElement.Color Then
            HComment.ColorIndex = XElement.ColorIndex
        ElseIf HComment.Color = XAttributeKey.Color Then
            HComment.ColorIndex = XAttributeKey.ColorIndex
        ElseIf HComment.Color = XAttributeValue.Color Then
            HComment.ColorIndex = XAttributeValue.ColorIndex
        ElseIf HComment.Color = XValue.Color Then
            HComment.ColorIndex = XValue.ColorIndex
        ElseIf HComment.Color = XComment.Color Then
            HComment.ColorIndex = XComment.ColorIndex
        ElseIf HComment.Color = HText.Color Then
            HComment.ColorIndex = HText.ColorIndex
        ElseIf HComment.Color = HElement.Color Then
            HComment.ColorIndex = HElement.ColorIndex
        ElseIf HComment.Color = HAttribute.Color Then
            HComment.ColorIndex = HAttribute.ColorIndex
        Else
            NColors = NColors + 1
            HComment.ColorIndex = NColors
        End If

        'Process HTML Special Character:
        If HChar.Color = XTag.Color Then
            HChar.ColorIndex = 1
        ElseIf HChar.Color = XElement.Color Then
            HChar.ColorIndex = XElement.ColorIndex
        ElseIf HChar.Color = XAttributeKey.Color Then
            HChar.ColorIndex = XAttributeKey.ColorIndex
        ElseIf HChar.Color = XAttributeValue.Color Then
            HChar.ColorIndex = XAttributeValue.ColorIndex
        ElseIf HChar.Color = XValue.Color Then
            HChar.ColorIndex = XValue.ColorIndex
        ElseIf HChar.Color = XComment.Color Then
            HChar.ColorIndex = XComment.ColorIndex
        ElseIf HChar.Color = HText.Color Then
            HChar.ColorIndex = HText.ColorIndex
        ElseIf HChar.Color = HElement.Color Then
            HChar.ColorIndex = HElement.ColorIndex
        ElseIf HChar.Color = HAttribute.Color Then
            HChar.ColorIndex = HAttribute.ColorIndex
        ElseIf HChar.Color = HComment.Color Then
            HChar.ColorIndex = HComment.ColorIndex
        Else
            NColors = NColors + 1
            HChar.ColorIndex = NColors
        End If

        'Process HTML Value:
        If HValue.Color = XTag.Color Then
            HValue.ColorIndex = 1
        ElseIf HValue.Color = XElement.Color Then
            HValue.ColorIndex = XElement.ColorIndex
        ElseIf HValue.Color = XAttributeKey.Color Then
            HValue.ColorIndex = XAttributeKey.ColorIndex
        ElseIf HValue.Color = XAttributeValue.Color Then
            HValue.ColorIndex = XAttributeValue.ColorIndex
        ElseIf HValue.Color = XValue.Color Then
            HValue.ColorIndex = XValue.ColorIndex
        ElseIf HValue.Color = XComment.Color Then
            HValue.ColorIndex = XComment.ColorIndex
        ElseIf HValue.Color = HText.Color Then
            HValue.ColorIndex = HText.ColorIndex
        ElseIf HValue.Color = HElement.Color Then
            HValue.ColorIndex = HElement.ColorIndex
        ElseIf HValue.Color = HAttribute.Color Then
            HValue.ColorIndex = HAttribute.ColorIndex
        ElseIf HValue.Color = HComment.Color Then
            HValue.ColorIndex = HComment.ColorIndex
        ElseIf HValue.Color = HChar.Color Then
            HValue.ColorIndex = HChar.ColorIndex
        Else
            NColors = NColors + 1
            HValue.ColorIndex = NColors
        End If

        'Process HTML Style:
        If HStyle.Color = XTag.Color Then
            HStyle.ColorIndex = 1
        ElseIf HStyle.Color = XElement.Color Then
            HStyle.ColorIndex = XElement.ColorIndex
        ElseIf HStyle.Color = XAttributeKey.Color Then
            HStyle.ColorIndex = XAttributeKey.ColorIndex
        ElseIf HStyle.Color = XAttributeValue.Color Then
            HStyle.ColorIndex = XAttributeValue.ColorIndex
        ElseIf HStyle.Color = XValue.Color Then
            HStyle.ColorIndex = XValue.ColorIndex
        ElseIf HStyle.Color = XComment.Color Then
            HStyle.ColorIndex = XComment.ColorIndex
        ElseIf HStyle.Color = HText.Color Then
            HStyle.ColorIndex = HText.ColorIndex
        ElseIf HStyle.Color = HElement.Color Then
            HStyle.ColorIndex = HElement.ColorIndex
        ElseIf HStyle.Color = HAttribute.Color Then
            HStyle.ColorIndex = HAttribute.ColorIndex
        ElseIf HStyle.Color = HComment.Color Then
            HStyle.ColorIndex = HComment.ColorIndex
        ElseIf HStyle.Color = HChar.Color Then
            HStyle.ColorIndex = HChar.ColorIndex
        ElseIf HStyle.Color = HValue.Color Then
            HStyle.ColorIndex = HValue.ColorIndex
        Else
            NColors = NColors + 1
            HStyle.ColorIndex = NColors
        End If

        'Process Other Text Types: -------------------------------------------------------
        Dim I As Integer
        For I = 1 To TextType.Count
            If TextType.ElementAt(I - 1).Value.Color = XTag.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = 1
            ElseIf TextType.ElementAt(I - 1).Value.Color = XElement.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = XElement.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = XAttributeKey.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = XAttributeKey.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = XAttributeValue.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = XAttributeValue.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = XValue.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = XValue.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = XComment.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = XComment.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = HText.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = HText.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = HElement.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = HElement.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = HAttribute.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = HAttribute.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = HComment.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = HComment.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = HChar.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = HChar.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = HValue.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = HValue.ColorIndex
            ElseIf TextType.ElementAt(I - 1).Value.Color = HStyle.Color Then
                TextType.ElementAt(I - 1).Value.ColorIndex = HStyle.ColorIndex
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
        'Debug.Print("Running RtfHeader.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

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
        'Debug.Print("Running RtfFontTableFormatString.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        Dim format As String = "{{\f{0}\fnil {1};}}"
        Dim rtfFormatString As New StringBuilder()

        'The first font is the Tag font:
        rtfFormatString.AppendFormat(format, 0, XTag.FontName)

        If NFonts > 1 Then
            Dim I As Integer
            For I = 2 To NFonts
                If XElement.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, XElement.FontName)
                ElseIf XAttributeKey.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, XAttributeKey.FontName)
                ElseIf XAttributeValue.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, XAttributeValue.FontName)
                ElseIf XValue.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, XValue.FontName)
                ElseIf XComment.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, XComment.FontName)
                ElseIf HText.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, HText.FontName)
                ElseIf HElement.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, HElement.FontName)
                ElseIf HAttribute.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, HAttribute.FontName)
                ElseIf HComment.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, HComment.FontName)
                ElseIf HChar.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, HChar.FontName)
                ElseIf HValue.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, HValue.FontName)
                ElseIf HStyle.FontIndex = I Then
                    rtfFormatString.AppendFormat(format, I - 1, HStyle.FontName)
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
        'Debug.Print("Running RtfColorTableFormatString.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        'XML types:
        '  XTag
        '  XElement
        '  XAttributeKey
        '  XAttributeValue
        '  XValue
        '  XComment

        'HTML types:
        '  HElement       HTML Elements
        '  HAttribute     HTML Attributes
        '  HComment       HTML Comments
        '  HChar          HTML Special Characters
        '  HValue         HTML Attribute or Styles Values 
        '  HStyle         HTML Style Attribute


        Dim format As String = "\red{0}\green{1}\blue{2};"
        Dim rtfFormatString As New StringBuilder()

        'The first color table defintion is ";" to indicate "default text color".
        'This is set up by the code that calls RtfColorTableFormatString()

        'The first non-default color is the Tag color:
        rtfFormatString.AppendFormat(format, XTag.Color.R, XTag.Color.G, XTag.Color.B)

        If NColors > 1 Then 'There are more colors to add to the ColorTable:
            Dim I As Integer
            For I = 2 To NColors
                If XElement.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, XElement.Color.R, XElement.Color.G, XElement.Color.B)
                ElseIf XAttributeKey.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, XAttributeKey.Color.R, XAttributeKey.Color.G, XAttributeKey.Color.B)
                ElseIf XAttributeValue.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, XAttributeValue.Color.R, XAttributeValue.Color.G, XAttributeValue.Color.B)
                ElseIf XValue.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, XValue.Color.R, XValue.Color.G, XValue.Color.B)
                ElseIf XComment.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, XComment.Color.R, XComment.Color.G, XComment.Color.B)
                ElseIf HText.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, HText.Color.R, HText.Color.G, HText.Color.B)
                ElseIf HElement.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, HElement.Color.R, HElement.Color.G, HElement.Color.B)
                ElseIf HAttribute.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, HAttribute.Color.R, HAttribute.Color.G, HAttribute.Color.B)
                ElseIf HComment.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, HComment.Color.R, HComment.Color.G, HComment.Color.B)
                ElseIf HChar.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, HChar.Color.R, HChar.Color.G, HChar.Color.B)
                ElseIf HValue.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, HValue.Color.R, HValue.Color.G, HValue.Color.B)
                ElseIf HStyle.ColorIndex = I Then
                    rtfFormatString.AppendFormat(format, HStyle.Color.R, HStyle.Color.G, HStyle.Color.B)
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
        'Debug.Print("Running RtfTextSettings.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

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

        'Add a space after the RTF format string:
        'rtfString.Append(" ") 'This fixed issue with HTML formatting but causes a new issue with XML formatting!!!

        Return rtfString.ToString()

    End Function

    ''' <summary>
    ''' Adds a new entry to the TextType dictionary.
    ''' </summary>
    Public Sub AddNewTextType(ByVal TypeName As String)
        'Add a new text type to the TextType dictionary.
        'Debug.Print("Running AddNewTextType.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        If TextType.ContainsKey(TypeName) Then
            RaiseEvent ErrorMessage("Text type already exists: " & TypeName & vbCrLf)
        Else
            TextType.Add(TypeName, New TextSettings)
        End If

    End Sub

    ''' <summary>
    ''' Clears all the entries from the TextType dictionary.
    ''' </summary>
    Public Sub ClearAllTextTypes()
        'Debug.Print("Running ClearAllTextTypes.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        'Clear all the text types from the TextType dictionary.
        TextType.Clear()
    End Sub

    ''' <summary>
    ''' Removes the specified entry from the TextType dictionary.
    ''' </summary>
    Public Sub RemoveTextType(ByVal TypeName As String)
        'Remove the selected text type, with name TypeName, from the TextType dictionary.
        'Debug.Print("Running RemoveTextType.")
        'Debug.Print("HAttribute.Color.Name = " & HAttribute.Color.Name)

        If TextType.ContainsKey(TypeName) Then
            TextType.Remove(TypeName)
        Else
            'TypeName entry not found.
        End If
    End Sub

#Region " Events" '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Event ErrorMessage(ByVal Message As String)
    Event Message(ByVal Message As String)

#End Region 'Events -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class



