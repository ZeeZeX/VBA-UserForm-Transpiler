Attribute VB_Name = "VBAForm2HTML"

' VBAForm2HTML v1.1.0
' https://github.com/GUI-Conversion-Tools/VBAForm2HTML
' Copyright (c) 2026 ZeeZeX
' This software is released under the MIT License.
' https://opensource.org/licenses/MIT

Option Explicit

#If VBA7 = 0 Then
    ' Define a placeholder LongPtr type for VBA6 or earlier.
    ' This allows code that uses LongPtr to compile in older VBA versions,
    ' even though those versions do not natively support the LongPtr type.
    Private Enum LongPtr
        [_]
    End Enum
#End If

#If VBA7 Then
    ' 64bit Office / VBA7 or later
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
    Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef image As LongPtr) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
    Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
    Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal stride As Long, ByVal format As Long, ByVal scan0 As LongPtr, ByRef bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, ByRef graphics As LongPtr) As Long
    Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
    Private Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal Mode As Long) As Long
    Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As LongPtr, ByVal image As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
    Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As LongPtr, ByVal filename As LongPtr, ByRef clsidEncoder As GUID, ByVal encoderParams As LongPtr) As Long
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef pclsid As GUID) As Long
    
    Private Declare PtrSafe Function CryptBinaryToStringW Lib "crypt32" (ByVal pbBinary As LongPtr, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As LongPtr, ByRef pcchString As Long) As Long
    
    Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cArgs As Long, ByRef rgvt As Integer, ByRef rgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
#Else
    ' 32bit Office
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    
    Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
    Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef image As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
    Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, ByRef width As Long) As Long
    Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, ByRef height As Long) As Long
    Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal stride As Long, ByVal format As Long, ByVal scan0 As Long, ByRef bitmap As Long) As Long
    Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, ByRef graphics As Long) As Long
    Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
    Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal Mode As Long) As Long
    Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
    Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal filename As Long, ByRef clsidEncoder As GUID, ByVal encoderParams As Long) As Long
    Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef pclsid As GUID) As Long
    
    Private Declare Function CryptBinaryToStringW Lib "crypt32" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
    
    Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cArgs As Long, ByRef rgvt As Integer, ByRef rgpvarg As Long, ByRef pvargResult As Variant) As Long
#End If

Private Type GDIPlusStartupInput: GdiPlusVersion As Long: DebugEventCallback As LongPtr: SuppressBackgroundThread As Long: SuppressExternalCodecs As Long: End Type
Private Type GUID: Data1 As Long: Data2 As Integer: Data3 As Integer: Data4(0 To 7) As Byte: End Type

Private Const OUTPUT_FOLDER_NAME As String = "VBAForm2HTML_output"

Public Sub TestRunConversion2Html()
    Call ConvertForm2HTML(UserForm1)
End Sub

Public Sub TestRunConversion2Html_2()
    Call ConvertForm2HTML(Array(UserForm1, UserForm2))
End Sub

Public Sub ConvertForm2HTML(ByVal frms As Variant, Optional ByVal usePrefix As Boolean = False, Optional ByVal imageMode As String = "file", Optional ByVal langAttribute As String = "")
    
    ' frms: Variant
    '   Accepts a single UserForm object or an Array of UserForm objects to be converted.
    ' usePrefix: Boolean
    '   If set to True, the form name will be added to each element name.
    '   This is automatically set to True if frms is an array.
    ' imageMode: String
    '   Determines how image files used in the UserForm are handled during conversion. You can choose one of the following options:
    '     "file" (Default): Images are saved as separate external files in the output directory, and the generated code references these files.
    '     "disabled": Image processing is disabled, and no image-related code is generated.
    '     "reference-only": Similar to "file", generates code that references image files, but does not export the image files. Useful when the image files already exist.
    '     "base64": Images are embedded directly into the generated code as Base64-encoded strings, keeping everything in a single file.
    ' langAttribute: String
    '   Specifies the language code for the "lang" attribute of the generated <html> tag (e.g., "en" for English, "zh" for Chinese, "ja" for Japanese).
    '   If omitted, the attribute will be left empty.
    
    Dim code As String
    Dim filePath As String
    Dim saveDir As String
    code = GenerateHTMLCode(frms, usePrefix, imageMode, langAttribute)
    If code <> "" Then
        saveDir = GetSaveDirPath()
        Call CreateFolderIfDoesNotExist(saveDir)
        filePath = saveDir & "\output.html"
        Call SaveUtf8TextNoBom(filePath, code)
        MsgBox "Saved: " & filePath
    Else
        MsgBox "Conversion failed."
    End If
    
End Sub


Public Function GenerateHTMLCode(ByVal frms As Variant, Optional ByVal usePrefix As Boolean = False, Optional ByVal imageMode As String = "file", Optional ByVal langAttribute As String = "") As String
    Dim root As Variant
    Dim indent As String
    Dim prefix As String
    Dim formName As String
    Dim controlVarName As String
    Dim parentVarName As String
    Dim unavailableNames() As Variant
    Dim ctrl As Object
    Dim ctrls As Collection
    Dim ctrlsWithZOrderInfo As Collection
    Dim item As Variant
    Dim result As String
    Dim codeList As New Collection
    Const q As String = """"

    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim pixelTop As Long
    Dim pixelLeft As Long
    Dim i As Long
    Dim cursorType As String
    Dim caption As String
    Dim colorSetting As String
    Dim cssSelectorProperties As Collection
    Dim tabPageProperties As Collection
    Dim activeTabProperties As Collection
    Dim inactiveTabProperties As Collection
    Dim tabContainerProperties As Collection
    Dim tabScrollWrapperProperties As Collection
    Dim tabNavButtonProperties As Collection
    Dim tabNavButtonTriangleProperties As Collection
    Dim toggleButtonInputProperties As Collection
    Dim toggleButtonSettingProperties As Collection
    Dim toggleButtonCheckedProperties As Collection
    Dim buttonTextLabelProperties As Collection
    Dim listViewHeaderProperties As Collection
    Dim listViewHeaderAndDataProperties As Collection
    Dim treeViewDetailsUlProperties As Collection
    Dim ctrlValue As String
    Dim spaceCnt As Long
    Dim temp As Variant
    Dim picSize As String
    Dim picPosition As String
    Dim allElementStyleList As New Collection
    Dim allJsCodeList As New Collection
    Dim allHtmlBodyList As New Collection
    Dim currentFormElementStyleList As Collection
    Dim currentFormJsCodeList As Collection
    Dim currentFormHtmlBodyList As Collection
    Dim elementStyle As String
    Dim jsCode As String
    Dim htmlBody As String
    Dim tabHeight As Long
    Dim tabPixelHeight As Long
    Dim positionProperties As Collection
    Dim colorProperties As Collection
    Dim fontProperties As Collection
    Dim otherProperties As Collection
    Dim htmlTitle As String
    Dim isHtmlTitleSet As Boolean: isHtmlTitleSet = False
    Dim enableScrollBar As Boolean
    Dim hasPicture As Boolean
    Dim tempPath As String
    Dim picturePath As String
    Dim pictureName As String
    Dim saveDir As String
    Dim base64Str As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim supportedImageModeValues() As Variant
    
    imageMode = LCase$(imageMode)
    supportedImageModeValues = Array("file", "disabled", "reference-only", "base64")
    If Not ContainsValue(supportedImageModeValues, imageMode) Then
        MsgBox "[imageMode] Invalid value: " & q & imageMode & q & vbLf & "Supported values are " & q & Join(supportedImageModeValues, q & ", " & q) & q
        GenerateHTMLCode = ""
        Exit Function
    End If
    
    saveDir = GetSaveDirPath()
    Call CreateFolderIfDoesNotExist(saveDir)
    
    
    If IsArray(frms) Then
        usePrefix = True
    Else
        frms = VBA.Array(frms)
    End If
    
    Set cssSelectorProperties = New Collection

    With cssSelectorProperties
        .Add "background: #ffffff"
        .Add "display: flex"
        .Add "flex-direction: column"
        .Add "align-items: flex-start"
        .Add "min-width: max-content"
        .Add "gap: 0px"
    End With

    elementStyle = GenerateCssSelector("body", cssSelectorProperties, "") & vbLf
    allElementStyleList.Add elementStyle
    
    Set cssSelectorProperties = New Collection
    
    ' Always show up/down button of <input type="number">
    Set temp = New Collection
    With temp
        .Add "input[type=""number""]::-webkit-inner-spin-button,"
        .Add "input[type=""number""]::-webkit-outer-spin-button {"
        .Add "  opacity: 1;", "  display: block;"
        .Add "}"
    End With
    allElementStyleList.Add JoinCollection(temp, vbLf) & vbLf
    
    For Each root In frms
    
        Set cssSelectorProperties = New Collection
        
        Set currentFormElementStyleList = New Collection
        Set currentFormJsCodeList = New Collection
        Set currentFormHtmlBodyList = New Collection
        
        unavailableNames = VBA.Array("class", "script", "body", "id", "style")
        
        For i = LBound(unavailableNames) To UBound(unavailableNames)
            unavailableNames(i) = LCase$(unavailableNames(i))
        Next
        
        If ContainsValue(unavailableNames, LCase$(root.Name)) Then
            MsgBox GenerateUnavailableNameMessage(root)
            result = ""
            GenerateHTMLCode = result
            Exit Function
        End If
        
        If usePrefix Then
            prefix = root.Name & "-"
        Else
            prefix = ""
        End If
        
        pixelWidth = UserFormSizeToPixel(root.InsideWidth)
        pixelHeight = UserFormSizeToPixel(root.InsideHeight)
        
        formName = root.Name
        
        
        caption = root.caption
        caption = Convert2HTMLFormatText(caption)
        If Not isHtmlTitleSet Then
            htmlTitle = caption
            isHtmlTitleSet = True
        End If
        
        Set cssSelectorProperties = New Collection
        
        With cssSelectorProperties
            .Add "position: relative"
            .Add "margin-left: auto"
            .Add "margin-right: auto"
            .Add "box-sizing: border-box"
            .Add "overflow: hidden"
            .Add "width: " & pixelWidth & "px"
            .Add "height: " & pixelHeight & "px"
            .Add "background: " & LCase$(FormColorToHex(root.BackColor))
            .Add "cursor: " & GetControlCursorType(root)
        End With
        
        Set temp = GetBorderSetting(root)
        Call ExtendCollection(cssSelectorProperties, temp)
        
        currentFormElementStyleList.Add GenerateCssSelector(formName, cssSelectorProperties) & vbLf & vbLf
        
        Set cssSelectorProperties = New Collection
        

        currentFormElementStyleList.Add vbLf

        Set ctrlsWithZOrderInfo = GetAllCtrlsInZOrder(root, zOrderType:="Visual", bfsSort:=False)
        Set ctrls = New Collection
        
        For Each item In ctrlsWithZOrderInfo
            ctrls.Add item(0)
        Next
        
        For Each ctrl In ctrls
            controlVarName = GenerateCtrlVarName(ctrl, prefix)
            parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix)
            enableScrollBar = False
            hasPicture = False
            base64Str = ""
            
            If usePrefix Then
                pictureName = "img_" & root.Name & "_" & ctrl.Name
            Else
                pictureName = "img_" & ctrl.Name
            End If
            
            Set cssSelectorProperties = New Collection
            Set tabPageProperties = New Collection
            Set activeTabProperties = New Collection
            Set inactiveTabProperties = New Collection
            Set tabContainerProperties = New Collection
            Set tabScrollWrapperProperties = New Collection
            Set tabNavButtonProperties = New Collection
            Set tabNavButtonTriangleProperties = New Collection
            Set toggleButtonInputProperties = New Collection
            Set toggleButtonSettingProperties = New Collection
            Set toggleButtonCheckedProperties = New Collection
            Set buttonTextLabelProperties = New Collection
            Set listViewHeaderProperties = New Collection
            Set listViewHeaderAndDataProperties = New Collection
            Set treeViewDetailsUlProperties = New Collection
            Set positionProperties = New Collection
            Set colorProperties = New Collection
            Set fontProperties = New Collection
            Set otherProperties = New Collection
            
            If ContainsValue(unavailableNames, LCase$(ctrl.Name)) Then
                MsgBox GenerateUnavailableNameMessage(ctrl)
                result = ""
                GenerateHTMLCode = result
                Exit Function
            End If
            
            If IsSupportedCtrlType(ctrl) Then
                
                
                If TypeName(ctrl) <> "Page" Then
                    pixelLeft = UserFormSizeToPixel(ctrl.Left)
                    pixelTop = UserFormSizeToPixel(ctrl.Top)
                    pixelWidth = UserFormSizeToPixel(ctrl.width)
                    pixelHeight = UserFormSizeToPixel(ctrl.height)
                
                
                    With positionProperties
                        .Add "position: absolute"
                        .Add "box-sizing: border-box"
                        If TypeName(ctrl) = "Frame" Then
                            .Add "overflow: hidden"
                        End If
                        .Add "left: " & pixelLeft & "px"
                        .Add "top: " & pixelTop & "px"
                        .Add "width: " & pixelWidth & "px"
                        .Add "height: " & pixelHeight & "px"
                    End With
                End If
                
                
                If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox"), TypeName(ctrl)) Or IsListView(ctrl) Then
                    ' Set ForeColor
                    colorProperties.Add "color: " & LCase$(FormColorToHex(ctrl.ForeColor))
                End If
                
                If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "Image", "ComboBox"), TypeName(ctrl)) Or IsListView(ctrl) Then
                    ' Set BackColor
                    colorSetting = LCase$(FormColorToHex(ctrl.BackColor))
                    If ContainsValue(Array("ComboBox", "Label", "TextBox", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "Image"), TypeName(ctrl)) Then
                        If ctrl.BackStyle = fmBackStyleTransparent Then
                            colorSetting = "transparent"
                        End If
                    End If

                    colorProperties.Add "background-color: " & colorSetting
                    
                End If
                
                If TypeName(ctrl) = "ToggleButton" Then
                    toggleButtonInputProperties.Add "display: none"
                    With buttonTextLabelProperties
                        .Add "border-top: 2px solid #ffffff"
                        .Add "border-left: 2px solid #ffffff"
                        .Add "border-right: 2px solid #7a7a7a"
                        .Add "border-bottom: 2px solid #7a7a7a"
                        .Add "box-sizing: border-box"
                        .Add "user-select: none"
                        .Add "-ms-user-select: none"
                    End With
                                        
                    With toggleButtonCheckedProperties
                        .Add "border-top: 2px solid #7a7a7a"
                        .Add "border-left: 2px solid #7a7a7a"
                        .Add "border-right: 2px solid #ffffff"
                        .Add "border-bottom: 2px solid #ffffff"
                    End With
                                        
                End If
                
                If TypeName(ctrl) = "TextBox" Then
                    If ctrl.MultiLine Then
                        With otherProperties
                            Select Case ctrl.ScrollBars
                                Case fmScrollBarsNone
                                    .Add "overflow-x: hidden"
                                    .Add "overflow-y: hidden"
                                Case fmScrollBarsHorizontal
                                    .Add "overflow-x: hidden"
                                    .Add "overflow-y: auto"
                                Case fmScrollBarsVertical
                                    .Add "overflow-x: auto"
                                    .Add "overflow-y: hidden"
                                Case fmScrollBarsBoth
                                    .Add "overflow-x: auto"
                                    .Add "overflow-y: auto"
                            End Select
                        End With
                    End If
                End If
                
                If TypeName(ctrl) = "MultiPage" Then
                    activeTabProperties.Add "padding: 2px 10px"
                    activeTabProperties.Add "background-color: " & LCase$(FormColorToHex(&H8000000F))
                    inactiveTabProperties.Add "padding: 2px 8px"
                    inactiveTabProperties.Add "background-color:  " & LCase$(AddRGB(FormColorToHex(&H8000000F), -20, -20, -20))

                    With tabContainerProperties
                        .Add "position: relative"
                        .Add "width: 100%"
                        .Add "box-sizing: border-box"
                        .Add "white-space: nowrap"
                        .Add "font-size: 0"
                        .Add "padding-right: 36px"
                    End With

                    With tabScrollWrapperProperties
                        .Add "overflow: hidden"
                        .Add "white-space: nowrap"
                        .Add "width: 100%"
                    End With

                    With tabNavButtonProperties
                        .Add "position: absolute"
                        If ctrl.TabOrientation = fmTabOrientationBottom Then
                            .Add "top: 0"
                        Else
                            .Add "bottom: 0"
                        End If
                        .Add "width: 16px"
                        .Add "height: 16px"
                        .Add "box-sizing: border-box"
                        .Add "overflow: hidden"
                        .Add "padding: 0"
                        .Add "margin: 0"
                        .Add "cursor: pointer"
                        .Add "vertical-align: middle"
                    End With
                    
                    With tabNavButtonTriangleProperties
                        .Add "display: block"
                        .Add "width: 0"
                        .Add "height: 0"
                        .Add "border-left: 4px solid transparent"
                        .Add "border-right: 4px solid transparent"
                        .Add "border-bottom: 7px solid #000000"
                        .Add "margin: auto"
                    End With
                    
                    With tabPageProperties
                        .Add "box-sizing: border-box"
                        .Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
                        .Add "cursor: default"
                        If ctrl.Style = fmTabStyleNone Or ctrl.Visible = False Then
                            .Add "display: none;"
                        Else
                            .Add "display: inline-block;"
                        End If
                        .Add "vertical-align: top;"
                        .Add "color: " & LCase$(FormColorToHex(ctrl.ForeColor))
                    End With
                    
                    
                End If
                
                
                If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "ListBox", "CheckBox", "OptionButton", "ComboBox", "ToggleButton", "MultiPage"), TypeName(ctrl)) Or IsListView(ctrl) Or IsTreeView(ctrl) Then
                    With fontProperties
                        .Add "font-family: " & ctrl.Font.Name
                        .Add "font-size: " & Round(ctrl.Font.size * 1.33, 2) & "px"
                        
                        If ctrl.Font.Bold Then .Add "font-weight: bold"
                        If ctrl.Font.Italic Then .Add "font-style: italic"
                        If ctrl.Font.Underline And ctrl.Font.Strikethrough Then
                             .Add "text-decoration: underline line-through"
                        ElseIf ctrl.Font.Underline Then
                             .Add "text-decoration: underline"
                        ElseIf ctrl.Font.Strikethrough Then
                             .Add "text-decoration: line-through"
                        End If
                    End With
                    
                End If
                
                If TypeName(ctrl) = "Page" Then
                    
                    tabHeight = GetTextSizeFromCtrlFontSetting(ctrl.Parent, "TEST")(1)
                    tabPixelHeight = UserFormSizeToPixel(tabHeight)
                    pixelWidth = UserFormSizeToPixel(ctrl.Parent.width)
                    pixelHeight = UserFormSizeToPixel(ctrl.Parent.height)
                    
                    If ctrl.Parent.Style <> fmTabStyleNone Then
                        pixelHeight = pixelHeight - tabPixelHeight - 10
                    End If
                    
                    With positionProperties
                        .Add "position: relative"
                        .Add "box-sizing: border-box"
                        .Add "overflow: hidden"
                        .Add "width: " & pixelWidth & "px"
                        .Add "height: " & pixelHeight & "px"
                        
                    End With
                    
                    With otherProperties
                        .Add "padding: 2px 8px"
                        .Add "background-color: " & LCase$(FormColorToHex(&H8000000F))
                        If ctrl.Parent.Style = fmTabStyleTabs Then
                            .Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
                        Else
                            .Add "border: none"
                        End If
                        .Add "display: none"
                    End With
                End If
                
                If IsListView(ctrl) Then
                    positionProperties.Add "border: 1px solid #666666"
                    positionProperties.Add "outline: 1px solid #ffffff"
                    positionProperties.Add "outline-offset: -2px"
                    otherProperties.Add "overflow: auto"
                    With listViewHeaderProperties
                        .Add "position: sticky"
                        .Add "top: 0"
                        .Add "background-color: " & LCase$(FormColorToHex(&H8000000F))
                        
                        .Add "color: #000000"
                        
                        ' To avoid the strikethrough or underline on table header elements being affected by the font color of the parent element, set them explicitly on the header elements.
                        ' If they are not set individually, according to the HTML specification, the underline and strikethrough defined on the parent element will use the same color as the parent element's font color.
                        ' As a result, when the font colors of the parent and header elements differ, the line color may not match the text color.
                        If ctrl.Font.Underline And ctrl.Font.Strikethrough Then
                             .Add "text-decoration: underline line-through"
                        ElseIf ctrl.Font.Underline Then
                             .Add "text-decoration: underline"
                        ElseIf ctrl.Font.Strikethrough Then
                             .Add "text-decoration: line-through"
                        End If
                    End With
                    
                    With listViewHeaderAndDataProperties
                        .Add "border: 1px solid #ccc"
                        .Add "padding: 8px 12px"
                        .Add "word-wrap: break-word"
                        .Add "overflow: break-word"
                        .Add "text-overflow: ellipsis"
                    End With
                    
                End If
                
                If IsTreeView(ctrl) Then
                    otherProperties.Add "overflow: auto"
                    positionProperties.Add "box-shadow: 1px  1px 0 #ffffff, -1px -1px 0 #666666, 2px  2px 0 #eeeeee, -2px -2px 0 #444444"
                    colorProperties.Add "color: #000000"
                    colorProperties.Add "background-color: #ffffff"
                    treeViewDetailsUlProperties.Add "list-style: none"
                    treeViewDetailsUlProperties.Add "margin: 0"
                    
                    If HasScrollProperty(ctrl) Then
                        enableScrollBar = ctrl.Scroll
                    Else
                        enableScrollBar = True
                    End If
                    
                    With otherProperties
                        If enableScrollBar Then
                            .Add "overflow-x: auto"
                            .Add "overflow-y: auto"
                        Else
                            .Add "overflow-x: hidden"
                            .Add "overflow-y: hidden"
                        End If
                    End With
                End If
                
                If ContainsValue(Array("Frame", "TextBox", "ComboBox", "Label", "ListBox", "Image"), TypeName(ctrl)) Then
                    Set temp = GetBorderSetting(ctrl)
                    Call ExtendCollection(positionProperties, temp)
                End If

                If ContainsValue(Array("Label", "TextBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton", "ListBox"), TypeName(ctrl)) Then
                    Set temp = GetTextAlignSetting(ctrl)
                    Call ExtendCollection(positionProperties, temp)
                End If
                
                If TypeName(ctrl) = "ScrollBar" Then
                    If IsVerticalScrollBar(ctrl) Then
                        otherProperties.Add "writing-mode: vertical-lr"
                    End If
                End If
                
                ' Set mouse cursor
                If TypeName(ctrl) <> "MultiPage" And TypeName(ctrl) <> "Page" Then
                    cursorType = GetControlCursorType(ctrl)
                    otherProperties.Add "cursor: " & cursorType
                End If
                
                
                If TypeName(ctrl) = "Image" Then
                    Select Case ctrl.PictureSizeMode
                        Case fmPictureSizeModeClip
                            picSize = "auto"
                        Case fmPictureSizeModeStretch
                            picSize = "cover"
                        Case fmPictureSizeModeZoom
                            picSize = "contain"
                        Case Else
                            picSize = "auto"
                    End Select
                    
                    
                    Select Case ctrl.PictureAlignment
                        Case fmPictureAlignmentTopLeft
                            picPosition = "left top"
                        Case fmPictureAlignmentTopRight
                            picPosition = "right top"
                        Case fmPictureAlignmentCenter
                            picPosition = "center center"
                        Case fmPictureAlignmentBottomLeft
                            picPosition = "left bottom"
                        Case fmPictureAlignmentBottomRight
                            picPosition = "right bottom"
                        Case Else
                            picPosition = "left top"
                    End Select
                    
                    If ctrl.Picture Is Nothing Or imageMode = "disabled" Then
                        hasPicture = False
                    Else
                        hasPicture = True
                    End If
                    
                    If hasPicture Then
                        
                        tempPath = fso.BuildPath(fso.GetSpecialFolder(2).Path, fso.GetTempName())
                        
                        If ContainsValue(Array("base64"), imageMode) Then
                            picturePath = tempPath & "tmp2.png"
                        Else
                            picturePath = saveDir & "\" & pictureName & ".png"
                        End If
                        
                        tempPath = tempPath & ".bmp"
                        
                        If ContainsValue(Array("file", "base64"), imageMode) Then
                            Call SavePicture(ctrl.Picture, tempPath)
                            Call ConvertImageFormat(tempPath, picturePath, False)
                            If fso.FileExists(tempPath) Then Call fso.DeleteFile(tempPath)
                        End If
                        
                        If ContainsValue(Array("base64"), imageMode) Then
                            base64Str = FileToBase64(picturePath)
                            If fso.FileExists(picturePath) Then Call fso.DeleteFile(picturePath)
                        End If
                        
                        If ContainsValue(Array("file", "reference-only"), imageMode) Then
                            otherProperties.Add "background-image: url(" & q & fso.GetFileName(picturePath) & q & ")"
                        ElseIf imageMode = "base64" Then
                            otherProperties.Add "background-image: url(" & q & "data:image/png;base64," & base64Str & q & ")"
                        End If
                    
                    Else
                        otherProperties.Add "background-image: url(" & q & "" & q & ")"
                    End If
                    
                    
                    With otherProperties
                        .Add "background-size: " & picSize
                        .Add "background-position: " & picPosition
                        .Add "background-repeat: no-repeat"
                    End With
                    
                End If
                
                If ctrl.Visible = False Then
                    otherProperties.Add "display: none"
                End If
                
                If TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Then
                    Call ExtendCollection(buttonTextLabelProperties, positionProperties)
                    Call ExtendCollection(buttonTextLabelProperties, colorProperties)
                    Call ExtendCollection(buttonTextLabelProperties, fontProperties)
                    Call ExtendCollection(buttonTextLabelProperties, otherProperties)
                ElseIf TypeName(ctrl) = "ToggleButton" Then
                    Call ExtendCollection(buttonTextLabelProperties, positionProperties)
                    Call ExtendCollection(toggleButtonSettingProperties, colorProperties)
                    Call ExtendCollection(toggleButtonSettingProperties, fontProperties)
                    Call ExtendCollection(toggleButtonSettingProperties, otherProperties)
                ElseIf TypeName(ctrl) = "MultiPage" Then
                    Call ExtendCollection(tabPageProperties, fontProperties)
                    Call ExtendCollection(cssSelectorProperties, positionProperties)
                    Call ExtendCollection(cssSelectorProperties, otherProperties)
                Else
                    Call ExtendCollection(cssSelectorProperties, positionProperties)
                    Call ExtendCollection(cssSelectorProperties, colorProperties)
                    Call ExtendCollection(cssSelectorProperties, fontProperties)
                    Call ExtendCollection(cssSelectorProperties, otherProperties)
                End If
                
                currentFormElementStyleList.Add GenerateCssSelector(controlVarName, cssSelectorProperties) & vbLf & vbLf
                
                If TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Then
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-label", buttonTextLabelProperties, ".") & vbLf & vbLf
                End If
                
                
                If TypeName(ctrl) = "ToggleButton" Then
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-label", buttonTextLabelProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-input", toggleButtonInputProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-setting", toggleButtonSettingProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-input:checked + ." & controlVarName & "-label", toggleButtonCheckedProperties, ".") & vbLf & vbLf
                End If
                
                If TypeName(ctrl) = "MultiPage" Then
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-TabPages", tabPageProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-ActiveTab", activeTabProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-InactiveTab", inactiveTabProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-TabContainer", tabContainerProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-TabScrollWrapper", tabScrollWrapperProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-TabNavBtn", tabNavButtonProperties, ".") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & "-TabNavBtnTriangle", tabNavButtonTriangleProperties, ".") & vbLf & vbLf
                End If
                
                If IsListView(ctrl) Then
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & " th", listViewHeaderProperties, "#") & vbLf & vbLf
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & " th, #" & controlVarName & " td", listViewHeaderAndDataProperties, "#") & vbLf & vbLf
                End If
                
                If IsTreeView(ctrl) Then
                    currentFormElementStyleList.Add GenerateCssSelector(controlVarName & " details ul", treeViewDetailsUlProperties, "#") & vbLf & vbLf
                End If
                
                currentFormElementStyleList.Add vbLf
                
            Else
                MsgBox GenerateUnsupportedControlMessage(ctrl)
                result = ""
                GenerateHTMLCode = result
                Exit Function
            End If
        Next ctrl
        
        
        currentFormHtmlBodyList.Add "<div id=" & q & formName & q & ">"
        currentFormHtmlBodyList.Add vbLf
        
        currentFormHtmlBodyList.Add SetCssAllCtrlsDfs(root, prefix, ctrls)

        currentFormHtmlBodyList.Add vbLf
        currentFormHtmlBodyList.Add "</div>" & vbLf
        
        For Each ctrl In ctrls
            If TypeName(ctrl) = "MultiPage" Then
                currentFormJsCodeList.Add GenerateTabPageSwitchingFunc(ctrl, prefix)
                currentFormJsCodeList.Add vbLf
            End If
        Next
        
        elementStyle = JoinCollection(currentFormElementStyleList, "")
        htmlBody = JoinCollection(currentFormHtmlBodyList, "")
        jsCode = JoinCollection(currentFormJsCodeList, "")
        
        allElementStyleList.Add elementStyle
        allHtmlBodyList.Add htmlBody
        allJsCodeList.Add jsCode
        
    Next root
    
    With codeList
        .Add "<!DOCTYPE html>" & vbLf
        .Add "<html lang=" & q & langAttribute & q & ">" & vbLf
        .Add "<head>" & vbLf
        .Add Space$(2) & "<meta charset=""UTF-8"">" & vbLf
        .Add Space$(2) & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" & vbLf
        .Add Space$(2) & "<title>" & htmlTitle & "</title>" & vbLf
        .Add vbLf
        .Add Space$(2) & "<style>" & vbLf
        temp = JoinCollection(allElementStyleList, vbLf)
        temp = AdjustIndent(temp, 4)
        .Add temp & vbLf
        .Add Space$(2) & "</style>" & vbLf
        .Add "</head>" & vbLf
        
        .Add "<body>" & vbLf
        temp = JoinCollection(allHtmlBodyList, vbLf) & vbLf
        temp = AdjustIndent(temp, 2)
        .Add temp & vbLf
        
        .Add Space$(2) & "<script>" & vbLf
        temp = JoinCollection(allJsCodeList, vbLf)
        temp = AdjustIndent(temp, 4)
        .Add temp & vbLf
        .Add Space$(2) & "</script>" & vbLf
        
        .Add "</body>" & vbLf
        .Add "</html>"
    End With
    result = JoinCollection(codeList, "")
    GenerateHTMLCode = result
End Function

Private Function GetSaveDirPath() As String
    Dim result As String
    Dim wsh As Object
    result = ""
    On Error Resume Next
    result = CallByName(Application, "ThisWorkbook", VbGet).Path ' Excel VBA
    result = CallByName(Application, "MacroContainer", VbGet).Path ' Word VBA
    On Error GoTo 0
    
    If result = "" Then ' Other Office / Unsaved Workbook
        On Error Resume Next
        Set wsh = CreateObject("WScript.Shell")
        result = wsh.SpecialFolders("MyDocuments")
        On Error GoTo 0
    End If
    
    If result = "" Then ' If fail to retrieve the path to the Documents folder
        result = "C:"
    End If
    
    result = result & "\" & OUTPUT_FOLDER_NAME
    GetSaveDirPath = result
End Function

Private Sub CreateFolderIfDoesNotExist(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Call fso.CreateFolder(folderPath)
    End If
    
End Sub

Private Function IsSupportedCtrlType(ByVal ctrl As Object) As Boolean
    ' Check if the control can be converted to CSS element
    Dim result As Boolean
    Select Case TypeName(ctrl)
        Case "Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", _
            "OptionButton", "Image", "ScrollBar", "ComboBox", "MultiPage", "Page", "ToggleButton"
            result = True
        Case Else
            result = False
            If result = False Then result = IsListView(ctrl)
            If result = False Then result = IsTreeView(ctrl)
    End Select
    
    IsSupportedCtrlType = result
End Function

Private Function SetCssElement(ByVal ctrl As Variant, ByVal prefix As String) As String
    Dim codeList As New Collection
    Dim controlVarName As String
    Dim parentVarName As String
    Dim ctrlValue As String
    Dim spaceCnt As Long
    Dim result As String
    Const q As String = """"
    Dim ctrlDepth As Long
    Dim temp As Variant
    Dim listBoxSize As Long
    
    controlVarName = GenerateCtrlVarName(ctrl, prefix)
    parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix)
    ctrlValue = ""

    If ContainsValue(Array("Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        ctrlValue = Convert2HTMLFormatText(ctrl.caption)
    End If

    If ContainsValue(Array("TextBox", "ComboBox"), TypeName(ctrl)) Then
        ctrlValue = Convert2HTMLFormatText(ctrl.value, useBrTag:=False)
    End If
    
    ctrlDepth = GetFormControlDepth(ctrl)
    spaceCnt = ctrlDepth * 2

    Select Case TypeName(ctrl)
        Case "Label"
            codeList.Add Space$(spaceCnt) & "<div id=" & q & controlVarName & q & ">" & ctrlValue & "</div>"
        Case "TextBox"
            If ctrl.MultiLine Then
                codeList.Add Space$(spaceCnt) & "<textarea id=" & q & controlVarName & q
                
                If ctrl.WordWrap Then
                    codeList.Add " wrap=""soft"""
                Else
                    codeList.Add " wrap=""off"""
                End If
                
                If ctrl.Locked Then
                    codeList.Add " disabled"
                End If
                
                codeList.Add ">" & ctrlValue & "</textarea>"
            Else
                codeList.Add Space$(spaceCnt) & "<input id=" & q & controlVarName & q
                If ctrl.PasswordChar <> "" Then
                    codeList.Add " type=""password"" "
                Else
                    codeList.Add " type=""text"" "
                End If
                codeList.Add "value=" & q & ctrlValue & q
                If ctrl.Locked Then
                    codeList.Add " disabled"
                End If
                codeList.Add ">"
            End If
        Case "SpinButton"
            codeList.Add Space$(spaceCnt) & "<input id=" & q & controlVarName & q & " type=""number"">"
        Case "ComboBox"

            If ctrl.Style = fmStyleDropDownList Then
                codeList.Add Space$(spaceCnt) & "<select id=" & q & controlVarName & q & ">" & vbLf & GenerateCssSelectOptions(ctrl, spaceCnt) & Space$(spaceCnt) & "</select>"
            Else
                codeList.Add Space$(spaceCnt) & "<input id=" & q & controlVarName & q & " list=" & q & controlVarName & "-items-value" & q & " value=" & q & ctrlValue & q & ">" & vbLf
                codeList.Add GenerateCssDataList(ctrl, spaceCnt, prefix)
            End If

        Case "ListBox"
            codeList.Add Space$(spaceCnt) & "<select "
            If ctrl.MultiSelect = fmMultiSelectMulti Or ctrl.MultiSelect = fmMultiSelectExtended Then
                codeList.Add "multiple "
            End If
            ' If the size is 1 or less, <select> will appear as a combo box, so specify a value of at least 2.
            If ctrl.ListCount > 1 Then
                listBoxSize = ctrl.ListCount
            Else
                listBoxSize = 2
            End If
            codeList.Add "id=" & q & controlVarName & q & " size=" & listBoxSize & ">" & vbLf & GenerateCssSelectOptions(ctrl, spaceCnt) & Space$(spaceCnt) & "</select>"
            
        Case "CheckBox"
            codeList.Add Space$(spaceCnt) & "<label for=" & q & controlVarName & q & " class=" & q & controlVarName & "-label" & q & ">"
            temp = "<input " & "id=" & q & controlVarName & q & " type=""checkbox"""
            If ctrl.value Then
                temp = temp & " checked>"
            Else
                temp = temp & ">"
            End If
            If ctrl.Alignment = fmAlignmentLeft Then
                codeList.Add ctrlValue & temp & "</label>"
            Else
                codeList.Add temp & ctrlValue & "</label>"
            End If
            
        Case "OptionButton"
            codeList.Add Space$(spaceCnt) & "<label for=" & q & controlVarName & q & " class=" & q & controlVarName & "-label" & q & ">"
            temp = "<input " & "id=" & q & controlVarName & q & " type=""radio"" " & "name=" & q & parentVarName & "-radio" & q & " value=" & q & controlVarName & "-value" & q
            If ctrl.value Then
                temp = temp & " checked>"
            Else
                temp = temp & ">"
            End If
            If ctrl.Alignment = fmAlignmentLeft Then
                codeList.Add ctrlValue & temp & "</label>"
            Else
                codeList.Add temp & ctrlValue & "</label>"
            End If
        Case "ToggleButton"
            
            temp = Space$(spaceCnt) & "<input " & "id=" & q & controlVarName & q & " type=""checkbox"""
            temp = temp & " class=" & q & controlVarName & "-input " & controlVarName & "-setting" & q

            If ctrl.value Then
                temp = temp & " checked>"
            Else
                temp = temp & ">"
            End If
            
            temp = temp & "<label for=" & q & controlVarName & q & " class=" & q & controlVarName & "-label " & controlVarName & "-setting" & q & ">"
            temp = temp & ctrlValue & "</label>"
            codeList.Add temp
            
        Case "Image"
            codeList.Add Space$(spaceCnt) & "<div id=" & q & controlVarName & q & "></div>"
        Case "ScrollBar"
            codeList.Add Space$(spaceCnt) & "<input id=" & q & controlVarName & q & " type=""range"" min=" & q & ctrl.Min & q & " max=" & q & ctrl.Max & q & " step=""1"" value=" & q & ctrl.value & q & ">"
        Case "CommandButton"
            codeList.Add Space$(spaceCnt) & "<button id=" & q & controlVarName & q & ">" & ctrlValue & "</button>"
    End Select

    If IsListView(ctrl) Then
        codeList.Add Space$(spaceCnt) & "<div " & "id=" & q & controlVarName & q & ">" & vbLf
        codeList.Add Space$(spaceCnt + 2) & "<table style=""border-collapse: collapse; table-layout: fixed"">" & vbLf
        codeList.Add DefineListViewColumns(ctrl, spaceCnt + 4) & vbLf
        codeList.Add GetListViewItems(ctrl, spaceCnt + 4) & vbLf
        codeList.Add Space$(spaceCnt + 2) & "</table>" & vbLf
        codeList.Add Space$(spaceCnt) & "</div>"
    End If

    If IsTreeView(ctrl) Then
        codeList.Add Space$(spaceCnt) & "<div " & "id=" & q & controlVarName & q & ">" & vbLf
        codeList.Add GenerateTreeViewHtmlDfs(ctrl, spaceCnt + 2) & vbLf
        codeList.Add Space$(spaceCnt) & "</div>"
    End If

    codeList.Add vbLf
    result = JoinCollection(codeList, "")
    SetCssElement = result
End Function

Private Function SetCssAllCtrlsDfs(ByVal parentCtrl As Object, ByVal prefix As String, ByVal allCtrls As Variant) As String
    Dim codeList As New Collection
    Dim resultCode As String
    Dim children As Variant
    Dim recursiveResult As Collection
    Set children = ExtractDirectChildCtrlsFromList(parentCtrl, allCtrls)
    Set recursiveResult = SetCssAllCtrlsDfsRecursive(children, codeList, prefix, allCtrls)
    resultCode = JoinCollection(codeList, "")
    SetCssAllCtrlsDfs = resultCode
End Function

Private Function SetCssAllCtrlsDfsRecursive(ByVal ctrls As Variant, ByRef codeList As Collection, ByVal prefix As String, ByVal allCtrls As Variant) As Collection
    Dim shouldProcess As Boolean: shouldProcess = False
    Dim results As New Collection
    Dim ctrl As Variant
    Dim ctrl2 As Variant
    Dim ctrlDepth As Long
    Dim children As Variant
    Dim controlVarName As String
    Dim i As Long
    Const q As String = """"
    If ctrls.count > 0 Then shouldProcess = True
    If shouldProcess Then
        For Each ctrl In ctrls
            results.Add ctrl
            controlVarName = GenerateCtrlVarName(ctrl, prefix)
            ctrlDepth = GetFormControlDepth(ctrl)
            codeList.Add SetCssElement(ctrl, prefix)
            Set children = ExtractDirectChildCtrlsFromList(ctrl, allCtrls)
            
            If TypeName(ctrl) = "Frame" Then
                codeList.Add Space$(ctrlDepth * 2) & "<fieldset id=" & q & controlVarName & q & ">"
                If ctrl.caption <> "" Then
                    codeList.Add "<legend>" & Convert2HTMLFormatText(ctrl.caption) & "</legend>"
                End If
                codeList.Add vbLf
            ElseIf TypeName(ctrl) = "MultiPage" Then
                codeList.Add Space$(ctrlDepth * 2) & "<div id=" & q & controlVarName & q & ">" & vbLf
                If ctrl.TabOrientation <> fmTabOrientationBottom Then
                    codeList.Add GenerateCssMultiPageTabsCode(ctrl, controlVarName, prefix, ctrlDepth) & vbLf
                End If
            ElseIf TypeName(ctrl) = "Page" Then
                codeList.Add Space$(ctrlDepth * 2) & "<div id=" & q & controlVarName & q
                If ctrl.Visible = False Then
                    codeList.Add " style=""display: none"""
                Else
                    For Each ctrl2 In ctrl.Parent.Pages
                        If ctrl2.Visible Then
                            If ctrl Is ctrl2 Then
                                ' Setting for first visible tab page
                                codeList.Add " style=""display: block"""
                            End If
                            Exit For
                        End If
                    Next

                End If
                
                codeList.Add ">" & vbLf
            End If

            Call ExtendCollection(results, SetCssAllCtrlsDfsRecursive(children, codeList, prefix, allCtrls))
            
            If TypeName(ctrl) = "Frame" Then
                codeList.Add Space$(ctrlDepth * 2) & "</fieldset>" & vbLf
            ElseIf TypeName(ctrl) = "MultiPage" Then
                
                If ctrl.TabOrientation = fmTabOrientationBottom Then
                    codeList.Add GenerateCssMultiPageTabsCode(ctrl, controlVarName, prefix, ctrlDepth) & vbLf
                End If
                
                codeList.Add Space$(ctrlDepth * 2) & "</div>" & vbLf
            ElseIf TypeName(ctrl) = "Page" Then
                codeList.Add Space$(ctrlDepth * 2) & "</div>" & vbLf
            End If
        Next
    End If
    Set SetCssAllCtrlsDfsRecursive = results
End Function

Private Function GenerateCssMultiPageTabsCode(ByVal ctrl As Object, ByVal controlVarName As String, ByVal prefix As String, ByVal ctrlDepth As Long) As String
    ' Return code string with scroll container and navigation buttons

    Dim codeList As New Collection
    Dim result As String
    Dim multiPageFuncName As String
    Dim multiPageScrollFuncName As String
    Dim multiPageTabClassName As String
    Dim multiPageTabState As String
    Dim pageName As String
    Dim IsActiveTabSet As Boolean
    Dim i As Long
    Dim ctrl2 As Variant
    Const q As String = """"

    multiPageFuncName = GenerateMultiPageFuncName(ctrl, prefix)
    multiPageScrollFuncName = GenerateMultiPageScrollFuncName(ctrl, prefix)
    multiPageTabClassName = controlVarName & "-TabPages"

    codeList.Add Space$(ctrlDepth * 2 + 2) & "<div class=" & q & controlVarName & "-TabContainer" & q & ">" & vbLf
    codeList.Add Space$(ctrlDepth * 2 + 4) & "<div id=" & q & controlVarName & "-TabScrollWrapper" & q & " class=" & q & controlVarName & "-TabScrollWrapper" & q & ">" & vbLf

    i = 0
    IsActiveTabSet = False
    For Each ctrl2 In ctrl.Pages
        pageName = GenerateCtrlVarName(ctrl2, prefix)
        If ctrl2.Visible And IsActiveTabSet = False Then
            multiPageTabState = controlVarName & "-ActiveTab"
            IsActiveTabSet = True
        Else
            multiPageTabState = controlVarName & "-InactiveTab"
        End If
        
        codeList.Add Space$(ctrlDepth * 2 + 6) & "<div id=" & q & pageName & "-Tab" & q & " class=" & q & multiPageTabClassName & " " & multiPageTabState & q & " onclick=" & q & multiPageFuncName & "(" & i & ")" & q
        If ctrl2.Visible = False Then
            codeList.Add " style=""display: none"""
        End If
        codeList.Add ">" & Convert2HTMLFormatText(ctrl2.caption) & "</div>" & vbLf
        i = i + 1
    Next

    codeList.Add Space$(ctrlDepth * 2 + 4) & "</div>" & vbLf
    codeList.Add Space$(ctrlDepth * 2 + 4) & "<button type=""button"" id=" & q & controlVarName & "-TabNavBtnLeft" & q & " class=" & q & controlVarName & "-TabNavBtn" & q & " style=""right: 18px;"" onclick=" & q & multiPageScrollFuncName & "(-50)" & q & "><span id=" & q & controlVarName & "-TabNavBtnTriangleLeft" & q & " class=" & q & controlVarName & "-TabNavBtnTriangle" & q & " style=""transform: rotate(270deg);""></span></button>" & vbLf
    codeList.Add Space$(ctrlDepth * 2 + 4) & "<button type=""button"" id=" & q & controlVarName & "-TabNavBtnRight" & q & " class=" & q & controlVarName & "-TabNavBtn" & q & " style=""right: 0px;"" onclick=" & q & multiPageScrollFuncName & "(50)" & q & "><span id=" & q & controlVarName & "-TabNavBtnTriangleRight" & q & " class=" & q & controlVarName & "-TabNavBtnTriangle" & q & " style=""transform: rotate(90deg);""></span></button>" & vbLf
    codeList.Add Space$(ctrlDepth * 2 + 2) & "</div>"
    result = JoinCollection(codeList, "")
    GenerateCssMultiPageTabsCode = result
End Function

Private Function ExtractDirectChildCtrlsFromList(ByVal parentCtrl As Object, ByVal allCtrls As Variant) As Collection
    ' Extracts the direct child controls of the specified parent from the given list.
    '
    ' The controls are extracted from the existing list rather than from the
    ' parent's Controls collection to preserve the order established by the
    ' previous sorting process.
    '
    ' Only controls whose immediate parent matches parentCtrl are included.
    ' Descendant controls at deeper levels are not included in the result.
    Dim results As Collection
    Set results = New Collection
    Dim ctrl As Object
    
    For Each ctrl In allCtrls
        If parentCtrl Is ctrl.Parent Then
            results.Add ctrl
        End If
    Next ctrl

    Set ExtractDirectChildCtrlsFromList = results
End Function

Private Function GetUserFormObjectFromCtrl(ByVal ctrl As Object) As Object
    ' Get the ancestor (UserForm) of the control.
    
    Dim root As Object
    
    If ctrl Is Nothing Then
        Err.Raise 13
    End If
    
    Set root = ctrl
    ' Loop to get root(UserForm) object
    On Error GoTo Finally:
    Do While True
        Set root = root.Parent
    Loop
    On Error GoTo 0

Finally:
    Set GetUserFormObjectFromCtrl = root
End Function

Private Function GenerateCtrlVarName(ByVal ctrl As Object, ByVal prefix As String) As String
    ' Generates a valid, unique identifier for a control in the target language.
    Dim controlVarName As String

    If TypeName(ctrl) = "Page" Then
    ' VBA allows duplicate names for Page objects if they belong to different MultiPage controls.
    ' To ensure unique variable names in the target language (which typically uses a flat
    ' namespace), namespace the Page by prepending its parent MultiPage's name.
    ' Example: "Page1" inside "MultiPage1" becomes "MultiPage1_Page1"
        controlVarName = prefix & ctrl.Parent.Name & "_" & ctrl.Name
    Else
        controlVarName = prefix & ctrl.Name
    End If

    GenerateCtrlVarName = controlVarName
End Function


Private Function GenerateCssSelector(ByVal ctrlName As String, ByVal properties As Variant, Optional ByVal selectorSymbol As String = "#") As String
    ' Generate CSS Selector from the control name and the list of properties (Array/Collection)
    
    ' Example:
    ' ctrlName:="Label1", properties:=Array("Left: 299px", "Top: 129px", "width: 97px", "height: 16px", "color: #000000")
    ' ->
    ' #Label1 {
    '   Left: 299px;
    '   Top: 129px;
    '   width: 97px;
    '   height: 16px;
    ' color: #000000;
    ' }
    
    Dim codeList As New Collection
    Dim item As Variant
    Dim result As String
    codeList.Add selectorSymbol & ctrlName & " {"
    For Each item In properties
        codeList.Add vbLf & Space$(2) & item & ";"
    Next
    codeList.Add vbLf & "}"
    result = JoinCollection(codeList, "")
    GenerateCssSelector = result
End Function

Private Function GenerateMultiPageFuncName(ByVal multiPageCtrl As Object, ByVal prefix As String) As String
    ' Generate JavaScript Function Name such as "MultiPage1_SwitchPage" / "UserForm1_MultiPage1_SwitchPage"
    Dim root As Object
    Dim result As String
    Set root = GetUserFormObjectFromCtrl(multiPageCtrl)
    result = prefix & multiPageCtrl.Name & "_SwitchPage"
    result = VBA.Replace$(result, "-", "_")
    GenerateMultiPageFuncName = result
End Function

Private Function GenerateMultiPageScrollFuncName(ByVal multiPageCtrl As Object, ByVal prefix As String) As String
    ' Generate JavaScript Function Name for scrolling such as "MultiPage1_ScrollTab" / "UserForm1_MultiPage1_ScrollTab"
    Dim root As Object
    Dim result As String
    Set root = GetUserFormObjectFromCtrl(multiPageCtrl)
    result = prefix & multiPageCtrl.Name & "_ScrollTab"
    result = VBA.Replace$(result, "-", "_")
    GenerateMultiPageScrollFuncName = result
End Function


Private Function GenerateTabPageSwitchingFunc(ByVal multiPageCtrl As Object, ByVal prefix As String) As String
    '------------------------------------------------------------------------------
    ' Description:
    '   Generates client-side JavaScript functions to manage tab switching, tab strip
    '   scrolling, and navigation button states for a MultiPage control converted to HTML.
    '
    ' Generated JavaScript Functions:
    '   1. <MultiPageName>_SwitchPage(index)
    '      - Displays the page corresponding to the specified tab index (display: "block")
    '        and hides all other pages (display: "none").
    '      - Updates the active and inactive tab styles (padding and background color).
    '      - Automatically scrolls the tab wrapper container to bring the selected tab
    '        fully into view if it is partially or completely out of sight.
    '      - Calls <MultiPageName>_UpdateTabNavBtns() to update navigation button visibility.
    '
    '   2. <MultiPageName>_ScrollTab(delta)
    '      - Scrolls the tab header container horizontally by the specified pixel delta.
    '      - Calls <MultiPageName>_UpdateTabNavBtns() after scrolling.
    '
    '   3. <MultiPageName>_UpdateTabNavBtns()
    '      - Checks if the total tab width exceeds the visible container width.
    '      - Displays or hides the left/right scroll navigation buttons dynamically based
    '        on overflow state.
    '      - Updates the triangle arrow colors (enabled/disabled appearance) depending
    '        on whether the tab container is scrolled to its leftmost or rightmost limits.
    '      - Automatically bound to the window "load" (or "onload") event to ensure proper
    '        initialization when the page finishes loading.
    '------------------------------------------------------------------------------
    Const q As String = """"
    
    Dim ctrl As Variant
    Dim collPages As New Collection
    Dim collTabs As New Collection
    Dim collFuncStr As New Collection
    Dim jsArrPages As String
    Dim jsArrTabs As String
    Dim jsFuncName As String
    Dim jsScrollFuncName As String
    Dim jsUpdateNavFuncName As String
    Dim jsFuncStr As String
    Dim controlVarName As String
    Dim tabVarName As String
    Dim activeTabColor As String
    Dim inactiveTabColor As String
    
    activeTabColor = FormColorToHex(&H8000000F)
    inactiveTabColor = AddRGB(activeTabColor, -20, -20, -20)
    activeTabColor = LCase$(activeTabColor)
    inactiveTabColor = LCase$(inactiveTabColor)
    
    controlVarName = GenerateCtrlVarName(multiPageCtrl, prefix)
    
    For Each ctrl In multiPageCtrl.Pages
        tabVarName = GenerateCtrlVarName(ctrl, prefix)
        collPages.Add q & tabVarName & q
        collTabs.Add q & tabVarName & "-Tab" & q
    Next
    jsArrPages = "[" & JoinCollection(collPages, ", ") & "]"
    jsArrTabs = "[" & JoinCollection(collTabs, ", ") & "]"
    jsFuncName = GenerateMultiPageFuncName(multiPageCtrl, prefix)
    jsScrollFuncName = GenerateMultiPageScrollFuncName(multiPageCtrl, prefix)
    jsUpdateNavFuncName = prefix & multiPageCtrl.Name & "_UpdateTabNavBtns"
    jsUpdateNavFuncName = VBA.Replace$(jsUpdateNavFuncName, "-", "_")
    With collFuncStr
        .Add "function " & jsFuncName & "(index) {"
        .Add "  var pages = " & jsArrPages & ";"
        .Add "  var tabs = " & jsArrTabs & ";"
        .Add "  var wrapper = document.getElementById(" & q & controlVarName & "-TabScrollWrapper" & q & ");"
        .Add "  for (var i = 0; i < pages.length; i++) {"
        .Add "    var page = document.getElementById(pages[i]);"
        .Add "    var tab = document.getElementById(tabs[i]);"
        .Add "      if (i === index) {"
        .Add "        page.style.display = ""block"";"
        .Add "        tab.style.padding = ""2px 10px"";"
        .Add "        tab.style.backgroundColor = " & q & activeTabColor & q & ";"
        .Add "        if (wrapper && tab) {"
        .Add "          if (tab.offsetLeft < wrapper.scrollLeft) {"
        .Add "            wrapper.scrollLeft = tab.offsetLeft;"
        .Add "          } else if ((tab.offsetLeft + tab.offsetWidth) > (wrapper.scrollLeft + wrapper.clientWidth)) {"
        .Add "            wrapper.scrollLeft = tab.offsetLeft + tab.offsetWidth - wrapper.clientWidth;"
        .Add "          }"
        .Add "        }"
        .Add "      } else {"
        .Add "        page.style.display = ""none"";"
        .Add "        tab.style.padding = ""2px 8px"";"
        .Add "        tab.style.backgroundColor = " & q & inactiveTabColor & q & ";"
        .Add "      }"
        .Add "  }"
        .Add "  " & jsUpdateNavFuncName & "();"
        .Add "}"
        .Add ""
        .Add "function " & jsScrollFuncName & "(delta) {"
        .Add "  var wrapper = document.getElementById(" & q & controlVarName & "-TabScrollWrapper" & q & ");"
        .Add "  if (wrapper) {"
        .Add "    wrapper.scrollLeft += delta;"
        .Add "    " & jsUpdateNavFuncName & "();"
        .Add "  }"
        .Add "}"
        .Add ""
        .Add "function " & jsUpdateNavFuncName & "() {"
        .Add "  var wrapper = document.getElementById(" & q & controlVarName & "-TabScrollWrapper" & q & ");"
        .Add "  var btnLeft = document.getElementById(" & q & controlVarName & "-TabNavBtnLeft" & q & ");"
        .Add "  var btnRight = document.getElementById(" & q & controlVarName & "-TabNavBtnRight" & q & ");"
        .Add "  var triLeft = document.getElementById(" & q & controlVarName & "-TabNavBtnTriangleLeft" & q & ");"
        .Add "  var triRight = document.getElementById(" & q & controlVarName & "-TabNavBtnTriangleRight" & q & ");"
        .Add "  if (!wrapper || !btnLeft || !btnRight || !triLeft || !triRight) return;"
        ' In IE, scrollWidth/clientWidth return incorrect values when child elements are hidden,
        ' incorrectly causing maxScroll > 0 and showing the buttons.
        ' Use hasVisibleTab to check if any tabs are visible first.
        ' Additionally, wrap getElementsByClassName in a try-catch block to prevent errors in IE8
        ' and earlier, which do not support getElementsByClassName.
        .Add "  try{"
        .Add "    var tabs = wrapper.getElementsByClassName(" & q & controlVarName & "-TabPages" & q & ");"
        .Add "  }catch(e){"
        .Add "    var tabs = [];"
        .Add "  }"
        .Add "  var hasVisibleTab = false;"
        .Add "  for (var i = 0; i < tabs.length; i++) {"
        ' Check computed style using currentStyle for older IE (IE8-)
        ' and getComputedStyle for standard browsers (IE9+ / modern browsers).
        .Add "    var displayStyle = tabs[i].currentStyle"
        .Add "      ? tabs[i].currentStyle.display"
        .Add "      : window.getComputedStyle(tabs[i]).display;"
        .Add "    if (displayStyle !== ""none"") {"
        .Add "      hasVisibleTab = true;"
        .Add "      break;"
        .Add "    }"
        .Add "  }"
        .Add "  if (!hasVisibleTab) {"
        .Add "    btnLeft.style.display = ""none"";"
        .Add "    btnRight.style.display = ""none"";"
        .Add "    return;"
        .Add "  }"
        .Add "  var maxScroll = wrapper.scrollWidth - wrapper.clientWidth;"
        .Add "  if (maxScroll <= 0) {"
        .Add "    btnLeft.style.display = ""none"";"
        .Add "    btnRight.style.display = ""none"";"
        .Add "  } else {"
        .Add "    btnLeft.style.display = ""inline-block"";"
        .Add "    btnRight.style.display = ""inline-block"";"
        .Add "    if (wrapper.scrollLeft <= 0) {"
        .Add "      triLeft.style.borderBottomColor = ""#a0a0a0"";"
        .Add "    } else {"
        .Add "      triLeft.style.borderBottomColor = ""#000000"";"
        .Add "    }"
        .Add "    if (wrapper.scrollLeft >= maxScroll - 1) {"
        .Add "      triRight.style.borderBottomColor = ""#a0a0a0"";"
        .Add "    } else {"
        .Add "      triRight.style.borderBottomColor = ""#000000"";"
        .Add "    }"
        .Add "  }"
        .Add "}"
        .Add ""
        .Add "if (window.attachEvent) {"
        .Add "  window.attachEvent(""onload"", " & jsUpdateNavFuncName & ");"
        .Add "} else if (window.addEventListener) {"
        .Add "  window.addEventListener(""load"", " & jsUpdateNavFuncName & ");"
        .Add "}"
        .Add ""
    End With
    jsFuncStr = JoinCollection(collFuncStr, vbLf)
    GenerateTabPageSwitchingFunc = jsFuncStr
End Function


Private Function GetBorderSetting(ByVal ctrl As Object) As Collection
    Dim hexBorderColor As String
    Dim result As New Collection
    hexBorderColor = FormColorToHex(ctrl.BorderColor)
    hexBorderColor = LCase$(hexBorderColor)
    
    Select Case ctrl.BorderStyle
        Case fmBorderStyleSingle
            ' SpecialEffect is 0 if BorderStyle is 1
            result.Add "border: 1px solid " & hexBorderColor
        Case fmBorderStyleNone
            Select Case ctrl.SpecialEffect
                Case fmSpecialEffectFlat
                    result.Add "border: 0px solid " & hexBorderColor
                Case fmSpecialEffectRaised
                    result.Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
                Case fmSpecialEffectSunken
                    result.Add "box-shadow: 1px  1px 0 #ffffff, -1px -1px 0 #666666, 2px  2px 0 #eeeeee, -2px -2px 0 #444444"
                Case fmSpecialEffectEtched
                    result.Add "border: 1px solid #666666"
                    result.Add "outline: 1px solid #ffffff"
                    result.Add "outline-offset: -2px"
                Case fmSpecialEffectBump
                    result.Add "border: 1px solid #999999"
                    result.Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
            End Select
    End Select

    Set GetBorderSetting = result
End Function

Private Function GetTextAlignSetting(ByVal ctrl As Object) As Collection
    Dim result As New Collection
   
    If ContainsValue(Array("CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                result.Add "justify-content: flex-start"
            Case fmTextAlignCenter
                result.Add "justify-content: center"
            Case fmTextAlignRight
                result.Add "justify-content: flex-end"
            Case Else
                result.Add "justify-content: center"
        End Select
    Else
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                result.Add "text-align: left"
            Case fmTextAlignCenter
                result.Add "text-align: center"
            Case fmTextAlignRight
                result.Add "text-align: right"
            Case Else
                result.Add "text-align: center"
        End Select
    End If
    If ContainsValue(Array("CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        result.Add "align-items: center"
        result.Add "display: flex"
    End If
    
    Set GetTextAlignSetting = result
End Function

Private Function IsVerticalScrollBar(ByVal ctrl As Object) As Boolean
    ' Vertical -> True, Horizontal -> False
    Dim result As Boolean
    Select Case ctrl.orientation
        Case fmOrientationAuto
            If ctrl.width > ctrl.height Then
                result = False
            Else
                result = True
            End If
            
        Case fmOrientationVertical
            result = True
        Case fmOrientationHorizontal
            result = False
        Case Else
            result = True
    End Select
    IsVerticalScrollBar = result
End Function

Private Function IsListView(ByVal ctrl As Object) As Boolean
    ' Since the class name of the ListView may vary depending on the version, so use InStr to check it.
    ' e.g ListView/ListView2/ListView3/ListView4
    If InStr(TypeName(ctrl), "ListView") = 1 Then
        IsListView = True
    Else
        IsListView = False
    End If
End Function

Private Function IsTreeView(ByVal ctrl As Object) As Boolean
    ' Since the class name of the TreeView may vary depending on the version, so use InStr to check it.
    ' e.g TreeView/TreeView2/TreeView3/TreeView4
    If InStr(TypeName(ctrl), "TreeView") = 1 Then
        IsTreeView = True
    Else
        IsTreeView = False
    End If
End Function

Private Function GetControlCursorType(ByVal ctrl As Object) As String
    Dim cursorType As String
    
    Select Case ctrl.MousePointer
        Case fmMousePointerDefault
            cursorType = "auto"
            
        Case fmMousePointerArrow
            cursorType = "default"
            
        Case fmMousePointerCross
            cursorType = "crosshair"
            
        Case fmMousePointerIBeam
            cursorType = "text"
            
        Case fmMousePointerSizeNESW
            cursorType = "nesw-resize"
            
        Case fmMousePointerSizeNS
            cursorType = "ns-resize"
            
        Case fmMousePointerSizeNWSE
            cursorType = "nwse-resize"
            
        Case fmMousePointerSizeWE
            cursorType = "ew-resize"
            
        Case fmMousePointerUpArrow
            cursorType = "n-resize" ' closest match
            
        Case fmMousePointerHourGlass
            cursorType = "wait"
            
        Case fmMousePointerNoDrop
            cursorType = "not-allowed"
            
        Case fmMousePointerAppStarting
            cursorType = "progress" ' better CSS equivalent
            
        Case fmMousePointerHelp
            cursorType = "help"
            
        Case fmMousePointerSizeAll
            cursorType = "move"
            
        Case Else
            cursorType = "auto"
    End Select
    
    GetControlCursorType = cursorType
End Function


Private Function GenerateCssDataList(ByVal ctrl As Object, ByVal indentOffset As Long, ByVal prefix As String) As String
    ' Generates HTML <option> elements from a ListBox or ComboBox control
    ' Each item in the control is converted into:
    
    ' <datalist id="fruits">
    '   <option value="1">
    '   <option value="2">
    '   <option value="3">
    ' </datalist>
    
    Const q As String = """"
    Dim codeList As New Collection
    Dim item As Variant
    Dim i As Long: i = 0
    Dim result As String
    Dim controlVarName As String
    Dim itemValue As String
    controlVarName = GenerateCtrlVarName(ctrl, prefix)
    codeList.Add Space$(0 + indentOffset) & "<datalist id=" & q & controlVarName & "-items-value" & q & ">" & vbLf
    If ctrl.ListCount > 0 Then
        For Each item In ctrl.List
            i = i + 1
            itemValue = Convert2HTMLFormatText(item)
            codeList.Add Space$(2 + indentOffset) & "<option value=" & q & itemValue & q & ">" & vbLf
            If i = ctrl.ListCount Then Exit For
        Next item
    Else
        ' If no items exist, output an empty <option> element
        codeList.Add Space$(2 + indentOffset) & "<option value=" & q & " " & q & ">" & vbLf
    End If
    
    codeList.Add Space$(0 + indentOffset) & "</datalist>"
    
    result = JoinCollection(codeList, "")
    
    GenerateCssDataList = result
End Function

Private Function GetListViewItems(ByVal ctrl As Object, ByVal indentCount As Long) As String
    ' Retrieve the items of a ListView as a string in the format:
    ' <tbody>
    '
    '   <tr>
    '     <td>Item1-1</td>
    '     <td>Item1-2</td>
    '     <td>Item1-3</td>
    '   </tr>
    '
    '   <tr>
    '     <td>Item2-1</td>
    '     <td>Item2-2</td>
    '     <td>Item2-3</td>
    '   </tr>
    
    '   <tr>
    '     <td>Item3-1</td>
    '     <td>Item3-2</td>
    '     <td>Item3-3</td>
    '   </tr>
    '
    ' </tbody>
    Dim item As Object
    Dim i As Long
    Dim coll As Collection
    Dim resultColl As New Collection
    Dim result As String
    Dim indent As String
    Dim indent2 As String
    Dim indent3 As String
    indent = Space$(indentCount)
    indent2 = indent & Space$(2)
    indent3 = indent & Space$(4)
    For Each item In ctrl.ListItems
        Set coll = New Collection
        coll.Add "<td>" & Convert2HTMLFormatText(item.text) & "</td>"
        
        ' Because older versions of the ListView control do not support For Each for SubItems, use index-based access.
        For i = 1 To ctrl.ColumnHeaders.count - 1
            coll.Add "<td>" & Convert2HTMLFormatText(item.SubItems(i)) & "</td>"
        Next
        
        resultColl.Add "<tr>" & vbLf & indent3 & JoinCollection(coll, vbLf & indent3) & vbLf & indent2 & "</tr>"
        
    Next
    
    If resultColl.count > 0 Then
        result = indent & "<tbody>" & vbLf & vbLf & indent2 & JoinCollection(resultColl, vbLf & vbLf & indent2) & vbLf & vbLf & indent & "</tbody>"
    Else
        result = indent & "<tbody>" & vbLf & indent & "</tbody>"
    End If
    GetListViewItems = result
End Function

Private Function DefineListViewColumns(ByVal ctrl As Object, ByVal indentCount As Long) As String
    ' Generate code for the ListView headers.
    ' Example:
    ' <colgroup>
    '   <col style="width: 133px">
    '   <col style="width: 133px">
    '   <col style="width: 133px">
    ' </colgroup>
    '
    ' <thead>
    '   <tr>
    '     <th>Header1</th>
    '     <th>Header2</th>
    '     <th>Header3</th>
    '   </tr>
    ' </thead>
    Dim codeList As New Collection
    Dim objHeaders As Object
    Set objHeaders = ctrl.ColumnHeaders
    Dim controlVarName As String
    Dim i As Long
    Dim item As Variant
    Dim result As String
    Dim colName As String
    Dim headerText As String
    Dim colWidth As Long
    Dim anchor As String
    Const q As String = """"
    Dim indent As String
    Dim indent2 As String
    Dim indent3 As String
    indent = Space$(indentCount)
    indent2 = indent & Space$(2)
    indent3 = indent & Space$(4)
    
    i = 0
    codeList.Add indent & "<colgroup>" & vbLf
    For Each item In objHeaders
        i = i + 1
        colWidth = UserFormSizeToPixel(item.width)
        codeList.Add indent2 & "<col style=""width: " & colWidth & "px"">" & vbLf
    Next item
    codeList.Add indent & "</colgroup>" & vbLf
    codeList.Add vbLf
    codeList.Add indent & "<thead>" & vbLf
    codeList.Add indent2 & "<tr>" & vbLf
    i = 0
    For Each item In objHeaders
        i = i + 1
        headerText = Convert2HTMLFormatText(item.text)
        codeList.Add indent3 & "<th>" & headerText & "</th>" & vbLf
    Next item
    codeList.Add indent2 & "</tr>" & vbLf
    codeList.Add indent & "</thead>" & vbLf
    
    result = JoinCollection(codeList, "")
    
    DefineListViewColumns = result
End Function

Private Function GenerateTreeViewHtmlDfs(ByVal treeviewCtrl As Object, ByVal indentCount As Long) As String
    ' Retrieve the items of a TreeView as a string in the format:
    ' <details>
    '   <summary>Parent-Node-1</summary>
    '   <ul>
    '     <li>
    '       <details>
    '         <summary>Child-Node-1</summary>
    '         <ul>
    '           <li>Grandchild-Node-1</li>
    '         </ul>
    '       </details>
    '     </li>
    '     <li>Child-Node-2</li>
    '     <li>Child-Node-3</li>
    '   </ul>
    ' </details>
    
    ' Note: Variable resultColl is not used directly, but is used for debugging purposes.
    
    Dim resultColl As Collection
    Set resultColl = New Collection
    Dim codeList As New Collection
    Dim result As String
    
    Dim nd As Object
    Dim rootIndex As Long
    rootIndex = 1
    
    ' Start DFS from each root node
    For Each nd In treeviewCtrl.nodes
        If nd.Parent Is Nothing Then
            GenerateTreeViewHtmlDfsHelper nd, CStr(rootIndex), resultColl, codeList, indentCount, 1
            rootIndex = rootIndex + 1
        End If
    Next nd
    
    result = JoinCollection(codeList, "")
    
    GenerateTreeViewHtmlDfs = result
End Function


Private Sub GenerateTreeViewHtmlDfsHelper(ByVal node As Object, _
                       ByVal hierarchy As String, _
                       ByRef resultColl As Collection, _
                       ByRef codeList As Collection, _
                       ByVal indentCount As Long, _
                       ByVal currentDepth As Long)
    
    ' Note: Variables resultColl and hierarchy are not used directly, but are used for debugging purposes.
    Dim child As Object
    Dim childIndex As Long
    
    ' Add current node first (pre-order DFS)
    resultColl.Add VBA.Array(node, hierarchy)
    
    ' Get the first child node
    Set child = node.child
    childIndex = 1
    
    Dim nodeText As String
    Dim indent As String
    Dim hasChild As Boolean
    Dim isExpanded As Boolean
    Dim detailsTag As String
    
    indent = Space$(indentCount)
    
    nodeText = node.text
    nodeText = Convert2HTMLFormatText(nodeText)
    
    hasChild = Not child Is Nothing
    isExpanded = node.Expanded
    
    If isExpanded Then
        detailsTag = "<details open>"
    Else
        detailsTag = "<details>"
    End If
    
    If currentDepth = 1 Then
        
        ' The root node does not need a list item
        If hasChild Then
            
            codeList.Add indent & detailsTag & vbLf
            codeList.Add indent & Space$(2) & "<summary>" & nodeText & "</summary>" & vbLf
            codeList.Add indent & Space$(2) & "<ul>" & vbLf
            
            ' Process child nodes recursively
            Do While Not child Is Nothing
                
                GenerateTreeViewHtmlDfsHelper _
                    child, _
                    hierarchy & "-" & childIndex, _
                    resultColl, _
                    codeList, _
                    indentCount + 4, _
                    currentDepth + 1
                
                childIndex = childIndex + 1
                Set child = child.Next
            Loop
            
            codeList.Add indent & Space$(2) & "</ul>" & vbLf
            codeList.Add indent & "</details>" & vbLf
            
        Else
            
            ' Output a root node without a list item
            codeList.Add indent & nodeText & vbLf
            
        End If
        
    Else
        
        ' Child nodes are contained in list items
        If hasChild Then
            
            codeList.Add indent & "<li>" & vbLf
            codeList.Add indent & Space$(2) & detailsTag & vbLf
            codeList.Add indent & Space$(4) & "<summary>" & nodeText & "</summary>" & vbLf
            codeList.Add indent & Space$(4) & "<ul>" & vbLf
            
            ' Process child nodes recursively
            Do While Not child Is Nothing
                
                GenerateTreeViewHtmlDfsHelper _
                    child, _
                    hierarchy & "-" & childIndex, _
                    resultColl, _
                    codeList, _
                    indentCount + 6, _
                    currentDepth + 1
                
                childIndex = childIndex + 1
                Set child = child.Next
            Loop
            
            codeList.Add indent & Space$(4) & "</ul>" & vbLf
            codeList.Add indent & Space$(2) & "</details>" & vbLf
            codeList.Add indent & "</li>" & vbLf
            
        Else
            
            ' Output a leaf node on a single line
            codeList.Add indent & "<li>" & nodeText & "</li>" & vbLf
            
        End If
        
    End If

End Sub

Private Function HasScrollProperty(ByVal ctrl As Object) As Boolean
    ' Since the Scroll property does not exist in older versions of TreeView, use this function to check for the property beforehand.
    Dim temp As Variant
    On Error GoTo Exception
    temp = VBA.Array(ctrl.Scroll)
    HasScrollProperty = True
    On Error GoTo 0
    Exit Function
Exception:
    HasScrollProperty = False
End Function

Private Function GenerateCssSelectOptions(ByVal ctrl As Object, ByVal indentOffset As Long) As String
    ' Generates HTML <option> elements from a ListBox or ComboBox control
    ' Each item in the control is converted into:
    '
    '   <option value="1">1</option>
    '   <option value="2">2</option>
    '   <option value="3">3</option>
    '
    
    Const q As String = """"
    Dim codeList As New Collection
    Dim item As Variant
    Dim i As Long: i = 0
    Dim result As String
    Dim itemValue As String
    If ctrl.ListCount > 0 Then
        For Each item In ctrl.List
            i = i + 1
            itemValue = Convert2HTMLFormatText(item)
            codeList.Add Space$(2 + indentOffset) & "<option value=" & q & itemValue & q & ">" & itemValue & "</option>" & vbLf
            If i = ctrl.ListCount Then Exit For
        Next item
    Else
        ' If no items exist, output an empty <option> element
        codeList.Add Space$(2 + indentOffset) & "<option value=" & q & q & ">" & item & "</option>" & vbLf
    End If
    
    result = JoinCollection(codeList, "")
    
    GenerateCssSelectOptions = result
End Function

Private Function GetTextSizeFromCtrlFontSetting(ByVal ctrl As Object, ByVal targetText As String) As Variant()
    '------------------------------------------------------------------------------
    ' Returns the rendered text size (Width, Height) for a given text string
    ' using the same font settings as the specified control.
    ' Size is measured in points, not pixels.
    '
    ' Parameters:
    '   ctrl        - The reference control whose font settings will be used.
    '   targetText  - The text to measure. If empty, "i" is used to ensure a measurable size is returned.
    '                 (The letter "i" is one of the ASCII characters with the narrowest rendering width.)
    '
    ' Returns:
    '   Variant() Array
    '       (0) = Text width
    '       (1) = Text height
    '
    ' Notes:
    '   - A temporary hidden Label control is dynamically created on the parent
    '     UserForm to calculate the actual rendered text dimensions.
    '   - AutoSize is enabled so the Label automatically resizes to fit the text.
    '   - The temporary control is removed immediately after measurement.
    '
    ' Compatibility Note:
    '   In Excel 2013 and earlier, it was confirmed that enabling .AutoSize does not
    '   correctly update the .Width and .Height properties, which remain 0.
    '   To avoid returning invalid measurements, this function falls back to
    '   an estimated size calculation based on the font size and text length.
    '   This fallback is less accurate than actual rendered text measurement.
    '
    '------------------------------------------------------------------------------
    Dim rootForm As Object
    Dim tempLabel As Object
    Dim tempName As String
    Dim textWidthSize As Double
    Dim textHeightSize As Double
    ' Prevent zero-size measurement for empty strings.
    If targetText = "" Then targetText = "i"
    ' Generate a unique temporary control name.
    tempName = "TempLabel_" & VBA.Replace$(GenerateUUIDv4(), "-", "_")
    ' Get the parent UserForm from the specified control.
    Set rootForm = GetUserFormObjectFromCtrl(ctrl)
    ' Create a temporary Label control for text measurement.
    Set tempLabel = rootForm.Controls.Add("Forms.Label.1", tempName, True)
    ' Initialize control properties.
    tempLabel.height = 0
    tempLabel.width = 0
    tempLabel.caption = ""
    tempLabel.AutoSize = True
    tempLabel.WordWrap = False
    ' Optional debug background color.
    tempLabel.BackColor = &H80C0FF
    ' Copy font settings from the source control.
    tempLabel.Font.Name = ctrl.Font.Name
    tempLabel.Font.size = ctrl.Font.size
    tempLabel.Font.Bold = ctrl.Font.Bold
    tempLabel.Font.Italic = ctrl.Font.Italic
    tempLabel.Font.Underline = ctrl.Font.Underline
    tempLabel.Font.Strikethrough = ctrl.Font.Strikethrough
    ' Apply target text so AutoSize calculates the rendered dimensions.
    tempLabel.caption = targetText
    ' Read calculated size.
    textWidthSize = tempLabel.width
    textHeightSize = tempLabel.height
    
    ' In Excel 2013 and earlier, it was confirmed that the result of .AutoSize
    ' is not reflected in .Width/.Height and remains 0.
    ' As a fallback, the font size is used instead,
    ' although the measurement accuracy is reduced.
    If textWidthSize = 0 Then textWidthSize = ctrl.Font.size * Len(targetText)
    If textHeightSize = 0 Then textHeightSize = ctrl.Font.size
    
    ' The .Controls.Remove method does not accept a String argument; the argument must be of type Variant (String).
    ' example: tempLabel.Name or CVar(tempName)
    Call rootForm.Controls.Remove(tempLabel.Name)
    ' Release object reference.
    Set tempLabel = Nothing
    ' Return width and height as an array.
    GetTextSizeFromCtrlFontSetting = VBA.Array(textWidthSize, textHeightSize)
End Function

Private Function Convert2HTMLFormatText(ByVal text As String, Optional useBrTag As Boolean = True) As String
    ' Escape special characters in the string
    ' "&" should be replaced first
    text = VBA.Replace$(text, "&", "&amp;")
    text = VBA.Replace$(text, "<", "&lt;")
    text = VBA.Replace$(text, ">", "&gt;")
    text = VBA.Replace$(text, """", "&quot;")
    text = VBA.Replace$(text, "'", "&#39;") ' Use numeric character reference for compatibility ("&apos;" is not officially supported in HTML 4.)
    text = VBA.Replace$(text, " ", "&nbsp;")
    ' Convert VBA line breaks to HTML format
    ' vbCrLf should be replaced first
    text = VBA.Replace$(text, vbCrLf, vbLf)
    text = VBA.Replace$(text, vbCr, vbLf)
    If useBrTag Then
        text = VBA.Replace$(text, vbLf, "<br>") ' For div
    Else
        text = VBA.Replace$(text, vbLf, "&#10;") ' For textarea
    End If
    Convert2HTMLFormatText = text
End Function

Private Function AddRGB(ByVal hexColor As String, _
                       ByVal addR As Long, _
                       ByVal addG As Long, _
                       ByVal addB As Long) As String
    ' Example:
    ' AddRGB("#F0F0F0", -20, -20, -20) -> "#DCDCDC"
    ' AddRGB("#000000", -20, -20, -20) -> "#000000"
    ' AddRGB("#000000", 13, 14, 15) -> "#0D0E0F"
    ' AddRGB("#FEFE00", 15, 15, 15) -> "#FFFF0F"
    Dim r As Long, g As Long, b As Long

    ' Remove "#" if included
    hexColor = VBA.Replace$(hexColor, "#", "")

    ' Validate length
    If Len(hexColor) <> 6 Then
        AddRGB = "#000000"
        Exit Function
    End If

    ' Convert HEX -> RGB
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))

    ' Add values
    r = r + addR
    g = g + addG
    b = b + addB

    ' Clamp between 0 and 255
    If r < 0 Then r = 0
    If r > 255 Then r = 255

    If g < 0 Then g = 0
    If g > 255 Then g = 255

    If b < 0 Then b = 0
    If b > 255 Then b = 255

    ' Convert back to HEX
    AddRGB = "#" & _
             Right$("0" & Hex(r), 2) & _
             Right$("0" & Hex(g), 2) & _
             Right$("0" & Hex(b), 2)

End Function

Private Function FormColorToHex(ByVal clr As Long) As String
    ' Example:
    ' 16777215 -> "#FFFFFF"
    ' 0 -> "#000000"
    ' &H000000FF& (255) -> "#FF0000"
    ' &H00B4769E& (11826846) -> "#9E76B4"
    ' &H8000000F& (-2147483633) -> "#F0F0F0"(Windows XP[Luna Theme]/10/11), "#D4D0C8"(Windows 2000/XP[Classic Theme])
    Dim r As Long, g As Long, b As Long
    ' Convert a system color to its decimal color code when the parameter is a system color
    If 0 > clr Or clr >= 2147483648# Then
        clr = GetSysColor(clr And &HFF)
    End If
    ' Retrieve each component of the RGB color.
    r = clr And &HFF            ' Extract low-order 8 bits
    g = (clr \ &H100) And &HFF  ' Extract bits 8-15
    b = (clr \ &H10000) And &HFF ' Extract bits 16-23
    
    ' Convert the decimal RGB values to a #RRGGBB hex string and return it
    FormColorToHex = "#" & _
                     Right$("0" & Hex(r), 2) & _
                     Right$("0" & Hex(g), 2) & _
                     Right$("0" & Hex(b), 2)
End Function

Private Sub ConvertImageFormat(ByVal srcPath As String, ByVal dstPath As String, _
    Optional ByVal resize As Boolean = False, _
    Optional ByVal dstWidth As Long = -1, _
    Optional ByVal dstHeight As Long = -1, _
    Optional ByVal preserveAspectRatio As Boolean = False)
    
    '----------------------------------------------------------------------------------------------------
    ' Converts an image file to a specified output format and optionally resizes the image before saving.
    ' Uses GDI+ API without relying on WIA components (for compatibility with older versions of Windows).
    '
    ' Parameters:
    '   srcPath              - Full path of the source image file.
    '   dstPath              - Full path of the destination image file. Output format is
    '                          determined by the file extension (.png, .jpg, .jpeg, .bmp, .gif).
    '   resize               - If True, resizes the image before conversion.
    '                          Default: False.
    '   dstWidth             - Target width for resizing in pixels.
    '                          If -1, uses the original image width.
    '                          Default: -1.
    '   dstHeight            - Target height for resizing in pixels.
    '                          If -1, uses the original image height.
    '                          Default: -1.
    '   preserveAspectRatio  - If True, maintains the original aspect ratio during resizing.
    '                          Default: False.
    '
    '
    ' Supported Formats:
    '   PNG (.png)
    '   JPEG (.jpg, .jpeg)
    '   BMP (.bmp)
    '   GIF (.gif)
    '
    ' Errors:
    '   Raises Error if the destination file extension is not supported.
    '
    '----------------------------------------------------------------------------------------------------
    
    Const PixelFormat32bppARGB As Long = &H26200A
    Const InterpolationModeHighQualityBicubic As Long = 7
    
    ' Encoder CLSID Constants
    Const CLSID_BMP As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
    Const CLSID_JPG As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
    Const CLSID_GIF As String = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
    Const CLSID_PNG As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
    
    Dim fso As Object
    Dim ext As String
    Dim encoderCLSID As String
    
    ' Determine output format by file extension
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase$(fso.GetExtensionName(dstPath))
    
    Select Case ext
        Case "png"
            encoderCLSID = CLSID_PNG
        Case "jpg", "jpeg"
            encoderCLSID = CLSID_JPG
        Case "bmp"
            encoderCLSID = CLSID_BMP
        Case "gif"
            encoderCLSID = CLSID_GIF
        Case Else
            Err.Raise Number:=513, Description:="[ConvertImageFormat] [dstPath] Invalid Extension (." & ext & ")"
    End Select
    
    ' Initialize GDI+
    Dim gsi As GDIPlusStartupInput
    Dim token As LongPtr
    gsi.GdiPlusVersion = 1
    If GdiplusStartup(token, gsi) <> 0 Then
        Err.Raise Number:=514, Description:="[ConvertImageFormat] Failed to initialize GDI+"
    End If
    
    On Error GoTo CleanUp
    
    ' Load source image
    Dim hSrcImage As LongPtr
    If GdipLoadImageFromFile(StrPtr(srcPath), hSrcImage) <> 0 Then
        Err.Raise Number:=515, Description:="[ConvertImageFormat] Failed to load source image: " & srcPath
    End If
    
    ' Get original dimensions
    Dim origWidth As Long, origHeight As Long
    GdipGetImageWidth hSrcImage, origWidth
    GdipGetImageHeight hSrcImage, origHeight
    
    ' Determine target dimensions
    Dim targetW As Long, targetH As Long
    targetW = IIf(dstWidth = -1, origWidth, dstWidth)
    targetH = IIf(dstHeight = -1, origHeight, dstHeight)
    
    ' Calculate aspect-ratio adjusted dimensions if required
    Dim finalW As Long, finalH As Long
    If resize And preserveAspectRatio Then
        Dim scaleW As Double, scaleH As Double, scaleFactor As Double
        scaleW = CDbl(targetW) / CDbl(origWidth)
        scaleH = CDbl(targetH) / CDbl(origHeight)
        
        ' Calculate ratio based on min(scaleW, scaleH) when preserving aspect ratio
        If scaleW < scaleH Then
            scaleFactor = scaleW
        Else
            scaleFactor = scaleH
        End If
        
        finalW = CLng(origWidth * scaleFactor)
        finalH = CLng(origHeight * scaleFactor)
        If finalW < 1 Then finalW = 1
        If finalH < 1 Then finalH = 1
    ElseIf resize Then
        finalW = targetW
        finalH = targetH
    Else
        finalW = origWidth
        finalH = origHeight
    End If
    
    ' Create new bitmap and draw resized image if necessary or convert directly
    Dim hDstBitmap As LongPtr
    Dim hGraphics As LongPtr
    
    If resize Or (finalW <> origWidth Or finalH <> origHeight) Then
        GdipCreateBitmapFromScan0 finalW, finalH, 0, PixelFormat32bppARGB, 0, hDstBitmap
        GdipGetImageGraphicsContext hDstBitmap, hGraphics
        GdipSetInterpolationMode hGraphics, InterpolationModeHighQualityBicubic
        GdipDrawImageRectI hGraphics, hSrcImage, 0, 0, finalW, finalH
    Else
        hDstBitmap = hSrcImage
    End If
    
    ' Prepare Encoder CLSID
    Dim tCLSID As GUID
    CLSIDFromString StrPtr(encoderCLSID), tCLSID
    
    ' Save image using temporary file pattern
    Dim tmpPath As String
    tmpPath = dstPath & ".tmp"
    
    If fso.FileExists(tmpPath) Then fso.DeleteFile tmpPath
    
    Dim status As Long
    status = GdipSaveImageToFile(hDstBitmap, StrPtr(tmpPath), tCLSID, 0)
    
    ' Clean up GDI+ image resources prior to moving file
    If hGraphics <> 0 Then GdipDeleteGraphics hGraphics
    If hDstBitmap <> 0 And hDstBitmap <> hSrcImage Then GdipDisposeImage hDstBitmap
    If hSrcImage <> 0 Then GdipDisposeImage hSrcImage
    
    hGraphics = 0
    hDstBitmap = 0
    hSrcImage = 0
    
    If status <> 0 Then
        If fso.FileExists(tmpPath) Then fso.DeleteFile tmpPath
        Err.Raise Number:=516, Description:="[ConvertImageFormat] Failed to save output image."
    End If
    
    ' Overwrite target file
    If fso.FileExists(dstPath) Then fso.DeleteFile dstPath
    fso.MoveFile tmpPath, dstPath

CleanUp:
    ' Ensure GDI+ resources are freed
    If hGraphics <> 0 Then GdipDeleteGraphics hGraphics
    If hDstBitmap <> 0 And hDstBitmap <> hSrcImage Then GdipDisposeImage hDstBitmap
    If hSrcImage <> 0 Then GdipDisposeImage hSrcImage
    If token <> 0 Then GdiplusShutdown token
    
    If Err.Number <> 0 Then
        Dim errNum As Long, errDesc As String
        errNum = Err.Number
        errDesc = Err.Description
        Err.Raise errNum, Description:=errDesc
    End If
End Sub


Private Function FileToBase64(ByVal filePath As String) As String
    ' Converts a file to a Base64-encoded string.
    Dim stream As Object
    Dim bytes() As Byte
    Dim emptyBytes() As Byte
    emptyBytes = VBA.vbNullString ' Empty Byte Array
    
    ' Load file as binary
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' binary
    stream.Open
    stream.LoadFromFile filePath
    
    If stream.size > 0 Then
        bytes = stream.Read
    Else
        bytes = emptyBytes
    End If
    
    stream.Close
    Set stream = Nothing
    
    ' Convert binary to Base64
    FileToBase64 = BytesToBase64(bytes)

End Function

Private Function BytesToBase64(ByRef bytes() As Byte) As String
    ' Converts a byte array to a Base64 string.
    ' Implemented using the Windows API for high performance.
    ' The implementation supports byte arrays up to 1,610,612,733 bytes by design,
    ' but in practice, memory limitations may be reached with smaller arrays.
    Dim cch As Long
    Dim requiredCharCount As Long
    Dim buffer As String
    Dim dwFlags As Long
    Dim charCountOffset As Long
    Dim dataSize As Double
    Const STRING_FUNCTION_MAXIMUM_LIMIT As Long = 1073741823
    Const STRING_VARIABLE_MAXIMUM_LIMIT As Long = 2147483645
    Const MAXIMUM_DATA_SIZE As Long = 1610612733 ' The largest value that ensures the Base64-encoded size does not exceed the String type's maximum length of 2,147,483,645 characters.
    Const CRYPT_STRING_BASE64 As Long = &H1
    Const CRYPT_STRING_NOCRLF As Long = &H40000000

    dwFlags = CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF
    dataSize = 0
    On Error Resume Next
    dataSize = UBound(bytes) - LBound(bytes) + 1
    On Error GoTo 0
    
    If dataSize = 0 Then
        BytesToBase64 = ""
        Exit Function
    End If
    
    If dataSize > MAXIMUM_DATA_SIZE Then
        Err.Raise Number:=9000, Description:="Error: Data size (" & dataSize & ") exceeds the maximum limit (" & MAXIMUM_DATA_SIZE & ")."
    End If
    
    ' Get required output length (includes terminating null)
    If CryptBinaryToStringW( _
            VarPtr(bytes(LBound(bytes))), _
            UBound(bytes) - LBound(bytes) + 1, _
            dwFlags, _
            0, _
            cch) = 0 Then
        Err.Raise vbObjectError + 1, , "CryptBinaryToStringW failed."
    End If
    
    If 0 >= cch Then
        Err.Raise Number:=9500, Description:="Error: Overflow occurred because the data size is too large"
    End If
    
    ' CryptBinaryToStringW returns the required character count including
    ' the terminating null character. VBA strings (BSTR) store their length
    ' separately, so the null terminator is not part of the string length.
    ' Allocate one fewer character to avoid including the terminating null
    ' in the resulting VBA string.
    requiredCharCount = cch - 1
    
    If requiredCharCount > STRING_VARIABLE_MAXIMUM_LIMIT Then
        Err.Raise Number:=10000, Description:="Error: Data size (" & requiredCharCount & ") exceeds the maximum limit (" & STRING_VARIABLE_MAXIMUM_LIMIT & ")."
    End If
    
    ' Allocate output buffer
    'The String$ function is limited to 1,073,741,823 characters per call,
    'although a String variable can hold up to 2,147,483,645 characters.
    'Allocate larger buffers by concatenating multiple String$ calls.
    If requiredCharCount > STRING_FUNCTION_MAXIMUM_LIMIT Then
        charCountOffset = requiredCharCount - STRING_FUNCTION_MAXIMUM_LIMIT
        buffer = String$(requiredCharCount - charCountOffset, vbNullChar) & String$(charCountOffset, vbNullChar)
    Else
        buffer = String$(requiredCharCount, vbNullChar)
    End If
    

    ' Convert
    If CryptBinaryToStringW( _
            VarPtr(bytes(LBound(bytes))), _
            UBound(bytes) - LBound(bytes) + 1, _
            dwFlags, _
            StrPtr(buffer), _
            cch) = 0 Then
        Err.Raise vbObjectError + 2, , "CryptBinaryToStringW failed."
    End If
    
    ' Windows XP does not support CRYPT_STRING_NOCRLF with CryptBinaryToStringW, so the generated Base64 string may contain line breaks.
    ' Therefore, if the string contains CRLF, replace them with an empty string to remove them.
    If InStr(buffer, VBA.vbCrLf) > 0 Then
        buffer = VBA.Replace$(buffer, VBA.vbCrLf, "")
    End If
    
    BytesToBase64 = buffer
End Function

Private Function ContainsValue(ByVal itemList As Variant, ByVal value As Variant) As Boolean
    ' Check if a specific value exists in Array/Collection/Dictionary
    ' itemList - Array/Collection/Dictionary to search
    ' value - value to check
    ' Performs strict type comparison for non-numeric values
    ' Nested arrays are not supported. Objects are compared by reference
    ' Dependency: IsStrictlyEqual(helper function)
    Dim item As Variant
    Dim temp As Variant
    If LCase$(TypeName(itemList)) = "dictionary" Then
        itemList = itemList.items
    End If
    If IsArray(itemList) Then
        On Error GoTo Finally
        ' Uninitialized Array -> False
        temp = LBound(itemList)
        On Error GoTo 0
    End If
    For Each item In itemList
    
        If IsStrictlyEqual(item, value) Then
            ContainsValue = True
            Exit Function
        End If
    Next
Finally:
    ContainsValue = False
    
End Function

Private Function IsStrictlyEqual(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean
    ' Performs a strict equality comparison including data types.
    ' Numeric types (Integer, Long, Double, etc.) are treated as compatible.
    ' Boolean and Date types are NOT treated as numeric.
    Dim t1 As VbVarType, t2 As VbVarType
    t1 = VarType(value1)
    t2 = VarType(value2)
    
    ' Returns True if objects point to the same reference.
    ' Objects are evaluated first to prevent false matches (e.g., Empty vs empty Cells).
    ' (Also applies to variables holding both objects and other data types)
    If IsObject(value1) Or IsObject(value2) Then
        If IsObject(value1) And IsObject(value2) Then
            IsStrictlyEqual = (value1 Is value2)
        End If
        Exit Function
    End If
    
    ' Null / Empty
    If IsNull(value1) Or IsNull(value2) Then
        IsStrictlyEqual = (IsNull(value1) And IsNull(value2))
        Exit Function
    ElseIf IsEmpty(value1) Or IsEmpty(value2) Then
        IsStrictlyEqual = (IsEmpty(value1) And IsEmpty(value2))
        Exit Function
    End If
    
    
    ' Arrays are not supported (Extend if necessary).
    If IsArray(value1) Or IsArray(value2) Then
        IsStrictlyEqual = False
        Exit Function
    End If
    
    ' Error values
    If t1 = vbError Or t2 = vbError Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' String, Date, Boolean
    If (t1 = vbString Or t2 = vbString) Or (t1 = vbDate Or t2 = vbDate) Or (t1 = vbBoolean Or t2 = vbBoolean) Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' Other data types (e.g., Numeric)
    On Error Resume Next
    IsStrictlyEqual = (value1 = value2)
    Exit Function
    On Error GoTo 0
    IsStrictlyEqual = False
End Function


Private Function UserFormSizeToPixel(ByVal ufSize As Double) As Long
    ' Function to convert the size of a UserForm or control to pixels
    ' Excel VBA UserForm dimensions are internally handled as
    ' DPI-independent logical points based on a fixed 96 DPI system.
    ' Therefore, point-to-pixel conversion can be calculated as:
    '     pixel = point * (96 / 72)
    ' and works consistently regardless of the monitor DPI setting.
    Dim pixelSize As Long
    pixelSize = Round(ufSize * (96 / 72))
    UserFormSizeToPixel = pixelSize
End Function

Private Function GenerateUUIDv4() As String
    Dim i As Long
    Dim b(15) As Byte
    Dim s As String
    Dim hexStr As String
    
    ' Initialize random number generator
    Randomize
    
    ' Generate 16 bytes of random values
    For i = 0 To 15
        b(i) = Int(Rnd() * 256)
    Next i
    
    ' Set version (4) (set bits 7-4 to 0100)
    b(6) = (b(6) And &HF) Or &H40
    
    ' Set variant (10xx)
    b(8) = (b(8) And &H3F) Or &H80
    
    ' Convert the 16 bytes to a string (with hyphen format)
    hexStr = ""
    For i = 0 To 15
        hexStr = hexStr & Right$("0" & Hex(b(i)), 2)
        Select Case i
            Case 3, 5, 7, 9
                hexStr = hexStr & "-"
        End Select
    Next i
    
    GenerateUUIDv4 = LCase$(hexStr)
End Function

Private Sub SaveUtf8TextNoBom(ByVal filePath As String, ByVal textData As String)
    ' Save the specified string as UTF-8 without BOM
    Dim textStream As Object
    Dim binaryStream As Object
    Dim bytes() As Byte
    Dim emptyBytes() As Byte
    emptyBytes = VBA.vbNullString ' Empty Byte Array
    
    ' Normalize line endings
    textData = VBA.Replace$(textData, vbCrLf, vbLf)
    textData = VBA.Replace$(textData, vbCr, vbLf)
    textData = VBA.Replace$(textData, vbLf, vbNewLine)
    
    ' Convert to UTF-8 and remove BOM
    Set textStream = CreateObject("ADODB.Stream")
    With textStream
        .Type = 2 ' Text mode
        .Charset = "utf-8"
        .Open
        .WriteText textData
        .position = 0
        .Type = 1 ' Switch to binary mode
        
        If .size > 0 Then
            bytes = .Read
        Else
            bytes = emptyBytes
        End If
        
        .Close
    End With
    Set textStream = Nothing
    
    ' Remove BOM if present
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            bytes = MidB(bytes, 4) ' Remove BOM (EF BB BF)
        End If
    End If
    
    ' Save file in binary mode
    Set binaryStream = CreateObject("ADODB.Stream")
    With binaryStream
        .Type = 1
        .Open
        .Write bytes
        .SaveToFile filePath, 2
        .Close
    End With
    Set binaryStream = Nothing
End Sub


Private Function GenerateUnsupportedControlMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnsupportedControlMessage = "Control type " & q & TypeName(ctrl) & q & " is not supported."
End Function

Private Function GenerateUnavailableNameMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnavailableNameMessage = "Object Name " & q & ctrl.Name & q & " is not available." & vbLf & "Please use a different name instead."
End Function

Private Function GetFormControlDepth(ByVal ctrl As Object) As Long
    ' Get the hierarchy depth of the control
    Dim depth As Long
    Dim temp As Variant
    depth = 0
    Set temp = ctrl
    Do While True
        If depth Mod 10 = 0 Then DoEvents
        On Error GoTo Finally
        Set temp = temp.Parent
        depth = depth + 1
        On Error GoTo 0
    Loop
Finally:
    
    If Err.Number <> 438 Then
        Err.Raise Number:=Err.Number
    End If
    
    GetFormControlDepth = depth
    
End Function

Private Function SortFormControlsByDepth(ByVal frmControls As Variant) As Collection
    ' Sort the list of UserForm controls in ascending order of hierarchy depth
    Dim tempColl As Collection
    Set tempColl = New Collection
    Dim sortedColl As Collection
    Set sortedColl = New Collection
    Dim ctrl As Variant
    Dim tempArray() As Variant
    Dim depth As Long
    Dim item As Variant
    For Each ctrl In frmControls
        depth = GetFormControlDepth(ctrl)
        tempColl.Add VBA.Array(depth, ctrl)
    Next ctrl
    If tempColl.count > 0 Then
        tempArray = CollectionToArray(tempColl)
        Call InsertionSortJaggedArray(tempArray, reverse:=False)
        For Each item In tempArray
            sortedColl.Add item(1)
        Next item
    End If
    Set SortFormControlsByDepth = sortedColl
End Function


Private Function CollectionToArray(ByVal coll As Collection, Optional ByVal isStartIdx1 As Boolean = False) As Variant()
    ' Convert a Collection to an array
    ' If isStartIdx1 is True, create an array starting from index 1 (to match Collection numbering)
    Dim arr() As Variant
    Dim item As Variant
    Dim idx As Long
    If coll.count > 0 Then
        If isStartIdx1 Then
            ReDim arr(1 To coll.count)
        Else
            ReDim arr(0 To coll.count - 1)
        End If
        idx = LBound(arr)
        For Each item In coll
            ' Use "Set" when assigning objects.
            If IsObject(item) Then
                Set arr(idx) = item
            Else
                arr(idx) = item
            End If
            idx = idx + 1
        Next
    Else
        arr = VBA.Array()
    End If
    CollectionToArray = arr
End Function


Private Sub InsertionSortJaggedArray(ByRef arr As Variant, _
    Optional ByVal reverse As Boolean = False, _
    Optional ByVal strSort As Boolean = False, _
    Optional ByVal ignoreCase As Boolean = True)
    
    ' Sorts a jagged array using the Insertion Sort algorithm based on the first element of each nested array.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[1, "A"], [2, "C"], [3, "B"]]
    '   Does not affect the relative order of items with the same numeric value
    '   e.g., [[3, "C"], [3, "A"], [1, "A"], [3, "B"]] -> [[1, "A"], [3, "C"], [3, "A"], [3, "B"]]
    ' reverse: Set to True for descending order.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[3, "B"], [2, "C"], [1, "A"]]
    ' strSort: Set to True for string-based comparison, False for numeric comparison.
    ' ignoreCase: Valid only when strSort is True. Set to True to perform case-insensitive comparison.
    ' Dependency: DynamicCompare
    If Not IsArray(arr) Then Err.Raise Number:=13
    Dim minIndex As Long
    Dim maxIndex As Long
    Dim idxToRef1 As Long
    Dim idxToRef2 As Long
    Dim op As String
    
    If reverse Then
        op = "<"
    Else
        op = ">"
    End If
    
    minIndex = LBound(arr)
    maxIndex = UBound(arr)
    Dim i As Long, j As Long
    Dim swap As Variant
    For i = minIndex + 1 To maxIndex
        swap = arr(i)
        For j = i - 1 To minIndex Step -1
            idxToRef1 = LBound(arr(j))
            idxToRef2 = LBound(swap)
            If DynamicCompare(arr(j)(idxToRef1), swap(idxToRef2), op, strSort, ignoreCase) Then
                arr(j + 1) = arr(j)
            Else
                Exit For
            End If
        Next
        arr(j + 1) = swap
    Next
End Sub


Private Function DynamicCompare(ByVal a As Variant, ByVal b As Variant, ByVal op As String, _
    Optional ByVal shouldStrComp As Boolean = False, Optional ByVal ignoreCase As Boolean = True) As Boolean
    ' Performs dynamic comparison using a string representation of an operator.
    ' a, b: Values to compare.
    ' op: Comparison operator as a string (">", ">=", "<", "<=", "=", "<>").
    ' shouldStrComp: Set to True for string comparison mode, False for numeric/default comparison.
    ' ignoreCase: Valid only when shouldStrComp is True. Set to True to ignore case sensitivity.
    Dim result As Boolean
    Dim compareMode As VbCompareMethod
    
    If shouldStrComp Then
        If ignoreCase Then
            compareMode = vbTextCompare
        Else
            compareMode = vbBinaryCompare
        End If
        
        Select Case op
            Case ">"
                result = StrComp(a, b, compareMode) > 0
            Case ">="
                result = StrComp(a, b, compareMode) >= 0
            Case "<"
                result = StrComp(a, b, compareMode) < 0
            Case "<="
                result = StrComp(a, b, compareMode) <= 0
            Case "="
                result = StrComp(a, b, compareMode) = 0
            Case "<>"
                result = StrComp(a, b, compareMode) <> 0
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    Else
        Select Case op
            Case ">"
                result = (a > b)
            Case ">="
                result = (a >= b)
            Case "<"
                result = (a < b)
            Case "<="
                result = (a <= b)
            Case "="
                result = (a = b)
            Case "<>"
                result = (a <> b)
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    End If
    DynamicCompare = result
End Function

Private Function CollContainsKey(ByVal coll As Collection, ByVal strKey As String) As Boolean
    ' Check if a specific key exists in the Collection
    CollContainsKey = False
    If coll Is Nothing Then Exit Function
    If coll.count = 0 Then Exit Function
     
    On Error GoTo Exception
    Call coll.item(strKey)
    On Error GoTo 0
    CollContainsKey = True
    
    Exit Function
Exception:
    CollContainsKey = False
    Exit Function
End Function


Private Sub ExtendCollection(ByRef originalColl As Variant, ByVal additionalColl As Variant)
    ' Merges two Collections or ArrayLists (modifies the original object in-place).
    ' Keys from originalColl are preserved, but keys from additionalColl will be lost.
    ' Compatible with any object that has an .Add method (excluding Dictionaries).
    
    ' Parameters:
    '   originalColl   : The target collection to be extended.
    '   additionalColl : The collection containing items to add.
    Dim item As Variant
    For Each item In additionalColl
        originalColl.Add item
    Next item
End Sub

Private Function JoinCollection(ByVal coll As Collection, Optional ByVal delimiter As String = "") As String
    Dim arr() As Variant
    Dim result As String
    arr = CollectionToArray(coll)
    result = Join(arr, delimiter)
    JoinCollection = result
End Function


Private Function AdjustIndent(ByVal text As String, ByVal indentSize As Long) As String
    '------------------------------------------------------------
    ' Function: AdjustIndent
    '
    ' Description:
    '   Adjusts the indentation of each line in a given text.
    '   The text may contain multiple lines separated by line breaks.
    '
    ' Specification:
    '   - Normalizes all line breaks to vbLf before processing.
    '   - If indentSize > 0:
    '       Adds the specified number of spaces to the beginning of each line.
    '       HOWEVER, if a line is empty (i.e., contains no characters),
    '       no spaces are added to that line.
    '   - If indentSize < 0:
    '       Removes the specified number of spaces from the beginning of each line.
    '       If a line has fewer leading spaces than the absolute value of indentSize,
    '       all leading spaces are removed.
    '   - If indentSize = 0:
    '       Returns the text unchanged (except for normalized line breaks).
    '
    ' Parameters:
    '   text (String)       : Input text (may include line breaks).
    '   indentSize (Long)   : Number of spaces to add (positive) or remove (negative).
    '
    ' Returns:
    '   String              : Text with adjusted indentation.
    '------------------------------------------------------------
    Dim lines() As String
    Dim i As Long
    Dim spaces As String
    Dim removeCount As Long
    Dim leadingSpaces As Long
    
    ' Normalize line breaks
    text = VBA.Replace$(text, vbCrLf, vbLf)
    text = VBA.Replace$(text, vbCr, vbLf)
    
    ' Split into lines
    lines = Split(text, vbLf)
    
    If indentSize > 0 Then
        ' Add spaces (skip empty lines)
        spaces = String(indentSize, " ")
        For i = LBound(lines) To UBound(lines)
            If Len(lines(i)) > 0 Then
                lines(i) = spaces & lines(i)
            End If
        Next i
        
    ElseIf indentSize < 0 Then
        ' Remove spaces
        removeCount = -indentSize
        
        For i = LBound(lines) To UBound(lines)
            leadingSpaces = 0
            
            ' Count leading spaces
            Do While leadingSpaces < Len(lines(i)) _
                And Mid(lines(i), leadingSpaces + 1, 1) = " "
                leadingSpaces = leadingSpaces + 1
            Loop
            
            ' Remove spaces safely
            If leadingSpaces >= removeCount Then
                lines(i) = Mid(lines(i), removeCount + 1)
            Else
                lines(i) = Mid(lines(i), leadingSpaces + 1)
            End If
        Next i
    End If
    
    ' Join lines back
    AdjustIndent = Join(lines, vbLf)
End Function


' ========================================================================================
' VBA-UserForm-Z-Order-Tracer v1.0.6
' https://github.com/ZeeZeX/VBA-UserForm-Z-Order-Tracer
' Copyright (c) 2026 ZeeZeX
' License: MIT
'
' Based originally on tarboh's VBA-UserForm-Z-Order-Tool with extensive changes/additions
' https://github.com/tarboh/VBA-UserForm-Z-Order-Tool
' Copyright (c) 2026 tarboh
' License: MIT
'
' https://opensource.org/licenses/MIT
' ========================================================================================


Private Function GetAllCtrlsInZOrder(ByVal objUserForm As MSForms.UserForm, Optional ByVal zOrderType As String = "Internal", Optional bfsSort As Boolean = False) As Collection
    ''
    ' Traverses and retrieves all controls in a UserForm in Z-Order using Depth-First Search (DFS).
    '
    ' @param objUserForm [In] The target UserForm object (MsForms.UserForm) to scan.
    ' @param zOrderType  [In] Optional mode to determine how Z-Order is evaluated for Page controls inside a MultiPage.
    '                         Default is "Internal".
    '                           - "Internal": Prioritizes internal Z-Order retrieval for Page controls inside a MultiPage.
    '                                         Controls are retrieved in addition order, even if the visual page order has been moved.
    '                           - "Visual":   Prioritizes visual tab order, assigning Z-Order sequentially starting from the left tab.
    '                                         Reflects any reordering of pages.
    '                           - "VisualKeepZ": Retrieves Z-Order values identical to "Internal", but sorts the resulting collection
    '                                            in the same order as "Visual".
    '
    ' @param bfsSort     [In] Optional flag to sort the returned collection in breadth-first order (BFS).
    '                             If True, controls are ordered by nesting depth, from shallowest to deepest,
    '                             while preserving their original order within the same depth.
    '                             Default is False (depth-first order).
    '
    ' @return Collection A Collection containing Variant arrays for each control found.
    '                     Each array element consists of:
    '                       (0): Control object (Object)
    '                       (1): Control name (String)
    '                       (2): Z-Order index within its immediate container (Long)
    '                       (3): Hierarchical Z-Order string (e.g., "1-2-1") (String)
    '                       (4): Nesting depth level starting from 1 (Long)
    '
    '                     Items in the returned collection can be retrieved by key using `Hex(ObjPtr(ctrl))`.
    '                     Note: Variables holding controls should be declared as `Object` or `Variant` instead of
    '                     `MSForms.Control` or specific control types, as assigning to `MSForms.Control` may
    '                     cause `ObjPtr` to return a different pointer address. (this pointer mismatch has been
    '                     specifically confirmed with Page controls within a MultiPage)
    '
    ' @raises 515 Thrown if the provided object is not the top-level (root) UserForm.
    ' @raises 516 Thrown if an invalid or unsupported zOrderType string is specified.
    ''
    Dim result As Collection
    Dim ctrlsInternalOrder As Collection
    Dim ctrlsVisialOrder As Collection
    Dim item As Variant
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim key As String
    Dim tempColl As New Collection
    Dim tempColl2 As Collection
    Dim maxDepth As Long
    Dim minDepth As Long
    Dim depth As Long
    Dim i As Long
    Const q = """"
    Dim errMsg As String
    Set result = New Collection
    ' Raise an error if the passed object is not the root UserForm.
    If Not objUserForm Is GetUserFormObjectFromCtrl(objUserForm) Then
        Err.Raise Number:=515, Description:="Invalid UserForm"
    End If
    
    Select Case UCase$(Trim$(zOrderType))
    
        Case "INTERNAL"
            TraceZOrderRecursive objUserForm, 1, "", result, useVisualOrder:=False
            
        Case "VISUAL"
            TraceZOrderRecursive objUserForm, 1, "", result, useVisualOrder:=True
            
        Case "VISUALKEEPZ"
            Set ctrlsInternalOrder = New Collection
            Set ctrlsVisialOrder = New Collection
            TraceZOrderRecursive objUserForm, 1, "", ctrlsInternalOrder, useVisualOrder:=False
            TraceZOrderRecursive objUserForm, 1, "", ctrlsVisialOrder, useVisualOrder:=True
            
            For Each item In ctrlsVisialOrder
                key = Hex(ObjPtr(item(0)))
                arr = item
                arr2 = ctrlsInternalOrder(key)
                arr(2) = arr2(2)
                arr(3) = arr2(3)
                result.Add arr, key
            Next item
            
        Case Else
            errMsg = "[zOrderType] Invalid value: " & q & zOrderType & q & vbLf & "Supported values are " & q & Join(Array("Internal", "Visual", "VisualkeepZ"), q & ", " & q) & q
            Err.Raise Number:=516, Description:=errMsg
    End Select
    
    If bfsSort Then
        maxDepth = 0
        minDepth = 2147483647
        For Each item In result
            depth = item(4)
            If depth > maxDepth Then
                maxDepth = depth
            End If
            
            If minDepth > depth Then
                minDepth = depth
            End If
            
            key = "key" & CStr(depth)
            On Error Resume Next
            tempColl.Add New Collection, key
            On Error GoTo 0
            tempColl(key).Add item
        Next
        Set result = New Collection
        For i = minDepth To maxDepth
            key = "key" & CStr(i)
            Set tempColl2 = Nothing
            On Error Resume Next
            Set tempColl2 = tempColl(key)
            On Error GoTo 0
            If Not tempColl2 Is Nothing Then
                For Each item In tempColl2
                    result.Add item, Hex(ObjPtr(item(0)))
                Next
            End If
        Next i
    End If
    
    Set GetAllCtrlsInZOrder = result
    
End Function

Private Function GetCtrlZOrderIndex(ByVal ctrl As Object, Optional ByVal zOrderType As String = "Internal") As Long
    ' Retrieves the 1-based Z-Order index of a specific control within its immediate parent container.
    '
    ' @param ctrl [In] The target control or UserForm object to locate.
    ' @param zOrderType  [In] Optional mode to determine how Z-Order is evaluated for Page controls inside a MultiPage.
    '                         Default is "Internal".
    '                           - "Internal": Prioritizes internal Z-Order retrieval for Page controls inside a MultiPage.
    '                                         Controls are retrieved in addition order, even if the visual page order has been moved.
    '                           - "Visual":   Prioritizes visual tab order, assigning Z-Order sequentially starting from the left tab.
    '                                         Reflects any reordering of pages.
    ' @return Long The 1-based Z-Order index of the control within its parent container.
    '              Returns 0 if the provided object is the top-level UserForm itself.
    '              Returns -1 if the control is not found.
    '
    ' @raises 516 Thrown if an invalid or unsupported zOrderType string is specified.
    Dim result As Long
    Dim key As String
    result = -1
    Dim ctrls As Collection
    Dim item As Variant
    Dim root As Object
    Const q = """"
    Dim errMsg As String
    
    If UCase$(Trim$(zOrderType)) <> "INTERNAL" And UCase$(Trim$(zOrderType)) <> "VISUAL" Then
        errMsg = "[zOrderType] Invalid value: " & q & zOrderType & q & vbLf & "Supported values are " & q & Join(Array("Internal", "Visual"), q & ", " & q) & q
        Err.Raise Number:=516, Description:=errMsg
    End If
    
    Set root = GetUserFormObjectFromCtrl(ctrl)
    
    If ctrl Is root Then
        GetCtrlZOrderIndex = 0
        Exit Function
    End If
    
    Set ctrls = GetAllCtrlsInZOrder(root, zOrderType)
    key = Hex(ObjPtr(ctrl))
    On Error Resume Next
    result = ctrls(key)(2)
    On Error GoTo 0
    
    ' Fall back to linear search using the `Is` operator if key retrieval fails.
    ' This fallback is necessary because Page controls (inside a MultiPage) may yield different
    ' ObjPtr values depending on whether they are typed as Object or MSForms.Control/Page.
    If result = -1 Then
        For Each item In ctrls
            If item(0) Is ctrl Then
                result = item(2)
                Exit For
            End If
        Next item
    End If
    GetCtrlZOrderIndex = result
End Function

Private Function GetCtrlZOrderHierarchy(ByVal ctrl As Object, Optional ByVal zOrderType As String = "Internal") As String
    ' Retrieves the hierarchical Z-Order path string for a target control within a UserForm.
    '
    ' @param ctrl [In] The target control or UserForm object to query.
    ' @param zOrderType  [In] Optional mode to determine how Z-Order is evaluated for Page controls inside a MultiPage.
    '                         Default is "Internal".
    '                           - "Internal": Prioritizes internal Z-Order retrieval for Page controls inside a MultiPage.
    '                                         Controls are retrieved in addition order, even if the visual page order has been moved.
    '                           - "Visual":   Prioritizes visual tab order, assigning Z-Order sequentially starting from the left tab.
    '                                         Reflects any reordering of pages.
    ' @return String A hyphen-delimited string representing the full Z-Order path from top-level container
    '                down to the target control (e.g., "1-3-2").
    '                Returns "0" if the provided object is the top-level UserForm itself.
    '                Returns an empty string ("") if the control is not found.
    '
    ' @raises 516 Thrown if an invalid or unsupported zOrderType string is specified.
    Dim result As String
    Dim key As String
    result = ""
    Dim ctrls As Collection
    Dim item As Variant
    Dim root As Object
    Const q = """"
    Dim errMsg As String
    
    If UCase$(Trim$(zOrderType)) <> "INTERNAL" And UCase$(Trim$(zOrderType)) <> "VISUAL" Then
        errMsg = "[zOrderType] Invalid value: " & q & zOrderType & q & vbLf & "Supported values are " & q & Join(Array("Internal", "Visual"), q & ", " & q) & q
        Err.Raise Number:=516, Description:=errMsg
    End If
    
    Set root = GetUserFormObjectFromCtrl(ctrl)
    
    If ctrl Is root Then
        GetCtrlZOrderHierarchy = "0"
        Exit Function
    End If
    
    Set ctrls = GetAllCtrlsInZOrder(root, zOrderType)
    key = Hex(ObjPtr(ctrl))
    On Error Resume Next
    result = ctrls(key)(3)
    On Error GoTo 0
    
    ' Fall back to linear search using the `Is` operator if key retrieval fails.
    ' This fallback is necessary because Page controls (inside a MultiPage) may yield different
    ' ObjPtr values depending on whether they are typed as Object or MSForms.Control/Page.
    If result = "" Then
        For Each item In ctrls
            If item(0) Is ctrl Then
                result = item(3)
                Exit For
            End If
        Next item
    End If
    GetCtrlZOrderHierarchy = result
End Function

Private Function GetDirectChildCtrls(ByVal parentCtrl As Object) As Collection
    Dim results As Collection
    Set results = New Collection
    Dim ctrl As Object
    Dim root As Object

    Set root = GetUserFormObjectFromCtrl(parentCtrl)

    If TypeName(parentCtrl) = "MultiPage" Then
        For Each ctrl In parentCtrl.Pages
            results.Add ctrl
        Next
    Else
        For Each ctrl In root.Controls
            If parentCtrl Is ctrl.Parent Then
                results.Add ctrl
            End If
        Next ctrl
    End If
    Set GetDirectChildCtrls = results
End Function

'Private Function GetUserFormObjectFromCtrl(ByVal ctrl As Object) As Object
'    ' Get the ancestor (UserForm) of the control.
'
'    Dim root As Object
'
'    If ctrl Is Nothing Then
'        Err.Raise 13
'    End If
'
'    Set root = ctrl
'    ' Loop to get root(UserForm) object
'    On Error GoTo Finally:
'    Do While True
'        Set root = root.Parent
'    Loop
'    On Error GoTo 0
'
'Finally:
'    Set GetUserFormObjectFromCtrl = root
'End Function

Private Function IsContainerCtrl(ByVal ctrl As Object) As Boolean
    ''
    ' Determines whether the specified control is a container control (e.g., Frame, MultiPage, Page, UserForm)
    ' by checking if it implements the IOleContainer COM interface.
    '
    ' @param ctrl [In] The target control or UserForm object to check.
    ' @return Boolean True if the object implements IOleContainer; otherwise, False.
    ''
    IsContainerCtrl = False
    If ctrl Is Nothing Then Exit Function
    
    Dim pUnk As LongPtr
    pUnk = ObjPtr(ctrl)
    If pUnk = 0 Then Exit Function
    
    ' IID_IOleContainer: {0000011B-0000-0000-C000-000000000046}
    Dim iidContainer(3) As Long
    iidContainer(0) = &H11B: iidContainer(1) = 0: iidContainer(2) = &HC0: iidContainer(3) = &H46000000
    
    Dim pContainer As LongPtr
    Dim hr As Long
    
    ' QueryInterface (QI: VTable Index 0) for IOleContainer
    hr = dcf(pUnk, 0, "QI", VarPtr(iidContainer(0)), VarPtr(pContainer))
    
    If hr = 0 And pContainer <> 0 Then
        IsContainerCtrl = True
        ' Release the retrieved interface (Release: VTable Index 2)
        Call dcf(pContainer, 2, "Release")
    End If
End Function

' IID_IOleContainer: {0000011B-0000-0000-C000-000000000046}
' IID_IEnumUnknown:  {00000100-0000-0000-C000-000000000046}
Private Sub TraceZOrderRecursive(ByVal objContainer As Object, ByVal depth As Long, ByVal currentHierarchy As String, ByRef resultColl As Collection, ByVal useVisualOrder As Boolean)
    ' Internal recursive subroutine that enumerates child controls of a container via COM interfaces
    ' (IOleContainer and IEnumObjects) to accurately determine Z-Order precedence.
    '
    ' @param objContainer      [In]     The container object (UserForm, Frame, MultiPage, or Page).
    ' @param depth             [In]     Current recursion depth level (1-based).
    ' @param currentHierarchy  [In]     Hierarchical Z-Order string of the parent container (e.g., "1-2").
    ' @param resultColl        [In/Out] Target Collection to populate with control metadata.
    ' @param useVisualOrder    [In]     If True, enumerates Page controls within a MultiPage based on their visual
    '                                   tab position rather than internal COM Z-Order.
    Dim pUnk As LongPtr: pUnk = ObjPtr(objContainer)
    Dim pContainer As LongPtr, pEnum As LongPtr
    Dim hr As Long
    Dim children As Collection
    Dim foundType As String
    Dim zOrderIndex As Long
    Dim zOrderHierarchy As String
    
    ' 1. Get IOleContainer
    Dim iidContainer(3) As Long
    iidContainer(0) = &H11B: iidContainer(1) = 0: iidContainer(2) = &HC0: iidContainer(3) = &H46000000
    hr = dcf(pUnk, 0, "QI", VarPtr(iidContainer(0)), VarPtr(pContainer))
    If hr <> 0 Then Exit Sub

    ' 2. EnumObjects
    hr = dcf(pContainer, 4, "EnumObjects", 3, VarPtr(pEnum))
    
    If hr = 0 And pEnum <> 0 Then
        Dim cnt As Long
        
        ' MultiPage -> Get direct Page controls,
        ' Other Containers(e.g., UserForm/Frame/Page) -> Get direct child controls (only controls whose .Parent matches),
        Set children = GetDirectChildCtrls(objContainer)
        cnt = children.count
        
        If cnt = 0 Then GoTo CleanUp ' Exit if empty
        
        Dim pElements() As LongPtr: ReDim pElements(cnt) ' Allocate with extra capacity
        Dim fetchedCount As Long
        
        Call dcf(pEnum, 5, "Reset")
        
        ' Fetch all specified items at once
        hr = dcf(pEnum, 3, "Next", cnt + 1, VarPtr(pElements(0)), VarPtr(fetchedCount))


        Dim i As Long
        Dim validCount As Long ' Tracks 1-based Z-Order index among matched child controls only (skips hidden internal COM objects)
        ' Calculate 1-based Z-Order index using validCount instead of array index 'i'.
        ' This avoids skipping indices (e.g., "2-2" instead of "2-1") when IOleContainer
        ' returns internal non-control COM objects at the beginning (such as in MultiPage).
        validCount = 0
        
        Dim c As Object
        ' Always declare variable c as Object or Variant.
        ' Declaring it as MsForms.Control may cause some properties of the retrieved control to behave incorrectly.
        ' For example, the Caption property of a Page control is retrieved correctly when c is declared as Object,
        ' but returns an empty string when c is declared as MsForms.Control.
        
        If TypeName(objContainer) = "MultiPage" And useVisualOrder Then
            For Each c In children
                validCount = validCount + 1
                zOrderIndex = validCount
                If currentHierarchy = "" Then
                    zOrderHierarchy = CStr(zOrderIndex)
                Else
                    zOrderHierarchy = currentHierarchy & "-" & CStr(zOrderIndex)
                End If
                resultColl.Add VBA.Array(c, c.Name, zOrderIndex, zOrderHierarchy, depth), Hex(ObjPtr(c))
            
                If IsContainerCtrl(c) Then
                    TraceZOrderRecursive c, depth + 1, zOrderHierarchy, resultColl, useVisualOrder
                End If
            
            Next c
                        
        Else
            For i = 0 To fetchedCount - 1
                foundType = ""
                If pElements(i) <> 0 Then
                    Dim foundName As String: foundName = "Unknown"
                    Dim foundObj As Object: Set foundObj = Nothing
                    
                    ' Normalize IUnknown (for comparison)
                    Dim pUnkReal As LongPtr
                    Dim iidUnk(3) As Long: iidUnk(0) = &H0: iidUnk(1) = 0: iidUnk(2) = &HC0: iidUnk(3) = &H46000000
                    Call dcf(pElements(i), 0, "QI", VarPtr(iidUnk(0)), VarPtr(pUnkReal))
                    
                    ' Match with controls in the current container
    
                    For Each c In children
                        Dim pUnkCtrl As LongPtr
                        Call dcf(ObjPtr(c), 0, "QI", VarPtr(iidUnk(0)), VarPtr(pUnkCtrl))
                        
                        If pUnkReal = pUnkCtrl Then
                            foundName = c.Name
                            foundType = TypeName(c)
                            Set foundObj = c
                            Call dcf(pUnkCtrl, 2, "Release")
                            Exit For
                        End If
                        Call dcf(pUnkCtrl, 2, "Release")
                    Next
                    
                    If foundType <> "" Then
                        validCount = validCount + 1 ' Increment index only when a valid child control is matched
                        zOrderIndex = validCount
                        If currentHierarchy = "" Then
                            zOrderHierarchy = CStr(zOrderIndex)
                        Else
                            zOrderHierarchy = currentHierarchy & "-" & CStr(zOrderIndex)
                        End If
                        resultColl.Add VBA.Array(foundObj, foundName, zOrderIndex, zOrderHierarchy, depth), Hex(ObjPtr(foundObj))
                    End If
                    
                    ' If the found item is a container control(Frame/MultiPage/Page/UserForm), scan inside it recursively
                    If Not foundObj Is Nothing Then
                        If IsContainerCtrl(foundObj) Then
                            TraceZOrderRecursive foundObj, depth + 1, zOrderHierarchy, resultColl, useVisualOrder
                        End If
                    End If
                    
                    ' Cleanup
                    Call dcf(pUnkReal, 2, "Release")
                    Call dcf(pElements(i), 2, "Release")
                End If
            Next
        End If
    End If

CleanUp:
    If pEnum <> 0 Then Call dcf(pEnum, 2, "Release")
    If pContainer <> 0 Then Call dcf(pContainer, 2, "Release")
End Sub


Private Function dcf(ByVal ptr As LongPtr, ByVal vtblIndex As Long, ByVal funcName As String, ParamArray args() As Variant) As Long
    ' Helper wrapper around the Windows API 'DispCallFunc' to invoke COM VTable functions by index.
    '
    ' Note: Copies ParamArray elements into a heap-allocated dynamic array before invocation
    ' to prevent pointer invalidation during execution.
    '
    ' @param ptr       [In] Memory address of the COM interface pointer (IUnknown/IOleContainer/etc.).
    ' @param vtblIndex [In] 0-based method index in the COM Virtual Method Table (VTable).
    ' @param funcName  [In] Descriptive name of the COM method (used for logging/debugging).
    ' @param args()    [In] Variable parameter list passed to the target COM method.
    ' @return Long HRESULT status code from DispCallFunc or the return value of the COM method.
    
    'Debug.Print "dcf called for " & funcName
    Dim l As Long: l = LBound(args)
    Dim u As Long: u = UBound(args)
    Dim cnt As Long: cnt = u - l + 1
    Dim hr As Long, res As Variant
    Dim args_Type() As Integer
    Dim args_Ptr() As LongPtr
    Dim localVar() As Variant
    Const CC_STDCALL As Long = 4 ' Calling convention for the 3rd argument of DispCallFunc (__stdcall)
    ' IMPORTANT: Do NOT use VarPtr(args(i)) directly.
    ' ParamArray elements are temporary Variants managed by the VBA runtime stack.
    ' Their addresses become invalid by the time DispCallFunc internally reads rgpvarg,
    ' causing the COM method to receive garbage values.
    ' Copying into a heap-allocated dynamic array (localArgs) ensures the Variant
    ' addresses remain stable throughout the DispCallFunc call.
    If cnt > 0 Then
        ReDim args_Type(l To u): ReDim args_Ptr(l To u): ReDim localVar(l To u)
        Dim i As Long
        For i = l To u
            localVar(i) = args(i)
            args_Type(i) = VarType(localVar(i))
            args_Ptr(i) = VarPtr(localVar(i))
            'Debug.Print "args(" & i & ")", "Type:" & args_Type(i), "Ptr:" & Hex(args_Ptr(i)),"Value:" & localVar(i)
        Next
        hr = DispCallFunc(ptr, vtblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, args_Type(l), args_Ptr(l), res)
    Else
        hr = DispCallFunc(ptr, vtblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, 0, 0, res)
    End If
    If hr = 0 Then
        If res <> 0 Then
            'Debug.Print funcName & " res:" & res
        End If
        dcf = res
    Else
        'Debug.Print funcName & "  hr:" & hr
        dcf = hr
    End If
End Function
