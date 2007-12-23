Attribute VB_Name = "modApplication"
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com(JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Darío Cutillas)
'   Izubiaurre: izubiaurre@users.sourceforge.net (Imanol Izubiaurre)
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

Option Explicit

Public Const CONF_FILE As String = "\conf\config.ini"
Public Const RES_FOLDER As String = "\Resources"

Public Ini As New cInifile

Public FSO As New FileSystemObject

Public fenixDir As String
Public R_Debug As Boolean
Public R_filter As Boolean
Public R_DoubleBuf As Boolean
Public R_SaveBeforeCompiling As Integer
Public A_StyleXP As Boolean 'Use XP menu and toolbars look
Public A_Bitmaps As Boolean
Public M_Style As Variant 'Menu style
Public T_Style As Variant 'Toolbar style


Public allOperatorsList() As String
Public allTypeList() As String

Public Propiedades As Globals
Public Fenix As Language

'Type for storing cmColorItems info (for editor configuration)
Public Type clrStyle
    name As String
    cmItem As cmColorItem
    cmStyleItem As cmFontStyleItem
    extended As Boolean 'If true, the clrStyle has two colors (background and foreground)
End Type
Private styles() As clrStyle

'Stores opened project
Public openedProject As cProject

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START ACTIVE X & DLL
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'Registers Active X & dlls controls form the related directory
Private Sub RegisterAppComponents()
    Dim fileString As String
    Dim result As Variant
    fileString = Dir(App.Path & "\Related\")
    fileString = LCase(fileString)
    Do Until fileString = ""
        If Right(fileString, 4) = ".dll" Or _
        Right(fileString, 4) = ".ocx" Or _
        Right(fileString, 4) = ".exe" Then
            
            SetSplashMessage "Registering " & fileString
            
            DoEvents
            result = Register(App.Path & "\Related\" & fileString)
            
            Select Case result
                Case 1: MsgBox "File Could Not Be Loaded Into Memory Space "
                Case 2: MsgBox "Not A Valid ActiveX Component"
                Case 3: MsgBox "ActiveX Component Registration Failed"
                Case 5: MsgBox "ActiveX Component UnRegister Successful"
                Case 6: MsgBox "ActiveX Component UnRegistration Failed"
                Case 7: MsgBox "No File Provided"
            End Select
            
            DoEvents
        End If
        fileString = Dir
        fileString = LCase(fileString)
    Loop
    
End Sub
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END ACTIVE X & DLL
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START LANGUAGE
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'Loads language definition
Private Sub LoadLan()
    Dim linea As String
    Dim Seccion As String, Str_func As String, Str_keyw As String
    Dim Str_op As String, Str_data As String
    Dim num As Integer, num1 As Integer, Contador1 As Integer
    Dim cntTypes As Integer, cntOperators As Integer
    
    On Error GoTo errhandler
    Str_func = ""
    Str_keyw = ""
    Seccion = -2
    num = FreeFile
    
    If FSO.FileExists(App.Path & "\Help\fdl.lan") = False Then
        MsgBox App.Path & "\Help\fdl.lan" & " is missing!. Applicatio termited.", vbCritical, "Fatal error"
        End
    End If
    
    'Clear types and operators lists
    Erase allTypeList
    Erase allOperatorsList
    cntTypes = 0
    cntOperators = 0

    Open App.Path & "\Help\fdl.lan" For Input As #num
        Do Until EOF(num)
            'lee una linea
            Line Input #num, linea
            If InStr(linea, "//#") <> 1 And linea <> "" Then 'si no es una palabra de seccion
                linea = Trim(linea)
                Select Case Seccion
                Case "SENTENS", "PREPROCESS", "DATATYPE": 'keywords
                    'la agrega a la lista de palabras reservadas
                    If Str_keyw = "" Then
                        Str_keyw = linea
                    Else
                        Str_keyw = Str_keyw & Chr(10) & linea
                    End If
                    
                    If Seccion = "DATATYPE" Then
                        ReDim Preserve allTypeList(cntTypes) As String
                        allTypeList(cntTypes) = linea
                        cntTypes = cntTypes + 1
                    End If
                    
                Case "SIMBOLS": 'operadores
                    ReDim Preserve allOperatorsList(cntOperators) As String
                    allOperatorsList(cntOperators) = linea
                    cntOperators = cntOperators + 1
                                    
                    'la agrega a la lista de palabras reservadas
                    If Str_op = "" Then
                        Str_op = linea
                    Else
                        Str_op = Str_op & Chr(10) & linea
                    End If
                Case "MISC": 'locales, globales y constantes
                    
                    'la agrega a la lista de palabras reservadas
                    If Str_data = "" Then
                        Str_data = linea
                    Else
                        Str_data = Str_data & Chr(10) & linea
                    End If
                End Select
            Else
                linea = replace(linea, "-", "")
                linea = Trim(linea)
                linea = Mid(linea, 4)
                
                Seccion = UCase(linea)
            End If
        Loop
    Close #num
    Set Propiedades = New Globals
    Set Fenix = New Language
    
    With Fenix
        .TagAttributeNames = Str_data
        .TagElementNames = "//#Section:" 'usado para IDEkeywords
        .CaseSensitive = False
        .StringDelims = Chr(34)
        .Keywords = Str_keyw
        .SingleLineComments = "//"
        .Operators = Str_op
        .MultiLineComments1 = "/*"
        .MultiLineComments2 = "*/"
    End With
    Propiedades.RegisterLanguage "Fenix", Fenix
    
    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("LoadLan"): Resume Next
End Sub
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END LANGUAGE
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START EDITOR STYLES
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'Add a new style to our ColorStyle array
Public Property Get StyleItem(index As Integer) As clrStyle
    StyleItem = styles(index)
End Property
Public Property Get StyleItemCount() As Integer
    StyleItemCount = UBound(styles) + 1
End Property

Private Sub AddStyle(name As String, cmItem As cmColorItem, extended As Boolean, cmStyleItem As cmFontStyleItem)
    Static stylescount As Integer
    
    ReDim Preserve styles(stylescount) As clrStyle
    With styles(stylescount)
        .name = name
        .cmItem = cmItem
        .cmStyleItem = cmStyleItem
        .extended = extended
    End With
    stylescount = stylescount + 1
End Sub

Private Sub InitStyles()
    Dim i As Integer
    
    AddStyle "Window", cmClrWindow, False, -1
    AddStyle "Text", cmClrText, True, cmStyText
    AddStyle "KeyWord", cmClrKeyword, True, cmStyKeyword
    AddStyle "Comment", cmClrComment, True, cmStyComment
    AddStyle "LineNumber", cmClrLineNumber, True, cmStyLineNumber
    AddStyle "Operator", cmClrOperator, True, cmStyOperator
    AddStyle "Number", cmClrNumber, True, cmStyNumber
    AddStyle "String", cmClrString, True, cmStyString
    AddStyle "ScopeKeyword", cmClrScopeKeyword, True, cmStyScopeKeyword
    AddStyle "LeftMargin", cmClrLeftMargin, False, -1
    AddStyle "Bookmark", cmClrBookmark, True, -1
    AddStyle "HighlightedLine", cmClrHighlightedLine, False, -1
    AddStyle "HDividerLines", cmClrHDividerLines, False, -1
    AddStyle "VDividerLines", cmClrVDividerLines, False, -1
    AddStyle "TagText", cmClrTagText, True, cmStyTagText
    AddStyle "TagEntity", cmClrTagEntity, True, cmStyTagEntity
    AddStyle "TagElementName", cmClrTagElementName, True, cmStyTagElementName
    AddStyle "TagAttributeName", cmClrTagAttributeName, True, cmStyTagAttributeName
End Sub

Public Sub LoadCSConf(cs As Object, Optional ByVal sConfFile As String)
    Dim i As Integer

    With Ini
        .Path = App.Path & "\Conf\editor.ini"
        If sConfFile <> "" Then
            If FSO.FileExists(sConfFile) Then
                .Path = sConfFile
            End If
        End If
        .Section = "EditorConfig"
        .Key = "Font"
        .Default = "Courier New"
        cs.font.name = .value
        .Key = "FontSize"
        .Default = "12"
        cs.font.Size = CLng(Val(.value))
        For i = 0 To UBound(styles)
            .Key = styles(i).name
            cs.SetColor styles(i).cmItem, CLng(Val(.value))
            If styles(i).extended Then
                .Key = styles(i).name & "Bk"
                cs.SetColor styles(i).cmItem + 1, CLng(Val(.value))
            End If
            If styles(i).cmStyleItem > -1 Then
                .Key = styles(i).name & "Style"
                cs.SetFontStyle styles(i).cmStyleItem, CLng(Val(.value))
            End If
        Next
        .Key = "LineNumering"
        cs.LineNumbering = CBool(CInt(.value))
        .Key = "BookmarkMargin"
        cs.DisplayLeftMargin = CBool(CInt(.value))
        .Key = "ColorSyntax"
        cs.ColorSyntax = CBool(CInt(.value))
        .Key = "NormalizeCase"
        cs.NormalizeCase = CBool(CInt(.value))
        .Key = "DisplayWhiteSpaces"
        cs.DisplayWhitespace = CBool(CInt(.value))
        .Key = "SmoothScrolling"
        cs.SmoothScrolling = CBool(CInt(.value))
        .Key = "ConfineCaretToText"
        cs.SelBounds = CInt(.value)
        .Key = "IndentMode"
        cs.AutoIndentMode = CInt(.value)
        .Key = "TabSize"
        cs.TabSize = CInt(.value)
    End With
End Sub

Public Sub SaveCSConf(cs As Object, Optional sConfFile As String)
    Dim i As Integer
    With Ini
        .Path = App.Path & "\Conf\editor.ini"
        If sConfFile <> "" Then
            .Path = sConfFile
        End If
        .Section = "EditorConfig"
        .Key = "Font"
        .value = cs.font.name
        .Key = "FontSize"
        .value = CStr(cs.font.Size)
        'Styles
        For i = 0 To UBound(styles)
            .Key = styles(i).name
            .value = "&H" & Hex(cs.GetColor(styles(i).cmItem)) & "&"
            If styles(i).extended Then
                .Key = styles(i).name & "Bk"
                .value = "&H" & Hex(cs.GetColor(styles(i).cmItem + 1)) & "&"
            End If
            If styles(i).cmStyleItem > -1 Then
                .Key = styles(i).name & "Style"
                .value = "&H" & Hex(cs.GetFontStyle(styles(i).cmStyleItem)) & "&"
            End If
        Next
        .Key = "LineNumering"
        .value = CStr(CInt(cs.LineNumbering))
        .Key = "BookmarkMargin"
        .value = CStr(CInt(cs.DisplayLeftMargin))
        .Key = "ColorSyntax"
        .value = CStr(CInt(cs.ColorSyntax))
        .Key = "NormalizeCase"
        .value = CStr(CInt(cs.NormalizeCase))
        .Key = "DisplayWhiteSpaces"
        .value = CStr(CInt(cs.DisplayWhitespace))
        .Key = "SmoothScrolling"
        .value = CStr(CInt(cs.SmoothScrolling))
        .Key = "ConfineCaretToText"
        .value = CStr(CInt(cs.SelBounds))
        .Key = "IndentMode"
        .value = CStr(cs.AutoIndentMode)
        .Key = "TabSize"
        .value = CStr(cs.TabSize)
    End With
End Sub
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END EDITOR STYLES
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

'Returns an string containig all the opened files belonging to the project
Public Function FileArray() As String
    Dim sFiles As String
    Dim ff As IFileForm
    Dim f As Form
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If openedProject.FileExist(ff.FilePath) Then 'The file belongs to the project
                sFiles = sFiles & ff.FilePath & "|"
            End If
        End If
    Next
    If Not sFiles = "" Then sFiles = Left(sFiles, Len(sFiles) - 1) 'Remove the last |
    FileArray = sFiles
End Function

'Loads the configuration file Config.ini
Private Sub LoadConf()
    With Ini
        .Path = App.Path & CONF_FILE
        .Section = "Appearance"
        
        .Key = "XPStyle"
        .Default = "1"
        A_StyleXP = IIf(.value = "1", True, False)
        
        .Key = "BitmapBacks"
        .Default = "0"
        A_Bitmaps = IIf(.value = "1", True, False)
        
        .Section = "Run"
        
        .Key = "FenixPath"
        .Default = " "
        fenixDir = .value
        
        .Key = "Debug"
        .Default = "1"
        R_Debug = IIf(.value = "1", True, False)
        
        .Key = "Filter"
        .Default = "0"
        R_filter = IIf(.value = "1", True, False)
        
        .Key = "DoubleBuffer"
        .Default = "0"
        R_DoubleBuf = IIf(.value = "1", True, False)
        
        .Key = "SaveBeforeCompiling"
        .Default = "0"
        R_SaveBeforeCompiling = IIf(.value = "1", 1, IIf(.value = "2", 2, IIf(.value = "3", 3, 0)))
    End With
  
    'Determine whether to use or not XP styles for menus and toolbars
    If A_StyleXP = True Then
        M_Style = mds_XP
        T_Style = CTBDrawOfficeXPStyle
    Else
        M_Style = mds_3D
        T_Style = CTBDrawStandard
    End If
End Sub

'Verify if supported file types are registered with FB
Private Sub CheckFileAssoc()
    Dim ask As Boolean
    
    'THIS SHOULD BE CHANGED TO SOMETHING MORE GENERAL
    If FileAssociated(".fbp", "FlameBird.Project") = False _
        Or FileAssociated(".prg", "Fenix.Source") = False _
        Or FileAssociated(".map", "Fenix.ImageFile") = False Or _
        FileAssociated(".dcb", "Fenix.Bin") = False Then
        With Ini
            .Path = App.Path & CONF_FILE
            .Section = "General"
            .Key = "AskFileRegister"
            .Default = "1"
            ask = IIf(.value = "1", True, False)
        End With
        
        If ask = True Then
            frmRegisterFiletypes.Show vbModal
        End If
    End If
End Sub

'Sets the lbl message of the Splash form
Public Sub SetSplashMessage(sMsg As String)
    frmSplash.lblMessage.text = sMsg
    DoEvents
End Sub

'APPLICATION ENTRY POINT
Public Sub Main()
    On Error GoTo errhandler
    
    frmSplash.Show
    
    'Register dlls and ocx of the related dir
    RegisterAppComponents
    
    'Language definition
    SetSplashMessage "Loading languaje definition"
    LoadLan
    
    'Editor Styles
    SetSplashMessage "Initializating editor styles"
    InitStyles
    
    'General configuration
    SetSplashMessage "Loading general configuration"
    LoadConf
    
    'Load External Tools
    SetSplashMessage "Loading external tools configuration"
    LoadExternalTools
    
    'FileFilter
    SetSplashMessage "Creating file filters"
    CreateFileFilters
    
    'Register user-defined plugins
    RegisterPlugins
    
    'Create color conversion tables
    SetSplashMessage "Loading conversion tables"
    initConversionTables
    
    'Verify if supported file types are registered with FB
    SetSplashMessage "Checking registration info"
    CheckFileAssoc
    
    'Load MDI form
    SetSplashMessage ("Loading main application")
    Load frmMain
    Unload frmSplash
    
    'FBMX Starts :D. GO Go go!
    frmMain.Show
    
    Exit Sub
    
errhandler:
    If Err.Number > 0 Then ShowError ("Main"): Resume Next
End Sub

'Just fot debuggin purposes
Public Sub ShowError(str As String)
    MsgBox "Error in " & str & vbCrLf & "Description: " & Err.Description _
            & vbCrLf & "Number: " & Err.Number, vbCritical
    Err.Clear
End Sub
