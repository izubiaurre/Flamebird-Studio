Attribute VB_Name = "modApplication"
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com            (JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Darío Cutillas)
'   Zubiaurre:  izubiaurre@users.sourceforge.net (Imanol Zubiaurre)
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

Public G_ProcHelpLine As Integer  ' -1: don't show, 0: upper, 1: under

Public fenixDir As String
Public R_Compiler As Integer
Public R_Debug As Boolean
Public R_Stub As Boolean
Public R_AutoDeclare As Boolean
Public R_MsDos As Boolean
Public R_DebugDCB As Boolean
Public R_Paths As Boolean
'Public R_filter As Boolean
'Public R_DoubleBuf As Boolean
Public R_SaveBeforeCompiling As Integer
Public A_StyleXP As Boolean 'Use XP menu and toolbars look
Public A_Bitmaps As Boolean
Public A_Color As Integer
Public M_Style As Variant 'Menu style
Public T_Style As Variant 'Toolbar style

Public IS_Show As Boolean
Public IS_Sensitive As Integer
Public IS_LangDefConst As Boolean
Public IS_LangDefVar As Boolean
Public IS_LangDefFunc As Boolean
Public IS_UserDefConst As Boolean
Public IS_UserDefVar As Boolean
Public IS_UserDefFunc As Boolean
Public IS_UserDefProc As Boolean

Public PI_ShowConsts As Boolean
Public PI_ShowGlobals As Boolean
Public PI_ShowLocals As Boolean
Public PI_ShowPrivates As Boolean
Public PI_OnlyConstHeader As Boolean
Public PI_OnlyGlobalHeader As Boolean
Public PI_OnlyLocalHeader As Boolean

Public operatorList() As String
Public typeList() As String
Public sentenceList() As String
Public preprocessList() As String
Public constList() As String
Public globalList() As String
Public localList() As String
Public globalStructList() As String
Public localStructList() As String

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

'XP Theme
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

'Stores opened project
Public openedProject As cProject

' Command history ---------------------------------------------------------------
Private Type CALL_LIST
    commands(100) As String
    lastCommandIndex As Integer
    paths(100) As String
    lastPathIndex As Integer
    ' to run out this form the last command
    publicLastCommand As String
    publicLastPath As String
End Type

Public callList As CALL_LIST
Public lastCommandEnabled As Boolean


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
    Dim cntTypes As Integer, cntOperators As Integer, cntSentences As Integer
    Dim cntPreprocesses As Integer, cntConsts As Integer, cntGlobals As Integer
    Dim cntLocals As Integer, cntGlobalStructs As Integer, cntLocalStructs As Integer
    
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
    Erase typeList
    Erase operatorList
    Erase sentenceList
    Erase preprocessList
    Erase constList
    Erase globalList
    Erase localList
    Erase globalStructList
    Erase localStructList
    
    cntTypes = 0
    cntOperators = 0
    cntSentences = 0
    cntPreprocesses = 0
    cntConsts = 0
    cntGlobals = 0
    cntLocals = 0
    cntGlobalStructs = 0
    cntLocalStructs = 0

    Open App.Path & "\Help\fdl.lan" For Input As #num
        Do Until EOF(num)
            ' reads a line
            Line Input #num, linea
            If InStr(linea, "//#") <> 1 And linea <> "" Then 'si no es una palabra de seccion
                linea = Trim(linea)
                Select Case Seccion
                Case "SENTENCE"
                    ReDim Preserve sentenceList(cntSentences) As String
                    sentenceList(cntSentences) = linea
                    cntSentences = cntSentences + 1
                    
                    If Str_keyw = "" Then
                        Str_keyw = linea
                    Else
                        Str_keyw = Str_keyw & Chr(10) & linea
                    End If
                Case "PREPROCESS"
                    ReDim Preserve preprocessList(cntPreprocesses) As String
                    preprocessList(cntPreprocesses) = linea
                    cntPreprocesses = cntPreprocesses + 1
                    
                     If Str_keyw = "" Then
                        Str_keyw = linea
                    Else
                        Str_keyw = Str_keyw & Chr(10) & linea
                    End If
                Case "TYPE"
                    ReDim Preserve typeList(cntTypes) As String
                    typeList(cntTypes) = linea
                    cntTypes = cntTypes + 1
                    
                    If Str_keyw = "" Then
                        Str_keyw = linea
                    Else
                        Str_keyw = Str_keyw & Chr(10) & linea
                    End If
                Case "OPERATOR"
                    ReDim Preserve operatorList(cntOperators) As String
                    operatorList(cntOperators) = linea
                    cntOperators = cntOperators + 1

                    If Str_op = "" Then
                        Str_op = linea
                    Else
                        Str_op = Str_op & Chr(10) & linea
                    End If
                Case "CONST"
                    ReDim Preserve constList(cntConsts) As String
                    constList(cntConsts) = linea
                    cntConsts = cntConsts + 1
                    
                    If Str_data = "" Then
                        Str_data = linea
                    Else
                        Str_data = Str_data & Chr(10) & linea
                    End If
                Case "GLOBAL"
                    ReDim Preserve globalList(cntGlobals) As String
                    globalList(cntGlobals) = linea
                    cntGlobals = cntGlobals + 1
                    
                    If Str_data = "" Then
                        Str_data = linea
                    Else
                        Str_data = Str_data & Chr(10) & linea
                    End If
                Case "LOCAL"
                    ReDim Preserve localList(cntLocals) As String
                    localList(cntLocals) = linea
                    cntLocals = cntLocals + 1
                    
                    If Str_data = "" Then
                        Str_data = linea
                    Else
                        Str_data = Str_data & Chr(10) & linea
                    End If
                Case "GLOBAL_STRUCT"
                    ReDim Preserve globalStructList(cntGlobalStructs) As String
                    globalStructList(cntGlobalStructs) = linea
                    cntGlobalStructs = cntGlobalStructs + 1
                    
                     If Str_data = "" Then
                        Str_data = linea
                    Else
                        Str_data = Str_data & Chr(10) & linea
                    End If
                Case "LOCAL_STRUCT"
                    ReDim Preserve localStructList(cntLocalStructs) As String
                    localStructList(cntLocalStructs) = linea
                    cntLocalStructs = cntLocalStructs + 1
                    
                    If Str_data = "" Then
                        Str_data = linea
                    Else
                        Str_data = Str_data & Chr(10) & linea
                    End If


'                Case "SENTENS", "PREPROCESS", "DATATYPE": 'keywords
'                    'la agrega a la lista de palabras reservadas
'                    If Str_keyw = "" Then
'                        Str_keyw = linea
'                    Else
'                        Str_keyw = Str_keyw & Chr(10) & linea
'                    End If
'
'                    If Seccion = "DATATYPE" Then
'                        ReDim Preserve typeList(cntTypes) As String
'                        typeList(cntTypes) = linea
'                        cntTypes = cntTypes + 1
'                    End If
'
'                Case "SIMBOLS": 'operadores
'                    ReDim Preserve operatorList(cntOperators) As String
'                    operatorList(cntOperators) = linea
'                    cntOperators = cntOperators + 1
'
'                    'la agrega a la lista de palabras reservadas
'                    If Str_op = "" Then
'                        Str_op = linea
'                    Else
'                        Str_op = Str_op & Chr(10) & linea
'                    End If
'                Case "MISC": 'locales, globales y constantes
'
'                    'la agrega a la lista de palabras reservadas
'                    If Str_data = "" Then
'                        Str_data = linea
'                    Else
'                        Str_data = Str_data & Chr(10) & linea
'                    End If
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
Public Property Get StyleItem(Index As Integer) As clrStyle
    StyleItem = styles(Index)
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
        cs.font.name = .Value
        .Key = "FontSize"
        .Default = "12"
        cs.font.Size = CLng(Val(.Value))
        For i = 0 To UBound(styles)
            .Key = styles(i).name
            cs.SetColor styles(i).cmItem, CLng(Val(.Value))
            If styles(i).extended Then
                .Key = styles(i).name & "Bk"
                cs.SetColor styles(i).cmItem + 1, CLng(Val(.Value))
            End If
            If styles(i).cmStyleItem > -1 Then
                .Key = styles(i).name & "Style"
                cs.SetFontStyle styles(i).cmStyleItem, CLng(Val(.Value))
            End If
        Next
        .Key = "LineNumering"
        cs.LineNumbering = CBool(CInt(.Value))
        .Key = "BookmarkMargin"
        cs.DisplayLeftMargin = CBool(CInt(.Value))
        .Key = "ColorSyntax"
        cs.ColorSyntax = CBool(CInt(.Value))
        .Key = "NormalizeCase"
        cs.NormalizeCase = CBool(CInt(.Value))
        .Key = "DisplayWhiteSpaces"
        cs.DisplayWhitespace = CBool(CInt(.Value))
        .Key = "SmoothScrolling"
        cs.SmoothScrolling = CBool(CInt(.Value))
        .Key = "ConfineCaretToText"
        cs.SelBounds = CInt(.Value)
        .Key = "IndentMode"
        cs.AutoIndentMode = CInt(.Value)
        .Key = "TabSize"
        cs.TabSize = CInt(.Value)
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
        .Value = cs.font.name
        .Key = "FontSize"
        .Value = CStr(cs.font.Size)
        'Styles
        For i = 0 To UBound(styles)
            .Key = styles(i).name
            .Value = "&H" & hex(cs.GetColor(styles(i).cmItem)) & "&"
            If styles(i).extended Then
                .Key = styles(i).name & "Bk"
                .Value = "&H" & hex(cs.GetColor(styles(i).cmItem + 1)) & "&"
            End If
            If styles(i).cmStyleItem > -1 Then
                .Key = styles(i).name & "Style"
                .Value = "&H" & hex(cs.GetFontStyle(styles(i).cmStyleItem)) & "&"
            End If
        Next
        .Key = "LineNumering"
        .Value = CStr(CInt(cs.LineNumbering))
        .Key = "BookmarkMargin"
        .Value = CStr(CInt(cs.DisplayLeftMargin))
        .Key = "ColorSyntax"
        .Value = CStr(CInt(cs.ColorSyntax))
        .Key = "NormalizeCase"
        .Value = CStr(CInt(cs.NormalizeCase))
        .Key = "DisplayWhiteSpaces"
        .Value = CStr(CInt(cs.DisplayWhitespace))
        .Key = "SmoothScrolling"
        .Value = CStr(CInt(cs.SmoothScrolling))
        .Key = "ConfineCaretToText"
        .Value = CStr(CInt(cs.SelBounds))
        .Key = "IndentMode"
        .Value = CStr(cs.AutoIndentMode)
        .Key = "TabSize"
        .Value = CStr(cs.TabSize)
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

' Loads commands history
Public Sub loadCommandHistory()
    Dim sCommands As String, sPaths As String
    Dim i As Integer
    
    On Error GoTo errhandler:
    
       With Ini
        .Path = App.Path & "\Conf\command.ini"
        .Section = "General"
        .Key = "lastCommand"
        callList.lastCommandIndex = .Value
        .Key = "lastPath"

        callList.lastPathIndex = .Value
        
        .Section = "Commands"
        For i = 0 To 100
            .Key = "cmd_" & i
            If i <= callList.lastCommandIndex Then
                callList.commands(i) = .Value
            Else
                callList.commands(i) = ""
            End If
        Next i

        .Section = "Paths"
        For i = 0 To 100
            .Key = "path_" & i
            If i <= callList.lastPathIndex Then
                callList.paths(i) = .Value
            Else
                callList.paths(i) = ""
            End If
        Next i
    End With
    
    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("loadCommandHistory")
End Sub

'Loads the configuration file Config.ini
Private Sub LoadConf()
    With Ini
        .Path = App.Path & CONF_FILE
        
        .Section = "General"
        .Key = "ProcHelpLine"
        .Default = "1"
        G_ProcHelpLine = IIf(.Value = "0", 0, IIf(.Value = "1", 1, -1))
        
        .Section = "Appearance"
        
        .Key = "XPStyle"
        .Default = "1"
        A_StyleXP = IIf(.Value = "1", True, False)
        
        .Key = "BitmapBacks"
        .Default = "0"
        A_Bitmaps = IIf(.Value = "1", True, False)
        
        .Key = "Color"
        .Default = "1"
        A_Color = IIf(.Value = "1" Or .Value = "2" Or .Value = "3" Or .Value = "4" Or .Value = "5" Or .Value = "6" Or .Value = "7" Or .Value = "8" Or .Value = "9" Or .Value = "0", .Value, 1)
                
        .Section = "Run"
        
        .Key = "Compiler"
        
        If .Value = 0 Then
            .Key = "FenixPath"
        Else
            .Key = "BennuPath"
        End If
        
        .Default = " "
        fenixDir = .Value
        
        .Key = "Compiler"
        .Default = "1"
        R_Compiler = IIf(.Value = True, "0", "1")

        .Key = "Debug"
        .Default = "1"
        R_Debug = IIf(.Value = "1", True, False)
        
        .Key = "AutoDeclare"
        .Default = "1"
        R_AutoDeclare = IIf(.Value = 1, True, False)
        
        .Key = "Stub"
        .Default = "0"
        R_Stub = IIf(.Value = 1, True, False)
        
        .Key = "MsDos"
        .Default = "0"
        R_MsDos = IIf(.Value = 1, True, False)
        
        .Key = "DebugDCB"
        .Default = "1"
        R_DebugDCB = IIf(.Value = "1", True, False)
        
        .Key = "Paths"
        .Default = "1"
        R_Paths = IIf(.Value = "1", True, False)
        
'        .Key = "Filter"
'        .Default = "0"
'        R_filter = IIf(.value = "1", True, False)
'
'        .Key = "DoubleBuffer"
'        .Default = "0"
'        R_DoubleBuf = IIf(.value = "1", True, False)
        
        .Key = "SaveBeforeCompiling"
        .Default = "0"
        R_SaveBeforeCompiling = IIf(.Value = "1", 1, IIf(.Value = "2", 2, IIf(.Value = "3", 3, 0)))
        
        .Section = "IntelliSense"

        .Key = "Show"
        .Default = "1"
        IS_Show = IIf(.Value = 1, True, False)

        'If IS_Show Then

            .Key = "Sensitive"
            .Default = "2"
            IS_Sensitive = CLng(.Value)

            .Key = "LangDefConst"
            .Default = "1"
            IS_LangDefConst = IIf(.Value = 1, True, False)

            .Key = "LangDefVar"
            .Default = "1"
            IS_LangDefVar = IIf(.Value = 1, True, False)

            .Key = "LangDefFunc"
            .Default = "1"
            IS_LangDefFunc = IIf(.Value = 1, True, False)

            .Key = "UserDefConst"
            .Default = "1"
            IS_UserDefConst = IIf(.Value = 1, True, False)

            .Key = "UserDefvar"
            .Default = "1"
            IS_UserDefVar = IIf(.Value = 1, True, False)

            .Key = "UserDefFunc"
            .Default = "1"
            IS_UserDefFunc = IIf(.Value = 1, True, False)
            
            .Key = "UserDefProc"
            .Default = "1"
            IS_UserDefProc = IIf(.Value = 1, True, False)

        'End If

        .Section = "ProgramInspector"

        .Key = "ShowConsts"
        .Default = "1"
        PI_ShowConsts = IIf(.Value = 1, True, False)

        .Key = "ShowGlobals"
        .Default = "1"
        PI_ShowGlobals = IIf(.Value = 1, True, False)

        .Key = "ShowLocals"
        .Default = "1"
        PI_ShowLocals = IIf(.Value = 1, True, False)

        .Key = "ShowPrivates"
        .Default = "1"
        PI_ShowPrivates = IIf(.Value = 1, True, False)

        .Key = "OnlyConstHeader"
        .Default = "1"
        PI_OnlyConstHeader = IIf(.Value = 1, True, False)

        .Key = "OnlyGlobalHeader"
        .Default = "1"
        PI_OnlyGlobalHeader = IIf(.Value = 1, True, False)

        .Key = "OnlyLocalHeader"
        .Default = "1"
        PI_OnlyLocalHeader = IIf(.Value = 1, True, False)
        
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
        Or FileAssociated(".prg", "Bennu/Fenix.Source") = False _
        Or FileAssociated(".map", "Bennu/Fenix.ImageFile") = False Or _
        FileAssociated(".dcb", "Bennu/Fenix.Bin") = False Then
        With Ini
            .Path = App.Path & CONF_FILE
            .Section = "General"
            .Key = "AskFileRegister"
            .Default = "1"
            ask = IIf(.Value = "1", True, False)
        End With
        
        If ask = True Then
            frmRegisterFiletypes.Show vbModal
        End If
    End If
End Sub

'Sets the lbl message of the Splash form
Public Sub SetSplashMessage(sMsg As String)
    frmSplash.lblMessage.Caption = sMsg
    DoEvents
End Sub

'APPLICATION ENTRY POINT
Public Sub Main()
    On Error GoTo errhandler
    
    LoadConf
    
    ' Select XP / Normal style
    If A_StyleXP Then
        InitCommonControlsVB
    End If
    
    frmSplash.Show
    
    'Register dlls and ocx of the related dir
    RegisterAppComponents
    
    'Language definition
    SetSplashMessage "Loading languaje definition"
    LoadLan
    
    'Command History
    SetSplashMessage "Loading command history"
    loadCommandHistory
    
    'Editor Styles
    SetSplashMessage "Initializating editor styles"
    InitStyles
    
    'General configuration
    'SetSplashMessage "Loading general configuration"
    'LoadConf
    '
    'If A_StyleXP Then
    '    InitCommonControlsVB
    'End If
    
    'Init FMOD sound system
    SetSplashMessage "Initializing Audio System"
    initFMOD
    
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
    MsgBox "Error in " & str & vbCrLf & "Description: " & Err.description _
            & vbCrLf & "Number: " & Err.Number, vbCritical '& vbCrLf & Err.Source
    Err.Clear
End Sub

' Just for XP themes
Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function
